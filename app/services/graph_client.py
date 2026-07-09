"""
Microsoft Graph API client — async version using httpx.
All logic from the original graph_client.py is preserved; requests replaced by httpx.AsyncClient.
"""
import logging
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional

import httpx

from app.core.config import settings

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"


class GraphClient:
    """
    Async Microsoft Graph API client.
    Uses a single shared httpx.AsyncClient for connection pooling.
    Token is cached in-memory and refreshed automatically before expiry.
    """

    def __init__(self) -> None:
        self._http: Optional[httpx.AsyncClient] = None
        self._access_token: Optional[str] = None
        self._token_expires: Optional[datetime] = None

    # ------------------------------------------------------------------ #
    # Lifecycle — call open() / close() via FastAPI lifespan
    # ------------------------------------------------------------------ #

    async def open(self) -> None:
        """Create the underlying HTTP client. Call once at app startup."""
        self._http = httpx.AsyncClient(
            timeout=30.0,
            headers={"Content-Type": "application/json"},
        )
        logger.info("GraphClient HTTP session opened")

    async def close(self) -> None:
        """Close the HTTP client. Call on app shutdown."""
        if self._http:
            await self._http.aclose()
            logger.info("GraphClient HTTP session closed")

    # ------------------------------------------------------------------ #
    # Token management
    # ------------------------------------------------------------------ #

    async def _get_token(self) -> str:
        """Get a valid OAuth2 client-credentials token, refreshing if needed."""
        now = datetime.now(timezone.utc)
        if self._access_token and self._token_expires and now < self._token_expires:
            return self._access_token

        url = TOKEN_URL.format(tenant_id=settings.TENANT_ID)
        data = {
            "grant_type": "client_credentials",
            "client_id": settings.CLIENT_ID,
            "client_secret": settings.CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
        }

        assert self._http, "GraphClient not opened — call open() first"
        resp = await self._http.post(url, data=data)
        resp.raise_for_status()

        result = resp.json()
        self._access_token = result["access_token"]
        expires_in = result.get("expires_in", 3600)
        # Refresh 5 minutes before actual expiry
        self._token_expires = now + timedelta(seconds=expires_in - 300)

        logger.info("MS Graph token refreshed")
        return self._access_token

    async def _headers(self) -> Dict[str, str]:
        token = await self._get_token()
        return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # ------------------------------------------------------------------ #
    # User operations
    # ------------------------------------------------------------------ #

    async def get_users(self, filter_query: Optional[str] = None) -> List[Dict]:
        """Get all non-guest users in the organisation."""
        params: Dict = {
            "$select": "id,displayName,mail,userPrincipalName",
            "$top": 100,
        }
        if filter_query:
            params["$filter"] = filter_query

        assert self._http
        resp = await self._http.get(
            f"{GRAPH_BASE}/users", headers=await self._headers(), params=params
        )
        resp.raise_for_status()

        users = resp.json().get("value", [])
        return [
            u for u in users
            if u.get("mail") and "#EXT#" not in u.get("userPrincipalName", "")
        ]

    # ------------------------------------------------------------------ #
    # Email operations
    # ------------------------------------------------------------------ #

    async def get_unread_emails(
        self, user_id: str, since: Optional[datetime] = None
    ) -> List[Dict]:
        """Get unread emails for a user since a given datetime."""
        if since is None:
            since = datetime.now(timezone.utc) - timedelta(minutes=settings.LOOKBACK_MINUTES)

        since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")
        params = {
            "$filter": f"isRead eq false and receivedDateTime ge {since_str}",
            "$select": (
                "id,subject,bodyPreview,body,from,toRecipients,"
                "receivedDateTime,isRead,categories,hasAttachments"
            ),
            "$top": settings.EMAILS_PER_BATCH,
            "$orderby": "receivedDateTime desc",
        }

        assert self._http
        resp = await self._http.get(
            f"{GRAPH_BASE}/users/{user_id}/messages",
            headers=await self._headers(),
            params=params,
        )

        if resp.status_code == 404:
            logger.warning("User not found or no mailbox access", extra={"user_id": user_id})
            return []

        resp.raise_for_status()
        return resp.json().get("value", [])

    async def apply_category(
        self, user_id: str, message_id: str, category_name: str
    ) -> bool:
        """Apply an Outlook category to a message (preserving existing categories)."""
        assert self._http
        headers = await self._headers()

        # Fetch existing categories
        get_resp = await self._http.get(
            f"{GRAPH_BASE}/users/{user_id}/messages/{message_id}",
            headers=headers,
            params={"$select": "categories"},
        )
        existing = get_resp.json().get("categories", []) if get_resp.status_code == 200 else []

        new_categories = list(set(existing + [category_name]))

        patch_resp = await self._http.patch(
            f"{GRAPH_BASE}/users/{user_id}/messages/{message_id}",
            headers=headers,
            json={"categories": new_categories},
        )

        if patch_resp.status_code == 200:
            return True
        else:
            logger.warning(
                "Failed to apply Outlook category",
                extra={"status": patch_resp.status_code, "body": patch_resp.text[:200]},
            )
            return False

    async def mark_as_read(self, user_id: str, message_id: str) -> bool:
        """Mark an email as read."""
        assert self._http
        resp = await self._http.patch(
            f"{GRAPH_BASE}/users/{user_id}/messages/{message_id}",
            headers=await self._headers(),
            json={"isRead": True},
        )
        return resp.status_code == 200

    async def create_outlook_category(
        self, user_id: str, category_name: str, color: str = "preset0"
    ) -> bool:
        """Create an Outlook master category for a user if it doesn't exist."""
        assert self._http
        headers = await self._headers()
        url = f"{GRAPH_BASE}/users/{user_id}/outlook/masterCategories"

        # Check if it already exists
        resp = await self._http.get(url, headers=headers)
        if resp.status_code == 200:
            existing = resp.json().get("value", [])
            if any(c.get("displayName") == category_name for c in existing):
                return True

        create_resp = await self._http.post(
            url,
            headers=headers,
            json={"displayName": category_name, "color": color},
        )
        return create_resp.status_code == 201

    async def move_to_folder(
        self, user_id: str, message_id: str, folder_name: str
    ) -> bool:
        """Move a message to a named folder (creates it if not found)."""
        folder_id = await self._get_or_create_folder(user_id, folder_name)
        if not folder_id:
            return False

        assert self._http
        resp = await self._http.post(
            f"{GRAPH_BASE}/users/{user_id}/messages/{message_id}/move",
            headers=await self._headers(),
            json={"destinationId": folder_id},
        )
        return resp.status_code == 201

    async def _get_or_create_folder(
        self, user_id: str, folder_name: str
    ) -> Optional[str]:
        assert self._http
        headers = await self._headers()

        resp = await self._http.get(
            f"{GRAPH_BASE}/users/{user_id}/mailFolders", headers=headers
        )
        resp.raise_for_status()

        for folder in resp.json().get("value", []):
            if folder.get("displayName") == folder_name:
                return folder["id"]

        create_resp = await self._http.post(
            f"{GRAPH_BASE}/users/{user_id}/mailFolders",
            headers=headers,
            json={"displayName": folder_name, "isHidden": False},
        )
        if create_resp.status_code == 201:
            return create_resp.json().get("id")

        logger.error("Failed to create mail folder", extra={"folder": folder_name})
        return None


# ---------------------------------------------------------------------------
# Module-level singleton managed by FastAPI lifespan
# ---------------------------------------------------------------------------
graph_client = GraphClient()
