"""
Provident Packaging — Graph API Client v2.1
Scans ALL mailbox folders (Inbox, Junk, Archive, Sent, custom folders).
Also detects if a reply was sent in a conversation thread for auto-status-advance.
"""
import requests
from datetime import datetime, timedelta
from typing import List, Dict, Optional
from config import config

# Well-known folder names supported by Microsoft Graph
WELL_KNOWN_FOLDERS = [
    "inbox",
    "junkemail",
    "archive",
    "sentitems",
    "deleteditems",
    "drafts",
]

# Folders to SKIP when looking for new emails to classify
# (we do check sentitems but only for reply detection, not classification)
SKIP_CLASSIFY_FOLDERS = {"sentitems", "drafts"}


class GraphClient:
    """Microsoft Graph API client — all-folder email scanning + reply detection"""

    def __init__(self):
        self.tenant_id = config.TENANT_ID
        self.client_id = config.CLIENT_ID
        self.client_secret = config.CLIENT_SECRET
        self.access_token = None
        self.token_expires = None
        self.base_url = "https://graph.microsoft.com/v1.0"

    # ─── Auth ─────────────────────────────────────────────────────────────────

    def _get_token(self) -> str:
        if self.access_token and self.token_expires and datetime.utcnow() < self.token_expires:
            return self.access_token

        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default"
        }
        response = requests.post(url, data=data, timeout=30)
        response.raise_for_status()

        result = response.json()
        self.access_token = result["access_token"]
        expires_in = result.get("expires_in", 3600)
        self.token_expires = datetime.utcnow() + timedelta(seconds=expires_in - 300)
        return self.access_token

    def _get_headers(self) -> Dict:
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json"
        }

    # ─── Users ────────────────────────────────────────────────────────────────

    def get_users(self, filter_query: str = None) -> List[Dict]:
        url = f"{self.base_url}/users"
        params = {
            "$select": "id,displayName,mail,userPrincipalName",
            "$top": 100
        }
        if filter_query:
            params["$filter"] = filter_query

        response = requests.get(url, headers=self._get_headers(), params=params, timeout=30)
        response.raise_for_status()
        users = response.json().get("value", [])
        return [u for u in users if u.get("mail") and "#EXT#" not in u.get("userPrincipalName", "")]

    # ─── All-folder email scanning ────────────────────────────────────────────

    def get_all_mail_folders(self, user_id: str) -> List[Dict]:
        """
        Returns all mailbox folders for a user — well-known + custom.
        Each dict has: id, displayName, wellKnownName (may be None for custom).
        """
        all_folders = []

        # First fetch well-known folders (guaranteed IDs)
        for wk in WELL_KNOWN_FOLDERS:
            url = f"{self.base_url}/users/{user_id}/mailFolders/{wk}"
            try:
                r = requests.get(url, headers=self._get_headers(), timeout=15)
                if r.status_code == 200:
                    f = r.json()
                    all_folders.append({
                        "id": f["id"],
                        "displayName": f.get("displayName", wk),
                        "wellKnownName": wk
                    })
            except Exception:
                pass  # Folder doesn't exist for this user — skip

        # Also fetch custom (non-system) folders
        url = f"{self.base_url}/users/{user_id}/mailFolders"
        params = {"$top": 50}
        try:
            r = requests.get(url, headers=self._get_headers(), params=params, timeout=30)
            if r.status_code == 200:
                known_ids = {f["id"] for f in all_folders}
                for f in r.json().get("value", []):
                    if f["id"] not in known_ids:
                        all_folders.append({
                            "id": f["id"],
                            "displayName": f.get("displayName", "Unknown"),
                            "wellKnownName": None
                        })
        except Exception:
            pass

        return all_folders

    def get_emails_from_all_folders(self, user_id: str, since: datetime = None) -> List[Dict]:
        """
        Fetches new/unread emails from ALL mailbox folders.
        - Classifiable folders: Inbox, Junk, Archive, custom, DeletedItems
        - Skips: Sent, Drafts (those are outgoing, not incoming)
        - Tags each email with source_folder name.
        - De-duplicates by message_id.
        Returns list of email dicts, each with an extra 'source_folder' key.
        """
        if since is None:
            since = datetime.utcnow() - timedelta(minutes=config.LOOKBACK_MINUTES)

        since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")
        folders = self.get_all_mail_folders(user_id)

        seen_ids: set = set()
        all_emails: List[Dict] = []

        for folder in folders:
            wk = folder.get("wellKnownName", "")
            if wk in SKIP_CLASSIFY_FOLDERS:
                continue  # Don't classify outgoing mail

            folder_id = folder["id"]
            folder_name = folder["displayName"]

            url = f"{self.base_url}/users/{user_id}/mailFolders/{folder_id}/messages"
            params = {
                "$filter": f"receivedDateTime ge {since_str}",
                "$select": "id,conversationId,subject,bodyPreview,body,from,toRecipients,"
                           "receivedDateTime,isRead,categories,hasAttachments",
                "$top": config.EMAILS_PER_BATCH,
                "$orderby": "receivedDateTime desc"
            }

            try:
                response = requests.get(url, headers=self._get_headers(), params=params, timeout=30)
                if response.status_code == 404:
                    continue
                response.raise_for_status()
                emails = response.json().get("value", [])
                for email in emails:
                    mid = email.get("id", "")
                    if mid and mid not in seen_ids:
                        seen_ids.add(mid)
                        email["source_folder"] = folder_name
                        all_emails.append(email)
            except Exception as exc:
                print(f"  Warning: Could not fetch folder '{folder_name}': {exc}")

        return all_emails

    # ─── Legacy method (kept for backward compatibility) ──────────────────────

    def get_unread_emails(self, user_id: str, since: datetime = None) -> List[Dict]:
        """Deprecated: use get_emails_from_all_folders instead."""
        return self.get_emails_from_all_folders(user_id, since)

    # ─── Reply detection for auto-advance ─────────────────────────────────────

    def get_sent_replies_for_conversation(self, user_id: str, conversation_id: str,
                                          since: datetime = None) -> List[Dict]:
        """
        Checks the Sent Items folder for outgoing messages in a conversation.
        Returns list of sent message dicts if a reply was found.
        Used to auto-advance lifecycle status when a rep replies to a thread.
        """
        if since is None:
            since = datetime.utcnow() - timedelta(minutes=config.LOOKBACK_MINUTES)

        since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")
        url = f"{self.base_url}/users/{user_id}/mailFolders/sentitems/messages"
        params = {
            "$filter": f"conversationId eq '{conversation_id}' and sentDateTime ge {since_str}",
            "$select": "id,conversationId,subject,sentDateTime",
            "$top": 5
        }
        try:
            response = requests.get(url, headers=self._get_headers(), params=params, timeout=30)
            if response.status_code == 200:
                return response.json().get("value", [])
        except Exception:
            pass
        return []

    # ─── Email actions ────────────────────────────────────────────────────────

    def apply_category(self, user_id: str, message_id: str, category_name: str) -> bool:
        url = f"{self.base_url}/users/{user_id}/messages/{message_id}"
        get_response = requests.get(url, headers=self._get_headers(),
                                    params={"$select": "categories"}, timeout=30)
        existing = get_response.json().get("categories", []) if get_response.status_code == 200 else []
        new_categories = existing + [category_name] if category_name not in existing else existing
        response = requests.patch(url, headers=self._get_headers(),
                                  json={"categories": new_categories}, timeout=30)
        return response.status_code == 200

    def mark_as_read(self, user_id: str, message_id: str) -> bool:
        url = f"{self.base_url}/users/{user_id}/messages/{message_id}"
        response = requests.patch(url, headers=self._get_headers(), json={"isRead": True}, timeout=30)
        return response.status_code == 200

    def move_to_folder(self, user_id: str, message_id: str, folder_name: str) -> bool:
        folder_id = self._get_or_create_folder(user_id, folder_name)
        if not folder_id:
            return False
        url = f"{self.base_url}/users/{user_id}/messages/{message_id}/move"
        response = requests.post(url, headers=self._get_headers(),
                                 json={"destinationId": folder_id}, timeout=30)
        return response.status_code == 201

    def _get_or_create_folder(self, user_id: str, folder_name: str) -> Optional[str]:
        url = f"{self.base_url}/users/{user_id}/mailFolders"
        response = requests.get(url, headers=self._get_headers(), timeout=30)
        response.raise_for_status()
        folders = response.json().get("value", [])
        for folder in folders:
            if folder.get("displayName") == folder_name:
                return folder["id"]

        create_response = requests.post(url, headers=self._get_headers(),
                                        json={"displayName": folder_name, "isHidden": False}, timeout=30)
        if create_response.status_code == 201:
            return create_response.json().get("id")
        return None

    def create_outlook_category(self, user_id: str, category_name: str, color: str = "preset0") -> bool:
        url = f"{self.base_url}/users/{user_id}/outlook/masterCategories"
        response = requests.get(url, headers=self._get_headers(), timeout=30)
        if response.status_code == 200:
            existing = response.json().get("value", [])
            if any(c.get("displayName") == category_name for c in existing):
                return True
        response = requests.post(url, headers=self._get_headers(),
                                 json={"displayName": category_name, "color": color}, timeout=30)
        return response.status_code == 201
