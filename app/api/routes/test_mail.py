"""
Verification endpoint for testing Microsoft Graph mailbox access.
"""
import logging
from typing import Dict, List, Optional
import httpx
from fastapi import APIRouter, Header, HTTPException, status
from pydantic import BaseModel

logger = logging.getLogger(__name__)
router = APIRouter(tags=["Mail Verification"])


class TestMailResponse(BaseModel):
    success: bool
    user: Optional[Dict] = None
    emails: Optional[List[str]] = None
    error: Optional[str] = None


@router.get(
    "/api/test-mail",
    response_model=TestMailResponse,
    summary="Test Microsoft Graph Mailbox Access",
)
async def test_mail(
    authorization: Optional[str] = Header(None)
) -> TestMailResponse:
    """
    Call Microsoft Graph API using the delegated access token provided in the Authorization header.
    Tests:
    1. GET /me (Retrieve authenticated user's name/email)
    2. GET /me/messages?$top=5 (Retrieve the first 5 email subjects)
    """
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing or invalid Authorization header. Expected 'Bearer <microsoft_access_token>'.",
        )

    token = authorization.removeprefix("Bearer ").strip()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    async with httpx.AsyncClient(timeout=15.0) as client:
        try:
            # 1. GET https://graph.microsoft.com/v1.0/me
            me_resp = await client.get("https://graph.microsoft.com/v1.0/me", headers=headers)
            if me_resp.status_code != 200:
                logger.error(
                    "Graph API /me call failed",
                    extra={"status_code": me_resp.status_code, "response": me_resp.text},
                )
                return TestMailResponse(
                    success=False,
                    error=f"Graph API /me call failed ({me_resp.status_code}): {me_resp.text[:200]}",
                )

            me_data = me_resp.json()

            # 2. GET https://graph.microsoft.com/v1.0/me/messages?$top=5
            msg_resp = await client.get(
                "https://graph.microsoft.com/v1.0/me/messages?$top=5",
                headers=headers,
            )
            if msg_resp.status_code != 200:
                logger.error(
                    "Graph API /me/messages call failed",
                    extra={"status_code": msg_resp.status_code, "response": msg_resp.text},
                )
                return TestMailResponse(
                    success=False,
                    error=f"Graph API /me/messages call failed ({msg_resp.status_code}): {msg_resp.text[:200]}",
                )

            msg_data = msg_resp.json()

            user_info = {
                "displayName": me_data.get("displayName"),
                "mail": me_data.get("mail") or me_data.get("userPrincipalName"),
            }

            emails = [
                msg.get("subject", "(No Subject)")
                for msg in msg_data.get("value", [])
            ]

            return TestMailResponse(
                success=True,
                user=user_info,
                emails=emails,
                error=None,
            )

        except httpx.RequestError as exc:
            logger.exception("HTTP Request failed while calling Microsoft Graph")
            return TestMailResponse(
                success=False,
                error=f"Network request failed: {str(exc)}",
            )
        except Exception as exc:
            logger.exception("Unexpected error during Graph API verification")
            return TestMailResponse(
                success=False,
                error=f"Unexpected error: {str(exc)}",
            )
