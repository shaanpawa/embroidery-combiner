"""
Authentication dependency for FastAPI.
Validates NextAuth JWT tokens and checks email whitelist.

In development mode (AUTH_DISABLED=true), auth is bypassed.
"""

import logging
import os
from fastapi import Depends, HTTPException, Request

logger = logging.getLogger(__name__)

# Whitelist of allowed emails (comma-separated in env var)
ALLOWED_EMAILS = set(
    e.strip().lower()
    for e in os.environ.get("ALLOWED_EMAILS", "").split(",")
    if e.strip()
)

# Secure by default: auth is ON unless explicitly disabled
AUTH_DISABLED = os.environ.get("AUTH_DISABLED", "false").lower() == "true"


async def get_current_user(request: Request) -> dict:
    """Extract and validate the current user from the request.

    In production: validates NextAuth JWT from cookie/header and checks whitelist.
    In dev mode (AUTH_DISABLED=true): returns a default local user.
    """
    if AUTH_DISABLED:
        return {"email": "local@dev", "name": "Local Dev"}

    # Try Authorization header first (Bearer token)
    auth_header = request.headers.get("authorization", "")
    token = None
    if auth_header.startswith("Bearer "):
        token = auth_header[7:]

    # Try NextAuth session cookie
    if not token:
        token = request.cookies.get("next-auth.session-token") or \
                request.cookies.get("__Secure-next-auth.session-token")

    if not token:
        logger.warning("Unauthenticated request from %s to %s", request.client.host if request.client else "unknown", request.url.path)
        raise HTTPException(401, "Not authenticated")

    # Decode the NextAuth JWT
    try:
        from jose import jwt as jose_jwt
        secret = os.environ.get("NEXTAUTH_SECRET", "")
        if not secret:
            raise HTTPException(500, "NEXTAUTH_SECRET not configured")
        payload = jose_jwt.decode(token, secret, algorithms=["HS256"])
        email = payload.get("email", "").lower()
        name = payload.get("name", "")
    except ImportError:
        # Fallback: try PyJWT
        try:
            import jwt
            secret = os.environ.get("NEXTAUTH_SECRET", "")
            payload = jwt.decode(token, secret, algorithms=["HS256"])
            email = payload.get("email", "").lower()
            name = payload.get("name", "")
        except Exception:
            logger.warning("Invalid JWT token from %s", request.client.host if request.client else "unknown")
            raise HTTPException(401, "Invalid token")
    except Exception:
        logger.warning("Invalid JWT token from %s", request.client.host if request.client else "unknown")
        raise HTTPException(401, "Invalid token")

    if not email:
        raise HTTPException(401, "No email in token")

    # Check whitelist (if configured)
    if ALLOWED_EMAILS and email not in ALLOWED_EMAILS:
        logger.warning("Email %s not in whitelist, rejected from %s", email, request.url.path)
        raise HTTPException(403, f"Email {email} is not authorized")

    return {"email": email, "name": name}
