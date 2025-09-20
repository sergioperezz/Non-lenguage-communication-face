from __future__ import annotations
from datetime import datetime, timedelta
from secrets import token_urlsafe
from sqlmodel import Session, select
from ..models import AccessToken, User

TOKEN_TTL_HOURS = 12


def create_or_get_user(session: Session, email: str, name: str, avatar_url: str | None = None) -> User:
    statement = select(User).where(User.email == email)
    user = session.exec(statement).first()
    if user:
        if avatar_url and user.avatar_url != avatar_url:
            user.avatar_url = avatar_url
            session.add(user)
            session.commit()
            session.refresh(user)
        return user
    user = User(id=token_urlsafe(8), email=email, name=name, avatar_url=avatar_url)
    session.add(user)
    session.commit()
    session.refresh(user)
    return user


def issue_token(session: Session, user: User) -> str:
    token = token_urlsafe(32)
    session.add(AccessToken(token=token, user_id=user.id))
    session.commit()
    return token


def validate_token(session: Session, token: str) -> User | None:
    statement = select(AccessToken, User).join(User, AccessToken.user_id == User.id).where(
        AccessToken.token == token
    )
    result = session.exec(statement).first()
    if not result:
        return None
    access_token, user = result
    if datetime.utcnow() - access_token.created_at > timedelta(hours=TOKEN_TTL_HOURS):
        session.delete(access_token)
        session.commit()
        return None
    return user
