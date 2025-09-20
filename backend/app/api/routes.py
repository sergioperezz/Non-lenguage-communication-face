from datetime import datetime
from typing import List, Optional
from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from sqlmodel import Session, select

from ..api.dependencies import get_current_user, get_db_session
from ..models import CoachingSession, User
from ..schemas import (
    AuthTokenResponse,
    HistoryItem,
    SessionMetrics,
    SessionResponse,
    SessionSummary,
    UploadSessionRequest,
    UserProfile
)
from ..services.auth import create_or_get_user, issue_token
from ..services.metrics import synthesize_metrics
from ..services.recommendations import build_summary
from ..services.storage import save_upload

router = APIRouter()


@router.get("/health")
def healthcheck():
    return {"status": "ok", "timestamp": datetime.utcnow().isoformat()}


@router.get("/auth/google/url")
def get_google_auth_url():
    return {"url": "https://accounts.google.com/o/oauth2/v2/auth?mock=true"}


@router.post("/auth/google/callback", response_model=AuthTokenResponse)
def google_callback(
    code: str,
    session: Session = Depends(get_db_session)
):
    if not code:
        raise HTTPException(status_code=400, detail="Código de autorización requerido")
    # Simulación de intercambio de código
    email = f"user-{code}@example.com"
    name = "Usuario Demo"
    user = create_or_get_user(session, email=email, name=name)
    token = issue_token(session, user)
    return AuthTokenResponse(
        token=token,
        user=UserProfile(id=user.id, email=user.email, name=user.name, avatarUrl=user.avatar_url)
    )


@router.get("/accounts/me", response_model=UserProfile)
def get_profile(current_user: User = Depends(get_current_user)):
    return UserProfile(
        id=current_user.id,
        email=current_user.email,
        name=current_user.name,
        avatarUrl=current_user.avatar_url
    )


def _latest_session(session: Session, user_id: str, scenario: Optional[str] = None) -> Optional[CoachingSession]:
    statement = select(CoachingSession).where(CoachingSession.user_id == user_id)
    if scenario:
        statement = statement.where(CoachingSession.scenario == scenario)
    statement = statement.order_by(CoachingSession.created_at.desc())
    return session.exec(statement).first()


@router.get("/sessions/current", response_model=SessionMetrics)
def get_current_metrics(
    scenario: str,
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_db_session)
):
    latest = _latest_session(session, current_user.id, scenario)
    if latest:
        return SessionMetrics.model_validate(latest.metrics)
    metrics = synthesize_metrics(scenario, audio_path=None, video_path=None)
    summary = build_summary(metrics)
    new_session = CoachingSession(
        user_id=current_user.id,
        scenario=scenario,
        metrics=metrics.model_dump(),
        summary=summary.model_dump()
    )
    session.add(new_session)
    session.commit()
    session.refresh(new_session)
    return metrics


@router.get("/sessions/summary", response_model=SessionSummary)
def get_summary(
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_db_session)
):
    latest = _latest_session(session, current_user.id)
    if not latest:
        metrics = synthesize_metrics("interview", audio_path=None, video_path=None)
        return build_summary(metrics)
    return SessionSummary.model_validate(latest.summary)


@router.get("/sessions/history", response_model=List[HistoryItem])
def get_history(
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_db_session)
):
    statement = (
        select(CoachingSession)
        .where(CoachingSession.user_id == current_user.id)
        .order_by(CoachingSession.created_at.desc())
        .limit(10)
    )
    sessions = session.exec(statement).all()
    return [
        HistoryItem(
            sessionId=item.id or 0,
            metrics=SessionMetrics.model_validate(item.metrics),
            summary=SessionSummary.model_validate(item.summary)
        )
        for item in sessions
    ]


@router.post("/sessions/upload", response_model=SessionResponse)
def upload_session(
    scenario: str = Form(...),
    video: UploadFile | None = File(default=None),
    audio: UploadFile | None = File(default=None),
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_db_session)
):
    video_path = save_upload(video, "video") if video else None
    audio_path = save_upload(audio, "audio") if audio else None
    metrics = synthesize_metrics(scenario, str(audio_path) if audio_path else None, str(video_path) if video_path else None)
    summary = build_summary(metrics)

    db_session = CoachingSession(
        user_id=current_user.id,
        scenario=scenario,
        video_path=str(video_path) if video_path else None,
        audio_path=str(audio_path) if audio_path else None,
        metrics=metrics.model_dump(),
        summary=summary.model_dump()
    )
    session.add(db_session)
    session.commit()
    session.refresh(db_session)

    return SessionResponse(sessionId=db_session.id or 0, metrics=metrics, summary=summary)


@router.post("/sessions/register", response_model=SessionResponse)
def register_session(
    payload: UploadSessionRequest,
    current_user: User = Depends(get_current_user),
    session: Session = Depends(get_db_session)
):
    db_session = CoachingSession(
        user_id=current_user.id,
        scenario=payload.scenario,
        video_path=payload.video_path,
        audio_path=payload.audio_path,
        metrics=payload.metrics.model_dump(),
        summary=payload.summary.model_dump()
    )
    session.add(db_session)
    session.commit()
    session.refresh(db_session)
    return SessionResponse(sessionId=db_session.id or 0, metrics=payload.metrics, summary=payload.summary)
