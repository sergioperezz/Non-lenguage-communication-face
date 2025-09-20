from datetime import datetime
from typing import List, Optional
from pydantic import BaseModel, Field

class MetricScore(BaseModel):
    label: str
    value: float = Field(ge=0, le=1)
    trend: Optional[str] = Field(default=None, pattern="^(up|down|steady)$")
    description: Optional[str] = None

class SessionMetrics(BaseModel):
    scenario: str
    voice: MetricScore
    prosody: MetricScore
    pronunciation: MetricScore
    speed: MetricScore
    facialExpression: MetricScore
    bodyLanguage: MetricScore
    fillerWords: MetricScore
    overall: MetricScore
    timestamp: datetime

class Recommendation(BaseModel):
    title: str
    description: str
    suggestedExercises: List[str]

class SessionSummary(BaseModel):
    generalComment: str
    recommendations: List[Recommendation]

class SessionResponse(BaseModel):
    sessionId: int
    metrics: SessionMetrics
    summary: SessionSummary

class AuthTokenResponse(BaseModel):
    token: str
    user: "UserProfile"

class UserProfile(BaseModel):
    id: str
    email: str
    name: str
    avatarUrl: Optional[str] = None

class UploadSessionRequest(BaseModel):
    scenario: str
    video_path: Optional[str] = None
    audio_path: Optional[str] = None
    metrics: SessionMetrics
    summary: SessionSummary

class HistoryItem(BaseModel):
    sessionId: int
    metrics: SessionMetrics
    summary: SessionSummary

AuthTokenResponse.model_rebuild()
