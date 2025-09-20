from datetime import datetime
from typing import Optional
from sqlmodel import Field, Relationship, SQLModel, Column, JSON

class User(SQLModel, table=True):
    id: str = Field(primary_key=True)
    email: str
    name: str
    avatar_url: Optional[str] = None
    google_sub: Optional[str] = Field(default=None, index=True)
    created_at: datetime = Field(default_factory=datetime.utcnow)
    sessions: list["CoachingSession"] = Relationship(back_populates="user")


class CoachingSession(SQLModel, table=True):
    id: Optional[int] = Field(default=None, primary_key=True)
    user_id: str = Field(foreign_key="user.id")
    scenario: str
    metrics: dict = Field(sa_column=Column(JSON))
    summary: dict = Field(sa_column=Column(JSON))
    video_path: Optional[str] = None
    audio_path: Optional[str] = None
    created_at: datetime = Field(default_factory=datetime.utcnow, index=True)
    user: Optional[User] = Relationship(back_populates="sessions")


class AccessToken(SQLModel, table=True):
    token: str = Field(primary_key=True)
    user_id: str = Field(foreign_key="user.id")
    created_at: datetime = Field(default_factory=datetime.utcnow, index=True)
