from fastapi.testclient import TestClient
from sqlmodel import Session

from ..main import app
from ..core.database import engine, init_db
from ..models import AccessToken, User

client = TestClient(app)


def setup_module(module):
    init_db()


def create_user_with_token():
    with Session(engine) as session:
        user = session.get(User, "test-user")
        if not user:
            user = User(id="test-user", email="test@example.com", name="Test User")
            session.add(user)
            session.commit()
        token = session.get(AccessToken, "test-token")
        if not token:
            token = AccessToken(token="test-token", user_id=user.id)
            session.add(token)
            session.commit()
    return "test-token"


def test_healthcheck():
    response = client.get("/api/health")
    assert response.status_code == 200
    assert response.json()["status"] == "ok"


def test_current_metrics_creates_session():
    token = create_user_with_token()
    response = client.get(
        "/api/sessions/current",
        params={"scenario": "interview"},
        headers={"Authorization": f"Bearer {token}"}
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["scenario"] == "interview"
    assert "voice" in payload


def test_summary_without_sessions_returns_defaults():
    token = create_user_with_token()
    response = client.get(
        "/api/sessions/summary",
        headers={"Authorization": f"Bearer {token}"}
    )
    assert response.status_code == 200
    data = response.json()
    assert "generalComment" in data
