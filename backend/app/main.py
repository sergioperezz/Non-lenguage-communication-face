from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .api import routes
from .core.config import get_settings
from .core.database import init_db

settings = get_settings()

app = FastAPI(title=settings.app_name)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)


@app.on_event("startup")
def on_startup() -> None:
    init_db()


app.include_router(routes.router, prefix="/api")


@app.get("/")
def root():
    return {"message": "Servidor activo"}
