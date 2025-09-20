from __future__ import annotations
from datetime import datetime
from ..schemas import SessionMetrics
from .prosody import analyze_prosody
from .facial import analyze_facial


def synthesize_metrics(scenario: str, audio_path: str | None, video_path: str | None) -> SessionMetrics:
    prosody_metrics = analyze_prosody(audio_path)
    facial_metrics = analyze_facial(video_path)
    all_metrics = {**prosody_metrics, **facial_metrics}

    voice_score = all_metrics["voice"].value
    body_score = all_metrics["bodyLanguage"].value
    filler_score = all_metrics["fillerWords"].value
    overall_value = max(0.0, min(1.0, (voice_score + body_score + filler_score) / 3))
    all_metrics["overall"] = all_metrics["overall"] if "overall" in all_metrics else all_metrics["voice"].model_copy(update={"label": "Puntaje global"})
    all_metrics["overall"].value = overall_value
    all_metrics["overall"].label = "Puntaje global"

    return SessionMetrics(
        scenario=scenario,
        voice=all_metrics["voice"],
        prosody=all_metrics["prosody"],
        pronunciation=all_metrics["pronunciation"],
        speed=all_metrics["speed"],
        facialExpression=all_metrics["facialExpression"],
        bodyLanguage=all_metrics["bodyLanguage"],
        fillerWords=all_metrics["fillerWords"],
        overall=all_metrics["overall"],
        timestamp=datetime.utcnow()
    )
