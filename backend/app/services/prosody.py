from __future__ import annotations
from typing import Dict
import random

from ..schemas import MetricScore


def analyze_prosody(audio_path: str | None) -> Dict[str, MetricScore]:
    base_score = random.uniform(0.6, 0.9)
    return {
        "voice": MetricScore(label="Claridad de voz", value=base_score, trend="steady"),
        "prosody": MetricScore(
            label="Entonación",
            value=max(0.0, min(1.0, base_score - 0.05)),
            description="Variación tonal estimada"
        ),
        "pronunciation": MetricScore(
            label="Pronunciación",
            value=max(0.0, min(1.0, base_score + 0.05))
        ),
        "speed": MetricScore(
            label="Velocidad",
            value=max(0.0, min(1.0, base_score - 0.1)),
            description="Palabras por minuto normalizadas"
        ),
        "fillerWords": MetricScore(
            label="Muletillas",
            value=max(0.0, min(1.0, base_score - 0.15)),
            description="Uso de muletillas detectado"
        )
    }
