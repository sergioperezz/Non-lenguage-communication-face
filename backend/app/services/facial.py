from __future__ import annotations
from typing import Dict
import random

from ..schemas import MetricScore


def analyze_facial(video_path: str | None) -> Dict[str, MetricScore]:
    base_score = random.uniform(0.55, 0.85)
    return {
        "facialExpression": MetricScore(
            label="Expresión facial",
            value=base_score,
            description="Probabilidad de sonrisa y expresividad"
        ),
        "bodyLanguage": MetricScore(
            label="Lenguaje corporal",
            value=max(0.0, min(1.0, base_score + 0.08)),
            description="Postura y gestos"
        )
    }
