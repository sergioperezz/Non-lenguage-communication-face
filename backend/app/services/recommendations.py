from __future__ import annotations
from typing import List

from ..schemas import Recommendation, SessionMetrics, SessionSummary


FEEDBACK_TEMPLATES = {
    "interview": {
        "comment": "En entrevistas, destaca manteniendo contacto visual estable y respuestas concisas.",
        "exercises": [
            "Simula preguntas de fortalezas y debilidades grabándote",
            "Practica respuestas STAR frente al espejo"
        ]
    },
    "presentation": {
        "comment": "Buen dominio del contenido. Trabaja en modular la entonación y pausar entre ideas.",
        "exercises": [
            "Ensaya con diapositivas marcando pausas",
            "Graba tu pitch y analiza la velocidad"
        ]
    },
    "date": {
        "comment": "Mantén un tono cercano y controla gestos nerviosos para generar confianza.",
        "exercises": [
            "Practica storytelling breve",
            "Ejercicio de respiración diafragmática"
        ]
    }
}


def build_summary(metrics: SessionMetrics) -> SessionSummary:
    template = FEEDBACK_TEMPLATES.get(metrics.scenario, FEEDBACK_TEMPLATES["interview"])
    recommendations: List[Recommendation] = [
        Recommendation(
            title="Prosodia",
            description="Varía tu ritmo y enfatiza palabras clave para sostener la atención.",
            suggestedExercises=["Lectura expresiva", "Reforzar pausas controladas"]
        ),
        Recommendation(
            title="Lenguaje corporal",
            description="Alinea tus gestos con el mensaje y evita movimientos repetitivos.",
            suggestedExercises=["Practicar frente al espejo", "Ejercicios de postura neutra"]
        )
    ]
    if metrics.fillerWords.value < 0.6:
        recommendations.append(
            Recommendation(
                title="Reducción de muletillas",
                description="Usa silencios estratégicos en lugar de \"eh\" o \"mmm\".",
                suggestedExercises=["Responder preguntas con pausas", "Meditación guiada"]
            )
        )

    return SessionSummary(generalComment=template["comment"], recommendations=recommendations)
