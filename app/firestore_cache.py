"""Persistencia de la caché dinámica de módulos en Google Cloud Firestore."""

from __future__ import annotations

import logging
import os
import re
import threading
import unicodedata

from google.api_core.exceptions import AlreadyExists
from google.cloud import firestore
from google.cloud.firestore_v1 import Client


LOGGER = logging.getLogger(__name__)

DEFAULT_COLLECTION = "modulos_cache"
DEFAULT_DATABASE = "(default)"

_client: Client | None = None
_client_lock = threading.Lock()


def normalizar_clave_tema(texto: str) -> str:
    """Devuelve la misma clave estable que utilizaba la caché JSON."""
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = texto.encode("ascii", "ignore").decode("ascii")
    texto = texto.lower().strip()
    texto = re.sub(r"[^a-z0-9]+", "-", texto)
    texto = re.sub(r"-+", "-", texto).strip("-")
    return texto or "sin-tema"


def extraer_modulos_de_registro(registro: object) -> list[str] | None:
    """Extrae módulos de los dos formatos históricos admitidos por el proyecto."""
    if isinstance(registro, list):
        modulos = registro
    elif isinstance(registro, dict):
        modulos = registro.get("modulos")
    else:
        return None

    if not isinstance(modulos, list):
        return None

    return [str(modulo).upper().strip() for modulo in modulos if str(modulo).strip()]


def _get_client() -> Client:
    """Crea el cliente bajo demanda para no bloquear el arranque de FastAPI."""
    global _client

    if _client is not None:
        return _client

    with _client_lock:
        if _client is None:
            project_id = os.getenv("FIRESTORE_PROJECT_ID") or None
            database_id = os.getenv("FIRESTORE_DATABASE_ID", DEFAULT_DATABASE).strip()
            database_id = database_id or DEFAULT_DATABASE
            _client = firestore.Client(project=project_id, database=database_id)

    return _client


def _collection_name() -> str:
    return os.getenv("FIRESTORE_COLLECTION", DEFAULT_COLLECTION).strip() or DEFAULT_COLLECTION


def buscar_modulos_en_firestore(
    cache_key: str,
    cantidad_esperada: int,
) -> list[str] | None:
    """
    Busca una entrada persistente. Los errores se registran y se tratan como miss;
    una indisponibilidad temporal de Firestore no impide iniciar ni usar la API.
    """
    try:
        snapshot = _get_client().collection(_collection_name()).document(cache_key).get()
        if not snapshot.exists:
            return None

        modulos = extraer_modulos_de_registro(snapshot.to_dict())
        if not modulos or len(modulos) != cantidad_esperada:
            LOGGER.warning(
                "Registro inválido en Firestore para cache_key=%s; se esperaban %s módulos",
                cache_key,
                cantidad_esperada,
            )
            return None

        return modulos
    except Exception:
        LOGGER.exception("Error buscando módulos en Firestore para cache_key=%s", cache_key)
        return None


def guardar_modulos_en_firestore(
    cache_key: str,
    grupo: str,
    tipo: str,
    tema: str,
    modulos: list[str],
) -> bool:
    """
    Guarda un resultado válido de OpenAI y conserva created_at si ya existía.

    Devuelve False ante un error, sin descartar los módulos ya generados ni hacer
    fallar la petición actual.
    """
    try:
        document = _get_client().collection(_collection_name()).document(cache_key)
        payload = {
            "grupo": grupo,
            "tipo_referencia": (tipo or "").upper().strip(),
            "tema_original": (tema or "").strip(),
            "modulos": [str(modulo).upper().strip() for modulo in modulos],
            "updated_at": firestore.SERVER_TIMESTAMP,
        }

        try:
            document.create({**payload, "created_at": firestore.SERVER_TIMESTAMP})
        except AlreadyExists:
            document.set(payload, merge=True)

        LOGGER.info("Módulos guardados en Firestore: cache_key=%s", cache_key)
        return True
    except Exception:
        LOGGER.exception("Error guardando módulos en Firestore para cache_key=%s", cache_key)
        return False
