"""Un registro por PPTX en Firestore; Sheets es una copia recuperable del detalle."""

from __future__ import annotations

import logging
import os
import re
import threading
import uuid
from datetime import datetime, timedelta, timezone

import google.auth
import google_auth_httplib2
import httplib2
from google.cloud import firestore
from google.cloud.firestore_v1.base_query import FieldFilter
from googleapiclient.discovery import build


LOGGER = logging.getLogger(__name__)
COLLECTION = "diplomas_generados"
SHEET_NAME = "Registros"
SHEETS_SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]
# Perú usa UTC-05:00 sin horario de verano; funciona sin tzdata en Windows.
PERU_TZ = timezone(timedelta(hours=-5))
FIRESTORE_TIMEOUT = 10
SHEETS_TIMEOUT = 20
HEADERS = [
    "generation_id", "fecha_registro", "estado", "modelo_certificado",
    "tipo_modelo", "nombres", "apellidos", "tema", "fecha_inicio", "fecha_fin",
    "horas", "creditos", "folio", "fecha_emision", "codigo_estudiante", "ciudad",
    "confirmado_frontend_at",
]
_firestore_client = None
_client_lock = threading.Lock()


class RegistroNoDisponible(RuntimeError):
    """No se pudo persistir el evento principal."""


def _sheet_id() -> str:
    # Se lee después de load_dotenv(), no al importar este módulo.
    return os.getenv("GOOGLE_SHEET_ID", "").strip()


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _local_iso(value: datetime) -> str:
    return value.astimezone(PERU_TZ).isoformat(timespec="seconds")


def _get_firestore():
    global _firestore_client
    with _client_lock:
        if _firestore_client is None:
            _firestore_client = firestore.Client(
                project=os.getenv("FIRESTORE_PROJECT_ID") or None,
                database=os.getenv("FIRESTORE_DATABASE_ID", "(default)").strip() or "(default)",
            )
    return _firestore_client


def _get_sheets():
    # httplib2 no es thread-safe: cada sincronización tiene su propio transporte.
    credentials, _ = google.auth.default(scopes=SHEETS_SCOPE)
    http = google_auth_httplib2.AuthorizedHttp(
        credentials, http=httplib2.Http(timeout=SHEETS_TIMEOUT)
    )
    return build("sheets", "v4", http=http, cache_discovery=False)


def _document(generation_id: str):
    return _get_firestore().collection(COLLECTION).document(generation_id)


def registrar_generacion(items, *, filename: str = "") -> str:
    """Persistir una generación completa ANTES de comenzar la respuesta PPTX."""
    generation_id = uuid.uuid4().hex
    fecha = _local_iso(_now())
    payload = {
        "generation_id": generation_id,
        "estado": "GENERADO",
        "created_at": firestore.SERVER_TIMESTAMP,
        "fecha_registro": fecha,
        "fecha_peru": fecha[:10],
        "hora_peru": fecha[11:19],
        "filename": filename,
        "cantidad_certificados": len(items),
        "frontend_confirmed": False,
        "sheet_status": "PENDIENTE" if _sheet_id() else "NO_CONFIGURADO",
        "items": [item.model_dump() for item in items],
    }
    try:
        _document(generation_id).create(payload, retry=None, timeout=FIRESTORE_TIMEOUT)
    except Exception as exc:
        LOGGER.error("No se pudo registrar %s (%s)", generation_id, type(exc).__name__)
        raise RegistroNoDisponible("No se pudo guardar el registro de generación") from exc
    return generation_id


def _confirm_transaction(transaction, document):
    snapshot = document.get(transaction=transaction, retry=None, timeout=FIRESTORE_TIMEOUT)
    if not snapshot.exists:
        return False
    data = snapshot.to_dict()
    if not data.get("frontend_confirmed"):
        transaction.update(document, {
            "estado": "CONFIRMADO_FRONTEND",
            "frontend_confirmed": True,
            "frontend_confirmed_at": firestore.SERVER_TIMESTAMP,
            "confirmado_frontend_at": _local_iso(_now()),
            "sheet_status": "PENDIENTE" if _sheet_id() else "NO_CONFIGURADO",
        })
    return True


def confirmar_generacion(generation_id: str) -> bool:
    """Conserva la primera confirmación incluso con reintentos concurrentes."""
    if not re.fullmatch(r"[0-9a-f]{32}", generation_id):
        return False
    try:
        return firestore.transactional(_confirm_transaction)(
            _get_firestore().transaction(max_attempts=3), _document(generation_id)
        )
    except Exception as exc:
        LOGGER.error("No se pudo confirmar %s (%s)", generation_id, type(exc).__name__)
        raise RegistroNoDisponible("No se pudo guardar la confirmación") from exc


def _claim_transaction(transaction, document, token):
    snapshot = document.get(transaction=transaction, retry=None, timeout=FIRESTORE_TIMEOUT)
    if not snapshot.exists:
        return None
    data = snapshot.to_dict()
    now = _now()
    # Evita dos append simultáneos para el mismo PPTX, incluso desde el script.
    if data.get("sheet_lease_until") and data["sheet_lease_until"] > now:
        return None
    transaction.update(document, {
        "sheet_lease_token": token,
        "sheet_lease_until": now + timedelta(minutes=5),
        "sheet_status": "SINCRONIZANDO",
    })
    return data


def _finish_transaction(transaction, document, token, synced_state, error):
    snapshot = document.get(transaction=transaction, retry=None, timeout=FIRESTORE_TIMEOUT)
    data = snapshot.to_dict() or {}
    if data.get("sheet_lease_token") != token:
        return False
    changed = data.get("estado") != synced_state
    transaction.update(document, {
        "sheet_status": "ERROR" if error else ("PENDIENTE" if changed else "OK"),
        "sheet_error": error or "",
        "sheet_lease_token": firestore.DELETE_FIELD,
        "sheet_lease_until": firestore.DELETE_FIELD,
        "sheet_attempted_at": firestore.SERVER_TIMESTAMP,
    })
    return changed and not error


def _rows(data):
    fecha = data.get("fecha_registro") or _local_iso(data["created_at"])
    confirmed = data.get("confirmado_frontend_at", "")
    if not confirmed and data.get("frontend_confirmed_at"):
        confirmed = _local_iso(data["frontend_confirmed_at"])
    return [[
        data["generation_id"], fecha, data["estado"],
        item["modeloCertificado"], item["tipoModelo"], item["nombres"],
        item["apellidos"], item["temaDiplomado"], item["fechaInicio"], item["fechaFin"],
        item["horasAcademicas"], item["creditosAcademicos"], item["folioNumero"],
        item["fechaEmision"], item.get("codigoEstudiante", ""), item.get("ciudad", ""),
        confirmed,
    ] for item in data["items"]]


def _write_sheet(service, sheet_id: str, data):
    values_api = service.spreadsheets().values()
    header = values_api.get(
        spreadsheetId=sheet_id, range=f"'{SHEET_NAME}'!A1:Q1"
    ).execute(num_retries=0).get("values", [])
    if header != [HEADERS]:
        raise ValueError("Encabezados de Registros incorrectos; consultar docs/GENERATION_LOG.md")
    # Tras un timeout Google pudo aceptar el append. Buscar antes de repetirlo.
    existing = values_api.get(
        spreadsheetId=sheet_id, range=f"'{SHEET_NAME}'!A2:A"
    ).execute(num_retries=0).get("values", [])
    row_numbers = [
        i + 2 for i, row in enumerate(existing)
        if row and row[0] == data["generation_id"]
    ]
    rows = _rows(data)
    if row_numbers:
        if len(row_numbers) != len(rows):
            raise ValueError("Cantidad de filas del lote inconsistente en Registros")
        # Solo cambia estado/confirmación: conserva participante-tema al ordenar.
        updates = []
        for number in row_numbers:
            updates.extend([
                {"range": f"'{SHEET_NAME}'!C{number}", "values": [[data["estado"]]]},
                {"range": f"'{SHEET_NAME}'!Q{number}", "values": [[rows[0][16]]]},
            ])
        values_api.batchUpdate(
            spreadsheetId=sheet_id,
            body={"valueInputOption": "RAW", "data": updates},
        ).execute(num_retries=0)
    else:
        values_api.append(
            spreadsheetId=sheet_id, range=f"'{SHEET_NAME}'!A:Q",
            valueInputOption="RAW", insertDataOption="INSERT_ROWS",
            body={"values": rows},
        ).execute(num_retries=0)


def sincronizar_generacion(generation_id: str) -> bool:
    """Tarea posterior a la respuesta. Los errores quedan pendientes en Firestore."""
    sheet_id = _sheet_id()
    if not sheet_id:
        return False
    try:
        client = _get_firestore()
        document = _document(generation_id)
        # Una confirmación puede llegar durante el primer envío a Sheets.
        for _ in range(2):
            token = uuid.uuid4().hex
            data = firestore.transactional(_claim_transaction)(
                client.transaction(max_attempts=3), document, token
            )
            if data is None:
                return False
            error = ""
            try:
                with _get_sheets() as service:
                    _write_sheet(service, sheet_id, data)
            except Exception as exc:
                # No guardar respuestas HTTP que puedan contener datos personales.
                error = type(exc).__name__
                status = getattr(getattr(exc, "resp", None), "status", None)
                if status:
                    error += f" HTTP {status}"
                LOGGER.error("Sheets pendiente para %s: %s", generation_id, error)
            changed = firestore.transactional(_finish_transaction)(
                client.transaction(max_attempts=3), document, token, data["estado"], error
            )
            if error or not changed:
                return not error
        return False
    except Exception as exc:
        LOGGER.error("Sincronización pendiente para %s (%s)", generation_id, type(exc).__name__)
        return False


def sincronizar_pendientes(*, limit: int = 10) -> dict:
    """Recuperación explícita; funciona también tras reiniciar Render."""
    if not _sheet_id():
        raise ValueError("Falta GOOGLE_SHEET_ID")
    query = _get_firestore().collection(COLLECTION).where(filter=FieldFilter(
        "sheet_status", "in", ["PENDIENTE", "ERROR", "NO_CONFIGURADO", "SINCRONIZANDO"]
    )).limit(limit)
    ids = [snapshot.id for snapshot in query.stream(retry=None, timeout=FIRESTORE_TIMEOUT)]
    synced = sum(sincronizar_generacion(generation_id) for generation_id in ids)
    return {"revisados": len(ids), "sincronizados": synced, "pendientes": len(ids) - synced}
