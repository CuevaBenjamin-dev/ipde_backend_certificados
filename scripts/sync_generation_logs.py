"""Validación de acceso (solo lectura) y recuperación explícita de Sheets."""

import argparse
import json
import sys
from pathlib import Path

from dotenv import load_dotenv


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
load_dotenv(ROOT / ".env")

from app import generation_log as log


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="Solo verificar acceso y encabezados")
    parser.add_argument("--limit", type=int, default=10, help="Máximo de pendientes a sincronizar")
    args = parser.parse_args()
    if not 1 <= args.limit <= 100:
        parser.error("--limit debe estar entre 1 y 100")
    if not log._sheet_id():
        parser.error("Falta configurar GOOGLE_SHEET_ID")
    try:
        if args.check:
            # Verifica también acceso a Firestore sin crear documentos.
            list(log._get_firestore().collection(log.COLLECTION).limit(1).stream(
                retry=None, timeout=log.FIRESTORE_TIMEOUT
            ))
            with log._get_sheets() as service:
                result = service.spreadsheets().values().get(
                    spreadsheetId=log._sheet_id(), range=f"'{log.SHEET_NAME}'!A1:Q1"
                ).execute(num_retries=0)
            if result.get("values", []) != [log.HEADERS]:
                print("ERROR: Registros!A1:Q1 no coincide con los 17 encabezados de la guía.")
                return 1
            print("OK: Firestore y Sheets accesibles; encabezados correctos. No se escribió ningún registro.")
            return 0
        result = log.sincronizar_pendientes(limit=args.limit)
        print(json.dumps(result, ensure_ascii=False))
        return 1 if result["pendientes"] else 0
    except Exception as exc:
        status = getattr(getattr(exc, "resp", None), "status", None)
        print(f"ERROR: {type(exc).__name__}" + (f" (HTTP {status})" if status else ""))
        print("Revisa credenciales, API habilitada y permisos del Sheet; consulta docs/GENERATION_LOG.md.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
