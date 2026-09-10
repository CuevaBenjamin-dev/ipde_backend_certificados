"""Importa app/data/modulos_cache.json en Firestore de forma idempotente."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from google.cloud import firestore
from google.oauth2 import service_account


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.firestore_cache import (  # noqa: E402
    DEFAULT_COLLECTION,
    DEFAULT_DATABASE,
    extraer_modulos_de_registro,
    normalizar_clave_tema,
)


DEFAULT_SOURCE = PROJECT_ROOT / "app" / "data" / "modulos_cache.json"
VALID_GROUPS = {"5_MODULOS": 5, "8_MODULOS": 8}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Migra la caché JSON de módulos a Firebase Cloud Firestore.",
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=DEFAULT_SOURCE,
        help=f"Archivo JSON de origen (predeterminado: {DEFAULT_SOURCE})",
    )
    parser.add_argument(
        "--project",
        help="ID del proyecto Firebase. Si se omite, se obtiene de las credenciales.",
    )
    parser.add_argument(
        "--credentials-file",
        type=Path,
        help=(
            "Ruta local al JSON de la cuenta de servicio. Debe estar fuera del "
            "repositorio. Si se omite, se utilizan Application Default Credentials."
        ),
    )
    parser.add_argument(
        "--database",
        default=DEFAULT_DATABASE,
        help=f"ID de la base de Firestore (predeterminado: {DEFAULT_DATABASE})",
    )
    parser.add_argument(
        "--collection",
        default=DEFAULT_COLLECTION,
        help=f"Colección de destino (predeterminado: {DEFAULT_COLLECTION})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Valida y cuenta los registros sin conectarse a Google Cloud.",
    )
    return parser.parse_args()


def leer_items(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(f"No existe el archivo de caché: {path}")

    with path.open("r", encoding="utf-8") as source_file:
        data = json.load(source_file)

    if not isinstance(data, dict):
        raise ValueError("El JSON raíz debe ser un objeto")

    items = data.get("items", data)
    if not isinstance(items, dict):
        raise ValueError("El campo 'items' debe ser un objeto")

    return items


def preparar_documento(cache_key: str, registro: Any) -> tuple[str, dict[str, Any]]:
    modulos = extraer_modulos_de_registro(registro)
    if not modulos:
        raise ValueError("el registro no contiene una lista válida de módulos")

    registro_dict = registro if isinstance(registro, dict) else {}
    raw_group, separator, raw_topic = cache_key.partition("::")
    group = str(registro_dict.get("grupo") or raw_group).upper().strip()

    if not separator and not registro_dict.get("grupo"):
        group = f"{len(modulos)}_MODULOS"

    expected_count = VALID_GROUPS.get(group)
    if expected_count is None:
        raise ValueError(f"grupo no soportado: {group!r}")
    if len(modulos) != expected_count:
        raise ValueError(
            f"el grupo {group} requiere {expected_count} módulos y recibió {len(modulos)}"
        )

    original_topic = str(registro_dict.get("tema_original") or raw_topic).strip()
    normalized_topic = normalizar_clave_tema(original_topic)
    normalized_key = f"{group}::{normalized_topic}"

    document: dict[str, Any] = {
        "grupo": group,
        "tipo_referencia": str(registro_dict.get("tipo_referencia") or "").upper().strip(),
        "tema_original": original_topic,
        "modulos": modulos,
    }

    if registro_dict.get("created_at") is not None:
        document["created_at"] = registro_dict["created_at"]
    if registro_dict.get("updated_at") is not None:
        document["updated_at"] = registro_dict["updated_at"]

    return normalized_key, document


def main() -> int:
    args = parse_args()

    try:
        items = leer_items(args.source.resolve())
    except Exception as exc:
        print(f"ERROR: no se pudo leer la caché: {exc}", file=sys.stderr)
        return 1

    prepared: list[tuple[str, dict[str, Any]]] = []
    validation_errors = 0

    for cache_key, registro in items.items():
        try:
            prepared.append(preparar_documento(str(cache_key), registro))
        except Exception as exc:
            validation_errors += 1
            print(f"ERROR de validación [{cache_key}]: {exc}", file=sys.stderr)

    print(f"Registros encontrados: {len(items)}")
    print(f"Registros válidos: {len(prepared)}")
    print(f"Errores de validación: {validation_errors}")

    if args.dry_run:
        print("DRY RUN finalizado: no se realizó ninguna conexión ni escritura en Firestore.")
        return 1 if validation_errors else 0

    try:
        credentials = None
        project_id = args.project or None

        if args.credentials_file:
            credentials_path = args.credentials_file.expanduser().resolve()
            if not credentials_path.is_file():
                raise FileNotFoundError(
                    f"No existe el archivo de credenciales: {credentials_path}"
                )

            credentials = service_account.Credentials.from_service_account_file(
                str(credentials_path)
            )
            project_id = project_id or credentials.project_id

        client = firestore.Client(
            project=project_id,
            credentials=credentials,
            database=args.database,
        )
    except Exception as exc:
        print(f"ERROR: no se pudo crear el cliente de Firestore: {exc}", file=sys.stderr)
        return 1

    migrated = 0
    write_errors = 0
    collection = client.collection(args.collection)

    for document_id, document in prepared:
        try:
            # set(..., merge=True) permite repetir la migración sin duplicar documentos.
            collection.document(document_id).set(document, merge=True)
            migrated += 1
        except Exception as exc:
            write_errors += 1
            print(f"ERROR escribiendo [{document_id}]: {exc}", file=sys.stderr)

    print(f"Registros migrados/actualizados: {migrated}")
    print(f"Errores de escritura: {write_errors}")
    return 1 if validation_errors or write_errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
