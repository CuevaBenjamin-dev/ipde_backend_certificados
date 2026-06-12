from __future__ import annotations

from io import BytesIO
from pathlib import PurePosixPath
import posixpath
import re
from xml.etree import ElementTree as ET
import zipfile


CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

SLIDE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
)
SLIDE_MASTER_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)

ET.register_namespace("p", PRESENTATION_NS)
ET.register_namespace("r", OFFICE_REL_NS)
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")


def _xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def _read_zip(pptx_bytes: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(BytesIO(pptx_bytes), "r") as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError("El PPTX de entrada contiene partes ZIP duplicadas")
        return {name: archive.read(name) for name in names}


def _write_zip(files: dict[str, bytes]) -> bytes:
    output = BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)
    return output.getvalue()


def _rels_path_for_part(part_path: str) -> str:
    folder, filename = posixpath.split(part_path)
    return posixpath.join(folder, "_rels", f"{filename}.rels")


def _part_path_for_rels(rels_path: str) -> str:
    if rels_path == "_rels/.rels":
        return ""

    folder, filename = rels_path.split("/_rels/", 1)
    return posixpath.join(folder, filename[:-5])


def _normalize_internal_target(base_part_path: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))

    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base_part_path), target)
    )


def _relative_target(from_part_path: str, to_part_path: str) -> str:
    return posixpath.relpath(to_part_path, start=posixpath.dirname(from_part_path))


def _max_number_for_pattern(names: set[str], pattern: str) -> int:
    regex = re.compile(pattern)
    values = [
        int(match.group(1))
        for name in names
        if (match := regex.match(name)) is not None
    ]
    return max(values, default=0)


def _init_part_counters(existing_names: set[str]) -> dict[str, int]:
    return {
        "slide": _max_number_for_pattern(
            existing_names, r"^ppt/slides/slide(\d+)\.xml$"
        ),
        "layout": _max_number_for_pattern(
            existing_names, r"^ppt/slideLayouts/slideLayout(\d+)\.xml$"
        ),
        "master": _max_number_for_pattern(
            existing_names, r"^ppt/slideMasters/slideMaster(\d+)\.xml$"
        ),
        "theme": _max_number_for_pattern(
            existing_names, r"^ppt/theme/theme(\d+)\.xml$"
        ),
        "chart": _max_number_for_pattern(
            existing_names, r"^ppt/charts/chart(\d+)\.xml$"
        ),
        "media": _max_number_for_pattern(
            existing_names, r"^ppt/media/[A-Za-z_]*(\d+)\.[A-Za-z0-9]+$"
        ),
        "embedding": _max_number_for_pattern(
            existing_names, r"^ppt/embeddings/embedding(\d+)\.[A-Za-z0-9]+$"
        ),
        "ole": _max_number_for_pattern(
            existing_names, r"^ppt/embeddings/oleObject(\d+)\.bin$"
        ),
    }


def _next_named_part(
    folder: str,
    prefix: str,
    extension: str,
    counter_key: str,
    existing_names: set[str],
    counters: dict[str, int],
) -> str:
    while True:
        counters[counter_key] += 1
        candidate = f"{folder}/{prefix}{counters[counter_key]}{extension}"
        if candidate not in existing_names:
            existing_names.add(candidate)
            return candidate


def _allocate_new_part_name(
    source_part_path: str,
    existing_names: set[str],
    counters: dict[str, int],
) -> str:
    folder, filename = posixpath.split(source_part_path)
    base, extension = posixpath.splitext(filename)

    known_parts = {
        "ppt/slides": ("slide", ".xml", "slide"),
        "ppt/slideLayouts": ("slideLayout", ".xml", "layout"),
        "ppt/slideMasters": ("slideMaster", ".xml", "master"),
        "ppt/theme": ("theme", ".xml", "theme"),
        "ppt/charts": ("chart", ".xml", "chart"),
    }
    if folder in known_parts:
        prefix, fixed_extension, counter_key = known_parts[folder]
        return _next_named_part(
            folder,
            prefix,
            fixed_extension,
            counter_key,
            existing_names,
            counters,
        )

    if folder == "ppt/media":
        return _next_named_part(
            folder,
            "image",
            extension,
            "media",
            existing_names,
            counters,
        )

    if folder == "ppt/embeddings":
        prefix = "oleObject" if extension.lower() == ".bin" else "embedding"
        counter_key = "ole" if extension.lower() == ".bin" else "embedding"
        return _next_named_part(
            folder,
            prefix,
            extension,
            counter_key,
            existing_names,
            counters,
        )

    if source_part_path not in existing_names:
        existing_names.add(source_part_path)
        return source_part_path

    copy_index = 1
    while True:
        candidate = posixpath.join(folder, f"{base}_copy{copy_index}{extension}")
        if candidate not in existing_names:
            existing_names.add(candidate)
            return candidate
        copy_index += 1


def _content_type_maps(content_types_xml: bytes):
    root = ET.fromstring(content_types_xml)
    overrides: dict[str, str] = {}
    defaults: dict[str, str] = {}

    for child in root:
        if child.tag.endswith("Override"):
            part_name = child.attrib.get("PartName", "").lstrip("/")
            content_type = child.attrib.get("ContentType", "")
            if part_name and content_type:
                overrides[part_name] = content_type
        elif child.tag.endswith("Default"):
            extension = child.attrib.get("Extension", "").lower()
            content_type = child.attrib.get("ContentType", "")
            if extension and content_type:
                defaults[extension] = content_type

    return root, overrides, defaults


def _ensure_content_type(
    destination_root: ET.Element,
    destination_overrides: dict[str, str],
    destination_defaults: dict[str, str],
    source_overrides: dict[str, str],
    source_defaults: dict[str, str],
    source_part_path: str,
    destination_part_path: str,
) -> None:
    content_type = source_overrides.get(source_part_path)
    if content_type:
        if destination_part_path not in destination_overrides:
            ET.SubElement(
                destination_root,
                f"{{{CONTENT_TYPES_NS}}}Override",
                {
                    "PartName": f"/{destination_part_path}",
                    "ContentType": content_type,
                },
            )
            destination_overrides[destination_part_path] = content_type
        return

    extension = PurePosixPath(source_part_path).suffix.lstrip(".").lower()
    content_type = source_defaults.get(extension)
    if extension and content_type and extension not in destination_defaults:
        ET.SubElement(
            destination_root,
            f"{{{CONTENT_TYPES_NS}}}Default",
            {"Extension": extension, "ContentType": content_type},
        )
        destination_defaults[extension] = content_type


def _next_relationship_id(relationships_root: ET.Element) -> str:
    values = []
    for relationship in relationships_root:
        match = re.fullmatch(r"rId(\d+)", relationship.attrib.get("Id", ""))
        if match:
            values.append(int(match.group(1)))
    return f"rId{max(values, default=0) + 1}"


def _next_slide_id(presentation_root: ET.Element) -> str:
    values = []
    for element in presentation_root.iter():
        if element.tag.endswith("sldId"):
            try:
                values.append(int(element.attrib["id"]))
            except (KeyError, ValueError):
                continue
    return str(max(values, default=255) + 1)


def _find_or_create_id_list(
    presentation_root: ET.Element,
    list_name: str,
) -> ET.Element:
    element = presentation_root.find(f"{{{PRESENTATION_NS}}}{list_name}")
    if element is not None:
        return element

    element = ET.Element(f"{{{PRESENTATION_NS}}}{list_name}")
    presentation_root.insert(0, element)
    return element


def _source_slide_parts(source_files: dict[str, bytes]) -> list[str]:
    presentation_root = ET.fromstring(source_files["ppt/presentation.xml"])
    relationships_root = ET.fromstring(
        source_files["ppt/_rels/presentation.xml.rels"]
    )
    relationships = {
        relationship.attrib.get("Id"): relationship
        for relationship in relationships_root
    }

    slide_parts = []
    slide_id_list = presentation_root.find(
        f"{{{PRESENTATION_NS}}}sldIdLst"
    )
    if slide_id_list is None:
        return slide_parts

    for slide_id in slide_id_list.findall(f"{{{PRESENTATION_NS}}}sldId"):
        relationship_id = slide_id.attrib.get(f"{{{OFFICE_REL_NS}}}id")
        relationship = relationships.get(relationship_id)
        if (
            relationship is None
            or relationship.attrib.get("Type") != SLIDE_REL_TYPE
        ):
            continue

        target = relationship.attrib.get("Target", "")
        part_path = _normalize_internal_target("ppt/presentation.xml", target)
        if part_path in source_files:
            slide_parts.append(part_path)

    return slide_parts


def _normalize_master_and_layout_ids(files: dict[str, bytes]) -> None:
    """Make master and layout IDs unique across the whole presentation."""
    presentation_root = ET.fromstring(files["ppt/presentation.xml"])
    relationships_root = ET.fromstring(
        files["ppt/_rels/presentation.xml.rels"]
    )
    relationship_targets = {
        relationship.attrib.get("Id"): relationship.attrib.get("Target", "")
        for relationship in relationships_root
        if relationship.attrib.get("Type") == SLIDE_MASTER_REL_TYPE
    }

    used_ids: set[int] = set()
    next_id = 2_147_483_648
    presentation_changed = False

    for master_id in presentation_root.findall(
        f".//{{{PRESENTATION_NS}}}sldMasterId"
    ):
        try:
            value = int(master_id.attrib["id"])
        except (KeyError, ValueError) as error:
            raise ValueError("El PPTX contiene un ID de master inválido") from error

        if value < 2_147_483_648 or value in used_ids:
            while next_id in used_ids:
                next_id += 1
            value = next_id
            master_id.attrib["id"] = str(value)
            presentation_changed = True

        used_ids.add(value)
        next_id = max(next_id, value + 1)

        relationship_id = master_id.attrib.get(f"{{{OFFICE_REL_NS}}}id")
        target = relationship_targets.get(relationship_id)
        if not target:
            raise ValueError("No se encontró la relación de un master del PPTX")

        master_path = _normalize_internal_target("ppt/presentation.xml", target)
        master_root = ET.fromstring(files[master_path])
        master_changed = False

        for layout_id in master_root.findall(
            f".//{{{PRESENTATION_NS}}}sldLayoutId"
        ):
            try:
                layout_value = int(layout_id.attrib["id"])
            except (KeyError, ValueError) as error:
                raise ValueError("El PPTX contiene un ID de layout inválido") from error

            if layout_value in used_ids:
                while next_id in used_ids:
                    next_id += 1
                layout_value = next_id
                layout_id.attrib["id"] = str(layout_value)
                master_changed = True

            used_ids.add(layout_value)
            next_id = max(next_id, layout_value + 1)

        if master_changed:
            files[master_path] = _xml_bytes(master_root)

    if presentation_changed:
        files["ppt/presentation.xml"] = _xml_bytes(presentation_root)


def _validate_package(files: dict[str, bytes]) -> None:
    for rels_path, rels_data in files.items():
        if not rels_path.endswith(".rels"):
            continue

        base_part_path = _part_path_for_rels(rels_path)
        relationships_root = ET.fromstring(rels_data)
        for relationship in relationships_root:
            if relationship.attrib.get("TargetMode", "").lower() == "external":
                continue

            target = relationship.attrib.get("Target", "")
            if not target:
                continue

            target_path = _normalize_internal_target(base_part_path, target)
            if target_path not in files:
                raise ValueError(
                    f"Relación interna rota en {rels_path}: {target_path}"
                )

    content_types_root = ET.fromstring(files["[Content_Types].xml"])
    for element in content_types_root:
        if not element.tag.endswith("Override"):
            continue
        part_name = element.attrib.get("PartName", "").lstrip("/")
        if part_name and part_name not in files:
            raise ValueError(
                f"[Content_Types].xml referencia una parte inexistente: {part_name}"
            )


def merge_pptx_packages(pptx_files: list[bytes]) -> bytes:
    """
    Merge PPTX packages while preserving each slide's master, layout and media.

    This is intentionally package-level. Recreating a slide with python-pptx assigns
    it a layout from the first presentation, which corrupts Universidad 2QRS when
    certificate types with different visual masters are combined.
    """
    if not pptx_files:
        raise ValueError("No hay PPTX para unir")

    destination_files = _read_zip(pptx_files[0])
    existing_names = set(destination_files)
    counters = _init_part_counters(existing_names)

    (
        destination_content_types_root,
        destination_overrides,
        destination_defaults,
    ) = _content_type_maps(destination_files["[Content_Types].xml"])
    destination_presentation_root = ET.fromstring(
        destination_files["ppt/presentation.xml"]
    )
    destination_relationships_root = ET.fromstring(
        destination_files["ppt/_rels/presentation.xml.rels"]
    )
    destination_slide_ids = _find_or_create_id_list(
        destination_presentation_root, "sldIdLst"
    )
    destination_master_ids = _find_or_create_id_list(
        destination_presentation_root, "sldMasterIdLst"
    )

    for pptx_bytes in pptx_files[1:]:
        source_files = _read_zip(pptx_bytes)
        _, source_overrides, source_defaults = _content_type_maps(
            source_files["[Content_Types].xml"]
        )
        part_map: dict[str, str] = {}
        copied_master_parts: set[str] = set()

        def copy_part(source_part_path: str) -> str:
            if source_part_path in part_map:
                return part_map[source_part_path]
            if source_part_path not in source_files:
                raise ValueError(f"Parte faltante en PPTX fuente: {source_part_path}")

            destination_part_path = _allocate_new_part_name(
                source_part_path,
                existing_names,
                counters,
            )
            part_map[source_part_path] = destination_part_path
            destination_files[destination_part_path] = source_files[source_part_path]

            _ensure_content_type(
                destination_content_types_root,
                destination_overrides,
                destination_defaults,
                source_overrides,
                source_defaults,
                source_part_path,
                destination_part_path,
            )

            if destination_part_path.startswith("ppt/slideMasters/"):
                copied_master_parts.add(destination_part_path)

            source_rels_path = _rels_path_for_part(source_part_path)
            if source_rels_path in source_files:
                relationships_root = ET.fromstring(source_files[source_rels_path])
                for relationship in relationships_root:
                    if (
                        relationship.attrib.get("TargetMode", "").lower()
                        == "external"
                    ):
                        continue

                    target = relationship.attrib.get("Target", "")
                    if not target:
                        continue

                    source_target_path = _normalize_internal_target(
                        source_part_path, target
                    )
                    if source_target_path not in source_files:
                        continue

                    destination_target_path = copy_part(source_target_path)
                    relationship.attrib["Target"] = _relative_target(
                        destination_part_path,
                        destination_target_path,
                    )

                destination_rels_path = _rels_path_for_part(destination_part_path)
                destination_files[destination_rels_path] = _xml_bytes(
                    relationships_root
                )

            return destination_part_path

        for source_slide_part in _source_slide_parts(source_files):
            destination_slide_part = copy_part(source_slide_part)
            relationship_id = _next_relationship_id(
                destination_relationships_root
            )
            ET.SubElement(
                destination_relationships_root,
                f"{{{PACKAGE_REL_NS}}}Relationship",
                {
                    "Id": relationship_id,
                    "Type": SLIDE_REL_TYPE,
                    "Target": _relative_target(
                        "ppt/presentation.xml", destination_slide_part
                    ),
                },
            )
            ET.SubElement(
                destination_slide_ids,
                f"{{{PRESENTATION_NS}}}sldId",
                {
                    "id": _next_slide_id(destination_presentation_root),
                    f"{{{OFFICE_REL_NS}}}id": relationship_id,
                },
            )

        for master_part in sorted(copied_master_parts):
            relationship_id = _next_relationship_id(
                destination_relationships_root
            )
            ET.SubElement(
                destination_relationships_root,
                f"{{{PACKAGE_REL_NS}}}Relationship",
                {
                    "Id": relationship_id,
                    "Type": SLIDE_MASTER_REL_TYPE,
                    "Target": _relative_target(
                        "ppt/presentation.xml", master_part
                    ),
                },
            )
            ET.SubElement(
                destination_master_ids,
                f"{{{PRESENTATION_NS}}}sldMasterId",
                {
                    # This provisional value is normalized with all layout IDs below.
                    "id": "0",
                    f"{{{OFFICE_REL_NS}}}id": relationship_id,
                },
            )

    destination_files["ppt/presentation.xml"] = _xml_bytes(
        destination_presentation_root
    )
    destination_files["ppt/_rels/presentation.xml.rels"] = _xml_bytes(
        destination_relationships_root
    )
    destination_files["[Content_Types].xml"] = _xml_bytes(
        destination_content_types_root
    )

    _normalize_master_and_layout_ids(destination_files)
    _validate_package(destination_files)
    return _write_zip(destination_files)
