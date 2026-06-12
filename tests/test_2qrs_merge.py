import io
import json
import os
from pathlib import Path
import posixpath
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile


os.environ.setdefault("OPENAI_API_KEY", "test-key")

from app import main


ROOT = Path(__file__).resolve().parents[1]
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _topic_for_module_count(module_count: int) -> str:
    data = json.loads(
        (ROOT / "app" / "modulos_base.json").read_text(encoding="utf-8")
    )
    prefix = f"{module_count}_MODULOS::"
    return next(
        value["tema_original"]
        for key, value in data["items"].items()
        if key.startswith(prefix)
    )


def _request(certificate_type: str, participant_number: int):
    module_count = main.MODULOS_COUNT[certificate_type]
    return main.DiplomaRequest(
        modeloCertificado="UNIVERSIDAD_2QRS",
        tipoModelo=certificate_type,
        nombres=f"PERSONA{participant_number}",
        apellidos="PRUEBA",
        temaDiplomado=_topic_for_module_count(module_count),
        fechaInicio="2025-01-01",
        fechaFin="2025-12-01",
        horasAcademicas=1200 if module_count == 8 else 240,
        creditosAcademicos=40 if module_count == 8 else 15,
        folioNumero=f"FOLIO-{participant_number}",
        fechaEmision="2025-12-19",
        codigoEstudiante=f"codigo-{participant_number}",
        ciudad="Trujillo",
    )


def _normalize_target(base_part_path: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base_part_path), target)
    )


class Universidad2QRSMergeTests(unittest.TestCase):
    def _assert_valid_package(self, pptx_bytes: bytes, expected_slides: int):
        with ZipFile(io.BytesIO(pptx_bytes)) as archive:
            names = archive.namelist()
            self.assertEqual(len(names), len(set(names)))
            self.assertIsNone(archive.testzip())
            name_set = set(names)

            for rels_path in names:
                if not rels_path.endswith(".rels"):
                    continue

                if rels_path == "_rels/.rels":
                    base_part_path = ""
                else:
                    folder, filename = rels_path.split("/_rels/", 1)
                    base_part_path = posixpath.join(folder, filename[:-5])

                relationships = ET.fromstring(archive.read(rels_path))
                for relationship in relationships:
                    if (
                        relationship.attrib.get("TargetMode", "").lower()
                        == "external"
                    ):
                        continue
                    target = relationship.attrib.get("Target", "")
                    if target:
                        self.assertIn(
                            _normalize_target(base_part_path, target),
                            name_set,
                        )

            presentation = ET.fromstring(
                archive.read("ppt/presentation.xml")
            )
            presentation_relationships = ET.fromstring(
                archive.read("ppt/_rels/presentation.xml.rels")
            )
            relationship_targets = {
                relationship.attrib["Id"]: relationship.attrib["Target"]
                for relationship in presentation_relationships
            }

            global_ids = []
            master_ids = presentation.findall(
                f".//{{{PRESENTATION_NS}}}sldMasterId"
            )
            for master_id in master_ids:
                global_ids.append(int(master_id.attrib["id"]))
                relationship_id = master_id.attrib[
                    f"{{{OFFICE_REL_NS}}}id"
                ]
                master_path = _normalize_target(
                    "ppt/presentation.xml",
                    relationship_targets[relationship_id],
                )
                master = ET.fromstring(archive.read(master_path))
                global_ids.extend(
                    int(layout_id.attrib["id"])
                    for layout_id in master.findall(
                        f".//{{{PRESENTATION_NS}}}sldLayoutId"
                    )
                )

            self.assertEqual(len(global_ids), len(set(global_ids)))

        presentation = main.Presentation(io.BytesIO(pptx_bytes))
        self.assertEqual(len(presentation.slides), expected_slides)

    def test_mixed_certificate_types_preserve_participant_and_structure(self):
        items = [
            _request("DIPLOMADO", 1),
            _request("CURSO", 2),
        ]
        pptx_bytes = main.generar_pptx_universidad_2qrs(items)

        self._assert_valid_package(pptx_bytes, expected_slides=4)

        presentation = main.Presentation(io.BytesIO(pptx_bytes))
        slide_texts = [
            "\n".join(
                shape.text
                for shape in slide.shapes
                if getattr(shape, "has_text_frame", False)
            ).upper()
            for slide in presentation.slides
        ]

        self.assertIn("PERSONA1", slide_texts[0])
        self.assertIn("PERSONA1", slide_texts[1])
        self.assertIn("PERSONA2", slide_texts[2])
        self.assertIn("PERSONA2", slide_texts[3])

    def test_all_2qrs_certificate_types_share_no_master_or_layout_ids(self):
        items = [
            _request(certificate_type, index)
            for index, certificate_type in enumerate(
                main.TEMPLATE_FILENAME_MAP,
                start=1,
            )
        ]
        pptx_bytes = main.generar_pptx_universidad_2qrs(items)

        self._assert_valid_package(pptx_bytes, expected_slides=10)


if __name__ == "__main__":
    unittest.main()
