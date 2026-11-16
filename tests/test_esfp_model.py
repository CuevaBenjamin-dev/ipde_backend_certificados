"""ESFP keeps its source artwork while filling five or eight module rows."""

import io
import os
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from fastapi.testclient import TestClient
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Cm


os.environ.setdefault("OPENAI_API_KEY", "test-key")

from app import main


LONG_MODULES = [
    "PLANIFICACIÓN, EVALUACIÓN Y DESARROLLO DE ESTRATEGIAS EDUCATIVAS PARA "
    "LA ENSEÑANZA Y EL APRENDIZAJE EN CONTEXTOS COMPLEJOS " + str(number)
    for number in range(1, 9)
]


def item(number: int, certificate_type: str) -> dict:
    return {
        "modeloCertificado": "ESFP",
        "tipoModelo": certificate_type,
        "nombres": f"PERSONA{number}",
        "apellidos": "PRUEBA",
        "temaDiplomado": f"GESTIÓN EDUCATIVA {number}",
        "fechaInicio": "2026-03-01",
        "fechaFin": "2026-06-30",
        "horasAcademicas": 420,
        "creditosAcademicos": 26,
        "folioNumero": f"FOLIO-{number}",
        "fechaEmision": "2026-07-01",
        "codigoEstudiante": f"ESFP-{number}",
        "ciudad": "Trujillo",
    }


def all_text(slide) -> str:
    pieces = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            pieces.append(shape.text)
        if shape.has_table:
            pieces.extend(cell.text for row in shape.table.rows for cell in row.cells)
    return "\n".join(pieces)


def module_table(slide):
    return next(
        shape
        for shape in slide.shapes
        if shape.has_table and any(
            "MÓDULO I" in cell.text
            for row in shape.table.rows for cell in row.cells
        )
    )


class ESFPModelTests(unittest.TestCase):
    def test_only_three_types_are_available(self):
        supported = {"CURSO", "DIPLOMADO", "PROGRAMA DE ESPECIALIZACIÓN"}
        for certificate_type in supported:
            with self.subTest(certificate_type=certificate_type):
                self.assertTrue(os.path.isfile(main.resolve_template_path("ESFP", certificate_type)))
        for certificate_type in main.TEMPLATE_FILENAME_MAP:
            if certificate_type not in supported:
                with self.subTest(certificate_type=certificate_type):
                    with self.assertRaises(main.HTTPException) as error:
                        main.resolve_template_path("ESFP", certificate_type)
                    self.assertEqual(error.exception.status_code, 400)

    def test_rows_and_font_size_are_consistent_for_long_titles(self):
        for certificate_type, count in (
            ("CURSO", 5),
            ("DIPLOMADO", 8),
            ("PROGRAMA DE ESPECIALIZACIÓN", 8),
        ):
            with self.subTest(certificate_type=certificate_type):
                with patch.object(main, "obtener_modulos_por_tema", return_value=LONG_MODULES[:count]):
                    presentation = main.generar_presentacion_por_item(
                        main.DiplomaRequest(**item(1, certificate_type))
                    )
                self.assertEqual(len(presentation.slides), 2)
                front, back = presentation.slides
                table_shape = module_table(back)
                rows = table_shape.table.rows
                self.assertEqual(len(rows), count)
                self.assertLess(table_shape.top + table_shape.height, Cm(15.7))
                self.assertEqual(
                    [row.cells[1].text for row in rows], LONG_MODULES[:count]
                )
                sizes = {
                    run.font.size.pt
                    for row in rows
                    for paragraph in row.cells[1].text_frame.paragraphs
                    for run in paragraph.runs
                    if run.text.strip()
                }
                self.assertEqual(len(sizes), 1)
                self.assertGreaterEqual(next(iter(sizes)), 5)
                self.assertTrue(all(
                    row.cells[1].text_frame.auto_size == MSO_AUTO_SIZE.NONE
                    for row in rows
                ))
                self.assertIn("PERSONA1 PRUEBA", all_text(front))
                self.assertIn("Trujillo,", all_text(front))
                self.assertIn("FOLIO-1", all_text(back))
                self.assertIn("esfp-1", all_text(back))
                self.assertNotIn("{{", (all_text(front) + all_text(back)).replace("{{QR_CODE}}", ""))

    def test_long_front_topic_does_not_overlap_the_text_below(self):
        long_topic = "GESTIÓN PEDAGÓGICA Y EVALUACIÓN CURRICULAR"
        for certificate_type in ("CURSO", "DIPLOMADO", "PROGRAMA DE ESPECIALIZACIÓN"):
            payload = item(1, certificate_type)
            payload["temaDiplomado"] = long_topic
            with self.subTest(certificate_type=certificate_type):
                with patch.object(
                    main,
                    "obtener_modulos_por_tema",
                    return_value=LONG_MODULES[:main.MODULOS_COUNT[certificate_type]],
                ):
                    presentation = main.generar_presentacion_por_item(main.DiplomaRequest(**payload))
                slide = presentation.slides[0]
                lead = main.find_shape_by_name(slide, "TextBox 16")
                topic = main.find_shape_by_name(slide, "TextBox 17")
                dates = main.find_shape_by_name(slide, "TextBox 18")
                closing = main.find_shape_by_name(slide, "TextBox 19")
                location = main.find_shape_by_name(slide, "TextBox 20")
                self.assertLessEqual(lead.top + lead.height, topic.top)
                self.assertLessEqual(topic.top + topic.height, dates.top)
                self.assertLessEqual(dates.top + dates.height, closing.top)
                self.assertLessEqual(location.top + location.height, Cm(14.35))

    def test_mixed_batch_has_each_persons_data_and_qr(self):
        inputs = [
            item(1, "CURSO"),
            item(2, "DIPLOMADO"),
            item(3, "PROGRAMA DE ESPECIALIZACIÓN"),
        ]
        with (
            patch.object(main, "obtener_modulos_por_tema", side_effect=lambda tipo, _: LONG_MODULES[:main.MODULOS_COUNT[tipo]]),
            patch.object(main, "registrar_generacion", return_value="b" * 32),
            patch.object(main, "sincronizar_generacion"),
            TestClient(main.app) as client,
        ):
            response = client.post("/api/diplomas", json={"items": inputs})
        self.assertEqual(response.status_code, 200, response.text)
        with ZipFile(io.BytesIO(response.content)) as archive:
            self.assertIsNone(archive.testzip())
            self.assertEqual(len(archive.namelist()), len(set(archive.namelist())))
            self.assertIn(
                b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                archive.read("[Content_Types].xml"),
            )
        presentation = Presentation(io.BytesIO(response.content))
        self.assertEqual(len(presentation.slides), 6)
        for number, offset in enumerate((0, 2, 4), start=1):
            front, back = presentation.slides[offset], presentation.slides[offset + 1]
            self.assertIn(f"PERSONA{number} PRUEBA", all_text(front))
            self.assertIn(f"FOLIO-{number}", all_text(back))
            qr_pictures = [
                shape for shape in back.shapes
                if shape.shape_type == main.MSO_SHAPE_TYPE.PICTURE
                and shape.left < Cm(5) and shape.top > Cm(14)
                and Cm(2.5) <= shape.width <= Cm(2.7)
            ]
            self.assertEqual(len(qr_pictures), 1)


if __name__ == "__main__":
    unittest.main()
