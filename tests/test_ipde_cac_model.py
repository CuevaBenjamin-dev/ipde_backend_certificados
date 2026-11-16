"""El modelo IPDE_CAC debe producir lotes PPTX legibles con datos propios."""

import io
import os
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from fastapi.testclient import TestClient
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches


os.environ.setdefault("OPENAI_API_KEY", "test-key")

from app import main


MODULES = [
    "FUNDAMENTOS DE ELECTRICIDAD Y SEGURIDAD ELÉCTRICA",
    "CIRCUITOS ELÉCTRICOS Y MEDICIÓN",
    "MATERIALES Y PROTECCIONES ELÉCTRICAS",
    "MONTAJE DE ALUMBRADO",
    "MANTENIMIENTO PREVENTIVO",
]


def _item(number: int) -> dict:
    return {
        "modeloCertificado": "IPDE_CAC",
        "tipoModelo": "CURSO",
        "nombres": f"PERSONA{number}",
        "apellidos": "PRUEBA",
        "temaDiplomado": f"ELECTRICIDAD TEMA {number}",
        "fechaInicio": "2026-03-01",
        "fechaFin": "2026-06-30",
        "horasAcademicas": 420,
        "creditosAcademicos": 26,
        "folioNumero": f"FOLIO-{number}",
        "fechaEmision": "2026-06-30",
        "codigoEstudiante": f"codigo-{number}",
        "ciudad": "Trujillo",
    }


class IPDECACModelTests(unittest.TestCase):
    def test_three_certificate_types_are_available(self):
        available = {"CURSO", "DIPLOMADO", "PROGRAMA DE ESPECIALIZACIÓN"}
        for certificate_type in available:
            with self.subTest(certificate_type=certificate_type):
                self.assertTrue(os.path.isfile(main.resolve_template_path("IPDE_CAC", certificate_type)))
        for certificate_type in main.TEMPLATE_FILENAME_MAP:
            if certificate_type in available:
                continue
            with self.subTest(certificate_type=certificate_type):
                with self.assertRaises(main.HTTPException) as error:
                    main.resolve_template_path("IPDE_CAC", certificate_type)
                self.assertEqual(error.exception.status_code, 400)

    def test_all_module_titles_share_one_font_size_for_each_type(self):
        for certificate_type, count in (
            ("CURSO", 5),
            ("DIPLOMADO", 8),
            ("PROGRAMA DE ESPECIALIZACIÓN", 8),
        ):
            modules = [
                f"MÓDULO DE PRUEBA {number}"
                if number != count else
                "INSTALACIÓN Y MANTENIMIENTO DE SISTEMAS ELÉCTRICOS RESIDENCIALES CON SEGURIDAD AVANZADA"
                for number in range(1, count + 1)
            ]
            payload = _item(1)
            payload["tipoModelo"] = certificate_type
            with self.subTest(certificate_type=certificate_type):
                with patch.object(main, "obtener_modulos_por_tema", return_value=modules):
                    presentation = main.generar_presentacion_por_item(main.DiplomaRequest(**payload))
                back = presentation.slides[1]
                text_shapes = [
                    shape for shape in back.shapes
                    if shape.has_text_frame and shape.text in modules
                ]
                self.assertEqual(len(text_shapes), count)
                font_sizes = {
                    run.font.size.pt
                    for shape in text_shapes
                    for paragraph in shape.text_frame.paragraphs
                    for run in paragraph.runs
                    if run.text.strip()
                }
                self.assertEqual(len(font_sizes), 1)
                self.assertTrue(5.0 <= next(iter(font_sizes)) <= 9.0)
                self.assertTrue(all(
                    shape.text_frame.auto_size == MSO_AUTO_SIZE.NONE
                    for shape in text_shapes
                ))

    def test_http_batch_uses_each_participants_data_and_qr(self):
        with (
            patch.object(main, "obtener_modulos_por_tema", return_value=MODULES),
            patch.object(main, "registrar_generacion", return_value="a" * 32),
            patch.object(main, "sincronizar_generacion"),
            TestClient(main.app) as client,
        ):
            response = client.post("/api/diplomas", json={"items": [_item(1), _item(2)]})

        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.headers["x-generation-id"], "a" * 32)
        with ZipFile(io.BytesIO(response.content)) as archive:
            names = archive.namelist()
            self.assertEqual(len(names), len(set(names)))
            self.assertIsNone(archive.testzip())

        presentation = Presentation(io.BytesIO(response.content))
        self.assertEqual(len(presentation.slides), 4)
        for number, front_index in ((1, 0), (2, 2)):
            front = presentation.slides[front_index]
            back = presentation.slides[front_index + 1]
            front_text = "\n".join(shape.text for shape in front.shapes if shape.has_text_frame)
            back_text = "\n".join(shape.text for shape in back.shapes if shape.has_text_frame)
            self.assertIn(f"PERSONA{number} PRUEBA", front_text)
            self.assertIn(f"ELECTRICIDAD TEMA {number}", front_text)
            self.assertIn(f"FOLIO-{number}", back_text)
            self.assertIn(f"codigo-{number}", back_text)
            self.assertIn("Trujillo,", front_text)
            self.assertEqual(back_text.count("MÓDULO "), 5)
            self.assertFalse("{{" in front_text or "{{" in back_text)
            self.assertNotIn("ELMER DE LA CRUZ", front_text + back_text)
            qr_pictures = [
                shape for shape in back.shapes
                if shape.shape_type == main.MSO_SHAPE_TYPE.PICTURE
                and abs(shape.left - Inches(1.30)) < Inches(0.1)
                and abs(shape.top - Inches(5.49)) < Inches(0.1)
            ]
            self.assertEqual(len(qr_pictures), 1)


if __name__ == "__main__":
    unittest.main()
