import os
import unittest


os.environ.setdefault("OPENAI_API_KEY", "test-key")

from app import main


class ColegioProfesoresNuevoModelTests(unittest.TestCase):
    MODEL_KEY = "COLEGIO_DE_PROFESORES_DEL_PERU_NUEVO"
    EXPECTED_FOLDER = "colegio_de_profesores_del_peru_nuevo"

    def test_new_model_resolves_all_existing_certificate_types(self):
        for certificate_type, filename in main.TEMPLATE_FILENAME_MAP.items():
            with self.subTest(certificate_type=certificate_type):
                self.assertEqual(
                    main.resolve_template_path(self.MODEL_KEY, certificate_type),
                    os.path.join("app", "templates", self.EXPECTED_FOLDER, filename),
                )

    def test_new_model_uses_same_long_date_format_as_original_cpp(self):
        self.assertIn("COLEGIO_DE_PROFESORES_DEL_PERU", main.MODELOS_FECHA_LARGA)
        self.assertIn(self.MODEL_KEY, main.MODELOS_FECHA_LARGA)

    def test_new_model_has_a_distinct_download_filename(self):
        self.assertEqual(main.MODEL_FILENAME_LABELS[self.MODEL_KEY], "CPP NUEVO")
        self.assertEqual(
            main.MODEL_FILENAME_LABELS["COLEGIO_DE_PROFESORES_DEL_PERU"],
            "CPP",
        )


if __name__ == "__main__":
    unittest.main()
