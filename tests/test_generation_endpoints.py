"""HTTP contract for one audit event per downloaded PPTX batch."""

import os
import unittest
from unittest.mock import MagicMock, patch

from fastapi.testclient import TestClient


os.environ.setdefault("OPENAI_API_KEY", "test-key")

from app import generation_log, main


GENERATION_ID = "a1" * 16
PPTX_CONTENT = b"synthetic-pptx-content"
PPTX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


def _item(model, number):
    return {
        "modeloCertificado": model,
        "tipoModelo": "DIPLOMADO" if number == 1 else "CURSO",
        "nombres": "Ana",
        "apellidos": "Perez",
        "temaDiplomado": f"Tema {number}",
        "fechaInicio": "2026-01-01",
        "fechaFin": "2026-09-01",
        "horasAcademicas": 240,
        "creditosAcademicos": 15,
        "folioNumero": f"FOLIO-{number}",
        "fechaEmision": "2026-09-24",
        "codigoEstudiante": "codigo-prueba",
        "ciudad": "Trujillo",
    }


class GenerationEndpointTests(unittest.TestCase):
    def setUp(self):
        self.events = []
        self.register = self.enterContext(
            patch.object(main, "registrar_generacion", side_effect=self._register)
        )
        self.confirm = self.enterContext(
            patch.object(main, "confirmar_generacion", return_value=True)
        )
        self.sync = self.enterContext(
            patch.object(
                main,
                "sincronizar_generacion",
                side_effect=lambda generation_id: self.events.append("synced"),
            )
        )
        self.generate_item = self.enterContext(
            patch.object(main, "generar_presentacion_por_item")
        )
        self.merged = MagicMock()
        self.merged.slides = [object(), object()]
        self.merged.save.side_effect = self._save
        self.merge = self.enterContext(
            patch.object(main, "merge_presentations", return_value=self.merged)
        )
        self.generate_2qrs = self.enterContext(
            patch.object(
                main, "generar_pptx_universidad_2qrs", side_effect=self._save_2qrs
            )
        )
        self.enterContext(patch.object(main, "generate_qr_image"))
        self.enterContext(patch.object(main, "insert_qr_at_placeholder"))
        self.client = self.enterContext(
            TestClient(main.app, raise_server_exceptions=False)
        )

    def _register(self, items, *, filename):
        self.events.append("registered")
        return GENERATION_ID

    def _save(self, output):
        output.write(PPTX_CONTENT)
        self.events.append("pptx_saved")

    def _save_2qrs(self, items):
        self.events.append("pptx_saved")
        return PPTX_CONTENT

    def _payload(self, model="INSTITUTO"):
        return {"items": [_item(model, 1), _item(model, 2)]}

    def _assert_success(self, model):
        payload = self._payload(model)
        response = self.client.post(
            "/api/diplomas",
            json=payload,
            headers={"Origin": "https://ipdefrontendcertificados.vercel.app"},
        )

        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.content, PPTX_CONTENT)
        self.assertEqual(response.headers["content-type"], PPTX_MEDIA_TYPE)
        self.assertEqual(response.headers["x-generation-id"], GENERATION_ID)
        self.assertIn("attachment;", response.headers["content-disposition"])
        self.register.assert_called_once()
        registered_items = self.register.call_args.args[0]
        self.assertEqual([item.model_dump() for item in registered_items], payload["items"])
        self.assertEqual(
            self.register.call_args.kwargs["filename"],
            main.construir_nombre_archivo_pptx(registered_items, model),
        )
        self.sync.assert_called_once_with(GENERATION_ID)
        self.assertEqual(self.events, ["pptx_saved", "registered", "synced"])
        self.confirm.assert_not_called()

        exposed_headers = {
            header.strip().lower()
            for header in response.headers["access-control-expose-headers"].split(",")
        }
        self.assertTrue({"content-disposition", "x-generation-id"} <= exposed_headers)
        self.assertEqual(
            response.headers["access-control-allow-origin"],
            "https://ipdefrontendcertificados.vercel.app",
        )

    def test_normal_batch_records_once_after_complete_pptx_and_exposes_header(self):
        self._assert_success("INSTITUTO")
        self.assertEqual(self.generate_item.call_count, 2)
        self.merge.assert_called_once()
        self.merged.save.assert_called_once()
        self.generate_2qrs.assert_not_called()

    def test_2qrs_batch_records_once_after_complete_pptx_and_exposes_header(self):
        self._assert_success("UNIVERSIDAD_2QRS")
        self.generate_2qrs.assert_called_once()
        self.assertEqual(len(self.generate_2qrs.call_args.args[0]), 2)
        self.generate_item.assert_not_called()
        self.merge.assert_not_called()

    def test_failed_item_generation_does_not_create_a_log(self):
        self.generate_item.side_effect = RuntimeError("Template generation failed")
        response = self.client.post("/api/diplomas", json=self._payload())
        self.assertEqual(response.status_code, 500)
        self.register.assert_not_called()
        self.sync.assert_not_called()

    def test_failed_final_pptx_save_does_not_create_a_log(self):
        self.merged.save.side_effect = RuntimeError("PPTX export failed")
        response = self.client.post("/api/diplomas", json=self._payload())
        self.assertEqual(response.status_code, 500)
        self.register.assert_not_called()
        self.sync.assert_not_called()

    def test_failed_2qrs_generation_does_not_create_a_log(self):
        self.generate_2qrs.side_effect = RuntimeError("PPTX merge failed")
        response = self.client.post(
            "/api/diplomas", json=self._payload("UNIVERSIDAD_2QRS")
        )
        self.assertEqual(response.status_code, 500)
        self.register.assert_not_called()
        self.sync.assert_not_called()

    def test_failed_registration_does_not_deliver_an_untracked_pptx(self):
        self.register.side_effect = generation_log.RegistroNoDisponible("unavailable")
        for model in ("INSTITUTO", "UNIVERSIDAD_2QRS"):
            with self.subTest(model=model):
                response = self.client.post("/api/diplomas", json=self._payload(model))
                self.assertEqual(response.status_code, 503, response.text)
                self.assertIn("detail", response.json())
                self.assertNotEqual(response.content, PPTX_CONTENT)
                self.assertNotIn("x-generation-id", response.headers)
                self.assertNotIn("content-disposition", response.headers)
        self.sync.assert_not_called()

    def test_confirmation_returns_id_and_schedules_sheet_update(self):
        response = self.client.post(f"/api/diplomas/confirmar/{GENERATION_ID}", json={})
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json(), {"ok": True, "generation_id": GENERATION_ID})
        self.confirm.assert_called_once_with(GENERATION_ID)
        self.sync.assert_called_once_with(GENERATION_ID)
        self.register.assert_not_called()

    def test_confirmation_for_missing_generation_returns_404(self):
        self.confirm.return_value = False
        response = self.client.post(f"/api/diplomas/confirmar/{GENERATION_ID}", json={})
        self.assertEqual(response.status_code, 404)
        self.confirm.assert_called_once_with(GENERATION_ID)
        self.sync.assert_not_called()

    def test_confirmation_when_log_is_unavailable_returns_503(self):
        self.confirm.side_effect = generation_log.RegistroNoDisponible("unavailable")
        response = self.client.post(f"/api/diplomas/confirmar/{GENERATION_ID}", json={})
        self.assertEqual(response.status_code, 503)
        self.sync.assert_not_called()

    def test_invalid_confirmation_ids_are_rejected_before_firestore(self):
        for invalid_id in ("not-an-id", "a" * 31, "a" * 33, "z" * 32, "A" * 32):
            with self.subTest(generation_id=invalid_id):
                response = self.client.post(
                    f"/api/diplomas/confirmar/{invalid_id}", json={}
                )
                self.assertEqual(response.status_code, 422, response.text)
        self.confirm.assert_not_called()
        self.sync.assert_not_called()


if __name__ == "__main__":
    unittest.main()
