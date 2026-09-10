import os
import unittest
from unittest.mock import MagicMock, patch

from google.api_core.exceptions import AlreadyExists


os.environ.setdefault("OPENAI_API_KEY", "test-key")

from app import firestore_cache
from app import main


class FirestoreCacheTests(unittest.TestCase):
    def test_normalization_reuses_the_same_key(self):
        variants = ["Derecho Penal", " derecho penal ", "DERECHO PENAL"]
        keys = {firestore_cache.normalizar_clave_tema(value) for value in variants}
        self.assertEqual(keys, {"derecho-penal"})

    @patch("app.firestore_cache._get_client")
    def test_firestore_lookup_returns_valid_modules(self, get_client):
        snapshot = MagicMock()
        snapshot.exists = True
        snapshot.to_dict.return_value = {"modulos": ["uno", "dos", "tres", "cuatro", "cinco"]}
        get_client.return_value.collection.return_value.document.return_value.get.return_value = snapshot

        result = firestore_cache.buscar_modulos_en_firestore("5_MODULOS::tema", 5)

        self.assertEqual(result, ["UNO", "DOS", "TRES", "CUATRO", "CINCO"])

    @patch("app.firestore_cache._get_client", side_effect=RuntimeError("Firestore no disponible"))
    def test_firestore_lookup_failure_is_a_cache_miss(self, _get_client):
        with self.assertLogs("app.firestore_cache", level="ERROR"):
            self.assertIsNone(firestore_cache.buscar_modulos_en_firestore("5_MODULOS::tema", 5))

    @patch("app.firestore_cache._get_client")
    def test_firestore_save_is_idempotent_when_document_exists(self, get_client):
        document = get_client.return_value.collection.return_value.document.return_value
        document.create.side_effect = AlreadyExists("el documento ya existe")

        saved = firestore_cache.guardar_modulos_en_firestore(
            cache_key="5_MODULOS::derecho-penal",
            grupo="5_MODULOS",
            tipo="CURSO",
            tema="Derecho Penal",
            modulos=["uno", "dos", "tres", "cuatro", "cinco"],
        )

        self.assertTrue(saved)
        document.set.assert_called_once()
        self.assertTrue(document.set.call_args.kwargs["merge"])

    @patch("app.main.buscar_modulos_en_firestore")
    @patch("app.main.obtener_modulos_desde_base")
    def test_base_cache_has_priority_over_firestore(self, base_lookup, firestore_lookup):
        base_lookup.return_value = [f"BASE {index}" for index in range(5)]

        result = main.obtener_modulos_por_tema("CURSO", "Derecho Penal")

        self.assertEqual(result, base_lookup.return_value)
        firestore_lookup.assert_not_called()

    @patch("app.main.buscar_modulos_en_firestore")
    @patch("app.main.obtener_modulos_desde_base", return_value=None)
    def test_firestore_prevents_an_openai_call(self, _base_lookup, firestore_lookup):
        firestore_lookup.return_value = [f"FIRESTORE {index}" for index in range(5)]

        with patch.object(main.client.responses, "create") as openai_create:
            result = main.obtener_modulos_por_tema("CURSO", "Derecho Penal")

        self.assertEqual(result, firestore_lookup.return_value)
        openai_create.assert_not_called()

    @patch("app.main.guardar_modulos_en_firestore", return_value=False)
    @patch("app.main.buscar_modulos_en_firestore", return_value=None)
    @patch("app.main.obtener_modulos_desde_base", return_value=None)
    def test_firestore_write_failure_does_not_discard_openai_result(
        self,
        _base_lookup,
        _firestore_lookup,
        save_firestore,
    ):
        response = MagicMock()
        response.output_text = '{"modulos":["a","b","c","d","e"]}'

        with patch.object(main.client.responses, "create", return_value=response):
            result = main.obtener_modulos_por_tema("CURSO", "Tema nuevo")

        self.assertEqual(result, ["A", "B", "C", "D", "E"])
        save_firestore.assert_called_once()


if __name__ == "__main__":
    unittest.main()
