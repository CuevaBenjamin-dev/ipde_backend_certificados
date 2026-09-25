"""Audit persistence and recoverable Sheets delivery, entirely without network."""

from copy import deepcopy
from datetime import datetime, timedelta, timezone
import os
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

from app import generation_log as log


NOW = datetime(2026, 9, 24, 2, 30, 45, tzinfo=timezone.utc)


def _items(count=5):
    result = []
    for number in range(count):
        data = {
            "modeloCertificado": "INSTITUTO",
            "tipoModelo": "DIPLOMADO",
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
        result.append(SimpleNamespace(model_dump=lambda data=data: deepcopy(data)))
    return result


class MemoryDocument:
    def __init__(self, generation_id):
        self.id = generation_id
        self.data = None
        self.create = MagicMock(side_effect=self._create)
        self.get = MagicMock(side_effect=self._get)

    def _create(self, data, **kwargs):
        if self.data is not None:
            raise ValueError("Document already exists")
        self.data = deepcopy(data)

    def _get(self, **kwargs):
        data = deepcopy(self.data)
        return SimpleNamespace(
            id=self.id,
            exists=data is not None,
            to_dict=lambda: data,
        )


class MemoryFirestore:
    def __init__(self):
        self.documents = {}
        self.updates = []
        self.collection = MagicMock(return_value=SimpleNamespace(document=self.document))
        self.transaction = MagicMock(
            side_effect=lambda **kwargs: SimpleNamespace(update=self.update)
        )

    def document(self, generation_id):
        return self.documents.setdefault(generation_id, MemoryDocument(generation_id))

    def update(self, document, data):
        self.updates.append((document.id, dict(data)))
        for key, value in data.items():
            if value is log.firestore.DELETE_FIELD:
                document.data.pop(key, None)
            else:
                document.data[key] = value


class MemorySheets:
    def __init__(self):
        self.rows = []
        self.headers = [log.HEADERS.copy()]
        self.fail_before_append = False
        self.timeout_after_append = False
        self.on_append = None
        self.values = MagicMock()
        self.values.get.side_effect = self.get
        self.values.append.side_effect = self.append
        self.values.batchUpdate.side_effect = self.batch_update
        self.service = MagicMock()
        self.service.__enter__.return_value = self.service
        self.service.__exit__.return_value = False
        self.service.spreadsheets.return_value.values.return_value = self.values

    @staticmethod
    def request(callback):
        return SimpleNamespace(execute=lambda **kwargs: callback())

    def get(self, *, spreadsheetId, range):
        if range == "'Registros'!A1:Q1":
            return self.request(lambda: {"values": deepcopy(self.headers)})
        if range == "'Registros'!A2:A":
            return self.request(lambda: {"values": [[row[0]] for row in self.rows]})
        raise AssertionError(f"Unexpected range: {range}")

    def append(self, *, body, **kwargs):
        def execute():
            if self.fail_before_append:
                raise ConnectionError("Private response contents must not be logged")
            self.rows.extend(deepcopy(body["values"]))
            if self.on_append:
                self.on_append()
            if self.timeout_after_append:
                raise TimeoutError("Append accepted but response lost")
            return {}

        return self.request(execute)

    def batch_update(self, *, body, **kwargs):
        def execute():
            for update in body["data"]:
                cell = update["range"].split("!")[1]
                column = {"C": 2, "Q": 16}[cell[0]]
                row = int(cell[1:]) - 2
                self.rows[row][column] = update["values"][0][0]
            return {}

        return self.request(execute)


class GenerationLogTests(unittest.TestCase):
    def setUp(self):
        self.db = MemoryFirestore()
        self.sheets = MemorySheets()
        self.firestore_client = self.enterContext(
            patch.object(log, "_get_firestore", return_value=self.db)
        )
        self.sheets_client = self.enterContext(
            patch.object(log, "_get_sheets", return_value=self.sheets.service)
        )
        self.enterContext(
            patch.object(log.firestore, "transactional", side_effect=lambda function: function)
        )
        self.now = self.enterContext(patch.object(log, "_now", return_value=NOW))
        self.enterContext(patch.dict(os.environ, {"GOOGLE_SHEET_ID": "test-sheet"}))

    def register(self, count=5):
        generation_id = log.registrar_generacion(_items(count), filename="Ana - IPDE.pptx")
        return generation_id, self.db.document(generation_id)

    def test_five_certificates_create_exactly_one_pptx_record(self):
        generation_id, document = self.register()
        self.assertRegex(generation_id, r"^[0-9a-f]{32}$")
        self.assertEqual(len(self.db.documents), 1)
        document.create.assert_called_once()
        self.assertEqual(document.data["generation_id"], generation_id)
        self.assertEqual(document.data["cantidad_certificados"], 5)
        self.assertEqual(document.data["items"], [item.model_dump() for item in _items()])
        self.assertEqual(document.data["filename"], "Ana - IPDE.pptx")
        self.assertEqual(document.data["estado"], "GENERADO")
        self.assertFalse(document.data["frontend_confirmed"])
        self.assertEqual(document.data["sheet_status"], "PENDIENTE")
        self.assertEqual(document.create.call_args.kwargs, {"retry": None, "timeout": 10})
        self.sheets_client.assert_not_called()

    def test_date_and_hour_follow_peru_even_when_utc_is_next_day(self):
        _, document = self.register()
        self.assertEqual(document.data["fecha_registro"], "2026-09-23T21:30:45-05:00")
        self.assertEqual(document.data["fecha_peru"], "2026-09-23")
        self.assertEqual(document.data["hora_peru"], "21:30:45")

    def test_primary_write_failure_is_reported_without_using_sheets(self):
        document = MemoryDocument("b" * 32)
        document.create.side_effect = ConnectionError("Unavailable")
        with patch.object(log, "_document", return_value=document):
            with self.assertLogs(log.LOGGER, level="ERROR"):
                with self.assertRaises(log.RegistroNoDisponible):
                    log.registrar_generacion(_items())
        document.create.assert_called_once()
        self.sheets_client.assert_not_called()

    def test_sheet_configuration_is_read_at_call_time(self):
        with patch.dict(os.environ, {"GOOGLE_SHEET_ID": ""}):
            _, document_without_sheet = self.register()
        with patch.dict(os.environ, {"GOOGLE_SHEET_ID": "  later-sheet  "}):
            generation_id, configured_document = self.register()
            self.assertTrue(log.sincronizar_generacion(generation_id))
        self.assertEqual(document_without_sheet.data["sheet_status"], "NO_CONFIGURADO")
        self.assertEqual(configured_document.data["sheet_status"], "OK")
        self.assertEqual(self.sheets.values.append.call_args.kwargs["spreadsheetId"], "later-sheet")

    def test_confirmation_is_idempotent_and_preserves_first_timestamp(self):
        generation_id, document = self.register()
        self.assertTrue(log.confirmar_generacion(generation_id))
        original_time = document.data["confirmado_frontend_at"]
        self.now.return_value = NOW + timedelta(hours=2)
        self.assertTrue(log.confirmar_generacion(generation_id))
        self.assertEqual(len(self.db.updates), 1)
        self.assertEqual(document.data["confirmado_frontend_at"], original_time)
        self.assertEqual(original_time, "2026-09-23T21:30:45-05:00")
        self.assertTrue(document.data["frontend_confirmed"])
        self.assertEqual(document.data["estado"], "CONFIRMADO_FRONTEND")
        self.sheets_client.assert_not_called()

    def test_unknown_confirmation_does_not_create_a_record(self):
        self.assertFalse(log.confirmar_generacion("b" * 32))
        self.db.document("b" * 32).create.assert_not_called()
        self.assertEqual(self.db.updates, [])

    def test_sheet_failure_remains_recoverable_and_does_not_lose_event(self):
        generation_id, document = self.register()
        self.sheets.fail_before_append = True
        with self.assertLogs(log.LOGGER, level="ERROR") as captured:
            self.assertFalse(log.sincronizar_generacion(generation_id))
        self.assertEqual(document.data["sheet_status"], "ERROR")
        self.assertEqual(document.data["sheet_error"], "ConnectionError")
        self.assertNotIn("Private response", "\n".join(captured.output))
        self.assertEqual(document.data["cantidad_certificados"], 5)
        self.assertNotIn("sheet_lease_token", document.data)
        self.sheets.fail_before_append = False
        self.assertTrue(log.sincronizar_generacion(generation_id))
        self.assertEqual(document.data["sheet_status"], "OK")
        self.assertEqual(len(self.sheets.rows), 5)
        self.assertEqual({row[0] for row in self.sheets.rows}, {generation_id})

    def test_retry_after_accepted_append_timeout_does_not_duplicate_rows(self):
        generation_id, document = self.register()
        self.sheets.timeout_after_append = True
        with self.assertLogs(log.LOGGER, level="ERROR"):
            self.assertFalse(log.sincronizar_generacion(generation_id))
        self.assertEqual(len(self.sheets.rows), 5)
        self.sheets.timeout_after_append = False
        self.assertTrue(log.sincronizar_generacion(generation_id))
        self.sheets.values.append.assert_called_once()
        self.sheets.values.batchUpdate.assert_called_once()
        self.assertEqual(len(self.sheets.rows), 5)
        self.assertEqual(document.data["sheet_status"], "OK")

    def test_active_lease_prevents_another_writer(self):
        generation_id, document = self.register()
        document.data.update({
            "sheet_lease_token": "another-worker",
            "sheet_lease_until": NOW + timedelta(minutes=1),
            "sheet_status": "SINCRONIZANDO",
        })
        self.assertFalse(log.sincronizar_generacion(generation_id))
        self.sheets_client.assert_not_called()
        self.assertEqual(self.db.updates, [])
        self.assertEqual(document.data["sheet_lease_token"], "another-worker")

    def test_expired_lease_can_be_recovered_after_worker_restart(self):
        generation_id, document = self.register()
        document.data.update({
            "sheet_lease_token": "stopped-worker",
            "sheet_lease_until": NOW - timedelta(seconds=1),
            "sheet_status": "SINCRONIZANDO",
        })
        self.assertTrue(log.sincronizar_generacion(generation_id))
        self.assertEqual(document.data["sheet_status"], "OK")
        self.assertNotIn("sheet_lease_token", document.data)
        self.assertEqual(len(self.sheets.rows), 5)

    def test_confirmation_during_sync_gets_second_pass_without_duplicate_append(self):
        generation_id, document = self.register()
        self.sheets.on_append = lambda: self.assertTrue(log.confirmar_generacion(generation_id))
        self.assertTrue(log.sincronizar_generacion(generation_id))
        self.sheets.values.append.assert_called_once()
        self.sheets.values.batchUpdate.assert_called_once()
        self.assertEqual(len(self.sheets.rows), 5)
        self.assertEqual(document.data["sheet_status"], "OK")
        self.assertEqual(document.data["estado"], "CONFIRMADO_FRONTEND")
        self.assertTrue(all(row[2] == "CONFIRMADO_FRONTEND" for row in self.sheets.rows))
        self.assertTrue(all(row[16] == "2026-09-23T21:30:45-05:00" for row in self.sheets.rows))

    def test_sorted_rows_keep_their_participant_and_topic_when_confirmed(self):
        generation_id, _ = self.register()
        self.assertTrue(log.sincronizar_generacion(generation_id))
        self.sheets.rows.reverse()
        topics_before = [row[7] for row in self.sheets.rows]
        self.assertTrue(log.confirmar_generacion(generation_id))
        self.assertTrue(log.sincronizar_generacion(generation_id))
        self.assertEqual([row[7] for row in self.sheets.rows], topics_before)
        self.assertTrue(all(row[2] == "CONFIRMADO_FRONTEND" for row in self.sheets.rows))
        self.sheets.values.append.assert_called_once()


if __name__ == "__main__":
    unittest.main()
