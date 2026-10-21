from __future__ import annotations

import unittest
from dataclasses import dataclass

from cruce_referencia_valor import cruzar_por_referencia_y_valor, normalizar_referencia
from dinero import a_unidades_minimas
from pse_conciliador import PseConciliador
from test_pse_exactitud import _cruces, _pse, _states


@dataclass
class FakeMov:
    sheet_name: str
    row: int
    value: float
    units: int
    reference_original: str
    matched: bool = False
    pendiente_revision: bool = False


def _mov(row: int, value: float, referencia: str, sheet: str = "PSE") -> FakeMov:
    units = a_unidades_minimas(value)
    assert units is not None
    return FakeMov(sheet, row, value, units, referencia)


class TestNormalizacionReferencia(unittest.TestCase):
    def test_espacios_extremos(self) -> None:
        self.assertEqual(normalizar_referencia("  ABC-01  "), "ABC-01")

    def test_conserva_ceros_iniciales_texto(self) -> None:
        self.assertEqual(normalizar_referencia("00123"), "00123")

    def test_alfanumerica(self) -> None:
        self.assertEqual(normalizar_referencia("REF/2026-01"), "REF/2026-01")

    def test_vacia(self) -> None:
        self.assertIsNone(normalizar_referencia("   "))
        self.assertIsNone(normalizar_referencia(None))


class TestCruceReferenciaValor(unittest.TestCase):
    def test_referencia_y_valor_coinciden(self) -> None:
        origen = [_mov(2, 150000, "  REF-01")]
        destino = [_mov(5, 150000, "REF-01", "CRUCES")]
        confirmados, pendientes = cruzar_por_referencia_y_valor(origen, destino)
        self.assertEqual(len(confirmados), 1)
        self.assertEqual(pendientes, [])
        self.assertTrue(origen[0].matched)
        self.assertEqual(confirmados[0].referencia_origen, "  REF-01")

    def test_misma_referencia_valor_distinto(self) -> None:
        origen = [_mov(2, 150000, "REF-01")]
        destino = [_mov(5, 150001, "REF-01", "CRUCES")]
        confirmados, pendientes = cruzar_por_referencia_y_valor(origen, destino)
        self.assertEqual(confirmados, [])
        self.assertEqual(pendientes, [])
        self.assertFalse(origen[0].matched)

    def test_mismo_valor_referencia_distinta(self) -> None:
        origen = [_mov(2, 150000, "REF-01")]
        destino = [_mov(5, 150000, "REF-02", "CRUCES")]
        confirmados, pendientes = cruzar_por_referencia_y_valor(origen, destino)
        self.assertEqual(confirmados, [])
        self.assertEqual(pendientes, [])

    def test_duplicados_pendientes(self) -> None:
        origen = [_mov(2, 150000, "REF-01"), _mov(3, 150000, "REF-01")]
        destino = [_mov(5, 150000, "REF-01", "CRUCES")]
        confirmados, pendientes = cruzar_por_referencia_y_valor(origen, destino)
        self.assertEqual(confirmados, [])
        self.assertEqual(len(pendientes), 1)
        self.assertTrue(all(item.pendiente_revision for item in origen))

    def test_ya_conciliado_no_reutiliza(self) -> None:
        origen = [_mov(2, 150000, "REF-01")]
        destino = [_mov(5, 150000, "REF-01", "CRUCES")]
        destino[0].matched = True
        confirmados, pendientes = cruzar_por_referencia_y_valor(origen, destino)
        self.assertEqual(confirmados, [])
        self.assertEqual(pendientes, [])

    def test_integrado_en_pse_con_ceros_iniciales(self) -> None:
        pse = _pse([["2026-04-15", 25000, "Pago PSE", "00123"]])
        cruces = _cruces([["2026-04-15", 25000, "Pago virtual PSE", "00123"]])
        resultado = PseConciliador(pse, cruces).procesar()
        self.assertEqual(_states(resultado)[2], "Conciliado por referencia y valor")
        self.assertIn("00123", resultado["dataset"][0]["comentario_conciliacion"])

    def test_integrado_referencia_igual_valor_no(self) -> None:
        pse = _pse([["2026-04-15", 25000, "Pago PSE", "00123"]])
        cruces = _cruces([["2026-04-15", 25001, "Pago virtual PSE", "00123"]])
        resultado = PseConciliador(pse, cruces).procesar()
        self.assertEqual(_states(resultado)[2], "Sin coincidencia")

    def test_integrado_valor_igual_referencia_no(self) -> None:
        pse = _pse([["2026-04-15", 25000, "Pago PSE", "AA"]])
        cruces = _cruces([["2026-04-15", 25000, "Pago virtual PSE", "BB"]])
        resultado = PseConciliador(pse, cruces).procesar()
        self.assertEqual(_states(resultado)[2], "Conciliado 1:1")

    def test_referencia_vacia_no_confirma_por_referencia(self) -> None:
        pse = _pse([["2026-04-15", 88000, "Pago PSE", ""]])
        cruces = _cruces([["2026-04-15", 88000, "Pago virtual PSE", ""]])
        resultado = PseConciliador(pse, cruces).procesar()
        self.assertEqual(_states(resultado)[2], "Conciliado 1:1")


if __name__ == "__main__":
    unittest.main()
