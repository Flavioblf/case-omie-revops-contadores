"""Abre um .xlsx no Excel (via COM), recalcula todas as fórmulas e salva — deixa os valores em cache
para quem abrir o arquivo em outro visualizador.

Uso: python src/recalc_excel.py output/Desafio2_Propensao.xlsx
"""
import sys
from pathlib import Path

import win32com.client as win32


def recalc(path: str) -> None:
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(Path(path).resolve()))
        excel.CalculateFull()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


if __name__ == "__main__":
    for p in sys.argv[1:]:
        recalc(p)
        print("Recalculado:", p)
