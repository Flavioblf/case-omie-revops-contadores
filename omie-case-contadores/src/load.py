"""Carrega as 5 abas do Excel no DuckDB (mesmos nomes de tabela do enunciado) e cria as views de staging.

Uso: python src/load.py
"""
from pathlib import Path

import duckdb
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "data" / "raw" / "base_case_contador_omie.xlsx"
DB = ROOT / "data" / "omie.duckdb"
STAGING = ROOT / "sql" / "00_staging.sql"

TABELAS = ["mercado_contadores", "franquias", "leads_pipeline", "execucao_comercial", "clientes_ativos"]
COLUNAS_DATA = {
    "franquias": ["data_inicio"],
    "leads_pipeline": ["data_entrada", "data_conversao"],
    "clientes_ativos": ["data_ativacao", "data_churn"],
}


def connect(read_only: bool = False) -> duckdb.DuckDBPyConnection:
    return duckdb.connect(str(DB), read_only=read_only)


def main() -> None:
    abas = pd.read_excel(XLSX, sheet_name=None)
    con = connect()
    for nome in TABELAS:
        df = abas[nome]
        for col in COLUNAS_DATA.get(nome, []):
            df[col] = pd.to_datetime(df[col]).dt.date
        con.register("df_tmp", df)
        con.execute(f"CREATE OR REPLACE TABLE {nome} AS SELECT * FROM df_tmp")
        con.unregister("df_tmp")
        print(f"{nome:<20} {len(df):>4} linhas")
    con.execute(STAGING.read_text(encoding="utf-8"))
    for view in ["stg_lead", "stg_contador", "stg_execucao_franquia_mes"]:
        n = con.execute(f"SELECT COUNT(*) FROM {view}").fetchone()[0]
        print(f"{view:<26} {n:>4} linhas")
    con.close()


if __name__ == "__main__":
    main()
