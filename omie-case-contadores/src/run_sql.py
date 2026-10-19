"""Executa um arquivo .sql no DuckDB, imprime cada resultado e salva em output/<nome>.csv.

Cada SELECT a exportar deve ser precedido de um comentário `-- @name: <nome>`.
O trecho antes do primeiro `@name` (ex.: CREATE VIEW) é executado sem exportar.
Uso: python src/run_sql.py sql/01c_tiers_prioridade.sql
"""
import re
import sys
from pathlib import Path

import pandas as pd

from load import ROOT, connect

pd.set_option("display.width", 220)
pd.set_option("display.max_columns", 30)
pd.set_option("display.max_rows", 200)


def main(path: str) -> None:
    sql = Path(path).read_text(encoding="utf-8")
    con = connect()
    preambulo, *blocos = re.split(r"^-- @name:\s*", sql, flags=re.M)
    con.execute(preambulo)  # DDL antes do primeiro @name (views auxiliares); comentários puros são no-op
    for bloco in blocos:
        nome, query = bloco.split("\n", 1)
        df = con.execute(query.strip().rstrip(";")).df()
        # SUM de inteiros no DuckDB volta como HUGEINT -> float no pandas; restaura inteiros
        for col in df.select_dtypes("float").columns:
            if (df[col].dropna() % 1 == 0).all():
                df[col] = df[col].astype("Int64")
        print(f"\n=== {nome.strip()} ({len(df)} linhas)")
        print(df.to_string(index=False))
        df.to_csv(ROOT / "output" / f"{nome.strip()}.csv", index=False, encoding="utf-8-sig")
    con.close()


if __name__ == "__main__":
    for p in sys.argv[1:]:
        main(p)
