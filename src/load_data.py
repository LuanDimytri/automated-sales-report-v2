"""Leitura e validação básica do arquivo CSV de vendas."""

from pathlib import Path

import pandas as pd


REQUIRED_COLUMNS = [
    "sale_id",
    "sale_date",
    "product",
    "category",
    "quantity",
    "unit_price",
    "total_value",
    "seller",
    "region",
    "payment_method",
]


def load_sales_data(file_path: str | Path) -> pd.DataFrame:
    """Carrega o CSV de vendas após verificar arquivo, conteúdo e colunas."""
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"Arquivo de vendas não encontrado: {path}")

    try:
        df = pd.read_csv(path)
    except pd.errors.EmptyDataError as exc:
        raise ValueError(f"O arquivo CSV está vazio: {path}") from exc

    if df.empty:
        raise ValueError(f"O arquivo CSV não possui registros: {path}")

    missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing_columns:
        missing = ", ".join(missing_columns)
        raise ValueError(f"O CSV não possui as colunas obrigatórias: {missing}")

    return df
