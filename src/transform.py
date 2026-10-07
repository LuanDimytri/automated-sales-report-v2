"""Transformações e fluxo ETL dos dados de vendas."""

from pathlib import Path

import pandas as pd

from src.load_data import load_sales_data


MONTH_NAMES_PT_BR = [
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro",
]


def clean_and_prep_data(df: pd.DataFrame) -> pd.DataFrame:
    """Converte a data da venda e adiciona campos úteis para análise."""
    prepared = df.copy()
    prepared["sale_date"] = pd.to_datetime(prepared["sale_date"], errors="raise")
    prepared["year"] = prepared["sale_date"].dt.year
    prepared["month"] = prepared["sale_date"].dt.month
    prepared["year_month"] = prepared["sale_date"].dt.to_period("M").astype(str)
    prepared["month_name"] = prepared["month"].map(
        lambda month: MONTH_NAMES_PT_BR[month - 1]
    )
    return prepared


def process_sales_pipeline() -> Path:
    """Lê, prepara e salva os dados processados, retornando o caminho final."""
    project_root = Path(__file__).resolve().parents[1]
    raw_path = project_root / "data" / "raw" / "vendas.csv"
    processed_path = project_root / "data" / "processed" / "vendas_processadas.csv"
    processed_path.parent.mkdir(parents=True, exist_ok=True)

    sales = load_sales_data(raw_path)
    processed_sales = clean_and_prep_data(sales)
    processed_sales.to_csv(processed_path, index=False, encoding="utf-8")
    print(f"Dados processados salvos em: {processed_path}")
    print(f"Quantidade de registros: {len(processed_sales):,}".replace(",", "."))
    return processed_path


if __name__ == "__main__":
    process_sales_pipeline()
