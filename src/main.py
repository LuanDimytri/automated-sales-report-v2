"""Ponto de entrada para executar o pipeline completo de vendas."""

import sys
from pathlib import Path

# Permite executar tanto `python src/main.py` quanto `python -m src.main`.
if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.generate_data import main as generate_data
from src.report import main as generate_report
from src.transform import process_sales_pipeline


def run_pipeline() -> None:
    """Executa geração dos dados, ETL e criação do relatório, em sequência."""
    print("\n=== [1/3] Gerando dados sintéticos... ===")
    generate_data()

    print("\n=== [2/3] Processando e preparando os dados... ===")
    process_sales_pipeline()

    print("\n=== [3/3] Calculando KPIs e gerando o relatório Excel... ===")
    generate_report()

    print("\n=== Pipeline concluído com sucesso. ===")


if __name__ == "__main__":
    run_pipeline()
