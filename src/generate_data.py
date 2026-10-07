"""Gera um conjunto reproduzível de vendas sintéticas em português."""

from pathlib import Path

import pandas as pd


SEED = 42
RECORD_COUNT = 2_000
# Data de referência fixa para garantir que o dataset seja reproduzível.
# O período cobre 12 meses completos (out/2025 a set/2026).
END_DATE = pd.Timestamp("2026-09-30")
START_DATE = END_DATE - pd.DateOffset(months=12) + pd.Timedelta(days=1)

PRODUCTS = {
    "Notebook Atlas 14": ("Computadores", 3_899.90),
    "Notebook Boreal Pro": ("Computadores", 5_499.90),
    "Monitor Horizonte 24": ("Monitores", 899.90),
    "Monitor Prisma 27": ("Monitores", 1_599.90),
    "Teclado Mecânico Aurora": ("Acessórios", 349.90),
    "Mouse Óptico Vento": ("Acessórios", 129.90),
    "Headset Imersão": ("Áudio", 279.90),
    "Fone Bluetooth Nuvem": ("Áudio", 459.90),
    "Webcam Clareza HD": ("Periféricos", 219.90),
    "SSD Rápido 1 TB": ("Armazenamento", 599.90),
    "Roteador Conecta Wi-Fi 6": ("Redes", 429.90),
    "Tablet Caminho 10": ("Dispositivos móveis", 1_799.90),
}

SELLERS = [
    "Ana Souza", "Bruno Lima", "Camila Rocha", "Daniel Alves",
    "Eduarda Martins", "Felipe Costa", "Gabriela Ribeiro", "Henrique Melo",
]
REGIONS = ["Norte", "Nordeste", "Centro-Oeste", "Sudeste", "Sul"]
PAYMENT_METHODS = ["Pix", "Cartão de Crédito", "Cartão de Débito", "Boleto"]
REQUIRED_COLUMNS = [
    "sale_id", "sale_date", "product", "category", "quantity",
    "unit_price", "total_value", "seller", "region", "payment_method",
]


def generate_sales() -> pd.DataFrame:
    """Cria 2.000 vendas sintéticas em um período fixo de 12 meses."""
    import numpy as np

    rng = np.random.default_rng(SEED)
    product_names = list(PRODUCTS)
    selected_products = rng.choice(product_names, size=RECORD_COUNT)
    quantities = rng.integers(1, 6, size=RECORD_COUNT)
    sellers = rng.choice(SELLERS, size=RECORD_COUNT)
    regions = rng.choice(REGIONS, size=RECORD_COUNT, p=[0.08, 0.18, 0.10, 0.42, 0.22])
    payment_methods = rng.choice(
        PAYMENT_METHODS, size=RECORD_COUNT, p=[0.38, 0.36, 0.16, 0.10]
    )

    date_range_days = (END_DATE - START_DATE).days
    sale_dates = START_DATE + pd.to_timedelta(
        rng.integers(0, date_range_days + 1, size=RECORD_COUNT), unit="D"
    )

    unit_prices = []
    categories = []
    for product in selected_products:
        category, reference_price = PRODUCTS[product]
        categories.append(category)
        variation_cents = int(reference_price * 100 * 0.08)
        price_cents = int(reference_price * 100) + rng.integers(
            -variation_cents, variation_cents + 1
        )
        unit_prices.append(round(price_cents / 100, 2))

    sales = pd.DataFrame(
        {
            "sale_date": sale_dates,
            "product": selected_products,
            "category": categories,
            "quantity": quantities,
            "unit_price": unit_prices,
            "total_value": [
                round(quantity * price, 2)
                for quantity, price in zip(quantities, unit_prices)
            ],
            "seller": sellers,
            "region": regions,
            "payment_method": payment_methods,
        }
    )

    # Ordena por data antes de gerar a sequência de IDs
    sales = sales.sort_values("sale_date").reset_index(drop=True)
    sales.insert(0, "sale_id", range(1, RECORD_COUNT + 1))

    return sales[REQUIRED_COLUMNS]


def validate_sales(sales: pd.DataFrame) -> None:
    """Interrompe a geração se algum requisito de qualidade não for atendido."""
    assert len(sales) == RECORD_COUNT, "Quantidade de registros incorreta."
    assert not sales["sale_id"].duplicated().any(), "sale_id contém duplicidades."
    assert (sales["quantity"] > 0).all(), "quantity deve ser maior que zero."
    assert (sales["unit_price"] > 0).all(), "unit_price deve ser maior que zero."
    assert (sales["total_value"] > 0).all(), "total_value deve ser maior que zero."
    assert (
        sales["total_value"]
        == (sales["quantity"] * sales["unit_price"]).round(2)
    ).all(), "total_value não corresponde a quantity * unit_price."
    assert sales[REQUIRED_COLUMNS].notna().all().all(), "Há valores nulos."
    assert sales["sale_id"].tolist() == list(range(1, RECORD_COUNT + 1))


def main() -> None:
    project_root = Path(__file__).resolve().parents[1]
    output_path = project_root / "data" / "raw" / "vendas.csv"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    sales = generate_sales()
    validate_sales(sales)
    sales.to_csv(output_path, index=False, encoding="utf-8")

    period_start = sales["sale_date"].min().strftime("%d/%m/%Y")
    period_end = sales["sale_date"].max().strftime("%d/%m/%Y")
    revenue = sales["total_value"].sum()

    print(f"Arquivo gerado: {output_path}")
    print(f"Quantidade de vendas: {len(sales):,}".replace(",", "."))
    print(f"Período dos dados: {period_start} a {period_end}")
    print(f"Receita total: R$ {revenue:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))


if __name__ == "__main__":
    main()