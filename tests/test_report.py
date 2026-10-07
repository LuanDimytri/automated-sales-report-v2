import pandas as pd

from src.report import calculate_kpis, generate_excel_report, get_monthly_trend, get_sales_by_product


def sample_sales():
    return pd.DataFrame({
        "sale_date": pd.to_datetime(["2026-01-10", "2026-01-15", "2026-02-02"]),
        "year_month": ["2026-01", "2026-01", "2026-02"],
        "product": ["Produto A", "Produto B", "Produto A"],
        "category": ["Categoria 1", "Categoria 1", "Categoria 1"],
        "quantity": [2, 1, 3],
        "unit_price": [100.0, 200.0, 100.0],
        "total_value": [200.0, 200.0, 300.0],
        "seller": ["Ana", "Bruno", "Ana"],
        "region": ["Sul", "Sudeste", "Sul"],
        "payment_method": ["Pix", "Cartão", "Pix"],
    })


def test_calculate_kpis():
    kpis = calculate_kpis(sample_sales())

    assert kpis["faturamento_total"] == 700.0
    assert kpis["quantidade_total"] == 6
    assert kpis["total_vendas"] == 3
    assert kpis["ticket_medio"] == 700 / 3


def test_product_ranking_and_monthly_trend():
    df = sample_sales()
    products = get_sales_by_product(df)
    monthly = get_monthly_trend(df)

    assert products.iloc[0]["product"] == "Produto A"
    assert products.iloc[0]["faturamento_total"] == 500.0
    assert monthly["year_month"].tolist() == ["2026-01", "2026-02"]


def test_excel_report_contains_expected_sheets(tmp_path):
    output = tmp_path / "relatorio.xlsx"
    generate_excel_report(sample_sales(), output)

    workbook = pd.ExcelFile(output)
    assert workbook.sheet_names == [
        "Dashboard",
        "Categorias",
        "Produtos",
        "Vendedores",
        "Regiões",
        "Pagamentos",
        "Evolução Mensal",
        "Dados",
        "_dados_dashboard",
    ]
