import pandas as pd

from src.generate_data import END_DATE, START_DATE, RECORD_COUNT, generate_sales, validate_sales


def test_generated_sales_meet_quality_rules():
    sales = generate_sales()

    validate_sales(sales)

    assert len(sales) == RECORD_COUNT
    assert sales["sale_id"].is_unique
    assert sales["sale_date"].min() >= START_DATE
    assert sales["sale_date"].max() <= END_DATE
    assert (sales["total_value"] == (sales["quantity"] * sales["unit_price"]).round(2)).all()


def test_generation_is_reproducible():
    first = generate_sales()
    second = generate_sales()

    pd.testing.assert_frame_equal(first, second)
