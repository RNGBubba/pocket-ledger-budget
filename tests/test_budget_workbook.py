from pathlib import Path
from zipfile import ZipFile

from budget_workbook import create_workbook


def test_create_workbook_writes_expected_sheets_and_input_headers(tmp_path: Path):
    output = tmp_path / "budget.xlsx"

    create_workbook(output)

    with ZipFile(output) as workbook:
        workbook_xml = workbook.read("xl/workbook.xml").decode("utf-8")
        transactions_xml = workbook.read("xl/worksheets/sheet2.xml").decode("utf-8")

    assert '<sheet name="Dashboard" sheetId="1"' in workbook_xml
    assert '<sheet name="Transactions" sheetId="2"' in workbook_xml
    assert '<sheet name="Monthly Plan" sheetId="3"' in workbook_xml
    assert 'Transaction date' in transactions_xml
    assert 'Amount' in transactions_xml
    assert 'Category' in transactions_xml


def test_create_workbook_includes_working_budget_formulas(tmp_path: Path):
    output = tmp_path / "budget.xlsx"

    create_workbook(output, months=6)

    with ZipFile(output) as workbook:
        dashboard_xml = workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")
        plan_xml = workbook.read("xl/worksheets/sheet3.xml").decode("utf-8")

    assert "SUMIFS" in dashboard_xml
    assert "Monthly Plan" in dashboard_xml
    assert "SUMIFS" in plan_xml
    assert "Remaining" in plan_xml


def test_create_workbook_rejects_non_positive_month_count(tmp_path: Path):
    output = tmp_path / "budget.xlsx"

    try:
        create_workbook(output, months=0)
    except ValueError as error:
        assert str(error) == "months must be positive"
    else:
        raise AssertionError("expected ValueError")
