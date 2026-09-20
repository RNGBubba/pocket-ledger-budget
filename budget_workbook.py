"""Generate an original, self-contained personal budget workbook in XLSX format."""

from __future__ import annotations

from datetime import date
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


MAX_TRANSACTIONS = 200
CURRENCY_FORMAT = '$#,##0.00;[Red]-$#,##0.00'


def _cell_ref(column: int, row: int) -> str:
    letters = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"


def _inline_cell(ref: str, value: object, style: int = 0) -> str:
    style_attr = f' s="{style}"' if style else ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'
    text = escape(str(value))
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{text}</t></is></c>'


def _formula_cell(ref: str, formula: str, style: int = 0) -> str:
    style_attr = f' s="{style}"' if style else ""
    return f'<c r="{ref}"{style_attr}><f>{escape(formula)}</f><v></v></c>'


def _row(row_number: int, cells: list[str], height: int | None = None) -> str:
    height_attr = f' ht="{height}" customHeight="1"' if height else ""
    return f'<row r="{row_number}"{height_attr}>{"".join(cells)}</row>'


def _worksheet(rows: list[str], dimension: str, widths: dict[str, int], frozen: str) -> str:
    cols = "".join(
        f'<col min="{column}" max="{column}" width="{width}" customWidth="1"/>'
        for column, width in widths.items()
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="{dimension}"/><sheetViews><sheetView workbookViewId="0">'
        f'<pane ySplit="{frozen[1:]}" topLeftCell="{frozen}" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A1" sqref="A1"/></sheetView></sheetViews>'
        f'<cols>{cols}</cols><sheetData>{"".join(rows)}</sheetData>'
        '<autoFilter ref="A4:F204"/>'
        '</worksheet>'
    )


def _dashboard() -> str:
    rows = [
        _row(1, [_inline_cell("A1", "Pocket Ledger Budget", 1)], 28),
        _row(2, [_inline_cell("A2", "A calm monthly view of what is planned, spent, and still available.", 2)]),
        _row(4, [_inline_cell("A4", "At a glance", 3)]),
        _row(5, [_inline_cell("A5", "Planned income", 4), _formula_cell("B5", "SUM('Monthly Plan'!B5:B16)", 5)]),
        _row(6, [_inline_cell("A6", "Planned expenses", 4), _formula_cell("B6", "SUM('Monthly Plan'!C5:C16)", 5)]),
        _row(7, [_inline_cell("A7", "Logged spending", 4), _formula_cell("B7", "SUM(Transactions!D5:D204)", 5)]),
        _row(8, [_inline_cell("A8", "Planned leftover", 4), _formula_cell("B8", "B5-B6", 5)]),
        _row(10, [_inline_cell("A10", "Start here", 3)]),
        _row(11, [_inline_cell("A11", "1. Set monthly income and expense targets on Monthly Plan.")]),
        _row(12, [_inline_cell("A12", "2. Log purchases or deposits on Transactions; use positive amounts for spending.")]),
        _row(13, [_inline_cell("A13", "3. Review the totals above before making a new commitment.")]),
        _row(15, [_inline_cell("A15", "Common category view", 3)]),
        _row(16, [_inline_cell("A16", "Category", 4), _inline_cell("B16", "Logged expense", 4)]),
    ]
    for row_number, category in enumerate(("Housing", "Food", "Transport", "Health", "Other"), 17):
        rows.append(_row(row_number, [
            _inline_cell(f"A{row_number}", category, 6),
            _formula_cell(f"B{row_number}", f'SUMIFS(Transactions!$D$5:$D$204,Transactions!$C$5:$C$204,A{row_number},Transactions!$E$5:$E$204,"Expense")', 7),
        ]))
    rows.append(_row(23, [_inline_cell("A23", "Workbook notes", 3)]))
    rows.append(_row(24, [_inline_cell("A24", "This template is educational and does not provide financial advice.")]))
    return _worksheet(rows, "A1:B24", {"1": 28, "2": 18}, "A4")


def _transactions() -> str:
    headers = ["Transaction date", "Description", "Category", "Amount", "Type", "Notes"]
    rows = [
        _row(1, [_inline_cell("A1", "Transactions", 1)], 28),
        _row(2, [_inline_cell("A2", "Enter one transaction per row. Keep amounts positive; choose Expense or Income in Type.", 2)]),
        _row(3, [_inline_cell("A3", "Tip: use a consistent category name so your monthly review stays meaningful.", 2)]),
        _row(4, [_inline_cell(_cell_ref(i, 4), header, 3) for i, header in enumerate(headers, 1)]),
    ]
    for row_number in range(5, MAX_TRANSACTIONS + 5):
        cells = [
            _inline_cell(f"A{row_number}", "", 6),
            _inline_cell(f"B{row_number}", "", 6),
            _inline_cell(f"C{row_number}", "", 6),
            _inline_cell(f"D{row_number}", "", 7),
            _inline_cell(f"E{row_number}", "", 6),
            _inline_cell(f"F{row_number}", "", 6),
        ]
        rows.append(_row(row_number, cells))
    return _worksheet(rows, "A1:F204", {"1": 16, "2": 28, "3": 20, "4": 14, "5": 14, "6": 32}, "A5")


def _monthly_plan(months: int) -> str:
    rows = [
        _row(1, [_inline_cell("A1", "Monthly Plan", 1)], 28),
        _row(2, [_inline_cell("A2", "Use a simple target for each month. Adjust it when life changes.", 2)]),
        _row(4, [_inline_cell("A4", "Month", 3), _inline_cell("B4", "Income target", 3), _inline_cell("C4", "Expense target", 3), _inline_cell("D4", "Remaining", 3), _inline_cell("E4", "Logged expense", 3), _inline_cell("F4", "Review note", 3)]),
    ]
    for index in range(months):
        row_number = index + 5
        month_number = (index % 12) + 1
        year = 2026 + index // 12
        label = date(year, month_number, 1).strftime("%b %Y")
        rows.append(_row(row_number, [
            _inline_cell(f"A{row_number}", label, 6),
            _inline_cell(f"B{row_number}", "", 7),
            _inline_cell(f"C{row_number}", "", 7),
            _formula_cell(f"D{row_number}", f"B{row_number}-C{row_number}", 7),
            _formula_cell(f"E{row_number}", f'SUMIFS(Transactions!$D$5:$D$204,Transactions!$A$5:$A$204,\">=\"&DATEVALUE(A{row_number}&\" 1\"),Transactions!$A$5:$A$204,\"<\"&EDATE(DATEVALUE(A{row_number}&\" 1\"),1),Transactions!$E$5:$E$204,\"Expense\")', 7),
            _inline_cell(f"F{row_number}", "", 6),
        ]))
    end_row = months + 4
    return _worksheet(rows, f"A1:F{end_row}", {"1": 16, "2": 18, "3": 18, "4": 16, "5": 16, "6": 36}, "A5").replace(
        '<autoFilter ref="A4:F204"/>', f'<autoFilter ref="A4:F{end_row}"/>'
    )


def _styles() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<numFmts count="1"><numFmt numFmtId="164" formatCode="$#,##0.00;[Red]-$#,##0.00"/></numFmts>'
        '<fonts count="3"><font><sz val="11"/><color theme="1"/><name val="Aptos"/></font>'
        '<font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Aptos Display"/></font>'
        '<font><b/><sz val="11"/><color rgb="FF234E52"/><name val="Aptos"/></font></fonts>'
        '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF234E52"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFE7F1EF"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left style="thin"><color rgb="FFD4E2DF"/></left><right style="thin"><color rgb="FFD4E2DF"/></right>'
        '<top style="thin"><color rgb="FFD4E2DF"/></top><bottom style="thin"><color rgb="FFD4E2DF"/></bottom><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="8">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1"/>'
        '<xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )


def _content_types() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>')


def _root_rels() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')


def _workbook() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<fileVersion appName="xl" lastEdited="7" lowestEdited="7"/> '
            '<workbookPr defaultThemeVersion="164011"/><bookViews><workbookView xWindow="240" yWindow="120" windowWidth="16000" windowHeight="9000"/></bookViews>'
            '<sheets><sheet name="Dashboard" sheetId="1" r:id="rId1"/><sheet name="Transactions" sheetId="2" r:id="rId2"/><sheet name="Monthly Plan" sheetId="3" r:id="rId3"/></sheets>'
            '<calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>')


def _workbook_rels() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
            '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>')


def create_workbook(output: str | Path, *, months: int = 12) -> Path:
    """Write a formatted budget workbook and return its path."""
    if months <= 0:
        raise ValueError("months must be positive")
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    parts = {
        "[Content_Types].xml": _content_types(),
        "_rels/.rels": _root_rels(),
        "xl/workbook.xml": _workbook(),
        "xl/_rels/workbook.xml.rels": _workbook_rels(),
        "xl/styles.xml": _styles(),
        "xl/worksheets/sheet1.xml": _dashboard(),
        "xl/worksheets/sheet2.xml": _transactions(),
        "xl/worksheets/sheet3.xml": _monthly_plan(months),
    }
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return output


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Create a personal budget XLSX template")
    parser.add_argument("output", type=Path)
    parser.add_argument("--months", type=int, default=12)
    args = parser.parse_args()
    create_workbook(args.output, months=args.months)
    print(args.output)
