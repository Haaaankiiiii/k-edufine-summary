from datetime import date
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
import os

import pytest

from core import (DocumentOptions, InputError, budget_text, date_text, decimal_value, demo_rows,
                  generate_document, item_lines, korean_number, period_text, read_workbook,
                  total_amount, validate_rows)


def xlsx_fixture(rows, sheet_name="품목내역"):
    """Tiny in-memory OOXML test input; no source/private workbook is committed."""
    xml_rows = []
    for i, row in enumerate(rows, 1):
        cells = []
        for j, value in enumerate(row):
            if value is None:
                continue
            ref = f"{chr(65 + j)}{i}"
            if isinstance(value, str) and value.startswith("="):
                cell = f'<c r="{ref}"><f>{escape(value[1:])}</f></c>'
            elif isinstance(value, (int, float)):
                cell = f'<c r="{ref}" t="n"><v>{value}</v></c>'
            else:
                cell = f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'
            cells.append(cell)
        xml_rows.append(f'<row r="{i}">{"".join(cells)}</row>')
    files = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:I{len(rows)}"/><sheetData>{"".join(xml_rows)}</sheetData></worksheet>',
    }
    buffer = BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as output:
        for path, content in files.items():
            output.writestr(path, content)
    return buffer.getvalue()


HEADERS = ["순번", "* 내용", "S2B물품번호", "규격", "수량", "단위", "예상단가", "* 예상금액", "수수료"]


@pytest.mark.parametrize("value,expected", [(0,"영"),(1,"일"),(10,"일십"),(100,"일백"),(10000,"일만"),
    (13560,"일만삼천오백육십"),(186000,"일십팔만육천"),(396000,"삼십구만육천"),
    (2600000,"이백육십만"),(100000001,"일억일"),(1000000000001,"일조일")])
def test_korean_money(value, expected):
    assert korean_number(value) == expected
    assert budget_text(Decimal(value), True).endswith(f"금{expected}원정)")


def test_dates():
    assert period_text(date(2026,3,3),date(2026,3,31)) == "2026. 3. 3.(화)~3. 31.(화)"
    assert period_text(date(2026,12,31), date(2027,1,1)).endswith("~2027. 1. 1.(금)")
    assert date_text(date(2026,9,11)) == "2026. 9. 11.(금)"
    with pytest.raises(InputError):
        period_text(date(2026,3,31),date(2026,3,3))


def test_official_item_sequence_and_large_lists():
    lines = item_lines([f"품목{i}" for i in range(85)])
    assert lines[0] == "  가. 품목별 내역"
    assert lines[-1] == "    85) 품목84"
    short = item_lines(["물품"] * 29)
    assert short[13].startswith("  하.")
    assert short[14].startswith("  거.")
    assert short[28].startswith("  고.")


def test_real_layout_blank_unit_no_guessing():
    data = xlsx_fixture([["품목내역"], [], HEADERS, ["1", "교육용 구독 2개월", None, None, 1, None, 300000, 300000, 0]])
    sheet = read_workbook(data)[0]
    assert sheet.header_row == 3
    checked = validate_rows(sheet.rows)
    assert not checked.errors
    assert checked.items[0].quantity == 1
    _, body, total = generate_document(checked.items, DocumentOptions("수업 준비", "교육용 서비스", "신청"))
    assert "300,000원 X 수량 1 = 300,000원" in body
    assert total == 300000
    assert "300,000원(금삼십만원)" in body
    assert "2개월" in body


def test_totals_and_invalid_rows_are_not_silently_dropped():
    data = xlsx_fixture([HEADERS, [1,"용지",None,None,2,"상자",10000,20000,0], [],
                         [2,None,None,None,1,"개",5000,5000,0], [None,"합계",None,None,None,None,None,25000,None]])
    sheet = read_workbook(data)[0]
    assert len(sheet.rows) == 2
    assert sheet.totals == [(5,"25000")]
    assert validate_rows(sheet.rows).errors


def test_uncached_formulas_require_manual_correction():
    data = xlsx_fixture([HEADERS, [1,"용지",None,None,2,"상자",10000,"=E2*G2",0]])
    sheet = read_workbook(data)[0]
    assert sheet.rows[0]["예상금액"] == "[수식 결과 없음]"
    assert validate_rows(sheet.rows).errors


@pytest.mark.parametrize("value", ["NaN", "Infinity", "1e8", "12,34", "-1", "삼만원", "#VALUE!", "0.1"])
def test_invalid_currency(value):
    with pytest.raises(InputError):
        decimal_value(value, "금액", whole=True)


def test_fees_and_mismatch_never_make_false_equation():
    rows = demo_rows()[:1]
    rows[0]["예상금액"] = "100000"
    rows[0]["수수료"] = "3000"
    result = validate_rows(rows)
    assert result.mismatches
    with pytest.raises(InputError):
        total_amount(result.items)
    with pytest.raises(InputError):
        generate_document(result.items, DocumentOptions("수업용", fee_mode="add"))
    _, body, total = generate_document(result.items, DocumentOptions("수업용", fee_mode="add", allow_mismatch=True))
    assert total == 103000
    assert "100,000원(예상금액 기준), 수수료 3,000원, 합계 103,000원" in body
    assert "24,000원 X 5" not in body
    assert total_amount(result.items,"included") == 100000
    assert total_amount(result.items,"exclude") == 100000


def test_fractional_qty_exact_and_missing_amount():
    rows = demo_rows()[:1]
    rows[0].update({"수량":"1.5", "예상단가":"1000", "예상금액":""})
    result = validate_rows(rows)
    assert not result.errors
    assert result.items[0].amount == 1500
    rows[0]["예상단가"] = "1000.5"
    assert validate_rows(rows).errors


def test_attachments_and_all_items():
    result = validate_rows(demo_rows())
    title, body, total = generate_document(result.items, DocumentOptions("교과 수업 지원", "수업용 물품", attachments="견적서 1부\n품목내역서 1부"))
    assert title == "수업용 물품 구입"
    assert body.startswith("수업용 물품을 아래와 같이 구입하고자 합니다.")
    assert total == 186000
    assert "3. 소요 예산: 금186,000원(금일십팔만육천원)" in body
    assert "  다. 클리어파일" in body
    assert body.endswith("붙임  1. 견적서 1부.\n      2. 품목내역서 1부.  끝.")
    assert body.count("끝.") == 1


def test_bad_file_and_headers():
    with pytest.raises(InputError):
        read_workbook(b"not excel")
    with pytest.raises(InputError):
        read_workbook(xlsx_fixture([["성명", "주소"], ["홍길동", "서울"]]))


@pytest.mark.skipif(not os.getenv("KEDUFINE_TEST_XLSX"), reason="개인 원본 파일은 저장소에 포함하지 않음")
def test_user_supplied_workbook():
    sheets = read_workbook(Path(os.environ["KEDUFINE_TEST_XLSX"]).read_bytes())
    checked = validate_rows(sheets[0].rows)
    assert not checked.errors
    assert checked.items
    assert total_amount(checked.items) >= 0
