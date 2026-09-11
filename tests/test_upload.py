from io import BytesIO
from pathlib import Path
import os
from unittest.mock import patch

import pytest
from streamlit.testing.v1 import AppTest
from core import budget_text, read_workbook, total_amount, validate_rows

APP = Path(__file__).resolve().parents[1] / "app.py"


@pytest.mark.skipif(not os.getenv("KEDUFINE_TEST_XLSX"), reason="사용자 원본은 로컬에서만 검사")
def test_actual_upload_through_app():
    data = Path(os.environ["KEDUFINE_TEST_XLSX"]).read_bytes()
    upload = BytesIO(data)
    items = validate_rows(read_workbook(data)[0].rows).items
    with patch("streamlit.file_uploader", return_value=upload):
        app = AppTest.from_file(str(APP), default_timeout=20).run()
        app.text_input(key="purpose").set_value("교과 수업 준비").run()
        assert not app.exception
        assert budget_text(total_amount(items)) in app.code[1].value
        assert all(item.name in app.code[1].value for item in items)
