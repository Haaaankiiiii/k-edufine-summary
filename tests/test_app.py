from pathlib import Path
from streamlit.testing.v1 import AppTest

APP = Path(__file__).resolve().parents[1] / "app.py"


def test_empty_then_demo_then_purpose_then_reset():
    app = AppTest.from_file(str(APP), default_timeout=20).run()
    assert not app.exception
    assert not app.code
    app.toggle[0].set_value(True).run()
    assert not app.exception
    assert not app.code
    app.text_input(key="purpose").set_value("2026학년도 교과 수업 운영").run()
    assert not app.exception
    assert "금186,000원(금일십팔만육천원)" in app.code[1].value
    assert "보드마카" in app.code[1].value
    app.text_area(key="manual_draft").set_value("사용자가 다듬은 문구").run()
    assert app.text_area(key="manual_draft").value == "사용자가 다듬은 문구"
    app.text_input(key="purpose").set_value("2학기 수업 운영").run()
    assert "2학기 수업 운영" in app.text_area(key="manual_draft").value
    app.toggle[0].set_value(False).run()
    assert not app.exception
    assert not app.code
