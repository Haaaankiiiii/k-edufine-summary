from __future__ import annotations

from datetime import date
from hashlib import sha256

import streamlit as st

from core import (
    FIELDS, DocumentOptions, InputError, SheetData, budget_text, clean, decimal_value,
    demo_rows, generate_document, number_text, period_text, read_workbook, total_amount, validate_rows,
)

st.set_page_config(page_title="품의 한 장 · K-에듀파인 개요", page_icon="📝", layout="wide")
st.markdown("""
<style>
.stApp {background:#f6f7fb;}
.block-container {max-width:1280px; padding-top:4.4rem; padding-bottom:4rem;}
h1,h2,h3 {letter-spacing:-.045em; color:#18233b;}
h1 {font-size:2.65rem!important; font-weight:800!important;}
h3 {font-size:1.18rem!important;}
.brand {color:#5261b5; font-size:.76rem; font-weight:800; letter-spacing:.14em; margin-bottom:.45rem;}
.intro {font-size:1.04rem; color:#626c7f; margin-top:-.3rem; margin-bottom:1.65rem;}
.steps {display:flex; gap:8px; margin-bottom:1.5rem; flex-wrap:wrap;}
.step {background:#e9ecf7; color:#44527d; border-radius:7px; padding:9px 15px; font-size:.84rem;}
[data-testid="stVerticalBlockBorderWrapper"]>div {background:white;}
[data-testid="stMetric"] {background:#f1f3fa; padding:12px 16px; border-radius:10px;}
[data-testid="stMetricLabel"] {color:#63708a;}
[data-testid="stMetricValue"] {font-size:1.55rem;}
textarea {line-height:1.8!important;}
[data-testid="stCode"] pre {font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif!important; font-size:14px!important; line-height:1.8!important;}
.footer {color:#81899a; font-size:.78rem; margin-top:1.3rem;}
@media(max-width:700px) {.block-container {padding:4rem 1.2rem 2rem;} h1 {font-size:2rem!important;}}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="brand">K-EDUFINE · 교사를 위한 문서 도구</div>', unsafe_allow_html=True)
st.title("품의 한 장")
st.markdown('<div class="intro">품목내역을 올리면, 산출 근거부터 한글 금액까지 한 번에.</div>', unsafe_allow_html=True)
st.markdown('<div class="steps"><span class="step">01 &nbsp; 엑셀 올리기</span><span class="step">02 &nbsp; 목적 입력하기</span><span class="step">03 &nbsp; 개요 복사하기</span></div>', unsafe_allow_html=True)

left, right = st.columns([1.02, 1.1], gap="large")
sheet = None
fingerprint = "empty"
with left:
    with st.container(border=True):
        st.subheader("01  품목내역 불러오기")
        uploaded = st.file_uploader("K-에듀파인에서 내려받은 품목내역", type=["xlsx"],
                                    help="품목내역 화면에서 엑셀을 다운로드한 뒤 올려 주세요. 10MB 이하, 5,000행까지 지원합니다.")
        demo = st.toggle("예시 품목으로 먼저 사용해 보기", value=False)
        if uploaded is not None:
            data = uploaded.getvalue()
            fingerprint = sha256(data).hexdigest()[:16]
            try:
                # Per-session memory only. No disk writes and no shared data cache.
                if st.session_state.get("upload_hash") != fingerprint:
                    st.session_state["parsed_sheets"] = read_workbook(data)
                    st.session_state["upload_hash"] = fingerprint
                sheets = st.session_state["parsed_sheets"]
                name = st.selectbox("품목이 있는 시트", [s.name for s in sheets]) if len(sheets) > 1 else sheets[0].name
                sheet = next(s for s in sheets if s.name == name)
                fingerprint += "_" + name
                st.caption(f"{sheet.name} · 제목 {sheet.header_row}행 · 품목 {len(sheet.rows):,}행을 읽었습니다.")
            except InputError as exc:
                st.error(str(exc))
            if demo:
                st.caption("현재 업로드한 파일을 사용하고 있습니다. 파일을 지우면 예시가 표시됩니다.")
        elif demo:
            sheet = SheetData("예시 품목", 3, demo_rows())
            fingerprint = "demo"
            st.caption("연습용 가상 데이터입니다. 실제 품의에는 다운로드한 엑셀을 올려 주세요.")
        else:
            st.session_state.pop("parsed_sheets", None)
            st.session_state.pop("upload_hash", None)
            st.caption("내용 · 규격 · 수량 · 단위 · 예상단가 · 예상금액 · 수수료를 자동으로 찾습니다.")

    with st.container(border=True):
        st.subheader("02  품의 내용 입력하기")
        purpose = st.text_input("목적 *", placeholder="예: 2026학년도 2학기 교과 수업 운영에 필요한 소모품 확보", key="purpose")
        subject = st.text_input("구입·지급 대상 *", value="수업용 물품", help="제목과 첫 문장에 사용합니다. 예: 교무실 소모품, 교육용 소프트웨어 이용료")
        action = st.selectbox("품의 유형", ["구입", "지급", "신청"])
        with st.expander("기간·장소·관련 문서 추가"):
            related = st.text_input("관련 문서", placeholder="예: 교무부-1234(2026. 3. 2.)")
            use_period = st.checkbox("기간 넣기")
            a, b = st.columns(2)
            start = a.date_input("시작일", value=date.today(), disabled=not use_period)
            end = b.date_input("종료일", value=date.today(), disabled=not use_period)
            place = st.text_input("사용 장소", placeholder="예: 2학년 교실")
            account = st.text_input("예산 과목", placeholder="입력한 경우에만 포함합니다.")
        attachments = st.text_area("붙임", placeholder="견적서 1부\n품목내역서 1부", height=90,
                                   help="실제로 첨부할 문서명과 수량을 한 줄에 하나씩 입력하세요. 번호·마침표·끝.은 자동으로 붙습니다. 없으면 비워 둡니다.")
        with st.expander("문서 표기 설정"):
            show_spec = st.checkbox("규격 함께 표기", value=True)
            all_names = st.checkbox("품명 항목에 모든 품목 나열", value=False,
                                    help="기본값은 첫 품목 외 N건입니다. 산출 근거에는 항상 모든 품목을 표시합니다.")
            suffix = st.checkbox("한글 금액 끝에 '정' 붙이기", value=False,
                                 help="첨부 공문서 작성법은 '정'을 생략합니다. 기관 관행에 따라 선택할 수 있습니다.")

validation = None
fee_mode = "none"
allow_mismatch = False
period = ""
form_errors = []
if use_period:
    try:
        period = period_text(start, end)
    except InputError as exc:
        form_errors.append(str(exc))

# The editor is below both columns; reserve the right side for the live output.
with st.container(border=True):
    st.subheader("품목 확인·수정")
    if sheet is None:
        st.caption("엑셀을 올리면 이곳에 품목이 표시됩니다.")
    else:
        st.caption("셀을 두 번 눌러 수정할 수 있습니다. 금액 수정 후에는 Enter를 눌러 반영해 주세요. 단위가 없는 값은 임의로 '개'로 바꾸지 않습니다.")
        edited = st.data_editor(sheet.rows, key=f"items_{fingerprint}", hide_index=True,
                                num_rows="fixed", width="stretch", disabled=["원본 행"],
                                column_order=["원본 행", *FIELDS],
                                column_config={key: st.column_config.TextColumn(key, width="medium" if key in {"내용", "규격"} else "small") for key in ["원본 행", *FIELDS]})
        validation = validate_rows(edited)
        if sheet.notes or validation.notes:
            with st.expander(f"읽기 안내 {len(sheet.notes) + len(validation.notes)}건", expanded=True):
                for note in sheet.notes + validation.notes:
                    st.text(note)
        if validation.errors:
            for error in validation.errors:
                st.error(error)
        if validation.mismatches:
            st.warning("단가 × 수량과 예상금액이 다른 품목이 있습니다.")
            for message in validation.mismatches:
                st.text(message)
            mismatch_signature = sha256("\n".join(validation.mismatches).encode()).hexdigest()[:16]
            allow_mismatch = st.checkbox("차이를 확인했습니다. 해당 품목은 계산식 없이 예상금액만 표기합니다.", key=f"mismatch_{fingerprint}_{mismatch_signature}")
        if any(item.fee for item in validation.items):
            mode = st.radio("수수료를 어떻게 반영할까요?", ["예상금액에 별도 가산", "예상금액에 이미 포함", "이번 품의에서 제외"], index=None,
                            key=f"fees_{fingerprint}", help="엑셀의 수수료가 예상금액에 포함되는지는 파일만으로 확정할 수 없어 직접 선택합니다.")
            fee_mode = {"예상금액에 별도 가산": "add", "예상금액에 이미 포함": "included", "이번 품의에서 제외": "exclude"}.get(mode, "none")
        if sheet.totals and not validation.errors:
            item_sum = sum(item.amount for item in validation.items)
            for row, value in sheet.totals:
                try:
                    original = decimal_value(value, f"원본 {row}행 합계", whole=True)
                    if original != item_sum:
                        st.warning(f"엑셀 {row}행 합계 {number_text(original)}원과 현재 품목 예상금액 합계 {number_text(item_sum)}원이 다릅니다. 현재 품목 합계를 사용합니다.")
                except InputError as exc:
                    st.warning(str(exc))

with right:
    with st.container(border=True):
        st.subheader("03  완성된 품의 개요")
        ready = sheet is not None and validation is not None and not validation.errors
        if not ready:
            st.info("품목내역을 올리고 목적을 입력하면 개요가 자동으로 완성됩니다.")
            st.markdown("**자동으로 작성되는 내용**\n\n- 품명과 규격\n- 모든 품목의 산출 근거\n- 총 소요 예산과 한글 금액\n- 항목 번호·들여쓰기·붙임·끝 표시")
        else:
            try:
                total = total_amount(validation.items, fee_mode)
                c1, c2 = st.columns(2)
                c1.metric("품목", f"{len(validation.items):,}건")
                c2.metric("소요 예산", f"{number_text(total)}원")
                st.caption(budget_text(total, suffix))
                if form_errors:
                    raise InputError(form_errors[0])
                if not purpose.strip():
                    raise InputError("왼쪽에 구입·지급 목적을 입력해 주세요.")
                options = DocumentOptions(purpose, subject, action, related, period, place, account, attachments,
                                          show_spec, all_names, suffix, fee_mode, allow_mismatch)
                title, body, _ = generate_document(validation.items, options)
                st.caption("제목")
                st.code(title, language=None)
                st.caption("개요 · 오른쪽 위 복사 버튼을 눌러 K-에듀파인에 붙여 넣으세요.")
                st.code(body, language=None, wrap_lines=True)
                st.download_button("개요 텍스트 내려받기", data=body.encode("utf-8-sig"),
                                   file_name="품의_개요.txt", mime="text/plain", width="stretch")
                with st.expander("문구 직접 다듬기"):
                    st.caption("아래 수정 내용은 품목 계산에 반영되지 않습니다. 위 입력값을 바꾸면 새 초안으로 초기화됩니다.")
                    signature = sha256(body.encode()).hexdigest()
                    if st.session_state.get("draft_signature") != signature:
                        st.session_state["manual_draft"] = body
                        st.session_state["draft_signature"] = signature
                    manual = st.text_area("최종 문구", height=340, key="manual_draft")
                    st.code(manual, language=None, wrap_lines=True)
                    st.download_button("수정한 개요 내려받기", data=manual.encode("utf-8-sig"),
                                       file_name="품의_개요_수정본.txt", mime="text/plain")
            except InputError as exc:
                st.info(str(exc))

with st.expander("사용 방법과 적용한 문서 형식"):
    st.markdown("""
1. K-에듀파인 품목내역에 물품을 입력하고 **엑셀을 다운로드**합니다.
2. 위에 파일을 올리고 **구입·지급 목적**을 입력합니다.
3. 품명·단위·금액을 확인한 뒤 **개요 복사 버튼**을 누릅니다.
4. K-에듀파인 개요 칸에 붙여 넣고 실제 붙임 문서를 첨부합니다.

첨부하신 「공문서 바로쓰기」의 지출 품의서(인쇄 쪽 3), 항목 구분(쪽 5), 날짜·금액(쪽 6)을 참고했습니다.
기본 순서는 목적 → 품명 → 소요 예산 → 산출 근거입니다. 하위 항목은 두 칸 들여쓰고, `하.` 다음에는 `거.`로 이어집니다.
84건을 넘으면 `가. 품목별 내역` 아래에 `1)`, `2)` 형식으로 모든 품목을 나열합니다.
품목을 자동으로 합치거나 인건비의 사람·시간 정보를 추측하지 않습니다. 인건비도 엑셀의 각 행을 그대로 산출 근거로 작성합니다.

앱은 업로드 파일을 서버 메모리에서 처리하며 파일·품의 내용을 별도 저장하거나 외부 AI로 보내지 않습니다.
온라인 배포 시 업로드 파일은 Streamlit 서버로 전송됩니다. 연결이 끝난 뒤의 메모리 정리는 호스팅 환경에 따릅니다.
K-에듀파인 로그인 정보나 API 키는 필요하지 않습니다.
""")
st.markdown('<div class="footer">품의 한 장 · Excel → 개요 · K-에듀파인 공식 서비스가 아닌 작성 보조 도구입니다.</div>', unsafe_allow_html=True)
