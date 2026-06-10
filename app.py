import os
import re
import csv
import io
import time
import json
import base64
import shutil
import zipfile
import calendar
from datetime import date, datetime, timedelta
from urllib.parse import quote
from typing import Dict, List, Any, Optional, Union 

from dotenv import load_dotenv
load_dotenv()

# 특정 네트워크에서 gRPC 통신 타임아웃 방지
os.environ['GRPC_DNS_RESOLVER'] = 'native'

import streamlit as st
import pandas as pd

# ==========================================
# 🛑 HWP -> PDF 변환용 라이브러리 (윈도우 전용)
# ==========================================
try:
    import win32com.client as win32
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

def hwp_to_pdf(hwp_path: str) -> str:
    if not WIN32_AVAILABLE:
        st.error("pywin32 라이브러리가 설치되지 않아 PDF 변환을 수행할 수 없습니다.")
        return hwp_path
        
    pdf_path = hwp_path[:-4] + ".pdf"
    abs_hwp = os.path.abspath(hwp_path)
    abs_pdf = os.path.abspath(pdf_path)
    
    try:
        pythoncom.CoInitialize() 
        hwp = win32.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(abs_hwp)
        hwp.HAction.Run("FilePrint")
        hwp.HAction.Run("FileSaveAsPdf") 
        hwp.Quit()
        return pdf_path
    except Exception as e: 
        print(f"HWP->PDF 변환 에러: {e}")
        return hwp_path 
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

# ==========================================
# ⚙️ 1. 기본 설정 및 전역 변수
# ==========================================
st.set_page_config(page_title="현장점검 통합관리 시스템", page_icon="🏛️", layout="wide")

SHARED_USER_ID = os.environ.get("ADMIN_ID", st.secrets.get("ADMIN_ID", "molitdj_default"))
SHARED_PASSWORD = os.environ.get("ADMIN_PW", st.secrets.get("ADMIN_PW", "change_me!"))

# AI 요약 기능용 라이브러리 로드 시도
try:
    from google import genai
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", st.secrets.get("GEMINI_API_KEY", ""))
except ImportError:
    GEMINI_API_KEY = ""

DB_FILENAME = "penalty_database.csv"
ATTACH_DIR = "attachments"
ITEMS_PER_PAGE = 10 

if not os.path.exists(ATTACH_DIR):
    os.makedirs(ATTACH_DIR)

PENALTY_INTERVALS = [
    (30, "확인서 이의제기 접수"), (14, "확인서 이의제기 의견 통보"), (14, "벌점 사전부과 통보"),
    (15, "벌점 사전부과 통보 의견제출 마감"), (15, "벌점 사전부과 의견 검토회의"), 
    (1, "벌점 사전부과 의견 검토회의 결과 통보 및 벌점 부과"), (30, "벌점 부과 이의제기 접수"), 
    (40, "벌점 심의위원회 개최 및 최종 결과 통보")
]

# ==========================================
# 🛡️ 유틸리티 함수 (보안 및 데이터 안정성)
# ==========================================
def secure_filename(filename: str) -> str:
    filename = os.path.basename(filename)
    return re.sub(r'[^a-zA-Z0-9가-힣_\-\.]', '_', filename)

def atomic_save_csv(filename: str, headers: List[str], rows: List[List[Any]]) -> None:
    temp_filename = f"{filename}.tmp"
    try:
        with open(temp_filename, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)
        os.replace(temp_filename, filename)
    except Exception as e:
        st.error(f"데이터 안전 저장 실패: {e}")
        if os.path.exists(temp_filename):
            os.remove(temp_filename)

# ==========================================
# 🖱️ 2. 새창(다이얼로그) 기능
# ==========================================
def make_dialog_draggable():
    import streamlit.components.v1 as components
    drag_js = """
    <script>
    const doc = window.parent.document;
    const setupDrag = () => {
        const dialogs = doc.querySelectorAll('div[data-testid="stDialog"]');
        dialogs.forEach(dialog => {
            const header = dialog.querySelector('header');
            if (header && !dialog.dataset.dragEnabled) {
                dialog.dataset.dragEnabled = 'true';
                header.style.cursor = 'grab';
                let isDragging = false; let startX, startY; let currentX = 0; let currentY = 0;
                dialog.style.position = 'relative';
                header.addEventListener('mousedown', (e) => { isDragging = true; header.style.cursor = 'grabbing'; startX = e.clientX; startY = e.clientY; doc.body.style.userSelect = 'none'; });
                doc.addEventListener('mousemove', (e) => { if (!isDragging) return; const dx = e.clientX - startX; const dy = e.clientY - startY; currentX += dx; currentY += dy; dialog.style.left = currentX + 'px'; dialog.style.top = currentY + 'px'; startX = e.clientX; startY = e.clientY; });
                const stopDrag = () => { if (isDragging) { isDragging = false; header.style.cursor = 'grab'; doc.body.style.userSelect = ''; } };
                doc.addEventListener('mouseup', stopDrag); doc.addEventListener('mouseleave', stopDrag);
            }
        });
    };
    const observer = new MutationObserver(setupDrag); observer.observe(doc.body, { childList: true, subtree: true }); setTimeout(setupDrag, 100);
    </script>
    """
    components.html(drag_js, height=0, width=0)

# 💡 [수정됨] 요약 결과를 기억(캐싱)하여 창 튕김 및 무한 반복을 방지하는 스마트 AI 대화창
@st.dialog("✨ AI 심의안건 보고서 작성 (새창)", width="large")
def show_summary_dialog(file_path, file_name):
    make_dialog_draggable() 
    
    # 실수로 창 바깥을 눌러서 닫히는 현상을 안내
    st.warning("💡 **안내:** 요약이 진행되는 동안 **어두운 배경(창 바깥쪽)을 클릭하면 창이 강제로 닫힙니다.** 완료될 때까지 기다려주세요.")
    st.markdown(f"### 📄 [{file_name}] 분석 결과")
    
    if not GEMINI_API_KEY:
        st.error("⚠️ 시스템에 AI API 키가 설정되어 있지 않습니다.")
        return

    # 파일 이름을 기반으로 고유한 메모리(세션) 방 이름 만들기
    state_key = f"ai_summary_{file_name}"
    
    # 1. 만약 이 파일을 요약한 기억(메모리)이 없다면 -> 새로 요약 시작
    if state_key not in st.session_state:
        with st.spinner("AI가 공무원 양식으로 보고서를 작성 중입니다..."):
            # 요약 결과를 받아서 바로 출력함과 동시에 메모리에 저장
            summary_result = st.write_stream(get_ai_summary_stream(file_path))
            st.session_state[state_key] = summary_result
            
    # 2. 이미 요약해 둔 기억이 있다면 -> 즉시 메모리에서 불러와서 보여줌
    else:
        st.markdown(st.session_state[state_key])

    st.divider()
    
    # 하단 버튼 배치 (닫기 / 다시 요약하기)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        # 이제 메모리에 결과가 있으므로, 이 버튼을 눌러도 다시 요약하지 않고 0.1초만에 즉시 창이 닫힙니다.
        if st.button("닫기", type="primary", use_container_width=True):
            st.rerun()
    with c3:
        # 혹시 처음부터 새로 요약하고 싶을 때를 대비한 버튼 (메모리를 지우고 다시 시작)
        if st.button("🔄 다시 요약하기", use_container_width=True):
            del st.session_state[state_key]
            st.rerun()

@st.dialog("📄 첨부 문서 뷰어 (새창)", width="large")
def show_file_dialog(file_path, file_name):
    make_dialog_draggable() 
    st.markdown(f"### 📎 {file_name}")
    ext = os.path.splitext(file_path)[1].lower()
    try: 
        if ext in ['.png', '.jpg', '.jpeg']: st.image(file_path, use_column_width=True)
        elif ext == '.pdf':
            with open(file_path, "rb") as f: base64_pdf = base64.b64encode(f.read()).decode('utf-8')
            st.markdown(f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700" type="application/pdf"></iframe>', unsafe_allow_html=True)
        elif ext == '.txt':
            with open(file_path, "r", encoding="utf-8") as f: st.text_area("문서 내용", f.read(), height=500)
        else: st.warning("⚠️ 지원하지 않는 형식입니다. 체크박스를 통해 다운로드해 주세요.")
    except Exception as e: st.error(f"오류가 발생했습니다: {e}")

# ==========================================
# 📌 현장 상세정보 공통 처리
# ==========================================
SITE_DETAIL_FIELDS = [
    ("site_office", "현장사무실", ["현장사무실", "현장 사무실", "현장사무소", "현장 사무소", "현장사무실 주소", "사무실주소"]),
    ("postal_address", "별도 우편 주소", ["별도 우편 주소", "별도우편주소", "우편주소", "우편 주소", "별도주소", "주소"]),
    ("construction_period", "착공일~준공일", ["착공일~준공일", "착공일 ~ 준공일", "착공일-준공일", "착공일 - 준공일", "착공/준공", "착공 및 준공", "공사기간", "공사 기간", "공사기간(착공~준공)", "공사기간(착공일~준공일)"]),
    ("construction_start", "착공일", ["착공일", "착공 일자", "공사시작일", "공사 시작일", "착수일"]),
    ("construction_end", "준공일", ["준공일", "준공 일자", "공사종료일", "공사 종료일", "완료일"]),
    ("total_cost", "총공사비", ["총공사비", "총 공사비", "공사비", "도급액", "계약금액", "총사업비"]),
    ("progress_rate", "공정률", ["공정률", "공정율", "진도율", "공사진행률", "공사진행율"]),
    ("builder", "시공사", ["시공사", "시공 회사", "시공회사", "시공회사명", "시공자", "시공자명", "건설사", "시공업체"]),
    ("supervisor", "감리사", ["감리사", "감리 회사", "감리회사", "감리회사명", "감리자", "감리자명", "감리업체"]),
    ("site_manager", "현장대리인", ["현장대리인", "현장 대리인", "현장대리인 성명", "현장대리인명", "대리인", "성명"]),
    ("manager_phone", "현장대리인 전화번호", ["현장대리인 전화번호", "현장대리인 연락처", "현장대리인 휴대폰", "현장대리인 핸드폰", "대리인 전화번호", "대리인 연락처", "전화번호", "연락처", "휴대폰", "핸드폰"]),
    ("manager_email", "현장대리인 이메일", ["현장대리인 이메일", "현장대리인 메일", "현장대리인 E-mail", "현장대리인 email", "대리인 이메일", "이메일", "메일", "E-mail", "Email", "email"]),
    ("client", "발주처", ["발주처\n(인·허가 기관)", "발주처(인·허가 기관)", "발주처", "인허가기관", "인·허가 기관", "인허가 기관"]),
    ("status", "공사진행상태", ["공사진행상태", "공사 진행 상태", "진행상태", "공사상태", "현장상태"]),
]

STEP_EXTRA_FIELDS = [
    ("inspection_period", "점검시기", ["점검시기", "점검 시기", "점검구분", "점검 구분", "점검유형", "점검 유형", "점검명"]),
    ("team", "담당조", ["담당조", "점검조", "조", "반", "담당반"]),
    ("inspectors", "점검자", ["점검자", "점검자명", "점검자 명", "담당자", "검사자", "참석자"]),
]

MONTHLY_INSPECTION_OPTIONS = [f"{m}월 상시점검" for m in range(1, 13)]
INSPECTION_PERIOD_OPTIONS = ["해빙기 점검", "우기대비 점검", "동절기 점검"] + MONTHLY_INSPECTION_OPTIONS + ["기타"]

DB_COLUMNS = (
    ['No', '현장명', '날짜', '업무명', '메모', '파일경로']
    + [label for _, label, _ in STEP_EXTRA_FIELDS]
    + [label for _, label, _ in SITE_DETAIL_FIELDS]
)

def clean_cell(value: Any) -> str:
    if value is None: return ""
    try:
        if pd.isna(value): return ""
    except Exception: pass
    if isinstance(value, pd.Timestamp) or isinstance(value, datetime) or isinstance(value, date):
        return value.strftime('%Y-%m-%d')
    value_str = str(value).strip()
    if value_str.lower() in ["nan", "nat", "none", "null"]: return ""
    return value_str

def normalize_column_name(value: Any) -> str: return re.sub(r"\s+", "", str(value).replace("\n", "")).lower()

def get_row_value(row: pd.Series, aliases: List[str]) -> str:
    normalized_row = {normalize_column_name(col): val for col, val in row.items()}
    for alias in aliases:
        cleaned = clean_cell(normalized_row.get(normalize_column_name(alias), ""))
        if cleaned: return cleaned
    return ""

def parse_date_value(value: Any, default_year: Optional[int] = None) -> Optional[date]:
    if value is None: return None
    try:
        if pd.isna(value): return None
    except Exception: pass
    if isinstance(value, pd.Timestamp) or isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value

    value_str = clean_cell(value)
    if not value_str: return None
    numbers = re.findall(r"\d+", value_str)
    try:
        if len(numbers) >= 3:
            year = int(numbers[0])
            if year < 100: year += 2000
            return date(year, int(numbers[1]), int(numbers[2]))
        if len(numbers) >= 2:
            return date(default_year or date.today().year, int(numbers[0]), int(numbers[1]))
    except Exception: return None
    return None

def infer_inspection_period(file_name: str = "", row: Optional[pd.Series] = None, plan_date: Optional[date] = None, desc: str = "") -> str:
    candidates = []
    if row is not None:
        try: candidates.append(get_row_value(row, ["점검시기", "점검 시기", "점검구분", "점검 구분", "점검유형", "점검 유형", "점검명"]))
        except Exception: pass
    candidates.extend([file_name, desc])
    source = " ".join(clean_cell(v) for v in candidates if clean_cell(v))

    if "해빙" in source: return "해빙기 점검"
    if "우기" in source: return "우기대비 점검"
    if "동절" in source or "겨울" in source: return "동절기 점검"
    if "상시" in source or "월점검" in source or "월 점검" in source:
        return f"{plan_date.month}월 상시점검" if plan_date else "상시점검"

    month_match = re.search(r"(1[0-2]|[1-9])\s*월", source)
    if month_match: return f"{int(month_match.group(1))}월 상시점검"
    return "기타"

def selectbox_options_with_current(current_value: str) -> List[str]:
    current_value = clean_cell(current_value)
    options = INSPECTION_PERIOD_OPTIONS.copy()
    if current_value and current_value not in options: options.insert(0, current_value)
    return options

def extract_team_from_desc(desc: str) -> str:
    match = re.search(r"\[([^\]]*조)\]", clean_cell(desc))
    return match.group(1).strip() if match else ""

def strip_wrapping_brackets(value: str) -> str:
    value = clean_cell(value)
    return value[1:-1].strip() if value.startswith("[") and value.endswith("]") else value

def normalize_inspection_period_label(period: str) -> str:
    period = strip_wrapping_brackets(period)
    if not period or period == "기타": return ""
    if "우기" in period: return "우기"
    if "해빙" in period: return "해빙기"
    if "동절" in period or "겨울" in period: return "동절기"
    month_match = re.search(r"(1[0-2]|[1-9])\s*월", period)
    if month_match and "상시" in period: return f"{int(month_match.group(1))}월 상시"
    if "상시" in period: return "상시"
    return period.replace(" 점검", "").replace("점검", "").strip()

def truncate_label(label: str, max_chars: int = 34) -> str:
    label = clean_cell(label)
    return label[:max_chars - 3].rstrip() + "..." if max_chars and len(label) > max_chars else label

def make_calendar_event_label(site: str, step: dict) -> str:
    desc = clean_cell(step.get("desc", ""))
    period = clean_cell(step.get("inspection_period", "")) or infer_inspection_period(plan_date=step.get("date"), desc=desc)
    team = clean_cell(step.get("team", "")) or extract_team_from_desc(desc)
    prefix = ""
    if normalize_inspection_period_label(period): prefix += f"[{normalize_inspection_period_label(period)}]"
    if strip_wrapping_brackets(team): prefix += f"[{strip_wrapping_brackets(team)}]"
    return f"{prefix}{clean_cell(site)}" if prefix else clean_cell(site)

def get_representative_step_for_site(steps: List[dict]) -> dict:
    if not steps: return {}
    dated_steps = [step for step in steps if step.get("date")]
    if not dated_steps: return steps[0]
    upcoming = [step for step in dated_steps if step.get("date") >= date.today()]
    return min(upcoming, key=lambda x: x.get("date")) if upcoming else max(dated_steps, key=lambda x: x.get("date"))

def make_site_list_label(site: str, site_data: dict, max_chars: int = 34) -> str:
    if site == "전체 현장": return site
    return truncate_label(make_calendar_event_label(site, get_representative_step_for_site(site_data.get(site, []))), max_chars=max_chars)

def normalize_team_for_sort(team_value: str) -> str: return re.sub(r"\s+", "", strip_wrapping_brackets(team_value)).upper()

def get_team_sort_rank(team_value: str) -> int:
    normalized = normalize_team_for_sort(team_value)
    if re.search(r"^TF0*1조?$", normalized): return 4
    if re.search(r"^TF0*2조?$", normalized): return 5
    plain_match = re.search(r"^(?:제)?0*([1-3])조?$", normalized)
    if plain_match: return int(plain_match.group(1))
    number_match = re.search(r"(\d+)", normalized)
    return 100 + int(number_match.group(1)) if number_match else 999

def get_step_team_value(step: dict) -> str: return clean_cell(step.get("team", "")) or extract_team_from_desc(step.get("desc", ""))

def calendar_event_sort_key(event_tuple: tuple) -> tuple:
    site, step_idx, step = event_tuple
    return (get_team_sort_rank(get_step_team_value(step)), normalize_team_for_sort(get_step_team_value(step)), clean_cell(site), step_idx)

def get_site_detail_defaults(steps: List[dict]) -> dict:
    defaults = {key: "" for key, _, _ in SITE_DETAIL_FIELDS}
    for step in steps:
        for key, _, _ in SITE_DETAIL_FIELDS:
            if not defaults[key] and clean_cell(step.get(key, "")): defaults[key] = clean_cell(step.get(key, ""))
    return defaults

def apply_site_details_to_all_steps(site_name: str, detail_values: dict) -> None:
    for step in st.session_state.site_data.get(site_name, []):
        for key, value in detail_values.items(): step[key] = value

def get_construction_period_text(step: dict) -> str:
    explicit = clean_cell(step.get("construction_period", ""))
    if explicit: return explicit
    start, end = clean_cell(step.get("construction_start", "")), clean_cell(step.get("construction_end", ""))
    return f"{start or '-'} ~ {end or '-'}" if start or end else ""

def format_site_detail_for_popup(step: dict) -> str:
    lines = [f"점검일정: {step.get('date').strftime('%Y-%m-%d') if step.get('date') else '-'}"]
    for label, value in [("점검시기", step.get("inspection_period", "")), ("담당조", step.get("team", "")), ("점검자", step.get("inspectors", ""))]:
        if clean_cell(value): lines.append(f"{label}: {clean_cell(value)}")
    detail_pairs = [("현장사무실", step.get("site_office", "")), ("별도 우편 주소", step.get("postal_address", "")), ("착공일~준공일", get_construction_period_text(step)), ("총공사비", step.get("total_cost", "")), ("공정률", step.get("progress_rate", "")), ("시공사", step.get("builder", "")), ("감리사", step.get("supervisor", "")), ("현장대리인", step.get("site_manager", "")), ("현장대리인 전화번호", step.get("manager_phone", "")), ("현장대리인 이메일", step.get("manager_email", "")), ("발주처", step.get("client", "")), ("공사진행상태", step.get("status", ""))]
    has_detail = False
    for label, value in detail_pairs:
        if clean_cell(value):
            lines.append(f"{label}: {clean_cell(value)}")
            has_detail = True
    if not has_detail: lines.append("등록된 현장 상세정보가 없습니다.")
    return "\n".join(lines)

def render_site_detail_inputs(defaults: dict, key_prefix: str) -> dict:
    detail_values = {}
    d1, d2 = st.columns(2)
    with d1:
        detail_values["site_office"] = st.text_input("현장사무실", value=defaults.get("site_office", ""), key=f"{key_prefix}_site_office")
        detail_values["construction_period"] = st.text_input("착공일~준공일", value=defaults.get("construction_period", ""), placeholder="예: 2026-03-01 ~ 2027-12-31", key=f"{key_prefix}_construction_period")
        detail_values["construction_start"] = st.text_input("착공일", value=defaults.get("construction_start", ""), placeholder="예: 2026-03-01", key=f"{key_prefix}_construction_start")
        detail_values["total_cost"] = st.text_input("총공사비", value=defaults.get("total_cost", ""), placeholder="예: 120억 원", key=f"{key_prefix}_total_cost")
        detail_values["builder"] = st.text_input("시공사", value=defaults.get("builder", ""), key=f"{key_prefix}_builder")
        detail_values["site_manager"] = st.text_input("현장대리인", value=defaults.get("site_manager", ""), key=f"{key_prefix}_site_manager")
        detail_values["manager_phone"] = st.text_input("현장대리인 전화번호", value=defaults.get("manager_phone", ""), key=f"{key_prefix}_manager_phone")
    with d2:
        detail_values["postal_address"] = st.text_input("별도 우편 주소", value=defaults.get("postal_address", ""), key=f"{key_prefix}_postal_address")
        detail_values["construction_end"] = st.text_input("준공일", value=defaults.get("construction_end", ""), placeholder="예: 2027-12-31", key=f"{key_prefix}_construction_end")
        detail_values["progress_rate"] = st.text_input("공정률", value=defaults.get("progress_rate", ""), placeholder="예: 42%", key=f"{key_prefix}_progress_rate")
        detail_values["supervisor"] = st.text_input("감리사", value=defaults.get("supervisor", ""), key=f"{key_prefix}_supervisor")
        detail_values["manager_email"] = st.text_input("현장대리인 이메일", value=defaults.get("manager_email", ""), key=f"{key_prefix}_manager_email")
        detail_values["client"] = st.text_input("발주처", value=defaults.get("client", ""), key=f"{key_prefix}_client")
        detail_values["status"] = st.text_input("공사진행상태", value=defaults.get("status", ""), key=f"{key_prefix}_status")
    return detail_values

@st.dialog("🛠️ 현장정보 및 점검일정 수정", width="large")
def show_schedule_edit_dialog(site_name: str, step_idx: int):
    make_dialog_draggable()
    try: step_idx = int(step_idx)
    except ValueError: return st.error("수정할 일정 정보를 찾을 수 없습니다.")
    if site_name not in st.session_state.site_data or step_idx < 0 or step_idx >= len(st.session_state.site_data[site_name]):
        return st.error("선택한 현장 또는 일정이 존재하지 않습니다.")

    steps = st.session_state.site_data[site_name]
    step = steps[step_idx]
    st.markdown(f"### {site_name}")
    st.info(format_site_detail_for_popup(step))

    with st.form(f"calendar_edit_form_{site_name}_{step_idx}"):
        c1, c2 = st.columns([1, 2])
        with c1: new_date = st.date_input("점검 예정일", value=step.get("date", date.today()))
        with c2:
            current_period = clean_cell(step.get("inspection_period", "")) or infer_inspection_period(plan_date=step.get("date"), desc=step.get("desc", ""))
            period_options = selectbox_options_with_current(current_period)
            new_period = st.selectbox("점검시기", period_options, index=period_options.index(current_period) if current_period in period_options else 0)

        c3, c4 = st.columns([1, 2])
        with c3: new_team = st.text_input("담당조", value=clean_cell(step.get("team", "")) or extract_team_from_desc(step.get("desc", "")), placeholder="예: 1조")
        with c4: new_desc = st.text_input("점검/업무명", value=clean_cell(step.get("desc", "")))

        new_inspectors = st.text_area("점검자", value=clean_cell(step.get("inspectors", "")), height=80)
        new_memo = st.text_area("메모", value=clean_cell(step.get("memo", "")), height=100)

        st.markdown("#### 현장 상세정보 수정")
        detail_values = render_site_detail_inputs(get_site_detail_defaults(steps), f"calendar_detail_{site_name}_{step_idx}")
        save_btn, close_btn = st.columns(2)
        with save_btn: submitted = st.form_submit_button("변경사항 저장", type="primary", use_container_width=True)
        with close_btn: closed = st.form_submit_button("닫기", use_container_width=True)

    if submitted:
        steps[step_idx].update({"date": new_date, "inspection_period": new_period, "team": new_team, "inspectors": new_inspectors, "desc": new_desc, "memo": new_memo})
        apply_site_details_to_all_steps(site_name, detail_values)
        steps.sort(key=lambda x: x['date'])
        save_data(st.session_state.site_data)
        st.success("저장되었습니다.")
        time.sleep(0.5)
        st.rerun()

    if closed: st.rerun()

# ==========================================
# 📅 4. Streamlit 네이티브 달력 렌더링 (순수 CSS 완벽 고정)
# ==========================================
def make_streamlit_key(*parts) -> str:
    raw = "_".join(clean_cell(part) for part in parts)
    return re.sub(r"[^0-9a-zA-Z가-힣_]+", "_", raw)[:180]

def get_team_style_class(teamRaw: str) -> str:
    if not teamRaw: return "default"
    team = normalize_team_for_sort(teamRaw)
    if re.search(r"^TF0*1조?$", team): return "tf1"
    if re.search(r"^TF0*2조?$", team): return "tf2"
    plain = re.search(r"^(?:제)?0*([1-3])조?$", team)
    if plain: return f"team{plain.group(1)}"
    return "default"

def inject_pure_css_calendar_style():
    CSS = """
    <style>
    div[data-testid="column"]:has(.cal-cell-marker)::-webkit-scrollbar { width: 4px; height: 4px; }
    div[data-testid="column"]:has(.cal-cell-marker)::-webkit-scrollbar-track { background: transparent; }
    div[data-testid="column"]:has(.cal-cell-marker)::-webkit-scrollbar-thumb { background-color: #CBD5E1; border-radius: 10px; }
    
    div[data-testid="stHorizontalBlock"]:has(.cal-header-marker),
    div[data-testid="stHorizontalBlock"]:has(.cal-cell-marker) {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        min-width: 800px !important;
        gap: 0 !important;
        border-left: 1px solid #E5E7EB !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.cal-header-marker) {
        border-top: 1px solid #E5E7EB !important;
    }

    div[data-testid="column"]:has(.cal-header-marker),
    div[data-testid="column"]:has(.cal-cell-marker) {
        width: 14.285% !important;
        flex: 1 1 0% !important;
        min-width: 110px !important;
        border-right: 1px solid #E5E7EB !important;
        border-bottom: 1px solid #E5E7EB !important;
        padding: 4px !important;
    }
    
    div[data-testid="column"]:has(.cal-header-marker) {
        background-color: #F8FAFC !important;
        padding: 8px 0 !important;
        border-bottom: 2px solid #94A3B8 !important;
    }
    
    div[data-testid="column"]:has(.cal-cell-marker) {
        height: 140px !important;
        min-height: 140px !important;
        max-height: 140px !important;
        overflow-y: auto !important;
        overflow-x: hidden !important;
    }

    div[data-testid="column"]:has(.cal-cell-marker) div.element-container {
        margin-bottom: 0 !important;
        margin-top: 0 !important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.cal-cell-marker) {
        border: none !important;
        border-radius: 0px !important;
        padding: 0 !important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.cal-cell-marker) div[data-testid="stVerticalBlock"] {
        padding: 0 !important;
        gap: 0 !important;
    }

    .cal-header-marker, .cal-cell-marker, .event-marker {
        display: none !important;
    }

    div.element-container:has(div.event-marker) + div.element-container button {
        border: none !important;
        border-radius: 0px !important;
        padding: 2px 4px !important;
        min-height: 22px !important;
        height: auto !important;
        margin: 1px 0 !important;
        width: 100% !important;
        display: block !important;
        text-align: left !important;
        background-color: #E2E8F0 !important;
        color: #1A202C !important;
    }
    div.element-container:has(div.event-marker) + div.element-container button p {
        font-size: 11.5px !important;
        font-weight: 600 !important;
        margin: 0 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        color: inherit !important;
    }
    div.element-container:has(div.event-marker) + div.element-container button:hover {
        filter: brightness(0.92) !important;
        border-color: transparent !important;
    }

    div.element-container:has(div[data-team="team1"]) + div.element-container button { background-color: #AEE4FF !important; color: #1A202C !important; }
    div.element-container:has(div[data-team="team2"]) + div.element-container button { background-color: #60B65C !important; color: #FFFFFF !important; }
    div.element-container:has(div[data-team="team3"]) + div.element-container button { background-color: #F99B62 !important; color: #FFFFFF !important; }
    div.element-container:has(div[data-team="tf1"]) + div.element-container button { background-color: #FFF2A8 !important; color: #1A202C !important; }
    div.element-container:has(div[data-team="tf2"]) + div.element-container button { background-color: transparent !important; color: #7A5299 !important; border-left: 3px solid #7A5299 !important; }
    </style>
    """
    st.markdown(CSS, unsafe_allow_html=True)

def render_streamlit_calendar(site_data: dict, year: int, month: int, selected_site: Optional[str] = None):
    calendar.setfirstweekday(calendar.SUNDAY)
    cal = calendar.monthcalendar(year, month)
    
    inject_pure_css_calendar_style()

    header_cols = st.columns(7)
    for col, day_name in zip(header_cols, ['일', '월', '화', '수', '목', '금', '토']):
        color = "#e53e3e" if day_name == '일' else "#4a5568"
        with col:
            st.markdown(f"<div class='cal-header-marker'></div><div style='text-align:center; font-size:13px; font-weight:bold; color:{color}; padding:6px;'>{day_name}</div>", unsafe_allow_html=True)

    for week_idx, week in enumerate(cal):
        cols = st.columns(7)
        for col_idx, day in enumerate(week):
            with cols[col_idx]:
                if day == 0:
                    with st.container(height=120):
                        st.markdown("<div class='cal-cell-marker'></div>", unsafe_allow_html=True)
                    continue

                with st.container(height=120):
                    st.markdown("<div class='cal-cell-marker'></div>", unsafe_allow_html=True)
                    current_date = date(year, month, day)
                    day_color = "#e53e3e" if col_idx == 0 else "#4a5568"
                    
                    if current_date == date.today():
                        bg_style = "background-color: #1a202c; color: white; border-radius: 50%; width: 22px; height: 22px; display: inline-flex; align-items: center; justify-content: center; font-weight: bold; font-size: 13px;"
                        day_html = f"<div style='padding:0 2px 2px 2px; text-align:left; margin-bottom: 2px;'><span style='{bg_style}'>{day}</span></div>"
                    else:
                        day_html = f"<div style='color:{day_color}; padding:2px; font-size:13px; font-weight:600; text-align:left; margin-bottom: 2px;'>{day}</div>"
                    
                    st.markdown(day_html, unsafe_allow_html=True)

                    day_events = []
                    for site, steps in site_data.items():
                        if selected_site and selected_site != "전체 현장" and site != selected_site: continue
                        for step_idx, step in enumerate(steps):
                            if step.get('date') == current_date:
                                day_events.append((site, step_idx, step))

                    if not day_events: continue
                    day_events.sort(key=calendar_event_sort_key)

                    for event_no, (site, step_idx, step) in enumerate(day_events):
                        label = truncate_label(make_calendar_event_label(site, step), max_chars=28)
                        btn_key = make_streamlit_key("cal_btn", year, month, week_idx, col_idx, event_no, site, step_idx)
                        style_class = get_team_style_class(get_step_team_value(step))
                        
                        st.markdown(f"<div class='event-marker' data-team='{style_class}'></div>", unsafe_allow_html=True)
                        if st.button(label, key=btn_key, use_container_width=True):
                            show_schedule_edit_dialog(site, step_idx)

# ==========================================
# 🤖 5. 공무원 양식 AI 요약 프롬프트 적용
# ==========================================
def get_ai_summary_stream(file_path: str):
    yield "🔄 AI 분석 엔진을 초기화하는 중입니다...\n\n"
    from google import genai
    client = genai.Client(api_key=GEMINI_API_KEY)
    ext = os.path.splitext(file_path)[1].lower()
    prompt = """당신은 관공서(국토관리청 등)의 벌점심의위원회 또는 현장점검 결과 보고서를 작성하는 전문 행정관입니다.
    제공된 문서를 철저히 분석하여, 아래의 [공식 심의안건 보고서 양식]에 맞추어 완벽하게 요약 및 재작성하십시오.
    [주의사항]
    - 반드시 명조체 느낌의 정중하고 딱딱한 공문서 개조식 어투(~함, ~임)를 사용하십시오.
    - 문서에 없는 내용을 절대 임의로 지어내지 마십시오. 정보가 부족한 항목은 과감히 생략하십시오.
    """
    uploaded_file = safe_filepath = None
    try:
        if ext in ['.pdf', '.png', '.jpg', '.jpeg', '.txt']:
            safe_filename = secure_filename(f"temp_ai_upload_{int(time.time())}{ext}") 
            safe_filepath = os.path.join(ATTACH_DIR, safe_filename)
            shutil.copy2(file_path, safe_filepath)
            uploaded_file = client.files.upload(file=safe_filepath)
            
            if ext == '.pdf':
                yield "📄 PDF 문서를 스캔하고 있습니다. (약 5~10초 소요)...\n\n"
                max_retries, retries = 15, 0
                while retries < max_retries:
                    file_info = client.files.get(name=uploaded_file.name)
                    if "ACTIVE" in str(file_info.state).upper(): break
                    if "FAILED" in str(file_info.state).upper():
                        yield "❌ 서버에서 문서를 읽는 데 실패했습니다."
                        return
                    time.sleep(2)
                    retries += 1
                if retries >= max_retries:
                    yield "❌ PDF 분석 시간이 초과되었습니다."
                    return
                    
            yield "💡 스캔 완료! 보고서 작성을 시작합니다...\n\n---\n\n"
            for chunk in client.models.generate_content_stream(model='gemini-2.5-flash', contents=[prompt, uploaded_file]):
                if chunk.text: yield chunk.text
        else:
            yield "⚠️ 스트림릿 환경에서는 PDF, TXT, 이미지 요약만 지원합니다."
    except Exception as e:
        yield f"\n\n❌ AI 분석 중 오류가 발생했습니다.\n상세: {e}"
    finally:
        try:
            if uploaded_file: client.files.delete(name=uploaded_file.name)
            if safe_filepath and os.path.exists(safe_filepath): os.remove(safe_filepath)
        except Exception: pass 

# ==========================================
# 💾 6. 데이터 처리 (엑셀/CSV 파싱 포함)
# ==========================================
def check_password() -> bool:
    if st.session_state.get("logged_in"): return True
    st.markdown("## 🏛️ 현장점검 통합관리 시스템 Login")
    with st.form("login_form"):
        user_id = st.text_input("아이디")
        password = st.text_input("비밀번호", type="password")
        if st.form_submit_button("접속하기"):
            if user_id == SHARED_USER_ID and password == SHARED_PASSWORD:
                st.session_state["logged_in"] = True
                st.rerun()
            else: st.error("아이디 또는 비밀번호가 일치하지 않습니다.")
    return False

def adjust_weekend(date_obj: date) -> date:
    wd = date_obj.weekday()
    if wd == 5: return date_obj + timedelta(days=2)
    if wd == 6: return date_obj + timedelta(days=1)
    return date_obj

def load_data() -> dict:
    site_data = {}
    if not os.path.exists(DB_FILENAME): return site_data
    try:
        with open(DB_FILENAME, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            if not reader.fieldnames: return site_data
            for row in reader:
                name, date_str, desc = clean_cell(row.get('현장명', '')), clean_cell(row.get('날짜', '')), clean_cell(row.get('업무명', ''))
                if not name or not date_str or not desc: continue
                date_obj = parse_date_value(date_str)
                if not date_obj: continue
                files_str = clean_cell(row.get('파일경로', ''))
                step = {
                    "date": date_obj, "desc": desc, "memo": clean_cell(row.get('메모', '')),
                    "files": files_str.split("|") if files_str else [],
                }
                for key, label, _ in STEP_EXTRA_FIELDS: step[key] = clean_cell(row.get(label, ''))
                if not step.get("inspection_period"): step["inspection_period"] = infer_inspection_period(plan_date=date_obj, desc=desc)
                if not step.get("team"): step["team"] = extract_team_from_desc(desc)
                for key, label, _ in SITE_DETAIL_FIELDS: step[key] = clean_cell(row.get(label, ''))
                site_data.setdefault(name, []).append(step)
        for name in site_data: site_data[name].sort(key=lambda x: x['date'])
    except Exception as e: st.error(f"데이터 로드 오류: {e}")
    return site_data

def save_data(site_data: dict) -> None:
    rows = []
    row_num = 1
    for name in sorted(site_data.keys()):
        for step in site_data[name]:
            row = [row_num, name, step['date'].strftime('%Y-%m-%d'), step.get('desc', ''), step.get('memo', ''), "|".join(step.get('files', []))]
            for key, _, _ in STEP_EXTRA_FIELDS: row.append(step.get(key, ''))
            for key, _, _ in SITE_DETAIL_FIELDS: row.append(step.get(key, ''))
            rows.append(row)
            row_num += 1
    atomic_save_csv(DB_FILENAME, DB_COLUMNS, rows)

def process_excel_schedule(file) -> None:
    try:
        if file.name.lower().endswith('.csv'):
            try: df = pd.read_csv(file, encoding='utf-8-sig', header=None)
            except Exception: df = pd.read_csv(file, encoding='cp949', header=None)
        else: df = pd.read_excel(file, header=None)

        header_idx = -1
        for i, row in df.iterrows():
            if "공사명" in "".join(str(val) for val in row.values) and "점검예정일" in "".join(str(val) for val in row.values):
                header_idx = i
                break

        if header_idx != -1:
            df.columns = df.iloc[header_idx]
            df = df.iloc[header_idx + 1:]
        else:
            return st.error("엑셀 파일 양식이 맞지 않습니다. '공사명' 및 '점검예정일' 열을 찾을 수 없습니다.")

        success_count, updated_count, default_year = 0, 0, date.today().year

        for idx, row in df.iterrows():
            site_name = get_row_value(row, ["공사명", "현장명", "프로젝트명"])
            plan_date = parse_date_value(get_row_value(row, ["점검예정일", "점검 예정일", "점검일", "날짜"]), default_year)
            if not site_name or not plan_date: continue

            team, inspectors = get_row_value(row, ["담당조", "점검조", "조", "반"]), get_row_value(row, ["점검자", "점검자명", "담당자"])
            inspection_period = infer_inspection_period(file.name, row=row, plan_date=plan_date)
            detail_values = {key: get_row_value(row, aliases) for key, _, aliases in SITE_DETAIL_FIELDS}
            desc = f"[{team}] {inspection_period}".strip() if team else inspection_period

            memo = "\n".join([f"🏢 발주처: {detail_values['client']}" if detail_values.get("client") else "",
                              f"👷 시공사: {detail_values['builder']}" if detail_values.get("builder") else "",
                              f"🔍 감리사: {detail_values['supervisor']}" if detail_values.get("supervisor") else ""]).strip()

            st.session_state.site_data.setdefault(site_name, [])
            existing_step = next((s for s in st.session_state.site_data[site_name] if s['date'] == plan_date and s['desc'] == desc), None)

            if existing_step:
                existing_step.update({"inspection_period": inspection_period, "team": team, "inspectors": inspectors})
                for k, v in detail_values.items():
                    if v: existing_step[k] = v
                if memo and not existing_step.get('memo'): existing_step['memo'] = memo
                updated_count += 1
            else:
                inherited_details = get_site_detail_defaults(st.session_state.site_data[site_name])
                for k, v in detail_values.items():
                    if v: inherited_details[k] = v
                st.session_state.site_data[site_name].append({
                    "date": plan_date, "desc": desc, "memo": memo, "files": [],
                    "inspection_period": inspection_period, "team": team, "inspectors": inspectors, **inherited_details,
                })
                st.session_state.site_data[site_name].sort(key=lambda x: x['date'])
                success_count += 1

        if success_count > 0 or updated_count > 0:
            save_data(st.session_state.site_data)
            st.success(f"신규 {success_count}건 등록, 기존 {updated_count}건 보강 완료")
            time.sleep(1)
            st.rerun()
    except Exception as e: st.error(f"엑셀 처리 중 오류 발생: {e}")

# ==========================================
# 🌐 7. 메인 앱 UI 
# ==========================================
def main():
    if not check_password(): return

    if "site_data" not in st.session_state: st.session_state.site_data = load_data()
    if "cal_year" not in st.session_state: st.session_state.cal_year = date.today().year
    if "cal_month" not in st.session_state: st.session_state.cal_month = date.today().month

    st.title("🏗️ 현장점검 통합관리 시스템")

    with st.sidebar:
        st.header("💾 데이터 백업 및 복구")
        st.info("⚠️ 클라우드 수면 모드로 인한 데이터 초기화 대비용입니다. 주기적으로 백업을 다운로드 해두세요.")
        if os.path.exists(DB_FILENAME):
            with open(DB_FILENAME, "rb") as f:
                st.download_button(label="⬇️ 현재 데이터 백업 (다운로드)", data=f, file_name=f"backup_{date.today().strftime('%Y%m%d')}.csv", mime="text/csv", use_container_width=True)
        backup_file = st.file_uploader("⬆️ 백업 파일 복구 (업로드)", type=['csv'], label_visibility="collapsed")
        if backup_file and st.button("🔄 시스템 복구 실행", use_container_width=True):
            with open(DB_FILENAME, "wb") as f: f.write(backup_file.getbuffer())
            st.session_state.site_data = load_data()
            st.success("데이터가 성공적으로 복구되었습니다!")
            time.sleep(1)
            st.rerun()
            
        st.divider()
        st.header("📁 엑셀 일정 일괄 등록")
        excel_file = st.file_uploader("엑셀/CSV 파일 업로드", type=['csv', 'xlsx', 'xls'])
        if st.button("🚀 일정 자동 등록하기", type="primary", use_container_width=True) and excel_file:
            process_excel_schedule(excel_file)
        
        st.divider()
        st.header("➕ 개별 프로젝트 등록")
        with st.form("add_project_form"):
            new_site_name = st.text_input("프로젝트(현장)명")
            start_date = st.date_input("점검 예정일", value=date.today())
            default_period = f"{start_date.month}월 상시점검"
            new_inspection_period = st.selectbox("점검시기", selectbox_options_with_current(default_period))
            pp1, pp2 = st.columns(2)
            with pp1: new_team = st.text_input("담당조", placeholder="예: 1조")
            with pp2: new_inspectors = st.text_input("점검자", placeholder="예: 홍길동, 김철수")
            new_detail_values = render_site_detail_inputs({}, "new_project_detail")
            if st.form_submit_button("초기 점검일정 생성") and new_site_name:
                if new_site_name not in st.session_state.site_data:
                    st.session_state.site_data[new_site_name] = [{
                        "date": start_date, "desc": "현장점검 실시", "memo": "", "files": [],
                        "inspection_period": new_inspection_period, "team": new_team, "inspectors": new_inspectors,
                        **new_detail_values,
                    }]
                    save_data(st.session_state.site_data)
                    st.rerun()

        st.divider()
        st.header("📋 프로젝트 선택")
        search_query = st.text_input("🔍 현장명 검색")
        all_sites = sorted(list(st.session_state.site_data.keys()))
        site_options = ["전체 현장"] + [s for s in all_sites if search_query.lower() in s.lower()]
        with st.container(height=300, border=True):
            selected_site = st.radio("일정을 볼 현장 선택", site_options, label_visibility="collapsed", format_func=lambda site: make_site_list_label(site, st.session_state.site_data))
        if selected_site != "전체 현장" and st.button("🗑️ 현재 프로젝트 삭제", type="primary", use_container_width=True):
            del st.session_state.site_data[selected_site]
            save_data(st.session_state.site_data)
            st.rerun()

    st.subheader("🗓️ 프로젝트 전체 일정 캘린더")
    c1, c2, c3 = st.columns([1, 4, 1])
    with c1:
        if st.button("◀ 이전 달", use_container_width=True):
            st.session_state.cal_month, st.session_state.cal_year = (12, st.session_state.cal_year - 1) if st.session_state.cal_month == 1 else (st.session_state.cal_month - 1, st.session_state.cal_year)
            st.rerun()
    with c2: st.markdown(f"<h3 style='text-align:center;'>{st.session_state.cal_year}년 {st.session_state.cal_month}월</h3>", unsafe_allow_html=True)
    with c3:
        if st.button("다음 달 ▶", use_container_width=True):
            st.session_state.cal_month, st.session_state.cal_year = (1, st.session_state.cal_year + 1) if st.session_state.cal_month == 12 else (st.session_state.cal_month + 1, st.session_state.cal_year)
            st.rerun()

    render_streamlit_calendar(st.session_state.site_data, st.session_state.cal_year, st.session_state.cal_month, selected_site)
    st.divider()

    if selected_site != "전체 현장":
        st.subheader(f"📂 [{selected_site}] 세부 일정 및 파일 관리")
        steps = st.session_state.site_data[selected_site]

        with st.expander("🏗️ 현장 기본정보 수정", expanded=False):
            with st.form(f"site_detail_form_{selected_site}"):
                detail_values = render_site_detail_inputs(get_site_detail_defaults(steps), f"site_detail_{selected_site}")
                if st.form_submit_button("저장", type="primary"):
                    apply_site_details_to_all_steps(selected_site, detail_values)
                    save_data(st.session_state.site_data)
                    st.rerun()

        add_col1, add_col2 = st.columns(2)
        with add_col1:
            with st.expander("📌 단순 일정 수동 추가"):
                e1, e2 = st.columns([1, 2])
                with e1:
                    custom_date = st.date_input("날짜", key="c_date")
                    custom_period = st.selectbox("점검시기", selectbox_options_with_current(f"{custom_date.month}월 상시점검"), key="c_period")
                with e2:
                    custom_desc = st.text_input("업무 내용", key="c_desc")
                    custom_team = st.text_input("담당조", key="c_team")
                    custom_inspectors = st.text_input("점검자", key="c_inspectors")
                if st.button("일정 끼워넣기", use_container_width=True):
                    steps.append({"date": adjust_weekend(custom_date), "desc": custom_desc or custom_period, "memo": "", "files": [], "inspection_period": custom_period, "team": custom_team, "inspectors": custom_inspectors, **get_site_detail_defaults(steps)})
                    steps.sort(key=lambda x: x['date'])
                    save_data(st.session_state.site_data)
                    st.rerun()

        with add_col2:
            with st.expander("🚨 벌점/과태료 발생 시 (후속 행정절차 자동 생성)"):
                base_step = next((s for s in steps if "현장점검" in s['desc']), None)
                penalty_base_date = st.date_input("기준일", value=base_step['date'] if base_step else date.today())
                if st.button("⚠️ 일괄 생성", type="primary", use_container_width=True):
                    if "확인서 이의제기 접수" not in [s['desc'] for s in steps]:
                        curr = penalty_base_date
                        for days, desc in PENALTY_INTERVALS:
                            curr = adjust_weekend(curr + timedelta(days=days))
                            steps.append({"date": curr, "desc": desc, "memo": "", "files": [], "inspection_period": "기타", "team": "", "inspectors": "", **get_site_detail_defaults(steps)})
                        steps.sort(key=lambda x: x['date'])
                        save_data(st.session_state.site_data)
                        st.rerun()

        if "current_page" not in st.session_state: st.session_state.current_page = 1
        total_pages = max(1, (len(steps) - 1) // ITEMS_PER_PAGE + 1)
        st.session_state.current_page = min(st.session_state.current_page, total_pages)
        start_idx = (st.session_state.current_page - 1) * ITEMS_PER_PAGE
        current_page_steps = steps[start_idx:start_idx + ITEMS_PER_PAGE]

        for i, step in enumerate(current_page_steps):
            actual_idx = start_idx + i  
            with st.container(border=True):
                c1, c2, c3 = st.columns([2, 5, 4])
                with c1:
                    new_date = st.date_input("기한", value=step['date'], key=f"date_{actual_idx}")
                    cur_per = clean_cell(step.get('inspection_period', '')) or infer_inspection_period(plan_date=step.get('date'), desc=step.get('desc', ''))
                    new_period = st.selectbox("점검시기", selectbox_options_with_current(cur_per), index=selectbox_options_with_current(cur_per).index(cur_per), key=f"period_{actual_idx}")
                    new_team = st.text_input("담당조", value=clean_cell(step.get('team', '')) or extract_team_from_desc(step.get('desc', '')), key=f"team_{actual_idx}")
                    new_desc = st.text_input("업무명", value=step['desc'], key=f"desc_{actual_idx}")
                    new_inspectors = st.text_area("점검자", value=clean_cell(step.get('inspectors', '')), height=70, key=f"inspectors_{actual_idx}")
                    
                    if any([new_date != step['date'], new_desc != step['desc'], new_period != cur_per, new_team != step.get('team'), new_inspectors != step.get('inspectors')]):
                        steps[actual_idx].update({"date": new_date, "inspection_period": new_period, "team": new_team, "desc": new_desc, "inspectors": new_inspectors})
                        steps.sort(key=lambda x: x['date'])
                        save_data(st.session_state.site_data)
                        st.rerun()
                    if st.button("❌ 일정 전체 삭제", key=f"del_{actual_idx}"):
                        steps.pop(actual_idx)
                        save_data(st.session_state.site_data)
                        st.rerun()
                
                with c2:
                    new_memo = st.text_area("📝 메모", value=step.get('memo', ''), height=100, key=f"memo_{actual_idx}")
                    if new_memo != step.get('memo', ''):
                        steps[actual_idx]['memo'] = new_memo
                        save_data(st.session_state.site_data)
                
                with c3:
                    st.markdown("**📂 첨부 파일 (드래그 앤 드롭)**")
                    uploaded_files = st.file_uploader("업로드", accept_multiple_files=True, key=f"up_{actual_idx}", label_visibility="collapsed")
                    
                    if uploaded_files:
                        files_changed = False
                        for uf in uploaded_files:
                            safe_name = secure_filename(uf.name)
                            original_path = os.path.join(ATTACH_DIR, f"{selected_site}_{safe_name}")
                            if original_path not in steps[actual_idx].get('files', []):
                                with open(original_path, "wb") as f: f.write(uf.getbuffer())
                                steps[actual_idx].setdefault('files', []).append(original_path)
                                files_changed = True
                                if original_path.lower().endswith(".hwp"):
                                    pdf_path = hwp_to_pdf(original_path)
                                    if pdf_path != original_path and os.path.exists(pdf_path) and pdf_path not in steps[actual_idx]['files']:
                                        steps[actual_idx]['files'].append(pdf_path)
                        
                        if files_changed:
                            save_data(st.session_state.site_data)
                            st.rerun()

                    existing_files = [f for f in steps[actual_idx].get('files', []) if os.path.exists(f)]
                    if len(existing_files) != len(steps[actual_idx].get('files', [])):
                        steps[actual_idx]['files'] = existing_files
                        save_data(st.session_state.site_data)

                    checked_files_to_download = []
                    for file_path in existing_files:
                        file_name = os.path.basename(file_path)
                        ext = file_name.lower().split('.')[-1]
                        
                        chk_col, btn_col1, btn_col2, btn_col3 = st.columns([4.5, 2.5, 2.5, 2.5])
                        
                        with chk_col:
                            if st.checkbox(f"📎 {file_name}", key=f"chk_{actual_idx}_{file_path}"):
                                checked_files_to_download.append(file_path)
                        
                        with btn_col1:
                            with open(file_path, "rb") as f:
                                st.download_button(label="파일보기", data=f, file_name=file_name, use_container_width=True, key=f"v_{actual_idx}_{file_path}")
                        
                        if ext in ['pdf', 'png', 'jpg', 'jpeg', 'txt']:
                            with btn_col2:
                                if st.button("AI요약", key=f"ai_{actual_idx}_{file_path}", use_container_width=True):
                                    show_summary_dialog(file_path, file_name)
                        else:
                            with btn_col2: st.write("")
                            
                        with btn_col3:
                            if st.button("삭제", key=f"delf_{actual_idx}_{file_path}", use_container_width=True):
                                steps[actual_idx]['files'].remove(file_path)
                                try: os.remove(file_path) 
                                except: pass
                                save_data(st.session_state.site_data)
                                st.rerun()

                    if checked_files_to_download:
                        st.markdown("<br>", unsafe_allow_html=True)
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for fpath in checked_files_to_download: zip_file.write(fpath, arcname=os.path.basename(fpath))
                        st.download_button(
                            label=f"💾 체크된 파일 {len(checked_files_to_download)}개 전체 다운로드 (.zip)",
                            data=zip_buffer.getvalue(),
                            file_name=f"첨부파일_다운로드_{date.today().strftime('%Y%m%d')}.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True,
                            key=f"zip_dl_{actual_idx}"
                        )

if __name__ == "__main__":
    main()
