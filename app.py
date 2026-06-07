import os
# 특정 네트워크에서 gRPC 통신 타임아웃 방지
os.environ.setdefault("GRPC_DNS_RESOLVER", "native")

import base64
import calendar
import csv
import hashlib
import hmac
import html
import io
import re
import shutil
import time
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from uuid import uuid4

import streamlit as st
import streamlit.components.v1 as components

try:
    from google import genai
except ImportError:
    genai = None

# ==========================================
# 🛑 HWP -> PDF 변환용 라이브러리 (윈도우 전용)
# ==========================================
try:
    import win32com.client as win32
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# ==========================================
# ⚙️ 1. 기본 설정 및 전역 변수
# ==========================================
st.set_page_config(page_title="현장점검 관리 시스템", page_icon="🏛️", layout="wide")

BASE_DIR = Path(__file__).resolve().parent


def get_config(name: str, default: str = "") -> str:
    """환경변수 또는 Streamlit secrets.toml에서 설정값을 가져옴."""
    env_value = os.getenv(name)
    if env_value not in (None, ""):
        return env_value
    try:
        secret_value = st.secrets.get(name)
        if secret_value not in (None, ""):
            return str(secret_value)
    except Exception:
        pass
    return default


APP_USER_ID = get_config("APP_USER_ID")
APP_PASSWORD = get_config("APP_PASSWORD")
APP_PASSWORD_HASH = get_config("APP_PASSWORD_HASH")  # sha256 해시 권장
GEMINI_API_KEY = get_config("GEMINI_API_KEY")

DB_PATH = Path(get_config("DB_FILENAME", str(BASE_DIR / "penalty_database.csv")))
if not DB_PATH.is_absolute():
    DB_PATH = BASE_DIR / DB_PATH

ATTACH_DIR = Path(get_config("ATTACH_DIR", str(BASE_DIR / "attachments")))
if not ATTACH_DIR.is_absolute():
    ATTACH_DIR = BASE_DIR / ATTACH_DIR
ATTACH_DIR.mkdir(parents=True, exist_ok=True)

ITEMS_PER_PAGE = 10
MAX_UPLOAD_MB = int(get_config("MAX_UPLOAD_MB", "50"))
SUPPORTED_PREVIEW_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".txt"}
SUPPORTED_AI_EXTS = {".pdf", ".png", ".jpg", ".jpeg", ".txt"}

PENALTY_INTERVALS = [
    (30, "확인서 이의제기 접수"),
    (14, "확인서 이의제기 의견 통보"),
    (14, "벌점 사전부과 통보"),
    (15, "벌점 사전부과 통보 의견제출 마감"),
    (15, "벌점 사전부과 의견 검토회의"),
    (1, "벌점 사전부과 의견 검토회의 결과 통보 및 벌점 부과"),
    (30, "벌점 부과 이의제기 접수"),
    (40, "벌점 심의위원회 개최 및 최종 결과 통보"),
]


# ==========================================
# 🛠️ 2. 공통 유틸리티
# ==========================================
def make_step_id() -> str:
    return uuid4().hex[:12]


def safe_slug(value: str, max_len: int = 80) -> str:
    value = str(value or "").strip()
    value = re.sub(r"[^0-9A-Za-z가-힣._-]+", "_", value)
    value = value.strip("._-")
    return (value[:max_len] or "file")


def to_storage_path(path: Path) -> str:
    """DB에는 가능하면 앱 폴더 기준 상대경로로 저장하여 배포 위치 변경에 대응."""
    try:
        return str(path.resolve().relative_to(BASE_DIR))
    except ValueError:
        return str(path.resolve())


def resolve_path(path_value: str) -> Path:
    path = Path(path_value)
    if path.is_absolute():
        return path
    return BASE_DIR / path


def unique_attachment_path(site_name: str, original_filename: str) -> Path:
    """수동 저장용 예비 함수. 업로드 파일은 아래 deterministic_attachment_path를 우선 사용."""
    original = Path(original_filename)
    stem = safe_slug(original.stem, 70)
    ext = original.suffix.lower()
    site = safe_slug(site_name, 50)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return ATTACH_DIR / f"{site}_{stamp}_{uuid4().hex[:8]}_{stem}{ext}"


def deterministic_attachment_path(site_name: str, step_id: str, original_filename: str, file_digest: str) -> Path:
    """
    같은 파일이 Streamlit 재실행 때 반복 저장되지 않도록,
    현장명 + 일정ID + 파일내용 해시 + 원본 파일명으로 항상 같은 저장 경로를 생성함.
    """
    original = Path(original_filename)
    stem = safe_slug(original.stem, 70)
    ext = original.suffix.lower()
    site = safe_slug(site_name, 50)
    step = safe_slug(step_id, 20)
    digest = safe_slug(file_digest[:12], 12)
    return ATTACH_DIR / f"{site}_{step}_{digest}_{stem}{ext}"


def read_text_file(file_path: Path) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp949", "euc-kr"):
        try:
            return file_path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return file_path.read_text(encoding="utf-8", errors="replace")


def verify_password(password: str) -> bool:
    if APP_PASSWORD_HASH:
        candidate = hashlib.sha256(password.encode("utf-8")).hexdigest()
        return hmac.compare_digest(candidate, APP_PASSWORD_HASH)
    if APP_PASSWORD:
        return hmac.compare_digest(password, APP_PASSWORD)
    return False


def make_backup_zip() -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        if DB_PATH.exists():
            zip_file.write(DB_PATH, arcname=DB_PATH.name)
        if ATTACH_DIR.exists():
            for file_path in ATTACH_DIR.rglob("*"):
                if file_path.is_file():
                    zip_file.write(file_path, arcname=str(Path("attachments") / file_path.relative_to(ATTACH_DIR)))
    return buffer.getvalue()


# ==========================================
# 🛑 HWP -> PDF 변환
# ==========================================
def hwp_to_pdf(hwp_path: Path) -> Path:
    """HWP/HWPX 파일을 PDF로 변환. Windows + 한글 설치 + pywin32 환경에서만 작동."""
    if not WIN32_AVAILABLE:
        st.warning("이 서버에는 pywin32가 없어 HWP PDF 자동 변환을 건너뜁니다. Windows PC에서 실행하거나 PDF로 변환 후 업로드해 주세요.")
        return hwp_path

    pdf_path = hwp_path.with_suffix(".pdf")
    hwp = None
    co_initialized = False

    try:
        pythoncom.CoInitialize()
        co_initialized = True
        hwp = win32.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(str(hwp_path.resolve()))
        hwp.SaveAs(str(pdf_path.resolve()), "PDF")
        return pdf_path if pdf_path.exists() else hwp_path
    except Exception as e:
        st.error(f"HWP -> PDF 변환 실패: {e}")
        return hwp_path
    finally:
        try:
            if hwp:
                hwp.Quit()
        except Exception:
            pass
        if co_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ==========================================
# 🖱️ 3. 새창(다이얼로그) 마우스 드래그 기능 주입
# ==========================================
def make_dialog_draggable():
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
                let isDragging = false;
                let startX, startY;
                let currentX = 0;
                let currentY = 0;
                dialog.style.position = 'relative';

                header.addEventListener('mousedown', (e) => {
                    isDragging = true;
                    header.style.cursor = 'grabbing';
                    startX = e.clientX;
                    startY = e.clientY;
                    doc.body.style.userSelect = 'none';
                });

                doc.addEventListener('mousemove', (e) => {
                    if (!isDragging) return;
                    const dx = e.clientX - startX;
                    const dy = e.clientY - startY;
                    currentX += dx;
                    currentY += dy;
                    dialog.style.left = currentX + 'px';
                    dialog.style.top = currentY + 'px';
                    startX = e.clientX;
                    startY = e.clientY;
                });

                const stopDrag = () => {
                    if (isDragging) {
                        isDragging = false;
                        header.style.cursor = 'grab';
                        doc.body.style.userSelect = '';
                    }
                };
                doc.addEventListener('mouseup', stopDrag);
                doc.addEventListener('mouseleave', stopDrag);
            }
        });
    };
    const observer = new MutationObserver(setupDrag);
    observer.observe(doc.body, { childList: true, subtree: true });
    setTimeout(setupDrag, 100);
    </script>
    """
    components.html(drag_js, height=0, width=0)


# ==========================================
# 🔐 4. 새창(다이얼로그) 팝업 UI 구현
# ==========================================
@st.dialog("✨ AI 심의안건 보고서 작성", width="large")
def show_summary_dialog(file_path_value: str, file_name: str):
    make_dialog_draggable()
    file_path = resolve_path(file_path_value)
    st.markdown(f"### 📄 {html.escape(file_name)} 분석 결과")
    if not file_path.exists():
        st.error("파일을 찾을 수 없습니다. 첨부파일 목록을 새로고침해 주세요.")
        return
    with st.spinner("AI가 보고서를 작성 중입니다..."):
        st.write_stream(get_ai_summary_stream(file_path))
    if st.button("닫기", type="primary"):
        st.rerun()


@st.dialog("📄 첨부 문서 뷰어", width="large")
def show_file_dialog(file_path_value: str, file_name: str):
    make_dialog_draggable()
    file_path = resolve_path(file_path_value)
    st.markdown(f"### 📎 {html.escape(file_name)}")

    if not file_path.exists():
        st.error("파일을 찾을 수 없습니다. 첨부파일 목록을 새로고침해 주세요.")
        return

    ext = file_path.suffix.lower()
    if ext in {".png", ".jpg", ".jpeg"}:
        st.image(str(file_path), use_container_width=True)
    elif ext == ".pdf":
        with file_path.open("rb") as f:
            base64_pdf = base64.b64encode(f.read()).decode("utf-8")
        pdf_display = (
            f'<iframe src="data:application/pdf;base64,{base64_pdf}" '
            'width="100%" height="700" type="application/pdf"></iframe>'
        )
        st.markdown(pdf_display, unsafe_allow_html=True)
    elif ext == ".txt":
        st.text_area("문서 내용", read_text_file(file_path), height=500)
    else:
        st.warning("웹 미리보기를 지원하지 않는 형식입니다. 체크박스를 선택해 ZIP으로 다운로드해 주세요.")


# ==========================================
# 📅 5. 중앙 달력 렌더링 함수
# ==========================================
def render_html_calendar(site_data, year, month, selected_site=None):
    cal = calendar.monthcalendar(year, month)
    html_parts = [
        "<style>",
        ".cal-cell-scroll::-webkit-scrollbar { width: 6px; }",
        ".cal-cell-scroll::-webkit-scrollbar-thumb { background: #c1c1c1; border-radius: 3px; }",
        ".cal-cell-scroll::-webkit-scrollbar-track { background: #f1f1f1; }",
        "</style>",
        "<table style='width:100%; table-layout:fixed; border-collapse: collapse; text-align:left; background-color:#ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.1);'>",
        "<tr>",
    ]

    for d in ["월", "화", "수", "목", "금", "토", "일"]:
        html_parts.append(f"<th style='border:1px solid #ddd; padding:10px; text-align:center; background-color:#f0f2f6; font-size:16px;'>{d}</th>")
    html_parts.append("</tr>")

    for week in cal:
        html_parts.append("<tr>")
        for day in week:
            if day == 0:
                html_parts.append("<td style='border:1px solid #ddd; background-color:#fafafa;'></td>")
                continue

            current_date = date(year, month, day)
            bg_color = "#e6f2ff" if current_date == date.today() else "#ffffff"
            day_events = []

            for site, steps in site_data.items():
                if selected_site and selected_site != "전체 현장" and site != selected_site:
                    continue
                for step in steps:
                    if step.get("date") == current_date:
                        desc = step.get("desc", "")
                        event_bg = "#d1e7dd" if "현장점검" in desc else "#ffefc2"
                        event_border = "#0f5132" if "현장점검" in desc else "#ffc107"
                        safe_site = html.escape(site)
                        safe_desc = html.escape(desc)
                        day_events.append(
                            f"<div style='font-size:12px; margin-top:4px; padding:4px; background-color:{event_bg}; border-left:3px solid {event_border}; border-radius:3px;'>"
                            f"<b>[{safe_site}]</b> {safe_desc}</div>"
                        )

            max_display = 20
            events_html = ""
            for i, event in enumerate(day_events):
                if i < max_display:
                    events_html += event
                elif i == max_display:
                    events_html += (
                        f"<div style='font-size:12px; margin-top:4px; padding:4px; text-align:center; background-color:#e0e0e0; color:#555; border-radius:3px;'>"
                        f"...외 {len(day_events) - max_display}건 더 있음</div>"
                    )
                    break

            html_parts.append(f"<td style='border:1px solid #ddd; padding:0; vertical-align:top; background-color:{bg_color};'>")
            html_parts.append("<div style='display:flex; flex-direction:column; height:130px; padding:6px;'>")
            html_parts.append(f"<strong style='font-size:16px; margin-bottom:2px; flex-shrink:0;'>{day}</strong>")
            html_parts.append("<div class='cal-cell-scroll' style='flex-grow:1; overflow-y:auto; padding-right:4px;'>")
            html_parts.append(events_html)
            html_parts.append("</div></div></td>")
        html_parts.append("</tr>")
    html_parts.append("</table><br>")
    return "".join(html_parts)


# ==========================================
# 🤖 6. 공무원 양식 AI 요약 프롬프트 적용
# ==========================================
def get_ai_summary_stream(file_path: Path):
    yield "🔄 AI 분석 엔진을 초기화하는 중입니다...\n\n"

    if genai is None:
        yield "❌ google-genai 라이브러리가 설치되어 있지 않습니다. `pip install google-genai` 후 다시 실행해 주세요."
        return
    if not GEMINI_API_KEY:
        yield "❌ GEMINI_API_KEY가 설정되어 있지 않습니다. Streamlit secrets 또는 환경변수에 키를 등록해 주세요."
        return

    ext = file_path.suffix.lower()
    if ext not in SUPPORTED_AI_EXTS:
        yield "⚠️ 현재 AI 요약은 PDF, TXT, PNG, JPG, JPEG 파일만 지원합니다."
        return

    prompt = """당신은 관공서(국토관리청 등)의 벌점심의위원회 또는 현장점검 결과 보고서를 작성하는 전문 행정관입니다.
제공된 문서를 철저히 분석하여, 아래의 [공식 심의안건 보고서 양식]에 맞추어 요약 및 재작성하십시오.

[공식 심의안건 보고서 양식]

(안건) [문서의 핵심 주제를 제목으로 작성]

1. 안건 개요

가. 안건명 : [핵심 주제]

나. 공사개요 (문서에 내용이 있는 경우만 작성, 없으면 생략)
- 공사명:
- 시공자/감리자:

다. 지적/부실내용
- [주요 위반사항 및 관련 법령, 기준 등 명시]

라. 처분내용(벌점 등)
- [어떤 처분/벌점이 부과되었는지]

2. 당사자 이의 신청 (해당 내용이 있을 경우에만 작성)
- [시공사, 감리사 등의 이의제기 및 주장 요약]

[주의사항]
- 반드시 정중하고 딱딱한 공문서 개조식 어투(~함, ~임)를 사용하십시오.
- 문서에 없는 내용을 절대 임의로 지어내지 마십시오.
- 정보가 부족한 항목은 "자료에서 확인되지 않음"이라고 쓰지 말고 과감히 생략하십시오.
"""

    client = genai.Client(api_key=GEMINI_API_KEY)
    uploaded_file = None
    temp_path = None

    try:
        yield "🚀 문서를 안전하게 복사하여 AI 서버로 전송합니다...\n\n"
        temp_path = ATTACH_DIR / f"temp_ai_upload_{uuid4().hex}{ext}"
        shutil.copy2(file_path, temp_path)
        uploaded_file = client.files.upload(file=str(temp_path))

        yield "📄 문서를 처리 중입니다...\n\n"
        started = time.time()
        while True:
            file_info = client.files.get(name=uploaded_file.name)
            state_str = str(file_info.state).upper()
            if "ACTIVE" in state_str:
                break
            if "FAILED" in state_str:
                yield "❌ AI 서버에서 문서를 읽는 데 실패했습니다. PDF가 암호화되어 있거나 손상되었는지 확인해 주세요."
                return
            if time.time() - started > 90:
                yield "❌ AI 서버 문서 처리 시간이 너무 오래 걸립니다. 파일 용량을 줄인 뒤 다시 시도해 주세요."
                return
            time.sleep(2)

        yield "💡 스캔 완료. 보고서 작성을 시작합니다...\n\n---\n\n"
        response_stream = client.models.generate_content_stream(
            model="gemini-2.5-flash",
            contents=[prompt, uploaded_file],
        )
        for chunk in response_stream:
            if getattr(chunk, "text", None):
                yield chunk.text
    except Exception as e:
        error_msg = str(e)
        if "503" in error_msg or "high demand" in error_msg.lower():
            yield "\n\n❌ 현재 AI 서버가 혼잡합니다. 잠시 후 다시 요약 버튼을 눌러주세요."
        else:
            yield f"\n\n❌ AI 분석 중 오류가 발생했습니다.\n상세: {error_msg}"
    finally:
        if uploaded_file:
            try:
                client.files.delete(name=uploaded_file.name)
            except Exception:
                pass
        if temp_path and temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass


# ==========================================
# 💾 7. 데이터 처리 함수
# ==========================================
def check_password():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if st.session_state["logged_in"]:
        return True

    st.markdown("## 🏛️ 현장점검 관리 시스템 Login")

    if not APP_USER_ID or not (APP_PASSWORD or APP_PASSWORD_HASH):
        st.error("로그인 설정이 없습니다. APP_USER_ID와 APP_PASSWORD 또는 APP_PASSWORD_HASH를 설정해 주세요.")
        with st.expander("설정 예시 보기"):
            st.code(
                """# .streamlit/secrets.toml 예시
APP_USER_ID = "admin"
APP_PASSWORD_HASH = "여기에_sha256_해시값"
GEMINI_API_KEY = "여기에_Gemini_API_Key"

# 비밀번호 해시 생성 예시
# python -c "import hashlib; print(hashlib.sha256('원하는비밀번호'.encode()).hexdigest())"
""",
                language="toml",
            )
        return False

    with st.form("login_form"):
        user_id = st.text_input("아이디")
        password = st.text_input("비밀번호", type="password")
        if st.form_submit_button("접속하기"):
            if hmac.compare_digest(user_id, APP_USER_ID) and verify_password(password):
                st.session_state["logged_in"] = True
                st.rerun()
            else:
                st.error("아이디 또는 비밀번호가 일치하지 않습니다.")
    return False


def adjust_weekend(date_obj):
    wd = date_obj.weekday()
    if wd == 5:
        return date_obj + timedelta(days=2)
    if wd == 6:
        return date_obj + timedelta(days=1)
    return date_obj


def parse_date(date_str: str):
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").date()
    except Exception:
        return None


def add_step(site_data, site_name: str, date_obj, desc: str, memo: str = "", files=None, step_id: str = ""):
    if site_name not in site_data:
        site_data[site_name] = []
    site_data[site_name].append(
        {
            "id": step_id or make_step_id(),
            "date": date_obj,
            "desc": desc,
            "memo": memo or "",
            "files": files or [],
        }
    )


def load_data():
    site_data = {}
    if not DB_PATH.exists():
        return site_data

    try:
        with DB_PATH.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            if not header:
                return site_data
            header = [h.strip() for h in header]

            if "ID" in header and "현장명" in header:
                idx = {name: header.index(name) for name in header}
                for row in reader:
                    if not row:
                        continue
                    try:
                        site_name = row[idx["현장명"]].strip()
                        date_obj = parse_date(row[idx["날짜"]].strip())
                        desc = row[idx["업무명"]].strip()
                        memo = row[idx["메모"]] if idx.get("메모", 999) < len(row) else ""
                        files_str = row[idx["파일경로"]] if idx.get("파일경로", 999) < len(row) else ""
                        step_id = row[idx["ID"]].strip() if idx.get("ID", 999) < len(row) else ""
                    except Exception:
                        continue

                    if not site_name or not date_obj or not desc:
                        continue
                    files_list = [p for p in files_str.split("|") if p]
                    add_step(site_data, site_name, date_obj, desc, memo, files_list, step_id)
            else:
                # 구버전 CSV 호환: No, 현장명, 날짜, 업무명, 메모, 파일경로 또는 옆으로 긴 형식
                for row in reader:
                    if len(row) < 6:
                        continue
                    site_name = row[1].strip()
                    if not site_name:
                        continue
                    for i in range(2, len(row), 4):
                        if i + 3 >= len(row):
                            break
                        date_obj = parse_date(row[i].strip())
                        desc = row[i + 1].strip()
                        memo = row[i + 2]
                        files_str = row[i + 3]
                        if not date_obj or not desc:
                            continue
                        files_list = [p for p in files_str.split("|") if p]
                        add_step(site_data, site_name, date_obj, desc, memo, files_list)

        for name in site_data:
            site_data[name].sort(key=lambda x: x["date"])
    except Exception as e:
        st.error(f"데이터 로드 오류: {e}")
    return site_data


def save_data(site_data):
    try:
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        temp_path = DB_PATH.with_suffix(DB_PATH.suffix + ".tmp")
        backup_path = DB_PATH.with_suffix(DB_PATH.suffix + ".bak")

        with temp_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["No", "ID", "현장명", "날짜", "업무명", "메모", "파일경로"])
            row_num = 1
            for name in sorted(site_data.keys()):
                for step in sorted(site_data[name], key=lambda x: x["date"]):
                    step.setdefault("id", make_step_id())
                    files_str = "|".join(step.get("files", []))
                    writer.writerow(
                        [
                            row_num,
                            step["id"],
                            name,
                            step["date"].strftime("%Y-%m-%d"),
                            step.get("desc", ""),
                            step.get("memo", ""),
                            files_str,
                        ]
                    )
                    row_num += 1

        if DB_PATH.exists():
            shutil.copy2(DB_PATH, backup_path)
        os.replace(temp_path, DB_PATH)
    except Exception as e:
        st.error(f"저장 실패: {e}")


# ==========================================
# 📎 8. 첨부파일 UI
# ==========================================
def handle_file_uploads(selected_site: str, step: dict):
    step_id = step.setdefault("id", make_step_id())
    uploaded_files = st.file_uploader(
        "파일 업로드",
        accept_multiple_files=True,
        key=f"up_{step_id}",
        label_visibility="collapsed",
    )

    if not uploaded_files:
        return False

    has_new = False
    step.setdefault("files", [])

    for uf in uploaded_files:
        if uf.size > MAX_UPLOAD_MB * 1024 * 1024:
            st.error(f"{uf.name}: 파일 용량이 {MAX_UPLOAD_MB}MB를 초과하여 업로드하지 않았습니다.")
            continue

        # 핵심 수정: Streamlit은 파일 업로드 후 앱을 다시 실행함.
        # 기존 코드는 재실행될 때마다 uuid가 붙은 새 파일명을 만들었기 때문에
        # 같은 파일을 계속 새 파일로 판단하여 무한 새로고침처럼 보였음.
        # 파일 내용 해시 기반의 고정 경로를 사용하면 같은 파일은 한 번만 저장됨.
        file_bytes = uf.getvalue()
        file_digest = hashlib.sha256(file_bytes).hexdigest()
        save_path = deterministic_attachment_path(selected_site, step_id, uf.name, file_digest)
        storage_path = to_storage_path(save_path)

        if storage_path in step["files"] and save_path.exists():
            continue

        if not save_path.exists():
            save_path.parent.mkdir(parents=True, exist_ok=True)
            save_path.write_bytes(file_bytes)

        if storage_path not in step["files"]:
            step["files"].append(storage_path)
        has_new = True

        if save_path.suffix.lower() in {".hwp", ".hwpx"}:
            with st.spinner("한글(HWP/HWPX) 문서를 PDF로 자동 변환 중입니다..."):
                pdf_path = hwp_to_pdf(save_path)
                if pdf_path != save_path and pdf_path.exists():
                    pdf_storage_path = to_storage_path(pdf_path)
                    if pdf_storage_path not in step["files"]:
                        step["files"].append(pdf_storage_path)
                        has_new = True

    return has_new


def render_file_manager(selected_site: str, step: dict):
    st.markdown("**📂 첨부 파일 관리 (드래그 앤 드롭)**")
    changed = handle_file_uploads(selected_site, step)
    if changed:
        save_data(st.session_state.site_data)
        st.success("파일이 업로드되었습니다.")
        st.rerun()

    existing_files = []
    for file_path_value in step.get("files", []):
        if resolve_path(file_path_value).exists():
            existing_files.append(file_path_value)
    if existing_files != step.get("files", []):
        step["files"] = existing_files
        save_data(st.session_state.site_data)

    checked_files_to_download = []
    for file_path_value in step.get("files", []):
        file_path = resolve_path(file_path_value)
        file_name = file_path.name
        ext = file_path.suffix.lower()
        file_key = f"{step['id']}_{safe_slug(file_path_value, 120)}"
        chk_col, btn_col1, btn_col2, btn_col3 = st.columns([5, 2, 2, 2])

        with chk_col:
            is_checked = st.checkbox(f"📎 {file_name}", key=f"chk_{file_key}")
            if is_checked:
                checked_files_to_download.append(file_path)

        if ext in SUPPORTED_PREVIEW_EXTS:
            with btn_col1:
                if st.button("열기", key=f"view_{file_key}", help="새창에서 문서 보기", use_container_width=True):
                    show_file_dialog(file_path_value, file_name)
            with btn_col2:
                if ext in SUPPORTED_AI_EXTS and st.button("요약", key=f"ai_{file_key}", help="새창에서 공문서 양식으로 요약", use_container_width=True):
                    show_summary_dialog(file_path_value, file_name)
        else:
            with btn_col1:
                st.write("")
            with btn_col2:
                st.write("")

        with btn_col3:
            if st.button("삭제", key=f"delf_{file_key}", use_container_width=True):
                if file_path_value in step.get("files", []):
                    step["files"].remove(file_path_value)
                try:
                    file_path.unlink()
                except Exception:
                    pass
                save_data(st.session_state.site_data)
                st.rerun()

    if checked_files_to_download:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for fpath in checked_files_to_download:
                if fpath.exists():
                    zip_file.write(fpath, arcname=fpath.name)
        st.download_button(
            label=f"💾 체크된 파일 {len(checked_files_to_download)}개 다운로드 (.zip)",
            data=zip_buffer.getvalue(),
            file_name=f"첨부파일_다운로드_{date.today().strftime('%Y%m%d')}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
            key=f"zip_dl_{step['id']}",
        )


# ==========================================
# 🌐 9. 메인 앱 UI
# ==========================================
def main():
    if not check_password():
        return

    st.markdown(
        """
        <style>
        .stMarkdown pre { white-space: pre-wrap !important; word-wrap: break-word !important; overflow-wrap: break-word !important; }
        .stMarkdown code { white-space: pre-wrap !important; word-break: keep-all !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    if "site_data" not in st.session_state:
        st.session_state.site_data = load_data()
    if "cal_year" not in st.session_state:
        st.session_state.cal_year = date.today().year
    if "cal_month" not in st.session_state:
        st.session_state.cal_month = date.today().month
    if "current_page" not in st.session_state:
        st.session_state.current_page = 1

    st.title("🏗️ 현장점검 관리 시스템")

    with st.sidebar:
        st.header("➕ 신규 건설현장 등록")
        with st.form("add_project_form"):
            new_site_name = st.text_input("건설현장명")
            start_date = st.date_input("점검 예정일", value=date.today())
            if st.form_submit_button("초기 점검일정 생성"):
                site_name = new_site_name.strip()
                if not site_name:
                    st.warning("건설현장명을 입력해 주세요.")
                elif site_name in st.session_state.site_data:
                    st.warning("이미 등록된 현장명입니다.")
                else:
                    st.session_state.site_data[site_name] = [
                        {"id": make_step_id(), "date": start_date, "desc": "현장점검 실시", "memo": "", "files": []}
                    ]
                    save_data(st.session_state.site_data)
                    st.rerun()

        st.divider()
        st.header("📋 건설현장 선택")
        search_query = st.text_input("🔍 현장명 검색", placeholder="현장 이름을 입력하세요")
        all_sites = sorted(st.session_state.site_data.keys())
        filtered_sites = [site for site in all_sites if search_query.lower() in site.lower()] if search_query else all_sites
        site_options = ["전체 현장"] + filtered_sites

        with st.container(height=300, border=True):
            selected_site = st.radio("일정을 볼 현장 선택", site_options, label_visibility="collapsed")

        if st.session_state.get("last_selected_site") != selected_site:
            st.session_state.current_page = 1
            st.session_state.last_selected_site = selected_site

        if selected_site != "전체 현장":
            st.markdown("---")
            confirm_delete = st.checkbox("삭제 확인", key=f"confirm_delete_site_{safe_slug(selected_site)}")
            if st.button("🗑️ 현재 건설현장 삭제", type="primary", use_container_width=True, disabled=not confirm_delete):
                del st.session_state.site_data[selected_site]
                save_data(st.session_state.site_data)
                st.rerun()

        st.divider()
        if DB_PATH.exists() or ATTACH_DIR.exists():
            st.download_button(
                "📦 전체 데이터 백업 다운로드",
                data=make_backup_zip(),
                file_name=f"현장점검_백업_{date.today().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                use_container_width=True,
            )

    st.subheader("🗓️ 현장점검 일정 캘린더")
    c1, c2, c3 = st.columns([1, 4, 1])
    with c1:
        if st.button("◀ 이전 달", use_container_width=True):
            if st.session_state.cal_month == 1:
                st.session_state.cal_month = 12
                st.session_state.cal_year -= 1
            else:
                st.session_state.cal_month -= 1
            st.rerun()
    with c2:
        st.markdown(f"<h3 style='text-align:center;'>{st.session_state.cal_year}년 {st.session_state.cal_month}월</h3>", unsafe_allow_html=True)
    with c3:
        if st.button("다음 달 ▶", use_container_width=True):
            if st.session_state.cal_month == 12:
                st.session_state.cal_month = 1
                st.session_state.cal_year += 1
            else:
                st.session_state.cal_month += 1
            st.rerun()

    st.markdown(
        render_html_calendar(st.session_state.site_data, st.session_state.cal_year, st.session_state.cal_month, selected_site),
        unsafe_allow_html=True,
    )
    st.divider()

    if selected_site == "전체 현장":
        st.info("좌측에서 특정 건설현장을 선택하면 세부 일정과 첨부파일을 관리할 수 있습니다.")
        return

    st.subheader(f"📂 [{selected_site}] 세부 일정 및 파일 관리")
    steps = st.session_state.site_data[selected_site]

    add_col1, add_col2 = st.columns(2)
    with add_col1:
        with st.expander("📌 단순 일정 수동 추가"):
            e1, e2 = st.columns([1, 2])
            with e1:
                custom_date = st.date_input("날짜", key="c_date")
            with e2:
                custom_desc = st.text_input("업무 내용", key="c_desc")
            if st.button("일정 끼워넣기", use_container_width=True):
                desc = custom_desc.strip()
                if not desc:
                    st.warning("업무 내용을 입력해 주세요.")
                else:
                    steps.append({"id": make_step_id(), "date": adjust_weekend(custom_date), "desc": desc, "memo": "", "files": []})
                    steps.sort(key=lambda x: x["date"])
                    save_data(st.session_state.site_data)
                    st.rerun()

    with add_col2:
        with st.expander("🚨 벌점/과태료 발생 시 (후속 행정절차 자동 생성)"):
            st.markdown(
                "<span style='font-size:14px;'>현장점검 결과 벌점 등이 부과된 경우, 점검일을 기준으로 이후 법정 행정절차 일정을 자동 생성합니다.</span>",
                unsafe_allow_html=True,
            )
            base_step = next((s for s in steps if "현장점검" in s.get("desc", "")), None)
            default_base_date = base_step["date"] if base_step else date.today()
            penalty_base_date = st.date_input("기준일 (현장점검일)", value=default_base_date)

            if st.button("⚠️ 후속 행정절차 일괄 생성", type="primary", use_container_width=True):
                existing_descs = [s.get("desc", "") for s in steps]
                if "확인서 이의제기 접수" in existing_descs:
                    st.warning("이미 벌점 부과 후속 일정이 생성되어 있습니다.")
                else:
                    curr = penalty_base_date
                    for days, desc in PENALTY_INTERVALS:
                        curr = adjust_weekend(curr + timedelta(days=days))
                        steps.append({"id": make_step_id(), "date": curr, "desc": desc, "memo": "", "files": []})
                    steps.sort(key=lambda x: x["date"])
                    save_data(st.session_state.site_data)
                    st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    total_items = len(steps)
    total_pages = (total_items - 1) // ITEMS_PER_PAGE + 1 if total_items > 0 else 1
    if st.session_state.current_page > total_pages:
        st.session_state.current_page = total_pages

    start_idx = (st.session_state.current_page - 1) * ITEMS_PER_PAGE
    end_idx = start_idx + ITEMS_PER_PAGE
    current_page_steps = steps[start_idx:end_idx]

    p_col1, p_col2, p_col3 = st.columns([1, 3, 1])
    with p_col1:
        if st.button("◀ 이전 페이지", disabled=(st.session_state.current_page == 1), use_container_width=True):
            st.session_state.current_page -= 1
            st.rerun()
    with p_col2:
        st.markdown(
            f"<div style='text-align:center; font-size:16px; padding-top:5px;'><b>페이지 {st.session_state.current_page} / {total_pages}</b> (총 {total_items}건)</div>",
            unsafe_allow_html=True,
        )
    with p_col3:
        if st.button("다음 페이지 ▶", disabled=(st.session_state.current_page == total_pages), use_container_width=True):
            st.session_state.current_page += 1
            st.rerun()

    for i, step in enumerate(current_page_steps):
        actual_idx = start_idx + i
        step.setdefault("id", make_step_id())
        step_id = step["id"]

        with st.container(border=True):
            c1, c2, c3 = st.columns([2, 5, 4])

            with c1:
                new_date = st.date_input("기한", value=step["date"], key=f"date_{step_id}")
                new_desc = st.text_input("업무명", value=step.get("desc", ""), key=f"desc_{step_id}")
                if new_date != step["date"] or new_desc != step.get("desc", ""):
                    if new_desc.strip():
                        steps[actual_idx]["date"] = new_date
                        steps[actual_idx]["desc"] = new_desc.strip()
                        steps.sort(key=lambda x: x["date"])
                        save_data(st.session_state.site_data)
                        st.rerun()
                    else:
                        st.warning("업무명은 비워둘 수 없습니다.")

                if st.button("❌ 일정 삭제", key=f"del_step_{step_id}"):
                    steps.pop(actual_idx)
                    save_data(st.session_state.site_data)
                    st.rerun()

            with c2:
                new_memo = st.text_area("📝 메모", value=step.get("memo", ""), height=100, key=f"memo_{step_id}")
                if new_memo != step.get("memo", ""):
                    steps[actual_idx]["memo"] = new_memo
                    save_data(st.session_state.site_data)

            with c3:
                render_file_manager(selected_site, step)


if __name__ == "__main__":
    main()
