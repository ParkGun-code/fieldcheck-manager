import os
# 특정 네트워크에서 gRPC 통신 타임아웃 방지
os.environ['GRPC_DNS_RESOLVER'] = 'native'

import streamlit as st
import streamlit.components.v1 as components
import csv
import calendar
import base64
from datetime import date, datetime, timedelta
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from google import genai
import time
import shutil
import zipfile
import io
import pandas as pd

import json
import html as html_lib
from textwrap import dedent

# ==========================================
# 🛑 HWP -> PDF 변환용 라이브러리 (윈도우 전용)
# ==========================================
try:
    import win32com.client as win32
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

def hwp_to_pdf(hwp_path):
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
        pythoncom.CoUninitialize()

# ==========================================
# ⚙️ 1. 기본 설정 및 전역 변수
# ==========================================
st.set_page_config(page_title="건설현장 벌점 통합 관리 웹", page_icon="🏛️", layout="wide")

SHARED_USER_ID = "molitdj"
SHARED_PASSWORD = "eowjscjd1!"

GEMINI_API_KEY = "AIzaSyBRSKqPy-IVLqAqICwaJmli5YKifNcRdoA"  
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

DB_FILENAME = "penalty_database.csv"
ATTACH_DIR = "attachments"

ITEMS_PER_PAGE = 10 

if not os.path.exists(ATTACH_DIR):
    os.makedirs(ATTACH_DIR)

PENALTY_INTERVALS = [
    (30, "확인서 이의제기 접수"),
    (14, "확인서 이의제기 의견 통보"),
    (14, "벌점 사전부과 통보"),
    (15, "벌점 사전부과 통보 의견제출 마감"), 
    (15, "벌점 사전부과 의견 검토회의"), 
    (1, "벌점 사전부과 의견 검토회의 결과 통보 및 벌점 부과"),
    (30, "벌점 부과 이의제기 접수"), 
    (40, "벌점 심의위원회 개최 및 최종 결과 통보")
]

# ==========================================
# 🖱️ 2. 새창(다이얼로그) 마우스 드래그 기능 주입
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
# 🔐 3. 새창(다이얼로그) 팝업 UI 구현
# ==========================================
@st.dialog("✨ AI 심의안건 보고서 작성 (새창)", width="large")
def show_summary_dialog(file_path, file_name):
    make_dialog_draggable() 
    st.markdown(f"### 📄 [{file_name}] 분석 결과")
    with st.spinner("AI가 공무원 양식으로 보고서를 작성 중입니다..."):
        st.write_stream(get_ai_summary_stream(file_path))
    if st.button("닫기", type="primary"):
        st.rerun()

@st.dialog("📄 첨부 문서 뷰어 (새창)", width="large")
def show_file_dialog(file_path, file_name):
    make_dialog_draggable() 
    st.markdown(f"### 📎 {file_name}")
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext in ['.png', '.jpg', '.jpeg']:
        st.image(file_path, use_column_width=True)
    elif ext == '.pdf':
        with open(file_path, "rb") as f:
            base64_pdf = base64.b64encode(f.read()).decode('utf-8')
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700" type="application/pdf"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)
    elif ext == '.txt':
        with open(file_path, "r", encoding="utf-8") as f:
            st.text_area("문서 내용", f.read(), height=500)
    else:
        st.warning("⚠️ 웹 미리보기를 지원하지 않는 형식입니다. 리스트의 체크박스를 통해 다운로드해 주세요.")

# ==========================================
# 📅 4. 중앙 달력 렌더링 함수
# ==========================================
def render_html_calendar(site_data, year, month, selected_site=None):
    cal = calendar.monthcalendar(year, month)

    html_parts = []

    html_parts.append(dedent("""
    <style>
        body {
            font-family: 'Malgun Gothic', sans-serif;
            margin: 0;
            padding: 0;
        }

        .cal-cell-scroll::-webkit-scrollbar {
            width: 6px;
        }

        .cal-cell-scroll::-webkit-scrollbar-thumb {
            background: #c1c1c1;
            border-radius: 3px;
        }

        .cal-cell-scroll::-webkit-scrollbar-track {
            background: #f1f1f1;
        }

        .calendar-table {
            width: 100%;
            table-layout: fixed;
            border-collapse: collapse;
            text-align: left;
            background-color: #ffffff;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }

        .calendar-table th {
            border: 1px solid #ddd;
            padding: 10px;
            text-align: center;
            background-color: #f0f2f6;
            font-size: 16px;
        }

        .calendar-table td {
            border: 1px solid #ddd;
            padding: 0;
            vertical-align: top;
        }

        .calendar-cell {
            display: flex;
            flex-direction: column;
            height: 130px;
            padding: 6px;
        }

        .calendar-day {
            font-size: 16px;
            margin-bottom: 2px;
            flex-shrink: 0;
        }

        .calendar-event {
            cursor: pointer;
            font-size: 12px;
            margin-top: 4px;
            padding: 4px;
            border-left: 3px solid;
            border-radius: 3px;
            transition: 0.2s;
            word-break: keep-all;
        }

        .calendar-event:hover {
            filter: brightness(0.95);
        }

        #customModal {
            display: none;
            position: fixed;
            z-index: 9999;
            left: 50%;
            top: 50%;
            transform: translate(-50%, -50%);
            background: #ffffff;
            padding: 20px;
            border: 1px solid #ccc;
            border-radius: 12px;
            box-shadow: 0 8px 16px rgba(0,0,0,0.3);
            width: 450px;
            max-width: 90%;
        }

        #modalSiteName {
            margin-top: 0;
            color: #1f2937;
            border-bottom: 2px solid #e5e7eb;
            padding-bottom: 10px;
        }

        #modalDesc {
            font-weight: bold;
            color: #2563eb;
            margin-bottom: 5px;
        }

        #modalMemo {
            white-space: pre-wrap;
            font-family: 'Malgun Gothic', sans-serif;
            font-size: 14px;
            color: #4b5563;
            background: #f3f4f6;
            padding: 15px;
            border-radius: 8px;
            line-height: 1.5;
            max-height: 300px;
            overflow-y: auto;
        }

        .modal-close-btn {
            float: right;
            padding: 8px 16px;
            cursor: pointer;
            background: #3b82f6;
            color: white;
            border: none;
            border-radius: 6px;
            font-weight: bold;
        }
    </style>

    <div id="customModal">
        <h3 id="modalSiteName">현장명</h3>
        <p id="modalDesc"></p>
        <pre id="modalMemo"></pre>
        <button class="modal-close-btn" onclick="closeSiteInfo()">닫기</button>
        <div style="clear: both;"></div>
    </div>

    <script>
        function openSiteInfo(site, desc, memo) {
            document.getElementById('modalSiteName').innerText = site || '';
            document.getElementById('modalDesc').innerText = desc || '';
            document.getElementById('modalMemo').innerText = memo || '';
            document.getElementById('customModal').style.display = 'block';
        }

        function closeSiteInfo() {
            document.getElementById('customModal').style.display = 'none';
        }
    </script>
    """))

    html_parts.append("<table class='calendar-table'>")
    html_parts.append("<tr>")

    for d in ['월', '화', '수', '목', '금', '토', '일']:
        html_parts.append(f"<th>{d}</th>")

    html_parts.append("</tr>")

    for week in cal:
        html_parts.append("<tr>")

        for day in week:
            if day == 0:
                html_parts.append("<td style='background-color:#fafafa;'></td>")
                continue

            current_date = date(year, month, day)
            bg_color = "#e6f2ff" if current_date == date.today() else "#ffffff"

            day_events = []

            for site, steps in site_data.items():
                if selected_site and selected_site != "전체 현장" and site != selected_site:
                    continue

                for step in steps:
                    if step.get('date') != current_date:
                        continue

                    desc = str(step.get('desc', ''))
                    memo = str(step.get('memo', ''))

                    if "우기" in desc:
                        event_bg, event_border = "#dbeafe", "#1d4ed8"
                    elif "상시" in desc or "월점검" in desc:
                        event_bg, event_border = "#d1fae5", "#047857"
                    elif "현장점검" in desc:
                        event_bg, event_border = "#f3f4f6", "#374151"
                    else:
                        event_bg, event_border = "#fef3c7", "#b45309"

                    site_label = html_lib.escape(site)
                    desc_label = html_lib.escape(desc)

                    site_arg = html_lib.escape(json.dumps(site, ensure_ascii=False), quote=True)
                    desc_arg = html_lib.escape(json.dumps(desc, ensure_ascii=False), quote=True)
                    memo_arg = html_lib.escape(json.dumps(memo, ensure_ascii=False), quote=True)

                    event_html = f"""
                    <div
                        class="calendar-event"
                        onclick="openSiteInfo({site_arg}, {desc_arg}, {memo_arg})"
                        style="background-color:{event_bg}; border-left-color:{event_border};"
                    >
                        <b>[{site_label}]</b> {desc_label}
                    </div>
                    """

                    day_events.append(dedent(event_html))

            max_display = 20
            events_html = ""

            for i, event in enumerate(day_events):
                if i < max_display:
                    events_html += event
                elif i == max_display:
                    remaining = len(day_events) - max_display
                    events_html += f"""
                    <div style="
                        font-size:12px;
                        margin-top:4px;
                        padding:4px;
                        text-align:center;
                        background-color:#e0e0e0;
                        color:#555;
                        border-radius:3px;
                    ">
                        ...외 {remaining}건 더 있음
                    </div>
                    """
                    break

            html_parts.append(f"<td style='background-color:{bg_color};'>")
            html_parts.append("<div class='calendar-cell'>")
            html_parts.append(f"<strong class='calendar-day'>{day}</strong>")
            html_parts.append("<div class='cal-cell-scroll' style='flex-grow:1; overflow-y:auto; padding-right:4px;'>")
            html_parts.append(events_html)
            html_parts.append("</div>")
            html_parts.append("</div>")
            html_parts.append("</td>")

        html_parts.append("</tr>")

    html_parts.append("</table>")

    return "".join(html_parts)

# ==========================================
# 🤖 5. 공무원 양식 AI 요약 프롬프트 적용
# ==========================================
def get_ai_summary_stream(file_path):
    yield "🔄 AI 분석 엔진을 초기화하는 중입니다...\n\n"
    client = genai.Client(api_key=GEMINI_API_KEY)
    ext = os.path.splitext(file_path)[1].lower()
    prompt = """당신은 관공서(국토관리청 등)의 벌점심의위원회 또는 현장점검 결과 보고서를 작성하는 전문 행정관입니다.
    제공된 문서를 철저히 분석하여, 아래의 [공식 심의안건 보고서 양식]에 맞추어 완벽하게 요약 및 재작성하십시오.

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
    3. 점검자 검토 의견
    - [문서 상 행정관/점검자의 검토 결과 요약]
    4. 심의 요구사항 (또는 종합 결론)
    - [최종 요약 및 심의/조치 필요사항]

    [주의사항]
    - 반드시 명조체 느낌의 정중하고 딱딱한 공문서 개조식 어투(~함, ~임)를 사용하십시오.
    - 문서에 없는 내용을 절대 임의로 지어내지 마십시오. 정보가 부족한 항목은 과감히 생략하십시오.
    """
    uploaded_file = None
    safe_filepath = None
    try:
        if ext in ['.pdf', '.png', '.jpg', '.jpeg', '.txt']:
            yield "🚀 문서를 안전하게 변환하여 구글 서버로 전송합니다...\n\n"
            safe_filename = f"temp_ai_upload_{int(time.time())}{ext}"
            safe_filepath = os.path.join(ATTACH_DIR, safe_filename)
            shutil.copy2(file_path, safe_filepath)
            uploaded_file = client.files.upload(file=safe_filepath)
            if ext == '.pdf':
                yield "📄 PDF 문서를 스캔하고 있습니다. (약 5~10초 소요)...\n\n"
                while True:
                    file_info = client.files.get(name=uploaded_file.name)
                    state_str = str(file_info.state).upper()
                    if "ACTIVE" in state_str:    
                        break
                    elif "FAILED" in state_str:  
                        yield "❌ 구글 서버에서 PDF 문서를 읽는 데 실패했습니다."
                        return
                    time.sleep(2)  
            yield "💡 스캔 완료! 보고서 작성을 시작합니다...\n\n---\n\n"
            response_stream = client.models.generate_content_stream(model='gemini-2.5-flash', contents=[prompt, uploaded_file])
            for chunk in response_stream:
                if chunk.text: yield chunk.text
        else:
            yield "⚠️ 스트림릿 환경에서는 PDF, TXT, 이미지 요약만 안정적으로 지원합니다."
    except Exception as e:
        error_msg = str(e)
        if "503" in error_msg or "high demand" in error_msg.lower():
            yield "\n\n❌ 현재 구글 AI 서버에 사용자가 몰려 일시적으로 혼잡한 상태입니다.\n(약 1~2분 뒤에 다시 요약 버튼을 눌러주세요.)"
        else:
            yield f"\n\n❌ AI 분석 중 오류가 발생했습니다.\n상세: {error_msg}"
    finally:
        if uploaded_file:
            try:
                client.files.delete(name=uploaded_file.name)
            except:
                pass
        if safe_filepath and os.path.exists(safe_filepath):
            try:
                os.remove(safe_filepath)
            except:
                pass

# ==========================================
# 💾 6. 데이터 처리 (엑셀/CSV 파싱 포함)
# ==========================================
def check_password():
    if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
    if st.session_state["logged_in"]: return True
    st.markdown("## 🏛️ 건설현장 벌점 통합 관리 시스템 Login")
    with st.form("login_form"):
        user_id = st.text_input("아이디")
        password = st.text_input("비밀번호", type="password")
        if st.form_submit_button("접속하기"):
            if user_id == SHARED_USER_ID and password == SHARED_PASSWORD:
                st.session_state["logged_in"] = True
                st.rerun()
            else:
                st.error("아이디 또는 비밀번호가 일치하지 않습니다.")
    return False

def adjust_weekend(date_obj):
    wd = date_obj.weekday()
    if wd == 5: return date_obj + timedelta(days=2)
    if wd == 6: return date_obj + timedelta(days=1)
    return date_obj

def load_data():
    site_data = {}
    if not os.path.exists(DB_FILENAME): return site_data
    try:
        with open(DB_FILENAME, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            header = next(reader, None)
            if not header: return site_data
            for row in reader:
                if len(row) < 6: continue
                name = row[1].strip()
                if not name: continue
                if name not in site_data: site_data[name] = []
                date_str, desc, memo, files_str = row[2], row[3], row[4], row[5]
                if not date_str or not desc: continue
                files_list = files_str.split("|") if files_str else []
                try: 
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
                    site_data[name].append({"date": date_obj, "desc": desc, "memo": memo, "files": files_list})
                except:
                    continue
        for name in site_data: site_data[name].sort(key=lambda x: x['date'])
    except Exception as e: st.error(f"데이터 로드 오류: {e}")
    return site_data

def save_data(site_data):
    try:
        with open(DB_FILENAME, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['No', '현장명', '날짜', '업무명', '메모', '파일경로'])
            row_num = 1
            for name in sorted(site_data.keys()):
                for step in site_data[name]:
                    files_str = "|".join(step.get('files', []))
                    writer.writerow([row_num, name, step['date'].strftime('%Y-%m-%d'), step['desc'], step.get('memo', ''), files_str])
                    row_num += 1
    except Exception as e:
        st.error(f"저장 실패: {e}")

# 💡 엑셀/CSV 일괄 등록 로직
def process_excel_schedule(file):
    try:
        if file.name.lower().endswith('.csv'):
            try:
                df = pd.read_csv(file, encoding='utf-8-sig', header=None)
            except:
                df = pd.read_csv(file, encoding='cp949', header=None)
        else:
            df = pd.read_excel(file, header=None)
        
        header_idx = -1
        for i, row in df.iterrows():
            row_str = "".join(str(val) for val in row.values)
            if "공사명" in row_str and "점검예정일" in row_str:
                header_idx = i
                break
                
        if header_idx != -1:
            df.columns = df.iloc[header_idx]
            df = df.iloc[header_idx + 1:]
        else:
            st.error("엑셀 파일 양식이 맞지 않습니다. '공사명' 및 '점검예정일' 열을 찾을 수 없습니다.")
            return

        success_count = 0
        for idx, row in df.iterrows():
            site_name = str(row.get('공사명', '')).strip()
            if not site_name or site_name == 'nan': continue
            
            date_raw = str(row.get('점검예정일', '')).strip()
            team = str(row.get('담당조', '')).strip()
            client = str(row.get('발주처\n(인·허가 기관)', row.get('발주처', ''))).strip()
            status = str(row.get('공사진행상태', '')).strip()
            builder = str(row.get('시공회사명', '')).strip()
            supervisor = str(row.get('감리회사명', '')).strip()
            manager = str(row.get('성명', '')).strip()
            phone = str(row.get('전화번호', '')).strip()
            
            try:
                parts = [p for p in date_raw.split('.') if p.strip().isdigit()]
                if len(parts) >= 2:
                    plan_date = date(2026, int(parts[0]), int(parts[1]))
                else:
                    continue 
            except:
                continue 
            
            inspection_type = "상시점검" if "상시" in file.name else "우기대비 점검"
            team_str = f"[{team}]" if team and team != 'nan' else ""
            desc = f"{team_str} {inspection_type}".strip()
            
            memo_lines = []
            if client and client != 'nan': memo_lines.append(f"🏢 발주처: {client}")
            if builder and builder != 'nan': memo_lines.append(f"👷 시공사: {builder}")
            if supervisor and supervisor != 'nan': memo_lines.append(f"🔍 감리사: {supervisor}")
            if manager and manager != 'nan': memo_lines.append(f"👤 현장대리인: {manager} ({phone})")
            if status and status != 'nan': memo_lines.append(f"📌 공사진행상태: {status}")
            
            memo = "\n".join(memo_lines)
            
            if site_name not in st.session_state.site_data:
                st.session_state.site_data[site_name] = []
                
            existing_descs = [s['desc'] for s in st.session_state.site_data[site_name] if s['date'] == plan_date]
            if desc not in existing_descs:
                st.session_state.site_data[site_name].append({
                    "date": plan_date,
                    "desc": desc,
                    "memo": memo,
                    "files": []
                })
                st.session_state.site_data[site_name].sort(key=lambda x: x['date'])
                success_count += 1
                
        if success_count > 0:
            save_data(st.session_state.site_data)
            st.success(f"총 {success_count}건의 점검 일정이 엑셀에서 성공적으로 등록되었습니다!")
            time.sleep(1) 
            st.rerun()
        else:
            st.warning("등록할 유효한 일정이 없습니다. (날짜가 '06.17.' 형식인지 확인하세요)")
            
    except Exception as e:
        st.error(f"엑셀 처리 중 오류 발생: {e}")

# ==========================================
# 🌐 7. 메인 앱 UI 
# ==========================================
def main():
    if not check_password(): return

    st.markdown("""
    <style>
    .stMarkdown pre {
        white-space: pre-wrap !important;
        word-wrap: break-word !important;
        overflow-wrap: break-word !important;
    }
    </style>
    """, unsafe_allow_html=True)

    if "site_data" not in st.session_state: st.session_state.site_data = load_data()
    if "cal_year" not in st.session_state: st.session_state.cal_year = date.today().year
    if "cal_month" not in st.session_state: st.session_state.cal_month = date.today().month

    st.title("🏗️ 건설현장 벌점 및 문서 통합 관리 시스템")

    with st.sidebar:
        st.header("📁 엑셀 일정 일괄 등록")
        excel_file = st.file_uploader("엑셀/CSV 파일 업로드", type=['csv', 'xlsx', 'xls'])
        if st.button("🚀 일정 자동 등록하기", type="primary", use_container_width=True):
            if excel_file:
                process_excel_schedule(excel_file)
            else:
                st.warning("먼저 엑셀 또는 CSV 파일을 올려주세요.")
        
        st.divider()

        st.header("➕ 개별 프로젝트 등록")
        with st.form("add_project_form"):
            new_site_name = st.text_input("프로젝트(현장)명")
            start_date = st.date_input("점검 예정일", value=date.today())
            if st.form_submit_button("초기 점검일정 생성"):
                if new_site_name and new_site_name not in st.session_state.site_data:
                    st.session_state.site_data[new_site_name] = [
                        {"date": start_date, "desc": "현장점검 실시", "memo": "", "files": []}
                    ]
                    save_data(st.session_state.site_data)
                    st.rerun()

        st.divider()
        st.header("📋 프로젝트 선택")
        search_query = st.text_input("🔍 현장명 검색", placeholder="현장 이름을 입력하세요")
        all_sites = sorted(list(st.session_state.site_data.keys()))
        if search_query:
            filtered_sites = [site for site in all_sites if search_query.lower() in site.lower()]
            site_options = ["전체 현장"] + filtered_sites
        else:
            site_options = ["전체 현장"] + all_sites
            
        with st.container(height=300, border=True):
            selected_site = st.radio("일정을 볼 현장 선택", site_options, label_visibility="collapsed")
        
        st.markdown("<br>", unsafe_allow_html=True)
        if selected_site != "전체 현장":
            if st.button("🗑️ 현재 프로젝트 삭제", type="primary", use_container_width=True):
                del st.session_state.site_data[selected_site]
                save_data(st.session_state.site_data)
                st.rerun()

    st.subheader("🗓️ 프로젝트 전체 일정 캘린더")
    c1, c2, c3 = st.columns([1, 4, 1])
    with c1:
        if st.button("◀ 이전 달", use_container_width=True):
            if st.session_state.cal_month == 1:
                st.session_state.cal_month = 12
                st.session_state.cal_year -= 1
            else: st.session_state.cal_month -= 1
            st.rerun()
    with c2:
        st.markdown(f"<h3 style='text-align:center;'>{st.session_state.cal_year}년 {st.session_state.cal_month}월</h3>", unsafe_allow_html=True)
    with c3:
        if st.button("다음 달 ▶", use_container_width=True):
            if st.session_state.cal_month == 12:
                st.session_state.cal_month = 1
                st.session_state.cal_year += 1
            else: st.session_state.cal_month += 1
            st.rerun()

    calendar_html = render_html_calendar(
        st.session_state.site_data,
        st.session_state.cal_year,
        st.session_state.cal_month,
        selected_site
    )

    components.html(calendar_html, height=900, scrolling=True)
		
	st.divider()

    if selected_site != "전체 현장":
        st.subheader(f"📂 [{selected_site}] 세부 일정 및 파일 관리")
        steps = st.session_state.site_data[selected_site]

        add_col1, add_col2 = st.columns(2)
        
        with add_col1:
            with st.expander("📌 단순 일정 수동 추가"):
                e1, e2 = st.columns([1, 2])
                with e1: custom_date = st.date_input("날짜", key="c_date")
                with e2: custom_desc = st.text_input("업무 내용", key="c_desc")
                if st.button("일정 끼워넣기", use_container_width=True):
                    steps.append({"date": adjust_weekend(custom_date), "desc": custom_desc, "memo": "", "files": []})
                    steps.sort(key=lambda x: x['date'])
                    save_data(st.session_state.site_data)
                    st.rerun()

        with add_col2:
            with st.expander("🚨 벌점/과태료 발생 시 (후속 행정절차 자동 생성)"):
                st.markdown("<span style='font-size:14px;'>현장점검 결과 벌점 등이 부과된 경우, 점검일을 기준으로 이후의 모든 법정 행정절차 일정을 자동으로 덧붙입니다.</span>", unsafe_allow_html=True)
                base_step = next((s for s in steps if "현장점검" in s['desc']), None)
                default_base_date = base_step['date'] if base_step else date.today()
                penalty_base_date = st.date_input("기준일 (현장점검일)", value=default_base_date)
                if st.button("⚠️ 후속 행정절차 일괄 생성", type="primary", use_container_width=True):
                    existing_descs = [s['desc'] for s in steps]
                    if "확인서 이의제기 접수" in existing_descs:
                        st.warning("이미 벌점 부과 후속 일정이 생성되어 있습니다.")
                    else:
                        curr = penalty_base_date
                        for days, desc in PENALTY_INTERVALS:
                            curr = adjust_weekend(curr + timedelta(days=days))
                            steps.append({"date": curr, "desc": desc, "memo": "", "files": []})
                        steps.sort(key=lambda x: x['date'])
                        save_data(st.session_state.site_data)
                        st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        total_items = len(steps)
        total_pages = (total_items - 1) // ITEMS_PER_PAGE + 1 if total_items > 0 else 1
        
        if "current_page" not in st.session_state: st.session_state.current_page = 1
        if st.session_state.current_page > total_pages: st.session_state.current_page = total_pages

        start_idx = (st.session_state.current_page - 1) * ITEMS_PER_PAGE
        end_idx = start_idx + ITEMS_PER_PAGE
        current_page_steps = steps[start_idx:end_idx]

        p_col1, p_col2, p_col3 = st.columns([1, 3, 1])
        with p_col1:
            if st.button("◀ 이전 페이지", disabled=(st.session_state.current_page == 1), use_container_width=True):
                st.session_state.current_page -= 1
                st.rerun()
        with p_col2:
            st.markdown(f"<div style='text-align:center; font-size:16px; padding-top:5px;'><b>페이지 {st.session_state.current_page} / {total_pages}</b> (총 {total_items}건)</div>", unsafe_allow_html=True)
        with p_col3:
            if st.button("다음 페이지 ▶", disabled=(st.session_state.current_page == total_pages), use_container_width=True):
                st.session_state.current_page += 1
                st.rerun()

        for i, step in enumerate(current_page_steps):
            actual_idx = start_idx + i  
            with st.container(border=True):
                c1, c2, c3 = st.columns([2, 5, 4])
                with c1:
                    new_date = st.date_input("기한", value=step['date'], key=f"date_{actual_idx}")
                    new_desc = st.text_input("업무명", value=step['desc'], key=f"desc_{actual_idx}")
                    if new_date != step['date'] or new_desc != step['desc']:
                        steps[actual_idx]['date'] = new_date
                        steps[actual_idx]['desc'] = new_desc
                        steps.sort(key=lambda x: x['date'])
                        save_data(st.session_state.site_data)
                        st.rerun()
                    if st.button("❌ 삭제", key=f"del_step_{actual_idx}"):
                        steps.pop(actual_idx)
                        save_data(st.session_state.site_data)
                        st.rerun()
                with c2:
                    new_memo = st.text_area("📝 메모", value=step.get('memo', ''), height=100, key=f"memo_{actual_idx}")
                    if new_memo != step.get('memo', ''):
                        steps[actual_idx]['memo'] = new_memo
                        save_data(st.session_state.site_data)
                with c3:
                    st.markdown("**📂 첨부 파일 관리 (드래그 앤 드롭)**")
                    uploaded_files = st.file_uploader("파일 업로드", accept_multiple_files=True, key=f"up_{actual_idx}", label_visibility="collapsed")
                    if uploaded_files:
                        has_new = False
                        for uf in uploaded_files:
                            original_path = os.path.join(ATTACH_DIR, f"{selected_site}_{uf.name}")
                            if original_path not in steps[actual_idx].get('files', []):
                                with open(original_path, "wb") as f: f.write(uf.getbuffer())
                                steps[actual_idx].setdefault('files', []).append(original_path)
                                has_new = True
                                if original_path.lower().endswith(".hwp"):
                                    with st.spinner("🔄 한글(HWP) 문서를 PDF로 자동 변환 중입니다... (최대 10초 소요)"):
                                        pdf_path = hwp_to_pdf(original_path)
                                        if pdf_path != original_path and os.path.exists(pdf_path):
                                            if pdf_path not in steps[actual_idx]['files']:
                                                steps[actual_idx]['files'].append(pdf_path)
                        if has_new:
                            save_data(st.session_state.site_data)
                            st.rerun()
                    existing_files = []
                    needs_sync = False
                    for file_path in steps[actual_idx].get('files', []):
                        if os.path.exists(file_path): existing_files.append(file_path)
                        else: needs_sync = True
                    if needs_sync:
                        steps[actual_idx]['files'] = existing_files
                        save_data(st.session_state.site_data)
                    checked_files_to_download = []
                    for file_path in steps[actual_idx].get('files', []):
                        file_name = os.path.basename(file_path)
                        ext = file_name.lower().split('.')[-1]
                        chk_col, btn_col1, btn_col2, btn_col3 = st.columns([5, 2, 2, 2])
                        with chk_col:
                            is_checked = st.checkbox(f"📎 {file_name}", key=f"chk_{actual_idx}_{file_path}")
                            if is_checked: checked_files_to_download.append(file_path)
                        if ext in ['pdf', 'png', 'jpg', 'jpeg', 'txt']:
                            with btn_col1:
                                if st.button("👁️ 열기", key=f"view_{actual_idx}_{file_path}", help="새창에서 문서 보기", use_container_width=True):
                                    show_file_dialog(file_path, file_name)
                            with btn_col2:
                                if st.button("✨ 요약", key=f"ai_{actual_idx}_{file_path}", help="새창에서 공문서 양식으로 요약", use_container_width=True):
                                    show_summary_dialog(file_path, file_name)
                        else:
                            with btn_col1: st.write("")
                            with btn_col2: st.write("")
                        with btn_col3:
                            if st.button("🗑️ 삭제", key=f"delf_{actual_idx}_{file_path}", use_container_width=True):
                                steps[actual_idx]['files'].remove(file_path)
                                try:
                                    os.remove(file_path) 
                                except:
                                    pass
                                save_data(st.session_state.site_data)
                                st.rerun()
                    if checked_files_to_download:
                        st.markdown("<br>", unsafe_allow_html=True)
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for fpath in checked_files_to_download:
                                zip_file.write(fpath, arcname=os.path.basename(fpath))
                        st.download_button(
                            label=f"💾 체크된 파일 {len(checked_files_to_download)}개 다운로드 (.zip)",
                            data=zip_buffer.getvalue(),
                            file_name=f"첨부파일_다운로드_{date.today().strftime('%Y%m%d')}.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True,
                            key=f"zip_dl_{actual_idx}"
                        )

if __name__ == "__main__":
    main()
