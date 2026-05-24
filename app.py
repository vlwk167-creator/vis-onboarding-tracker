# =============================================================================
# VIS 이슈어 온보딩 트래커  ·  DealMe Internal
# 개선 버전: 자동저장 / 액션 고도화 / URL검증 / 진행차트
# =============================================================================
import streamlit as st
import anthropic
import json
import re
import hashlib
import base64
import copy
import os
import urllib.parse
from io import BytesIO
from datetime import datetime, date

# 로컬 데이터 파일 경로 (app.py와 같은 폴더)
_DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")

def load_local_data() -> dict | None:
    """data.json이 있으면 읽어서 반환, 없으면 None."""
    try:
        if os.path.exists(_DATA_FILE):
            with open(_DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return None

def save_local_data():
    """현재 상태를 data.json에 저장."""
    try:
        payload = {
            "bank_state":  st.session_state.bank_state,
            "bank_order":  st.session_state.bank_order,
            "actions":     st.session_state.actions,
            "update_time": st.session_state.update_time,
        }
        with open(_DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

import openpyxl
import gspread
from google.oauth2.service_account import Credentials
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Google Sheets 클라이언트 (세션 초기화보다 먼저 정의) ───────────────────────
_SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

def get_gsheet_client():
    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=_SCOPES
        )
        return gspread.authorize(creds)
    except Exception:
        return None

# ── 페이지 설정 ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="VIS 이슈어 온보딩 트래커", page_icon="💳", layout="wide")

st.markdown("""
<style>
  .tag-urgent  { background:#FCEBEB; color:#A32D2D; padding:2px 10px; border-radius:20px; font-size:12px; font-weight:600; }
  .tag-check   { background:#E6F1FB; color:#185FA5; padding:2px 10px; border-radius:20px; font-size:12px; font-weight:600; }
  .tag-pending { background:#FAEEDA; color:#854F0B; padding:2px 10px; border-radius:20px; font-size:12px; font-weight:600; }
  .tag-done    { background:#EAF3DE; color:#2D6A0E; padding:2px 10px; border-radius:20px; font-size:12px; font-weight:600; }
  .bin-chip    { display:inline-block; background:#F0F4FF; color:#1a3a8f; border:1px solid #C7D4F5;
                 padding:3px 10px; border-radius:16px; font-size:12px; font-family:monospace; margin:2px; }
  .log-row     { font-size:12px; padding:4px 0; border-bottom:1px solid #f0f0f0; }
</style>
""", unsafe_allow_html=True)


# ── 1. 비밀번호 로그인 ─────────────────────────────────────────────────────────
def check_password():
    try:
        correct = st.secrets["passwords"]["correct_password"]
    except Exception:
        correct = os.environ.get("APP_PASSWORD", "dealme2026")

    def _hash(pw): return hashlib.sha256(pw.encode()).hexdigest()

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True

    st.markdown("""
    <div style="max-width:400px;margin:80px auto 0;text-align:center;">
      <div style="font-size:40px;margin-bottom:8px;">💳</div>
      <div style="font-size:22px;font-weight:700;margin-bottom:4px;">VIS 온보딩 트래커</div>
      <div style="font-size:13px;color:#888;margin-bottom:32px;">DealMe · 내부 전용</div>
    </div>
    """, unsafe_allow_html=True)

    _, col, _ = st.columns([1, 2, 1])
    with col:
        pw = st.text_input("비밀번호", type="password", placeholder="비밀번호를 입력하세요")
        if st.button("로그인", use_container_width=True, type="primary"):
            if _hash(pw) == _hash(correct):
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
    st.stop()

check_password()


# ── 2. 상수 / 기본 데이터 ──────────────────────────────────────────────────────
EMPTY_BANK = {
    "krw_vnd": "미완료", "usd": "미완료",
    "tc": "미완료",      "tc_url": "",
    "bin": "미완료",     "bin_list": [],
    "status": "미확인",  "note": "온보딩 시작 전",
}

DEFAULT_STATE = {
    "ACB": {
        "krw_vnd": "완료", "usd": "확인중", "tc": "완료",
        "tc_url": "https://acb.com.vn/acbwebsite/files/Ban_Dieu_khoan_dieu_kien_tra_gop_the_tin_dung_09.25.pdf",
        "bin": "완료", "bin_list": [],
        "status": "확인필요", "note": "USD 플랜 확인 중",
    },
    "Sacombank": {
        "krw_vnd": "완료", "usd": "완료", "tc": "완료",
        "tc_url": "https://www.sacombank.com.vn/content/dam/sacombank/files/the-le/sacombank-installment-en.pdf",
        "bin": "미완료", "bin_list": [],
        "status": "BIN 대기", "note": "BIN 리스트 미제출",
    },
    "TCB": {
        "krw_vnd": "완료", "usd": "진행중", "tc": "완료",
        "tc_url": "https://aemapp.techcombank.com/content/dam/techcombank/public-site/en/images/personal-banking/pl-06-hd-card-17-en-apendix-conditions-and-terms-of-use-of-credit-card-installments-a189549880-f3cae49033.pdf",
        "bin": "미완료", "bin_list": [],
        "status": "진행중", "note": "USD 세팅 중",
    },
    "Vietcombank": {
        "krw_vnd": "미완료", "usd": "미완료", "tc": "미완료", "tc_url": "",
        "bin": "미완료", "bin_list": [],
        "status": "심각지연", "note": "테스트 미통과",
    },
}

DEFAULT_ACTIONS = [
    {"tag": "확인필요", "bank": "ACB",         "text": "USD 플랜 최종 등록 완료 여부 Irene에게 확인 요청",
     "assignee": "", "due_date": "", "completed": False},
    {"tag": "긴급",    "bank": "Sacombank",    "text": "Irene/Wisid에게 BIN 리스트 제출 촉구 필요",
     "assignee": "", "due_date": "", "completed": False},
    {"tag": "대기중",  "bank": "TCB",          "text": "USD 플랜 세팅 완료 확인 후 BIN 리스트 수급 진행",
     "assignee": "", "due_date": "", "completed": False},
    {"tag": "긴급",    "bank": "Vietcombank",  "text": "로직 수정 이슈 단기 해결 불가. 신세계 런칭 시 제외 여부 의사결정 필요",
     "assignee": "", "due_date": "", "completed": False},
]

def load_from_gsheet(client, spreadsheet_id: str) -> dict | None:
    """구글 시트에서 bank_state, actions, update_time을 읽어와서 반환."""
    try:
        wb = client.open_by_key(spreadsheet_id)

        # dashboard 시트 → bank_state
        ws_dash = wb.worksheet("dashboard")
        records = ws_dash.get_all_records()
        bank_state = {}
        bank_order = []
        for row in records:
            bank = row.get("은행", "")
            if not bank:
                continue
            bank_order.append(bank)
            bank_state[bank] = {
                "krw_vnd": row.get("KRW/VND", ""),
                "usd":     row.get("USD", ""),
                "tc":      row.get("T&C", ""),
                "tc_url":  row.get("T&C URL", ""),
                "bin":     row.get("BIN상태", ""),
                "status":  row.get("전체상태", ""),
                "note":    row.get("비고", ""),
                "bin_list": [],
            }

        # bin_list 시트 → bin_list 복원
        try:
            ws_bin = wb.worksheet("bin_list")
            for row in ws_bin.get_all_records():
                bank = row.get("은행", "")
                bin_no = str(row.get("BIN 번호", "")).strip()
                if bank in bank_state and bin_no:
                    bank_state[bank]["bin_list"].append(bin_no)
        except Exception:
            pass

        # actions 시트 → actions
        actions = []
        try:
            ws_act = wb.worksheet("actions")
            for row in ws_act.get_all_records():
                tag  = row.get("태그", "")
                text = row.get("내용", "")
                if not text:
                    continue
                actions.append({
                    "tag":       tag,
                    "bank":      row.get("은행", ""),
                    "text":      text,
                    "assignee":  row.get("담당자", ""),
                    "due_date":  row.get("마감일", ""),
                    "completed": row.get("완료여부", "") == "완료",
                })
        except Exception:
            pass

        if not bank_order:
            return None

        # input_hashes 시트 복원
        input_hashes = []
        try:
            ws_hash = wb.worksheet("input_hashes")
            for row in ws_hash.get_all_records():
                h = row.get("hash", "")
                if h:
                    input_hashes.append(h)
        except Exception:
            pass

        return {
            "bank_state":   bank_state,
            "bank_order":   bank_order,
            "actions":      actions if actions else list(DEFAULT_ACTIONS),
            "update_time":  "구글 시트에서 복원",
            "input_hashes": input_hashes,
        }
    except Exception:
        return None

# API 키: secrets → 환경변수 → 빈 문자열 순으로 자동 로드
def _load_api_key() -> str:
    try:
        return st.secrets["anthropic"]["api_key"]
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY", "")

_defaults = {
    "bank_state":   dict(DEFAULT_STATE),
    "bank_order":   list(DEFAULT_STATE.keys()),
    "actions":      list(DEFAULT_ACTIONS),
    "update_time":  "2026년 5월 15일 (이메일 기준)",
    "api_key":      _load_api_key(),
    "input_hashes": [],
    "history":      [],
}
# ── 시작 시 데이터 로드: 구글 시트 → 로컬 JSON → 기본값 순으로 시도 ──────────
if "bank_state" not in st.session_state:
    _gsheet_data = None
    try:
        _gs_sid = st.secrets["google"]["spreadsheet_id"]
        _gs_cli = get_gsheet_client()
        if _gs_cli and _gs_sid:
            _gsheet_data = load_from_gsheet(_gs_cli, _gs_sid)
    except Exception:
        pass

    if _gsheet_data:
        st.session_state.bank_state   = _gsheet_data["bank_state"]
        st.session_state.bank_order   = _gsheet_data["bank_order"]
        st.session_state.actions      = _gsheet_data["actions"]
        st.session_state.update_time  = _gsheet_data["update_time"]
        st.session_state.input_hashes = _gsheet_data.get("input_hashes", [])
    else:
        _local = load_local_data()
        if _local:
            st.session_state.bank_state  = _local.get("bank_state",  dict(DEFAULT_STATE))
            st.session_state.bank_order  = _local.get("bank_order",  list(DEFAULT_STATE.keys()))
            st.session_state.actions     = _local.get("actions",     list(DEFAULT_ACTIONS))
            st.session_state.update_time = _local.get("update_time", "2026년 5월 15일 (이메일 기준)")

for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# 기존 세션 호환 처리
for _bank in st.session_state.bank_state:
    if "bin_list" not in st.session_state.bank_state[_bank]:
        st.session_state.bank_state[_bank]["bin_list"] = []

# 구버전 세션키 제거 (호환)
for _old_key in ["manual_log", "slack_webhook", "editor_name", "auto_save"]:
    st.session_state.pop(_old_key, None)

# 기존 actions에 bank 필드 없을 경우 보완 (구버전 호환)
for _act in st.session_state.actions:
    _act.setdefault("bank", "")

# 기존 actions에 새 필드 없을 경우 보완
for _act in st.session_state.actions:
    _act.setdefault("assignee", "")
    _act.setdefault("due_date", "")
    _act.setdefault("completed", False)
    _act.setdefault("bank", "")


# ── 3. 헬퍼 함수 ───────────────────────────────────────────────────────────────
def status_icon(s):
    return {"완료": "✅", "미완료": "❌", "진행중": "🔄", "확인중": "❓",
            "확인필요": "⚠️", "심각지연": "🚨", "BIN 대기": "⏳"}.get(s, "❓")

def card_bg(status):
    if "심각" in status or "지연" in status: return "#FFF0F0"
    if status == "완료":                      return "#F0FBF0"
    if status in ("미확인", ""):              return "#F8F9FA"
    if "확인필요" in status:                  return "#FFFBF0"
    return "#FFF8F0"

def card_border(status):
    if "심각" in status or "지연" in status: return "#FFAAAA"
    if status == "완료":                      return "#AADDAA"
    if "확인필요" in status:                  return "#FFD580"
    return "#E5E7EB"

def make_hash(text: str) -> str:
    return hashlib.sha256(text.strip().encode()).hexdigest()[:16]

def is_duplicate(h: str) -> bool:
    return h in st.session_state.input_hashes

def save_snapshot(label: str):
    snap = {
        "label":       label,
        "time":        datetime.now().strftime("%m/%d %H:%M"),
        "bank_state":  copy.deepcopy(st.session_state.bank_state),
        "bank_order":  list(st.session_state.bank_order),
        "actions":     copy.deepcopy(st.session_state.actions),
        "update_time": st.session_state.update_time,
    }
    st.session_state.history.append(snap)
    if len(st.session_state.history) > 10:
        st.session_state.history.pop(0)

def restore_snapshot(idx: int):
    snap = st.session_state.history[idx]
    st.session_state.bank_state  = copy.deepcopy(snap["bank_state"])
    st.session_state.bank_order  = list(snap["bank_order"])
    st.session_state.actions     = copy.deepcopy(snap["actions"])
    st.session_state.update_time = snap["update_time"]

def merge_bin_list(existing: list, new_bins: list) -> list:
    """기존 BIN 목록에 새 BIN을 중복 없이 추가."""
    merged = list(existing)
    for b in new_bins:
        b_str = str(b).strip()
        if b_str and b_str not in merged:
            merged.append(b_str)
    return sorted(merged)

# ── [NEW] URL 유효성 검사 ──────────────────────────────────────────────────────
def validate_url(url: str) -> tuple[bool, str]:
    """URL 형식 및 접근 가능 여부를 검사. (bool, 메시지) 반환."""
    if not url:
        return False, "URL이 비어 있습니다."
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in ("http", "https"):
        return False, "http 또는 https로 시작하는 URL이어야 합니다."
    if not parsed.netloc:
        return False, "올바른 도메인이 없습니다."
    return True, "✅ URL 형식 정상"

# ── 저장 & 동기화 ─────────────────────────────────────────────────────────────
def save_and_sync():
    """로컬 JSON 저장 + Google Sheets 연결 시 자동 동기화."""
    save_local_data()
    sid = st.session_state.get("_sheet_id", "")
    if sid:
        gs = get_gsheet_client()
        if gs:
            save_to_gsheet(gs, sid)


# ── 4. AI 파싱 엔진 ────────────────────────────────────────────────────────────
_BUSINESS_RULES = """
[핵심 비즈니스 규칙 — 반드시 준수]
1. 최신 정보 우선: 이메일 스레드는 아래로 갈수록 오래된 인용입니다.
   반드시 '가장 최근(상단) 메시지'의 결론을 최종값으로 채택하고, 하단 인용문은 참고만 하세요.
2. 롤백 인지: A→B→A처럼 원복되는 흐름이 있으면 최종값은 A입니다. 중간값 B를 최종값으로 착각하지 마세요.
3. 특이사항 자동 감지:
   - 8자리 BIN 번호(예: 43116707)가 언급되면 → status='확인필요', 액션 아이템에 추가
   - T&C 업로드 오류 / 시스템 에러 문맥 → status='확인필요', 액션 아이템에 추가
   - 테스트 실패, 로직 오류 → status='심각지연'
4. 은행 약자 매핑: VCB=Vietcombank, SCB=Sacombank, TCB=TCB(Techcombank)
5. 필드 허용값:
   - krw_vnd / usd / bin: 완료 | 미완료 | 진행중 | 확인중
   - tc: 완료 | 미완료
   - status: 완료 | 진행중 | 확인필요 | BIN대기 | 심각지연 | 미확인
   - bin_list: 6~8자리 숫자 BIN 번호 문자열 배열. 언급된 모든 BIN을 추출하세요.
6. 온보딩과 무관한 내용이면 반드시 {"irrelevant": true, "reason": "..."} 형태로만 응답하세요.
7. 언급이 없는 은행 필드는 무조건 빈 문자열("") 또는 빈 배열([])로 남기세요.
"""

def _bank_json_template() -> str:
    lines = [
        f'  "{b}": {{"krw_vnd":"","usd":"","tc":"","tc_url":"","bin":"","bin_list":[],"status":"","note":""}}'
        for b in st.session_state.bank_order
    ]
    return "{\n" + ",\n".join(lines) + "\n}"

def _call_claude_text(prompt: str) -> dict:
    client = anthropic.Anthropic(api_key=st.session_state.api_key)
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = re.sub(r"```json|```", "", msg.content[0].text).strip()
    return json.loads(raw)

def _call_claude_vision(image_bytes: bytes, media_type: str, prompt: str) -> dict:
    client = anthropic.Anthropic(api_key=st.session_state.api_key)
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": media_type,
                    "data": base64.b64encode(image_bytes).decode(),
                }},
                {"type": "text", "text": prompt},
            ],
        }],
    )
    raw = re.sub(r"```json|```", "", msg.content[0].text).strip()
    return json.loads(raw)

def analyze_email(text: str) -> dict:
    prompt = f"""{_BUSINESS_RULES}

[작업]
아래 이메일 스레드를 읽고 VIS 이슈어 온보딩 현황을 분석하세요.
등록 은행: {', '.join(st.session_state.bank_order)}

- tc_url: 이메일에 T&C 원문 PDF/웹 URL이 있으면 해당 은행의 tc_url에 넣으세요.
- bin_list: 본문에 등장하는 모든 6~8자리 BIN 번호를 배열로 추출하세요. 없으면 [].
- actions: 후속 조치 항목 최대 5개. tag는 긴급|확인필요|대기중 중 하나. bank는 해당 은행명(등록 은행 중 하나), 여러 은행이면 가장 관련 높은 은행 하나만.
- update_date: 확인되는 가장 최근 날짜 "YYYY년 MM월 DD일" 형식.

[이메일 본문]
{text}

[출력 — JSON만 반환, 다른 텍스트 절대 금지]
관련 없음: {{"irrelevant": true, "reason": "이유"}}
관련 있음:
{{
  "irrelevant": false,
  "banks": {_bank_json_template()},
  "new_banks": [],
  "actions": [{{"tag": "긴급|확인필요|대기중", "bank": "은행명", "text": ""}}],
  "update_date": ""
}}"""
    return _call_claude_text(prompt)

def analyze_excel(sheet_data: dict) -> dict:
    content = ""
    for sname, rows in sheet_data.items():
        content += f"\n[시트: {sname}]\n"
        content += "".join("\t".join(r) + "\n" for r in rows[:80])

    prompt = f"""{_BUSINESS_RULES}

[작업]
아래는 VIS 온보딩 관련 엑셀 파일 내용입니다.
등록 은행: {', '.join(st.session_state.bank_order)}

- tc_url: 엑셀에 T&C URL이 있으면 해당 은행의 tc_url에 넣으세요.
- bin_list: 파일에 등장하는 모든 6~8자리 BIN 번호를 배열로 추출하세요. 없으면 [].
- bin_count: 파악된 BIN 번호 총 개수.
- submitting_bank: 이 파일을 제출한 은행 이름.

[엑셀 내용]
{content}

[출력 — JSON만 반환]
관련 없음: {{"irrelevant": true, "reason": "이유"}}
관련 있음:
{{
  "irrelevant": false,
  "banks": {_bank_json_template()},
  "new_banks": [],
  "actions": [{{"tag": "긴급|확인필요|대기중", "bank": "은행명", "text": ""}}],
  "update_date": "",
  "submitting_bank": "",
  "bin_count": 0
}}"""
    return _call_claude_text(prompt)

def analyze_image(image_bytes: bytes, media_type: str) -> dict:
    prompt = f"""{_BUSINESS_RULES}

[작업]
이 이미지는 VIS 이슈어 온보딩 관련 문서나 이메일 캡처일 수 있습니다.
등록 은행: {', '.join(st.session_state.bank_order)}

- tc_url: 이미지에 T&C URL이 보이면 반드시 추출하세요.
- bin_list: 이미지에서 보이는 모든 6~8자리 BIN 번호를 배열로 추출하세요. 없으면 [].

[출력 — JSON만 반환]
관련 없는 이미지: {{"irrelevant": true, "reason": "관련없는 이미지입니다"}}
관련 있는 이미지:
{{
  "irrelevant": false,
  "banks": {_bank_json_template()},
  "new_banks": [],
  "actions": [{{"tag": "긴급|확인필요|대기중", "bank": "은행명", "text": ""}}],
  "update_date": ""
}}"""
    return _call_claude_vision(image_bytes, media_type, prompt)


# ── 5. 상태 업데이트 로직 ──────────────────────────────────────────────────────
def read_excel_bytes(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    out = {}
    for name in wb.sheetnames:
        rows = []
        for row in wb[name].iter_rows(values_only=True):
            if any(c is not None for c in row):
                rows.append([str(c) if c is not None else "" for c in row])
        out[name] = rows
    return out

def apply_update(parsed: dict, input_hash: str = None):
    """파싱 결과를 session_state에 반영. irrelevant면 완전 무시."""
    if parsed.get("irrelevant"):
        return

    save_snapshot("업데이트 전")
    if input_hash:
        st.session_state.input_hashes.append(input_hash)

    # 새 은행 추가
    for nb in parsed.get("new_banks", []):
        if nb and nb not in st.session_state.bank_state:
            st.session_state.bank_state[nb] = dict(EMPTY_BANK)
            st.session_state.bank_state[nb]["bin_list"] = []
            st.session_state.bank_order.append(nb)

    # 은행 상태 업데이트
    for bank, info in parsed.get("banks", {}).items():
        if not info:
            continue
        if bank not in st.session_state.bank_state:
            st.session_state.bank_state[bank] = dict(EMPTY_BANK)
            st.session_state.bank_state[bank]["bin_list"] = []
            st.session_state.bank_order.append(bank)

        for k, v in info.items():
            if k == "bin_list":
                if isinstance(v, list) and v:
                    existing = st.session_state.bank_state[bank].get("bin_list", [])
                    st.session_state.bank_state[bank]["bin_list"] = merge_bin_list(existing, v)
            elif v:
                st.session_state.bank_state[bank][k] = v

    # 액션 아이템: 완료된 항목 보존 + 신규 항목 추가
    new_actions_raw = [a for a in parsed.get("actions", []) if a.get("text", "").strip()]
    if new_actions_raw:
        # 기존에서 완료된 것들 보존
        completed_old = [a for a in st.session_state.actions if a.get("completed")]
        new_actions = []
        for a in new_actions_raw:
            new_actions.append({
                "tag":       a.get("tag", "확인필요"),
                "bank":      a.get("bank", ""),
                "text":      a.get("text", ""),
                "assignee":  "",
                "due_date":  "",
                "completed": False,
            })
        st.session_state.actions = new_actions + completed_old

    if parsed.get("update_date"):
        st.session_state.update_time = parsed["update_date"]

    # 자동 저장
    save_and_sync()


# ── 6. Google Sheets 연동 ──────────────────────────────────────────────────────
_FIELD_LABELS = {
    "krw_vnd": "KRW/VND 플랜", "usd": "USD 플랜",
    "tc": "영어 T&C", "bin": "BIN 리스트",
    "status": "전체 상태", "note": "비고",
}

def save_to_gsheet(client, spreadsheet_id: str):
    try:
        wb    = client.open_by_key(spreadsheet_id)
        banks = st.session_state.bank_order
        state = st.session_state.bank_state
        now   = datetime.now().strftime("%Y-%m-%d %H:%M")

        # dashboard 시트
        try:
            ws_dash = wb.worksheet("dashboard")
            ws_dash.clear()
        except gspread.WorksheetNotFound:
            ws_dash = wb.add_worksheet(title="dashboard", rows=60, cols=12)

        header = ["은행", "KRW/VND", "USD", "T&C", "T&C URL", "BIN상태", "BIN수", "전체상태", "비고", "마지막업데이트"]
        rows = [header]
        for bank in banks:
            info = state.get(bank, dict(EMPTY_BANK))
            bin_cnt = len(info.get("bin_list", []))
            rows.append([
                bank,
                info.get("krw_vnd", ""), info.get("usd", ""),
                info.get("tc", ""),      info.get("tc_url", ""),
                info.get("bin", ""),     bin_cnt,
                info.get("status", ""),  info.get("note", ""), now,
            ])
        ws_dash.update("A1", rows)
        ws_dash.format("A1:J1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
        })

        # bin_list 시트
        try:
            ws_bin = wb.worksheet("bin_list")
            ws_bin.clear()
        except gspread.WorksheetNotFound:
            ws_bin = wb.add_worksheet(title="bin_list", rows=2000, cols=3)
        bin_header = ["은행", "BIN 번호", "저장일시"]
        bin_rows = [bin_header]
        for bank in banks:
            info = state.get(bank, dict(EMPTY_BANK))
            for b in info.get("bin_list", []):
                bin_rows.append([bank, b, now])
        ws_bin.update("A1", bin_rows)
        ws_bin.format("A1:C1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
        })

        # history 시트
        try:
            ws_hist = wb.worksheet("history")
        except gspread.WorksheetNotFound:
            ws_hist = wb.add_worksheet(title="history", rows=2000, cols=6)
            ws_hist.append_row(["날짜", "은행", "항목", "이전값", "현재값", "비고"])
            ws_hist.format("A1:F1", {
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
            })

        last_state: dict = {}
        for rec in ws_hist.get_all_records():
            b, f, v = rec.get("은행", ""), rec.get("항목", ""), rec.get("현재값", "")
            if b and f:
                last_state.setdefault(b, {})[f] = v

        new_rows = []
        for bank in banks:
            info = state.get(bank, dict(EMPTY_BANK))
            for field in ["krw_vnd", "usd", "tc", "bin", "status"]:
                label    = _FIELD_LABELS[field]
                cur_val  = info.get(field, "")
                prev_val = last_state.get(bank, {}).get(label)
                if prev_val is None or prev_val != cur_val:
                    new_rows.append([now, bank, label,
                                     prev_val if prev_val else "신규",
                                     cur_val, info.get("note", "")])
        if new_rows:
            ws_hist.append_rows(new_rows)

        # [NEW] actions 시트
        try:
            ws_act_gs = wb.worksheet("actions")
            ws_act_gs.clear()
        except gspread.WorksheetNotFound:
            ws_act_gs = wb.add_worksheet(title="actions", rows=100, cols=6)
        act_header = ["태그", "은행", "내용", "담당자", "마감일", "완료여부", "저장일시"]
        act_rows = [act_header]
        for a in st.session_state.actions:
            act_rows.append([
                a.get("tag",""), a.get("bank",""), a.get("text",""),
                a.get("assignee",""), a.get("due_date",""),
                "완료" if a.get("completed") else "미완료", now,
            ])
        ws_act_gs.update("A1", act_rows)
        ws_act_gs.format("A1:G1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
        })

        # input_hashes 시트 (중복 제출 방지)
        try:
            ws_hash = wb.worksheet("input_hashes")
            ws_hash.clear()
        except gspread.WorksheetNotFound:
            ws_hash = wb.add_worksheet(title="input_hashes", rows=500, cols=2)
        hashes = st.session_state.get("input_hashes", [])
        hash_rows = [["hash", "저장일시"]] + [[h, now] for h in hashes]
        ws_hash.update("A1", hash_rows)

        return True, len(new_rows)
    except Exception as e:
        return False, str(e)

def load_history_from_gsheet(client, spreadsheet_id: str) -> list:
    try:
        return client.open_by_key(spreadsheet_id).worksheet("history").get_all_records()
    except Exception:
        return []


# ── 7. 엑셀 내보내기 ───────────────────────────────────────────────────────────
def export_excel_file() -> BytesIO:
    wb = openpyxl.Workbook()

    hdr_fill  = PatternFill("solid", fgColor="1F4E79")
    hdr_font  = Font(bold=True, color="FFFFFF", size=11)
    center    = Alignment(horizontal="center", vertical="center")
    wrap_left = Alignment(vertical="center", wrap_text=True)
    thin      = Side(style="thin", color="CCCCCC")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    sc = {"완료": "EAF3DE", "미완료": "FCEBEB", "진행중": "FFF3CD",
          "확인중": "E6F1FB", "확인필요": "FFF3CD", "해당없음": "F1EFE8"}

    def write_header(ws, headers, col_widths):
        for i, (h, w) in enumerate(zip(headers, col_widths), 1):
            c = ws.cell(row=1, column=i, value=h)
            c.font, c.fill, c.alignment, c.border = hdr_font, hdr_fill, center, border
            ws.column_dimensions[c.column_letter].width = w

    # 시트1: 온보딩 현황
    ws1 = wb.active
    ws1.title = "온보딩 현황"
    write_header(ws1,
                 ["은행명", "KRW/VND 플랜", "USD 플랜", "영어 T&C", "BIN 리스트", "BIN 수", "전체 상태", "비고"],
                 [15, 16, 14, 14, 14, 10, 14, 35])
    for r, bank in enumerate(st.session_state.bank_order, 2):
        info    = st.session_state.bank_state.get(bank, dict(EMPTY_BANK))
        bin_cnt = len(info.get("bin_list", []))
        vals = [bank, info["krw_vnd"], info["usd"], info["tc"],
                info["bin"], bin_cnt, info["status"], info["note"]]
        for c_idx, val in enumerate(vals, 1):
            cell = ws1.cell(row=r, column=c_idx, value=val)
            cell.border = border
            cell.alignment = center if c_idx != 8 else wrap_left
            if c_idx in (2, 3, 4, 5, 7) and str(val) in sc:
                cell.fill = PatternFill("solid", fgColor=sc[str(val)])
        ws1.row_dimensions[r].height = 22

    # 시트2: BIN 리스트
    ws_bin = wb.create_sheet("BIN 리스트")
    write_header(ws_bin, ["은행명", "BIN 번호", "자리수"], [15, 14, 10])
    bin_row = 2
    for bank in st.session_state.bank_order:
        info     = st.session_state.bank_state.get(bank, dict(EMPTY_BANK))
        bin_list = info.get("bin_list", [])
        if bin_list:
            for b in bin_list:
                ws_bin.cell(row=bin_row, column=1, value=bank).border = border
                bc = ws_bin.cell(row=bin_row, column=2, value=b)
                bc.border, bc.alignment = border, center
                bc.font = Font(name="Courier New", size=11)
                dc = ws_bin.cell(row=bin_row, column=3, value=len(str(b)))
                dc.border, dc.alignment = border, center
                ws_bin.row_dimensions[bin_row].height = 20
                bin_row += 1
        else:
            ws_bin.cell(row=bin_row, column=1, value=bank).border = border
            nc = ws_bin.cell(row=bin_row, column=2, value="(미제출)")
            nc.border, nc.alignment, nc.font = border, center, Font(color="999999", italic=True)
            ws_bin.cell(row=bin_row, column=3, value="").border = border
            ws_bin.row_dimensions[bin_row].height = 20
            bin_row += 1

    ws_bin.cell(row=1, column=5, value="은행별 BIN 수 요약").font = Font(bold=True)
    for sr, bank in enumerate(st.session_state.bank_order, 2):
        info = st.session_state.bank_state.get(bank, dict(EMPTY_BANK))
        ws_bin.cell(row=sr, column=5, value=bank)
        ws_bin.cell(row=sr, column=6, value=len(info.get("bin_list", []))).alignment = center
    ws_bin.column_dimensions["E"].width = 15
    ws_bin.column_dimensions["F"].width = 10

    # 시트3: T&C 현황
    ws_tc = wb.create_sheet("T&C 현황")
    write_header(ws_tc, ["은행명", "T&C 상태", "원문 URL"], [15, 12, 90])
    for r, bank in enumerate(st.session_state.bank_order, 2):
        info  = st.session_state.bank_state.get(bank, dict(EMPTY_BANK))
        tc_v  = info.get("tc", "미완료")
        tc_u  = info.get("tc_url", "")
        ws_tc.cell(row=r, column=1, value=bank).border = border
        tc_c = ws_tc.cell(row=r, column=2, value=tc_v)
        tc_c.alignment, tc_c.border = center, border
        if tc_v in sc:
            tc_c.fill = PatternFill("solid", fgColor=sc[tc_v])
        url_c = ws_tc.cell(row=r, column=3, value=tc_u)
        url_c.alignment, url_c.border = wrap_left, border
        if tc_u:
            url_c.font = Font(color="185FA5", underline="single")
        ws_tc.row_dimensions[r].height = 22

    # 시트4: 액션 아이템 (담당자·마감일·완료 포함)
    ws_act = wb.create_sheet("액션 아이템")
    write_header(ws_act, ["태그", "내용", "담당자", "마감일", "완료"], [14, 60, 14, 14, 10])
    tag_c = {"긴급": "FCEBEB", "확인필요": "E6F1FB", "대기중": "FAEEDA"}
    for r, a in enumerate(st.session_state.actions, 2):
        t = ws_act.cell(row=r, column=1, value=a["tag"])
        t.alignment, t.border = center, border
        t.fill = PatternFill("solid", fgColor=tag_c.get(a["tag"], "FFFFFF"))
        n = ws_act.cell(row=r, column=2, value=a["text"])
        n.alignment, n.border = wrap_left, border
        asn = ws_act.cell(row=r, column=3, value=a.get("assignee", ""))
        asn.alignment, asn.border = center, border
        dd = ws_act.cell(row=r, column=4, value=a.get("due_date", ""))
        dd.alignment, dd.border = center, border
        cmp = ws_act.cell(row=r, column=5, value="✅" if a.get("completed") else "")
        cmp.alignment, cmp.border = center, border
        ws_act.row_dimensions[r].height = 22

    # 시트5: 요약
    ws_sum = wb.create_sheet("요약")
    ws_sum.column_dimensions["A"].width = 22
    ws_sum.column_dimensions["B"].width = 18
    bstate  = st.session_state.bank_state
    total   = len(bstate)
    ready   = sum(1 for b in bstate.values() if b["status"] == "완료")
    crit    = sum(1 for b in bstate.values() if "심각" in b["status"])
    tc_done = sum(1 for b in bstate.values() if b.get("tc") == "완료")
    total_bins = sum(len(b.get("bin_list", [])) for b in bstate.values())
    act_done = sum(1 for a in st.session_state.actions if a.get("completed"))
    for r, (k, v) in enumerate([
        ("내보내기 일시",    datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("마지막 업데이트", st.session_state.update_time),
        ("전체 은행 수",     total),
        ("런칭 준비 완료",  ready),
        ("진행 중",          total - ready - crit),
        ("심각한 지연",     crit),
        ("T&C 완료",         tc_done),
        ("등록된 BIN 총계",  total_bins),
    ], 1):
        ws_sum.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws_sum.cell(row=r, column=2, value=v)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf




# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

# ── 사이드바 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 설정")
    st.divider()
    # secrets에 키가 있으면 입력창 숨김
    _key_from_secrets = bool(_load_api_key())
    if _key_from_secrets:
        st.success("✅ API 키 설정됨")
    else:
        key_in = st.text_input("Anthropic API Key", type="password",
                               value=st.session_state.api_key, placeholder="sk-ant-...")
        if key_in:
            st.session_state.api_key = key_in
        if st.session_state.api_key:
            st.success("✅ API 키 설정됨")
        else:
            st.warning("AI 분석을 위해 API 키를 입력하세요")

    st.divider()
    st.header("📊 Google Sheets")
    try:
        _sheet_id = st.secrets["google"]["spreadsheet_id"]
        st.caption(f"시트 ID: `{_sheet_id[:18]}...`")
    except Exception:
        _sheet_id = st.text_input("Spreadsheet ID", placeholder="구글 시트 URL의 ID")
    st.session_state["_sheet_id"] = _sheet_id if _sheet_id else ""

    _gs_client = get_gsheet_client()
    if _gs_client and _sheet_id:
        st.success("✅ 구글 시트 연결됨")
        st.caption("🟢 데이터 변경 시 자동으로 동기화됩니다")
        if st.button("💾 지금 동기화"):
            with st.spinner("동기화 중..."):
                ok, result = save_to_gsheet(_gs_client, _sheet_id)
                if ok:
                    st.success(f"완료! 변경 {result}건")
                else:
                    st.error(f"실패: {result}")
    else:
        st.info("Google Sheets 미연결 — 로컬 파일에 자동 저장됩니다")

    st.divider()
    st.header("🏦 은행 관리")
    new_bank = st.text_input("은행 이름 추가", placeholder="예: BIDV, VPBank")
    if st.button("➕ 추가"):
        n = new_bank.strip()
        if not n:
            st.warning("이름을 입력해주세요.")
        elif n in st.session_state.bank_state:
            st.warning(f"'{n}'은 이미 등록되어 있습니다.")
        else:
            st.session_state.bank_state[n] = dict(EMPTY_BANK)
            st.session_state.bank_state[n]["bin_list"] = []
            st.session_state.bank_order.append(n)
            st.success(f"'{n}' 추가 완료!")
            st.rerun()

    del_target = st.selectbox("은행 삭제", ["선택..."] + st.session_state.bank_order)
    if st.button("🗑️ 삭제"):
        if del_target != "선택...":
            st.session_state.bank_state.pop(del_target, None)
            if del_target in st.session_state.bank_order:
                st.session_state.bank_order.remove(del_target)
            st.success(f"'{del_target}' 삭제 완료")
            st.rerun()

    st.divider()
    st.markdown("**범례** ✅완료 ❌미완료 🔄진행중 ❓확인중 ⚠️확인필요 🚨심각지연")
    st.divider()
    c_lo, c_re = st.columns(2)
    with c_lo:
        if st.button("🚪 로그아웃", use_container_width=True):
            st.session_state.authenticated = False
            st.rerun()
    with c_re:
        if st.button("🔄 초기화", use_container_width=True):
            for k in ["bank_state","bank_order","actions","update_time","input_hashes","history"]:
                st.session_state.pop(k, None)
            st.rerun()


# ── 메인 헤더 ──────────────────────────────────────────────────────────────────
st.title("💳 VIS 이슈어 온보딩 트래커")
st.caption(f"신세계면세점 런칭 전 베트남 은행 온보딩 현황 | 마지막 업데이트: **{st.session_state.update_time}**")

# KPI
_st         = st.session_state.bank_state
_total      = len(_st)
_ready      = sum(1 for b in _st.values() if b["status"] == "완료")
_crit       = sum(1 for b in _st.values() if "심각" in b["status"])
_tc_done    = sum(1 for b in _st.values() if b.get("tc") == "완료")
_tc_url     = sum(1 for b in _st.values() if b.get("tc_url"))
_total_bins = sum(len(b.get("bin_list", [])) for b in _st.values())

k1, k2, k3, k4 = st.columns(4)
k1.metric("🏦 전체 은행",     _total)
k2.metric("✅ 런칭 준비완료", _ready)
k3.metric("🚨 심각 지연",     _crit)
k4.metric("🔢 등록 BIN",      _total_bins)

st.divider()

# 내보내기
with st.expander("📤 보고서 내보내기", expanded=False):
    buf_x = export_excel_file()
    st.download_button("📊 엑셀 다운로드 (.xlsx)", data=buf_x,
                       file_name=f"VIS_온보딩_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)


# ── 대시보드 탭 ──────────────────────────────────────────────────────────────
dash_tab1, dash_tab2, dash_tab3, dash_tab4 = st.tabs([
    "📊 온보딩 현황", "📄 T&C 현황", "🔢 BIN 리스트", "📈 진행 현황"
])

# ── 탭1: 온보딩 현황 ───────────────────────────────────────────────────────────
with dash_tab1:
    st.subheader("은행별 온보딩 현황")
    _fields = [("krw_vnd","KRW/VND 플랜"),("usd","USD 플랜"),("tc","영어 T&C"),("bin","BIN 상태")]
    cols_card = st.columns(2)

    for i, bank in enumerate(st.session_state.bank_order):
        if bank not in _st:
            continue
        info   = _st[bank]
        tc_url = info.get("tc_url", "")
        tc_lnk = (f' <a href="{tc_url}" target="_blank" style="font-size:11px;color:#185FA5;">📄 원문</a>') if tc_url else ""

        rows_html = ""
        for fk, lbl in _fields:
            extra = tc_lnk if fk == "tc" else ""
            _bl = info.get("bin_list", [])
            bin_cnt_str = f" <span style='font-size:11px;color:#888;'>({len(_bl)}개)</span>" if fk == "bin" else ""
            rows_html += (
                f'<div style="display:flex;justify-content:space-between;'
                f'padding:5px 0;font-size:13px;border-bottom:1px solid rgba(0,0,0,0.06);">'
                f'<span style="color:#555;">{lbl}</span>'
                f'<span style="font-weight:500;">{status_icon(info[fk])} {info[fk]}{bin_cnt_str}{extra}</span></div>'
            )

        with cols_card[i % 2]:
            st.markdown(f"""
            <div style="background:{card_bg(info['status'])};border-radius:14px;padding:18px 22px;
                        margin-bottom:14px;border:1.5px solid {card_border(info['status'])};">
              <div style="font-size:17px;font-weight:700;margin-bottom:10px;">
                {bank}
                <span style="font-size:12px;font-weight:400;color:#777;margin-left:8px;">
                  {status_icon(info['status'])} {info['status']}
                </span>
              </div>
              {rows_html}
              <div style="font-size:12px;color:#888;margin-top:8px;">📝 {info.get('note','')}</div>
            </div>""", unsafe_allow_html=True)

            with st.expander(f"✏️ {bank} 직접 편집"):
                _opts = {
                    "krw_vnd": ["완료","미완료","진행중","확인중"],
                    "usd":     ["완료","미완료","진행중","확인중"],
                    "tc":      ["완료","미완료"],
                    "bin":     ["완료","미완료","진행중","확인중"],
                    "status":  ["완료","진행중","확인필요","BIN 대기","심각지연","미확인"],
                }
                edited = {}
                for fk, opts in _opts.items():
                    cur = info.get(fk, opts[0])
                    idx = opts.index(cur) if cur in opts else 0
                    edited[fk] = st.selectbox(fk.upper().replace("_","/"), opts,
                                              index=idx, key=f"edit_{bank}_{fk}")
                edited["note"]   = st.text_input("비고", value=info.get("note",""), key=f"edit_{bank}_note")

                # [ENHANCED] T&C URL 수정 + 유효성 검사
                new_tc_url = st.text_input("T&C URL", value=info.get("tc_url",""),
                                           key=f"edit_{bank}_tc_url", placeholder="https://...")
                if new_tc_url:
                    url_ok, url_msg = validate_url(new_tc_url)
                    if url_ok:
                        st.caption(url_msg)
                    else:
                        st.warning(f"⚠️ URL 오류: {url_msg}")
                edited["tc_url"] = new_tc_url

                if st.button("💾 저장", key=f"save_{bank}"):
                    save_snapshot(f"{bank} 수동편집 전")
                    for fk, val in edited.items():
                        st.session_state.bank_state[bank][fk] = val
                    st.success(f"✅ {bank} 저장 완료")
                    save_and_sync()
                    st.rerun()

    st.divider()

    # ── [ENHANCED] 액션 아이템 (담당자·마감일·완료 처리) ─────────────────────
    st.subheader("📌 현재 액션 아이템")
    _tag_cls = {"긴급":"tag-urgent","확인필요":"tag-check","대기중":"tag-pending","완료":"tag-done"}

    if st.session_state.actions:
        pending_actions = [a for a in st.session_state.actions if not a.get("completed")]
        done_actions    = [a for a in st.session_state.actions if a.get("completed")]

        for idx, a in enumerate(st.session_state.actions):
            cls = _tag_cls.get(a.get("tag",""), "tag-check")
            completed = a.get("completed", False)
            text_style = "text-decoration:line-through;color:#aaa;" if completed else ""

            col_check, col_tag, col_bank, col_text, col_assign, col_date = st.columns([0.5, 1.2, 1.5, 4, 1.5, 1.5])

            with col_check:
                new_done = st.checkbox("", value=completed, key=f"act_done_{idx}",
                                       label_visibility="collapsed")
                if new_done != completed:
                    st.session_state.actions[idx]["completed"] = new_done
                    save_and_sync()
                    st.rerun()

            with col_tag:
                st.markdown(f'<span class="{cls}">{a["tag"]}</span>', unsafe_allow_html=True)

            with col_bank:
                bank_name = a.get("bank", "")
                if bank_name:
                    st.markdown(
                        f'<span style="display:inline-block;background:#F0F4FF;color:#1a3a8f;'
                        f'border:1px solid #C7D4F5;padding:2px 10px;border-radius:16px;'
                        f'font-size:12px;font-weight:600;">{bank_name}</span>',
                        unsafe_allow_html=True,
                    )

            with col_text:
                st.markdown(f'<span style="{text_style}">{a["text"]}</span>', unsafe_allow_html=True)

            with col_assign:
                new_assignee = st.text_input("담당자", value=a.get("assignee",""),
                                              key=f"act_assign_{idx}", placeholder="담당자",
                                              label_visibility="collapsed")
                if new_assignee != a.get("assignee",""):
                    st.session_state.actions[idx]["assignee"] = new_assignee

            with col_date:
                new_due = st.text_input("마감일", value=a.get("due_date",""),
                                         key=f"act_due_{idx}", placeholder="MM/DD",
                                         label_visibility="collapsed")
                if new_due != a.get("due_date",""):
                    st.session_state.actions[idx]["due_date"] = new_due

        # 완료 항목 정리 버튼
        if done_actions:
            st.caption(f"✅ 완료된 항목 {len(done_actions)}개 포함")
            if st.button("🗑️ 완료 항목 모두 제거"):
                st.session_state.actions = pending_actions
                save_and_sync()
                st.rerun()

        # 액션 아이템 수동 추가
        with st.expander("➕ 액션 아이템 직접 추가"):
            c1, c2 = st.columns(2)
            with c1:
                new_tag  = st.selectbox("태그", ["긴급","확인필요","대기중"], key="new_act_tag")
            with c2:
                new_bank = st.selectbox("은행", [""] + st.session_state.bank_order, key="new_act_bank")
            new_text = st.text_input("내용", key="new_act_text", placeholder="조치 내용을 입력하세요")
            new_asn  = st.text_input("담당자", key="new_act_asn", placeholder="담당자 이름")
            new_due_str = st.text_input("마감일", key="new_act_due", placeholder="예: 06/01")
            if st.button("추가", key="add_act_btn", type="primary"):
                if new_text.strip():
                    st.session_state.actions.append({
                        "tag": new_tag, "bank": new_bank, "text": new_text.strip(),
                        "assignee": new_asn.strip(), "due_date": new_due_str.strip(),
                        "completed": False,
                    })
                    save_and_sync()
                    st.rerun()
    else:
        st.info("등록된 액션 아이템이 없습니다.")


# ── 탭2: T&C 현황 ─────────────────────────────────────────────────────────────
with dash_tab2:
    st.subheader("카드사별 영문 T&C 현황")
    st.caption("이미지·이메일·엑셀 분석 시 T&C URL이 감지되면 자동 반영됩니다.")

    m1, m2, m3 = st.columns(3)
    m1.metric("전체 은행",     len(st.session_state.bank_order))
    m2.metric("T&C 제출 완료", _tc_done)
    m3.metric("URL 확보",      _tc_url)

    st.markdown("---")
    _h1,_h2,_h3,_h4 = st.columns([2,1.2,1,5])
    _h1.markdown("**은행명**"); _h2.markdown("**T&C 상태**")
    _h3.markdown("**URL**");    _h4.markdown("**원문 링크**")
    st.markdown("---")

    for bank in st.session_state.bank_order:
        info      = st.session_state.bank_state.get(bank, dict(EMPTY_BANK))
        tc_status = info.get("tc","미완료")
        tc_url    = info.get("tc_url","")
        c1,c2,c3,c4 = st.columns([2,1.2,1,5])
        c1.markdown(f"**{bank}**")
        c2.markdown(f"{status_icon(tc_status)} {tc_status}")
        c3.markdown("✅" if tc_url else "❌")
        if tc_url:
            c4.markdown(f"[📄 T&C 원문 보기]({tc_url})")
        else:
            c4.caption("URL 미확보")

        with st.expander(f"✏️ {bank} T&C URL 수정"):
            new_url = st.text_input("T&C URL", value=tc_url,
                                    key=f"tc_url_{bank}", placeholder="https://...")
            # [NEW] URL 유효성 검사
            if new_url:
                url_ok, url_msg = validate_url(new_url)
                if url_ok:
                    st.caption(url_msg)
                else:
                    st.warning(f"⚠️ {url_msg}")

            if st.button("저장", key=f"tc_save_{bank}"):
                if new_url.strip():
                    url_ok, url_msg = validate_url(new_url.strip())
                    if not url_ok:
                        st.error(f"유효하지 않은 URL입니다: {url_msg}")
                    else:
                        save_snapshot(f"{bank} T&C 수정 전")
                        st.session_state.bank_state[bank]["tc_url"] = new_url.strip()
                        st.session_state.bank_state[bank]["tc"] = "완료"
                        st.success(f"✅ {bank} T&C URL 저장됨")
                        save_and_sync()
                        st.rerun()
                else:
                    save_snapshot(f"{bank} T&C 수정 전")
                    st.session_state.bank_state[bank]["tc_url"] = ""
                    st.success(f"✅ {bank} T&C URL 초기화됨")
                    save_and_sync()
                    st.rerun()


# ── 탭3: BIN 리스트 ───────────────────────────────────────────────────────────
with dash_tab3:
    st.subheader("🔢 은행별 BIN 리스트")
    st.caption("이메일·엑셀·이미지 분석 시 6~8자리 BIN 번호가 감지되면 자동으로 누적 추가됩니다. 수동 입력도 가능합니다.")

    b1, b2, b3 = st.columns(3)
    _banks_with_bin = sum(1 for b in _st.values() if b.get("bin_list"))
    b1.metric("BIN 제출 은행",  _banks_with_bin)
    b2.metric("BIN 미제출 은행", len(_st) - _banks_with_bin)
    b3.metric("전체 BIN 수",    _total_bins)

    st.markdown("---")

    for bank in st.session_state.bank_order:
        info     = st.session_state.bank_state.get(bank, dict(EMPTY_BANK))
        bin_list = info.get("bin_list", [])
        bin_cnt  = len(bin_list)

        col_title, col_badge = st.columns([5, 1])
        with col_title:
            st.markdown(f"### 🏦 {bank}")
        with col_badge:
            badge_color = "#EAF3DE" if bin_cnt > 0 else "#FCEBEB"
            badge_text  = "#2D6A0E" if bin_cnt > 0 else "#A32D2D"
            st.markdown(
                f'<div style="background:{badge_color};color:{badge_text};'
                f'border-radius:20px;padding:4px 14px;text-align:center;'
                f'font-weight:700;font-size:14px;margin-top:8px;">'
                f'{bin_cnt}개</div>',
                unsafe_allow_html=True,
            )

        if bin_list:
            chips_html = "".join(f'<span class="bin-chip">{b}</span>' for b in bin_list)
            st.markdown(f'<div style="padding:10px 0 4px;">{chips_html}</div>', unsafe_allow_html=True)

            with st.expander(f"📋 {bank} BIN 전체 목록 보기 / 편집 ({bin_cnt}개)"):
                col_headers = st.columns([1, 2, 2])
                col_headers[0].markdown("**#**")
                col_headers[1].markdown("**BIN 번호**")
                col_headers[2].markdown("**자리수**")
                for idx, b in enumerate(bin_list, 1):
                    r1, r2, r3 = st.columns([1, 2, 2])
                    r1.write(idx)
                    r2.code(b)
                    r3.write(f"{len(str(b))}자리")

                st.markdown("---")
                del_bin = st.selectbox("삭제할 BIN 선택", ["선택..."] + bin_list, key=f"del_bin_{bank}")
                if st.button("🗑️ 선택 BIN 삭제", key=f"del_bin_btn_{bank}"):
                    if del_bin != "선택...":
                        save_snapshot(f"{bank} BIN 삭제 전")
                        st.session_state.bank_state[bank]["bin_list"].remove(del_bin)
                        st.success(f"✅ {del_bin} 삭제됨")
                        save_and_sync()
                        st.rerun()

                if st.button(f"⚠️ {bank} BIN 전체 삭제", key=f"clear_bin_{bank}", type="secondary"):
                    save_snapshot(f"{bank} BIN 전체삭제 전")
                    st.session_state.bank_state[bank]["bin_list"] = []
                    st.session_state.bank_state[bank]["bin"] = "미완료"
                    st.success(f"✅ {bank} BIN 목록 초기화됨")
                    save_and_sync()
                    st.rerun()
        else:
            st.caption("아직 등록된 BIN이 없습니다.")

        with st.expander(f"➕ {bank} BIN 수동 입력"):
            st.caption("여러 개는 쉼표(,) 또는 줄바꿈으로 구분하세요. 예: 431167, 43116707, 456789")
            manual_bin_input = st.text_area(
                "BIN 번호 입력", height=100,
                key=f"manual_bin_{bank}",
                placeholder="431167\n43116707\n456789",
            )
            if st.button("➕ 추가", key=f"add_bin_{bank}", type="primary"):
                raw = manual_bin_input.replace(",", "\n")
                new_bins = [
                    b.strip() for b in raw.split("\n")
                    if b.strip().isdigit() and 6 <= len(b.strip()) <= 8
                ]
                # [NEW] 유효하지 않은 항목 상세 안내
                invalid = [
                    b.strip() for b in raw.split("\n")
                    if b.strip() and not (b.strip().isdigit() and 6 <= len(b.strip()) <= 8)
                ]
                if new_bins:
                    save_snapshot(f"{bank} BIN 수동추가 전")
                    existing = st.session_state.bank_state[bank].get("bin_list", [])
                    merged   = merge_bin_list(existing, new_bins)
                    added    = len(merged) - len(existing)
                    st.session_state.bank_state[bank]["bin_list"] = merged
                    if merged:
                        st.session_state.bank_state[bank]["bin"] = "완료"
                    st.success(f"✅ {added}개 추가됨 (전체 {len(merged)}개)")
                    save_and_sync()
                if invalid:
                    inv_detail = []
                    for b in invalid:
                        if not b.isdigit():
                            inv_detail.append(f"{b} (숫자만 허용)")
                        elif len(b) < 6:
                            inv_detail.append(f"{b} ({len(b)}자리 — 최소 6자리)")
                        elif len(b) > 8:
                            inv_detail.append(f"{b} ({len(b)}자리 — 최대 8자리)")
                    st.warning(f"⚠️ 유효하지 않은 항목 무시됨:\n" + "\n".join(inv_detail))
                if new_bins:
                    st.rerun()

        st.markdown("---")


# ── 탭4: [NEW] 진행 현황 차트 ─────────────────────────────────────────────────
with dash_tab4:
    st.subheader("📈 진행 현황 차트")

    try:
        import plotly.graph_objects as go
        import plotly.express as px
        import pandas as pd
        _plotly_ok = True
    except ImportError:
        _plotly_ok = False
        st.warning("plotly가 설치되어 있지 않습니다. `pip install plotly` 후 재시작하세요.")

    if _plotly_ok:
        _fields_check = ["krw_vnd", "usd", "tc", "bin"]

        # ── 은행별 완료율 바 차트 ────────────────────────────────────────────
        st.markdown("#### 🏦 은행별 완료율")
        chart_data = []
        for bank in st.session_state.bank_order:
            info = st.session_state.bank_state.get(bank, dict(EMPTY_BANK))
            done = sum(1 for fk in _fields_check if info.get(fk) == "완료")
            pct  = round(done / len(_fields_check) * 100)
            chart_data.append({"은행": bank, "완료율(%)": pct, "완료": done,
                                "상태": info.get("status","미확인")})

        df_chart = pd.DataFrame(chart_data)
        color_map = {
            "완료": "#639922", "진행중": "#EF9F27", "확인필요": "#EF9F27",
            "BIN 대기": "#EF9F27", "심각지연": "#E24B4A", "미확인": "#888780",
        }
        bar_colors = [color_map.get(s, "#888780") for s in df_chart["상태"]]

        fig_bar = go.Figure(go.Bar(
            x=df_chart["은행"],
            y=df_chart["완료율(%)"],
            marker_color=bar_colors,
            text=[f"{p}%" for p in df_chart["완료율(%)"]],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>완료율: %{y}%<extra></extra>",
        ))
        fig_bar.update_layout(
            yaxis=dict(range=[0, 115], title="완료율 (%)"),
            xaxis_title="은행",
            plot_bgcolor="white",
            margin=dict(t=30, b=20),
            height=320,
        )
        fig_bar.add_hline(y=100, line_dash="dot", line_color="#639922",
                          annotation_text="목표 100%", annotation_position="right")
        st.plotly_chart(fig_bar, use_container_width=True)

        # ── 전체 상태 분포 파이 차트 ─────────────────────────────────────────
        col_pie1, col_pie2 = st.columns(2)

        with col_pie1:
            st.markdown("#### 📊 전체 상태 분포")
            status_counts = {}
            for b in st.session_state.bank_state.values():
                s = b.get("status", "미확인")
                status_counts[s] = status_counts.get(s, 0) + 1
            pie_colors = [color_map.get(s, "#888780") for s in status_counts.keys()]
            fig_pie = go.Figure(go.Pie(
                labels=list(status_counts.keys()),
                values=list(status_counts.values()),
                marker_colors=pie_colors,
                hole=0.4,
                textinfo="label+percent",
            ))
            fig_pie.update_layout(margin=dict(t=10, b=10), height=280,
                                   showlegend=False)
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_pie2:
            st.markdown("#### 📌 액션 아이템 현황")
            act_pending = sum(1 for a in st.session_state.actions if not a.get("completed"))
            act_done_n  = sum(1 for a in st.session_state.actions if a.get("completed"))
            act_urgent  = sum(1 for a in st.session_state.actions
                              if not a.get("completed") and a.get("tag") == "긴급")

            fig_act = go.Figure(go.Pie(
                labels=["완료", "진행중(일반)", "긴급"],
                values=[act_done_n, max(0, act_pending - act_urgent), act_urgent],
                marker_colors=["#639922", "#EF9F27", "#E24B4A"],
                hole=0.4,
                textinfo="label+value",
            ))
            fig_act.update_layout(margin=dict(t=10, b=10), height=280,
                                   showlegend=False)
            st.plotly_chart(fig_act, use_container_width=True)

        # ── 항목별 완료 현황 히트맵 ──────────────────────────────────────────
        st.markdown("#### 🗂️ 항목별 완료 현황")
        field_labels = {"krw_vnd": "KRW/VND", "usd": "USD", "tc": "T&C", "bin": "BIN"}
        heat_data = []
        for bank in st.session_state.bank_order:
            info = st.session_state.bank_state.get(bank, dict(EMPTY_BANK))
            row = {"은행": bank}
            for fk, lbl in field_labels.items():
                v = info.get(fk, "미완료")
                row[lbl] = 1 if v == "완료" else (0.5 if v in ("진행중","확인중") else 0)
            heat_data.append(row)
        df_heat = pd.DataFrame(heat_data).set_index("은행")
        fig_heat = px.imshow(
            df_heat,
            color_continuous_scale=["#FCEBEB", "#FFF3CD", "#EAF3DE"],
            zmin=0, zmax=1,
            text_auto=False,
            aspect="auto",
        )
        fig_heat.update_traces(
            text=[[
                {1.0:"완료", 0.5:"진행중", 0.0:"미완료"}.get(v, "미완료")
                for v in row
            ] for row in df_heat.values],
            texttemplate="%{text}",
        )
        fig_heat.update_layout(
            coloraxis_showscale=False,
            margin=dict(t=10, b=10),
            height=max(200, len(st.session_state.bank_order) * 50),
        )
        st.plotly_chart(fig_heat, use_container_width=True)

        # ── Google Sheets 기반 시계열 (연동 시) ──────────────────────────────
        _gs2 = get_gsheet_client()
        _sid = st.session_state.get("_sheet_id", "")
        if _gs2 and _sid:
            st.markdown("#### 📅 날짜별 변경 이력 (Google Sheets 기반)")
            with st.spinner("이력 불러오는 중..."):
                hist_records = load_history_from_gsheet(_gs2, _sid)
            if hist_records:
                df_hist = pd.DataFrame(hist_records)
                if not df_hist.empty and "날짜" in df_hist.columns:
                    df_hist["날짜_dt"] = pd.to_datetime(df_hist["날짜"], errors="coerce")
                    df_hist = df_hist.dropna(subset=["날짜_dt"])
                    if not df_hist.empty:
                        df_hist["날짜만"] = df_hist["날짜_dt"].dt.date
                        daily_counts = df_hist.groupby("날짜만").size().reset_index(name="변경 건수")
                        fig_line = px.line(daily_counts, x="날짜만", y="변경 건수",
                                           markers=True, title="일별 변경 건수")
                        fig_line.update_layout(margin=dict(t=40, b=10), height=280)
                        st.plotly_chart(fig_line, use_container_width=True)
            else:
                st.info("저장된 이력이 없습니다. 사이드바 '💾 지금 저장'을 눌러주세요.")
        else:
            st.info("Google Sheets를 연동하면 날짜별 변경 추이 차트를 볼 수 있습니다.")


# ── 입력 섹션 ─────────────────────────────────────────────────────────────────
st.subheader("📥 새 데이터 입력 → AI 분석 → 자동 업데이트")
st.caption("세 가지 방법 모두 동일한 템플릿에 반영됩니다. BIN 번호도 자동 추출·누적됩니다.")

inp1, inp2, inp3 = st.tabs(["📧 이메일 붙여넣기", "📊 엑셀 업로드", "🖼️ 이미지 업로드"])

with inp1:
    st.caption("아웃룩 이메일 스레드를 통째로 붙여넣으세요. AI가 최신 메시지 기준으로 상태·BIN·T&C URL 등을 자동 반영합니다.")
    email_text = st.text_area("이메일 본문", height=220, placeholder="이메일 전문을 그대로 붙여넣으세요...")
    if st.button("🤖 이메일 분석 후 업데이트", type="primary", key="btn_email"):
        if not st.session_state.api_key:
            st.error("먼저 사이드바에서 API 키를 입력하세요.")
        elif not email_text.strip():
            st.warning("이메일 내용을 입력해주세요.")
        else:
            h = make_hash(email_text)
            if is_duplicate(h):
                st.warning("⚠️ 이미 처리한 내용입니다.")
            else:
                with st.spinner("AI가 이메일을 분석하는 중..."):
                    try:
                        parsed = analyze_email(email_text)
                        if parsed.get("irrelevant"):
                            st.warning(f"⚠️ {parsed.get('reason','온보딩과 관련 없는 내용입니다.')}")
                            st.session_state.input_hashes.append(h)
                        else:
                            new_b = parsed.get("new_banks", [])
                            apply_update(parsed, input_hash=h)
                            msg = "✅ 현황판 업데이트 완료!"
                            if new_b: msg += f" 새 은행 추가: {', '.join(new_b)}"
                            st.success(msg); st.rerun()
                    except Exception as e:
                        st.error(f"오류: {e}")

with inp2:
    st.caption("은행이 보낸 온보딩 시트(.xlsx)를 업로드하세요. BIN 번호 목록도 자동 추출합니다.")
    uploaded_xlsx = st.file_uploader("엑셀 파일 (.xlsx)", type=["xlsx"])
    if uploaded_xlsx:
        st.info(f"📂 업로드됨: **{uploaded_xlsx.name}**")
        if st.button("🤖 엑셀 분석 후 업데이트", type="primary", key="btn_excel"):
            if not st.session_state.api_key:
                st.error("먼저 사이드바에서 API 키를 입력하세요.")
            else:
                file_bytes = uploaded_xlsx.read()
                h = make_hash(file_bytes.decode("latin-1"))
                if is_duplicate(h):
                    st.warning("⚠️ 이미 처리한 파일입니다.")
                else:
                    with st.spinner("AI가 엑셀을 분석하는 중..."):
                        try:
                            parsed = analyze_excel(read_excel_bytes(file_bytes))
                            if parsed.get("irrelevant"):
                                st.warning(f"⚠️ {parsed.get('reason','온보딩과 관련 없는 파일입니다.')}")
                                st.session_state.input_hashes.append(h)
                            else:
                                new_b = parsed.get("new_banks", [])
                                apply_update(parsed, input_hash=h)
                                msg = f"✅ 업데이트 완료! (제출 은행: {parsed.get('submitting_bank','?')}, BIN: {parsed.get('bin_count',0)}개)"
                                if new_b: msg += f" | 새 은행: {', '.join(new_b)}"
                                st.success(msg); st.rerun()
                        except Exception as e:
                            st.error(f"오류: {e}")

with inp3:
    st.caption("이메일 스크린샷·표 캡처 등을 업로드하세요. AI가 BIN 번호와 T&C URL까지 이미지에서 직접 추출합니다.")
    uploaded_img = st.file_uploader("이미지 파일 (PNG, JPG, JPEG, WEBP)", type=["png","jpg","jpeg","webp"])
    if uploaded_img:
        st.image(uploaded_img, caption=f"📷 {uploaded_img.name}", use_container_width=True)
        if st.button("🤖 이미지 분석 후 업데이트", type="primary", key="btn_img"):
            if not st.session_state.api_key:
                st.error("먼저 사이드바에서 API 키를 입력하세요.")
            else:
                img_bytes  = uploaded_img.read()
                ext        = uploaded_img.name.rsplit(".", 1)[-1].lower()
                mt_map     = {"jpg":"image/jpeg","jpeg":"image/jpeg","png":"image/png","webp":"image/webp"}
                media_type = mt_map.get(ext, "image/png")
                h = make_hash(img_bytes.decode("latin-1"))
                if is_duplicate(h):
                    st.warning("⚠️ 이미 처리한 이미지입니다.")
                else:
                    with st.spinner("AI가 이미지를 분석하는 중..."):
                        try:
                            parsed = analyze_image(img_bytes, media_type)
                            if parsed.get("irrelevant"):
                                st.warning(f"⚠️ {parsed.get('reason','온보딩과 관련 없는 이미지입니다.')}")
                                st.session_state.input_hashes.append(h)
                            else:
                                new_b = parsed.get("new_banks", [])
                                apply_update(parsed, input_hash=h)
                                msg = "✅ 현황판 업데이트 완료!"
                                if new_b: msg += f" 새 은행 추가: {', '.join(new_b)}"
                                st.success(msg); st.rerun()
                        except Exception as e:
                            st.error(f"오류: {e}")

st.divider()

# ── 히스토리 & 롤백 ────────────────────────────────────────────────────────────
if st.session_state.history:
    with st.expander(f"🕓 업데이트 히스토리 ({len(st.session_state.history)}개)", expanded=False):
        st.caption("최대 10개 보관. 복원 버튼을 누르면 해당 시점으로 돌아갑니다.")
        for i, snap in reversed(list(enumerate(st.session_state.history))):
            ca, cb = st.columns([6, 1])
            ca.markdown(f'`{snap["time"]}` &nbsp; {snap["label"]} → **{snap["update_time"]}**')
            with cb:
                if st.button("↩ 복원", key=f"restore_{i}"):
                    restore_snapshot(i)
                    st.success("이전 상태로 복원했습니다.")
                    st.rerun()
    st.divider()

st.caption("VIS Issuer Onboarding Tracker · Powered by Claude AI · DealMe Internal · v2.0")
