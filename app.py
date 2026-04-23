import streamlit as st
import pandas as pd
import math
import numbers
import re
from copy import copy
from io import BytesIO
from datetime import datetime

try:
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell
except Exception:
    load_workbook = None

    class MergedCell:  # fallback to keep module importable
        pass

# -------------------------------
# 기본 설정
# -------------------------------
st.set_page_config(page_title="서무봇(업무택시)", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 2rem;}
.stButton>button {
    border-radius: 14px;
    background-color: #f5f5f7;
}
.stTextInput input {
    border-radius: 12px;
}
</style>
""", unsafe_allow_html=True)

# -------------------------------
# 시작 알림
# -------------------------------
if "ack" not in st.session_state:
    st.session_state.ack = False

if not st.session_state.ack:
    st.warning("⚠️ 본 프로그램은 실습용으로 만들었으며 제출 전 다시 확인하셔야 합니다")
    if st.button("확인"):
        st.session_state.ack = True
        st.rerun()
    st.stop()

# -------------------------------
# 제목
# -------------------------------
title_col_logo, title_col_text = st.columns([0.85, 4.15], gap="small")
with title_col_logo:
    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)
    st.image(r"c:\Users\USER\AI-Education\5.DatabaseSQL\code\logo.png", width=180)
with title_col_text:
    st.markdown(
        "<h1 style='margin: 10px 0 0 -18px; padding: 0;'>서무봇(업무택시)</h1>",
        unsafe_allow_html=True,
    )
st.caption("재정경제부 서무 자동화 시스템")

# -------------------------------
# 과 이름
# -------------------------------
_, dept_mid = st.columns([1, 2])
with dept_mid:
    dept_name = st.text_input("과 이름 입력")

# -------------------------------
# 인원 입력
# -------------------------------
st.subheader("👥 인원 입력")

positions = ["과장", "주무사무관(서기관)", "사무관", "주무관", "그 외"]

if "members" not in st.session_state:
    st.session_state.members = {p: [""] for p in positions}


def members_snapshot():
    """text_input은 key로 session_state에만 저장되므로, 위젯 값을 모아 필터·정렬에 씁니다."""
    out = {}
    for pos in positions:
        row = []
        for i in range(len(st.session_state.members[pos])):
            key = f"{pos}_{i}"
            val = st.session_state.get(key, "")
            row.append("" if val is None else str(val))
        out[pos] = row
    return out


def render_member_block(pos):
    """이름 입력란을 짧게(좁은 열) + 추가 버튼."""
    st.markdown(f"**{pos}**")
    for i, name in enumerate(st.session_state.members[pos]):
        col_in, col_btn = st.columns([1, 0.22])
        with col_in:
            st.text_input("", key=f"{pos}_{i}", value=name, label_visibility="collapsed")
        with col_btn:
            if st.button("➕", key=f"add_{pos}_{i}"):
                st.session_state.members[pos].append("")


row_top = st.columns(3)
for idx, pos in enumerate(positions[:3]):
    with row_top[idx]:
        render_member_block(pos)

row_bot = st.columns(3)
with row_bot[0]:
    render_member_block(positions[3])
with row_bot[1]:
    render_member_block(positions[4])
with row_bot[2]:
    pass

# -------------------------------
# 파일 업로드
# -------------------------------
st.subheader("📂 파일 업로드")

total_file = st.file_uploader("부처 전체 이용내역", type=["xlsx"])
form_file = st.file_uploader(
    "과 양식 파일 (예: `2.업무택시이용내역(양식).xlsx` — 10행 헤더·7행 집계 수식 유지)",
    type=["xlsx"],
)

# -------------------------------
# 이름 정규화
# -------------------------------
def normalize_name(name):
    return str(name).strip().replace(" ", "")


def _col_label(c):
    return str(c).strip()


def find_name_col(df):
    for key in ("이름", "성명", "이용자", "승객"):
        for c in df.columns:
            if key in _col_label(c):
                return c
    return None


def find_date_col(df):
    for key in ("날짜", "일자", "일시", "승차", "배차", "이용일", "호출일", "시각"):
        for c in df.columns:
            if key in _col_label(c):
                return c
    for c in df.columns:
        low = _col_label(c).lower()
        if "date" in low or "time" in low:
            return c
    skip_keys = ("금액", "요금", "거리", "과금", "원", "번호")
    best_col, best_ok = None, -1
    for c in df.columns:
        lab = _col_label(c)
        if any(s in lab for s in skip_keys):
            continue
        parsed = pd.to_datetime(df[c], errors="coerce")
        ok = int(parsed.notna().sum())
        if ok > best_ok:
            best_ok, best_col = ok, c
    return best_col if best_ok > 0 else None


def find_phone_col(df):
    for key in ("핸드폰", "휴대폰", "휴대전화", "전화번호", "연락처", "phone", "mobile", "Mobile"):
        for c in df.columns:
            if key in _col_label(c):
                return c
    return None


def format_usage_datetime_minutes(val):
    """이용일시: 분 단위까지만 (초 제거)."""
    ts = pd.to_datetime(val, errors="coerce")
    if pd.isna(ts):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ""
        return str(val).strip()
    ts = ts.replace(second=0, microsecond=0)
    return ts.strftime("%Y-%m-%d %H:%M")


def strip_address_from_gu(val):
    """출발/도착: 앞쪽 ○도·○시 등 제거 후 첫 '○구'부터 표기."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip().replace("\n", " ")
    if not s:
        return ""
    m = re.search(r"[가-힣0-9]{1,40}구", s)
    if m:
        return s[m.start() :].strip()
    return s


def format_amount_no_minus(val):
    """사용 금액: 음수/마이너스 표기 제거 후 숫자만."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, numbers.Real) and not isinstance(val, bool):
        v = float(val)
        if math.isnan(v):
            return ""
        return int(abs(v)) if v.is_integer() else abs(v)
    s = str(val).strip().replace(",", "").replace(" ", "")
    for ch in ("−", "–", "—"):
        s = s.replace(ch, "-")
    s = re.sub(r"^[\-+]+", "", s)
    s = s.replace("-", "")
    if not s:
        return ""
    try:
        n = float(s)
        return int(n) if n.is_integer() else n
    except ValueError:
        return s


def parse_amount_numeric(val):
    """금액을 숫자로 (총액 합산용)."""
    x = format_amount_no_minus(val)
    if x == "" or x is None:
        return 0.0
    if isinstance(x, numbers.Real) and not isinstance(x, bool):
        return float(x)
    try:
        return float(str(x).replace(",", ""))
    except ValueError:
        return 0.0

# -------------------------------
# 데이터 추출
# -------------------------------
def extract_data(file, members):
    file.seek(0)
    xls = pd.ExcelFile(file)
    result = pd.DataFrame()

    all_names = [
        normalize_name(n)
        for p in members
        for n in members[p]
        if normalize_name(n)
    ]

    if not all_names:
        return result

    for sheet in xls.sheet_names:
        df = xls.parse(sheet)

        name_col = find_name_col(df)

        if name_col:
            df[name_col] = df[name_col].astype(str).apply(normalize_name)
            filtered = df[df[name_col].isin(all_names)]
            result = pd.concat([result, filtered])

    return result

# -------------------------------
# 정렬
# -------------------------------
def sort_data(df, members):

    order_map = {}
    priority = 0

    for pos in positions:
        for name in members[pos]:
            order_map[normalize_name(name)] = priority
        priority += 1

    name_col = find_name_col(df)
    if not name_col:
        raise ValueError(
            "이름/성명/이용자 열을 찾지 못했습니다. 엑셀 첫 행(헤더)에 해당 열이 있는지 확인해 주세요."
        )

    date_col = find_date_col(df)

    df = df.copy()
    df["정렬"] = df[name_col].map(order_map)
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        df = df.sort_values(by=[date_col, "정렬"])
    else:
        df = df.sort_values(by=["정렬"])

    return df

# -------------------------------
# 이용시간 → 출장/야근, 주간/심야 (6시 미만·22시 초과 = 야간대)
# -------------------------------
def is_overnight_usage_time(val):
    """6:00 미만 또는 22:00 초과(초 단위)이면 True."""
    try:
        dt = pd.to_datetime(val, errors="coerce")
        if pd.isna(dt):
            return False
        sec = dt.hour * 3600 + dt.minute * 60 + dt.second
        return sec < 6 * 3600 or sec > 22 * 3600
    except Exception:
        return False


def 이용사유_출장_야근(row, date_col):
    """G열 이용사유: '출장' 또는 '야근'."""
    try:
        dt = pd.to_datetime(row[date_col], errors="coerce")
    except Exception:
        dt = pd.NaT
    if pd.isna(dt):
        return ""
    return "야근" if is_overnight_usage_time(dt) else "출장"


def 주간야근_표시(row, date_col):
    """I열 주간/심야: '주간' 또는 '심야'만."""
    try:
        dt = pd.to_datetime(row[date_col], errors="coerce")
    except Exception:
        dt = pd.NaT
    if pd.isna(dt):
        return ""
    return "심야" if is_overnight_usage_time(dt) else "주간"


def normalize_shift_for_excel(v):
    """I열: '주간' 또는 '심야'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    if not s:
        return None
    if s in ("주간", "심야"):
        return s
    if s in ("야근", "야간"):
        return "심야"
    return None


def shift_from_usage_time_value(v):
    """C열 사용시간 기준으로 I열 값을 강제 계산: 6시 미만/22시 초과=심야, 아니면 주간."""
    dt = pd.to_datetime(v, errors="coerce")
    if pd.isna(dt):
        return None
    return "심야" if is_overnight_usage_time(dt) else "주간"


def sanitize_i_column_no_yageun(ws, start_row, row_count):
    """I열(주간/심야)에서 '야근'을 절대 허용하지 않음."""
    end_row = start_row + max(row_count - 1, 0)
    for r in range(start_row, end_row + 1):
        c_val = ws.cell(r, 3).value
        i_cell = ws.cell(r, 9)
        i_val = i_cell.value

        # 값이 비어 있으면 C열 기준으로 채움
        if i_val is None or (isinstance(i_val, str) and not i_val.strip()):
            i_cell.value = shift_from_usage_time_value(c_val)
            continue

        # 어떤 경로로든 '야근'이 들어오면 반드시 '심야'로 치환
        if isinstance(i_val, str):
            s = i_val.strip()
            if "야근" in s or "야간" in s:
                i_cell.value = "심야"
            elif s not in ("주간", "심야"):
                # 예상 외 값은 C열 기준으로 정정
                i_cell.value = shift_from_usage_time_value(c_val)


def safe_cell(ws, row, col):
    c = ws.cell(row=row, column=col)
    if isinstance(c, MergedCell):
        return None
    return c


def set_cell_safe(ws, row, col, value):
    """병합 셀이면 건너뜀."""
    cell = safe_cell(ws, row, col)
    if cell is not None:
        cell.value = value


# 양식 `2.업무택시이용내역(양식).xlsx` 10행 헤더와 동일 (A~J, 11행부터 데이터)
OUTPUT_COLUMNS = [
    "이름",
    "휴대폰번호",
    "사용시간",
    "출발지(장소)",
    "도착지(장소)",
    "이용요금",
    "이용사유",
    "내용(상세)",
    "주간/심야",
    "적정사용여부",
]


def build_final_output_table(source_df):
    """전체본(추출) 데이터프레임 → 최종 파일과 같은 열 구성의 표."""
    df = source_df
    name_col = find_name_col(df)
    date_col = find_date_col(df)
    if not name_col or not date_col:
        raise ValueError(
            "필수 열을 찾지 못했습니다. (이름·성명·이용자, 날짜·일시 등) "
            f"현재 열: {list(df.columns)}"
        )
    phone_col = find_phone_col(df)
    amount_col = next((c for c in df.columns if "금액" in c or "요금" in c), None)
    start_col = next((c for c in df.columns if "출발" in c), None)
    end_col = next((c for c in df.columns if "도착" in c), None)

    def _cell_txt(row, col):
        if not col or col not in row.index:
            return ""
        v = row[col]
        return "" if pd.isna(v) else v

    rows = []
    total_amount = 0.0
    for _, row in df.iterrows():
        if amount_col:
            total_amount += parse_amount_numeric(row[amount_col])
        phone = ""
        if phone_col and phone_col in row.index and pd.notna(row[phone_col]):
            phone = str(row[phone_col]).strip()
        rows.append(
            {
                "이름": row[name_col],
                "휴대폰번호": phone,
                "사용시간": format_usage_datetime_minutes(row[date_col]),
                "출발지(장소)": strip_address_from_gu(_cell_txt(row, start_col)),
                "도착지(장소)": strip_address_from_gu(_cell_txt(row, end_col)),
                "이용요금": format_amount_no_minus(row[amount_col]) if amount_col else "",
                "이용사유": 이용사유_출장_야근(row, date_col),
                "내용(상세)": "",
                "주간/심야": 주간야근_표시(row, date_col),
                "적정사용여부": "O",
            }
        )

    amt_display = ""
    if amount_col:
        amt = total_amount
        amt_display = int(amt) if amt == int(amt) else amt

    rows.append(
        {
            "이름": "합계",
            "휴대폰번호": "",
            "사용시간": "",
            "출발지(장소)": "",
            "도착지(장소)": "",
            "이용요금": amt_display,
            "이용사유": "",
            "내용(상세)": "",
            "주간/심야": "",
            "적정사용여부": "",
        }
    )

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


def usage_year_month_from_output(out_df):
    """출력 데이터의 사용시간에서 대표 연/월을 계산."""
    body = out_df[out_df["이름"].astype(str).str.strip() != "합계"]
    y, m = datetime.now().year, datetime.now().month
    col = "사용시간" if "사용시간" in body.columns else None
    if col and not body.empty:
        for v in body[col]:
            ts = pd.to_datetime(v, errors="coerce")
            if pd.notna(ts):
                y, m = ts.year, ts.month
                break
    return y, m


def title_merged_a1_from_usage(out_df):
    """양식 A1:J1 제목 — ( YYYY. MM 월 ) 업무택시 이용내역"""
    y, m = usage_year_month_from_output(out_df)
    return f"( {y}. {m:02d} 월 ) 업무택시 이용내역"


def dept_label_a7(dept_name):
    """A7 병합: ~관으로 끝나면 접미 '과' 없이 그대로, 그 외는 OOO과."""
    s = str(dept_name).strip() if dept_name else ""
    if not s:
        return ""
    if s.endswith("관"):
        return s
    if s.endswith("과"):
        return s
    return f"{s}과"


def clear_template_data_rows(ws, start_row, col_end=10, num_rows=320):
    """
    10행 헤더 아래(11행~) 데이터 영역만 비움(A~J). 5~7행 집계 수식은 건드리지 않음.
    """
    end_row = start_row + num_rows
    for r in range(start_row, end_row + 1):
        for c in range(1, col_end + 1):
            cell = safe_cell(ws, r, c)
            if cell is not None:
                cell.value = None

def _excel_value_for_column(col_name, raw, row_name):
    """엑셀 셀에 넣을 값(사용시간은 datetime 권장 — 양식 수식 HOUR() 호환)."""
    if col_name == "적정사용여부":
        return None if str(row_name).strip() == "합계" else "O"
    if col_name == "주간/심야":
        # I열은 항상 C열 시간 규칙(주간/심야)으로 맞춘다.
        return shift_from_usage_time_value(raw) or normalize_shift_for_excel(raw)
    if col_name == "사용시간":
        if raw is None or (isinstance(raw, float) and pd.isna(raw)):
            return None
        ts = pd.to_datetime(raw, errors="coerce")
        if pd.isna(ts):
            return raw
        return ts.to_pydatetime().replace(second=0, microsecond=0)
    if col_name == "이용요금":
        if raw is None or raw == "":
            return None
        return format_amount_no_minus(raw)
    if pd.isna(raw):
        return None
    if isinstance(raw, str) and raw.strip() == "":
        return None
    return raw


# -------------------------------
# 양식 적용 (미리보기와 동일한 output_df 사용)
# -------------------------------
def apply_template(form_file, output_df, dept_name):
    if load_workbook is None:
        raise RuntimeError(
            "openpyxl가 설치되지 않았습니다. requirements.txt에 openpyxl를 추가 후 재배포해 주세요."
        )

    form_file.seek(0)
    wb = load_workbook(form_file)
    ws = wb.active

    start_row = 11

    # C10 사용시간 열: 양식 C11~C13과 동일한 표시 형식(숫자 서식)
    sample_c = ws.cell(start_row, 3)
    usage_time_numfmt = sample_c.number_format if sample_c.number_format else "m/d/yy h:mm"

    ref_cell = ws.cell(start_row, 10)
    ref_font = copy(ref_cell.font) if ref_cell.font else None
    ref_align = copy(ref_cell.alignment) if ref_cell.alignment else None

    out = output_df.reindex(columns=OUTPUT_COLUMNS)
    excel_rows = out[out["이름"].astype(str).str.strip() != "합계"]
    n = len(excel_rows)
    clear_template_data_rows(ws, start_row, col_end=10, num_rows=max(320, n + 5))

    # A1:J1 병합 제목 (양식 문구 형식)
    set_cell_safe(ws, 1, 1, title_merged_a1_from_usage(out))

    # A7:B7 병합 — 소속 과명 (양식 예: OOO과)
    set_cell_safe(ws, 7, 1, dept_label_a7(dept_name))

    def apply_row_style(target_row, cols):
        for c in cols:
            cell = ws.cell(target_row, c)
            if ref_font is not None:
                cell.font = copy(ref_font)
            if ref_align is not None:
                cell.alignment = copy(ref_align)

    for offset in range(len(excel_rows)):
        r = start_row + offset
        row = excel_rows.iloc[offset]
        nm = str(row["이름"]).strip() if "이름" in row.index else ""
        for j, col in enumerate(OUTPUT_COLUMNS, start=1):
            raw = row[col] if col in row.index else None
            ws.cell(r, j).value = _excel_value_for_column(col, raw, nm)
        # I열(9)은 어떤 값이 들어왔든 C열(3) 기준으로 마지막에 강제 고정
        c_val = ws.cell(r, 3).value
        ws.cell(r, 9).value = shift_from_usage_time_value(c_val)
        apply_row_style(r, tuple(range(1, 11)))
        ccell = ws.cell(r, 3)
        if ccell.value is not None and isinstance(ccell.value, datetime):
            ccell.number_format = usage_time_numfmt

    # 안전장치: I열에는 '주간'/'심야'만 남기고 '야근'은 전부 제거
    sanitize_i_column_no_yageun(ws, start_row, len(excel_rows))

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# -------------------------------
# 실행
# -------------------------------
if st.button("🚀 내역 생성"):

    if not total_file or not form_file:
        st.error("파일 업로드 필요")
    else:
        members = members_snapshot()
        if not any(normalize_name(n) for p in members for n in members[p]):
            st.warning("인원 입력란에 이름을 한 명 이상 입력해 주세요.")
        else:
            df = extract_data(total_file, members)

            if len(df) == 0:
                st.warning(
                    "데이터 없음: 엑셀에 해당 이름(이름·성명·이용자 열)이 없거나, "
                    "철자·띄어쓰기가 다를 수 있습니다."
                )
            else:
                try:
                    df = sort_data(df, members)
                    st.session_state.output_df = build_final_output_table(df)
                except ValueError as e:
                    st.error(str(e))

# -------------------------------
# 미리보기 + 다운로드 (최종 파일과 동일한 표)
# -------------------------------
if "output_df" in st.session_state:

    st.subheader("📋 미리보기 (다운로드 파일과 동일)")
    edited_df = st.data_editor(
        st.session_state.output_df,
        use_container_width=True,
        key="output_preview_editor",
    )
    # 미리보기 편집값을 항상 세션에 반영해서 다운로드와 1:1로 맞춘다.
    st.session_state.output_df = edited_df

    data_only = st.session_state.output_df[
        st.session_state.output_df["이름"].astype(str).str.strip() != "합계"
    ]
    _parts = []
    if len(data_only) and "이용요금" in st.session_state.output_df.columns:
        _tot = data_only["이용요금"].map(parse_amount_numeric).sum()
        _parts.append(f"총액 {_tot:,.0f}원" if _tot == int(_tot) else f"총액 {_tot:,.1f}원")
    if len(data_only) and "주간/심야" in st.session_state.output_df.columns:
        _sx = data_only["주간/심야"].astype(str).str.strip()
        _sx = _sx.replace({"야근": "심야", "야간": "심야"})
        _parts.append(f"주간 {(_sx == '주간').sum()}건 · 심야 {(_sx == '심야').sum()}건")
    if _parts:
        st.caption(" · ".join(_parts))

    if st.button("🔄 수정 반영"):
        st.session_state.output_df = edited_df

    y, m = usage_year_month_from_output(st.session_state.output_df)
    filename = f"{dept_name} {y}.{m:02d}월 업무택시 내역.xlsx"

    try:
        output_bytes = apply_template(form_file, st.session_state.output_df, dept_name)
    except RuntimeError as e:
        st.error(str(e))
    else:
        st.download_button(
            "📥 파일 다운로드",
            data=output_bytes,
            file_name=filename
        )
