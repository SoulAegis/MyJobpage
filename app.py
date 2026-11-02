"""스탠다드 뱅크 인증 프로젝트 접수·진행현황 대시보드 (교육용 가상 데이터)

실행: streamlit run app.py
- 외부 시스템 연동, 메일 발송 기능 없음 (로컬 CSV 조회 전용)
- 시험분야 변경 내용은 session_state 에만 보관 (원본 CSV 는 수정하지 않음)
"""
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
# 과제 지정 파일명 우선, 실제 배포된 파일명도 함께 탐색
CANDIDATE_FILES = ["SB_projects_demo.csv", "SB_project_demo.csv"]

COLUMNS = [
    "의뢰번호", "고객사", "제품명", "모델명", "시험분야", "적용규격",
    "담당엔지니어", "접수일", "완료예정일", "진행상태", "시료상태", "메모",
]
STATUSES = ["접수", "자료확인", "시험진행", "성적서검토", "보완요청", "완료"]
FILTER_FIELDS = ["시험분야", "적용규격", "담당엔지니어", "진행상태", "시료상태"]
DEFAULT_BASE_DATE = date(2026, 10, 23)
DEFAULT_CATEGORIES = ["EMC", "무선(RF)", "인증서류", "SAR"]
MISSING_LABEL = "(미입력)"

# 일정 상태 라벨
SCHED_DONE = "완료"
SCHED_LATE = "지연 후보(일정 확인 필요)"
SCHED_TODAY = "오늘 마감"
SCHED_PLANNED = "진행 예정"
SCHED_UNKNOWN = "일정 확인 필요"
SCHED_ORDER = [SCHED_LATE, SCHED_TODAY, SCHED_PLANNED, SCHED_UNKNOWN, SCHED_DONE]

# Excel
EXCEL_NOTICE = "교육용 검토 파일 — 공식 보고서 아님"
EXCEL_PREFIX = "standardbank_project_status"

NAVY = "#0B2545"
BLUE = "#1D4ED8"
LIGHT_BLUE = "#EAF1FB"
LATE_BG = "#FDE2E2"
LATE_TEXT = "#9B1C1C"
STATUS_COLORS = ["#0B2545", "#1D4ED8", "#3B82F6", "#60A5FA", "#93C5FD", "#94A3B8"]

st.set_page_config(
    page_title="인증 프로젝트 진행현황",
    page_icon="📋",
    layout="wide",
)

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: #FFFFFF; }}
    h1, h2, h3 {{ color: {NAVY}; }}
    .demo-banner {{
        background: {LIGHT_BLUE};
        border-left: 5px solid {BLUE};
        color: {NAVY};
        padding: 10px 16px;
        border-radius: 6px;
        font-weight: 600;
        margin-bottom: 12px;
    }}
    .card {{
        background: #FFFFFF;
        border: 1px solid #D6E2F3;
        border-top: 4px solid {BLUE};
        border-radius: 10px;
        padding: 14px 12px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(11, 37, 69, 0.08);
    }}
    .card.total {{ background: {NAVY}; border-top-color: {NAVY}; }}
    .card .label {{ font-size: 0.9rem; color: #4B5B73; }}
    .card .value {{ font-size: 1.9rem; font-weight: 700; color: {NAVY}; }}
    .card.total .label {{ color: #C9D6EA; }}
    .card.total .value {{ color: #FFFFFF; }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# 데이터 로딩 / 계산
# ---------------------------------------------------------------------------
def find_data_file() -> Path | None:
    for name in CANDIDATE_FILES:
        path = DATA_DIR / name
        if path.exists():
            return path
    return None


@st.cache_data
def load_csv(source) -> pd.DataFrame:
    try:
        df = pd.read_csv(source, encoding="utf-8-sig", dtype=str)
    except UnicodeDecodeError:
        if hasattr(source, "seek"):
            source.seek(0)
        df = pd.read_csv(source, encoding="cp949", dtype=str)
    df.columns = df.columns.str.strip()
    for col in ["접수일", "완료예정일"]:
        if col in df.columns:
            # 빈 값·잘못된 날짜 형식은 NaT 로 변환
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def schedule_status(progress: str, remaining) -> str:
    """진행상태와 잔여일수(완료예정일 - 기준일)로 일정 상태를 판정한다."""
    if progress == "완료":
        return SCHED_DONE
    if pd.isna(remaining):
        return SCHED_UNKNOWN
    if remaining < 0:
        return SCHED_LATE
    if remaining == 0:
        return SCHED_TODAY
    return SCHED_PLANNED


def card(label: str, value: int, total: bool = False) -> str:
    cls = "card total" if total else "card"
    return (
        f'<div class="{cls}"><div class="label">{label}</div>'
        f'<div class="value">{value}</div></div>'
    )


def highlight_late(row: pd.Series) -> list[str]:
    if row["일정상태"] == SCHED_LATE:
        return [f"background-color: {LATE_BG}; color: {LATE_TEXT}"] * len(row)
    return [""] * len(row)


def display_table(data: pd.DataFrame) -> pd.DataFrame:
    """화면·Excel 공용 목록 형태 (날짜 문자열, 잔여일수/일정상태 열 삽입)."""
    table = data.copy()
    for col in ["접수일", "완료예정일"]:
        table[col] = table[col].dt.strftime("%Y-%m-%d").fillna("")
    return table[COLUMNS[:9] + ["잔여일수", "일정상태"] + COLUMNS[9:]]


# ---------------------------------------------------------------------------
# 시험분야(카테고리) 관리 — session_state 기반
# ---------------------------------------------------------------------------
def init_state(df: pd.DataFrame, source_key: str) -> None:
    """데이터 출처가 바뀌었을 때만 작업용 데이터·카테고리를 초기화한다."""
    if st.session_state.get("source_key") == source_key:
        return
    st.session_state.source_key = source_key
    st.session_state.projects = df
    categories = list(DEFAULT_CATEGORIES)
    # 데이터에만 있는 분야도 목록에 포함 (관리 대상에서 빠지지 않도록)
    for value in df["시험분야"].unique():
        if value != MISSING_LABEL and value not in categories:
            categories.append(value)
    st.session_state.categories = categories
    st.session_state.change_log = []


def set_message(kind: str, text: str) -> None:
    st.session_state.cat_msg = (kind, text)


def category_usage(name: str) -> int:
    return int((st.session_state.projects["시험분야"] == name).sum())


def log_change(text: str) -> None:
    st.session_state.change_log.append(f"{datetime.now():%H:%M:%S} {text}")


def add_category() -> None:
    name = st.session_state.cat_new.strip()
    if not name:
        set_message("warning", "등록할 시험분야 이름을 입력해 주세요.")
    elif name in st.session_state.categories:
        set_message("warning", f"'{name}'은(는) 이미 등록된 시험분야입니다.")
    else:
        st.session_state.categories.append(name)
        log_change(f"등록: {name}")
        set_message("success", f"'{name}' 시험분야를 등록했습니다.")
        st.session_state.cat_new = ""


def rename_category() -> None:
    old = st.session_state.cat_rename_src
    new = st.session_state.cat_rename_dst.strip()
    categories = st.session_state.categories
    if not new:
        set_message("warning", "새 이름을 입력해 주세요.")
    elif new == old:
        set_message("warning", "기존 이름과 같습니다.")
    elif new in categories:
        set_message(
            "warning",
            f"'{new}'은(는) 이미 있는 시험분야입니다. 합치려면 삭제 메뉴의 재분류를 이용하세요.",
        )
    else:
        categories[categories.index(old)] = new
        projects = st.session_state.projects
        mask = projects["시험분야"] == old
        projects.loc[mask, "시험분야"] = new
        log_change(f"이름 변경: {old} → {new} (의뢰 {int(mask.sum())}건 갱신)")
        set_message(
            "success",
            f"'{old}' → '{new}' 변경 완료. 관련 의뢰 {int(mask.sum())}건도 함께 갱신했습니다.",
        )
        st.session_state.cat_rename_dst = ""
        st.session_state.pop("cat_rename_src", None)


def reclassify_category() -> None:
    src = st.session_state.cat_delete_src
    dst = st.session_state.cat_reclass_dst
    projects = st.session_state.projects
    mask = projects["시험분야"] == src
    projects.loc[mask, "시험분야"] = dst
    log_change(f"재분류: {src} → {dst} (의뢰 {int(mask.sum())}건)")
    set_message(
        "success",
        f"'{src}' 의뢰 {int(mask.sum())}건을 '{dst}'(으)로 재분류했습니다. 이제 삭제할 수 있습니다.",
    )


def delete_category() -> None:
    name = st.session_state.cat_delete_src
    used = category_usage(name)
    if used:
        set_message("warning", f"'{name}'은(는) 의뢰 {used}건에서 사용 중이라 삭제할 수 없습니다.")
        return
    st.session_state.categories.remove(name)
    log_change(f"삭제: {name}")
    set_message("success", f"'{name}' 시험분야를 삭제했습니다.")
    st.session_state.pop("cat_delete_src", None)


def reset_changes() -> None:
    st.session_state.pop("source_key", None)
    set_message("success", "원본 데이터와 기본 시험분야로 초기화했습니다.")


# ---------------------------------------------------------------------------
# Excel 내보내기
# ---------------------------------------------------------------------------
def count_table(data: pd.DataFrame, row: str, col: str, col_order: list[str]) -> pd.DataFrame:
    """row × col 건수 교차표 (+ 합계 열). 빈 데이터에서도 동작."""
    table = (
        data.groupby([row, col]).size().unstack(fill_value=0)
        if not data.empty
        else pd.DataFrame()
    )
    extra = [c for c in table.columns if c not in col_order]
    table = table.reindex(columns=col_order + extra, fill_value=0)
    table.insert(0, "합계", table.sum(axis=1).astype(int))
    table.index.name = row
    return table.reset_index()


def build_excel(data: pd.DataFrame, conditions: list[tuple[str, str]]) -> bytes:
    status_order = [s for s in STATUSES if s in set(data["진행상태"])] + sorted(
        set(data["진행상태"]) - set(STATUSES)
    )
    status_sheet = count_table(data, "진행상태", "일정상태", SCHED_ORDER)
    status_sheet = (
        status_sheet.set_index("진행상태").reindex(status_order).reset_index()
    )
    owner_sheet = count_table(data, "담당엔지니어", "진행상태", STATUSES)
    owner_sheet["지연 후보"] = owner_sheet["담당엔지니어"].map(
        data[data["일정상태"] == SCHED_LATE]["담당엔지니어"].value_counts()
    ).fillna(0).astype(int)
    condition_sheet = pd.DataFrame(conditions, columns=["항목", "값"])

    sheets = {
        "프로젝트목록": display_table(data),
        "상태별집계": status_sheet,
        "담당자별집계": owner_sheet,
        "분석조건": condition_sheet,
    }
    header_fill = PatternFill("solid", fgColor=NAVY.lstrip("#"))
    late_fill = PatternFill("solid", fgColor=LATE_BG.lstrip("#"))

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            # 1행: 안내 문구, 3행부터 표
            frame.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
            ws = writer.sheets[sheet_name]
            ws["A1"] = EXCEL_NOTICE
            ws["A1"].font = Font(bold=True, color="C00000", size=12)
            for cell in ws[3]:
                cell.fill = header_fill
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center")
            for idx, column in enumerate(frame.columns, start=1):
                width = max([len(str(column))] + [len(str(v)) for v in frame[column]])
                ws.column_dimensions[get_column_letter(idx)].width = min(width * 1.8 + 2, 50)
            ws.freeze_panes = "A4"
            if sheet_name == "프로젝트목록":
                for offset, sched in enumerate(frame["일정상태"]):
                    if sched == SCHED_LATE:
                        for cell in ws[4 + offset]:
                            cell.fill = late_fill
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# 화면
# ---------------------------------------------------------------------------
st.title("📋 인증 프로젝트 접수·진행현황 대시보드")
st.markdown(
    '<div class="demo-banner">⚠️ 교육용 가상 데이터 기준 — 실제 업무 시스템 아님</div>',
    unsafe_allow_html=True,
)

data_path = find_data_file()
if data_path is not None:
    raw = load_csv(data_path)
    source_label = f"data/{data_path.name}"
    source_key = str(data_path)
else:
    st.info(
        "data 폴더에서 프로젝트 데이터 파일(SB_projects_demo.csv)을 찾을 수 없습니다. "
        "아래에서 CSV 파일을 업로드해 주세요."
    )
    uploaded = st.file_uploader("프로젝트 CSV 업로드", type=["csv"])
    if uploaded is None:
        st.stop()
    raw = load_csv(uploaded)
    source_label = f"업로드: {uploaded.name}"
    source_key = f"upload:{uploaded.name}:{uploaded.size}"

missing = [c for c in COLUMNS if c not in raw.columns]
if missing:
    st.warning(f"CSV에 다음 필드가 없습니다: {', '.join(missing)}")
    st.stop()

base = raw[COLUMNS].copy()
for col in FILTER_FIELDS:
    base[col] = base[col].fillna(MISSING_LABEL).str.strip()
init_state(base, source_key)

st.caption(
    "💾 시험분야 등록·변경·재분류 내용은 현재 브라우저 세션(session_state)에만 보관됩니다. "
    "새로고침하거나 세션이 종료되면 사라질 수 있으며, 원본 CSV 파일은 수정되지 않습니다."
)

# ---------------- 사이드바 ----------------
with st.sidebar:
    st.header("⚙️ 조회 설정")
    base_date = st.date_input("기준일", value=DEFAULT_BASE_DATE)
    st.caption(f"데이터 출처: {source_label}")

    # ---- 필터 영역 ----
    st.divider()
    st.header("🔎 필터")
    st.caption("선택하지 않으면 전체가 조회됩니다.")
    projects = st.session_state.projects
    selections = {}
    for col in FILTER_FIELDS:
        values = list(projects[col].unique())
        if col == "시험분야":
            options = st.session_state.categories + sorted(
                v for v in values if v not in st.session_state.categories
            )
        elif col == "진행상태":
            options = [s for s in STATUSES if s in values] + sorted(
                v for v in values if v not in STATUSES
            )
        else:
            options = sorted(values)
        selections[col] = st.multiselect(col, options, placeholder="전체")

    # ---- 카테고리 관리 영역 ----
    st.divider()
    st.header("🗂️ 카테고리 관리")
    st.caption("시험분야 등록 / 이름 변경 / 삭제")

    if "cat_msg" in st.session_state:
        kind, text = st.session_state.pop("cat_msg")
        getattr(st, kind)(text)

    categories = st.session_state.categories
    st.markdown(
        " · ".join(f"**{c}** {category_usage(c)}건" for c in categories)
        or "등록된 시험분야가 없습니다."
    )

    action = st.radio(
        "작업", ["등록", "이름 변경", "삭제"], horizontal=True, key="cat_action"
    )
    if action == "등록":
        st.text_input("새 시험분야 이름", key="cat_new")
        st.button("등록", on_click=add_category, type="primary")

    elif action == "이름 변경":
        if categories:
            st.selectbox("변경할 시험분야", categories, key="cat_rename_src")
            st.text_input("새 이름", key="cat_rename_dst")
            st.button("이름 변경", on_click=rename_category, type="primary")
            st.caption("이름을 바꾸면 해당 분야의 의뢰 데이터도 함께 갱신됩니다.")
        else:
            st.caption("변경할 시험분야가 없습니다.")

    else:  # 삭제
        if categories:
            target = st.selectbox("삭제할 시험분야", categories, key="cat_delete_src")
            used = category_usage(target)
            others = [c for c in categories if c != target]
            if used:
                st.warning(
                    f"'{target}'은(는) 의뢰 {used}건에서 사용 중이라 바로 삭제할 수 없습니다. "
                    "아래에서 다른 분야로 재분류한 뒤 삭제해 주세요."
                )
                if others:
                    st.selectbox("재분류할 분야", others, key="cat_reclass_dst")
                    st.button("재분류 실행", on_click=reclassify_category)
                else:
                    st.caption("재분류할 다른 분야가 없습니다. 먼저 새 분야를 등록해 주세요.")
            else:
                st.button("삭제", on_click=delete_category, type="primary")
        else:
            st.caption("삭제할 시험분야가 없습니다.")

    if st.session_state.change_log:
        with st.expander(f"변경 이력 ({len(st.session_state.change_log)}건)"):
            for line in st.session_state.change_log:
                st.text(line)
    st.button("변경 내용 초기화", on_click=reset_changes)

# ---------------- 계산 ----------------
df = st.session_state.projects.copy()
base_ts = pd.Timestamp(base_date)
df["잔여일수"] = (df["완료예정일"] - base_ts).dt.days.astype("Int64")
df["일정상태"] = [
    schedule_status(p, r) for p, r in zip(df["진행상태"], df["잔여일수"])
]

# 필터 적용 — 이후 카드·표·차트는 모두 filtered 기준
filtered = df
for col, selected in selections.items():
    if selected:
        filtered = filtered[filtered[col].isin(selected)]

# 요약 카드
st.subheader(f"요약 (기준일 {base_date:%Y-%m-%d})")
counts = filtered["진행상태"].value_counts()
cols = st.columns(len(STATUSES) + 1)
cols[0].markdown(card("전체", len(filtered), total=True), unsafe_allow_html=True)
for col, status in zip(cols[1:], STATUSES):
    col.markdown(card(status, int(counts.get(status, 0))), unsafe_allow_html=True)

if filtered.empty:
    st.info("조건에 맞는 인증 프로젝트가 없습니다")
else:
    sched_counts = filtered["일정상태"].value_counts()
    st.markdown(
        f"<p style='color:{NAVY};margin-top:10px;'>"
        f"<b style='color:{LATE_TEXT};'>{SCHED_LATE}: {int(sched_counts.get(SCHED_LATE, 0))}건</b>"
        f" · {SCHED_TODAY}: {int(sched_counts.get(SCHED_TODAY, 0))}건"
        f" · {SCHED_UNKNOWN}: {int(sched_counts.get(SCHED_UNKNOWN, 0))}건</p>",
        unsafe_allow_html=True,
    )

    # 차트
    chart_left, chart_right = st.columns(2)

    with chart_left:
        st.subheader("시험분야별 건수")
        field_counts = filtered["시험분야"].value_counts().reset_index()
        field_counts.columns = ["시험분야", "건수"]
        bar = (
            alt.Chart(field_counts)
            .mark_bar(color=BLUE, cornerRadiusEnd=4)
            .encode(
                x=alt.X("건수:Q", axis=alt.Axis(tickMinStep=1)),
                y=alt.Y("시험분야:N", sort="-x", title=None),
                tooltip=["시험분야", "건수"],
            )
            .properties(height=280)
        )
        st.altair_chart(bar)

    with chart_right:
        st.subheader("진행상태 비율")
        status_counts = filtered["진행상태"].value_counts().reset_index()
        status_counts.columns = ["진행상태", "건수"]
        status_counts["비율"] = status_counts["건수"] / status_counts["건수"].sum()
        donut = (
            alt.Chart(status_counts)
            .mark_arc(innerRadius=60)
            .encode(
                theta=alt.Theta("건수:Q"),
                color=alt.Color(
                    "진행상태:N",
                    scale=alt.Scale(domain=STATUSES, range=STATUS_COLORS),
                    sort=STATUSES,
                ),
                tooltip=["진행상태", "건수", alt.Tooltip("비율:Q", format=".0%")],
            )
            .properties(height=280)
        )
        st.altair_chart(donut)

    # 의뢰 목록
    st.subheader(f"의뢰 목록 ({len(filtered)}건)")
    st.dataframe(
        display_table(filtered).style.apply(highlight_late, axis=1),
        hide_index=True,
        column_config={
            "잔여일수": st.column_config.NumberColumn(
                "잔여일수", help="완료예정일 - 기준일 (일)"
            ),
        },
    )
    st.caption("빨간색 행은 '지연 후보(일정 확인 필요)'입니다.")

# ---------------- Excel 다운로드 ----------------
st.subheader("📥 Excel 다운로드")
today = date.today()
generated_at = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
common_conditions = [
    ("안내", EXCEL_NOTICE),
    ("기준일", f"{base_date:%Y-%m-%d}"),
    ("생성일시", generated_at),
    ("데이터 출처", source_label),
    ("시험분야 목록", ", ".join(st.session_state.categories)),
    ("시험분야 변경 이력", " / ".join(st.session_state.change_log) or "없음"),
]
filter_conditions = [
    (f"필터: {col}", ", ".join(selected) if selected else "전체")
    for col, selected in selections.items()
]

dl_all, dl_filtered = st.columns(2)
with dl_all:
    st.download_button(
        f"전체 목록 다운로드 ({len(df)}건)",
        data=build_excel(
            df,
            common_conditions
            + [("범위", "전체 목록 (필터 미적용)"), ("건수", f"{len(df)}건")],
        ),
        file_name=f"{EXCEL_PREFIX}_{today:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
with dl_filtered:
    st.download_button(
        f"현재 필터 결과 다운로드 ({len(filtered)}건)",
        data=build_excel(
            filtered,
            common_conditions
            + [("범위", "현재 필터 결과"), ("건수", f"{len(filtered)}건")]
            + filter_conditions,
        ),
        file_name=f"{EXCEL_PREFIX}_filtered_{today:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=filtered.empty,
    )
st.caption(
    "파일명의 날짜는 다운로드한 실제 날짜입니다. 각 시트 상단에 "
    f"'{EXCEL_NOTICE}' 문구가 표시되며, 분석조건 시트에 기준일과 적용한 필터가 기록됩니다. "
    "본 화면은 로컬 CSV 조회 전용이며, 메일 발송·외부 시스템 연동 기능은 없습니다."
)
