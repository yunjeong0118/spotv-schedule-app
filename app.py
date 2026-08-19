import io
from datetime import datetime

import streamlit as st

from schedule_parser import (
    parse_full_schedule,
    diff_entries,
    build_change_ranges,
    write_upload_excel,
    entry_dt,
)

st.set_page_config(page_title="SPOTV 편성표 → 업로드파일 변환", page_icon="📺", layout="centered")

st.title("📺 SPOTV 편성표 → 업로드파일 변환기")
st.caption("편성표 엑셀을 업로드파일 형식으로 변환하거나, 두 편성표를 비교해 변경구간을 뽑고 업로드파일을 수정합니다.")

mode = st.radio(
    "작업 선택",
    ["편성표 → 업로드파일 변환", "변경 편성표 비교 (기존 vs 신규)"],
    horizontal=False,
)

st.divider()


def sorted_entries(entries):
    return sorted(entries, key=lambda e: (e['date'], e['hour'], e['minute']))


def to_excel_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


if mode == "편성표 → 업로드파일 변환":
    st.subheader("1) 편성표 파일 업로드")
    sched_file = st.file_uploader("SPOTV 주간 편성표(.xlsx)", type=["xlsx"], key="single")

    if sched_file is not None:
        with st.spinner("편성표를 분석하는 중..."):
            file_bytes = sched_file.read()
            try:
                entries, unmatched = parse_full_schedule(file_bytes)
            except Exception as e:
                st.error(f"파싱 중 오류가 발생했습니다: {e}")
                st.stop()

        entries = sorted_entries(entries)
        st.success(f"총 {len(entries)}건의 편성 항목을 인식했습니다.")

        if unmatched:
            st.warning(f"자막/해설/수어 표기 중 {len(unmatched)}건은 프로그램에 정확히 매칭하지 못했습니다. "
                       f"(파일 서식이 크게 다른 경우 발생할 수 있습니다)")

        with st.expander("인식된 편성 미리보기 (처음 20건)"):
            preview_rows = []
            for e in entries[:20]:
                flags = e.get('flags', set())
                preview_rows.append({
                    "날짜": e['date'].isoformat(),
                    "시각": f"{e['hour']:02d}:{e['minute']:02d}",
                    "프로그램명": e['title'],
                    "구분": e['cat'],
                    "자": "Y" if "자" in flags else "",
                    "해": "Y" if "해" in flags else "",
                    "수": "Y" if "수" in flags else "",
                })
            st.dataframe(preview_rows, use_container_width=True)

        wb = write_upload_excel(entries)
        buf = to_excel_bytes(wb)

        out_name = sched_file.name.replace(".xlsx", "") + "_업로드파일.xlsx"
        st.download_button(
            "📥 업로드파일 다운로드",
            data=buf,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.subheader("1) 기존 편성표와 변경(신규) 편성표를 각각 업로드")
    col1, col2 = st.columns(2)
    with col1:
        old_file = st.file_uploader("기존 편성표(.xlsx)", type=["xlsx"], key="old")
    with col2:
        new_file = st.file_uploader("변경 편성표(.xlsx)", type=["xlsx"], key="new")

    if old_file is not None and new_file is not None:
        with st.spinner("두 편성표를 분석하고 비교하는 중..."):
            old_bytes = old_file.read()
            new_bytes = new_file.read()
            try:
                old_entries, _ = parse_full_schedule(old_bytes)
                new_entries, _ = parse_full_schedule(new_bytes)
            except Exception as e:
                st.error(f"파싱 중 오류가 발생했습니다: {e}")
                st.stop()

        old_sorted = sorted_entries(old_entries)
        new_sorted = sorted_entries(new_entries)

        changes = diff_entries(old_entries, new_entries)

        if not changes:
            st.info("두 편성표 사이에 차이가 없습니다.")
        else:
            report, details = build_change_ranges(changes, new_sorted, old_sorted)

            st.subheader("2) 변경구간")
            st.code(report, language="markdown")

            with st.expander("세부 변경 내역"):
                for d, span, desc in details:
                    st.markdown(f"- **{d.month}/{d.day} {span}** — {desc}")

        st.subheader("3) 변경 반영된 업로드파일")
        wb = write_upload_excel(new_entries, old_entries=old_entries)
        buf = to_excel_bytes(wb)

        out_name = new_file.name.replace(".xlsx", "") + "_업로드파일_변경반영.xlsx"
        st.download_button(
            "📥 업로드파일 다운로드 (변경 셀 노란색 표시)",
            data=buf,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("기존 편성표와 변경 편성표 파일을 모두 업로드하면 비교가 시작됩니다.")

st.divider()
with st.expander("ℹ️ 이 도구가 판정하는 규칙"):
    st.markdown("""
- **시간축**: B열의 시(03~24~01~02) 라벨과 셀 위치로 30분 단위 시각을 추정하되, 셀 안에 `HH:MM 생중계/생방송` 문구가 있으면 그 시각을 우선합니다.
- **항목 경계**: 프로그램 블록이 새 항목인지 직전 항목의 연속(대진정보 등)인지는 **셀 테두리 유무**로 판정합니다.
- **방송구분**: 셀 배경색(테마)으로 본방송/재방송/LIVE를 구분합니다.
- **자막/해설/수어**: 셀 값이 아니라 셀 위의 작은 도형(텍스트박스)이며, 파일 내부 구조를 직접 읽어 프로그램과 매칭합니다.
- **24시 이후(24,01,02)**: 다음 날 새벽으로 자동 처리합니다.

이 로직은 규칙 기반이며 AI를 호출하지 않으므로, 매번 결과가 동일하고 빠릅니다.
""")
