import io
from datetime import datetime

import streamlit as st

from datetime import date, timedelta

from schedule_parser import (
    parse_full_schedule,
    diff_entries,
    build_change_ranges,
    write_upload_excel,
    entry_dt,
)
import schedule_store as store

st.set_page_config(page_title="SPOTV 편성표 → 업로드파일 변환", page_icon="📺", layout="centered")

st.title("📺 SPOTV 편성표 → 업로드파일 변환기")
st.caption("편성표 엑셀을 업로드파일 형식으로 변환하거나, 두 편성표를 비교해 변경구간을 뽑고 업로드파일을 수정합니다.")

mode = st.radio(
    "작업 선택",
    ["편성표 → 업로드파일 변환", "변경 편성표 비교 (기존 vs 신규)", "채널별 누적 관리 (마스터 → 주차 파일)"],
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

        review_entries = [e for e in entries if e.get('needs_review')]
        if review_entries:
            st.warning(
                f"⚠️ {len(review_entries)}건은 해당 날짜 열의 첫 방송이 테두리 없이 시작되어 "
                f"정확한 시작 시각을 파일만으로 확정할 수 없습니다(기본 30분 앞당겨 추정). "
                f"아래 목록과 다운로드 파일의 주황색 셀을 실제 편성 사실과 대조해 확인해주세요."
            )
            for e in review_entries:
                st.markdown(f"- **{e['date'].isoformat()} {e['hour']:02d}:{e['minute']:02d}** — {e['title']}")

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

elif mode.startswith("채널별"):
    st.subheader("채널별 누적 마스터 · 주차 파일 생성")
    st.caption("편성표를 올릴 때마다 채널 마스터에 쌓고, 원하는 주(월요일 기준)·기준시각으로 잘라 업로드파일을 만듭니다.")

    channel = st.selectbox("채널", list(store.CHANNELS.keys()))
    bases = store.CHANNELS[channel]['bases']
    st.caption("작성 기준 시각: " + " / ".join(f"{b:02d}:00" for b in bases) + (" (2가지 버전 생성)" if len(bases) > 1 else ""))

    st.markdown("**1) 마스터 불러오기** — 처음이면 비워두세요")
    master_file = st.file_uploader("해당 채널의 누적 마스터(.xlsx)", type=["xlsx"], key="master")

    skey = f"master::{channel}"
    fid = (master_file.name, master_file.size) if master_file else None
    if skey not in st.session_state or st.session_state.get(skey + "::src") != fid:
        if master_file is not None:
            try:
                m = store.load_master(master_file.getvalue())
            except Exception as e:
                st.error(f"마스터를 읽지 못했습니다: {e}")
                st.stop()
            if m['channel'] and m['channel'] != channel:
                st.error(f"이 마스터는 '{m['channel']}' 채널용입니다. 채널 선택을 확인해주세요.")
                st.stop()
            m['channel'] = channel
        else:
            m = store.new_master(channel)
        st.session_state[skey] = m
        st.session_state[skey + "::src"] = fid
        st.session_state[skey + "::report"] = None
    master = st.session_state[skey]
    st.info(f"현재 마스터: {len(master['entries'])}건 · 적용 이력 {len(master['log'])}회"
            + (f" · 적용 구간 {', '.join(store.fmt_gap(c) for c in store.coverage(master))}" if master['log'] else ""))

    st.markdown("**2) 편성표 적용** — 해당 주 구간이 통째로 교체됩니다(변경 편성표 재적용 = 개정). 다른 주는 그대로입니다. 여러 파일은 파일명 순(날짜 순)으로 적용됩니다.")
    sched_files = st.file_uploader("주간 편성표(.xlsx)", type=["xlsx"], key="store_sched", accept_multiple_files=True)
    if sched_files and st.button("마스터에 적용"):
        reports = []
        for sf in sorted(sched_files, key=lambda f: f.name):  # 파일명(YYMMDD) 순, 나중 파일이 우선
            try:
                fb = sf.getvalue()
                ents, _ = parse_full_schedule(fb)
                win = store.grid_window(fb)
                rep = store.apply_schedule(master, ents, win, sf.name)
                rep['name'] = sf.name
                reports.append(rep)
            except Exception as e:
                st.error(f"{sf.name} 적용 실패: {e}")
        st.session_state[skey + "::report"] = reports
    for rep in (st.session_state.get(skey + "::report") or []):
        s0, e0 = rep['window']
        c = rep['counts']
        st.success(f"{rep['name']} 적용 — {s0:%m/%d %H:%M} ~ {e0:%m/%d %H:%M} · "
                   + ("신규 " + str(rep['new_count']) + "건" if rep['first_apply']
                      else f"추가 {c['ADDED']} / 삭제 {c['REMOVED']} / 변경 {c['CHANGED']}"))
        for n in rep['notes']:
            st.warning(n)
        if not rep['first_apply'] and rep['changes']:
            in_win = lambda e: s0 <= entry_dt(e) < e0
            new_sorted = sorted([e for e in master['entries'] if in_win(e)], key=entry_dt)
            old_sorted = sorted([x['old'] for x in rep['changes'] if x['old']], key=entry_dt)
            report, _ = build_change_ranges(rep['changes'], new_sorted, old_sorted)
            st.code(report, language="markdown")

    if master['entries']:
        ch_name = channel.replace(" ", "_")
        st.download_button("📥 갱신된 마스터 다운로드 (다음에 다시 올려 쓰세요)", data=store.master_bytes(master),
                           file_name=f"{ch_name}_누적마스터.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("**3) 주차 파일 생성**")
        last_start = max(g['start'] for g in master['log']) if master['log'] else None
        default_mon = store.monday_of(last_start.date()) if last_start else store.monday_of(date.today())
        pick = st.date_input("주 시작일(월요일이 아니면 해당 주 월요일로 보정)", value=default_mon)
        monday = store.monday_of(pick)
        y, mo, n = store.week_label(monday)
        st.write(f"대상: **{y}년 {mo}월 {n}주차** ({monday:%m/%d} 월 시작)")
        for b in bases:
            sl = store.slice_week(master, monday, b)
            st.markdown(f"##### {b:02d}:00 기준 — {sl['start']:%m/%d %H:%M} ~ {sl['end']:%m/%d %H:%M}")
            if sl['gaps']:
                st.error("마스터에 데이터가 없는 구간이 있습니다: " + " / ".join(store.fmt_gap(g) for g in sl['gaps'])
                         + " → 해당 구간의 편성표를 먼저 적용해야 완전한 파일이 됩니다.")
            for note in sl['notes']:
                st.warning(note)
            if sl['review']:
                st.warning(f"시작 시각 확인 필요 {len(sl['review'])}건 (주황색): "
                           + ", ".join(f"{entry_dt(e):%m/%d %H:%M}" for e in sl['review']))
            if sl['entries']:
                st.write(f"{len(sl['entries'])}건")
                st.download_button(f"📥 {b:02d}시 기준 업로드파일 다운로드",
                                   data=store.week_bytes(sl['entries'], channel, b),
                                   file_name=store.week_filename(channel, monday, b),
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key=f"dl_{b}")
            else:
                st.info("이 구간에 해당하는 편성이 없습니다.")
    else:
        st.info("편성표를 적용하면 주차 파일을 만들 수 있습니다.")

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

        review_entries = [e for e in new_entries if e.get('needs_review')]
        if review_entries:
            st.warning(
                f"⚠️ {len(review_entries)}건은 날짜 열 첫 방송이 테두리 없이 시작되어 "
                f"정확한 시작 시각을 파일만으로 확정할 수 없습니다(기본 30분 앞당겨 추정). "
                f"다운로드 파일의 주황색 셀을 실제 편성 사실과 대조해 확인해주세요."
            )
            for e in review_entries:
                st.markdown(f"- **{e['date'].isoformat()} {e['hour']:02d}:{e['minute']:02d}** — {e['title']}")

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
- **시간축**: B열의 시(03~24~01~02) 라벨과 셀 위치로 30분 단위 시각을 추정하되, 셀 안에 `생중계`/`생방송` 문구가 있으면(앞에 HH:MM이 있든 없든) 그 시각을 우선하고, 해당 단어는 프로그램명에서 제거하고 방송구분은 LIVE로 확정합니다.
- **항목 경계**: 프로그램 블록이 새 항목인지 직전 항목의 연속(대진정보 등)인지는 **셀 테두리 유무**로 판정합니다.
- **방송구분**: 셀 배경색(테마)으로 판정합니다 — 현장생중계·수신생중계=LIVE, 본방송=본방송, 배경없음=재방송. (`생중계/생방송` 문구는 LIVE, `녹화중계` 문구는 `녹화방송`으로 확정하며 두 문구 모두 프로그램명에서는 제거)
- **자막/해설/수어**: 셀 값이 아니라 셀 위의 작은 도형(텍스트박스)이며, 파일 내부 구조를 직접 읽어 프로그램과 매칭합니다.
- **24시 이후(24,01,02)**: 다음 날 새벽으로 자동 처리합니다.
- **연령고지**: `⑮` 등 원문자 등급 표기는 프로그램명에서 제거합니다.
- **⚠️ 알려진 한계**: 어떤 날짜의 첫 방송이 테두리 없이 시작되는 경우, 그리드가 보여주는 시각보다 실제로는 더 일찍 시작된 것인데 정확히 얼마나 앞당겨야 하는지는 파일 구조만으로 항상 안전하게 알 수 없습니다(같은 신호가 파일마다 30분 앞일 수도, 1시간 앞일 수도 있음이 확인됨). 이런 항목은 기본 30분으로 추정하고 화면과 출력 파일(주황색 셀)에 표시하니, 실제 편성 사실과 대조해 필요하면 직접 수정해주세요.

이 로직은 규칙 기반이며 AI를 호출하지 않으므로, 매번 결과가 동일하고 빠릅니다.
""")
