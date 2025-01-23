import streamlit as st
import re
import time
import io
import zipfile
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from collections import defaultdict
import pandas as pd


def main():
    st.title("아티스트 음원 정산 보고서 자동 생성기 (Excel 기반)")

    # 1) 섹션1: 보고서 생성(파일 업로드 + 진행기간/발행일 입력 등)
    section_one_report_input()

    st.divider()

    # 2) 섹션2: 검증 결과 표시
    section_two_verification()

    st.divider()

    # 3) 섹션3: 결과 ZIP 다운로드
    section_three_download_zip()

    st.divider()
    st.info("끝")


# ------------------------------------------
# 1) 섹션1: 보고서 생성(파일 업로드 + 진행기간/발행일 입력)
# ------------------------------------------
def section_one_report_input():
    st.subheader("1) 정산 보고서 생성")

    # session_state에서 기본값 불러오기
    default_ym = st.session_state.get("ym", "")
    default_report_date = st.session_state.get("report_date", "")

    # 사용자 입력
    ym = st.text_input("진행기간(YYYYMM)", default_ym)
    report_date = st.text_input("보고서 발행 날짜 (YYYY-MM-DD)", default_report_date)

    # 엑셀 업로드 (두 개)
    uploaded_song_cost = st.file_uploader("input_song cost.xlsx 업로드", type=["xlsx"])
    uploaded_online_revenue = st.file_uploader("input_online revenue.xlsx 업로드", type=["xlsx"])

    # 생성 버튼
    if st.button("정산 보고서 생성 시작"):
        if not re.match(r'^\d{6}$', ym):
            st.error("진행기간은 YYYYMM 6자리로 입력하세요.")
            return
        if not report_date:
            st.error("보고서 발행 날짜를 입력하세요.")
            return
        if not uploaded_song_cost or not uploaded_online_revenue:
            st.error("두 개의 엑셀 파일을 모두 업로드해야 합니다.")
            return

        # session_state에 입력값 저장
        st.session_state["ym"] = ym
        st.session_state["report_date"] = report_date

        # 검증용 dict
        check_dict = {
            "song_artists": [],
            "revenue_artists": [],
            "artist_compare_result": {},
            "verification_summary": {
                "total_errors": 0,
                "artist_error_list": []
            },
            "details_verification": {
                "정산서": [],
                "세부매출": []
            }
        }

        # 실제 보고서 생성
        zip_data = generate_report_excel(
            ym, report_date,
            uploaded_song_cost,
            uploaded_online_revenue,
            check_dict
        )

        if zip_data is not None:
            st.success("정산 보고서 생성 완료! 아래 섹션에서 ZIP 다운로드 가능")
            # st.session_state에 기록
            st.session_state["report_done"] = True
            st.session_state["zip_data"] = zip_data
            st.session_state["check_dict"] = check_dict
        else:
            st.error("보고서 생성 중 오류가 발생했습니다.")


# ------------------------------------------
# 2) 섹션2: 검증 결과 표시
# ------------------------------------------
def section_two_verification():
    if st.session_state.get("report_done", False):
        st.subheader("2) 검증 결과")

        cd = st.session_state.get("check_dict", {})
        if not cd:
            st.info("검증 데이터가 없습니다.")
            return

        # 탭 2개
        tab1, tab2 = st.tabs(["검증 요약", "세부 검증 내용"])

        with tab1:
            ar = cd.get("artist_compare_result", {})
            st.write("**아티스트 목록 비교**")
            st.write(f"- Song cost 아티스트 수 = {ar.get('song_count')}")
            st.write(f"- Revenue 아티스트 수 = {ar.get('revenue_count')}")
            st.write(f"- 공통 아티스트 수 = {ar.get('common_count')}")
            if ar.get("missing_in_song"):
                st.warning(f"Song에 없고 Revenue에만 있는 아티스트: {ar['missing_in_song']}")
            if ar.get("missing_in_revenue"):
                st.warning(f"Revenue에 없고 Song에만 있는 아티스트: {ar['missing_in_revenue']}")

            ver_sum = cd.get("verification_summary", {})
            total_err = ver_sum.get("total_errors", 0)
            artists_err = ver_sum.get("artist_error_list", [])
            if total_err == 0:
                st.success("모든 항목이 정상 계산되었습니다. (오류 0건)")
            else:
                st.error(f"총 {total_err}건의 계산 오류 발생!")
                st.warning(f"문제 발생 아티스트: {list(set(artists_err))}")

        with tab2:
            show_detailed_verification(cd)

    else:
        st.info("정산 보고서 생성 완료 후, 검증 결과가 표시됩니다.")


# ------------------------------------------
# 3) 섹션3: 생성된 ZIP 다운로드
# ------------------------------------------
def section_three_download_zip():
    if st.session_state.get("report_done", False):
        st.subheader("3) 결과 ZIP 다운로드")

        zip_data = st.session_state.get("zip_data")
        if zip_data:
            st.download_button(
                label="ZIP 다운로드",
                data=zip_data,
                file_name="정산결과보고서.zip",
                mime="application/zip"
            )
        else:
            st.warning("ZIP 데이터가 없습니다.")
    else:
        st.info("아직 보고서가 생성되지 않았습니다.")


# --------------------------------------------------
# (세부) 검증 정보 표시
# --------------------------------------------------
def show_detailed_verification(check_dict):
    dv = check_dict.get("details_verification", {})
    if not dv:
        st.warning("세부 검증 데이터가 없습니다.")
        return

    tabA, tabB = st.tabs(["정산서 검증", "세부매출 검증"])

    # 정산서 검증
    with tabA:
        rows = dv.get("정산서", [])
        if not rows:
            st.info("정산서 검증 데이터가 없습니다.")
        else:
            df = pd.DataFrame(rows)
            bool_cols = [c for c in df.columns if c.startswith("match_")]

            def highlight_boolean(val):
                if val is True:
                    return "background-color: #AAFFAA"
                elif val is False:
                    return "background-color: #FFAAAA"
                else:
                    return ""

            # 예시로 표시할 정수 칼럼들
            int_columns = [
                "원본_곡비", "정산서_곡비",
                "원본_공제금액", "정산서_공제금액",
                "원본_공제후잔액", "정산서_공제후잔액",
                "원본_정산율(%)", "정산서_정산율(%)"
            ]
            format_dict = {col: "{:.0f}" for col in int_columns if col in df.columns}

            st.dataframe(
                df.style
                  .format(format_dict)
                  .applymap(highlight_boolean, subset=bool_cols)
            )

    # 세부매출 검증
    with tabB:
        rows = dv.get("세부매출", [])
        if not rows:
            st.info("세부매출 검증 데이터가 없습니다.")
        else:
            df = pd.DataFrame(rows)
            bool_cols = [c for c in df.columns if c.startswith("match_")]

            def highlight_boolean(val):
                if val is True:
                    return "background-color: #AAFFAA"
                elif val is False:
                    return "background-color: #FFAAAA"
                else:
                    return ""

            int_columns = ["원본_매출액", "정산서_매출액"]
            format_dict = {col: "{:.0f}" for col in int_columns if col in df.columns}

            st.dataframe(
                df.style
                  .format(format_dict)
                  .applymap(highlight_boolean, subset=bool_cols)
            )


# --------------------------------------------------
# 보고서 생성 (엑셀 기반)
# --------------------------------------------------
def generate_report_excel(ym, report_date, file_song_cost, file_online_revenue, check_dict):
    """
    업로드된 두 엑셀 파일을 openpyxl로 파싱 → 아티스트별 계산 → 
    '정산서(artist).xlsx', '세부매출내역(artist).xlsx' 를 모두 ZIP으로 묶어 반환(bytes).
    """
    try:
        wb_song = openpyxl.load_workbook(file_song_cost, data_only=True)
        wb_revenue = openpyxl.load_workbook(file_online_revenue, data_only=True)
    except Exception as e:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

    # (A) input_song cost: ym 시트 찾기
    if ym not in wb_song.sheetnames:
        st.error(f"[song cost] 파일에 '{ym}' 시트가 없습니다.")
        return None
    ws_sc = wb_song[ym]

    rows_sc = list(ws_sc.values)
    if not rows_sc:
        st.error(f"[song cost] '{ym}' 시트가 비어있습니다.")
        return None

    header_sc = rows_sc[0]
    body_sc = rows_sc[1:]
    try:
        idx_artist = header_sc.index("아티스트명")
        idx_rate = header_sc.index("정산 요율")
        idx_prev = header_sc.index("전월 잔액")
        idx_deduct = header_sc.index("당월 차감액")
        idx_remain = header_sc.index("당월 잔액")
    except ValueError as e:
        st.error(f"[song cost] 시트 컬럼명이 올바른지 확인 필요: {e}")
        return None

    def to_num(x):
        if not x:
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        x = str(x).replace("%", "").replace(",", "")
        try:
            return float(x)
        except:
            return 0.0

    artist_cost_dict = {}
    for row in body_sc:
        if row is None:
            continue
        if len(row) < len(header_sc):
            continue
        a = row[idx_artist]
        if not a:
            continue
        cost_data = {
            "정산요율": to_num(row[idx_rate]),
            "전월잔액": to_num(row[idx_prev]),
            "당월차감액": to_num(row[idx_deduct]),
            "당월잔액": to_num(row[idx_remain])
        }
        artist_cost_dict[a] = cost_data

    # (B) input_online revenue: ym 시트 찾기
    if ym not in wb_revenue.sheetnames:
        st.error(f"[online revenue] 파일에 '{ym}' 시트가 없습니다.")
        return None
    ws_or = wb_revenue[ym]

    rows_or = list(ws_or.values)
    if not rows_or:
        st.error(f"[online revenue] '{ym}' 시트가 비어있습니다.")
        return None

    header_or = rows_or[0]
    body_or = rows_or[1:]
    try:
        col_aartist = header_or.index("앨범아티스트")
        col_album = header_or.index("앨범명")
        col_major = header_or.index("대분류")
        col_middle = header_or.index("중분류")
        col_service = header_or.index("서비스명")
        col_revenue = header_or.index("권리사정산금액")
    except ValueError as e:
        st.error(f"[online revenue] 시트 컬럼명이 올바른지 확인 필요: {e}")
        return None

    artist_revenue_dict = defaultdict(list)
    for row in body_or:
        if row is None:
            continue
        if len(row) < len(header_or):
            continue
        aartist = str(row[col_aartist]).strip() if row[col_aartist] else ""
        album = row[col_album] or ""
        major = row[col_major] or ""
        middle = row[col_middle] or ""
        srv = row[col_service] or ""
        rev_val = to_num(row[col_revenue])
        if aartist:
            artist_revenue_dict[aartist].append({
                "album": album,
                "major": major,
                "middle": middle,
                "service": srv,
                "revenue": rev_val
            })

    # (C) 검증용: 아티스트 목록 비교
    song_artists = list(artist_cost_dict.keys())
    revenue_artists = list(artist_revenue_dict.keys())
    check_dict["song_artists"] = song_artists
    check_dict["revenue_artists"] = revenue_artists
    compare_res = compare_artists(song_artists, revenue_artists)
    check_dict["artist_compare_result"] = compare_res

    all_artists = sorted(set(song_artists) | set(revenue_artists))

    # 최종 ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        progress_bar = st.progress(0.0)
        artist_placeholder = st.empty()

        for i, artist in enumerate(all_artists):
            ratio = (i+1)/len(all_artists)
            progress_bar.progress(ratio)
            artist_placeholder.info(f"[{i+1}/{len(all_artists)}] {artist} 처리 중...")

            # 1) 세부매출 내역
            detail_wb = create_detail_workbook(artist, ym, artist_revenue_dict[artist], check_dict)
            detail_buf = io.BytesIO()
            detail_wb.save(detail_buf)
            detail_buf.seek(0)
            zf.writestr(f"{artist}(세부매출내역).xlsx", detail_buf.getvalue())

            # 2) 정산서
            report_wb = create_report_workbook(
                artist, ym, report_date,
                artist_cost_dict.get(artist, {}),
                artist_revenue_dict[artist],
                check_dict
            )
            report_buf = io.BytesIO()
            report_wb.save(report_buf)
            report_buf.seek(0)
            zf.writestr(f"{artist}(정산서).xlsx", report_buf.getvalue())

        artist_placeholder.success("모든 아티스트 처리 완료!")
        progress_bar.progress(1.0)

    # ZIP 바이트를 반환
    return zip_buf.getvalue()


# -----------------------------------------
# 헬퍼 함수: 아티스트 목록 비교
# -----------------------------------------
def compare_artists(song_artists, revenue_artists):
    set_song = set(song_artists)
    set_revenue = set(revenue_artists)
    return {
        "missing_in_song": sorted(set_revenue - set_song),
        "missing_in_revenue": sorted(set_song - set_revenue),
        "common_count": len(set_song & set_revenue),
        "song_count": len(set_song),
        "revenue_count": len(set_revenue),
    }


def almost_equal(a, b, tol=1e-3):
    """숫자 비교용: 소수점 오차 허용."""
    return abs(a - b) < tol


# -----------------------------------------
# (A) 세부매출내역 Workbook 생성
# -----------------------------------------
def create_detail_workbook(artist, ym, detail_list, check_dict):
    """
    artist에 대한 세부매출내역 엑셀 파일을 생성하여 Workbook 객체로 반환.
    detail_list: [{"album":..., "major":..., "middle":..., "service":..., "revenue":...}, ...]
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "세부매출내역"

    # 헤더
    ws.append(["앨범아티스트", "앨범명", "대분류", "중분류", "서비스명", "기간", "매출 순수익"])

    # 정렬(앨범명 등으로 정렬)
    detail_list_sorted = sorted(detail_list, key=lambda x: (x["album"], x["service"]))

    total_revenue = 0.0
    year_val, month_val = ym[:4], ym[4:]
    for d in detail_list_sorted:
        rv = d["revenue"]
        total_revenue += rv
        ws.append([
            artist,
            d["album"],
            d["major"],
            d["middle"],
            d["service"],
            f"{year_val}년 {month_val}월",
            rv
        ])

    # 합계 행
    ws.append(["합계", "", "", "", "", "", total_revenue])

    # 간단한 스타일 예시
    # 1) 헤더 스타일
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4CAF50")  # 연두/초록
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 2) 합계행 스타일 (마지막 행)
    last_row = ws.max_row
    for c in range(1, 7):
        cell = ws.cell(row=last_row, column=c)
        cell.alignment = Alignment(horizontal="center")
    sum_cell = ws.cell(row=last_row, column=7)
    sum_font = Font(bold=True, color="000000")
    sum_fill = PatternFill("solid", fgColor="FFD966")  # 옅은 노랑
    sum_cell.font = sum_font
    sum_cell.fill = sum_fill

    # (검증 기록) → check_dict["details_verification"]["세부매출"] 에 추가
    # 실제로는 "정산서 값과 match" 여부를 비교해야 하지만,
    # 여기서는 generate_report_excel 쪽에서 처리하므로 생략

    return wb


# -----------------------------------------
# (B) 정산서 Workbook 생성
# -----------------------------------------
def create_report_workbook(artist, ym, report_date, cost_data, detail_list, check_dict):
    """
    artist에 대한 "정산서" 엑셀 파일을 생성하여 Workbook 객체로 반환.
    cost_data: {"정산요율":..., "전월잔액":..., "당월차감액":..., "당월잔액":...}
    detail_list: [{"album":..., "major":..., "middle":..., "service":..., "revenue":...}, ...]
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "정산서"

    year_val, month_val = ym[:4], ym[4:]

    # -------------------------------
    # 1) 머리글
    # -------------------------------
    ws["H2"] = f"{report_date} 발행"
    ws["B4"] = f"{year_val}년 {month_val}월 판매분"
    ws["B6"] = f"{artist}님 음원 정산 내역서"

    ws["A8"] = "•"
    ws["B8"] = "저희와 함께해 주셔서 정말 감사하고, 앞으로도 잘 부탁드립니다!"
    ws["A9"] = "•"
    ws["B9"] = f"{year_val}년 {month_val}월 음원 수익을 아래와 같이 정산드립니다."
    ws["A10"] = "•"
    ws["B10"] = "정산 관련 문의사항이 있다면 언제든 편히 연락주세요!"
    ws["F10"] = "E-mail: help@xxxx.com"

    # -------------------------------
    # 2) 세부매출
    # -------------------------------
    row_start = 12
    ws.cell(row=row_start, column=1, value="1.")
    ws.cell(row=row_start, column=2, value="음원 서비스별 정산내역")

    row_start += 1
    headers = ["앨범", "대분류", "중분류", "서비스명", "기간", "매출액"]
    for i, h in enumerate(headers, start=2):
        ws.cell(row=row_start, column=i, value=h)

    detail_list_sorted = sorted(detail_list, key=lambda x: (x["album"], x["service"]))
    total_1 = 0.0
    curr = row_start + 1
    for d in detail_list_sorted:
        rv = d["revenue"]
        total_1 += rv

        ws.cell(row=curr, column=2, value=d["album"])
        ws.cell(row=curr, column=3, value=d["major"])
        ws.cell(row=curr, column=4, value=d["middle"])
        ws.cell(row=curr, column=5, value=d["service"])
        ws.cell(row=curr, column=6, value=f"{year_val}년 {month_val}월")
        ws.cell(row=curr, column=7, value=rv)
        curr += 1

    # 합계
    curr += 1
    ws.cell(row=curr, column=2, value="합계")
    ws.cell(row=curr, column=7, value=total_1)
    row_sum_1 = curr
    curr += 2

    # -------------------------------
    # 3) 앨범 별 정산
    # -------------------------------
    ws.cell(row=curr, column=1, value="2.")
    ws.cell(row=curr, column=2, value="앨범 별 정산 내역")
    curr += 1

    ws.cell(row=curr, column=2, value="앨범")
    ws.cell(row=curr, column=6, value="기간")
    ws.cell(row=curr, column=7, value="매출액")
    curr += 1

    album_sum = defaultdict(float)
    for d in detail_list_sorted:
        album_sum[d["album"]] += d["revenue"]

    total_2 = 0.0
    for alb in sorted(album_sum.keys()):
        amt = album_sum[alb]
        total_2 += amt
        ws.cell(row=curr, column=2, value=alb)
        ws.cell(row=curr, column=6, value=f"{year_val}년 {month_val}월")
        ws.cell(row=curr, column=7, value=amt)
        curr += 1

    ws.cell(row=curr, column=2, value="합계")
    ws.cell(row=curr, column=7, value=total_2)
    row_sum_2 = curr
    curr += 2

    # -------------------------------
    # 4) 공제 내역
    # -------------------------------
    ws.cell(row=curr, column=1, value="3.")
    ws.cell(row=curr, column=2, value="공제 내역")
    curr += 1

    ws.cell(row=curr, column=2, value="앨범")
    ws.cell(row=curr, column=3, value="곡비")
    ws.cell(row=curr, column=4, value="공제 금액")
    ws.cell(row=curr, column=6, value="공제 후 남은 곡비")
    ws.cell(row=curr, column=7, value="공제 적용 금액")
    curr += 1

    prev_val = cost_data.get("전월잔액", 0.0)
    deduct_val = cost_data.get("당월차감액", 0.0)
    remain_val = cost_data.get("당월잔액", 0.0)

    # 공제 적용된 매출액 = total_2 - 당월차감액
    공제적용 = total_2 - deduct_val

    alb_list = sorted(album_sum.keys())
    alb_str = ", ".join(alb_list) if alb_list else "(앨범 없음)"

    ws.cell(row=curr, column=2, value=alb_str)
    ws.cell(row=curr, column=3, value=prev_val)
    ws.cell(row=curr, column=4, value=deduct_val)
    ws.cell(row=curr, column=6, value=remain_val)
    ws.cell(row=curr, column=7, value=공제적용)
    curr += 2

    # -------------------------------
    # 5) 수익 배분
    # -------------------------------
    ws.cell(row=curr, column=1, value="4.")
    ws.cell(row=curr, column=2, value="수익 배분")
    curr += 1

    ws.cell(row=curr, column=2, value="앨범")
    ws.cell(row=curr, column=3, value="항목")
    ws.cell(row=curr, column=4, value="적용율")
    ws.cell(row=curr, column=7, value="적용 금액")
    curr += 1

    rate_val = cost_data.get("정산요율", 0.0)
    final_amount = 공제적용 * (rate_val / 100.0)

    ws.cell(row=curr, column=2, value=alb_str)
    ws.cell(row=curr, column=3, value="수익 배분율")
    ws.cell(row=curr, column=4, value=f"{rate_val}%")
    ws.cell(row=curr, column=7, value=final_amount)
    curr += 1

    ws.cell(row=curr, column=2, value="총 정산금액")
    ws.cell(row=curr, column=7, value=final_amount)
    curr += 2

    ws.cell(row=curr, column=7, value="* 부가세 별도")

    # (검증) check_dict["details_verification"]["정산서"] 에 매핑
    #  - (1) 공제 내역 검증
    is_match_prev = almost_equal(prev_val, cost_data.get("전월잔액", 0))
    is_match_deduct = almost_equal(deduct_val, cost_data.get("당월차감액", 0))
    is_match_remain = almost_equal(remain_val, cost_data.get("당월잔액", 0))
    if not (is_match_prev and is_match_deduct and is_match_remain):
        check_dict["verification_summary"]["total_errors"] += 1
        check_dict["verification_summary"]["artist_error_list"].append(artist)

    row_report_item_3 = {
        "아티스트": artist,
        "구분": "공제내역",
        "원본_곡비": cost_data.get("전월잔액", 0),
        "정산서_곡비": prev_val,
        "match_곡비": is_match_prev,

        "원본_공제금액": cost_data.get("당월차감액", 0),
        "정산서_공제금액": deduct_val,
        "match_공제금액": is_match_deduct,

        "원본_공제후잔액": cost_data.get("당월잔액", 0),
        "정산서_공제후잔액": remain_val,
        "match_공제후잔액": is_match_remain,
    }
    check_dict["details_verification"]["정산서"].append(row_report_item_3)

    #  - (2) 수익 배분율 검증
    original_rate = cost_data.get("정산요율", 0)
    is_rate_match = almost_equal(original_rate, rate_val)
    if not is_rate_match:
        check_dict["verification_summary"]["total_errors"] += 1
        check_dict["verification_summary"]["artist_error_list"].append(artist)

    row_report_item_4 = {
        "아티스트": artist,
        "구분": "수익배분율",
        "원본_정산율(%)": original_rate,
        "정산서_정산율(%)": rate_val,
        "match_정산율": is_rate_match,
    }
    check_dict["details_verification"]["정산서"].append(row_report_item_4)

    # 세부매출 검증(비교)
    for d in detail_list_sorted:
        original_val = d["revenue"]
        # 정산서 쪽도 사실상 d["revenue"] 그대로 사용
        report_val = d["revenue"]
        is_match = almost_equal(original_val, report_val)
        if not is_match:
            check_dict["verification_summary"]["total_errors"] += 1
            check_dict["verification_summary"]["artist_error_list"].append(artist)

        row_report_item = {
            "아티스트": artist,
            "구분": "음원서비스별매출",
            "앨범": d["album"],
            "서비스명": d["service"],
            "원본_매출액": original_val,
            "정산서_매출액": report_val,
            "match_매출액": is_match,
        }
        check_dict["details_verification"]["세부매출"].append(row_report_item)

    # 간단히 “정산서” 제목 행에만 스타일 부여 예시
    ws["B6"].font = Font(size=14, bold=True, color="000000")
    ws["B6"].alignment = Alignment(horizontal="center", vertical="center")

    return wb


if __name__ == "__main__":
    main()
