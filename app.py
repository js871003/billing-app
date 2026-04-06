"""
플레이태그 과금 자동화 웹 앱 v2.0
- 기준파일 기반 가격 매핑
- 28.02 / 29.02 계약 구분
- 한결교육 프로모션 지원
- 디렉토리 직접판매 분리
- 정합성 체크
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from billing_logic import (
    load_price_criteria, process_billing, assign_prices,
    split_by_directory, create_invoice_excel, create_summary_sheet,
    create_detail_excel, check_plan_consistency, check_issue_list_excluded,
    classify_plan, SUPPLIER_INFO
)

st.set_page_config(page_title="과금 자동화 v2", page_icon="📊", layout="wide")
st.title("📊 과금 자동화 시스템 v2.0")

# ===== 세션 상태 관리 =====
if 'billing_raw' not in st.session_state:
    st.session_state.billing_raw = None
if 'price_lookup' not in st.session_state:
    st.session_state.price_lookup = None
if 'directory_sites' not in st.session_state:
    st.session_state.directory_sites = None
if 'stats' not in st.session_state:
    st.session_state.stats = None
if 'generated' not in st.session_state:
    st.session_state.generated = False
if 'criteria_path' not in st.session_state:
    st.session_state.criteria_path = None

# ===== 사이드바: 기준파일 업로드 =====
with st.sidebar:
    st.header("⚙️ 설정")

    st.subheader("1. 청구요금 기준파일")
    criteria_file = st.file_uploader(
        "기준파일 업로드 (.xlsx)",
        type=["xlsx"],
        key="criteria_upload",
        help="청구요금_기준파일_XX_XX_XX.xlsx"
    )

    if criteria_file:
        # 임시 파일로 저장
        import tempfile, os
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(criteria_file.read())
            tmp_path = tmp.name

        try:
            price_lookup, directory_sites = load_price_criteria(tmp_path)
            st.session_state.price_lookup = price_lookup
            st.session_state.directory_sites = directory_sites
            st.session_state.criteria_path = tmp_path
            st.success(f"✅ 기준파일 로드 완료\n- 전체: {len(price_lookup)}건\n- 디렉토리: {len(directory_sites)}건")
        except Exception as e:
            st.error(f"❌ 기준파일 로드 실패: {e}")

    st.divider()

    st.subheader("2. 정산 기간")
    col1, col2 = st.columns(2)
    with col1:
        billing_year = st.number_input("연도", value=2026, min_value=2024, max_value=2030)
    with col2:
        billing_month = st.number_input("월", value=3, min_value=1, max_value=12)

    billing_month_str = f"{billing_year % 100:02d}.{billing_month:02d}"

# ===== 탭 =====
tab1, tab2 = st.tabs(["1단계: 과금 Raw 생성", "2단계: 거래명세서 생성"])

# ===== 1단계: 과금 Raw 생성 =====
with tab1:
    st.subheader("과금 가능 여부 파일 → 과금 Raw")

    if st.session_state.price_lookup is None:
        st.warning("⬅️ 사이드바에서 청구요금 기준파일을 먼저 업로드해주세요.")
        st.stop()

    uploaded_file = st.file_uploader(
        "과금 가능 여부 파일 업로드 (.xlsx)",
        type=["xlsx"],
        help="2026_기관별_과금_가능_여부.xlsx",
        key="step1_upload"
    )

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        selected_sheet = st.selectbox(
            "처리할 시트 선택 (정산 대상 기간)",
            xls.sheet_names,
            index=len(xls.sheet_names) - 1,
        )

        if st.button("🚀 과금 Raw 생성", type="primary"):
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            df.columns = df.columns.str.strip()

            # 필수 컬럼 확인
            required_cols = ['사이트아이디', '기관명', '스토리라인 성공률', '요금제', '과금 가능 여부']
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                st.error(f"❌ 필수 컬럼이 없습니다: {', '.join(missing)}\n\n현재 컬럼: {', '.join(df.columns)}")
                st.stop()

            # 과금 판정
            result_df, stats, review_items = process_billing(df)

            # 요금 할당
            result_df = assign_prices(
                result_df,
                st.session_state.price_lookup,
                billing_year=billing_year,
                billing_month=billing_month
            )

            # 세션 저장
            st.session_state.billing_raw = result_df
            st.session_state.stats = stats
            st.session_state.generated = True

    # ===== 결과 표시 =====
    if st.session_state.generated and st.session_state.billing_raw is not None:
        result_df = st.session_state.billing_raw
        stats = st.session_state.stats

        st.divider()
        st.subheader("결과 요약")

        # 실시간 통계
        ok_count = (result_df['과금 가능 여부'] == '가능').sum()
        fail_count = (result_df['과금 가능 여부'] == '불가능').sum()
        review_count = (result_df['과금 가능 여부'] == '확인필요').sum()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("전체", f"{stats['total_count']}건")
        col2.metric("가능 ✅", f"{ok_count}건")
        col3.metric("확인필요 ⚠️", f"{review_count}건")
        col4.metric("불가능 ❌", f"{fail_count}건")

        # 총판 / 디렉토리 분리
        df_wholesale, df_directory = split_by_directory(
            result_df, st.session_state.directory_sites
        )

        df_wh_ok = df_wholesale[df_wholesale['과금 가능 여부'] == '가능']
        df_dir_ok = df_directory[df_directory['과금 가능 여부'] == '가능']

        wh_total = df_wh_ok['요금'].sum()
        dir_total = df_dir_ok['요금'].sum()
        grand_total = wh_total + dir_total

        st.divider()
        st.subheader("과금 금액")

        col1, col2, col3 = st.columns(3)
        col1.metric("총판 과금액", f"₩{wh_total:,.0f}", f"{len(df_wh_ok)}건")
        col2.metric("디렉토리 과금액", f"₩{dir_total:,.0f}", f"{len(df_dir_ok)}건")
        col3.metric("전체 합계 (VAT별도)", f"₩{grand_total:,.0f}")

        st.metric("전체 합계 (VAT 포함)", f"₩{grand_total * 1.1:,.0f}")

        # ===== 정합성 체크 =====
        st.divider()
        st.subheader("🔍 정합성 체크")

        # 요금제 불일치
        mismatches = check_plan_consistency(result_df, st.session_state.price_lookup)
        if mismatches:
            st.warning(f"⚠️ 요금제 불일치 {len(mismatches)}건")
            mismatch_df = pd.DataFrame(mismatches)
            mismatch_df.columns = ['사이트아이디', '기관명', '과금리스트', '기준파일']
            st.dataframe(mismatch_df, use_container_width=True, hide_index=True)
        else:
            st.success("✅ 요금제 불일치 없음")

        # 이슈리스트 확인
        if st.session_state.criteria_path:
            excl, total, found = check_issue_list_excluded(
                result_df, st.session_state.criteria_path
            )
            if total > 0:
                if len(found) == 0:
                    st.success(f"✅ 이슈리스트 {total}건 모두 과금 리스트에서 제외됨")
                else:
                    st.error(f"🚨 이슈리스트 중 {len(found)}건이 과금 리스트에 포함되어 있습니다!")
                    st.write(found)

        # ===== 확인필요 목록 =====
        review_mask = result_df['과금 가능 여부'] == '확인필요'
        review_df = result_df[review_mask]

        if len(review_df) > 0:
            st.divider()
            st.subheader(f"⚠️ 확인필요 목록 ({len(review_df)}건)")
            st.caption("각 항목을 확인하고 가능/불가능을 선택해주세요.")

            for idx, row in review_df.iterrows():
                with st.container():
                    cols = st.columns([3, 2, 1.5, 1.5, 1, 1])
                    cols[0].write(f"**{row['기관명']}**")
                    cols[1].write(f"{row.get('반명', '')}")
                    cols[2].write(f"성공률: {row['스토리라인 성공률']:.1%}")
                    cols[3].write(f"{row.get('담당지사', '')}")

                    if cols[4].button("✅ 가능", key=f"ok_{idx}", type="primary"):
                        result_df.loc[idx, '과금 가능 여부'] = '가능'
                        # 요금 재할당
                        info = st.session_state.price_lookup.get(row['사이트아이디'], {})
                        result_df.loc[idx, '요금'] = info.get('price', 0)
                        st.session_state.billing_raw = result_df
                        st.rerun()

                    if cols[5].button("❌ 불가", key=f"fail_{idx}"):
                        result_df.loc[idx, '과금 가능 여부'] = '불가능'
                        result_df.loc[idx, '요금'] = 0
                        st.session_state.billing_raw = result_df
                        st.rerun()
        else:
            st.success("✅ 확인필요 항목이 없습니다.")

        # ===== 요약 테이블 =====
        st.divider()
        st.subheader("담당지사별 요약 (총판)")

        if len(df_wh_ok) > 0:
            summary = create_summary_sheet(df_wholesale)
            st.dataframe(
                summary.style.format('{:,.0f}'),
                use_container_width=True
            )

        # ===== 단가별 분포 =====
        st.divider()
        st.subheader("단가별 분포")

        price_dist = result_df[result_df['과금 가능 여부'] == '가능'].groupby('요금').agg(
            건수=('요금', 'count'),
            소계=('요금', 'sum')
        ).sort_index(ascending=False)
        st.dataframe(
            price_dist.style.format({'소계': '{:,.0f}'}),
            use_container_width=True
        )

        # ===== 전체 결과 =====
        st.divider()
        st.subheader("전체 결과")

        filter_col1, filter_col2 = st.columns(2)
        with filter_col1:
            status_filter = st.multiselect(
                "과금 가능 여부",
                ['가능', '확인필요', '불가능'],
                default=['가능', '확인필요', '불가능']
            )

        display_df = result_df[result_df['과금 가능 여부'].isin(status_filter)]
        display_cols = [c for c in ['사이트아이디', '기관명', '반명', '스토리라인 성공률',
                                     '담당지사', '요금제_분류', '과금 가능 여부', '요금', '계약구분']
                        if c in display_df.columns]

        st.dataframe(
            display_df[display_cols].style.format({
                '스토리라인 성공률': '{:.1%}',
                '요금': '{:,.0f}'
            }),
            use_container_width=True, hide_index=True, height=400
        )

        # ===== 다운로드 =====
        st.divider()

        output = BytesIO()
        result_df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="📥 과금 Raw 다운로드 (.xlsx)",
            data=output,
            file_name=f"과금_Raw_{billing_month_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        st.success("✅ 2단계 탭에서 거래명세서를 생성할 수 있습니다!")


# ===== 2단계: 거래명세서 생성 =====
with tab2:
    st.subheader("과금 Raw → 거래명세서 / 별도 제공자료")

    if st.session_state.price_lookup is None:
        st.warning("⬅️ 사이드바에서 청구요금 기준파일을 먼저 업로드해주세요.")
        st.stop()

    # 데이터 소스
    data_source = st.radio(
        "데이터 소스",
        ["1단계에서 생성한 데이터 사용", "과금 Raw 파일 직접 업로드"],
        horizontal=True
    )

    df_for_invoice = None

    if data_source == "1단계에서 생성한 데이터 사용":
        if st.session_state.billing_raw is not None:
            df_for_invoice = st.session_state.billing_raw
            st.success(f"✅ 1단계 데이터 로드 완료 ({len(df_for_invoice)}건)")
        else:
            st.warning("1단계에서 먼저 과금 Raw를 생성해주세요.")
    else:
        raw_upload = st.file_uploader(
            "과금 Raw 파일 업로드 (.xlsx)",
            type=["xlsx"],
            key="step2_upload"
        )
        if raw_upload:
            xls = pd.ExcelFile(raw_upload)
            raw_sheet = st.selectbox("시트 선택", xls.sheet_names, key="raw_sheet")
            df_for_invoice = pd.read_excel(raw_upload, sheet_name=raw_sheet)
            st.success(f"✅ 파일 로드 완료 ({len(df_for_invoice)}건)")

    if df_for_invoice is not None:
        # 총판 / 디렉토리 분리
        df_wholesale, df_directory = split_by_directory(
            df_for_invoice, st.session_state.directory_sites
        )

        st.divider()
        st.subheader("📋 거래명세서 설정")

        invoice_type = st.radio(
            "생성 대상",
            ["총판 거래명세서", "디렉토리 거래명세서", "둘 다"],
            horizontal=True
        )

        col1, col2 = st.columns(2)

        with col1:
            # 총판용
            if invoice_type in ["총판 거래명세서", "둘 다"]:
                st.markdown("**총판 설정**")
                wh_ok = df_wholesale[df_wholesale['과금 가능 여부'] == '가능']
                jisa_count = wh_ok['담당지사'].nunique()
                default_recipient = f"문화사 외 {jisa_count - 1}개 지사" if jisa_count > 1 else "문화사"
                wh_recipient = st.text_input("총판 수신처", value=default_recipient)
                wh_address = st.text_input("총판 주소", value="")
                wh_biz_no = st.text_input("총판 사업자번호", value="")

        with col2:
            # 디렉토리용
            if invoice_type in ["디렉토리 거래명세서", "둘 다"]:
                st.markdown("**디렉토리 설정**")
                dir_recipient = st.text_input("디렉토리 수신처", value="(주)디렉토리")
                dir_address = st.text_input("디렉토리 주소", value="")
                dir_biz_no = st.text_input("디렉토리 사업자번호", value="")

        if st.button("📄 거래명세서 생성", type="primary"):
            outputs = []

            # === 총판 거래명세서 ===
            if invoice_type in ["총판 거래명세서", "둘 다"]:
                wh_info = {'address': wh_address, 'biz_no': wh_biz_no, 'email': ''}

                wb_wh = create_invoice_excel(
                    df_wholesale, recipient_name=wh_recipient,
                    billing_month=billing_month_str, recipient_info=wh_info
                )
                wh_excel = BytesIO()
                wb_wh.save(wh_excel)
                wh_excel.seek(0)
                outputs.append(('wholesale_invoice', wh_excel))

                # 별도 제공자료
                detail_wb = create_detail_excel(df_wholesale, billing_month_str)
                detail_output = BytesIO()
                detail_wb.save(detail_output)
                detail_output.seek(0)
                outputs.append(('wholesale_detail', detail_output))

            # === 디렉토리 거래명세서 ===
            if invoice_type in ["디렉토리 거래명세서", "둘 다"]:
                dir_info = {'address': dir_address, 'biz_no': dir_biz_no, 'email': ''}

                wb_dir = create_invoice_excel(
                    df_directory, recipient_name=dir_recipient,
                    billing_month=billing_month_str, recipient_info=dir_info
                )
                dir_excel = BytesIO()
                wb_dir.save(dir_excel)
                dir_excel.seek(0)
                outputs.append(('directory_invoice', dir_excel))

            # === 결과 표시 ===
            st.divider()
            st.subheader("거래명세서 요약")

            if invoice_type in ["총판 거래명세서", "둘 다"]:
                wh_ok = df_wholesale[df_wholesale['과금 가능 여부'] == '가능']
                wh_total = wh_ok['요금'].sum()
                st.markdown("**총판**")
                c1, c2, c3 = st.columns(3)
                c1.metric("공급가", f"₩{wh_total:,.0f}")
                c2.metric("부가세", f"₩{wh_total * 0.1:,.0f}")
                c3.metric("합계", f"₩{wh_total * 1.1:,.0f}")

            if invoice_type in ["디렉토리 거래명세서", "둘 다"]:
                dir_ok = df_directory[df_directory['과금 가능 여부'] == '가능']
                dir_total = dir_ok['요금'].sum()
                st.markdown("**디렉토리**")
                c1, c2, c3 = st.columns(3)
                c1.metric("공급가", f"₩{dir_total:,.0f}")
                c2.metric("부가세", f"₩{dir_total * 0.1:,.0f}")
                c3.metric("합계", f"₩{dir_total * 1.1:,.0f}")

            # 다운로드 버튼
            st.divider()
            month_code = billing_month_str.replace('.', '')

            dl_cols = st.columns(3)
            col_idx = 0

            for name, data in outputs:
                if name == 'wholesale_invoice':
                    dl_cols[col_idx].download_button(
                        label="📥 총판 거래명세서",
                        data=data,
                        file_name=f"거래명세서_{month_code}_총판.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    col_idx += 1
                elif name == 'wholesale_detail':
                    dl_cols[col_idx].download_button(
                        label="📥 총판 별도제공자료",
                        data=data,
                        file_name=f"과금_{month_code}_총판_상세.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    col_idx += 1
                elif name == 'directory_invoice':
                    dl_cols[col_idx].download_button(
                        label="📥 디렉토리 거래명세서",
                        data=data,
                        file_name=f"거래명세서_{month_code}_디렉토리.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    col_idx += 1

            st.success("✅ 거래명세서 생성 완료!")
