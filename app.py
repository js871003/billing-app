"""
플레이태그 과금 자동화 웹 앱
1단계: 과금 Raw 생성
2단계: 거래명세서 PDF 생성
"""
import streamlit as st
import pandas as pd
from io import BytesIO
from billing_logic import (
    process_billing, assign_prices, create_invoice_excel,
    create_summary_sheet, PRICE_MAP, DEFAULT_1YEAR_SITES
)
from invoice_pdf import create_invoice_pdf

st.set_page_config(page_title="과금 자동화", page_icon="📊", layout="wide")

st.title("📊 과금 자동화 시스템")

# 탭으로 구분
tab1, tab2 = st.tabs(["1단계: 과금 Raw 생성", "2단계: 거래명세서 생성"])

# ===== 세션 상태 관리 =====
if 'billing_raw' not in st.session_state:
    st.session_state.billing_raw = None
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = None


# ===== 1단계: 과금 Raw 생성 =====
with tab1:
    st.subheader("기관별 과금 가능 여부 → 과금 Raw")

    uploaded_file = st.file_uploader(
        "엑셀 파일 업로드 (.xlsx)",
        type=["xlsx"],
        help="기관별 과금 가능 여부 파일을 올려주세요",
        key="step1_upload"
    )

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        selected_sheet = st.selectbox(
            "처리할 시트 선택",
            xls.sheet_names,
            index=len(xls.sheet_names) - 1,
        )

        # 1년 약정 기관 관리
        with st.expander("⚙️ 1년 약정 기관 설정"):
            st.caption("1년 약정 기관의 사이트아이디를 입력하세요 (한 줄에 하나씩)")
            one_year_text = st.text_area(
                "사이트아이디 목록",
                value='\n'.join(DEFAULT_1YEAR_SITES),
                height=100,
                help="예: koreauniv_316"
            )
            one_year_sites = [s.strip() for s in one_year_text.split('\n') if s.strip()]

        if st.button("🚀 과금 Raw 생성", type="primary"):
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            result_df, stats, review_items = process_billing(df)

            # 요금 할당
            result_df = assign_prices(result_df, one_year_sites)

            # 세션에 저장 (2단계에서 사용)
            st.session_state.billing_raw = result_df
            st.session_state.selected_sheet = selected_sheet

            # 결과 요약
            st.divider()
            st.subheader("결과 요약")

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("원본", f"{stats['original_count']}건")
            col2.metric("가능", f"{stats['ok_count']}건")
            col3.metric("확인필요", f"{stats['review_count']}건")
            col4.metric("불가능", f"{stats['fail_count']}건")

            total_amount = result_df[result_df['과금 가능 여부'] == '가능']['요금'].sum()
            st.metric("총 과금액 (VAT 별도)", f"₩{total_amount:,.0f}")

            st.caption(
                f"담당지사 제외: {stats['excluded_jisa_count']}건 / "
                f"디렉토리 별도계약 제외: {stats['excluded_dir_count']}건"
            )

            # 확인필요 목록
            if len(review_items) > 0:
                st.divider()
                st.subheader("⚠️ 확인필요 목록 (40~60% 경계 구간)")
                st.dataframe(
                    review_items.style.format({'스토리라인 성공률': '{:.1%}'}),
                    use_container_width=True, hide_index=True
                )

            # 요약 테이블
            st.divider()
            st.subheader("담당지사별 요약")
            summary = create_summary_sheet(result_df)
            st.dataframe(
                summary.style.format('{:,.0f}'),
                use_container_width=True
            )

            # 전체 결과
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
            st.dataframe(
                display_df.style.format({
                    '스토리라인 성공률': '{:.1%}',
                    '요금': '{:,.0f}'
                }),
                use_container_width=True, hide_index=True, height=400
            )

            # 다운로드
            st.divider()
            output = BytesIO()
            result_df.to_excel(output, index=False, sheet_name=selected_sheet)
            output.seek(0)

            st.download_button(
                label="📥 과금 Raw 다운로드 (.xlsx)",
                data=output,
                file_name=f"과금_Raw_{selected_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            st.success("✅ 2단계 탭에서 거래명세서를 생성할 수 있습니다!")


# ===== 2단계: 거래명세서 생성 =====
with tab2:
    st.subheader("과금 Raw → 거래명세서 PDF / Excel")

    # 데이터 소스 선택
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
        st.divider()

        # 거래명세서 설정
        col1, col2 = st.columns(2)
        with col1:
            # 담당지사 수 계산
            jisa_count = df_for_invoice[df_for_invoice['과금 가능 여부'] == '가능']['담당지사'].nunique()
            default_recipient = f"문화사 외 {jisa_count - 1}개 지사" if jisa_count > 1 else "문화사"

            recipient_name = st.text_input("수신처 (대리점명)", value=default_recipient)
            billing_month = st.text_input("이용월", value="26.01", help="예: 26.01")

        with col2:
            recipient_address = st.text_input("수신처 주소", value="")
            recipient_biz_no = st.text_input("수신처 사업자번호", value="")
            recipient_email = st.text_input("수신처 이메일", value="")

        if st.button("📄 거래명세서 생성", type="primary"):
            recipient_info = {
                'address': recipient_address,
                'biz_no': recipient_biz_no,
                'email': recipient_email,
            }

            # === Excel 생성 ===
            wb = create_invoice_excel(
                df_for_invoice,
                recipient_name=recipient_name,
                billing_month=billing_month,
                recipient_info=recipient_info
            )
            excel_output = BytesIO()
            wb.save(excel_output)
            excel_output.seek(0)

            # === PDF 생성 ===
            pdf_output = create_invoice_pdf(
                df_for_invoice,
                recipient_name=recipient_name,
                billing_month=billing_month,
                recipient_info=recipient_info
            )

            st.divider()

            # 미리보기: 과금 요약
            billing_ok = df_for_invoice[df_for_invoice['과금 가능 여부'] == '가능']
            total = billing_ok['요금'].sum()

            st.subheader("거래명세서 요약")
            col1, col2, col3 = st.columns(3)
            col1.metric("공급가", f"₩{total:,.0f}")
            col2.metric("부가세", f"₩{total * 0.1:,.0f}")
            col3.metric("합계금액", f"₩{total * 1.1:,.0f}")

            # 항목별 내역
            price_groups = billing_ok.groupby('요금').size().reset_index(name='수량')
            price_groups = price_groups[price_groups['요금'] > 0].sort_values('요금', ascending=False)
            price_groups['금액'] = price_groups['요금'] * price_groups['수량']
            price_groups.columns = ['단가', '수량', '금액']
            st.dataframe(
                price_groups.style.format({'단가': '{:,.0f}', '금액': '{:,.0f}'}),
                use_container_width=True, hide_index=True
            )

            # 다운로드 버튼
            st.divider()
            col1, col2 = st.columns(2)

            with col1:
                st.download_button(
                    label="📥 거래명세서 다운로드 (PDF)",
                    data=pdf_output,
                    file_name=f"거래명세서_{billing_month}.pdf",
                    mime="application/pdf",
                    type="primary"
                )

            with col2:
                st.download_button(
                    label="📥 거래명세서 다운로드 (Excel)",
                    data=excel_output,
                    file_name=f"거래명세서_{billing_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.success("✅ 거래명세서 생성 완료!")
