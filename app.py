"""
플레이태그 과금 자동화 웹 앱 v2.1
- 1단계: 과금 Raw 생성
- 2단계: 거래명세서 생성 (총판 + 디렉토리 별도 양식)
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date
from billing_logic import (
    load_price_criteria, process_billing, assign_prices,
    split_by_directory, create_invoice_excel, create_directory_invoice_excel,
    create_summary_sheet, create_detail_excel,
    check_plan_consistency, check_issue_list_excluded,
    classify_plan, SUPPLIER_INFO, excel_to_pdf_bytes
)
from invoice_pdf import create_invoice_pdf, create_directory_invoice_pdf

st.set_page_config(page_title="과금 자동화 v2.3", page_icon="📊", layout="wide")
st.title("📊 과금 자동화 시스템 v2.3")

# ===== 세션 상태 =====
for key, default in [
    ('billing_raw', None),
    ('price_lookup', None),
    ('directory_sites', None),
    ('stats', None),
    ('generated', False),
    ('criteria_path', None),
    ('criteria_file_id', None),
    ('unmatched_sites', None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ===== 사이드바: 기준파일 + 정산 기간 =====
with st.sidebar:
    st.header("⚙️ 설정")

    st.subheader("1. 청구요금 기준파일")
    criteria_file = st.file_uploader(
        "기준파일 업로드 (.xlsx)",
        type=["xlsx"],
        key="criteria_upload",
        help="청구요금_기준파일_XX_XX_XX.xlsx"
    )

    if criteria_file is not None:
        # 파일 식별자 (같은 파일이면 재로드 방지)
        file_id = (criteria_file.name, criteria_file.size)

        if st.session_state.get('criteria_file_id') != file_id:
            # 새 파일 → 로드
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                # CRITICAL: .getvalue() 사용 (.read()는 스트림 위치 문제)
                tmp.write(criteria_file.getvalue())
                tmp_path = tmp.name
            try:
                price_lookup, directory_sites = load_price_criteria(tmp_path)
                if len(price_lookup) == 0:
                    st.error("❌ 기준파일이 비어있습니다. 파일을 확인해주세요.")
                else:
                    st.session_state.price_lookup = price_lookup
                    st.session_state.directory_sites = directory_sites
                    st.session_state.criteria_path = tmp_path
                    st.session_state.criteria_file_id = file_id
                    st.success(
                        f"✅ 기준파일 로드 완료\n\n"
                        f"- 전체 사이트: {len(price_lookup)}건\n"
                        f"- 디렉토리 직접판매: {len(directory_sites)}건"
                    )
            except Exception as e:
                st.error(f"❌ 기준파일 로드 실패: {e}")
        else:
            # 이미 로드된 파일
            st.success(
                f"✅ 기준파일 로드됨\n\n"
                f"- 전체 사이트: {len(st.session_state.price_lookup)}건\n"
                f"- 디렉토리 직접판매: {len(st.session_state.directory_sites)}건"
            )

    st.divider()

    st.subheader("2. 정산 기간")
    col1, col2 = st.columns(2)
    with col1:
        billing_year = st.number_input("연도", value=2026, min_value=2024, max_value=2030)
    with col2:
        billing_month = st.number_input("월", value=3, min_value=1, max_value=12)
    billing_month_str = f"{billing_year % 100:02d}.{billing_month:02d}"
    st.caption(f"이용월: {billing_month_str}")

    st.subheader("3. 거래명세서 발행일")
    invoice_date = st.date_input("발행일", value=date.today())

# ===== 탭 구조 (1단계 / 2단계) =====
tab1, tab2 = st.tabs(["📝 1단계: 과금 Raw 생성", "📄 2단계: 거래명세서 생성"])

# ============================================================
# 1단계: 과금 Raw 생성
# ============================================================
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
        # 스트림 위치 문제 방지를 위해 바이트를 한 번만 읽어둠
        from io import BytesIO
        file_bytes = BytesIO(uploaded_file.getvalue())
        xls = pd.ExcelFile(file_bytes)
        selected_sheet = st.selectbox(
            "처리할 시트 선택 (정산 대상 기간)",
            xls.sheet_names,
            index=len(xls.sheet_names) - 1,
            help="시트명에 정산 기간이 표시됩니다."
        )

        if st.button("🚀 과금 Raw 생성", type="primary"):
            # 재사용을 위해 BytesIO 새로 생성
            df = pd.read_excel(BytesIO(uploaded_file.getvalue()), sheet_name=selected_sheet)
            df.columns = df.columns.str.strip()

            required_cols = ['사이트아이디', '기관명', '스토리라인 성공률',
                             '요금제', '과금 가능 여부']
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                st.error(f"❌ 필수 컬럼 누락: {', '.join(missing)}")
                st.stop()

            result_df, stats, _ = process_billing(df)
            result_df = assign_prices(
                result_df,
                st.session_state.price_lookup,
                billing_year=billing_year,
                billing_month=billing_month
            )

            # 매칭 진단: 과금 가능인데 요금이 0이거나 contract_year가 비어있는 건 체크
            ok_mask = result_df['과금 가능 여부'] == '가능'
            df_ok_all = result_df[ok_mask]
            matched_mask = (df_ok_all['요금'] > 0) & (df_ok_all['contract_year'] != '')
            matched = matched_mask.sum()
            total_ok = ok_mask.sum()
            unmatched = total_ok - matched

            # 미매칭 사이트 상세 정보 저장
            df_unmatched = df_ok_all[~matched_mask].copy()
            st.session_state.unmatched_sites = df_unmatched

            if unmatched == total_ok:
                st.error(
                    "🚨 **모든 항목이 매칭 실패했습니다!**\n\n"
                    "기준파일을 다시 업로드하거나, 사이드바에서 기준파일이 제대로 "
                    "로드되었는지 확인해주세요."
                )
                st.stop()

            st.session_state.billing_raw = result_df
            st.session_state.stats = stats
            st.session_state.generated = True
            st.rerun()

    # ===== 결과 표시 =====
    if st.session_state.generated and st.session_state.billing_raw is not None:
        result_df = st.session_state.billing_raw
        stats = st.session_state.stats

        st.divider()
        st.subheader("📊 결과 요약")

        ok_count = (result_df['과금 가능 여부'] == '가능').sum()
        fail_count = (result_df['과금 가능 여부'] == '불가능').sum()
        review_count = (result_df['과금 가능 여부'] == '확인필요').sum()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("전체", f"{stats['total_count']}건")
        col2.metric("과금 가능", f"{ok_count}건", delta=f"{ok_count/stats['total_count']:.0%}")
        col3.metric("확인필요", f"{review_count}건")
        col4.metric("과금 불가", f"{fail_count}건")

        # ===== 기준파일 미매칭 사이트 리포트 (최우선 표시) =====
        unmatched_df = st.session_state.get('unmatched_sites')
        if unmatched_df is not None and len(unmatched_df) > 0:
            st.divider()
            st.error(
                f"## 🚨 기준파일 업데이트 필요: {len(unmatched_df)}건\n\n"
                f"아래 사이트들은 **과금 가능 리스트에는 있지만 기준파일에 없습니다.** "
                f"이 사이트들의 단가를 알 수 없어 현재 요금 0원으로 처리되고 있습니다.\n\n"
                f"**조치사항:**\n"
                f"1. 담당자에게 기준파일 업데이트 요청 (신규 계약 반영)\n"
                f"2. 기준파일 업데이트 완료 후 사이드바에서 다시 업로드\n"
                f"3. 이 페이지에서 '과금 Raw 생성' 다시 클릭"
            )

            # 미매칭 사이트 테이블
            display_cols = [c for c in
                            ['사이트아이디', '기관명', '반명', '요금제', '스토리라인 성공률']
                            if c in unmatched_df.columns]
            display_df = unmatched_df[display_cols].copy()
            if '스토리라인 성공률' in display_df.columns:
                display_df['스토리라인 성공률'] = display_df['스토리라인 성공률'].apply(
                    lambda x: f"{x:.1%}" if pd.notna(x) else ""
                )

            st.markdown(f"**미매칭 사이트 {len(unmatched_df)}건 (기준파일에 추가 필요):**")
            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True,
                height=min(400, 50 + len(unmatched_df) * 35)
            )

            # 기관별 집계
            if '기관명' in unmatched_df.columns:
                inst_count = unmatched_df.groupby('기관명').size().reset_index(name='클래스 수')
                inst_count = inst_count.sort_values('클래스 수', ascending=False)
                st.markdown(f"**기관별 요약 ({len(inst_count)}개 기관):**")
                st.dataframe(inst_count, use_container_width=True, hide_index=True,
                             height=min(300, 50 + len(inst_count) * 35))

            # 다운로드 버튼
            unmatched_output = BytesIO()
            unmatched_df.to_excel(unmatched_output, index=False)
            unmatched_output.seek(0)
            st.download_button(
                label="📥 미매칭 사이트 목록 다운로드 (기준파일 업데이트용)",
                data=unmatched_output,
                file_name=f"기준파일_업데이트_필요_{billing_month_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        # 🚨 매칭 실패 감지: 가능한데 요금이 0원인 건 확인
        unmatched = result_df[
            (result_df['과금 가능 여부'] == '가능') & (result_df['요금'] == 0)
        ]
        if len(unmatched) > 0:
            st.error(
                f"🚨 **가격 매칭 실패 {len(unmatched)}건!**\n\n"
                f"과금 가능으로 판정되었지만 기준파일에서 가격을 찾지 못했습니다. "
                f"기준파일과 과금 파일의 사이트아이디가 일치하는지 확인해주세요."
            )
            with st.expander(f"미매칭 사이트 목록 ({len(unmatched)}건)", expanded=False):
                st.dataframe(
                    unmatched[['사이트아이디', '기관명', '반명', '요금제']].head(20),
                    use_container_width=True, hide_index=True
                )

        # 총판 / 디렉토리 분리
        df_wholesale, df_directory = split_by_directory(
            result_df, st.session_state.directory_sites
        )
        df_wh_ok = df_wholesale[df_wholesale['과금 가능 여부'] == '가능']
        df_dir_ok = df_directory[df_directory['과금 가능 여부'] == '가능']

        wh_total = df_wh_ok['요금'].sum()
        # 디렉토리는 50% 요율 적용
        dir_total = df_dir_ok['요금'].sum() * 0.5

        st.divider()
        st.subheader("💰 과금 금액")
        col1, col2, col3 = st.columns(3)
        col1.metric(
            "총판 (요율 100%)",
            f"₩{wh_total:,.0f}",
            f"{len(df_wh_ok)}건"
        )
        col2.metric(
            "디렉토리 (요율 50%)",
            f"₩{dir_total:,.0f}",
            f"{len(df_dir_ok)}건"
        )
        col3.metric(
            "전체 합계 (VAT포함)",
            f"₩{(wh_total + dir_total) * 1.1:,.0f}"
        )

        # ===== 정합성 체크 =====
        st.divider()
        st.subheader("🔍 정합성 체크")

        check_col1, check_col2 = st.columns(2)

        with check_col1:
            mismatches = check_plan_consistency(
                result_df, st.session_state.price_lookup
            )
            if mismatches:
                st.warning(f"⚠️ 요금제 불일치 {len(mismatches)}건")
                mismatch_df = pd.DataFrame(mismatches)
                mismatch_df.columns = ['사이트아이디', '기관명', '리스트', '기준파일']
                st.dataframe(mismatch_df, use_container_width=True, hide_index=True)
                st.caption("* 가격은 기준파일 기준으로 부과됩니다.")
            else:
                st.success("✅ 요금제 불일치 없음")

        with check_col2:
            if st.session_state.criteria_path:
                excl, total, found = check_issue_list_excluded(
                    result_df, st.session_state.criteria_path
                )
                if total > 0:
                    if len(found) == 0:
                        st.success(
                            f"✅ 이슈리스트 {total}건 모두 과금 리스트에서 제외됨"
                        )
                    else:
                        st.error(f"🚨 이슈 {len(found)}건이 과금 리스트에 포함")
                        st.write(found)

        # ===== 확인필요 목록 =====
        review_mask = result_df['과금 가능 여부'] == '확인필요'
        review_df = result_df[review_mask]

        if len(review_df) > 0:
            st.divider()
            st.subheader(f"⚠️ 확인필요 목록 ({len(review_df)}건)")

            for idx, row in review_df.iterrows():
                cols = st.columns([3, 2, 1.5, 1.5, 1, 1])
                cols[0].write(f"**{row['기관명']}**")
                cols[1].write(f"{row.get('반명', '')}")
                cols[2].write(f"{row['스토리라인 성공률']:.1%}")
                cols[3].write(f"{row.get('담당지사', '')}")

                if cols[4].button("✅", key=f"ok_{idx}", type="primary"):
                    result_df.loc[idx, '과금 가능 여부'] = '가능'
                    info = st.session_state.price_lookup.get(row['사이트아이디'], {})
                    result_df.loc[idx, '요금'] = info.get('price', 0)
                    st.session_state.billing_raw = result_df
                    st.rerun()

                if cols[5].button("❌", key=f"fail_{idx}"):
                    result_df.loc[idx, '과금 가능 여부'] = '불가능'
                    result_df.loc[idx, '요금'] = 0
                    st.session_state.billing_raw = result_df
                    st.rerun()

        # ===== 단가별 분포 =====
        st.divider()
        st.subheader("📈 단가별 분포")
        price_dist = result_df[result_df['과금 가능 여부'] == '가능'].groupby('요금').agg(
            건수=('요금', 'count'),
            소계=('요금', 'sum')
        ).sort_index(ascending=False)
        st.dataframe(price_dist.style.format({'소계': '{:,.0f}'}), use_container_width=True)

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

        st.success("✅ 2단계 탭에서 거래명세서를 생성하세요!")


# ============================================================
# 2단계: 거래명세서 생성
# ============================================================
with tab2:
    st.subheader("과금 Raw → 거래명세서")

    if st.session_state.price_lookup is None:
        st.warning("⬅️ 사이드바에서 청구요금 기준파일을 먼저 업로드해주세요.")
        st.stop()

    data_source = st.radio(
        "데이터 소스",
        ["1단계에서 생성한 데이터 사용", "과금 Raw 파일 직접 업로드"],
        horizontal=True
    )

    df_for_invoice = None

    if data_source == "1단계에서 생성한 데이터 사용":
        if st.session_state.billing_raw is not None:
            df_for_invoice = st.session_state.billing_raw
            st.success(f"✅ 1단계 데이터 로드 ({len(df_for_invoice)}건)")
        else:
            st.warning("1단계에서 먼저 과금 Raw를 생성해주세요.")
    else:
        raw_upload = st.file_uploader(
            "과금 Raw 파일 (.xlsx)",
            type=["xlsx"],
            key="step2_upload"
        )
        if raw_upload:
            from io import BytesIO
            file_bytes = BytesIO(raw_upload.getvalue())
            xls = pd.ExcelFile(file_bytes)
            raw_sheet = st.selectbox("시트", xls.sheet_names, key="raw_sheet")
            df_for_invoice = pd.read_excel(
                BytesIO(raw_upload.getvalue()), sheet_name=raw_sheet
            )
            st.success(f"✅ 파일 로드 ({len(df_for_invoice)}건)")

    if df_for_invoice is not None:
        df_wholesale, df_directory = split_by_directory(
            df_for_invoice, st.session_state.directory_sites
        )

        st.divider()
        st.subheader("📋 거래명세서 설정")

        invoice_type = st.radio(
            "생성할 거래명세서",
            ["총판 + 디렉토리 (둘 다)", "총판만", "디렉토리만"],
            horizontal=True
        )

        # 기본 이름
        wh_ok = df_wholesale[df_wholesale['과금 가능 여부'] == '가능']
        jisa_count = wh_ok['담당지사'].nunique()
        default_wh_recipient = (
            f"문화사 외 {jisa_count - 1}개 지사"
            if jisa_count > 1 else "문화사"
        )

        wh_recipient = wh_address = wh_biz_no = wh_email = ''
        dir_recipient = dir_address = dir_biz_no = dir_email = ''

        col1, col2 = st.columns(2)

        with col1:
            if "총판" in invoice_type or "둘" in invoice_type:
                st.markdown("**📦 총판 설정**")
                wh_recipient = st.text_input(
                    "총판 수신처", value=default_wh_recipient, key="wh_recipient"
                )
                wh_address = st.text_input(
                    "총판 주소", value="인천광역시 서구 파랑로 495 2동 3층 302호 (청라 에이스)",
                    key="wh_addr"
                )
                wh_biz_no = st.text_input(
                    "총판 사업자번호", value="406-81-66140", key="wh_biz"
                )
                wh_email = st.text_input(
                    "총판 이메일", value="goldengate2021@naver.com", key="wh_email"
                )

        with col2:
            if "디렉토리" in invoice_type or "둘" in invoice_type:
                st.markdown("**📦 디렉토리 설정**")
                dir_recipient = st.text_input(
                    "디렉토리 수신처", value="디렉토리", key="dir_recipient"
                )
                dir_address = st.text_input(
                    "디렉토리 주소", value="부산광역시 부산진구 새싹로 253번길 11",
                    key="dir_addr"
                )
                dir_biz_no = st.text_input(
                    "디렉토리 사업자번호", value="605-86-06399", key="dir_biz"
                )
                dir_email = st.text_input(
                    "디렉토리 이메일", value="hongkr1@dirsys.co.kr", key="dir_email"
                )

        if st.button("📄 거래명세서 생성", type="primary"):
            outputs = {}

            invoice_dt = datetime.combine(invoice_date, datetime.min.time())

            # 총판
            if "총판" in invoice_type or "둘" in invoice_type:
                wh_info = {'address': wh_address, 'biz_no': wh_biz_no, 'email': wh_email}

                # Excel (도장 포함, 미리 계산된 값)
                wb_wh = create_invoice_excel(
                    df_wholesale, recipient_name=wh_recipient,
                    billing_month=billing_month_str,
                    recipient_info=wh_info,
                    billing_date=invoice_dt
                )
                buf = BytesIO()
                wb_wh.save(buf)
                buf.seek(0)
                outputs['총판 거래명세서 (Excel)'] = ('xlsx', buf)

                # PDF: Excel → PDF 변환 시도 (LibreOffice 사용)
                pdf_bytes = excel_to_pdf_bytes(wb_wh)
                if pdf_bytes:
                    outputs['총판 거래명세서 (PDF)'] = ('pdf', BytesIO(pdf_bytes))
                else:
                    # LibreOffice 미설치 시 fallback: reportlab PDF
                    pdf_wh = create_invoice_pdf(
                        df_wholesale, recipient_name=wh_recipient,
                        billing_month=billing_month_str,
                        recipient_info=wh_info
                    )
                    outputs['총판 거래명세서 (PDF)'] = ('pdf', pdf_wh)

                # 별도 제공자료
                detail_wb = create_detail_excel(df_wholesale, billing_month_str)
                detail_buf = BytesIO()
                detail_wb.save(detail_buf)
                detail_buf.seek(0)
                outputs['총판 별도제공자료'] = ('xlsx', detail_buf)

            # 디렉토리
            if "디렉토리" in invoice_type or "둘" in invoice_type:
                dir_info = {'address': dir_address, 'biz_no': dir_biz_no, 'email': dir_email}

                # Excel (도장 포함)
                wb_dir = create_directory_invoice_excel(
                    df_directory, recipient_name=dir_recipient,
                    billing_month=billing_month_str,
                    recipient_info=dir_info,
                    billing_date=invoice_dt
                )
                buf = BytesIO()
                wb_dir.save(buf)
                buf.seek(0)
                outputs['디렉토리 거래명세서 (Excel)'] = ('xlsx', buf)

                # PDF: Excel → PDF 변환 시도
                pdf_bytes = excel_to_pdf_bytes(wb_dir)
                if pdf_bytes:
                    outputs['디렉토리 거래명세서 (PDF)'] = ('pdf', BytesIO(pdf_bytes))
                else:
                    # Fallback
                    pdf_dir = create_directory_invoice_pdf(
                        df_directory, recipient_name=dir_recipient,
                        billing_month=billing_month_str,
                        recipient_info=dir_info
                    )
                    outputs['디렉토리 거래명세서 (PDF)'] = ('pdf', pdf_dir)

            # ===== 결과 표시 =====
            st.divider()
            st.subheader("거래명세서 요약")

            if "총판" in invoice_type or "둘" in invoice_type:
                wh_total = df_wholesale[df_wholesale['과금 가능 여부'] == '가능']['요금'].sum()
                st.markdown("**📦 총판**")
                c1, c2, c3 = st.columns(3)
                c1.metric("공급가", f"₩{wh_total:,.0f}")
                c2.metric("부가세", f"₩{wh_total * 0.1:,.0f}")
                c3.metric("합계", f"₩{wh_total * 1.1:,.0f}")

            if "디렉토리" in invoice_type or "둘" in invoice_type:
                dir_classes = len(df_directory[df_directory['과금 가능 여부'] == '가능'])
                dir_supply = dir_classes * 72000 * 0.5
                st.markdown("**📦 디렉토리** (요율 50%)")
                c1, c2, c3 = st.columns(3)
                c1.metric("공급가", f"₩{dir_supply:,.0f}")
                c2.metric("부가세", f"₩{dir_supply * 0.1:,.0f}")
                c3.metric("합계", f"₩{dir_supply * 1.1:,.0f}")

            # ===== 다운로드 =====
            st.divider()
            month_code = billing_month_str.replace('.', '')

            # 다운로드 버튼들을 행 단위로 배치
            num_outputs = len(outputs)
            cols_per_row = min(3, num_outputs)
            rows_needed = (num_outputs + cols_per_row - 1) // cols_per_row

            output_items = list(outputs.items())
            for row_idx in range(rows_needed):
                cols = st.columns(cols_per_row)
                for col_idx in range(cols_per_row):
                    item_idx = row_idx * cols_per_row + col_idx
                    if item_idx >= num_outputs:
                        break
                    name, (filetype, buf) = output_items[item_idx]

                    if 'PDF' in name:
                        if '총판' in name:
                            fname = f"거래명세서_{month_code}_총판.pdf"
                        else:
                            fname = f"거래명세서_{month_code}_디렉토리.pdf"
                        mime = "application/pdf"
                    elif '별도' in name:
                        fname = f"과금_{month_code}_총판_상세.xlsx"
                        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    elif '총판' in name:
                        fname = f"거래명세서_{month_code}_총판.xlsx"
                        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    else:
                        fname = f"거래명세서_{month_code}_디렉토리.xlsx"
                        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                    cols[col_idx].download_button(
                        label=f"📥 {name}",
                        data=buf,
                        file_name=fname,
                        mime=mime,
                        type="primary" if 'PDF' in name else "secondary",
                        key=f"dl_{item_idx}"
                    )

            st.success("✅ 거래명세서 생성 완료!")
