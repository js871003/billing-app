"""
과금 자동화 로직 v2.1
- 28.02 / 29.02 계약 구분 (2024 / 2025)
- 한결교육 프로모션 지원
- 디렉토리 직접판매 분리 (별도 양식)
- 새 거래명세서 양식: 2024_주N회 / 2025_주N회 / 2025_한결프로모션
"""

import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime


# ========== 기준파일 로딩 ==========

def load_price_criteria(criteria_path):
    """청구요금 기준파일 로드 → 사이트아이디 → 요금 매핑"""
    xls = pd.ExcelFile(criteria_path)
    price_lookup = {}
    directory_sites = set()

    # 1) 직접판매_디렉토리
    for sheet in xls.sheet_names:
        if '직접판매' in sheet and '디렉토리' in sheet:
            df = pd.read_excel(xls, sheet_name=sheet)
            for _, row in df.iterrows():
                site_id = row['사이트아이디']
                directory_sites.add(site_id)
                price_lookup[site_id] = {
                    'price': row['요금'],
                    'price_after_promo': row['요금'],
                    'source': sheet,
                    'contract_year': '디렉토리',
                    '담당지사': _normalize_jisa(row['담당지사']),
                    '요금제': row['요금제'],
                    'promo_end': None,
                }

    # 2) 총판판매_28.02월 종료 (2024 계약분)
    for sheet in xls.sheet_names:
        if '총판판매' in sheet and '28.02' in sheet:
            df = pd.read_excel(xls, sheet_name=sheet)
            price_col = [c for c in df.columns if '요금' in str(c) and c != '요금제'][0]
            for _, row in df.iterrows():
                site_id = row['사이트아이디']
                price_lookup[site_id] = {
                    'price': row[price_col],
                    'price_after_promo': row[price_col],
                    'source': sheet,
                    'contract_year': '2024',
                    '담당지사': _normalize_jisa(row['담당지사']),
                    '요금제': row['요금제'],
                    'promo_end': None,
                }

    # 3) 총판판매_29.02월 종료 (2025 계약분, 프로모션 포함)
    for sheet in xls.sheet_names:
        if '총판판매' in sheet and '29.02' in sheet:
            df = pd.read_excel(xls, sheet_name=sheet)
            price_cols = [c for c in df.columns if '요금' in str(c) and c != '요금제']
            price_cols_sorted = sorted(price_cols)

            for _, row in df.iterrows():
                site_id = row['사이트아이디']
                if site_id in directory_sites:
                    continue
                if len(price_cols_sorted) >= 2:
                    price_promo = row[price_cols_sorted[0]]
                    price_normal = row[price_cols_sorted[1]]
                    promo_end = _extract_date_from_col(price_cols_sorted[0])
                else:
                    price_promo = row[price_cols_sorted[0]]
                    price_normal = price_promo
                    promo_end = None

                price_lookup[site_id] = {
                    'price': price_promo,
                    'price_after_promo': price_normal,
                    'source': sheet,
                    'contract_year': '2025',
                    '담당지사': _normalize_jisa(row['담당지사']),
                    '요금제': row['요금제'],
                    'promo_end': promo_end,
                }

    return price_lookup, directory_sites


def _normalize_jisa(value):
    if isinstance(value, datetime):
        return '오월오일'
    return str(value).strip()


def _extract_date_from_col(col_name):
    import re
    match = re.search(r'~(\d{4})\.(\d{2})', col_name)
    if match:
        return (int(match.group(1)), int(match.group(2)))
    return None


def is_promo_active(promo_end, billing_year, billing_month):
    if promo_end is None:
        return True
    end_year, end_month = promo_end
    if billing_year < end_year:
        return True
    if billing_year == end_year and billing_month <= end_month:
        return True
    return False


# ========== 요금제 매핑 ==========

def classify_plan(plan_str):
    """요일 표기 → 주N회"""
    if plan_str in ['주3회', '주5회']:
        return plan_str
    days = [d.strip() for d in str(plan_str).split(',')]
    return f"주{len(days)}회"


# ========== 거래명세서용 카테고리 분류 ==========

def categorize_for_invoice(row):
    """
    각 행을 거래명세서 카테고리로 분류 (가격 기반).
    
    - 2024 (28.02 계약): 60000=주5회/1년, 54000=주5회/3년, 48000=주3회/1년, 43200=주3회/3년
    - 2025 (29.02 계약): 95000=주5회/3년, 83000=주3회/3년 (모든 2025 약정은 3년)
    - 2025 한결 프로모션: 60000=한결_주5회/3년, 48000=한결_주3회/3년
    """
    contract_year = row.get('contract_year', '')
    price = row.get('요금', 0)
    jisa = row.get('담당지사', '')

    if contract_year == '디렉토리':
        return {'category': '디렉토리', '약정': 'BASIC(3)', 'sort_order': 999}

    # 2024 계약 (28.02 종료)
    if contract_year == '2024':
        if price == 60000:
            return {'category': '2024_주5회', '약정': '1년', 'sort_order': 1}
        elif price == 54000:
            return {'category': '2024_주5회', '약정': '3년', 'sort_order': 2}
        elif price == 48000:
            return {'category': '2024_주3회', '약정': '1년', 'sort_order': 3}
        elif price == 43200:
            return {'category': '2024_주3회', '약정': '3년', 'sort_order': 4}

    # 2025 계약 (29.02 종료) - 약정 모두 3년
    if contract_year == '2025':
        # 한결교육 프로모션 (할인가)
        if jisa == '한결교육' and price == 60000:
            return {'category': '2025_한결_주5회', '약정': '3년', 'sort_order': 7}
        if jisa == '한결교육' and price == 48000:
            return {'category': '2025_한결_주3회', '약정': '3년', 'sort_order': 8}
        # 정상가
        if price == 95000:
            return {'category': '2025_주5회', '약정': '3년', 'sort_order': 5}
        if price == 83000:
            return {'category': '2025_주3회', '약정': '3년', 'sort_order': 6}

    return {'category': '기타', '약정': '', 'sort_order': 999}


# ========== 1단계: 과금 판단 ==========

def process_billing(df):
    """과금 가능 여부 파일 처리 (O/X → 가능/불가능)"""
    df = df.copy()
    df.columns = df.columns.str.strip()
    df['요금제_분류'] = df['요금제'].apply(classify_plan)

    if set(df['과금 가능 여부'].unique()) <= {'O', 'X'}:
        df['과금 가능 여부'] = df['과금 가능 여부'].map({'O': '가능', 'X': '불가능'})

    stats = {
        'total_count': len(df),
        'ok_count': (df['과금 가능 여부'] == '가능').sum(),
        'fail_count': (df['과금 가능 여부'] == '불가능').sum(),
        'review_count': (df['과금 가능 여부'] == '확인필요').sum(),
    }
    review_items = df[df['과금 가능 여부'] == '확인필요'] if stats['review_count'] > 0 else pd.DataFrame()
    return df, stats, review_items


def assign_prices(df, price_lookup, billing_year=2026, billing_month=3):
    """과금 Raw에 요금/계약구분/카테고리 컬럼 추가"""
    df = df.copy()

    def get_info(row):
        if row['과금 가능 여부'] != '가능':
            return {'price': 0, 'jisa': '', 'source': '', 'contract_year': ''}
        info = price_lookup.get(row['사이트아이디'])
        if not info:
            return {'price': 0, 'jisa': '', 'source': '', 'contract_year': ''}
        if is_promo_active(info.get('promo_end'), billing_year, billing_month):
            price = info['price']
        else:
            price = info['price_after_promo']
        return {
            'price': price,
            'jisa': info.get('담당지사', ''),
            'source': info.get('source', ''),
            'contract_year': info.get('contract_year', ''),
        }

    info_series = df.apply(get_info, axis=1)
    df['요금'] = info_series.apply(lambda x: x['price'])
    df['담당지사'] = info_series.apply(lambda x: x['jisa'])
    df['계약구분'] = info_series.apply(lambda x: x['source'])
    df['contract_year'] = info_series.apply(lambda x: x['contract_year'])
    df['비고'] = ''
    return df


# ========== 정합성 체크 ==========

def check_plan_consistency(df, price_lookup):
    mismatches = []
    df_ok = df[df['과금 가능 여부'] == '가능']
    for _, row in df_ok.iterrows():
        site_id = row['사이트아이디']
        billing_plan = row.get('요금제_분류', classify_plan(row['요금제']))
        info = price_lookup.get(site_id)
        if info and billing_plan != info['요금제']:
            mismatches.append({
                'site_id': site_id,
                'institution': row['기관명'],
                'billing_plan': billing_plan,
                'criteria_plan': info['요금제'],
            })
    return mismatches


def check_issue_list_excluded(df, criteria_path):
    xls = pd.ExcelFile(criteria_path)
    issue_sheets = [s for s in xls.sheet_names if '이슈' in s]
    if not issue_sheets:
        return 0, 0, []
    df_issue = pd.read_excel(xls, sheet_name=issue_sheets[0])
    issue_sites = set(df_issue['사이트아이디'])
    billing_sites = set(df['사이트아이디'])
    found = issue_sites & billing_sites
    return len(issue_sites) - len(found), len(issue_sites), list(found)


# ========== 공급자 정보 ==========

SUPPLIER_INFO = {
    'biz_no': '870-88-02332',
    'company': '플레이태그',
    'ceo': '박현수',
    'phone': '02-553-0214',
    'address': '서울 강남구 삼성동 천해빌딩 10층',
    'bank': '하나은행 / 플레이태그주식회사 / 403-910059-30704',
}


def split_by_directory(df, directory_sites):
    """과금 Raw → 총판용 / 디렉토리용 분리"""
    mask = df['사이트아이디'].isin(directory_sites)
    df_directory = df[mask].copy()
    df_wholesale = df[~mask].copy()
    return df_wholesale, df_directory


# ========== 거래명세서 양식 헬퍼 ==========

def _setup_invoice_layout(ws, recipient_name, recipient_info, billing_date,
                          col3_header='요금제', col_d_header='약정', 비고_text=''):
    """공통 레이아웃 (Row 1~10 + Row 33~37)"""
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_area = 'A1:M37'
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    font_title = Font(name='Dotum', size=25)
    font_normal = Font(name='Dotum', size=11)
    font_normal_bold = Font(name='Dotum', size=11, bold=True)
    font_recipient = Font(name='Dotum', size=13, bold=True)

    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    col_widths = {'A': 8.86, 'B': 17.57, 'C': 45.29, 'D': 5.43, 'E': 1.86,
                  'F': 13.0, 'G': 10.71, 'H': 8.86, 'I': 3.14, 'J': 7.29,
                  'K': 2.14, 'L': 5.14, 'M': 16.43}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    ws.row_dimensions[1].height = 12
    ws.row_dimensions[2].height = 38.25
    ws.merge_cells('A2:M2')
    ws['A2'] = '거 래 명 세 서'
    ws['A2'].font = font_title
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    ws.row_dimensions[3].height = 28.5
    ws.row_dimensions[4].height = 21.75
    ws['A4'] = '날짜 :'
    ws['A4'].font = font_normal_bold
    ws.merge_cells('B4:C4')
    ws['B4'] = billing_date if billing_date else datetime.now()
    ws['B4'].font = font_normal_bold
    ws['B4'].number_format = 'YYYY"년" M"월" D"일"'

    for r in range(5, 9):
        ws.row_dimensions[r].height = 30

    ws.merge_cells('D5:D8')
    ws['D5'] = '공\n\n급\n\n자'
    ws['D5'].font = font_normal_bold
    ws['D5'].alignment = center
    ws['D5'].border = thin

    ws.merge_cells('E5:G5')
    ws['E5'] = '사업자등록번호'
    ws['E5'].font = font_normal_bold
    ws['E5'].border = thin
    ws['E5'].alignment = center
    ws.merge_cells('H5:M5')
    ws['H5'] = SUPPLIER_INFO['biz_no']
    ws['H5'].font = font_normal
    ws['H5'].border = thin
    ws['H5'].alignment = center

    ws['A6'] = f' {recipient_name} 귀하'
    ws['A6'].font = font_recipient

    ws.merge_cells('E6:G6')
    ws['E6'] = '상호'
    ws['E6'].font = font_normal_bold
    ws['E6'].border = thin
    ws['E6'].alignment = center
    ws.merge_cells('H6:J6')
    ws['H6'] = SUPPLIER_INFO['company']
    ws['H6'].font = font_normal
    ws['H6'].border = thin
    ws['H6'].alignment = center
    ws.merge_cells('K6:L6')
    ws['K6'] = '대표자'
    ws['K6'].font = font_normal_bold
    ws['K6'].border = thin
    ws['K6'].alignment = center
    ws['M6'] = f' {SUPPLIER_INFO["ceo"]} (인)'
    ws['M6'].font = font_normal
    ws['M6'].border = thin

    addr = recipient_info.get('address', '')
    biz_no = recipient_info.get('biz_no', '')
    email = recipient_info.get('email', '')
    left_info = f'주소 : {addr}'
    if biz_no:
        left_info += f'\n사업자 번호 : {biz_no}'
    if email:
        left_info += f'\n이메일 : {email}'

    ws.merge_cells('A7:C8')
    ws['A7'] = left_info
    ws['A7'].font = font_normal
    ws['A7'].alignment = Alignment(vertical='center', wrap_text=True)

    ws.merge_cells('E7:G7')
    ws['E7'] = '전화번호'
    ws['E7'].font = font_normal_bold
    ws['E7'].border = thin
    ws['E7'].alignment = center
    ws.merge_cells('H7:M7')
    ws['H7'] = SUPPLIER_INFO['phone']
    ws['H7'].font = font_normal
    ws['H7'].border = thin
    ws['H7'].alignment = center

    ws.merge_cells('E8:G8')
    ws['E8'] = '주소'
    ws['E8'].font = font_normal_bold
    ws['E8'].border = thin
    ws['E8'].alignment = center
    ws.merge_cells('H8:M8')
    ws['H8'] = SUPPLIER_INFO['address']
    ws['H8'].font = font_normal
    ws['H8'].border = thin
    ws['H8'].alignment = center

    # Row 9: 합계금액
    ws.row_dimensions[9].height = 28.5
    ws['B9'] = '합계금액 : '
    ws['B9'].font = font_recipient
    ws.merge_cells('C9:G9')
    ws['C9'] = '=K36'
    ws['C9'].font = font_recipient
    ws['C9'].number_format = '#,##0'
    ws['H9'] = '원정'
    ws['H9'].font = font_recipient
    ws.merge_cells('J9:M9')
    ws['J9'] = '=K36'
    ws['J9'].font = font_recipient
    ws['J9'].number_format = '#,##0'

    # Row 10: 헤더
    ws.row_dimensions[10].height = 24
    ws['A10'] = '이용월'
    ws['B10'] = '대리점'
    ws['C10'] = col3_header
    ws['G10'] = '수량'
    ws['H10'] = '요율'

    for c in ['A10', 'B10', 'C10', 'G10', 'H10']:
        ws[c].font = font_normal_bold
        ws[c].border = thin
        ws[c].alignment = center

    ws.merge_cells('D10:F10')
    ws['D10'] = col_d_header
    ws['D10'].font = font_normal_bold
    ws['D10'].border = thin
    ws['D10'].alignment = center

    ws.merge_cells('I10:J10')
    ws['I10'] = '단가'
    ws['I10'].font = font_normal_bold
    ws['I10'].border = thin
    ws['I10'].alignment = center

    ws.merge_cells('K10:M10')
    ws['K10'] = '금 액'
    ws['K10'].font = font_normal_bold
    ws['K10'].border = thin
    ws['K10'].alignment = center

    # Row 33~37: 합계
    ws.row_dimensions[33].height = 6
    for r in [34, 35, 36]:
        ws.row_dimensions[r].height = 22.5

    ws.merge_cells('A34:G36')
    ws['A34'] = f'비 고 : \n{비고_text}' if 비고_text else '비 고 : '
    ws['A34'].font = Font(name='Dotum', size=10, bold=True)
    ws['A34'].alignment = Alignment(vertical='top', wrap_text=True)
    ws['A34'].border = thin

    ws.merge_cells('H34:J34')
    ws['H34'] = '공 급 가'
    ws['H34'].font = font_normal_bold
    ws['H34'].alignment = center
    ws['H34'].border = thin
    ws.merge_cells('K34:M34')
    ws['K34'] = '=SUM(K11:M32)'
    ws['K34'].font = font_normal_bold
    ws['K34'].alignment = Alignment(horizontal='right', vertical='center')
    ws['K34'].border = thin
    ws['K34'].number_format = '#,##0'

    ws.merge_cells('H35:J35')
    ws['H35'] = '부가세'
    ws['H35'].font = font_normal_bold
    ws['H35'].alignment = center
    ws['H35'].border = thin
    ws.merge_cells('K35:M35')
    ws['K35'] = '=K34*10%'
    ws['K35'].font = font_normal_bold
    ws['K35'].alignment = Alignment(horizontal='right', vertical='center')
    ws['K35'].border = thin
    ws['K35'].number_format = '#,##0'

    ws.merge_cells('H36:J36')
    ws['H36'] = '합 계 금 액'
    ws['H36'].font = font_normal_bold
    ws['H36'].alignment = center
    ws['H36'].border = thin
    ws.merge_cells('K36:M36')
    ws['K36'] = '=SUM(K34:M35)'
    ws['K36'].font = font_normal_bold
    ws['K36'].alignment = Alignment(horizontal='right', vertical='center')
    ws['K36'].border = thin
    ws['K36'].number_format = '#,##0'

    ws.row_dimensions[37].height = 34.5
    ws.merge_cells('A37:M37')
    ws['A37'] = f' 입금정보 : {SUPPLIER_INFO["bank"]}'
    ws['A37'].font = font_normal_bold
    ws['A37'].border = thin
    ws['A37'].alignment = Alignment(vertical='center')

    for r in range(5, 9):
        for c in ['D', 'E', 'H', 'K', 'M']:
            ws[f'{c}{r}'].border = thin


def _fill_data_row(ws, row_idx, billing_month, agency, col_c, col_d,
                   qty, rate, unit_price, thin):
    """데이터 행 한 줄 채우기"""
    font_small = Font(name='Dotum', size=10, bold=True)
    font_data = Font(name='Dotum', size=8, bold=True)
    font_num = Font(name='Malgun Gothic', size=10)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    right = Alignment(horizontal='right', vertical='center')

    ws.row_dimensions[row_idx].height = 22.5

    ws[f'A{row_idx}'] = billing_month
    ws[f'A{row_idx}'].font = font_small
    ws[f'A{row_idx}'].alignment = center
    ws[f'A{row_idx}'].border = thin

    ws[f'B{row_idx}'] = agency
    ws[f'B{row_idx}'].font = font_data
    ws[f'B{row_idx}'].alignment = center
    ws[f'B{row_idx}'].border = thin

    ws[f'C{row_idx}'] = col_c
    ws[f'C{row_idx}'].font = font_small
    ws[f'C{row_idx}'].alignment = center
    ws[f'C{row_idx}'].border = thin

    ws.merge_cells(f'D{row_idx}:F{row_idx}')
    ws[f'D{row_idx}'] = col_d
    ws[f'D{row_idx}'].font = font_data
    ws[f'D{row_idx}'].alignment = center
    ws[f'D{row_idx}'].border = thin

    ws[f'G{row_idx}'] = qty
    ws[f'G{row_idx}'].font = font_small
    ws[f'G{row_idx}'].alignment = center
    ws[f'G{row_idx}'].border = thin

    ws[f'H{row_idx}'] = rate
    ws[f'H{row_idx}'].font = font_small
    ws[f'H{row_idx}'].alignment = center
    ws[f'H{row_idx}'].border = thin
    ws[f'H{row_idx}'].number_format = '0%'

    ws.merge_cells(f'I{row_idx}:J{row_idx}')
    ws[f'I{row_idx}'] = unit_price
    ws[f'I{row_idx}'].font = font_num
    ws[f'I{row_idx}'].alignment = center
    ws[f'I{row_idx}'].border = thin
    ws[f'I{row_idx}'].number_format = '#,##0'

    ws.merge_cells(f'K{row_idx}:M{row_idx}')
    ws[f'K{row_idx}'] = f'=I{row_idx}*G{row_idx}*H{row_idx}'
    ws[f'K{row_idx}'].font = font_small
    ws[f'K{row_idx}'].alignment = right
    ws[f'K{row_idx}'].border = thin
    ws[f'K{row_idx}'].number_format = '#,##0'


def _fill_empty_rows(ws, start_row, end_row, thin):
    """빈 데이터 행 (테두리만)"""
    for r in range(start_row, end_row + 1):
        ws.row_dimensions[r].height = 22.5
        for col in ['A', 'B', 'C', 'G', 'H']:
            ws[f'{col}{r}'].border = thin
        ws.merge_cells(f'D{r}:F{r}')
        ws[f'D{r}'].border = thin
        ws.merge_cells(f'I{r}:J{r}')
        ws[f'I{r}'].border = thin
        ws.merge_cells(f'K{r}:M{r}')
        ws[f'K{r}'].border = thin


# ========== 총판 거래명세서 ==========

def create_invoice_excel(df_raw, recipient_name, billing_month, recipient_info=None,
                          billing_date=None):
    """
    총판 거래명세서 생성
    카테고리: 2024_주5회 / 2024_주3회 / 2025_주5회 / 2025_주3회 / 2025_한결프로모션
    """
    if recipient_info is None:
        recipient_info = {'address': '', 'biz_no': '', 'email': ''}

    billing_ok = df_raw[df_raw['과금 가능 여부'] == '가능'].copy()

    cats = billing_ok.apply(categorize_for_invoice, axis=1)
    billing_ok['category'] = cats.apply(lambda x: x['category'])
    billing_ok['약정_label'] = cats.apply(lambda x: x['약정'])
    billing_ok['sort_order'] = cats.apply(lambda x: x['sort_order'])

    grouped = billing_ok.groupby(
        ['sort_order', 'category', '약정_label', '요금']
    ).size().reset_index(name='수량')
    grouped = grouped.sort_values(['sort_order', '요금'], ascending=[True, False])

    wb = Workbook()
    ws = wb.active
    ws.title = '거래명세서'

    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    _setup_invoice_layout(
        ws, recipient_name, recipient_info, billing_date,
        col3_header='요금제', col_d_header='약정',
        비고_text='1) 상세 데이터 별도 제공'
    )

    data_start = 11
    data_end = 32
    row_idx = data_start

    for _, grp in grouped.iterrows():
        if row_idx > data_end:
            break
        _fill_data_row(
            ws, row_idx,
            billing_month=billing_month,
            agency=recipient_name,
            col_c=grp['category'],
            col_d=grp['약정_label'],
            qty=int(grp['수량']),
            rate=1.0,
            unit_price=int(grp['요금']),
            thin=thin
        )
        row_idx += 1

    _fill_empty_rows(ws, row_idx, data_end, thin)
    return wb


# ========== 디렉토리 거래명세서 ==========

def create_directory_invoice_excel(df_directory, recipient_name='(주)디렉토리',
                                    billing_month='', recipient_info=None,
                                    billing_date=None):
    """
    디렉토리 거래명세서 생성 (기관별 그룹핑, 50% 요율)
    각 행: 기관명 / BASIC(3) / 수량 / 50% / 72,000원 / 금액
    """
    if recipient_info is None:
        recipient_info = {'address': '', 'biz_no': '', 'email': ''}

    billing_ok = df_directory[df_directory['과금 가능 여부'] == '가능'].copy()

    grouped = billing_ok.groupby('기관명').agg(
        수량=('사이트아이디', 'count'),
        요금=('요금', 'first'),
    ).reset_index()
    grouped = grouped.sort_values('수량', ascending=False)

    wb = Workbook()
    ws = wb.active
    ws.title = '거래명세서'

    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    _setup_invoice_layout(
        ws, recipient_name, recipient_info, billing_date,
        col3_header='기관명', col_d_header='구분',
        비고_text=''
    )

    data_start = 11
    data_end = 32
    row_idx = data_start

    for _, grp in grouped.iterrows():
        if row_idx > data_end:
            break
        _fill_data_row(
            ws, row_idx,
            billing_month=billing_month,
            agency='디렉토리',
            col_c=grp['기관명'],
            col_d='BASIC(3)',
            qty=int(grp['수량']),
            rate=0.5,
            unit_price=int(grp['요금']),
            thin=thin
        )
        row_idx += 1

    _fill_empty_rows(ws, row_idx, data_end, thin)
    return wb


# ========== 별도 제공자료 ==========

def create_summary_sheet(df_raw):
    """담당지사별 요약 테이블"""
    billing_ok = df_raw[df_raw['과금 가능 여부'] == '가능'].copy()

    pivot = billing_ok.pivot_table(
        values='요금', index='담당지사', columns='요금제_분류',
        aggfunc='sum', fill_value=0
    )
    for col in ['주3회', '주5회']:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot['합계'] = pivot.sum(axis=1)
    pivot = pivot[['주3회', '주5회', '합계']]
    pivot.loc['합계'] = pivot.sum()

    vat_row = pivot.loc['합계'] * 1.1
    vat_row.name = 'VAT 포함'
    pivot = pd.concat([pivot, pd.DataFrame([vat_row])])
    return pivot


def create_detail_excel(df_raw, billing_month):
    """별도 제공자료 엑셀 (요약 + Raw)"""
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = '요약'

    billing_ok = df_raw[df_raw['과금 가능 여부'] == '가능'].copy()

    pivot = billing_ok.pivot_table(
        values='요금', index='담당지사', columns='요금제_분류',
        aggfunc='sum', fill_value=0
    )
    for col in ['주3회', '주5회']:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot['Grand Total'] = pivot.sum(axis=1)
    pivot = pivot[['주3회', '주5회', 'Grand Total']]
    pivot = pivot.sort_index()

    header_font = Font(name='Dotum', size=11, bold=True)
    data_font = Font(name='Dotum', size=10)
    num_fmt = '#,##0'
    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    center = Alignment(horizontal='center', vertical='center')
    right = Alignment(horizontal='right', vertical='center')

    headers = ['담당지사', '주3회', '주5회', 'Grand Total']
    for col_idx, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.border = thin
        cell.alignment = center

    for row_idx, (jisa, row_data) in enumerate(pivot.iterrows(), 2):
        cell = ws_summary.cell(row=row_idx, column=1, value=jisa)
        cell.font = data_font
        cell.border = thin
        for col_idx, col_name in enumerate(['주3회', '주5회', 'Grand Total'], 2):
            val = row_data.get(col_name, 0)
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=val if val != 0 else None)
            cell.font = data_font
            cell.border = thin
            cell.alignment = right
            if val != 0:
                cell.number_format = num_fmt

    ws_summary.column_dimensions['A'].width = 20
    ws_summary.column_dimensions['B'].width = 15
    ws_summary.column_dimensions['C'].width = 15
    ws_summary.column_dimensions['D'].width = 15

    month_code = billing_month.replace('.', '')
    ws_raw = wb.create_sheet(title=f'Raw_{month_code}')

    raw_columns = [
        '사이트아이디', '기관명', '반명', '가능한 일자 수', '성공 일자 수',
        '스토리라인 성공률', '담당지사', '요금제', '요금제_분류',
        '과금 가능 여부', '요금', '계약구분', 'contract_year', '비고'
    ]
    available_cols = [c for c in raw_columns if c in billing_ok.columns]

    for col_idx, h in enumerate(available_cols, 1):
        cell = ws_raw.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.border = thin
        cell.alignment = center

    for row_idx, (_, row) in enumerate(billing_ok.iterrows(), 2):
        for col_idx, col_name in enumerate(available_cols, 1):
            val = row.get(col_name, '')
            if pd.isna(val):
                val = ''
            cell = ws_raw.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin
            if col_name == '스토리라인 성공률' and isinstance(val, (int, float)):
                cell.number_format = '0.0000'
            elif col_name == '요금' and isinstance(val, (int, float)):
                cell.number_format = num_fmt

    col_widths = [18, 15, 12, 12, 12, 15, 15, 12, 10, 12, 12, 18, 10, 10]
    for i, w in enumerate(col_widths[:len(available_cols)]):
        ws_raw.column_dimensions[get_column_letter(i + 1)].width = w

    return wb
