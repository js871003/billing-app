"""
과금 Raw 자동 생성 + 거래명세서 생성 로직
"""
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter
from datetime import datetime
import copy

# ========== 1단계: 과금 판단 설정 ==========
EXCLUDED_JISA = ['플레이태그', '?']

EXCLUDED_DIRECTORY_INSTITUTIONS = [
    '수영남천더샵', '동래행복주택', '동래효성해링턴',
    '동래꿈에그린', '동래명륜자이', '동래안락현대',
]

THRESHOLD_AUTO_OK = 0.85
THRESHOLD_REVIEW_HIGH = 0.60
THRESHOLD_REVIEW_LOW = 0.40

# ========== 2단계: 요금 설정 ==========
# 요금제 + 약정 → 단가
PRICE_MAP = {
    ('주5회', '1년'): 60000,
    ('주5회', '3년'): 54000,
    ('주3회', '1년'): 48000,
    ('주3회', '3년'): 43200,
}

# 1년 약정 기관 사이트아이디 목록 (기본값, 웹에서 수정 가능)
DEFAULT_1YEAR_SITES = []

# ========== 공급자 정보 (고정) ==========
SUPPLIER_INFO = {
    'biz_no': '870-88-02332',
    'company': '플레이태그',
    'ceo': '박현수',
    'phone': '02-553-0214',
    'address': '서울 강남구 강남대로140길 9, 5층',
    'bank': '하나은행 / 플레이태그주식회사 / 403-910059-30704',
}


def process_billing(df):
    """1단계: 과금 Raw 생성"""
    # 컬럼명 공백 제거
    df.columns = df.columns.str.strip()
    original_count = len(df)

    df = df[~df['담당지사'].isin(EXCLUDED_JISA)].copy()
    excluded_jisa_count = original_count - len(df)

    before = len(df)
    mask = (df['담당지사'].str.contains('디렉토리', na=False)) & \
           (df['기관명'].isin(EXCLUDED_DIRECTORY_INSTITUTIONS))
    df = df[~mask].copy()
    excluded_dir_count = before - len(df)

    def determine_billing(row):
        # 서비스 종료 → 자동 불가능
        if row['담당지사'] == '서비스 종료':
            return '불가능(서비스종료)'
        if row['리포트 전송 여부'] == '미전송':
            return '불가능'
        if row['스토리라인 성공률'] >= THRESHOLD_AUTO_OK:
            return '가능'
        if row['스토리라인 성공률'] >= THRESHOLD_REVIEW_HIGH:
            return '가능'
        if row['스토리라인 성공률'] >= THRESHOLD_REVIEW_LOW:
            return '확인필요'
        return '불가능'

    df['과금 가능 여부'] = df.apply(determine_billing, axis=1)

    stats = {
        'original_count': original_count,
        'excluded_jisa_count': excluded_jisa_count,
        'excluded_dir_count': excluded_dir_count,
        'final_count': len(df),
        'ok_count': (df['과금 가능 여부'] == '가능').sum(),
        'review_count': (df['과금 가능 여부'] == '확인필요').sum(),
        'fail_count': (df['과금 가능 여부'].isin(['불가능', '불가능(서비스종료)'])).sum(),
        'service_end_count': (df['과금 가능 여부'] == '불가능(서비스종료)').sum(),
    }

    review_items = df[df['과금 가능 여부'] == '확인필요'][
        ['기관명', '반명', '스토리라인 성공률', '담당지사']
    ] if stats['review_count'] > 0 else pd.DataFrame()

    return df, stats, review_items


def assign_prices(df, one_year_sites=None):
    """과금 Raw에 요금 컬럼 추가"""
    if one_year_sites is None:
        one_year_sites = DEFAULT_1YEAR_SITES

    one_year_set = set(one_year_sites)

    def get_price(row):
        if row['과금 가능 여부'] != '가능':
            return 0
        contract = '1년' if row['사이트아이디'] in one_year_set else '3년'
        plan = row['요금제']
        return PRICE_MAP.get((plan, contract), 0)

    def get_service(row):
        if row['요금제'] == '주5회':
            return '주 5일'
        elif row['요금제'] == '주3회':
            return '주3일'
        return ''

    df = df.copy()
    df['요금'] = df.apply(get_price, axis=1)
    df['서비스'] = df.apply(get_service, axis=1)
    df['비고'] = ''
    return df


def create_invoice_excel(df_raw, recipient_name, billing_month, recipient_info=None):
    """
    거래명세서 엑셀 생성 (기존 양식 그대로 재현)

    Parameters:
        df_raw: 과금 Raw DataFrame (요금 포함)
        recipient_name: 수신처 이름 (예: "문화사 외 16개 지사")
        billing_month: 이용월 (예: "26.01")
        recipient_info: 수신처 정보 dict (address, biz_no, email)
    """
    if recipient_info is None:
        recipient_info = {'address': '', 'biz_no': '', 'email': ''}

    # 과금 가능한 건만
    billing_ok = df_raw[df_raw['과금 가능 여부'] == '가능']

    # 단가별 그룹핑
    price_groups = billing_ok.groupby('요금').size().reset_index(name='수량')
    price_groups = price_groups[price_groups['요금'] > 0].sort_values('요금', ascending=False)

    # 요금제/약정 매핑
    price_to_plan = {
        60000: ('Standard', '1년'),
        54000: ('Standard', '3년'),
        48000: ('Basic', '1년'),
        43200: ('Basic', '3년'),
    }

    wb = Workbook()
    ws = wb.active
    ws.title = '거래명세서'

    # 인쇄 설정 (A4 1페이지)
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

    # 스타일
    font_title = Font(name='Dotum', size=25)
    font_normal = Font(name='Dotum', size=11)
    font_normal_bold = Font(name='Dotum', size=11, bold=True)
    font_small = Font(name='Dotum', size=10, bold=True)
    font_small_plain = Font(name='Malgun Gothic', size=10)
    font_recipient = Font(name='Dotum', size=13, bold=True)
    font_data = Font(name='Dotum', size=8, bold=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    right_align = Alignment(horizontal='right', vertical='center')

    # 열 너비
    col_widths = {'A': 8.86, 'B': 17.57, 'C': 45.29, 'D': 5.43, 'E': 1.86,
                  'F': 13.0, 'G': 10.71, 'H': 8.86, 'I': 3.14, 'J': 7.29,
                  'K': 2.14, 'L': 5.14, 'M': 16.43}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Row 1: 빈 행
    ws.row_dimensions[1].height = 12
    ws['A1'] = ' '

    # Row 2: 제목
    ws.row_dimensions[2].height = 38.25
    ws.merge_cells('A2:M2')
    ws['A2'] = '거 래 명 세 서'
    ws['A2'].font = font_title
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    # Row 3: 빈 행
    ws.row_dimensions[3].height = 28.5

    # Row 4: 날짜
    ws.row_dimensions[4].height = 21.75
    ws['A4'] = '날짜 :'
    ws['A4'].font = font_normal_bold
    ws.merge_cells('B4:C4')
    now = datetime.now()
    # 이용월의 마지막 날
    ws['B4'] = now
    ws['B4'].font = font_normal_bold
    ws['B4'].number_format = 'YYYY-MM-DD'

    # Row 5~8: 공급자/수신자 정보
    for r in range(5, 9):
        ws.row_dimensions[r].height = 30

    # 공급자 섹션
    ws.merge_cells('D5:D8')
    ws['D5'] = '공\n\n급\n\n자'
    ws['D5'].font = font_normal_bold
    ws['D5'].alignment = center
    ws['D5'].border = thin_border

    ws.merge_cells('E5:G5')
    ws['E5'] = '사업자등록번호'
    ws['E5'].font = font_normal_bold
    ws['E5'].border = thin_border
    ws['E5'].alignment = center

    ws.merge_cells('H5:M5')
    ws['H5'] = SUPPLIER_INFO['biz_no']
    ws['H5'].font = font_normal
    ws['H5'].border = thin_border
    ws['H5'].alignment = center

    # 수신자
    ws['A6'] = f'    {recipient_name} 귀하'
    ws['A6'].font = font_recipient

    ws.merge_cells('E6:G6')
    ws['E6'] = '상호'
    ws['E6'].font = font_normal_bold
    ws['E6'].border = thin_border
    ws['E6'].alignment = center

    ws.merge_cells('H6:J6')
    ws['H6'] = SUPPLIER_INFO['company']
    ws['H6'].font = font_normal
    ws['H6'].border = thin_border
    ws['H6'].alignment = center

    ws.merge_cells('K6:L6')
    ws['K6'] = '대표자'
    ws['K6'].font = font_normal_bold
    ws['K6'].border = thin_border
    ws['K6'].alignment = center

    ws['M6'] = f'   {SUPPLIER_INFO["ceo"]}    (인)'
    ws['M6'].font = font_normal
    ws['M6'].border = thin_border

    # 수신자 주소/사업자번호
    addr_text = recipient_info.get('address', '')
    biz_no = recipient_info.get('biz_no', '')
    email = recipient_info.get('email', '')
    left_info = f'주소 : {addr_text}'
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
    ws['E7'].border = thin_border
    ws['E7'].alignment = center

    ws.merge_cells('H7:M7')
    ws['H7'] = SUPPLIER_INFO['phone']
    ws['H7'].font = font_normal
    ws['H7'].border = thin_border
    ws['H7'].alignment = center

    ws.merge_cells('E8:G8')
    ws['E8'] = '주소'
    ws['E8'].font = font_normal_bold
    ws['E8'].border = thin_border
    ws['E8'].alignment = center

    ws.merge_cells('H8:M8')
    ws['H8'] = SUPPLIER_INFO['address']
    ws['H8'].font = font_normal
    ws['H8'].border = thin_border
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
    headers = {'A10': '이용월', 'B10': '대리점', 'C10': '요금제',
               'G10': '수량', 'H10': '요율', 'K10': '금    액'}
    ws.merge_cells('D10:F10')
    ws['D10'] = '약정'
    ws['D10'].font = font_normal_bold
    ws['D10'].border = thin_border
    ws['D10'].alignment = center

    ws.merge_cells('I10:J10')
    ws['I10'] = '단가'
    ws['I10'].font = font_normal_bold
    ws['I10'].border = thin_border
    ws['I10'].alignment = center

    ws.merge_cells('K10:M10')
    ws['K10'] = '금    액'
    ws['K10'].font = font_normal_bold
    ws['K10'].border = thin_border
    ws['K10'].alignment = center

    for cell_ref, value in headers.items():
        if cell_ref not in ['K10']:
            ws[cell_ref] = value
            ws[cell_ref].font = font_normal_bold
            ws[cell_ref].border = thin_border
            ws[cell_ref].alignment = center

    # Row 11~32: 데이터 행 (최대 22행)
    data_start_row = 11
    data_end_row = 32
    row_idx = data_start_row

    for _, grp in price_groups.iterrows():
        price = int(grp['요금'])
        qty = int(grp['수량'])
        plan_info = price_to_plan.get(price, ('Unknown', ''))

        ws.row_dimensions[row_idx].height = 22.5

        ws[f'A{row_idx}'] = billing_month
        ws[f'A{row_idx}'].font = font_small
        ws[f'A{row_idx}'].alignment = center
        ws[f'A{row_idx}'].border = thin_border

        ws[f'B{row_idx}'] = recipient_name
        ws[f'B{row_idx}'].font = font_data
        ws[f'B{row_idx}'].alignment = center
        ws[f'B{row_idx}'].border = thin_border

        ws[f'C{row_idx}'] = plan_info[0]
        ws[f'C{row_idx}'].font = font_small
        ws[f'C{row_idx}'].alignment = center
        ws[f'C{row_idx}'].border = thin_border

        ws.merge_cells(f'D{row_idx}:F{row_idx}')
        ws[f'D{row_idx}'] = plan_info[1]
        ws[f'D{row_idx}'].font = font_data
        ws[f'D{row_idx}'].alignment = center
        ws[f'D{row_idx}'].border = thin_border

        ws[f'G{row_idx}'] = qty
        ws[f'G{row_idx}'].font = font_small
        ws[f'G{row_idx}'].alignment = center
        ws[f'G{row_idx}'].border = thin_border

        ws[f'H{row_idx}'] = 1
        ws[f'H{row_idx}'].font = font_small
        ws[f'H{row_idx}'].alignment = center
        ws[f'H{row_idx}'].border = thin_border

        ws.merge_cells(f'I{row_idx}:J{row_idx}')
        ws[f'I{row_idx}'] = price
        ws[f'I{row_idx}'].font = font_small_plain
        ws[f'I{row_idx}'].alignment = center
        ws[f'I{row_idx}'].border = thin_border
        ws[f'I{row_idx}'].number_format = '#,##0'

        ws.merge_cells(f'K{row_idx}:M{row_idx}')
        ws[f'K{row_idx}'] = f'=I{row_idx}*G{row_idx}*H{row_idx}'
        ws[f'K{row_idx}'].font = font_small
        ws[f'K{row_idx}'].alignment = right_align
        ws[f'K{row_idx}'].border = thin_border
        ws[f'K{row_idx}'].number_format = '#,##0'

        row_idx += 1

    # 빈 데이터 행 채우기 (테두리만)
    for r in range(row_idx, data_end_row + 1):
        ws.row_dimensions[r].height = 22.5
        for col in ['A', 'B', 'C', 'G', 'H']:
            ws[f'{col}{r}'].border = thin_border
        ws.merge_cells(f'D{r}:F{r}')
        ws[f'D{r}'].border = thin_border
        ws.merge_cells(f'I{r}:J{r}')
        ws[f'I{r}'].border = thin_border
        ws.merge_cells(f'K{r}:M{r}')
        ws[f'K{r}'].border = thin_border

    # Row 33: 구분선 (작은 행)
    ws.row_dimensions[33].height = 6

    # Row 34~36: 공급가/부가세/합계
    for r in [34, 35, 36]:
        ws.row_dimensions[r].height = 22.5

    ws.merge_cells('A34:G36')
    ws['A34'] = '비  고 : \n1) 상세 데이터 별도 제공'
    ws['A34'].font = font_small
    ws['A34'].alignment = Alignment(vertical='top', wrap_text=True)
    ws['A34'].border = thin_border

    ws.merge_cells('H34:J34')
    ws['H34'] = '공  급  가'
    ws['H34'].font = font_normal_bold
    ws['H34'].alignment = center
    ws['H34'].border = thin_border

    ws.merge_cells('K34:M34')
    ws['K34'] = f'=SUM(K{data_start_row}:M{data_end_row})'
    ws['K34'].font = font_normal_bold
    ws['K34'].alignment = right_align
    ws['K34'].border = thin_border
    ws['K34'].number_format = '#,##0'

    ws.merge_cells('H35:J35')
    ws['H35'] = '부가세'
    ws['H35'].font = font_normal_bold
    ws['H35'].alignment = center
    ws['H35'].border = thin_border

    ws.merge_cells('K35:M35')
    ws['K35'] = '=K34*10%'
    ws['K35'].font = font_normal_bold
    ws['K35'].alignment = right_align
    ws['K35'].border = thin_border
    ws['K35'].number_format = '#,##0'

    ws.merge_cells('H36:J36')
    ws['H36'] = '합 계 금 액'
    ws['H36'].font = font_normal_bold
    ws['H36'].alignment = center
    ws['H36'].border = thin_border

    ws.merge_cells('K36:M36')
    ws['K36'] = '=SUM(K34:M35)'
    ws['K36'].font = font_normal_bold
    ws['K36'].alignment = right_align
    ws['K36'].border = thin_border
    ws['K36'].number_format = '#,##0'

    # Row 37: 입금정보
    ws.row_dimensions[37].height = 34.5
    ws.merge_cells('A37:M37')
    ws['A37'] = f'     입금정보 : {SUPPLIER_INFO["bank"]}'
    ws['A37'].font = font_normal_bold
    ws['A37'].border = thin_border
    ws['A37'].alignment = Alignment(vertical='center')

    # 테두리 추가 (공급자 영역)
    for r in range(5, 9):
        for c in ['D', 'E', 'H', 'K', 'M']:
            ws[f'{c}{r}'].border = thin_border

    return wb


def create_summary_sheet(df_raw):
    """담당지사별 요약 테이블 생성"""
    billing_ok = df_raw[df_raw['과금 가능 여부'] == '가능']

    # 서비스 매핑
    def categorize(row):
        if row['요금'] in [54000, 60000]:
            return '스탠다드'
        elif row['요금'] in [43200, 48000]:
            return '베이직'
        return '기타'

    billing_ok = billing_ok.copy()
    billing_ok['카테고리'] = billing_ok.apply(categorize, axis=1)

    pivot = billing_ok.pivot_table(
        values='요금', index='담당지사', columns='카테고리',
        aggfunc='sum', fill_value=0
    )

    if '베이직' not in pivot.columns:
        pivot['베이직'] = 0
    if '스탠다드' not in pivot.columns:
        pivot['스탠다드'] = 0

    pivot['합계'] = pivot.sum(axis=1)
    pivot = pivot[['베이직', '스탠다드', '합계']]
    pivot.loc['합계'] = pivot.sum()

    # VAT 포함 행
    vat_row = pivot.loc['합계'] * 1.1
    vat_row.name = 'VAT 포함'
    pivot = pd.concat([pivot, pd.DataFrame([vat_row])])

    return pivot


def create_detail_excel(df_raw, billing_month):
    """
    별도 제공자료 엑셀 생성 (요약 시트 + Raw 시트)

    Parameters:
        df_raw: 과금 Raw DataFrame (요금 포함)
        billing_month: 이용월 (예: "26.01" → 시트명 Raw_2601)
    """
    wb = Workbook()

    # ===== 시트1: 요약 =====
    ws_summary = wb.active
    ws_summary.title = '요약'

    billing_ok = df_raw[df_raw['과금 가능 여부'] == '가능'].copy()

    # 카테고리 분류
    def categorize(row):
        if row['요금'] in [54000, 60000]:
            return '스탠다드'
        elif row['요금'] in [43200, 48000]:
            return '베이직'
        return '기타'

    billing_ok['카테고리'] = billing_ok.apply(categorize, axis=1)

    pivot = billing_ok.pivot_table(
        values='요금', index='담당지사', columns='카테고리',
        aggfunc='sum', fill_value=0
    )

    if '베이직' not in pivot.columns:
        pivot['베이직'] = 0
    if '스탠다드' not in pivot.columns:
        pivot['스탠다드'] = 0

    pivot['Grand Total'] = pivot.sum(axis=1)
    pivot = pivot[['베이직', '스탠다드', 'Grand Total']]
    pivot = pivot.sort_index()

    # 헤더 스타일
    header_font = Font(name='Dotum', size=11, bold=True)
    data_font = Font(name='Dotum', size=10)
    num_fmt = '#,##0'
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    center = Alignment(horizontal='center', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')

    # 헤더 작성
    headers = ['담당지사', '베이직', '스탠다드', 'Grand Total']
    for col_idx, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center

    # 데이터 작성
    for row_idx, (jisa, row_data) in enumerate(pivot.iterrows(), 2):
        cell = ws_summary.cell(row=row_idx, column=1, value=jisa)
        cell.font = data_font
        cell.border = thin_border

        for col_idx, col_name in enumerate(['베이직', '스탠다드', 'Grand Total'], 2):
            val = row_data.get(col_name, 0)
            cell = ws_summary.cell(row=row_idx, column=col_idx, value=val if val != 0 else None)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = right_align
            if val != 0:
                cell.number_format = num_fmt

    # 열 너비
    ws_summary.column_dimensions['A'].width = 20
    ws_summary.column_dimensions['B'].width = 15
    ws_summary.column_dimensions['C'].width = 15
    ws_summary.column_dimensions['D'].width = 15

    # ===== 시트2: Raw 데이터 =====
    month_code = billing_month.replace('.', '')
    raw_sheet_name = f'Raw_{month_code}'
    ws_raw = wb.create_sheet(title=raw_sheet_name)

    # Raw 컬럼 정의
    raw_columns = [
        '사이트아이디', '기관명', '반명', '가능한 일자 수', '성공 일자 수',
        '스토리라인 성공률', '담당지사', '요금제', '리포트 전송 여부',
        '과금 가능 여부', '요금', '서비스', '비고'
    ]

    # 헤더
    for col_idx, h in enumerate(raw_columns, 1):
        cell = ws_raw.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center

    # 데이터 (과금 가능한 건만)
    for row_idx, (_, row) in enumerate(billing_ok.iterrows(), 2):
        for col_idx, col_name in enumerate(raw_columns, 1):
            val = row.get(col_name, '')
            if pd.isna(val):
                val = ''
            cell = ws_raw.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            if col_name == '스토리라인 성공률' and isinstance(val, (int, float)):
                cell.number_format = '0.0000'
            elif col_name == '요금' and isinstance(val, (int, float)):
                cell.number_format = num_fmt

    # 열 너비 자동조절
    col_widths = [18, 15, 12, 12, 12, 15, 15, 10, 12, 12, 12, 10, 10]
    for i, w in enumerate(col_widths):
        ws_raw.column_dimensions[get_column_letter(i + 1)].width = w

    return wb
