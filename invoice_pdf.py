"""
거래명세서 PDF 생성 모듈 v2.2
- 폰트 크기 통일 (데이터 셀 일괄 9pt)
- 합계금액 콤마 표시
- 좌우 여백 조정 (표가 페이지 끝까지 안 차도록)
- 2025 카테고리: 2025_주5회/주3회 (약정 3년) + 2025_한결_주5회/주3회 (약정 3년)
- 도장(stamp.png) 자동 삽입
- 공급자 주소: 서울 강남구 삼성동 천해빌딩 10층
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from io import BytesIO
import calendar
import os

# 한글 CID 폰트 등록
pdfmetrics.registerFont(UnicodeCIDFont('HYSMyeongJo-Medium'))
pdfmetrics.registerFont(UnicodeCIDFont('HYGothic-Medium'))
FONT_G = 'HYGothic-Medium'  # 고딕 (한글 텍스트)
FONT_M = 'HYSMyeongJo-Medium'  # 명조 (숫자)

# ========== 폰트 크기 (통일) ==========
SIZE_TITLE = 22
SIZE_RECIPIENT = 13
SIZE_TOTAL_LABEL = 12
SIZE_TOTAL_VALUE = 12
SIZE_HEADER_LABEL = 9    # 사업자등록번호 등
SIZE_HEADER_VALUE = 10   # 870-88-02332 등
SIZE_TABLE_HEADER = 9    # 이용월, 대리점 등
SIZE_TABLE_DATA = 9      # 데이터 셀 (통일)
SIZE_SUMMARY_LABEL = 9   # 공급가, 부가세
SIZE_SUMMARY_VALUE = 10  # 금액
SIZE_BANK = 10

# ========== 공급자 정보 ==========
SUPPLIER = {
    'biz_no': '870-88-02332',
    'company': '플레이태그',
    'ceo': '박현수',
    'phone': '02-553-0214',
    'address': '서울 강남구 삼성동 천해빌딩 10층',
    'bank': '하나은행 / 플레이태그주식회사 / 403-910059-30704',
}


def _categorize_row(row):
    """
    거래명세서 카테고리 분류.
    우선순위: contract_year → 가격 + 담당지사 추정
    Returns: (category, 약정_label, sort_order)
    """
    price = int(row.get('요금', 0) or 0)
    jisa = str(row.get('담당지사', '') or '').strip()
    contract_year = str(row.get('contract_year', '') or '').strip()

    if price == 72000 or contract_year == '디렉토리':
        return ('디렉토리', 'BASIC(3)', 999)

    if price == 54000:
        return ('2024_주5회', '3년', 2)
    if price == 43200:
        return ('2024_주3회', '3년', 4)

    if price == 95000:
        return ('2025_주5회', '3년', 5)
    if price == 83000:
        return ('2025_주3회', '3년', 6)

    if price == 60000:
        if contract_year == '2025':
            return ('2025_한결_주5회', '3년', 7)
        if contract_year == '2024':
            return ('2024_주5회', '1년', 1)
        if jisa == '한결교육':
            return ('2025_한결_주5회', '3년', 7)
        return ('2024_주5회', '1년', 1)

    if price == 48000:
        if contract_year == '2025':
            return ('2025_한결_주3회', '3년', 8)
        if contract_year == '2024':
            return ('2024_주3회', '1년', 3)
        if jisa == '한결교육':
            return ('2025_한결_주3회', '3년', 8)
        return ('2024_주3회', '1년', 3)

    return ('기타', '', 999)


def _fmt(n):
    return f"{int(n):,}"


def _draw_stamp(c, cx, cy, size=38):
    """도장 이미지 삽입 (없으면 (인) 텍스트로 대체)"""
    stamp_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'stamp.png')
    if os.path.exists(stamp_path):
        c.drawImage(stamp_path, cx - size/2, cy - size/2, size, size, mask='auto')
    else:
        c.setFont(FONT_G, 9)
        c.setFillColorRGB(0.85, 0.1, 0.1)
        c.drawCentredString(cx, cy - 4, '(인)')
        c.setFillColorRGB(0, 0, 0)


def _draw_cell(c, x, y, w, h, text='', font=FONT_G, size=9,
               align='center', border=True):
    if border:
        c.rect(x, y, w, h)
    if not text:
        return
    c.setFont(font, size)
    ty = y + h / 2 - size * 0.35
    if align == 'center':
        c.drawCentredString(x + w / 2, ty, str(text))
    elif align == 'right':
        c.drawRightString(x + w - 4, ty, str(text))
    elif align == 'left':
        c.drawString(x + 4, ty, str(text))


def _parse_billing_date(billing_month):
    """billing_month '26.03' → '2026년 3월 31일'"""
    try:
        yy, mm_ = billing_month.split('.')
        yr = 2000 + int(yy)
        mo = int(mm_)
        last_day = calendar.monthrange(yr, mo)[1]
        return f"{yr}년 {mo}월 {last_day}일"
    except Exception:
        from datetime import datetime
        return datetime.now().strftime("%Y년 %m월 %d일")


def _draw_invoice_header(c, W, H, LM, RM, recipient_name, recipient_info,
                          date_str, total_amount):
    """거래명세서 상단 영역 (제목, 날짜, 공급자/수신자, 합계금액)"""

    # 1. 상단 구분선
    y = H - 42
    c.setLineWidth(1.5)
    c.line(LM, y, RM, y)

    # 2. 제목
    y -= 38
    c.setFont(FONT_G, SIZE_TITLE)
    c.drawCentredString(W / 2, y, '거 래 명 세 서')
    y -= 10
    c.setLineWidth(0.3)
    c.line(LM, y, RM, y)

    # 3. 날짜
    y -= 20
    c.setFont(FONT_G, SIZE_HEADER_VALUE)
    c.drawString(LM, y, f'날짜 : {date_str}')

    # 4. 공급자 테이블 (우측 절반)
    st = y - 10
    rh = 21
    sl = LM + (RM - LM) * 0.45  # 좌측 마진 기준으로 위치 조정
    lbl_w = 24
    fld_w = 80
    val_x = sl + lbl_w + fld_w

    c.setLineWidth(0.5)
    c.rect(sl, st - rh * 4, lbl_w, rh * 4)
    c.setFont(FONT_G, SIZE_HEADER_VALUE)
    cx = sl + lbl_w / 2
    c.drawCentredString(cx, st - rh * 1 + 5, '공')
    c.drawCentredString(cx, st - rh * 2 + 5, '급')
    c.drawCentredString(cx, st - rh * 3 + 5, '자')

    # R1: 사업자등록번호
    ry = st - rh
    _draw_cell(c, sl + lbl_w, ry, fld_w, rh, '사업자등록번호', size=SIZE_HEADER_LABEL)
    _draw_cell(c, val_x, ry, RM - val_x, rh, SUPPLIER['biz_no'],
               font=FONT_M, size=SIZE_HEADER_VALUE)

    # R2: 상호 + 대표자 (도장!)
    ry = st - rh * 2
    _draw_cell(c, sl + lbl_w, ry, fld_w, rh, '상호', size=SIZE_HEADER_LABEL)
    mid = val_x + (RM - val_x) * 0.48
    rep_lbl = mid + (RM - mid) * 0.38
    _draw_cell(c, val_x, ry, mid - val_x, rh, SUPPLIER['company'],
               font=FONT_M, size=SIZE_HEADER_VALUE)
    _draw_cell(c, mid, ry, rep_lbl - mid, rh, '대표자', size=SIZE_HEADER_LABEL)
    _draw_cell(c, rep_lbl, ry, RM - rep_lbl, rh,
               f'{SUPPLIER["ceo"]}', font=FONT_M, size=SIZE_HEADER_VALUE, align='left')
    # 도장
    _draw_stamp(c, RM - 24, ry + rh / 2, size=38)

    # R3: 전화번호
    ry = st - rh * 3
    _draw_cell(c, sl + lbl_w, ry, fld_w, rh, '전화번호', size=SIZE_HEADER_LABEL)
    _draw_cell(c, val_x, ry, RM - val_x, rh, SUPPLIER['phone'],
               font=FONT_M, size=SIZE_HEADER_VALUE)

    # R4: 주소
    ry = st - rh * 4
    _draw_cell(c, sl + lbl_w, ry, fld_w, rh, '주소', size=SIZE_HEADER_LABEL)
    _draw_cell(c, val_x, ry, RM - val_x, rh, SUPPLIER['address'],
               font=FONT_M, size=SIZE_HEADER_VALUE)

    # 5. 수신자 정보 (좌측)
    c.setFont(FONT_G, SIZE_RECIPIENT)
    c.drawString(LM + 12, st - rh - 8, f'{recipient_name} 귀하')

    c.setFont(FONT_M, SIZE_HEADER_LABEL)
    iy = st - rh * 2 + 2
    addr = recipient_info.get('address', '')
    if addr:
        if len(addr) > 22:
            c.drawString(LM + 12, iy, f'주소 : {addr[:22]}')
            iy -= 13
            c.drawString(LM + 12 + 36, iy, addr[22:])
        else:
            c.drawString(LM + 12, iy, f'주소 : {addr}')
        iy -= 13

    biz = recipient_info.get('biz_no', '')
    if biz:
        c.drawString(LM + 12, iy, f'사업자 번호 : {biz}')
        iy -= 13
    email = recipient_info.get('email', '')
    if email:
        c.drawString(LM + 12, iy, f'이메일 : {email}')

    # 6. 합계금액 행 (콤마 적용!)
    y_sum = st - rh * 4 - 20
    c.setFont(FONT_G, SIZE_TOTAL_LABEL)
    c.drawString(LM + 12, y_sum, '합계금액 :')
    c.drawString(LM + 100, y_sum, f'{_fmt(total_amount)} 원정')
    c.drawString(W / 2 + 80, y_sum, '₩')
    c.drawRightString(RM - 4, y_sum, _fmt(total_amount))

    return y_sum


def _draw_invoice_footer(c, W, LM, RM, hdr_y, max_rows, drh,
                          col_table, supply, vat, total, 비고_text):
    """비고 + 합계 + 입금정보"""
    PW = RM - LM

    # 비고 영역의 우측 경계 = 단가 칸 시작 위치
    danga_x = col_table[6][1]    # 단가 컬럼 x
    danga_w = col_table[6][2]
    geum_x = col_table[7][1]
    geum_w = col_table[7][2]

    bottom = hdr_y - drh - (max_rows + 1) * drh
    sy = bottom - 6
    sh = drh

    # 비고 영역 (3행 높이)
    bigo_h = sh * 3
    c.rect(LM, sy - bigo_h, danga_x - LM, bigo_h)
    c.setFont(FONT_G, SIZE_SUMMARY_LABEL)
    c.drawString(LM + 5, sy - 14, '비 고 :')
    if 비고_text:
        c.setFont(FONT_M, SIZE_SUMMARY_LABEL)
        c.drawString(LM + 5, sy - 28, 비고_text)

    # 공급가 / 부가세 / 합계
    items = [
        ('공 급 가', supply),
        ('부가세', vat),
        ('합 계 금 액', total),
    ]
    for i, (label, value) in enumerate(items):
        ry = sy - (i + 1) * sh
        _draw_cell(c, danga_x, ry, danga_w, sh, label, size=SIZE_SUMMARY_LABEL)
        _draw_cell(c, geum_x, ry, geum_w, sh, _fmt(value),
                   font=FONT_M, size=SIZE_SUMMARY_VALUE, align='right')

    # 입금정보
    bank_y = sy - bigo_h - 14
    c.setFont(FONT_G, SIZE_BANK)
    c.drawString(LM + 5, bank_y, f'입금정보 : {SUPPLIER["bank"]}')


# ========== 총판 거래명세서 PDF ==========

def create_invoice_pdf(df_raw, recipient_name, billing_month, recipient_info=None):
    """
    총판 거래명세서 PDF 생성
    카테고리: 2024_주5회/주3회 + 2025_주5회/주3회 + 2025_한결_주5회/주3회
    """
    if recipient_info is None:
        recipient_info = {'address': '', 'biz_no': '', 'email': ''}

    billing_ok = df_raw[df_raw['과금 가능 여부'] == '가능'].copy()

    cats = billing_ok.apply(_categorize_row, axis=1)
    billing_ok['category'] = cats.apply(lambda x: x[0])
    billing_ok['약정_label'] = cats.apply(lambda x: x[1])
    billing_ok['sort_order'] = cats.apply(lambda x: x[2])

    grouped = billing_ok.groupby(
        ['sort_order', 'category', '약정_label', '요금']
    ).size().reset_index(name='수량')
    grouped = grouped.sort_values(['sort_order', '요금'], ascending=[True, False])

    data_rows = []
    for _, grp in grouped.iterrows():
        price = int(grp['요금'])
        qty = int(grp['수량'])
        amount = price * qty  # 100% rate
        data_rows.append({
            'plan': grp['category'],
            'contract': grp['약정_label'],
            'qty': qty,
            'price': price,
            'amount': amount,
            'rate': '100%',
        })

    supply = sum(r['amount'] for r in data_rows)
    vat = int(supply * 0.1)
    total = supply + vat

    date_str = _parse_billing_date(billing_month)

    # === PDF 생성 ===
    buf = BytesIO()
    W, H = A4
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle('거래명세서_총판')

    # 여백 설정 (이전보다 넓게)
    LM = 55
    RM = W - 55
    PW = RM - LM

    y_sum = _draw_invoice_header(c, W, H, LM, RM, recipient_name, recipient_info,
                                  date_str, total)

    # 데이터 테이블
    hdr_y = y_sum - 16
    hdr_h = 20
    drh = 19
    MAX_ROWS = 18

    # 컬럼 정의 (총 PW = 485)
    # 이용월(45) + 대리점(105) + 요금제(95) + 약정(35) + 수량(35) + 요율(35) + 단가(55) + 금액(80)
    col_widths = [45, 105, 95, 35, 35, 35, 55, 80]
    col_x = [LM]
    for w in col_widths[:-1]:
        col_x.append(col_x[-1] + w)
    col_names = ['이용월', '대리점', '요금제', '약정', '수량', '요율', '단가', '금 액']
    C = list(zip(col_names, col_x, col_widths))

    # 헤더
    for name, x, w in C:
        _draw_cell(c, x, hdr_y - hdr_h, w, hdr_h, name, size=SIZE_TABLE_HEADER)

    # 데이터 행 (모든 셀 SIZE_TABLE_DATA로 통일)
    for i in range(MAX_ROWS):
        ry = hdr_y - hdr_h - (i + 1) * drh
        if i < len(data_rows):
            d = data_rows[i]
            vals = [
                (billing_month, FONT_G, 'center'),
                (recipient_name, FONT_G, 'center'),
                (d['plan'], FONT_G, 'center'),
                (str(d['contract']), FONT_G, 'center'),
                (str(d['qty']), FONT_G, 'center'),
                (d['rate'], FONT_G, 'center'),
                (_fmt(d['price']), FONT_M, 'right'),
                (_fmt(d['amount']), FONT_M, 'right'),
            ]
        else:
            vals = [
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('- -', FONT_M, 'right'),
                ('- -', FONT_M, 'right'),
            ]
        for j, (name, x, w) in enumerate(C):
            text, font, al = vals[j]
            _draw_cell(c, x, ry, w, drh, text,
                       font=font, size=SIZE_TABLE_DATA, align=al)

    # 푸터
    _draw_invoice_footer(c, W, LM, RM, hdr_y, MAX_ROWS, drh, C,
                         supply, vat, total, '1) 상세 데이터 별도 제공')

    c.save()
    buf.seek(0)
    return buf


# ========== 디렉토리 거래명세서 PDF ==========

def create_directory_invoice_pdf(df_directory, recipient_name='디렉토리',
                                   billing_month='', recipient_info=None):
    """
    디렉토리 거래명세서 PDF 생성
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

    data_rows = []
    for _, grp in grouped.iterrows():
        qty = int(grp['수량'])
        price = int(grp['요금'])
        amount = int(qty * price * 0.5)  # 50% rate
        data_rows.append({
            'institution': grp['기관명'],
            'category': 'BASIC(3)',
            'qty': qty,
            'price': price,
            'amount': amount,
            'rate': '50%',
        })

    supply = sum(r['amount'] for r in data_rows)
    vat = int(supply * 0.1)
    total = supply + vat

    date_str = _parse_billing_date(billing_month)

    # === PDF 생성 ===
    buf = BytesIO()
    W, H = A4
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle('거래명세서_디렉토리')

    LM = 55
    RM = W - 55
    PW = RM - LM

    y_sum = _draw_invoice_header(c, W, H, LM, RM, recipient_name, recipient_info,
                                  date_str, total)

    # 데이터 테이블
    hdr_y = y_sum - 16
    hdr_h = 20
    drh = 19
    MAX_ROWS = 18

    # 컬럼 정의 (총 PW = 485)
    # 이용월(45) + 대리점(60) + 기관명(140) + 구분(35) + 수량(35) + 요율(35) + 단가(55) + 금액(80)
    col_widths = [45, 60, 140, 35, 35, 35, 55, 80]
    col_x = [LM]
    for w in col_widths[:-1]:
        col_x.append(col_x[-1] + w)
    col_names = ['이용월', '대리점', '기관명', '구분', '수량', '요율', '단가', '금 액']
    C = list(zip(col_names, col_x, col_widths))

    # 헤더
    for name, x, w in C:
        _draw_cell(c, x, hdr_y - hdr_h, w, hdr_h, name, size=SIZE_TABLE_HEADER)

    # 데이터 행
    for i in range(MAX_ROWS):
        ry = hdr_y - hdr_h - (i + 1) * drh
        if i < len(data_rows):
            d = data_rows[i]
            vals = [
                (billing_month, FONT_G, 'center'),
                ('디렉토리', FONT_G, 'center'),
                (d['institution'], FONT_G, 'center'),
                (d['category'], FONT_G, 'center'),
                (str(d['qty']), FONT_G, 'center'),
                (d['rate'], FONT_G, 'center'),
                (_fmt(d['price']), FONT_M, 'right'),
                (_fmt(d['amount']), FONT_M, 'right'),
            ]
        else:
            vals = [
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('', FONT_G, 'center'),
                ('- -', FONT_M, 'right'),
                ('- -', FONT_M, 'right'),
            ]
        for j, (name, x, w) in enumerate(C):
            text, font, al = vals[j]
            _draw_cell(c, x, ry, w, drh, text,
                       font=font, size=SIZE_TABLE_DATA, align=al)

    _draw_invoice_footer(c, W, LM, RM, hdr_y, MAX_ROWS, drh, C,
                         supply, vat, total, '')

    c.save()
    buf.seek(0)
    return buf
