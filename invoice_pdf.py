"""
거래명세서 PDF 직접 생성 모듈
- reportlab CID 한글 폰트 사용 (외부 폰트 불필요)
- 완성본 양식과 동일한 레이아웃
- Mac Preview / Chrome / Adobe Reader 에서 정상 한글 출력
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
pdfmetrics.registerFont(UnicodeCIDFont('HYSMyeongJo-Medium'))  # 명조
pdfmetrics.registerFont(UnicodeCIDFont('HYGothic-Medium'))      # 고딕

FONT_G = 'HYGothic-Medium'
FONT_M = 'HYSMyeongJo-Medium'

# 공급자 정보
SUPPLIER = {
    'biz_no': '870-88-02332',
    'company': '플레이태그',
    'ceo': '박현수',
    'phone': '02-553-0214',
    'address': '서울 강남구 강남대로140길 9, 5층',
    'bank': '하나은행 / 플레이태그주식회사 / 403-910059-30704',
}

PRICE_TO_PLAN = {
    60000: ('Standard', '1년'),
    54000: ('Standard', '3년'),
    48000: ('Basic', '1년'),
    43200: ('Basic', '3년'),
}


def _fmt(n):
    return f"{int(n):,}"


def _draw_stamp(c, cx, cy, size=38):
    """
    도장 이미지를 PDF에 삽입
    cx, cy: 도장 중심 좌표
    size: 도장 크기 (가로=세로)
    """
    stamp_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'stamp.png')
    if os.path.exists(stamp_path):
        c.drawImage(stamp_path, cx - size/2, cy - size/2, size, size, mask='auto')
    else:
        # fallback: (인) 텍스트
        c.setFont(FONT_G, 9)
        c.setFillColorRGB(0.85, 0.1, 0.1)
        c.drawCentredString(cx, cy - 4, '(인)')
        c.setFillColorRGB(0, 0, 0)


def _draw_cell(c, x, y, w, h, text='', font=FONT_G, size=9,
               align='center', border=True, bold=False):
    """셀 하나를 그리는 유틸리티"""
    if border:
        c.rect(x, y, w, h)
    if not text:
        return

    c.setFont(font, size)
    ty = y + h / 2 - size * 0.35  # 수직 중앙

    if align == 'center':
        c.drawCentredString(x + w / 2, ty, str(text))
    elif align == 'right':
        c.drawRightString(x + w - 4, ty, str(text))
    elif align == 'left':
        c.drawString(x + 4, ty, str(text))


def create_invoice_pdf(df_raw, recipient_name, billing_month, recipient_info=None):
    """
    거래명세서 PDF 생성 (완성본 양식 재현)

    Returns: BytesIO (PDF 데이터)
    """
    if recipient_info is None:
        recipient_info = {'address': '', 'biz_no': '', 'email': ''}

    # === 데이터 준비 ===
    billing_ok = df_raw[df_raw['과금 가능 여부'] == '가능']
    price_groups = billing_ok.groupby('요금').size().reset_index(name='수량')
    price_groups = price_groups[price_groups['요금'] > 0].sort_values('요금', ascending=False)

    data_rows = []
    for _, grp in price_groups.iterrows():
        price = int(grp['요금'])
        qty = int(grp['수량'])
        plan, contract = PRICE_TO_PLAN.get(price, ('Unknown', ''))
        data_rows.append({
            'plan': plan, 'contract': contract,
            'qty': qty, 'price': price,
            'amount': price * qty,
        })

    supply = sum(r['amount'] for r in data_rows)
    vat = int(supply * 0.1)
    total = supply + vat

    # 날짜 계산
    try:
        yy, mm_ = billing_month.split('.')
        yr = 2000 + int(yy)
        mo = int(mm_)
        last_day = calendar.monthrange(yr, mo)[1]
        date_str = f"{yr}년 {mo}월 {last_day}일"
    except Exception:
        from datetime import datetime
        date_str = datetime.now().strftime("%Y년 %m월 %d일")

    # === PDF 캔버스 ===
    buf = BytesIO()
    W, H = A4
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle('거래명세서')

    LM = 42        # 좌측 마진
    RM = W - 42     # 우측 끝
    PW = RM - LM    # 페이지 폭

    # ──────────────────────────────────────────
    # 1. 상단 구분선
    # ──────────────────────────────────────────
    y = H - 42
    c.setLineWidth(1.5)
    c.line(LM, y, RM, y)

    # ──────────────────────────────────────────
    # 2. 제목
    # ──────────────────────────────────────────
    y -= 38
    c.setFont(FONT_G, 22)
    c.drawCentredString(W / 2, y, '거  래  명  세  서')
    y -= 10
    c.setLineWidth(0.3)
    c.line(LM, y, RM, y)

    # ──────────────────────────────────────────
    # 3. 날짜
    # ──────────────────────────────────────────
    y -= 20
    c.setFont(FONT_G, 10)
    c.drawString(LM, y, f'날짜 :    {date_str}')

    # ──────────────────────────────────────────
    # 4. 공급자 테이블 (우측 절반)
    # ──────────────────────────────────────────
    st = y - 10  # supplier table top
    rh = 21      # row height
    sl = W * 0.42  # 공급자 테이블 좌측 시작
    lbl_w = 24   # "공급자" 라벨 칸
    fld_w = 76   # 필드명 칸 (사업자등록번호 등)
    val_x = sl + lbl_w + fld_w  # 값 시작

    # 공/급/자 세로 병합 셀
    c.setLineWidth(0.5)
    c.rect(sl, st - rh * 4, lbl_w, rh * 4)
    c.setFont(FONT_G, 10)
    cx = sl + lbl_w / 2
    c.drawCentredString(cx, st - rh * 1 + 5, '공')
    c.drawCentredString(cx, st - rh * 2 + 5, '급')
    c.drawCentredString(cx, st - rh * 3 + 5, '자')

    # R1: 사업자등록번호
    ry = st - rh
    _draw_cell(c, sl + lbl_w, ry, fld_w, rh, '사업자등록번호', size=9)
    _draw_cell(c, val_x, ry, RM - val_x, rh, SUPPLIER['biz_no'], font=FONT_M, size=10)

    # R2: 상호 + 대표자
    ry = st - rh * 2
    _draw_cell(c, sl + lbl_w, ry, fld_w, rh, '상호', size=9)
    mid = val_x + (RM - val_x) * 0.48
    rep_lbl = mid + (RM - mid) * 0.38
    _draw_cell(c, val_x, ry, mid - val_x, rh, SUPPLIER['company'], font=FONT_M, size=10)
    _draw_cell(c, mid, ry, rep_lbl - mid, rh, '대표자', size=9)
    _draw_cell(c, rep_lbl, ry, RM - rep_lbl, rh,
               f'{SUPPLIER["ceo"]}', font=FONT_M, size=10, align='left')
    # 도장 찍기 (대표자 이름 오른쪽)
    _draw_stamp(c, RM - 24, ry + rh / 2, size=38)

    # R3: 전화번호
    ry = st - rh * 3
    _draw_cell(c, sl + lbl_w, ry, fld_w, rh, '전화번호', size=9)
    _draw_cell(c, val_x, ry, RM - val_x, rh, SUPPLIER['phone'], font=FONT_M, size=10)

    # R4: 주소
    ry = st - rh * 4
    _draw_cell(c, sl + lbl_w, ry, fld_w, rh, '주소', size=9)
    _draw_cell(c, val_x, ry, RM - val_x, rh, SUPPLIER['address'], font=FONT_M, size=10)

    # ──────────────────────────────────────────
    # 5. 수신자 정보 (좌측)
    # ──────────────────────────────────────────
    c.setFont(FONT_G, 13)
    c.drawString(LM + 12, st - rh - 8, f'{recipient_name}  귀하')

    c.setFont(FONT_M, 9)
    iy = st - rh * 2 + 2  # info y
    addr = recipient_info.get('address', '')
    if addr:
        # 긴 주소 줄바꿈
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

    # ──────────────────────────────────────────
    # 6. 합계금액 행
    # ──────────────────────────────────────────
    y_sum = st - rh * 4 - 20
    c.setFont(FONT_G, 12)
    c.drawString(LM + 12, y_sum, '합계금액 :')
    c.setFont(FONT_G, 11)
    c.drawString(LM + 90, y_sum, f'{int(total)}  원정')
    c.setFont(FONT_G, 11)
    c.drawString(W / 2 + 55, y_sum, '₩')
    c.setFont(FONT_G, 12)
    c.drawRightString(RM - 4, y_sum, _fmt(total))

    # ──────────────────────────────────────────
    # 7. 데이터 테이블
    # ──────────────────────────────────────────
    hdr_y = y_sum - 16
    hdr_h = 20
    drh = 19  # data row height
    MAX_ROWS = 18

    # 열 정의: (이름, x, w)
    C = [
        ('이용월',  LM,       50),
        ('대리점',  LM + 50,  116),
        ('요금제',  LM + 166, 78),
        ('약정',    LM + 244, 40),
        ('수량',    LM + 284, 40),
        ('요율',    LM + 324, 40),
        ('단가',    LM + 364, 56),
        ('금  액',  LM + 420, PW - 420 + LM),
    ]

    # 헤더
    for name, x, w in C:
        _draw_cell(c, x, hdr_y - hdr_h, w, hdr_h, name, size=9)

    # 데이터 행
    for i in range(MAX_ROWS):
        ry = hdr_y - hdr_h - (i + 1) * drh

        if i < len(data_rows):
            d = data_rows[i]
            vals = [
                (billing_month, FONT_G, 9, 'center'),
                (recipient_name, FONT_G, 7, 'center'),
                (d['plan'], FONT_G, 9, 'center'),
                (str(d['contract']), FONT_G, 9, 'center'),
                (str(d['qty']), FONT_G, 9, 'center'),
                ('100%', FONT_G, 9, 'center'),
                (_fmt(d['price']), FONT_M, 9, 'right'),
                (_fmt(d['amount']), FONT_M, 9, 'right'),
            ]
        else:
            vals = [
                ('', FONT_G, 9, 'center'),
                ('', FONT_G, 9, 'center'),
                ('', FONT_G, 9, 'center'),
                ('', FONT_G, 9, 'center'),
                ('', FONT_G, 9, 'center'),
                ('', FONT_G, 9, 'center'),
                ('- -', FONT_M, 9, 'right'),
                ('- -', FONT_M, 9, 'right'),
            ]

        for j, (name, x, w) in enumerate(C):
            text, font, sz, al = vals[j]
            _draw_cell(c, x, ry, w, drh, text, font=font, size=sz, align=al)

    # ──────────────────────────────────────────
    # 8. 하단: 비고 + 공급가/부가세/합계
    # ──────────────────────────────────────────
    bottom = hdr_y - hdr_h - (MAX_ROWS + 1) * drh
    sy = bottom - 6  # summary top
    sh = drh  # summary row height

    # 비고 영역 (좌측, 3행 높이)
    bigo_right = C[5][1] + C[5][2]  # 요율 칸 끝
    bigo_h = sh * 3
    c.rect(LM, sy - bigo_h, bigo_right - LM, bigo_h)
    c.setFont(FONT_G, 9)
    c.drawString(LM + 5, sy - 14, '비  고 :')
    c.setFont(FONT_M, 9)
    c.drawString(LM + 5, sy - 28, '1) 상세 데이터 별도 제공')

    # 공급가 / 부가세 / 합계금액 (우측)
    lbl_x = C[6][1]  # 단가 칸 시작
    lbl_w2 = C[6][2]
    val_x2 = C[7][1]
    val_w2 = C[7][2]

    items = [
        ('공  급  가', supply),
        ('부가세', vat),
        ('합 계 금 액', total),
    ]
    for i, (label, value) in enumerate(items):
        ry = sy - (i + 1) * sh
        _draw_cell(c, lbl_x, ry, lbl_w2, sh, label, size=9)
        _draw_cell(c, val_x2, ry, val_w2, sh, _fmt(value), font=FONT_M, size=10, align='right')

    # ──────────────────────────────────────────
    # 9. 입금정보
    # ──────────────────────────────────────────
    bank_y = sy - bigo_h - 14
    c.setFont(FONT_G, 10)
    c.drawString(LM + 5, bank_y, f'입금정보 : {SUPPLIER["bank"]}')

    # 완료
    c.save()
    buf.seek(0)
    return buf


# ===== 테스트 =====
if __name__ == '__main__':
    import pandas as pd
    import os

    # 완성본과 동일한 테스트 데이터
    test_data = (
        [{'과금 가능 여부': '가능', '요금': 60000}] * 11 +
        [{'과금 가능 여부': '가능', '요금': 54000}] * 225 +
        [{'과금 가능 여부': '가능', '요금': 48000}] * 68 +
        [{'과금 가능 여부': '가능', '요금': 43200}] * 159
    )
    df = pd.DataFrame(test_data)

    recipient_info = {
        'address': '인천광역시 서구 파랑로 495 2동 3층 302호 (청라 에이스)',
        'biz_no': '406-81-66140',
        'email': 'goldengate2021@naver.com',
    }

    pdf = create_invoice_pdf(
        df,
        recipient_name='문화사 외 16개 지사',
        billing_month='25.12',
        recipient_info=recipient_info,
    )

    out = '/sessions/dreamy-exciting-tesla/mnt/jin/Downloads/billing-app/test_invoice.pdf'
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'wb') as f:
        f.write(pdf.read())
    print(f"✅ PDF 생성: {out}")
