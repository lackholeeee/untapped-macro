import io
import os
import re
import zipfile
from collections import defaultdict
import streamlit as st
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 화면 전체 너비(wide) 사용
st.set_page_config(
    page_title="UNTAPPED STUDIO | 나염·발주 엑셀 자동 가공 웹매크로",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# UI 스타일 시트 (CSS)
st.markdown("""
<style>
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2rem !important;
        max-width: 98% !important;
    }
    body, [data-testid="stAppViewContainer"] {
        background-color: #EBF3FC !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }
    #MainMenu, footer, header {visibility: hidden;}

    /* 최상단 네이비 바 */
    .main-top-header {
        background-color: #1558B0;
        color: #ffffff;
        padding: 16px 28px;
        border-radius: 14px;
        box-shadow: 0 4px 12px rgba(21, 88, 176, 0.25);
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 12px;
    }
    .main-top-header h1 {
        font-size: 1.4rem !important;
        font-weight: 800 !important;
        margin: 0 !important;
        color: #ffffff !important;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .main-top-header h1 span {
        font-weight: 500;
        font-size: 1.2rem;
        color: #CFE2FE;
    }
    .pc-badge {
        background: #1D4ED8;
        border: 1px solid #60A5FA;
        border-radius: 9999px;
        padding: 6px 18px;
        font-size: 0.85rem;
        color: #ffffff;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* 상단 안내 배너 */
    .notice-banner {
        background-color: #EFF6FF;
        border: 1px solid #BFDBFE;
        border-radius: 12px;
        padding: 12px 20px;
        color: #1E3A8A;
        font-size: 0.95rem;
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 16px;
    }

    /* 변환 결과 카드 */
    .file-res-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 12px 16px;
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 8px;
    }
    .file-res-name {
        font-weight: 700;
        font-size: 0.88rem;
        color: #1E293B;
    }
    .file-res-desc {
        font-size: 0.76rem;
        color: #64748B;
        margin-top: 2px;
    }

    /* 관리자 확인요청 경고 박스 */
    .admin-warn-box {
        background-color: #FEF2F2;
        border: 2px solid #EF4444;
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 14px;
        color: #991B1B;
    }

    /* 다운로드 버튼 스타일링 */
    div[data-testid="stDownloadButton"] button {
        background-color: #1F7A4D !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-size: 0.8rem !important;
        font-weight: 600 !important;
        padding: 7px 14px !important;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #18633E !important;
    }
    .zip-wrap div[data-testid="stDownloadButton"] button {
        background-color: #1A73E8 !important;
        width: 100% !important;
        padding: 12px !important;
        font-size: 0.95rem !important;
        border-radius: 8px !important;
    }

    /* 우측 엑셀 미리보기 테이블 스타일 */
    .excel-view-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.85rem;
        text-align: center;
        background: #ffffff;
    }
    .excel-view-table th, .excel-view-table td {
        border: 1px solid #CBD5E1;
        padding: 8px 6px;
        white-space: nowrap;
    }
    .excel-head {
        background-color: #F8FAFC;
        font-weight: 700;
        color: #1E293B;
    }
    .row-yellow {
        background-color: #FFF2CC !important;
    }
    .row-red-alert {
        background-color: #FFF5F5 !important;
    }
    .callout-tag-blue {
        background: #2563EB;
        color: #ffffff;
        padding: 3px 12px;
        border-radius: 9999px;
        font-size: 0.76rem;
        font-weight: 600;
        display: inline-block;
    }
    .callout-tag-yellow {
        background: #D97706;
        color: #ffffff;
        padding: 3px 12px;
        border-radius: 9999px;
        font-size: 0.76rem;
        font-weight: 600;
        display: inline-block;
    }
    .callout-tag-red {
        background: #DC2626;
        color: #ffffff;
        padding: 3px 12px;
        border-radius: 9999px;
        font-size: 0.76rem;
        font-weight: 600;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# 최상단 헤더
st.markdown("""
<div class="main-top-header">
    <h1>⚙️ UNTAPPED STUDIO <span>| 나염·발주 엑셀 자동 가공 웹매크로</span></h1>
    <div class="pc-badge">💻 다른 PC에서도 동일하게 사용 가능</div>
</div>
<div class="notice-banner">
    <span style="font-size: 1.25rem;">💡</span>
    <span>파이썬 미설치 PC에서도 링크 접속 후 원본 <b>.xls</b> 파일을 끌어다 놓으면 <b>1초 만에 [가공완료].xlsx</b> 양식으로 자동 변환됩니다.</span>
</div>
""", unsafe_allow_html=True)

def clean_cell_value(val):
    if isinstance(val, float) and val.is_integer():
        return int(val)
    return val

def extract_base_option(opt_str):
    opt_str = str(opt_str).strip()
    m = re.match(r'\[(.*?)\]', opt_str)
    inner = m.group(1) if m else opt_str
    cleaned = re.sub(r'\(.*?\)', '', inner).strip()
    return f"[{cleaned}]"

# 의류 규격 사이즈 우선순위 정렬 (S, M, L, XL, 2XL, 3XL, 4XL)
size_order_map = {
    'XXS': 1, 'XS': 2, 'S': 3, 'M': 4, 'L': 5, 'XL': 6,
    '2XL': 7, 'XXL': 7, '3XL': 8, 'XXXL': 8, '4XL': 9, 'XXXXL': 9, '5XL': 10,
    'FREE': 99
}

def get_size_rank(size_str):
    s = size_str.upper().strip()
    return size_order_map.get(s, 50)

def inventory_sort_key(item):
    prod, opt = item[0]
    m = re.match(r'\[(.*?)-(.*?)\]', opt)
    if m:
        color, size = m.group(1).strip(), m.group(2).strip()
        return (prod, color, get_size_rank(size), size)
    return (prod, opt, 999, opt)

def process_single_file_bytes(file_bytes):
    wb_target = xlrd.open_workbook(file_contents=file_bytes)
    sheet = wb_target.sheet_by_index(0)
    top_title = sheet.cell_value(0, 0)

    rows_data = []
    raw_inventory_items = []
    admin_alerts = []

    for r in range(10, sheet.nrows - 1):
        c0 = clean_cell_value(sheet.cell_value(r, 0)) # 공급처상품명
        c2 = clean_cell_value(sheet.cell_value(r, 2)) # 상품코드
        c3 = clean_cell_value(sheet.cell_value(r, 3)) # 옵션
        c5 = clean_cell_value(sheet.cell_value(r, 5)) # 입고수량
        c4 = clean_cell_value(sheet.cell_value(r, 4)) # 바코드

        # 미배송 수량 확인 (7번 열)
        undelivered = 0
        if sheet.ncols > 7:
            try:
                undelivered = int(float(sheet.cell_value(r, 7)))
            except (ValueError, TypeError):
                undelivered = 0

        try:
            qty_num = int(float(c5))
        except (ValueError, TypeError):
            qty_num = 0

        is_over = False
        if sheet.ncols > 7 and qty_num > undelivered:
            is_over = True
            admin_alerts.append({
                "item_code": str(c2),
                "option": str(c3),
                "in_qty": qty_num,
                "undelivered": undelivered
            })

        rows_data.append([c0, c2, c3, qty_num, c4, is_over])
        raw_inventory_items.append((str(c0), str(c3), qty_num))

    # 구분 행 분할
    grouped_rows = []
    prev_item_code = None
    prev_logo_part = None

    for row in rows_data:
        item_code = str(row[1])
        opt_str = str(row[2])
        logo_part = opt_str.rsplit('-', 1)[0] if '-' in opt_str else opt_str

        if prev_item_code is not None:
            if item_code != prev_item_code:
                grouped_rows.append((['', '', '', '', '', False], 'item_change'))
            elif logo_part != prev_logo_part:
                grouped_rows.append((['', '', '', '', '', False], 'logo_change'))

        grouped_rows.append((row, 'data'))
        prev_item_code = item_code
        prev_logo_part = logo_part

    wb = openpyxl.Workbook()

    # ----------------------------------------------------
    # [시트 1] 가공리스트
    # ----------------------------------------------------
    ws1 = wb.active
    ws1.title = "가공리스트"
    ws1.views.sheetView[0].showGridLines = True

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 1~2행 전표 타이틀 병합
    ws1.merge_cells("A1:F2")
    cell_a1 = ws1["A1"]
    cell_a1.value = top_title
    cell_a1.font = Font(name="Malgun Gothic", size=16, bold=True)
    cell_a1.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(1, 3):
        ws1.row_dimensions[r].height = 25
        for c in range(1, 7):
            ws1.cell(row=r, column=c).border = border_all

    headers1 = ['공급처상품명', '상품코드', '옵션', '입고수량', '바코드', '비고']
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    ws1.row_dimensions[3].height = 30

    for c_idx, h in enumerate(headers1, 1):
        cell = ws1.cell(row=3, column=c_idx, value=h)
        cell.font = Font(name="Malgun Gothic", size=13, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill
        cell.border = border_all

    item_change_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    warn_qty_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_warn = Font(name="Malgun Gothic", size=13, bold=True, color="9C0006")

    curr_row1 = 4
    total_qty1 = 0

    for r_data, r_type in grouped_rows:
        ws1.row_dimensions[curr_row1].height = 30
        if r_type == 'item_change':
            for c_idx in range(1, 7):
                cell = ws1.cell(row=curr_row1, column=c_idx, value="")
                cell.border = border_all
                cell.fill = item_change_fill
        elif r_type == 'logo_change':
            for c_idx in range(1, 7):
                cell = ws1.cell(row=curr_row1, column=c_idx, value="")
                cell.border = border_all
        else:
            is_over = r_data[5]
            row_display_vals = [r_data[0], r_data[1], r_data[2], r_data[3], r_data[4]]
            
            for c_idx, val in enumerate(row_display_vals, 1):
                cell = ws1.cell(row=curr_row1, column=c_idx, value=val)
                cell.font = Font(name="Malgun Gothic", size=13, bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
                cell.border = border_all

                if c_idx == 4:
                    cell.number_format = '#,##0'
                    total_qty1 += val
                    if is_over:
                        cell.fill = warn_qty_fill
                        cell.font = font_warn
                elif c_idx in (2, 5) and isinstance(val, (int, float)):
                    cell.number_format = '@'

            cell_remark = ws1.cell(row=curr_row1, column=6)
            cell_remark.border = border_all
            cell_remark.alignment = Alignment(horizontal="center", vertical="center")
            if is_over:
                cell_remark.value = "⚠️ 관리자확인요청 (미배송 초과)"
                cell_remark.font = font_warn
                cell_remark.fill = warn_qty_fill
            else:
                cell_remark.value = ""

        curr_row1 += 1

    # 시트 1 합계 행
    ws1.row_dimensions[curr_row1].height = 30
    cell_total1 = ws1.cell(row=curr_row1, column=1, value="합계")
    cell_total1.font = Font(name="Malgun Gothic", size=13, bold=True)
    cell_total1.alignment = Alignment(horizontal="center", vertical="center")
    cell_total1.border = border_all

    for c_idx in range(2, 7):
        cell = ws1.cell(row=curr_row1, column=c_idx)
        cell.border = border_all
        if c_idx == 4:
            cell.value = total_qty1
            cell.font = Font(name="Malgun Gothic", size=13, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = '#,##0'

    for col in ws1.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < 3:
                continue
            if cell.value is not None:
                val_str = str(cell.value)
                length = sum(2 if ord(char) > 127 else 1 for char in val_str)
                if length > max_len:
                    max_len = length
        ws1.column_dimensions[col_letter].width = max(max_len * 1.3 + 4, 12)

    # ----------------------------------------------------
    # [시트 2] 원재고취합 (S, M, L, XL, 2XL, 3XL, 4XL 완벽 정렬)
    # ----------------------------------------------------
    ws2 = wb.create_sheet(title="원재고취합")
    ws2.views.sheetView[0].showGridLines = True

    agg_dict = defaultdict(int)
    for prod_name, opt_raw, qty_num in raw_inventory_items:
        base_opt = extract_base_option(opt_raw)
        agg_dict[(prod_name, base_opt)] += qty_num

    sorted_items = sorted(agg_dict.items(), key=inventory_sort_key)

    headers2 = ['공급처상품명', '옵션', '입고수량']
    ws2.row_dimensions[1].height = 30
    for c_idx, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=c_idx, value=h)
        cell.font = Font(name="Malgun Gothic", size=13, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill
        cell.border = border_all

    curr_row2 = 2
    total_qty2 = 0
    prev_color = None

    for (prod_name, base_opt), qty_val in sorted_items:
        m_color = re.match(r'\[(.*?)-', base_opt)
        curr_color = m_color.group(1).strip() if m_color else base_opt

        if prev_color is not None and curr_color != prev_color:
            ws2.row_dimensions[curr_row2].height = 28
            for c_idx in range(1, 4):
                cell = ws2.cell(row=curr_row2, column=c_idx, value="")
                cell.border = border_all
            curr_row2 += 1

        ws2.row_dimensions[curr_row2].height = 30
        row_vals = [prod_name, base_opt, qty_val]
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws2.cell(row=curr_row2, column=c_idx, value=val)
            cell.font = Font(name="Malgun Gothic", size=13, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_all
            if c_idx == 3:
                cell.number_format = '#,##0'

        total_qty2 += qty_val
        prev_color = curr_color
        curr_row2 += 1

    # 시트 2 합계
    ws2.row_dimensions[curr_row2].height = 30
    cell_total_label = ws2.cell(row=curr_row2, column=1, value="총합계")
    cell_total_label.font = Font(name="Malgun Gothic", size=13, bold=True)
    cell_total_label.alignment = Alignment(horizontal="center", vertical="center")
    cell_total_label.border = border_all

    cell_blank_opt = ws2.cell(row=curr_row2, column=2, value="")
    cell_blank_opt.border = border_all

    cell_total_val = ws2.cell(row=curr_row2, column=3, value=total_qty2)
    cell_total_val.font = Font(name="Malgun Gothic", size=13, bold=True)
    cell_total_val.alignment = Alignment(horizontal="center", vertical="center")
    cell_total_val.border = border_all
    cell_total_val.number_format = '#,##0'

    for col in ws2.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                val_str = str(cell.value)
                length = sum(2 if ord(char) > 127 else 1 for char in val_str)
                if length > max_len:
                    max_len = length
        ws2.column_dimensions[col_letter].width = max(max_len * 1.3 + 4, 14)

    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer.getvalue(), len(rows_data), total_qty1, admin_alerts

# =========================================================================
# 좌우 2열 배치
# =========================================================================
col_left, col_right = st.columns([1, 1], gap="medium")

if "all_alerts" not in st.session_state:
    st.session_state["all_alerts"] = []

# --- 좌측 열 ---
with col_left:
    st.markdown("""
    <div style="background:#0F52BA; color:white; padding:10px 16px; border-radius:10px 10px 0 0; font-weight:700; font-size:0.95rem; display:flex; align-items:center; gap:8px;">
        <span>💻</span> <span>타 PC 크롬 접속 시 보이는 웹매크로 화면</span>
    </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown("""
        <div style="background:#FFFFFF; border:1px solid #CBD5E1; border-top:none; border-radius:0 0 12px 12px; padding:18px;">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:12px;">
                <span style="font-size:1.4rem;">⚙️</span>
                <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:800; color:#1E293B;">UNTAPPED STUDIO</h3>
                    <p style="margin:0; font-size:0.8rem; color:#64748B;">나염·발주 엑셀 자동 가공 웹매크로</p>
                </div>
            </div>
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px; font-size:0.82rem; color:#1E40AF; margin-bottom:16px;">
                ℹ️ 파이썬 미설치 PC에서도 링크 접속 후 원본 <b>.xls</b> 파일을 끌어다 놓으면 <b>1초 만에 [가공완료].xlsx</b> 양식으로 자동 변환됩니다.
            </div>
        """, unsafe_allow_html=True)

        uploaded_files = st.file_uploader(
            "여기에 원본 .xls 파일을 끌어다 놓으세요 (다중 파일 일괄 변환 지원)",
            type=["xls"],
            accept_multiple_files=True
        )

        all_current_alerts = []

        if uploaded_files:
            processed_results = []
            for up_file in uploaded_files:
                try:
                    file_bytes = up_file.read()
                    res_bytes, item_cnt, total_q, alerts = process_single_file_bytes(file_bytes)
                    base_name = os.path.splitext(up_file.name)[0]
                    out_name = f"[가공완료]_{base_name}.xlsx"
                    
                    if alerts:
                        for al in alerts:
                            all_current_alerts.append({
                                "file": out_name,
                                **al
                            })

                    processed_results.append({
                        "filename": out_name,
                        "item_count": item_cnt,
                        "total_qty": total_q,
                        "data": res_bytes
                    })
                except Exception as e:
                    st.error(f"❌ {up_file.name} 변환 실패: {e}")

            st.session_state["all_alerts"] = all_current_alerts

            if processed_results:
                st.markdown("""
                <div style="display:flex; align-items:center; gap:6px; margin-top:16px; margin-bottom:8px;">
                    <span style="color:#2563EB; font-weight:bold;">✔️</span>
                    <span style="font-weight:700; color:#1E293B; font-size:0.95rem;">변환 완료 리스트</span>
                </div>
                """, unsafe_allow_html=True)

                for idx, item in enumerate(processed_results):
                    c_info, c_btn = st.columns([3.3, 1.3])
                    with c_info:
                        st.markdown(f"""
                        <div class="file-res-card">
                            <span style="font-size:1.6rem; color:#2563EB;">📄</span>
                            <div>
                                <div class="file-res-name">{item['filename']}</div>
                                <div class="file-res-desc">품목 {item['item_count']}건 / 총 입고수량 {item['total_qty']:,}장 (원재고취합 포함)</div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                    with c_btn:
                        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
                        st.download_button(
                            label="⬇️ 개별 다운로드",
                            data=item["data"],
                            file_name=item["filename"],
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_single_{idx}"
                        )

                if len(processed_results) > 1:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for item in processed_results:
                            zf.writestr(item["filename"], item["data"])
                    zip_buffer.seek(0)

                    st.markdown("<div class='zip-wrap' style='margin-top:14px;'>", unsafe_allow_html=True)
                    st.download_button(
                        label="📦 전체 가공완료 파일 ZIP 일괄 다운로드",
                        data=zip_buffer.getvalue(),
                        file_name="가공완료_나염발주_전체목록.zip",
                        mime="application/zip",
                        key="dl_zip_all"
                    )
                    st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.session_state["all_alerts"] = []

        st.markdown("</div>", unsafe_allow_html=True)

# --- 우측 열 (예시 행 복원 + 안내 태그 완벽 배치) ---
with col_right:
    if st.session_state.get("all_alerts"):
        alerts = st.session_state["all_alerts"]
        alert_rows_html = "".join([
            f"<li style='margin-top:4px;'><b>[{al['file']}]</b> 상품코드: <code>{al['item_code']}</code> | 옵션: <code>{al['option']}</code><br>"
            f"&nbsp;&nbsp;👉 <b>입고(생산)요청: {al['in_qty']}장</b> > <b>미배송: {al['undelivered']}장</b></li>"
            for al in alerts
        ])

        st.markdown(f"""
        <div class="admin-warn-box">
            <div style="display:flex; align-items:center; gap:8px; font-weight:800; font-size:1.05rem;">
                <span>🚨</span> <span>[긴급] 관리자 확인요청 (미배송 수량 초과 발생)</span>
            </div>
            <p style="margin:6px 0 8px 0; font-size:0.85rem; line-height:1.4;">
                미배송 수량보다 더 많은 수량을 생산하도록 전표가 생성되었습니다.<br>
                <b>관리자에게 즉시 확인 후 생산을 진행해야 합니다!</b>
            </p>
            <ul style="font-size:0.8rem; margin:0; padding-left:18px;">
                {alert_rows_html}
            </ul>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div style="background:#1F7A4D; color:white; padding:10px 16px; border-radius:10px 10px 0 0; font-weight:700; font-size:0.95rem; display:flex; align-items:center; gap:8px;">
        <span>📗</span> <span>변환 후 다운로드되는 엑셀 결과값 (가공리스트 시트)</span>
    </div>
    """, unsafe_allow_html=True)

    excel_preview_html = """
    <div style="background:#FFFFFF; border:1px solid #CBD5E1; border-top:none; border-radius:0 0 12px 12px; padding:16px;">
        <div style="background:#F1F5F9; border:1px solid #CBD5E1; padding:6px 12px; border-radius:6px; font-size:0.78rem; color:#475569; margin-bottom:12px; font-family:monospace;">
            fx | 2026-10-15 나염 발주 전표
        </div>
        <table class="excel-view-table">
            <thead>
                <tr>
                    <th colspan="6" style="background:#FFFFFF; font-size:1.15rem; font-weight:800; padding:12px;">
                        2026-10-15 나염 발주 전표
                    </th>
                </tr>
                <tr class="excel-head">
                    <th>공급처상품명</th>
                    <th>상품코드</th>
                    <th>옵션</th>
                    <th>입고수량</th>
                    <th>바코드</th>
                    <th>비고</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>오버핏 헤비 후드</td>
                    <td>100234</td>
                    <td>블랙-스폴로고-M</td>
                    <td>50</td>
                    <td>880912345678</td>
                    <td></td>
                </tr>
                <tr>
                    <td>오버핏 헤비 후드</td>
                    <td>100234</td>
                    <td>블랙-스폴로고-L</td>
                    <td>30</td>
                    <td>880912345679</td>
                    <td></td>
                </tr>
                <tr style="background:#F8FAFC;">
                    <td colspan="6" style="padding:6px;">
                        <span class="callout-tag-blue">⬅ 동일 상품코드 내 로고/컬러 변경 시 일반 빈 행 자동 삽입</span>
                    </td>
                </tr>
                <tr>
                    <td>오버핏 헤비 후드</td>
                    <td>100234</td>
                    <td>블랙-빅로고-L</td>
                    <td>40</td>
                    <td>880912345680</td>
                    <td></td>
                </tr>
                <tr class="row-yellow">
                    <td colspan="6" style="padding:6px;">
                        <span class="callout-tag-yellow">⬅ 상품코드가 변경될 때 연노란색(#FFF2CC) 구분 행 자동 삽입</span>
                    </td>
                </tr>
                <tr>
                    <td>피그먼트 맨투맨</td>
                    <td>100599</td>
                    <td>차콜-기본-FREE</td>
                    <td>80</td>
                    <td>880912345999</td>
                    <td></td>
                </tr>
                <tr class="row-red-alert">
                    <td>MUN053 맨투맨</td>
                    <td>ACUN053</td>
                    <td>블랙-3XL</td>
                    <td style="background:#FFC7CE; color:#9C0006; font-weight:bold;">5</td>
                    <td>ACUN053BK3XL</td>
                    <td style="background:#FFC7CE; color:#9C0006; font-weight:bold;">⚠️ 관리자확인요청</td>
                </tr>
                <tr style="background:#FEF2F2;">
                    <td colspan="6" style="padding:6px;">
                        <span class="callout-tag-red">⬅ 미배송 초과 건 발생 시 비고란에 붉은색 관리자확인요청 자동 표기</span>
                    </td>
                </tr>
                <tr style="font-weight:700; background:#F8FAFC;">
                    <td>합계</td>
                    <td></td>
                    <td></td>
                    <td>205</td>
                    <td></td>
                    <td></td>
                </tr>
            </tbody>
        </table>
        <div style="display:flex; gap:8px; margin-top:14px; border-top:1px solid #E2E8F0; padding-top:10px; font-size:0.75rem;">
            <div style="background:#FFFFFF; border:1px solid #CBD5E1; border-bottom:2px solid #16A34A; padding:4px 12px; font-weight:700; color:#15803D; border-radius:4px;">
                가공리스트 (시트1)
            </div>
            <div style="background:#F8FAFC; border:1px solid #E2E8F0; padding:4px 12px; color:#64748B; border-radius:4px;">
                원재고취합 (시트2)
            </div>
        </div>
    </div>
    """
    st.markdown(excel_preview_html, unsafe_allow_html=True)
