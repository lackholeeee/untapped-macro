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

# 페이지 레이아웃 및 브라우저 탭 설정
st.set_page_config(
    page_title="UNTAPPED STUDIO | 나염·발주 엑셀 자동 가공 웹매크로",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- UNTAPPED STUDIO 맞춤형 UI 스타일링 (CSS) ---
st.markdown("""
<style>
    /* 기본 여백 및 배경색 지정 */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 3rem !important;
        max-width: 900px !important;
    }
    body, [data-testid="stAppViewContainer"] {
        background-color: #F4F7FB !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }
    #MainMenu, footer, header {visibility: hidden;}

    /* 상단 UNTAPPED STUDIO 네이비 헤더 바 */
    .custom-header {
        background-color: #1A5CBA;
        color: #ffffff;
        padding: 16px 24px;
        border-radius: 14px;
        box-shadow: 0 4px 14px rgba(26, 92, 186, 0.25);
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 20px;
    }
    .custom-header h1 {
        font-size: 1.25rem !important;
        font-weight: 800 !important;
        margin: 0 !important;
        color: #ffffff !important;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .custom-header h1 span {
        font-weight: 400;
        font-size: 1.05rem;
        color: #BFDBFE;
    }
    .custom-badge {
        background: rgba(255, 255, 255, 0.18);
        border: 1px solid rgba(255, 255, 255, 0.35);
        border-radius: 9999px;
        padding: 5px 14px;
        font-size: 0.75rem;
        color: #EFF6FF;
        font-weight: 500;
    }

    /* 알림 배너 박스 */
    .notice-box {
        background-color: #EFF6FF;
        border: 1px solid #BFDBFE;
        border-radius: 12px;
        padding: 14px 18px;
        color: #1E40AF;
        font-size: 0.88rem;
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 24px;
    }

    /* 파일 카드 스타일 */
    .file-card {
        background-color: #ffffff;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 14px 20px;
        display: flex;
        align-items: center;
        gap: 16px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        margin-bottom: 10px;
    }
    .file-card .icon-box {
        font-size: 1.8rem;
        color: #1A5CBA;
    }
    .file-card .file-info {
        flex-grow: 1;
    }
    .file-card .file-name {
        font-weight: 700;
        font-size: 0.92rem;
        color: #1E293B;
    }
    .file-card .file-desc {
        font-size: 0.78rem;
        color: #64748B;
        margin-top: 3px;
    }

    /* 다운로드 버튼 녹색/파란색 커스텀 */
    div[data-testid="stDownloadButton"] button {
        background-color: #1F7A4D !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        font-size: 0.82rem !important;
        font-weight: 600 !important;
        padding: 6px 14px !important;
        box-shadow: 0 2px 4px rgba(31, 122, 77, 0.2) !important;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #18633E !important;
    }

    /* 전체 ZIP 다운로드 버튼 스타일 */
    .zip-btn-wrap div[data-testid="stDownloadButton"] button {
        background-color: #1A73E8 !important;
        width: 100% !important;
        padding: 12px !important;
        font-size: 0.95rem !important;
        border-radius: 10px !important;
        box-shadow: 0 3px 8px rgba(26, 115, 232, 0.25) !important;
    }
    .zip-btn-wrap div[data-testid="stDownloadButton"] button:hover {
        background-color: #1557B0 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 템플릿 상단 헤더 & 배너 렌더링 ---
st.markdown("""
<div class="custom-header">
    <h1>⚙️ UNTAPPED STUDIO <span>| 나염·발주 엑셀 자동 가공 웹매크로</span></h1>
    <div class="custom-badge">💻 다른 PC에서도 동일하게 사용 가능</div>
</div>
<div class="notice-box">
    <span>💡</span>
    <span>파이썬 미설치 PC에서도 링크 접속 후 원본 <b>.xls</b> 파일을 끌어다 놓으면 <b>1초 만에 [가공완료].xlsx</b> 양식으로 자동 변환됩니다.</span>
</div>
""", unsafe_allow_html=True)

# --- 엑셀 변환 핵심 로직 ---
def clean_cell_value(val):
    if isinstance(val, float) and val.is_integer():
        return int(val)
    return val

def extract_base_option(opt_str):
    opt_str = str(opt_str).strip()
    m = re.match(r'\[(.*)\]', opt_str)
    inner = m.group(1) if m else opt_str
    cleaned = re.sub(r'\(.*?\)', '', inner).strip()
    return f"[{cleaned}]"

def process_single_file_bytes(file_bytes):
    # 원본 XLS 읽기
    wb_target = xlrd.open_workbook(file_contents=file_bytes)
    sheet = wb_target.sheet_by_index(0)
    top_title = sheet.cell_value(0, 0)

    rows_data = []
    raw_inventory_items = []

    for r in range(10, sheet.nrows - 1):
        c0 = clean_cell_value(sheet.cell_value(r, 0)) # 공급처상품명
        c2 = clean_cell_value(sheet.cell_value(r, 2)) # 상품코드
        c3 = clean_cell_value(sheet.cell_value(r, 3)) # 옵션
        c5 = clean_cell_value(sheet.cell_value(r, 5)) # 요청수량
        c4 = clean_cell_value(sheet.cell_value(r, 4)) # 바코드

        rows_data.append([c0, c2, c3, c5, c4])

        try:
            qty_num = int(float(c5))
        except (ValueError, TypeError):
            qty_num = 0
        raw_inventory_items.append((str(c0), str(c3), qty_num))

    # 구분 행 분류
    grouped_rows = []
    prev_item_code = None
    prev_logo_part = None

    for row in rows_data:
        item_code = str(row[1])
        opt_str = str(row[2])
        logo_part = opt_str.rsplit('-', 1)[0] if '-' in opt_str else opt_str

        if prev_item_code is not None:
            if item_code != prev_item_code:
                grouped_rows.append((['', '', '', '', ''], 'item_change'))
            elif logo_part != prev_logo_part:
                grouped_rows.append((['', '', '', '', ''], 'logo_change'))

        grouped_rows.append((row, 'data'))
        prev_item_code = item_code
        prev_logo_part = logo_part

    wb = openpyxl.Workbook()

    # -------------------------------------------------------------------------
    # [시트 1] 가공리스트 (기존 서식 100% 동일 유지)
    # -------------------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "가공리스트"
    ws1.views.sheetView[0].showGridLines = True

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 1~2행 전표 타이틀 병합
    ws1.merge_cells("A1:E2")
    cell_a1 = ws1["A1"]
    cell_a1.value = top_title
    cell_a1.font = Font(name="Malgun Gothic", size=16, bold=True)
    cell_a1.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(1, 3):
        ws1.row_dimensions[r].height = 25
        for c in range(1, 6):
            ws1.cell(row=r, column=c).border = border_all

    # 3행 헤더
    headers1 = ['공급처상품명', '상품코드', '옵션', '요청수량', '바코드']
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    ws1.row_dimensions[3].height = 30

    for c_idx, h in enumerate(headers1, 1):
        cell = ws1.cell(row=3, column=c_idx, value=h)
        cell.font = Font(name="Malgun Gothic", size=13, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill
        cell.border = border_all

    item_change_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    curr_row1 = 4
    total_qty1 = 0

    for r_data, r_type in grouped_rows:
        ws1.row_dimensions[curr_row1].height = 30
        if r_type == 'item_change':
            for c_idx in range(1, 6):
                cell = ws1.cell(row=curr_row1, column=c_idx, value="")
                cell.border = border_all
                cell.fill = item_change_fill
        elif r_type == 'logo_change':
            for c_idx in range(1, 6):
                cell = ws1.cell(row=curr_row1, column=c_idx, value="")
                cell.border = border_all
        else:
            for c_idx, val in enumerate(r_data, 1):
                cell = ws1.cell(row=curr_row1, column=c_idx, value=val)
                cell.font = Font(name="Malgun Gothic", size=13, bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
                cell.border = border_all

                if c_idx == 4:
                    try:
                        qty_val = int(val)
                        total_qty1 += qty_val
                        cell.value = qty_val
                        cell.number_format = '#,##0'
                    except (ValueError, TypeError):
                        pass
                elif c_idx in (2, 5) and isinstance(val, (int, float)):
                    cell.number_format = '@'

        curr_row1 += 1

    # 시트 1 합계 행
    ws1.row_dimensions[curr_row1].height = 30
    cell_total1 = ws1.cell(row=curr_row1, column=1, value="합계")
    cell_total1.font = Font(name="Malgun Gothic", size=13, bold=True)
    cell_total1.alignment = Alignment(horizontal="center", vertical="center")
    cell_total1.border = border_all

    for c_idx in range(2, 6):
        cell = ws1.cell(row=curr_row1, column=c_idx)
        cell.border = border_all
        if c_idx == 4:
            cell.value = total_qty1
            cell.font = Font(name="Malgun Gothic", size=13, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = '#,##0'

    # 시트 1 열 너비 자동 맞춤
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

    # -------------------------------------------------------------------------
    # [시트 2] 원재고취합 (12333.xls 양식)
    # -------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="원재고취합")
    ws2.views.sheetView[0].showGridLines = True

    agg_dict = defaultdict(int)
    for prod_name, opt_raw, qty_num in raw_inventory_items:
        base_opt = extract_base_option(opt_raw)
        agg_dict[(prod_name, base_opt)] += qty_num

    sorted_items = sorted(agg_dict.items(), key=lambda x: (x[0][0], x[0][1]))

    headers2 = ['공급처상품명', '옵션', '수량']
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

    # 시트 2 총합계 행
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
    return out_buffer.getvalue(), len(rows_data), total_qty1

# --- 파일 업로더 ---
uploaded_files = st.file_uploader(
    "여기에 원본 .xls 파일을 끌어다 놓으세요 (다중 파일 일괄 변환 지원)",
    type=["xls"],
    accept_multiple_files=True,
    help="여러 개의 원본 발주 .xls 파일을 동시에 등록할 수 있습니다."
)

# --- 변환 완료 카드 리스트 렌더링 ---
if uploaded_files:
    processed_results = []
    for up_file in uploaded_files:
        try:
            file_bytes = up_file.read()
            res_bytes, item_cnt, total_q = process_single_file_bytes(file_bytes)
            base_name = os.path.splitext(up_file.name)[0]
            out_name = f"[가공완료]_{base_name}.xlsx"
            processed_results.append({
                "filename": out_name,
                "item_count": item_cnt,
                "total_qty": total_q,
                "data": res_bytes
            })
        except Exception as e:
            st.error(f"❌ {up_file.name} 변환 중 오류: {e}")

    if processed_results:
        st.markdown("""
        <div style="display:flex; align-items:center; gap:8px; margin-top:24px; margin-bottom:12px;">
            <span style="color:#1A5CBA; font-size:1.15rem;">✔️</span>
            <span style="font-weight:700; color:#1E293B; font-size:1.05rem;">변환 완료 리스트</span>
        </div>
        """, unsafe_allow_html=True)

        # 각 파일별 요약 카드 및 개별 다운로드 버튼
        for idx, item in enumerate(processed_results):
            col_info, col_btn = st.columns([4, 1])
            with col_info:
                st.markdown(f"""
                <div class="file-card">
                    <div class="icon-box">📄</div>
                    <div class="file-info">
                        <div class="file-name">{item['filename']}</div>
                        <div class="file-desc">품목 {item['item_count']}건 / 총 수량 {item['total_qty']:,}장 (가공리스트 + 원재고취합 포함)</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with col_btn:
                st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
                st.download_button(
                    label="📥 개별 다운로드",
                    data=item["data"],
                    file_name=item["filename"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_single_{idx}"
                )

        # 전체 ZIP 일괄 다운로드 버튼
        if len(processed_results) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for item in processed_results:
                    zf.writestr(item["filename"], item["data"])
            zip_buffer.seek(0)

            st.markdown("<div class='zip-btn-wrap' style='margin-top:16px;'>", unsafe_allow_html=True)
            st.download_button(
                label="📦 전체 가공완료 파일 ZIP 일괄 다운로드",
                data=zip_buffer.getvalue(),
                file_name="가공완료_나염발주_전체목록.zip",
                mime="application/zip",
                key="dl_zip_all"
            )
            st.markdown("</div>", unsafe_allow_html=True)
