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

# 화면 전체 폭(wide) 사용 설정
st.set_page_config(
    page_title="UNTAPPED STUDIO | 나염·발주 엑셀 자동 가공 웹매크로",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 전체 화면 템플릿 CSS ---
st.markdown("""
<style>
    /* 전체 여백 초기화 및 와이드 레이아웃 */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        max-width: 98% !important;
    }
    body, [data-testid="stAppViewContainer"] {
        background-color: #EBF3FC !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }
    #MainMenu, footer, header {visibility: hidden;}

    /* 최상단 네이비 헤더 바 */
    .main-top-header {
        background-color: #1558B0;
        color: #ffffff;
        padding: 16px 28px;
        border-radius: 14px;
        box-shadow: 0 4px 12px rgba(21, 88, 176, 0.25);
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 14px;
    }
    .main-top-header h1 {
        font-size: 1.45rem !important;
        font-weight: 800 !important;
        margin: 0 !important;
        color: #ffffff !important;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .main-top-header h1 span {
        font-weight: 500;
        font-size: 1.25rem;
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

    /* 전폭 상단 알림 배너 */
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

    /* 윈도우 창 공통 프레임 */
    .window-card {
        background: #ffffff;
        border: 1px solid #CBD5E1;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.06);
        height: 100%;
    }
    .window-header {
        background: #E2E8F0;
        padding: 10px 16px;
        border-bottom: 1px solid #CBD5E1;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .traffic-lights {
        display: flex;
        gap: 6px;
    }
    .traffic-dot {
        width: 10px;
        height: 10px;
        border-radius: 50%;
    }
    .dot-red { background: #EF4444; }
    .dot-yellow { background: #F59E0B; }
    .dot-green { background: #10B981; }

    .window-title-tab {
        background: #ffffff;
        padding: 4px 14px;
        border-radius: 6px 6px 0 0;
        font-size: 0.78rem;
        font-weight: 600;
        color: #334155;
    }
    .url-bar {
        background: #F1F5F9;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        padding: 4px 12px;
        font-size: 0.78rem;
        color: #475569;
        flex-grow: 1;
    }

    /* 변환 완료 결과 카드 */
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

    /* 개별 다운로드 버튼 (초록색) */
    div[data-testid="stDownloadButton"] button {
        background-color: #1F7A4D !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-size: 0.78rem !important;
        font-weight: 600 !important;
        padding: 6px 12px !important;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #18633E !important;
    }

    /* ZIP 일괄 다운로드 버튼 (파란색) */
    .zip-wrap div[data-testid="stDownloadButton"] button {
        background-color: #1A73E8 !important;
        width: 100% !important;
        padding: 12px !important;
        font-size: 0.92rem !important;
        border-radius: 8px !important;
    }

    /* 우측 엑셀 설명서 테이블 */
    .excel-view-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.78rem;
        text-align: center;
    }
    .excel-view-table th, .excel-view-table td {
        border: 1px solid #CBD5E1;
        padding: 6px 4px;
        white-space: nowrap;
    }
    .excel-head {
        background-color: #F1F5F9;
        font-weight: 700;
        color: #1E293B;
    }
    .row-yellow {
        background-color: #FFF2CC !important;
    }
    .callout-tag-blue {
        background: #2563EB;
        color: white;
        padding: 2px 8px;
        border-radius: 9999px;
        font-size: 0.72rem;
        font-weight: 600;
        display: inline-block;
    }
    .callout-tag-yellow {
        background: #D97706;
        color: white;
        padding: 2px 8px;
        border-radius: 9999px;
        font-size: 0.72rem;
        font-weight: 600;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# 1. 상단 타이틀 헤더 바
st.markdown("""
<div class="main-top-header">
    <h1>⚙️ UNTAPPED STUDIO <span>| 나염·발주 엑셀 자동 가공 웹매크로</span></h1>
    <div class="pc-badge">💻 다른 PC에서도 동일하게 사용 가능</div>
</div>
<div class="notice-banner">
    <span style="font-size: 1.2rem;">💡</span>
    <span>파이썬 미설치 PC에서도 링크 접속 후 원본 <b>.xls</b> 파일을 끌어다 놓으면 <b>1초 만에 [가공완료].xlsx</b> 양식으로 자동 변환됩니다.</span>
</div>
""", unsafe_allow_html=True)

# --- 엑셀 변환 로직 (시트1 가공리스트 + 시트2 원재고취합) ---
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

    # 시트 1용 그룹화
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

    # ----------------------------------------------------
    # [시트 1] 가공리스트
    # ----------------------------------------------------
    ws1 = wb.active
    ws1.title = "가공리스트"
    ws1.views.sheetView[0].showGridLines = True

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws1.merge_cells("A1:E2")
    cell_a1 = ws1["A1"]
    cell_a1.value = top_title
    cell_a1.font = Font(name="Malgun Gothic", size=16, bold=True)
    cell_a1.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(1, 3):
        ws1.row_dimensions[r].height = 25
        for c in range(1, 6):
            ws1.cell(row=r, column=c).border = border_all

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

    # 시트 1 합계
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
    # [시트 2] 원재고취합
    # ----------------------------------------------------
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

# =========================================================================
# 화면 레이아웃 (좌측: 웹매크로 도구 / 우측: 설명서 및 엑셀 결과값 안내)
# =========================================================================
col_left, col_right = st.columns([1, 1], gap="medium")

# --- [좌측 열] 타 PC 크롬 접속 시 보이는 웹매크로 화면 ---
with col_left:
    st.markdown("""
    <div style="background:#0F52BA; color:white; padding:8px 16px; border-radius:8px 8px 0 0; font-weight:700; font-size:0.92rem; display:flex; align-items:center; gap:8px;">
        <span>💻</span> <span>타 PC 크롬 접속 시 보이는 웹매크로 화면</span>
    </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown("""
        <div style="background:#FFFFFF; border:1px solid #CBD5E1; border-top:none; border-radius:0 0 12px 12px; padding:18px;">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:12px;">
                <span style="font-size:1.3rem;">⚙️</span>
                <div>
                    <h3 style="margin:0; font-size:1.05rem; font-weight:800; color:#1E293B;">UNTAPPED STUDIO</h3>
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
                    st.error(f"❌ {up_file.name} 변환 실패: {e}")

            if processed_results:
                st.markdown("""
                <div style="display:flex; align-items:center; gap:6px; margin-top:16px; margin-bottom:8px;">
                    <span style="color:#2563EB; font-weight:bold;">✔️</span>
                    <span style="font-weight:700; color:#1E293B; font-size:0.92rem;">변환 완료 리스트</span>
                </div>
                """, unsafe_allow_html=True)

                for idx, item in enumerate(processed_results):
                    c_info, c_btn = st.columns([3.2, 1.2])
                    with c_info:
                        st.markdown(f"""
                        <div class="file-res-card">
                            <span style="font-size:1.6rem; color:#2563EB;">📄</span>
                            <div>
                                <div class="file-res-name">{item['filename']}</div>
                                <div class="file-res-desc">품목 {item['item_count']}건 / 총 수량 {item['total_qty']:,}장 (원재고취합 포함)</div>
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

                    st.markdown("<div class='zip-wrap' style='margin-top:12px;'>", unsafe_allow_html=True)
                    st.download_button(
                        label="📦 전체 가공완료 파일 ZIP 일괄 다운로드",
                        data=zip_buffer.getvalue(),
                        file_name="가공완료_나염발주_전체목록.zip",
                        mime="application/zip",
                        key="dl_zip_all"
                    )
                    st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

# --- [우측 열] 변환 후 다운로드되는 엑셀 결과값 (설명서 및 가이드) ---
with col_right:
    st.markdown("""
    <div style="background:#1F7A4D; color:white; padding:8px 16px; border-radius:8px 8px 0 0; font-weight:700; font-size:0.92rem; display:flex; align-items:center; gap:8px;">
        <span>📗</span> <span>변환 후 다운로드되는 엑셀 결과값 (가공리스트 시트)</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="background:#FFFFFF; border:1px solid #CBD5E1; border-top:none; border-radius:0 0 12px 12px; padding:16px;">
        <div style="background:#F1F5F9; border:1px solid #E2E8F0; padding:6px 12px; border-radius:6px; font-size:0.75rem; color:#475569; margin-bottom:12px; font-family:monospace;">
            fx | 2026-10-15 나염 발주 전표
        </div>

        <table class="excel-view-table">
            <thead>
                <tr>
                    <th colspan="5" style="background:#FFFFFF; font-size:1.15rem; font-weight:800; padding:12px;">
                        2026-10-15 나염 발주 전표
                    </th>
                </tr>
                <tr class="excel-head">
                    <th>공급처상품명</th>
                    <th>상품코드</th>
                    <th>옵션</th>
                    <th>요청수량</th>
                    <th>바코드</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>오버핏 헤비 후드</td>
                    <td>100234</td>
                    <td>블랙-스폴로고-M</td>
                    <td>50</td>
                    <td>880912345678</td>
                </tr>
                <tr>
                    <td>오버핏 헤비 후드</td>
                    <td>100234</td>
                    <td>블랙-스폴로고-L</td>
                    <td>30</td>
                    <td>880912345679</td>
                </tr>
                <tr style="background:#F8FAFC;">
                    <td colspan="5" style="padding:6px;">
                        <span class="callout-tag-blue">⬅ 동일 상품코드 내 로고/컬러 변경 시 일반 빈 행 자동 삽입</span>
                    </td>
                </tr>
                <tr>
                    <td>오버핏 헤비 후드</td>
                    <td>100234</td>
                    <td>블랙-빅로고-L</td>
                    <td>40</td>
                    <td>880912345680</td>
                </tr>
                <tr class="row-yellow">
                    <td colspan="5" style="padding:6px;">
                        <span class="callout-tag-yellow">⬅ 상품코드가 변경될 때 연노란색(#FFF2CC) 구분 행 자동 삽입</span>
                    </td>
                </tr>
                <tr>
                    <td>피그먼트 맨투맨</td>
                    <td>100599</td>
                    <td>차콜-기본-FREE</td>
                    <td>80</td>
                    <td>880912345999</td>
                </tr>
                <tr style="font-weight:700; background:#F8FAFC;">
                    <td>합계</td>
                    <td></td>
                    <td></td>
                    <td>200</td>
                    <td></td>
                </tr>
            </tbody>
        </table>

        <!-- 하단 엑셀 시트 탭 설명 -->
        <div style="display:flex; gap:8px; margin-top:14px; border-top:1px solid #E2E8F0; padding-top:10px; font-size:0.75rem;">
            <div style="background:#FFFFFF; border:1px solid #CBD5E1; border-bottom:2px solid #16A34A; padding:4px 12px; font-weight:700; color:#15803D; border-radius:4px;">
                가공리스트 (시트1)
            </div>
            <div style="background:#F8FAFC; border:1px solid #E2E8F0; padding:4px 12px; color:#64748B; border-radius:4px;">
                원재고취합 (시트2)
            </div>
            <div style="margin-left:auto; font-size:0.72rem; color:#94A3B8; align-self:center;">
                * 엑셀 파일 다운로드 시 시트1과 시트2가 모두 포함되어 있습니다.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
