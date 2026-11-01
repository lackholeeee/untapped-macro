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

st.set_page_config(page_title="발주 전표 자동 가공기", layout="centered")

def clean_cell_value(val):
    if isinstance(val, float) and val.is_integer():
        return int(val)
    return val

def extract_base_option(opt_str):
    """
    나염 로고 표기를 제거하여 원단/무지 원재고 옵션으로 변환
    예: '[블랙(화이트로고)-S]' -> '[블랙-S]'
        '[그레이(네이비로고)-2XL]' -> '[그레이-2XL]'
    """
    opt_str = str(opt_str).strip()
    m = re.match(r'\[(.*)\]', opt_str)
    inner = m.group(1) if m else opt_str
    cleaned = re.sub(r'\(.*?\)', '', inner).strip()
    return f"[{cleaned}]"

def process_single_file_bytes(file_bytes):
    # 1. 원본 XLS 바이너리 데이터 읽기
    wb_target = xlrd.open_workbook(file_contents=file_bytes)
    sheet = wb_target.sheet_by_index(0)

    top_title = sheet.cell_value(0, 0) # 1행 전표명

    # 2. 데이터 추출 (10행 헤더 기준, 마지막 합계 행 제외)
    rows_data = []
    raw_inventory_items = []

    for r in range(10, sheet.nrows - 1):
        c0 = clean_cell_value(sheet.cell_value(r, 0)) # 공급처상품명
        c2 = clean_cell_value(sheet.cell_value(r, 2)) # 상품코드
        c3 = clean_cell_value(sheet.cell_value(r, 3)) # 옵션
        c5 = clean_cell_value(sheet.cell_value(r, 5)) # 요청수량
        c4 = clean_cell_value(sheet.cell_value(r, 4)) # 바코드

        rows_data.append([c0, c2, c3, c5, c4])
        
        # 2번 시트(원재고 취합)용 원천 데이터
        try:
            qty_num = int(float(c5))
        except (ValueError, TypeError):
            qty_num = 0
        raw_inventory_items.append((str(c0), str(c3), qty_num))

    # 3. 1번 시트용 상품코드 및 옵션 변경 구분 처리
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

    # 4. Openpyxl 엑셀 생성
    wb = openpyxl.Workbook()
    
    # ----------------------------------------------------
    # [시트 1] 기존 가공리스트 (기존 기능 100% 동일 유지)
    # ----------------------------------------------------
    ws1 = wb.active
    ws1.title = "가공리스트"
    ws1.views.sheetView[0].showGridLines = True

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 1~2행 병합 및 제목
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

    # ----------------------------------------------------
    # [시트 2] 신규 원재고 취합 수량 시트 (12333.xls 양식)
    # ----------------------------------------------------
    ws2 = wb.create_sheet(title="원재고취합")
    ws2.views.sheetView[0].showGridLines = True

    # 원재고 옵션별 수량 합산
    agg_dict = defaultdict(int)
    for prod_name, opt_raw, qty_num in raw_inventory_items:
        base_opt = extract_base_option(opt_raw)
        agg_dict[(prod_name, base_opt)] += qty_num

    # 옵션명 기준 정렬 (그레이 -> 네이비 -> 블랙 순)
    sorted_items = sorted(agg_dict.items(), key=lambda x: (x[0][0], x[0][1]))

    # 시트 2 헤더 (1행)
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
        # 옵션에서 색상 추출 (예: '[그레이-2XL]' -> '그레이')
        m_color = re.match(r'\[(.*?)-', base_opt)
        curr_color = m_color.group(1).strip() if m_color else base_opt

        # 색상이 변경될 때 구분 빈 행 삽입 (테두리 포함)
        if prev_color is not None and curr_color != prev_color:
            ws2.row_dimensions[curr_row2].height = 28
            for c_idx in range(1, 4):
                cell = ws2.cell(row=curr_row2, column=c_idx, value="")
                cell.border = border_all
            curr_row2 += 1

        # 데이터 입력 행
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

    # 시트 2 최하단 총합계 행
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

    # 시트 2 열 너비 자동 맞춤
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

    # 엑셀 파일 저장 및 바이너리 반환
    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer.getvalue()

# --- 웹 화면 구성 (Streamlit UI) ---
st.title("📦 발주 전표 자동 가공기")
st.markdown("가공할 **`.xls`** 파일을 드래그하여 올리면, **가공리스트(시트1)**와 **원재고취합(시트2)**이 포함된 **`.xlsx`** 파일로 즉시 변환됩니다.")

uploaded_files = st.file_uploader(
    "XLS 전표 파일을 선택하세요 (복수 선택 가능)", 
    type=["xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 변환 시작", type="primary"):
        with st.spinner("엑셀 가공 작업 중입니다..."):
            processed_results = []
            
            for up_file in uploaded_files:
                file_bytes = up_file.read()
                try:
                    res_bytes = process_single_file_bytes(file_bytes)
                    base_name = os.path.splitext(up_file.name)[0]
                    out_name = f"[가공완료]_{base_name}.xlsx"
                    processed_results.append((out_name, res_bytes))
                except Exception as e:
                    st.error(f"❌ {up_file.name} 변환 실패: {e}")

            if processed_results:
                st.success(f"총 {len(processed_results)}개 파일 변환이 완료되었습니다!")

                if len(processed_results) == 1:
                    out_name, res_bytes = processed_results[0]
                    st.download_button(
                        label=f"📥 {out_name} 다운로드",
                        data=res_bytes,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for out_name, res_bytes in processed_results:
                            zf.writestr(out_name, res_bytes)
                    zip_buffer.seek(0)

                    st.download_button(
                        label="📦 전체 가공 파일 ZIP 다운로드",
                        data=zip_buffer.getvalue(),
                        file_name="가공완료_엑셀목록.zip",
                        mime="application/zip"
                    )
