import io
import os
import zipfile
from typing import List
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, Response
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

app = FastAPI(title="UNTAPPED STUDIO 웹매크로")

def clean_cell_value(val):
    if isinstance(val, float) and val.is_integer():
        return int(val)
    return val

def process_xls_to_xlsx(file_bytes: bytes):
    wb_target = xlrd.open_workbook(file_contents=file_bytes)
    sheet = wb_target.sheet_by_index(0)

    top_title = sheet.cell_value(0, 0)

    rows_data = []
    for r in range(10, sheet.nrows - 1):
        c0 = clean_cell_value(sheet.cell_value(r, 0))
        c2 = clean_cell_value(sheet.cell_value(r, 2))
        c3 = clean_cell_value(sheet.cell_value(r, 3))
        c5 = clean_cell_value(sheet.cell_value(r, 5))
        c4 = clean_cell_value(sheet.cell_value(r, 4))
        rows_data.append([c0, c2, c3, c5, c4])

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
    ws = wb.active
    ws.title = "가공리스트"
    ws.views.sheetView[0].showGridLines = True

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 1~2행 전표 타이틀 병합
    ws.merge_cells("A1:E2")
    cell_a1 = ws["A1"]
    cell_a1.value = top_title
    cell_a1.font = Font(name="Malgun Gothic", size=16, bold=True)
    cell_a1.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(1, 3):
        ws.row_dimensions[r].height = 25
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = border_all

    # 3행 헤더
    headers = ['공급처상품명', '상품코드', '옵션', '요청수량', '바코드']
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    ws.row_dimensions[3].height = 30

    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c_idx, value=h)
        cell.font = Font(name="Malgun Gothic", size=13, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill
        cell.border = border_all

    item_change_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    curr_row = 4
    total_qty = 0
    item_count = len(rows_data)

    for r_data, r_type in grouped_rows:
        ws.row_dimensions[curr_row].height = 30
        if r_type == 'item_change':
            for c_idx in range(1, 6):
                cell = ws.cell(row=curr_row, column=c_idx, value="")
                cell.border = border_all
                cell.fill = item_change_fill
        elif r_type == 'logo_change':
            for c_idx in range(1, 6):
                cell = ws.cell(row=curr_row, column=c_idx, value="")
                cell.border = border_all
        else:
            for c_idx, val in enumerate(r_data, 1):
                cell = ws.cell(row=curr_row, column=c_idx, value=val)
                cell.font = Font(name="Malgun Gothic", size=13, bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
                cell.border = border_all

                if c_idx == 4:
                    try:
                        qty_val = int(val)
                        total_qty += qty_val
                        cell.value = qty_val
                        cell.number_format = '#,##0'
                    except (ValueError, TypeError):
                        pass
                elif c_idx in (2, 5) and isinstance(val, (int, float)):
                    cell.number_format = '@'

        curr_row += 1

    # 합계 행
    ws.row_dimensions[curr_row].height = 30
    cell_total = ws.cell(row=curr_row, column=1, value="합계")
    cell_total.font = Font(name="Malgun Gothic", size=13, bold=True)
    cell_total.alignment = Alignment(horizontal="center", vertical="center")
    cell_total.border = border_all

    for c_idx in range(2, 6):
        cell = ws.cell(row=curr_row, column=c_idx)
        cell.border = border_all
        if c_idx == 4:
            cell.value = total_qty
            cell.font = Font(name="Malgun Gothic", size=13, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = '#,##0'

    # 너비 자동 설정
    for col in ws.columns:
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
        ws.column_dimensions[col_letter].width = max(max_len * 1.3 + 4, 12)

    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer.getvalue(), item_count, total_qty

@app.get("/", response_class=HTMLResponse)
async def get_index():
    # 이미지 디자인을 그대로 살린 HTML / Tailwind UI
    return """
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>UNTAPPED STUDIO | 나염·발주 엑셀 자동 가공 웹매크로</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    </head>
    <body class="bg-[#F4F7FB] min-h-screen text-slate-800">
        <!-- 상단 헤더 -->
        <header class="bg-[#1A5CBA] text-white px-6 py-4 shadow-md flex justify-between items-center">
            <div class="flex items-center space-x-3">
                <i class="fa-solid fa-gear text-2xl"></i>
                <h1 class="text-xl font-bold tracking-wide">UNTAPPED STUDIO <span class="font-normal text-blue-200 text-lg">| 나염·발주 엑셀 자동 가공 웹매크로</span></h1>
            </div>
            <div class="bg-blue-600/60 border border-blue-300/40 rounded-full px-4 py-1.5 text-xs text-blue-100 flex items-center space-x-2">
                <i class="fa-solid fa-desktop"></i>
                <span>다른 PC에서도 동일하게 사용 가능</span>
            </div>
        </header>

        <main class="max-w-4xl mx-auto py-8 px-4">
            <!-- 알림 배너 -->
            <div class="bg-blue-50 border border-blue-200 text-blue-800 rounded-xl p-4 flex items-center space-x-3 mb-6 shadow-sm">
                <i class="fa-regular fa-lightbulb text-blue-600 text-xl"></i>
                <p class="text-sm font-medium">파이썬 미설치 PC에서도 링크 접속 후 원본 <span class="font-bold">.xls</span> 파일을 끌어다 놓으면 <span class="text-blue-600 font-bold">1초 만에 [가공완료].xlsx</span> 양식으로 자동 변환됩니다.</p>
            </div>

            <!-- 파일 드롭존 -->
            <div id="drop-zone" class="border-2 border-dashed border-blue-400 bg-white hover:bg-blue-50/50 transition-all rounded-2xl p-10 text-center cursor-pointer shadow-sm relative">
                <input type="file" id="file-input" multiple accept=".xls" class="hidden">
                <div class="flex flex-col items-center pointer-events-none">
                    <div class="w-16 h-16 bg-blue-100 text-blue-600 rounded-2xl flex items-center justify-center text-3xl mb-3 shadow-inner">
                        <i class="fa-solid fa-folder-open"></i>
                    </div>
                    <p class="text-lg font-bold text-blue-700 mb-1">여기에 원본 .xls 파일을 끌어다 놓으세요</p>
                    <p class="text-xs text-slate-400">(다중 파일 일괄 변환 지원)</p>
                </div>
            </div>

            <!-- 로딩 상태 -->
            <div id="loading" class="hidden my-8 text-center">
                <div class="inline-block animate-spin rounded-full h-8 w-8 border-4 border-blue-600 border-t-transparent"></div>
                <p class="text-sm text-slate-600 mt-2 font-medium">엑셀 파일을 가공하고 있습니다...</p>
            </div>

            <!-- 변환 완료 섹션 -->
            <div id="result-section" class="hidden mt-8">
                <div class="flex items-center space-x-2 mb-3">
                    <i class="fa-solid fa-circle-check text-blue-600 text-lg"></i>
                    <h2 class="font-bold text-slate-700">변환 완료 리스트</h2>
                </div>

                <!-- 파일 카드 목록 -->
                <div id="cards-container" class="space-y-3 mb-4"></div>

                <!-- 일괄 다운로드 버튼 -->
                <button id="zip-download-btn" class="w-full bg-[#1A73E8] hover:bg-blue-600 text-white font-bold py-3.5 px-4 rounded-xl shadow-md transition flex items-center justify-center space-x-2">
                    <i class="fa-solid fa-box-archive"></i>
                    <span>전체 가공완료 파일 ZIP 일괄 다운로드</span>
                </button>
            </div>
        </main>

        <script>
            const dropZone = document.getElementById('drop-zone');
            const fileInput = document.getElementById('file-input');
            const loading = document.getElementById('loading');
            const resultSection = document.getElementById('result-section');
            const cardsContainer = document.getElementById('cards-container');
            const zipBtn = document.getElementById('zip-download-btn');

            let processedFiles = [];

            dropZone.addEventListener('click', () => fileInput.click());

            ['dragover', 'dragenter'].forEach(evt => {
                dropZone.addEventListener(evt, (e) => {
                    e.preventDefault();
                    dropZone.classList.add('border-blue-600', 'bg-blue-50');
                });
            });

            ['dragleave', 'drop'].forEach(evt => {
                dropZone.addEventListener(evt, (e) => {
                    e.preventDefault();
                    dropZone.classList.remove('border-blue-600', 'bg-blue-50');
                });
            });

            dropZone.addEventListener('drop', (e) => {
                if (e.dataTransfer.files.length > 0) {
                    handleFiles(e.dataTransfer.files);
                }
            });

            fileInput.addEventListener('change', (e) => {
                if (e.target.files.length > 0) {
                    handleFiles(e.target.files);
                }
            });

            async function handleFiles(files) {
                const formData = new FormData();
                let validCount = 0;
                for (let file of files) {
                    if (file.name.endsWith('.xls')) {
                        formData.append('files', file);
                        validCount++;
                    }
                }

                if (validCount === 0) {
                    alert('.xls 확장자 파일만 업로드할 수 있습니다.');
                    return;
                }

                loading.classList.remove('hidden');
                resultSection.classList.add('hidden');
                cardsContainer.innerHTML = '';
                processedFiles = [];

                try {
                    const response = await fetch('/convert', {
                        method: 'POST',
                        body: formData
                    });

                    if (!response.ok) throw new Error('변환 실패');

                    const data = await response.json();
                    processedFiles = data.results;

                    renderResults(data.results);
                    resultSection.classList.remove('hidden');
                } catch (err) {
                    alert('오류가 발생했습니다: ' + err.message);
                } finally {
                    loading.classList.add('hidden');
                }
            }

            function renderResults(results) {
                cardsContainer.innerHTML = '';
                results.forEach((item, index) => {
                    const card = document.createElement('div');
                    card.className = "bg-white border border-slate-200 rounded-xl p-4 flex items-center justify-between shadow-sm hover:shadow transition";
                    card.innerHTML = `
                        <div class="flex items-center space-x-3.5">
                            <div class="text-blue-600 text-2xl">
                                <i class="fa-regular fa-file-excel"></i>
                            </div>
                            <div>
                                <p class="font-bold text-slate-800 text-sm leading-tight">${item.filename}</p>
                                <p class="text-xs text-slate-500 mt-1">품목 ${item.item_count}건 / 총 수량 ${item.total_qty.toLocaleString()}장</p>
                            </div>
                        </div>
                        <button onclick="downloadSingle(${index})" class="bg-[#1F7A4D] hover:bg-[#18633e] text-white text-xs font-semibold px-3.5 py-2 rounded-lg flex items-center space-x-1.5 transition">
                            <i class="fa-solid fa-arrow-down"></i>
                            <span>개별 다운로드</span>
                        </button>
                    `;
                    cardsContainer.appendChild(card);
                });
            }

            window.downloadSingle = function(idx) {
                const item = processedFiles[idx];
                const byteCharacters = atob(item.data);
                const byteNumbers = new Array(byteCharacters.length);
                for (let i = 0; i < byteCharacters.length; i++) {
                    byteNumbers[i] = byteCharacters.charCodeAt(i);
                }
                const byteArray = new Uint8Array(byteNumbers);
                const blob = new Blob([byteArray], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
                const link = document.createElement('a');
                link.href = window.URL.createObjectURL(blob);
                link.download = item.filename;
                link.click();
            };

            zipBtn.addEventListener('click', async () => {
                if (processedFiles.length === 0) return;
                const response = await fetch('/download-zip', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({files: processedFiles})
                });
                const blob = await response.blob();
                const link = document.createElement('a');
                link.href = window.URL.createObjectURL(blob);
                link.download = "가공완료_나염발주_일괄.zip";
                link.click();
            });
        </script>
    </body>
    </html>
    """

import base64
from pydantic import BaseModel

class FileItem(BaseModel):
    filename: str
    data: str

class ZipRequest(BaseModel):
    files: List[FileItem]

@app.post("/convert")
async def convert_files(files: List[UploadFile] = File(...)):
    results = []
    for up_file in files:
        file_bytes = await up_file.read()
        try:
            out_bytes, item_cnt, total_q = process_xls_to_xlsx(file_bytes)
            base_name = os.path.splitext(up_file.filename)[0]
            out_name = f"[가공완료]_{base_name}.xlsx"
            results.append({
                "filename": out_name,
                "item_count": item_cnt,
                "total_qty": total_q,
                "data": base64.b64encode(out_bytes).decode('utf-8')
            })
        except Exception as e:
            print(f"변환 실패 {up_file.filename}: {e}")
    return {"results": results}

@app.post("/download-zip")
async def download_zip(req: ZipRequest):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in req.files:
            file_bytes = base64.b64decode(f.data)
            zf.writestr(f.filename, file_bytes)
    zip_buffer.seek(0)
    return Response(
        content=zip_buffer.getvalue(),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=converted_files.zip"}
    )

if __name__ == "__main__":
    import uvicorn
    # 0.0.0.0으로 열어 같은 사무실/네트워크의 다른 PC에서도 접속 가능하도록 설정
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
