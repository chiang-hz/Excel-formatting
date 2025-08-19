import os
import win32com.client as win32
from flask import (Flask, request, render_template, send_from_directory, 
                   session, redirect, url_for, make_response)
from werkzeug.utils import secure_filename
import pythoncom
from win32com.client import gencache

# ============================================================================== 
# 應用程式設定
# ==============================================================================
app = Flask(__name__)
app.secret_key = 'your-very-secret-key-for-excel-formatter'
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
DOWNLOAD_FOLDER = os.path.join(BASE_DIR, 'downloads')
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# ============================================================================== 
# 核心處理函式 (無變動)
# ==============================================================================
def apply_format_to_file(template_path, target_path, source_sheet_name, selected_target_sheets):
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        wb_template = excel.Workbooks.Open(template_path)
        wb_target = excel.Workbooks.Open(target_path)
        try:
            ws_source = wb_template.Worksheets(source_sheet_name)
        except Exception:
            ws_source = wb_template.Worksheets(1)
        for sheet_name in selected_target_sheets:
            try:
                ws_target = wb_target.Worksheets(sheet_name)
            except Exception:
                print(f"警告：在目標檔案中找不到名為 '{sheet_name}' 的工作表，已跳過。")
                continue
            print(f"--- 開始處理工作表: {ws_target.Name} ---")
            print("步驟一: 提取並清理原始資料至記憶體...")
            original_data = []
            if ws_target.UsedRange.Rows.Count > 1 or ws_target.UsedRange.Columns.Count > 1 or ws_target.UsedRange.Value is not None :
                original_data_raw = ws_target.UsedRange.Value
                if isinstance(original_data_raw, tuple):
                    for row in original_data_raw:
                        processed_row = []
                        row_iterable = row if isinstance(row, tuple) else (row,)
                        for cell_value in row_iterable:
                            processed_row.append(cell_value)
                        original_data.append(processed_row)
                else:
                    cell_value = original_data_raw
                    original_data.append([cell_value])
            print("步驟一: 完成。")
            print("步驟二: 全面複製範本格式...")
            ws_target.PageSetup.Orientation = ws_source.PageSetup.Orientation
            ws_target.PageSetup.PaperSize = ws_source.PageSetup.PaperSize
            ws_target.PageSetup.TopMargin = ws_source.PageSetup.TopMargin
            ws_target.PageSetup.BottomMargin = ws_source.PageSetup.BottomMargin
            ws_target.PageSetup.LeftMargin = ws_source.PageSetup.LeftMargin
            ws_target.PageSetup.RightMargin = ws_source.PageSetup.RightMargin
            ws_target.PageSetup.HeaderMargin = ws_source.PageSetup.HeaderMargin
            ws_target.PageSetup.FooterMargin = ws_source.PageSetup.FooterMargin
            ws_target.PageSetup.LeftHeader = ws_source.PageSetup.LeftHeader
            ws_target.PageSetup.CenterHeader = ws_source.PageSetup.CenterHeader
            ws_target.PageSetup.RightHeader = ws_source.PageSetup.RightHeader
            ws_target.PageSetup.LeftFooter = ws_source.PageSetup.LeftFooter
            ws_target.PageSetup.CenterFooter = ws_source.PageSetup.CenterFooter
            ws_target.PageSetup.RightFooter = ws_source.PageSetup.RightFooter
            ws_target.PageSetup.OddAndEvenPagesHeaderFooter = ws_source.PageSetup.OddAndEvenPagesHeaderFooter
            ws_target.PageSetup.DifferentFirstPageHeaderFooter = ws_source.PageSetup.DifferentFirstPageHeaderFooter
            ws_target.PageSetup.ScaleWithDocHeaderFooter = ws_source.PageSetup.ScaleWithDocHeaderFooter
            ws_target.PageSetup.AlignMarginsHeaderFooter = ws_source.PageSetup.AlignMarginsHeaderFooter
            ws_target.PageSetup.Zoom = ws_source.PageSetup.Zoom
            ws_target.PageSetup.FitToPagesWide = ws_source.PageSetup.FitToPagesWide
            ws_target.PageSetup.FitToPagesTall = ws_source.PageSetup.FitToPagesTall
            print("正在複製列高...")
            for i in range(1, ws_source.UsedRange.Rows.Count + 1):
                if i <= ws_target.Rows.Count:
                    ws_target.Rows(i).RowHeight = ws_source.Rows(i).RowHeight
            print("列高複製完成。")
            ws_source.UsedRange.Copy()
            ws_target.Activate()
            ws_target.Range("A1").PasteSpecial(win32.constants.xlPasteAllUsingSourceTheme)
            ws_target.Range("A1").PasteSpecial(win32.constants.xlPasteColumnWidths)
            excel.CutCopyMode = False
            ws_target.PageSetup.PrintArea = ""
            ws_target.ResetAllPageBreaks()
            print("步驟二: 完成。")
            print("步驟三: 清空目標工作表所有內容 (保留格式)...")
            ws_target.Cells.ClearContents()
            print("步驟三: 完成。")
            print("步驟四: 將乾淨資料回填至工作表...")
            if original_data:
                max_cols = max(len(row) for row in original_data) if original_data else 0
                final_data = [(row + [None] * (max_cols - len(row))) for row in original_data]
                start_cell = ws_target.Cells(1, 1)
                end_cell = ws_target.Cells(len(final_data), max_cols if max_cols > 0 else 1)
                write_range = ws_target.Range(start_cell, end_cell)
                write_range.Value = final_data
            print("步驟四: 完成。")
            print(f"--- 工作表 {ws_target.Name} 處理完畢 ---")
        
        # 這裡仍然使用安全檔名來儲存
        safe_original_filename = os.path.basename(target_path)
        filename_root, filename_ext = os.path.splitext(safe_original_filename)
        output_filename = f"{filename_root}_formatted{filename_ext}"
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
        wb_target.SaveAs(output_path)
        wb_template.Close(SaveChanges=False)
        wb_target.Close(SaveChanges=False)
        return output_filename # 返回安全檔名
    except Exception as e:
        print(f"處理 Excel 時發生錯誤: {e}")
        return None
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()

# ... (get_sheet_names 和 allowed_file 函式不變) ...
def get_sheet_names(file_path):
    pythoncom.CoInitialize()
    excel = None
    sheet_names = []
    try:
        excel = gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        workbook = excel.Workbooks.Open(file_path)
        sheet_names = [sheet.Name for sheet in workbook.Worksheets]
        workbook.Close(SaveChanges=False)
    except Exception as e:
        print(f"讀取工作表名稱時發生錯誤: {e}")
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()
    return sheet_names

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ============================================================================== 
# Flask 路由
# ==============================================================================
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload_template', methods=['POST'])
def upload_template():
    if 'template_file' not in request.files:
        return "錯誤：請求中沒有檔案", 400
    file = request.files['template_file']
    if file.filename == '' or not allowed_file(file.filename):
        return "錯誤：未選擇檔案或檔案類型不支援", 400
    original_filename = file.filename
    safe_filename = secure_filename(original_filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
    file.save(filepath)
    session['template_path'] = filepath
    session['template_filename_original'] = original_filename
    sheet_names = get_sheet_names(filepath)
    if not sheet_names:
        return "錯誤：無法讀取此 Excel 檔案的工作表", 500
    return render_template('select_format.html', sheet_names=sheet_names, template_filename=original_filename)

@app.route('/select_target_sheets', methods=['POST'])
def select_target_sheets():
    if 'template_path' not in session:
        return "錯誤：Session 遺失，請返回首頁重新開始。", 400
    if 'target_file' not in request.files:
        return "錯誤：請求中沒有目標檔案", 400
    file = request.files['target_file']
    if file.filename == '' or not allowed_file(file.filename):
        return "錯誤：未選擇目標檔案或檔案類型不支援", 400
    original_filename = file.filename
    safe_filename = secure_filename(original_filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
    file.save(filepath)
    session['target_path'] = filepath
    session['target_filename_original'] = original_filename
    session['source_sheet_name'] = request.form.get('source_sheet')
    target_sheet_names = get_sheet_names(filepath)
    if not target_sheet_names:
        return "錯誤：無法讀取目標檔案的工作表", 500
    return render_template('confirm_apply.html', 
                           target_sheet_names=target_sheet_names,
                           template_filename=session.get('template_filename_original'),
                           target_filename=session.get('target_filename_original'),
                           source_sheet_name=session.get('source_sheet_name'))

@app.route('/apply_format', methods=['POST'])
def apply_format():
    template_path = session.get('template_path')
    target_path = session.get('target_path')
    source_sheet_name = session.get('source_sheet_name')
    # 【修改點】一併獲取原始檔名
    target_filename_original = session.get('target_filename_original')

    if not all([template_path, target_path, source_sheet_name, target_filename_original]):
        return "錯誤：Session 資訊不完整，請返回首頁重新開始。", 400

    selected_sheets = request.form.getlist('target_sheets')
    if not selected_sheets:
        return "錯誤：您沒有選擇任何要套用格式的工作表。", 400

    processed_safe_filename = apply_format_to_file(template_path, target_path, source_sheet_name, selected_sheets)
    
    # 清理 session
    session.pop('template_path', None)
    session.pop('template_filename_original', None)
    session.pop('target_path', None)
    session.pop('target_filename_original', None)
    session.pop('source_sheet_name', None)

    if processed_safe_filename:
        # 【修改點】根據原始檔名，建立使用者看到的下載檔名
        root, ext = os.path.splitext(target_filename_original)
        download_filename_original = f"{root}_formatted{ext}"
        
        # 使用 send_from_directory 的 download_name 參數
        response = make_response(send_from_directory(
            app.config['DOWNLOAD_FOLDER'],
            path=processed_safe_filename, # 在伺服器上找這個安全檔名
            download_name=download_filename_original, # 告訴瀏覽器下載時用這個原始檔名
            as_attachment=True
        ))
        response.set_cookie('fileDownloadInProgress', 'true', path='/')
        return response
    else:
        return "處理檔案時發生內部錯誤，請檢查伺服器日誌。", 500

if __name__ == '__main__':
    app.run(debug=True)