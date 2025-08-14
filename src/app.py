
# 批次儲存裁切區塊與正確答案

from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
import os
from PIL import Image
import pytesseract
import openpyxl
import random

app = Flask(__name__)
app.secret_key = 'your_secret_key'
UPLOAD_FOLDER = 'src/static/uploads'
EXCEL_FILE = 'wrong_questions.xlsx'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def create_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['題目', '解析', '標註'])
        wb.save(EXCEL_FILE)

@app.route('/')
def index():
    return render_template('index.html')
@app.route('/questions')
def questions():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    questions = list(ws.iter_rows(min_row=2, values_only=True))
    return render_template('questions.html', questions=questions)

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        flash('未選擇檔案')
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash('未選擇檔案')
        return redirect(url_for('index'))
    # 檢查副檔名
    allowed_ext = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.mpo']
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in allowed_ext:
        flash('只支援圖片格式：jpg, jpeg, png, bmp, tiff, mpo')
        return redirect(url_for('index'))
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    return redirect(url_for('crop', filename=file.filename))
@app.route('/crop/<filename>')
def crop(filename):
    return render_template('crop.html', filename=filename)

@app.route('/save_crop', methods=['POST'])
def save_crop():
    cropped = request.files['cropped_image']
    origin = request.form.get('origin')
    save_name = f"crop_{origin}_{random.randint(1000,9999)}.png"
    save_path = os.path.join(UPLOAD_FOLDER, save_name)
    cropped.save(save_path)
    # 可在此記錄到 Excel 或資料庫
    return f'裁切圖片已儲存：{save_name}'

@app.route('/save_crops', methods=['POST'])
def save_crops():
    files = request.files.getlist('cropped_images')
    answers = request.form.getlist('answers')
    origin = request.form.get('origin')
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    saved_names = []
    for idx, f in enumerate(files):
        save_name = f"crop_{origin}_{random.randint(1000,9999)}_{idx}.png"
        save_path = os.path.join(UPLOAD_FOLDER, save_name)
        f.save(save_path)
        answer = answers[idx] if idx < len(answers) else ''
        ws.append([save_name, answer, ''])  # 檔名、正確答案、標註
        saved_names.append(save_name)
    wb.save(EXCEL_FILE)
    return f'已儲存 {len(saved_names)} 個裁切區塊到 Excel！'

@app.route('/add', methods=['POST'])
def add():
    question = request.form.get('question')
    solution = request.form.get('solution', '')
    tag = request.form.get('tag', '')
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([question, solution, tag])
    wb.save(EXCEL_FILE)
    flash('已加入Excel錯題本')
    return redirect(url_for('index'))


@app.route('/random', methods=['GET', 'POST'])
def random_question():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    questions = list(ws.iter_rows(min_row=2, values_only=True))
    if not questions:
        return render_template('random.html', questions=None, num=0)
    num = 1
    if request.method == 'POST':
        try:
            num = int(request.form.get('num', 1))
        except ValueError:
            num = 1
        num = max(1, min(num, len(questions)))
    selected = random.sample(questions, min(num, len(questions)))
    return render_template('random.html', questions=selected, num=num)

# 刪除指定列的錯題
@app.route('/delete_question', methods=['POST'])
def delete_question():
    row_index = request.form.get('row_index')
    print(f"收到刪除請求 row_index: {row_index}")
    try:
        row_index = int(row_index)
    except Exception as e:
        flash(f'row_index 解析失敗: {row_index}, 錯誤: {e}')
        return redirect(url_for('questions'))
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    max_row = ws.max_row
    print(f"Excel 最大行: {max_row}")
    # Excel 第一列是標題，row_index+2 才是正確資料列
    target_row = row_index + 2
    print(f"實際要刪除的 Excel 行: {target_row}")
    if target_row <= max_row:
        ws.delete_rows(target_row)
        wb.save(EXCEL_FILE)
        flash(f'已刪除第 {target_row} 行 (row_index={row_index})')
    else:
        flash(f'刪除失敗，索引超出範圍 (row_index={row_index}, target_row={target_row}, max_row={max_row})')
    return redirect(url_for('questions'))


if __name__ == '__main__':
    create_excel()
    app.run(host='0.0.0.0', port=5000, debug=True)
