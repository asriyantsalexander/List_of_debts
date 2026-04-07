import re
import os
import io
import tempfile
import subprocess
from flask import Flask, request, send_file, render_template, jsonify
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
 
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB
 
# ── Extraction logic ────────────────────────────────────────────────────────
 
def last_day(d):
    _, m, y = map(int, d.split('.')); m -= 1
    if m == 0: m = 12; y -= 1
    ends = {1:31,2:28,3:31,4:30,5:31,6:30,7:31,8:31,9:30,10:31,11:30,12:31}
    if m == 2 and (y%4==0 and (y%100!=0 or y%400==0)): ends[2] = 29
    return f"{ends[m]:02d}.{m:02d}.{y}"
 
def prev_1st(d):
    _, m, y = map(int, d.split('.')); m -= 1
    if m == 0: m = 12; y -= 1
    return f"01.{m:02d}.{y}"
 
def prev_11th(d):
    _, m, y = map(int, d.split('.')); m -= 1
    if m == 0: m = 12; y -= 1
    return f"11.{m:02d}.{y}"
 
def detect_columns(text):
    """
    Dynamically detect service column order from PDF header.
    Returns list of roles: 'teplo', 'pod', 'nos', 'nos_pov', 'skip', 'итого'
    Works for both Люберецкая теплосеть and ДМУП ЭКПО card formats.
    """
    m = re.search(r'периодам\n(.*?)Сальдо на', text, re.DOTALL)
    if not m:
        return None
    raw = re.sub(r'\n+', ' ', m.group(1)).strip()
 
    # Order matters: longer/more specific patterns must come first
    patterns = [
        (r'Горячее в/с \(носитель\)\s*\(Повышающий[\w\s,]*?\)', 'nos_pov'),
        (r'ГВС носитель\s*\(повышающий[\w\s,]*?\)',              'nos_pov'),
        (r'Холодное в/с\s*\(Повышающий[\w\s,]*?\)',              'skip'),
        (r'Горячее в/с \(носитель\)',                             'nos'),
        (r'Горячее в/с \(энергия\)',                              'pod'),
        (r'Отопление',                                            'teplo'),
        (r'Водоотведение',                                        'skip'),
        (r'Холодное в/с',                                         'skip'),
        (r'ИТОГО',                                                'итого'),
    ]
 
    found = []
    for pattern, role in patterns:
        for mf in re.finditer(pattern, raw, re.IGNORECASE):
            found.append((mf.start(), role, mf.group()))
    found.sort(key=lambda x: x[0])
 
    result, last_end = [], 0
    for pos, role, matched in found:
        if pos >= last_end:
            result.append(role)
            last_end = pos + len(matched)
    return result
 
 
def parse_row(nums, cols):
    """Map numeric values to (teplo, pod, tep) using column roles."""
    svc_cols = [c for c in cols if c != 'итого']
    teplo = sum(nums[i] for i, c in enumerate(svc_cols) if c == 'teplo'            and i < len(nums))
    pod   = sum(nums[i] for i, c in enumerate(svc_cols) if c == 'pod'              and i < len(nums))
    tep   = sum(nums[i] for i, c in enumerate(svc_cols) if c in ('nos', 'nos_pov') and i < len(nums))
    return teplo, pod, tep
 
 
def extract_card(pdf_path):
    r = PdfReader(pdf_path)
    text = ''.join(p.extract_text() for p in r.pages) + '\n'
 
    cols = detect_columns(text)
    if not cols:
        return None
 
    n_svc = len([c for c in cols if c != 'итого'])
 
    vp = r'\s+([-\d.]+)'
    all_saldo = []
    for nc in range(9, 0, -1):
        found = re.findall(r'Сальдо на (\d{2}\.\d{2}\.\d{4})' + vp * (nc + 1), text)
        if found:
            all_saldo = found
            break
    if not all_saldo:
        return None
 
    def get_vals(entry):
        return [float(x) for x in entry[1:]][:-1]  # drop ИТОГО
 
    last_date = all_saldo[-1][0]
    end_date  = last_day(last_date)
 
    lv = parse_row(get_vals(all_saldo[-1]), cols)
    td, pd, ted = (max(0.0, x) for x in lv)
 
    # First positive saldo date per service
    fps = [None, None, None]
    for e in all_saldo:
        t, p, te = parse_row(get_vals(e), cols)
        dt = e[0]
        if fps[0] is None and t  > 0: fps[0] = dt
        if fps[1] is None and p  > 0: fps[1] = dt
        if fps[2] is None and te > 0: fps[2] = dt
    first_d = [prev_1st(x) if x else None for x in fps]
 
    # Last peni saldo (remaining debt, not total charged)
    peni_pat = r'Сальдо пени на конец периода' + vp * (n_svc + 1)
    all_peni = re.findall(peni_pat, text)
    if all_peni:
        lp = parse_row([float(x) for x in all_peni[-1]][:-1], cols)
        tp, pp, tep_p = (max(0.0, x) for x in lp)
    else:
        tp = pp = tep_p = 0.0
 
    # First positive peni date per service — position-based scan
    fpp = [None, None, None]
    pm_list = list(re.finditer(peni_pat, text))
    sd_list = list(re.finditer(r'Сальдо на (\d{2}\.\d{2}\.\d{4})', text))
    for pm in pm_list:
        next_sd = next((m.group(1) for m in sd_list if m.start() > pm.start()), None)
        if not next_sd:
            continue
        ps = prev_11th(next_sd)
        pv = parse_row([float(x) for x in pm.groups()][:-1], cols)
        if fpp[0] is None and pv[0] > 0: fpp[0] = ps
        if fpp[1] is None and pv[1] > 0: fpp[1] = ps
        if fpp[2] is None and pv[2] > 0: fpp[2] = ps
        if all(x is not None for x in fpp):
            break
 
    return dict(
        teplo_dolg=round(td, 2),  pod_dolg=round(pd, 2),   tep_dolg=round(ted, 2),
        teplo_peni=round(tp, 2),  pod_peni=round(pp, 2),   tep_peni=round(tep_p, 2),
        first_d=first_d, first_p=fpp, end_date=end_date
    )
 
 
def build_xlsx(results, acc_order):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Лист2'
 
    thin = Side(style='thin', color='000000')
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    h9   = Font(name='Calibri', bold=True, size=9)
    d11  = Font(name='Calibri', size=11)
    ca   = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cc   = Alignment(horizontal='center', vertical='center')
 
    def hc(c, v): c.value=v; c.font=h9;  c.alignment=ca; c.border=bdr
    def dc(c, v): c.value=v; c.font=d11; c.alignment=cc; c.border=bdr
 
    ws.merge_cells('C2:F2'); ws.merge_cells('G2:J2'); ws.merge_cells('K2:N2')
    hc(ws['A2'], None); hc(ws['B2'], None)
    hc(ws['C2'], 'Теплоснабжение')
    hc(ws['G2'], 'Горячее водоснабжение - подогрев')
    hc(ws['K2'], 'Горячее водоснабжение - теплоноситель')
 
    for col, h in enumerate([
        'Лицевой счет', 'Назначение платежа',
        'Задолженность', 'Период', 'Пени', 'Период',
        'Задолженность', 'Период', 'Пени', 'Период',
        'Задолженность', 'Период', 'Пени', 'Период'], 1):
        hc(ws.cell(3, col), h)
 
    for col, w in zip('ABCDEFGHIJKLMN',
                      [12, 40, 12, 22, 10, 22, 12, 22, 10, 22, 12, 22, 10, 22]):
        ws.column_dimensions[col].width = w
 
    def ps(start, val, end):
        return f"{start}\u2013{end}" if (start and val > 0) else None
 
    for row_i, ls in enumerate(acc_order, start=4):
        if ls not in results:
            continue
        d   = results[ls]
        end = d['end_date']
        fd  = d['first_d']
        fp  = d['first_p']
        for col, v in enumerate([
            int(ls),
            f'взыскание задолженности лс {ls}',
            d['teplo_dolg'], ps(fd[0], d['teplo_dolg'], end),
            d['teplo_peni'], ps(fp[0], d['teplo_peni'], end),
            d['pod_dolg'],   ps(fd[1], d['pod_dolg'],   end),
            d['pod_peni'],   ps(fp[1], d['pod_peni'],   end),
            d['tep_dolg'],   ps(fd[2], d['tep_dolg'],   end),
            d['tep_peni'],   ps(fp[2], d['tep_peni'],   end),
        ], 1):
            dc(ws.cell(row_i, col), v)
 
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
 
 
# ── Routes ──────────────────────────────────────────────────────────────────
 
@app.route('/')
def index():
    return render_template('index.html')
 
 
@app.route('/process', methods=['POST'])
def process():
    if 'archive' not in request.files:
        return jsonify({'error': 'Файл не загружен'}), 400
 
    f = request.files['archive']
    if not f.filename:
        return jsonify({'error': 'Файл не выбран'}), 400
 
    with tempfile.TemporaryDirectory() as tmp:
        archive_path = os.path.join(tmp, f.filename)
        f.save(archive_path)
 
        extract_dir = os.path.join(tmp, 'cards')
        os.makedirs(extract_dir)
 
        fname_lower = f.filename.lower()
        try:
            if fname_lower.endswith('.rar'):
                subprocess.run(['unrar', 'x', archive_path, extract_dir],
                               check=True, capture_output=True)
            elif fname_lower.endswith('.zip'):
                subprocess.run(['unzip', '-o', archive_path, '-d', extract_dir],
                               check=True, capture_output=True)
            elif fname_lower.endswith('.7z'):
                subprocess.run(['7z', 'x', archive_path, f'-o{extract_dir}'],
                               check=True, capture_output=True)
            else:
                return jsonify({'error': 'Поддерживаются форматы: .rar, .zip, .7z'}), 400
        except subprocess.CalledProcessError as e:
            return jsonify({'error': f'Ошибка распаковки: {e.stderr.decode()}'}), 500
 
        pdfs = []
        for root, dirs, files in os.walk(extract_dir):
            for fname in files:
                if fname.lower().endswith('.pdf'):
                    ls = re.match(r'(\d+)', fname)
                    if ls:
                        pdfs.append((ls.group(1), os.path.join(root, fname)))
        pdfs.sort(key=lambda x: x[0])
 
        if not pdfs:
            return jsonify({'error': 'PDF-файлы не найдены в архиве'}), 400
 
        results, acc_order, errors = {}, [], []
        for ls, path in pdfs:
            try:
                d = extract_card(path)
                if d:
                    results[ls] = d
                    acc_order.append(ls)
                else:
                    errors.append(ls)
            except Exception as e:
                errors.append(f'{ls} ({e})')
 
        if not results:
            return jsonify({'error': 'Не удалось обработать ни один файл'}), 500
 
        buf = build_xlsx(results, acc_order)
 
    return send_file(
        buf,
        as_attachment=True,
        download_name='задолженность.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
 
 
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
 
