from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import csv, io, os, re
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app, origins="*", allow_headers=["Content-Type"], methods=["GET","POST","OPTIONS"])

KAKAO = ['kakaonline2','kakaonline3','kakaonline4','kakaonline7','kakaonline8']
DARK="1F3864"; MID="2F5496"; LIGHT="D6E4F7"; WHITE="FFFFFF"; GRAY="F2F5FA"

def clean(v):
    if not isinstance(v, str): return v
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', v)

def hdr(cell, bg=DARK, fg=WHITE, size=10):
    cell.font = Font(name='Arial', bold=True, color=fg, size=size)
    cell.fill = PatternFill('solid', start_color=bg)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def brd(cell):
    s = Side(style='thin', color="BBBBBB")
    cell.border = Border(left=s, right=s, top=s, bottom=s)

def parse_csvs(files):
    # key: (channel, month, shop, shop_id, product, product_id, category)
    gmv_map = defaultdict(float)
    hh_map = defaultdict(lambda: defaultdict(int))

    for f in files:
        raw = f.read()
        # detect encoding
        for enc in ['utf-8-sig', 'utf-8', 'cp1252', 'latin1']:
            try:
                text = raw.decode(enc)
                break
            except:
                continue
        else:
            continue

        reader = csv.DictReader(io.StringIO(text))
        if 'Tên tài khoản KOL' not in (reader.fieldnames or []):
            continue

        for row in reader:
            ch = row.get('Tên tài khoản KOL','').strip()
            if ch not in KAKAO: continue

            date = row.get('Thời gian đặt hàng','')
            month = date[:7] if len(date) >= 7 else ''
            if not month: continue

            shop    = clean(row.get('Tên Shop','').strip())
            shop_id = row.get('ID Shop','').strip()
            product = clean(row.get('Tên sản phẩm','').strip())
            prod_id = row.get('Mã sản phẩm','').strip()
            cat     = clean(row.get('Danh mục sản phẩm cấp 1','').strip())
            hh      = row.get('Tỷ lệ hoa hồng người bán','').strip()

            try:
                gmv = float(row.get('Giá trị mua hàng(₫)','0').replace(',','') or 0)
            except:
                gmv = 0.0

            key = (ch, month, shop, shop_id, product, prod_id, cat)
            gmv_map[key] += gmv
            if hh:
                hh_map[key][hh] += 1

    # Build result list
    rows = []
    for key, gmv in gmv_map.items():
        ch, month, shop, shop_id, product, prod_id, cat = key
        hh = max(hh_map[key], key=hh_map[key].get) if hh_map[key] else ''
        try:
            sid = str(int(float(shop_id))) if shop_id else ''
        except:
            sid = shop_id
        try:
            mid = str(int(float(prod_id))) if prod_id else ''
        except:
            mid = prod_id
        link = f'https://shopee.vn/product/{sid}/{mid}' if sid and mid else ''
        rows.append((ch, month, shop, sid, product, mid, cat, gmv, hh, link))

    return rows

def build_excel(rows, months):
    # Sort by channel, month, gmv desc
    rows.sort(key=lambda r: (r[0], r[1], -r[7]))

    # Group top 500 per (channel, month)
    groups = defaultdict(list)
    for r in rows:
        k = (r[0], r[1])
        if len(groups[k]) < 500:
            groups[k].append(r)

    # Also build "all" per month
    all_by_month = defaultdict(list)
    for r in rows:
        all_by_month[r[1]].append(r)

    wb = Workbook()

    DASH_COLS = ['#','Tên Shop','ID Shop','Ngành hàng','Tên sản phẩm','Mã sản phẩm','Tổng GMV (₫)','Tỷ lệ HH NB','Link sản phẩm']
    TAB = {'kakaonline2':'E74C3C','kakaonline3':'E67E22','kakaonline4':'27AE60','kakaonline7':'2980B9','kakaonline8':'8E44AD'}
    latest = sorted(months)[-1]

    def write_sheet(wb, sheet_name, data_rows, tab_color):
        ws = wb.create_sheet(sheet_name)
        ws.sheet_properties.tabColor = tab_color
        ws.freeze_panes = 'A2'
        for c, h in enumerate(DASH_COLS, 1):
            hdr(ws.cell(1, c, h), bg="444444")
        for i, r in enumerate(data_rows, 1):
            ch,mo,shop,sid,product,mid,cat,gmv,hh,link = r
            bg = "FFD700" if i==1 else "E8E8E8" if i==2 else "CDAA7D" if i==3 else (LIGHT if i%2==0 else WHITE)
            rc = ws.cell(i+1, 1, i)
            rc.font = Font(name='Arial', size=9, bold=(i<=3))
            rc.fill = PatternFill('solid', start_color=bg)
            rc.alignment = Alignment(horizontal='center', vertical='center')
            brd(rc)
            vals = [shop, sid, cat, product, mid, gmv, hh, link]
            for ci, val in enumerate(vals, 2):
                cell = ws.cell(i+1, ci, val)
                cell.font = Font(name='Arial', size=9)
                cell.fill = PatternFill('solid', start_color=bg)
                cell.alignment = Alignment(vertical='center', wrap_text=(ci==5))
                brd(cell)
                if ci==7: cell.number_format='#,##0'
                if ci==8: cell.alignment=Alignment(horizontal='center',vertical='center')
                if ci==9 and link:
                    cell.hyperlink=link
                    cell.font=Font(name='Arial',size=9,color="0563C1",underline='single')
        for col,w in zip('ABCDEFGHI',[5,30,14,20,55,16,18,12,40]):
            ws.column_dimensions[col].width=w

    for ch in sorted(KAKAO):
        data = groups.get((ch, latest), [])
        write_sheet(wb, ch, data, TAB.get(ch,'999999'))

    # All channels sheet
    all_rows = sorted(all_by_month.get(latest,[]), key=lambda r: -r[7])[:500]
    write_sheet(wb, 'Tat ca kenh', all_rows, DARK)

    # Dashboard summary
    ws_d = wb.create_sheet("Dashboard", 0)
    ws_d.sheet_properties.tabColor = DARK
    ws_d.row_dimensions[1].height = 40
    ws_d.merge_cells('A1:I1')
    ws_d['A1'] = "TOP 500 GMV THEO KÊNH & THÁNG — KakaOnline"
    ws_d['A1'].font = Font(name='Arial', bold=True, size=16, color=WHITE)
    ws_d['A1'].fill = PatternFill('solid', start_color=DARK)
    ws_d['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws_d.row_dimensions[2].height = 22
    ws_d.merge_cells('A2:I2')
    ws_d['A2'] = f"Dữ liệu tháng {latest} — Chọn sheet kênh bên dưới để xem Top 500"
    ws_d['A2'].font = Font(name='Arial', italic=True, size=10, color="DDDDDD")
    ws_d['A2'].fill = PatternFill('solid', start_color=DARK)
    ws_d['A2'].alignment = Alignment(horizontal='center', vertical='center')

    ws_d.row_dimensions[3].height = 10
    ws_d.row_dimensions[4].height = 36
    ws_d.merge_cells('A4:I4')
    ws_d['A4'] = "👇 Click vào tab sheet bên dưới để xem Top 500 GMV từng kênh"
    ws_d['A4'].font = Font(name='Arial', bold=True, size=12, color=DARK)
    ws_d['A4'].alignment = Alignment(horizontal='center', vertical='center')

    ws_d.row_dimensions[5].height = 10
    ws_d.row_dimensions[6].height = 28
    for c, h in enumerate(['Kênh','Số SP','Tổng GMV (₫)','SP Top 1','GMV Top 1 (₫)'], 1):
        hdr(ws_d.cell(6,c,h), bg=MID)
        brd(ws_d.cell(6,c))

    row_idx = 7
    for ch in sorted(KAKAO):
        data = groups.get((ch, latest), [])
        total_gmv = sum(r[7] for r in data)
        top1_name = data[0][4] if data else ''
        top1_gmv  = data[0][7] if data else 0
        bg = LIGHT if row_idx%2==0 else WHITE
        for c, v in enumerate([ch, len(data), total_gmv, top1_name[:50], top1_gmv], 1):
            cell = ws_d.cell(row_idx, c, v)
            cell.font = Font(name='Arial', size=10)
            cell.fill = PatternFill('solid', start_color=bg)
            cell.alignment = Alignment(vertical='center')
            brd(cell)
            if c in [3,5]: cell.number_format='#,##0'
        row_idx += 1

    for col,w in zip('ABCDE',[16,10,18,50,18]):
        ws_d.column_dimensions[col].width=w

    # Remove empty sheets
    for name in list(wb.sheetnames):
        ws = wb[name]
        if ws.max_row<=1 and ws.max_column<=1:
            wb.remove(ws)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status':'ok'})

@app.route('/process', methods=['POST'])
def process():
    try:
        files = request.files.getlist('files')
        if not files:
            return jsonify({'error':'Không có file nào được upload'}), 400

        rows = parse_csvs(files)
        if not rows:
            return jsonify({'error':'Không tìm thấy data của 5 kênh KakaOnline'}), 400

        months = sorted(set(r[1] for r in rows))
        latest = months[-1]
        excel = build_excel(rows, months)

        return send_file(
            excel,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'Top_GMV_KakaOnline_{latest}.xlsx'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
