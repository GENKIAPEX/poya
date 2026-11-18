"""Read any 寶雅新品報價單 (new 實體專用 xlsx or old 國際櫃/東京櫃 format) and return
product info for the BA site: retail price (市價, shown as 牌價) only — never cost/wholesale/margin. Columns are located by header text, so both layouts work."""
import re, warnings
import openpyxl
warnings.filterwarnings('ignore')

def _norm(v):
    return re.sub(r'\s+', '', str(v)) if v is not None else ''

def ean_ok(e):
    if not re.fullmatch(r'\d{13}', e):
        return False
    d = list(map(int, e))
    return (10 - sum(d[i] * (3 if i % 2 else 1) for i in range(12)) % 10) % 10 == d[12]

def extract(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = next((wb[n] for n in wb.sheetnames if '報價單' in n), wb.worksheets[0])
    hdr_row, cols = None, {}
    for r in range(1, 15):
        row = {c.column: _norm(c.value) for c in ws[r]}
        if any('國際條碼' in v for v in row.values()):
            hdr_row = r
            for col, v in row.items():
                for key, pat in [('ean', '國際條碼'), ('code', '廠商商品編碼'), ('name', '品名'),
                                 ('days', '保存天數'), ('period', '新品檔期'), ('unit', '單位'),
                                 ('price', '市價')]:
                    if pat in v and key not in cols:
                        cols[key] = col
            break
    if not hdr_row:
        return None
    meta = ' '.join(_norm(c.value) for r in range(1, hdr_row) for c in ws[r] if c.value)
    m = re.search(r'製表日期[:：](\d{4})[./](\d{1,2})[./](\d{1,2})', meta)
    made = f'{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}' if m else ''
    m = re.search(r'新品檔期[:：]([\d./]+)', meta)
    period = m.group(1) if m else ''
    items = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        ev = ws.cell(r, cols['ean']).value
        if ev is None:
            continue
        ean = str(int(ev)) if isinstance(ev, float) else str(ev).strip()
        if not ean.isdigit() or len(ean) < 8:
            continue
        if len(ean) == 12:
            ean = '0' + ean
        name = ws.cell(r, cols['name']).value
        name = re.sub(r'[ \t]+', ' ', '\n'.join(l.strip() for l in str(name or '').splitlines() if l.strip()))
        code = ws.cell(r, cols['code']).value if 'code' in cols else ''
        code = str(int(code)) if isinstance(code, float) else str(code or '').strip()
        days = ws.cell(r, cols['days']).value if 'days' in cols else None
        vols = re.findall(r'(\d+(?:\.\d+)?\s*(?:ml|ML|g|入|片))(?![a-zA-Z])', name.splitlines()[-1] if name else '')
        vol = vols[-1].replace(' ', '').replace('ML', 'ml') if vols else ''
        p = ws.cell(r, cols['period']).value if 'period' in cols else None
        pv = ws.cell(r, cols['price']).value if 'price' in cols else None
        price = int(pv) if isinstance(pv, (int, float)) and float(pv).is_integer() else (pv if isinstance(pv, (int, float)) else None)
        items.append(dict(ean=ean, ok=ean_ok(ean), code=code, name=name, price=price,
                          vol=vol,
                          years=(round(days / 365) if abs(days - 365 * round(days / 365)) <= 30 else None)
                                if isinstance(days, (int, float)) and days else None,
                          days=int(days) if isinstance(days, (int, float)) and days else None,
                          period=str(p).strip() if p else period))
    for i, it in enumerate(items, 1):
        it['no'] = i
    return dict(made=made, period=period, items=items)


def images(path, size=300):
    """Return {ean: data-uri} from the 圖檔 sheet (matched by 國際條碼 in col B, else by 品項 no. in col A)."""
    import io, base64, zipfile, posixpath
    from PIL import Image
    out_e, out_n = {}, {}
    try:
        z = zipfile.ZipFile(path)
        wbx = z.read('xl/workbook.xml').decode()
        rels = z.read('xl/_rels/workbook.xml.rels').decode()
        m = re.search(r'<sheet name="圖檔"[^>]*r:id="(rId\d+)"', wbx)
        if not m:
            return out_e, out_n
        tgt = re.search(rf'Id="{m.group(1)}"[^>]*Target="([^"]+)"', rels).group(1)
        sheet = posixpath.normpath(posixpath.join('xl', tgt.lstrip('/').replace('xl/', '', 1)))
        srel = posixpath.join(posixpath.dirname(sheet), '_rels', posixpath.basename(sheet) + '.rels')
        if srel not in z.namelist():
            return out_e, out_n
        dm = re.search(r'Target="([^"]*drawings/[^"]+)"', z.read(srel).decode())
        if not dm:
            return out_e, out_n
        drawing = posixpath.normpath(posixpath.join(posixpath.dirname(sheet), dm.group(1)))
        drel = posixpath.join(posixpath.dirname(drawing), '_rels', posixpath.basename(drawing) + '.rels')
        dmap = {i: posixpath.normpath(posixpath.join(posixpath.dirname(drawing), t))
                for i, t in re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', z.read(drel).decode())}
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb['圖檔']
        EMU_PT = 12700
        dflt = (ws.sheet_format.defaultRowHeight or 15)
        def row_h(r):
            h = ws.row_dimensions[r].height if r in ws.row_dimensions else None
            return (h or dflt) * EMU_PT
        def row_at(y):  # y in EMU from sheet top -> 1-based row
            r, top = 1, 0
            while r < 2000:
                h = row_h(r)
                if y < top + h:
                    return r
                top += h; r += 1
            return r
        def top_of(r):
            return sum(row_h(k) for k in range(1, r))
        by_row = {}
        for a in re.findall(r'<xdr:(?:oneCellAnchor|twoCellAnchor).*?</xdr:(?:oneCellAnchor|twoCellAnchor)>', z.read(drawing).decode(), re.S):
            fm = re.search(r'<xdr:from>.*?<xdr:row>(\d+)</xdr:row>\s*<xdr:rowOff>(-?\d+)</xdr:rowOff>', a, re.S)
            emb = re.search(r'r:embed="(rId\d+)"', a)
            if not (fm and emb) or emb.group(1) not in dmap:
                continue
            y0 = top_of(int(fm.group(1)) + 1) + int(fm.group(2))
            tm = re.search(r'<xdr:to>.*?<xdr:row>(\d+)</xdr:row>\s*<xdr:rowOff>(-?\d+)</xdr:rowOff>', a, re.S)
            em = re.search(r'<xdr:ext cx="\d+" cy="(\d+)"', a)
            y1 = top_of(int(tm.group(1)) + 1) + int(tm.group(2)) if tm else y0 + (int(em.group(1)) if em else 0)
            row = row_at((y0 + y1) / 2)
            try:
                im = Image.open(io.BytesIO(z.read(dmap[emb.group(1)]))).convert('RGB')
            except Exception:
                continue
            im.thumbnail((size, size), Image.LANCZOS)
            buf = io.BytesIO(); im.save(buf, 'JPEG', quality=80)
            by_row.setdefault(row, 'data:image/jpeg;base64,' + base64.b64encode(buf.getvalue()).decode())
        prev = ''
        for row in range(2, ws.max_row + 1):
            uri = by_row.get(row)
            note = str(ws[f'D{row}'].value or '')
            if not uri and '同上' in note:
                uri = prev
            if not uri:
                continue
            prev = uri
            b = ws[f'B{row}'].value
            e = str(int(b)) if isinstance(b, float) else str(b or '').strip()
            if len(e) == 12 and e.isdigit():
                e = '0' + e
            if e.isdigit():
                out_e.setdefault(e, uri)
            n = ws[f'A{row}'].value
            if isinstance(n, (int, float)):
                out_n.setdefault(int(n), uri)
    except Exception as ex:
        print('images: skipped', path, ex)
    return out_e, out_n
