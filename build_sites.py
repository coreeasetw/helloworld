import os
import re
import html
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_FILE = BASE_DIR / '新增 Microsoft Excel 工作表.xlsx'
DOCS_DIR = BASE_DIR / 'docs'
SHOPS_DIR = DOCS_DIR / 'shops'

ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def load_shared_strings(zf):
    shared = []
    ss_root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
    for si in ss_root.findall('.//a:si', ns):
        text = ''.join(t.text or '' for t in si.findall('.//a:t', ns))
        shared.append(text)
    return shared


def load_rows():
    with zipfile.ZipFile(DATA_FILE) as zf:
        shared = load_shared_strings(zf)
        sheet = ET.fromstring(zf.read('xl/worksheets/sheet1.xml'))
        rows = []
        for row in sheet.findall('.//a:sheetData/a:row', ns):
            values = []
            for cell in row.findall('a:c', ns):
                cell_type = cell.get('t')
                value_node = cell.find('a:v', ns)
                if value_node is None:
                    values.append('')
                    continue
                if cell_type == 's':
                    values.append(shared[int(value_node.text)])
                else:
                    values.append(value_node.text)
            rows.append(values)
        return rows


def to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value):
    f = to_float(value)
    return int(abs(f)) if f is not None else None


def clean_token(token):
    if token is None:
        return None
    token = token.strip()
    if not token or token in {'·', '\uf54a', '\ue934', '\ue5d4'}:
        return None
    return token


def detect_fields(tokens):
    tokens = [t for t in (clean_token(tok) for tok in tokens) if t]
    address = next((t for t in tokens if any(ch.isdigit() for ch in t) and any(key in t for key in ['路', '街', '巷', '段', '號'])), None)
    phone = next((t for t in tokens if re.search(r'\d', t) and (t.startswith('0') or '-' in t) and len(re.sub(r'\D', '', t)) >= 7), None)
    status_parts = [t for t in tokens if '營業' in t or '打烊' in t]
    status = ' / '.join(status_parts) if status_parts else None
    category = next((t for t in tokens if '店' in t or '公司' in t or '工程' in t or '服務' in t or '承辦' in t or '商店' in t), None)
    image = next((t for t in tokens if 'googleusercontent' in t or 'streetviewpixels' in t), None)
    review_snippet = next((t for t in tokens if '\"' in t or '"' in t), None)

    remaining = []
    for t in tokens:
        if t in {address, phone, category, image, review_snippet} or (status and t in status_parts):
            continue
        remaining.append(t)

    return {
        'category': category,
        'address': address,
        'phone': phone,
        'status': status,
        'image': image,
        'review_snippet': review_snippet.strip('"') if review_snippet else None,
        'extra': remaining,
    }


def slugify(name, index):
    base = re.sub(r'[^a-zA-Z0-9]+', '-', name).strip('-').lower()
    if not base:
        base = f'shop-{index:02d}'
    return base


def parse_shops():
    shops = []
    for row in load_rows():
        if not row or not row[0].startswith('http'):
            continue
        name = row[1] if len(row) > 1 else '未命名水電行'
        rating = to_float(row[2])
        review_count = to_int(row[3])
        fields = detect_fields(row[4:])
        shops.append({
            'map_url': row[0],
            'name': name,
            'rating': rating,
            'review_count': review_count,
            **fields,
        })
    return shops


def write_file(path: Path, content: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding='utf-8')


def render_shop_page(shop, slug, index):
    rating = f"{shop['rating']:.1f}" if shop['rating'] is not None else '—'
    reviews = shop['review_count'] if shop['review_count'] is not None else '—'
    details = []
    if shop.get('category'):
        details.append(f"<li><strong>類型：</strong>{html.escape(shop['category'])}</li>")
    if shop.get('address'):
        details.append(f"<li><strong>地址：</strong>{html.escape(shop['address'])}</li>")
    if shop.get('phone'):
        details.append(f"<li><strong>電話：</strong>{html.escape(shop['phone'])}</li>")
    if shop.get('status'):
        details.append(f"<li><strong>營業資訊：</strong>{html.escape(shop['status'])}</li>")
    if shop.get('review_snippet'):
        details.append(f"<li><strong>評論摘錄：</strong>{html.escape(shop['review_snippet'])}</li>")
    for extra in shop.get('extra', []):
        details.append(f"<li>{html.escape(extra)}</li>")

    image_html = ''
    if shop.get('image'):
        image_html = f"<div class='photo'><img src='{html.escape(shop['image'])}' alt='{html.escape(shop['name'])}'></div>"

    return f"""
<!doctype html>
<html lang='zh-Hant'>
<head>
  <meta charset='utf-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1'>
  <title>{html.escape(shop['name'])}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;600&display=swap" rel="stylesheet">
  <link rel='stylesheet' href='../styles.css'>
</head>
<body>
  <header class='hero'>
    <div class='hero__inner'>
      <a class='back' href='../index.html'>← 返回列表</a>
      <h1>{html.escape(shop['name'])}</h1>
      <div class='meta'>
        <span class='rating'>⭐ {rating}</span>
        <span class='reviews'>{reviews} 則評論</span>
      </div>
      <a class='map-button' href='{html.escape(shop['map_url'])}' target='_blank' rel='noopener'>在 Google Maps 開啟</a>
    </div>
  </header>
  <main class='content'>
    {image_html}
    <section>
      <h2>店家資訊</h2>
      <ul class='details'>
        {''.join(details) if details else '<li>目前沒有更多資訊。</li>'}
      </ul>
    </section>
  </main>
  <footer class='footer'>
    <p>本頁面由資料集自動產生，供在 GitHub Pages 上瀏覽。</p>
  </footer>
</body>
</html>
"""


def render_index(shops, slugs):
    cards = []
    for shop, slug in zip(shops, slugs):
        rating = f"{shop['rating']:.1f}" if shop['rating'] is not None else '—'
        address = shop.get('address') or '地址資訊未提供'
        cards.append(f"""
        <article class='card'>
          <h2><a href='shops/{slug}.html'>{html.escape(shop['name'])}</a></h2>
          <p class='meta'>⭐ {rating}｜{shop.get('review_count', '—')} 則評論</p>
          <p class='address'>{html.escape(address)}</p>
          <a class='map-inline' href='{html.escape(shop['map_url'])}' target='_blank' rel='noopener'>地圖</a>
        </article>
        """)

    return f"""
<!doctype html>
<html lang='zh-Hant'>
<head>
  <meta charset='utf-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1'>
  <title>水電行列表</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;600&display=swap" rel="stylesheet">
  <link rel='stylesheet' href='styles.css'>
</head>
<body>
  <header class='hero'>
    <div class='hero__inner'>
      <h1>水電行資料庫</h1>
      <p>自動依 Excel 資料生成的靜態網站，每一間店都有獨立頁面。</p>
    </div>
  </header>
  <main class='grid'>
    {''.join(cards)}
  </main>
  <footer class='footer'>
    <p>將 docs 資料夾推送到 GitHub 主分支或 gh-pages 分支即可透過 GitHub Pages 瀏覽。</p>
  </footer>
</body>
</html>
"""


def render_styles():
    return """
:root { color-scheme: light; }
* { box-sizing: border-box; }
body { font-family: 'Noto Sans TC', system-ui, -apple-system, sans-serif; margin: 0; background:#f6f8fb; color:#202124; }
a { color:#0b57d0; text-decoration:none; }
a:hover { text-decoration:underline; }
.hero { background: linear-gradient(120deg, #0b57d0, #3f8efc); color: white; padding: 32px 16px; box-shadow:0 2px 10px rgba(0,0,0,0.15); }
.hero__inner { max-width: 1024px; margin: 0 auto; }
.hero h1 { margin: 0 0 8px; font-size: 32px; }
.hero p { margin: 0; opacity: 0.9; }
.grid { max-width: 1024px; margin: 24px auto; padding: 0 16px; display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; }
.card { background: white; border-radius: 12px; padding: 16px; box-shadow:0 2px 8px rgba(0,0,0,0.08); display:flex; flex-direction:column; gap:6px; }
.card h2 { margin: 0; font-size: 18px; }
.card .meta, .card .address { margin: 0; color: #444; }
.card .map-inline { align-self: flex-start; margin-top: auto; display:inline-flex; align-items:center; gap:6px; background:#eaf2ff; color:#0b57d0; padding:8px 10px; border-radius:8px; text-decoration:none; font-weight:600; }
.card .map-inline:hover { background:#d8e6ff; }
.meta { display:flex; gap:8px; align-items:center; }
.content { max-width: 800px; margin: 24px auto; padding: 0 16px 32px; background:white; border-radius: 12px; box-shadow:0 2px 8px rgba(0,0,0,0.08); }
.back { color: white; text-decoration:none; display:inline-block; margin-bottom:12px; }
.map-button { display:inline-block; margin-top:10px; padding:10px 16px; background:white; color:#0b57d0; border-radius:8px; font-weight:700; text-decoration:none; }
.photo { text-align:center; padding:16px; }
.photo img { max-width:100%; border-radius:12px; box-shadow:0 2px 10px rgba(0,0,0,0.1); }
section { padding: 0 16px 16px; }
section h2 { margin-top:0; }
.details { list-style: none; padding: 0; margin: 0; display:flex; flex-direction:column; gap:8px; }
.footer { text-align:center; color:#666; padding:16px; }
@media (max-width: 640px) {
  .hero h1 { font-size: 24px; }
}
"""


def main():
    shops = parse_shops()
    slugs = [slugify(shop['name'], i + 1) for i, shop in enumerate(shops)]

    write_file(DOCS_DIR / 'styles.css', render_styles())
    write_file(DOCS_DIR / 'index.html', render_index(shops, slugs))
    for i, (shop, slug) in enumerate(zip(shops, slugs), start=1):
        write_file(SHOPS_DIR / f'{slug}.html', render_shop_page(shop, slug, i))


if __name__ == '__main__':
    main()
