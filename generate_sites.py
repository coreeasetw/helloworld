from __future__ import annotations

import html
import json
import re
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional
from xml.etree import ElementTree as ET

XL_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@dataclass
class Store:
    map_url: str
    name: str
    rating: Optional[float]
    review_count: Optional[int]
    category: str
    address: str
    status: str
    closing_time: str
    phone: str
    hero_image: str
    avatar_image: str
    review_snippet: str
    slug: str


DEFAULT_AVATAR = "https://ssl.gstatic.com/local/servicebusiness/default_user.png"
DEFAULT_HERO = "https://ssl.gstatic.com/local/generic/default_logo.png"


def load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    content = zf.read("xl/sharedStrings.xml")
    root = ET.fromstring(content)
    strings: List[str] = []
    for si in root.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
        texts = []
        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            texts.append(t.text or "")
        strings.append("".join(texts))
    return strings


def column_index(cell_ref: str) -> int:
    """Return a 1-based column index extracted from a cell reference (e.g., A1 -> 1)."""
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return 1
    letters = match.group(1)
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def load_sheet_rows(zf: zipfile.ZipFile) -> List[List[str]]:
    shared = load_shared_strings(zf)
    sheet_root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
    rows: List[List[str]] = []
    for row in sheet_root.findall(".//main:row", XL_NS):
        row_vals: List[str] = []
        current_col = 1
        for cell in row.findall("main:c", XL_NS):
            value = cell.find("main:v", XL_NS)
            if value is None:
                row_vals.append("")
                continue
            val = value.text or ""
            if cell.get("t") == "s":
                val = shared[int(val)]
            idx = column_index(cell.get("r", ""))
            while current_col < idx:
                row_vals.append("")
                current_col += 1
            row_vals.append(val)
            current_col += 1
        rows.append(row_vals)
    return rows


def safe_float(value: str) -> Optional[float]:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_int(value: str) -> Optional[int]:
    try:
        # Some rows contain values like "-63"; preserve magnitude only.
        return abs(int(float(value)))
    except (TypeError, ValueError):
        return None


def slugify(name: str) -> str:
    slug = re.sub(r"[^\w]+", "-", name.strip().lower())
    slug = re.sub(r"-+", "-", slug).strip("-")
    return slug or "store"


def row_to_store(row: List[str]) -> Optional[Store]:
    if len(row) < 2:
        return None

    def get(index: int) -> str:
        return row[index] if len(row) > index else ""

    name = get(5).strip()
    if not name:
        return None

    map_url = get(4).strip()
    rating = safe_float(get(6).strip())
    review_count = safe_int(get(7).strip())
    category = get(8).strip()
    address = get(11).strip()
    status = get(12).strip()
    closing_time = get(13).strip()
    phone = get(15).strip()
    hero_image = get(16).strip() or DEFAULT_HERO
    avatar_image = get(17).strip() or DEFAULT_AVATAR
    review_snippet = "".join(part.strip() for part in row[18:] if part).strip()

    return Store(
        map_url=map_url,
        name=name,
        rating=rating,
        review_count=review_count,
        category=category,
        address=address,
        status=status,
        closing_time=closing_time,
        phone=phone,
        hero_image=hero_image,
        avatar_image=avatar_image,
        review_snippet=review_snippet,
        slug=slugify(name),
    )


def load_stores(xlsx_path: Path) -> List[Store]:
    with zipfile.ZipFile(xlsx_path) as zf:
        rows = load_sheet_rows(zf)
    stores = [row_to_store(row) for row in rows[1:]]
    return [s for s in stores if s is not None]


def render_index(stores: List[Store]) -> str:
    cards = []
    for store in stores:
        rating = f"{store.rating:.1f}" if store.rating is not None else "N/A"
        reviews = f"（{store.review_count} 則評論）" if store.review_count is not None else ""
        snippet = html.escape(store.review_snippet or "尚無評論摘錄")
        cards.append(
            f"""
            <article class='card'>
              <img src='{html.escape(store.hero_image)}' alt='{html.escape(store.name)} 店面圖片' class='hero'>
              <div class='card__body'>
                <h2>{html.escape(store.name)}</h2>
                <p class='meta'>{html.escape(store.category)} · 評分 {rating} {reviews}</p>
                <p class='address'>📍 {html.escape(store.address)}</p>
                <p class='status'>{html.escape(store.status)} {html.escape(store.closing_time)}</p>
                <p class='snippet'>{snippet}</p>
                <div class='actions'>
                  <a class='button' href='stores/{store.slug}.html'>查看詳情</a>
                  <a class='button button--ghost' href='{html.escape(store.map_url)}' target='_blank' rel='noopener'>Google 地圖</a>
                </div>
              </div>
            </article>
            """
        )

    return PAGE_TEMPLATE.format(
        title="士林水電行指南",
        description="從 Excel 產生的水電行清單，點擊卡片可瀏覽各店獨立頁面。",
        content="\n".join(cards),
        asset_prefix="",
    )


def render_detail(store: Store) -> str:
    rating = f"{store.rating:.1f}" if store.rating is not None else "N/A"
    reviews = f"（{store.review_count} 則評論）" if store.review_count is not None else ""
    review_section = (
        f"<p class='snippet'>{html.escape(store.review_snippet)}</p>" if store.review_snippet else "<p class='snippet'>尚無評論摘錄</p>"
    )
    return PAGE_TEMPLATE.format(
        title=f"{store.name}｜士林水電行指南",
        description=f"{store.name} 的服務資訊與聯絡方式。",
        content=f"""
        <article class='detail'>
          <header class='detail__header'>
            <img src='{html.escape(store.hero_image)}' alt='{html.escape(store.name)} 店面圖片' class='hero hero--large'>
            <div>
              <p class='kicker'>{html.escape(store.category)}</p>
              <h1>{html.escape(store.name)}</h1>
              <p class='meta'>評分 {rating} {reviews}</p>
            </div>
          </header>
          <section class='info'>
            <h2>店家資訊</h2>
            <ul>
              <li>📍 地址：{html.escape(store.address)}</li>
              <li>⏰ 營業資訊：{html.escape(store.status)} {html.escape(store.closing_time)}</li>
              <li>☎️ 電話：<a href='tel:{html.escape(store.phone)}'>{html.escape(store.phone)}</a></li>
              <li>🗺️ 地圖：<a href='{html.escape(store.map_url)}' target='_blank' rel='noopener'>在 Google 地圖開啟</a></li>
            </ul>
          </section>
          <section class='info'>
            <h2>顧客評論摘錄</h2>
            {review_section}
          </section>
          <section class='info avatar'>
            <img src='{html.escape(store.avatar_image)}' alt='{html.escape(store.name)} 的代表頭像'>
            <div>
              <p>網站由 Excel 資料自動生成，便於透過 GitHub Pages 快速發布。</p>
              <p><a class='button' href='../index.html'>返回列表</a></p>
            </div>
          </section>
        </article>
        """,
        asset_prefix="../",
    )


PAGE_TEMPLATE = """<!doctype html>
<html lang='zh-Hant'>
<head>
  <meta charset='utf-8'>
  <meta name='viewport' content='width=device-width, initial-scale=1'>
  <title>{title}</title>
  <meta name='description' content='{description}'>
  <link rel='stylesheet' href='{asset_prefix}assets/site.css'>
</head>
<body>
  <nav class='topbar'>
    <div class='brand'>士林水電行資料集</div>
    <a class='topbar__link' href='https://github.com/'>GitHub Pages</a>
  </nav>
  <main class='container'>
    {content}
  </main>
  <footer class='footer'>
    以 Excel 自動生成的靜態網站，部署至 GitHub Pages 即可瀏覽。
  </footer>
</body>
</html>
"""


def write_site(stores: List[Store], output_dir: Path) -> None:
    assets_dir = output_dir / "assets"
    stores_dir = output_dir / "stores"
    assets_dir.mkdir(parents=True, exist_ok=True)
    stores_dir.mkdir(parents=True, exist_ok=True)

    (assets_dir / "site.css").write_text(STYLE_CSS, encoding="utf-8")
    (output_dir / "index.html").write_text(render_index(stores), encoding="utf-8")

    for store in stores:
        (stores_dir / f"{store.slug}.html").write_text(render_detail(store), encoding="utf-8")

    data_path = output_dir / "stores.json"
    data_path.write_text(json.dumps([asdict(s) for s in stores], ensure_ascii=False, indent=2), encoding="utf-8")


STYLE_CSS = """
:root {
  --bg: #0e1117;
  --card: #161b22;
  --text: #e6edf3;
  --muted: #9ba3b4;
  --accent: #8dd0ff;
  --border: #30363d;
  --shadow: 0 10px 30px rgba(0,0,0,0.35);
  font-family: "Noto Sans TC", "Inter", "Segoe UI", system-ui, sans-serif;
}

* { box-sizing: border-box; }
body {
  margin: 0;
  background: var(--bg);
  color: var(--text);
}

.topbar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 16px 24px;
  border-bottom: 1px solid var(--border);
  background: rgba(12,16,22,0.9);
  position: sticky;
  top: 0;
  z-index: 5;
  backdrop-filter: blur(10px);
}

.brand { font-weight: 700; letter-spacing: 0.02em; }
.topbar__link { color: var(--accent); text-decoration: none; }

.container {
  padding: 32px 24px 80px;
  max-width: 1200px;
  margin: auto;
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 20px;
}

.card, .detail {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 16px;
  box-shadow: var(--shadow);
  overflow: hidden;
}

.card__body { padding: 16px; display: flex; flex-direction: column; gap: 6px; }
.card h2 { margin: 0; }
.hero {
  width: 100%;
  height: 180px;
  object-fit: cover;
  background: #111;
}
.hero--large { height: 260px; }

.meta { color: var(--muted); font-size: 14px; }
.address, .status { margin: 0; font-size: 15px; }
.snippet { color: var(--muted); }
.actions { display: flex; gap: 8px; margin-top: 6px; flex-wrap: wrap; }
.button {
  display: inline-block;
  padding: 8px 12px;
  border-radius: 10px;
  background: var(--accent);
  color: #0a0c10;
  font-weight: 600;
  text-decoration: none;
  box-shadow: 0 8px 18px rgba(141,208,255,0.25);
}
.button--ghost { background: transparent; color: var(--accent); border: 1px solid var(--accent); box-shadow: none; }

.footer {
  padding: 24px;
  text-align: center;
  color: var(--muted);
  border-top: 1px solid var(--border);
  background: #0a0c10;
}

.detail { grid-column: 1 / -1; padding: 0 0 24px; }
.detail__header { display: grid; grid-template-columns: 1fr; gap: 12px; }
.detail__header h1 { margin: 4px 0; }
.kicker { color: var(--muted); margin: 0; text-transform: uppercase; letter-spacing: 0.08em; }
.info { padding: 0 24px; margin-top: 12px; }
.info ul { list-style: none; padding: 0; margin: 0; display: grid; gap: 6px; }
.info li { color: var(--muted); }
.avatar { display: grid; grid-template-columns: 96px 1fr; align-items: center; gap: 12px; }
.avatar img { width: 96px; height: 96px; border-radius: 50%; object-fit: cover; border: 1px solid var(--border); }

@media (min-width: 820px) {
  .detail__header { grid-template-columns: 1.1fr 1fr; align-items: center; }
  .container { grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); }
}
"""


def main() -> None:
    xlsx_path = Path("新增 Microsoft Excel 工作表.xlsx")
    output_dir = Path("docs")
    stores = load_stores(xlsx_path)
    write_site(stores, output_dir)
    print(f"Generated {len(stores)} store pages in {output_dir}/stores/")


if __name__ == "__main__":
    main()
