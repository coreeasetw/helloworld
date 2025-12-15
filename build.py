from __future__ import annotations
import html
import math
import re
import zipfile
from pathlib import Path
from typing import List, Dict
from xml.etree import ElementTree as ET

SOURCE_PATH = Path("新增 Microsoft Excel 工作表.xlsx")
DOCS_DIR = Path("docs")
ASSETS_DIR = DOCS_DIR / "assets"
STORES_DIR = DOCS_DIR / "stores"


def slugify(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-")
    text = text[:60].rstrip("-")
    return text or "store"


def load_strings(zip_path: Path) -> List[str]:
    with zipfile.ZipFile(zip_path) as z:
        shared_strings = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return [t.text or "" for t in shared_strings.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]


def load_rows(zip_path: Path, strings: List[str]) -> List[List[str]]:
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(zip_path) as z:
        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))

    rows: List[List[str]] = []
    for row in sheet.iter(ns + "row"):
        values: List[str] = []
        for cell in row.iter(ns + "c"):
            cell_type = cell.get("t")
            value_node = cell.find(ns + "v")
            value = value_node.text if value_node is not None else ""
            if cell_type == "s":
                value = strings[int(value)] if value else ""
            values.append(value)
        rows.append(values)
    return rows


def parse_store(row: List[str]) -> Dict[str, str]:
    def get(index: int) -> str:
        return row[index] if index < len(row) else ""

    def to_float(value: str) -> float | None:
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    name = get(1)
    if not name or name.startswith("http"):
        name = "未命名水電行"
    rating_value = to_float(get(2))
    rating = f"{rating_value:.1f}" if rating_value is not None else "N/A"

    review_value = to_float(get(3))
    review_count = "" if review_value is None else str(abs(int(review_value)))

    category = get(4)
    address = get(7)
    status = get(8)
    closing = get(9).lstrip("· ")
    phone = get(11)
    image_url = get(12)
    default_avatar = get(13)
    snippet = "".join(get(i) for i in range(14, len(row))).replace("_x000D_", " ").strip()

    return {
        "name": name,
        "slug": slugify(name),
        "map_url": get(0),
        "rating": rating,
        "review_count": review_count,
        "category": category,
        "address": address,
        "status": status,
        "closing": closing,
        "phone": phone,
        "image_url": image_url,
        "default_avatar": default_avatar,
        "snippet": snippet,
    }


def build_card(store: Dict[str, str]) -> str:
    rating_block = f"<strong>{html.escape(store['rating'])}</strong>" if store["rating"] != "N/A" else "N/A"
    reviews = f"（{store['review_count']} 則評論）" if store["review_count"] else ""
    status_parts = [part for part in [store["status"], store["closing"]] if part]
    status_text = " · ".join(status_parts)
    return f"""
    <article class=\"card\">
      <img src=\"{html.escape(store['image_url'] or store['default_avatar'])}\" alt=\"{html.escape(store['name'])}\" loading=\"lazy\" />
      <div class=\"card__body\">
        <h2><a href=\"stores/{store['slug']}/index.html\">{html.escape(store['name'])}</a></h2>
        <p class=\"card__meta\">{rating_block} {reviews}</p>
        <p class=\"card__meta\">{html.escape(store['category'])}</p>
        <p>{html.escape(store['address'])}</p>
        <p class=\"card__status\">{html.escape(status_text)}</p>
        <div class=\"card__actions\">
          <a class=\"button\" href=\"stores/{store['slug']}/index.html\">檢視店家頁面</a>
          <a class=\"button button--ghost\" href=\"{html.escape(store['map_url'])}\">查看地圖</a>
        </div>
      </div>
    </article>
    """


def render_index(stores: List[Dict[str, str]]) -> str:
    cards = "\n".join(build_card(store) for store in stores)
    return f"""<!doctype html>
<html lang=\"zh-Hant\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>水電行列表</title>
  <link rel=\"stylesheet\" href=\"assets/style.css\" />
</head>
<body>
  <header class=\"hero\">
    <div class=\"hero__content\">
      <p class=\"eyebrow\">GitHub Pages</p>
      <h1>水電行小百科</h1>
      <p>每一家水電行都有自己的獨立介紹頁面，點擊卡片即可前往。</p>
    </div>
  </header>
  <main class=\"grid\">
    {cards}
  </main>
</body>
</html>
"""


def render_store(store: Dict[str, str]) -> str:
    status_parts = [part for part in [store["status"], store["closing"]] if part]
    status_text = " · ".join(status_parts)
    review_text = f"共有 {store['review_count']} 則評論" if store["review_count"] else ""
    snippet = html.escape(store["snippet"]) if store["snippet"] else "暫無評論摘錄"
    phone_line = f"<p><strong>電話：</strong>{html.escape(store['phone'])}</p>" if store["phone"] else ""

    return f"""<!doctype html>
<html lang=\"zh-Hant\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>{html.escape(store['name'])}</title>
  <link rel=\"stylesheet\" href=\"../assets/style.css\" />
</head>
<body class=\"store-page\">
  <div class=\"store\">
    <a class=\"back-link\" href=\"../index.html\">← 返回列表</a>
    <header>
      <h1>{html.escape(store['name'])}</h1>
      <p class=\"store__meta\">{html.escape(store['category'])}</p>
      <p class=\"store__meta\">⭐ {html.escape(store['rating'])} {review_text}</p>
      <p class=\"store__meta\">{html.escape(status_text)}</p>
      <p>{html.escape(store['address'])}</p>
      {phone_line}
      <p><a class=\"button\" href=\"{html.escape(store['map_url'])}\">在 Google 地圖查看</a></p>
    </header>
    <img class=\"store__image\" src=\"{html.escape(store['image_url'] or store['default_avatar'])}\" alt=\"{html.escape(store['name'])}\" loading=\"lazy\" />
    <section class=\"store__section\">
      <h2>顧客回饋</h2>
      <p>{snippet}</p>
    </section>
  </div>
</body>
</html>
"""


def write_files(stores: List[Dict[str, str]]) -> None:
    DOCS_DIR.mkdir(exist_ok=True)
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    STORES_DIR.mkdir(parents=True, exist_ok=True)

    (DOCS_DIR / "index.html").write_text(render_index(stores), encoding="utf-8")

    for store in stores:
        store_dir = STORES_DIR / store["slug"]
        store_dir.mkdir(parents=True, exist_ok=True)
        (store_dir / "index.html").write_text(render_store(store), encoding="utf-8")

    (ASSETS_DIR / "style.css").write_text(STYLE, encoding="utf-8")


STYLE = """
:root {
  --bg: #0f172a;
  --panel: #111827;
  --text: #e2e8f0;
  --muted: #94a3b8;
  --accent: #38bdf8;
  --border: #1f2937;
  --radius: 16px;
  font-family: 'Noto Sans TC', 'Inter', system-ui, -apple-system, sans-serif;
}

* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--text); }
a { color: var(--accent); text-decoration: none; }
a:hover { text-decoration: underline; }

.hero {
  background: radial-gradient(circle at 20% 20%, rgba(56,189,248,0.15), transparent 40%),
              radial-gradient(circle at 80% 0%, rgba(94,234,212,0.1), transparent 30%),
              linear-gradient(135deg, #0b1224, #0f172a);
  padding: 64px 24px;
}
.hero__content { max-width: 960px; margin: 0 auto; }
.hero h1 { font-size: 2.4rem; margin: 8px 0; }
.hero p { color: var(--muted); }
.eyebrow { text-transform: uppercase; letter-spacing: 0.1em; font-size: 0.75rem; color: var(--accent); }

.grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 20px;
  padding: 24px;
  max-width: 1200px;
  margin: 0 auto;
}

.card {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  overflow: hidden;
  display: flex;
  flex-direction: column;
  box-shadow: 0 10px 40px rgba(0,0,0,0.35);
}
.card img { width: 100%; height: 180px; object-fit: cover; }
.card__body { padding: 16px; display: flex; flex-direction: column; gap: 8px; flex: 1; }
.card__meta { color: var(--muted); margin: 0; }
.card__status { color: #4ade80; }
.card__actions { margin-top: auto; display: flex; gap: 8px; flex-wrap: wrap; }

.button {
  display: inline-block;
  padding: 8px 12px;
  border-radius: 999px;
  background: var(--accent);
  color: #0b1224;
  font-weight: 600;
}
.button--ghost {
  background: transparent;
  border: 1px solid var(--border);
  color: var(--text);
}

.store-page { background: radial-gradient(circle at 10% 20%, rgba(56,189,248,0.2), transparent 25%), #0b1224; }
.store {
  max-width: 800px;
  margin: 24px auto;
  padding: 24px;
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: 0 10px 40px rgba(0,0,0,0.35);
}
.store h1 { margin-top: 8px; }
.store__meta { color: var(--muted); margin: 4px 0; }
.store__image { width: 100%; border-radius: 12px; object-fit: cover; margin: 16px 0; }
.store__section h2 { margin-bottom: 8px; }
.back-link { color: var(--muted); }
"""


def main() -> None:
    strings = load_strings(SOURCE_PATH)
    rows = load_rows(SOURCE_PATH, strings)
    stores = [parse_store(row) for row in rows[1:]]  # skip header row
    write_files(stores)
    print(f"Generated {len(stores)} store pages in {STORES_DIR}")


if __name__ == "__main__":
    main()
