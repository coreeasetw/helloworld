# 士林水電行 GitHub Pages 靜態網站

此專案會從儲存在 `新增 Microsoft Excel 工作表.xlsx` 的水電行資料自動產生 GitHub Pages 網站。每間店家都會得到獨立頁面，並在首頁提供彙整卡片與連結。

## 快速開始
1. 使用內建的 Python 腳本產生靜態檔案：
   ```bash
   python generate_sites.py
   ```
   腳本會在 `docs/` 下輸出首頁、每家店的內頁（共 56 頁）以及共用的 `assets/site.css` 與 `stores.json`。
2. 在 GitHub 專案的 Pages 設定中，選擇 **Deploy from a branch**，並將資料夾設為 `docs/`，分支選擇目前的主要分支。
3. 儲存設定後，GitHub Pages 會自動部署首頁 `docs/index.html` 與所有獨立店家頁面。

## 技術細節
- 不需額外套件：腳本以 Python 標準函式庫解讀 `.xlsx` 檔，解析欄位並輸出 HTML/CSS。
- 產生的 `stores.json` 可供未來前端互動功能或搜尋使用。
- 頁面樣式使用暗色系卡片布局，提供 Google 地圖與電話連結，方便訪客瀏覽。

若要更新資料，只需更新 Excel 檔並重新執行腳本即可重新生成整個網站。
