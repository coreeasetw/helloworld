# helloworld

這個專案會把提供的水電行清單轉成 GitHub Pages 靜態網站。首頁顯示所有店家的卡片，點擊即可進到各自的獨立頁面查看地址、評價、營業狀態與地圖連結。

## 如何重新產生頁面
1. 確認 Excel 檔案維持在根目錄（`新增 Microsoft Excel 工作表.xlsx`）。
2. 執行 `python build.py`，會在 `docs/` 內輸出首頁與每間水電行的獨立頁面。
3. 將 GitHub Pages 指向 `docs/` 作為發佈來源即可。

> 備註：`build.py` 使用標準函式庫解析 Excel，所以不需要額外安裝套件。
