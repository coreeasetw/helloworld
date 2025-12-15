# helloworld

這個專案會把 Excel 檔案中的水電行資料轉成靜態網站，可以直接放在 GitHub Pages 瀏覽。只要執行下方指令就會在 `docs/` 產生首頁與每間水電行的獨立頁面。

```bash
python build_sites.py
```

生成後將 `docs/` 目錄推送到 GitHub 的預設分支（或 gh-pages 分支）即可啟用 GitHub Pages。
