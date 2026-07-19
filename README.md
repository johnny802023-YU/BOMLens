# BOMLens

BOMLens 是供公司內部使用的離線 BOM 與線路圖版本比對工具。所有 Excel、PDF 與比較結果都只在目前電腦的瀏覽器記憶體中處理，不需要登入 GPT，也不會上傳到外部網站。

## 主要功能

- 依 Excel 第一列欄位名稱辨識 BOM，不依賴固定欄位位置
- 比較新增元件、新增替料、移除元件、刪除替料、同位置換料與插件位置變更
- B 料號取末 12 碼作為標準料號，P 料號與製造商名稱僅供結果辨識
- BOM 插件位置可連動舊版與新版線路圖 PDF
- PDF 文字索引、分批處理、旋轉文字辨識、同步縮放與移動及本機快取管理
- 匯出易讀的 Excel 差異報告與獨立 HTML 報告
- Windows 離線啟動與可攜式封裝
- GitHub Actions 自動產生 Windows x64 Portable ZIP 與 Setup EXE

## 本機執行

需求：Node.js 22.13.0 以上。

```bash
npm install
npm run offline
```

瀏覽器開啟 `http://127.0.0.1:3784/`。詳細操作方式請參考 [LOCAL_OFFLINE_GUIDE.md](LOCAL_OFFLINE_GUIDE.md)。

Windows 可直接執行：

```text
scripts\start-offline.bat
```

如果 Windows 電腦沒有 Node.js，請從 GitHub Actions 的 `Build Windows packages` 成功紀錄下載：

- `BOMLens-Windows-x64-Portable`
- `BOMLens-Setup-x64`

兩種版本都已包含 Windows Node.js runtime，使用端不需另外安裝 Node.js。Setup 採每位使用者安裝，不需管理員權限，並會建立桌面捷徑。

## 驗證

```bash
npm run lint
npm test
```

## 機密資料注意事項

- 請勿將公司 BOM、線路圖、匯出報告或可攜式封裝提交到 Git。
- `.gitignore` 已預設排除常見 Excel、PDF、報告與輸入資料夾。
- 本專案不包含雲端部署設定；請在公司允許的本機或內網環境執行。
