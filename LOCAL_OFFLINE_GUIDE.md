# BOMLens 公司內部離線使用

## 安全邊界

- 服務只監聽 `127.0.0.1:3784`，區域網路內其他電腦無法連入。
- BOM 與線路圖只在瀏覽器記憶體解析，不會送到伺服器、API 或雲端儲存。
- 沒有登入、同步、分析追蹤或外部連線功能。
- 關閉瀏覽器頁面後，匯入的檔案內容即從工作階段清除。
- 網頁安全政策只允許載入本機來源；相機、麥克風、定位、付款與 USB 權限均停用。

## BOM 與線路圖連動

在線路圖比對頁載入含可搜尋文字的 PDF 後，可直接點擊 BOM 差異表中的插件位置（例如 `U45`）。系統會在本機建立 PDF 文字索引，讓前後版線路圖各自跳到對應頁面並以紅框標示。若同一位置出現多次，可使用上一筆／下一筆切換。

索引採逐頁處理並顯示進度；指定位置一旦找到會先跳轉顯示，其餘頁面繼續在背景建立索引。同一工作階段切換頁籤時會沿用記憶體快取，不會重新掃描同一份 PDF。

搜尋會直接查詢插件位置索引表，PDF 則依電腦效能每批平行處理 2～4 頁。若 CAD 輸出的文字被拆成相鄰區塊，例如 `U`＋`45` 或 `U`＋`102`＋`A`，系統會依文字座標自動合併後再建立索引。

前後版 PDF 預設啟用同步檢視，縮放或捲動其中一邊時另一邊會跟著對齊，也可切換成獨立檢視。旋轉 90／270 度的文字會依方向計算框選範圍，垂直拆開的插件位置也會嘗試合併。

記憶體最多保留兩份 PDF 索引；更換檔案後，未使用的舊索引會依最近使用時間自動淘汰、取消背景處理並釋放 PDF。線路圖工具列也可手動清除閒置快取。

掃描型 PDF 或純圖片目前只能使用並排／像素差異功能，無法做文字定位；PDF 與文字索引都只存在瀏覽器記憶體，不會上傳。

## Mac 啟動方式

雙擊 `scripts/start-offline.command`。瀏覽器會自動開啟：

`http://127.0.0.1:3784/`

使用完畢後，關閉瀏覽器頁面，並在終端機視窗按 `Control + C`。

## Windows 啟動方式

雙擊 `scripts\start-offline.bat`。程式會等待本機服務完成啟動，再自動開啟瀏覽器：

`http://127.0.0.1:3784/`

請保持命令提示字元視窗開啟。使用完畢後，在該視窗按 `Ctrl + C` 停止服務。

Mac 的 `node_modules` 不能複製到 Windows 使用；Windows 電腦第一次使用時，必須在專案資料夾重新執行一次 `npm ci`。

## 從 GitHub 下載 Windows 成品（不需要 Node.js）

專案的 GitHub Actions 會在 Windows runner 使用 Node.js 22 完成建置與測試，並提供兩個可下載的 workflow artifacts：

- `BOMLens-Windows-x64-Portable`：內含 `BOMLens-Windows-x64-Portable.zip`，解壓縮後直接雙擊 `BOMLens.exe`，不需管理員權限。
- `BOMLens-Setup-x64`：內含 `BOMLens-Setup-x64.exe`，使用每位使用者的安裝精靈，安裝到目前帳號的 Local AppData，並建立桌面與開始功能表捷徑。

在 GitHub repository 開啟 `Actions` → `Build Windows packages` → 選擇成功的執行紀錄，於頁面最下方下載 Artifacts。Artifacts 保留 30 天。下載後可用同一個 artifact 內的 `SHA256SUMS.txt` 核對檔案完整性。

GitHub runner 需要 Node.js 是為了「建置」；上述 ZIP 與 Setup EXE 都已包含 Windows Node.js runtime，實際使用的公司電腦不需要安裝 Node.js，也不需要連網。

## Windows 正式可攜版（自行在 Windows 建置）

IT 人員在一台 Windows 建置電腦完成 `npm ci` 後，執行：

`npm run package:windows`

程式會產生 `release\BOMLens-Windows-x64-Portable.zip`。將它解壓縮後，使用者直接雙擊 `BOMLens.exe` 即可，不需要另外安裝 Node.js、不需要管理員權限，也不會連上外部網站。程式啟動後會出現在 Windows 通知區，可從通知區選擇「開啟 BOMLens」或「結束」。

若公司有程式碼簽章憑證，可用以下方式建立已簽章版本：

`powershell -ExecutionPolicy Bypass -File scripts\build-windows-portable.ps1 -CertificateThumbprint "憑證指紋"`

可攜版必須在 Windows 上建立，不能直接把 Mac 的套件資料夾複製到 Windows。

## 從原始碼第一次建置

電腦需有 Node.js 22 或更新版本，並需先執行一次 `npm ci` 安裝程式套件。套件準備完成後，日常使用不需要網路。

若公司電腦不允許安裝 Node.js，請直接下載 GitHub Actions 產生的 Portable ZIP 或 Setup EXE，不要在該電腦從原始碼建置。
