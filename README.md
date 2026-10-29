# TW Stock Radar｜台股技術分析儀表板

以 Streamlit、yfinance、Pandas 與 Plotly 製作的台股技術分析工具。可匯入 Google 試算表或手動輸入股票代號，查看技術指標、量化訊號與互動圖表，並下載 Excel 報表。

## 功能

- 台股上市／上櫃代號辨識與公司名稱查詢
- RSI、MACD、KD、布林通道、ADX、SAR、一目均衡表等技術指標
- 個股卡片、互動 K 線、技術指標明細與量化評分
- 匯出含圖表與摘要的 Excel 報表

## 本機執行

需要 Python 3.9 或更新版本。

```bash
python -m venv .venv
```

Windows PowerShell：

```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

macOS / Linux：

```bash
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

啟動後在瀏覽器開啟 Streamlit 顯示的本機網址。

## 部署至 Streamlit Community Cloud

GitHub 儲存庫只保存原始碼；要讓 Python Streamlit 應用在網路上執行，請另外使用 Streamlit Community Cloud：

1. 將此 GitHub 儲存庫連結至 Streamlit Community Cloud。
2. 選擇 `main` 分支及 `app.py` 作為主程式。
3. 部署時使用根目錄的 `requirements.txt` 安裝套件。

## Google 試算表與資料來源

應用程式預設會讀取程式中設定的 Google 試算表 ID。若使用該清單，試算表需設定為「知道連結的任何人皆可檢視」；也可以在側邊欄改用手動輸入股票代號。由於本儲存庫為公開，程式內的預設試算表 ID 也會公開，請勿將含有個人或敏感資料的試算表設為公開。

股價資料來自 Yahoo Finance，台股名稱對照來自臺灣證券交易所公開資料。第三方服務的可用性及資料更新時間可能有所差異。

## 免責聲明

本工具提供技術面量化分析，僅供資訊參考，不構成投資建議或買賣邀約。使用者應自行查證資料並承擔投資決策責任。
