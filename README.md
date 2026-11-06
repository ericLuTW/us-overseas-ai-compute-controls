# US Overseas AI-Compute Controls — Dashboard (Internal Review)

境外取得美國 AI 算力：政策工具資料庫的互動式 dashboard。

> ⚠️ **內部審閱用。目前為公開 repo（obscure 網址 + noindex），非最終付費發行版；請勿轉散佈。**
> Internal review copy — public repo, not for distribution.

## Live
- **Dashboard**: https://ericlutw.github.io/us-overseas-ai-compute-controls/
- **改版前舊版**（中英雙語、導讀）：https://ericlutw.github.io/us-overseas-ai-compute-controls/old/2026-11-06_v1-design/
- 已設 `robots.txt` noindex：搜尋引擎不收錄（知道網址才找得到，非真正私密）。

## Contents
- `index.html` — 互動式 dashboard（2026-11-06 起為新前端 v1：制度層級／標的矩陣／供應鏈版圖／政策資料庫四個分頁，中文介面）。資料內嵌於一個 `window.__DATA_ZH` script block；四個子畫板也內嵌在檔內，雙擊即可離線開啟。
- `support.js`、`vendor/react.js`、`vendor/react-dom.js` — Claude Design 的執行環境，必須和 `index.html` 放在同一層。
- `MatrixView.dc.html`、`LayersView.dc.html`、`SupplyView.dc.html`、`PolicyDB.dc.html` — 子畫板原始檔（`bundle` 由這四個檔產生內嵌內容；頁面執行時不需要它們）。
- `old/2026-11-06_v1-design/index.html` — 改版前的舊版單檔（含 EN 介面、導讀、進階矩陣），留存不再更新。
- `data/US-overseas-AI-compute-controls_v2_ZH.xlsx` — 來源 workbook（中文，資料真值）
- `data/US-overseas-AI-compute-controls_v2_EN.xlsx` — 來源 workbook（英文；v1 前端只用中文，EN 留待日後）
- `tools/gen_data_blobs.py` — 由 workbook 重生 dashboard 內嵌資料的產生器（2026-09-21 新增，2026-11-06 改為子命令）
- `tools/statics.json` — 設計期靜態內容（`GLOSS` 66 條縮寫、`nodes` N1–N15、`GRP_EN`），自舊版抽出一次

## 更新資料（workbook 改版時）
1. 把最新 canonical workbook 複製到 `data/`（沿用現有檔名）
2. 灌資料並自檢：
   ```
   python3 tools/gen_data_blobs.py splice data/US-overseas-AI-compute-controls_v2_ZH.xlsx data/US-overseas-AI-compute-controls_v2_EN.xlsx
   python3 tools/gen_data_blobs.py check  data/US-overseas-AI-compute-controls_v2_ZH.xlsx data/US-overseas-AI-compute-controls_v2_EN.xlsx
   ```
   `check` 應印 `[zh] diffs: 0 (allowed 0, outside 0)`、`[en] skipped: not in index.html`，exit 0。
3. `git add -A && git commit -m "update" && git push`
4. Pages 自動重建，reviewer 網址不變、重新整理即看到最新版

## 更新前端（Claude Design 畫布改版時）
1. 在畫布匯出「實作 v1・App」畫板，zip 解開到 `tmp/app_export/`（新的空資料夾）。
2. 把 `App.dc.html` 複製為 `index.html`，四個子畫板、`support.js`、`vendor/` 複製到 repo 根目錄（`assets/` 不用複製）。
3. 依序執行：
   ```
   python3 tools/gen_data_blobs.py inject --asset-id dfdd7576b54cc91759c0ad9b20d3b29a data/US-overseas-AI-compute-controls_v2_ZH.xlsx data/US-overseas-AI-compute-controls_v2_EN.xlsx
   python3 tools/gen_data_blobs.py bundle
   python3 tools/gen_data_blobs.py check  data/US-overseas-AI-compute-controls_v2_ZH.xlsx data/US-overseas-AI-compute-controls_v2_EN.xlsx
   ```
   - `inject`：把匯出檔裡指向畫布資料檔的 `<script src>` 換成內嵌資料（id 為畫布上 `data_zh.js` 的上傳 id；畫布重新上傳資料時 id 會變）。
   - `bundle`：把子畫板內嵌進 `index.html`。不做這步，用雙擊（`file://`）開會只剩頁首與篩選框，因為瀏覽器不允許 `file://` 頁面讀取旁邊的檔案；上 GitHub Pages 或用 `python3 -m http.server` 開則不受影響。可重複執行。
4. 本機檢查：`python3 -m http.server 8765` 後開 `http://localhost:8765/`，四個分頁都要有內容。

## Notes
- 資料現況：對齊 canonical v2（ZH `c65272e5…`／EN `a0b2a05e…`，2026-11-06）＝55 政策／144 工具，含 BL 基線制度與制度層級（`LAYERS`）。
- `LAYERS`（制度層級）一律由 ZH workbook 計算：BL 條目〈法源〉→ 依據關係；BL 條目寫「載於 …」→ 修訂關係；〈參照表〉引用碼 → 引用關係。
- **BL 呈現規則**（2026-09-24）：canonical 中 node_id=0 的 BL 框架性工具，產生器展開至全 15 節點呈現（全鏈適用）；有明確節點映射的 BL 工具（3B001→N2、4A003/3A001→N7、軍事終端用途→N15）只落自己的節點。非 BL 的 node_id=0 工具維持不入圖。
- 前端（UI／版面）由 Claude Design 產出；產生器只更新資料層並內嵌子畫板，不改前端程式。
