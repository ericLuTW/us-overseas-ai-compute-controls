# US Overseas AI-Compute Controls — Dashboard (Internal Review)

境外取得美國 AI 算力：政策工具資料庫的互動式 dashboard。

> ⚠️ **內部審閱用。目前為公開 repo（obscure 網址 + noindex），非最終付費發行版；請勿轉散佈。**
> Internal review copy — public repo, not for distribution.

## Live
- **Dashboard**: https://ericlutw.github.io/us-overseas-ai-compute-controls/
- 已設 `robots.txt` noindex：搜尋引擎不收錄（知道網址才找得到，非真正私密）。

## Contents
- `index.html` — 互動式 dashboard（自包含單檔，離線可開；資料內嵌於 `window.__DATA_ZH/__DATA_EN` 兩個 script block）
- `data/US-overseas-AI-compute-controls_v2_ZH.xlsx` — 來源 workbook（中文，資料真值）
- `data/US-overseas-AI-compute-controls_v2_EN.xlsx` — 來源 workbook（英文）
- `tools/gen_data_blobs.py` — 由 workbook 重生 dashboard 內嵌資料的產生器（2026-09-21 新增）

## 更新方式
1. 把最新 canonical workbook 複製到 `data/`（沿用現有檔名）
2. 回歸自檢＋灌資料：
   ```
   python3 tools/gen_data_blobs.py splice data/US-overseas-AI-compute-controls_v2_ZH.xlsx data/US-overseas-AI-compute-controls_v2_EN.xlsx
   ```
   （`check` 模式可先驗證產生器對現行 blob 重現度；splice 後再跑 `check` 應為 0 diff）
3. `git add -A && git commit -m "update" && git push`
4. Pages 自動重建，reviewer 網址不變、重新整理即看到最新版

## Notes
- 資料現況：對齊 canonical **v2.3**（ZH `f0693e9b…`／EN `0d7966ae…`，2026-09-21）＝55 政策／144 工具，含 BL 基線制度分區。
- `GLOSS`（66 條縮寫詞彙）、`nodes`（N1–N15 顯示名與分組）、`GRP_EN` 為設計期靜態內容，產生器自現行 `index.html` 沿用；要改這三者需直接編輯。
- workbook 的〈參照表〉分頁（112 列 referent 查詢層）目前**不在 dashboard 呈現範圍**（前端無對應版面，屬日後功能）。
- 前端（UI／版面）仍由 Claude Design 產出；本產生器只更新資料層，不動前端程式。
