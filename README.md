# US Overseas AI-Compute Controls — Dashboard (Internal Review)

境外取得美國 AI 算力：政策工具資料庫的互動式 dashboard。

> ⚠️ **內部審閱用。目前為公開 repo（obscure 網址 + noindex），非最終付費發行版；請勿轉散佈。**
> Internal review copy — public repo, not for distribution.

## Live
- **Dashboard**: https://ericlutw.github.io/us-overseas-ai-compute-controls/
- 已設 `robots.txt` noindex：搜尋引擎不收錄（知道網址才找得到，非真正私密）。

## Contents
- `index.html` — 互動式 dashboard（自包含單檔，離線可開）
- `data/` — 來源 workbook（Excel，資料真值）

## 更新方式（不用再丟檔案）
1. 重新產生 dashboard，覆蓋本機 `index.html`（及/或更新 `data/`）
2. `git add -A && git commit -m "update" && git push`
3. Pages 自動重建，reviewer 網址不變、重新整理即看到最新版

## Notes
- 現階段 dashboard 為靜態單檔、資料內嵌，**不自動讀取 workbook 更新**（改資料仍需重產 dashboard）。
- 「資料／前端解耦 → 只改後端自動反映」為後續另案。
