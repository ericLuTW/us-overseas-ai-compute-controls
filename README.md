# US Overseas AI-Compute Controls — Dashboard (Internal Review)

境外取得美國 AI 算力：政策工具資料庫的互動式 dashboard。

> ⚠️ **內部審閱用，非公開發行版本。請勿轉散佈。**
> Internal review copy — not for distribution.

## Contents
- `index.html` — 互動式 dashboard（自包含單檔，離線可開）
- `data/` — 來源 workbook（Excel，資料真值）

## Live (internal)
- Pages URL: _(建立後填入)_
- 僅供知道網址者審閱；已設 `robots.txt` noindex 阻擋搜尋引擎收錄（obscurity，非真正私密）。

## Notes
- 現階段 dashboard 為靜態單檔、資料內嵌，**不自動讀取 workbook 更新**。
- 更新方式：改後重新產生 dashboard、覆蓋 `index.html` 再推即可（reviewer 網址不變、看到最新版）。
- 「資料／前端解耦 → 只改後端自動反映」為後續另案。
