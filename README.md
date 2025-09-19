# Streamlit BOL 批次填寫工具

## 快速開始（本機）
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 使用方式
1. 開啟網頁後，於主畫面上傳 Excel（.xlsx）。
2. 若不使用內建模板，請於側邊欄開啟「改用上傳 PDF 模板」並上傳你的空白 BOL。
3. 按「開始生成」取得打包的 ZIP（內含多份已填好的 BOL PDF）。

## 對映說明
- 若生成的欄位沒填上，請展開頁面中的「🔎 檢視模板的表單欄位名稱」，把列出的 PDF 欄位名對應到 `app.py` 的 `FIELD_MAP`。
