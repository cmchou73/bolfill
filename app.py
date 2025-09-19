import io, zipfile, datetime
import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter

st.set_page_config(page_title="BOL 批次填寫", page_icon="📦")
st.title("📦 BOL 批次填寫工具")
st.caption("上傳 Excel → 自動灌進 PDF 表單 → 下載 ZIP")

DEFAULT_TEMPLATE_PATH = "templates/bill-of-lading-01.pdf"

# ===== 欄位對映（請按你的 Excel 欄名調整） =====
FIELD_MAP = {
    "BOLnum": "BOLnum",
    "FromName": "FromName", "FromAddr": "FromAddr", "FromCityStateZip": "FromCityStateZip",
    "FromSIDNum": "FromSIDNum", "FromFOB": "FromFOB",
    "ToName": "ToName", "ToAddress": "ToAddress", "ToCityStateZip": "ToCityStateZip",
    "ToLocNum": "ToLocNum", "ToCID": "ToCID", "ToFOB": "ToFOB",
    "CarrierName": "CarrierName", "SCAC": "SCAC", "PRO": "PRO",
    "TrailerNum": "TrailerNum", "SealNum": "SealNum",
    "BillName": "BillName", "BillAddress": "BillAddress",
    "BillCityStateZip": "BillCityStateZip", "BillInstructions": "BillInstructions",
    "PrePaid": "PrePaid", "Collect": "Collect", "3rdParty": "3rdParty",
    "MasterBOL": "MasterBOL", "Date": "Date",
}
for i in range(1, 9):
    FIELD_MAP[f"OrderNum{i}"] = f"OrderNum{i}"
    FIELD_MAP[f"NumPkgs{i}"] = f"NumPkgs{i}"
    FIELD_MAP[f"Weight{i}"] = f"Weight{i}"
    FIELD_MAP[f"Pallet{i}"] = f"Pallet{i}"
    FIELD_MAP[f"AddInfo{i}"] = f"AddInfo{i}"
for i in range(1, 9):
    FIELD_MAP[f"HU_QTY_{i}"] = f"HU_QTY_{i}"
    FIELD_MAP[f"HU_Type_{i}"] = f"HU_Type_{i}"
    FIELD_MAP[f"Pkg_QTY_{i}"] = f"Pkg_QTY_{i}"
    FIELD_MAP[f"Pkg_Type_{i}"] = f"Pkg_Type_{i}"
    FIELD_MAP[f"WT_{i}"] = f"WT_{i}"
    FIELD_MAP[f"HM_{i}"] = f"HM_{i}"
    FIELD_MAP[f"Desc_{i}"] = f"Desc_{i}"
    FIELD_MAP[f"NMFC{i}"] = f"NMFC{i}"
    FIELD_MAP[f"Class{i}"] = f"Class{i}"

CHECKBOX_FIELDS = {
    "FromFOB","ToFOB","PrePaid","Collect","3rdParty","MasterBOL",
    "Pallet1","Pallet2","Pallet3","Pallet4","Pallet5","Pallet6","Pallet7","Pallet8"
}

def set_need_appearances(writer: PdfWriter):
    if "/AcroForm" in writer._root_object:
        writer._root_object["/AcroForm"].update({"/NeedAppearances": True})

def fill_one(template_bytes: bytes, row: pd.Series, idx: int) -> bytes:
    reader = PdfReader(io.BytesIO(template_bytes))
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)
    set_need_appearances(writer)

    for excel_col, pdf_field in FIELD_MAP.items():
        if excel_col in row and pd.notna(row[excel_col]):
            val = str(row[excel_col]).strip()
            if pdf_field in CHECKBOX_FIELDS:
                on = val.lower() in {"y","yes","true","1","✔","✓"}
                writer.update_page_form_field_values(writer.pages[0], {pdf_field: "Yes" if on else "Off"})
            else:
                writer.update_page_form_field_values(writer.pages[0], {pdf_field: val})

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

def extract_field_names(template_bytes: bytes):
    # 回傳此 PDF 第一頁所有表單欄位名稱（/T）方便對映
    names = []
    r = PdfReader(io.BytesIO(template_bytes))
    page = r.pages[0]
    if "/Annots" in page:
        for annot in page["/Annots"]:
            obj = annot.get_object()
            if "/T" in obj:
                names.append(str(obj["/T"]))
    return sorted(set(names))

# === UI ===
with st.sidebar:
    st.markdown("### 模板來源")
    use_uploaded_template = st.toggle("改用上傳 PDF 模板", value=False)
    st.caption("未開啟則使用 repo 內建 `templates/bill-of-lading-01.pdf`")

if use_uploaded_template:
    tmpl_file = st.file_uploader("上傳 BOL 空白 PDF 模板", type=["pdf"])
else:
    tmpl_file = None

excel_file = st.file_uploader("上傳 Excel（*.xlsx）", type=["xlsx"])

# 顯示模板欄位名（協助對映）
with st.expander("🔎 檢視模板的表單欄位名稱（對映用）"):
    source_bytes = None
    try:
        if tmpl_file is not None:
            source_bytes = tmpl_file.read()
        else:
            with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
                source_bytes = f.read()
    except FileNotFoundError:
        st.warning("找不到預設模板（templates/bill-of-lading-01.pdf）。請改用上傳模板模式。")

    if source_bytes:
        names = extract_field_names(source_bytes)
        if names:
            st.write("找到以下欄位名稱（請在 FIELD_MAP 對應你的 Excel 欄名）：")
            st.code("\n".join(names))
        else:
            st.info("此 PDF 未偵測到可填寫的表單欄位。請確認模板是否為可填式 PDF。")

disabled = excel_file is None
if st.button("開始生成", type="primary", disabled=disabled):
    try:
        # 取得模板 bytes
        if tmpl_file is not None:
            template_bytes = tmpl_file.getvalue()
        else:
            with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
                template_bytes = f.read()

        # 讀取 Excel
        df = pd.read_excel(excel_file)

        # 逐列生成 PDF -> 打包 ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, row in df.iterrows():
                pdf_bytes = fill_one(template_bytes, row, i)
                out_name = str(row.get("BOLnum") or f"row{i+1}").strip()
                safe_name = "".join(c for c in out_name if c not in '\\/:*?"<>|').strip() or f"row{i+1}"
                zf.writestr(f"BOL_{safe_name}.pdf", pdf_bytes)

        zip_buf.seek(0)
        now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        st.success(f"生成完成，共 {len(df)} 份。")
        st.download_button(
            label="下載 ZIP",
            data=zip_buf,
            file_name=f"BOL_PDFs_{now}.zip",
            mime="application/zip",
        )
    except FileNotFoundError as e:
        st.error(f"找不到預設模板：{e}")
    except Exception as e:
        st.error("處理失敗，請檢查 Excel 欄名與 PDF 欄位是否對應。")
        st.exception(e)

st.markdown("---")
st.markdown("**提示**：若欄位對不上，請展開上方「檢視模板的表單欄位名稱」，把偵測到的 PDF 欄位名貼回 `FIELD_MAP` 對應到你的 Excel 欄名。")
