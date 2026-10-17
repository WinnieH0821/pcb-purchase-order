from pathlib import Path
from datetime import date, datetime
from zoneinfo import ZoneInfo
import re
import io

import streamlit as st
from docx import Document
from openpyxl import load_workbook


def normalize_quantity(value):
    """正規化數量文字；簡單數字可比較數值，其他格式比較去空白後的文字。"""
    text = str(value or "").strip().replace(",", "").replace("，", "")
    text = re.sub(r"\s+", " ", text).casefold()
    simple = re.fullmatch(r"([0-9]+(?:\.[0-9]+)?)\s*(pcs|pc|片|set|套|個|件)?", text)
    if simple:
        number = float(simple.group(1))
        unit = simple.group(2) or ""
        return ("simple", number, unit)
    return ("text", re.sub(r"\s+", "", text))


def quantities_match(request_value, order_value):
    """比較申請數量與實際採購數量；不確定的複合格式採保守文字比對。"""
    request_text = str(request_value or "").strip()
    order_text = str(order_value or "").strip()
    if not request_text or not order_text:
        return False

    request_norm = normalize_quantity(request_text)
    order_norm = normalize_quantity(order_text)
    if request_norm[0] == "simple" and order_norm[0] == "simple":
        _, request_number, request_unit = request_norm
        _, order_number, order_unit = order_norm
        # 若只有其中一邊填了單位，允許比對數值；兩邊單位都有填時必須一致。
        units_match = not request_unit or not order_unit or request_unit == order_unit
        return request_number == order_number and units_match
    return request_norm == order_norm

# 供應商主檔：目前先建立第一家；後續供應商可沿用相同欄位結構新增。
SUPPLIERS = {
    "比斯捷科技(深圳)有限公司": {
        "contact": "熊小輝",
        "phone": "137 14444 4149",
        "code": "901",
    },
}


def apply_supplier_code(part_number, supplier_name):
    """只移除已建檔的供應商代號；保留 3403105-EXP 這類本身屬於料號的數字前綴。"""
    pn = str(part_number or "").strip()
    if not pn:
        return ""
    # 僅移除供應商主檔中已知的公司代號，不把任意「數字-」誤認成公司代號。
    known_codes = {
        str(info.get("code", "")).strip()
        for info in SUPPLIERS.values()
        if str(info.get("code", "")).strip()
    }
    for known_code in sorted(known_codes, key=len, reverse=True):
        pn = re.sub(rf"^{re.escape(known_code)}-", "", pn)
    supplier_info = SUPPLIERS.get(str(supplier_name or "").strip(), {})
    code = str(supplier_info.get("code", "")).strip()
    if not code:
        return pn
    return f"{code}-{pn}"

def on_supplier_change():
    supplier_name = str(st.session_state.get("supplier", "") or "").strip()
    supplier_info = SUPPLIERS.get(supplier_name, {})
    st.session_state["contact"] = supplier_info.get("contact", "")
    st.session_state["phone"] = supplier_info.get("phone", "")
    st.session_state["customer_pn"] = apply_supplier_code(
        st.session_state.get("customer_pn", ""), supplier_name
    )


st.set_page_config(page_title="PCB 採購單產生器", page_icon="📄", layout="wide")
st.title("PCB 採購單自動產生器")
st.caption("申請資料支援 Word / Excel；輸出一律為 Word。輸出前請人工核對料號、數量與交期。")

DEFAULT_TEMPLATE = Path(__file__).with_name("採購單範本.docx")


def docx_to_text(uploaded):
    doc = Document(uploaded)
    lines = []

    for p in doc.paragraphs:
        if p.text.strip():
            lines.append(p.text.strip())

    for table_index, table in enumerate(doc.tables, start=1):
        lines.append(f"[Word 表格 {table_index}]")
        for row in table.rows:
            values = []
            seen_cells = set()
            for cell in row.cells:
                key = cell._tc
                if key in seen_cells:
                    continue
                seen_cells.add(key)
                value = " ".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                values.append(value)
            if any(values):
                lines.append(" | ".join(values))

    for section in doc.sections:
        for p in section.header.paragraphs:
            if p.text.strip():
                lines.append("[頁首] " + p.text.strip())
        for p in section.footer.paragraphs:
            if p.text.strip():
                lines.append("[頁尾] " + p.text.strip())

    return "\n".join(lines)


def xlsx_to_text(uploaded):
    # openpyxl 可讀取儲存格文字與數值；Excel 中的浮動圖形/部分繪圖物件不會被轉成文字。
    wb = load_workbook(uploaded, data_only=True, read_only=True)
    lines = []
    for ws in wb.worksheets:
        lines.append(f"===== 工作表：{ws.title} =====")
        for row in ws.iter_rows():
            values = []
            for cell in row:
                value = cell.value
                if value is None:
                    values.append("")
                else:
                    values.append(str(value).strip())
            while values and not values[-1]:
                values.pop()
            if any(values):
                lines.append(" | ".join(values))
    wb.close()
    return "\n".join(lines)



def extract_labeled_value(text, labels, value_pattern=r"([^|\n]+)"):
    """逐列從「標籤 | 值」格式擷取欄位，避免跨列誤抓到其他文字。"""
    label_pattern = "|".join(re.escape(label) for label in sorted(labels, key=len, reverse=True))
    pattern = rf"(?:{label_pattern})[ \t]*[:：]?[ \t]*\|[ \t]*{value_pattern}"
    for line in text.splitlines():
        match = re.search(pattern, line, re.I)
        if match:
            value = match.group(1).strip(" ：:|")
            if value:
                return value
    return ""


def extract_request_fields_by_rows(text):
    """針對實際 PCB 製作申請單常見的「欄位 | 值」表格列擷取資料。"""
    found = {}
    row_rules = {
        # 僅使用明確的料號標籤；訂單編號不視為料號。
        "客戶料號": ["產品料號", "客戶料號", "客戶產品 P/N", "客戶產品 PN", "Part Number", "Part No.", "P/N", "Model Number", "Product Code", "料號"],
        "產品名稱": ["產品名稱", "產品名", "品名", "Product Name", "Item Description", "Description"],
        "需求數量": ["PCB需求量", "PCB 需求量", "PCB需求數量", "採購數量", "訂購數量", "需求數量", "需求量", "數量", "Qty", "Quantity", "Order Quantity", "Required Qty"],
        "專案編號": ["專案編號", "Project No", "Project Number", "Project"],
        "版本": ["版本Version", "產品版本", "版本", "Version", "Revision", "Rev."],
        "PCB層數": ["PCB層數", "PCB 層數", "層數", "Layer Count", "Layers"],
        "PCB厚度": ["PCB厚度", "PCB 厚度", "板厚", "厚度", "Board Thickness", "Thickness"],
        "提出日期": ["提出日期", "申請日期", "需求提出日", "Request Date", "Issue Date"],
        "需求日期": ["需求日期", "要求交期", "需求交期", "交貨日期", "交期", "Due Date", "Delivery Date", "Required Date"],
    }
    for key, labels in row_rules.items():
        value = extract_labeled_value(text, labels)
        if value:
            # 需求數量只擷取前段數量，保留 PCS/SET 單位。
            if key == "需求數量":
                q = re.match(r"\s*([0-9][0-9,]*(?:\.[0-9]+)?\s*(?:PCS|PC|片|SET|套)?)", value, re.I)
                if q:
                    value = q.group(1).strip()
            found[key] = value
    return found


def extract_pcb_process_fields(uploaded):
    """專門讀取 PCB 製程流程表；只把能可靠判定的欄位自動帶入。"""
    result = {}
    wb = load_workbook(io.BytesIO(uploaded.getvalue()), data_only=True, read_only=False)
    try:
        # 正面 D3 與反面 M1 都是此版工單的料號欄位；優先使用有標籤對應的值。
        for sheet_name, cell_ref in [("流程表正面(23)", "D3"), ("流程表反面", "M1")]:
            if sheet_name in wb.sheetnames:
                value = wb[sheet_name][cell_ref].value
                if value is not None:
                    value = str(value).strip()
                    if value and re.fullmatch(r"[A-Z0-9][A-Z0-9._/-]*", value, re.I):
                        result["客戶料號"] = value
                        result["客戶料號來源"] = f"{sheet_name}!{cell_ref}"
                        break

        # S3 的 112 PCS 與 S2 的 6 PNL 是流程表上的製程/排版資訊，
        # 無法單憑該欄位確認為本次採購需求量，因此不填入需求數量或採購數量。
        if "流程表正面(23)" in wb.sheetnames:
            ws = wb["流程表正面(23)"]
            pnl = ws["S2"].value
            pcs = ws["S3"].value
            if pnl is not None or pcs is not None:
                candidates = []
                if pnl is not None and str(pnl).strip():
                    candidates.append(f"S2={str(pnl).strip()}")
                if pcs is not None and str(pcs).strip():
                    candidates.append(f"S3={str(pcs).strip()}")
                result["數量候選（需人工確認）"] = "；".join(candidates)
                result["數量辨識狀態"] = "未自動填入：工單中的排版/PCS 欄位不一定等於採購需求量"
    finally:
        wb.close()
    return result


def extract_request(uploaded):
    if uploaded is None:
        return {}

    suffix = Path(uploaded.name).suffix.lower()
    try:
        if suffix == ".docx":
            text = docx_to_text(uploaded)
            source_type = "Word"
        elif suffix == ".xlsx":
            text = xlsx_to_text(uploaded)
            source_type = "Excel"
        else:
            return {"_來源": "不支援的檔案格式", "_全文": ""}
    except Exception as exc:
        return {"_來源": suffix, "_讀取錯誤": str(exc), "_全文": ""}

    # 只有標籤明確且內容符合常見格式時才自動填入；不確定的欄位留給使用者人工輸入。
    compact = re.sub(r"[ \t]+", " ", text)
    patterns = {
        # Word 申請單的表格可能會重複標籤，規則允許跳過重複標籤後再取值。
        # 支援一般申請單的「客戶料號／產品料號」，也支援本次 Word 檢查表中的「料號：W00G-F1100TPA4」。
        # 「料號」必須緊接冒號或表格分隔符，避免誤抓內文中的「加料號&版本」。
        "客戶料號": r"(?:客戶產品\s*P/?N|產品料號|客戶料號|Part\s*(?:Number|No\.?)|P/N|料號|Model\s*Number|Product\s*Code)(?:\s*\|\s*(?:客戶產品\s*P/?N|產品料號|客戶料號|Part\s*(?:Number|No\.?)|P/N|料號|Model\s*Number|Product\s*Code))*\s*[:：|]\s*([A-Z0-9][A-Z0-9._/-]{3,})",
        "需求數量": r"(?:PCB\s*需求量|PCB\s*需求數量|採購數量|訂購數量|需求數量|需求量|數量|Qty|Quantity|Order\s*Quantity|Required\s*Qty)(?:\s*\|\s*(?:PCB\s*需求量|PCB\s*需求數量|採購數量|訂購數量|需求數量|需求量|數量|Qty|Quantity|Order\s*Quantity|Required\s*Qty))*\s*[:：|]\s*([0-9,]+(?:\.[0-9]+)?\s*(?:PCS|PC|片|SET|套|個|件)?)",
        "產品名稱": r"(?:產品名稱|產品名|品名|Product\s*Name|Item\s*Description|Description)(?:\s*\|\s*(?:產品名稱|產品名|品名|Product\s*Name|Item\s*Description|Description))*\s*[:：|]?\s*([^|\n]{2,80})",
        "專案編號": r"(?:專案編號|Project\s*(?:No\.?|Number)?)(?:\s*\|\s*(?:專案編號|Project\s*(?:No\.?|Number)?))*\s*[:：|]?\s*([A-Z0-9_-]+)",
        "版本": r"(?:產品版本|版本(?:\s*Version)?|Version|Revision|Rev\.?)(?:\s*\|\s*(?:產品版本|版本(?:\s*Version)?|Version|Revision|Rev\.?))*\s*[:：|]?\s*([A-Z0-9._-]+)",
        "PCB層數": r"(?:PCB\s*層數|層數|Layer\s*Count|Layers)(?:\s*\|\s*(?:PCB\s*層數|層數|Layer\s*Count|Layers))*\s*[:：|]?\s*([0-9]+\s*(?:層|L)?)",
        "PCB厚度": r"(?:PCB\s*厚度|板厚|厚度|Board\s*Thickness|Thickness)(?:\s*\|\s*(?:PCB\s*厚度|板厚|厚度|Board\s*Thickness|Thickness))*\s*[:：|]?\s*([0-9.±+\- ]{1,30}(?:mm|MM|公厘)?)",
        "提出日期": r"(?:提出日期|申請日期|需求提出日|Request\s*Date|Issue\s*Date)(?:\s*\|\s*(?:提出日期|申請日期|需求提出日|Request\s*Date|Issue\s*Date))*\s*[:：|]?\s*([0-9]{2,4}[./-][0-9]{1,2}[./-][0-9]{1,4})",
        "需求日期": r"(?:需求日期|要求交期|需求交期|交貨日期|交期|Due\s*Date|Delivery\s*Date|Required\s*Date)(?:\s*\|\s*(?:需求日期|要求交期|需求交期|交貨日期|交期|Due\s*Date|Delivery\s*Date|Required\s*Date))*\s*[:：|]?\s*([0-9]{2,4}[./-][0-9]{1,2}[./-][0-9]{1,4})",
    }
    found = {"_來源": source_type}
    for key, pattern in patterns.items():
        match = re.search(pattern, compact, re.I)
        if match:
            found[key] = match.group(1).strip(" :：|")

    # 先以實際表格列格式補強辨識；若成功，優先採用標籤和值成對的結果。
    row_fields = extract_request_fields_by_rows(text)
    for key, value in row_fields.items():
        if value:
            found[key] = value

    # 建立料號候選清單與辨識依據，避免文件有多個編號時只呈現單一猜測。
    pn_labels = r"(?:客戶產品\s*P/?N|產品料號|客戶料號|Part\s*(?:Number|No\.?)|P/N|料號|Model\s*Number|Product\s*Code)"
    pn_candidates = []
    for line_no, line in enumerate(text.splitlines(), start=1):
        m = re.search(pn_labels + r"\s*[:：|]\s*([A-Z0-9][A-Z0-9._/-]{3,})", line, re.I)
        if m:
            label = re.search(pn_labels, line, re.I).group(0)
            value = m.group(1).strip()
            confidence = "高" if re.search(r"料號|Part\s*(?:Number|No\.?)|P/N|Product\s*Code", label, re.I) else "中（請核對是否為客戶料號）"
            item = {"候選值": value, "欄位標籤": label, "來源行": line_no, "信心提示": confidence}
            if item not in pn_candidates:
                pn_candidates.append(item)
    if pn_candidates:
        found["料號候選清單"] = pn_candidates[:20]
        if not found.get("客戶料號"):
            found["辨識提醒"] = "找到可能的料號候選值，請核對來源欄位後再使用。"

    # 專用流程表採用有來源座標的規則，覆蓋一般文字猜測結果。
    # 這類工單的 PCS/PNL 不能直接當作採購需求量，避免把製程資訊誤填為採購數量。
    if suffix == ".xlsx":
        process_fields = extract_pcb_process_fields(uploaded)
        for key, value in process_fields.items():
            if value:
                found[key] = value
        if process_fields.get("數量辨識狀態"):
            found.pop("需求數量", None)

    # 部分 PCB Layout 檢查表只有料號，沒有本次下單數量；明確標示缺少數量，不從層數、專案編號或版本號猜數量。
    if suffix == ".docx" and found.get("客戶料號") and not found.get("需求數量"):
        found["數量辨識狀態"] = "文件中未找到明確的需求／採購數量，請在採購單欄位手動輸入本次下單量。"

    # 依常見 PCB 申請單標籤補充可安全辨識的規格欄位。
    if re.search(r"內、外層成品銅厚[^\n]*■\s*1\s*oz[^\n]*外層\s*1\s*oz", compact, re.I):
        found["銅厚"] = "內層 1 oz / 外層 1 oz"
    elif re.search(r"內、外層成品銅厚[^\n]*外層\s*1\s*oz", compact, re.I):
        found["銅厚"] = "外層 1 oz（內層請人工確認）"

    if re.search(r"防焊[^\n]*■\s*綠色[^\n]*文字[^\n]*■\s*白色", compact, re.I):
        found["防焊文字"] = "綠色防焊 / 白色文字"

    # 常見 PCB 規格：僅在標籤與值相鄰時擷取，避免將一般文字中的數字當規格。
    extra_specs = {
        "表面處理": r"(?:表面處理|表面加工|Surface\s*Finish)\s*[:：|]\s*([^|\n]{2,60})",
        "防焊顏色": r"(?:防焊顏色|阻焊顏色|Solder\s*Mask(?:\s*Color)?)\s*[:：|]\s*([^|\n]{2,40})",
        "文字顏色": r"(?:文字顏色|字符顏色|Silkscreen(?:\s*Color)?)\s*[:：|]\s*([^|\n]{2,40})",
        "板材": r"(?:板材|基材|Material)\s*[:：|]\s*([^|\n]{2,50})",
        "阻抗要求": r"(?:阻抗要求|特性阻抗|Impedance)\s*[:：|]\s*([^|\n]{2,60})",
    }
    for spec_name, spec_pattern in extra_specs.items():
        spec_match = re.search(spec_pattern, compact, re.I)
        if spec_match:
            found[spec_name] = spec_match.group(1).strip(" ：:|")

    if "PCB層數" in found or "PCB厚度" in found:
        parts = ["FR-4"]
        if found.get("PCB層數"):
            layer = found["PCB層數"].strip()
            if re.fullmatch(r"\d+", layer):
                layer += "L"
            parts.append(layer)
        if found.get("PCB厚度"):
            parts.append("板厚 " + found["PCB厚度"].strip())
        found["材料板厚"] = " / ".join(parts)

    special = re.search(r"特殊需求附加說明\s*\|\s*(.+?)(?=\n|$)", text, re.I)
    if special:
        found["特殊需求"] = re.sub(r"\s*/\s*", "；", special.group(1)).strip(" |")

    found["_全文"] = text
    return found


def set_paragraph_text(paragraph, value):
    if paragraph.runs:
        paragraph.runs[0].text = str(value)
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(str(value))


def replace_paragraph(paragraph, replacements):
    original = paragraph.text
    updated = original
    for old, new in replacements.items():
        if old and old in updated:
            updated = updated.replace(old, str(new))
    if updated != original:
        set_paragraph_text(paragraph, updated)


def replace_doc_text(doc, replacements, delivery_value):
    # 交貨地點欄位獨立處理，避免改到頁首公司抬頭。
    for table in doc.tables:
        for row in table.rows:
            cells = []
            seen_row = set()
            for cell in row.cells:
                key = cell._tc
                if key not in seen_row:
                    seen_row.add(key)
                    cells.append(cell)
            if any("交貨地點" in cell.text for cell in cells):
                # 交貨地點與付款方式位於同一個合併儲存格，僅在這一列替換公司名稱，避免改到頁首抬頭。
                for cell in cells:
                    if "交貨地點" not in cell.text and "銘星科技有限公司" in cell.text:
                        for paragraph in cell.paragraphs:
                            replace_paragraph(paragraph, {**replacements, "銘星科技有限公司": delivery_value})

    for paragraph in doc.paragraphs:
        replace_paragraph(paragraph, replacements)

    for table in doc.tables:
        seen = set()
        for row in table.rows:
            for cell in row.cells:
                key = cell._tc
                if key in seen:
                    continue
                seen.add(key)
                for paragraph in cell.paragraphs:
                    replace_paragraph(paragraph, replacements)

    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            replace_paragraph(paragraph, replacements)
        for paragraph in section.footer.paragraphs:
            replace_paragraph(paragraph, replacements)


st.subheader("1. 上傳本次 PCB 申請資料")
request_file = st.file_uploader("申請資料（Word .docx 或 Excel .xlsx）", type=["docx", "xlsx"])
extracted = extract_request(request_file) if request_file else {}

if extracted:
    display_fields = {key: value for key, value in extracted.items() if not key.startswith("_")}
    st.write(f"檔案格式：{extracted.get('_來源', '未知')}")
    if extracted.get("_讀取錯誤"):
        st.error("讀取檔案失敗：" + extracted["_讀取錯誤"])
    st.write("以下辨識結果僅供參考，請依原始申請資料人工確認：")
    st.json(display_fields if display_fields else {"提示": "沒有找到足夠明確的欄位。請以原始檔案為準，並在下方人工填寫。"})
    if extracted.get("料號候選清單"):
        st.markdown("#### 料號候選與辨識依據")
        st.caption("只把明確標示為料號、Part Number、P/N 等欄位當作料號；訂單號碼不會自動當成料號。")
        st.dataframe(extracted["料號候選清單"], use_container_width=True, hide_index=True)
    if extracted.get("數量辨識狀態"):
        if extracted.get("_來源") == "Excel":
            st.warning("這份 Excel 是 PCB 製程流程表。系統會顯示找到的 PNL／PCS 數字供參考，但不會直接當成採購數量；請依正式需求確認本次下單量。")
        else:
            st.warning("已讀取文件內容，但文件中沒有可確認的需求／採購數量。請在下方手動輸入本次實際下單量；系統不會自行猜測數量。")
    with st.expander("查看讀取到的文字內容"):
        st.text(extracted.get("_全文", "")[:30000])
        if len(extracted.get("_全文", "")) > 30000:
            st.caption("內容過長，畫面僅顯示前 30,000 個字。")

# 檔案更換時，將辨識結果同步到第二區塊的 widget state。
# 只在上傳檔案變更時同步，避免每次重新整理都覆蓋使用者手動修改的內容。
if request_file is not None:
    upload_signature = (request_file.name, request_file.size, hash(request_file.getvalue()))
else:
    upload_signature = None

if "_last_request_signature" not in st.session_state or upload_signature != st.session_state.get("_last_request_signature"):
    st.session_state["_last_request_signature"] = upload_signature
    ambiguous_process_qty = bool(extracted.get("數量辨識狀態"))
    auto_fields = {
        "customer_pn": extracted.get("客戶料號", ""),
        "product_name": extracted.get("產品名稱", "印刷電路板"),
        # 採購數量是本次下單量，不應直接沿用製程表上的 PNL/PCS 數字。
        "qty": "" if ambiguous_process_qty else extracted.get("需求數量", ""),
        "request_qty": "" if ambiguous_process_qty else extracted.get("需求數量", ""),
        "due": extracted.get("需求日期", ""),
        "material": extracted.get("材料板厚", "FR-4 TG150 / 6L / 板厚 1.6mm"),
        "copper": extracted.get("銅厚", "內層 1 oz / 外層 1 oz"),
        "solder": extracted.get("防焊文字", "亮面綠色 / 文字白色"),
        "production": extracted.get(
            "特殊需求",
            "請完全依照銘星提供之資料製作；請注意板件上不能有 Made In Taiwan 字樣"
        ),
        # 供應商相關欄位每次由使用者自行填寫；保留本公司交貨地點。
        "supplier": "比斯捷科技(深圳)有限公司",
        "contact": "熊小輝",
        "phone": "137 14444 4149",
        "price": "打包價 USD$450",
        "payment": "月結 60 天",
        "delivery": "銘星科技有限公司",
        "finish": '鍍金厚度：足 2 μ" (MIN)；鍍鎳厚度：足 120 μ" (MIN)',
        "inspection": "IPC 2 級",
        "notes": "打叉方式：黑筆雙面打叉，4 UP 允收單報，訂單總數 5% 內（此批打樣不收單報）；測試章：黑色 ET 章，有板邊的蓋在板邊料號旁，沒板邊的蓋在週期旁",
        "packaging": "透明真空袋＋濕度卡（放側邊）＋乾燥包 6g（放側邊）",
        "internal_pn": "M056D6026A2",
        "maker": "Sally 劉曼姣",
    }
    for field_key, field_value in auto_fields.items():
        st.session_state[field_key] = field_value
    st.session_state["customer_pn"] = apply_supplier_code(
        st.session_state.get("customer_pn", ""), st.session_state.get("supplier", "")
    )

# 將既有工作階段的舊預設值更新一次，並保留使用者已修改的銘星料號。
if st.session_state.get("maker") == "劉曼姣":
    st.session_state["maker"] = "Sally 劉曼姣"
if not st.session_state.get("internal_pn"):
    st.session_state["internal_pn"] = "M056D6026A2"

st.subheader("2. 每次重新輸入採購單資料")
taiwan_today = datetime.now(ZoneInfo("Asia/Taipei")).date()
current_order_prefix = f"MH{taiwan_today.strftime('%y')}A-{taiwan_today.strftime('%m')}-"
previous_default = st.session_state.get("_last_order_no_default")
current_order_value = st.session_state.get("order_no")
if (
    "order_no" not in st.session_state
    or current_order_value == previous_default
    or re.fullmatch(r"MH\d{2}A-\d{2}-XXX", str(current_order_value or ""))
):
    st.session_state["order_no"] = current_order_prefix
st.session_state["_last_order_no_default"] = current_order_prefix

st.info(f"訂單編號會依目前月份自動帶入前綴：{current_order_prefix}；請直接在同一欄位最後輸入流水號，例如 {current_order_prefix}035。客戶料號會依供應商代號自動補上前綴。")

st.markdown("#### 供應商與客戶料號")
supplier = st.selectbox(
    "廠商名稱 *（可下拉選擇，也可直接輸入新名稱）",
    options=list(SUPPLIERS.keys()),
    accept_new_options=True,
    key="supplier",
    on_change=on_supplier_change,
    help="目前已建檔：比斯捷科技(深圳)有限公司；選取後會自動帶入聯絡人、電話與已設定的公司代號。",
)
contact_col, phone_col = st.columns(2)
with contact_col:
    contact = st.text_input("聯絡人 *", key="contact")
with phone_col:
    phone = st.text_input("聯絡電話 *", key="phone")
customer_pn = st.text_input("客戶料號（選擇供應商後自動加上公司代號，可手動修改）*", key="customer_pn")

with st.form("po_form"):
    left, right = st.columns(2)
    with left:
        order_no_input = st.text_input("訂單號碼 *（月份自動帶入，流水號自行輸入）", key="order_no", help="月份會依目前日期自動更新；請在最後的連字號後輸入流水號，例如 MH26A-11-035。")
        order_date = st.text_input("訂購日期 *", value=date.today().strftime("%Y/%m/%d"))
        internal_pn = st.text_input("銘星料號 *", key="internal_pn")
        product_name = st.text_input("項目／產品名稱 *", key="product_name")
        qty = st.text_input("訂購數量（採購單實際下單量）*", placeholder="例如：60 pcs + 備品 2 set", key="qty")
        price = st.text_input("價格／打包價 *", key="price", placeholder="例如：打包價 USD$450")
        due_date_input = st.date_input("交期 *（點選日曆選日期）", value=date.today(), format="YYYY/MM/DD")
        due = due_date_input.strftime("%Y/%m/%d") + " 前到銘星"
        delivery = st.text_input("交貨地點 *", key="delivery")
        payment = st.text_input("付款方式 *", key="payment")
    with right:
        production = st.text_area("製作內容／特殊要求 *", height=90, key="production")
        material = st.text_input("材料板厚 *", key="material")
        copper = st.text_input("銅厚 *", key="copper")
        solder = st.text_input("防焊／文字顏色 *", key="solder")
        finish = st.text_input("表面處理 *（預填，請人工核對）", key="finish")
        inspection = st.text_input("檢驗標準 *（預填，請人工核對）", key="inspection")
        notes = st.text_area("重要備註 *（預填，請人工核對）", height=90, key="notes")
        packaging = st.text_input("包裝方式 *（預填，請人工核對）", key="packaging")
        maker = st.text_input("製表者 *", key="maker")

    submitted = st.form_submit_button("產生採購單 Word", type="primary")

if submitted:
    # 訂單號碼整欄可編輯；預設 MH26A-10-XXX，由使用者自行改成完整編號。
    order_no = order_no_input.strip()
    due = due_date_input.strftime("%Y/%m/%d") + " 前到銘星"
    final_customer_pn = apply_supplier_code(customer_pn.strip(), supplier)
    required = {
        "訂單號碼": order_no, "訂購日期": order_date, "客戶料號": customer_pn,
        "銘星料號": internal_pn, "項目／產品名稱": product_name, "訂購數量": qty,
        "價格／打包價": price, "交期": due, "交貨地點": delivery, "付款方式": payment,
        "廠商名稱": supplier, "聯絡人": contact, "聯絡電話": phone,
        "製作內容／特殊要求": production, "材料板厚": material, "銅厚": copper,
        "防焊／文字顏色": solder, "表面處理": finish, "檢驗標準": inspection,
        "重要備註": notes, "包裝方式": packaging, "製表者": maker,
    }
    missing = [key for key, value in required.items() if not value.strip()]
    if missing:
        st.error("請先填寫必填欄位：" + "、".join(missing))
    elif not DEFAULT_TEMPLATE.exists():
        st.error("找不到範本檔：請確認「採購單範本.docx」與 app.py 放在同一資料夾。")
    else:
        replacements = {
            "MH26A-10-035": order_no,
            "MH26A-10-XXX": order_no,
            "2026/10/16": order_date,
            "901-W00G-F1012TSA3": final_customer_pn,
            "M056D6026A2": internal_pn,
            "印刷電路板": product_name,
            "8pcs + 備品2 set ***打樣不收單報": qty,
            "打包價USD$450": price,
            "10/30/2026前到銘星": due,
            "月結 60天": payment,
            "比斯捷科技(深圳)有限公司": supplier,
            "熊小輝": contact,
            "137 14444 4149": phone,
            "請完全依照銘星提供之資料製作  *請注意板件上不能有Made In Taiwan 字樣": production,
            "FR-4 TG150 (或同等級) / 6L / 板厚1.6mm": material,
            "內層 1 oz / 外層 1 oz": copper,
            "亮面綠色 / 文字白色": solder,
            '鍍金厚度 : 足  2 μ" (MIN) ；鍍鎳厚度 : 足  120  μ" (MIN)': finish,
            "IPC 2級": inspection,
            "打叉方式 :黑筆雙面打叉 , 4 UP允收單報 , 訂單總數5%內 (此批打樣不收單報)": notes,
            "透明真空袋+濕度卡(放側邊)+乾燥包6g(放側邊)": packaging,
            "Sally 劉曼姣": maker,
        }
        try:
            doc = Document(DEFAULT_TEMPLATE)
            replace_doc_text(doc, replacements, delivery)
            output_bytes = io.BytesIO()
            doc.save(output_bytes)
            output_bytes.seek(0)
            # 檔名使用銘星料號＋供應商簡稱（例如「M056D6026A2_比斯捷.docx」）。
            safe_internal_pn = re.sub(r'[<>:"/\\|?*]', "_", internal_pn.strip()).strip(" ._")
            supplier_short = re.sub(r"(科技.*|有限公司.*|有限责任公司.*|股份有限公司.*|公司.*)$", "", supplier.strip()).strip()
            if not supplier_short:
                supplier_short = supplier.strip()
            safe_supplier_short = re.sub(r'[<>:"/\\|?*]', "_", supplier_short).strip(" ._")
            output_filename = f"{safe_internal_pn or '銘星料號'}_{safe_supplier_short or '供應商'}.docx"
            st.download_button(
                "下載採購單 Word",
                data=output_bytes.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            st.success(f"已重新產生最新版本：訂單號碼 {order_no}。若之後又修改欄位，請再次按「產生採購單 Word」再下載。")
            st.warning("目前依範本既有文字進行替換。若範本原文有變更，需同步更新 app.py 的 replacements 對照表。")
        except Exception as exc:
            st.error(f"產生失敗：{exc}")
