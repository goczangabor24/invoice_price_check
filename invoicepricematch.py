import base64
import io
import json
import math
import re
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
import pandas as pd
import pdfplumber
import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI


st.set_page_config(page_title="🐶 Price Check", page_icon="🐶", layout="wide")


# ---------------------------
# Styling
# ---------------------------

st.markdown(
    """
    <style>
    div.stButton > button {
        background-color: #28a745;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.2rem;
        font-weight: 600;
    }
    div.stButton > button:hover {
        background-color: #218838;
        color: white;
    }
    div.stButton > button:focus,
    div.stButton > button:focus-visible,
    div.stButton > button:focus:not(:active) {
        background-color: #8fdaa0;
        color: white;
        border: none;
        outline: none;
        box-shadow: 0 0 0 0.2rem rgba(143, 218, 160, 0.45);
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------
# Shared helpers
# ---------------------------

def get_api_key() -> str:
    try:
        return st.secrets["OPENAI_API_KEY"]
    except Exception:
        return ""


def normalize_code(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_eu_number(value) -> Optional[float]:
    if value is None:
        return None

    s = str(value).strip()
    if not s:
        return None

    s = s.replace("\u00a0", " ")
    s = s.replace("€", "")
    s = re.sub(r"\bEUR\b", "", s, flags=re.IGNORECASE)
    s = s.strip().replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)

    if not s:
        return None

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2:
            decimal_part = parts[-1]
            int_part = "".join(parts[:-1])
            if len(decimal_part) in (1, 2, 3):
                s = f"{int_part}.{decimal_part}"
            else:
                s = s.replace(".", "")

    try:
        return float(s)
    except Exception:
        return None


def format_eu_number(value: Optional[float], decimals: int = 2) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return f"{value:.{decimals}f}".replace(".", ",")


def normalize_european_number(value: str) -> str:
    if value is None:
        return ""

    s = str(value).strip()
    if not s:
        return ""

    s = s.replace("\u00a0", " ")
    s = re.sub(r"\bEUR\b", "", s, flags=re.IGNORECASE)
    s = s.replace("€", "").strip()
    s = s.replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)

    if not s or not re.search(r"\d", s):
        return ""

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            return s
        s = s.replace(",", "")
        s = s.replace(".", ",")
        return s

    if "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) in (1, 2, 3):
            return s
        return s.replace(",", "")

    if "." in s:
        parts = s.split(".")
        if len(parts) == 2 and len(parts[1]) in (1, 2, 3):
            return s.replace(".", ",")
        if len(parts) > 2:
            decimal_part = parts[-1]
            int_part = "".join(parts[:-1])
            if len(decimal_part) in (1, 2, 3):
                return f"{int_part},{decimal_part}"
            return "".join(parts)

    return s


def sanitize_cell(value: str, numeric: bool) -> str:
    if value is None:
        return ""

    text = str(value).strip()
    if not text:
        return ""

    if numeric:
        return normalize_european_number(text)

    return re.sub(r"\s+", " ", text).strip()


def looks_numeric_column(column_name: str) -> bool:
    name = column_name.strip().lower()
    numeric_keywords = [
        "price",
        "amount",
        "unit price",
        "total",
        "sum",
        "qty",
        "quantity",
        "cost",
        "value",
        "vat",
        "eur",
        "net",
        "gross",
        "number",
        "preis",
        "betrag",
        "menge",
        "anzahl",
        "einzelpreis",
        "unit",
    ]
    return any(keyword in name for keyword in numeric_keywords)


def map_matched_on(formula: str) -> str:
    if formula in ("F", "D*F"):
        return "List Price"
    if formula in ("G", "D*G"):
        return "Discounted Price"
    return ""


def highlight_comparison_rows(row):
    no_match = not bool(row.get("_found", False))
    ref_num = pd.to_numeric(pd.Series([row.get("_ref_num")]), errors="coerce").iloc[0]
    closest_num = pd.to_numeric(pd.Series([row.get("_closest_num")]), errors="coerce").iloc[0]

    if no_match:
        return ["background-color: #ffefef"] * len(row)

    if pd.isna(ref_num) or pd.isna(closest_num):
        return [""] * len(row)

    if ref_num > closest_num:
        return ["background-color: #ffefef"] * len(row)

    if ref_num < closest_num:
        return ["background-color: #effbef"] * len(row)

    return [""] * len(row)


def dataframe_to_tsv_without_headers(df: pd.DataFrame) -> str:
    return df.to_csv(sep="\t", index=False, header=False)


def render_copy_button(text_to_copy: str, button_label: str = "Copy to clipboard") -> None:
    escaped = json.dumps(text_to_copy)
    html = f"""
    <div style="margin-top: 0.5rem; margin-bottom: 1rem;">
      <button
        onclick='navigator.clipboard.writeText({escaped}).then(() => {{
            const msg = document.getElementById("copy-msg");
            msg.innerText = "Copied to clipboard";
            setTimeout(() => msg.innerText = "", 2000);
        }});'
        style="
          background-color:#28a745;
          color:white;
          border:none;
          border-radius:8px;
          padding:0.6rem 1.2rem;
          font-weight:600;
          cursor:pointer;
        "
      >
        {button_label}
      </button>
      <span id="copy-msg" style="margin-left: 12px; color: #28a745; font-weight: 600;"></span>
    </div>
    """
    components.html(html, height=60)


# ---------------------------
# Part 1: PDF extraction
# ---------------------------

def extract_text_and_tables_from_pdf(file_bytes: bytes) -> Tuple[str, str]:
    all_text: List[str] = []
    all_tables: List[str] = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text() or ""
            if page_text.strip():
                all_text.append(f"\n--- PAGE {page_num} ---\n{page_text}")

            try:
                tables = page.extract_tables() or []
            except Exception:
                tables = []

            for table_idx, table in enumerate(tables, start=1):
                if not table:
                    continue

                cleaned_rows = []
                for row in table:
                    if not row:
                        continue
                    cleaned = [
                        re.sub(r"\s+", " ", str(cell).strip()) if cell is not None else ""
                        for cell in row
                    ]
                    cleaned_rows.append(" | ".join(cleaned))

                if cleaned_rows:
                    all_tables.append(
                        f"\n--- PAGE {page_num} TABLE {table_idx} ---\n" + "\n".join(cleaned_rows)
                    )

    return "\n".join(all_text), "\n".join(all_tables)


def render_pdf_pages_to_base64_png(
    file_bytes: bytes,
    max_pages: int = 8,
    zoom: float = 2.0
) -> List[str]:
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    images_base64 = []

    page_count = min(len(doc), max_pages)
    for i in range(page_count):
        page = doc.load_page(i)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")
        images_base64.append(base64.b64encode(img_bytes).decode("utf-8"))

    doc.close()
    return images_base64


def build_text_prompt(columns: List[str], filename: str, text: str, table_preview: str) -> str:
    columns_text = ", ".join(columns)

    return f"""
You are extracting structured row-based data from a PDF.

Requested columns:
{columns_text}

Rules:
1. Return only rows that you can infer from the PDF content.
2. Match the requested columns as closely as possible, even if the PDF uses slightly different labels.
3. Do not invent values.
4. If a value is missing for a row, return an empty string for that field.
5. Return only JSON matching the required schema.
6. Prices and amounts should be returned as plain numeric strings without currency symbols.
7. For codes / IDs / article numbers, return only the relevant code value.

Source filename:
{filename}

PDF text:
{text[:8000]}

Extracted table preview:
{table_preview[:6000]}
""".strip()


def build_image_prompt(columns: List[str], filename: str) -> str:
    columns_text = ", ".join(columns)

    return f"""
You are extracting structured row-based data from scanned PDF page images.

Requested columns:
{columns_text}

Rules:
1. Read the uploaded page images carefully.
2. Return only rows that are actually visible in the document.
3. Match the requested columns as closely as possible, even if the document uses slightly different labels.
4. Do not invent values.
5. If a value is missing for a row, return an empty string for that field.
6. Return only JSON matching the required schema.
7. Prices and amounts should be returned as plain numeric strings without currency symbols.
8. For codes / IDs / article numbers, return only the relevant code value.

Source filename:
{filename}
""".strip()


def build_schema(columns: List[str]) -> Dict:
    return {
        "type": "object",
        "properties": {
            "rows": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {col: {"type": "string"} for col in columns},
                    "required": columns,
                    "additionalProperties": False,
                },
            }
        },
        "required": ["rows"],
        "additionalProperties": False,
    }


def clean_rows(rows: List[Dict[str, str]], columns: List[str]) -> List[Dict[str, str]]:
    cleaned_rows: List[Dict[str, str]] = []

    for row in rows:
        if not isinstance(row, dict):
            continue

        cleaned: Dict[str, str] = {}
        for col in columns:
            value = row.get(col, "")
            cleaned[col] = sanitize_cell(value, looks_numeric_column(col))

        if any(str(v).strip() for v in cleaned.values()):
            cleaned_rows.append(cleaned)

    return cleaned_rows


def extract_rows_from_text_with_openai(
    client: OpenAI,
    model: str,
    columns: List[str],
    filename: str,
    text: str,
    table_preview: str
) -> List[Dict[str, str]]:
    prompt = build_text_prompt(columns, filename, text, table_preview)
    schema = build_schema(columns)

    response = client.responses.create(
        model=model,
        instructions="You extract structured data from PDF text and return only valid JSON.",
        input=prompt,
        text={
            "format": {
                "type": "json_schema",
                "name": "pdf_rows_text",
                "schema": schema,
                "strict": True,
            }
        },
    )

    raw = (getattr(response, "output_text", "") or "").strip()
    if not raw:
        raise ValueError("The model returned an empty response for text extraction.")

    data = json.loads(raw)
    rows = data.get("rows", [])

    if not isinstance(rows, list):
        raise ValueError("The model response does not contain a valid 'rows' list.")

    return clean_rows(rows, columns)


def extract_rows_from_images_with_openai(
    client: OpenAI,
    model: str,
    columns: List[str],
    filename: str,
    images_base64: List[str]
) -> List[Dict[str, str]]:
    schema = build_schema(columns)
    prompt = build_image_prompt(columns, filename)

    content = [{"type": "input_text", "text": prompt}]
    for img_b64 in images_base64:
        content.append({
            "type": "input_image",
            "image_url": f"data:image/png;base64,{img_b64}"
        })

    response = client.responses.create(
        model=model,
        input=[{"role": "user", "content": content}],
        text={
            "format": {
                "type": "json_schema",
                "name": "pdf_rows_image",
                "schema": schema,
                "strict": True,
            }
        },
    )

    raw = (getattr(response, "output_text", "") or "").strip()
    if not raw:
        raise ValueError("The model returned an empty response for image extraction.")

    data = json.loads(raw)
    rows = data.get("rows", [])

    if not isinstance(rows, list):
        raise ValueError("The model response does not contain a valid 'rows' list.")

    return clean_rows(rows, columns)


# ---------------------------
# Part 2: Main table matching
# ---------------------------

def read_main_table(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()

    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(uploaded_file)

    raw = uploaded_file.read()

    attempts = [
        {"encoding": "utf-16", "sep": "\t"},
        {"encoding": "utf-8", "sep": "\t"},
        {"encoding": "utf-8-sig", "sep": "\t"},
        {"encoding": "latin1", "sep": "\t"},
        {"encoding": "utf-16", "sep": ","},
        {"encoding": "utf-8", "sep": ","},
        {"encoding": "utf-8-sig", "sep": ","},
        {"encoding": "latin1", "sep": ","},
        {"encoding": "utf-16", "sep": ";"},
        {"encoding": "utf-8", "sep": ";"},
        {"encoding": "utf-8-sig", "sep": ";"},
        {"encoding": "latin1", "sep": ";"},
    ]

    last_error = None
    for attempt in attempts:
        try:
            return pd.read_csv(
                io.BytesIO(raw),
                encoding=attempt["encoding"],
                sep=attempt["sep"]
            )
        except Exception as e:
            last_error = e

    raise ValueError(f"Could not read main file. Last error: {last_error}")


def read_orders_last_90_days(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()

    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(uploaded_file)

    raw = uploaded_file.read()

    attempts = [
        {"encoding": "utf-16", "sep": "\t"},
        {"encoding": "utf-8", "sep": "\t"},
        {"encoding": "utf-8-sig", "sep": "\t"},
        {"encoding": "latin1", "sep": "\t"},
        {"encoding": "utf-16", "sep": ","},
        {"encoding": "utf-8", "sep": ","},
        {"encoding": "utf-8-sig", "sep": ","},
        {"encoding": "latin1", "sep": ","},
        {"encoding": "utf-16", "sep": ";"},
        {"encoding": "utf-8", "sep": ";"},
        {"encoding": "utf-8-sig", "sep": ";"},
        {"encoding": "latin1", "sep": ";"},
    ]

    last_error = None
    for attempt in attempts:
        try:
            return pd.read_csv(
                io.BytesIO(raw),
                encoding=attempt["encoding"],
                sep=attempt["sep"]
            )
        except Exception as e:
            last_error = e

    raise ValueError(f"Could not read Orders last 90 days file. Last error: {last_error}")


def find_best_match(
    target: float,
    d: Optional[float],
    f: Optional[float],
    g: Optional[float],
    tolerance: float
):
    candidates = []

    if f is not None:
        candidates.append(("F", f))
    if g is not None:
        candidates.append(("G", g))
    if d is not None and f is not None:
        candidates.append(("D*F", d * f))
    if d is not None and g is not None:
        candidates.append(("D*G", d * g))

    if not candidates:
        return {
            "exact": False,
            "exact_formula": "",
            "exact_value": None,
            "closest_formula": "",
            "closest_value": None,
            "difference": None,
        }

    for formula, value in candidates:
        if abs(value - target) <= tolerance:
            return {
                "exact": True,
                "exact_formula": formula,
                "exact_value": value,
                "closest_formula": formula,
                "closest_value": value,
                "difference": 0.0,
            }

    closest_formula, closest_value = min(candidates, key=lambda x: abs(x[1] - target))
    diff = abs(closest_value - target)

    return {
        "exact": False,
        "exact_formula": "",
        "exact_value": None,
        "closest_formula": closest_formula,
        "closest_value": closest_value,
        "difference": diff,
    }


def build_reference_df_from_extracted(
    extracted_df: pd.DataFrame,
    code_column: str,
    value_column: str
) -> pd.DataFrame:
    ref_df = extracted_df[[code_column, value_column]].copy()
    ref_df.columns = ["ref_code", "ref_value"]

    ref_df["ref_code"] = ref_df["ref_code"].apply(normalize_code)
    ref_df["ref_value"] = ref_df["ref_value"].astype(str).str.strip()

    ref_df = ref_df[
        (ref_df["ref_code"] != "") |
        (ref_df["ref_value"] != "")
    ].copy()

    return ref_df


def build_results(main_df: pd.DataFrame, ref_df: pd.DataFrame, tolerance: float) -> pd.DataFrame:
    df = main_df.copy()

    if df.shape[1] < 7:
        raise ValueError("Main table must contain at least 7 columns so A, B, D, F, G exist.")

    col_a = df.columns[0]
    col_b = df.columns[1]
    col_d = df.columns[3]
    col_f = df.columns[5]
    col_g = df.columns[6]

    df["_A_code"] = df[col_a].apply(normalize_code)
    df["_B_code"] = df[col_b].apply(normalize_code)
    df["_D_num"] = df[col_d].apply(parse_eu_number)
    df["_F_num"] = df[col_f].apply(parse_eu_number)
    df["_G_num"] = df[col_g].apply(parse_eu_number)

    results = []

    for _, ref_row in ref_df.iterrows():
        ref_code = normalize_code(ref_row["ref_code"])
        ref_value_raw = ref_row["ref_value"]
        ref_value_num = parse_eu_number(ref_value_raw)

        matches_a = df[df["_A_code"] == ref_code]
        matches_b = df[df["_B_code"] == ref_code]
        matches = pd.concat([matches_a, matches_b]).drop_duplicates()

        if matches.empty:
            results.append({
                "reference_code": ref_code,
                "reference_value": str(ref_value_raw),
                "exact_match": "",
                "matched_on": "",
                "closest_value": "",
                "difference": "",
                "LA#": "",
                "Supplier ID": "",
                "Quantity": "",
                "list_price": "",
                "discounted_price": "",
                "_found": False,
                "_ref_num": ref_value_num,
                "_closest_num": None,
            })
            continue

        best_row_result = None
        best_main_row = None

        for _, main_row in matches.iterrows():
            if ref_value_num is None:
                comparison = {
                    "exact": False,
                    "exact_formula": "",
                    "exact_value": None,
                    "closest_formula": "",
                    "closest_value": None,
                    "difference": None,
                }
            else:
                comparison = find_best_match(
                    target=ref_value_num,
                    d=main_row["_D_num"],
                    f=main_row["_F_num"],
                    g=main_row["_G_num"],
                    tolerance=tolerance
                )

            if best_row_result is None:
                best_row_result = comparison
                best_main_row = main_row
            else:
                current_diff = comparison["difference"]
                best_diff = best_row_result["difference"]

                if comparison["exact"] and not best_row_result["exact"]:
                    best_row_result = comparison
                    best_main_row = main_row
                elif comparison["exact"] == best_row_result["exact"]:
                    if current_diff is not None and best_diff is not None and current_diff < best_diff:
                        best_row_result = comparison
                        best_main_row = main_row

        results.append({
            "reference_code": ref_code,
            "reference_value": str(ref_value_raw),
            "exact_match": "✓" if best_row_result["exact"] else "",
            "matched_on": map_matched_on(best_row_result["closest_formula"]),
            "closest_value": format_eu_number(best_row_result["closest_value"]),
            "difference": format_eu_number(best_row_result["difference"]) if best_row_result["difference"] is not None else "",
            "LA#": normalize_code(best_main_row[col_a]),
            "Supplier ID": normalize_code(best_main_row[col_b]),
            "Quantity": format_eu_number(best_main_row["_D_num"]) if best_main_row["_D_num"] is not None else "",
            "list_price": format_eu_number(best_main_row["_F_num"]) if best_main_row["_F_num"] is not None else "",
            "discounted_price": format_eu_number(best_main_row["_G_num"]) if best_main_row["_G_num"] is not None else "",
            "_found": True,
            "_ref_num": ref_value_num,
            "_closest_num": best_row_result["closest_value"],
        })

    return pd.DataFrame(results)


def build_jira_autofill_df(
    orders_df: pd.DataFrame,
    ticket_df: pd.DataFrame,
    result_df: pd.DataFrame
) -> pd.DataFrame:
    if orders_df.empty:
        raise ValueError("The Orders last 90 days file is empty.")

    if orders_df.shape[1] < 5:
        raise ValueError("The Orders last 90 days file must contain at least 5 columns so A, B and E are available.")

    first_row = orders_df.iloc[0]

    purchase_order = "" if pd.isna(first_row.iloc[0]) else str(first_row.iloc[0]).strip()
    supplier_1 = "" if pd.isna(first_row.iloc[1]) else str(first_row.iloc[1]).strip()
    shipping_fc = "" if pd.isna(first_row.iloc[4]) else str(first_row.iloc[4]).strip()

    purchasing_org = "Matina GmbH" if "matina" in supplier_1.lower() else "zooplus AG"

    if result_df["exact_match"].fillna("").eq("✓").sum() == 0:
        la_amount = "Whole invoice"
    elif len(ticket_df) == 1:
        la_amount = "Single LA"
    else:
        la_amount = "Multiple LA"

    la_values = [str(v).strip() for v in ticket_df["LA#"].fillna("").tolist() if str(v).strip()]
    la_joined = ", ".join(la_values)

    pd_in_favor_or_loss = "PD - Loss"

    jira_df = pd.DataFrame(
        [
            {
                "Shipping FC": shipping_fc,
                "Purchase Order": purchase_order,
                "Purchasing Organization": purchasing_org,
                "Category": "Wrong Price Supplier",
                "Supplier 1": supplier_1,
                "LA Amount": la_amount,
                "LA": la_joined,
                "PD: In favor or loss": pd_in_favor_or_loss,
            }
        ]
    )

    return jira_df


# ---------------------------
# UI
# ---------------------------

st.title("🐶 Price Check")

with st.sidebar:
    st.header("Settings")
    model = st.text_input("Model", value="gpt-4.1-mini")
    max_pages = st.number_input(
        "Max pages for scanned PDF fallback",
        min_value=1,
        max_value=20,
        value=3
    )
    tolerance = st.number_input(
        "Matching tolerance",
        min_value=0.0,
        value=0.00,
        step=0.01,
        format="%.2f",
        help="Two values are treated as equal if their difference is within this tolerance."
    )

st.markdown("### Upload invoice PDF")
pdf_file = st.file_uploader(
    "Invoice PDF",
    type=["pdf"],
    accept_multiple_files=False,
    key="pdf_file"
)

st.markdown("### Upload main table")
main_file = st.file_uploader(
    "Main table file",
    type=["csv", "tsv", "txt", "xlsx", "xls"],
    accept_multiple_files=False,
    key="main_file"
)

st.markdown("### Enter the 2 PDF columns to extract")
columns_input = st.text_area(
    "Exactly 2 columns, one per line",
    value="item code\nunit price w/o VAT",
    height=100,
    placeholder="Example:\nitem code\nunit price w/o VAT",
)

if "price_check_ready" not in st.session_state:
    st.session_state.price_check_ready = False

run = st.button("Price check", type="primary")

if run:
    api_key = get_api_key()

    if not api_key:
        st.error("OPENAI_API_KEY is missing from Streamlit secrets.")
        st.stop()

    if pdf_file is None:
        st.error("Please upload the invoice PDF.")
        st.stop()

    if main_file is None:
        st.error("Please upload the main table.")
        st.stop()

    columns = [line.strip() for line in columns_input.splitlines() if line.strip()]
    if len(columns) != 2:
        st.error("Please provide exactly 2 column names.")
        st.stop()

    client = OpenAI(api_key=api_key)

    try:
        file_bytes = pdf_file.read()
        text, table_preview = extract_text_and_tables_from_pdf(file_bytes)

        if text.strip() or table_preview.strip():
            extracted_rows = extract_rows_from_text_with_openai(
                client=client,
                model=model,
                columns=columns,
                filename=pdf_file.name,
                text=text,
                table_preview=table_preview,
            )
        else:
            st.info(
                f"{pdf_file.name}: No readable text layer found. Switching to image-based extraction."
            )
            images_base64 = render_pdf_pages_to_base64_png(
                file_bytes,
                max_pages=max_pages
            )

            extracted_rows = extract_rows_from_images_with_openai(
                client=client,
                model=model,
                columns=columns,
                filename=pdf_file.name,
                images_base64=images_base64,
            )

        if not extracted_rows:
            st.warning("No extractable rows were found in the uploaded PDF.")
            st.stop()

        extracted_df = pd.DataFrame(extracted_rows)

        for col in columns:
            if col not in extracted_df.columns:
                extracted_df[col] = ""

        extracted_df = extracted_df[columns]

        reference_df = build_reference_df_from_extracted(
            extracted_df=extracted_df,
            code_column=columns[0],
            value_column=columns[1]
        )

        main_df = read_main_table(main_file)
        result_df = build_results(main_df, reference_df, tolerance=tolerance)

        result_df["_ref_num"] = pd.to_numeric(result_df["_ref_num"], errors="coerce")
        result_df["_closest_num"] = pd.to_numeric(result_df["_closest_num"], errors="coerce")
        result_df["_found"] = result_df["_found"].fillna(False)

        ticket_mask = (
            result_df["_found"].eq(True)
            & result_df["_ref_num"].notna()
            & result_df["_closest_num"].notna()
            & (result_df["_ref_num"] > result_df["_closest_num"])
        )

        ticket_df = result_df.loc[ticket_mask].copy()

        st.session_state.price_check_ready = True
        st.session_state.price_check_result_df = result_df
        st.session_state.price_check_ticket_df = ticket_df

    except Exception as e:
        st.session_state.price_check_ready = False
        st.error(str(e))

if st.session_state.price_check_ready:
    result_df = st.session_state.price_check_result_df
    ticket_df = st.session_state.price_check_ticket_df

    visible_columns = [
        "reference_code",
        "reference_value",
        "exact_match",
        "matched_on",
        "closest_value",
        "difference",
        "LA#",
        "Supplier ID",
        "Quantity",
        "list_price",
        "discounted_price",
    ]

    st.markdown("### Match result")
    styled_result = result_df.style.apply(highlight_comparison_rows, axis=1)
    styled_result = styled_result.hide(axis="columns", subset=["_found", "_ref_num", "_closest_num"])
    st.dataframe(styled_result, use_container_width=True)

    if not ticket_df.empty:
        st.markdown("### Please open a ticket for the following price differences:")
        ticket_styled = ticket_df.style.apply(highlight_comparison_rows, axis=1)
        ticket_styled = ticket_styled.hide(axis="columns", subset=["_found", "_ref_num", "_closest_num"])
        st.dataframe(ticket_styled, use_container_width=True)

        st.markdown("### Please upload the Orders last 90 days table")
        orders_last_90_file = st.file_uploader(
            "Orders last 90 days table",
            type=["csv", "tsv", "txt", "xlsx", "xls"],
            accept_multiple_files=False,
            key="orders_last_90_file"
        )
        
        if orders_last_90_file is not None:
            try:
                orders_last_90_df = read_orders_last_90_days(orders_last_90_file)

                jira_autofill_df = build_jira_autofill_df(
                    orders_df=orders_last_90_df,
                    ticket_df=ticket_df,
                    result_df=result_df,
                )

                first_row = orders_last_90_df.iloc[0]

                supplier_value = "" if pd.isna(first_row.iloc[1]) else str(first_row.iloc[1]).strip()
                issue_type = "Matina" if "matina" in supplier_value.lower() else "Zooplus"

                vendor_manager = "" if pd.isna(first_row.iloc[2]) else str(first_row.iloc[2]).strip()

                st.markdown(
                    f" Don't forget to select\n"
                    f" Issue Type: **{issue_type}**\n"
                    f" Vendor Manager: **{vendor_manager}**\n"
                )

                st.markdown("### JIRA Ticket Autofill")
                edited_jira_df = st.data_editor(
                    jira_autofill_df,
                    use_container_width=True,
                    num_rows="fixed",
                    key="jira_autofill_editor"
                )

                clipboard_text = dataframe_to_tsv_without_headers(edited_jira_df)
                render_copy_button(clipboard_text, "Copy JIRA autofill to clipboard")
            except Exception as e:
                st.error(str(e))
