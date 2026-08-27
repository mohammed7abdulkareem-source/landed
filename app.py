
import streamlit as st
import pandas as pd
import io
import re
import urllib.parse
from datetime import datetime

st.set_page_config(page_title="حاسبة واصل المخزن", page_icon="📦", layout="wide")

st.markdown("""
<style>
.block-container {max-width: 1200px; padding-top: 1.3rem;}
h1 {text-align:center;}
.stButton button {width:100%; font-weight:700;}
div[data-testid="stMetricValue"] {font-size:1.45rem;}
</style>
""", unsafe_allow_html=True)

st.title("📦 حاسبة سعر البضاعة واصل المخزن")
st.caption("ارفع الفاتورة، أدخل تكلفة النقل، واختر طريقة توزيع النقل.")

def clean_number(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(",", "")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s) if s else None
    except:
        return None

def normalize_columns(df):
    df = df.copy()
    mapping = {}
    for c in df.columns:
        s = str(c).strip().lower()
        if any(k in s for k in ["item", "description", "product", "الصنف", "الوصف", "name"]):
            mapping[c] = "الصنف"
        elif any(k in s for k in ["qty", "quantity", "pcs", "الكمية", "عدد"]):
            mapping[c] = "الكمية"
        elif any(k in s for k in ["unit price", "price", "السعر", "سعر الوحدة"]):
            mapping[c] = "سعر الوحدة"
        elif any(k in s for k in ["weight", "kg", "الوزن"]):
            mapping[c] = "الوزن"
        elif any(k in s for k in ["amount", "total", "value", "المجموع", "القيمة"]):
            mapping[c] = "القيمة"
    df = df.rename(columns=mapping)

    needed = ["الصنف", "الكمية", "سعر الوحدة", "الوزن", "القيمة"]
    for c in needed:
        if c not in df.columns:
            df[c] = None

    df = df[needed]
    df["الصنف"] = df["الصنف"].fillna("").astype(str)

    for c in ["الكمية", "سعر الوحدة", "الوزن", "القيمة"]:
        df[c] = df[c].apply(clean_number)

    # infer value or unit price
    mask = df["القيمة"].isna() & df["الكمية"].notna() & df["سعر الوحدة"].notna()
    df.loc[mask, "القيمة"] = df.loc[mask, "الكمية"] * df.loc[mask, "سعر الوحدة"]

    mask = df["سعر الوحدة"].isna() & df["القيمة"].notna() & df["الكمية"].notna() & (df["الكمية"] != 0)
    df.loc[mask, "سعر الوحدة"] = df.loc[mask, "القيمة"] / df.loc[mask, "الكمية"]

    # remove clearly empty rows
    df = df[
        (df["الصنف"].str.strip() != "") |
        df["الكمية"].notna() |
        df["سعر الوحدة"].notna() |
        df["القيمة"].notna()
    ].reset_index(drop=True)
    return df

def read_excel(file):
    data = file.read()
    xls = pd.ExcelFile(io.BytesIO(data))
    best = None
    for sheet in xls.sheet_names:
        tmp = pd.read_excel(io.BytesIO(data), sheet_name=sheet)
        if len(tmp) > 0:
            best = tmp
            break
    return normalize_columns(best if best is not None else pd.DataFrame())

def read_csv(file):
    data = file.read()
    for enc in ["utf-8-sig", "utf-8", "latin1"]:
        try:
            return normalize_columns(pd.read_csv(io.BytesIO(data), encoding=enc))
        except:
            pass
    return pd.DataFrame()

def read_pdf_tables(file):
    try:
        import pdfplumber
        data = file.read()
        frames = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    headers = table[0]
                    rows = table[1:]
                    try:
                        tmp = pd.DataFrame(rows, columns=headers)
                        frames.append(tmp)
                    except:
                        continue
        if frames:
            return normalize_columns(pd.concat(frames, ignore_index=True))
    except Exception:
        pass
    return pd.DataFrame()

def make_pdf(df, transport, method):
    # English PDF labels to avoid Arabic shaping/font dependency issues.
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Landed Cost Report", styles["Title"]),
        Paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["Normal"]),
        Paragraph(f"Freight / Transport: {transport:,.2f}", styles["Normal"]),
        Paragraph(f"Allocation Method: {method}", styles["Normal"]),
        Spacer(1, 12)
    ]

    cols = ["الصنف", "الكمية", "سعر الوحدة", "القيمة", "حصة النقل", "النقل/وحدة", "واصل المخزن/وحدة", "الإجمالي واصل"]
    data = [["Item","Qty","Unit Price","Goods Value","Freight Share","Freight/Unit","Landed/Unit","Landed Total"]]
    for _, r in df[cols].iterrows():
        data.append([
            str(r["الصنف"])[:35],
            f'{r["الكمية"]:,.2f}' if pd.notna(r["الكمية"]) else "",
            f'{r["سعر الوحدة"]:,.4f}' if pd.notna(r["سعر الوحدة"]) else "",
            f'{r["القيمة"]:,.2f}' if pd.notna(r["القيمة"]) else "",
            f'{r["حصة النقل"]:,.2f}',
            f'{r["النقل/وحدة"]:,.4f}' if pd.notna(r["النقل/وحدة"]) else "",
            f'{r["واصل المخزن/وحدة"]:,.4f}' if pd.notna(r["واصل المخزن/وحدة"]) else "",
            f'{r["الإجمالي واصل"]:,.2f}',
        ])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("GRID",(0,0),(-1,-1),0.5,colors.grey),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("ALIGN",(1,1),(-1,-1),"RIGHT"),
    ]))
    story.append(table)
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

uploaded = st.file_uploader(
    "ارفع الفاتورة",
    type=["xlsx", "xls", "csv", "pdf"],
    help="أفضل نتيجة حالياً مع Excel/CSV. يدعم PDF إذا كان الجدول داخل الملف قابلاً للقراءة."
)

if uploaded:
    ext = uploaded.name.lower().split(".")[-1]
    if ext in ["xlsx", "xls"]:
        df = read_excel(uploaded)
    elif ext == "csv":
        df = read_csv(uploaded)
    elif ext == "pdf":
        df = read_pdf_tables(uploaded)
    else:
        df = pd.DataFrame()

    if df.empty:
        st.warning("ما قدرت أستخرج جدول الفاتورة تلقائياً. تقدر تدخل أو تلصق البيانات يدوياً بالجدول أدناه.")
        df = pd.DataFrame(columns=["الصنف","الكمية","سعر الوحدة","الوزن","القيمة"])

    st.subheader("1) راجع بيانات الفاتورة")
    st.info("إذا أي عمود مو صحيح، عدّله مباشرة من الجدول قبل الحساب.")
    df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "الصنف": st.column_config.TextColumn("الصنف"),
            "الكمية": st.column_config.NumberColumn("الكمية", min_value=0.0, format="%.2f"),
            "سعر الوحدة": st.column_config.NumberColumn("سعر الوحدة", min_value=0.0, format="%.4f"),
            "الوزن": st.column_config.NumberColumn("الوزن (كغم)", min_value=0.0, format="%.3f"),
            "القيمة": st.column_config.NumberColumn("القيمة", min_value=0.0, format="%.2f"),
        }
    )

    # recalculate missing values after edits
    for c in ["الكمية","سعر الوحدة","الوزن","القيمة"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    mask = df["القيمة"].isna() & df["الكمية"].notna() & df["سعر الوحدة"].notna()
    df.loc[mask, "القيمة"] = df.loc[mask, "الكمية"] * df.loc[mask, "سعر الوحدة"]
    mask = df["سعر الوحدة"].isna() & df["القيمة"].notna() & df["الكمية"].notna() & (df["الكمية"] != 0)
    df.loc[mask, "سعر الوحدة"] = df.loc[mask, "القيمة"] / df.loc[mask, "الكمية"]

    st.subheader("2) أدخل تكلفة النقل")
    c1, c2 = st.columns(2)
    with c1:
        transport = st.number_input("تكلفة النقل", min_value=0.0, value=0.0, step=100.0)
    with c2:
        method = st.selectbox("توزيع النقل حسب", ["الكمية", "الوزن", "السعر"])

    if st.button("🧮 احسب واصل المخزن", type="primary"):
        work = df.copy()
        work = work[work["القيمة"].notna() | work["الكمية"].notna()].reset_index(drop=True)

        if work.empty:
            st.error("ماكو بيانات كافية للحساب.")
            st.stop()

        if method == "الكمية":
            base = work["الكمية"].fillna(0)
            label_method = "Quantity"
        elif method == "الوزن":
            base = work["الوزن"].fillna(0)
            label_method = "Weight"
        else:
            base = work["القيمة"].fillna(0)
            label_method = "Goods Value"

        total_base = float(base.sum())
        if total_base <= 0:
            st.error(f"ما أقدر أوزع النقل حسب {method} لأن المجموع صفر أو البيانات ناقصة.")
            st.stop()

        work["حصة النقل"] = base / total_base * transport
        work["النقل/وحدة"] = work.apply(
            lambda r: r["حصة النقل"] / r["الكمية"] if pd.notna(r["الكمية"]) and r["الكمية"] != 0 else None,
            axis=1
        )
        work["واصل المخزن/وحدة"] = work["سعر الوحدة"] + work["النقل/وحدة"]
        work["الإجمالي واصل"] = work["القيمة"].fillna(0) + work["حصة النقل"]

        goods_total = work["القيمة"].fillna(0).sum()
        landed_total = work["الإجمالي واصل"].sum()

        st.subheader("3) النتيجة")
        m1, m2, m3 = st.columns(3)
        m1.metric("قيمة البضاعة", f"{goods_total:,.2f}")
        m2.metric("النقل", f"{transport:,.2f}")
        m3.metric("الإجمالي واصل المخزن", f"{landed_total:,.2f}")

        show_cols = ["الصنف","الكمية","سعر الوحدة","الوزن","القيمة","حصة النقل","النقل/وحدة","واصل المخزن/وحدة","الإجمالي واصل"]
        st.dataframe(work[show_cols], use_container_width=True, hide_index=True)

        # Excel
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            work[show_cols].to_excel(writer, index=False, sheet_name="Landed Cost")
        excel_buf.seek(0)

        # PDF
        pdf_bytes = make_pdf(work, transport, label_method)

        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "📄 تحميل PDF",
                data=pdf_bytes,
                file_name="landed_cost.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with d2:
            st.download_button(
                "📊 تحميل Excel",
                data=excel_buf.getvalue(),
                file_name="landed_cost.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        summary = (
            f"نتيجة حساب واصل المخزن\n"
            f"قيمة البضاعة: {goods_total:,.2f}\n"
            f"النقل: {transport:,.2f}\n"
            f"طريقة التوزيع: {method}\n"
            f"الإجمالي واصل المخزن: {landed_total:,.2f}"
        )
        wa = "https://wa.me/?text=" + urllib.parse.quote(summary)
        st.link_button("🟢 مشاركة ملخص على واتساب", wa, use_container_width=True)

        st.caption("ملاحظة: المتصفح لا يسمح بإرفاق ملف PDF تلقائياً داخل واتساب. نزّل الـPDF ثم أرسله من واتساب، أو استخدم زر مشاركة الملخص.")
else:
    st.info("ابدأ برفع فاتورة Excel / CSV / PDF.")
