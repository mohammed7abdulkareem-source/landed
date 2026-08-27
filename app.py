import streamlit as st
import pandas as pd
import io
import re
import urllib.parse
from datetime import datetime

st.set_page_config(page_title="حاسبة واصل المخزن", page_icon="📦", layout="wide")

st.markdown("""
<style>
.block-container{max-width:1100px;padding-top:1rem}
h1{text-align:center}
.stButton>button,.stDownloadButton>button{width:100%;font-weight:700}
</style>
""", unsafe_allow_html=True)

st.title("📦 حاسبة سعر البضاعة واصل المخزن")
st.write("ارفع الفاتورة، راجع البيانات، أدخل النقل، ثم احسب السعر واصل المخزن.")

def num(x):
    if pd.isna(x): return None
    if isinstance(x,(int,float)): return float(x)
    s=re.sub(r"[^0-9.\-]","",str(x).replace(",",""))
    try:return float(s) if s else None
    except:return None

def normalize(df):
    df=df.copy()
    rename={}
    for c in df.columns:
        s=str(c).strip().lower()
        if any(k in s for k in ["description","product","item","الصنف","الوصف","name"]): rename[c]="الصنف"
        elif any(k in s for k in ["quantity","qty","pcs","الكمية","عدد"]): rename[c]="الكمية"
        elif any(k in s for k in ["unit price","unitprice","price","سعر الوحدة","السعر"]): rename[c]="سعر الوحدة"
        elif any(k in s for k in ["weight","gross weight","net weight","kg","الوزن"]): rename[c]="الوزن"
        elif any(k in s for k in ["amount","total value","line total","value","القيمة","المجموع"]): rename[c]="القيمة"
    df=df.rename(columns=rename)
    # Handle duplicate mapped columns
    df=df.loc[:,~df.columns.duplicated()]
    for c in ["الصنف","الكمية","سعر الوحدة","الوزن","القيمة"]:
        if c not in df.columns: df[c]="" if c=="الصنف" else None
    df=df[["الصنف","الكمية","سعر الوحدة","الوزن","القيمة"]]
    df["الصنف"]=df["الصنف"].fillna("").astype(str)
    for c in ["الكمية","سعر الوحدة","الوزن","القيمة"]:
        df[c]=df[c].apply(num)
    m=df["القيمة"].isna() & df["الكمية"].notna() & df["سعر الوحدة"].notna()
    df.loc[m,"القيمة"]=df.loc[m,"الكمية"]*df.loc[m,"سعر الوحدة"]
    m=df["سعر الوحدة"].isna() & df["القيمة"].notna() & df["الكمية"].notna() & (df["الكمية"]!=0)
    df.loc[m,"سعر الوحدة"]=df.loc[m,"القيمة"]/df.loc[m,"الكمية"]
    return df.dropna(how="all").reset_index(drop=True)

def parse_file(f):
    ext=f.name.lower().rsplit(".",1)[-1]
    raw=f.getvalue()
    if ext in ["xlsx","xls"]:
        # Try each sheet and choose the one with most rows
        xls=pd.ExcelFile(io.BytesIO(raw))
        frames=[]
        for sh in xls.sheet_names:
            try:
                d=pd.read_excel(io.BytesIO(raw),sheet_name=sh)
                if not d.empty: frames.append(d)
            except: pass
        if not frames: return pd.DataFrame()
        return normalize(max(frames,key=len))
    if ext=="csv":
        for enc in ["utf-8-sig","utf-8","latin1"]:
            try:return normalize(pd.read_csv(io.BytesIO(raw),encoding=enc))
            except: pass
        return pd.DataFrame()
    if ext=="pdf":
        import pdfplumber
        frames=[]
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            for page in pdf.pages:
                for t in page.extract_tables() or []:
                    if t and len(t)>1:
                        try:
                            h=[str(x or f"col{i}") for i,x in enumerate(t[0])]
                            frames.append(pd.DataFrame(t[1:],columns=h))
                        except: pass
        if frames:return normalize(pd.concat(frames,ignore_index=True))
        return pd.DataFrame()
    return pd.DataFrame()

uploaded=st.file_uploader("1️⃣ ارفع الفاتورة",type=["xlsx","xls","csv","pdf"])

if uploaded is None:
    st.info("اضغط Browse files واختر الفاتورة.")
    st.stop()

st.success(f"تم رفع الملف: {uploaded.name}")

try:
    df=parse_file(uploaded)
except Exception as e:
    st.warning("لم أستطع قراءة جدول الفاتورة تلقائياً. أدخل البيانات يدوياً بالجدول.")
    df=pd.DataFrame(columns=["الصنف","الكمية","سعر الوحدة","الوزن","القيمة"])

if df.empty:
    st.warning("لم يتم العثور على أصناف تلقائياً. أضفها يدوياً بالجدول أدناه.")
    df=pd.DataFrame([{"الصنف":"","الكمية":None,"سعر الوحدة":None,"الوزن":None,"القيمة":None}])

st.subheader("2️⃣ راجع بيانات الفاتورة")
edited=st.data_editor(
    df,num_rows="dynamic",use_container_width=True,key="invoice_editor",
    column_config={
        "الصنف":st.column_config.TextColumn("الصنف"),
        "الكمية":st.column_config.NumberColumn("الكمية",min_value=0.0),
        "سعر الوحدة":st.column_config.NumberColumn("سعر الوحدة",min_value=0.0,format="%.4f"),
        "الوزن":st.column_config.NumberColumn("الوزن (كغم)",min_value=0.0),
        "القيمة":st.column_config.NumberColumn("قيمة الصنف",min_value=0.0,format="%.2f"),
    }
)

st.subheader("3️⃣ تكلفة النقل")
c1,c2=st.columns(2)
with c1:
    freight=st.number_input("مبلغ النقل",min_value=0.0,value=0.0,step=100.0)
with c2:
    method=st.selectbox("توزيع النقل حسب",["الكمية","الوزن","السعر"])

if st.button("🧮 احسب السعر واصل المخزن",type="primary"):
    w=edited.copy()
    for c in ["الكمية","سعر الوحدة","الوزن","القيمة"]:
        w[c]=pd.to_numeric(w[c],errors="coerce")
    m=w["القيمة"].isna() & w["الكمية"].notna() & w["سعر الوحدة"].notna()
    w.loc[m,"القيمة"]=w.loc[m,"الكمية"]*w.loc[m,"سعر الوحدة"]
    m=w["سعر الوحدة"].isna() & w["القيمة"].notna() & w["الكمية"].notna() & (w["الكمية"]!=0)
    w.loc[m,"سعر الوحدة"]=w.loc[m,"القيمة"]/w.loc[m,"الكمية"]

    if method=="الكمية": base=w["الكمية"].fillna(0)
    elif method=="الوزن": base=w["الوزن"].fillna(0)
    else: base=w["القيمة"].fillna(0)

    if base.sum()<=0:
        st.error(f"بيانات {method} ناقصة أو مجموعها صفر.")
        st.stop()

    w["حصة النقل"]=base/base.sum()*freight
    w["النقل لكل وحدة"]=w.apply(lambda r:r["حصة النقل"]/r["الكمية"] if pd.notna(r["الكمية"]) and r["الكمية"]>0 else None,axis=1)
    w["سعر الوحدة واصل المخزن"]=w["سعر الوحدة"]+w["النقل لكل وحدة"]
    w["الإجمالي واصل المخزن"]=w["القيمة"].fillna(0)+w["حصة النقل"]

    st.session_state["result"]=w
    st.session_state["freight"]=freight
    st.session_state["method"]=method

if "result" in st.session_state:
    w=st.session_state["result"]
    freight=st.session_state["freight"]
    method=st.session_state["method"]
    st.subheader("✅ النتيجة")
    a,b,c=st.columns(3)
    goods=w["القيمة"].fillna(0).sum()
    total=w["الإجمالي واصل المخزن"].sum()
    a.metric("قيمة البضاعة",f"{goods:,.2f}")
    b.metric("النقل",f"{freight:,.2f}")
    c.metric("واصل المخزن",f"{total:,.2f}")
    st.dataframe(w,use_container_width=True,hide_index=True)

    out=io.BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as writer:
        w.to_excel(writer,index=False,sheet_name="Landed Cost")
    st.download_button("📊 تحميل النتيجة Excel",out.getvalue(),"landed_cost.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    msg=f"""حساب واصل المخزن
قيمة البضاعة: {goods:,.2f}
النقل: {freight:,.2f}
التوزيع حسب: {method}
الإجمالي واصل المخزن: {total:,.2f}"""
    st.link_button("🟢 مشاركة الملخص على WhatsApp","https://wa.me/?text="+urllib.parse.quote(msg))
