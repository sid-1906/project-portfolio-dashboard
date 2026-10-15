import io, os, hashlib
from datetime import date
import pandas as pd
import streamlit as st
import plotly.express as px
from openpyxl import load_workbook

st.set_page_config(page_title="Project Portfolio Dashboard", page_icon="📊", layout="wide")
st.markdown("""
<style>
.block-container{padding-top:1rem;max-width:1550px}
div[data-testid="stMetric"]{background:#f5f8fc;border:1px solid #e2e8f0;padding:12px 15px;border-radius:12px}
.period-banner{padding:13px 17px;border-radius:12px;background:#eaf2fb;border:1px solid #cbdcf1;margin-bottom:14px}
</style>""",unsafe_allow_html=True)
st.title("📊 Project Portfolio Dashboard")
st.caption("Portfolio overview • UID1–UID9 • Excel workbook is the source of truth")

def amount(v):
    if v is None: return None
    if isinstance(v,(int,float)): return float(v)
    try: return float(str(v).replace("₹","").replace(",","").replace(" ","").strip())
    except: return None

def parse_workbook(data):
    wb=load_workbook(io.BytesIO(data),data_only=True)
    expected=[f"UID{i}" for i in range(1,10)]
    missing=[x for x in expected if x not in wb.sheetnames]
    if missing: raise ValueError("Workbook must contain all nine sheets UID1–UID9. Missing: "+", ".join(missing))
    projects=[]; expenses=[]; conducted=[]; attended=[]; warnings=[]
    for uid in expected:
        sh=wb[uid]
        budget=amount(sh["B7"].value); spent=amount(sh["C7"].value)
        if budget is None or spent is None: warnings.append(f"{uid}: could not read sanctioned budget or cumulative expenditure from row 7.")
        projects.append({"Project UID":uid,"Sanctioned Budget (₹)":budget,"Cumulative Expenditure (₹)":spent})
        for r in range(3,7):
            head=sh.cell(r,1).value
            if head:
                expenses.append({"Project UID":uid,"Expense Head":str(head),
                                 "Sanctioned (₹)":amount(sh.cell(r,2).value),
                                 "Spent (₹)":amount(sh.cell(r,3).value)})
        section=None
        for r in range(1,sh.max_row+1):
            label=str(sh.cell(r,1).value or "").upper()
            if "ACTIVITIES CONDUCTED" in label:
                section="conducted"; continue
            if "ACTIVITIES ATTENDED" in label:
                section="attended"; continue
            activity=sh.cell(r,1).value; count=sh.cell(r,2).value
            if section and activity and str(activity).strip().lower() not in ("activity","number") and isinstance(count,(int,float)):
                target=conducted if section=="conducted" else attended
                target.append({"Project UID":uid,"Activity":str(activity),"Count":float(count)})
    return pd.DataFrame(projects),pd.DataFrame(expenses),pd.DataFrame(conducted),pd.DataFrame(attended),warnings

def get_supabase():
    try:
        from supabase import create_client
        url=st.secrets.get("SUPABASE_URL","")
        key=st.secrets.get("SUPABASE_SERVICE_ROLE_KEY","")
        if url and key: return create_client(url,key)
    except Exception: pass
    return None

def download_saved_workbook(client):
    try:
        res=client.storage.from_("portfolio-files").download("Project Info.xlsx")
        return res if res else None
    except Exception:
        return None

def save_workbook(client,data):
    try:
        client.storage.from_("portfolio-files").upload(
            "Project Info.xlsx",data,
            file_options={"upsert":"true","content-type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
        return True,"Workbook saved to persistent storage."
    except Exception as e: return False,f"Persistent save failed: {e}"

def fingerprint(data): return hashlib.sha256(data).hexdigest()[:12]

# Prefer the latest persisted workbook on startup; bundled workbook is the safe fallback.
if "workbook_bytes" not in st.session_state:
    st.session_state.workbook_bytes=None
    client=get_supabase()
    if client:
        st.session_state.workbook_bytes=download_saved_workbook(client)
    if not st.session_state.workbook_bytes:
        with open(os.path.join(os.path.dirname(__file__),"Project Info.xlsx"),"rb") as f:
            st.session_state.workbook_bytes=f.read()

# Reporting controls appear before all dashboard metrics and visuals.
with st.container(border=True):
    st.subheader("Reporting period")
    c1,c2,c3,c4=st.columns([1.55,1,1,1])
    with c1: preset=st.selectbox("Period",["All Available Data","June–August","September–November","Custom"])
    years=list(range(date.today().year+2,1999,-1))
    with c2: year=st.selectbox("Year",years,index=years.index(date.today().year))
    with c3: custom_start=st.date_input("Custom start",value=date(year,1,1),disabled=preset!="Custom")
    with c4: custom_end=st.date_input("Custom end",value=date(year,12,31),disabled=preset!="Custom")
    if preset=="June–August": start,end=date(year,6,1),date(year,8,31)
    elif preset=="September–November": start,end=date(year,9,1),date(year,11,30)
    elif preset=="Custom": start,end=custom_start,custom_end
    else: start=end=None
    if preset=="Custom" and start>end: st.error("Custom start date must be on or before custom end date.")
    period_label=("All Available Data (cumulative source totals)" if preset=="All Available Data" else
                  (f"{start:%d %b %Y} – {end:%d %b %Y}" if start and end else "Choose a valid date range"))
    st.markdown(f'<div class="period-banner"><b>Selected reporting period:</b> {period_label}</div>',unsafe_allow_html=True)

with st.sidebar:
    st.header("Administration")
    st.caption("Public access is read-only. Only authorized administrators can replace the source workbook.")
    password=st.text_input("Administrator password",type="password")
    expected=st.secrets.get("ADMIN_PASSWORD","") if hasattr(st,"secrets") else ""
    is_admin=bool(expected and password and password==expected)
    if is_admin: st.success("Administrator access enabled")
    elif not expected: st.warning("Set ADMIN_PASSWORD in deployment secrets to enable administration.")
    else: st.caption("Sign in to manage the workbook.")
    st.divider()
    st.caption("Only upload the standard Excel workbook containing UID1–UID9 sheets.")

try:
    projects,expense_heads,conducted,attended,warnings=parse_workbook(st.session_state.workbook_bytes)
except Exception as e:
    st.error(f"Workbook could not be read: {e}")
    st.stop()

for warning in warnings: st.warning(warning)
for _,r in projects.iterrows():
    b,s=r["Sanctioned Budget (₹)"],r["Cumulative Expenditure (₹)"]
    if pd.notna(b) and pd.notna(s) and s>b:
        st.warning(f'{r["Project UID"]}: cumulative expenditure (₹{s:,.0f}) exceeds sanctioned budget (₹{b:,.0f}). Verify the source workbook and any approved revisions.')

total_budget=projects["Sanctioned Budget (₹)"].fillna(0).sum()
total_spent=projects["Cumulative Expenditure (₹)"].fillna(0).sum()
util=total_spent/total_budget if total_budget else None
if preset!="All Available Data":
    st.info("The selected period is displayed above, but this workbook contains cumulative expenditure and undated activity totals only. Period-specific expenditure and activity metrics cannot be calculated from this file. No values are assigned to months.")
k1,k2,k3,k4=st.columns(4)
k1.metric("Total sanctioned budget",f"₹{total_budget:,.0f}")
k2.metric("Cumulative expenditure",f"₹{total_spent:,.0f}")
k3.metric("Portfolio utilization",f"{util:.1%}" if util is not None else "Unavailable")
k4.metric("Projects",f"{projects['Project UID'].nunique()} / 9")
st.caption(f"Source workbook fingerprint: {fingerprint(st.session_state.workbook_bytes)} • Latest source values loaded in this app.")

# Single portfolio overview dashboard, with visuals that compare all nine UIDs.
left,right=st.columns([1.35,1])
with left:
    st.subheader("Sanctioned budget vs cumulative expenditure")
    plot=projects.melt(id_vars=["Project UID"],value_vars=["Sanctioned Budget (₹)","Cumulative Expenditure (₹)"],var_name="Metric",value_name="Amount")
    fig=px.bar(plot,x="Project UID",y="Amount",color="Metric",barmode="group",template="plotly_white",
               color_discrete_sequence=["#2D6A9F","#E9A23B"])
    fig.update_layout(legend_title_text="",yaxis_tickprefix="₹",yaxis_tickformat=",.0f",margin=dict(l=8,r=8,t=25,b=8))
    st.plotly_chart(fig,use_container_width=True)
with right:
    st.subheader("Cumulative expenditure by expense head")
    eh=expense_heads.groupby("Expense Head",as_index=False)["Spent (₹)"].sum()
    fig2=px.pie(eh,names="Expense Head",values="Spent (₹)",hole=.48,template="plotly_white",
                color_discrete_sequence=["#2D6A9F","#69A7A8","#E9A23B","#8C78B8"])
    fig2.update_traces(textposition="inside",textinfo="percent+label")
    st.plotly_chart(fig2,use_container_width=True)

st.subheader("All-project overview")
overview=projects.copy()
overview["Utilization"]=overview.apply(lambda r:r["Cumulative Expenditure (₹)"]/r["Sanctioned Budget (₹)"] if pd.notna(r["Sanctioned Budget (₹)"]) and r["Sanctioned Budget (₹)"] else None,axis=1)
st.dataframe(overview,hide_index=True,use_container_width=True,column_config={
    "Sanctioned Budget (₹)":st.column_config.NumberColumn(format="₹%.0f"),
    "Cumulative Expenditure (₹)":st.column_config.NumberColumn(format="₹%.0f"),
    "Utilization":st.column_config.ProgressColumn(min_value=0,max_value=1,format="%.1f%%")
})
left2,right2=st.columns(2)
with left2:
    st.subheader("Activity counts by project")
    if len(conducted):
        cfig=px.bar(conducted,x="Project UID",y="Count",color="Activity",barmode="stack",template="plotly_white")
        cfig.update_layout(margin=dict(l=8,r=8,t=20,b=8),legend_title_text="Activity")
        st.plotly_chart(cfig,use_container_width=True)
    else: st.info("No activity counts found in the workbook.")
with right2:
    st.subheader("Activity attended by project")
    if len(attended):
        af=px.bar(attended,x="Project UID",y="Count",color="Activity",barmode="stack",template="plotly_white")
        af.update_layout(margin=dict(l=8,r=8,t=20,b=8),legend_title_text="Activity")
        st.plotly_chart(af,use_container_width=True)
    else: st.info("No attended activity counts found in the workbook.")

with st.expander("Project-level details"):
    selected=st.selectbox("View a single UID",projects["Project UID"].tolist())
    row=projects[projects["Project UID"]==selected].iloc[0]
    a,b,c=st.columns(3)
    a.metric("Sanctioned budget",f'₹{(row["Sanctioned Budget (₹)"] or 0):,.0f}')
    b.metric("Cumulative expenditure",f'₹{(row["Cumulative Expenditure (₹)"] or 0):,.0f}')
    ratio=(row["Cumulative Expenditure (₹)"]/row["Sanctioned Budget (₹)"]) if row["Sanctioned Budget (₹)"] else None
    c.metric("Utilization",f"{ratio:.1%}" if ratio is not None else "Unavailable")
    st.markdown("**Expense heads**")
    st.dataframe(expense_heads[expense_heads["Project UID"]==selected],hide_index=True,use_container_width=True)
    x,y=st.columns(2)
    with x:
        st.markdown("**Activities conducted**")
        st.dataframe(conducted[conducted["Project UID"]==selected],hide_index=True,use_container_width=True)
    with y:
        st.markdown("**Activities attended**")
        st.dataframe(attended[attended["Project UID"]==selected],hide_index=True,use_container_width=True)

with st.expander("Data quality and reporting notes"):
    st.markdown("""
    - All project values are read independently from each UID worksheet in the uploaded workbook.
    - Portfolio utilization is total cumulative expenditure divided by total sanctioned budget.
    - Cumulative expenditure and undated activity counts cannot be filtered by date.
    - The reporting-period selector is ready for dated records, but this workbook format does not provide transaction dates or activity dates.
    - Research-output labels (for example, research papers and patents) are displayed as source activity labels; validate them before formal reporting.
    - Data freshness is identified by the workbook fingerprint above. Upload a revised workbook to refresh all dashboard metrics and visuals.
    """)

st.divider()
st.caption("Source of truth: Project Info.xlsx • Public dashboard is read-only • Admin changes replace the workbook as a whole")
# Admin panel at bottom; the only admin data-entry action is uploading the Excel workbook.
if is_admin:
    with st.expander("Administrator: upload updated Excel workbook",expanded=True):
        st.write("Upload the standard workbook with all nine UID worksheets. All KPI cards, charts, and tables will be recalculated from it.")
        uploaded=st.file_uploader("Upload updated Project Info.xlsx",type=["xlsx"],key="admin_workbook")
        if uploaded and st.button("Validate and publish workbook",type="primary"):
            raw=uploaded.getvalue()
            try:
                new_projects,*_=parse_workbook(raw)
                if new_projects["Project UID"].nunique()!=9:
                    st.error("Expected exactly UID1–UID9. Existing workbook was not changed.")
                else:
                    client=get_supabase()
                    if not client:
                        st.error("Persistent storage is not configured. Configure Supabase before publishing, so the upload survives app restarts.")
                    else:
                        ok,msg=save_workbook(client,raw)
                        if ok:
                            st.session_state.workbook_bytes=raw
                            st.success("Workbook validated and saved. Dashboard refreshed from the new workbook.")
                            st.rerun()
                        else: st.error(msg)
            except Exception as e:
                st.error(f"Workbook validation failed. Existing dashboard was not changed. Details: {e}")
