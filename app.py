import io, os, hashlib
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openpyxl import load_workbook

st.set_page_config(page_title="Project Portfolio | Executive View", page_icon="📊", layout="wide", initial_sidebar_state="collapsed")

# Presentation-first styling: compact, high contrast, and no tables or unnecessary controls.
st.markdown("""
<style>
:root { color-scheme: light; }
.stApp { background:#F3F6FB; color:#1D2A3A; }
.block-container { padding:1.05rem 1.2rem .65rem 1.2rem !important; max-width:1800px; }
[data-testid="stHeader"] { background:transparent; }
[data-testid="stToolbar"] { right:1rem; }
footer, #MainMenu { visibility:hidden; }
.hero {
 background:linear-gradient(115deg,#12233F 0%,#1E4774 100%);
 color:#FFFFFF; padding:.88rem 1.25rem; border-radius:12px; margin-top:.15rem; margin-bottom:.72rem;
 box-shadow:0 4px 14px rgba(32,60,100,.10);
}
.hero-title {font-size:1.38rem;font-weight:780;line-height:1.25;color:#FFFFFF!important;letter-spacing:-.01em;margin:0;}
.kpi {background:#FFFFFF;border:1px solid #E4EAF3;border-radius:12px;padding:.68rem .85rem;min-height:86px;box-shadow:0 4px 14px rgba(32,60,100,.06);}
.kpi-label {font-size:.72rem;color:#6B7890!important;font-weight:650;letter-spacing:.025em;}
.kpi-value {font-size:1.42rem;font-weight:780;color:#12233F!important;margin-top:.24rem;line-height:1.12;}
.kpi-foot {font-size:.68rem;color:#6B7890!important;margin-top:.18rem;}
.chart-title {font-size:1rem;font-weight:750;color:#12233F!important;margin:.12rem 0 .12rem 0;}
.chart-sub {font-size:.76rem;color:#6B7890!important;margin-bottom:.1rem;line-height:1.35;}
div[data-testid="stPlotlyChart"] {background:#FFFFFF;border:1px solid #E4EAF3;border-radius:12px;padding:.2rem;box-shadow:0 4px 14px rgba(32,60,100,.045);margin-bottom:.2rem;}
.stButton>button {border-radius:8px;border:1px solid #D8E2EE;background:#FFFFFF;color:#17365E;font-weight:700;}
[data-testid="stVerticalBlock"] {gap:.34rem;}\n.overview-divider {height:1px;background:#DCE5EF;margin:.5rem 0 .65rem 0;}\n.chart-title {margin-top:.18rem !important;}
[data-testid="stHorizontalBlock"] {gap:.7rem;}
@media(max-width:900px){.block-container{padding:.8rem .65rem .35rem .65rem!important}.hero-title{font-size:1.12rem}.kpi{padding:.48rem .55rem;min-height:72px}.kpi-value{font-size:1.12rem}}
</style>
""", unsafe_allow_html=True)

NAVY = "#183D5B"
BLUE = "#3E78A8"
TEAL = "#3E9B91"
GOLD = "#E5A84B"
PALE = "#E9F0F5"
MUTED = "#748391"
PLOT_CONFIG = {"displayModeBar": False, "responsive": True, "staticPlot": False}


def amount(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace("₹", "").replace(",", "").replace(" ", "").strip())
    except Exception: return None


def parse_workbook(data):
    wb = load_workbook(io.BytesIO(data), data_only=True)
    expected = [f"UID{i}" for i in range(1, 10)]
    missing = [x for x in expected if x not in wb.sheetnames]
    if missing: raise ValueError("Workbook must contain all nine sheets UID1–UID9. Missing: " + ", ".join(missing))
    projects, expenses, conducted, attended, warnings = [], [], [], [], []
    for uid in expected:
        sh = wb[uid]
        budget, spent = amount(sh["B7"].value), amount(sh["C7"].value)
        if budget is None or spent is None:
            warnings.append(f"{uid}: could not read sanctioned budget or cumulative expenditure from row 7.")
        projects.append({"Project UID": uid, "Sanctioned Budget (₹)": budget, "Cumulative Expenditure (₹)": spent})
        for r in range(3, 7):
            head = sh.cell(r, 1).value
            if head:
                expenses.append({"Project UID": uid, "Expense Head": str(head), "Sanctioned (₹)": amount(sh.cell(r, 2).value), "Spent (₹)": amount(sh.cell(r, 3).value)})
        section = None
        for r in range(1, sh.max_row + 1):
            label = str(sh.cell(r, 1).value or "").upper()
            if "ACTIVITIES CONDUCTED" in label:
                section = "conducted"; continue
            if "ACTIVITIES ATTENDED" in label:
                section = "attended"; continue
            activity, count = sh.cell(r, 1).value, sh.cell(r, 2).value
            if section and activity and str(activity).strip().lower() not in ("activity", "number") and isinstance(count, (int, float)):
                target = conducted if section == "conducted" else attended
                target.append({"Project UID": uid, "Activity": str(activity), "Count": float(count)})
    return pd.DataFrame(projects), pd.DataFrame(expenses), pd.DataFrame(conducted), pd.DataFrame(attended), warnings


def get_supabase():
    try:
        from supabase import create_client
        url = st.secrets.get("SUPABASE_URL", "")
        key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")
        if url and key: return create_client(url, key)
    except Exception:
        pass
    return None


def download_saved_workbook(client):
    try:
        res = client.storage.from_("portfolio-files").download("Project Info.xlsx")
        return res if res else None
    except Exception:
        return None


def save_workbook(client, data):
    try:
        client.storage.from_("portfolio-files").upload("Project Info.xlsx", data, file_options={"upsert": "true", "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
        return True, "Workbook saved to persistent storage."
    except Exception as e:
        return False, f"Persistent save failed: {e}"


def fingerprint(data): return hashlib.sha256(data).hexdigest()[:12]


# Load latest persisted workbook; use bundled workbook only as first-run fallback.
if "workbook_bytes" not in st.session_state:
    st.session_state.workbook_bytes = None
    client = get_supabase()
    if client: st.session_state.workbook_bytes = download_saved_workbook(client)
    if not st.session_state.workbook_bytes:
        with open(os.path.join(os.path.dirname(__file__), "Project Info.xlsx"), "rb") as f:
            st.session_state.workbook_bytes = f.read()



# Administrator controls stay inside the collapsed sidebar for a clean presentation view.
with st.sidebar:
    st.markdown("### Administrator access")
    st.caption("The presentation dashboard is read-only for visitors.")
    password = st.text_input("Administrator password", type="password", key="admin_password")
    expected = st.secrets.get("ADMIN_PASSWORD", "") if hasattr(st, "secrets") else ""
    is_admin = bool(expected and password and password == expected)
    if is_admin:
        st.success("Administrator access enabled")
    elif not expected:
        st.warning("Set ADMIN_PASSWORD in deployment secrets to enable administration.")
    with st.expander("Upload updated workbook", expanded=False):
        st.caption("Upload the standard Excel workbook containing UID1–UID9 sheets.")
        uploaded = st.file_uploader("Project Info.xlsx", type=["xlsx"], key="admin_workbook") if is_admin else None
        if uploaded and st.button("Validate and publish", type="primary", key="publish_workbook"):
            raw = uploaded.getvalue()
            try:
                new_projects, *_ = parse_workbook(raw)
                if new_projects["Project UID"].nunique() != 9:
                    st.error("Expected exactly UID1–UID9. Existing workbook was not changed.")
                else:
                    client = get_supabase()
                    if not client:
                        st.error("Persistent storage is not configured. Configure Supabase before publishing.")
                    else:
                        ok, msg = save_workbook(client, raw)
                        if ok:
                            st.session_state.workbook_bytes = raw
                            st.success("Workbook validated and saved.")
                            st.rerun()
                        else:
                            st.error(msg)
            except Exception as e:
                st.error(f"Workbook validation failed. Existing dashboard was not changed. Details: {e}")

try:
    projects, expense_heads, conducted, attended, warnings = parse_workbook(st.session_state.workbook_bytes)
except Exception as e:
    st.error(f"Workbook could not be read: {e}")
    st.stop()

projects["Sanctioned Budget (₹)"] = pd.to_numeric(projects["Sanctioned Budget (₹)"], errors="coerce")
projects["Cumulative Expenditure (₹)"] = pd.to_numeric(projects["Cumulative Expenditure (₹)"], errors="coerce")
projects["Budget (₹ lakh)"] = projects["Sanctioned Budget (₹)"] / 100000
projects["Spent (₹ lakh)"] = projects["Cumulative Expenditure (₹)"] / 100000
projects["Utilization (%)"] = projects.apply(lambda r: (r["Cumulative Expenditure (₹)"] / r["Sanctioned Budget (₹)"] * 100) if pd.notna(r["Sanctioned Budget (₹)"]) and r["Sanctioned Budget (₹)"] > 0 and pd.notna(r["Cumulative Expenditure (₹)"]) else None, axis=1)

total_budget = projects["Sanctioned Budget (₹)"].fillna(0).sum()
total_spent = projects["Cumulative Expenditure (₹)"].fillna(0).sum()
util = total_spent / total_budget if total_budget else None
overspent = projects[projects["Cumulative Expenditure (₹)"].fillna(0) > projects["Sanctioned Budget (₹)"].fillna(0)]

st.markdown(f"""
<div class="hero"><div class="hero-title">Overview of ANRF-PAIR Projects under COEP</div></div>
""", unsafe_allow_html=True)

# KPI cards are visual scorecards rather than Streamlit metric widgets.
def kpi(label, value, foot):
    st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-foot">{foot}</div></div>', unsafe_allow_html=True)

k1, k2, k3, k4 = st.columns(4, gap="medium")
with k1: kpi("SANCTIONED BUDGET", f"₹{total_budget/100000:,.2f} L", "Across all nine projects")
with k2: kpi("CUMULATIVE EXPENDITURE", f"₹{total_spent/100000:,.2f} L", "Reported spend to date")
with k3: kpi("PORTFOLIO UTILIZATION", f"{util:.1%}" if util is not None else "—", "Spend as a share of budget")
with k4: kpi("PROJECTS IN PORTFOLIO", f"{projects['Project UID'].nunique()} / 9", "UID1 through UID9")

st.markdown('<div class="overview-divider"></div>', unsafe_allow_html=True)


# Chart styling helpers.
def polish(fig, height=295):
    fig.update_layout(template="plotly_white", height=height, margin=dict(l=8, r=10, t=16, b=16),
                      font=dict(family="Arial, sans-serif", size=10, color="#26394A"),
                      paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                      legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0, font=dict(size=9, color="#26394A")),
                      hoverlabel=dict(bgcolor="#10243A", font_color="white"))
    fig.update_xaxes(showgrid=False, zeroline=False, linecolor="#DCE4EA", tickfont=dict(size=9, color="#34495E"), title_font=dict(size=9, color="#34495E"))
    fig.update_yaxes(gridcolor="#EAF0F4", zeroline=False, tickfont=dict(size=9, color="#34495E"), title_font=dict(size=9, color="#34495E"))
    return fig

# Presentation layout: page 1 has budget chart, then utilization + expense mix.
# Activity delivery is intentionally placed on page 2 so chart titles/descriptions never overlap.
view = st.radio(
    "Dashboard page",
    ["Portfolio overview", "Activity delivery"],
    horizontal=True,
    label_visibility="collapsed",
    key="dashboard_page"
)

if view == "Portfolio overview":
    st.markdown('<div class="chart-title">Budget vs. expenditure by project</div><div class="chart-sub">₹ lakh · compare approved allocation with cumulative spend</div>', unsafe_allow_html=True)
    budget_plot = projects.melt(id_vars=["Project UID"], value_vars=["Budget (₹ lakh)", "Spent (₹ lakh)"], var_name="Series", value_name="Amount")
    budget_plot["Series"] = budget_plot["Series"].map({"Budget (₹ lakh)": "Sanctioned budget", "Spent (₹ lakh)": "Cumulative expenditure"})
    fig = px.bar(budget_plot, x="Project UID", y="Amount", color="Series", barmode="group", text_auto=".1f",
                 color_discrete_map={"Sanctioned budget": BLUE, "Cumulative expenditure": GOLD},
                 category_orders={"Project UID": [f"UID{i}" for i in range(1,10)]})
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0, textfont=dict(color="#1D2A3A", size=10))
    fig.update_layout(yaxis_title="₹ lakh", xaxis_title="", bargap=.28, bargroupgap=.08,
                      margin=dict(l=12, r=16, t=22, b=26))
    fig.update_yaxes(rangemode="tozero")
    polish(fig, 330)
    st.plotly_chart(fig, use_container_width=True, config=PLOT_CONFIG)

    left, right = st.columns([1, 1], gap="medium")
    with left:
        st.markdown('<div class="chart-title">Portfolio utilization by project</div><div class="chart-sub">Cumulative expenditure as a percentage of sanctioned budget</div>', unsafe_allow_html=True)
        util_plot = projects.dropna(subset=["Utilization (%)"]).sort_values("Utilization (%)", ascending=True)
        util_colors = ["#D97757" if x > 100 else TEAL if x >= 70 else BLUE for x in util_plot["Utilization (%)"]]
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=util_plot["Utilization (%)"], y=util_plot["Project UID"], orientation="h", marker_color=util_colors,
                              text=[f"{x:.0f}%" for x in util_plot["Utilization (%)"]], textposition="outside", cliponaxis=False,
                              textfont=dict(color="#1D2A3A", size=10),
                              hovertemplate="%{y}<br>Utilization: %{x:.1f}%<extra></extra>"))
        fig2.add_vline(x=100, line_dash="dash", line_color="#9AA8B4", line_width=1.5,
                       annotation_text="100% budget", annotation_position="top")
        fig2.update_layout(xaxis_title="Utilization (%)", yaxis_title="", showlegend=False,
                           margin=dict(l=12, r=32, t=28, b=24))
        polish(fig2, 310)
        fig2.update_xaxes(range=[0, max(110, (util_plot["Utilization (%)"].max() * 1.18 if len(util_plot) else 110))])
        st.plotly_chart(fig2, use_container_width=True, config=PLOT_CONFIG)

    with right:
        st.markdown('<div class="chart-title">Where the expenditure goes</div><div class="chart-sub">Share of reported cumulative spend by expense head</div>', unsafe_allow_html=True)
        if not expense_heads.empty and "Spent (₹)" in expense_heads:
            eh = expense_heads.copy()
            eh["Spent (₹)"] = pd.to_numeric(eh["Spent (₹)"], errors="coerce").fillna(0)
            eh = eh.groupby("Expense Head", as_index=False)["Spent (₹)"].sum()
            eh = eh[eh["Spent (₹)"] > 0].sort_values("Spent (₹)", ascending=False)
        else:
            eh = pd.DataFrame(columns=["Expense Head", "Spent (₹)"])
        if len(eh):
            fig3 = go.Figure(go.Pie(labels=eh["Expense Head"], values=eh["Spent (₹)"], hole=.63, sort=False,
                                    marker=dict(colors=[BLUE, TEAL, GOLD, "#8D83B8", "#6B879B", "#D18B73"],
                                                line=dict(color="white", width=2)),
                                    textinfo="percent", textposition="inside",
                                    hovertemplate="%{label}<br>₹%{value:,.0f}<br>%{percent}<extra></extra>"))
            fig3.add_annotation(text=f"₹{eh['Spent (₹)'].sum()/100000:,.1f}L<br><span style='font-size:10px'>TOTAL SPEND</span>",
                                showarrow=False, font=dict(size=19, color=NAVY))
            fig3.update_layout(showlegend=True,
                               legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0,
                                           font=dict(size=9, color="#1D2A3A")),
                               margin=dict(l=8, r=8, t=26, b=18))
            polish(fig3, 310)
            st.plotly_chart(fig3, use_container_width=True, config=PLOT_CONFIG)
        else:
            st.info("No positive expense-head totals found in the source workbook.")

else:
    st.markdown('<div class="chart-title">Activity delivery across projects</div><div class="chart-sub">Reported activity counts · conducted versus attended</div>', unsafe_allow_html=True)
    cdf = conducted.groupby("Project UID", as_index=False)["Count"].sum().rename(columns={"Count": "Activities conducted"}) if not conducted.empty else pd.DataFrame(columns=["Project UID", "Activities conducted"])
    adf = attended.groupby("Project UID", as_index=False)["Count"].sum().rename(columns={"Count": "Activities attended"}) if not attended.empty else pd.DataFrame(columns=["Project UID", "Activities attended"])
    act = pd.DataFrame({"Project UID": [f"UID{i}" for i in range(1,10)]}).merge(cdf, on="Project UID", how="left").merge(adf, on="Project UID", how="left").fillna(0)
    act_long = act.melt(id_vars="Project UID", var_name="Activity type", value_name="Count")
    fig4 = px.bar(act_long, x="Project UID", y="Count", color="Activity type", barmode="group", text_auto=True,
                  color_discrete_map={"Activities conducted": TEAL, "Activities attended": "#8D83B8"},
                  category_orders={"Project UID": [f"UID{i}" for i in range(1,10)]})
    fig4.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0, textfont=dict(color="#1D2A3A", size=11))
    fig4.update_layout(xaxis_title="", yaxis_title="Activity count", bargap=.3,
                       margin=dict(l=16, r=18, t=28, b=28))
    polish(fig4, 470)
    fig4.update_yaxes(rangemode="tozero")
    st.plotly_chart(fig4, use_container_width=True, config=PLOT_CONFIG)

st.markdown(f'<div class="footer-note">ANRF–PAIR portfolio overview · Source workbook: Project Info.xlsx · Cumulative figures</div>', unsafe_allow_html=True)
