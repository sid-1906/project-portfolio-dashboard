import io, os, hashlib, re
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openpyxl import load_workbook
from streamlit_plotly_events import plotly_events

st.set_page_config(page_title="Project Portfolio | Executive View", page_icon="📊", layout="wide", initial_sidebar_state="collapsed")

# Presentation-first styling: compact, high contrast, and no tables or unnecessary controls.
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root { color-scheme: light; }
html, body, [class*="css"] { font-family:"DM Sans",Arial,sans-serif; }
.stApp { background:#F4F7FB; color:#172B45; }
.block-container { padding:1.05rem 1.55rem 1.4rem 1.55rem !important; max-width:1900px; }
[data-testid="stHeader"] { background:rgba(244,247,251,.92); }
[data-testid="stToolbar"] { right:1rem; }
footer, #MainMenu { visibility:hidden; }
.hero {
 background:linear-gradient(112deg,#10253F 0%,#17466B 58%,#176C72 100%);
 color:#FFFFFF; padding:1.25rem 1.55rem 1.3rem; border-radius:18px; margin:.1rem 0 1rem;
 box-shadow:0 10px 28px rgba(16,37,63,.14); position:relative; overflow:hidden;
 border:1px solid rgba(255,255,255,.14);
}
.hero:after {content:"";position:absolute;width:250px;height:250px;right:-70px;top:-135px;border:1px solid rgba(255,255,255,.15);border-radius:50%;box-shadow:0 0 0 28px rgba(255,255,255,.035),0 0 0 58px rgba(255,255,255,.025);}
.hero-eyebrow {font-size:.68rem;font-weight:700;letter-spacing:.17em;text-transform:uppercase;color:#9FE2D7!important;margin-bottom:.4rem;}
.hero-title {font-family:"Manrope",Arial,sans-serif;font-size:1.7rem;font-weight:800;line-height:1.2;color:#FFFFFF!important;letter-spacing:-.035em;margin:0;}
.hero-subtitle {font-size:.88rem;color:#D5E4F1!important;margin-top:.45rem;line-height:1.5;}
.kpi {background:#FFFFFF;border:1px solid #E2EAF2;border-radius:15px;padding:.95rem 1rem .9rem;min-height:112px;box-shadow:0 5px 18px rgba(26,54,84,.045);position:relative;overflow:hidden;}
.kpi:before {content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:#3E78A8;}
[data-testid="column"]:nth-child(2) .kpi:before {background:#E5A84B;}
[data-testid="column"]:nth-child(3) .kpi:before {background:#3E9B91;}
[data-testid="column"]:nth-child(4) .kpi:before {background:#8D83B8;}
.kpi-label {font-size:.68rem;color:#6B7C90!important;font-weight:700;letter-spacing:.085em;text-transform:uppercase;}
.kpi-value {font-family:"Manrope",Arial,sans-serif;font-size:1.62rem;font-weight:800;color:#132C49!important;margin-top:.45rem;line-height:1.05;letter-spacing:-.035em;}
.kpi-foot {font-size:.73rem;color:#73849A!important;margin-top:.38rem;}
.section-kicker {font-size:.67rem;font-weight:800;letter-spacing:.14em;text-transform:uppercase;color:#71839A!important;margin:.85rem 0 .2rem;}
.section-heading {font-family:"Manrope",Arial,sans-serif;font-size:1.18rem;font-weight:800;letter-spacing:-.025em;color:#142E4B!important;margin:.05rem 0 .18rem;line-height:1.35;}
.chart-title {font-family:"Manrope",Arial,sans-serif;font-size:1rem;font-weight:800;color:#173653!important;margin:.65rem 0 .18rem;line-height:1.35;letter-spacing:-.02em;display:block;}
.chart-sub {font-size:.78rem;color:#73849A!important;margin-bottom:.48rem;line-height:1.5;display:block;}
/* Clean chart surfaces: no heavy black frame around the interactive budget chart. */
div[data-testid="stPlotlyChart"] {
 background:#FFFFFF;border:1px solid #E3EAF2;border-radius:14px;
 padding:.12rem;box-shadow:0 4px 14px rgba(26,54,84,.035);
 margin-bottom:.55rem;overflow:hidden;
}
div[data-testid="stCustomComponentV1"] {
 border:0!important;outline:0!important;border-radius:14px!important;
 background:transparent!important;padding:0!important;overflow:hidden!important;
 margin-bottom:.55rem;box-shadow:none!important;
}
div[data-testid="stCustomComponentV1"] > div,
div[data-testid="stCustomComponentV1"] > iframe,
iframe[title*="streamlit_plotly_events"],
div[data-testid="stCustomComponentV1"] iframe {
 display:block;border:0!important;outline:0!important;box-shadow:none!important;
 background:transparent!important;border-radius:14px!important;
}
div[data-testid="stCustomComponentV1"] iframe {width:100%!important;}

/* streamlit-plotly-events is rendered in a custom-component iframe.
   Remove borders/backgrounds from every Streamlit wrapper around that iframe. */
div[data-testid="stElementContainer"]:has(div[data-testid="stCustomComponentV1"]),
div[data-testid="stVerticalBlock"] > div:has(div[data-testid="stCustomComponentV1"]),
div[data-testid="stCustomComponentV1"],
div[data-testid="stCustomComponentV1"] > div,
div[data-testid="stCustomComponentV1"] > iframe,
div[data-testid="stCustomComponentV1"] iframe,
iframe[title*="streamlit_plotly_events"] {
  border:0!important;
  outline:0!important;
  box-shadow:none!important;
  background:transparent!important;
}
div[data-testid="stCustomComponentV1"] {padding:0!important;margin-top:0!important;}
div[data-testid="stCustomComponentV1"] iframe {border-radius:14px!important;}

.stButton>button {border-radius:9px;border:1px solid #D8E3EF;background:#FFFFFF;color:#17365E;font-weight:700;min-height:2.55rem;padding:.45rem .9rem;}
.stButton>button:hover {border-color:#3E78A8;color:#12395F;background:#F5F9FD;}
[data-testid="stVerticalBlock"] {gap:.48rem;}
[data-testid="stHorizontalBlock"] {gap:1rem;}
div[role="radiogroup"] {background:#EAF0F7;border:1px solid #D9E3EF;border-radius:12px;padding:.48rem .7rem;margin:.35rem 0 .9rem;gap:1rem;width:max-content;}
div[role="radiogroup"] label, div[role="radiogroup"] p, div[role="radiogroup"] span {color:#17365E!important;font-weight:700!important;opacity:1!important;}
div[role="radiogroup"] label {padding:.24rem .5rem;}
.overview-divider {height:1px;background:linear-gradient(90deg,#D8E3EF,transparent);margin:.85rem 0 .75rem;}
div[data-testid="stCaptionContainer"] p {color:#728399;font-size:.75rem;}
@media(max-width:900px){
 .block-container{padding:.75rem .7rem 1rem!important}
 .hero{padding:1rem;border-radius:14px}
 .hero-title{font-size:1.3rem}
 .hero-subtitle{font-size:.78rem}
 .kpi{padding:.65rem .7rem;min-height:90px}
 .kpi-value{font-size:1.15rem}
 .section-heading{font-size:1rem}
}
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
    # Accept any number of project sheets named UID1, UID2, ...; sort numerically.
    uid_sheets = []
    for name in wb.sheetnames:
        match = re.fullmatch(r"UID\s*(\d+)", str(name).strip(), flags=re.I)
        if match:
            uid_sheets.append((int(match.group(1)), name))
    uid_sheets.sort(key=lambda x: x[0])
    if not uid_sheets:
        raise ValueError("Workbook must contain project sheets named UID1, UID2, ...")
    projects, expenses, conducted, attended, progress, warnings = [], [], [], [], [], []

    def norm_header(v):
        return re.sub(r"\s+", " ", str(v or "").strip().lower())

    for _, uid in uid_sheets:
        sh = wb[uid]
        section = None
        section_header_row = None
        section_headers = {}
        project_budget = project_spent = None

        for r in range(1, sh.max_row + 1):
            first = str(sh.cell(r, 1).value or "").strip()
            first_upper = first.upper()
            if "SECTION 1" in first_upper or "EXPENDITURE" in first_upper and "SECTION" in first_upper:
                section = "expense"; section_header_row = None; section_headers = {}; continue
            if "SECTION 2" in first_upper or "ACTIVITIES CONDUCTED" in first_upper:
                section = "conducted"; section_header_row = None; section_headers = {}; continue
            if "SECTION 3" in first_upper or "ACTIVITIES ATTENDED" in first_upper:
                section = "attended"; section_header_row = None; section_headers = {}; continue
            if "SECTION 4" in first_upper or ("ACTIVIT" in first_upper and ("PROGRES" in first_upper or "PUBLICATION" in first_upper)):
                section = "progress"; section_header_row = None; section_headers = {}; continue

            if not first:
                continue

            # Detect a section's column headers by recognizable column names.
            row_values = [sh.cell(r, c).value for c in range(1, sh.max_column + 1)]
            normalized = [norm_header(v) for v in row_values]
            if section == "expense" and any("sanction" in x for x in normalized) and any("spent" in x for x in normalized):
                section_header_row = r
                section_headers = {i + 1: x for i, x in enumerate(normalized)}
                continue
            if section in ("conducted", "attended") and any(x == "activity" for x in normalized) and any(x in ("number", "count") for x in normalized):
                section_header_row = r
                continue
            if section == "progress" and any(x == "activity" for x in normalized) and len([x for x in normalized[1:] if x]) >= 1:
                section_header_row = r
                section_headers = {i + 1: str(v).strip() for i, v in enumerate(row_values) if v is not None}
                continue

            if section == "expense":
                if first_upper == "TOTAL":
                    # Use the total row, irrespective of its row number.
                    project_budget = amount(sh.cell(r, 2).value)
                    project_spent = amount(sh.cell(r, 3).value)
                    continue
                if section_header_row and norm_header(first) not in ("need / expense head", "expense head", "total"):
                    # Only collect expense lines with a numeric value in either numeric column.
                    sanctioned = amount(sh.cell(r, 2).value)
                    spent = amount(sh.cell(r, 3).value)
                    if sanctioned is not None or spent is not None:
                        expenses.append({"Project UID": uid, "Expense Head": first, "Sanctioned (₹)": sanctioned, "Spent (₹)": spent})
            elif section in ("conducted", "attended"):
                if section_header_row and norm_header(first) != "activity":
                    count = amount(sh.cell(r, 2).value)
                    if count is not None:
                        target = conducted if section == "conducted" else attended
                        target.append({"Project UID": uid, "Activity": first, "Count": count})
            elif section == "progress" and section_header_row and norm_header(first) != "activity":
                for c in range(2, sh.max_column + 1):
                    status = str(sh.cell(section_header_row, c).value or "").strip()
                    count = amount(sh.cell(r, c).value)
                    if status and count is not None:
                        # Normalize common variants while retaining distinct status meanings.
                        s = re.sub(r"\s+", " ", status).strip()
                        sl = s.lower()
                        if "progress" in sl:
                            s = "In progress"
                        elif "publish" in sl or "grant" in sl:
                            s = "Published / granted"
                        elif "communicat" in sl or "filed" in sl or "fill" in sl:
                            s = "In communication / filed"
                        progress.append({"Project UID": uid, "Activity": first, "Status": s, "Count": count})

        if project_budget is None or project_spent is None:
            # Fallback to conventional total row (not fixed row number).
            for r in range(1, sh.max_row + 1):
                if str(sh.cell(r, 1).value or "").strip().upper() == "TOTAL":
                    project_budget = amount(sh.cell(r, 2).value)
                    project_spent = amount(sh.cell(r, 3).value)
                    break
        if project_budget is None or project_spent is None:
            warnings.append(f"{uid}: could not read the Section 1 TOTAL sanctioned/spent values.")
        projects.append({"Project UID": uid, "Sanctioned Budget (₹)": project_budget, "Cumulative Expenditure (₹)": project_spent})

    return (
        pd.DataFrame(projects),
        pd.DataFrame(expenses),
        pd.DataFrame(conducted),
        pd.DataFrame(attended),
        pd.DataFrame(progress),
        warnings,
    )


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
        st.caption("Upload the standard Excel workbook with one or more UID-numbered project sheets and Sections 1–4.")
        uploaded = st.file_uploader("Project Info.xlsx", type=["xlsx"], key="admin_workbook") if is_admin else None
        if uploaded and st.button("Validate and publish", type="primary", key="publish_workbook"):
            raw = uploaded.getvalue()
            try:
                new_projects, *_ = parse_workbook(raw)
                if new_projects.empty:
                    st.error("No valid UID-numbered project sheets were found. Existing workbook was not changed.")
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
    projects, expense_heads, conducted, attended, progress, warnings = parse_workbook(st.session_state.workbook_bytes)
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

st.markdown("""
<div class="hero">
  <div class="hero-eyebrow">Executive portfolio dashboard</div>
  <div class="hero-title">Overview of ANRF-PAIR Projects under COEP</div>
  <div class="hero-subtitle">Funding overview · Expenditure distribution · Project utilization</div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-kicker">Portfolio at a glance</div>', unsafe_allow_html=True)

# KPI cards are visual scorecards rather than Streamlit metric widgets.
def kpi(label, value, foot):
    st.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-foot">{foot}</div></div>', unsafe_allow_html=True)

k1, k2, k3, k4 = st.columns(4, gap="medium")
with k1: kpi("SANCTIONED BUDGET", f"₹{total_budget/100000:,.2f} L", f"Across all {len(projects)} projects")
with k2: kpi("CUMULATIVE EXPENDITURE", f"₹{total_spent/100000:,.2f} L", "Reported spend to date")
with k3: kpi("PORTFOLIO UTILIZATION", f"{util:.1%}" if util is not None else "—", "Spend as a share of budget")
with k4: kpi("PROJECTS IN PORTFOLIO", f"{projects['Project UID'].nunique()} projects", f"{projects['Project UID'].min()} through {projects['Project UID'].max()}")

st.markdown('<div class="overview-divider"></div>', unsafe_allow_html=True)


# Chart styling helpers.
def polish(fig, height=295):
    fig.update_layout(template="plotly_white", height=height, margin=dict(l=16, r=18, t=28, b=28),
                      font=dict(family="Arial, sans-serif", size=11, color="#26394A"),
                      paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                      legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0,
                                  font=dict(size=10, color="#26394A"), bgcolor="rgba(255,255,255,0)"),
                      hoverlabel=dict(bgcolor="#10243A", font_color="white"),
                      bargap=.24, uniformtext_minsize=9, uniformtext_mode="hide")
    fig.update_xaxes(showgrid=False, zeroline=False, linecolor="#DCE4EA",
                     tickfont=dict(size=10, color="#34495E"), title_font=dict(size=10, color="#34495E"),
                     automargin=True)
    fig.update_yaxes(gridcolor="#EAF0F4", gridwidth=1, zeroline=False,
                     tickfont=dict(size=10, color="#34495E"), title_font=dict(size=10, color="#34495E"),
                     automargin=True)
    return fig

# Presentation layout: page 1 has budget chart, then utilization + expense mix.
# Activity delivery is intentionally placed on page 2 so chart titles/descriptions never overlap.
page_options = ["Portfolio overview", "Activity delivery"]
if st.session_state.get("dashboard_page") not in page_options:
    st.session_state["dashboard_page"] = "Portfolio overview"
view = st.radio(
    "Select dashboard view",
    page_options,
    horizontal=True,
    label_visibility="visible",
    key="dashboard_page"
)
# When an expenditure bar is clicked, show the drill-down branch without
# mutating the radio widget's state after it has been instantiated.
if st.session_state.get("selected_project_uid"):
    view = "Project expenditure details"

if view == "Portfolio overview":
    st.markdown('<div class="section-kicker">Financial performance</div><div class="section-heading">Budget health & expenditure mix</div>', unsafe_allow_html=True)
    left, right = st.columns([1, 1], gap="medium")
    with left:
        st.markdown('<div class="chart-title">Portfolio utilization by project</div><div class="chart-sub">Cumulative expenditure as a percentage of sanctioned budget</div>', unsafe_allow_html=True)
        util_plot = projects.copy()
        util_plot["Utilization (%)"] = pd.to_numeric(util_plot["Utilization (%)"], errors="coerce")
        util_plot = util_plot.dropna(subset=["Utilization (%)"]).sort_values("Utilization (%)", ascending=True)
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
            eh["Spent (₹)"] = pd.to_numeric(eh["Spent (₹)"], errors="coerce").fillna(0.0)
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

    st.markdown('<div class="section-kicker">Project comparison</div><div class="section-heading">Budget vs. expenditure by project</div><div class="chart-sub">Values in ₹ lakh · compare approved allocation with cumulative spend. Select an orange bar to explore a project.</div>', unsafe_allow_html=True)
    # Build explicit numeric traces. This avoids categorical/index coercion and
    # ensures the visible bar height is the same value shown in its label.
    chart_projects = projects.copy()
    chart_projects["Budget (₹ lakh)"] = pd.to_numeric(chart_projects["Budget (₹ lakh)"], errors="coerce").fillna(0.0)
    chart_projects["Spent (₹ lakh)"] = pd.to_numeric(chart_projects["Spent (₹ lakh)"], errors="coerce").fillna(0.0)
    chart_projects["Project UID"] = chart_projects["Project UID"].astype(str)
    uids = chart_projects["Project UID"].tolist()
    budgets = chart_projects["Budget (₹ lakh)"].astype(float).tolist()
    spends = chart_projects["Spent (₹ lakh)"].astype(float).tolist()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Sanctioned budget",
        x=uids, y=budgets,
        marker=dict(color=BLUE, line=dict(width=0)),
        text=[f"{v:.1f}" for v in budgets],
        textposition="outside", cliponaxis=False,
        customdata=[[uid, "Sanctioned budget"] for uid in uids],
        hovertemplate="%{x}<br>Sanctioned budget: ₹%{y:.2f} lakh<extra></extra>",
        textfont=dict(color="#26394A", size=11)
    ))
    fig.add_trace(go.Bar(
        name="Cumulative expenditure",
        x=uids, y=spends,
        marker=dict(color=GOLD, line=dict(width=0)),
        text=[f"{v:.1f}" for v in spends],
        textposition="outside", cliponaxis=False,
        customdata=[[uid, "Cumulative expenditure"] for uid in uids],
        hovertemplate="%{x}<br>Cumulative expenditure: ₹%{y:.2f} lakh<extra></extra>",
        textfont=dict(color="#26394A", size=11)
    ))
    max_value = max(budgets + spends, default=0)
    fig.update_layout(
        barmode="group", yaxis_title="₹ lakh", xaxis_title="",
        bargap=.24, bargroupgap=.10,
        margin=dict(l=22, r=26, t=38, b=34),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0,
                    font=dict(size=10, color="#26394A")),
        yaxis=dict(rangemode="tozero", range=[0, max_value * 1.18 if max_value else 1],
                   tickformat=".1f", automargin=True)
    )
    polish(fig, 355)
    fig.update_layout(
        paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
        margin=dict(l=22, r=24, t=42, b=36),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0,
                    font=dict(size=10, color="#26394A"), bgcolor="rgba(255,255,255,0)")
    )
    fig.update_yaxes(range=[0, max_value * 1.18 if max_value else 1], tickformat=".1f")
    st.caption("Select an orange bar to view the selected project's expenditure breakdown.")
    clicked_points = plotly_events(fig, click_event=True, select_event=False, hover_event=False,
                                   override_height=355, key="budget_expenditure_click")
    if clicked_points:
        point = clicked_points[0]
        custom = point.get("customdata") or []
        clicked_uid = str(custom[0]) if len(custom) > 0 else str(point.get("x", ""))
        clicked_series = str(custom[1]) if len(custom) > 1 else ""
        if clicked_uid in uids and (
            "Cumulative expenditure" in clicked_series or point.get("curveNumber") == 1
        ):
            st.session_state["selected_project_uid"] = clicked_uid
            st.rerun()

elif view == "Project expenditure details":
    st.markdown('<div class="section-kicker">Project drill-down</div>', unsafe_allow_html=True)
    selected_uid = st.session_state.get("selected_project_uid")
    project_row = projects[projects["Project UID"].astype(str) == str(selected_uid)]
    if project_row.empty:
        st.warning("Select a project from the orange expenditure bars to view its details.")
    else:
        row = project_row.iloc[0]
        project_budget = float(row["Sanctioned Budget (₹)"] or 0)
        project_spent = float(row["Cumulative Expenditure (₹)"] or 0)
        remaining = project_budget - project_spent
        st.markdown(f'<div class="chart-title">{selected_uid} · Expenditure details</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-sub">Breakdown of expenditure reported in Section 1 of the uploaded workbook</div>', unsafe_allow_html=True)
        if st.button("← Back to portfolio overview"):
            st.session_state.pop("selected_project_uid", None)
            st.rerun()
        d1, d2, d3, d4 = st.columns(4, gap="medium")
        with d1: kpi("SANCTIONED BUDGET", f"₹{project_budget/100000:,.2f} L", selected_uid)
        with d2: kpi("CUMULATIVE EXPENDITURE", f"₹{project_spent/100000:,.2f} L", "Reported spend")
        with d3: kpi("UTILIZATION", f"{(project_spent/project_budget*100):.1f}%" if project_budget > 0 else "—", "Spent / sanctioned")
        with d4: kpi("BALANCE / OVERSPEND", f"₹{abs(remaining)/100000:,.2f} L", "Overspent" if remaining < 0 else "Remaining")

        uid_expenses = expense_heads[expense_heads["Project UID"].astype(str) == str(selected_uid)].copy() if not expense_heads.empty else pd.DataFrame()
        if not uid_expenses.empty and "Spent (₹)" in uid_expenses:
            uid_expenses["Spent (₹)"] = pd.to_numeric(uid_expenses["Spent (₹)"], errors="coerce").fillna(0)
            if "Sanctioned (₹)" in uid_expenses:
                uid_expenses["Sanctioned (₹)"] = pd.to_numeric(uid_expenses["Sanctioned (₹)"], errors="coerce").fillna(0)
            else:
                uid_expenses["Sanctioned (₹)"] = 0
            uid_expenses = uid_expenses[uid_expenses["Spent (₹)"] > 0].sort_values("Spent (₹)", ascending=True)
            left_detail, right_detail = st.columns([1.2, 1], gap="medium")
            with left_detail:
                st.markdown('<div class="chart-title">Expenditure by expense head</div><div class="chart-sub">Reported cumulative spend for this project</div>', unsafe_allow_html=True)
                exp_fig = px.bar(uid_expenses, x="Spent (₹)", y="Expense Head", orientation="h",
                                 text=uid_expenses["Spent (₹)"].map(lambda x: f"₹{x:,.0f}"),
                                 color_discrete_sequence=[BLUE])
                exp_fig.update_layout(xaxis_title="Expenditure (₹)", yaxis_title="", showlegend=False,
                                      margin=dict(l=12, r=18, t=18, b=28))
                exp_fig.update_xaxes(tickformat=",.0f")
                polish(exp_fig, max(300, 46 * len(uid_expenses) + 95))
                st.plotly_chart(exp_fig, use_container_width=True, config=PLOT_CONFIG)
            with right_detail:
                st.markdown('<div class="chart-title">Share of project expenditure</div><div class="chart-sub">Percentage of reported spend by expense head</div>', unsafe_allow_html=True)
                pie_fig = px.pie(uid_expenses, names="Expense Head", values="Spent (₹)", hole=.58,
                                 color_discrete_sequence=[BLUE, TEAL, GOLD, "#8D83B9", "#D97757"])
                pie_fig.update_traces(textposition="inside", textinfo="percent", sort=False)
                pie_fig.update_layout(margin=dict(l=8, r=8, t=18, b=18), showlegend=True,
                                      legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0))
                polish(pie_fig, 330)
                st.plotly_chart(pie_fig, use_container_width=True, config=PLOT_CONFIG)
            st.markdown('<div class="chart-title">Sanctioned amount vs. actual expenditure by expense head</div>', unsafe_allow_html=True)
            uid_expenses["Sanctioned (₹)"] = pd.to_numeric(uid_expenses["Sanctioned (₹)"], errors="coerce").fillna(0.0)
            uid_expenses["Spent (₹)"] = pd.to_numeric(uid_expenses["Spent (₹)"], errors="coerce").fillna(0.0)
            compare = uid_expenses.melt(id_vars=["Expense Head"], value_vars=["Sanctioned (₹)", "Spent (₹)"],
                                        var_name="Series", value_name="Amount")
            compare["Amount"] = pd.to_numeric(compare["Amount"], errors="coerce").fillna(0.0)
            compare["Series"] = compare["Series"].map({"Sanctioned (₹)": "Sanctioned", "Spent (₹)": "Spent"})
            compare["Amount label"] = compare["Amount"].map(lambda value: f"₹{value:,.0f}")
            compare_fig = px.bar(compare, x="Expense Head", y="Amount", color="Series", barmode="group",
                                 color_discrete_map={"Sanctioned": BLUE, "Spent": GOLD}, text="Amount label")
            compare_fig.update_layout(xaxis_title="", yaxis_title="Amount (₹)", margin=dict(l=12, r=12, t=20, b=35))
            compare_fig.update_yaxes(tickformat=",.0f", rangemode="tozero")
            compare_fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(color="#26394A", size=10))
            polish(compare_fig, 360)
            st.plotly_chart(compare_fig, use_container_width=True, config=PLOT_CONFIG)
        else:
            st.info("No expense-head line items with reported spend were found for this project in Section 1.")

else:
    st.markdown('<div class="section-kicker">Delivery performance</div><div class="section-heading">Activity delivery across projects</div><div class="chart-sub">Reported activity counts · conducted versus attended</div>', unsafe_allow_html=True)
    cdf = conducted.groupby("Project UID", as_index=False)["Count"].sum().rename(columns={"Count": "Activities conducted"}) if not conducted.empty else pd.DataFrame(columns=["Project UID", "Activities conducted"])
    adf = attended.groupby("Project UID", as_index=False)["Count"].sum().rename(columns={"Count": "Activities attended"}) if not attended.empty else pd.DataFrame(columns=["Project UID", "Activities attended"])
    act = projects[["Project UID"]].copy().merge(cdf, on="Project UID", how="left").merge(adf, on="Project UID", how="left").fillna(0)
    act["Project UID"] = act["Project UID"].astype(str)
    act["Activities conducted"] = pd.to_numeric(act["Activities conducted"], errors="coerce").fillna(0)
    act["Activities attended"] = pd.to_numeric(act["Activities attended"], errors="coerce").fillna(0)
    act_long = act.melt(id_vars="Project UID", var_name="Activity type", value_name="Count")
    act_long["Count"] = pd.to_numeric(act_long["Count"], errors="coerce").fillna(0)
    act_long["Count label"] = act_long["Count"].map(lambda value: f"{value:g}")
    fig4 = px.bar(act_long, x="Project UID", y="Count", color="Activity type", barmode="group", text="Count label",
                  color_discrete_map={"Activities conducted": TEAL, "Activities attended": "#8D83B8"},
                  category_orders={"Project UID": projects["Project UID"].astype(str).tolist(),
                                   "Activity type": ["Activities conducted", "Activities attended"]})
    fig4.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0, textfont=dict(color="#26394A", size=11),
                       hovertemplate="%{x}<br>%{fullData.name}: %{y}<extra></extra>")
    fig4.update_layout(xaxis_title="", yaxis_title="Activity count", bargap=.3,
                       margin=dict(l=16, r=18, t=28, b=28))
    polish(fig4, 470)
    fig4.update_yaxes(rangemode="tozero")
    st.plotly_chart(fig4, use_container_width=True, config=PLOT_CONFIG)

    st.markdown('<div class="chart-title">Research and innovation progress</div><div class="chart-sub">Section 4 · status of research papers, patents, and other reported activities</div>', unsafe_allow_html=True)
    if not progress.empty:
        pvt = progress.copy()
        pvt["Count"] = pd.to_numeric(pvt["Count"], errors="coerce").fillna(0)
        pvt["Project UID"] = pvt["Project UID"].astype(str)
        pvt = pvt.groupby(["Project UID", "Status"], as_index=False)["Count"].sum()
        pvt["Count label"] = pvt["Count"].map(lambda value: f"{value:g}" if value > 0 else "")
        fig5 = px.bar(
            pvt, x="Project UID", y="Count", color="Status", barmode="stack", text="Count label",
            color_discrete_map={
                "Published / granted": BLUE,
                "In communication / filed": TEAL,
                "In progress": GOLD,
            },
            category_orders={"Project UID": projects["Project UID"].astype(str).tolist()},
        )
        fig5.update_traces(textposition="inside", insidetextanchor="middle", cliponaxis=False,
                           textfont=dict(color="white", size=10), marker_line_width=0,
                           hovertemplate="%{x}<br>%{fullData.name}: %{y}<extra></extra>")
        fig5.update_layout(xaxis_title="", yaxis_title="Reported items", bargap=.28,
                           margin=dict(l=16, r=18, t=22, b=28),
                           legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0,
                                       font=dict(size=9, color="#1D2A3A")))
        polish(fig5, 390)
        fig5.update_yaxes(rangemode="tozero")
        st.plotly_chart(fig5, use_container_width=True, config=PLOT_CONFIG)
    else:
        st.info("Section 4 was not found or contains no numeric status counts in this workbook.")

st.markdown(f'<div class="footer-note">ANRF–PAIR portfolio overview · Source workbook: Project Info.xlsx · Cumulative figures</div>', unsafe_allow_html=True)
