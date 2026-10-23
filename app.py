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
.chart-title {font-size:1rem;font-weight:750;color:#12233F!important;margin:.45rem 0 .18rem 0;line-height:1.35;display:block;}
.chart-sub {font-size:.76rem;color:#6B7890!important;margin-bottom:.35rem;line-height:1.45;display:block;}
div[data-testid="stPlotlyChart"] {background:#FFFFFF;border:1px solid #E4EAF3;border-radius:12px;padding:.2rem;box-shadow:0 4px 14px rgba(32,60,100,.045);margin-bottom:.2rem;}
.stButton>button {border-radius:8px;border:1px solid #D8E2EE;background:#FFFFFF;color:#17365E;font-weight:700;}
[data-testid="stVerticalBlock"] {gap:.34rem;}\n.overview-divider {height:1px;background:#DCE5EF;margin:.5rem 0 .65rem 0;}\n.chart-title {margin-top:.18rem !important;}
[data-testid="stHorizontalBlock"] {gap:.7rem;}
div[role="radiogroup"] {background:#E7EEF7;border:1px solid #D4DFEC;border-radius:10px;padding:.45rem .7rem;margin:.35rem 0 .85rem 0;gap:1rem;}
div[role="radiogroup"] label, div[role="radiogroup"] p, div[role="radiogroup"] span {color:#17365E !important;font-weight:700 !important;opacity:1 !important;}
div[role="radiogroup"] label {padding:.2rem .45rem;}
div[data-testid="stPlotlyChart"] {margin-top:.12rem;}

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

st.markdown(f"""
<div class="hero"><div class="hero-title">Overview of ANRF-PAIR Projects under COEP</div></div>
""", unsafe_allow_html=True)

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
page_options = ["Portfolio overview", "Activity delivery"]
if st.session_state.get("selected_project_uid"):
    page_options.append("Project expenditure details")
if st.session_state.get("dashboard_page") not in page_options:
    st.session_state["dashboard_page"] = "Portfolio overview"
view = st.radio(
    "Select dashboard view",
    page_options,
    horizontal=True,
    label_visibility="visible",
    key="dashboard_page"
)

if view == "Portfolio overview":
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

    st.markdown('<div class="chart-title">Budget vs. expenditure by project</div><div class="chart-sub">₹ lakh · compare approved allocation with cumulative spend</div>', unsafe_allow_html=True)
    budget_plot = projects.melt(id_vars=["Project UID"], value_vars=["Budget (₹ lakh)", "Spent (₹ lakh)"], var_name="Series", value_name="Amount")
    budget_plot["Series"] = budget_plot["Series"].map({"Budget (₹ lakh)": "Sanctioned budget", "Spent (₹ lakh)": "Cumulative expenditure"})
    budget_plot["Value label"] = budget_plot["Amount"].map(lambda value: f"{value:.1f}")
    fig = px.bar(
        budget_plot, x="Project UID", y="Amount", color="Series", barmode="group", text="Value label",
        custom_data=["Project UID", "Series"],
        color_discrete_map={"Sanctioned budget": BLUE, "Cumulative expenditure": GOLD},
        category_orders={"Project UID": projects["Project UID"].tolist(),
                         "Series": ["Sanctioned budget", "Cumulative expenditure"]}
    )
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0,
                      textfont=dict(color="#1D2A3A", size=10))
    fig.update_layout(yaxis_title="₹ lakh", xaxis_title="", bargap=.28, bargroupgap=.08,
                      margin=dict(l=12, r=16, t=22, b=26))
    fig.update_yaxes(rangemode="tozero")
    polish(fig, 330)
    st.caption("Click an orange cumulative-expenditure bar to open that project's expenditure details.")
    clicked_points = plotly_events(fig, click_event=True, select_event=False, hover_event=False,
                                   override_height=330, key="budget_expenditure_click")
    if clicked_points:
        point = clicked_points[0]
        custom = point.get("customdata") or []
        clicked_uid = str(custom[0]) if len(custom) > 0 else str(point.get("x", ""))
        clicked_series = str(custom[1]) if len(custom) > 1 else ""
        if clicked_uid in projects["Project UID"].astype(str).tolist() and (
            "Cumulative expenditure" in clicked_series or point.get("curveNumber") == 1
        ):
            st.session_state["selected_project_uid"] = clicked_uid
            st.rerun()

elif view == "Project expenditure details":
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
            st.session_state["dashboard_page"] = "Portfolio overview"
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
            compare = uid_expenses.melt(id_vars=["Expense Head"], value_vars=["Sanctioned (₹)", "Spent (₹)"],
                                        var_name="Series", value_name="Amount")
            compare["Series"] = compare["Series"].map({"Sanctioned (₹)": "Sanctioned", "Spent (₹)": "Spent"})
            compare_fig = px.bar(compare, x="Expense Head", y="Amount", color="Series", barmode="group",
                                 color_discrete_map={"Sanctioned": BLUE, "Spent": GOLD}, text_auto=",.0f")
            compare_fig.update_layout(xaxis_title="", yaxis_title="Amount (₹)", margin=dict(l=12, r=12, t=20, b=35))
            compare_fig.update_yaxes(tickformat=",.0f", rangemode="tozero")
            polish(compare_fig, 360)
            st.plotly_chart(compare_fig, use_container_width=True, config=PLOT_CONFIG)
        else:
            st.info("No expense-head line items with reported spend were found for this project in Section 1.")

else:
    st.markdown('<div class="chart-title">Activity delivery across projects</div><div class="chart-sub">Reported activity counts · conducted versus attended</div>', unsafe_allow_html=True)
    cdf = conducted.groupby("Project UID", as_index=False)["Count"].sum().rename(columns={"Count": "Activities conducted"}) if not conducted.empty else pd.DataFrame(columns=["Project UID", "Activities conducted"])
    adf = attended.groupby("Project UID", as_index=False)["Count"].sum().rename(columns={"Count": "Activities attended"}) if not attended.empty else pd.DataFrame(columns=["Project UID", "Activities attended"])
    act = projects[["Project UID"]].copy().merge(cdf, on="Project UID", how="left").merge(adf, on="Project UID", how="left").fillna(0)
    act_long = act.melt(id_vars="Project UID", var_name="Activity type", value_name="Count")
    fig4 = px.bar(act_long, x="Project UID", y="Count", color="Activity type", barmode="group", text_auto=True,
                  color_discrete_map={"Activities conducted": TEAL, "Activities attended": "#8D83B8"},
                  category_orders={"Project UID": projects["Project UID"].tolist()})
    fig4.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0, textfont=dict(color="#1D2A3A", size=11))
    fig4.update_layout(xaxis_title="", yaxis_title="Activity count", bargap=.3,
                       margin=dict(l=16, r=18, t=28, b=28))
    polish(fig4, 470)
    fig4.update_yaxes(rangemode="tozero")
    st.plotly_chart(fig4, use_container_width=True, config=PLOT_CONFIG)

    st.markdown('<div class="chart-title">Research and innovation progress</div><div class="chart-sub">Section 4 · status of research papers, patents, and other reported activities</div>', unsafe_allow_html=True)
    if not progress.empty:
        pvt = progress.groupby(["Project UID", "Status"], as_index=False)["Count"].sum()
        fig5 = px.bar(
            pvt, x="Project UID", y="Count", color="Status", barmode="stack", text_auto=True,
            color_discrete_map={
                "Published / granted": BLUE,
                "In communication / filed": TEAL,
                "In progress": GOLD,
            },
            category_orders={"Project UID": projects["Project UID"].tolist()},
        )
        fig5.update_traces(textposition="inside", insidetextanchor="middle", cliponaxis=False,
                           textfont=dict(color="white", size=10), marker_line_width=0)
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
