# Project Portfolio Dashboard — Excel-upload workflow

## Intended operating model
- Public visitors can view a single portfolio dashboard covering UID1–UID9.
- Only an authorized administrator can upload a revised Excel workbook.
- The administrator's only routine data update is uploading the same Excel workbook format. The app validates that all nine UID sheets exist, then refreshes KPI cards, charts, tables, and project comparisons from that workbook.
- The uploaded workbook is stored in a private Supabase Storage bucket so the latest approved upload can survive app restarts. The bundled `Project Info.xlsx` is used as the initial fallback.
- Dashboard visuals include sanctioned budget vs cumulative expenditure by UID, cumulative expenditure by expense head, a nine-project summary table with utilization, and activity-count charts by UID.
- Budget values and metrics are read independently from each UID worksheet.

## Important reporting limitation
The supplied workbook contains cumulative expense-head totals and cumulative activity counts, but no transaction dates or activity dates. The app therefore does not allocate these totals to June–August, September–November, or custom months. When a period other than All Available Data is selected, it displays a clear warning that period-specific metrics cannot be calculated from the current workbook. To make the timeline genuinely filter expenditure and activities, the workbook needs actual dated transaction/activity records (or an additional dated worksheet format). Never manufacture dates to make a period chart appear populated.

## Recommended hosting
- **Streamlit Community Cloud**: easiest managed hosting for this Python dashboard.
- **Supabase**: private persistent storage for the latest workbook.
- The dashboard is public read-only. Admin upload requires a shared password in this initial version. For a larger production deployment, replace the shared password with individual authenticated accounts and role-based access.

## Run locally
Requires Python 3.10+.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy
1. Create a GitHub repository and upload `app.py`, `requirements.txt`, `Project Info.xlsx`, `supabase_schema.sql`, `.gitignore`, and this README. Never commit real credentials.
2. Create a Supabase project.
3. Run `supabase_schema.sql` in Supabase SQL Editor. This creates the private `portfolio-files` storage bucket.
4. Copy the Supabase project URL and service-role key. Treat the service-role key as a highly sensitive server secret.
5. Deploy the repository in Streamlit Community Cloud using `app.py` as the entry point.
6. In the app's Settings → Secrets, set:

```toml
ADMIN_PASSWORD = "use-a-long-unique-password"
SUPABASE_URL = "https://YOUR-PROJECT.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "YOUR-SUPABASE-SERVICE-ROLE-KEY"
```

7. Deploy and test the public URL in a private/incognito browser. Confirm that viewers can see the dashboard but cannot see or use the admin upload panel.
8. Sign in as administrator and upload a copy of the source workbook. Confirm the dashboard updates and that the update remains after app restart.

## Testing checklist
- [ ] Workbook has exactly the required UID1–UID9 project sheets.
- [ ] Missing UID worksheet or unreadable budget totals are shown as validation errors/warnings.
- [ ] Each UID's budget, cumulative expenditure, expense heads, and activity counts come from that UID only.
- [ ] Portfolio budget equals the sum of sanctioned budgets.
- [ ] Portfolio utilization equals total cumulative expenditure divided by total sanctioned budget.
- [ ] Overspending appears as a warning, not silently corrected.
- [ ] Period controls update the selected-period label and never imply cumulative totals are date-filtered.
- [ ] Workbook upload with missing sheets is rejected without replacing the previous workbook.
- [ ] Valid workbook upload updates KPIs, charts, tables, and project comparisons.
- [ ] Latest workbook remains available after restart with Supabase configured.
- [ ] Public users cannot upload a workbook; incorrect admin password grants no admin controls.
- [ ] Test the layout on desktop and mobile widths.

## Known limits
- The supplied workbook does not support historical period expenditure or activity calculations without dated records.
- Admin authentication uses one shared password; use individual identities and role-based permissions before a broad institutional rollout.
- This package is prepared for deployment but has not been deployed to a live hosting account from this environment.
