-- Run this in Supabase SQL Editor before deploying.
-- Keep the service-role key ONLY in Streamlit server secrets. Never expose it in frontend code.
create table if not exists public.dated_expenses (
  transaction_id text primary key,
  project_uid text not null check (project_uid in ('UID1','UID2','UID3','UID4','UID5','UID6','UID7','UID8','UID9')),
  transaction_date date not null,
  expense_head text not null,
  amount numeric(14,2) not null check (amount >= 0),
  description text,
  constraint expense_head_allowed check (expense_head in ('Travel','Contingency','Consumables','Other'))
);
create table if not exists public.dated_activities (
  activity_id text primary key,
  project_uid text not null check (project_uid in ('UID1','UID2','UID3','UID4','UID5','UID6','UID7','UID8','UID9')),
  activity_date date not null,
  activity_type text not null,
  count numeric(12,2) not null check (count >= 0),
  attended text,
  research_output text,
  description text
);
create table if not exists public.budget_revisions (
  revision_id text primary key,
  project_uid text not null check (project_uid in ('UID1','UID2','UID3','UID4','UID5','UID6','UID7','UID8','UID9')),
  effective_date date not null,
  previous_budget numeric(14,2) not null check (previous_budget >= 0),
  new_budget numeric(14,2) not null check (new_budget >= 0),
  approval_reference text not null,
  approved_by text not null,
  recorded_at timestamptz not null default now(),
  status text not null
);
alter table public.dated_expenses enable row level security;
alter table public.dated_activities enable row level security;
alter table public.budget_revisions enable row level security;
-- Do not create public write policies. This app uses the service-role key only on the server,
-- and protects write tools behind ADMIN_PASSWORD. For stronger per-user audit, use Supabase Auth
-- and server-side role checks before production deployment.

-- Persistent storage for the latest approved source workbook.
insert into storage.buckets (id, name, public)
values ('portfolio-files', 'portfolio-files', false)
on conflict (id) do nothing;
-- No public storage policies are added. Server-side service-role credentials are used by Streamlit.
