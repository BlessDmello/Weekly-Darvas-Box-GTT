-- Run this once in Supabase: Project -> SQL Editor -> New query -> paste -> Run

-- One row per weekly entry per stock, owned by the logged-in user.
create table if not exists public.tracker_weeks (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users (id) on delete cascade,
  symbol text not null default '',
  week_start date,
  week_end date,
  mon_high text default '',
  tue_high text default '',
  wed_high text default '',
  thu_high text default '',
  fri_high text default '',
  shares text default '',
  execution text default 'pending',
  execution_week date,
  next_week_open text default '',
  next_week_high text default '',
  exit_hit boolean default false,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- A per-user settings row (currently just brokerage/DP charge estimate).
create table if not exists public.tracker_settings (
  user_id uuid primary key references auth.users (id) on delete cascade,
  charges_per_exit numeric not null default 65,
  updated_at timestamptz not null default now()
);

-- Lock both tables down: row level security is OFF by default in Postgres,
-- so without this, anyone with the anon key could read/write all rows.
alter table public.tracker_weeks enable row level security;
alter table public.tracker_settings enable row level security;

-- Each user may only see and modify their own rows.
create policy "Users can view their own weeks"
  on public.tracker_weeks for select
  using (auth.uid() = user_id);

create policy "Users can insert their own weeks"
  on public.tracker_weeks for insert
  with check (auth.uid() = user_id);

create policy "Users can update their own weeks"
  on public.tracker_weeks for update
  using (auth.uid() = user_id);

create policy "Users can delete their own weeks"
  on public.tracker_weeks for delete
  using (auth.uid() = user_id);

create policy "Users can view their own settings"
  on public.tracker_settings for select
  using (auth.uid() = user_id);

create policy "Users can insert their own settings"
  on public.tracker_settings for insert
  with check (auth.uid() = user_id);

create policy "Users can update their own settings"
  on public.tracker_settings for update
  using (auth.uid() = user_id);

-- Keep updated_at fresh on every edit.
create or replace function public.set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

create trigger tracker_weeks_set_updated_at
  before update on public.tracker_weeks
  for each row execute function public.set_updated_at();

create trigger tracker_settings_set_updated_at
  before update on public.tracker_settings
  for each row execute function public.set_updated_at();
