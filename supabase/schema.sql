-- Elite Haulerz worker accounts.
-- Run this once in Supabase: Dashboard → SQL Editor → New query → paste → Run.

-- One row per worker, linking their Supabase auth account to a username.
create table if not exists public.profiles (
  id         uuid primary key references auth.users (id) on delete cascade,
  username   text not null unique
             check (username ~ '^[a-z0-9_]{3,24}$'),
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;

-- Defense in depth on top of the policies below: logged-out visitors (anon)
-- get no access at all, and workers can only read and insert -- never edit,
-- delete, or wipe the table -- even if a policy is added by mistake later.
revoke all on public.profiles from anon;
revoke all on public.profiles from authenticated;
grant select, insert on public.profiles to authenticated;

-- Workers can see only their own profile row.
drop policy if exists "profiles: read own" on public.profiles;
create policy "profiles: read own"
  on public.profiles for select
  to authenticated
  using (id = auth.uid());

-- Workers can create their own profile row (once, during account setup).
drop policy if exists "profiles: insert own" on public.profiles;
create policy "profiles: insert own"
  on public.profiles for insert
  to authenticated
  with check (id = auth.uid());

-- No update/delete policies: usernames are fixed once chosen.
-- An admin can change them in the Table Editor if needed.
--
-- The username → email lookup happens in the Netlify login function using the
-- secret key, so worker emails are never exposed to the browser.
