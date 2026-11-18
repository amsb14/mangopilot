-- Tighten row-level security.
--
-- Verified from the app (anon key, no password):
--   * decision_journal was readable (policy "anon all on own pin" is USING (true)).
--   * user_memory was readable. The client filters with .eq('pin_hash', ...),
--     which is not a security boundary — anyone with the anon key can omit it.
--
-- Public market tables stay readable. The site loads them before sign-in.
-- Writes to those tables stay denied for anon; sync uses the service role.
--
-- Private tables (journal, memory) require the x-auth-hash header the app
-- sends after sign-in, and that header must equal the password hash stored
-- in app_auth. Copy that value from the existing edge-function secret
-- APP_PASSWORD_HASH (same SHA-256 the sign-in screen already checks).
-- Do not commit the hash.
--
-- Apply in the Supabase SQL editor, AFTER the new site is deployed
-- (the new app.js sends x-auth-hash on these requests).
--
-- Then insert the secret (run once, replace the placeholder):
--
--   insert into public.app_auth (id, password_hash)
--   values (1, 'PASTE_APP_PASSWORD_HASH_HERE')
--   on conflict (id) do update set password_hash = excluded.password_hash, updated_at = now();
--
-- Verify (REST, anon key):
--   1. GET /decision_journal with no x-auth-hash  -> 0 rows
--   2. GET /user_memory with no x-auth-hash       -> 0 rows
--   3. Same requests with x-auth-hash set to the real hash -> only that user's rows
--   4. GET /company_profiles still returns rows (the public site depends on this)
--
-- If step 3 returns nothing, Postgres is not seeing the header. Check
-- current_setting('request.headers', true) from a request that sends it.
-- Rollback for the private tables only (restores the old open access — temporary):
--
--   drop policy if exists "signed-in journal" on public.decision_journal;
--   create policy "anon all on own pin" on public.decision_journal
--     for all to anon, authenticated using (true) with check (true);
--   drop policy if exists "signed-in memory" on public.user_memory;
--   create policy "anon all memory" on public.user_memory
--     for all to anon, authenticated using (true) with check (true);

-- ── Password hash the Data API can check (not readable by anon) ────────────
create table if not exists public.app_auth (
  id integer primary key check (id = 1),
  password_hash text not null,
  updated_at timestamptz default now()
);

alter table public.app_auth enable row level security;
revoke all on public.app_auth from anon, authenticated;

-- Header lookup that fails closed if PostgREST didn't pass headers.
create or replace function public.request_header(header_name text)
returns text
language plpgsql
stable
as $$
declare
  raw text;
  hdr jsonb;
begin
  raw := current_setting('request.headers', true);
  if raw is null or btrim(raw) = '' then
    return '';
  end if;
  begin
    hdr := raw::jsonb;
  exception when others then
    return '';
  end;
  return coalesce(hdr->>header_name, '');
end;
$$;

create or replace function public.request_auth_ok()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.request_header('x-auth-hash') <> ''
    and public.request_header('x-auth-hash') = (
      select password_hash from public.app_auth where id = 1
    );
$$;

revoke all on function public.request_header(text) from public;
revoke all on function public.request_auth_ok() from public;
grant execute on function public.request_header(text) to anon, authenticated;
grant execute on function public.request_auth_ok() to anon, authenticated;

-- ── decision_journal: only the signed-in password, only its own rows ───────
alter table public.decision_journal enable row level security;

do $$
declare r record;
begin
  for r in select policyname from pg_policies
    where schemaname = 'public' and tablename = 'decision_journal'
  loop
    execute format('drop policy %I on public.decision_journal', r.policyname);
  end loop;
end $$;

create policy "signed-in journal" on public.decision_journal
  for all to anon, authenticated
  using (
    public.request_auth_ok()
    and pin_hash = public.request_header('x-auth-hash')
  )
  with check (
    public.request_auth_ok()
    and pin_hash = public.request_header('x-auth-hash')
  );

-- ── user_memory: shared rows, but only after the password checks out ───────
-- The app stores watchlist / portfolio / profile under pin_hash = 'shared_user'.
create table if not exists public.user_memory (
  pin_hash text not null,
  key text not null,
  value jsonb,
  updated_at timestamptz default now(),
  primary key (pin_hash, key)
);

alter table public.user_memory enable row level security;
grant select, insert, update, delete on public.user_memory to anon, authenticated;

do $$
declare r record;
begin
  for r in select policyname from pg_policies
    where schemaname = 'public' and tablename = 'user_memory'
  loop
    execute format('drop policy %I on public.user_memory', r.policyname);
  end loop;
end $$;

create policy "signed-in memory" on public.user_memory
  for all to anon, authenticated
  using (public.request_auth_ok() and pin_hash = 'shared_user')
  with check (public.request_auth_ok() and pin_hash = 'shared_user');

-- ── Public market data: anon may read, not write ───────────────────────────
do $$
declare
  t text;
  r record;
begin
  foreach t in array array[
    'company_profiles',
    'annual_financials',
    'quarterly_financials',
    'analyst_estimates',
    'sync_metadata'
  ]
  loop
    if to_regclass('public.' || t) is null then
      continue;
    end if;
    execute format('alter table public.%I enable row level security', t);
    for r in select policyname from pg_policies
      where schemaname = 'public' and tablename = t
    loop
      execute format('drop policy %I on public.%I', r.policyname, t);
    end loop;
    execute format(
      'create policy %I on public.%I for select to anon, authenticated using (true)',
      'public read ' || t, t
    );
  end loop;
end $$;
