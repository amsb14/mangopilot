-- Thesis, exit conditions, and saved price checkpoints on the decision journal.
-- Same access rule as migration 007/008: the anon key can read or write a row
-- only when x-auth-hash matches app_auth (request_auth_ok) and the row's pin_hash.
-- No new edge function. Run this in the Supabase SQL editor after the site deploy.

alter table public.decision_journal
  add column if not exists thesis text,
  add column if not exists conditions jsonb not null default '[]'::jsonb,
  add column if not exists price_marks jsonb not null default '{}'::jsonb;

alter table public.decision_journal enable row level security;

-- Drop leftover open policies, then recreate the signed-in policy so a re-run
-- cannot leave USING (true) in place.
drop policy if exists "anon all on own pin" on public.decision_journal;
drop policy if exists "signed-in journal" on public.decision_journal;

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

-- Confirm the new columns exist and anon still cannot read the table with no header.
-- (Run by hand. The second query should return 0 rows when called with the anon key
-- and no x-auth-hash.)
--   select column_name from information_schema.columns
--   where table_name = 'decision_journal'
--     and column_name in ('thesis', 'conditions', 'price_marks');
