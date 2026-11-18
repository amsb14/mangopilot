-- The signed-in app sends a clean lowercase SHA-256 in x-auth-hash, and that
-- header does reach Postgres (request_header('x-auth-hash') echoes it).
-- Watchlist and journal still come back empty because request_auth_ok()
-- requires that header to equal app_auth.password_hash exactly. A copy of
-- APP_PASSWORD_HASH pasted into SQL often picks up a newline, spaces, or
-- quotes, and the comparison then fails closed: HTTP 200 and zero rows,
-- which the UI shows as an empty watchlist and an empty journal.
--
-- This keeps the same check, but compares the hex digits only.
-- Run in the Supabase SQL editor. No new secret.
--
-- Then, still in the SQL editor (returns no secret, only shape):
--
--   select public.app_auth_status();
--
-- is_sha256_hex must be true and length 64. If it is false, the value in
-- app_auth is not a SHA-256 at all — paste APP_PASSWORD_HASH again:
--
--   update public.app_auth
--   set password_hash = 'PASTE_APP_PASSWORD_HASH_HERE', updated_at = now()
--   where id = 1;
--
-- Optional boolean check you run yourself (do not share the hash). Compute
-- the sign-in hash locally and compare; the query prints true or false only:
--
--   select public.normalized_auth_hash(password_hash) = 'LOCALLY_COMPUTED_SHA256' as matches
--   from public.app_auth where id = 1;

create or replace function public.normalized_auth_hash(raw text)
returns text
language sql
immutable
as $$
  select regexp_replace(lower(coalesce(raw, '')), '[^0-9a-f]', '', 'g');
$$;

revoke all on function public.normalized_auth_hash(text) from public, anon, authenticated;

create or replace function public.request_auth_ok()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select length(public.normalized_auth_hash(public.request_header('x-auth-hash'))) = 64
    and public.normalized_auth_hash(public.request_header('x-auth-hash')) = (
      select public.normalized_auth_hash(password_hash)
      from public.app_auth
      where id = 1
    );
$$;

revoke all on function public.request_auth_ok() from public;
grant execute on function public.request_auth_ok() to anon, authenticated;

-- Shape of the stored hash only. No characters of the secret.
create or replace function public.app_auth_status()
returns jsonb
language plpgsql
stable
security definer
set search_path = public
as $$
declare
  raw text;
  norm text;
begin
  select password_hash into raw from public.app_auth where id = 1;
  if raw is null then
    return jsonb_build_object('row_present', false);
  end if;
  norm := public.normalized_auth_hash(raw);
  return jsonb_build_object(
    'row_present', true,
    'length', length(raw),
    'trimmed_length', length(btrim(raw)),
    'hex_length', length(norm),
    'is_sha256_hex', length(norm) = 64,
    'has_whitespace', raw ~ '\s',
    'has_non_hex', raw ~ '[^0-9A-Fa-f]'
  );
end;
$$;

revoke all on function public.app_auth_status() from public, anon, authenticated;
-- SQL editor (postgres) can still call it. anon cannot.
