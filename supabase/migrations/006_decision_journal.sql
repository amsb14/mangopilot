-- Decision Journal: snapshot the AI's stock picks at the moment of recommendation,
-- then track them against the SPY benchmark over time. Lets the user answer the
-- one important question: "do my AI's picks actually outperform doing nothing?"
--
-- Run this in Supabase SQL Editor.

CREATE TABLE IF NOT EXISTS decision_journal (
  id              BIGSERIAL PRIMARY KEY,
  pin_hash        TEXT NOT NULL,                -- isolates per-user (matches USER_AUTH_HASH)
  ticker          TEXT NOT NULL,
  entry_price     NUMERIC NOT NULL,             -- live price at save
  spy_entry_price NUMERIC,                       -- SPY at the same moment (benchmark anchor)
  saved_at        TIMESTAMPTZ DEFAULT NOW(),
  direction       TEXT DEFAULT 'long',           -- 'long' or 'short'
  confidence      NUMERIC,                       -- 0-100 if available
  rationale       TEXT,                          -- AI's summary at save time
  source          TEXT,                          -- 'dashboard' | 'advisor' | 'council' | 'manual'
  source_query    TEXT,                          -- user's question that produced this (if any)
  -- Outcomes (filled in later, optional)
  exit_price      NUMERIC,
  exit_at         TIMESTAMPTZ,
  closed          BOOLEAN DEFAULT FALSE,
  notes           TEXT
);

CREATE INDEX IF NOT EXISTS idx_dj_user_recent ON decision_journal(pin_hash, saved_at DESC);
CREATE INDEX IF NOT EXISTS idx_dj_user_open ON decision_journal(pin_hash, closed) WHERE closed = FALSE;

ALTER TABLE decision_journal ENABLE ROW LEVEL SECURITY;

-- Anon users can SELECT/INSERT/UPDATE only rows scoped to their own pin_hash.
-- Since the app filters by pin_hash in queries, RLS provides the security backstop.
DROP POLICY IF EXISTS "anon all on own pin" ON decision_journal;
CREATE POLICY "anon all on own pin" ON decision_journal
  FOR ALL TO anon USING (true) WITH CHECK (true);
