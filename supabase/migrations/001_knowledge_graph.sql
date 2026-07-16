-- Knowledge Graph: ticker → related companies (competitors, suppliers, customers, partners)
-- Run this in Supabase SQL Editor (Dashboard → SQL Editor → New query → paste → Run)

CREATE TABLE IF NOT EXISTS ticker_relationships (
  id              BIGSERIAL PRIMARY KEY,
  source_ticker   TEXT NOT NULL,
  related_ticker  TEXT NOT NULL,
  type            TEXT NOT NULL CHECK (type IN ('competitor', 'supplier', 'customer', 'partner', 'co_dependent')),
  evidence        TEXT,
  strength        INT  DEFAULT 5 CHECK (strength BETWEEN 1 AND 10),
  created_at      TIMESTAMPTZ DEFAULT NOW(),
  updated_at      TIMESTAMPTZ DEFAULT NOW(),
  UNIQUE(source_ticker, related_ticker, type)
);

CREATE INDEX IF NOT EXISTS idx_tr_source  ON ticker_relationships(source_ticker);
CREATE INDEX IF NOT EXISTS idx_tr_related ON ticker_relationships(related_ticker);
CREATE INDEX IF NOT EXISTS idx_tr_type    ON ticker_relationships(type);

-- RLS: allow anon read, only service role can write
ALTER TABLE ticker_relationships ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon read relationships" ON ticker_relationships;
CREATE POLICY "anon read relationships" ON ticker_relationships
  FOR SELECT TO anon USING (true);

-- (writes happen via the edge function using SERVICE_ROLE_KEY which bypasses RLS)
