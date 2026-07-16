-- Adds 'thematic_peer' relationship type for secular-tailwind beneficiaries
-- Run this in Supabase SQL Editor

ALTER TABLE ticker_relationships DROP CONSTRAINT IF EXISTS ticker_relationships_type_check;
ALTER TABLE ticker_relationships ADD CONSTRAINT ticker_relationships_type_check
  CHECK (type IN ('competitor', 'supplier', 'customer', 'partner', 'co_dependent', 'thematic_peer'));
