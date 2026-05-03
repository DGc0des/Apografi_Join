-- Run this once in your Supabase project: SQL Editor → New query → paste & run

-- Inventory counts per store per date
CREATE TABLE inventory_counts (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  store_id integer NOT NULL,
  count_date date NOT NULL DEFAULT CURRENT_DATE,
  item_row integer NOT NULL,
  quantity numeric(10,3),
  updated_at timestamptz DEFAULT now(),
  UNIQUE(store_id, count_date, item_row)
);
ALTER TABLE inventory_counts ENABLE ROW LEVEL SECURITY;
CREATE POLICY "anon_all" ON inventory_counts FOR ALL TO anon USING (true) WITH CHECK (true);

-- Enable real-time for inventory_counts
ALTER PUBLICATION supabase_realtime ADD TABLE inventory_counts;

-- Container (tare) weights — per store
CREATE TABLE container_weights (
  store_id integer NOT NULL,
  item_row integer NOT NULL,
  weight_kg numeric(10,3) NOT NULL,
  updated_at timestamptz DEFAULT now(),
  PRIMARY KEY (store_id, item_row)
);
ALTER TABLE container_weights ENABLE ROW LEVEL SECURITY;
CREATE POLICY "anon_all" ON container_weights FOR ALL TO anon USING (true) WITH CHECK (true);

-- Store access codes (hashed) — one row per store key + 'master'
CREATE TABLE store_codes (
  key text PRIMARY KEY,  -- 'store_1'..'store_5' or 'master'
  code_hash text NOT NULL,
  updated_at timestamptz DEFAULT now()
);
ALTER TABLE store_codes ENABLE ROW LEVEL SECURITY;
CREATE POLICY "anon_all" ON store_codes FOR ALL TO anon USING (true) WITH CHECK (true);

-- Item slot/tare configuration — per store per item
CREATE TABLE item_config (
  store_id integer NOT NULL,
  item_row integer NOT NULL,
  num_inputs integer NOT NULL DEFAULT 1,
  tare_count integer NOT NULL DEFAULT 0,
  PRIMARY KEY (store_id, item_row)
);
ALTER TABLE item_config ENABLE ROW LEVEL SECURITY;
CREATE POLICY "anon_all" ON item_config FOR ALL TO anon USING (true) WITH CHECK (true);

-- ⚠️  If you already ran this file before, run this migration instead:
-- ALTER TABLE container_weights DROP CONSTRAINT container_weights_pkey;
-- ALTER TABLE container_weights ADD COLUMN store_id integer NOT NULL DEFAULT 1;
-- ALTER TABLE container_weights ADD PRIMARY KEY (store_id, item_row);
-- ALTER TABLE container_weights ALTER COLUMN store_id DROP DEFAULT;
