CREATE TABLE IF NOT EXISTS records (
  code        text PRIMARY KEY,
  gen         text,
  kiban       text,
  method      text,
  time        bigint,
  lat         double precision,
  lng         double precision,
  acc         double precision,
  listed      boolean DEFAULT true,
  photos      jsonb   DEFAULT '[]'::jsonb,
  updated_at  timestamptz DEFAULT now()
);
CREATE TABLE IF NOT EXISTS photos (
  name        text PRIMARY KEY,
  code        text NOT NULL,
  kind        text,
  time        bigint,
  data        bytea NOT NULL,
  updated_at  timestamptz DEFAULT now()
);
CREATE INDEX IF NOT EXISTS photos_code_idx ON photos(code);
