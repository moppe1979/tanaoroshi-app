'use strict';
const fs = require('fs'), path = require('path'), http = require('http'), https = require('https');
const express = require('express');
const { Pool } = require('pg');

const env = process.env;
if (!env.API_KEY) { console.error('API_KEY が未設定です (.env を確認してください)'); process.exit(1); }
const pool = new Pool({ max: 5, connectionTimeoutMillis: 8000 }); // PGHOST 等の環境変数を使用
const app = express();

/* CORS */
app.use((req, res, next) => {
  const o = req.headers.origin;
  if (o && (!env.ALLOW_ORIGIN || env.ALLOW_ORIGIN.split(',').map(s => s.trim()).includes(o))) {
    res.setHeader('Access-Control-Allow-Origin', o);
    res.setHeader('Vary', 'Origin');
  }
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-API-Key');
  res.setHeader('Access-Control-Allow-Methods', 'GET,PUT,DELETE,OPTIONS');
  res.setHeader('Access-Control-Allow-Private-Network', 'true');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

/* 認証(合言葉) */
const auth = (req, res, next) => {
  const k = req.get('X-API-Key') || req.query.key;
  if (k !== env.API_KEY) return res.status(401).json({ error: 'unauthorized' });
  next();
};
const wrap = fn => (req, res) => fn(req, res).catch(e => { console.error(e); res.status(500).json({ error: String(e.message || e) }); });

app.get('/api/health', auth, wrap(async (req, res) => {
  const r = await pool.query('SELECT now() AS now');
  const c = await pool.query('SELECT (SELECT count(*) FROM records)::int AS records, (SELECT count(*) FROM photos)::int AS photos');
  res.json({ ok: true, now: r.rows[0].now, ...c.rows[0] });
}));

app.get('/api/records', auth, wrap(async (req, res) => {
  const r = await pool.query('SELECT code,gen,kiban,method,time::double precision AS time,lat,lng,acc,listed,photos FROM records ORDER BY code');
  res.json(r.rows);
}));

app.put('/api/records/:code', auth, express.json({ limit: '1mb' }), wrap(async (req, res) => {
  const b = req.body || {}, code = req.params.code;
  await pool.query(
    `INSERT INTO records(code,gen,kiban,method,time,lat,lng,acc,listed,photos,updated_at)
     VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10::jsonb,now())
     ON CONFLICT(code) DO UPDATE SET gen=$2,kiban=$3,method=$4,time=$5,lat=$6,lng=$7,acc=$8,listed=$9,photos=$10::jsonb,updated_at=now()`,
    [code, b.gen ?? null, b.kiban ?? null, b.method ?? null, b.time ?? null, b.lat ?? null, b.lng ?? null, b.acc ?? null, b.listed !== false, JSON.stringify(b.photos || [])]);
  res.json({ ok: true });
}));

app.delete('/api/records/:code', auth, wrap(async (req, res) => {
  await pool.query('DELETE FROM photos WHERE code=$1', [req.params.code]);
  await pool.query('DELETE FROM records WHERE code=$1', [req.params.code]);
  res.json({ ok: true });
}));

app.get('/api/photos', auth, wrap(async (req, res) => {
  const r = await pool.query('SELECT name,code,kind,time::double precision AS time,octet_length(data) AS size FROM photos ' + (req.query.code ? 'WHERE code=$1 ' : '') + 'ORDER BY name', req.query.code ? [req.query.code] : []);
  res.json(r.rows);
}));

app.put('/api/photos/:name', auth, express.raw({ type: '*/*', limit: '40mb' }), wrap(async (req, res) => {
  if (!Buffer.isBuffer(req.body) || !req.body.length) return res.status(400).json({ error: 'empty body' });
  const q = req.query;
  await pool.query(
    `INSERT INTO photos(name,code,kind,time,data,updated_at) VALUES($1,$2,$3,$4,$5,now())
     ON CONFLICT(name) DO UPDATE SET code=$2,kind=$3,time=$4,data=$5,updated_at=now()`,
    [req.params.name, q.code || '', q.kind || null, q.time ? Number(q.time) : null, req.body]);
  res.json({ ok: true });
}));

app.get('/api/photos/:name', auth, wrap(async (req, res) => {
  const r = await pool.query('SELECT data FROM photos WHERE name=$1', [req.params.name]);
  if (!r.rows.length) return res.sendStatus(404);
  res.type('image/jpeg').send(r.rows[0].data);
}));

app.delete('/api/photos/:name', auth, wrap(async (req, res) => {
  await pool.query('DELETE FROM photos WHERE name=$1', [req.params.name]);
  res.json({ ok: true });
}));

/* アプリ画面の配信(任意) */
if (env.SERVE_APP === '1') {
  app.use(express.static(path.join(__dirname, '..'), { index: 'index.html', dotfiles: 'ignore', setHeaders: (res, p) => { if (/[\\/]server[\\/]/.test(p)) res.statusCode = 404; } }));
  app.use('/server', (req, res) => res.sendStatus(404));
}

(async () => {
  await pool.query(fs.readFileSync(path.join(__dirname, 'schema.sql'), 'utf8'));
  const port = Number(env.PORT) || 8443;
  const srv = (env.TLS_CERT && env.TLS_KEY)
    ? https.createServer({ cert: fs.readFileSync(env.TLS_CERT), key: fs.readFileSync(env.TLS_KEY) }, app)
    : http.createServer(app);
  srv.listen(port, () => console.log(`棚卸API 起動: ${env.TLS_CERT ? 'https' : 'http'}://0.0.0.0:${port}`));
})().catch(e => { console.error('起動に失敗しました:', e.message); process.exit(1); });
