// Harvest every shipment from SNKRDUNK: tracking number + the orders in it.
// Writes prein-shipments.json keyed by tracking number, ready for build-prein-xlsx.py.
//
// Required env: SNKRDUNK_ENSID, SNKRDUNK_SESSION
// Optional config: prein.config.json in cwd (tracking_regex to filter results).
import { readFileSync, writeFileSync, existsSync } from 'node:fs';

const ENSID = process.env.SNKRDUNK_ENSID;
const SESSION = process.env.SNKRDUNK_SESSION;
if (!ENSID || !SESSION) {
  console.error('set SNKRDUNK_ENSID + SNKRDUNK_SESSION');
  process.exit(1);
}

const config = existsSync('prein.config.json')
  ? JSON.parse(readFileSync('prein.config.json', 'utf8'))
  : {};
const trackingRe = config.tracking_regex ? new RegExp(config.tracking_regex) : null;

const COOKIE = `ENSID=${ENSID}; session=${SESSION}`;
const H = { 'User-Agent': 'Mozilla/5.0', Cookie: COOKIE, Accept: 'application/json' };

async function get(path) {
  const r = await fetch('https://snkrdunk.com' + path, { headers: H });
  if (!r.ok) throw new Error(`${path} → ${r.status}`);
  return r.json();
}

function findCert(o) {
  if (!o || typeof o !== 'object') return null;
  const keys = ['psaCert', 'certNumber', 'certificationNumber', 'gradingCertNumber'];
  for (const k of keys) {
    const v = o[k];
    if (v != null && String(v).match(/^\d{6,12}$/)) return String(v);
  }
  for (const v of Object.values(o)) {
    if (v && typeof v === 'object') { const r = findCert(v); if (r) return r; }
  }
  return null;
}
function findGrade(o) {
  if (!o || typeof o !== 'object') return null;
  const keys = ['gradeName', 'grade', 'psaGrade'];
  for (const k of keys) if (typeof o[k] === 'string' && /PSA/i.test(o[k])) return o[k];
  for (const v of Object.values(o)) {
    if (v && typeof v === 'object') { const r = findGrade(v); if (r) return r; }
  }
  return null;
}

// 1. Carts (trading + completed).
const cartIds = new Set();
for (const filter of ['trading', 'completed']) {
  for (let page = 1; ; page++) {
    let r;
    try { r = await get(`/en/v1/account/orders?perPage=20&page=${page}&filter=${filter}&orderTransactionType=buying`); }
    catch (e) { console.error(`carts ${filter} p${page}: ${e.message}`); break; }
    const carts = r.orders || r.carts || r.items || r.data || [];
    if (!carts.length) break;
    for (const c of carts) cartIds.add(c.id || c.cartId);
    if (carts.length < 20) break;
  }
}
console.error(`Carts: ${cartIds.size}`);

// 2. Shipped orders.
const shipped = [];
for (const cid of cartIds) {
  for (const status of ['shipped', 'completed', 'delivered']) {
    for (let page = 1; ; page++) {
      let r;
      try { r = await get(`/en/v1/account/orders/cart/${cid}?status=${status}&perPage=50&page=${page}`); }
      catch { break; }
      const list = r.orders || [];
      if (!list.length) break;
      for (const o of list) {
        shipped.push({
          cart: cid, orderId: o.id, tx: o.transactionId, name: o.name,
          cert: findCert(o), grade: findGrade(o), status,
        });
      }
      if (list.length < 50) break;
    }
  }
}
console.error(`Shipped orders: ${shipped.length}`);

// 3. Chat → tracking.
const trackRe = /tracking number is\s*\[([^\]]+)\]/i;
const shipments = {};
let cursor = 0, done = 0;
const t0 = Date.now();
async function worker() {
  while (cursor < shipped.length) {
    const o = shipped[cursor++];
    try {
      const r = await get(`/en/v1/account/orders/${o.orderId}/contacts`);
      const msgs = r.contacts || r.messages || r.items || r.data || [];
      let tracking = null, sentAt = null;
      for (const m of msgs) {
        const text = (m.content || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
        const mm = text.match(trackRe);
        if (mm) { tracking = mm[1].trim(); sentAt = m.sentAt; break; }
      }
      if (tracking && (!trackingRe || trackingRe.test(tracking))) {
        if (!shipments[tracking]) shipments[tracking] = { date: sentAt, orderIds: [], rows: [] };
        shipments[tracking].orderIds.push(o.orderId);
        shipments[tracking].rows.push({
          orderId: o.orderId, tx: o.tx, name: o.name, cert: o.cert, grade: o.grade,
        });
      }
    } catch (e) {
      console.error(`  ✗ ${o.orderId}: ${e.message.slice(0, 80)}`);
    }
    done++;
    if (done % 25 === 0) {
      const rate = done / ((Date.now() - t0) / 1000);
      console.error(`  ${done}/${shipped.length} (${rate.toFixed(1)}/s)`);
    }
  }
}
const CONC = 8;
await Promise.all(Array.from({ length: CONC }, worker));

for (const t of Object.keys(shipments)) {
  const seen = new Set();
  shipments[t].rows = shipments[t].rows.filter(r => seen.has(r.orderId) ? false : seen.add(r.orderId));
  shipments[t].orderIds = [...new Set(shipments[t].orderIds)];
}

writeFileSync('prein-shipments.json', JSON.stringify(shipments, null, 2));
console.log(`\nWrote ${Object.keys(shipments).length} shipments → prein-shipments.json`);
for (const [t, s] of Object.entries(shipments)) {
  console.log(`  ${t}  (${s.rows.length} items, ${s.date || '?'})`);
}
