// netlify/functions/data.js
// All secrets come from Netlify environment variables - nothing hardcoded here.

const { createClient } = require('@supabase/supabase-js');
const XLSX = require('xlsx');

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_KEY
);

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, x-upload-password',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
};

function json(status, body) {
  return {
    statusCode: status,
    headers: { ...CORS, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  };
}

// ── GET: load all leases + cashflows ────────────────────────
async function handleGet() {
  const [{ data: leases, error: le }, { data: cashflows, error: ce }] =
    await Promise.all([
      supabase.from('leases').select('*').order('city'),
      supabase.from('cashflows').select('*').order('payment_date'),
    ]);

  if (le || ce) return json(500, { error: 'Database read failed.' });

  const cfMap = {};
  (cashflows || []).forEach(r => {
    if (!cfMap[r.lease_id]) cfMap[r.lease_id] = [];
    cfMap[r.lease_id].push({ d: r.payment_date, b: +r.base_rent, t: +r.total_cash });
  });

  const mapped = (leases || []).map(l => ({
    id: l.id, city: l.city, state: l.state,
    tenant: l.tenant, tenantShort: l.tenant_short,
    address: l.address, sf: +l.sf,
    start: l.start_date, end: l.end_date,
    monthlyBase: +l.monthly_base, annualCost: +l.annual_cost,
    freeMonths: +l.free_months, npvRemaining: +l.npv_remaining || 0,
    lat: +l.lat, lng: +l.lng,
  }));

  return json(200, { leases: mapped, cashflows: cfMap });
}

// ── POST: password check or Excel upload ────────────────────
async function handlePost(event) {
  const pw = event.headers['x-upload-password'] || '';
  if (pw !== process.env.UPLOAD_PASSWORD) {
    return json(401, { error: 'Incorrect password.' });
  }

  let body;
  try { body = JSON.parse(event.body); }
  catch { return json(400, { error: 'Invalid request.' }); }

  if (body.checkOnly) return json(200, { ok: true });
  if (!body.file) return json(400, { error: 'No file provided.' });

  let wb;
  try {
    wb = XLSX.read(Buffer.from(body.file, 'base64'), { type: 'buffer', cellDates: true });
  } catch (e) {
    return json(400, { error: 'Could not read Excel file: ' + e.message });
  }

  // ── Parse Summary sheet ──────────────────────────────────
  const summarySheet = wb.Sheets['Summary'];
  if (!summarySheet) return json(400, { error: 'No Summary sheet found.' });

  const rows = XLSX.utils.sheet_to_json(summarySheet, { header: 1, defval: null });

  // Find header row containing 'Tab Name'
  let hdrRow = -1;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i] && rows[i].some(c => String(c || '').includes('Tab Name'))) {
      hdrRow = i; break;
    }
  }
  if (hdrRow < 0) return json(400, { error: 'Cannot find header row in Summary sheet.' });

  const hdrs = rows[hdrRow].map(h => String(h || '').trim());
  const ci = k => hdrs.findIndex(h => h.includes(k));
  const iTab    = ci('Tab Name');
  const iTenant = ci('Tenant');
  const iRSF    = ci('RSF');
  const iStart  = ci('Commencement');
  const iEnd    = ci('Lease End');
  const iFree   = ci('Free Mo');
  const iNPV    = ci('NPV Total Cash');  // NPV Total Cash — Rem ($)

  const COORDS = {
    'Abilene':     [32.4487,-99.7331], 'Allen TX':    [33.1032,-96.6706],
    'Atlanta':     [33.9526,-84.5499], 'Austin':      [30.3579,-97.6924],
    'Birmingham':  [33.5186,-86.8104], 'Chattanooga': [35.0456,-85.3097],
    'Dallas':      [32.9177,-96.7767], 'Houston':     [29.7419,-95.5588],
    'Huntsville':  [34.7304,-86.5861], 'Nashville':   [36.0274,-86.7308],
    'Orlando':     [28.5383,-81.3792], 'Pittsburgh':  [40.5109,-79.9352],
    'Tampa':       [27.9506,-82.4572], 'Vienna':      [38.9012,-77.2653],
  };
  const STATES = {
    'Abilene':'Texas','Allen TX':'Texas','Atlanta':'Georgia','Austin':'Texas',
    'Birmingham':'Alabama','Chattanooga':'Tennessee','Dallas':'Texas','Houston':'Texas',
    'Huntsville':'Alabama','Nashville':'Tennessee','Orlando':'Florida',
    'Pittsburgh':'Pennsylvania','Tampa':'Florida','Vienna':'Virginia',
  };
  const ADDRESSES = {
    'Abilene':     'Alexander Building, 104 Pine Street, Suite A10, Abilene, TX 79601',
    'Allen TX':    '950 W. Bethany Drive, Suite 580, Allen, TX 75013',
    'Atlanta':     '1100 Circle 75 Parkway, Suite 300, Atlanta, GA 30339',
    'Austin':      '2800-B Industrial Terrace, Suites A & B, Austin, TX 78758',
    'Birmingham':  'Two Metroplex Drive, Birmingham, AL 35209',
    'Chattanooga': 'Suites 201 & 202, 1300 Broad Street, Chattanooga, TN 37402',
    'Dallas':      '5495 Belt Line Road, Suite 335, Dallas, TX 75254',
    'Houston':     '10370 Richmond Avenue, Level 7, Houston, TX',
    'Huntsville':  '200 Clinton Avenue, Suite 703, Huntsville, AL 35801',
    'Nashville':   '215 Centerview Drive, Suite 3-300 & 3-360, Brentwood, TN',
    'Orlando':     '111 N. Orange Avenue, Suites 1125/1150, Orlando, FL',
    'Pittsburgh':  '1000 Main Street, Pittsburgh, PA 15215',
    'Tampa':       '100 South Ashley Drive, Suite 1120, Tampa, FL 33602',
    'Vienna':      '8200 Greensboro Drive, Suite 1150, McLean, VA 22102',
  };
  const TS = t => t.includes('Holdings') ? 'HH'
                : (t.includes('Program') || t.includes('HPM')) ? 'HPM' : 'HC';

  const parsedLeases = [];

  for (let i = hdrRow + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[iTab]) continue;
    const tab = String(r[iTab]).trim();
    if (!tab || tab === 'null') continue;

    const toDate = v => v instanceof Date ? v.toISOString().split('T')[0]
                      : v ? String(v).split('T')[0] : null;
    const sd = toDate(r[iStart]);
    const ed = toDate(r[iEnd]);
    if (!sd || !ed) continue;

    const tenantFull = String(r[iTenant] || '').replace(/, LLC|, Inc\./g, '').trim();
    const coords = COORDS[tab] || [0, 0];

    // ── Pull monthly base rent from the individual tab ──
    let monthlyBase = 0;
    let annualCost  = 0;
    const leaseSheet = wb.Sheets[tab];
    if (leaseSheet) {
      const lrows = XLSX.utils.sheet_to_json(leaseSheet, { header: 1, defval: null });
      // Find rent schedule header row (contains 'Monthly Base Rent')
      for (let j = 0; j < lrows.length; j++) {
        const lr = lrows[j];
        if (lr && lr.some(c => String(c||'').includes('Monthly Base Rent'))) {
          // First data row after header is the current/first rent period
          const dataRow = lrows[j + 1];
          if (dataRow) {
            const mIdx = lr.findIndex(c => String(c||'').includes('Monthly Base Rent'));
            monthlyBase = parseFloat(dataRow[mIdx]) || 0;
            annualCost  = Math.round(monthlyBase * 12 * 100) / 100;
          }
          break;
        }
      }
    }

    // NPV remaining from Summary sheet
    const npvRemaining = iNPV >= 0 ? (parseFloat(r[iNPV]) || 0) : 0;

    parsedLeases.push({
      id:            tab,
      city:          tab === 'Allen TX' ? 'Allen' : tab === 'Vienna' ? 'McLean' : tab,
      state:         STATES[tab] || '',
      tenant:        tenantFull.includes('Program') || tenantFull.includes('HPM') ? 'HPM'
                   : tenantFull.includes('Holdings') ? 'Hoar Holdings' : 'Hoar Construction',
      tenant_short:  TS(tenantFull),
      address:       ADDRESSES[tab] || '',
      sf:            parseInt(r[iRSF]) || 0,
      start_date:    sd,
      end_date:      ed,
      monthly_base:  Math.round(monthlyBase * 100) / 100,
      annual_cost:   annualCost,
      npv_remaining: Math.round(npvRemaining * 100) / 100,
      free_months:   parseInt(r[iFree]) || 0,
      lat:           coords[0],
      lng:           coords[1],
      updated_at:    new Date().toISOString(),
    });
  }

  if (!parsedLeases.length) return json(400, { error: 'No valid leases found in Summary sheet.' });

  // ── Parse monthly cash flows from each tab ───────────────
  const allCF = [];
  for (const lease of parsedLeases) {
    const sheet = wb.Sheets[lease.id];
    if (!sheet) continue;
    const srows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    let cfHdr = -1;
    for (let i = 0; i < srows.length; i++) {
      if (srows[i] && srows[i].some(c => String(c||'').includes('Base Rent') &&
          srows[i].some(c2 => String(c2||'').includes('Date')))) {
        cfHdr = i; break;
      }
    }
    // Fallback: find row with Date and Base Rent columns
    if (cfHdr < 0) {
      for (let i = 0; i < srows.length; i++) {
        if (srows[i] && srows[i].some(c => String(c||'') === 'Date' ||
            String(c||'').includes('Month #'))) {
          cfHdr = i; break;
        }
      }
    }
    if (cfHdr < 0) continue;

    const sh = srows[cfHdr].map(h => String(h||'').trim());
    const iD = sh.findIndex(h => h === 'Date' || h.includes('Date'));
    const iB = sh.findIndex(h => h.includes('Base Rent'));
    const iT = sh.findIndex(h => h.includes('Total Cash'));
    if (iD < 0) continue;

    for (let i = cfHdr + 1; i < srows.length; i++) {
      const r = srows[i];
      if (!r || !r[iD]) continue;
      const dv = r[iD];
      const ds = dv instanceof Date ? dv.toISOString().split('T')[0]
                : typeof dv === 'string' ? dv.split('T')[0] : null;
      if (!ds || !ds.match(/^\d{4}-\d{2}-\d{2}$/)) continue;
      const b = iB >= 0 ? (+r[iB] || 0) : 0;
      const t = iT >= 0 ? (+r[iT] || 0) : b;
      if (b === 0 && t === 0) continue;
      allCF.push({ lease_id: lease.id, payment_date: ds, base_rent: b, total_cash: t });
    }
  }

  // ── Add npv_remaining column if missing ─────────────────
  // First ensure the column exists (safe to run multiple times)
  await supabase.rpc('exec_sql', {
    sql: 'ALTER TABLE leases ADD COLUMN IF NOT EXISTS npv_remaining NUMERIC(14,2) DEFAULT 0'
  }).catch(() => {}); // ignore if RPC not available, column may already exist

  // ── Write to Supabase ────────────────────────────────────
  const { error: le } = await supabase.from('leases').upsert(parsedLeases, { onConflict: 'id' });
  if (le) return json(500, { error: 'Failed to save leases: ' + le.message });

  const ids = parsedLeases.map(l => l.id);
  const { error: de } = await supabase.from('cashflows').delete().in('lease_id', ids);
  if (de) return json(500, { error: 'Failed to clear cash flows: ' + de.message });

  for (let i = 0; i < allCF.length; i += 500) {
    const { error: ce } = await supabase.from('cashflows').insert(allCF.slice(i, i + 500));
    if (ce) return json(500, { error: 'Failed to save cash flows: ' + ce.message });
  }

  return json(200, {
    success: true,
    message: `Updated ${parsedLeases.length} leases and ${allCF.length} monthly payment records.`,
  });
}

// ── Router ───────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: CORS, body: '' };
  if (event.httpMethod === 'GET')  return handleGet();
  if (event.httpMethod === 'POST') return handlePost(event);
  return json(405, { error: 'Method not allowed.' });
};
