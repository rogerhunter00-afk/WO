(function(global){
  'use strict';

  const SHEET_INDEX_CSV_URL =
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vRJocigDhxneJtrUmezFU7FcWpzSSah8-Wb6Rce8NA1f7jKcINgYU29iYRqt5QQymWATX5zs5k8_rK0/pub?single=true&output=csv&gid=105348743';

  const ACTIVE_STATUS_ALLOWLIST = [
    'open',
    'in progress',
    'planned',
    'planning',
    'scheduled',
    'ready',
    'awaiting',
    'on hold',
    'pending',
    'assigned',
    'new'
  ];

  const INACTIVE_STATUS_BLOCKLIST = [
    'cancel',
    'void',
    'archive',
    'inactive',
    'closed',
    'complete',
    'done',
    'rejected'
  ];

  const WEEK_MS = 7 * 24 * 60 * 60 * 1000;

  function toStringSafe(value){
    return String(value ?? '').trim();
  }

  function normalizeKey(value){
    return toStringSafe(value).toLowerCase();
  }

  function slugify(s){
    return String(s || '')
      .toLowerCase()
      .trim()
      .normalize('NFKD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '');
  }

  function parseDelimited(text){
    const sep = text.includes('\t') ? '\t' : ',';
    const rows = text.trim().split(/\r?\n/).map(r =>
      r.split(new RegExp(`${sep}(?=(?:[^"]*"[^"]*")*[^"]*$)`)).map(c => c.replace(/^"|"$/g,''))
    );
    const rawHeaders = rows.shift() || [];
    const headers = rawHeaders.map((h, i)=>String(h ?? '').replace(/^\uFEFF/, '').trim() || `col${i}`);
    const data = rows.map(r => Object.fromEntries(r.map((v, i)=>[headers[i] || `col${i}`, v])));
    return { headers, rows: data };
  }

  function parseDateLoose(v){
    if(!v) return null;
    const str = String(v).trim();
    if(!str) return null;

    const iso = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if(iso){
      const [,y,M,d,hh='00',mm='00',ss='00'] = iso;
      const dt = new Date(`${y}-${M}-${d}T${hh.padStart(2,'0')}:${mm}:${ss.padStart(2,'0')}`);
      return Number.isNaN(dt.getTime()) ? null : dt;
    }

    const slashed = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/);
    if(!slashed) return null;
    let [,d,M,y,hh='00',mm='00'] = slashed;
    d = d.padStart(2,'0');
    M = M.padStart(2,'0');
    hh = hh.padStart(2,'0');
    const dt = new Date(`${y}-${M}-${d}T${hh}:${mm}:00`);
    return Number.isNaN(dt.getTime()) ? null : dt;
  }

  function parseWeekEndDate(weekEnd){
    if(!weekEnd) return null;
    const parsed = parseDateLoose(weekEnd);
    if(parsed) return parsed;
    const dt = new Date(String(weekEnd));
    return Number.isNaN(dt.getTime()) ? null : dt;
  }

  function withBust(u){
    const x = new URL(u);
    x.searchParams.set('_ts', Date.now());
    return x.toString();
  }

  function csvUrlForGid(gid){
    const u = new URL(SHEET_INDEX_CSV_URL);
    u.searchParams.set('gid', String(gid));
    u.searchParams.set('single', 'true');
    u.searchParams.set('output', 'csv');
    return u.toString();
  }

  function sanitizePlanScalar(value){
    if(value == null) return null;
    if(typeof value === 'string'){
      const trimmed = value.trim();
      if(!trimmed || trimmed === '-' || /^null$/i.test(trimmed) || /^undefined$/i.test(trimmed)) return null;
      return trimmed;
    }
    if(typeof value === 'number') return Number.isFinite(value) ? String(value) : null;
    if(typeof value === 'boolean') return value ? '1' : null;
    return null;
  }

  async function loadWeeksIndex({ bust = false } = {}){
    const finalUrl = bust ? withBust(SHEET_INDEX_CSV_URL) : SHEET_INDEX_CSV_URL;
    const res = await fetch(finalUrl, { cache: bust ? 'reload' : 'no-store' });
    if(!res.ok) throw new Error('Index load failed: ' + res.status);

    const text = await res.text();
    if(!text || !text.trim()) throw new Error('Index CSV was empty');

    const { headers, rows } = parseDelimited(text);
    const normHeaders = headers.map(h => String(h || '').trim());
    const findHeader = (...candidates) => {
      const wanted = candidates.map(c => String(c).trim().toLowerCase());
      const idx = normHeaders.findIndex(h => wanted.includes(h.toLowerCase()));
      return idx >= 0 ? headers[idx] : null;
    };

    const H_LABEL = findHeader('label', 'week label', 'title', 'name');
    const H_WEEKEND = findHeader('week_end', 'week end', 'week', 'weekending', 'week-ending', 'week_end_date', 'date', 'week ending (iso)');
    const H_GID = findHeader('gid', 'sheet_gid', 'tab_gid', 'tab id', 'sheet id', 'sheetid', 'tabid', 'sheet');
    const H_URL = findHeader('url', 'sheet url', 'link');

    if(!H_GID && !H_URL){
      throw new Error('Index is missing a "gid" or "url" column');
    }

    const gidFromUrl = (u) => {
      if(!u) return null;
      const m = String(u).match(/[?#&]gid=(\d+)/);
      return m ? m[1] : null;
    };

    const out = rows.map(r => {
      const label = H_LABEL ? r[H_LABEL] : (r[H_WEEKEND] || '').trim();
      const weekEnd = H_WEEKEND ? r[H_WEEKEND] : '';
      const urlRaw = H_URL ? String(r[H_URL] || '').trim() : '';
      let gid = sanitizePlanScalar(H_GID ? r[H_GID] : null);
      if(!gid && urlRaw) gid = gidFromUrl(urlRaw);
      const url = sanitizePlanScalar(urlRaw) || (urlRaw || null);

      return { label, weekEnd, gid: gid || null, url };
    }).filter(x => x && (x.gid || x.url));

    out.sort((a,b) => {
      const aDate = parseWeekEndDate(a.weekEnd)?.getTime() || 0;
      const bDate = parseWeekEndDate(b.weekEnd)?.getTime() || 0;
      if(aDate !== bDate) return bDate - aDate;
      return String(b.label || '').localeCompare(String(a.label || ''), undefined, { numeric:true, sensitivity:'base' });
    });

    return out;
  }

  function chooseInitialWeek(weeks){
    const q = new URLSearchParams(global.location?.search || '');
    const byGid = q.get('gid');
    const byDate = q.get('week');
    const forceLatest = q.get('latest') === '1';

    let chosen = null;
    if(byGid) chosen = weeks.find(w => String(w.gid) === String(byGid));
    if(!chosen && byDate) chosen = weeks.find(w => String(w.weekEnd) === String(byDate));
    if(!chosen && !forceLatest){
      try {
        const last = global.localStorage?.getItem('alsLastWeekEnd');
        if(last) chosen = weeks.find(w => String(w.weekEnd) === last);
      } catch {}
    }
    return chosen || weeks[0] || null;
  }

  function normalizeId(v){
    return String(v ?? '')
      .trim()
      .replace(/^WO-?/i,'')
      .replace(/[^\w]/g,'')
      .toLowerCase();
  }

  function bestTimestamp(r){
    return (parseDateLoose(r['Last updated'])?.getTime())
      || (parseDateLoose(r['Completed on'])?.getTime())
      || (parseDateLoose(r['Created on'])?.getTime())
      || 0;
  }

  function dedupeRows(rows){
    const map = new Map();
    for(const r of (rows || [])){
      const id = normalizeId(r['ID']);
      const fallback = String(r['Title'] || '').trim().toLowerCase()
        + '|' + (parseDateLoose(r['Created on'])?.toISOString()?.slice(0,10) || '');
      const key = id || fallback;

      const prev = map.get(key);
      if(!prev){
        map.set(key, r);
        continue;
      }
      if(bestTimestamp(r) > bestTimestamp(prev)){
        map.set(key, r);
      }
    }
    return [...map.values()];
  }

  function toSiteKey(loc){
    const s = String(loc || '').toLowerCase();
    if(s.includes('byron')) return 'by';
    if(s.includes('mugiemoss') || s.includes('bucksburn')) return 'mm';
    if(s.includes('keith')) return 'keith';
    if(s.includes('cathkin') || s.includes('east kilbride') || s.includes('ek')) return 'ek';
    return 'other';
  }

  const SITE_KEY_LABELS = Object.freeze({
    all: 'All',
    ek: 'East Kilbride',
    mm: 'Mugiemoss',
    keith: 'Keith',
    by: 'Byron',
    other: 'Other'
  });

  function statusClass(v){
    const map = new Map([
      ['done','done'],
      ['complete','done'],
      ['completed','done'],
      ['open','open'],
      ['in-progress','in-progress'],
      ['in progress','in-progress'],
      ['progress','in-progress'],
      ['inprogress','in-progress'],
      ['on-hold','on-hold'],
      ['on hold','on-hold'],
      ['hold','on-hold'],
      ['onhold','on-hold']
    ]);
    const s = slugify(String(v || ''));
    return map.get(s) || s;
  }

  function parseDateValue(value){
    if(value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
    return parseDateLoose(value);
  }

  function startOfDay(date){
    if(!(date instanceof Date) || Number.isNaN(date.getTime())) return null;
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  function startOfWeek(date){
    const d = new Date(date);
    const day = d.getDay();
    const mondayOffset = (day + 6) % 7;
    d.setDate(d.getDate() - mondayOffset);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  function extractField(raw, names){
    for(const name of names){
      const value = raw?.[name];
      const text = toStringSafe(value);
      if(text) return text;
    }
    return '';
  }

  function getNormalizedWorkOrders(rawData){
    const list = Array.isArray(rawData) ? rawData : [];
    return list.map((row, idx) => {
      const woNumber = extractField(row, ['ID', 'WO', 'WO Number', 'Work Order', 'Work Order Number', 'WO/Link']);
      const title = extractField(row, ['Title', 'Description', 'Task', 'Summary']);
      const category = extractField(row, ['Categories', 'Category', 'Type', 'Classification']);
      const status = extractField(row, ['Status', 'State']);
      const site = extractField(row, ['Location', 'Site', 'Facility']);
      const asset = extractField(row, ['Asset', 'Asset Name', 'Equipment', 'Equipment Name']);
      const dueDateRaw = extractField(row, ['Due date', 'Due Date', 'Planned Start Date', 'Planned Start', 'Planned Start Date (Local)', 'Scheduled date', 'Scheduled Date']);
      const dueDate = parseDateValue(dueDateRaw);
      return {
        id: woNumber || `row-${idx}`,
        woNumber,
        title,
        category,
        status,
        statusKey: statusClass(status),
        site,
        siteKey: toSiteKey(site),
        assetName: asset,
        dueDateRaw,
        dueDate,
        raw: row
      };
    });
  }

  function isPPMCategory(categoryValue){
    const raw = toStringSafe(categoryValue);
    if(!raw) return false;
    const normalized = raw.replace(/\//g, ',');
    const parts = normalized.split(/[;,]+/).map(part => normalizeKey(part)).filter(Boolean);
    if(parts.length === 0) return /\bppm\b/i.test(raw);
    return parts.some(part => /\bppm\b/i.test(part));
  }

  function isActivePlanningStatus(status){
    const normalized = normalizeKey(status);
    if(!normalized) return false;
    if(INACTIVE_STATUS_BLOCKLIST.some(token => normalized.includes(token))) return false;
    return ACTIVE_STATUS_ALLOWLIST.some(token => normalized.includes(token));
  }

  function getActivePPMWorkOrders(workOrders){
    return (Array.isArray(workOrders) ? workOrders : []).filter(order => {
      if(!isPPMCategory(order.category)) return false;
      if(!isActivePlanningStatus(order.status)) return false;
      if(!order.assetName) return false;
      if(!order.woNumber) return false;
      return true;
    });
  }

  function computeWeekBucket(date, selectedPeriodStart){
    const workDate = startOfDay(parseDateValue(date));
    const periodStart = startOfDay(parseDateValue(selectedPeriodStart));

    if(!workDate || !periodStart) return 'outOfRange';

    const deltaMs = workDate.getTime() - periodStart.getTime();
    if(deltaMs < 0) return 'outOfRange';

    const bucketConfig = [
      { key: 'week1', from: 0, to: WEEK_MS },
      { key: 'week2', from: WEEK_MS, to: 2 * WEEK_MS },
      { key: 'week3', from: 2 * WEEK_MS, to: 3 * WEEK_MS }
    ];

    for(const bucket of bucketConfig){
      if(deltaMs >= bucket.from && deltaMs < bucket.to) return bucket.key;
    }
    return 'outOfRange';
  }

  function createEmptyWeeks(){
    return { week1: [], week2: [], week3: [] };
  }

  function groupPPMWorkOrdersByAssetAndWeek(workOrders, selectedSite, selectedPeriodStart){
    const filteredSiteKey = selectedSite && selectedSite !== 'all' ? selectedSite : 'all';
    const bySite = new Map();

    for(const order of (Array.isArray(workOrders) ? workOrders : [])){
      if(filteredSiteKey !== 'all' && order.siteKey !== filteredSiteKey) continue;

      const bucket = computeWeekBucket(order.dueDate || order.dueDateRaw, selectedPeriodStart);
      if(bucket === 'outOfRange') continue;

      const siteKey = order.siteKey || 'other';
      if(!bySite.has(siteKey)) bySite.set(siteKey, new Map());
      const assets = bySite.get(siteKey);

      const assetName = order.assetName;
      if(!assets.has(assetName)){
        assets.set(assetName, { assetName, siteKey, site: order.site || SITE_KEY_LABELS[siteKey] || 'Unknown', weeks: createEmptyWeeks() });
      }

      assets.get(assetName).weeks[bucket].push(order);
    }

    const sites = [...bySite.keys()].sort((a, b) => {
      const rank = ['ek', 'mm', 'keith', 'by', 'other'];
      const ia = rank.indexOf(a);
      const ib = rank.indexOf(b);
      if(ia !== ib) return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
      return a.localeCompare(b, undefined, { sensitivity: 'base' });
    });

    const assetsBySite = {};
    const rows = [];
    let totalAssets = 0;
    let totalWorkOrders = 0;

    for(const siteKey of sites){
      const sortedAssets = [...bySite.get(siteKey).values()]
        .sort((a,b) => a.assetName.localeCompare(b.assetName, undefined, { sensitivity: 'base' }));

      assetsBySite[siteKey] = sortedAssets;
      totalAssets += sortedAssets.length;

      sortedAssets.forEach(asset => {
        totalWorkOrders += asset.weeks.week1.length + asset.weeks.week2.length + asset.weeks.week3.length;
        rows.push(asset);
      });
    }

    return {
      sites,
      assetsBySite,
      rows,
      summary: {
        totalAssets,
        totalWorkOrders
      }
    };
  }

  async function loadMaintainXRawData({ bust = false } = {}){
    const weeks = await loadWeeksIndex({ bust });
    if(!Array.isArray(weeks) || !weeks.length){
      throw new Error('No week entries were found in the MaintainX index sheet.');
    }

    const selectedWeek = chooseInitialWeek(weeks);
    if(!selectedWeek) throw new Error('Could not resolve a selected week from the index sheet.');

    let sourceUrl = selectedWeek.url || null;
    if(!sourceUrl && selectedWeek.gid != null){
      sourceUrl = csvUrlForGid(selectedWeek.gid);
    }
    if(!sourceUrl){
      throw new Error('Selected week is missing both URL and gid source references.');
    }

    const res = await fetch(sourceUrl, { cache: 'no-store' });
    if(!res.ok){
      throw new Error(`MaintainX week load failed (${res.status} ${res.statusText})`);
    }

    const text = await res.text();
    if(!text || !text.trim()){
      throw new Error('MaintainX week CSV is empty.');
    }

    const parsed = parseDelimited(text);
    const dedupedRows = dedupeRows(parsed.rows || []);

    return {
      selectedWeek,
      headers: parsed.headers || [],
      rawRows: dedupedRows
    };
  }

  async function buildPPMPlannerModelFromMaintainX(options = {}){
    const source = await loadMaintainXRawData(options);
    const normalized = getNormalizedWorkOrders(source.rawRows);
    const activePPM = getActivePPMWorkOrders(normalized);
    const selectedPeriodStart = options.selectedPeriodStart || startOfWeek(new Date());
    const grouped = groupPPMWorkOrdersByAssetAndWeek(activePPM, options.siteKey || 'all', selectedPeriodStart);

    return {
      ...grouped,
      selectedWeek: source.selectedWeek,
      summary: {
        ...grouped.summary,
        totalRowsInput: source.rawRows.length,
        totalNormalizedRows: normalized.length,
        totalActivePPMRows: activePPM.length
      }
    };
  }

  const api = {
    SHEET_INDEX_CSV_URL,
    SITE_KEY_LABELS,
    loadMaintainXRawData,
    getNormalizedWorkOrders,
    getActivePPMWorkOrders,
    groupPPMWorkOrdersByAssetAndWeek,
    buildPPMPlannerModelFromMaintainX,
    buildPPMPlannerModel(rawRows, { siteKey = 'all', selectedPeriodStart } = {}){
      const normalized = getNormalizedWorkOrders(rawRows);
      const activePPM = getActivePPMWorkOrders(normalized);
      const grouped = groupPPMWorkOrdersByAssetAndWeek(activePPM, siteKey, selectedPeriodStart || startOfWeek(new Date()));
      return {
        ...grouped,
        summary: {
          ...grouped.summary,
          totalRowsInput: Array.isArray(rawRows) ? rawRows.length : 0,
          totalNormalizedRows: normalized.length,
          totalActivePPMRows: activePPM.length
        }
      };
    }
  };

  if(typeof module !== 'undefined' && module.exports){
    module.exports = api;
  }

  global.PPMWeeklyAssetTransform = api;
})(typeof window !== 'undefined' ? window : globalThis);
