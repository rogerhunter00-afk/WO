(function(global){
  'use strict';

  const FIELD_ALIASES = Object.freeze({
    woNumber: ['ID', 'WO', 'WO Number', 'Work Order', 'Work Order Number', 'WO/Link', 'Reference', 'Number'],
    category: ['Categories', 'Category', 'Type', 'Classification'],
    asset: ['Asset', 'Asset Name', 'Asset ID', 'Asset Description', 'Equipment', 'Equipment Name', 'Name', 'Title'],
    site: ['Location', 'Site', 'Facility'],
    status: ['Status', 'State'],
    dueDate: ['Due date', 'Due Date', 'Planned Start Date', 'Planned Start', 'Planned Start Date (Local)'],
    title: ['Title', 'Description', 'Task', 'Summary', 'Asset', 'Asset Name']
  });

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

  function parseDateValue(value){
    if(value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
    const text = toStringSafe(value);
    if(!text) return null;
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function startOfDay(date){
    if(!(date instanceof Date) || Number.isNaN(date.getTime())) return null;
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  function extractField(raw, aliases){
    if(!raw || typeof raw !== 'object') return '';
    for(const key of aliases){
      if(Object.prototype.hasOwnProperty.call(raw, key)){
        const direct = toStringSafe(raw[key]);
        if(direct) return direct;
      }
    }

    const lowered = new Map();
    for(const [key, value] of Object.entries(raw)){
      const cleaned = normalizeKey(key);
      if(!cleaned || lowered.has(cleaned)) continue;
      lowered.set(cleaned, value);
    }

    for(const key of aliases){
      const fallback = lowered.get(normalizeKey(key));
      const cleaned = toStringSafe(fallback);
      if(cleaned) return cleaned;
    }

    return '';
  }

  function mapMaintainXRow(raw){
    const woNumber = extractField(raw, FIELD_ALIASES.woNumber);
    const category = extractField(raw, FIELD_ALIASES.category);
    const assetName = extractField(raw, FIELD_ALIASES.asset);
    const site = extractField(raw, FIELD_ALIASES.site);
    const status = extractField(raw, FIELD_ALIASES.status);
    const dueDateRaw = extractField(raw, FIELD_ALIASES.dueDate);
    const dueDate = parseDateValue(dueDateRaw);
    const title = extractField(raw, FIELD_ALIASES.title) || assetName;

    return {
      woNumber,
      category,
      assetName,
      site,
      status,
      dueDateRaw,
      dueDate,
      title,
      raw
    };
  }

  function isActivePlanningStatus(status){
    const normalized = normalizeKey(status);
    if(!normalized) return false;
    if(INACTIVE_STATUS_BLOCKLIST.some(token => normalized.includes(token))) return false;
    return ACTIVE_STATUS_ALLOWLIST.some(token => normalized.includes(token));
  }

  function isPPMCategory(categoryValue){
    if(Array.isArray(categoryValue)){
      return categoryValue.some(isPPMCategory);
    }
    const raw = toStringSafe(categoryValue);
    if(!raw) return false;
    const normalized = raw.replace(/\//g, ',');
    const parts = normalized.split(/[;,]+/).map(part => normalizeKey(part)).filter(Boolean);
    if(parts.length === 0){
      return /\bppm\b/i.test(raw);
    }
    return parts.some(part => /\bppm\b/i.test(part));
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

  function buildPPMPlannerModel(rawRows, { siteKey = 'all', selectedPeriodStart } = {}){
    const mappedRows = (Array.isArray(rawRows) ? rawRows : []).map(mapMaintainXRow);

    const malformed = {
      missingCoreFields: [],
      invalidDates: []
    };

    const siteOrder = [];
    const assetsBySite = {};

    for(const row of mappedRows){
      if(!isPPMCategory(row.category)) continue;
      if(!isActivePlanningStatus(row.status)) continue;

      if(!row.woNumber || !row.assetName || !row.site){
        malformed.missingCoreFields.push({
          woNumber: row.woNumber,
          assetName: row.assetName,
          site: row.site,
          status: row.status,
          category: row.category
        });
        continue;
      }

      const weekBucket = computeWeekBucket(row.dueDate || row.dueDateRaw, selectedPeriodStart);
      if(weekBucket === 'outOfRange'){
        if(!row.dueDate){
          malformed.invalidDates.push({ woNumber: row.woNumber, dueDateRaw: row.dueDateRaw });
        }
        continue;
      }

      const siteName = row.site;
      if(siteKey !== 'all' && normalizeKey(siteName) !== normalizeKey(siteKey)) continue;

      if(!assetsBySite[siteName]){
        assetsBySite[siteName] = [];
        siteOrder.push(siteName);
      }

      let assetEntry = assetsBySite[siteName].find(item => item.assetName === row.assetName);
      if(!assetEntry){
        assetEntry = {
          assetName: row.assetName,
          weeks: createEmptyWeeks()
        };
        assetsBySite[siteName].push(assetEntry);
      }

      assetEntry.weeks[weekBucket].push({
        woNumber: row.woNumber,
        title: row.title,
        status: row.status,
        dueDate: row.dueDate,
        dueDateRaw: row.dueDateRaw,
        category: row.category,
        site: row.site,
        raw: row.raw
      });
    }

    const rows = [];
    let totalWorkOrders = 0;
    let totalAssets = 0;

    for(const site of siteOrder){
      const siteAssets = assetsBySite[site]
        .filter(asset => (asset.weeks.week1.length + asset.weeks.week2.length + asset.weeks.week3.length) > 0)
        .sort((a, b) => a.assetName.localeCompare(b.assetName, undefined, { sensitivity: 'base' }));

      assetsBySite[site] = siteAssets;
      totalAssets += siteAssets.length;

      for(const asset of siteAssets){
        totalWorkOrders += asset.weeks.week1.length + asset.weeks.week2.length + asset.weeks.week3.length;
        rows.push(asset);
      }
    }

    if(malformed.missingCoreFields.length || malformed.invalidDates.length){
      console.warn('[PPM Planner] Malformed rows skipped:', {
        missingCoreFields: malformed.missingCoreFields,
        invalidDates: malformed.invalidDates
      });
    }

    return {
      sites: siteOrder.filter(site => assetsBySite[site] && assetsBySite[site].length > 0),
      assetsBySite,
      rows,
      summary: {
        totalRowsInput: Array.isArray(rawRows) ? rawRows.length : 0,
        totalPPMRowsMapped: mappedRows.length,
        totalAssets,
        totalWorkOrders,
        malformed
      }
    };
  }

  const api = {
    FIELD_ALIASES,
    mapMaintainXRow,
    isActivePlanningStatus,
    isPPMCategory,
    computeWeekBucket,
    buildPPMPlannerModel
  };

  if(typeof module !== 'undefined' && module.exports){
    module.exports = api;
  }

  global.PPMWeeklyAssetTransform = api;
})(typeof window !== 'undefined' ? window : globalThis);
