const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { put, get } = require('@vercel/blob');

// 로컬 개발 환경에서 .env.local 자동 로드
function ensureEnvLoaded() {
  if (process.env.BLOB_READ_WRITE_TOKEN) return;
  try {
    const envPath = path.join(process.cwd(), '.env.local');
    if (fs.existsSync(envPath)) {
      const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) continue;
        const eqIdx = trimmed.indexOf('=');
        if (eqIdx > 0) {
          const key = trimmed.slice(0, eqIdx).trim();
          let val = trimmed.slice(eqIdx + 1).trim();
          if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
            val = val.slice(1, -1);
          }
          if (!process.env[key]) {
            process.env[key] = val;
          }
        }
      }
    }
  } catch (e) {
    // ignore
  }
}

const COMMON_COST_CATEGORIES = [
  '종묘비',
  '무기질비료비',
  '유기질비료비',
  '농약비',
  '수도광열비',
  '기타재료비',
  '고용노동비',
  '기타'
];

const DEFAULT_CHANNELS = ['농라', '스마트스토어', '카톡', '고흥몰', '남도장터', '로컬푸드', '공판장', '기타'];

const DEFAULT_TARGET_MONTHS = [
  { idx: 1, ym: '2026-11', label: '26년 11월', start: '2026-11-01', end: '2026-11-30', days: 30, target: 0 },
  { idx: 2, ym: '2026-12', label: '26년 12월', start: '2026-12-01', end: '2026-12-31', days: 31, target: 0 },
  { idx: 3, ym: '2027-01', label: '27년 1월',  start: '2027-01-01', end: '2027-01-31', days: 31, target: 0 },
  { idx: 4, ym: '2027-02', label: '27년 2월',  start: '2027-02-01', end: '2027-02-28', days: 28, target: 0 },
  { idx: 5, ym: '2027-03', label: '27년 3월',  start: '2027-03-01', end: '2027-03-31', days: 31, target: 0 },
  { idx: 6, ym: '2027-04', label: '27년 4월',  start: '2027-04-30', end: '2027-04-30', days: 30, target: 0 },
  { idx: 7, ym: '2027-05', label: '27년 5월',  start: '2027-05-01', end: '2027-05-31', days: 31, target: 0 },
  { idx: 8, ym: '2027-06', label: '27년 6월',  start: '2027-06-01', end: '2027-06-30', days: 30, target: 0 },
  { idx: 9, ym: '2027-07', label: '27년 7월',  start: '2027-07-01', end: '2027-07-31', days: 31, target: 0 },
  { idx: 10, ym: '2027-08', label: '27년 8월', start: '2027-08-01', end: '2027-08-31', days: 31, target: 0 },
  { idx: 11, ym: '2027-09', label: '27년 9월', start: '2027-09-01', end: '2027-09-30', days: 30, target: 0 },
  { idx: 12, ym: '2027-10', label: '27년 10월', start: '2027-10-01', end: '2027-10-31', days: 31, target: 0 }
];
// Fix 2027-04 start date
DEFAULT_TARGET_MONTHS[5].start = '2027-04-01';

function sanitizeState(raw) {
  if (!raw || typeof raw !== 'object') return null;

  const channels = Array.isArray(raw.channels) && raw.channels.length > 0
    ? Array.from(new Set([...DEFAULT_CHANNELS, ...raw.channels.map(c => String(c || '').trim()).filter(Boolean)]))
    : [...DEFAULT_CHANNELS];

  const targetMonths = DEFAULT_TARGET_MONTHS.map(def => {
    const found = Array.isArray(raw.targetMonths) ? raw.targetMonths.find(x => x && x.ym === def.ym) : null;
    return {
      ...def,
      target: found ? Math.max(0, Number(found.target) || 0) : 0
    };
  });

  const channelRates = Array.isArray(raw.channelRates) && raw.channelRates.length > 0
    ? raw.channelRates.map((r, i) => ({
        id: r.id || (i + 1),
        channel: String(r.channel || DEFAULT_CHANNELS[i] || '기타'),
        start: String(r.start || '2026-11-01'),
        end: r.end ? String(r.end) : '',
        feeRate: Math.max(0, Number(r.feeRate) || 0),
        shipUnder10: Math.max(0, Number(r.shipUnder10) || 0),
        shipOver10: Math.max(0, Number(r.shipOver10) || 0),
        boxCost: Math.max(0, Number(r.boxCost) || 0),
        vinylCost: Math.max(0, Number(r.vinylCost) || 0),
        leafletCost: Math.max(0, Number(r.leafletCost) || 0),
        note: String(r.note || '')
      }))
    : DEFAULT_CHANNELS.map((ch, i) => ({
        id: i + 1,
        channel: ch,
        start: '2026-11-01',
        end: '',
        feeRate: 0,
        shipUnder10: 0,
        shipOver10: 0,
        boxCost: 0,
        vinylCost: 0,
        leafletCost: 0,
        note: ''
      }));

  const dailySales = Array.isArray(raw.dailySales)
    ? raw.dailySales.map((s, i) => ({
        id: s.id || (Date.now() + i),
        date: String(s.date || '2026-11-01'),
        channel: String(s.channel || '농라'),
        ordersUnder10: Math.max(0, Math.round(Number(s.ordersUnder10) || 0)),
        ordersOver10: Math.max(0, Math.round(Number(s.ordersOver10) || 0)),
        sales: Math.max(0, Number(s.sales) || 0),
        extraCost: Math.max(0, Number(s.extraCost) || 0),
        note: String(s.note || '')
      }))
    : [];

  const commonCosts = Array.isArray(raw.commonCosts)
    ? raw.commonCosts.map((c, i) => {
        const isValidCat = COMMON_COST_CATEGORIES.includes(c.category);
        return {
          id: c.id || (Date.now() + 10000 + i),
          date: String(c.date || '2026-11-01'),
          category: isValidCat ? c.category : '기타',
          detail: c.detail !== undefined ? String(c.detail) : (!isValidCat && c.category ? String(c.category) : ''),
          amount: Math.max(0, Number(c.amount) || 0),
          note: String(c.note || '')
        };
      })
    : [];

  return {
    channels,
    commonCostCategories: [...COMMON_COST_CATEGORIES],
    targetMonths,
    channelRates,
    dailySales,
    commonCosts
  };
}

function isRateConfigured(r) {
  if (!r) return false;
  return (
    (Number(r.feeRate) || 0) > 0 ||
    (Number(r.shipUnder10) || 0) > 0 ||
    (Number(r.shipOver10) || 0) > 0 ||
    (Number(r.boxCost) || 0) > 0 ||
    (Number(r.vinylCost) || 0) > 0 ||
    (Number(r.leafletCost) || 0) > 0 ||
    Boolean((r.end || '').trim()) ||
    Boolean((r.note || '').trim())
  );
}

function mergeStates(cloudRaw, localRaw) {
  const cloud = sanitizeState(cloudRaw);
  const local = sanitizeState(localRaw);
  if (!cloud && !local) return sanitizeState({});
  if (!cloud) return local;
  if (!local) return cloud;

  const channels = Array.from(new Set([...cloud.channels, ...local.channels]));

  const targetMonths = DEFAULT_TARGET_MONTHS.map(def => {
    const cTm = cloud.targetMonths.find(x => x.ym === def.ym);
    const lTm = local.targetMonths.find(x => x.ym === def.ym);
    const cVal = cTm ? Number(cTm.target) || 0 : 0;
    const lVal = lTm ? Number(lTm.target) || 0 : 0;
    return {
      ...def,
      target: cVal > 0 ? cVal : lVal
    };
  });

  // channelRates 병합: 설정된 단가가 있는 항목 우선 보존 + 신규 추가 행 합집합
  const rateMap = new Map();
  cloud.channelRates.forEach(r => {
    rateMap.set(`${r.channel}|${r.start}|${r.id}`, { ...r });
  });
  local.channelRates.forEach(lRate => {
    // 동일 id 또는 동일 채널+시작일의 기본 행이 있는지 확인
    let matchedKey = null;
    for (const [k, cRate] of rateMap.entries()) {
      if (cRate.id === lRate.id || (cRate.channel === lRate.channel && cRate.start === lRate.start && (cRate.end || '') === (lRate.end || ''))) {
        matchedKey = k;
        break;
      }
    }
    if (matchedKey) {
      const existing = rateMap.get(matchedKey);
      if (!isRateConfigured(existing) && isRateConfigured(lRate)) {
        rateMap.set(matchedKey, { ...lRate });
      }
    } else if (isRateConfigured(lRate)) {
      rateMap.set(`${lRate.channel}|${lRate.start}|${lRate.id}`, { ...lRate });
    }
  });
  const channelRates = Array.from(rateMap.values());

  // dailySales 병합 (id 또는 date+channel+sales+ordersUnder10+ordersOver10 기준 중복 방지 합집합)
  const salesMap = new Map();
  const makeSaleSig = s => `${s.id}::${s.date}::${s.channel}::${s.ordersUnder10}::${s.ordersOver10}::${s.sales}`;
  cloud.dailySales.forEach(s => salesMap.set(String(s.id), { ...s, _sig: makeSaleSig(s) }));
  local.dailySales.forEach(s => {
    const idKey = String(s.id);
    if (!salesMap.has(idKey)) {
      salesMap.set(idKey, { ...s });
    }
  });
  const dailySales = Array.from(salesMap.values()).map(({ _sig, ...rest }) => rest);

  // commonCosts 병합 (id 기준 합집합)
  const commonMap = new Map();
  cloud.commonCosts.forEach(c => commonMap.set(String(c.id), { ...c }));
  local.commonCosts.forEach(c => {
    const idKey = String(c.id);
    if (!commonMap.has(idKey)) {
      commonMap.set(idKey, { ...c });
    }
  });
  const commonCosts = Array.from(commonMap.values());

  return {
    channels,
    commonCostCategories: [...COMMON_COST_CATEGORIES],
    targetMonths,
    channelRates,
    dailySales,
    commonCosts
  };
}

function hashUserKey(userId) {
  return crypto.createHash('sha256').update(`spm_user_v1:${userId}`).digest('hex');
}

function hashPassword(userId, password) {
  return crypto.createHash('sha256').update(`spm_pw_v1:${userId}:${password}`).digest('hex');
}

function safeCompareHex(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  try {
    return crypto.timingSafeEqual(Buffer.from(a, 'utf8'), Buffer.from(b, 'utf8'));
  } catch (e) {
    return false;
  }
}

function extractCredentials(req) {
  const authHeader = req.headers && (req.headers.authorization || req.headers.Authorization);
  if (authHeader && authHeader.startsWith('Bearer ')) {
    const token = authHeader.slice(7).trim();
    try {
      const decoded = JSON.parse(Buffer.from(token, 'base64').toString('utf8'));
      return {
        userId: String(decoded.userId || '').trim().toLowerCase(),
        password: String(decoded.password || '')
      };
    } catch (e) {
      // ignore
    }
  }
  if (req.body && req.body.auth) {
    return {
      userId: String(req.body.auth.userId || '').trim().toLowerCase(),
      password: String(req.body.auth.password || '')
    };
  }
  return { userId: '', password: '' };
}

// 비공개(Private) 저장소 읽기/쓰기 헬퍼
async function readAccountRecord(userHash) {
  ensureEnvLoaded();
  const blobPath = `users/${userHash}/account_state_v1.json`;

  if (process.env.SPM_LOCAL_FALLBACK === '1' || !process.env.BLOB_READ_WRITE_TOKEN) {
    const localFile = path.join(process.cwd(), '.local_cloud_store', `${userHash}.json`);
    if (!fs.existsSync(localFile)) return null;
    return JSON.parse(fs.readFileSync(localFile, 'utf8'));
  }

  try {
    const result = await get(blobPath, {
      access: 'private',
      useCache: false
    });
    if (!result || !result.stream) return null;
    const text = await new Response(result.stream).text();
    if (!text) return null;
    return JSON.parse(text);
  } catch (err) {
    if (err && (err.name === 'BlobNotFoundError' || String(err.message || '').includes('not found') || err.status === 404)) {
      return null;
    }
    throw err;
  }
}

async function writeAccountRecord(userHash, record) {
  ensureEnvLoaded();
  const blobPath = `users/${userHash}/account_state_v1.json`;
  const payload = JSON.stringify(record);

  if (process.env.SPM_LOCAL_FALLBACK === '1' || !process.env.BLOB_READ_WRITE_TOKEN) {
    const dir = path.join(process.cwd(), '.local_cloud_store');
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, `${userHash}.json`), payload, 'utf8');
    return;
  }

  await put(blobPath, payload, {
    access: 'private',
    addRandomSuffix: false,
    allowOverwrite: true,
    contentType: 'application/json',
    cacheControlMaxAge: 0
  });
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.setHeader('Content-Type', 'application/json; charset=utf-8');

  if (req.method === 'OPTIONS') {
    res.statusCode = 200;
    return res.end(JSON.stringify({ ok: true }));
  }

  try {
    const { userId, password } = extractCredentials(req);
    if (!userId || userId.length < 2 || !password || password.length < 4) {
      res.statusCode = 401;
      return res.end(JSON.stringify({
        ok: false,
        code: 'AUTH_REQUIRED',
        message: '계정 아이디(2자 이상)와 비밀번호(4자 이상)로 로그인해 주세요.'
      }));
    }

    const userHash = hashUserKey(userId);
    const pwHash = hashPassword(userId, password);

    const existing = await readAccountRecord(userHash);

    if (existing && !safeCompareHex(existing.passwordHash, pwHash)) {
      res.statusCode = 401;
      return res.end(JSON.stringify({
        ok: false,
        code: 'INVALID_CREDENTIALS',
        message: '비밀번호가 일치하지 않습니다. PC와 모바일에서 동일한 아이디·비밀번호를 입력해 주세요.'
      }));
    }

    if (req.method === 'GET') {
      if (!existing) {
        const newRecord = {
          userId,
          passwordHash: pwHash,
          revision: 0,
          updatedAt: Date.now(),
          state: null
        };
        await writeAccountRecord(userHash, newRecord);
        res.statusCode = 200;
        return res.end(JSON.stringify({
          ok: true,
          isNewAccount: true,
          userId,
          revision: 0,
          updatedAt: newRecord.updatedAt,
          state: null
        }));
      }

      res.statusCode = 200;
      return res.end(JSON.stringify({
        ok: true,
        isNewAccount: false,
        userId: existing.userId || userId,
        revision: Number(existing.revision) || 0,
        updatedAt: existing.updatedAt || Date.now(),
        state: existing.state ? sanitizeState(existing.state) : null
      }));
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
      const action = body.action || 'save';
      const incomingState = body.state ? sanitizeState(body.state) : null;

      let finalState = incomingState;
      if (action === 'merge') {
        finalState = mergeStates(existing ? existing.state : null, incomingState);
      } else if (!finalState && existing && existing.state) {
        finalState = sanitizeState(existing.state);
      } else if (!finalState) {
        finalState = sanitizeState({});
      }

      const nextRevision = (existing ? (Number(existing.revision) || 0) : 0) + 1;
      const nowTs = Date.now();
      const nextRecord = {
        userId,
        passwordHash: pwHash,
        revision: nextRevision,
        updatedAt: nowTs,
        state: finalState
      };

      await writeAccountRecord(userHash, nextRecord);

      res.statusCode = 200;
      return res.end(JSON.stringify({
        ok: true,
        userId,
        revision: nextRevision,
        updatedAt: nowTs,
        state: finalState
      }));
    }

    res.statusCode = 405;
    return res.end(JSON.stringify({ ok: false, message: '허용되지 않는 요청 방식입니다.' }));
  } catch (err) {
    console.error('[api/state error]', err);
    res.statusCode = 500;
    return res.end(JSON.stringify({
      ok: false,
      code: 'SERVER_ERROR',
      message: '클라우드 저장소와 통신하는 중 일시적인 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.'
    }));
  }
};
