/**
 * xlife-wallet.js
 * BFF API — 幣錢包聚合服務
 *
 * Endpoints:
 *   GET  /wallet/summary        — 錢包總覽（聚合四個來源）
 *   GET  /wallet/transactions   — 交易明細（游標分頁）
 *   POST /wallet/redeem         — 幣兌換 / 扣點
 */

const express = require('express');
const { requireAuth } = require('../middleware/auth');

const router = express.Router();

// ---------------------------------------------------------------------------
// 模擬資料（TODO: 替換為呼叫子系統 API — 遠傳點數系統 / friDay Pay / 影音幣服務）
// ---------------------------------------------------------------------------

const NOW = new Date();

/** 計算距今 N 天後的 ISO 字串 */
function daysFromNow(n) {
  const d = new Date(NOW);
  d.setDate(d.getDate() + n);
  return d.toISOString();
}

/** 判斷是否在 30 天內到期 */
function isExpiringSoon(isoDate) {
  if (!isoDate) return false;
  const diff = new Date(isoDate) - NOW;
  return diff > 0 && diff <= 30 * 24 * 60 * 60 * 1000;
}

// 四個幣種來源定義
const SOURCES = {
  friday_points: {
    source: 'friday_points',
    label: 'friDay 購物點數',
    balance: 3200,
    expires_at: daysFromNow(25),    // 25 天後到期 → is_expiring_soon = true
  },
  friday_coin: {
    source: 'friday_coin',
    label: 'friDay 影音幣',
    balance: 150,
    expires_at: daysFromNow(90),
  },
  far_pay: {
    source: 'far_pay',
    label: '遠傳 Pay 餘額',
    balance: 1800,
    expires_at: null,               // 無期限
  },
  telecom_reward: {
    source: 'telecom_reward',
    label: '電信回饋金',
    balance: 480,
    expires_at: daysFromNow(10),    // 10 天後到期 → is_expiring_soon = true
  },
};

// 模擬交易明細（10 筆）
// TODO: 替換為呼叫遠傳點數帳務系統 / friDay Pay 交易記錄 API
const MOCK_TRANSACTIONS = [
  { id: 'txn_001', source: 'friday_points',  type: 'credit', amount: 500,  description: '消費回饋', occurred_at: daysFromNow(-3),  ref_order_id: 'ORD-20260429-001' },
  { id: 'txn_002', source: 'far_pay',        type: 'debit',  amount: 200,  description: '門市消費', occurred_at: daysFromNow(-5),  ref_order_id: 'ORD-20260427-009' },
  { id: 'txn_003', source: 'friday_coin',    type: 'credit', amount: 30,   description: '影音訂閱回饋', occurred_at: daysFromNow(-6),  ref_order_id: null },
  { id: 'txn_004', source: 'friday_points',  type: 'debit',  amount: 100,  description: '點數折抵', occurred_at: daysFromNow(-8),  ref_order_id: 'ORD-20260424-003' },
  { id: 'txn_005', source: 'telecom_reward', type: 'credit', amount: 480,  description: '月租帳單回饋', occurred_at: daysFromNow(-10), ref_order_id: null },
  { id: 'txn_006', source: 'far_pay',        type: 'credit', amount: 1000, description: '儲值', occurred_at: daysFromNow(-12), ref_order_id: null },
  { id: 'txn_007', source: 'friday_points',  type: 'expired', amount: 200, description: '點數到期', occurred_at: daysFromNow(-15), ref_order_id: null },
  { id: 'txn_008', source: 'friday_coin',    type: 'credit', amount: 50,   description: '活動贈幣', occurred_at: daysFromNow(-18), ref_order_id: null },
  { id: 'txn_009', source: 'far_pay',        type: 'debit',  amount: 350,  description: '繳費代扣', occurred_at: daysFromNow(-20), ref_order_id: 'BILL-20260412' },
  { id: 'txn_010', source: 'friday_points',  type: 'credit', amount: 2800, description: '首購大禮包', occurred_at: daysFromNow(-30), ref_order_id: 'ORD-20260402-001' },
];

// 兌換規則白名單：source → 可轉入的 target_service 集合
// TODO: 替換為呼叫幣兌換規則引擎 API
const REDEEM_RULES = {
  friday_points:  ['friday_shopping', 'far_pay', 'friday_video'],
  friday_coin:    ['friday_video'],             // 影音幣只能在影音服務使用，不可轉 Pay
  far_pay:        ['friday_shopping', 'bill_payment', 'store_payment'],
  telecom_reward: ['friday_shopping', 'far_pay', 'bill_payment'],
};

// ---------------------------------------------------------------------------
// GET /wallet/summary
// ---------------------------------------------------------------------------
router.get('/wallet/summary', requireAuth, (req, res) => {
  // TODO: 呼叫各子系統查詢真實餘額（遠傳點數 API、friDay Pay API、影音幣 API、電信帳務 API）
  const sources = Object.values(SOURCES).map(s => ({
    source:          s.source,
    label:           s.label,
    balance:         s.balance,
    expires_at:      s.expires_at,
    is_expiring_soon: isExpiringSoon(s.expires_at),
  }));

  const total_balance = sources.reduce((sum, s) => sum + s.balance, 0);

  res.json({
    total_balance,
    sources,
    last_synced_at: NOW.toISOString(),
  });
});

// ---------------------------------------------------------------------------
// GET /wallet/transactions
// Query params:
//   source  — 過濾來源（選填）
//   limit   — 每頁筆數（預設 20，最大 100）
//   cursor  — 游標（上一頁最後一筆 id，選填）
// ---------------------------------------------------------------------------
router.get('/wallet/transactions', requireAuth, (req, res) => {
  // TODO: 呼叫帳務子系統 API，傳入 user_id + source + limit + cursor

  const { source, cursor } = req.query;
  const limit = Math.min(parseInt(req.query.limit, 10) || 20, 100);

  let list = [...MOCK_TRANSACTIONS];

  // 來源過濾
  if (source) {
    if (!SOURCES[source]) {
      return res.status(400).json({ error: `Unknown source: ${source}` });
    }
    list = list.filter(t => t.source === source);
  }

  // 游標分頁：找到 cursor 所在位置後取後續資料
  if (cursor) {
    const idx = list.findIndex(t => t.id === cursor);
    if (idx === -1) {
      return res.status(400).json({ error: 'Invalid cursor' });
    }
    list = list.slice(idx + 1);
  }

  const page = list.slice(0, limit);
  const next_cursor = page.length === limit ? page[page.length - 1].id : null;

  res.json({
    transactions: page,
    pagination: {
      limit,
      next_cursor,
      has_more: next_cursor !== null,
    },
  });
});

// ---------------------------------------------------------------------------
// POST /wallet/redeem
// Headers: Idempotency-Key (必要)
// Body:    { source, amount, target_service, ref_order_id }
// ---------------------------------------------------------------------------
router.post('/wallet/redeem', requireAuth, (req, res) => {
  // TODO: 呼叫幣兌換服務 API（需冪等 key 傳遞至下游）
  const idempotencyKey = req.headers['idempotency-key'];
  if (!idempotencyKey) {
    return res.status(400).json({ error: 'Missing required header: Idempotency-Key' });
  }

  const { source, amount, target_service, ref_order_id } = req.body;

  // 必填欄位驗證
  if (!source || amount === undefined || !target_service) {
    return res.status(400).json({ error: 'Missing required fields: source, amount, target_service' });
  }

  // 金額必須為正整數
  const parsedAmount = parseInt(amount, 10);
  if (!Number.isInteger(parsedAmount) || parsedAmount <= 0) {
    return res.status(400).json({ error: 'amount must be a positive integer' });
  }

  // 來源存在性驗證
  const sourceData = SOURCES[source];
  if (!sourceData) {
    return res.status(400).json({ error: `Unknown source: ${source}` });
  }

  // 兌換規則驗證
  const allowedTargets = REDEEM_RULES[source] || [];
  if (!allowedTargets.includes(target_service)) {
    return res.status(422).json({
      error: `Source "${source}" cannot be redeemed to target "${target_service}"`,
      allowed_targets: allowedTargets,
    });
  }

  // 餘額不足驗證
  if (parsedAmount > sourceData.balance) {
    return res.status(422).json({
      error: 'Insufficient balance',
      current_balance: sourceData.balance,
    });
  }

  // TODO: 以 idempotencyKey 查詢快取，避免重複扣款
  // TODO: 呼叫扣款子系統 API，取得真實 redeem_id 與 new_balance

  // 模擬扣款（僅回傳計算結果，不修改 SOURCES 常數）
  const new_balance = sourceData.balance - parsedAmount;
  const redeem_id = `rdm_${crypto.randomUUID().replace(/-/g, '').slice(0, 12)}`;

  res.status(201).json({
    redeem_id,
    idempotency_key: idempotencyKey,
    source,
    target_service,
    ref_order_id: ref_order_id || null,
    deducted: parsedAmount,
    new_balance,
    status: 'success',
    redeemed_at: new Date().toISOString(),
  });
});

module.exports = router;
