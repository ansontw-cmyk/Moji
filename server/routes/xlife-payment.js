/**
 * xlife-payment.js
 * BFF API — 支付回饋聚合服務
 *
 * Endpoints:
 *   GET /payment/reward-preview/:transactionId
 *     — 查詢指定交易的回饋資訊及會員等級進度
 */

'use strict';

const express = require('express');
const { requireAuth } = require('../middleware/auth');

const router = express.Router();

// ---------------------------------------------------------------------------
// 模擬資料與輔助函式
// ---------------------------------------------------------------------------

/**
 * 計算距今 N 天後（或前，負數）的 ISO 字串。
 * 使用固定基準時間，避免每次請求產生不同的「昨天」。
 */
const BASE_NOW = new Date();

function daysFromNow(n) {
  const d = new Date(BASE_NOW);
  d.setDate(d.getDate() + n);
  return d.toISOString();
}

// 會員等級定義（門檻為累計消費金額 TWD）
// TODO: 替換為呼叫會員等級服務 API 取得真實設定
const LEVEL_TIERS = [
  { level: 'bronze',   label: '銅卡',   threshold: 0,      next_threshold: 5000  },
  { level: 'silver',   label: '銀卡',   threshold: 5000,   next_threshold: 20000 },
  { level: 'gold',     label: '金卡',   threshold: 20000,  next_threshold: 50000 },
  { level: 'platinum', label: '白金卡', threshold: 50000,  next_threshold: null  },
];

/**
 * 根據累計消費金額決定當前等級及下一級資訊。
 * @param {number} totalSpend - 累計消費金額（TWD）
 */
function resolveLevel(totalSpend) {
  let current = LEVEL_TIERS[0];
  for (const tier of LEVEL_TIERS) {
    if (totalSpend >= tier.threshold) current = tier;
  }
  const next = LEVEL_TIERS.find(t => t.threshold > current.threshold) || null;
  return { current, next };
}

/**
 * 根據交易金額與會員等級計算回饋點數。
 * TODO: 替換為呼叫回饋計算引擎 API（支援多種回饋規則、活動加碼）
 *
 * 回饋率：bronze 1%, silver 1.5%, gold 2%, platinum 3%
 */
const REWARD_RATE = {
  bronze:   0.01,
  silver:   0.015,
  gold:     0.02,
  platinum: 0.03,
};

function calcPoints(paymentAmount, level) {
  const rate = REWARD_RATE[level] ?? 0.01;
  return Math.floor(paymentAmount * rate);
}

// 模擬交易資料庫（以 transactionId 為索引）
// TODO: 替換為呼叫交易查詢 API（遠傳 Pay / friDay 購物 交易記錄服務）
const MOCK_TRANSACTIONS = {
  'TXN-20260501-001': { payment_amount: 1280, user_total_spend: 18500, source: 'friday_points', is_settled: true,  days_ago: -1 },
  'TXN-20260430-002': { payment_amount: 450,  user_total_spend: 17220, source: 'far_pay',        is_settled: true,  days_ago: -2 },
  'TXN-20260429-003': { payment_amount: 3200, user_total_spend: 16770, source: 'friday_points',  is_settled: false, days_ago: -3 },
  'TXN-20260428-004': { payment_amount: 99,   user_total_spend: 13570, source: 'telecom_reward', is_settled: true,  days_ago: -4 },
  'TXN-20260425-005': { payment_amount: 5800, user_total_spend: 13471, source: 'far_pay',        is_settled: true,  days_ago: -7 },
};

// 推薦動作清單（依會員等級差異化）
// TODO: 替換為呼叫個人化推薦 API
const RECOMMENDATIONS_BY_LEVEL = {
  bronze: [
    { type: 'upgrade_hint',    title: '再消費即可升級銀卡',     description: '銀卡享 1.5% 回饋，提升消費加速升等',      cta: { label: '查看優惠', action: 'navigate', target: '/membership/benefits' } },
    { type: 'offer',           title: 'friDay 購物首購 95 折',  description: '使用遠傳 Pay 結帳再享 3% 回饋',            cta: { label: '立即購物', action: 'deep_link', target: 'friday://mall/home' } },
  ],
  silver: [
    { type: 'upgrade_hint',    title: '衝刺金卡！提升回饋率',   description: '金卡享 2% 回饋，距離升等還差一點點',       cta: { label: '查看進度', action: 'navigate', target: '/membership/progress' } },
    { type: 'offer',           title: '影音月租方案 9 折',      description: '訂閱 friDay 影音享雙倍回饋幣',             cta: { label: '立即訂閱', action: 'deep_link', target: 'friday://video/subscribe' } },
  ],
  gold: [
    { type: 'upgrade_hint',    title: '白金卡在望！',          description: '白金卡享 3% 最高回饋，加速衝刺',           cta: { label: '查看進度', action: 'navigate', target: '/membership/progress' } },
    { type: 'offer',           title: '門市獨家金卡優惠',      description: '持金卡消費享消費金額 5 倍點數',             cta: { label: '查看門市', action: 'navigate', target: '/store-locator' } },
  ],
  platinum: [
    { type: 'exclusive_offer', title: '白金專屬：機場接送免費', description: '每月乘車一次，遠傳尊榮接送服務',            cta: { label: '立即預約', action: 'navigate', target: '/platinum/airport-pickup' } },
    { type: 'exclusive_offer', title: '白金尊享：3% 最高回饋', description: '所有消費均享 3% 回饋，無上限累積',          cta: { label: '查看回饋', action: 'navigate', target: '/membership/benefits' } },
  ],
};

// ---------------------------------------------------------------------------
// GET /payment/reward-preview/:transactionId
// ---------------------------------------------------------------------------
router.get('/payment/reward-preview/:transactionId', requireAuth, (req, res) => {
  // TODO: 使用 req.user.sub 驗證交易是否屬於該用戶（呼叫交易查詢 API）
  const { transactionId } = req.params;

  // 查詢模擬交易資料
  // TODO: 替換為呼叫交易查詢 API
  const txn = MOCK_TRANSACTIONS[transactionId];
  if (!txn) {
    return res.status(404).json({
      error: 'Transaction not found',
      transaction_id: transactionId,
    });
  }

  // 等級計算
  const { current: currentTier, next: nextTier } = resolveLevel(txn.user_total_spend);

  // 回饋點數計算
  // TODO: 替換為呼叫回饋計算引擎 API（含活動加碼規則）
  const points_earned = calcPoints(txn.payment_amount, currentTier.level);

  // 升等進度
  const current_total   = txn.user_total_spend;
  const next_threshold  = nextTier ? nextTier.threshold : null;
  const gap             = nextTier ? Math.max(0, next_threshold - current_total) : 0;
  const percentage      = nextTier
    ? Math.min(100, Math.round(((current_total - currentTier.threshold) / (next_threshold - currentTier.threshold)) * 100))
    : 100;

  res.json({
    transaction_id: transactionId,
    payment_amount: txn.payment_amount,
    reward: {
      points_earned,
      source:      txn.source,
      credited_at: daysFromNow(txn.days_ago + 1), // 模擬次日入帳
      is_settled:  txn.is_settled,
    },
    progress: {
      current_level:   currentTier.level,
      current_label:   currentTier.label,
      next_level:      nextTier ? nextTier.level  : null,
      next_label:      nextTier ? nextTier.label  : null,
      current_total,
      next_threshold,
      gap,
      percentage,
    },
    recommendations: RECOMMENDATIONS_BY_LEVEL[currentTier.level] ?? [],
  });
});

module.exports = router;
