/**
 * xlife-home.js
 * BFF API — 個人化首頁聚合服務
 *
 * Endpoints:
 *   GET  /home/cards              — 個人化首頁卡片列表
 *   POST /user/preference-quiz    — 新用戶偏好問卷初始化
 */

const express = require('express');
const { requireAuth } = require('../middleware/auth');

const router = express.Router();

// ---------------------------------------------------------------------------
// 模擬資料與輔助函式
// ---------------------------------------------------------------------------

/** 計算距今 N 天後的 ISO 字串 */
function daysFromNow(n) {
  const d = new Date();
  d.setDate(d.getDate() + n);
  return d.toISOString();
}

// 卡片工廠：各卡片類型的模擬內容
// TODO: 替換為呼叫個人化推薦引擎 API（含 AB 測試分流邏輯）

function buildTelecomUsageCard(userId) {
  // TODO: 呼叫電信帳務 API 取得真實用量資料
  return {
    card_id:   'card_telecom_usage',
    type:      'telecom_usage',
    priority:  1,
    data: {
      plan_name:      '遠傳 5G 月租 999',
      data_used_gb:   18.4,
      data_total_gb:  30,
      data_pct:       61,
      days_remaining: 12,
      voice_used_min: 210,
      voice_total_min: 500,
    },
    cta: { label: '查看帳單', action: 'navigate', target: '/telecom/bill' },
  };
}

function buildOfferCards() {
  // TODO: 呼叫優惠推薦 API（依用戶消費標籤個人化）
  return [
    {
      card_id:  'card_offer_friday_mall',
      type:     'offer',
      priority: 2,
      data: {
        campaign_id:   'camp_20260501_mall',
        title:         'friDay 購物節 — 指定品牌 88 折',
        subtitle:      '使用遠傳 Pay 再享 5% 回饋',
        image_url:     'https://cdn.friday.com/campaigns/20260501_mall.jpg', // TODO: 替換為實際 CDN
        expires_at:    daysFromNow(7),
        badge:         'hot',
      },
      cta: { label: '立即搶購', action: 'deep_link', target: 'friday://mall/campaign/camp_20260501_mall' },
    },
    {
      card_id:  'card_offer_video',
      type:     'offer',
      priority: 3,
      data: {
        campaign_id:   'camp_20260501_video',
        title:         'friDay 影音 30 天免費體驗',
        subtitle:      '訂閱即送 50 影音幣',
        image_url:     'https://cdn.friday.com/campaigns/20260501_video.jpg', // TODO: 替換為實際 CDN
        expires_at:    daysFromNow(14),
        badge:         'new',
      },
      cta: { label: '免費體驗', action: 'deep_link', target: 'friday://video/trial' },
    },
  ];
}

function buildNearbyMerchantCard(lat, lng) {
  // TODO: 呼叫附近商家服務 API（傳入 lat/lng，回傳半徑 1km 內合作商家）
  return {
    card_id:  'card_nearby_merchant',
    type:     'nearby_merchant',
    priority: 4,
    data: {
      location:  { lat: parseFloat(lat), lng: parseFloat(lng) },
      merchants: [
        {
          merchant_id: 'mcht_001',
          name:        '全家便利商店（信義店）',
          distance_m:  120,
          cashback_pct: 3,
          logo_url:    'https://cdn.friday.com/merchants/family_mart.png', // TODO: 替換
        },
        {
          merchant_id: 'mcht_002',
          name:        '路易莎咖啡（松仁路店）',
          distance_m:  350,
          cashback_pct: 5,
          logo_url:    'https://cdn.friday.com/merchants/louisa.png', // TODO: 替換
        },
      ],
    },
    cta: { label: '查看更多', action: 'navigate', target: '/nearby-merchants' },
  };
}

function buildVideoRecommendationCard() {
  // TODO: 呼叫影音推薦引擎 API（協同過濾或內容推薦）
  return {
    card_id:  'card_video_recommendation',
    type:     'video_recommendation',
    priority: 5,
    data: {
      section_title: '為你推薦',
      items: [
        { content_id: 'vid_101', title: '你的孩子不是你的孩子', type: 'series', thumbnail_url: 'https://cdn.friday.com/thumbnails/vid_101.jpg' },
        { content_id: 'vid_205', title: '模仿犯', type: 'series', thumbnail_url: 'https://cdn.friday.com/thumbnails/vid_205.jpg' },
        { content_id: 'vid_312', title: 'KPOP 演唱會精選', type: 'live', thumbnail_url: 'https://cdn.friday.com/thumbnails/vid_312.jpg' },
      ],
    },
    cta: { label: '前往影音', action: 'deep_link', target: 'friday://video/home' },
  };
}

// AB 測試分組（模擬）
// TODO: 替換為呼叫 Feature Flag / AB 測試服務
function resolveAbVariant(userId) {
  const hash = userId.split('').reduce((acc, c) => acc + c.charCodeAt(0), 0);
  return hash % 2 === 0 ? 'variant_A' : 'variant_B';
}

// ---------------------------------------------------------------------------
// GET /home/cards
// Query params:
//   lat, lng       — 使用者當前位置（選填，提供時才回傳 nearby_merchant 卡）
//   context_from   — 進入點標記（選填，用於埋點 / 個人化調整）
// ---------------------------------------------------------------------------
router.get('/home/cards', requireAuth, (req, res) => {
  // TODO: 所有資料來源替換為真實子系統呼叫（電信、推薦、商家、影音）

  const userId = String(req.user.sub);
  const { lat, lng, context_from } = req.query;

  const hasLocation = lat !== undefined && lng !== undefined;

  // 冷啟動判斷（TODO: 查詢用戶偏好服務，確認是否已完成問卷）
  const is_cold_start = false;

  const cards = [
    buildTelecomUsageCard(userId),
    ...buildOfferCards(),
    ...(hasLocation ? [buildNearbyMerchantCard(lat, lng)] : []),
    buildVideoRecommendationCard(),
  ];

  // 依 priority 排序（越小越靠前）
  cards.sort((a, b) => a.priority - b.priority);

  res.json({
    user_id:              userId,
    is_cold_start,
    cards,
    context_from:         context_from || null,
    digest_generated_at:  new Date().toISOString(),
    ab_variant:           resolveAbVariant(userId),
  });
});

// ---------------------------------------------------------------------------
// POST /user/preference-quiz
// Body: { answers: [{ question_id, value }] }
//   — 至少需要 3 個 answers
// ---------------------------------------------------------------------------
router.post('/user/preference-quiz', requireAuth, (req, res) => {
  // TODO: 呼叫用戶偏好服務 API，持久化問卷結果並觸發個人化模型更新

  const { answers } = req.body;

  if (!Array.isArray(answers) || answers.length < 3) {
    return res.status(400).json({
      error: 'At least 3 answers are required',
      received: Array.isArray(answers) ? answers.length : 0,
    });
  }

  // 每個 answer 必須有 question_id 與 value
  const invalid = answers.filter(a => !a.question_id || a.value === undefined);
  if (invalid.length > 0) {
    return res.status(400).json({
      error: 'Each answer must include question_id and value',
      invalid_count: invalid.length,
    });
  }

  // 模擬：根據 question_id 映射到興趣標籤
  // TODO: 替換為呼叫個人化引擎的標籤推斷模型
  const INTEREST_MAP = {
    q_content_pref:  { video: 'entertainment', music: 'music', news: 'news_media', sport: 'sports' },
    q_shopping_freq: { daily: 'frequent_shopper', weekly: 'regular_shopper', monthly: 'casual_shopper' },
    q_telecom_use:   { data_heavy: 'data_user', call_heavy: 'voice_user', balance: 'balanced_user' },
    q_payment_pref:  { mobile_pay: 'digital_payment', credit_card: 'credit_user', cash: 'cash_user' },
    q_lifestyle:     { outdoor: 'outdoor_lifestyle', homebody: 'home_lifestyle', travel: 'traveler' },
  };

  const initial_interests = answers.reduce((interests, ans) => {
    const mapping = INTEREST_MAP[ans.question_id];
    if (mapping && mapping[ans.value]) {
      interests.push(mapping[ans.value]);
    }
    return interests;
  }, []);

  // 去重
  const unique_interests = [...new Set(initial_interests)];

  res.status(201).json({
    profile_initialized: true,
    user_id:             String(req.user.sub),
    initial_interests:   unique_interests,
    answered_count:      answers.length,
    next_step:           unique_interests.length > 0 ? 'view_home' : 'complete_more_questions',
    initialized_at:      new Date().toISOString(),
  });
});

module.exports = router;
