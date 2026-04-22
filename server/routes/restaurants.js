const express = require('express');
const db = require('../db');
const { requireManager } = require('../middleware/auth');

const router = express.Router();

router.get('/', (req, res) => {
  const isManager = req.user.role === 'manager';
  let restaurants;

  if (isManager) {
    restaurants = db.prepare(`
      SELECT r.*,
             a.user_id as assigned_to,
             u.name as assigned_to_name,
             rp.status as latest_status,
             rp.visited_at as latest_visit
      FROM restaurants r
      LEFT JOIN assignments a ON a.restaurant_id = r.id
      LEFT JOIN users u ON u.id = a.user_id
      LEFT JOIN reports rp ON rp.id = (
        SELECT id FROM reports WHERE restaurant_id = r.id ORDER BY created_at DESC LIMIT 1
      )
      ORDER BY r.name
    `).all();
  } else {
    restaurants = db.prepare(`
      SELECT r.*,
             rp.status as latest_status,
             rp.visited_at as latest_visit
      FROM restaurants r
      JOIN assignments a ON a.restaurant_id = r.id AND a.user_id = ?
      LEFT JOIN reports rp ON rp.id = (
        SELECT id FROM reports WHERE restaurant_id = r.id AND user_id = ? ORDER BY created_at DESC LIMIT 1
      )
      ORDER BY r.name
    `).all(req.user.sub, req.user.sub);
  }
  res.json(restaurants);
});

router.get('/:id', (req, res) => {
  const r = db.prepare(`
    SELECT r.*, a.user_id as assigned_to, u.name as assigned_to_name
    FROM restaurants r
    LEFT JOIN assignments a ON a.restaurant_id = r.id
    LEFT JOIN users u ON u.id = a.user_id
    WHERE r.id = ?
  `).get(req.params.id);
  if (!r) return res.status(404).json({ error: 'Not found' });
  res.json(r);
});

router.post('/', requireManager, (req, res) => {
  const { name, address, lat, lng, phone, category } = req.body;
  if (!name || !lat || !lng) return res.status(400).json({ error: 'Missing required fields' });
  const result = db.prepare('INSERT INTO restaurants (name, address, lat, lng, phone, category, created_by) VALUES (?,?,?,?,?,?,?)')
    .run(name, address, lat, lng, phone, category, req.user.sub);
  res.status(201).json({ id: result.lastInsertRowid });
});

router.patch('/:id', requireManager, (req, res) => {
  const { name, address, lat, lng, phone, category } = req.body;
  db.prepare('UPDATE restaurants SET name=COALESCE(?,name), address=COALESCE(?,address), lat=COALESCE(?,lat), lng=COALESCE(?,lng), phone=COALESCE(?,phone), category=COALESCE(?,category) WHERE id=?')
    .run(name, address, lat, lng, phone, category, req.params.id);
  res.json({ ok: true });
});

router.delete('/:id', requireManager, (req, res) => {
  db.prepare('DELETE FROM restaurants WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

module.exports = router;
