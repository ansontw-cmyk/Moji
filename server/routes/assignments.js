const express = require('express');
const db = require('../db');

const router = express.Router();

router.get('/', (req, res) => {
  const assignments = db.prepare(`
    SELECT a.*, r.name as restaurant_name, r.lat, r.lng, r.address,
           u.name as user_name
    FROM assignments a
    JOIN restaurants r ON r.id = a.restaurant_id
    JOIN users u ON u.id = a.user_id
    ORDER BY a.assigned_at DESC
  `).all();
  res.json(assignments);
});

router.post('/', (req, res) => {
  const { restaurant_id, user_id } = req.body;
  if (!restaurant_id || !user_id) return res.status(400).json({ error: 'Missing fields' });
  try {
    db.prepare('INSERT INTO assignments (restaurant_id, user_id, assigned_by) VALUES (?,?,?) ON CONFLICT(restaurant_id) DO UPDATE SET user_id=excluded.user_id, assigned_by=excluded.assigned_by, assigned_at=datetime(\'now\')')
      .run(restaurant_id, user_id, req.user.sub);
    res.status(201).json({ ok: true });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

router.delete('/:id', (req, res) => {
  db.prepare('DELETE FROM assignments WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

router.get('/user/:userId', (req, res) => {
  const assignments = db.prepare(`
    SELECT a.*, r.name as restaurant_name, r.lat, r.lng, r.address, r.category
    FROM assignments a
    JOIN restaurants r ON r.id = a.restaurant_id
    WHERE a.user_id = ?
  `).all(req.params.userId);
  res.json(assignments);
});

module.exports = router;
