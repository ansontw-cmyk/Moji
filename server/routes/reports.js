const express = require('express');
const db = require('../db');
const upload = require('../middleware/upload');
const { requireManager } = require('../middleware/auth');

const router = express.Router();

router.get('/', (req, res) => {
  const isManager = req.user.role === 'manager';
  const { restaurant_id, user_id } = req.query;

  let sql = `
    SELECT rp.*, r.name as restaurant_name, u.name as user_name
    FROM reports rp
    JOIN restaurants r ON r.id = rp.restaurant_id
    JOIN users u ON u.id = rp.user_id
    WHERE 1=1
  `;
  const params = [];

  if (!isManager) { sql += ' AND rp.user_id = ?'; params.push(req.user.sub); }
  if (restaurant_id) { sql += ' AND rp.restaurant_id = ?'; params.push(restaurant_id); }
  if (isManager && user_id) { sql += ' AND rp.user_id = ?'; params.push(user_id); }

  sql += ' ORDER BY rp.created_at DESC LIMIT 100';
  res.json(db.prepare(sql).all(...params));
});

router.post('/', upload.single('photo'), (req, res) => {
  const { restaurant_id, contact_name, contact_phone, status, notes } = req.body;
  if (!restaurant_id || !status) return res.status(400).json({ error: 'restaurant_id and status required' });

  const photo_path = req.file ? `/uploads/${req.file.filename}` : null;
  const result = db.prepare(`
    INSERT INTO reports (restaurant_id, user_id, contact_name, contact_phone, status, notes, photo_path)
    VALUES (?,?,?,?,?,?,?)
  `).run(restaurant_id, req.user.sub, contact_name, contact_phone, status, notes, photo_path);

  res.status(201).json({ id: result.lastInsertRowid });
});

router.get('/restaurant/:restaurantId', (req, res) => {
  const reports = db.prepare(`
    SELECT rp.*, u.name as user_name FROM reports rp
    JOIN users u ON u.id = rp.user_id
    WHERE rp.restaurant_id = ?
    ORDER BY rp.created_at DESC
  `).all(req.params.restaurantId);
  res.json(reports);
});

router.get('/:id', (req, res) => {
  const report = db.prepare(`
    SELECT rp.*, r.name as restaurant_name, u.name as user_name
    FROM reports rp
    JOIN restaurants r ON r.id = rp.restaurant_id
    JOIN users u ON u.id = rp.user_id
    WHERE rp.id = ?
  `).get(req.params.id);
  if (!report) return res.status(404).json({ error: 'Not found' });
  if (req.user.role !== 'manager' && report.user_id !== req.user.sub)
    return res.status(403).json({ error: 'Forbidden' });
  res.json(report);
});

router.patch('/:id', upload.single('photo'), (req, res) => {
  const { contact_name, contact_phone, status, notes } = req.body;
  const report = db.prepare('SELECT * FROM reports WHERE id = ?').get(req.params.id);
  if (!report) return res.status(404).json({ error: 'Not found' });
  if (req.user.role !== 'manager' && report.user_id !== req.user.sub)
    return res.status(403).json({ error: 'Forbidden' });
  const photo_path = req.file ? `/uploads/${req.file.filename}` : report.photo_path;
  db.prepare('UPDATE reports SET contact_name=COALESCE(?,contact_name), contact_phone=COALESCE(?,contact_phone), status=COALESCE(?,status), notes=COALESCE(?,notes), photo_path=? WHERE id=?')
    .run(contact_name, contact_phone, status, notes, photo_path, req.params.id);
  res.json({ ok: true });
});

router.delete('/:id', requireManager, (req, res) => {
  db.prepare('DELETE FROM reports WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

module.exports = router;
