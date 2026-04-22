const express = require('express');
const bcrypt = require('bcrypt');
const db = require('../db');

const router = express.Router();

router.get('/', (req, res) => {
  const users = db.prepare(`
    SELECT u.id, u.name, u.email, u.role, u.is_active, u.created_at,
           COUNT(a.id) as assigned_count
    FROM users u
    LEFT JOIN assignments a ON a.user_id = u.id
    WHERE u.role = 'rep'
    GROUP BY u.id
    ORDER BY u.name
  `).all();
  res.json(users);
});

router.post('/', (req, res) => {
  const { name, email, password, role = 'rep' } = req.body;
  if (!name || !email || !password) return res.status(400).json({ error: 'Missing fields' });
  const hash = bcrypt.hashSync(password, 10);
  try {
    const result = db.prepare('INSERT INTO users (name, email, password_hash, role) VALUES (?,?,?,?)').run(name, email, hash, role);
    res.status(201).json({ id: result.lastInsertRowid });
  } catch (e) {
    res.status(400).json({ error: 'Email already exists' });
  }
});

router.get('/:id', (req, res) => {
  const user = db.prepare('SELECT id, name, email, role, is_active FROM users WHERE id = ?').get(req.params.id);
  if (!user) return res.status(404).json({ error: 'Not found' });
  const assignments = db.prepare(`
    SELECT r.*, a.assigned_at FROM restaurants r
    JOIN assignments a ON a.restaurant_id = r.id
    WHERE a.user_id = ?
  `).all(req.params.id);
  res.json({ ...user, assignments });
});

router.patch('/:id', (req, res) => {
  const { name, email, is_active } = req.body;
  db.prepare('UPDATE users SET name=COALESCE(?,name), email=COALESCE(?,email), is_active=COALESCE(?,is_active) WHERE id=?')
    .run(name, email, is_active, req.params.id);
  res.json({ ok: true });
});

module.exports = router;
