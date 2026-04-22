require('dotenv').config();
const express = require('express');
const path = require('path');

const { requireAuth, requireManager } = require('./middleware/auth');

const app = express();
app.use(express.json());
app.use('/uploads', express.static(path.join(__dirname, '../uploads')));

app.use('/api/auth', require('./routes/auth'));
app.use('/api/users',       requireAuth, requireManager, require('./routes/users'));
app.use('/api/restaurants', requireAuth, require('./routes/restaurants'));
app.use('/api/assignments', requireAuth, requireManager, require('./routes/assignments'));
app.use('/api/reports',     requireAuth, require('./routes/reports'));

app.use(express.static(path.join(__dirname, '../public')));
app.get('*', (req, res) => {
  if (!req.path.startsWith('/api')) {
    res.sendFile(path.join(__dirname, '../public/index.html'));
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`friDay CRM running on http://localhost:${PORT}`));
