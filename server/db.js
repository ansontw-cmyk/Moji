const Database = require('better-sqlite3');
const path = require('path');
const bcrypt = require('bcrypt');

const db = new Database(path.join(__dirname, '../data/crm.sqlite'));
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    name          TEXT    NOT NULL,
    email         TEXT    NOT NULL UNIQUE,
    password_hash TEXT    NOT NULL,
    role          TEXT    NOT NULL DEFAULT 'rep' CHECK(role IN ('rep','manager')),
    created_at    TEXT    NOT NULL DEFAULT (datetime('now')),
    is_active     INTEGER NOT NULL DEFAULT 1 CHECK(is_active IN (0,1))
  );

  CREATE TABLE IF NOT EXISTS restaurants (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT    NOT NULL,
    address     TEXT,
    lat         REAL    NOT NULL,
    lng         REAL    NOT NULL,
    phone       TEXT,
    category    TEXT,
    created_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    created_by  INTEGER REFERENCES users(id)
  );

  CREATE INDEX IF NOT EXISTS idx_restaurants_lat_lng ON restaurants(lat, lng);

  CREATE TABLE IF NOT EXISTS assignments (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    restaurant_id INTEGER NOT NULL REFERENCES restaurants(id) ON DELETE CASCADE,
    user_id       INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    assigned_by   INTEGER NOT NULL REFERENCES users(id),
    assigned_at   TEXT    NOT NULL DEFAULT (datetime('now')),
    UNIQUE(restaurant_id)
  );

  CREATE INDEX IF NOT EXISTS idx_assignments_user ON assignments(user_id);
  CREATE INDEX IF NOT EXISTS idx_assignments_restaurant ON assignments(restaurant_id);

  CREATE TABLE IF NOT EXISTS reports (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    restaurant_id INTEGER NOT NULL REFERENCES restaurants(id),
    user_id       INTEGER NOT NULL REFERENCES users(id),
    contact_name  TEXT,
    contact_phone TEXT,
    status        TEXT NOT NULL CHECK(status IN (
                    'not_visited','visited_interested','visited_not_interested',
                    'follow_up','closed_won','closed_lost'
                  )),
    notes         TEXT,
    photo_path    TEXT,
    visited_at    TEXT NOT NULL DEFAULT (datetime('now')),
    created_at    TEXT NOT NULL DEFAULT (datetime('now'))
  );

  CREATE INDEX IF NOT EXISTS idx_reports_restaurant ON reports(restaurant_id);
  CREATE INDEX IF NOT EXISTS idx_reports_user ON reports(user_id);
`);

// Seed demo data if empty
const userCount = db.prepare('SELECT COUNT(*) as c FROM users').get().c;
if (userCount === 0) {
  const hash = bcrypt.hashSync('password123', 10);
  db.prepare(`INSERT INTO users (name, email, password_hash, role) VALUES (?,?,?,?)`).run('Manager Wang', 'manager@demo.com', hash, 'manager');
  db.prepare(`INSERT INTO users (name, email, password_hash, role) VALUES (?,?,?,?)`).run('Rep A - 小明', 'rep1@demo.com', hash, 'rep');
  db.prepare(`INSERT INTO users (name, email, password_hash, role) VALUES (?,?,?,?)`).run('Rep B - 小美', 'rep2@demo.com', hash, 'rep');

  // Seed restaurants around Taipei (Xinyi District)
  const restaurants = [
    ['鼎泰豐 信義店', '台北市信義區信義路五段150巷', 25.0330, 121.5654, '02-8101-7799', '中式'],
    ['瓦城泰國料理', '台北市信義區松仁路58號', 25.0359, 121.5680, '02-2723-0677', '泰式'],
    ['胡同燒肉 微風信義', '台北市信義區忠孝東路五段68號', 25.0413, 121.5685, '02-2346-3499', '燒肉'],
    ['饗食天堂 信義店', '台北市信義區松壽路12號', 25.0352, 121.5679, '02-8780-2898', '吃到飽'],
    ['王品牛排', '台北市信義區松仁路100號', 25.0340, 121.5670, '02-2723-1122', '牛排'],
    ['一蘭拉麵 統一時代', '台北市信義區忠孝東路四段553號', 25.0413, 121.5655, '02-2748-3388', '日式'],
    ['海底撈火鍋', '台北市信義區松壽路9號', 25.0349, 121.5676, '02-8780-5678', '火鍋'],
    ['Pizza Hut 信義店', '台北市信義區信義路五段7號', 25.0336, 121.5649, '02-2720-1234', '西式'],
  ];

  const insertR = db.prepare(`INSERT INTO restaurants (name, address, lat, lng, phone, category, created_by) VALUES (?,?,?,?,?,?,1)`);
  restaurants.forEach(r => insertR.run(...r));

  // Assign some restaurants
  db.prepare(`INSERT INTO assignments (restaurant_id, user_id, assigned_by) VALUES (1,2,1)`).run();
  db.prepare(`INSERT INTO assignments (restaurant_id, user_id, assigned_by) VALUES (2,2,1)`).run();
  db.prepare(`INSERT INTO assignments (restaurant_id, user_id, assigned_by) VALUES (3,3,1)`).run();
  db.prepare(`INSERT INTO assignments (restaurant_id, user_id, assigned_by) VALUES (4,3,1)`).run();
}

module.exports = db;
