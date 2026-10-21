// server.js — QRコード式出勤管理アプリ（sql.js版）

const express = require('express');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const initSqlJs = require('sql.js');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;
const DB_FILE = path.join(__dirname, 'attendance.db');

// ===== DB初期化 =====
let db;

async function initDB() {
  const SQL = await initSqlJs();

  if (fs.existsSync(DB_FILE)) {
    const fileBuffer = fs.readFileSync(DB_FILE);
    db = new SQL.Database(fileBuffer);
  } else {
    db = new SQL.Database();
  }

  db.run(`
    CREATE TABLE IF NOT EXISTS qr_locations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      qr_token TEXT NOT NULL UNIQUE,
      location_name TEXT,
      latitude REAL,
      longitude REAL,
      allowed_radius_m INTEGER DEFAULT 150,
      is_active INTEGER DEFAULT 1,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      attendance_code TEXT NOT NULL UNIQUE,
      login_id TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL,
      is_active INTEGER DEFAULT 1,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS attendance_records (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL,
      attendance_code TEXT NOT NULL,
      qr_location_id INTEGER NOT NULL,
      attendance_date TEXT NOT NULL,
      attendance_time TEXT NOT NULL,
      registered_at TEXT NOT NULL,
      latitude REAL,
      longitude REAL,
      location_accuracy REAL,
      distance_m REAL,
      location_judgement TEXT,
      status TEXT NOT NULL,
      memo TEXT,
      device_info TEXT,
      ip_address TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS operation_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER,
      attendance_code TEXT,
      role TEXT,
      action TEXT NOT NULL,
      detail TEXT,
      ip_address TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
  `);

  // DBをファイルに保存するヘルパー
  saveDB();

  // 初期データ投入
  const locations = [
    { name: '基本法制監理部', token: 'loc-001' },
    { name: '第1部',         token: 'loc-002' },
    { name: '第2部',         token: 'loc-003' },
    { name: '第3部',         token: 'loc-004' },
    { name: '第4部',         token: 'loc-005' },
    { name: '第5部',         token: 'loc-006' },
    { name: '法制主幹',      token: 'loc-007' },
    { name: '総務課人事総務係', token: 'loc-008' },
    { name: '総務課会計情報係', token: 'loc-009' },
    { name: '総務課図書係',  token: 'loc-010' },
    { name: '次室',          token: 'loc-011' },
  ];
  for (const l of locations) {
    db.run(
      `INSERT OR IGNORE INTO qr_locations (name, qr_token, latitude, longitude) VALUES (?, ?, ?, ?)`,
      [l.name, l.token, 35.681236, 139.767125]
    );
  }

  const initialUsers = [
    { code: '9001', role: 'admin', pass: 'admin123' },
    { code: '9002', role: 'admin', pass: 'admin123' },
    { code: '9003', role: 'admin', pass: 'admin123' },
    { code: '1234', role: 'staff', pass: 'user123' },
    { code: '5678', role: 'staff', pass: 'user123' },
    { code: '0425', role: 'staff', pass: 'user123' },
  ];
  for (const u of initialUsers) {
    const hash = bcrypt.hashSync(u.pass, 10);
    db.run(
      `INSERT OR IGNORE INTO users (attendance_code, login_id, password_hash, role) VALUES (?, ?, ?, ?)`,
      [u.code, u.code, hash, u.role]
    );
  }
  saveDB();
}

function saveDB() {
  const data = db.export();
  fs.writeFileSync(DB_FILE, Buffer.from(data));
}

// ===== DBヘルパー関数 =====
function dbGet(sql, params = []) {
  const stmt = db.prepare(sql);
  stmt.bind(params);
  if (stmt.step()) {
    const row = stmt.getAsObject();
    stmt.free();
    return row;
  }
  stmt.free();
  return null;
}

function dbAll(sql, params = []) {
  const results = [];
  const stmt = db.prepare(sql);
  stmt.bind(params);
  while (stmt.step()) {
    results.push(stmt.getAsObject());
  }
  stmt.free();
  return results;
}

function dbRun(sql, params = []) {
  db.run(sql, params);
  saveDB();
  return db;
}

function dbInsert(sql, params = []) {
  db.run(sql, params);
  const row = dbGet('SELECT last_insert_rowid() as id');
  saveDB();
  return row ? row.id : null;
}

// ===== ミドルウェア =====
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
  secret: 'qr-attendance-secret-key',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 }
}));

// ===== ユーティリティ =====
function requireLogin(req, res, next) {
  if (!req.session.userId) return res.redirect('/');
  next();
}

function requireAdmin(req, res, next) {
  if (!req.session.userId || req.session.role !== 'admin') {
    return res.status(403).send('管理者権限が必要です');
  }
  next();
}

function logOperation(userId, code, role, action, detail, ip) {
  dbRun(
    `INSERT INTO operation_logs (user_id, attendance_code, role, action, detail, ip_address) VALUES (?, ?, ?, ?, ?, ?)`,
    [userId || null, code || null, role || null, action, detail || null, ip || null]
  );
}

function calcDistance(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat/2)**2 +
    Math.cos(lat1 * Math.PI/180) * Math.cos(lat2 * Math.PI/180) * Math.sin(dLng/2)**2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function layout(title, body, user) {
  const nav = user
    ? `<header><h1>出勤管理システム</h1><a href="/logout">ログアウト</a></header>`
    : '';
  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title}</title>
  <link rel="stylesheet" href="/style.css">
</head>
<body>${nav}<main>${body}</main></body>
</html>`;
}

// ===== ルーティング =====

app.get('/', (req, res) => {
  if (req.session.userId) {
    return res.redirect(req.session.role === 'admin' ? '/admin' : '/home');
  }
  res.send(`<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ログイン</title>
  <link rel="stylesheet" href="/style.css">
</head>
<body>
<div class="login-wrapper">
  <div class="login-card">
    <h2>出勤管理システム</h2>
    <p class="subtitle">出勤用IDとパスワードでログイン</p>
    ${req.query.error ? '<div class="message message-error">IDまたはパスワードが正しくありません</div>' : ''}
    <form method="POST" action="/login">
      <div class="form-group">
        <label>出勤用ID</label>
        <input type="text" name="login_id" placeholder="例：1234" required autofocus>
      </div>
      <div class="form-group">
        <label>パスワード</label>
        <input type="password" name="password" required>
      </div>
      <button type="submit" class="btn btn-primary">ログイン</button>
    </form>
  </div>
</div>
</body></html>`);
});

app.post('/login', (req, res) => {
  const { login_id, password } = req.body;
  const user = dbGet('SELECT * FROM users WHERE login_id = ? AND is_active = 1', [login_id]);

  if (!user || !bcrypt.compareSync(password, user.password_hash)) {
    logOperation(null, login_id, null, 'LOGIN_FAIL', 'ログイン失敗', req.ip);
    return res.redirect('/?error=1');
  }

  req.session.userId = user.id;
  req.session.role = user.role;
  req.session.attendanceCode = user.attendance_code;

  logOperation(user.id, user.attendance_code, user.role, 'LOGIN', 'ログイン成功', req.ip);
  res.redirect(user.role === 'admin' ? '/admin' : '/home');
});

app.get('/logout', requireLogin, (req, res) => {
  logOperation(req.session.userId, req.session.attendanceCode, req.session.role, 'LOGOUT', null, req.ip);
  req.session.destroy();
  res.redirect('/');
});

app.get('/home', requireLogin, (req, res) => {
  if (req.session.role === 'admin') return res.redirect('/admin');
  res.send(layout('職員ホーム', `
    <div class="card">
      <h2>出勤登録</h2>
      <div class="message message-info">
        QRコードを読み取って出勤登録を行ってください。<br>
        <small>※ 氏名・職員ID・所属はこのシステムには表示されません。</small>
      </div>
      <p style="text-align:center; margin:20px 0;">
        <a href="/qr-scan" class="btn btn-success" style="display:inline-block; padding:16px 40px; font-size:1.1rem;">
          📷 QRコードを読み取る
        </a>
      </p>
    </div>
  `, true));
});

app.get('/scan/:token', requireLogin, (req, res) => {
  if (req.session.role === 'admin') return res.redirect('/admin');
  const loc = dbGet('SELECT * FROM qr_locations WHERE qr_token = ? AND is_active = 1', [req.params.token]);
  if (!loc) return res.send(layout('エラー', '<div class="card"><div class="message message-error">無効なQRコードです。</div></div>', true));

  res.send(layout('出勤登録確認', `
    <div class="card">
      <h2>出勤登録の確認</h2>
      <div class="confirm-item">
        <span class="label">本人確認</span>
        <span class="value">ログイン済みアカウント</span>
      </div>
      <div class="confirm-item">
        <span class="label">QR設置場所</span>
        <span class="value">${loc.name}</span>
      </div>
      <div class="message message-info" style="margin-top:16px; font-size:0.85rem;">
        位置情報を取得して登録します。取得できない場合や範囲外でも登録は完了しますが「要確認」となります。
      </div>
      <form id="attendanceForm" method="POST" action="/api/attendance">
        <input type="hidden" name="qr_token" value="${loc.qr_token}">
        <input type="hidden" name="latitude" id="lat">
        <input type="hidden" name="longitude" id="lng">
        <input type="hidden" name="accuracy" id="acc">
        <button type="button" onclick="getLocationAndSubmit()" class="btn btn-success" style="margin-top:20px;">
          ✅ 出勤登録する
        </button>
      </form>
    </div>
    <script>
    function getLocationAndSubmit() {
      const btn = document.querySelector('button');
      btn.disabled = true;
      btn.textContent = '位置情報取得中...';
      if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(
          pos => {
            document.getElementById('lat').value = pos.coords.latitude;
            document.getElementById('lng').value = pos.coords.longitude;
            document.getElementById('acc').value = pos.coords.accuracy;
            document.getElementById('attendanceForm').submit();
          },
          () => { document.getElementById('attendanceForm').submit(); },
          { timeout: 10000 }
        );
      } else {
        document.getElementById('attendanceForm').submit();
      }
    }
    </script>
  `, true));
});

app.post('/api/attendance', requireLogin, (req, res) => {
  const { qr_token, latitude, longitude, accuracy } = req.body;
  const loc = dbGet('SELECT * FROM qr_locations WHERE qr_token = ? AND is_active = 1', [qr_token]);
  if (!loc) return res.redirect('/home');

  const now = new Date();
  const date = now.toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' }).replace(/\//g, '-');
  const time = now.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

  let distanceM = null, judgement = '位置情報未取得', status = '要確認';
  const lat = parseFloat(latitude), lng = parseFloat(longitude), acc = parseFloat(accuracy);

  if (!isNaN(lat) && !isNaN(lng)) {
    if (loc.latitude && loc.longitude) {
      distanceM = Math.round(calcDistance(lat, lng, loc.latitude, loc.longitude));
      judgement = distanceM <= loc.allowed_radius_m ? '範囲内' : '範囲外';
      status = distanceM <= loc.allowed_radius_m ? '正常' : '要確認';
    } else {
      judgement = '基準未設定';
    }
  }

  const newId = dbInsert(
    `INSERT INTO attendance_records
     (user_id, attendance_code, qr_location_id, attendance_date, attendance_time,
      registered_at, latitude, longitude, location_accuracy, distance_m,
      location_judgement, status, device_info, ip_address)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [
      req.session.userId, req.session.attendanceCode, loc.id,
      date, time, now.toISOString(),
      isNaN(lat) ? null : lat, isNaN(lng) ? null : lng, isNaN(acc) ? null : acc,
      distanceM, judgement, status,
      req.headers['user-agent'], req.ip
    ]
  );

  logOperation(req.session.userId, req.session.attendanceCode, 'staff', 'ATTENDANCE', `場所:${loc.name} 状態:${status}`, req.ip);
  res.redirect(`/attendance-complete?id=${newId}`);
});

app.get('/attendance-complete', requireLogin, (req, res) => {
  const rec = dbGet(
    `SELECT r.*, l.name as location_name FROM attendance_records r
     JOIN qr_locations l ON r.qr_location_id = l.id WHERE r.id = ?`,
    [req.query.id]
  );
  if (!rec || rec.user_id != req.session.userId) return res.redirect('/home');

  const statusClass = rec.status === '正常' ? 'message-success' : 'message-warning';
  res.send(layout('出勤登録完了', `
    <div class="card">
      <h2>出勤登録完了</h2>
      <div class="message ${statusClass}">
        ${rec.status === '正常' ? '✅ 正常に登録されました' : '⚠️ 登録されましたが「要確認」です'}
      </div>
      <div class="confirm-item"><span class="label">出勤日</span><span class="value">${rec.attendance_date}</span></div>
      <div class="confirm-item"><span class="label">出勤時刻</span><span class="value">${rec.attendance_time}</span></div>
      <div class="confirm-item"><span class="label">QR設置場所</span><span class="value">${rec.location_name}</span></div>
      <div class="confirm-item"><span class="label">位置判定</span><span class="value">${rec.location_judgement}</span></div>
      <div class="confirm-item"><span class="label">登録状態</span>
        <span class="value"><span class="badge ${rec.status === '正常' ? 'badge-normal' : 'badge-caution'}">${rec.status}</span></span>
      </div>
      <p style="margin-top:24px; text-align:center;">
        <a href="/home" class="btn btn-secondary" style="display:inline-block;">ホームへ戻る</a>
      </p>
    </div>
  `, true));
});

app.get('/admin', requireAdmin, (req, res) => {
  res.send(layout('管理者ホーム', `
    <div class="card">
      <h2>管理者メニュー</h2>
      <div class="admin-menu">
        <a href="/admin/attendance" class="admin-menu-item">
          <div class="icon">📋</div><div class="title">出勤状況一覧</div>
        </a>
        <a href="/admin/qr" class="admin-menu-item">
          <div class="icon">🔲</div><div class="title">QRコード管理</div>
        </a>
        <a href="/admin/logs" class="admin-menu-item">
          <div class="icon">📝</div><div class="title">操作ログ</div>
        </a>
        <a href="/admin/attendance/excel" class="admin-menu-item">
          <div class="icon">📊</div><div class="title">Excel出力</div>
        </a>
      </div>
    </div>
  `, true));
});

app.get('/admin/attendance', requireAdmin, (req, res) => {
  const { start, end, location, status, code } = req.query;
  let sql = `SELECT r.*, l.name as location_name FROM attendance_records r
             JOIN qr_locations l ON r.qr_location_id = l.id WHERE 1=1`;
  const params = [];
  if (start)    { sql += ' AND r.attendance_date >= ?'; params.push(start); }
  if (end)      { sql += ' AND r.attendance_date <= ?'; params.push(end); }
  if (location) { sql += ' AND r.qr_location_id = ?'; params.push(location); }
  if (status)   { sql += ' AND r.status = ?'; params.push(status); }
  if (code)     { sql += ' AND r.attendance_code = ?'; params.push(code); }
  sql += ' ORDER BY r.registered_at DESC LIMIT 500';

  const records = dbAll(sql, params);
  const locs = dbAll('SELECT * FROM qr_locations WHERE is_active = 1');

  const rows = records.map(r => `
    <tr>
      <td>${r.attendance_code}</td>
      <td>${r.location_name}</td>
      <td>${r.attendance_date}</td>
      <td>${r.attendance_time}</td>
      <td><span class="badge ${r.status === '正常' ? 'badge-normal' : 'badge-caution'}">${r.status}</span></td>
      <td>${r.location_judgement || '-'}</td>
      <td>${r.distance_m !== null ? r.distance_m + 'm' : '-'}</td>
      <td><a href="/admin/attendance/${r.id}">詳細</a></td>
    </tr>`).join('');

  const locOptions = locs.map(l =>
    `<option value="${l.id}" ${location == l.id ? 'selected' : ''}>${l.name}</option>`).join('');

  res.send(layout('出勤状況一覧', `
    <div class="card">
      <h2>出勤状況一覧</h2>
      <form class="filter-form" method="GET">
        <input type="date" name="start" value="${start||''}">
        <input type="date" name="end" value="${end||''}">
        <select name="location"><option value="">全場所</option>${locOptions}</select>
        <select name="status">
          <option value="">全状態</option>
          <option value="正常" ${status==='正常'?'selected':''}>正常</option>
          <option value="要確認" ${status==='要確認'?'selected':''}>要確認</option>
        </select>
        <input type="text" name="code" value="${code||''}" placeholder="出勤用ID">
        <button type="submit">絞り込み</button>
      </form>
      <div class="table-wrapper">
        <table>
          <thead>
            <tr><th>出勤用ID</th><th>QR設置場所</th><th>出勤日</th><th>出勤時刻</th><th>状態</th><th>位置判定</th><th>距離</th><th>詳細</th></tr>
          </thead>
          <tbody>${rows || '<tr><td colspan="8" style="text-align:center">データがありません</td></tr>'}</tbody>
        </table>
      </div>
      <p style="margin-top:16px;">
        <a href="/admin/attendance/excel" class="btn btn-primary" style="display:inline-block; width:auto;">📊 Excel出力</a>
      </p>
    </div>
  `, true));
});

app.get('/admin/attendance/excel', requireAdmin, async (req, res) => {
  const records = dbAll(
    `SELECT r.*, l.name as location_name FROM attendance_records r
     JOIN qr_locations l ON r.qr_location_id = l.id
     ORDER BY r.attendance_date, r.attendance_time`
  );
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('出勤記録');
  sheet.columns = [
    { header: '出勤用ID',   key: 'attendance_code',   width: 12 },
    { header: 'QR設置場所', key: 'location_name',      width: 20 },
    { header: '出勤日',     key: 'attendance_date',    width: 14 },
    { header: '出勤時刻',   key: 'attendance_time',    width: 12 },
    { header: '状態',       key: 'status',             width: 10 },
    { header: '位置判定',   key: 'location_judgement', width: 14 },
    { header: '距離m',      key: 'distance_m',         width: 10 },
    { header: '緯度',       key: 'latitude',           width: 14 },
    { header: '経度',       key: 'longitude',          width: 14 },
    { header: '登録日時',   key: 'registered_at',      width: 22 },
    { header: '備考',       key: 'memo',               width: 20 },
  ];
  sheet.getRow(1).font = { bold: true };
  records.forEach(r => sheet.addRow(r));
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="attendance.xlsx"');
  await workbook.xlsx.write(res);
  res.end();
});

app.get('/admin/attendance/:id', requireAdmin, (req, res) => {
  const rec = dbGet(
    `SELECT r.*, l.name as location_name FROM attendance_records r
     JOIN qr_locations l ON r.qr_location_id = l.id WHERE r.id = ?`,
    [req.params.id]
  );
  if (!rec) return res.redirect('/admin/attendance');
  const items = [
    ['出勤用ID', rec.attendance_code], ['QR設置場所', rec.location_name],
    ['出勤日', rec.attendance_date],   ['出勤時刻', rec.attendance_time],
    ['登録日時', rec.registered_at],
    ['状態', `<span class="badge ${rec.status==='正常'?'badge-normal':'badge-caution'}">${rec.status}</span>`],
    ['位置判定', rec.location_judgement||'-'], ['距離', rec.distance_m!==null?rec.distance_m+'m':'-'],
    ['緯度', rec.latitude||'-'],       ['経度', rec.longitude||'-'],
    ['位置精度', rec.location_accuracy?rec.location_accuracy+'m':'-'],
    ['IPアドレス', rec.ip_address||'-'], ['端末情報', rec.device_info||'-'],
    ['備考', rec.memo||'-'],
  ];
  res.send(layout('出勤詳細', `
    <div class="card">
      <h2>出勤詳細</h2>
      ${items.map(([l,v])=>`<div class="confirm-item"><span class="label">${l}</span><span class="value">${v}</span></div>`).join('')}
      <p style="margin-top:20px;">
        <a href="/admin/attendance" class="btn btn-secondary" style="display:inline-block; width:auto;">← 一覧へ戻る</a>
      </p>
    </div>
  `, true));
});

app.get('/admin/qr', requireAdmin, async (req, res) => {
  const locs = dbAll('SELECT * FROM qr_locations WHERE is_active = 1');
  const baseUrl = `https://qr-attendance-app-60gi.onrender.com`;
  const qrItems = await Promise.all(locs.map(async l => {
    const url = `${baseUrl}/scan/${l.qr_token}`;
    const qrDataUrl = await QRCode.toDataURL(url, { width: 180 });
    return `<div class="qr-item">
      <img src="${qrDataUrl}" alt="${l.name}" style="width:160px;">
      <p>${l.name}</p>
    </div>`;
  }));
  res.send(layout('QRコード管理', `
    <div class="card">
      <h2>QRコード一覧（全11か所）</h2>
      <div class="message message-warning" style="margin-bottom:16px; font-size:0.85rem;">
        ⚠️ QRコードは印刷して各場所に掲示してください。
      </div>
      <div class="qr-grid">${qrItems.join('')}</div>
    </div>
  `, true));
});

app.get('/admin/logs', requireAdmin, (req, res) => {
  const logs = dbAll('SELECT * FROM operation_logs ORDER BY created_at DESC LIMIT 300');
  const rows = logs.map(l => `
    <tr>
      <td>${l.id}</td><td>${l.created_at}</td><td>${l.attendance_code||'-'}</td>
      <td>${l.role||'-'}</td><td>${l.action}</td>
      <td style="font-size:0.8rem;">${l.detail||'-'}</td>
      <td style="font-size:0.8rem;">${l.ip_address||'-'}</td>
    </tr>`).join('');
  res.send(layout('操作ログ', `
    <div class="card">
      <h2>操作ログ（最新300件）</h2>
      <div class="table-wrapper">
        <table>
          <thead><tr><th>ID</th><th>日時</th><th>出勤用ID</th><th>権限</th><th>操作</th><th>詳細</th><th>IP</th></tr></thead>
          <tbody>${rows||'<tr><td colspan="7" style="text-align:center">ログがありません</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `, true));
});

// ===== サーバー起動 =====
initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`QR Attendance App started.`);
    console.log(`http://localhost:${PORT}`);
    console.log(`\nテストアカウント:`);
    console.log(`  管理者: ID=9001 / PW=admin123`);
    console.log(`  職員:   ID=1234 / PW=user123`);
  });
}).catch(err => {
  console.error('DB初期化エラー:', err);
});