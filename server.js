// server.js — QRコード式出勤管理アプリ（PostgreSQL版）

const express = require('express');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const { Pool } = require('pg');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// ===== DB接続 =====
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

async function query(sql, params = []) {
  const client = await pool.connect();
  try {
    const result = await client.query(sql, params);
    return result.rows;
  } finally {
    client.release();
  }
}

async function queryOne(sql, params = []) {
  const rows = await query(sql, params);
  return rows[0] || null;
}

// ===== DB初期化 =====
async function initDB() {
  await query(`
    CREATE TABLE IF NOT EXISTS qr_locations (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL,
      qr_token TEXT NOT NULL UNIQUE,
      is_active INTEGER DEFAULT 1,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY,
      attendance_code TEXT NOT NULL UNIQUE,
      login_id TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL,
      is_active INTEGER DEFAULT 1,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS attendance_records (
      id SERIAL PRIMARY KEY,
      user_id INTEGER NOT NULL,
      attendance_code TEXT NOT NULL,
      qr_location_id INTEGER NOT NULL,
      attendance_date TEXT NOT NULL,
      attendance_time TEXT NOT NULL,
      registered_at TEXT NOT NULL,
      device_info TEXT,
      ip_address TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS operation_logs (
      id SERIAL PRIMARY KEY,
      user_id INTEGER,
      attendance_code TEXT,
      role TEXT,
      action TEXT NOT NULL,
      detail TEXT,
      ip_address TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

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
    await query(
      `INSERT INTO qr_locations (name, qr_token) VALUES ($1, $2) ON CONFLICT (qr_token) DO NOTHING`,
      [l.name, l.token]
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
    await query(
      `INSERT INTO users (attendance_code, login_id, password_hash, role) VALUES ($1, $2, $3, $4) ON CONFLICT (attendance_code) DO NOTHING`,
      [u.code, u.code, hash, u.role]
    );
  }

  console.log('DB初期化完了');
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

function requireLogin(req, res, next) {
  if (!req.session.userId) return res.redirect('/');
  next();
}

function requireAdmin(req, res, next) {
  if (!req.session.userId || req.session.role !== 'admin') return res.status(403).send('管理者権限が必要です');
  next();
}

async function logOperation(userId, code, role, action, detail, ip) {
  await query(
    `INSERT INTO operation_logs (user_id, attendance_code, role, action, detail, ip_address) VALUES ($1, $2, $3, $4, $5, $6)`,
    [userId || null, code || null, role || null, action, detail || null, ip || null]
  );
}

function layout(title, body, user) {
  const nav = user ? `<header><h1>出勤管理システム</h1><a href="/logout">ログアウト</a></header>` : '';
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

// ===== QRコード読み取り =====
app.get('/scan/:token', async (req, res) => {
  const loc = await queryOne('SELECT * FROM qr_locations WHERE qr_token = $1 AND is_active = 1', [req.params.token]);
  if (!loc) return res.send(layout('エラー', '<div class="card"><div class="message message-error">無効なQRコードです。</div></div>', false));
  if (req.session.userId) return res.redirect(`/confirm?token=${req.params.token}`);
  res.redirect(`/?token=${req.params.token}`);
});

// ===== ログイン画面 =====
app.get('/', async (req, res) => {
  if (req.session.userId) return res.redirect(req.session.role === 'admin' ? '/admin' : '/home');

  const token = req.query.token || '';
  const error = req.query.error ? '<div class="message message-error">IDまたはパスワードが正しくありません</div>' : '';

  let locationInfo = '';
  if (token) {
    const loc = await queryOne('SELECT * FROM qr_locations WHERE qr_token = $1 AND is_active = 1', [token]);
    if (loc) {
      locationInfo = `<div class="message message-info">📍 QR設置場所：<strong>${loc.name}</strong><br>ログインして出勤登録を完了してください。</div>`;
    }
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
    ${error}
    ${locationInfo}
    <form method="POST" action="/login">
      <input type="hidden" name="token" value="${token}">
      <div class="form-group">
        <label>出勤用ID</label>
        <input type="text" name="login_id" placeholder="例：1234" required autofocus>
      </div>
      <div class="form-group">
        <label>パスワード</label>
        <input type="password" name="password" required>
      </div>
      <button type="submit" class="btn btn-primary" style="font-size:1.2rem; padding:16px; margin-top:8px;">
        🔑 ログイン
      </button>
    </form>
    ${!token ? `
    <div style="margin-top:28px; border-top:1px solid #e0e7ff; padding-top:20px;">
      <p style="text-align:center; color:#666; font-size:0.85rem; margin-bottom:12px;">
        ── まだQRコードを読み取っていない場合 ──
      </p>
      <a href="/qr-reader" class="btn-qr-link">
        <span style="font-size:2rem;">📷</span>
        <span>
          <strong>QRコードを読み取る</strong><br>
          <small>出勤場所のQRコードをカメラで読み取ってから<br>ログインする場合はこちら</small>
        </span>
      </a>
    </div>` : `
    <p style="text-align:center; margin-top:16px; color:#888; font-size:0.8rem;">
      ✅ QRコード読み取り済み・ログインして登録を完了してください
    </p>`}
  </div>
</div>
</body>
</html>`);
});

// ===== ログイン処理 =====
app.post('/login', async (req, res) => {
  const { login_id, password, token } = req.body;
  const user = await queryOne('SELECT * FROM users WHERE login_id = $1 AND is_active = 1', [login_id]);

  if (!user || !bcrypt.compareSync(password, user.password_hash)) {
    await logOperation(null, login_id, null, 'LOGIN_FAIL', 'ログイン失敗', req.ip);
    const tokenParam = token ? `&token=${token}` : '';
    return res.redirect(`/?error=1${tokenParam}`);
  }

  req.session.userId = user.id;
  req.session.role = user.role;
  req.session.attendanceCode = user.attendance_code;
  req.session.qrToken = token || null;
  await logOperation(user.id, user.attendance_code, user.role, 'LOGIN', 'ログイン成功', req.ip);

  if (token) return res.redirect(`/confirm?token=${token}`);
  res.redirect(user.role === 'admin' ? '/admin' : '/home');
});

// ===== ログアウト =====
app.get('/logout', requireLogin, async (req, res) => {
  await logOperation(req.session.userId, req.session.attendanceCode, req.session.role, 'LOGOUT', null, req.ip);
  req.session.destroy();
  res.redirect('/');
});

// ===== QRコードリーダー画面 =====
app.get('/qr-reader', (req, res) => {
  res.send(layout('QRコード読み取り', `
    <div class="card">
      <h2>QRコードを読み取る</h2>
      <p style="color:#555; margin-bottom:16px; font-size:0.9rem;">出勤場所に掲示されているQRコードにカメラを向けてください。</p>
      <div id="reader" style="width:100%; max-width:400px; margin:0 auto;"></div>
      <p id="result" style="text-align:center; color:red; margin-top:12px;"></p>
      <script src="https://unpkg.com/html5-qrcode"></script>
      <script>
        const html5QrCode = new Html5Qrcode("reader");
        html5QrCode.start(
          { facingMode: "environment" },
          { fps: 10, qrbox: { width: 250, height: 250 } },
          (decodedText) => { html5QrCode.stop(); window.location.href = decodedText; },
          (err) => {}
        ).catch(err => {
          document.getElementById('result').textContent = 'カメラの起動に失敗しました。ブラウザのカメラ許可を確認してください。';
        });
      </script>
      <p style="text-align:center; margin-top:20px;">
        <a href="/" style="color:#4f46e5; font-size:0.9rem;">← ログイン画面へ戻る</a>
      </p>
    </div>
  `, false));
});

// ===== 出勤登録確認画面 =====
app.get('/confirm', requireLogin, async (req, res) => {
  if (req.session.role === 'admin') return res.redirect('/admin');
  const token = req.query.token || req.session.qrToken;
  const loc = await queryOne('SELECT * FROM qr_locations WHERE qr_token = $1 AND is_active = 1', [token]);
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
      <form method="POST" action="/api/attendance" style="margin-top:24px;">
        <input type="hidden" name="qr_token" value="${loc.qr_token}">
        <button type="submit" class="btn btn-success" style="font-size:1.2rem; padding:18px;">
          ✅ 出勤登録する
        </button>
      </form>
    </div>
  `, true));
});

// ===== 職員ホーム =====
app.get('/home', requireLogin, (req, res) => {
  if (req.session.role === 'admin') return res.redirect('/admin');
  res.send(layout('職員ホーム', `
    <div class="card">
      <h2>出勤登録</h2>
      <div class="message message-info">
        出勤場所のQRコードを読み取って出勤登録を行ってください。<br>
        <small>※ 氏名・職員ID・所属はこのシステムには表示されません。</small>
      </div>
      <div id="reader" style="width:100%; max-width:400px; margin:20px auto;"></div>
      <p id="result" style="text-align:center; color:red;"></p>
      <script src="https://unpkg.com/html5-qrcode"></script>
      <script>
        const html5QrCode = new Html5Qrcode("reader");
        html5QrCode.start(
          { facingMode: "environment" },
          { fps: 10, qrbox: { width: 250, height: 250 } },
          (decodedText) => { html5QrCode.stop(); window.location.href = decodedText; },
          (err) => {}
        ).catch(err => {
          document.getElementById('result').textContent = 'カメラの起動に失敗しました。ブラウザのカメラ許可を確認してください。';
        });
      </script>
    </div>
  `, true));
});

// ===== 出勤登録API =====
app.post('/api/attendance', requireLogin, async (req, res) => {
  const { qr_token } = req.body;
  const loc = await queryOne('SELECT * FROM qr_locations WHERE qr_token = $1 AND is_active = 1', [qr_token]);
  if (!loc) return res.redirect('/home');

  const now = new Date();
  const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const date = jstNow.toISOString().slice(0, 10);
  const time = jstNow.toISOString().slice(11, 19);

  const result = await query(
    `INSERT INTO attendance_records
     (user_id, attendance_code, qr_location_id, attendance_date, attendance_time, registered_at, device_info, ip_address)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8) RETURNING id`,
    [req.session.userId, req.session.attendanceCode, loc.id,
     date, time, now.toISOString(),
     req.headers['user-agent'], req.ip]
  );

  await logOperation(req.session.userId, req.session.attendanceCode, 'staff', 'ATTENDANCE', `場所:${loc.name}`, req.ip);
  res.redirect(`/attendance-complete?id=${result[0].id}`);
});

// ===== 出勤登録完了 =====
app.get('/attendance-complete', requireLogin, async (req, res) => {
  const rec = await queryOne(
    `SELECT r.*, l.name as location_name FROM attendance_records r
     JOIN qr_locations l ON r.qr_location_id = l.id WHERE r.id = $1`,
    [req.query.id]
  );
  if (!rec || rec.user_id != req.session.userId) return res.redirect('/home');

  res.send(layout('出勤登録完了', `
    <div class="card">
      <h2>出勤登録完了</h2>
      <div class="message message-success">✅ 出勤登録が完了しました</div>
      <div class="confirm-item"><span class="label">出勤日</span><span class="value">${rec.attendance_date}</span></div>
      <div class="confirm-item"><span class="label">出勤時刻</span><span class="value">${rec.attendance_time}</span></div>
      <div class="confirm-item"><span class="label">QR設置場所</span><span class="value">${rec.location_name}</span></div>
      <p style="margin-top:24px; text-align:center;">
        <a href="/logout" class="btn btn-secondary" style="display:inline-block;">ログアウト</a>
      </p>
    </div>
  `, true));
});

// ===== 管理者ホーム =====
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
        <a href="/admin/users" class="admin-menu-item">
          <div class="icon">👥</div><div class="title">ユーザー管理</div>
        </a>
      </div>
    </div>
  `, true));
});

// ===== 管理者：出勤状況一覧 =====
app.get('/admin/attendance', requireAdmin, async (req, res) => {
  const { start, end, location, code } = req.query;
  let sql = `SELECT r.*, l.name as location_name FROM attendance_records r
             JOIN qr_locations l ON r.qr_location_id = l.id WHERE 1=1`;
  const params = [];
  let i = 1;
  if (start)    { sql += ` AND r.attendance_date >= $${i++}`; params.push(start); }
  if (end)      { sql += ` AND r.attendance_date <= $${i++}`; params.push(end); }
  if (location) { sql += ` AND r.qr_location_id = $${i++}`; params.push(location); }
  if (code)     { sql += ` AND r.attendance_code = $${i++}`; params.push(code); }
  sql += ' ORDER BY r.registered_at DESC LIMIT 500';

  const records = await query(sql, params);
  const locs = await query('SELECT * FROM qr_locations WHERE is_active = 1');

  const rows = records.map(r => `
    <tr>
      <td>${r.attendance_code}</td>
      <td>${r.location_name}</td>
      <td>${r.attendance_date}</td>
      <td>${r.attendance_time}</td>
      <td><a href="/admin/attendance/${r.id}">詳細</a></td>
    </tr>`).join('');

  const locOptions = locs.map(l =>
    `<option value="${l.id}" ${location == l.id ? 'selected' : ''}>${l.name}</option>`).join('');

  res.send(layout('出勤状況一覧', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin" class="btn btn-secondary" style="display:inline-block; width:auto;">← 管理者ホームへ</a>
      </p>
      <h2>出勤状況一覧</h2>
      <form class="filter-form" method="GET">
        <input type="date" name="start" value="${start||''}">
        <input type="date" name="end" value="${end||''}">
        <select name="location"><option value="">全場所</option>${locOptions}</select>
        <input type="text" name="code" value="${code||''}" placeholder="出勤用ID">
        <button type="submit">絞り込み</button>
      </form>
      <div class="table-wrapper">
        <table>
          <thead>
            <tr><th>出勤用ID</th><th>QR設置場所</th><th>出勤日</th><th>出勤時刻</th><th>詳細</th></tr>
          </thead>
          <tbody>${rows || '<tr><td colspan="5" style="text-align:center">データがありません</td></tr>'}</tbody>
        </table>
      </div>
      <p style="margin-top:16px;">
        <a href="/admin/attendance/excel" class="btn btn-primary" style="display:inline-block; width:auto;">📊 Excel出力</a>
      </p>
    </div>
  `, true));
});

// ===== Excel出力 =====
app.get('/admin/attendance/excel', requireAdmin, async (req, res) => {
  const records = await query(
    `SELECT r.*, l.name as location_name FROM attendance_records r
     JOIN qr_locations l ON r.qr_location_id = l.id
     ORDER BY r.attendance_date, r.attendance_time`
  );
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('出勤記録');
  sheet.columns = [
    { header: '出勤用ID',   key: 'attendance_code', width: 12 },
    { header: 'QR設置場所', key: 'location_name',   width: 20 },
    { header: '出勤日',     key: 'attendance_date', width: 14 },
    { header: '出勤時刻',   key: 'attendance_time', width: 12 },
    { header: '登録日時',   key: 'registered_at',   width: 22 },
  ];
  sheet.getRow(1).font = { bold: true };
  records.forEach(r => sheet.addRow(r));
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="attendance.xlsx"');
  await workbook.xlsx.write(res);
  res.end();
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'EXCEL_EXPORT', null, req.ip);
});

// ===== 管理者：出勤詳細 =====
app.get('/admin/attendance/:id', requireAdmin, async (req, res) => {
  const rec = await queryOne(
    `SELECT r.*, l.name as location_name FROM attendance_records r
     JOIN qr_locations l ON r.qr_location_id = l.id WHERE r.id = $1`,
    [req.params.id]
  );
  if (!rec) return res.redirect('/admin/attendance');
  const items = [
    ['出勤用ID', rec.attendance_code],
    ['QR設置場所', rec.location_name],
    ['出勤日', rec.attendance_date],
    ['出勤時刻', rec.attendance_time],
    ['登録日時', rec.registered_at],
    ['IPアドレス', rec.ip_address || '-'],
    ['端末情報', rec.device_info || '-'],
  ];
  res.send(layout('出勤詳細', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin/attendance" class="btn btn-secondary" style="display:inline-block; width:auto;">← 出勤状況一覧へ</a>
      </p>
      <h2>出勤詳細</h2>
      ${items.map(([l,v]) => `<div class="confirm-item"><span class="label">${l}</span><span class="value">${v}</span></div>`).join('')}
    </div>
  `, true));
});

// ===== QRコード管理 =====
app.get('/admin/qr', requireAdmin, async (req, res) => {
  const locs = await query('SELECT * FROM qr_locations WHERE is_active = 1');
  const baseUrl = `https://qr-attendance-app-60gi.onrender.com`;
  const qrItems = await Promise.all(locs.map(async l => {
    const url = `${baseUrl}/scan/${l.qr_token}`;
    const qrDataUrl = await QRCode.toDataURL(url, { width: 200 });
    const mailBody = encodeURIComponent(`出勤登録用QRコードをお送りします。\n\n設置場所：${l.name}\n\nQRコードURL：${url}\n\nスマートフォンのカメラでQRコードを読み取るか、上記URLにアクセスして出勤登録を行ってください。`);
    const mailSubject = encodeURIComponent(`【出勤管理】${l.name} QRコードのご案内`);
    return `
      <div class="qr-item">
        <p style="font-weight:bold; font-size:1rem; color:#333; margin-bottom:8px;">${l.name}</p>
        <img src="${qrDataUrl}" alt="${l.name}" style="width:180px; display:block; margin:0 auto;">
        <small style="color:#888; font-size:0.7rem; word-break:break-all; display:block; margin:8px 0;">${url}</small>
        <div style="display:flex; gap:8px; justify-content:center; margin-top:8px;">
          <a href="/admin/qr/print/${l.qr_token}" target="_blank"
             style="background:#4f46e5; color:white; padding:6px 12px; border-radius:8px; font-size:0.8rem; text-decoration:none;">
            🖨️ 印刷
          </a>
          <a href="mailto:?subject=${mailSubject}&body=${mailBody}"
             style="background:#22c55e; color:white; padding:6px 12px; border-radius:8px; font-size:0.8rem; text-decoration:none;">
            📧 メール
          </a>
        </div>
      </div>`;
  }));
  res.send(layout('QRコード管理', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin" class="btn btn-secondary" style="display:inline-block; width:auto;">← 管理者ホームへ</a>
      </p>
      <h2>QRコード一覧（全11か所）</h2>
      <div class="message message-warning" style="margin-bottom:16px; font-size:0.85rem;">
        ⚠️ 各QRコードの「印刷」ボタンで個別印刷、「メール」ボタンでメール送付できます。
      </div>
      <p style="margin-bottom:16px;">
        <a href="/admin/qr/print-all" target="_blank" class="btn btn-primary" style="display:inline-block; width:auto;">
          🖨️ 全ロケーション一括印刷
        </a>
      </p>
      <div class="qr-grid">${qrItems.join('')}</div>
    </div>
  `, true));
});

// ===== QRコード全体印刷ページ =====
app.get('/admin/qr/print-all', requireAdmin, async (req, res) => {
  const locs = await query('SELECT * FROM qr_locations WHERE is_active = 1');
  const baseUrl = `https://qr-attendance-app-60gi.onrender.com`;
  const pages = await Promise.all(locs.map(async l => {
    const url = `${baseUrl}/scan/${l.qr_token}`;
    const qrDataUrl = await QRCode.toDataURL(url, { width: 300 });
    return `
      <div class="print-page">
        <div class="print-card">
          <p class="title">出勤登録QRコード</p>
          <p class="subtitle">${l.name}</p>
          <img src="${qrDataUrl}" alt="${l.name}" class="qr-img">
          <div class="instruction">
            📱 スマートフォンのカメラで<br>
            このQRコードを読み取って<br>
            出勤登録を行ってください
          </div>
        </div>
      </div>`;
  }));
  res.send(`<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>QRコード一括印刷</title>
  <style>
    * { margin:0; padding:0; box-sizing:border-box; }
    body { font-family:'Hiragino Sans','Meiryo',sans-serif; background:#f0f4f8; }
    .no-print { text-align:center; padding:24px; background:#4f46e5; color:white; }
    .no-print h1 { font-size:1.2rem; margin-bottom:12px; }
    .no-print-btns { display:flex; gap:12px; justify-content:center; }
    .btn { padding:10px 24px; border-radius:8px; font-size:0.95rem; cursor:pointer; border:none; font-weight:bold; }
    .btn-print { background:white; color:#4f46e5; }
    .btn-close { background:rgba(255,255,255,0.2); color:white; }
    .pages-wrapper { padding:24px; display:flex; flex-direction:column; gap:24px; align-items:center; }
    .print-page { background:white; width:210mm; min-height:297mm; display:flex; align-items:center; justify-content:center; box-shadow:0 2px 12px rgba(0,0,0,0.1); }
    .print-card { text-align:center; padding:48px 40px; width:100%; }
    .title { font-size:1.6rem; font-weight:bold; color:#1e1b4b; margin-bottom:10px; }
    .subtitle { font-size:1.3rem; font-weight:bold; color:#4f46e5; margin-bottom:32px; padding:10px 24px; background:#e0e7ff; border-radius:8px; display:inline-block; }
    .qr-img { width:280px; height:280px; margin:0 auto 32px; display:block; }
    .instruction { font-size:1rem; color:#444; line-height:1.8; background:#f0f4ff; border-radius:10px; padding:16px 24px; display:inline-block; }
    @media print {
      .no-print { display:none; }
      body { background:white; }
      .pages-wrapper { padding:0; gap:0; }
      .print-page { width:100%; min-height:100vh; page-break-after:always; box-shadow:none; margin:0; }
      .print-page:last-child { page-break-after:avoid; }
    }
  </style>
</head>
<body>
  <div class="no-print">
    <h1>🖨️ QRコード一括印刷（全${locs.length}か所・1ページ1ロケーション）</h1>
    <div class="no-print-btns">
      <button class="btn btn-print" onclick="window.print()">🖨️ 印刷する</button>
      <button class="btn btn-close" onclick="window.close()">✕ 閉じる</button>
    </div>
  </div>
  <div class="pages-wrapper">${pages.join('')}</div>
</body>
</html>`);
});

// ===== QRコード個別印刷ページ =====
app.get('/admin/qr/print/:token', requireAdmin, async (req, res) => {
  const loc = await queryOne('SELECT * FROM qr_locations WHERE qr_token = $1 AND is_active = 1', [req.params.token]);
  if (!loc) return res.redirect('/admin/qr');
  const baseUrl = `https://qr-attendance-app-60gi.onrender.com`;
  const url = `${baseUrl}/scan/${loc.qr_token}`;
  const qrDataUrl = await QRCode.toDataURL(url, { width: 300 });
  res.send(`<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>${loc.name} QRコード</title>
  <style>
    * { margin:0; padding:0; box-sizing:border-box; }
    body { font-family:'Hiragino Sans','Meiryo',sans-serif; display:flex; align-items:center; justify-content:center; min-height:100vh; background:white; }
    .print-card { text-align:center; padding:40px; border:3px solid #4f46e5; border-radius:16px; width:340px; }
    .title { font-size:1.4rem; font-weight:bold; color:#1e1b4b; margin-bottom:8px; }
    .subtitle { font-size:1.1rem; font-weight:bold; color:#4f46e5; margin-bottom:24px; padding:8px 16px; background:#e0e7ff; border-radius:8px; display:inline-block; }
    .qr-img { width:260px; height:260px; margin:0 auto 20px; display:block; }
    .instruction { font-size:0.85rem; color:#444; line-height:1.6; background:#f0f4ff; border-radius:8px; padding:12px; margin-bottom:20px; }
    .no-print { margin-top:20px; display:flex; gap:10px; justify-content:center; }
    .btn { padding:10px 20px; border-radius:8px; font-size:0.9rem; cursor:pointer; border:none; font-weight:bold; }
    .btn-print { background:#4f46e5; color:white; }
    .btn-close { background:#e5e7eb; color:#374151; }
    @media print { .no-print { display:none; } }
  </style>
</head>
<body>
  <div class="print-card">
    <p class="title">出勤登録QRコード</p>
    <p class="subtitle">${loc.name}</p>
    <img src="${qrDataUrl}" alt="${loc.name}" class="qr-img">
    <div class="instruction">
      📱 スマートフォンのカメラで<br>
      このQRコードを読み取って<br>
      出勤登録を行ってください
    </div>
    <div class="no-print">
      <button class="btn btn-print" onclick="window.print()">🖨️ 印刷する</button>
      <button class="btn btn-close" onclick="window.close()">✕ 閉じる</button>
    </div>
  </div>
</body>
</html>`);
});

// ===== 管理者：ユーザー管理 =====
app.get('/admin/users', requireAdmin, async (req, res) => {
  const result = req.query.result ? decodeURIComponent(req.query.result) : '';
  const users = await query(`SELECT id, attendance_code, role, is_active, created_at FROM users ORDER BY is_active DESC, attendance_code`);
  const rows = users.map(u => `
    <tr style="${!u.is_active ? 'opacity:0.5; background:#f9fafb;' : ''}">
      <td style="padding:14px 12px;">${u.attendance_code}</td>
      <td style="padding:14px 12px;">${u.role === 'admin' ? '管理者' : '職員'}</td>
      <td style="padding:14px 12px;"><span class="badge ${u.is_active ? 'badge-normal' : 'badge-caution'}">${u.is_active ? '有効' : '無効'}</span></td>
      <td style="padding:14px 12px;">${u.created_at}</td>
      <td style="padding:14px 12px; white-space:nowrap;">
        <a href="/admin/users/${u.id}/edit" style="color:#4f46e5; font-size:0.85rem; margin-right:10px;">✏️ 編集</a>
        <a href="/admin/users/${u.id}/toggle"
           style="color:${u.is_active ? '#f97316' : '#22c55e'}; font-size:0.85rem; margin-right:10px;"
           onclick="return confirm('${u.is_active ? '無効にしますか？' : '有効にしますか？'}')">
          ${u.is_active ? '🔒 無効化' : '🔓 有効化'}
        </a>
        <a href="/admin/users/${u.id}/delete"
           style="color:#ef4444; font-size:0.85rem;"
           onclick="return confirm('本当に削除しますか？この操作は元に戻せません。')">
          🗑️ 削除
        </a>
      </td>
    </tr>`).join('');

  res.send(layout('ユーザー管理', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin" class="btn btn-secondary" style="display:inline-block; width:auto;">← 管理者ホームへ</a>
      </p>
      <h2>CSVで一括登録</h2>
      ${result ? `<div class="message message-success">✅ ${result}</div>` : ''}
      <div class="message message-info" style="font-size:0.85rem;">
        CSVファイルの形式：<strong>出勤用ID,パスワード,権限</strong><br>
        権限は <code>staff</code>（職員）または <code>admin</code>（管理者）<br>
        例：<code>1234,user123,staff</code><br>
        ※ 既に登録済みのIDはスキップされます
      </div>
      <form id="uploadForm" method="POST" action="/admin/users/upload" enctype="multipart/form-data" style="margin-bottom:24px;">
        <div class="form-group">
          <label>CSVファイルを選択</label>
          <input type="file" name="csv" accept=".csv" required>
        </div>
        <button type="submit" class="btn btn-primary" style="width:auto; padding:10px 24px;">
          📤 アップロードして登録
        </button>
      </form>
      <h2>登録済みユーザー一覧（${users.length}名）</h2>
      <div class="table-wrapper">
        <table>
          <thead>
            <tr><th>出勤用ID</th><th>権限</th><th>状態</th><th>登録日時</th><th>操作</th></tr>
          </thead>
          <tbody>${rows || '<tr><td colspan="5" style="text-align:center">データがありません</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `, true));
});

// CSVアップロード処理
app.post('/admin/users/upload', requireAdmin, (req, res) => {
  let body = '';
  req.on('data', chunk => { body += chunk.toString('binary'); });
  req.on('end', async () => {
    const boundary = req.headers['content-type'].split('boundary=')[1];
    const parts = body.split('--' + boundary);

    let csvContent = '';
    for (const part of parts) {
      if (part.includes('filename=') && part.includes('.csv')) {
        const lines = part.split('\r\n');
        const dataStart = lines.findIndex(l => l === '') + 1;
        csvContent = lines.slice(dataStart).join('\n').trim();
        break;
      }
    }

    if (!csvContent) {
      return res.redirect('/admin/users?result=' + encodeURIComponent('CSVファイルが読み込めませんでした'));
    }

    const lines = csvContent.split('\n').filter(l => l.trim());
    let success = 0, skipped = 0, errors = 0;

    for (const line of lines) {
      const cols = line.trim().split(',');
      if (cols.length < 2) { errors++; continue; }
      const code = cols[0].trim();
      const pass = cols[1].trim();
      const role = (cols[2] || 'staff').trim();
      if (!code || !pass) { errors++; continue; }
      if (!['staff', 'admin'].includes(role)) { errors++; continue; }

      const exists = await queryOne('SELECT id FROM users WHERE attendance_code = $1', [code]);
      if (exists) { skipped++; continue; }

      try {
        const hash = bcrypt.hashSync(pass, 10);
        await query(
          `INSERT INTO users (attendance_code, login_id, password_hash, role) VALUES ($1, $2, $3, $4)`,
          [code, code, hash, role]
        );
        success++;
      } catch (e) {
        errors++;
      }
    }

    await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'CSV_UPLOAD',
      `成功:${success} スキップ:${skipped} エラー:${errors}`, req.ip);

    res.redirect(`/admin/users?result=${encodeURIComponent(`登録成功:${success}名 / スキップ(重複):${skipped}名 / エラー:${errors}件`)}`);
  });
});

// ユーザー編集画面
app.get('/admin/users/:id/edit', requireAdmin, async (req, res) => {
  const user = await queryOne('SELECT * FROM users WHERE id = $1', [req.params.id]);
  if (!user) return res.redirect('/admin/users');
  const msg = req.query.result ? `<div class="message message-success">✅ ${decodeURIComponent(req.query.result)}</div>` : '';

  res.send(layout('ユーザー編集', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin/users" class="btn btn-secondary" style="display:inline-block; width:auto;">← ユーザー管理へ</a>
      </p>
      <h2>ユーザー編集</h2>
      ${msg}
      <div class="confirm-item">
        <span class="label">出勤用ID</span>
        <span class="value">${user.attendance_code}</span>
      </div>
      <div style="margin-top:24px;">
        <h3 style="font-size:1rem; color:#4f46e5; margin-bottom:12px;">パスワード変更</h3>
        <form method="POST" action="/admin/users/${user.id}/password" style="margin-bottom:20px;">
          <div class="form-group">
            <label>新しいパスワード</label>
            <input type="password" name="password" required minlength="6" placeholder="6文字以上">
          </div>
          <div class="form-group">
            <label>確認用パスワード</label>
            <input type="password" name="password_confirm" required minlength="6" placeholder="もう一度入力">
          </div>
          <button type="submit" class="btn btn-primary" style="width:auto; padding:10px 24px;">🔑 パスワード変更</button>
        </form>
        <h3 style="font-size:1rem; color:#4f46e5; margin-bottom:12px;">権限変更</h3>
        <form method="POST" action="/admin/users/${user.id}/role">
          <div class="form-group">
            <label>権限</label>
            <select name="role">
              <option value="staff" ${user.role === 'staff' ? 'selected' : ''}>職員</option>
              <option value="admin" ${user.role === 'admin' ? 'selected' : ''}>管理者</option>
            </select>
          </div>
          <button type="submit" class="btn btn-primary" style="width:auto; padding:10px 24px;">👤 権限変更</button>
        </form>
      </div>
    </div>
  `, true));
});

// パスワード変更処理
app.post('/admin/users/:id/password', requireAdmin, async (req, res) => {
  const { password, password_confirm } = req.body;
  if (password !== password_confirm) {
    return res.redirect(`/admin/users/${req.params.id}/edit?result=${encodeURIComponent('パスワードが一致しません')}`);
  }
  const hash = bcrypt.hashSync(password, 10);
  await query('UPDATE users SET password_hash = $1 WHERE id = $2', [hash, req.params.id]);
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'PASSWORD_CHANGE', `対象ID:${req.params.id}`, req.ip);
  res.redirect(`/admin/users/${req.params.id}/edit?result=${encodeURIComponent('パスワードを変更しました')}`);
});

// 権限変更処理
app.post('/admin/users/:id/role', requireAdmin, async (req, res) => {
  const { role } = req.body;
  if (!['staff', 'admin'].includes(role)) return res.redirect('/admin/users');
  await query('UPDATE users SET role = $1 WHERE id = $2', [role, req.params.id]);
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'ROLE_CHANGE', `対象ID:${req.params.id} 権限:${role}`, req.ip);
  res.redirect(`/admin/users/${req.params.id}/edit?result=${encodeURIComponent('権限を変更しました')}`);
});

// 有効/無効切り替え
app.get('/admin/users/:id/toggle', requireAdmin, async (req, res) => {
  const user = await queryOne('SELECT * FROM users WHERE id = $1', [req.params.id]);
  if (!user) return res.redirect('/admin/users');
  const newStatus = user.is_active ? 0 : 1;
  await query('UPDATE users SET is_active = $1 WHERE id = $2', [newStatus, req.params.id]);
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'USER_TOGGLE',
    `対象ID:${req.params.id} 状態:${newStatus ? '有効' : '無効'}`, req.ip);
  res.redirect('/admin/users');
});

// ユーザー削除
app.get('/admin/users/:id/delete', requireAdmin, async (req, res) => {
  const user = await queryOne('SELECT * FROM users WHERE id = $1', [req.params.id]);
  if (!user) return res.redirect('/admin/users');
  await query('DELETE FROM users WHERE id = $1', [req.params.id]);
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'USER_DELETE',
    `出勤用ID:${user.attendance_code}を削除`, req.ip);
  res.redirect('/admin/users');
});

// ===== 操作ログ =====
app.get('/admin/logs', requireAdmin, async (req, res) => {
  const logs = await query('SELECT * FROM operation_logs ORDER BY created_at DESC LIMIT 300');
  const roleLabel = { admin: '管理者', staff: '職員' };
  const actionLabel = {
    LOGIN: 'ログイン',
    LOGIN_FAIL: 'ログイン失敗',
    LOGOUT: 'ログアウト',
    ATTENDANCE: '出勤登録',
    CSV_UPLOAD: 'CSVアップロード',
    USER_TOGGLE: 'ユーザー有効/無効変更',
    USER_DELETE: 'ユーザー削除',
    PASSWORD_CHANGE: 'パスワード変更',
    ROLE_CHANGE: '権限変更',
    EXCEL_EXPORT: 'Excel出力',
  };
  const rows = logs.map(l => `
    <tr>
      <td style="padding:14px 12px; white-space:nowrap;">${l.id}</td>
      <td style="padding:14px 12px; white-space:nowrap;">${l.created_at}</td>
      <td style="padding:14px 12px;">${l.attendance_code||'-'}</td>
      <td style="padding:14px 12px;">${roleLabel[l.role] || l.role || '-'}</td>
      <td style="padding:14px 12px;">${actionLabel[l.action] || l.action}</td>
      <td style="padding:14px 12px;">${l.detail||'-'}</td>
      <td style="padding:14px 12px; font-size:0.8rem;">${l.ip_address||'-'}</td>
    </tr>`).join('');
  res.send(layout('操作ログ', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin" class="btn btn-secondary" style="display:inline-block; width:auto;">← 管理者ホームへ</a>
      </p>
      <h2>操作ログ（最新300件）</h2>
      <div class="table-wrapper">
        <table>
          <thead>
            <tr>
              <th style="padding:12px 14px;">ID</th>
              <th style="padding:12px 14px;">日時</th>
              <th style="padding:12px 14px;">出勤用ID</th>
              <th style="padding:12px 14px;">権限</th>
              <th style="padding:12px 14px;">操作</th>
              <th style="padding:12px 14px;">詳細</th>
              <th style="padding:12px 14px;">IP</th>
            </tr>
          </thead>
          <tbody>${rows||'<tr><td colspan="7" style="text-align:center; padding:20px;">ログがありません</td></tr>'}</tbody>
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