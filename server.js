// server.js — QRコード式出勤管理アプリ（セキュリティ強化版）

const express = require('express');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const { Pool } = require('pg');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const path = require('path');
const crypto = require('crypto');

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

// QRコードに入れる文字：推測できないランダムな16文字（URL に使える文字だけ）
function newQrToken() {
  return crypto.randomBytes(12).toString('base64url');
}

// 前の形（loc-001 など、推測しやすい文字）のままのQRコードか
function isGuessableToken(token) {
  return /^loc-\d+$/.test(token || '');
}

// 使えるQRコードの設置場所。なければ null
async function activeLocation(token) {
  if (!token) return null;
  return queryOne('SELECT * FROM qr_locations WHERE qr_token = $1 AND is_active = 1', [token]);
}

// 新しくして使えなくなったQRコードなら、その設置場所の名前と新しくした日時。そうでなければ null
async function retiredLocation(token) {
  if (!token) return null;
  return queryOne(
    `SELECT l.name, h.retired_at FROM qr_token_history h JOIN qr_locations l ON h.qr_location_id = l.id WHERE h.token = $1`,
    [token]
  );
}

// 使えないQRコードを読んだときの画面（古いQRなら記録も残す）
async function invalidQrPage(req, token, user) {
  const old = await retiredLocation(token);
  if (old) {
    await logOperation(req.session.userId, req.session.attendanceCode, req.session.role, 'OLD_QR',
      `設置場所：${old.name}（${formatJst(old.retired_at)} に新しくしたQRコード）`, req.ip);
    return layout('使えないQRコード', `<div class="card"><div class="message message-error">
      このQRコードは使えなくなりました（${escapeHtml(old.name)}：${formatJst(old.retired_at)} に新しくなりました）。<br>
      設置場所に貼ってある新しいQRコードを読み取ってください。</div></div>`, user);
  }
  return layout('エラー', '<div class="card"><div class="message message-error">無効なQRコードです。</div></div>', user);
}

// 今の日本時間：日付 yyyy-mm-dd と時刻 HH:MM:SS（出勤記録の形）
function jstToday() {
  const now = new Date();
  const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000).toISOString();
  return { now, date: jst.slice(0, 10), time: jst.slice(11, 19) };
}

// その職員のその日の最初の出勤記録（設置場所の名前つき）。なければ null
async function firstAttendanceToday(userId, date) {
  return queryOne(
    `SELECT r.attendance_time, l.name AS location_name FROM attendance_records r
     JOIN qr_locations l ON r.qr_location_id = l.id
     WHERE r.user_id = $1 AND r.attendance_date = $2
     ORDER BY r.attendance_time LIMIT 1`,
    [userId, date]
  );
}

// その職員の出勤登録してよい設置場所を、ids で置き換える（空なら、どこでも登録できる状態に戻る）
async function setUserLocations(userId, ids) {
  await query('DELETE FROM user_locations WHERE user_id = $1', [userId]);
  for (const id of ids) {
    await query('INSERT INTO user_locations (user_id, qr_location_id) VALUES ($1, $2) ON CONFLICT DO NOTHING', [userId, id]);
  }
}

// その職員の出勤登録してよい設置場所（名前の順）。決めていなければ空の配列（＝どこでも登録できる）
async function allowedLocations(userId) {
  return query(
    `SELECT l.id, l.name FROM user_locations ul JOIN qr_locations l ON ul.qr_location_id = l.id
     WHERE ul.user_id = $1 ORDER BY l.loc_key`,
    [userId]
  );
}

// 決められた設置場所以外のQRコードを読んだときの画面：記録はせず、読んだことは操作ログに残す
async function wrongPlacePage(req, loc, allowed) {
  const names = allowed.map(a => a.name).join('・');
  await logOperation(req.session.userId, req.session.attendanceCode, req.session.role, 'WRONG_PLACE',
    `読んだ場所：${loc.name}（登録できる場所：${names}）`, req.ip);
  return layout('この場所では登録できません', `
    <div class="card">
      <h2>この場所では出勤登録できません</h2>
      <div class="message message-error">
        出勤登録は、決められた設置場所（${escapeHtml(names)}）のQRコードでだけできます。<br>
        設置場所のQRコードを読み取ってください。
      </div>
      <div class="confirm-item"><span class="label">読み取った場所</span><span class="value">${escapeHtml(loc.name)}</span></div>
      <p style="margin-top:24px; text-align:center;">
        <a href="/logout" class="btn btn-secondary" style="display:inline-block;">ログアウト</a>
      </p>
    </div>
  `, { role: req.session.role });
}

// 同じ日の2回目の出勤登録：記録はせず、最初の登録を知らせる。読んだことは操作ログに残す
async function alreadyRegisteredPage(req, first, loc) {
  const hm = String(first.attendance_time || '').slice(0, 5);
  await logOperation(req.session.userId, req.session.attendanceCode, req.session.role, 'ATTENDANCE_AGAIN',
    `読んだ場所：${loc.name}（最初の登録 ${hm} ${first.location_name}）`, req.ip);
  return layout('出勤登録済み', `
    <div class="card">
      <h2>出勤登録済みです</h2>
      <div class="message message-info">本日はすでに出勤登録が済んでいます。もう一度登録する必要はありません。</div>
      <div class="confirm-item"><span class="label">出勤時刻</span><span class="value">${escapeHtml(hm)}</span></div>
      <div class="confirm-item"><span class="label">QR設置場所</span><span class="value">${escapeHtml(first.location_name)}</span></div>
      <p style="margin-top:24px; text-align:center;">
        <a href="/logout" class="btn btn-secondary" style="display:inline-block;">ログアウト</a>
      </p>
    </div>
  `, { role: req.session.role });
}

// ISO の日時を日本時間の「2026/11/18 9:05」に
function formatJst(iso) {
  const d = new Date(new Date(iso).getTime() + 9 * 60 * 60 * 1000);
  if (isNaN(d)) return iso || '';
  return `${d.getUTCFullYear()}/${d.getUTCMonth() + 1}/${d.getUTCDate()} ${d.getUTCHours()}:${String(d.getUTCMinutes()).padStart(2, '0')}`;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
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
      failed_attempts INTEGER DEFAULT 0,
      locked_until TEXT DEFAULT NULL,
      must_change_password INTEGER DEFAULT 0,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);
  // 既存テーブルへのカラム追加（存在しない場合のみ）
  await query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS failed_attempts INTEGER DEFAULT 0`);
  await query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS locked_until TEXT DEFAULT NULL`);
  await query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS must_change_password INTEGER DEFAULT 0`);

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

  // QRコードの文字（qr_token）は設置場所の目印（loc_key）と分け、管理者が［QRコードを新しくする］で変えられるようにする。
  // 新しくした古い文字は qr_token_history に残し、読まれたら「使えなくなったQRコード」と案内して記録する
  await query(`ALTER TABLE qr_locations ADD COLUMN IF NOT EXISTS loc_key TEXT`);
  await query(`ALTER TABLE qr_locations ADD COLUMN IF NOT EXISTS token_updated_at TEXT`);
  await query(`UPDATE qr_locations SET loc_key = qr_token WHERE loc_key IS NULL`);   // 前の形（qr_token が loc-001 など）の行
  await query(`CREATE UNIQUE INDEX IF NOT EXISTS qr_locations_loc_key ON qr_locations (loc_key)`);
  await query(`
    CREATE TABLE IF NOT EXISTS qr_token_history (
      token TEXT PRIMARY KEY,
      qr_location_id INTEGER NOT NULL,
      retired_at TEXT NOT NULL
    )
  `);

  // 職員ごとの出勤登録してよい設置場所。行がない職員は、今までどおりどの設置場所でも登録できる
  await query(`
    CREATE TABLE IF NOT EXISTS user_locations (
      user_id INTEGER NOT NULL,
      qr_location_id INTEGER NOT NULL,
      PRIMARY KEY (user_id, qr_location_id)
    )
  `);

  // 設置場所（目印と名称）。QRコードの文字はここには書かない（はじめて作るときだけ、推測できないランダムな文字にする）
  const locations = [
    { key: 'loc-001', name: 'K' },
    { key: 'loc-002', name: '1' },
    { key: 'loc-003', name: '2' },
    { key: 'loc-004', name: '3' },
    { key: 'loc-005', name: '4' },
    { key: 'loc-006', name: '5' },
    { key: 'loc-007', name: 'H' },
    { key: 'loc-008', name: 'S-J' },
    { key: 'loc-009', name: 'S-K' },
    { key: 'loc-010', name: 'S-T' },
    { key: 'loc-011', name: 'J' },
  ];
  // 既存の設置場所は名称だけを更新し、QRコードの文字は変えない
  for (const l of locations) {
    await query(
      `INSERT INTO qr_locations (name, loc_key, qr_token, token_updated_at) VALUES ($1, $2, $3, $4)
       ON CONFLICT (loc_key) DO UPDATE SET name = EXCLUDED.name`,
      [l.name, l.key, newQrToken(), new Date().toISOString()]
    );
  }

  // ユーザーが1人もいない場合のみ、初期管理者（9001）を作成する
  const userCount = await queryOne('SELECT COUNT(*)::int AS count FROM users');
  if (userCount.count === 0) {
    const initialPass = process.env.INITIAL_ADMIN_PASSWORD || crypto.randomBytes(9).toString('base64url');
    const hash = bcrypt.hashSync(initialPass, 10);
    await query(
      `INSERT INTO users (attendance_code, login_id, password_hash, role, must_change_password)
       VALUES ($1, $2, $3, $4, 1)`,
      ['9001', '9001', hash, 'admin']
    );
    console.log('初期管理者を作成しました（出勤用ID: 9001）');
    if (!process.env.INITIAL_ADMIN_PASSWORD) {
      console.log(`初期パスワード: ${initialPass}（初回ログイン時に変更してください）`);
    }
  }
  console.log('DB初期化完了');
}

// ===== ミドルウェア =====
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
// セッション秘密鍵は環境変数 SESSION_SECRET から読み込む
const SESSION_SECRET = process.env.SESSION_SECRET || crypto.randomBytes(32).toString('hex');
if (!process.env.SESSION_SECRET) {
  console.warn('警告: SESSION_SECRET が未設定です。一時的な鍵を使うため、再起動のたびに全員ログアウトされます。');
}
app.use(session({
  secret: SESSION_SECRET,
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

// パスワード変更が必要な場合はリダイレクト
function requirePasswordChanged(req, res, next) {
  if (req.session.mustChangePassword && req.path !== '/change-password' && req.path !== '/logout') {
    return res.redirect('/change-password');
  }
  next();
}

async function logOperation(userId, code, role, action, detail, ip) {
  await query(
    `INSERT INTO operation_logs (user_id, attendance_code, role, action, detail, ip_address) VALUES ($1, $2, $3, $4, $5, $6)`,
    [userId || null, code || null, role || null, action, detail || null, ip || null]
  );
}

function layout(title, body, user) {
  const nav = user ? `<header><h1>出勤管理システム</h1><div style="display:flex;gap:12px;align-items:center;">
    ${user.role === 'staff' ? `<a href="/change-password" style="font-size:0.8rem;padding:4px 10px;">🔑 PW変更</a>` : ''}
    <a href="/logout">ログアウト</a>
  </div></header>` : '';
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

// ===== パスワードの決まり =====
// 英字・数字・記号をすべて含む10文字以上。使える文字は半角の英字・数字・記号（空白・全角は不可）
const PASSWORD_MIN = 10;
const PASSWORD_RULE_TEXT = `パスワードは<strong>英字・数字・記号をすべて含む${PASSWORD_MIN}文字以上</strong>にしてください（半角。記号の例：! # $ % & * + - = ? @ _）。`;

// 初期化用の仮のパスワード：見分けにくい文字（0 O 1 l I）を除いた英字・数字と記号で12文字。決まりに必ず合う
function tempPassword() {
  const letters = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz', digits = '23456789', marks = '!#$%&*+-=?@';
  const all = letters + digits + marks;
  const pick = s => s[crypto.randomInt(s.length)];
  const chars = [pick(letters), pick(digits), pick(marks)];
  while (chars.length < 12) chars.push(pick(all));
  for (let i = chars.length - 1; i > 0; i--) { const j = crypto.randomInt(i + 1); [chars[i], chars[j]] = [chars[j], chars[i]]; }
  return chars.join('');
}

// 決まりに合わなければ理由の文、合えば null
function passwordProblem(pw) {
  pw = typeof pw === 'string' ? pw : '';
  if (/[^\x21-\x7E]/.test(pw)) return 'パスワードに使えるのは半角の英字・数字・記号です（空白・全角の文字は使えません）';
  const missing = [];
  if (!/[A-Za-z]/.test(pw)) missing.push('英字');
  if (!/[0-9]/.test(pw)) missing.push('数字');
  if (!/[^A-Za-z0-9]/.test(pw)) missing.push('記号');
  if (pw.length < PASSWORD_MIN || missing.length) {
    return `パスワードは英字・数字・記号をすべて含む${PASSWORD_MIN}文字以上にしてください`
      + (missing.length ? `（${missing.join('・')}が入っていません）` : `（今は${pw.length}文字です）`);
  }
  return null;
}

// ===== ログイン失敗回数チェック =====
const MAX_ATTEMPTS = 5;
const LOCK_MINUTES = 30;

async function checkAndLock(user) {
  // ロック中かチェック
  if (user.locked_until) {
    const lockedUntil = new Date(user.locked_until);
    if (new Date() < lockedUntil) {
      const remaining = Math.ceil((lockedUntil - new Date()) / 60000);
      return { locked: true, remaining };
    } else {
      // ロック解除
      await query('UPDATE users SET failed_attempts = 0, locked_until = NULL WHERE id = $1', [user.id]);
    }
  }
  return { locked: false };
}

async function recordFailedAttempt(user) {
  const newAttempts = (user.failed_attempts || 0) + 1;
  if (newAttempts >= MAX_ATTEMPTS) {
    const lockedUntil = new Date(Date.now() + LOCK_MINUTES * 60 * 1000).toISOString();
    await query('UPDATE users SET failed_attempts = $1, locked_until = $2 WHERE id = $3',
      [newAttempts, lockedUntil, user.id]);
    return { locked: true, remaining: LOCK_MINUTES };
  } else {
    await query('UPDATE users SET failed_attempts = $1 WHERE id = $2', [newAttempts, user.id]);
    return { locked: false, remaining: MAX_ATTEMPTS - newAttempts };
  }
}

async function resetFailedAttempts(userId) {
  await query('UPDATE users SET failed_attempts = 0, locked_until = NULL WHERE id = $1', [userId]);
}

// ===== QRコード読み取り =====
app.get('/scan/:token', async (req, res) => {
  const loc = await activeLocation(req.params.token);
  if (!loc) return res.send(await invalidQrPage(req, req.params.token, false));
  const t = encodeURIComponent(loc.qr_token);
  if (req.session.userId) return res.redirect(`/confirm?token=${t}`);
  res.redirect(`/?token=${t}`);
});

// ===== ログイン画面 =====
app.get('/', async (req, res) => {
  if (req.session.userId) return res.redirect(req.session.role === 'admin' ? '/admin' : '/home');

  const token = req.query.token || '';
  const errorCode = req.query.error || '';

  let errorMsg = '';
  if (errorCode === '1') errorMsg = '<div class="message message-error">IDまたはパスワードが正しくありません</div>';
  if (errorCode === 'locked') {
    const min = req.query.min || LOCK_MINUTES;
    errorMsg = `<div class="message message-error">⚠️ ログイン失敗が続いたため、アカウントをロックしました。${min}分後に再度お試しください。</div>`;
  }
  if (errorCode === 'attempts') {
    const rem = req.query.rem || '';
    errorMsg = `<div class="message message-warning">⚠️ パスワードが違います。あと${rem}回失敗するとアカウントがロックされます。</div>`;
  }

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
    ${errorMsg}
    ${locationInfo}
    <form method="POST" action="/login">
      <input type="hidden" name="token" value="${escapeHtml(token)}">
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
    
  </div>
</div>
</body>
</html>
  `);
});

// ===== ログイン処理 =====
app.post('/login', async (req, res) => {
  const { login_id, password, token } = req.body;
  const user = await queryOne('SELECT * FROM users WHERE login_id = $1 AND is_active = 1', [login_id]);

  if (!user) {
    await logOperation(null, login_id, null, 'LOGIN_FAIL', 'IDが存在しない', req.ip);
    const tokenParam = token ? `&token=${encodeURIComponent(token)}` : '';
    return res.redirect(`/?error=1${tokenParam}`);
  }

  // ロックチェック
  const lockStatus = await checkAndLock(user);
  if (lockStatus.locked) {
    await logOperation(user.id, user.attendance_code, user.role, 'LOGIN_FAIL', 'アカウントロック中', req.ip);
    const tokenParam = token ? `&token=${encodeURIComponent(token)}` : '';
    return res.redirect(`/?error=locked&min=${lockStatus.remaining}${tokenParam}`);
  }

  // パスワードチェック
  if (!bcrypt.compareSync(password, user.password_hash)) {
    const result = await recordFailedAttempt(user);
    await logOperation(user.id, user.attendance_code, user.role, 'LOGIN_FAIL', `パスワード誤り(${user.failed_attempts + 1}回目)`, req.ip);
    const tokenParam = token ? `&token=${encodeURIComponent(token)}` : '';
    if (result.locked) {
      return res.redirect(`/?error=locked&min=${result.remaining}${tokenParam}`);
    }
    return res.redirect(`/?error=attempts&rem=${result.remaining}${tokenParam}`);
  }

  // ログイン成功
  await resetFailedAttempts(user.id);
  req.session.userId = user.id;
  req.session.role = user.role;
  req.session.attendanceCode = user.attendance_code;
  req.session.qrToken = token || null;
  req.session.mustChangePassword = user.must_change_password === 1;
  await logOperation(user.id, user.attendance_code, user.role, 'LOGIN', 'ログイン成功', req.ip);

  // パスワード変更が必要な場合
  if (user.must_change_password === 1) {
    return res.redirect('/change-password');
  }

  if (token) return res.redirect(`/confirm?token=${encodeURIComponent(token)}`);
  res.redirect(user.role === 'admin' ? '/admin' : '/home');
});

// ===== ログアウト =====
app.get('/logout', requireLogin, async (req, res) => {
  await logOperation(req.session.userId, req.session.attendanceCode, req.session.role, 'LOGOUT', null, req.ip);
  req.session.destroy();
  res.redirect('/');
});

// ===== パスワード変更画面（職員・管理者共通） =====
app.get('/change-password', requireLogin, (req, res) => {
  const isForced = req.session.mustChangePassword;
  const msg = req.query.result ? `<div class="message message-success">✅ ${escapeHtml(String(req.query.result))}</div>` : '';
  const err = req.query.error ? `<div class="message message-error">❌ ${escapeHtml(String(req.query.error))}</div>` : '';

  res.send(layout('パスワード変更', `
    <div class="card">
      <h2>🔑 パスワード変更</h2>
      ${isForced ? '<div class="message message-warning">⚠️ 初回ログインのため、パスワードを変更してください。</div>' : ''}
      ${msg}${err}
      <div class="message message-info" style="font-size:0.85rem;">${PASSWORD_RULE_TEXT}</div>
      <form method="POST" action="/change-password">
        <div class="form-group">
          <label>現在のパスワード</label>
          <input type="password" name="current_password" required placeholder="現在のパスワード">
        </div>
        <div class="form-group">
          <label>新しいパスワード（英字・数字・記号を含む${PASSWORD_MIN}文字以上）</label>
          <input type="password" name="new_password" required minlength="${PASSWORD_MIN}" autocomplete="new-password" placeholder="新しいパスワード">
        </div>
        <div class="form-group">
          <label>新しいパスワード（確認）</label>
          <input type="password" name="confirm_password" required minlength="${PASSWORD_MIN}" autocomplete="new-password" placeholder="もう一度入力">
        </div>
        <button type="submit" class="btn btn-primary" style="margin-top:8px;">
          🔑 パスワードを変更する
        </button>
      </form>
      ${!isForced ? `<p style="margin-top:16px;">
        <a href="${req.session.role === 'admin' ? '/admin' : '/home'}" class="btn btn-secondary" style="display:inline-block; width:auto;">← 戻る</a>
      </p>` : ''}
    </div>
  `, { role: req.session.role }));
});

app.post('/change-password', requireLogin, async (req, res) => {
  const { current_password, new_password, confirm_password } = req.body;

  const user = await queryOne('SELECT * FROM users WHERE id = $1', [req.session.userId]);

  if (!bcrypt.compareSync(current_password, user.password_hash)) {
    return res.redirect(`/change-password?error=${encodeURIComponent('現在のパスワードが正しくありません')}`);
  }
  if (new_password !== confirm_password) {
    return res.redirect(`/change-password?error=${encodeURIComponent('新しいパスワードが一致しません')}`);
  }
  const problem = passwordProblem(new_password);
  if (problem) {
    return res.redirect(`/change-password?error=${encodeURIComponent(problem)}`);
  }
  if (bcrypt.compareSync(new_password, user.password_hash)) {
    return res.redirect(`/change-password?error=${encodeURIComponent('今のパスワードと違うパスワードにしてください')}`);
  }

  const hash = bcrypt.hashSync(new_password, 10);
  await query('UPDATE users SET password_hash = $1, must_change_password = 0 WHERE id = $2', [hash, req.session.userId]);
  req.session.mustChangePassword = false;
  await logOperation(req.session.userId, req.session.attendanceCode, req.session.role, 'PASSWORD_CHANGE', '本人によるパスワード変更', req.ip);

  const isForced = user.must_change_password === 1;
  if (isForced) {
    const token = req.session.qrToken;
    if (token) return res.redirect(`/confirm?token=${encodeURIComponent(token)}`);
    return res.redirect(req.session.role === 'admin' ? '/admin' : '/home');
  }

  res.redirect(`/change-password?result=${encodeURIComponent('パスワードを変更しました')}`);
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
app.get('/confirm', requireLogin, requirePasswordChanged, async (req, res) => {
  if (req.session.role === 'admin') return res.redirect('/admin');
  const token = req.query.token || req.session.qrToken;
  const loc = await activeLocation(token);
  if (!loc) return res.send(await invalidQrPage(req, token, { role: req.session.role }));

  // 今日すでに出勤登録していれば、確認画面は出さずに知らせる
  const first = await firstAttendanceToday(req.session.userId, jstToday().date);
  if (first) return res.send(await alreadyRegisteredPage(req, first, loc));

  // 設置場所が決められている職員は、その場所のQRコードでだけ登録できる
  const allowed = await allowedLocations(req.session.userId);
  if (allowed.length > 0 && !allowed.some(a => a.id === loc.id)) return res.send(await wrongPlacePage(req, loc, allowed));

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
      <form method="POST" action="/api/attendance" style="margin-top:24px;"
            onsubmit="this.querySelector('button').disabled = true;">
        <input type="hidden" name="qr_token" value="${loc.qr_token}">
        <button type="submit" class="btn btn-success" style="font-size:1.2rem; padding:18px;">
          ✅ 出勤登録する
        </button>
      </form>
    </div>
  `, { role: req.session.role }));
});

// ===== 職員ホーム =====
app.get('/home', requireLogin, requirePasswordChanged, (req, res) => {
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
  `, { role: req.session.role }));
});

// ===== 出勤登録API =====
app.post('/api/attendance', requireLogin, requirePasswordChanged, async (req, res) => {
  const { qr_token } = req.body;
  const loc = await activeLocation(qr_token);
  // 確認画面を開いたあとでQRコードが新しくなった場合も、古いQRコードとして案内する
  if (!loc) return res.send(await invalidQrPage(req, qr_token, { role: req.session.role }));

  const allowed = await allowedLocations(req.session.userId);
  if (allowed.length > 0 && !allowed.some(a => a.id === loc.id)) return res.send(await wrongPlacePage(req, loc, allowed));

  const { now, date, time } = jstToday();

  // 1日1回：今日の記録がまだないときだけ入れる（ボタンの2度押しや、確認画面を2つ開いていた場合も1件になる）
  const result = await query(
    `INSERT INTO attendance_records
     (user_id, attendance_code, qr_location_id, attendance_date, attendance_time, registered_at, device_info, ip_address)
     SELECT $1::integer, $2::text, $3::integer, $4::text, $5::text, $6::text, $7::text, $8::text
     WHERE NOT EXISTS (SELECT 1 FROM attendance_records WHERE user_id = $1::integer AND attendance_date = $4::text)
     RETURNING id`,
    [req.session.userId, req.session.attendanceCode, loc.id,
     date, time, now.toISOString(),
     req.headers['user-agent'], req.ip]
  );
  if (result.length === 0) {
    const first = await firstAttendanceToday(req.session.userId, date);
    if (first) return res.send(await alreadyRegisteredPage(req, first, loc));
    return res.redirect('/home');
  }

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
  `, { role: req.session.role }));
});

// ===== 管理者ホーム =====
app.get('/admin', requireAdmin, requirePasswordChanged, (req, res) => {
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
  `, { role: 'admin' }));
});

// ===== 管理者：出勤状況一覧 =====
// 出勤状況一覧とExcel出力で共通の絞り込み（期間・設置場所・出勤用ID）
function attendanceFilter(q) {
  const pick = v => (typeof v === 'string' ? v.trim() : '');
  const f = { start: pick(q.start), end: pick(q.end), location: pick(q.location), code: pick(q.code) };
  let where = '';
  const params = [];
  if (f.start)    { params.push(f.start);    where += ` AND r.attendance_date >= $${params.length}`; }
  if (f.end)      { params.push(f.end);      where += ` AND r.attendance_date <= $${params.length}`; }
  if (f.location) { params.push(parseInt(f.location, 10) || 0); where += ` AND r.qr_location_id = $${params.length}`; }
  if (f.code)     { params.push(f.code);     where += ` AND r.attendance_code = $${params.length}`; }
  const qs = new URLSearchParams(Object.entries(f).filter(([, v]) => v)).toString();
  return { ...f, where, params, qs };
}

app.get('/admin/attendance', requireAdmin, async (req, res) => {
  const f = attendanceFilter(req.query);
  const { start, end, location, code } = f;
  const sql = `SELECT r.*, l.name as location_name FROM attendance_records r
             JOIN qr_locations l ON r.qr_location_id = l.id WHERE 1=1${f.where}
             ORDER BY r.registered_at DESC LIMIT 500`;

  const records = await query(sql, f.params);
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
        <a href="/admin/attendance/excel${f.qs ? '?' + escapeHtml(f.qs) : ''}" class="btn btn-primary" style="display:inline-block; width:auto;">📊 Excel出力</a>
        <span style="font-size:0.85rem; color:#555; margin-left:8px;">${f.qs ? '絞り込んだ条件のデータを出力します（件数の上限なし）' : '全期間のデータを出力します'}</span>
      </p>
    </div>
  `, { role: 'admin' }));
});

// ===== Excel出力 =====
app.get('/admin/attendance/excel', requireAdmin, async (req, res) => {
  const f = attendanceFilter(req.query);
  const records = await query(
    `SELECT r.*, l.name as location_name FROM attendance_records r
     JOIN qr_locations l ON r.qr_location_id = l.id WHERE 1=1${f.where}
     ORDER BY r.attendance_date, r.attendance_time`,
    f.params
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
  // ファイル名に期間を入れる（例 attendance_2026-11-01_2026-11-30.xlsx）。期間なしは今までどおり attendance.xlsx
  const span = (f.start || f.end) ? `_${f.start || ''}_${f.end || ''}`.replace(/[^0-9_-]/g, '') : '';
  res.setHeader('Content-Disposition', `attachment; filename="attendance${span}.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
  const cond = [f.start || f.end ? `期間:${f.start || '最初'}〜${f.end || '最後'}` : '全期間',
    f.location ? `場所ID:${f.location}` : '', f.code ? `出勤用ID:${f.code}` : ''].filter(Boolean).join(' ');
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'EXCEL_EXPORT', `${cond} ${records.length}件`, req.ip);
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
    // 端末情報（device_info）は記録は続けるが、画面には出さない（管理者の希望。2026-11-12）
  ];
  res.send(layout('出勤詳細', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin/attendance" class="btn btn-secondary" style="display:inline-block; width:auto;">← 出勤状況一覧へ</a>
      </p>
      <h2>出勤詳細</h2>
      ${items.map(([l,v]) => `<div class="confirm-item"><span class="label">${l}</span><span class="value">${v}</span></div>`).join('')}
    </div>
  `, { role: 'admin' }));
});

// ===== QRコード管理 =====
app.get('/admin/qr', requireAdmin, async (req, res) => {
  const locs = await query('SELECT * FROM qr_locations WHERE is_active = 1 ORDER BY loc_key, id');
  const baseUrl = `https://qr-attendance-app-60gi.onrender.com`;
  const guessable = locs.filter(l => isGuessableToken(l.qr_token)).length;
  const result = req.query.result ? `<div class="message message-success">✅ ${escapeHtml(req.query.result)}</div>` : '';
  const qrItems = await Promise.all(locs.map(async l => {
    const url = `${baseUrl}/scan/${l.qr_token}`;
    const qrDataUrl = await QRCode.toDataURL(url, { width: 200 });
    const mailBody = encodeURIComponent(`出勤登録用QRコードをお送りします。\n\n設置場所：${l.name}\n\nQRコードURL：${url}\n\nスマートフォンのカメラでQRコードを読み取るか、上記URLにアクセスして出勤登録を行ってください。`);
    const mailSubject = encodeURIComponent(`【出勤管理】${l.name} QRコードのご案内`);
    const updated = isGuessableToken(l.qr_token)
      ? '<span style="color:#b91c1c; font-weight:bold;">前の形のまま（推測しやすい）</span>'
      : `最終更新：${l.token_updated_at ? formatJst(l.token_updated_at) : '-'}`;
    return `
      <div class="qr-item">
        <p style="font-weight:bold; font-size:1rem; color:#333; margin-bottom:8px;">${escapeHtml(l.name)}</p>
        <img src="${qrDataUrl}" alt="${escapeHtml(l.name)}" style="width:180px; display:block; margin:0 auto;">
        <small style="color:#888; font-size:0.7rem; word-break:break-all; display:block; margin:8px 0;">${url}</small>
        <small style="display:block; font-size:0.75rem; margin-bottom:6px;">${updated}</small>
        <div style="display:flex; gap:8px; justify-content:center; margin-top:8px; flex-wrap:wrap;">
          <a href="/admin/qr/print/${encodeURIComponent(l.qr_token)}" target="_blank"
             style="background:#4f46e5; color:white; padding:6px 12px; border-radius:8px; font-size:0.8rem; text-decoration:none;">
            🖨️ 印刷
          </a>
          <a href="mailto:?subject=${mailSubject}&body=${mailBody}"
             style="background:#22c55e; color:white; padding:6px 12px; border-radius:8px; font-size:0.8rem; text-decoration:none;">
            📧 メール
          </a>
          <form method="POST" action="/admin/qr/rotate/${l.id}" style="margin:0;"
                onsubmit="return confirm('${escapeHtml(l.name)} のQRコードを新しくします。今貼ってあるQRコードはすぐに使えなくなります。よろしいですか？');">
            <button type="submit" style="background:#f59e0b; color:white; padding:6px 12px; border-radius:8px; font-size:0.8rem; border:none; cursor:pointer;">
              🔄 新しくする
            </button>
          </form>
        </div>
      </div>`;
  }));
  res.send(layout('QRコード管理', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin" class="btn btn-secondary" style="display:inline-block; width:auto;">← 管理者ホームへ</a>
      </p>
      <h2>QRコード一覧（全${locs.length}か所）</h2>
      ${result}
      ${guessable ? `<div class="message message-error" style="margin-bottom:16px; font-size:0.9rem;">
        ⚠️ ${guessable}か所のQRコードが前の形（loc-001 などの推測しやすい文字）のままです。［全部のQRコードを新しくする］を押し、印刷して貼り替えてください。</div>` : ''}
      <div class="message message-warning" style="margin-bottom:16px; font-size:0.85rem;">
        ⚠️ 不正を防ぐため、QRコードは定期的に（例：月に1回）新しくして貼り替えてください。新しくすると、それまでのQRコードはすぐに使えなくなります。
        貼り替えの直前に押し、印刷 → 貼り替えまでを続けて行ってください。古いQRコードを読んだ人には「使えなくなりました」と表示し、操作ログに記録します。<br>
        「メール」はQRコードの文字（URL）をそのまま送るので、送った相手はどこからでも出勤登録できてしまいます。使うときは送り先に注意してください。
      </div>
      <div style="display:flex; gap:12px; flex-wrap:wrap; margin-bottom:16px;">
        <a href="/admin/qr/print-all" target="_blank" class="btn btn-primary" style="display:inline-block; width:auto;">
          🖨️ 全ロケーション一括印刷
        </a>
        <form method="POST" action="/admin/qr/rotate-all" style="margin:0;"
              onsubmit="return confirm('全部（${locs.length}か所）のQRコードを新しくします。今貼ってあるQRコードはすぐに使えなくなります。よろしいですか？');">
          <button type="submit" class="btn" style="width:auto; background:#f59e0b; color:white;">🔄 全部のQRコードを新しくする</button>
        </form>
      </div>
      <div class="qr-grid">${qrItems.join('')}</div>
    </div>
  `, { role: 'admin' }));
});

// QRコードを新しくする：設置場所ごとに新しいランダムな文字にし、古い文字は履歴に残す（読まれたら「使えなくなった」と案内する）
async function rotateQrToken(locId) {
  const loc = await queryOne('SELECT * FROM qr_locations WHERE id = $1 AND is_active = 1', [locId]);
  if (!loc) return null;
  const now = new Date().toISOString();
  await query(
    `INSERT INTO qr_token_history (token, qr_location_id, retired_at) VALUES ($1, $2, $3) ON CONFLICT (token) DO NOTHING`,
    [loc.qr_token, loc.id, now]
  );
  await query('UPDATE qr_locations SET qr_token = $1, token_updated_at = $2 WHERE id = $3', [newQrToken(), now, loc.id]);
  return loc;
}

app.post('/admin/qr/rotate/:id', requireAdmin, async (req, res) => {
  const loc = await rotateQrToken(Number(req.params.id));
  if (!loc) return res.redirect('/admin/qr');
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'QR_ROTATE', `設置場所：${loc.name}`, req.ip);
  res.redirect('/admin/qr?result=' + encodeURIComponent(`${loc.name} のQRコードを新しくしました。印刷して貼り替えてください。`));
});

app.post('/admin/qr/rotate-all', requireAdmin, async (req, res) => {
  const locs = await query('SELECT id FROM qr_locations WHERE is_active = 1');
  for (const l of locs) await rotateQrToken(l.id);
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'QR_ROTATE', `全${locs.length}か所`, req.ip);
  res.redirect('/admin/qr?result=' + encodeURIComponent(`全${locs.length}か所のQRコードを新しくしました。［全ロケーション一括印刷］で印刷して貼り替えてください。`));
});

// ===== QRコード全体印刷 =====
app.get('/admin/qr/print-all', requireAdmin, async (req, res) => {
  const locs = await query('SELECT * FROM qr_locations WHERE is_active = 1 ORDER BY loc_key, id');
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
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:'Hiragino Sans','Meiryo',sans-serif;background:#f0f4f8}
    .no-print{text-align:center;padding:24px;background:#4f46e5;color:white}
    .no-print h1{font-size:1.2rem;margin-bottom:12px}
    .no-print-btns{display:flex;gap:12px;justify-content:center}
    .btn{padding:10px 24px;border-radius:8px;font-size:0.95rem;cursor:pointer;border:none;font-weight:bold}
    .btn-print{background:white;color:#4f46e5}
    .btn-close{background:rgba(255,255,255,0.2);color:white}
    .pages-wrapper{padding:24px;display:flex;flex-direction:column;gap:24px;align-items:center}
    .print-page{background:white;width:210mm;min-height:297mm;display:flex;align-items:center;justify-content:center;box-shadow:0 2px 12px rgba(0,0,0,0.1)}
    .print-card{text-align:center;padding:48px 40px;width:100%}
    .title{font-size:1.6rem;font-weight:bold;color:#1e1b4b;margin-bottom:10px}
    .subtitle{font-size:1.3rem;font-weight:bold;color:#4f46e5;margin-bottom:32px;padding:10px 24px;background:#e0e7ff;border-radius:8px;display:inline-block}
    .qr-img{width:280px;height:280px;margin:0 auto 32px;display:block}
    .instruction{font-size:1rem;color:#444;line-height:1.8;background:#f0f4ff;border-radius:10px;padding:16px 24px;display:inline-block}
    @media print{
      .no-print{display:none}
      body{background:white}
      .pages-wrapper{padding:0;gap:0}
      .print-page{width:100%;min-height:100vh;page-break-after:always;box-shadow:none;margin:0}
      .print-page:last-child{page-break-after:avoid}
    }
  </style>
</head>
<body>
  <div class="no-print">
    <h1>🖨️ QRコード一括印刷（全${locs.length}か所）</h1>
    <div class="no-print-btns">
      <button class="btn btn-print" onclick="window.print()">🖨️ 印刷する</button>
      <button class="btn btn-close" onclick="window.close()">✕ 閉じる</button>
    </div>
  </div>
  <div class="pages-wrapper">${pages.join('')}</div>
</body>
</html>`);
});

// ===== QRコード個別印刷 =====
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
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:'Hiragino Sans','Meiryo',sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;background:white}
    .print-card{text-align:center;padding:40px;border:3px solid #4f46e5;border-radius:16px;width:340px}
    .title{font-size:1.4rem;font-weight:bold;color:#1e1b4b;margin-bottom:8px}
    .subtitle{font-size:1.1rem;font-weight:bold;color:#4f46e5;margin-bottom:24px;padding:8px 16px;background:#e0e7ff;border-radius:8px;display:inline-block}
    .qr-img{width:260px;height:260px;margin:0 auto 20px;display:block}
    .instruction{font-size:0.85rem;color:#444;line-height:1.6;background:#f0f4ff;border-radius:8px;padding:12px;margin-bottom:20px}
    .no-print{margin-top:20px;display:flex;gap:10px;justify-content:center}
    .btn{padding:10px 20px;border-radius:8px;font-size:0.9rem;cursor:pointer;border:none;font-weight:bold}
    .btn-print{background:#4f46e5;color:white}
    .btn-close{background:#e5e7eb;color:#374151}
    @media print{.no-print{display:none}}
  </style>
</head>
<body>
  <div class="print-card">
    <p class="title">出勤登録QRコード</p>
    <p class="subtitle">${loc.name}</p>
    <img src="${qrDataUrl}" alt="${loc.name}" class="qr-img">
    <div class="instruction">📱 スマートフォンのカメラで<br>このQRコードを読み取って<br>出勤登録を行ってください</div>
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
  const users = await query(`SELECT id, attendance_code, role, is_active, failed_attempts, locked_until, must_change_password, created_at FROM users ORDER BY is_active DESC, attendance_code`);
  // 職員ごとの設置場所（名前を「・」でつないだもの）
  const placeRows = await query(
    `SELECT ul.user_id, l.name FROM user_locations ul JOIN qr_locations l ON ul.qr_location_id = l.id ORDER BY l.loc_key`);
  const places = {};
  for (const p of placeRows) (places[p.user_id] = places[p.user_id] || []).push(p.name);
  const now = new Date();
  const rows = users.map(u => {
    const isLocked = u.locked_until && new Date(u.locked_until) > now;
    return `
    <tr style="${!u.is_active ? 'opacity:0.5; background:#f9fafb;' : ''}">
      <td style="padding:14px 12px;">${u.attendance_code}</td>
      <td style="padding:14px 12px;">${u.role === 'admin' ? '管理者' : '職員'}</td>
      <td style="padding:14px 12px;">
        <span class="badge ${u.is_active ? 'badge-normal' : 'badge-caution'}">${u.is_active ? '有効' : '無効'}</span>
        ${isLocked ? '<span class="badge badge-caution" style="margin-left:4px;">🔒 ロック中</span>' : ''}
        ${u.must_change_password ? '<span class="badge" style="background:#fef3c7;color:#92400e;margin-left:4px;">PW要変更</span>' : ''}
      </td>
      <td style="padding:14px 12px;">${places[u.id] ? escapeHtml(places[u.id].join('・')) : '<span style="color:#888;">どこでも</span>'}</td>
      <td style="padding:14px 12px;">${u.failed_attempts || 0}回</td>
      <td style="padding:14px 12px; white-space:nowrap;">
        <a href="/admin/users/${u.id}/edit" style="color:#4f46e5; font-size:0.85rem; margin-right:10px;">✏️ 編集</a>
        ${isLocked ? `<a href="/admin/users/${u.id}/unlock" style="color:#22c55e; font-size:0.85rem; margin-right:10px;">🔓 ロック解除</a>` : ''}
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
    </tr>`;
  }).join('');

  res.send(layout('ユーザー管理', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin" class="btn btn-secondary" style="display:inline-block; width:auto;">← 管理者ホームへ</a>
      </p>
      <h2>CSVで一括登録</h2>
      ${result ? `<div class="message message-success">✅ ${result}</div>` : ''}
      <div class="message message-info" style="font-size:0.85rem;">
        CSVファイルの形式：<strong>出勤用ID,パスワード,権限,設置場所</strong>（設置場所は省略できます）<br>
        権限は <code>staff</code>（職員）または <code>admin</code>（管理者）<br>
        設置場所は、出勤登録してよい場所の名前（QRコード管理の名称。例 <code>K</code>）。2か所以上は <code>K;1</code> のように「;」で区切ります。空欄の人はどこでも登録できます。<br>
        ※ 初回登録時は必ずパスワード変更を促す設定になります（変更後のパスワードは英字・数字・記号をすべて含む${PASSWORD_MIN}文字以上）<br>
        ※ 登録済みの出勤用IDの行は、設置場所が書いてあればその設置場所だけを置き換えます（パスワード・権限は変えません）
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
            <tr><th>出勤用ID</th><th>権限</th><th>状態</th><th>設置場所</th><th>失敗回数</th><th>操作</th></tr>
          </thead>
          <tbody>${rows || '<tr><td colspan="6" style="text-align:center">データがありません</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `, { role: 'admin' }));
});

// CSVアップロード
app.post('/admin/users/upload', requireAdmin, (req, res) => {
  let body = '';
  req.on('data', chunk => { body += chunk.toString('binary'); });
  req.on('end', async () => {
    const boundary = req.headers['content-type'].split('boundary=')[1];
    const parts = body.split('--' + boundary);
    let csvContent = '';
    for (const part of parts) {
      if (part.includes('filename=') && part.includes('.csv')) {
        // ファイルの中身は、部分の見出し（Content-Disposition など）の後の空行から
        const headEnd = part.indexOf('\r\n\r\n');
        csvContent = (headEnd >= 0 ? part.slice(headEnd + 4) : '').replace(/\r\n/g, '\n').replace(/^﻿|^\xEF\xBB\xBF/, '').trim();
        break;
      }
    }
    if (!csvContent) return res.redirect('/admin/users?result=' + encodeURIComponent('CSVファイルが読み込めませんでした'));

    const lines = csvContent.split('\n').filter(l => l.trim());
    // 設置場所の名前 → id（名前は QRコード管理の名称。前後の空白は無視）
    const locByName = {};
    for (const l of await query('SELECT id, name FROM qr_locations WHERE is_active = 1')) locByName[l.name.trim()] = l.id;
    let success = 0, skipped = 0, placed = 0, errors = 0;
    const badPlaces = new Set();
    for (const line of lines) {
      const cols = line.trim().split(',');
      if (cols.length < 2) { errors++; continue; }
      const code = cols[0].trim();
      const pass = cols[1].trim();
      const role = (cols[2] || 'staff').trim();
      if (!code || !pass || !['staff', 'admin'].includes(role)) { errors++; continue; }
      // 4列目：設置場所（「;」区切り）。書いてない行は設置場所を変えない
      const placeNames = (cols[3] || '').split(/[;；]/).map(s => s.trim()).filter(Boolean);
      const unknown = placeNames.filter(n => !(n in locByName));
      if (unknown.length) { unknown.forEach(n => badPlaces.add(n)); errors++; continue; }
      const placeIds = [...new Set(placeNames.map(n => locByName[n]))];
      const exists = await queryOne('SELECT id FROM users WHERE attendance_code = $1', [code]);
      if (exists) {
        if (placeIds.length) { await setUserLocations(exists.id, placeIds); placed++; }
        else skipped++;
        continue;
      }
      try {
        const hash = bcrypt.hashSync(pass, 10);
        const created = await query(
          `INSERT INTO users (attendance_code, login_id, password_hash, role, must_change_password) VALUES ($1, $2, $3, $4, 1) RETURNING id`,
          [code, code, hash, role]
        );
        if (placeIds.length) await setUserLocations(created[0].id, placeIds);
        success++;
      } catch (e) { errors++; }
    }
    const summary = `登録成功:${success}名 / 設置場所を置き換え:${placed}名 / スキップ:${skipped}名 / エラー:${errors}件`
      + (badPlaces.size ? `（見つからない設置場所：${[...badPlaces].join('・')}）` : '');
    await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'CSV_UPLOAD', summary, req.ip);
    res.redirect(`/admin/users?result=${encodeURIComponent(summary)}`);
  });
});

// ユーザー編集
app.get('/admin/users/:id/edit', requireAdmin, async (req, res) => {
  const user = await queryOne('SELECT * FROM users WHERE id = $1', [req.params.id]);
  if (!user) return res.redirect('/admin/users');
  const msg = req.query.result ? `<div class="message message-success">✅ ${decodeURIComponent(req.query.result)}</div>` : '';
  const allLocs = await query('SELECT id, name FROM qr_locations WHERE is_active = 1 ORDER BY loc_key');
  const mine = new Set((await allowedLocations(user.id)).map(l => l.id));
  const placeChecks = allLocs.map(l => `
          <label style="display:inline-flex; align-items:center; gap:4px; margin:0 14px 8px 0; font-weight:normal;">
            <input type="checkbox" name="loc" value="${l.id}" ${mine.has(l.id) ? 'checked' : ''} style="width:auto;"> ${escapeHtml(l.name)}
          </label>`).join('');

  res.send(layout('ユーザー編集', `
    <div class="card">
      <p style="margin-bottom:16px;">
        <a href="/admin/users" class="btn btn-secondary" style="display:inline-block; width:auto;">← ユーザー管理へ</a>
      </p>
      <h2>ユーザー編集</h2>
      ${msg}
      <div class="confirm-item"><span class="label">出勤用ID</span><span class="value">${user.attendance_code}</span></div>
      <div style="margin-top:24px;">
        
        <h3 style="font-size:1rem; color:#4f46e5; margin-bottom:12px;">パスワード初期化</h3>
        <div class="message message-info" style="font-size:0.85rem;">初期化すると、推測されにくい仮のパスワードに変わり、画面に1回だけ表示されます。本人に伝えてください。<br>ロックも解除されます。次回ログイン時に本人がパスワードを変更する必要があります。</div>
        <form method="POST" action="/admin/users/${user.id}/reset-password" style="margin-bottom:20px;">
          <button type="submit" class="btn btn-danger" style="width:auto; padding:10px 24px;" onclick="return confirm('パスワードを初期化しますか？')">🔄 パスワードを初期化する</button>
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
        <h3 style="font-size:1rem; color:#4f46e5; margin:24px 0 12px;">出勤登録できる設置場所</h3>
        <div class="message message-info" style="font-size:0.85rem;">チェックした設置場所のQRコードでだけ出勤登録できます。1つもチェックしなければ、どこでも登録できます。</div>
        <form method="POST" action="/admin/users/${user.id}/locations">
          <div class="form-group">${placeChecks}</div>
          <button type="submit" class="btn btn-primary" style="width:auto; padding:10px 24px;">📍 設置場所を保存</button>
        </form>
      </div>
    </div>
  `, { role: 'admin' }));
});

// パスワード初期化：仮のパスワード（英字・数字・記号を含む12文字）にし、画面に1回だけ表示する。
// 仮のパスワードはアドレスに載せないよう、転送せずにこの応答で表示する
app.post('/admin/users/:id/reset-password', requireAdmin, async (req, res) => {
  const user = await queryOne('SELECT id, attendance_code FROM users WHERE id = $1', [req.params.id]);
  if (!user) return res.redirect('/admin/users');
  const temp = tempPassword();
  await query(
    'UPDATE users SET password_hash = $1, must_change_password = 1, failed_attempts = 0, locked_until = NULL WHERE id = $2',
    [bcrypt.hashSync(temp, 10), user.id]
  );
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'PASSWORD_RESET',
    `出勤用ID:${user.attendance_code}`, req.ip);
  res.set('Cache-Control', 'no-store');
  res.send(layout('パスワード初期化', `
    <div class="card">
      <h2>パスワードを初期化しました</h2>
      <div class="message message-success">✅ 出勤用ID ${escapeHtml(user.attendance_code)} のパスワードを、次の仮のパスワードにしました。</div>
      <div class="confirm-item"><span class="label">仮のパスワード</span>
        <span class="value" style="font-family:monospace; font-size:1.3rem; letter-spacing:1px;">${escapeHtml(temp)}</span></div>
      <div class="message message-warning" style="font-size:0.85rem; margin-top:16px;">
        この画面を閉じると、仮のパスワードはもう表示できません。控えてから本人に伝えてください。<br>
        本人が次にログインすると、パスワード変更の画面になります。
      </div>
      <p style="margin-top:20px;">
        <a href="/admin/users/${user.id}/edit" class="btn btn-secondary" style="display:inline-block; width:auto;">← ユーザー編集へ</a>
      </p>
    </div>
  `, { role: 'admin' }));
});

// パスワード変更（管理者による）
app.post('/admin/users/:id/password', requireAdmin, async (req, res) => {
  const { password, password_confirm, must_change } = req.body;
  if (password !== password_confirm) {
    return res.redirect(`/admin/users/${req.params.id}/edit?result=${encodeURIComponent('パスワードが一致しません')}`);
  }
  const problem = passwordProblem(password);
  if (problem) return res.redirect(`/admin/users/${req.params.id}/edit?result=${encodeURIComponent(problem)}`);
  const hash = bcrypt.hashSync(password, 10);
  const mustChange = must_change ? 1 : 0;
  await query('UPDATE users SET password_hash = $1, must_change_password = $2, failed_attempts = 0, locked_until = NULL WHERE id = $3',
    [hash, mustChange, req.params.id]);
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'PASSWORD_CHANGE', `対象ID:${req.params.id}`, req.ip);
  res.redirect(`/admin/users/${req.params.id}/edit?result=${encodeURIComponent('パスワードを変更しました')}`);
});

// 権限変更
app.post('/admin/users/:id/role', requireAdmin, async (req, res) => {
  const { role } = req.body;
  if (!['staff', 'admin'].includes(role)) return res.redirect('/admin/users');
  await query('UPDATE users SET role = $1 WHERE id = $2', [role, req.params.id]);
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'ROLE_CHANGE', `対象ID:${req.params.id} 権限:${role}`, req.ip);
  res.redirect(`/admin/users/${req.params.id}/edit?result=${encodeURIComponent('権限を変更しました')}`);
});

// 出勤登録できる設置場所（チェックなし＝どこでも）
app.post('/admin/users/:id/locations', requireAdmin, async (req, res) => {
  const user = await queryOne('SELECT id, attendance_code FROM users WHERE id = $1', [req.params.id]);
  if (!user) return res.redirect('/admin/users');
  const picked = [].concat(req.body.loc || []).map(v => parseInt(v, 10)).filter(Number.isInteger);
  const valid = (await query('SELECT id, name FROM qr_locations WHERE is_active = 1 ORDER BY loc_key'))
    .filter(l => picked.includes(l.id));
  await setUserLocations(user.id, valid.map(l => l.id));
  const names = valid.length ? valid.map(l => l.name).join('・') : 'どこでも';
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'USER_LOCATIONS',
    `出勤用ID:${user.attendance_code} 設置場所:${names}`, req.ip);
  res.redirect(`/admin/users/${user.id}/edit?result=${encodeURIComponent(`出勤登録できる設置場所を「${names}」にしました`)}`);
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

// ロック解除
app.get('/admin/users/:id/unlock', requireAdmin, async (req, res) => {
  await query('UPDATE users SET failed_attempts = 0, locked_until = NULL WHERE id = $1', [req.params.id]);
  await logOperation(req.session.userId, req.session.attendanceCode, 'admin', 'USER_UNLOCK', `対象ID:${req.params.id}`, req.ip);
  res.redirect('/admin/users');
});

// ユーザー削除
app.get('/admin/users/:id/delete', requireAdmin, async (req, res) => {
  const user = await queryOne('SELECT * FROM users WHERE id = $1', [req.params.id]);
  if (!user) return res.redirect('/admin/users');
  await query('DELETE FROM user_locations WHERE user_id = $1', [req.params.id]);
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
    LOGIN: 'ログイン', LOGIN_FAIL: 'ログイン失敗', LOGOUT: 'ログアウト',
    ATTENDANCE: '出勤登録', ATTENDANCE_AGAIN: '同じ日の2回目の読み取り（記録せず）',
    WRONG_PLACE: '決められた場所以外での読み取り（記録せず）', USER_LOCATIONS: '設置場所の変更', CSV_UPLOAD: 'CSVアップロード',
    USER_TOGGLE: 'ユーザー有効/無効変更', USER_DELETE: 'ユーザー削除',
    USER_UNLOCK: 'ロック解除', QR_ROTATE: 'QRコードを新しくした', OLD_QR: '使えなくなったQRコードを読んだ',
    PASSWORD_CHANGE: 'パスワード変更', PASSWORD_RESET: 'パスワード初期化', ROLE_CHANGE: '権限変更', EXCEL_EXPORT: 'Excel出力',
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
  `, { role: 'admin' }));
});

// ===== サーバー起動 =====
initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`QR Attendance App started.`);
    console.log(`http://localhost:${PORT}`);
  });
}).catch(err => {
  console.error('DB初期化エラー:', err);
});
