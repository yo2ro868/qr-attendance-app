
// ---- ここから下は4つのスクリプトで共通（office_scripts/共通.ts。make_package.py が各スクリプトの後ろに付ける） ----

const DAY_MIN = 465;   // 1日＝7時間45分＝465分
const KIND_ANNUAL = "年次休暇";
const KIND_SICK = "病気休暇";
const KIND_NURSING = "看護休暇（子の看護等）";
const KIND_SPECIAL = "特別休暇（看護休暇以外）";
const KIND_CARE_TIME = "介護時間";
const KIND_CHILDCARE_TIME = "育児時間";
const KIND_EXEMPT = "職務専念義務の免除";
const ST_WAIT = "承認待ち";
const ST_OK = "承認";
const ST_NG = "却下";

type Cell = string | number | boolean;

// 表（テーブル）を「見出し → 列番号」と全行の値で読む。空の行も含む（行番号を表の行と合わせるため）
interface Tbl {
  table: ExcelScript.Table;
  cols: Map<string, number>;
  rows: Cell[][];
}

function readTable(workbook: ExcelScript.Workbook, name: string): Tbl {
  const table = workbook.getTable(name);
  if (!table) {
    throw new Error(`表「${name}」が見つかりません`);
  }
  const cols = new Map<string, number>();
  table.getHeaderRowRange().getValues()[0].forEach((h, i) => cols.set(String(h).trim(), i));
  return { table, cols, rows: table.getRangeBetweenHeaderAndTotal().getValues() };
}

function col(t: Tbl, name: string): number {
  const i = t.cols.get(name);
  if (i === undefined) {
    throw new Error(`表の見出し「${name}」が見つかりません`);
  }
  return i;
}

function val(t: Tbl, row: Cell[], name: string): Cell {
  return row[col(t, name)];
}

function str(v: Cell): string {
  return v === undefined || v === null ? "" : String(v).trim();
}

function num(v: Cell): number {
  const n = Number(str(v));
  return isFinite(n) ? n : 0;
}

// 出勤用ID の照合用（前後の空白・先頭の0・文字列と数値の違いを吸収）
function normId(v: Cell): string {
  const s = str(v);
  return s !== "" && isFinite(Number(s)) ? String(Number(s)) : s;
}

function normEmail(v: Cell): string {
  return str(v).toLowerCase();
}

function isBlankRow(row: Cell[]): boolean {
  return row.every(v => str(v) === "");
}

// Forms の日付（2026-11-05 など）を Excel の日付シリアル値にする。読めなければ 0
function toSerial(text: Cell): number {
  if (typeof text === "number") {
    return text;
  }
  const s = str(text);
  let y = 0, m = 0, d = 0;
  let r = s.match(/^(\d{4})[-\/年](\d{1,2})[-\/月](\d{1,2})/);
  if (r) {
    y = Number(r[1]); m = Number(r[2]); d = Number(r[3]);
  } else {
    r = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);   // 11/5/2026（月/日/年）
    if (!r) {
      return 0;
    }
    m = Number(r[1]); d = Number(r[2]); y = Number(r[3]);
  }
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}

function serialDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
}

function dateText(serial: number): string {
  if (!serial) {
    return "";
  }
  const d = serialDate(serial);
  return `${d.getUTCFullYear()}/${d.getUTCMonth() + 1}/${d.getUTCDate()}`;
}

// 日本時間の「2026/10/29 13:05」。iso を省略すると今の時刻
function jstText(iso?: string): string {
  const t = iso ? new Date(iso) : new Date();
  if (isNaN(t.getTime())) {
    return str(iso);
  }
  const j = new Date(t.getTime() + 9 * 3600000);
  const mm = String(j.getUTCMinutes()).padStart(2, "0");
  return `${j.getUTCFullYear()}/${j.getUTCMonth() + 1}/${j.getUTCDate()} ${j.getUTCHours()}:${mm}`;
}

function jstYear(): number {
  return new Date(Date.now() + 9 * 3600000).getUTCFullYear();
}

// 分を「3時間30分」に
function hmText(min: number): string {
  const sign = min < 0 ? "−" : "";
  const a = Math.abs(Math.round(min));
  const h = Math.floor(a / 60), m = a % 60;
  if (a === 0) {
    return "0分";
  }
  return sign + (h > 0 ? `${h}時間` : "") + (m > 0 ? `${m}分` : "");
}

// 分を「15日3時間30分」に（1日＝7時間45分）
function daysText(min: number): string {
  const sign = min < 0 ? "−" : "";
  const a = Math.abs(Math.round(min));
  const d = Math.floor(a / DAY_MIN), rest = a % DAY_MIN;
  return sign + `${d}日` + (rest > 0 ? hmText(rest) : "");
}

// 申請1行の取得時間（分）。1日単位は1日＝7時間45分で換算
function rowMinutes(t: Tbl, row: Cell[]): number {
  return Math.round(num(val(t, row, "日数")) * DAY_MIN + num(val(t, row, "時間")) * 1440);
}

function rowYear(t: Tbl, row: Cell[]): number {
  const s = toSerial(val(t, row, "開始日"));
  return s ? serialDate(s).getUTCFullYear() : 0;
}

// 申請1行の内容を1行の文にする
function rowSummary(t: Tbl, row: Cell[]): string {
  const detail = str(val(t, row, "種別・内容"));
  const start = toSerial(val(t, row, "開始日"));
  const end = toSerial(val(t, row, "終了日"));
  const dates = !end || end === start ? dateText(start) : `${dateText(start)}〜${dateText(end)}`;
  const time = num(val(t, row, "時間"));
  const clock = str(val(t, row, "時刻"));
  const amount = time > 0
    ? hmText(time * 1440) + (clock ? `（${clock}）` : "")
    : `${num(val(t, row, "日数"))}日`;
  return `${str(val(t, row, "休暇の種類"))}${detail ? `（${detail}）` : ""}　${dates}　${amount}`;
}

interface Staff {
  email: string;
  id: string;
  name: string;
  notes: string;           // 通知先（Notes）
  approver: string;        // 承認者メールアドレス（Power Automate の承認の担当者）
  approverNotes: string;   // 承認者の通知先（Notes）
}

function findStaff(workbook: ExcelScript.Workbook, email: string): Staff | undefined {
  const t = readTable(workbook, "職員");
  const row = t.rows.find(r => normEmail(val(t, r, "メールアドレス")) === normEmail(email));
  if (!row) {
    return undefined;
  }
  return {
    email: str(val(t, row, "メールアドレス")),
    id: normId(val(t, row, "出勤用ID")),
    name: str(val(t, row, "氏名")),
    notes: str(val(t, row, "通知先（Notes）")),
    approver: str(val(t, row, "承認者メールアドレス")),
    approverNotes: str(val(t, row, "承認者の通知先（Notes）")),
  };
}

// その年の年次休暇の持ち日数・使用（承認済みと導入前）・承認待ち（分）
interface Balance {
  registered: boolean;
  total: number;
  used: number;
  pending: number;
}

function annualBalance(workbook: ExcelScript.Workbook, req: Tbl, id: string, year: number): Balance {
  const h = readTable(workbook, "持ち日数");
  const own = h.rows.find(r => normId(val(h, r, "出勤用ID")) === id && num(val(h, r, "年")) === year);
  const b: Balance = { registered: !!own, total: 0, used: 0, pending: 0 };
  if (own) {
    b.total = num(val(h, own, "合計（日）")) * DAY_MIN;
    b.used = Math.round(num(val(h, own, "導入前の使用（日）")) * DAY_MIN + num(val(h, own, "導入前の使用（時間）")) * 1440);
  }
  for (const r of req.rows) {
    if (normId(val(req, r, "出勤用ID")) !== id || str(val(req, r, "休暇の種類")) !== KIND_ANNUAL || rowYear(req, r) !== year) {
      continue;
    }
    const status = str(val(req, r, "状況"));
    if (status === ST_OK) {
      b.used += rowMinutes(req, r);
    } else if (status === ST_WAIT) {
      b.pending += rowMinutes(req, r);
    }
  }
  return b;
}

// その年の、ある種類の承認済みの取得時間（分）
function approvedMinutes(req: Tbl, id: string, year: number, kind: string): number {
  let sum = 0;
  for (const r of req.rows) {
    if (normId(val(req, r, "出勤用ID")) === id && str(val(req, r, "休暇の種類")) === kind
      && rowYear(req, r) === year && str(val(req, r, "状況")) === ST_OK) {
      sum += rowMinutes(req, r);
    }
  }
  return sum;
}

// 残日数などの説明の行
function balanceLines(workbook: ExcelScript.Workbook, req: Tbl, id: string, year: number): string[] {
  const b = annualBalance(workbook, req, id, year);
  const lines: string[] = [];
  if (!b.registered) {
    lines.push(`年次休暇（${year}年）：持ち日数が台帳に登録されていません`);
  } else {
    lines.push(`年次休暇（${year}年）：持ち日数 ${daysText(b.total)}／使用済み ${daysText(b.used)}／残り ${daysText(b.total - b.used)}`);
    if (b.pending > 0) {
      lines.push(`　承認待ち ${daysText(b.pending)}（すべて承認されると残り ${daysText(b.total - b.used - b.pending)}）`);
    }
  }
  lines.push(`病気休暇（${year}年・承認済み）：${hmText(approvedMinutes(req, id, year, KIND_SICK))}`);
  lines.push(`看護休暇（${year}年・承認済み）：${hmText(approvedMinutes(req, id, year, KIND_NURSING))}`);
  return lines;
}

// Notes で送るお知らせを「通知」表に1行書く（組織のパソコンの send_notices.ps1 が送る）。
// 同じ通知番号がすでにあれば書かない（フローの再実行で二重に送らないように）。
// 宛先が空欄なら、設定シートの B1（管理者の通知先）に「宛先未登録」として送る。
function addNotice(workbook: ExcelScript.Workbook, id: string, to: string, subject: string, lines: string[]) {
  const t = readTable(workbook, "通知");
  if (t.rows.some(r => str(val(t, r, "通知番号")) === id)) {
    return;
  }
  if (!to) {
    to = str(workbook.getWorksheet("設定").getRange("B1").getValue());
    subject = `【宛先未登録】${subject}`;
    lines = ["（職員表に通知先（Notes）が登録されていないため、管理者に送っています）", "", ...lines];
  }
  const values: Cell[] = new Array<Cell>(t.cols.size).fill("");
  values[col(t, "通知番号")] = id;
  values[col(t, "作成日時")] = jstText();
  values[col(t, "宛先")] = to;
  values[col(t, "件名")] = subject;
  values[col(t, "本文")] = lines.join("\n");
  const blank = t.rows.findIndex(isBlankRow);
  if (blank >= 0) {
    t.table.getRangeBetweenHeaderAndTotal().getRow(blank).setValues([values]);
  } else {
    t.table.addRow(-1, values);
  }
}

// 承認の詳細（Markdown）は行末に空白2つで改行になる
function toMarkdown(lines: string[]): string {
  return lines.join("  \n");
}
