/**
 * 申請を登録
 * Forms の休暇申請1件を「申請」表に追加し、承認依頼の宛先と本文を返す。
 * Power Automate の「スクリプトの実行」から呼ぶ。引数は Forms の回答をそのまま渡す（未回答は空欄）。
 * 同じ申請番号がすでにあれば追加しない（フローの再実行で二重に入らないように）。
 */
interface Result {
  found: boolean;          // 職員表にメールアドレスがあったか
  staffId: string;
  name: string;
  approverEmail: string;
  title: string;           // 承認依頼の件名
  details: string;         // 承認依頼の本文（Markdown）
  detailsHtml: string;     // 同じ内容をメール用に
  warnings: string;        // 確認事項（なければ空欄）
}

function main(workbook: ExcelScript.Workbook,
  responseId: string, submittedAt: string, email: string,
  kind: string, specialKind: string, exemptDetail: string, unit: string,
  dayStart: string, dayEnd: string, days: string,
  timeDate: string, hours: string, minutes: string, timeRange: string,
  reason: string): Result {
  const staff = findStaff(workbook, email);
  kind = str(kind);
  const byTime = str(unit).indexOf("時間") >= 0;
  const start = toSerial(byTime ? timeDate : dayStart);
  const end = byTime ? start : (toSerial(dayEnd) || start);
  const nDays = byTime ? 0 : num(days);
  const mins = byTime ? Math.round(num(hours) * 60 + num(minutes)) : 0;
  const detail = kind === KIND_SPECIAL ? str(specialKind) : kind === KIND_EXEMPT ? str(exemptDetail) : "";

  const warn: string[] = [];
  if (!staff) {
    warn.push(`職員表にメールアドレス ${email} がありません`);
  }
  if (!start) {
    warn.push("日付が読み取れません");
  } else if (end < start) {
    warn.push("終了日が開始日より前です");
  }
  if (!byTime && nDays <= 0) {
    warn.push("日数が入っていません");
  }
  if (!byTime && start && end >= start && nDays > end - start + 1) {
    warn.push("日数が期間の日数より多くなっています");
  }
  if (!byTime && start && end > start && [KIND_ANNUAL, KIND_SICK, KIND_NURSING].indexOf(kind) >= 0) {
    const a = serialDate(start), b = serialDate(end);
    if (a.getUTCFullYear() !== b.getUTCFullYear() || a.getUTCMonth() !== b.getUTCMonth()) {
      warn.push("月をまたいでいます。集計では開始日の月に入るため、月ごとに分けて申請してください");
    }
  }
  if (byTime && mins <= 0) {
    warn.push("時間が入っていません");
  }
  if (byTime && kind === KIND_ANNUAL && mins % 15 !== 0) {
    warn.push("年次休暇の時間は15分単位です");
  }
  if (byTime && (kind === KIND_CARE_TIME || kind === KIND_CHILDCARE_TIME) && mins % 30 !== 0) {
    warn.push(`${kind}は30分単位です`);
  }
  if (byTime && mins >= DAY_MIN) {
    warn.push("時間が7時間45分以上です（1日単位にしなくてよいか確認してください）");
  }

  // 申請表に追加（同じ申請番号があれば追加しない）
  let req = readTable(workbook, "申請");
  const exists = req.rows.some(r => str(val(req, r, "申請番号")) === str(responseId));
  if (!exists) {
    const values: Cell[] = new Array<Cell>(req.cols.size).fill("");
    const set = (name: string, v: Cell) => { values[col(req, name)] = v; };
    set("申請番号", num(responseId) || str(responseId));
    set("受付日時", jstText(submittedAt));
    set("メールアドレス", str(email));
    set("出勤用ID", staff ? staff.id : "");
    set("氏名", staff ? staff.name : "");
    set("休暇の種類", kind);
    set("種別・内容", detail);
    set("単位", byTime ? "時間単位" : "1日単位");
    set("開始日", start || "");
    set("終了日", byTime ? "" : (end || ""));
    set("日数", byTime ? "" : nDays);
    set("時間", byTime ? mins / 1440 : "");
    set("時刻", byTime ? str(timeRange) : "");
    set("事由", str(reason));
    set("状況", ST_WAIT);
    const blank = req.rows.findIndex(isBlankRow);
    if (blank >= 0) {
      req.table.getRangeBetweenHeaderAndTotal().getRow(blank).setValues([values]);
    } else {
      req.table.addRow(-1, values);
    }
    req = readTable(workbook, "申請");
  }
  const index = req.rows.findIndex(r => str(val(req, r, "申請番号")) === str(responseId));
  const row = req.rows[index];
  const id = staff ? staff.id : "";
  const year = rowYear(req, row) || jstYear();

  // 年次休暇の残りは表に追加したあと（この申請を承認待ちに含めて）確かめる
  if (staff && kind === KIND_ANNUAL) {
    const b = annualBalance(workbook, req, id, year);
    if (b.registered && b.total - b.used - b.pending < 0) {
      warn.push("承認待ちを含めると年次休暇の残りが足りません");
    }
  }
  if (!exists) {
    req.table.getRangeBetweenHeaderAndTotal().getCell(index, col(req, "確認事項")).setValue(warn.join("／"));
  }

  const lines: string[] = [
    `申請者：${staff ? `${staff.name}（ID ${id}）` : str(email)}`,
    `内容：${rowSummary(req, row)}`,
  ];
  if (str(reason)) {
    lines.push(`事由：${str(reason)}`);
  }
  lines.push(`申請番号：${str(responseId)}　受付：${jstText(submittedAt)}`);
  if (staff) {
    lines.push("", ...balanceLines(workbook, req, id, year));
  }
  if (warn.length > 0) {
    lines.push("", "【確認事項】", ...warn.map(w => `・${w}`));
  }
  return {
    found: !!staff,
    staffId: id,
    name: staff ? staff.name : "",
    approverEmail: staff ? staff.approver : "",
    title: `休暇申請：${staff ? staff.name : str(email)}　${rowSummary(req, row)}`.slice(0, 120),
    details: toMarkdown(lines),
    detailsHtml: toHtml(lines),
    warnings: warn.join("／"),
  };
}

// ---- ここから下は4つのスクリプトで共通（直すときは全部そろえる） ----

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
  approver: string;
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
    approver: str(val(t, row, "承認者メールアドレス")),
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

function toHtml(lines: string[]): string {
  return lines.map(s => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")).join("<br>");
}

// 承認の詳細（Markdown）は行末に空白2つで改行になる
function toMarkdown(lines: string[]): string {
  return lines.join("  \n");
}
