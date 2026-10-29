/**
 * 申請を登録
 * Forms の休暇申請1件を「申請」表に追加し、承認依頼の宛先と本文を返す。
 * 承認者への Notes のお知らせを「通知」表に書く（職員表にない人からの申請は管理者に知らせる）。
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
  const title = `休暇申請：${staff ? staff.name : str(email)}　${rowSummary(req, row)}`.slice(0, 120);
  if (staff) {
    addNotice(workbook, `${str(responseId)}-依頼`, staff.approverNotes, `【承認依頼】${title}`, [
      `${staff.name} さんから休暇の申請がありました。`,
      "Power Automate（https://make.powerautomate.com）の［承認］で、承認または却下してください。",
      "", ...lines]);
  } else {
    addNotice(workbook, `${str(responseId)}-未登録`, "", `【職員表にない申請】${title}`, [
      "職員表にないメールアドレスから休暇の申請がありました。承認依頼は出していません。",
      "職員表に登録してから、申請し直してもらってください。",
      "", ...lines]);
  }
  return {
    found: !!staff,
    staffId: id,
    name: staff ? staff.name : "",
    approverEmail: staff ? staff.approver : "",
    title,
    details: toMarkdown(lines),
    warnings: warn.join("／"),
  };
}
