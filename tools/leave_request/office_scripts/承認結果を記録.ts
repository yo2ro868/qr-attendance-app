/**
 * 承認結果を記録
 * 承認・却下の結果を「申請」表の該当行に書き込み、申請者への Notes のお知らせを「通知」表に書く。
 * Power Automate の「開始して承認を待機」のあとの「スクリプトの実行」から呼ぶ。
 * outcome は承認アクションの結果（Approve / Reject）をそのまま渡す。
 */
interface Result {
  found: boolean;          // 申請番号の行があったか
  status: string;          // 承認 / 却下
}

function main(workbook: ExcelScript.Workbook,
  responseId: string, outcome: string, approver: string, comments: string): Result {
  const req = readTable(workbook, "申請");
  const index = req.rows.findIndex(r => str(val(req, r, "申請番号")) === str(responseId));
  if (index < 0) {
    return { found: false, status: "" };
  }
  const o = str(outcome).toLowerCase();
  const status = o === "approve" || o === "approved" || o === ST_OK ? ST_OK : ST_NG;
  const body = req.table.getRangeBetweenHeaderAndTotal();
  body.getCell(index, col(req, "状況")).setValue(status);
  body.getCell(index, col(req, "承認者")).setValue(str(approver));
  body.getCell(index, col(req, "処理日時")).setValue(jstText());
  body.getCell(index, col(req, "承認者コメント")).setValue(str(comments));

  const after = readTable(workbook, "申請");
  const row = after.rows[index];
  const id = normId(val(after, row, "出勤用ID"));
  const year = rowYear(after, row) || jstYear();
  const lines: string[] = [
    `${str(val(after, row, "氏名"))} さん`,
    "",
    `休暇の申請が${status}されました。`,
    "",
    `内容：${rowSummary(after, row)}`,
    `申請番号：${str(responseId)}`,
    `承認者：${str(approver)}`,
  ];
  if (str(comments)) {
    lines.push(`コメント：${str(comments)}`);
  }
  if (id) {
    lines.push("", ...balanceLines(workbook, after, id, year));
  }
  const staff = findStaff(workbook, str(val(after, row, "メールアドレス")));
  addNotice(workbook, `${str(responseId)}-結果`, staff ? staff.notes : "",
    `【休暇申請】${status}：${rowSummary(after, row)}`.slice(0, 120), lines);
  return { found: true, status };
}
