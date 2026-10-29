/**
 * 残日数を照会
 * メールアドレスの職員の、年次休暇の残日数と病気休暇・看護休暇の取得時間を、Notes のお知らせとして「通知」表に書く。
 * 「休暇の残日数照会」フォームのフローから呼ぶ。year が空欄なら今年。
 * 職員表にないメールアドレスからの照会は管理者に知らせる。
 */
interface Result {
  found: boolean;
}

function main(workbook: ExcelScript.Workbook, responseId: string, email: string, year: string): Result {
  const staff = findStaff(workbook, email);
  const y = num(year) || jstYear();
  const noticeId = `照会-${str(responseId)}`;
  if (!staff) {
    addNotice(workbook, noticeId, "", "【休暇】職員表にない人からの残日数照会", [
      `職員表にないメールアドレス ${email} から残日数の照会がありました。`,
      "職員表に登録してから、照会し直してもらってください。"]);
    return { found: false };
  }
  const req = readTable(workbook, "申請");
  const lines: string[] = [`${staff.name} さん`, "", ...balanceLines(workbook, req, staff.id, y)];
  const waiting = req.rows.filter(r => normId(val(req, r, "出勤用ID")) === staff.id
    && str(val(req, r, "状況")) === ST_WAIT);
  if (waiting.length > 0) {
    lines.push("", "承認待ちの申請：", ...waiting.map(r => `・${rowSummary(req, r)}`));
  }
  lines.push("", "※ 承認済みの申請と、この仕組みを使い始める前の取得分から計算しています（1日＝7時間45分）。");
  addNotice(workbook, noticeId, staff.notes, `【休暇】${y}年の残日数`, lines);
  return { found: true };
}
