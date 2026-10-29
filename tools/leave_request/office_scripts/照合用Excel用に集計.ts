/**
 * 照合用Excel用に集計
 * 「集計」シートの B1 の年について、承認済みの申請を照合用Excel と同じ列の並びで出力する。
 *   A4〜：年次休暇シート用（出勤用ID・年・本年分（空欄）・本年使用（日）・本年使用（時間））
 *   G4〜：月別取得時間シート用（出勤用ID・年・区分・1月〜12月）
 * 年次休暇の本年使用には「年次休暇の持ち日数」表の導入前の使用を含める。
 * 病気休暇・看護休暇の1日単位の申請は1日＝7時間45分で時間にする。日付は開始日の月に入れる。
 * Excel の［自動化］タブから実行する。
 */
function main(workbook: ExcelScript.Workbook) {
  const ws = workbook.getWorksheet("集計");
  const year = num(ws.getRange("B1").getValue());
  if (!year) {
    throw new Error("集計シートの B1 に年（西暦）を入れてください");
  }
  const req = readTable(workbook, "申請");
  const staff = readTable(workbook, "職員");
  const hold = readTable(workbook, "持ち日数");

  // 職員表の順に並べる（職員表にない ID は後ろに）
  const order: string[] = [];
  const add = (id: string) => {
    if (id && order.indexOf(id) < 0) {
      order.push(id);
    }
  };
  staff.rows.forEach(r => add(normId(val(staff, r, "出勤用ID"))));

  const annualDays = new Map<string, number>();   // 本年使用（日）
  const annualMin = new Map<string, number>();    // 本年使用（時間）の分
  const monthly = new Map<string, number[]>();    // "ID|区分" → 1〜12月の分
  const addAnnual = (id: string, d: number, m: number) => {
    add(id);
    annualDays.set(id, (annualDays.get(id) || 0) + d);
    annualMin.set(id, (annualMin.get(id) || 0) + m);
  };
  hold.rows.forEach(r => {
    const id = normId(val(hold, r, "出勤用ID"));
    if (id && num(val(hold, r, "年")) === year) {
      addAnnual(id, num(val(hold, r, "導入前の使用（日）")), Math.round(num(val(hold, r, "導入前の使用（時間）")) * 1440));
    }
  });
  req.rows.forEach(r => {
    const id = normId(val(req, r, "出勤用ID"));
    if (!id || str(val(req, r, "状況")) !== ST_OK || rowYear(req, r) !== year) {
      return;
    }
    const kind = str(val(req, r, "休暇の種類"));
    if (kind === KIND_ANNUAL) {
      addAnnual(id, num(val(req, r, "日数")), Math.round(num(val(req, r, "時間")) * 1440));
    } else if (kind === KIND_SICK || kind === KIND_NURSING) {
      add(id);
      const key = `${id}|${kind === KIND_SICK ? "病気休暇" : "看護休暇"}`;
      const months = monthly.get(key) || [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
      months[serialDate(toSerial(val(req, r, "開始日"))).getUTCMonth()] += rowMinutes(req, r);
      monthly.set(key, months);
    }
  });

  const annual: Cell[][] = [["出勤用ID", "年（西暦）", "本年分（日）", "本年使用（日）", "本年使用（時間）"]];
  const month: Cell[][] = [["出勤用ID", "年（西暦）", "区分", "1月", "2月", "3月", "4月", "5月", "6月",
    "7月", "8月", "9月", "10月", "11月", "12月"]];
  for (const id of order) {
    if (annualDays.has(id)) {
      annual.push([id, year, "", annualDays.get(id) || 0, (annualMin.get(id) || 0) / 1440]);
    }
    for (const cat of ["病気休暇", "看護休暇"]) {
      const months = monthly.get(`${id}|${cat}`);
      if (months) {
        const cells: Cell[] = [id, year, cat];
        months.forEach(m => cells.push(m ? m / 1440 : ""));
        month.push(cells);
      }
    }
  }

  // 書式（ID は文字列、時間は「時間:分」）を先に付けてから値を入れる
  ws.getRange("A3:Z100000").clear(ExcelScript.ClearApplyTo.all);
  ws.getRange("A3").setValue("年次休暇シート用：本年使用（日・時間）を照合用Excel の同じ ID・年の行に写します");
  ws.getRange("G3").setValue("月別取得時間シート用：行ごと貼り付けできます（同じ ID・年・区分の古い行は消してください）");
  const a = ws.getRange("A4").getResizedRange(annual.length - 1, annual[0].length - 1);
  a.getColumn(0).setNumberFormat(annual.map(() => ["@"]));
  a.getColumn(4).setNumberFormat(annual.map((_, i) => [i === 0 ? "General" : "[h]:mm"]));
  a.setValues(annual);
  const m = ws.getRange("G4").getResizedRange(month.length - 1, month[0].length - 1);
  m.getColumn(0).setNumberFormat(month.map(() => ["@"]));
  for (let c = 3; c < 15; c++) {
    m.getColumn(c).setNumberFormat(month.map((_, i) => [i === 0 ? "General" : "[h]:mm"]));
  }
  m.setValues(month);
  ws.getRange("A4:E4").getFormat().getFont().setBold(true);
  ws.getRange("G4:U4").getFormat().getFont().setBold(true);
}
