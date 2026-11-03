// 照合用Excel の申請データ：申請用紙（申請データ_用紙）と Forms（Forms申請）を合わせる（接続のみ）。
// 列は tools/leave_sheet/queries/申請データ.pq と同じ。年休_自動・月別_自動・日別_自動・申請一覧がここから作る。
Table.Combine({申請データ_用紙, Forms申請})
