# Microsoft Forms の休暇申請システム

Microsoft Forms で休暇を申請し、Power Automate で上司の承認を取り、結果を Excel の台帳（休暇申請台帳.xlsx）に記録する仕組みです。
**構築と運用は別のPCで行います。** このフォルダはそのPCに引き渡す一式を作るためのもので、引き渡し先では Python などは使いません（ブラウザの Microsoft 365 と、組織のパソコンの Notes・PowerShell で作業します）。
照合用Excel は別に引き渡し、その様式（「年次休暇」「月別取得時間」シートの列の並び）に合わせて集計を出します。

組織では Teams・Outlook のメールを使えないため、お知らせ（承認依頼・承認結果・残日数）は Notes で送ります。
Office Script がお知らせを台帳の「通知」表に書き、組織のパソコンで動かす `send_notices.ps1` が、OneDrive で同期した台帳を読んで、まだ送っていないものを Notes のメールで送ります。

## 引き渡し一式の作り方

プロジェクトのフォルダで次を実行すると、プロジェクト直下に `休暇申請システム一式\` と `休暇申請システム一式.zip` ができます（どちらも Git には含めません）。

```
python tools\leave_request\make_package.py
```

一式の中身：

| ファイル | 内容 |
|---|---|
| `手順書.html` | 別のPCでの構築・運用の手順（ブラウザで開く。印刷もできる） |
| `休暇申請台帳.xlsx` | 台帳の空の様式（氏名などは入っていない） |
| `スクリプト\1_申請を登録.txt` など4つ | Office Script。メモ帳で開いて Excel（Web 版）の［自動化］に貼り付ける |
| `通知送信\` | 組織のパソコンで Notes の通知を送る `send_notices.ps1` と、それを起動する .cmd 3つ |

## このフォルダのファイル

| ファイル | 内容 |
|---|---|
| `手順書.html` | 引き渡し先向けの手順書（Forms の質問一覧・フローの設定・通知送信・運用・照合用Excel への反映） |
| `make_ledger.py` | 台帳の空の様式を作る |
| `make_package.py` | 引き渡し一式（フォルダと zip）を作る |
| `office_scripts/*.ts` | Office Script の各スクリプトの本体 |
| `office_scripts/共通.ts` | 4つのスクリプトで共通の部分。Office Script は1ファイルで完結させる必要があるので、`make_package.py` が各スクリプトの後ろに付けて .txt にする |
| `notice_sender/send_notices.ps1` | 台帳の「通知」表を読んで Notes で送る（BOM 付き UTF-8。Windows PowerShell 5.1 で動く） |
| `notice_sender/*.cmd` | send_notices.ps1 を起動する（中身は英数字だけにしてある。cmd.exe はコードページによって日本語のファイル名を読めないため） |

台帳の表の名前・見出しと、フォームの「休暇の種類」の選択肢の文字は、スクリプトの定数（`KIND_` など）と一致させる必要があります。

## send_notices.ps1 のしくみ

- 台帳（xlsx）は Excel を使わずに ZIP として読む（OneDrive や Excel が開いていても読めるように、共有モードで読み込む）。
- 送った通知の通知番号は、スクリプトと同じフォルダの `送信済み.txt` に記録する。初めて動かしたときは、その時点の通知は送らずに記録だけする。
- Notes へは `Notes.NotesSession`（起動中の Notes クライアント）で接続し、だめなら `Lotus.NotesSession` を使う。64 ビットの PowerShell で接続できないときは、32 ビットの PowerShell で起動し直す。
- このPCには Notes がないので、Notes の部分は確かめていない。台帳の読み取りと送信済みの管理は、Notes を置き換えたもので確かめた。
