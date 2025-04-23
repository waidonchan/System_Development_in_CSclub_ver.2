## 🧩 主な機能

- Slack での申請通知（個人・団体）
- リアクションによる承認・却下判定（マル/バツスタンプ）
- リマインド通知の送信（2 日前・1 週間前）
- 未承認申請の定期リマインド
- スプレッドシート行の背景色変更（承認＝緑／却下＝赤）
- 自動メール送信（代表者＋管理者）
- 新しく slack チャンネルに加入した人に DM でメッセージ

---

## 🔧 使用技術

- Google Apps Script (GAS)
- Slack API（chat.postMessage, reaction_added, reaction_removed）
- Google スプレッドシート

---

## 📂 ファイル構成

```
📁 slackApp/
└── slackApp.js         // メインスクリプト（Slack受信、通知処理、リマインド管理など）
```

---

## 🚀 セットアップ手順

1. **Google Apps Script プロジェクトの作成**

   - Google Drive 上で Apps Script を作成し、`Code.gs`の内容を貼り付ける。

2. **スクリプトプロパティの設定**

   - `SLACK_TOKEN`: Slack Bot Token（`xoxb-...`）
   - `CHANNEL_ID`: 通知を投稿する Slack チャンネルの ID
   - `SPREADSHEET_ID_KARI`: 仮申請スプレッドシートの ID
   - `SPREADSHEET_ID_HON`: 本申請（リマインド）スプレッドシートの ID
   - `ADMINISTRATOR_EMAIL`: 承認メールで cc に含める管理者アドレス
   - `FOLLOWUP_FORM_URL`: 本申請フォームの URL
   - `TARGET_MESSAGES`: 初期は空配列 (`[]`) を設定

3. **Slack アプリの構築と設定**

   - Slack API でアプリ作成
   - OAuth スコープ：`chat:write`, `reactions:read`, `reactions:write`, `users:read`
   - Event Subscriptions:
     - Request URL → GAS デプロイ URL
     - Subscribe to bot events: `reaction_added`, `reaction_removed`
   - Interactivity: OFF でも可（スタンプ処理のみ対応）

4. **デプロイ**
   - `公開 > ウェブアプリケーションとして導入`
   - トリガーの設定で `remindUnprocessedMessages` を定期実行（例: 毎日 9:00）

---

## 🔄 処理フロー概要

```
[フォーム送信]
    ↓
Slackに申請メッセージを投稿
    ↓
リアクション（マル or バツ）
    ↓
承認 or 却下の自動判定
    ↓
スプレッドシート反映 / メール送信
    ↓
2日前・1週間前にSlack＆メールでリマインド
```

---

## 📝 補足事項

- `club_name`が空白なら個人申請、それ以外は団体申請と判断されます。
- `remindUnprocessedMessages()`は Slack 通知未処理の申請に対し、スタンプを促すメッセージを再送します。
