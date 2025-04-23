### ✅ 主な機能

- 📑 **Google フォーム回答の受信トリガー処理（onFormSubmit）**
- 📝 **テンプレートドキュメントのコピー＆自動編集**
- 🕒 **開始・終了時刻のフォーマット変換（HH:mm）**
- 📄 **Google ドキュメント → PDF に変換**
- 📁 **Google ドライブ上に団体別フォルダを作成しファイルを整理**
- 📧 **PDF を添付したメールを大学に自動送信**

---

### 🗂️ ファイル構成の概要

```text
📦 report/
 ┣ 📄 club_report.js ... 団体用フォーム送信をトリガーに処理を行うメインスクリプト
 ┣ 📄 individual_report.js ... 個人用フォーム送信をトリガーに処理を行うメインスクリプト
 ┣ 📄 report.md ... このファイル
```

---

### ⚙️ 導入方法

1. Google Apps Script プロジェクトを作成し、`club_report.js` および `individual_report.js` の内容を貼り付け
2. [スクリプトプロパティ（`PropertiesService`）](https://developers.google.com/apps-script/reference/properties/properties-service) に以下の値を登録：

| プロパティ名       | 説明                                |
| ------------------ | ----------------------------------- |
| `PARENT_FOLDER_ID` | ファイルを格納する親フォルダの ID   |
| `TEMPLATE_DOC_ID`  | 団体用テンプレートドキュメントの ID |
| `UNIVERSITY_EMAIL` | 提出先（大学）のメールアドレス      |

3. Google フォームの回答をトリガーに `onFormSubmit` を設定  
   　※「トリガー」メニューから関数名 `onFormSubmit` を選び、「フォーム送信時」を選択してください。

---

### 📸 出力例

- 📝 `キッチンカー報告書(団体用) - ○○サークル`（Google ドキュメント）
- 📄 `キッチンカー利用報告書 - ○○サークル.pdf`（自動で PDF 変換）
- 📁 `○○サークル_年-月-日`（Google ドライブ内のフォルダ）
- 📧 自動メール送信（PDF 添付付き）
