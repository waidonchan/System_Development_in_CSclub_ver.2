let cachedProps = null;

function getProps() {
  if (!cachedProps) {
    cachedProps = PropertiesService.getScriptProperties();
  }
  return cachedProps;
}

// "HH:mm:ss" → "HH:mm" に変換する
function formatTimeToHHMM(rawTimeStr) {
  const date = new Date(`1970-01-01T${rawTimeStr}`);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
}

function onFormSubmit(e) {
  var values = e.values; // フォームの回答を取得
  var stamp = values[0]; // タイムスタンプを取得
  var mail = values[1]; // メールアドレスを取得
  var club_name = values[2]; // サークル名を取得
  var representative_name = values[3]; // 代表者氏名を取得
  var contact_mail = values[4]; // 連絡用メールアドレスを取得
  var application_date = values[5]; // 申請日を取得
  var start_date = values[6]; // 開始日を取得
  var start_time = values[7]; // 開始時刻を取得
  var end_date = values[8]; // 終了日を取得
  var end_time = values[9]; // 終了時刻を取得
  var result = values[10]; // 結果の情報を取得
  var memo = values[11]; // 備考を取得

  // テンプレートドキュメントのIDを指定
  const templateDocId = getProps().getProperty("TEMPLATE_DOC_ID");
  const templateDoc = DriveApp.getFileById(templateDocId);

  // テンプレートドキュメントをコピーして新しいドキュメントを作成
  const newDoc = templateDoc.makeCopy(
    "キッチンカー報告書(団体用) - " + club_name
  );
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  // ドキュメントの内容を置換
  body.replaceText("{{申請日}}", application_date);
  body.replaceText("{{サークル名}}", club_name);
  body.replaceText("{{開始年月日}}", start_date);
  body.replaceText("{{開始時刻}}", formatTimeToHHMM(start_time));
  body.replaceText("{{終了年月日}}", end_date);
  body.replaceText("{{終了時刻}}", formatTimeToHHMM(end_time));
  body.replaceText("{{結果}}", result);
  body.replaceText("{{備考}}", memo);

  // ドキュメントを保存して閉じる
  doc.saveAndClose();

  // 一時的にスクリプトを停止
  Utilities.sleep(10000);

  // ドキュメントをPDFとしてエクスポート
  const pdf = DriveApp.getFileById(newDocId).getAs("application/pdf");
  const pdfFileName = "キッチンカー利用報告書 - " + club_name + ".pdf";

  // 親フォルダのIDを指定
  const parentFolderId = getProps().getProperty("PARENT_FOLDER_ID");
  const parentFolder = DriveApp.getFolderById(parentFolderId);

  // フォルダ名を作成 (club_name + "_" + end_date)
  const folderName = club_name + "_" + end_date;

  // 新しいフォルダを作成し、フォルダ情報を格納
  const newFolder = parentFolder.createFolder(folderName);
  const folder_information = newFolder.getId();

  // ドキュメントとPDFを新しいフォルダに移動
  const newFolderDestination = DriveApp.getFolderById(folder_information);
  DriveApp.getFileById(newDocId).moveTo(newFolderDestination);
  const pdfFile = newFolderDestination.createFile(pdf);
  pdfFile.setName(pdfFileName);

  // メール本文を生成
  const subject = "キッチンカー出店の報告書の提出： " + club_name;
  const emailBody = `(提出先)課
  ご担当者様 (cc: ${club_name} ${representative_name}さん)

  お世話になっております。△△大学 ooサークルです。
  
  先日、キッチンカー販売を行っていた、${club_name}の報告書を提出いたします。

  ご不明点がございましたら、このメールへの返信にてお知らせください。CSサークルおよび${club_name}の代表である${representative_name}さんが対応させていただきます。

  よろしくお願いいたします。

  CSサークル`;

  // メールを送信
  MailApp.sendEmail({
    to: getProps().getProperty("UNIVERSITY_EMAIL"), // 就職・生活支援課のメールアドレス
    cc: contact_mail,
    subject: subject,
    body: emailBody,
    attachments: [pdfFile], // ここでPDFファイルを添付
  });
}
