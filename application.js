let cachedProps = null;

function getProps() {
  if (!cachedProps) {
    cachedProps = PropertiesService.getScriptProperties();
  }
  return cachedProps;
}

// 時間 分まで表示
function formatTime(rawTimeStr) {
  const date = new Date(`1970-01-01T${rawTimeStr}`);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
}

// 列7が「個人」の場合は個人用処理
function onFormSubmit(e) {
  var club_or_individual = e.values[50]; // 列50が個人の場合は個人用処理
  if (club_or_individual.trim() === "個人") {
    handleIndividualSubmit(e);
    Logger.log("個人用の通知を受け取りました");
  } else {
    handleClubSubmission(e);
    Logger.log("団体用の通知を受け取りました");
  }
}

function endEmail() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("リマインド");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues(); // A〜I列

  const today = new Date();

  data.forEach((row, index) => {
    const club_name = row[0]; // A列
    const name = row[1]; // B列
    const endDay = row[6]; // G列
    const email = row[8]; // I列

    if (areDatesEqual(new Date(endDay), today)) {
      if (!club_name || club_name.trim() === "") {
        individual_sendFormEmail(email, name);
        Logger.log(`個人用メール送信：${name}`);
      } else {
        club_sendFormEmail(email, club_name, name);
        Logger.log(`団体用メール送信：${club_name}の${name}`);
      }
    }
  });
}

// 日付を文字列にフォーマット (例: yyyy-mm-dd)
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// 新しいフォルダを作成する関数
function createFolderInParent(parentFolderId, newFolderName) {
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var newFolder = parentFolder.createFolder(newFolderName);
  return newFolder.getId(); // フォルダIDを返す
}

// ------------------------------------個人用-----------------------------------------------

function handleIndividualSubmit(e) {
  var values = e.values; // フォームの回答を取得
  var stamp = values[0]; // タイムスタンプを取得
  var mail = values[1]; // メールアドレスを取得
  var representative_name = values[2]; // 代表者氏名を取得
  var representative_number = values[51]; // 代表者氏名を取得
  var contact_mail = values[3]; // 連絡用メールアドレスを取得
  var application_date = values[4]; // 申請日を取得
  var club_name = values[5]; // クラブサークル名 (空欄) を取得
  var start_date = values[6]; // 開始日を取得
  var start_time = values[7]; // 開始時刻を取得
  var end_date = values[8]; // 終了日を取得
  var end_time = values[9]; // 終了時刻を取得
  var before_participants = values[10]; // 前日準備参加予定人数を取得
  var today_participants = values[11]; // 販売当日参加予定人数を取得

  // 1人目の情報を取得
  var post_first = values[12]; // 役職(1人目)を取得
  var name_first = values[13]; // 氏名(1人目)を取得
  var department_first = values[14]; // 学科(1人目)を取得
  var grade_first = values[15]; // 学年(1人目)を取得

  // 2人目の情報を取得
  var post_second = values[16]; // 役職(2人目)を取得
  var name_second = values[17]; // 氏名(2人目)を取得
  var department_second = values[18]; // 学科(2人目)を取得
  var grade_second = values[19]; // 学年(2人目)を取得

  // 3人目の情報を取得
  var post_third = values[20]; // 役職(3人目)を取得
  var name_third = values[21]; // 氏名(3人目)を取得
  var department_third = values[22]; // 学科(3人目)を取得
  var grade_third = values[23]; // 学年(3人目)を取得

  // 4人目の情報を取得
  var post_fourth = values[24]; // 役職(4人目)を取得
  var name_fourth = values[25]; // 氏名(4人目)を取得
  var department_fourth = values[26]; // 学科(4人目)を取得
  var grade_fourth = values[27]; // 学年(4人目)を取得

  // 5人目の情報を取得
  var post_fifth = values[28]; // 役職(5人目)を取得
  var name_fifth = values[29]; // 氏名(5人目)を取得
  var department_fifth = values[30]; // 学科(5人目)を取得
  var grade_fifth = values[31]; // 学年(5人目)を取得

  // 6人目の情報を取得
  var post_sixth = values[32]; // 役職(6人目)を取得
  var name_sixth = values[33]; // 氏名(6人目)を取得
  var department_sixth = values[34]; // 学科(6人目)を取得
  var grade_sixth = values[35]; // 学年(6人目)を取得

  // 7人目の情報を取得
  var post_seventh = values[36]; // 役職(7人目)を取得
  var name_seventh = values[37]; // 氏名(7人目)を取得
  var department_seventh = values[38]; // 学科(7人目)を取得
  var grade_seventh = values[39]; // 学年(7人目)を取得

  // 8人目の情報を取得
  var post_eighth = values[40]; // 役職(8人目)を取得
  var name_eighth = values[41]; // 氏名(8人目)を取得
  var department_eighth = values[42]; // 学科(8人目)を取得
  var grade_eighth = values[43]; // 学年(8人目)を取得

  // 販売物概要の情報を取得
  var sale_information = values[44]; // 販売物の情報を取得
  var sale_image = values[45]; // 販売物の写真1を取得
  var sale_image2 = values[46]; // 販売物の写真2を取得
  var sale_image3 = values[47]; // 販売物の写真3を取得
  var memo = values[48]; // 備考を取得

  // 写真を挿入
  sale_image = sale_image.replace("https://drive.google.com/open?id=", "");
  let attachImg = DriveApp.getFileById(sale_image).getBlob();

  if (sale_image2) {
    sale_image2 = sale_image2.replace("https://drive.google.com/open?id=", "");
    var attachImg2 = DriveApp.getFileById(sale_image2).getBlob();
  }

  if (sale_image3) {
    sale_image3 = sale_image3.replace("https://drive.google.com/open?id=", "");
    var attachImg3 = DriveApp.getFileById(sale_image3).getBlob();
  }

  // カレンダーの設定とイベント作成
  const calendarId = getProps().getProperty("CALENDAR_ID");
  let calendar = CalendarApp.getCalendarById(calendarId);

  let startDate = new Date(start_date + " " + start_time);
  let endDate = new Date(end_date + " " + end_time);

  if (start_time === "" || end_time === "") {
    endDate.setDate(endDate.getDate() + 1);
    event = calendar.createAllDayEvent(
      "キッチンカー利用 - " + representative_name,
      startDate,
      endDate,
      { description: sale_information }
    );
  } else {
    event = calendar.createEvent(
      "キッチンカー利用 - " + representative_name,
      startDate,
      endDate,
      { description: sale_information }
    );
  }

  // イベントの色をピンクに設定
  event.setColor("4");

  // フォルダ名をrepresentative_name + start_dateにして、それをfolder_informationに格納
  var folder_information = representative_name + "_" + start_date;

  // 親フォルダのIDを指定
  const parentFolderId = getProps().getProperty("PARENT_FOLDER_ID");
  const newFolderId = createFolderInParent(parentFolderId, folder_information);
  const newFolder = DriveApp.getFolderById(newFolderId); // 新しいフォルダを取得

  // テンプレートドキュメントのIDを指定
  const templateDocId = getProps().getProperty("INDIVIDUAL_TEMPLATE_DOC_ID");
  const templateDoc = DriveApp.getFileById(templateDocId);
  const newDoc = templateDoc.makeCopy(
    "キッチンカー利用申込書(個人用) - " + representative_name,
    newFolder
  );
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  const formattedStartTime = formatTime(start_time);
  const formattedEndTime = formatTime(end_time);

  // ドキュメントの内容を置換
  body.replaceText("{{申請日}}", application_date);
  body.replaceText("{{開始年月日}}", start_date);
  body.replaceText("{{開始時刻}}", formattedStartTime);
  body.replaceText("{{終了年月日}}", end_date);
  body.replaceText("{{終了時刻}}", formattedEndTime);
  body.replaceText("{{代表者氏名}}", representative_name);
  body.replaceText("{{前日準備参加予定人数}}", before_participants);
  body.replaceText("{{販売当日参加予定人数}}", today_participants);

  // 1人目の情報を置換
  body.replaceText("{{一人目役職}}", post_first);
  body.replaceText("{{一人目氏名}}", name_first);
  body.replaceText("{{一人目学科}}", department_first);
  body.replaceText("{{一人目学年}}", grade_first);

  // 2人目の情報を置換
  body.replaceText("{{二人目役職}}", post_second);
  body.replaceText("{{二人目氏名}}", name_second);
  body.replaceText("{{二人目学科}}", department_second);
  body.replaceText("{{二人目学年}}", grade_second);

  // 3人目の情報を置換
  body.replaceText("{{三人目役職}}", post_third);
  body.replaceText("{{三人目氏名}}", name_third);
  body.replaceText("{{三人目学科}}", department_third);
  body.replaceText("{{三人目学年}}", grade_third);

  // 4人目の情報を置換
  body.replaceText("{{四人目役職}}", post_fourth);
  body.replaceText("{{四人目氏名}}", name_fourth);
  body.replaceText("{{四人目学科}}", department_fourth);
  body.replaceText("{{四人目学年}}", grade_fourth);

  // 5人目の情報を置換
  body.replaceText("{{五人目役職}}", post_fifth);
  body.replaceText("{{五人目氏名}}", name_fifth);
  body.replaceText("{{五人目学科}}", department_fifth);
  body.replaceText("{{五人目学年}}", grade_fifth);

  // 6人目の情報を置換
  body.replaceText("{{六人目役職}}", post_sixth);
  body.replaceText("{{六人目氏名}}", name_sixth);
  body.replaceText("{{六人目学科}}", department_sixth);
  body.replaceText("{{六人目学年}}", grade_sixth);

  // 7人目の情報を置換
  body.replaceText("{{七人目役職}}", post_seventh);
  body.replaceText("{{七人目氏名}}", name_seventh);
  body.replaceText("{{七人目学科}}", department_seventh);
  body.replaceText("{{七人目学年}}", grade_seventh);

  // 8人目の情報を置換
  body.replaceText("{{八人目役職}}", post_eighth);
  body.replaceText("{{八人目氏名}}", name_eighth);
  body.replaceText("{{八人目学科}}", department_eighth);
  body.replaceText("{{八人目学年}}", grade_eighth);

  // 販売物概要の情報を置換
  body.replaceText("{{販売物の情報}}", sale_information);
  body.replaceText("{{備考}}", memo);

  // 画像の縦横比を取得
  let res = ImgApp.getSize(attachImg);
  let width = res.width;
  let height = res.height;
  // 画像を横300pxでアスペクト比を揃えて大きさを編集し最終行へ挿入
  body
    .appendImage(attachImg)
    .setWidth(300)
    .setHeight((300 * height) / width);

  // 2枚目の画像が存在する場合、挿入
  if (attachImg2) {
    let res2 = ImgApp.getSize(attachImg2);
    let width2 = res2.width;
    let height2 = res2.height;
    body
      .appendImage(attachImg2)
      .setWidth(300)
      .setHeight((300 * height2) / width2);
  }

  // 3枚目の画像が存在する場合、挿入
  if (attachImg3) {
    let res3 = ImgApp.getSize(attachImg3);
    let width3 = res3.width;
    let height3 = res3.height;
    body
      .appendImage(attachImg3)
      .setWidth(300)
      .setHeight((300 * height3) / width3);
  }

  // ドキュメントを保存して閉じる
  doc.saveAndClose();

  // 新たにスプレッドシートをコピー
  const spreadsheetTemplateId = getProps().getProperty("TEMPLATE_SHEET_ID");
  const spreadsheetTemplate = DriveApp.getFileById(spreadsheetTemplateId);
  const newSpreadsheet = spreadsheetTemplate.makeCopy(
    representative_name + "さん_持ち物リストと予定表",
    newFolder
  );

  // スプレッドシートの共有設定を変更
  newSpreadsheet.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.EDIT
  );

  // 新しく作成したドキュメントとスプレッドシートのURLを取得
  const newDocUrl = "https://docs.google.com/document/d/" + newDocId;
  const newSpreadsheetUrl =
    "https://docs.google.com/spreadsheets/d/" + newSpreadsheet.getId();

  // このスクリプト自体が動いているフォーム連携スプレッドシートの2枚目（リマインド）シートに書き込み
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reminderSheet = spreadsheet.getSheetByName("リマインド");

  // 日付を操作するために、start_dateをDateオブジェクトに変換
  let startDay = new Date(start_date);
  let twoDaysBefore = new Date(startDay);
  let oneWeekBefore = new Date(startDay);

  // 二日前と一週間前の日付を計算
  twoDaysBefore.setDate(startDay.getDate() - 2);
  oneWeekBefore.setDate(startDay.getDate() - 7);

  let formattedStartDay = formatDate(startDay);
  let formattedTwoDaysBefore = formatDate(twoDaysBefore);
  let formattedOneWeekBefore = formatDate(oneWeekBefore);

  // 新しい行にデータを書き込む
  reminderSheet.appendRow([
    club_name, // ① 空欄の団体名（識別の基本）
    representative_name, // ② 代表者名（連絡相手）

    start_time, // ③ 開始時刻（時系列要素その1）
    formattedStartDay, // ④ 開始日（中心日程）
    formattedTwoDaysBefore, // ⑤ 二日前（通知用）
    formattedOneWeekBefore, // ⑥ 一週間前（通知用）
    end_date, // ⑦ 終了日（後続処理に必要）

    newSpreadsheetUrl, // ⑧ スプレッドシートURL（参照用）
    contact_mail, // ⑨ 連絡先メールアドレス（通知に使用）
  ]);

  // ドキュメントの保存直後、PDF変換までのタイミング調整のため
  Utilities.sleep(10000);

  // 作成したドキュメントをPDFに変換し、フォルダに格納
  const pdfBlob = DriveApp.getFileById(newDocId).getAs("application/pdf");
  const pdfFileName =
    "キッチンカー利用申込書(個人用) - " + representative_name + ".pdf";
  const pdfFile = newFolder.createFile(pdfBlob).setName(pdfFileName);

  // メール本文を生成
  const subject = "キッチンカー出店に関して： " + representative_name + "さん";
  const emailBody = `就職・生活支援課
  ご担当者様 (cc: ${representative_name}さん)

  お世話になっております。福井県立大学 CSサークルです。
  
  この度、キッチンカー利用の申請をさせていただきたく、ご連絡いたしました。

  今回、利用を希望しているのは、学籍番号${representative_number}の${representative_name}さんです。

  詳細につきましては、添付の資料をご確認いただけますと幸いです。

  また、ご不明点やご質問がございましたら、このメールへの返信にてお知らせください。CSサークルおよび代表である${representative_name}さんが対応させていただきます。

  よろしくお願いいたします。

  CSサークル`;

  MailApp.sendEmail({
    to: getProps().getProperty("UNIVERSITY_EMAIL"), // 就職・生活支援課のメールアドレスに変更
    cc: contact_mail,
    subject: subject,
    body: emailBody,
    attachments: [pdfFile], // ここでPDFファイルを添付
  });

  // メールの件名と本文をまとめて定義
  const newEmailSubject = `タイムスケジュール及び持ち物チェックリスト作成のお願い： ${representative_name}`;
  const newEmailBody = `${representative_name}さん （cc:柴田代表）

  こんにちは、CSサークルです。

  資料作成フォームにご記入いただき、ありがとうございます。

  タイムスケジュール及び持ち物チェックリストの作成をお願いしたく、ご連絡いたしました。つきましては、以下のスプレッドシートに直接ご記入くださいますようお願い申し上げます。

  スプレッドシートリンク: ${newSpreadsheetUrl}

  記入が完了しましたら、このメールへの返信にてお知らせください。

  参考までに、就職・生活支援課に提出したPDFを添付いたします。

  また、キッチンカーのご利用に際してご質問がある場合も、このメールへの返信にてお問い合わせください。CSサークルおよびキッチンカー利用の責任者である柴田代表が対応いたします。

  よろしくお願いいたします。

  CSサークル`;

  // 新しいメールを送信
  MailApp.sendEmail({
    to: contact_mail,
    cc: getProps().getProperty("ADMINISTRATOR_EMAIL"), //柴田代表のメールアドレス
    subject: newEmailSubject,
    body: newEmailBody,
    attachments: [pdfFile],
  });
}

function individual_sendFormEmail(email, representativeName) {
  // スクリプトプロパティからフォームURLを取得
  const individual_formUrl = getProps().getProperty(
    "INDIVIDUAL_SUBMISSION_FORM_URL"
  );

  var bodyText = `
  ${representativeName}さん

  こんにちは、CSサークルです。

  キッチンカーでの販売、お疲れさまでした。

  今回のイベントの結果を学校側に報告するため、以下のフォームへの記入をお願いします。

  フォームリンク：${individual_formUrl}

  このフォームに記載すると、自動的にドキュメントが生成され、書類が提出されます。

  ご質問や不明点がございましたら、いつでもご連絡ください。

  よろしくお願いいたします。

  CSサークル
  `;

  // HTMLメールとして送信
  var htmlBody = bodyText.replace(/\n/g, "<br>");

  MailApp.sendEmail({
    to: email,
    subject: "【重要】提出書類作成用フォームへの記入のお願い",
    body: bodyText, // プレーンテキストの本文
    htmlBody: htmlBody, // HTML形式の本文
  });
}

// ------------------------------------団体用-----------------------------------------------

function handleClubSubmission(e) {
  var values = e.values; // フォームの回答を取得
  var stamp = values[0]; // タイムスタンプを取得
  var mail = values[1]; // メールアドレスを取得
  var representative_name = values[2]; // 代表者氏名を取得
  var contact_mail = values[3]; // 連絡用メールアドレスを取得
  var application_date = values[4]; // 申請日を取得
  var club_name = values[5]; // クラブサークル名を取得
  var start_date = values[6]; // 開始日を取得
  var start_time = values[7]; // 開始時刻を取得
  var end_date = values[8]; // 終了日を取得
  var end_time = values[9]; // 終了時刻を取得
  var before_participants = values[10]; // 前日準備参加予定人数を取得
  var today_participants = values[11]; // 販売当日参加予定人数を取得

  // 1人目の情報を取得
  var post_first = values[12]; // 役職(1人目)を取得
  var name_first = values[13]; // 氏名(1人目)を取得
  var department_first = values[14]; // 学科(1人目)を取得
  var grade_first = values[15]; // 学年(1人目)を取得

  // 2人目の情報を取得
  var post_second = values[16]; // 役職(2人目)を取得
  var name_second = values[17]; // 氏名(2人目)を取得
  var department_second = values[18]; // 学科(2人目)を取得
  var grade_second = values[19]; // 学年(2人目)を取得

  // 3人目の情報を取得
  var post_third = values[20]; // 役職(3人目)を取得
  var name_third = values[21]; // 氏名(3人目)を取得
  var department_third = values[22]; // 学科(3人目)を取得
  var grade_third = values[23]; // 学年(3人目)を取得

  // 4人目の情報を取得
  var post_fourth = values[24]; // 役職(4人目)を取得
  var name_fourth = values[25]; // 氏名(4人目)を取得
  var department_fourth = values[26]; // 学科(4人目)を取得
  var grade_fourth = values[27]; // 学年(4人目)を取得

  // 5人目の情報を取得
  var post_fifth = values[28]; // 役職(5人目)を取得
  var name_fifth = values[29]; // 氏名(5人目)を取得
  var department_fifth = values[30]; // 学科(5人目)を取得
  var grade_fifth = values[31]; // 学年(5人目)を取得

  // 6人目の情報を取得
  var post_sixth = values[32]; // 役職(6人目)を取得
  var name_sixth = values[33]; // 氏名(6人目)を取得
  var department_sixth = values[34]; // 学科(6人目)を取得
  var grade_sixth = values[35]; // 学年(6人目)を取得

  // 7人目の情報を取得
  var post_seventh = values[36]; // 役職(7人目)を取得
  var name_seventh = values[37]; // 氏名(7人目)を取得
  var department_seventh = values[38]; // 学科(7人目)を取得
  var grade_seventh = values[39]; // 学年(7人目)を取得

  // 8人目の情報を取得
  var post_eighth = values[40]; // 役職(8人目)を取得
  var name_eighth = values[41]; // 氏名(8人目)を取得
  var department_eighth = values[42]; // 学科(8人目)を取得
  var grade_eighth = values[43]; // 学年(8人目)を取得

  // 販売物概要の情報を取得
  var sale_information = values[44]; // 販売物の情報を取得
  var sale_image = values[45]; // 販売物の写真1を取得
  var sale_image2 = values[46]; // 販売物の写真2を取得
  var sale_image3 = values[47]; // 販売物の写真3を取得
  var memo = values[48]; // 備考を取得

  // 写真を挿入
  sale_image = sale_image.replace("https://drive.google.com/open?id=", "");
  let attachImg = DriveApp.getFileById(sale_image).getBlob();

  if (sale_image2) {
    sale_image2 = sale_image2.replace("https://drive.google.com/open?id=", "");
    var attachImg2 = DriveApp.getFileById(sale_image2).getBlob();
  }

  if (sale_image3) {
    sale_image3 = sale_image3.replace("https://drive.google.com/open?id=", "");
    var attachImg3 = DriveApp.getFileById(sale_image3).getBlob();
  }

  // カレンダーの設定とイベント作成
  const calendarId = getProps().getProperty("CALENDAR_ID");
  let calendar = CalendarApp.getCalendarById(calendarId);

  let startDate = new Date(start_date + " " + start_time);
  let endDate = new Date(end_date + " " + end_time);

  if (start_time === "" || end_time === "") {
    endDate.setDate(endDate.getDate() + 1);
    event = calendar.createAllDayEvent(
      "キッチンカー利用 - " + club_name,
      startDate,
      endDate,
      { description: sale_information }
    );
  } else {
    event = calendar.createEvent(
      "キッチンカー利用 - " + club_name,
      startDate,
      endDate,
      { description: sale_information }
    );
  }

  // イベントの色をピンクに設定
  event.setColor("4");

  // フォルダ名をclub_name + start_dateにして、それをfolder_informationに格納
  var folder_information = club_name + "_" + start_date;

  // 親フォルダのIDを指定
  const parentFolderId = getProps().getProperty("PARENT_FOLDER_ID");
  const newFolderId = createFolderInParent(parentFolderId, folder_information);
  const newFolder = DriveApp.getFolderById(newFolderId); // 新しいフォルダを取得

  // テンプレートドキュメントのIDを指定
  const templateDocId = getProps().getProperty("TEMPLATE_DOC_ID");
  const templateDoc = DriveApp.getFileById(templateDocId);
  const newDoc = templateDoc.makeCopy(
    "キッチンカー利用申込書(団体用) - " + club_name,
    newFolder
  );
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  const formattedStartTime = formatTime(start_time);
  const formattedEndTime = formatTime(end_time);

  // ドキュメントの内容を置換
  body.replaceText("{{申請日}}", application_date);
  body.replaceText("{{クラブサークル名}}", club_name);
  body.replaceText("{{開始年月日}}", start_date);
  body.replaceText("{{開始時刻}}", formattedStartTime);
  body.replaceText("{{終了年月日}}", end_date);
  body.replaceText("{{終了時刻}}", formattedEndTime);
  body.replaceText("{{前日準備参加予定人数}}", before_participants);
  body.replaceText("{{販売当日参加予定人数}}", today_participants);

  // 1人目の情報を置換
  body.replaceText("{{一人目役職}}", post_first);
  body.replaceText("{{一人目氏名}}", name_first);
  body.replaceText("{{一人目学科}}", department_first);
  body.replaceText("{{一人目学年}}", grade_first);

  // 2人目の情報を置換
  body.replaceText("{{二人目役職}}", post_second);
  body.replaceText("{{二人目氏名}}", name_second);
  body.replaceText("{{二人目学科}}", department_second);
  body.replaceText("{{二人目学年}}", grade_second);

  // 3人目の情報を置換
  body.replaceText("{{三人目役職}}", post_third);
  body.replaceText("{{三人目氏名}}", name_third);
  body.replaceText("{{三人目学科}}", department_third);
  body.replaceText("{{三人目学年}}", grade_third);

  // 4人目の情報を置換
  body.replaceText("{{四人目役職}}", post_fourth);
  body.replaceText("{{四人目氏名}}", name_fourth);
  body.replaceText("{{四人目学科}}", department_fourth);
  body.replaceText("{{四人目学年}}", grade_fourth);

  // 5人目の情報を置換
  body.replaceText("{{五人目役職}}", post_fifth);
  body.replaceText("{{五人目氏名}}", name_fifth);
  body.replaceText("{{五人目学科}}", department_fifth);
  body.replaceText("{{五人目学年}}", grade_fifth);

  // 6人目の情報を置換
  body.replaceText("{{六人目役職}}", post_sixth);
  body.replaceText("{{六人目氏名}}", name_sixth);
  body.replaceText("{{六人目学科}}", department_sixth);
  body.replaceText("{{六人目学年}}", grade_sixth);

  // 7人目の情報を置換
  body.replaceText("{{七人目役職}}", post_seventh);
  body.replaceText("{{七人目氏名}}", name_seventh);
  body.replaceText("{{七人目学科}}", department_seventh);
  body.replaceText("{{七人目学年}}", grade_seventh);

  // 8人目の情報を置換
  body.replaceText("{{八人目役職}}", post_eighth);
  body.replaceText("{{八人目氏名}}", name_eighth);
  body.replaceText("{{八人目学科}}", department_eighth);
  body.replaceText("{{八人目学年}}", grade_eighth);

  // 販売物概要の情報を置換
  body.replaceText("{{販売物の情報}}", sale_information);
  body.replaceText("{{備考}}", memo);

  // 画像の縦横比を取得
  let res = ImgApp.getSize(attachImg);
  let width = res.width;
  let height = res.height;
  // 画像を横300pxでアスペクト比を揃えて大きさを編集し最終行へ挿入
  body
    .appendImage(attachImg)
    .setWidth(300)
    .setHeight((300 * height) / width);

  // 2枚目の画像が存在する場合、挿入
  if (attachImg2) {
    let res2 = ImgApp.getSize(attachImg2);
    let width2 = res2.width;
    let height2 = res2.height;
    body
      .appendImage(attachImg2)
      .setWidth(300)
      .setHeight((300 * height2) / width2);
  }

  // 3枚目の画像が存在する場合、挿入
  if (attachImg3) {
    let res3 = ImgApp.getSize(attachImg3);
    let width3 = res3.width;
    let height3 = res3.height;
    body
      .appendImage(attachImg3)
      .setWidth(300)
      .setHeight((300 * height3) / width3);
  }

  // ドキュメントを保存して閉じる
  doc.saveAndClose();

  // 新たにスプレッドシートをコピー
  const spreadsheetTemplateId = getProps().getProperty("TEMPLATE_SHEET_ID");
  const spreadsheetTemplate = DriveApp.getFileById(spreadsheetTemplateId);
  const newSpreadsheet = spreadsheetTemplate.makeCopy(
    club_name + "_持ち物リストと予定表",
    newFolder
  );

  // スプレッドシートの共有設定を変更
  newSpreadsheet.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.EDIT
  );

  // 新しく作成したドキュメントとスプレッドシートのURLを取得
  const newDocUrl = "https://docs.google.com/document/d/" + newDocId;
  const newSpreadsheetUrl =
    "https://docs.google.com/spreadsheets/d/" + newSpreadsheet.getId();

  // このスクリプト自体が動いているフォーム連携スプレッドシートの2枚目（リマインド）シートに書き込み
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reminderSheet = spreadsheet.getSheetByName("リマインド");

  // 日付を操作するために、start_dateをDateオブジェクトに変換
  let startDay = new Date(start_date);
  let twoDaysBefore = new Date(startDay);
  let oneWeekBefore = new Date(startDay);

  // 二日前と一週間前の日付を計算
  twoDaysBefore.setDate(startDay.getDate() - 2);
  oneWeekBefore.setDate(startDay.getDate() - 7);

  let formattedStartDay = formatDate(startDay);
  let formattedTwoDaysBefore = formatDate(twoDaysBefore);
  let formattedOneWeekBefore = formatDate(oneWeekBefore);

  // 新しい行にデータを書き込む
  reminderSheet.appendRow([
    club_name, // ① 団体名（識別の基本）
    representative_name, // ② 代表者名（連絡相手）

    start_time, // ③ 開始時刻（時系列要素その1）
    formattedStartDay, // ④ 開始日（中心日程）
    formattedTwoDaysBefore, // ⑤ 二日前（通知用）
    formattedOneWeekBefore, // ⑥ 一週間前（通知用）
    end_date, // ⑦ 終了日（後続処理に必要）

    newSpreadsheetUrl, // ⑧ スプレッドシートURL（参照用）
    contact_mail, // ⑨ 連絡先メールアドレス（通知に使用）
  ]);

  // 一時的にスクリプトを停止
  Utilities.sleep(10000);

  // 作成したドキュメントをPDFに変換し、フォルダに格納
  const pdfBlob = DriveApp.getFileById(newDocId).getAs("application/pdf");
  const pdfFileName = "キッチンカー利用申込書(団体用) - " + club_name + ".pdf";
  const pdfFile = newFolder.createFile(pdfBlob).setName(pdfFileName);

  // メール本文を生成
  const subject = "キッチンカー出店に関して： " + club_name;
  const emailBody = `就職・生活支援課
  ご担当者様 (cc: ${club_name} ${representative_name}さん)

  お世話になっております。福井県立大学 CSサークルです。
  
  この度、キッチンカー利用の申請をさせていただきたく、ご連絡いたしました。

  今回、利用を希望している団体は${club_name}です。（代表者: ${representative_name}さん）

  詳細につきましては、添付の資料をご確認いただけますと幸いです。

  また、ご不明点やご質問がございましたら、このメールへの返信にてお知らせください。CSサークルおよび${club_name}の代表である${representative_name}さんが対応させていただきます。

  よろしくお願いいたします。

  CSサークル`;

  MailApp.sendEmail({
    to: getProps().getProperty("UNIVERSITY_EMAIL"), // 就職・生活支援課のメールアドレスに変更
    cc: contact_mail,
    subject: subject,
    body: emailBody,
    attachments: [pdfFile], // ここでPDFファイルを添付
  });

  // メールの件名と本文をまとめて定義
  const newEmailSubject = `タイムスケジュール及び持ち物チェックリスト作成のお願い： ${club_name}`;
  const newEmailBody = `${club_name}
  ${representative_name}さん （cc:柴田代表）

  こんにちは、CSサークルです。

  資料作成フォームにご記入いただき、ありがとうございます。

  タイムスケジュール及び持ち物チェックリストの作成をお願いしたく、ご連絡いたしました。つきましては、以下のスプレッドシートに直接ご記入くださいますようお願い申し上げます。

  スプレッドシートリンク: ${newSpreadsheetUrl}

  記入が完了しましたら、このメールへの返信にてお知らせください。

  参考までに、就職・生活支援課に提出したPDFを添付いたします。

  また、キッチンカーのご利用に際してご質問がある場合も、このメールへの返信にてお問い合わせください。CSサークルおよびキッチンカー利用の責任者である柴田代表が対応させていただきます。

  よろしくお願いいたします。

  CSサークル`;

  // 新しいメールを送信
  MailApp.sendEmail({
    to: contact_mail,
    cc: getProps().getProperty("ADMINISTRATOR_EMAIL"), //柴田代表のメールアドレス
    subject: newEmailSubject,
    body: newEmailBody,
    attachments: [pdfFile],
  });
}

function club_sendFormEmail(email, clubName, representativeName) {
  // スクリプトプロパティからフォームURLを取得
  const club_formUrl = getProps().getProperty("CLUB_SUBMISSION_FORM_URL");

  var bodyText = `
  ${clubName}
  ${representativeName}さん

  こんにちは、CSサークルです。

  キッチンカーでの販売、お疲れさまでした。

  今回のイベントの結果を学校側に報告するため、以下のフォームへの記入をお願いします。

  フォームリンク：${club_formUrl}

  このフォームに記載すると、自動的にドキュメントが生成され、書類が提出されます。

  ご質問や不明点がございましたら、いつでもご連絡ください。

  よろしくお願いいたします。

  CSサークル
  `;

  // HTMLメールとして送信
  var htmlBody = bodyText.replace(/\n/g, "<br>");

  MailApp.sendEmail({
    to: email,
    subject: "【重要】提出書類作成用フォームへの記入のお願い",
    body: bodyText, // プレーンテキストの本文
    htmlBody: htmlBody, // HTML形式の本文
  });
}

function areDatesEqual(date1, date2) {
  return date1.toDateString() === date2.toDateString();
}

// --------------------------------------------------------------------------------------------------

function sendSlackNotification(mail, name, club_name, row) {
  const webhookUrl = getProps().getProperty("WEBHOOK_URL"); // Slack専用スクリプトのURL

  const payload = {
    type: "external_notification",
    mail,
    name,
    club_name,
    row,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(webhookUrl, options); // Slack専用スクリプトへ通知
  Logger.log("Slack通知レスポンス: " + response.getContentText());

  try {
    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log("✅ Slack通知送信成功");
  } catch (e) {
    Logger.log("❌ Slack通知送信失敗: " + e.message);
  }
}

// --------------------------------------------共通------------------------------------------------

// スクリプトプロパティ一覧表示
function logScriptProperties() {
  const props = getProps().getProperties();
  Logger.log("📋 現在のスクリプトプロパティ一覧:");
  for (const key in props) {
    Logger.log(`${key} = ${props[key]}`);
  }
}
