let cachedProps = null;

function getProps() {
  if (!cachedProps) {
    cachedProps = PropertiesService.getScriptProperties();
  }
  return cachedProps;
}

// 列7が空欄の場合は個人用処理
function onFormSubmit(e) {
  var club_or_individual = e.values[7]; // 列7（サークル名）が空欄の場合は個人用処理
  if (!club_or_individual || club_or_individual.trim() === "") {
    handleIndividualSubmission(e);
    Logger.log("個人用の通知を受け取りました");
  } else {
    handleClubSubmission(e);
    Logger.log("団体用の通知を受け取りました");
  }
}

// ---------------------------------------------------個人用------------------------------------------------------------------

function handleIndividualSubmission(e) {
  Logger.log("個人用の通知です");
  var values = e.values;
  var name = values[6];
  var rawScore = values[2];
  var score = parseInt(rawScore.split("/")[0].trim());
  var responsible = values[8]; // 食品衛生責任者を取得
  var quiz = values[50]; // ひっかけクイズを取得
  var mail = values[1]; // メールアドレスを取得
  var row = e.range.getRow(); // 行番号を取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 条件分岐の処理
  if (score === 1 && responsible === "はい" && quiz === "") {
    // ★ 33点満点、responsibleが「はい」、quizが空欄の場合：元の処理を実行
    try {
      executeIndividualOriginalProcess(values, row);
    } catch (err) {
      Logger.log(
        "❌ executeIndividualOriginalProcess 実行中にエラー: " + err.message
      );
    }
  } else if (
    score === 33 &&
    responsible === "はい" &&
    quiz === "読みました。"
  ) {
    // ★ 33点満点、responsibleが「はい」、quizが「読みました。」の場合：行の色をグレーにし、不合格通知をメールで送信
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("gray");
    sendMail(
      mail,
      "仮申請審査結果：不合格",
      generateIndividualRejectionMessage("quiz_failed", name)
    );
  } else if (score === 33 && responsible === "いいえ") {
    // ★ 33点満点、responsibleが「いいえ」の場合：行の色をグレーにし、不合格通知をメールで送信
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("gray");
    sendMail(
      mail,
      "仮申請審査結果：不合格",
      generateIndividualRejectionMessage("no_responsible", name)
    );
  } else if (score < 33 && responsible === "はい") {
    // ★ 33点未満、responsibleが「はい」の場合：行の色をグレーにし、不合格通知をメールで送信
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("gray");
    sendMail(
      mail,
      "仮申請審査結果：不合格",
      generateIndividualRejectionMessage("low_score", name)
    );
  } else if (score < 33 && responsible === "いいえ") {
    // ★ 33点未満、responsibleが「いいえ」の場合：行の色をグレーにし、不合格通知をメールで送信
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("gray");
    sendMail(
      mail,
      "仮申請審査結果：不合格",
      generateIndividualRejectionMessage("low_score_no_responsible", name)
    );
  }
}

// 元の処理を実行する関数
function executeIndividualOriginalProcess(values, row) {
  var individual = " (個人利用希望) ";

  var stamp = values[0]; // タイムスタンプを取得
  var mail = values[1]; // メールアドレスを取得
  var score = values[2]; // 点数を取得
  var faculty = values[3]; // 学部を取得
  var department = values[4]; // 学科を取得
  var grade = values[5]; // 学年を取得
  var name = values[6]; // 名前を取得
  var club_name = values[7]; //サークル名を取得
  var responsible = values[8]; // 食品衛生責任者を取得
  var purpose = values[9]; // 目的を取得
  var hygiene = values[10]; // 衛生管理を取得
  var start_day = values[11]; // 出店日を取得
  var befor_preparation = values[12]; // 前日準備を取得
  var sale_image = values[13]; // 販売物の写真を取得
  var sale_image2 = values[14]; // 販売物の写真を取得
  var sale_image3 = values[15]; // 販売物の写真を取得
  var information = values[16]; // 販売物の情報を取得

  var quiz = values[50]; // ひっかけクイズを取得
  var memo = values[51]; // 備考を取得

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

  // テンプレートドキュメントのIDを指定
  const templateDocId = getProps().getProperty("TEMPLATE_DOC_ID");
  const templateDoc = DriveApp.getFileById(templateDocId);
  const newDoc = templateDoc.makeCopy("一次選考合格者 - " + name);
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  // ドキュメントの内容を置換
  body.replaceText("{{点数}}", score);
  body.replaceText("{{学部}}", faculty);
  body.replaceText("{{学科}}", department);
  body.replaceText("{{学年}}", grade);
  body.replaceText("{{名前}}", name);
  body.replaceText("{{サークル名}}", individual);
  body.replaceText("{{メールアドレス}}", mail);
  body.replaceText("{{食品衛生責任者}}", responsible);
  body.replaceText("{{目的}}", purpose);
  body.replaceText("{{衛生管理}}", hygiene);
  body.replaceText("{{出店日}}", start_day);
  body.replaceText("{{前日準備}}", befor_preparation);
  body.replaceText("{{販売物情報}}", information);
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

  // 生成されたドキュメントを指定フォルダに移動
  const destinationFolderId = getProps().getProperty("DESTINATION_FOLDER_ID");
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  DriveApp.getFileById(newDocId).moveTo(destinationFolder);

  // 新しく作成したドキュメントのURLを取得
  const newDocUrl = "https://docs.google.com/document/d/" + newDocId;

  // ドキュメントの共有設定
  DriveApp.getFileById(newDocId).setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.EDIT
  );

  // Slack通知メッセージの作成
  const message =
    `<!channel>\n` +
    `*${name}* さんがキッチンカーの仮申請に *合格* しました！以下のURLから情報を確認し、リアクションスタンプで対応をお願いします。\n\n` +
    `📌 *【△△】* \n` +
    `　✅ 承認 → :cs_マル:\n` +
    `　❌ 却下 → :cs_バツ:\n\n` +
    `📌 *【ooさん】* \n` +
    `　✅ 承認 → :管理者_マル:\n` +
    `　❌ 却下 → :管理者_バツ:\n\n` +
    `────────────────────────────────────\n` +
    `✅ *「:cs_マル:」と「:管理者_マル:」の両方が押された場合* → *承認処理* が実行され、申請者に *合格通知メール* が送られます。\n\n` +
    `❌ *「:cs_バツ:」または「:管理者_バツ:」が押された場合* → *却下処理* が実行されます。（申請者にはメールが送られません）\n\n` +
    `💡 判断に迷った場合は「却下スタンプ（:cs_バツ: or :管理者_バツ:）」を押し、手動で ${name} さん ( メールアドレス： ${mail} ) に確認を取りましょう。\n` +
    `────────────────────────────────────\n\n` +
    `🔍 *${name} さんの詳細はこちら：* \n` +
    `${newDocUrl}`;
  // Slackに通知を送信
  sendSlackNotification(message, mail, name, "", row);

  // 一時的にスクリプトを停止
  Utilities.sleep(10000);
}

// 条件に応じたメールの本文を生成
// 条件に応じたメールの本文を生成
function generateIndividualRejectionMessage(caseType, name) {
  var greeting = name + "様\n\n";

  switch (caseType) {
    case "quiz_failed":
      return (
        greeting +
        "この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございます。\n\n" +
        "残念ながら、選考の結果不合格となりました。\n" +
        "再挑戦を希望される場合は、問題文をよくお読みのうえ、再度ご応募いただけますと幸いです。\n\n" +
        "どうぞよろしくお願いいたします\n\n" +
        "△△"
      );

    case "no_responsible":
      return (
        greeting +
        "この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございます。\n\n" +
        "残念ながら、食品衛生責任者の資格をお持ちでないため、選考の結果不合格となりました。\n" +
        "資格を取得された後、再度ご応募いただけますと幸いです。\n\n" +
        "どうぞよろしくお願いいたします\n\n" +
        "△△"
      );

    case "low_score":
      return (
        greeting +
        "この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございます。\n\n" +
        "残念ながら、得点が基準に達していないため、選考の結果不合格となりました。\n" +
        "再挑戦を希望される場合は、マニュアルなどを参考にし、再度ご応募いただけますと幸いです。\n\n" +
        "どうぞよろしくお願いいたします\n\n" +
        "△△"
      );

    case "low_score_no_responsible":
      return (
        greeting +
        "この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございます。\n\n" +
        "食品衛生責任者の資格をお持ちでないこと、また得点が基準に達していないことを考慮した結果、残念ながら不合格となりました。\n" +
        "資格を取得し、マニュアルなどをよくお読みの上、再挑戦していただけますと幸いです。\n\n" +
        "どうぞよろしくお願いいたします\n\n" +
        "△△"
      );

    default:
      return "";
  }
}

// -----------------------------------------------団体用------------------------------------------------------

function handleClubSubmission(e) {
  var values = e.values;
  var name = values[6];
  var rawScore = values[2];
  var score = parseInt(rawScore.split("/")[0].trim());
  var responsible = values[8]; // 食品衛生責任者を取得
  var quiz = values[50]; // ひっかけクイズを取得
  var mail = values[1]; // メールアドレスを取得
  var club_name = values[7]; // サークル名を取得
  var row = e.range.getRow(); // 行番号を取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 条件分岐の処理
  if (score === 1 && responsible === "はい" && quiz === "") {
    // ★ 33点満点、responsibleが「はい」、quizが空欄の場合：元の処理を実行
    try {
      executeClubOriginalProcess(values, row);
    } catch (err) {
      Logger.log(
        "❌ executeClubOriginalProcess 実行中にエラー: " + err.message
      );
    }
  } else if (score === 1 && responsible === "はい" && quiz === "読みました。") {
    // ★ 33点満点、responsibleが「はい」、quizが「読みました。」の場合：行の色をグレーにし、不合格通知をメールで送信
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("gray");
    sendMail(
      mail,
      "仮申請審査結果：不合格",
      generateClubRejectionMessage("quiz_failed", club_name, name)
    );
  } else if (score === 1 && responsible === "いいえ") {
    // ★ 33点満点、responsibleが「いいえ」の場合：行の色をグレーにし、不合格通知をメールで送信
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("gray");
    sendMail(
      mail,
      "仮申請審査結果：不合格",
      generateClubRejectionMessage("no_responsible", club_name, name)
    );
  } else if (score < 1 && responsible === "はい") {
    // ★ 33点未満、responsibleが「はい」の場合：行の色をグレーにし、不合格通知をメールで送信
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("gray");
    sendMail(
      mail,
      "仮申請審査結果：不合格",
      generateClubRejectionMessage("low_score", club_name, name)
    );
  } else if (score < 1 && responsible === "いいえ") {
    // ★ 33点未満、responsibleが「いいえ」の場合：行の色をグレーにし、不合格通知をメールで送信
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("gray");
    sendMail(
      mail,
      "仮申請審査結果：不合格",
      generateClubRejectionMessage("low_score_no_responsible", club_name, name)
    );
  }
}

// 元の処理を実行する関数
function executeClubOriginalProcess(values, row) {
  var stamp = values[0]; // タイムスタンプを取得
  var mail = values[1]; // メールアドレスを取得
  var score = values[2]; // 点数を取得
  var faculty = values[3]; // 学部を取得
  var department = values[4]; // 学科を取得
  var grade = values[5]; // 学年を取得
  var name = values[6]; // 名前を取得
  var club_name = values[7]; //サークル名を取得
  var responsible = values[8]; // 食品衛生責任者を取得
  var purpose = values[9]; // 目的を取得
  var hygiene = values[10]; // 衛生管理を取得
  var start_day = values[11]; // 出店日を取得
  var befor_preparation = values[12]; // 前日準備を取得
  var sale_image = values[13]; // 販売物の写真を取得
  var sale_image2 = values[14]; // 販売物の写真を取得
  var sale_image3 = values[15]; // 販売物の写真を取得
  var information = values[16]; // 販売物の情報を取得

  var quiz = values[50]; // ひっかけクイズを取得
  var memo = values[51]; // 備考を取得

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

  // テンプレートドキュメントのIDを指定
  const templateDocId = getProps().getProperty("TEMPLATE_DOC_ID");
  const templateDoc = DriveApp.getFileById(templateDocId);
  const newDoc = templateDoc.makeCopy("一次選考合格者 - " + club_name);
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();

  // ドキュメントの内容を置換
  body.replaceText("{{点数}}", score);
  body.replaceText("{{サークル名}}", club_name);
  body.replaceText("{{学部}}", faculty);
  body.replaceText("{{学科}}", department);
  body.replaceText("{{学年}}", grade);
  body.replaceText("{{名前}}", name);
  body.replaceText("{{メールアドレス}}", mail);
  body.replaceText("{{食品衛生責任者}}", responsible);
  body.replaceText("{{目的}}", purpose);
  body.replaceText("{{衛生管理}}", hygiene);
  body.replaceText("{{出店日}}", start_day);
  body.replaceText("{{前日準備}}", befor_preparation);
  body.replaceText("{{販売物情報}}", information);
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

  // 生成されたドキュメントを指定フォルダに移動
  const destinationFolderId = getProps().getProperty("DESTINATION_FOLDER_ID");
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  DriveApp.getFileById(newDocId).moveTo(destinationFolder);

  // 新しく作成したドキュメントのURLを取得
  const newDocUrl = "https://docs.google.com/document/d/" + newDocId;

  // ドキュメントの共有設定
  DriveApp.getFileById(newDocId).setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.EDIT
  );

  // Slack通知メッセージの作成
  const message =
    `<!channel>\n` +
    `*${club_name}* がキッチンカーの仮申請に *合格* しました！以下のURLから情報を確認し、リアクションスタンプで対応をお願いします。\n\n` +
    `📌 *【△△】* \n` +
    `　✅ 承認 → :cs_マル:\n` +
    `　❌ 却下 → :cs_バツ:\n\n` +
    `📌 *【ooさん】* \n` +
    `　✅ 承認 → :管理者_マル:\n` +
    `　❌ 却下 → :管理者_バツ:\n\n` +
    `────────────────────────────────────\n` +
    `✅ *「:cs_マル:」と「:管理者_マル:」の両方が押された場合* → *承認処理* が実行され、申請者に *合格通知メール* が送られます。\n\n` +
    `❌ *「:cs_バツ:」または「:管理者_バツ:」が押された場合* → *却下処理* が実行されます。（申請者にはメールが送られません）\n\n` +
    `💡 判断に迷った場合は「却下スタンプ（:cs_バツ: or :管理者_バツ:）」を押し、手動で ${club_name} の ${name} さん ( メールアドレス： ${mail} ) に確認を取りましょう。\n` +
    `────────────────────────────────────\n\n` +
    `🔍 *${club_name} の詳細はこちら：* \n` +
    `${newDocUrl}`;
  // Slackに通知を送信
  sendSlackNotification(message, mail, name, club_name, row);

  // 一時的にスクリプトを停止
  Utilities.sleep(10000);
}

// 条件に応じたメールの本文を生成
// 条件に応じたメールの本文を生成 (club_nameを追加)
function generateClubRejectionMessage(caseType, club_name, name) {
  var greeting = club_name + "\n" + name + "様\n\n";

  switch (caseType) {
    case "quiz_failed":
      return (
        greeting +
        "この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございます。\n\n" +
        "残念ながら、選考の結果不合格となりました。\n" +
        "再挑戦を希望される場合は、問題文をよくお読みのうえ、再度ご応募いただけますと幸いです。\n\n" +
        "どうぞよろしくお願いいたします\n\n" +
        "△△"
      );

    case "no_responsible":
      return (
        greeting +
        "この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございます。\n\n" +
        "残念ながら、食品衛生責任者の資格をお持ちでないため、選考の結果不合格となりました。\n" +
        "資格を取得された後、再度ご応募いただけますと幸いです。\n\n" +
        "どうぞよろしくお願いいたします\n\n" +
        "△△"
      );

    case "low_score":
      return (
        greeting +
        "この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございます。\n\n" +
        "残念ながら、得点が基準に達していないため、選考の結果不合格となりました。\n" +
        "再挑戦を希望される場合は、マニュアルなどを参考にし、再度ご応募いただけますと幸いです。\n\n" +
        "どうぞよろしくお願いいたします\n\n" +
        "△△"
      );

    case "low_score_no_responsible":
      return (
        greeting +
        "この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございます。\n\n" +
        "食品衛生責任者の資格をお持ちでないこと、また得点が基準に達していないことを考慮した結果、残念ながら不合格となりました。\n" +
        "資格を取得し、マニュアルなどをよくお読みの上、再挑戦していただけますと幸いです。\n\n" +
        "どうぞよろしくお願いいたします\n\n" +
        "△△"
      );

    default:
      return "";
  }
}

// ------------------------------------------------共通--------------------------------------------------------------

function sendSlackNotification(message, mail, name, club_name, row) {
  Logger.log("📨 Slack通知関数が呼ばれました");
  Logger.log("🔤 message: " + message);

  const webhookUrl = getProps().getProperty("WEBHOOK_URL"); // Slack専用スクリプトのURL

  const payload = {
    type: "external_notification",
    message,
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

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    Logger.log("Slack通知レスポンス: " + response.getContentText());
    Logger.log("✅ Slack通知送信成功");
  } catch (e) {
    Logger.log("❌ Slack通知送信失敗: " + e.message);
  }
}

// メール送信を行う関数
function sendMail(to, subject, body) {
  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: body,
  });
}
