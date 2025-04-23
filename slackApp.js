/**
 * Slack Webhook受信用スクリプト（統合バージョン）
 * - スタンプ反応による承認/却下処理
 * - 外部スクリプトからの通知投稿
 * - SlackアプリのURL検証対応
 */

// キャッシュ用変数
let cachedProps = null;

// プロパティ取得関数
function getProps() {
  if (!cachedProps) {
    cachedProps = PropertiesService.getScriptProperties();
  }
  return cachedProps;
}

function getSlackToken() {
  return getProps().getProperty("SLACK_TOKEN");
}
const APPROVED_REACTIONS = ["cs_マル", "管理者_マル"];
const REJECTED_REACTIONS = ["cs_バツ", "管理者_バツ"];
const MESSAGES_KEY = "TARGET_MESSAGES"; // 全監視対象メッセージリスト

function getSheetByNameKari(sheetName) {
  const sheetId = getProps().getProperty("SPREADSHEET_ID_KARI");
  return SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
}

function getSheetByNameHon(sheetName) {
  const sheetId = getProps().getProperty("SPREADSHEET_ID_HON");
  return SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    Logger.log("📩 Slackからの受信データ: " + JSON.stringify(data));

    // -------------------------------
    // 🔹 SlackのURL検証（最初の登録時）
    // -------------------------------
    if (data.type === "url_verification") {
      Logger.log("✅ Slack URL確認成功");
      return ContentService.createTextOutput(data.challenge).setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    // -------------------------------
    // 🔹 外部スクリプトからの通知投稿
    // -------------------------------
    if (data.type === "external_notification") {
      const { message, mail, name, club_name, row } = data;

      if (!message || !mail || !name || typeof row === "undefined") {
        Logger.log("⚠️ external_notificationに必要なデータが不足しています");
        return ContentService.createTextOutput("Missing parameters");
      }

      // ▼ 処理の分岐（club_nameが未定義 or 空文字なら個人用）
      if (club_name === "") {
        Logger.log("📨 個人申請としてSlackに通知を投稿します");
        postIndividualSlackMessage(message, mail, name, row); // ← 新しい関数に分けてもOK
      } else {
        Logger.log("📨 団体申請としてSlackに通知を投稿します");
        postClubSlackMessage(message, mail, name, club_name, row);
      }

      return ContentService.createTextOutput("Slack投稿完了").setMimeType(
        ContentService.MimeType.TEXT
      );
    }

    // -------------------------------
    // 🔹 Slackからのイベント（リアクションなど）
    // -------------------------------
    if (data.type === "event_callback") {
      const event = data.event;

      Logger.log("📝 イベントタイプ: " + event.type);

      // reaction に対する処理
      if (
        event.type === "reaction_added" ||
        event.type === "reaction_removed"
      ) {
        Logger.log(
          `🎯 リアクション ${event.type}: ${event.reaction} by ${event.user}`
        );
        upsertReaction(event);
        return ContentService.createTextOutput("OK").setMimeType(
          ContentService.MimeType.TEXT
        );
      }

      // 初回DM挨拶----------------
      if (event && event.type === "team_join") {
        const userId = event.user.id;
        sendWelcomeMessage(userId);
      }
      // -------------------------

      // 他のイベントタイプがあればここに追加
      Logger.log("ℹ️ 未処理のイベントタイプ: " + event.type);
      return ContentService.createTextOutput(
        "Unhandled event type"
      ).setMimeType(ContentService.MimeType.TEXT);
    }

    // -------------------------------
    // 🔹 未対応のリクエストタイプ
    // -------------------------------
    Logger.log("⚠️ 未対応のリクエスト: " + JSON.stringify(data));
    return ContentService.createTextOutput(
      "Unsupported request type"
    ).setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    Logger.log("❌ doPost エラー: " + err.toString());
    return ContentService.createTextOutput("error").setMimeType(
      ContentService.MimeType.TEXT
    );
  }
}

function evaluateMessageStatus(channel, ts, stored) {
  if (!stored.hasOwnProperty("club_name") || stored.club_name.trim() === "") {
    return evaluateIndividualSubmission(channel, ts, stored);
  } else {
    return evaluateClubSubmission(channel, ts, stored);
  }
}

function isToday(date) {
  const today = new Date(); // 現在の日付・時刻を取得（例: 2025年4月10日 22:00 など）

  return (
    date instanceof Date && // dateがちゃんとDateオブジェクトであることを確認
    date.getFullYear() === today.getFullYear() && // 年が同じか
    date.getMonth() === today.getMonth() && // 月が同じか（0〜11で表現される）
    date.getDate() === today.getDate() // 日付が同じか
  );
}

// slackへのリマインドメッセージ
function sendReminderToSlack() {
  const sheet = getSheetByNameHon("リマインド");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues(); // 1行目は見出しなので除く
  const today = new Date();

  data.forEach((row, index) => {
    const clubName = row[0]; // A列（空欄なら個人）
    const name = row[1]; // B列（代表者名）
    const startTime = row[2]; // C列
    const startDate = row[3]; // D列
    const twoDaysBefore = row[4]; // E列
    const oneWeekBefore = row[5]; // F列
    const endDate = row[6]; // G列
    const sheetUrl = row[7]; // H列
    const email = row[8]; // I列

    const rowIndex = index + 2; // 実際の行番号

    if (isToday(twoDaysBefore) || isToday(oneWeekBefore)) {
      if (!clubName || clubName.trim() === "") {
        // 👤 個人用通知
        sendIndividualReminderToSlack(
          name,
          startDate,
          startTime,
          sheetUrl,
          email,
          rowIndex
        );
      } else {
        // 🏢 団体用通知
        sendClubReminderToSlack(
          clubName,
          name,
          startDate,
          startTime,
          sheetUrl,
          email,
          rowIndex
        );
      }
    }
  });
}

function remindUnprocessedMessages() {
  const keys = JSON.parse(getProps().getProperty(MESSAGES_KEY) || "[]");

  let hasClub = false;
  let hasIndividual = false;

  for (const key of keys) {
    const storedStr = getProps().getProperty(key);
    if (!storedStr) continue;

    const stored = JSON.parse(storedStr);
    const status = stored.status;

    if (status === "approved" || status === "rejected") continue;

    const isIndividual = !stored.club_name || stored.club_name.trim() === "";
    const isClub = stored.club_name && stored.club_name.trim() !== "";

    if (isIndividual) {
      hasIndividual = true;
    } else if (isClub) {
      hasClub = true;
    }

    // 最適化：両方確認できたら早期終了
    if (hasIndividual && hasClub) break;
  }

  if (hasIndividual && !hasClub) {
    Logger.log("✅ 個人リマインド 実行完了");
    remindIndividualUnprocessedMessages();
  } else if (hasClub && !hasIndividual) {
    Logger.log("✅ 団体リマインド 実行完了");
    remindClubUnprocessedMessages();
  } else if (hasIndividual && hasClub) {
    Logger.log("✅ 両方リマインド 実行開始");
    remindIndividualUnprocessedMessages();
    remindClubUnprocessedMessages();
    Logger.log("✅ 両方リマインド 実行完了");
  } else {
    Logger.log("ℹ️ 未処理の申請はありませんでした");
  }

  Logger.log("✅ remindUnprocessedMessages 実行完了");
}

// 個人用----------------------------------------------------------------------------------------------

function evaluateIndividualSubmission(channel, ts, stored) {
  const reactions = stored.reactions || [];
  const messageKey = `${channel}_${ts}`;
  const approved = APPROVED_REACTIONS.every((r) => reactions.includes(r));
  const rejected = REJECTED_REACTIONS.some((r) => reactions.includes(r));

  if (approved) {
    postToSlack(
      channel,
      ts,
      `✅ 承認されました！${stored.mail} に承認されたことを通知しました！`
    );
    stored.status = "approved";
    getProps().setProperty(messageKey, JSON.stringify(stored));

    // ★ スプレッドシートの行を緑に
    const row = parseInt(stored.row);
    const sheet = getSheetByNameKari("フォームの回答 1");
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#d9ead3");

    // ★ メール送信
    if (stored.mail && stored.name) {
      const administrator_email = getProps().getProperty("ADMINISTRATOR_EMAIL");
      const formUrl = getProps().getProperty("FOLLOWUP_FORM_URL");

      const to = stored.mail;
      const subject = "【重要】キッチンカー利用承認のお知らせ";
      const body =
        `${stored.name}さん（※このメールは(施設責任者氏名)代表にもccで送付しています）\nこんにちは、ooサークルです。\n\n` +
        `この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございました。\n\n` +
        `審査の結果、キッチンカーのご利用が【承認】されましたので、お知らせいたします。\n\n` +
        `──────────────────────\n` +
        `■ 今後のご対応について\n` +
        `──────────────────────\n\n` +
        `①学校への提出資料について\n` +
        `下記のフォームにご回答ください。\n` +
        `＊回答期限：販売日の2週間前まで\n\n` +
        `フォームURL：${formUrl}\n\n` +
        `※提出が遅れると、学校側から出店を認められない場合がありますのでご注意ください。\n\n` +
        `② 前日準備について（厨房利用の連絡）\n` +
        `前日に仕込みを希望される場合は、学食の厨房をご利用いただきます。\n` +
        `その際、キッチンカー運営責任者である(施設責任者氏名)さんに、事前にご連絡・ご相談をお願いいたします。\n` +
        `ご返信の際は、冒頭に「(施設責任者氏名)様」などの宛名をご記載いただきますようお願いいたします。\n\n` +
        `【確認方法】\n` +
        `・方法①：「全員に返信」で、このメールにご返信ください（ccに(施設責任者氏名)さんが含まれています）\n` +
        `・方法②：(施設責任者氏名)代表のメールアドレスに直接ご連絡ください\n` +
        `　▷ ${administrator_email}\n\n` +
        `ご不明な点がありましたら、本メールにご返信いただくか、ooサークルまでお気軽にご連絡ください。\n\n` +
        `今後とも、どうぞよろしくお願いいたします。\n\n` +
        `ooサークル`;

      MailApp.sendEmail({
        to,
        cc: administrator_email,
        subject,
        body,
      });
    }
    return;
  }

  if (rejected) {
    postToSlack(
      channel,
      ts,
      `❌ この申請は却下されました。却下理由は何ですか？まとまったら ${stored.name}さん (メールアドレス： ${stored.mail} ) にその内容を伝えましょう！`
    );
    stored.status = "rejected";
    getProps().setProperty(messageKey, JSON.stringify(stored));

    // ★ スプレッドシートの行を赤に
    const row = parseInt(stored.row);
    const sheet = getSheetByNameKari("フォームの回答 1");
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#f4cccc");
  }
}

// ▼ リマインド処理：未処理メッセージに毎日通知
function remindIndividualUnprocessedMessages() {
  const keys = JSON.parse(getProps().getProperty(MESSAGES_KEY) || "[]");

  keys.forEach((key) => {
    const stored = JSON.parse(getProps().getProperty(key) || "{}");
    const clubName = (stored.club_name || "").trim();
    const isIndividual = clubName === "";
    const isPending =
      stored.status !== "approved" && stored.status !== "rejected";

    if (isIndividual && isPending) {
      const [channel, ts] = key.split("_");
      const message = `${stored.name} さん ( メールアドレス： ${stored.mail} ) の申請はまだ「承認」または「却下」されていません。スタンプで対応をお願いします！`;
      postReminderToThread(channel, ts, message);
    }
  });
}

// slackに通知を送る関数
// ▼ フォーム送信時：Slackにメッセージ投稿 & メッセージキーを保存
function postIndividualSlackMessage(message, mail, name, row) {
  const channel = getProps().getProperty("CHANNEL_ID");

  const url = "https://slack.com/api/chat.postMessage";
  const payload = {
    channel,
    text: message,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${getSlackToken()}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      Logger.log("❌ Slack投稿失敗: " + result.error);
    }

    if (result.ok) {
      const messageKey = `${result.channel}_${result.ts}`;

      // ✅ メッセージ一覧に追記
      const all = JSON.parse(getProps().getProperty(MESSAGES_KEY) || "[]");
      all.push(messageKey);
      getProps().setProperty(MESSAGES_KEY, JSON.stringify(all));

      // ✅ 初期データを保存（ここが重要！）
      getProps().setProperty(
        messageKey,
        JSON.stringify({
          type: "申請",
          mail,
          name,
          row,
          club_name: "",
        })
      );

      Logger.log("Slackに投稿し、監視対象に追加: " + messageKey);
    } else {
      Logger.log("Slack投稿エラー: " + result.error);
      Logger.log(`❌ Slack投稿失敗 (HTTP ${response.getResponseCode()})`);
      Logger.log(`📩 レスポンス本文: ${response.getContentText()}`);
      Logger.log(`🔤 エラー内容: ${result.error}`);
    }
  } catch (err) {
    Logger.log("❌ Slack通信エラー: " + err.message);
  }
}

function sendIndividualReminderToSlack() {
  Logger.log("🔔 リマインド処理を開始します");
  let messagesSet = new Set();

  const sheet = getSheetByNameHon("リマインド");
  var today = new Date();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  data.forEach(function (row, index) {
    var clubName = row[0];
    if (clubName && clubName.trim() !== "") return;
    var representativeName = row[1];
    var startTime = row[2];
    var startDayStr = row[3];
    var twoDaysBeforeStr = row[4];
    var oneWeekBeforeStr = row[5];
    var endDate = row[6];
    var sheetLink = row[7];
    var email = row[8];

    var startDay = new Date(startDayStr);
    var twoDaysBefore = new Date(twoDaysBeforeStr);
    var oneWeekBefore = new Date(oneWeekBeforeStr);

    Logger.log(`📅 チェック中: ${representativeName}, 開始日: ${startDayStr}`);

    if (areDatesEqual(twoDaysBefore, today)) {
      var message2Days = `:alarm_clock: リマインド：${representativeName}さんのキッチンカー利用予定日まであと二日です！！\n詳細はこちら：${sheetLink}`;
      messagesSet.add(message2Days);
      Logger.log(`✅ 二日前リマインド対象: ${representativeName}さん`);
      sendIndividualReminderEmail(
        email,
        "リマインド：キッチンカー利用予定日（2日前）",
        message2Days,
        sheetLink,
        representativeName
      );
    }

    if (areDatesEqual(oneWeekBefore, today)) {
      var message1Week = `:alarm_clock: リマインド：${representativeName}さんのキッチンカー利用予定日まであと一週間です！\n詳細はこちら：${sheetLink}`;
      messagesSet.add(message1Week);
      Logger.log(`✅ 一週間前リマインド対象: ${representativeName}`);
      sendIndividualReminderEmail(
        email,
        "リマインド：キッチンカー利用予定日（一週間前）",
        message1Week,
        sheetLink,
        representativeName
      );
    }
  });

  if (messagesSet.size > 0) {
    const slackMessage =
      "以下のタスクの通知があります。\n" + Array.from(messagesSet).join("\n");
    postSimpleSlackMessage(slackMessage);
    Logger.log("📤 Slackにリマインドを送信しました");
  } else {
    Logger.log("ℹ️ 本日のリマインド対象はありませんでした");
  }
}

function sendIndividualReminderEmail(
  email,
  subject,
  message,
  sheetLink,
  representativeName
) {
  var bodyText = `
  ${representativeName}さん

  こんにちは、ooサークルです。

  ${representativeName}さんのキッチンカー利用予定日が近づいております。あと少しで当日となりますね。 準備や確認事項がありましたら、ぜひこの機会にご確認ください。

  こちらのリンクから、最終調整用のスプレッドシートをご確認いただけます： ${sheetLink}

  このメールは自動送信されていますので、返信はご遠慮ください。 ご質問がある場合は、先日お送りいたしました最終調整用のメールにご返信いただけますよう、お願いいたします。

  当日が楽しいイベントとなりますことを願っております。

  何卒よろしくお願い申し上げます。

  ooサークル`;

  var htmlBody = bodyText.replace(/\n/g, "<br>");

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: bodyText,
    htmlBody: htmlBody,
  });
}

// 団体用----------------------------------------------------------------------------------------------

function evaluateClubSubmission(channel, ts, stored) {
  const reactions = stored.reactions || [];
  const messageKey = `${channel}_${ts}`;
  const approved = APPROVED_REACTIONS.every((r) => reactions.includes(r));
  const rejected = REJECTED_REACTIONS.some((r) => reactions.includes(r));

  if (approved) {
    postToSlack(
      channel,
      ts,
      `✅ 承認されました！${stored.mail} に承認されたことを通知しました！`
    );
    stored.status = "approved";
    getProps().setProperty(messageKey, JSON.stringify(stored));

    // ★ スプレッドシートの行を緑に
    const row = parseInt(stored.row);
    const sheet = getSheetByNameKari("フォームの回答 1");
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#d9ead3");

    // ★ メール送信
    if (stored.mail && stored.name && stored.club_name) {
      const administrator_email = getProps().getProperty("ADMINISTRATOR_EMAIL");
      const formUrl = getProps().getProperty("FOLLOWUP_FORM_URL");

      const to = stored.mail;
      const subject = "【重要】キッチンカー利用承認のお知らせ";
      const body =
        `${stored.club_name}\n${stored.name}さん（※このメールは(施設責任者氏名)代表にもccで送付しています）\nこんにちは、ooサークルです。\n\n` +
        `この度はキッチンカー利用の仮申請フォームにご回答いただき、誠にありがとうございました。\n\n` +
        `審査の結果、キッチンカーのご利用が【承認】されましたので、お知らせいたします。\n\n` +
        `──────────────────────\n` +
        `■ 今後のご対応について\n` +
        `──────────────────────\n\n` +
        `①学校への提出資料について\n` +
        `下記のフォームにご回答ください。\n` +
        `＊回答期限：販売日の2週間前まで\n\n` +
        `フォームURL：${formUrl}\n\n` +
        `※提出が遅れると、学校側から出店を認められない場合がありますのでご注意ください。\n\n` +
        `② 前日準備について（厨房利用の連絡）\n` +
        `前日に仕込みを希望される場合は、学食の厨房をご利用いただきます。\n` +
        `その際、キッチンカー運営責任者である(施設責任者氏名)さんに、事前にご連絡・ご相談をお願いいたします。\n` +
        `ご返信の際は、冒頭に「(施設責任者氏名)様」などの宛名をご記載いただきますようお願いいたします。\n\n` +
        `【確認方法】\n` +
        `・方法①：「全員に返信」で、このメールにご返信ください（ccに(施設責任者氏名)さんが含まれています）\n` +
        `・方法②：(施設責任者氏名)代表のメールアドレスに直接ご連絡ください\n` +
        `　▷ ${administrator_email}\n\n` +
        `ご不明な点がありましたら、本メールにご返信いただくか、ooサークルまでお気軽にご連絡ください。\n\n` +
        `今後とも、どうぞよろしくお願いいたします。\n\n` +
        `ooサークル`;

      MailApp.sendEmail({
        to,
        cc: administrator_email,
        subject,
        body,
      });
    }
    return;
  }

  if (rejected) {
    postToSlack(
      channel,
      ts,
      `❌ この申請は却下されました。却下理由は何ですか？まとまったら ${stored.club_name}の${stored.name}さん (メールアドレス： ${stored.mail} ) にその内容を伝えましょう！`
    );
    stored.status = "rejected";
    getProps().setProperty(messageKey, JSON.stringify(stored));

    // ★ スプレッドシートの行を赤に
    const row = parseInt(stored.row);
    const sheet = getSheetByNameKari("フォームの回答 1");
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#f4cccc");
  }
}

// ▼ リマインド処理：未処理メッセージに毎日通知
function remindClubUnprocessedMessages() {
  const keys = JSON.parse(getProps().getProperty(MESSAGES_KEY) || "[]");

  keys.forEach((key) => {
    const stored = JSON.parse(getProps().getProperty(key) || "{}");

    const clubName = (stored.club_name || "").trim();
    const isClub = clubName !== "";
    const isPending =
      stored.status !== "approved" && stored.status !== "rejected";

    if (isClub && isPending) {
      const [channel, ts] = key.split("_");
      const message = `この申請はまだ「承認」または「却下」されていません。スタンプで対応をお願いします！もし承認するのに不安があるようであれば、「却下」ボタンを押して、手動で ${stored.club_name} の ${stored.name} さん ( メールアドレス： ${stored.mail} ) まで確認メールを打ちましょう！`;
      postReminderToThread(channel, ts, message);
    }
  });
}

// slackに通知を送る関数
// ▼ フォーム送信時：Slackにメッセージ投稿 & メッセージキーを保存
function postClubSlackMessage(message, mail, name, club_name, row) {
  const channel = getProps().getProperty("CHANNEL_ID");

  const url = "https://slack.com/api/chat.postMessage";
  const payload = {
    channel,
    text: message,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${getSlackToken()}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      Logger.log("❌ Slack投稿失敗: " + result.error);
    }

    if (result.ok) {
      const messageKey = `${result.channel}_${result.ts}`;

      // ✅ メッセージ一覧に追記
      const all = JSON.parse(getProps().getProperty(MESSAGES_KEY) || "[]");
      all.push(messageKey);
      getProps().setProperty(MESSAGES_KEY, JSON.stringify(all));

      // ✅ 初期データを保存（ここが重要！）
      getProps().setProperty(
        messageKey,
        JSON.stringify({
          type: "申請",
          mail,
          name,
          club_name,
          row,
        })
      );

      Logger.log("Slackに投稿し、監視対象に追加: " + messageKey);
    } else {
      Logger.log("Slack投稿エラー: " + result.error);
      Logger.log(`❌ Slack投稿失敗 (HTTP ${response.getResponseCode()})`);
      Logger.log(`📩 レスポンス本文: ${response.getContentText()}`);
      Logger.log(`🔤 エラー内容: ${result.error}`);
    }
  } catch (err) {
    Logger.log("❌ Slack通信エラー: " + err.message);
  }
}

function sendClubReminderToSlack() {
  Logger.log("🔔 リマインド処理を開始します");
  let messagesSet = new Set();

  const sheet = getSheetByNameHon("リマインド");
  var today = new Date();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  data.forEach(function (row, index) {
    var clubName = row[0];
    var representativeName = row[1];
    var startTime = row[2];
    var startDayStr = row[3];
    var twoDaysBeforeStr = row[4];
    var oneWeekBeforeStr = row[5];
    var endDate = row[6];
    var sheetLink = row[7];
    var email = row[8];

    var startDay = new Date(startDayStr);
    var twoDaysBefore = new Date(twoDaysBeforeStr);
    var oneWeekBefore = new Date(oneWeekBeforeStr);

    Logger.log(`📅 チェック中: ${clubName}, 開始日: ${startDayStr}`);

    if (!clubName || clubName.trim() === "") return;

    if (areDatesEqual(twoDaysBefore, today)) {
      var message2Days = `:alarm_clock: リマインド：${clubName}のキッチンカー利用予定日まであと二日です！！\n詳細はこちら：${sheetLink}`;
      messagesSet.add(message2Days);
      Logger.log(`✅ 二日前リマインド対象: ${clubName}`);
      sendClubReminderEmail(
        email,
        "リマインド：キッチンカー利用予定日（2日前）",
        message2Days,
        sheetLink,
        representativeName,
        clubName
      );
    }

    if (areDatesEqual(oneWeekBefore, today)) {
      var message1Week = `:alarm_clock: リマインド：${clubName}のキッチンカー利用予定日まであと一週間です！\n詳細はこちら：${sheetLink}`;
      messagesSet.add(message1Week);
      Logger.log(`✅ 一週間前リマインド対象: ${clubName}`);
      sendClubReminderEmail(
        email,
        "リマインド：キッチンカー利用予定日（一週間前）",
        message1Week,
        sheetLink,
        representativeName,
        clubName
      );
    }
  });

  if (messagesSet.size > 0) {
    var slackMessage =
      "以下のタスクの通知があります。\n" + Array.from(messagesSet).join("\n");
    postSimpleSlackMessage(slackMessage);
    Logger.log("📤 Slackにリマインドを送信しました");
  } else {
    Logger.log("ℹ️ 本日のリマインド対象はありませんでした");
  }
}

function sendClubReminderEmail(
  email,
  subject,
  message,
  sheetLink,
  representativeName,
  clubName
) {
  var bodyText = `
  ${clubName}
  ${representativeName}さん

  こんにちは、ooサークルです。

  ${clubName}のキッチンカー利用予定日が近づいております。あと少しで当日となりますね。 準備や確認事項がありましたら、ぜひこの機会にご確認ください。

  こちらのリンクから、最終調整用のスプレッドシートをご確認いただけます： ${sheetLink}

  このメールは自動送信されていますので、返信はご遠慮ください。 ご質問がある場合は、先日お送りいたしました最終調整用のメールにご返信いただけますよう、お願いいたします。

  当日が楽しいイベントとなりますことを願っております。

  何卒よろしくお願い申し上げます。

  ooサークル`;

  var htmlBody = bodyText.replace(/\n/g, "<br>");

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: bodyText,
    htmlBody: htmlBody,
  });
}

// --------------------------------------------------------------------------------------------------------------------

function upsertReaction(event) {
  const { item, reaction, user, type } = event;
  const channel = item.channel;
  const ts = item.ts;
  const messageKey = `${channel}_${ts}`;

  // 対象スタンプ以外は無視
  const isTargetReaction =
    APPROVED_REACTIONS.includes(reaction) ||
    REJECTED_REACTIONS.includes(reaction);

  if (!isTargetReaction) return;

  const stored = JSON.parse(getProps().getProperty(messageKey) || "{}");
  if (!stored.reactions) stored.reactions = [];

  // GASが送信した「申請メッセージ」でなければ無視
  if (stored.type !== "申請") return;
  if (stored.status === "approved" || stored.status === "rejected") return;

  const index = stored.reactions.indexOf(reaction);

  if (type === "reaction_added") {
    if (index === -1) stored.reactions.push(reaction);
    postToSlack(
      channel,
      ts,
      `✅ <@${user}> さんが「:${reaction}:」リアクションを追加しました！`
    );
  } else if (type === "reaction_removed") {
    if (index !== -1) stored.reactions.splice(index, 1);
    postToSlack(
      channel,
      ts,
      `❎ <@${user}> さんが「:${reaction}:」リアクションを削除しました！`
    );
  }

  getProps().setProperty(messageKey, JSON.stringify(stored));

  evaluateMessageStatus(channel, ts, stored);
}

// Slackにスレッド返信を送信
function postToSlack(channel, thread_ts, text) {
  const url = "https://slack.com/api/chat.postMessage";
  const payload = {
    channel,
    thread_ts,
    text,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${getSlackToken()}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      Logger.log("Slack投稿エラー: " + result.error);
    }
  } catch (err) {
    Logger.log("GASエラー: " + err.message);
  }
}

function postReminderToThread(channel, ts, message) {
  const url = "https://slack.com/api/chat.postMessage";
  const payload = {
    channel: channel,
    thread_ts: ts,
    text: `:alarm_clock: リマインド：${message}`,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${getSlackToken()}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      Logger.log("⚠️ リマインド投稿失敗: " + result.error);
    }
  } catch (err) {
    Logger.log("❌ スレッド投稿エラー: " + err.message);
  }
}

function areDatesEqual(date1, date2) {
  return (
    date1.getFullYear() === date2.getFullYear() &&
    date1.getMonth() === date2.getMonth() &&
    date1.getDate() === date2.getDate()
  );
}

function postSimpleSlackMessage(message) {
  const url = "https://slack.com/api/chat.postMessage";
  const channel = getProps().getProperty("CHANNEL_ID");
  const payload = {
    channel,
    text: message,
  };
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${getSlackToken()}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      Logger.log("Slack投稿失敗: " + result.error);
    }
  } catch (e) {
    Logger.log("Slack投稿エラー: " + e.message);
  }
}

// 初回DM通知----------------------------
function sendWelcomeMessage(userId) {
  const token = getSlackToken(); // すでに共通関数があるので活用

  try {
    const imOpenResponse = UrlFetchApp.fetch(
      "https://slack.com/api/conversations.open",
      {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + token },
        payload: JSON.stringify({ users: userId }),
        muteHttpExceptions: true,
      }
    );

    const imData = JSON.parse(imOpenResponse.getContentText());
    if (!imData.ok) {
      Logger.log("❌ DMチャネル作成失敗: " + imData.error);
      return;
    }

    const channelId = imData.channel.id;
    const notionUrl = getProps().getProperty("WELCOME_GUIDE_URL"); // チャレサポくんガイド
    const portalUrl = getProps().getProperty("CIRCLE_PORTAL_URL"); // サークルポータル
    const meetingUrl = getProps().getProperty("MEETING_URL"); // ミーティングURL

    const welcomeText =
      `🎉 *ようこそ、ooサークルへ！*\n\n` +
      `こんにちは！皆さんのサークル活動をサポートする *「チャレサポくん」* です 🤖\n` +
      `これから活動がよりスムーズで楽しくなるように、お手伝いしていきます！\n\n` +
      `📝 *まずはこちらをご確認ください！*\n` +
      `・チャレサポくんの説明書（Notion）：\n${notionUrl}\n` +
      `・サークル活動ポータルサイト：\n${portalUrl}\n` +
      `・ミーティングURL(オンライン参加の場合はこちらから参加)：\n${meetingUrl}\n\n` +
      `気軽に頼ってくださいね！今後ともよろしくお願いします 🌱✨`;

    const messageResponse = UrlFetchApp.fetch(
      "https://slack.com/api/chat.postMessage",
      {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + token },
        payload: JSON.stringify({
          channel: channelId,
          text: welcomeText,
        }),
        muteHttpExceptions: true,
      }
    );

    const messageResult = JSON.parse(messageResponse.getContentText());
    if (!messageResult.ok) {
      Logger.log("❌ ようこそメッセージ送信失敗: " + messageResult.error);
    }
  } catch (err) {
    Logger.log("❌ sendWelcomeMessage() エラー: " + err.message);
  }
}
