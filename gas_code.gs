/**
 * 8月既存生販売（梅・竹・松）お申し込みページ ── 申込管理シート連携 + 自動返信メール
 *
 * 【セットアップ手順】
 * 1. Googleスプレッドシートを新規作成（例：「8月既存生販売_申込管理」）
 * 2. 拡張機能 → Apps Script を開き、このコードを全て貼り付け
 * 3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *      - 次のユーザーとして実行：自分
 *      - アクセスできるユーザー：全員
 * 4. 発行されたウェブアプリURLを index.html の GAS_URL に貼り付け
 * 5. Stripe/PayPalの決済リンクが用意できたら、下のCONFIGに追記する
 */

const CONFIG = {
  // ── 銀行振込先（自動返信メールに記載）──
  BANK_INFO:
    "【銀行振込（日本国内）】\n" +
    "※ 弊社所在地がドイツのため、日本国内のお振込みは\n" +
    "　レムケなつこ個人口座になります。\n\n" +
    "三菱UFJ銀行\n" +
    "・店名：池上支店（116）\n" +
    "・預金種目：普通\n" +
    "・口座番号：0190673\n" +
    "・口座名義：ﾚﾑｹﾅﾂｺ\n\n" +
    "ドイツ法人口座へのお振込みをご希望の場合は、\n" +
    "このメールに「ドイツ法人振込希望」とご返信ください。詳細をご案内します。\n\n" +
    "※ 振込手数料は国内外問わずご負担をお願いいたします。\n" +
    "※ 銀行振込は、プランを問わず一括のみとなります。",

  // ── Stripe / PayPal 決済リンク（準備でき次第ここに貼る。空欄の間は「追ってご案内」で返信）──
  STRIPE_LINK: "",
  PAYPAL_LINK: "",

  SENDER_NAME: "IOBオーガニックスクール事務局",
  SENDER_EMAIL: "school@iob.bio",   // Gmailの送信元アドレスに登録済みの場合のみ反映
  NOTIFY_TO: "hi@iob.bio"           // 申込通知の受信先（カンマ区切りで複数可）
};

function checkAliases() {
  Logger.log(GmailApp.getAliases());
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["申込日時", "お名前", "メールアドレス", "プラン", "希望決済方法", "お支払い回数", "モニター表記希望", "ご質問・連絡事項", "通知結果", "返信結果"]);
    }

    const ts = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    sheet.appendRow([ts, d.name, d.email, d.plan, d.payment, d.installment, d.monitorName, d.message]);
    const lastRow = sheet.getLastRow();

    var notifyStatus = "";
    var replyStatus = "";

    // ── 事務局への控え通知 ──
    if (CONFIG.NOTIFY_TO) {
      try {
        MailApp.sendEmail({
          to: CONFIG.NOTIFY_TO,
          subject: "【8月既存生販売】" + d.name + " 様が申し込みました（" + d.plan + "）",
          body: "申込が入りました。\n\n" +
            "お名前：" + d.name + "\n" +
            "メール：" + d.email + "\n" +
            "プラン：" + d.plan + "\n" +
            "希望決済方法：" + d.payment + "\n" +
            "お支払い回数：" + d.installment + "\n" +
            "モニター表記希望：" + d.monitorName + "\n" +
            "ご質問・連絡事項：" + d.message + "\n" +
            "日時：" + ts
        });
        notifyStatus = "通知OK " + ts;
      } catch (e2) {
        notifyStatus = "通知ERR: " + e2;
      }
      sheet.getRange(lastRow, 9).setValue(notifyStatus);
    }

    // ── 申込者への自動返信 ──
    if (d.email) {
      try {
        var payBlock = "";
        if (d.payment && d.payment.indexOf("銀行振込") !== -1) {
          payBlock = CONFIG.BANK_INFO;
        } else if (d.payment === "Stripe（クレジットカード）") {
          payBlock = CONFIG.STRIPE_LINK
            ? "【Stripe（クレジットカード）でのお支払い】\n" + CONFIG.STRIPE_LINK + "\n\nこちらのリンクよりお手続きください。"
            : "【Stripe（クレジットカード）でのお支払い】\n決済リンクを準備中です。整い次第、改めてメールでご案内いたします。今しばらくお待ちください。";
        } else if (d.payment === "PayPal") {
          payBlock = CONFIG.PAYPAL_LINK
            ? "【PayPalでのお支払い】\n" + CONFIG.PAYPAL_LINK + "\n\nこちらのリンクよりお手続きください。"
            : "【PayPalでのお支払い】\n決済リンクを準備中です。整い次第、改めてメールでご案内いたします。今しばらくお待ちください。";
        } else {
          payBlock = "決済方法を確認のうえ、改めてご案内いたします。";
        }

        const body =
          d.name + " 様\n\n" +
          "IOBオーガニックスクール事務局です。\n" +
          "この度は「" + d.plan + "」のお申し込みをいただき、誠にありがとうございました。\n\n" +
          "■ お申し込み内容\n" +
          "プラン：" + d.plan + "\n" +
          "希望決済方法：" + d.payment + "\n" +
          "お支払い回数：" + d.installment + "\n\n" +
          "■ お支払いについて\n\n" +
          payBlock + "\n\n\n" +
          "■ 開始時期について\n\n" +
          "梅（メールレッスン）は9月中旬、竹・松（共同学習会・グループコンサル）は10月からの開始を予定しています。日程が決まり次第、改めてご案内いたします。\n\n\n" +
          "■ モニターについて\n\n" +
          "今回はモニター価格でのご案内です。アンケートと取材へのご協力をお願いしております。お名前の出し方（イニシャル／フルネーム）・お顔出しの有無は、その際に改めて個別にお伺いします。\n\n\n" +
          "ご不明な点は、お気軽に school@iob.bio までご連絡ください。\n" +
          "これから、また一緒に学んでいけたら嬉しいです。\n\n" +
          "──────\n" +
          "Institut für Organic Business GmbH\n" +
          "オーガニックビジネス研究所 スクール事務局\n" +
          "「オーガニックが、あたりまえ。」な社会へ\n" +
          "営業時間：日本時間 平日10〜16時／お問い合わせは3営業日以内にお答えします\n" +
          "──────";

        const opts = { name: CONFIG.SENDER_NAME };
        try {
          if (GmailApp.getAliases().indexOf(CONFIG.SENDER_EMAIL) !== -1) {
            opts.from = CONFIG.SENDER_EMAIL;
          }
        } catch (e3) {}
        GmailApp.sendEmail(d.email, "【お申し込みありがとうございます】" + d.plan + " のご案内", body, opts);
        replyStatus = "返信OK " + ts;
      } catch (e4) {
        replyStatus = "返信ERR: " + e4;
      }
      sheet.getRange(lastRow, 10).setValue(replyStatus);
    }

    return ContentService.createTextOutput(JSON.stringify({ result: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
