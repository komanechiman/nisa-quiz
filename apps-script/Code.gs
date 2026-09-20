/**
 * とうしクイズ ─ 結果メール送信サーバー（Google Apps Script）
 *
 * このファイルは「まるごとコピーして貼るだけ」で動きます。書きかえる必要はありません。
 * 合言葉は初回にこのプログラムが自動で作ってくれます。
 *
 * ▼ 手順（SETUP.md にスクリーンショット相当の詳しい説明があります）
 *   1. script.google.com →「新しいプロジェクト」
 *   2. 最初から入っている「function myFunction() {}」を全部消して、このファイルを全部貼る
 *   3. 右上「デプロイ」→「新しいデプロイ」→ 歯車から「ウェブアプリ」を選ぶ
 *        次のユーザーとして実行 ： 自分
 *        アクセスできるユーザー ： 全員          ← ここ重要
 *   4.「デプロイ」→ 権限を承認（「安全ではないページ」警告の通りかたは SETUP.md 参照）
 *   5. 出てきた「ウェブアプリのURL」(.../exec) を、そのままブラウザで開く
 *        → 合言葉が表示されます（表示されるのは最初の1回だけ）
 *   6. とうしクイズの ⚙️ に「URL」と「合言葉」を貼って、保存 → テスト送信
 *
 * ▼ 合言葉をもう一度見たい / 作りなおしたいとき
 *   このエディタの上にある関数リストから showSecret または resetSecret を選んで「実行」。
 *   結果は「実行ログ」に出ます。
 *
 * ▼「MailApp.sendEmail を呼び出す権限がありません」と出たら
 *   関数リストから sendTestMail を選んで「実行」→ 権限を承認。
 *   そのあと「デプロイ」→「デプロイを管理」→ 鉛筆 → バージョン「新バージョン」→「デプロイ」。
 *   それでも直らなければ SETUP.md の該当節（appsscript.json で権限を明示）を参照。
 *
 * ▼ コードを直したあと
 *   「デプロイ」→「デプロイを管理」→ 鉛筆アイコン → バージョン「新バージョン」→「デプロイ」。
 *   URL は変わりません。（「新しいデプロイ」を押すと別URLになるので注意）
 */

/** 送り先。空のままなら、このスクリプトを持っているGoogleアカウント宛に届きます。 */
var TO = "";

/** 合言葉を自分で決めたいときだけ書く。空なら自動生成されます。 */
var SECRET = "";

/** いたずら対策：1時間に送るメールの上限 */
var MAX_PER_HOUR = 30;

var PROP_SECRET = "secret";
var PROP_SHOWN  = "secret_shown";


/* ================= 受信 ================= */

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, error: "no body" });
    }
    var d = JSON.parse(e.postData.contents);

    var secret = currentSecret();
    if (!secret) {
      return json({ ok: false, error: "not set up: open this URL in a browser first" });
    }
    if (d.token !== secret) {
      return json({ ok: false, error: "bad token" });
    }
    if (!underLimit()) {
      return json({ ok: false, error: "rate limited" });
    }
    // 再送とすれちがって同じ結果が二重に届いたときは捨てる
    if (d.id && isDuplicate(d.id)) {
      return json({ ok: true, skipped: "duplicate" });
    }

    MailApp.sendEmail({
      to: TO || Session.getEffectiveUser().getEmail(),
      subject: subjectOf(d),
      htmlBody: htmlOf(d),
      body: textOf(d),
      name: "とうしクイズ"
    });
    return json({ ok: true });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}


/* ========== ブラウザで開いたときの画面（合言葉を出す） ========== */

function doGet() {
  var props  = PropertiesService.getScriptProperties();
  var secret = currentSecret();
  var first  = false;

  if (!secret) {
    secret = makeSecret();
    props.setProperty(PROP_SECRET, secret);
  }
  if (!props.getProperty(PROP_SHOWN)) {
    first = true;
    props.setProperty(PROP_SHOWN, "1");
  }

  // このページは URL を知っていれば誰でも開けるので、アドレスは伏せ字にする。
  // 全文を確認したいときは、エディタで showSecret を実行してください（所有者のみ）。
  var to = maskEmail(TO || Session.getEffectiveUser().getEmail());
  return HtmlService.createHtmlOutput(setupPage(secret, first, to))
    .setTitle("とうしクイズ 受信サーバー");
}

function setupPage(secret, first, to) {
  var F = "font-family:-apple-system,'Hiragino Sans','Yu Gothic',sans-serif;";
  var h = '<div style="' + F + 'max-width:520px;margin:32px auto;padding:0 20px;'
        + 'line-height:1.9;color:#26343A;">';
  h += '<p style="color:#6E8189;font-size:13px;margin:0;">とうしクイズ</p>';
  h += '<h1 style="font-size:22px;margin:4px 0 20px;">✅ 受信サーバーは動いています</h1>';

  if (first) {
    h += '<p style="margin:0 0 6px;font-size:14px;">この「合言葉」をアプリの ⚙️ に貼ってください。</p>';
    h += '<div style="background:#EAF7F2;border:2px solid #1E9E78;border-radius:14px;'
       + 'padding:18px;text-align:center;margin-bottom:8px;">'
       + '<code style="font-size:26px;font-weight:bold;letter-spacing:2px;color:#137A5B;">'
       + esc(secret) + '</code></div>';
    h += '<p style="font-size:12.5px;color:#E2566B;font-weight:bold;margin:0 0 20px;">'
       + '⚠️ この合言葉が表示されるのは、この1回だけです。いまメモしてください。</p>';
  } else {
    h += '<p style="font-size:14px;margin:0 0 20px;">合言葉はすでに発行ずみです。'
       + 'もう一度見たいときは、Apps Script のエディタで関数 <b>showSecret</b> を実行して'
       + '「実行ログ」を見てください。</p>';
  }

  h += '<div style="background:#F5F9F7;border-radius:14px;padding:16px 18px;font-size:13.5px;">';
  h += '<b>メールの送り先</b><br>' + esc(to);
  h += '<br><span style="font-size:11.5px;color:#6E8189;">'
     + '（このページは誰でも開けるので伏せ字にしています。'
     + '全文はエディタで showSecret を実行すると見られます）</span>';
  h += '</div>';

  h += '<p style="font-size:12.5px;color:#6E8189;margin-top:20px;">'
     + 'このページのURL（アドレス欄の .../exec）も、アプリの ⚙️ に貼る必要があります。</p>';
  h += '</div>';
  return h;
}


/* ========== エディタから手で実行する用 ========== */

/** 合言葉を表示する（実行ログに出ます） */
function showSecret() {
  var s = currentSecret();
  if (!s) {
    s = makeSecret();
    PropertiesService.getScriptProperties().setProperty(PROP_SECRET, s);
  }
  Logger.log("合言葉: " + s);
  Logger.log("メールの送り先: " + (TO || Session.getEffectiveUser().getEmail()));
  return s;
}

/** 合言葉を作りなおす（アプリ側にも入れなおしてください） */
function resetSecret() {
  var props = PropertiesService.getScriptProperties();
  var s = makeSecret();
  props.setProperty(PROP_SECRET, s);
  props.deleteProperty(PROP_SHOWN);
  Logger.log("新しい合言葉: " + s);
  return s;
}

/** 動作確認：自分あてにテストメールを1通送る */
function sendTestMail() {
  MailApp.sendEmail({
    to: TO || Session.getEffectiveUser().getEmail(),
    subject: "[とうしクイズ] テスト送信（エディタから）",
    htmlBody: htmlOf({ test: true, player: "テスト", atLocal: new Date().toLocaleString("ja-JP") }),
    name: "とうしクイズ"
  });
  Logger.log("送信しました: " + (TO || Session.getEffectiveUser().getEmail()));
}


/* ================= メールの中身 ================= */

function subjectOf(d) {
  if (d.test) return "[とうしクイズ] テスト送信";
  return "[とうしクイズ] " + (d.player || "だれか") + "　" + d.score + "/" + d.total
       + "（" + (d.category || "") + "）";
}

function textOf(d) {
  if (d.test) {
    return "とうしクイズのテスト送信です。\nこのメールが届いていれば設定は完了です。\n\n"
         + "なまえ: " + (d.player || "-") + "\n日時: " + (d.atLocal || "");
  }
  var L = [];
  L.push("なまえ　: " + (d.player || "-"));
  L.push("ジャンル: " + (d.category || "-"));
  L.push("スコア　: " + d.score + " / " + d.total + "（" + d.pct + "%）");
  L.push("かかった時間: " + fmtSec(d.seconds));
  L.push("日時　　: " + (d.atLocal || ""));
  if (d.wrong && d.wrong.length) {
    L.push("");
    L.push("まちがえた " + d.wrong.length + " 問:");
    d.wrong.forEach(function (w, i) {
      L.push((i + 1) + ". " + w.q);
      L.push("    えらんだ: " + w.picked);
      L.push("    正解　　: " + w.answer);
    });
  } else {
    L.push("");
    L.push("全問正解でした。");
  }
  return L.join("\n");
}

function htmlOf(d) {
  var F = "font-family:-apple-system,'Hiragino Sans','Yu Gothic',sans-serif;";

  if (d.test) {
    return '<div style="' + F + 'font-size:15px;line-height:1.8;color:#26343A;">'
      + '<h2 style="margin:0 0 8px;font-size:18px;">✅ テスト送信が届きました</h2>'
      + '<p style="margin:0;color:#6E8189;">とうしクイズの設定は完了です。'
      + 'これから、プレイするたびに結果が届きます。</p>'
      + '<p style="margin:12px 0 0;color:#6E8189;font-size:13px;">なまえ: ' + esc(d.player || "-")
      + '<br>日時: ' + esc(d.atLocal || "") + '</p></div>';
  }

  var pct = Number(d.pct) || 0;
  var bar = pct >= 80 ? "#1E9E78" : (pct >= 50 ? "#E8A317" : "#E2566B");

  var h = '<div style="' + F + 'font-size:15px;line-height:1.8;color:#26343A;max-width:560px;">';
  h += '<p style="margin:0 0 4px;color:#6E8189;font-size:13px;">とうしクイズ</p>';
  h += '<h2 style="margin:0 0 14px;font-size:20px;">' + esc(d.player || "だれか")
     + ' さんが「' + esc(d.category || "-") + '」をやりました</h2>';

  h += '<div style="background:#F5F9F7;border-radius:14px;padding:16px 18px;margin-bottom:16px;">';
  h += '<div style="font-size:15px;">' + d.total + '問中 <b style="font-size:30px;color:' + bar + ';">'
     + d.score + '</b> 問せいかい</div>';
  h += '<div style="height:8px;background:#DCE5E1;border-radius:4px;margin:10px 0 6px;overflow:hidden;">'
     + '<div style="height:8px;width:' + pct + '%;background:' + bar + ';border-radius:4px;"></div></div>';
  h += '<div style="font-size:13px;color:#6E8189;">正答率 ' + pct + '%　・　'
     + fmtSec(d.seconds) + '　・　' + esc(d.atLocal || "") + '</div>';
  h += '</div>';

  if (d.wrong && d.wrong.length) {
    h += '<p style="margin:0 0 8px;font-size:13px;color:#6E8189;font-weight:bold;">まちがえた '
       + d.wrong.length + ' 問</p>';
    d.wrong.forEach(function (w) {
      h += '<div style="border-left:4px solid #E2566B;background:#FFF;padding:10px 14px;'
         + 'margin-bottom:8px;border-radius:0 10px 10px 0;box-shadow:0 1px 3px rgba(0,0,0,.08);">';
      h += '<div style="font-size:14px;font-weight:bold;margin-bottom:4px;">' + esc(w.q) + '</div>';
      h += '<div style="font-size:13px;color:#E2566B;">えらんだ: ' + esc(w.picked) + '</div>';
      h += '<div style="font-size:13px;color:#1E9E78;">正解　　: ' + esc(w.answer) + '</div>';
      h += '</div>';
    });
  } else {
    h += '<p style="margin:0;font-size:15px;color:#1E9E78;font-weight:bold;">🏆 全問正解でした</p>';
  }

  h += '</div>';
  return h;
}


/* ================= ユーティリティ ================= */

function currentSecret() {
  if (SECRET) return SECRET;
  return PropertiesService.getScriptProperties().getProperty(PROP_SECRET);
}

/** メールアドレスを伏せ字にする（例: k****i@gmail.com） */
function maskEmail(e) {
  e = String(e || "");
  var at = e.indexOf("@");
  if (at < 1) return "（不明）";
  var name = e.slice(0, at), dom = e.slice(at);
  if (name.length <= 2) return name.charAt(0) + "****" + dom;
  return name.charAt(0) + "****" + name.charAt(name.length - 1) + dom;
}

/** 読み書きしやすい合言葉を作る（まぎらわしい 0/O/1/l は使わない） */
function makeSecret() {
  var chars = "abcdefghijkmnpqrstuvwxyz23456789";
  var s = "";
  for (var i = 0; i < 10; i++) {
    s += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return s;
}

function fmtSec(s) {
  s = Number(s) || 0;
  var m = Math.floor(s / 60);
  return m ? (m + "分" + (s % 60) + "秒") : (s + "秒");
}

function esc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 同じ id を6時間以内に受け取っていたら true */
function isDuplicate(id) {
  var cache = CacheService.getScriptCache();
  var k = "seen_" + id;
  if (cache.get(k)) return true;
  cache.put(k, "1", 21600);
  return false;
}

function underLimit() {
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();
  var hits;
  try { hits = JSON.parse(props.getProperty("hits") || "[]"); } catch (e) { hits = []; }
  hits = hits.filter(function (t) { return now - t < 3600000; });
  if (hits.length >= MAX_PER_HOUR) return false;
  hits.push(now);
  props.setProperty("hits", JSON.stringify(hits));
  return true;
}
