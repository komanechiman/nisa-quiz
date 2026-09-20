/**
 * とうしクイズ ─ 結果メール受信サーバー（Google Apps Script）
 *
 * 【セットアップ手順】
 *  1. https://script.google.com/ を開いて「新しいプロジェクト」
 *  2. このファイルの中身をぜんぶコピーして貼りつけ、プロジェクト名を「とうしクイズ」などに
 *  3. 下の SECRET に、すきな合言葉を入れる（例: "himitsu2027"）
 *  4. 右上「デプロイ」→「新しいデプロイ」→ 種類の選択で「ウェブアプリ」
 *       次のユーザーとして実行 : 自分
 *       アクセスできるユーザー : 全員
 *     →「デプロイ」→ 初回は権限の承認を求められるので許可する
 *  5. 表示される「ウェブアプリのURL」（.../exec で終わるもの）をコピー
 *  6. とうしクイズの ⚙️ せってい を開き、URL と 合言葉 を貼って「保存」→「テスト送信」
 *
 * 【コードを直したとき】
 *  「デプロイ」→「デプロイを管理」→ 鉛筆アイコン →
 *  バージョンを「新バージョン」にして「デプロイ」。URL は変わりません。
 *  （「新しいデプロイ」を押すと別のURLになってしまうので注意）
 */

/** 送り先。空にすると、このスクリプトを持っているGoogleアカウント宛に送ります。 */
var TO = "";

/** 合言葉。アプリの「せってい」に入れるものと同じ文字にしてください。
 *  空のままでも動きますが、そのURLを知った人は誰でもあなた宛にメールを送れてしまいます。
 *  かならず設定することをおすすめします。 */
var SECRET = "";

/** いたずら対策：1時間に送るメールの上限 */
var MAX_PER_HOUR = 30;


function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, error: "no body" });
    }
    var d = JSON.parse(e.postData.contents);

    if (SECRET && d.token !== SECRET) {
      return json({ ok: false, error: "bad token" });
    }
    if (!underLimit()) {
      return json({ ok: false, error: "rate limited" });
    }
    // 同じ結果が二重に届いたとき（アプリ側の再送とすれ違ったとき）は捨てる
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

/** ブラウザでURLを直接開いたときの動作確認用 */
function doGet() {
  return json({ ok: true, msg: "とうしクイズ 受信サーバーは動いています" });
}


/* ---------- メールの中身 ---------- */

function subjectOf(d) {
  if (d.test) return "[とうしクイズ] テスト送信";
  var who = d.player || "だれか";
  return "[とうしクイズ] " + who + "　" + d.score + "/" + d.total + "（" + (d.category || "") + "）";
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
      + '<p style="margin:0;color:#6E8189;">とうしクイズの設定は完了です。これから、プレイするたびに結果が届きます。</p>'
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
      h += '<div style="border-left:4px solid #E2566B;background:#FFF;padding:10px 14px;margin-bottom:8px;'
         + 'border-radius:0 10px 10px 0;box-shadow:0 1px 3px rgba(0,0,0,.08);">';
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


/* ---------- ユーティリティ ---------- */

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
