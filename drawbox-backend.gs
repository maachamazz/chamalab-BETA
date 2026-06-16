/* ═══════════════════════════════════════════════════
   DRAWBOX バックエンド (Google Apps Script)
   ───────────────────────────────────────────────────
   設置方法:
   1. https://script.google.com を開く(Googleアカウントでログイン)
   2. 「新しいプロジェクト」→ このコードを全部貼り付け → 保存
   3. 右上「デプロイ」→「新しいデプロイ」
      - 種類の選択(歯車アイコン)→「ウェブアプリ」
      - 次のユーザーとして実行 :「自分」
      - アクセスできるユーザー :「全員」 ←★重要
   4. 「デプロイ」→ アクセスを承認(自分のアカウントを選択)
      ※「このアプリは確認されていません」と出たら
        「詳細」→「(プロジェクト名)に移動」で進めばOK
   5. 表示された「ウェブアプリのURL」(…/exec で終わるもの)をコピーして、
      index.html の DRAWBOX_API_URL に貼り付ける

   ※ コードを修正した時は「デプロイ」→「デプロイを管理」→
     鉛筆アイコン →「バージョン: 新バージョン」→ デプロイ
     (これをしないと修正が反映されません)

   投稿された絵は、あなたのGoogleドライブの
   「drawbox-images」フォルダに保存されます。
   不適切な絵はそのフォルダから直接削除すればギャラリーからも消えます。
   ═══════════════════════════════════════════════════ */

const FOLDER_NAME = 'drawbox-images'; // 保存先フォルダ名(自動作成されます)
const MAX_IMAGES  = 60;               // ギャラリーに表示する最大枚数
const MAX_SIZE_KB = 400;              // 受け付ける画像の最大サイズ(KB)

// ── 絵の投稿を受け取る ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const base64 = String(data.image || '').split(',')[1];
    if (!base64) return json_({ ok: false, error: 'no image data' });

    const bytes = Utilities.base64Decode(base64);
    if (bytes.length > MAX_SIZE_KB * 1024) {
      return json_({ ok: false, error: 'image too large' });
    }

    const blob = Utilities.newBlob(bytes, 'image/png', 'drawbox_' + Date.now() + '.png');
    const file = getFolder_().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ── ギャラリー用の画像一覧を返す ──
function doGet() {
  try {
    const it = getFolder_().getFiles();
    const files = [];
    while (it.hasNext()) {
      const f = it.next();
      files.push({ date: f.getDateCreated().getTime(), id: f.getId() });
    }
    files.sort((a, b) => b.date - a.date); // 新しい順

    const images = files.slice(0, MAX_IMAGES).map(f => ({
      url: 'https://lh3.googleusercontent.com/d/' + f.id
    }));
    return json_({ ok: true, images: images });
  } catch (err) {
    return json_({ ok: false, error: String(err), images: [] });
  }
}

// ── ヘルパー ──
function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
