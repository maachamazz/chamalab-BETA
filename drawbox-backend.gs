// DRAWBOX バックエンド (Google Apps Script)
// 設置方法は drawbox-setup.md を参照してください。

const FOLDER_NAME         = 'drawbox-images';         // 公開(ギャラリー表示)用フォルダ名
const PRIVATE_FOLDER_NAME = 'drawbox-images-private';  // 非公開(展示NG)用フォルダ名
const MAX_IMAGES  = 60;               // ギャラリーに表示する最大枚数
const MAX_SIZE_KB = 400;              // 受け付ける画像の最大サイズ(KB)
const NOTIFY_EMAIL = 'chama.zzz58@gmail.com'; // 非公開投稿の通知を送るメールアドレス

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

    // allowGallery が false の場合はギャラリーに出さず、非公開フォルダに保存してメールで通知する
    const allowGallery = data.allowGallery !== false;
    const blob = Utilities.newBlob(bytes, 'image/png', 'drawbox_' + Date.now() + '.png');

    if (allowGallery) {
      const file = getFolder_().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } else {
      getPrivateFolder_().createFile(blob);
      try {
        MailApp.sendEmail({
          to: NOTIFY_EMAIL,
          subject: '[Drawbox] 非公開の投稿が届きました',
          body: '展示NGの絵が投稿されました。添付画像を確認してください。',
          attachments: [blob]
        });
      } catch (mailErr) {
        // メール送信に失敗しても、絵自体は非公開フォルダに保存済みなので投稿自体は成功扱いにする
        // (デバッグ用: Cloudログが見れない環境向けに、エラー内容をテキストファイルとしてDriveに残す)
        getPrivateFolder_().createFile('drawbox_mail_error_' + Date.now() + '.txt', String(mailErr));
      }
    }

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

function getPrivateFolder_() {
  const it = DriveApp.getFoldersByName(PRIVATE_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PRIVATE_FOLDER_NAME);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
