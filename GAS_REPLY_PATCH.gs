// THE PEACE: 注文通知からお客様へ返信するための追加コード
//
// 1) doPost(e) の JSON.parse(e.postData.contents) の直後に追加：
//
// if (data.type === 'customerReply') {
//   return handleCustomerReply(data);
// }
//
// 2) 朝注文を保存している箇所で、appendRow 後を下記の形に変更：
//
// const newRow = sheet2.getLastRow();
// const replyToken = Utilities.getUuid();
// sheet2.getRange(newRow, 12).setValue(replyToken); // L列：返信トークン
// setupReplyControlsForRow_(sheet2, newRow);
//
// const replyUrl =
//   'https://rk63kimura-dev.github.io/mps-site/reply.html' +
//   '?row=' + encodeURIComponent(newRow) +
//   '&token=' + encodeURIComponent(replyToken) +
//   '&name=' + encodeURIComponent(name);
//
// const ownerMessage =
//   '\n新しい注文！' +
//   '\n━━━━━━━━━━━━' +
//   '\n👤 ' + name + ' 様' +
//   '\n🕐 ピックアップ ' + pickupTime +
//   '\n━━━━━━━━━━━━' +
//   '\n' + lines +
//   '\n━━━━━━━━━━━━' +
//   '\n💰 合計　¥' + total +
//   '\n\n📩 このお客様に返信' +
//   '\n' + replyUrl;
//
// 3) どこかの関数外に以下を追加：

function handleCustomerReply(data) {
  try {
    const row = Number(data.row);
    const token = String(data.token || '');
    const message = String(data.message || '').trim();

    if (!Number.isInteger(row) || row < 2) {
      return jsonOutput_({ ok: false, error: 'invalid row' });
    }
    if (!token) {
      return jsonOutput_({ ok: false, error: 'missing token' });
    }
    if (!message || message.length > 500) {
      return jsonOutput_({ ok: false, error: 'invalid message' });
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('注文記録');

    const savedToken = String(sheet.getRange(row, 12).getValue() || ''); // L列
    if (!savedToken || savedToken !== token) {
      return jsonOutput_({ ok: false, error: 'invalid token' });
    }

    const customerName = String(sheet.getRange(row, 3).getValue() || 'お客様');
    const userId = String(sheet.getRange(row, 8).getValue() || '');

    if (!userId || userId === 'IDなし') {
      return jsonOutput_({ ok: false, error: 'LINE ID unavailable' });
    }

    const fullMessage =
      customerName + ' 様\n\n' +
      message +
      '\n\nTHE PEACE';

    sendLine(userId, fullMessage);

    sheet.getRange(row, 11).setValue(
      '返信済み ' +
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm')
    );

    return jsonOutput_({ ok: true });

  } catch (err) {
    return jsonOutput_({
      ok: false,
      error: err.message
    });
  }
}

// 4) 任意：注文記録シートの L1 に「返信トークン」と入れ、L列は非表示にしてください。
