// 既存の注文用 Apps Script に、このファイルを追加する。
// doGet(e) の先頭に次の1行を追加する。
// if (e.parameter.type === 'containerConfig') return containerConfig_();
// doPost(e) の JSON を data に読み込んだ直後、ほかの注文処理より前に追加する。
// if (data.type === 'containerPickup') return saveContainerPickup_(data);
// setupContainerPickup を一度実行し、「コンテナ設定」のA列に営業日を入力する。価格はC列で変更できる。
// 既存のウェブアプリのデプロイを新しいバージョンに更新する。
const CONTAINER_MENU = [
  ['pork', 'プルドポークサンド'],
  ['pork_cheese', 'プルドポークチーズサンド'],
  ['salmon', '厚切りスモークサーモンとクリームチーズのサンド', true],
  ['prosciutto', '生ハムサンド', true]
];

function setupContainerPickup() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const existing = ss.getSheetByName('コンテナ設定');
  const prices = [650, 750, 900, 900];
  if (existing) {
    const rows = existing.getDataRange().getValues();
    rows.forEach((row, index) => {
      const product = CONTAINER_MENU.findIndex(item => item[1] === row[1]);
      if (product >= 0 && (row[2] === '' || row[2] === undefined)) existing.getRange(index + 1, 3).setValue(prices[product]);
    });
    return;
  }
  const sheet = ss.insertSheet('コンテナ設定');
  sheet.appendRow(['営業日', '商品', '税込価格（円）']);
  sheet.getRange('B2:C5').setValues(CONTAINER_MENU.map((item, index) => [item[1], prices[index]]));
  sheet.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
}

function containerSettings_(ss) {
  const sheet = ss.getSheetByName('コンテナ設定');
  if (!sheet) throw Error('現在コンテナ受取の準備中です。');
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 4), 3).getValues();
  const days = values.map(row => row[0] instanceof Date
    ? Utilities.formatDate(row[0], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(row[0]).trim())
    .filter(day => /^\d{4}-\d{2}-\d{2}$/.test(day));
  const menu = CONTAINER_MENU.map(([id, name, reserved]) => {
    const row = values.find(row => row[1] === name);
    const price = Number(row && row[2]);
    if (!Number.isSafeInteger(price) || price <= 0) throw Error('現在メニューを準備中です。');
    return {id, name, reserved: Boolean(reserved), price};
  });
  return {type:'containerConfig', days:[...new Set(days)].sort(), menu};
}

function containerConfig_() {
  try {
    return jsonOutput_(containerSettings_(SpreadsheetApp.openById(SS_ID)));
  } catch (err) {
    return jsonOutput_({ok:false, error:err.message});
  }
}

function saveContainerPickup_(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return jsonOutput_({ok:false, error:'混み合っています。同じ内容で再送してください。'});
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('注文記録');
    if (!/^[0-9a-f-]{36}$/i.test(data.receiptId)) throw Error('受付番号を確認できません。');
    // A〜L列は既存の注文・返信に使用。M列に再送照合用の受付番号を保存する。
    if (sheet.getLastRow() >= 2) {
      const found = sheet.getRange(2, 13, sheet.getLastRow()-1, 1).createTextFinder(data.receiptId).matchEntireCell(true).findNext();
      if (found) return jsonOutput_({ok:true, total:sheet.getRange(found.getRow(), 7).getValue()});
    }
    const config = containerSettings_(ss);
    const pickup = Date.parse(data.day + 'T' + data.time + ':00+09:00');
    if (!config.days.includes(data.day) || !/^(1[1-4]:(00|15|30|45)|15:00)$/.test(data.time) ||
        !Number.isFinite(pickup) || pickup < Date.now() + 30*60000) throw Error('この受取日時の受付は終了しています。');
    const name = String(data.name || '').trim();
    if (!name || name.length > 20 || !Array.isArray(data.items) || !data.items.length || data.items.length > 4)
      throw Error('お名前と商品を確認してください。');
    let total = 0;
    const seen = new Set();
    const lines = data.items.map(({id, count}) => {
      const item = config.menu.find(m => m.id === id);
      if (!item || seen.has(id) || !Number.isInteger(count) || count < 1 || count > 99) throw Error('商品と数量を確認してください。');
      seen.add(id);
      total += item.price * count;
      return item.name + ' ×' + count;
    });
    if (data.phone) lines.push('電話：' + String(data.phone).slice(0, 15));
    if (data.note) lines.push('備考：' + String(data.note).slice(0, 120));
    const safeCell = value => /^[=+@-]/.test(value) ? "'" + value : value;
    sheet.appendRow([new Date(), 'コンテナ受取', safeCell(name), data.day, data.time,
      lines.join('\n'), total, safeCell(String(data.lineUserId || '')), '', '', '', Utilities.getUuid(), data.receiptId]);
    SpreadsheetApp.flush();
    return jsonOutput_({ok:true, total});
  } catch (err) {
    return jsonOutput_({ok:false, error:err.message});
  } finally {
    lock.releaseLock();
  }
}
