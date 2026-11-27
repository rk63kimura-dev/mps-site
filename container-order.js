(async function containerOrder() {
  document.title = 'お昼の予約注文 | THE PEACE';
  const el = id => document.getElementById(id);
  const nav = document.createElement('nav');
  nav.className = 'input-card';
  nav.setAttribute('aria-label', 'その他の予約');
  document.querySelector('header').after(nav);
  nav.innerHTML = '<a href="order-app.html">朝の注文</a>　／　<a href="focaccia-order.html">フォカッチャの予約</a>';
  document.querySelector('.header-tagline').textContent = 'お昼の予約注文（コンテナ受取）';
  document.querySelector('.section-label').textContent = 'コンテナ受取メニュー';
  el('morningStockBanner').remove();
  el('cancelBtn').remove();
  el('orderNote').placeholder = 'ご希望があればご記入ください';
  document.querySelector('.success-msg').textContent = 'ご注文ありがとうございます。ご指定の時間にコンテナでお受け取りください。';
  document.querySelector('.back-btn').onclick = () => location.reload();
  const date = el('pickupDate'), time = el('pickupTime'), button = el('submitBtn');
  date.onchange = updateTimes;
  button.disabled = true;
  const menuList = el('menuList');
  const notice = document.createElement('p');
  notice.setAttribute('role', 'status');
  notice.textContent = '営業日を確認しています…';
  menuList.append(notice);
  el('submitBtn').querySelector('.btn-label').textContent = 'コンテナ受取を予約する';
  let menu = [], receiptId = '', previousBody = '';
  const quantity = new Map();
  function slots(day) {
    const result = [];
    for (let minute = 660; minute <= 900; minute += 15) {
      const label = Math.floor(minute / 60) + ':' + String(minute % 60).padStart(2, '0');
      if (Date.parse(day + 'T' + label + ':00+09:00') >= Date.now() + 1800000) result.push(label);
    }
    return result;
  }
  function updateTimes() {
    time.replaceChildren(new Option('時間を選んでください', ''));
    for (const label of slots(date.value)) time.add(new Option(label, label));
    time.disabled = time.length === 1;
  }
  function orderItems() {
    return menu.map(item => ({id: item.id, count: Number(quantity.get(item.id).value)})).filter(item => item.count);
  }
  function summary() {
    const items = orderItems();
    el('orderSummary').style.display = items.length ? 'block' : 'none';
    el('summaryItems').textContent = items.map(i => menu.find(m => m.id === i.id).name + ' ×' + i.count).join('、');
    el('summaryTotal').textContent = '¥' + items.reduce((sum, i) => sum + i.count * menu.find(m => m.id === i.id).price, 0).toLocaleString();
  }
  await initLiff();
  try {
    const config = await (await fetch(GAS_URL + '?type=containerConfig', {cache:'no-store'})).json();
    if (config.type !== 'containerConfig') throw Error();
    menu = config.menu;
    const days = config.days.filter(day => slots(day).length);
    if (!days.length) {
      notice.textContent = '現在、予約できるコンテナ営業日はありません。';
      return;
    }
    notice.textContent = '受取は営業日の11:00〜15:00。ご注文は受取の30分前までです。';
    for (const item of menu) {
      const row = document.createElement('label');
      row.className = 'input-card';
      row.style.display = 'block';
      row.textContent = item.name + (item.reserved ? '（予約限定）' : '') + '　¥' + item.price.toLocaleString() + '　';
      const input = document.createElement('select');
      input.setAttribute('aria-label', item.name + 'の数量');
      for (let n = 0; n <= 99; n++) input.add(new Option(n + '個', n));
      input.onchange = summary;
      quantity.set(item.id, input);
      row.append(input);
      el('menuList').append(row);
    }
    days.forEach(day => date.add(new Option(day, day)));
    button.disabled = false;
  } catch {
    notice.textContent = '受付情報を読み込めませんでした。時間をおいて、もう一度開いてください。';
    return;
  }
  button.onclick = async () => {
    const items = orderItems(), name = el('nameInput').value.trim();
    if (!name || !items.length || !date.value || !time.value) return showError('商品・お名前・受取日時を選んでください。');
    if (!slots(date.value).includes(time.value)) {
      updateTimes();
      return showError('この時間の受付は終了しました。受取時間を選び直してください。');
    }
    const data = {type:'containerPickup', name, lineUserId, items, day:date.value, time:time.value,
      phone:el('phoneInput').value.trim(), note:el('orderNote').value.trim()};
    const body = JSON.stringify(data);
    if (body !== previousBody) receiptId = crypto.randomUUID();
    previousBody = body;
    button.disabled = true;
    el('errorMsg').classList.remove('show');
    try {
      const result = await (await fetch(GAS_URL, {method:'POST', body:JSON.stringify({...data, receiptId})})).json();
      if (!result.ok) return showError(result.error || '受付できませんでした。');
      el('successDetail').textContent = 'コンテナ受取　' + data.day + ' ' + data.time + '　' + name + ' 様　¥' + result.total.toLocaleString();
      el('successScreen').classList.add('show');
    } catch {
      showError('受付結果を確認できませんでした。同じ内容のまま再送して確認できます。');
    } finally {
      button.disabled = false;
    }
  };
})();
