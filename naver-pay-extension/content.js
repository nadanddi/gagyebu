(function () {
  if (document.getElementById('gagyebu-npay-sync')) return;

  const panel = document.createElement('div');
  panel.id = 'gagyebu-npay-sync';
  panel.innerHTML = '<input type="month" aria-label="동기화할 월"><button type="button">가계부 자동 동기화</button><span></span>';
  document.body.appendChild(panel);
  const button = panel.querySelector('button');
  const monthInput = panel.querySelector('input');
  const status = panel.querySelector('span');
  const now = new Date();
  monthInput.value = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  button.addEventListener('click', runSync);

  async function runSync() {
    button.disabled = true;
    monthInput.disabled = true;
    try {
      const targetMonth = monthInput.value;
      if (!/^20\d{2}-\d{2}$/.test(targetMonth)) throw new Error('동기화할 월을 선택하세요.');
      setStatus('결제 목록 확인 중…');
      const currentHtml = document.documentElement.outerHTML;
      const hintedMaximum = NaverPayParser.maximumPage(currentHtml);
      const maximum = hintedMaximum > 1 ? hintedMaximum : 20;
      const links = new Set(NaverPayParser.detailLinks(currentHtml, location.href));
      for (let page = 1; page <= maximum; page += 1) {
        const pageUrl = new URL(location.href);
        pageUrl.searchParams.set('page', String(page));
        const html = page === currentPage() ? currentHtml : await fetchText(pageUrl.href);
        const before = links.size;
        NaverPayParser.detailLinks(html, pageUrl.href).forEach((url) => links.add(url));
        if (hintedMaximum === 1 && page > 1 && links.size === before) break;
      }
      if (!links.size) throw new Error('상세 결제 링크를 찾지 못했습니다. 현장결제 탭인지 확인하세요.');

      const records = [];
      let processed = 0;
      for (const url of links) {
        setStatus('상세내역 읽는 중 ' + (++processed) + '/' + links.size);
        const html = await fetchText(url);
        const record = NaverPayParser.parseDetail(html, url);
        if (NaverPayParser.isComplete(record) && record.date.slice(0, 7) === targetMonth) records.push(record);
      }
      if (!records.length) throw new Error(targetMonth + '의 날짜·금액·사용처가 있는 결제내역을 찾지 못했습니다.');
      setStatus('가계부와 맞추는 중…');
      const response = await send({type: 'syncLedger', records});
      setStatus('완료: ' + response.result.matched + '건 자동입력 · ' + response.result.unmatched + '건 미일치', 'ok');
    } catch (error) {
      setStatus(error.message || String(error), 'error');
      if (/설정에서/.test(error.message || '')) await send({type: 'openOptions'});
    } finally {
      button.disabled = false;
      monthInput.disabled = false;
    }
  }

  function currentPage() {
    return Number(new URL(location.href).searchParams.get('page') || 1);
  }

  async function fetchText(url) {
    const response = await send({type: 'fetchNaver', url});
    return response.text;
  }

  function send(message) {
    return new Promise((resolve, reject) => chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
      else if (!response || !response.ok) reject(new Error(response && response.error || '확장 프로그램 응답이 없습니다.'));
      else resolve(response);
    }));
  }

  function setStatus(message, type) {
    status.textContent = message;
    status.className = type || '';
  }
})();
