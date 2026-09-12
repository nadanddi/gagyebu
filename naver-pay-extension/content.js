(function () {
  if (document.getElementById('gagyebu-npay-sync')) return;

  const MAX_HISTORY_PAGES = 100;
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
  send({type: 'getConfig'}).then((config) => {
    if (!config.configured) button.textContent = '네이버페이 내역 CSV 저장';
  }).catch(() => {});

  async function runSync() {
    button.disabled = true;
    monthInput.disabled = true;
    try {
      const targetMonth = monthInput.value;
      if (!/^20\d{2}-\d{2}$/.test(targetMonth)) throw new Error('동기화할 월을 선택하세요.');

      const records = [];
      const seenLinks = new Set();
      const seenPages = new Set();
      let foundTarget = false;
      let reachedEnd = false;

      for (let page = 1; page <= MAX_HISTORY_PAGES; page += 1) {
        setStatus('결제 목록 ' + page + '페이지 확인 중…');
        const pageUrl = new URL(location.href);
        pageUrl.searchParams.set('page', String(page));
        const html = page === currentPage()
          ? document.documentElement.outerHTML
          : await fetchText(pageUrl.href);
        const pageLinks = NaverPayParser.detailLinks(html, pageUrl.href);
        const signature = pageLinks.join('|');

        if (!pageLinks.length || seenPages.has(signature)) {
          reachedEnd = true;
          break;
        }
        seenPages.add(signature);

        const pageRecords = [];
        for (const url of pageLinks) {
          if (seenLinks.has(url)) continue;
          seenLinks.add(url);
          setStatus('결제 목록 ' + page + '페이지 상세내역 확인 중 (' + (pageRecords.length + 1) + '/' + pageLinks.length + ')');
          const detailHtml = await fetchText(url);
          const record = NaverPayParser.parseDetail(detailHtml, url);
          if (!NaverPayParser.isComplete(record)) continue;
          pageRecords.push(record);
          if (record.date.slice(0, 7) === targetMonth) {
            records.push(record);
            foundTarget = true;
          }
        }

        if (NaverPayParser.scanDecision(pageRecords, targetMonth, foundTarget) === 'stop') {
          reachedEnd = true;
          break;
        }
      }

      if (!reachedEnd) throw new Error('100페이지까지 확인했지만 ' + targetMonth + '의 끝에 도달하지 못했습니다. 기간을 나눠 다시 시도하세요.');
      if (!seenLinks.size) throw new Error('상세 결제 링크를 찾지 못했습니다. 전체 또는 현장결제 탭인지 확인하세요.');
      if (!records.length) throw new Error(targetMonth + '의 날짜·금액·사용처가 있는 결제내역을 찾지 못했습니다.');

      records.sort((left, right) => (left.date + left.time).localeCompare(right.date + right.time));
      const config = await send({type: 'getConfig'});
      if (!config.configured) {
        downloadCsv(records, targetMonth);
        setStatus('완료: 여러 페이지에서 찾은 ' + records.length + '건을 CSV로 저장했습니다.', 'ok');
        return;
      }
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

  function downloadCsv(records, targetMonth) {
    const headers = ['결제일', '결제시간', '결제번호', '실제매장', '상품', '금액', '상세URL'];
    const rows = records.map((record) => [
      record.date, record.time, record.paymentId, record.merchant,
      record.item, record.amount, record.detailUrl
    ]);
    const csv = '\uFEFF' + [headers].concat(rows).map((row) => row.map(csvCell).join(',')).join('\r\n');
    const url = URL.createObjectURL(new Blob([csv], {type: 'text/csv;charset=utf-8'}));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = '네이버페이_상세내역_' + targetMonth + '.csv';
    anchor.style.display = 'none';
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function csvCell(value) {
    const text = String(value === null || value === undefined ? '' : value);
    return '"' + text.replace(/"/g, '""') + '"';
  }
})();
