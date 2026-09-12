const ALLOWED_NAVER_HOSTS = new Set(['pay.naver.com', 'orders.pay.naver.com']);

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message && message.type === 'fetchNaver') {
    fetchNaver(message.url).then((text) => sendResponse({ok: true, text}))
      .catch((error) => sendResponse({ok: false, error: error.message}));
    return true;
  }
  if (message && message.type === 'syncLedger') {
    syncLedger(message.records).then((result) => sendResponse({ok: true, result}))
      .catch((error) => sendResponse({ok: false, error: error.message}));
    return true;
  }
  if (message && message.type === 'openOptions') {
    chrome.runtime.openOptionsPage();
    sendResponse({ok: true});
  }
  return false;
});

async function fetchNaver(rawUrl) {
  const url = new URL(rawUrl);
  if (url.protocol !== 'https:' || !ALLOWED_NAVER_HOSTS.has(url.hostname)) throw new Error('허용되지 않은 네이버 주소입니다.');
  const response = await fetch(url.href, {credentials: 'include', redirect: 'follow', cache: 'no-store'});
  if (!response.ok) throw new Error('네이버페이 페이지를 읽지 못했습니다: HTTP ' + response.status);
  const text = await response.text();
  if (/로그인|nidlogin/i.test(response.url) && !/결제내역|결제상품/.test(text)) throw new Error('네이버 로그인이 필요합니다.');
  return text;
}

async function syncLedger(records) {
  const config = await chrome.storage.local.get(['webAppUrl', 'syncToken']);
  if (!config.webAppUrl || !config.syncToken) throw new Error('확장 프로그램 설정에서 Apps Script 주소와 동기화 토큰을 입력하세요.');
  const endpoint = new URL(config.webAppUrl);
  if (endpoint.protocol !== 'https:' || endpoint.hostname !== 'script.google.com' || !/\/macros\/s\/[^/]+\/exec$/.test(endpoint.pathname)) {
    throw new Error('올바른 Apps Script 웹 앱 실행 주소가 아닙니다.');
  }
  const response = await fetch(endpoint.href, {
    method: 'POST',
    redirect: 'follow',
    headers: {'Content-Type': 'text/plain;charset=utf-8'},
    body: JSON.stringify({action: 'syncNaverPay', token: config.syncToken, records: records})
  });
  const text = await response.text();
  let result;
  try { result = JSON.parse(text); } catch (_error) { throw new Error('가계부 응답을 해석하지 못했습니다. 웹 앱 배포를 확인하세요.'); }
  if (!result.ok) throw new Error(result.error || '가계부 동기화에 실패했습니다.');
  return result;
}
