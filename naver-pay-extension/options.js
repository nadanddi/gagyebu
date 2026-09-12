const webAppUrl = document.getElementById('webAppUrl');
const syncToken = document.getElementById('syncToken');
const status = document.getElementById('status');

chrome.storage.local.get(['webAppUrl', 'syncToken'], (value) => {
  webAppUrl.value = value.webAppUrl || '';
  syncToken.value = value.syncToken || '';
});

document.getElementById('save').addEventListener('click', async () => {
  const url = webAppUrl.value.trim();
  const token = syncToken.value.trim();
  try {
    const parsed = new URL(url);
    if (parsed.hostname !== 'script.google.com' || !/\/macros\/s\/[^/]+\/exec$/.test(parsed.pathname)) throw new Error();
  } catch (_error) {
    status.textContent = '실행 URL을 확인하세요.';
    return;
  }
  if (token.length < 20) {
    status.textContent = '동기화 토큰을 확인하세요.';
    return;
  }
  await chrome.storage.local.set({webAppUrl: url, syncToken: token});
  status.textContent = '저장했습니다.';
});
