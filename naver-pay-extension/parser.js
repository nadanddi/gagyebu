(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.NaverPayParser = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  function decodeHtml(value) {
    return String(value || '')
      .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'")
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ');
  }

  function htmlToLines(html) {
    const stripped = String(html || '')
      .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(?:div|p|li|section|article|h\d|button|dd|dt)>/gi, '\n')
      .replace(/<[^>]+>/g, ' ');
    return decodeHtml(stripped).split(/\r?\n/)
      .map((line) => line.replace(/\s+/g, ' ').trim())
      .filter(Boolean);
  }

  function detailLinks(html, baseUrl) {
    const result = [];
    const seen = new Set();
    const pattern = /href\s*=\s*["']([^"']*(?:instantPay\/detail|\/detail\/)[^"']*)["']/gi;
    let match;
    while ((match = pattern.exec(String(html || ''))) !== null) {
      try {
        const parsed = new URL(decodeHtml(match[1]), baseUrl);
        const url = parsed.href;
        if (!['pay.naver.com', 'orders.pay.naver.com'].includes(parsed.hostname) || !/\/detail\//i.test(parsed.pathname) || seen.has(url)) continue;
        seen.add(url);
        result.push(url);
      } catch (_error) {
        // 관련 없는 위젯에 들어 있는 잘못된 링크는 무시합니다.
      }
    }
    return result;
  }

  function maximumPage(html) {
    const pages = [...String(html || '').matchAll(/[?&]page=(\d+)/g)].map((match) => Number(match[1]));
    return pages.length ? Math.max(...pages.filter(Number.isFinite), 1) : 1;
  }

  function firstUseful(lines, start, excluded) {
    const blocked = excluded || /^(문의하기|결제완료|영수증|포인트 뽑기|결제상품)$/;
    for (let index = start; index < Math.min(lines.length, start + 12); index += 1) {
      const value = lines[index];
      if (!value || blocked.test(value) || /^결제 승인번호/.test(value) || /^[\d,]+원$/.test(value)) continue;
      return value;
    }
    return '';
  }

  function parseDateTime(text) {
    const match = String(text).match(/(20\d{2})[.\/-]\s*(\d{1,2})[.\/-]\s*(\d{1,2})[.]?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (!match) return {date: '', time: ''};
    return {
      date: match[1] + '-' + match[2].padStart(2, '0') + '-' + match[3].padStart(2, '0'),
      time: match[4].padStart(2, '0') + ':' + match[5] + ':' + (match[6] || '00')
    };
  }

  function parseDetail(html, url) {
    const lines = htmlToLines(html);
    const text = lines.join('\n');
    const dt = parseDateTime(text);
    const payment = text.match(/결제번호\s*([0-9A-Z-]+)/i);
    const total = text.match(/결제금액[\s\S]{0,160}?총\s*([\d,]+)\s*원/) || text.match(/총\s*([\d,]+)\s*원/);
    const productSection = lines.findIndex((line) => line === '결제상품');
    const completeSection = lines.findIndex((line) => line === '결제완료');
    const merchant = productSection >= 0 ? firstUseful(lines, productSection + 1) : '';
    const item = completeSection >= 0
      ? firstUseful(lines, completeSection + 1, /^(문의하기|영수증|포인트 뽑기|결제완료)$/)
      : '';
    const amount = total ? Number(total[1].replace(/,/g, '')) : 0;
    return {
      paymentId: payment ? payment[1] : paymentIdFromUrl(url),
      date: dt.date,
      time: dt.time,
      merchant: merchant,
      item: item,
      amount: Number.isFinite(amount) ? amount : 0,
      detailUrl: String(url || '')
    };
  }

  function paymentIdFromUrl(url) {
    const match = String(url || '').match(/\/detail\/([^?/#]+)/);
    return match ? decodeURIComponent(match[1]) : '';
  }

  function isComplete(record) {
    return Boolean(record && record.date && record.amount > 0 && (record.merchant || record.item));
  }

  function parseHistoryCard(text, targetYear, detailUrl) {
    const lines = String(text || '').split(/\r?\n/)
      .map((line) => line.replace(/\s+/g, ' ').trim())
      .filter(Boolean);
    const joined = lines.join('\n');
    const dateMatch = joined.match(/(?:20\d{2}[.\/-]\s*)?(\d{1,2})[.\/-]\s*(\d{1,2})[.]?\s+(\d{1,2}):(\d{2})\s*결제/);
    const amountMatch = joined.match(/(?:^|\n)\s*([\d,]+)\s*원(?:\s|$)/);
    const statusIndex = lines.findIndex((line) => line === '결제완료');
    let title = '';
    for (let index = statusIndex >= 0 ? statusIndex + 1 : 0; index < lines.length; index += 1) {
      const line = lines[index];
      if (/^[\d,]+\s*원/.test(line) || /\d{1,2}[.\/-]\s*\d{1,2}[.]?\s+\d{1,2}:\d{2}\s*결제/.test(line)) break;
      if (/^(포인트 뽑기|문의하기|영수증|결제완료)$/.test(line) || /적립/.test(line)) continue;
      title = line.replace(/\s*[>›〉]$/, '').trim();
      if (title) break;
    }
    const date = dateMatch
      ? String(targetYear) + '-' + dateMatch[1].padStart(2, '0') + '-' + dateMatch[2].padStart(2, '0')
      : '';
    const amount = amountMatch ? Number(amountMatch[1].replace(/,/g, '')) : 0;
    return {
      paymentId: paymentIdFromUrl(detailUrl),
      date: date,
      time: dateMatch ? dateMatch[3].padStart(2, '0') + ':' + dateMatch[4] + ':00' : '',
      merchant: title,
      item: title,
      amount: Number.isFinite(amount) ? amount : 0,
      detailUrl: String(detailUrl || '')
    };
  }

  function scanDecision(records, targetMonth, foundTarget) {
    const months = (records || [])
      .map((record) => String(record.date || '').slice(0, 7))
      .filter((month) => /^20\d{2}-\d{2}$/.test(month));
    if (!months.length) return 'continue';
    if (foundTarget && months.some((month) => month < targetMonth)) return 'stop';
    if (!foundTarget && months.every((month) => month < targetMonth)) return 'stop';
    return 'continue';
  }

  return {
    htmlToLines, detailLinks, maximumPage, parseDateTime, parseDetail,
    parseHistoryCard, paymentIdFromUrl, isComplete, scanDecision
  };
});
