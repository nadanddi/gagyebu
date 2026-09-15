const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const vm = require('vm');

const context = {
  console,
  Utilities: {
    DigestAlgorithm: {SHA_256: 'SHA_256'},
    Charset: {UTF_8: 'UTF_8'},
    computeDigest(_algorithm, text) {
      return [...crypto.createHash('sha256').update(text, 'utf8').digest()].map((n) => n > 127 ? n - 256 : n);
    },
    formatDate() { return '2026-08'; }
  }
};
vm.createContext(context);
vm.runInContext(fs.readFileSync('apps-script/Code.gs', 'utf8'), context);

const toss = context.parseMatrix_([
  ['', '토스뱅크 거래내역'],
  ['', '거래 일시', '적요', '거래 유형', '거래 기관', '계좌번호', '거래 금액', '거래 후 잔액', '메모'],
  ['', '2026.08.10 12:00:00', '예시식당', '체크카드결제', '', '', '-12,000', '3,000', '']
], '토스뱅크_거래내역.xlsx');
assert.equal(toss.bank, '토스뱅크');
assert.equal(toss.rows.length, 1);
assert.equal(toss.rows[0].outgoing, 12000);

const ok = context.parseMatrix_([
  ['거래내역조회'],
  ['거래일자', '출금액', '입금액', '잔액', '적요'],
  ['2026.08.11', '10,000', '0', '90,000 원', '토뱅 본인이름']
], 'OK저축은행_8월.xls');
assert.equal(ok.bank, 'OK저축은행');
assert.equal(ok.rows[0].time, '');

const pdf = context.parseWooriPdfText_(
  '2026-08-12 09:21:04 펌뱅킹 네이버파이낸셜 3,320 - 306,200',
  '우리은행_n페이_8월.pdf'
);
assert.equal(pdf.bank, '우리은행 N페이');
assert.equal(pdf.rows[0].outgoing, 3320);

const fakeFile = {getId: () => 'file', getName: () => 'sample.xlsx', getUrl: () => 'https://example.invalid'};
const own = context.normalizeTransaction_(ok.rows[0], ok.bank, fakeFile, ['본인이름']);
assert.equal(own.bucket, '내부이체');
assert.equal(own.amount, 10000);

const external = context.normalizeTransaction_({
  date: '2026-08-11', time: '10:00:00', description: '다른사람', type: '계좌이체',
  outgoing: 5000, incoming: 0, rawAmount: -5000, balance: 0, sourceRow: 1
}, '토스뱅크', fakeFile, ['본인이름']);
assert.equal(external.bucket, '지출');

const topUpRows = [
  context.normalizeTransaction_({
    date: '2026-08-12', time: '', description: '토뱅 본인이름', type: '계좌이체',
    outgoing: 30000, incoming: 0, rawAmount: -30000, balance: 70000, sourceRow: 2
  }, 'OK저축은행', fakeFile, []),
  context.normalizeTransaction_({
    date: '2026-08-12', time: '12:34:57', description: '카드잔액 자동충전', type: '입금',
    outgoing: 0, incoming: 30000, rawAmount: 30000, balance: 30000, sourceRow: 3
  }, '토스뱅크', fakeFile, [])
];
assert.equal(context.reconcileTossCardTopUps_(topUpRows), 1);
assert.deepEqual(Array.from(topUpRows, (row) => row.bucket), ['내부이체', '내부이체']);
assert.deepEqual(Array.from(topUpRows, (row) => row.method), ['카드잔액 자동충전', '카드잔액 자동충전']);
assert.match(topUpRows[0].note, /OK저축은행→토스뱅크 자동충전 대응/);

assert.match(
  context.naverLedgerNote_('원본에 거래시간 없음', {
    paymentId: 'sample-payment', amount: 4500, merchant: '예시매장', item: '예시상품',
    detailUrl: 'https://orders.pay.naver.com/instantPay/detail/sample-payment'
  }, '네이버파이낸셜'),
  /원본에 거래시간 없음.*예시매장.*결제번호: sample-payment.*은행 표시내용: 네이버파이낸셜/
);
assert.equal(
  context.naverLedgerNote_('[네이버페이 사용처: 이전매장]', {
    paymentId: 'new-payment', amount: 1300, merchant: '새매장', item: '', detailUrl: ''
  }, '네이버파이낸셜'),
  '[네이버페이 결제: 새매장 / 결제번호: new-payment / 표시금액: 1,300원 / 은행 표시내용: 네이버파이낸셜]'
);
assert.equal(context.timeDistance_('13:47:30', '13:47:38'), 8);
assert.equal(context.naverMatchScore_('2026-08-20', 17500, '19:53:03', {
  date: '2026-08-20', time: '19:53:00', amount: 17500
}), 3);
assert.equal(context.naverMatchScore_('2026-08-20', 16266, '19:53:03', {
  date: '2026-08-20', time: '19:53:00', amount: 17500
}), 86403);
assert.equal(context.naverMatchScore_('2026-08-20', 16266, '20:00:00', {
  date: '2026-08-20', time: '19:53:00', amount: 17500
}), null);
assert.equal(context.category_('Apple Services', '', '출금', '지출', '네이버페이'), '구독·디지털');
assert.equal(context.category_('설빙전북대점', '', '출금', '지출', '네이버페이'), '카페·간식');
assert.equal(context.category_('비비고김치100G 외 2종', '', '출금', '지출', '네이버페이'), '식료품·편의점');
assert.equal(
  context.naverDetailKey_('2026.08.20', '19:53:00', 17500, ' 예시매장 '),
  '2026-08-20|19:53:00|17500|예시매장'
);
assert.deepEqual(
  JSON.parse(JSON.stringify(context.normalizeNaverPayDetail_({
    paymentId: 'sample-payment', date: '2026-09-11', time: '13:47:38',
    amount: 4500, merchant: '예시매장', item: '예시상품',
    detailUrl: 'https://orders.pay.naver.com/instantPay/detail/sample-payment'
  }))),
  {
    paymentId: 'sample-payment', date: '2026-09-11', time: '13:47:38', amount: 4500,
    merchant: '예시매장', item: '예시상품',
    detailUrl: 'https://orders.pay.naver.com/instantPay/detail/sample-payment'
  }
);
assert.equal(context.secureEquals_('same-token', 'same-token'), true);
assert.equal(context.secureEquals_('same-token', 'other-token'), false);
assert.equal(context.merchantKey_('(주)세광'), '세광');
assert.equal(context.merchantKey_('지에스 25 전북대점'), 'GS25전북대점');
assert.equal(
  context.matchSimpleRule_('GS25 전북대점', [{key: 'GS25전북대점', example: 'GS25 전북대점', category: '생활용품'}]).category,
  '생활용품'
);
assert.equal(context.simpleFallbackCategory_({
  direction: '출금', description: '예시식당', type: '체크카드결제', bucket: '지출'
}), '식비');
const simpleImported = context.simpleLedgerRow_({
  id: 'auto-sample', date: '2026-09-15', time: '12:00:00', bank: '토스뱅크',
  description: 'GS25 전북대점', type: '체크카드결제', direction: '출금', amount: 6500,
  bucket: '지출', note: ''
}, [{key: 'GS25전북대점', example: 'GS25 전북대점', category: '생활용품'}]);
assert.deepEqual(Array.from(simpleImported.slice(1, 12)), [
  '지출', '생활용품', '', 6500, '토스카드', '토스', '', 'GS25 전북대점',
  '원본 토스뱅크 12:00:00 / 유형: 체크카드결제 / 분류규칙: GS25 전북대점', '2026-09', 'auto-sample'
]);
console.log('parser tests passed');
