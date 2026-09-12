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

assert.equal(
  context.naverPayNote_('원본에 거래시간 없음', '예시매장', '간단 메모'),
  '원본에 거래시간 없음 [네이버페이 사용처: 예시매장 / 간단 메모]'
);
assert.equal(
  context.naverPayNote_('[네이버페이 사용처: 이전매장]', '새매장', ''),
  '[네이버페이 사용처: 새매장]'
);
assert.equal(context.timeDistance_('13:47:30', '13:47:38'), 8);
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
assert.match(
  context.naverSyncMemo_('', {
    paymentId: 'sample-payment', item: '예시상품',
    detailUrl: 'https://orders.pay.naver.com/instantPay/detail/sample-payment'
  }),
  /결제번호: sample-payment.*상품: 예시상품/
);

console.log('parser tests passed');
