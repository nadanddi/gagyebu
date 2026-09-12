const assert = require('assert');
const parser = require('../parser');

const history = `
  <html><body>
    <div>광고 배너 2 / 7</div>
    <a href="?page=4">4페이지</a>
    <a href="https://orders.pay.naver.com/instantPay/detail/20260911NP1000000001?backUrl=x">결제완료</a>
    <a href="/instantPay/detail/20260911NP1000000002">결제완료</a>
  </body></html>`;

assert.equal(parser.maximumPage(history), 4);
assert.deepEqual(parser.detailLinks(history, 'https://pay.naver.com/pc/history?page=2'), [
  'https://orders.pay.naver.com/instantPay/detail/20260911NP1000000001?backUrl=x',
  'https://pay.naver.com/instantPay/detail/20260911NP1000000002'
]);

const detail = `
  <html><body>
    <section><div>2026. 09. 11. 13:47:38</div><div>결제번호 20260911NP1000000001</div></section>
    <h2>결제상품</h2><div>예시편의점 중앙점</div><button>문의하기</button>
    <h3>결제완료</h3><div>예시상품 외 1종</div><div>결제 승인번호: 12345678</div><div>4,500원</div>
    <h2>결제정보</h2><div>결제금액</div><strong>총 4,500원</strong>
  </body></html>`;

assert.deepEqual(parser.parseDetail(detail, 'https://orders.pay.naver.com/instantPay/detail/20260911NP1000000001'), {
  paymentId: '20260911NP1000000001',
  date: '2026-09-11',
  time: '13:47:38',
  merchant: '예시편의점 중앙점',
  item: '예시상품 외 1종',
  amount: 4500,
  detailUrl: 'https://orders.pay.naver.com/instantPay/detail/20260911NP1000000001'
});

assert.equal(parser.isComplete(parser.parseDetail(detail, 'https://example.invalid')), true);
assert.equal(parser.scanDecision([
  {date: '2026-09-10'}, {date: '2026-09-01'}
], '2026-08', false), 'continue');
assert.equal(parser.scanDecision([
  {date: '2026-08-31'}, {date: '2026-08-01'}
], '2026-08', true), 'continue');
assert.equal(parser.scanDecision([
  {date: '2026-08-01'}, {date: '2026-07-31'}
], '2026-08', true), 'stop');
assert.equal(parser.scanDecision([
  {date: '2026-07-31'}
], '2026-08', false), 'stop');
console.log('naver pay parser tests passed');
