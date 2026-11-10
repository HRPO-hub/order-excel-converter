const test = require('node:test');
const assert = require('node:assert/strict');
const { convertOrders, safeFileName, purchaseOrderFiles, convertTracking } = require('./convert.js');

function order(overrides = {}) {
  return {
    상품주문번호: '2026111000001',
    상품명: '국산 햇양파 3kg',
    옵션정보: '특품',
    수량: '2',
    수취인명: '김철수',
    수취인연락처1: '010-1234-5678',
    수취인연락처2: '',
    우편번호: '06234',
    배송지: '서울특별시 강남구 테헤란로 1',
    배송메세지: '문 앞',
    ...overrides,
  };
}

test('송장: 열 매핑과 품목명 결합', () => {
  const r = convertOrders([order()], []);
  assert.equal(r.ok, true);
  assert.deepEqual(r.labels[0], {
    받는분성명: '김철수', 받는분전화번호: '010-1234-5678', 받는분기타연락처: '',
    받는분우편번호: '06234', 받는분주소: '서울특별시 강남구 테헤란로 1',
    품목명: '국산 햇양파 3kg / 특품', 내품수량: '2', 배송메세지: '문 앞', 고객주문번호: '2026111000001',
  });
});

test('송장: 옵션 없으면 상품명만', () => {
  const r = convertOrders([order({ 옵션정보: '' })], []);
  assert.equal(r.labels[0].품목명, '국산 햇양파 3kg');
});

test('열 순서가 달라도 동작', () => {
  const reversed = Object.fromEntries(Object.entries(order()).reverse());
  const r = convertOrders([reversed], []);
  assert.equal(r.ok, true);
  assert.equal(r.labels[0].받는분성명, '김철수');
});

test('필수 열 누락 시 빠진 열 이름', () => {
  const row = order();
  delete row.수량;
  delete row.배송지;
  const r = convertOrders([row], []);
  assert.equal(r.ok, false);
  assert.match(r.error, /수량/);
  assert.match(r.error, /배송지/);
});

test('열 이름 앞뒤 공백 무시', () => {
  const row = order();
  delete row.수량;
  row['수량 '] = '2';
  const r = convertOrders([row], []);
  assert.equal(r.ok, true);
  assert.equal(r.labels[0].내품수량, '2');
});

test('앞자리 0 유지·복구', () => {
  const kept = convertOrders([order()], []).labels[0];
  assert.equal(kept.받는분전화번호, '010-1234-5678');
  assert.equal(kept.받는분우편번호, '06234');
  const fixed = convertOrders([order({ 수취인연락처1: 1012345678, 수취인연락처2: 212345678, 우편번호: 6234 })], []).labels[0];
  assert.equal(fixed.받는분전화번호, '01012345678');
  assert.equal(fixed.받는분기타연락처, '0212345678');
  assert.equal(fixed.받는분우편번호, '06234');
});

test('수량 오류는 주문번호와 함께 실패', () => {
  for (const q of ['0', '-1', '1.5', '']) {
    const r = convertOrders([order({ 수량: q })], []);
    assert.equal(r.ok, false, `수량 '${q}'`);
    assert.match(r.error, /2026111000001/);
  }
  assert.equal(convertOrders([order({ 수량: ' 2 ' })], []).labels[0].내품수량, '2');
});

test('빈 행 건너뜀', () => {
  const empty = Object.fromEntries(Object.keys(order()).map(k => [k, '']));
  const r = convertOrders([order(), empty], []);
  assert.equal(r.labels.length, 1);
});

const MAP = [{ 상품명: '국산 햇양파 3kg', 거래처: '무안농산' }, { 상품명: '제주 감귤 5kg', 거래처: '제주팜' }];

test('발주서: 거래처별 분리와 열', () => {
  const r = convertOrders([order(), order({ 상품주문번호: '2', 상품명: '제주 감귤 5kg', 옵션정보: '' })], MAP);
  assert.deepEqual(Object.keys(r.purchaseOrders).sort(), ['무안농산', '제주팜']);
  assert.deepEqual(r.purchaseOrders['무안농산'][0], {
    상품주문번호: '2026111000001', 상품명: '국산 햇양파 3kg', 옵션: '특품', 수량: '2', 수취인명: '김철수',
    연락처: '010-1234-5678', 우편번호: '06234', 주소: '서울특별시 강남구 테헤란로 1', 배송메세지: '문 앞',
  });
  assert.equal(r.purchaseOrders['제주팜'][0].상품주문번호, '2');
});

test('미매칭: 송장 포함, 발주서 제외, 경고', () => {
  const r = convertOrders([order({ 상품명: '없는 상품' })], MAP);
  assert.equal(r.labels.length, 1);
  assert.deepEqual(r.purchaseOrders, {});
  assert.deepEqual(r.unmatched, [{ orderNo: '2026111000001', productName: '없는 상품' }]);
});

test('상품명 앞뒤 공백 무시', () => {
  const r = convertOrders([order({ 상품명: ' 국산 햇양파 3kg ' })], [{ 상품명: '국산 햇양파 3kg  ', 거래처: '무안농산' }]);
  assert.deepEqual(Object.keys(r.purchaseOrders), ['무안농산']);
});

test('매칭표 중복은 첫 행, 거래처 빈칸은 미매칭', () => {
  const dup = convertOrders([order()], [{ 상품명: '국산 햇양파 3kg', 거래처: 'A' }, { 상품명: '국산 햇양파 3kg', 거래처: 'B' }]);
  assert.deepEqual(Object.keys(dup.purchaseOrders), ['A']);
  const blank = convertOrders([order()], [{ 상품명: '국산 햇양파 3kg', 거래처: '' }]);
  assert.deepEqual(blank.purchaseOrders, {});
  assert.equal(blank.unmatched.length, 1);
});

test('매칭표 열 누락 시 실패', () => {
  const r = convertOrders([order()], [{ 이름: 'x' }]);
  assert.equal(r.ok, false);
  assert.match(r.error, /거래처/);
});

test('파일명 금지 문자 치환', () => {
  assert.equal(safeFileName('A/B:C*?"<>|\\'), 'A_B_C_______');
});

test('파일명 충돌 시 (2)', () => {
  const files = purchaseOrderFiles({ 'A/B': [{}], 'A_B': [{}] }, '2026-11-10');
  assert.deepEqual(files.map(f => f.fileName), ['발주서_A_B_2026-11-10.xlsx', '발주서_A_B (2)_2026-11-10.xlsx']);
});

test('일괄발송: 변환과 고정값, 앞자리 0 유지', () => {
  const r = convertTracking([{ 고객주문번호: '2026111000001', 운송장번호: '0123456789012' }]);
  assert.equal(r.ok, true);
  assert.deepEqual(r.rows, [{ 상품주문번호: '2026111000001', 배송방법: '택배,등기,소포', 택배사: 'CJ대한통운', 송장번호: '0123456789012' }]);
});

test('운송장번호 빈 행 제외 + 경고', () => {
  const r = convertTracking([{ 고객주문번호: '1', 운송장번호: '' }]);
  assert.deepEqual(r.rows, []);
  assert.deepEqual(r.missing, ['1']);
});

test('중복 주문번호는 첫 번째만 + 경고', () => {
  const r = convertTracking([{ 고객주문번호: '1', 운송장번호: 'A' }, { 고객주문번호: '1', 운송장번호: 'B' }]);
  assert.equal(r.rows.length, 1);
  assert.equal(r.rows[0].송장번호, 'A');
  assert.deepEqual(r.duplicates, ['1']);
});

test('송장번호 결과 필수 열 누락', () => {
  const r = convertTracking([{ 고객주문번호: '1' }]);
  assert.equal(r.ok, false);
  assert.match(r.error, /운송장번호/);
});
