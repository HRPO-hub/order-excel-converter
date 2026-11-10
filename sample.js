// 시연용 가상 데이터. 이름·주소·전화번호는 실존하지 않는 값이다.

const SAMPLE_MAPPING = [
  { 상품명: '국산 햇양파 3kg', 거래처: '무안농산' },
  { 상품명: '무안 황토 고구마 5kg', 거래처: '무안농산' },
  { 상품명: '제주 감귤 5kg', 거래처: '제주팜' },
  { 상품명: '제주 한라봉 3kg', 거래처: '제주팜' },
  { 상품명: '한우 불고기 500g', 거래처: '강원한우' },
  { 상품명: '한우 국거리 300g', 거래처: '강원한우' },
];

const SAMPLE_ORDERS = (() => {
  const names = ['김하늘', '이도윤', '박서연', '최지호', '정유나', '강민준', '조수아', '윤시우', '장예린', '임건우'];
  const addrs = [
    ['04524', '서울특별시 중구 가상로 12, 101동 1001호'],
    ['48058', '부산광역시 해운대구 예시대로 34'],
    ['21999', '인천광역시 연수구 샘플길 56, 3층'],
    ['61475', '광주광역시 동구 테스트로 7'],
    ['34126', '대전광역시 유성구 연습로 89, 202호'],
  ];
  const options = ['', '특품', '선물포장', ''];
  const messages = ['문 앞에 놓아주세요', '', '경비실에 맡겨주세요', '배송 전 연락 부탁드려요', ''];
  const products = SAMPLE_MAPPING.map(m => m.상품명);
  const orders = [];
  for (let i = 0; i < 20; i++) {
    const [zip, addr] = addrs[i % addrs.length];
    orders.push({
      주문번호: `20261110${String(1000 + Math.floor(i / 2)).padStart(6, '0')}`,
      상품주문번호: `202611100${String(10001 + i)}`,
      구매자명: names[i % names.length],
      상품명: i === 2 ? '수제 딸기잼 500g' : products[i % products.length],
      옵션정보: options[i % options.length],
      수량: String((i % 3) + 1),
      수취인명: names[(i + 3) % names.length],
      수취인연락처1: `010-0000-${String(1000 + i * 37).slice(-4)}`,
      수취인연락처2: i % 4 === 0 ? `02-000-${String(2000 + i).slice(-4)}` : '',
      우편번호: zip,
      배송지: addr,
      배송메세지: messages[i % messages.length],
    });
  }
  return orders;
})();

// 송장마다 가짜 운송장번호를 붙인다. 마지막 1건은 빈칸(경고 시연용).
function sampleTracking(labels) {
  return labels.map((l, i) => ({
    고객주문번호: l.고객주문번호,
    받는분성명: l.받는분성명,
    운송장번호: i === labels.length - 1 ? '' : `6${String(84100000000 + i * 7919).padStart(11, '0')}`,
  }));
}

const Sample = { SAMPLE_ORDERS, SAMPLE_MAPPING, sampleTracking };
if (typeof module !== 'undefined') module.exports = Sample;
else window.Sample = Sample;
