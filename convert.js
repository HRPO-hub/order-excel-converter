// 주문 엑셀 변환 로직. 행 객체 배열만 다룬다(파일 입출력은 index.html).
// 고객마다 양식이 다르면 아래 상수만 고치면 된다.

const ORDER_COLUMNS = ['상품주문번호', '상품명', '옵션정보', '수량', '수취인명', '수취인연락처1', '수취인연락처2', '우편번호', '배송지', '배송메세지'];
const MAPPING_COLUMNS = ['상품명', '거래처'];
const LABEL_COLUMNS = ['받는분성명', '받는분전화번호', '받는분기타연락처', '받는분우편번호', '받는분주소', '품목명', '내품수량', '배송메세지', '고객주문번호'];
const PO_COLUMNS = ['상품주문번호', '상품명', '옵션', '수량', '수취인명', '연락처', '우편번호', '주소', '배송메세지'];
const TRACKING_COLUMNS = { orderNo: '고객주문번호', trackingNo: '운송장번호' };
const DISPATCH_COLUMNS = ['상품주문번호', '배송방법', '택배사', '송장번호'];
const DISPATCH_FIXED = { method: '택배,등기,소포', carrier: 'CJ대한통운' };

function cleanRows(rows) {
  return rows
    .map(row => Object.fromEntries(Object.entries(row).map(([k, v]) => [String(k).trim(), String(v ?? '').trim()])))
    .filter(row => Object.values(row).some(v => v !== ''));
}

function missingColumns(rows, required) {
  const keys = Object.keys(rows[0] || {});
  return required.filter(c => !keys.includes(c));
}

// 엑셀이 숫자로 저장해 앞자리 0이 사라진 경우 복구
function fixPhone(s) {
  return /^\d{9,10}$/.test(s) && s[0] !== '0' ? '0' + s : s;
}

function fixZip(s) {
  return /^\d{1,4}$/.test(s) ? s.padStart(5, '0') : s;
}

function convertOrders(orderRows, mappingRows) {
  const orders = cleanRows(orderRows);
  const missing = missingColumns(orders, ORDER_COLUMNS);
  if (missing.length) return { ok: false, error: `필요한 열이 없어요: ${missing.join(', ')}` };

  const badQty = orders.filter(o => !/^\d+$/.test(o.수량) || Number(o.수량) <= 0).map(o => o.상품주문번호);
  if (badQty.length) return { ok: false, error: `수량이 올바르지 않은 주문이 있어요: ${badQty.join(', ')}` };

  const mapping = cleanRows(mappingRows);
  // 매칭표가 비어 있으면 전부 미매칭으로 보고, 행이 있을 때만 열을 검사한다
  const mapMissing = mapping.length ? missingColumns(mapping, MAPPING_COLUMNS) : [];
  if (mapMissing.length) return { ok: false, error: `매칭표에 필요한 열이 없어요: ${mapMissing.join(', ')}` };
  const vendorOf = new Map();
  for (const m of mapping) {
    if (m.거래처 && !vendorOf.has(m.상품명)) vendorOf.set(m.상품명, m.거래처);
  }

  const labels = orders.map(o => ({
    받는분성명: o.수취인명,
    받는분전화번호: fixPhone(o.수취인연락처1),
    받는분기타연락처: fixPhone(o.수취인연락처2),
    받는분우편번호: fixZip(o.우편번호),
    받는분주소: o.배송지,
    품목명: o.옵션정보 ? `${o.상품명} / ${o.옵션정보}` : o.상품명,
    내품수량: o.수량,
    배송메세지: o.배송메세지,
    고객주문번호: o.상품주문번호,
  }));

  const byVendor = new Map(); // 거래처명이 'constructor', '__proto__' 등이어도 안전
  const unmatched = [];
  for (const o of orders) {
    const vendor = vendorOf.get(o.상품명);
    if (!vendor) {
      unmatched.push({ orderNo: o.상품주문번호, productName: o.상품명 });
      continue;
    }
    if (!byVendor.has(vendor)) byVendor.set(vendor, []);
    byVendor.get(vendor).push({
      상품주문번호: o.상품주문번호,
      상품명: o.상품명,
      옵션: o.옵션정보,
      수량: o.수량,
      수취인명: o.수취인명,
      연락처: fixPhone(o.수취인연락처1),
      우편번호: fixZip(o.우편번호),
      주소: o.배송지,
      배송메세지: o.배송메세지,
    });
  }

  return { ok: true, labels, purchaseOrders: Object.fromEntries(byVendor), unmatched };
}

function convertTracking(rows) {
  const data = cleanRows(rows);
  const { orderNo, trackingNo } = TRACKING_COLUMNS;
  const missingCols = missingColumns(data, [orderNo, trackingNo]);
  if (missingCols.length) return { ok: false, error: `필요한 열이 없어요: ${missingCols.join(', ')}` };

  const seen = new Set();
  const out = [];
  const missing = [];
  const duplicates = [];
  const noOrderNo = []; // 수동 발송 등 주문번호 없는 송장: 스마트스토어에 올릴 수 없음
  for (const r of data) {
    const no = r[orderNo];
    if (!no) {
      noOrderNo.push(r[trackingNo]);
      continue;
    }
    if (seen.has(no)) {
      if (!duplicates.includes(no)) duplicates.push(no);
      continue;
    }
    seen.add(no);
    if (!r[trackingNo]) {
      missing.push(no);
      continue;
    }
    out.push({ 상품주문번호: no, 배송방법: DISPATCH_FIXED.method, 택배사: DISPATCH_FIXED.carrier, 송장번호: r[trackingNo] });
  }
  return { ok: true, rows: out, missing, duplicates, noOrderNo };
}

function safeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_');
}

function purchaseOrderFiles(purchaseOrders, date) {
  const used = new Map();
  return Object.entries(purchaseOrders).map(([vendor, rows]) => {
    const base = safeFileName(vendor);
    const n = (used.get(base) || 0) + 1;
    used.set(base, n);
    const name = n === 1 ? base : `${base} (${n})`;
    return { fileName: `발주서_${name}_${date}.xlsx`, rows };
  });
}

const Convert = {
  ORDER_COLUMNS, MAPPING_COLUMNS, LABEL_COLUMNS, PO_COLUMNS, TRACKING_COLUMNS, DISPATCH_COLUMNS, DISPATCH_FIXED,
  cleanRows, missingColumns, fixPhone, fixZip, convertOrders, safeFileName, purchaseOrderFiles, convertTracking,
};
if (typeof module !== 'undefined') module.exports = Convert;
else window.Convert = Convert;
