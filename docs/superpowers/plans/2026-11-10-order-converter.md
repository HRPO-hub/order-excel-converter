# 주문 엑셀 변환기 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 스마트스토어 주문 엑셀을 CJ대한통운 송장 양식과 거래처별 발주서(zip)로, 택배사 송장번호 결과를 스마트스토어 일괄발송 양식으로 바꾸는 브라우저 전용 웹페이지.

**Architecture:** 변환 로직은 `convert.js`에 순수 함수로 두고 행 객체 배열만 다룬다(파일 입출력 없음). `index.html`이 SheetJS로 엑셀을 읽고 쓰고 JSZip으로 묶는다. `convert.js`와 `sample.js`는 브라우저(전역 객체)와 Node(`module.exports`) 양쪽에서 쓸 수 있게 끝에서 내보낸다.

**Tech Stack:** 순수 HTML/JS, SheetJS `xlsx@0.18.5`, `jszip@3.10.1` (둘 다 cdnjs), 테스트는 Node 내장 `node:test` + `node:assert/strict` (npm 의존성 없음).

**Spec:** `docs/superpowers/specs/2026-11-10-order-converter-design.md`

## Global Constraints

- 서버 없음. 파일은 브라우저 밖으로 보내지 않는다.
- 빌드 도구·npm 패키지 없음. 테스트 실행: `node --test` (프로젝트 루트).
- 주문 필수 열: 상품주문번호, 상품명, 옵션정보, 수량, 수취인명, 수취인연락처1, 수취인연락처2, 우편번호, 배송지, 배송메세지
- 매칭표 열: 상품명, 거래처
- 송장 출력 열(순서 고정): 받는분성명, 받는분전화번호, 받는분기타연락처, 받는분우편번호, 받는분주소, 품목명, 내품수량, 배송메세지, 고객주문번호
- 발주서 출력 열(순서 고정): 상품주문번호, 상품명, 옵션, 수량, 수취인명, 연락처, 우편번호, 주소, 배송메세지
- 송장번호 결과 필수 열: 고객주문번호, 운송장번호
- 일괄발송 출력 열: 상품주문번호, 배송방법(`택배,등기,소포`), 택배사(`CJ대한통운`), 송장번호
- 품목명 = 상품명 + (옵션정보가 있으면 ` / ` + 옵션정보)
- 파일명: `CJ대한통운_송장_YYYY-MM-DD.xlsx`, `발주서_YYYY-MM-DD.zip`, `발주서_<거래처>_YYYY-MM-DD.xlsx`, `스마트스토어_일괄발송_YYYY-MM-DD.xlsx`
- 파일명 금지 문자 `\ / : * ? " < > |` → `_`
- 모든 셀 값은 문자열로 다룬다.
- 열 이름·순서 상수는 `convert.js` 맨 위 한 곳에 둔다.
- 화면 문구는 한국어.

## Review Focus

1. 원본 엑셀이 전화번호·우편번호를 숫자 셀로 저장해 앞자리 0이 이미 사라진 경우 → 우편번호는 5자리로 0을 채우고, 숫자만 9~10자리이면서 0으로 시작하지 않는 전화번호는 앞에 0을 붙인다. (Task 1 테스트)
2. 열 이름 앞뒤에 공백이 있는 경우(`"수량 "`) → 같은 열로 인식한다. (Task 1 테스트)
3. 수량이 `0`, `-1`, `1.5`, 빈칸 → 해당 상품주문번호와 함께 실패. `" 2 "`는 2로 통과. (Task 1 테스트)
4. 매칭표에 같은 상품명이 두 번 있거나 거래처가 빈칸 → 첫 번째 유효 행 사용, 거래처 빈칸은 미매칭 처리. (Task 2 테스트)
5. 서로 다른 거래처명이 치환 후 같은 파일명이 되는 경우(`A/B`, `A_B`) → zip 안에서 덮어쓰지 않고 두 번째에 ` (2)`를 붙인다. (Task 2 테스트)

---

### Task 1: 주문 읽기 검증과 송장 변환

**Files:**
- Create: `convert.js`
- Test: `convert.test.js`

**Interfaces:**
- Produces:
  - 상수 `ORDER_COLUMNS: string[]`, `LABEL_COLUMNS: string[]`, `PO_COLUMNS: string[]`, `TRACKING_COLUMNS: {orderNo: '고객주문번호', trackingNo: '운송장번호'}`, `DISPATCH_COLUMNS: string[]`, `DISPATCH_FIXED: {method: '택배,등기,소포', carrier: 'CJ대한통운'}`
  - `cleanRows(rows: object[]) -> object[]` — 키·값을 문자열로 바꾸고 trim, 모든 값이 빈 행 제거
  - `missingColumns(rows: object[], required: string[]) -> string[]` — 첫 행의 키 기준. 행이 없으면 required 전부
  - `fixPhone(s: string) -> string`, `fixZip(s: string) -> string`
  - `convertOrders(orderRows: object[], mappingRows: object[]) -> {ok: true, labels: object[], purchaseOrders: Object<string, object[]>, unmatched: {orderNo: string, productName: string}[]} | {ok: false, error: string}` — Task 1에서는 labels와 오류만, purchaseOrders는 `{}`, unmatched는 `[]`로 둔다
  - 파일 끝: `if (typeof module !== 'undefined') module.exports = {...}; else window.Convert = {...};`

- [ ] **Step 1: Node 설치 확인**

Run: `node --version` — 없으면 `winget install --id OpenJS.NodeJS.LTS -e --accept-package-agreements --accept-source-agreements --silent` 후 PATH 새로고침.
Expected: `v22` 이상 출력.

- [ ] **Step 2: 실패하는 테스트 작성** (`convert.test.js`)

테스트용 `order(overrides)` 헬퍼: 필수 열 10개를 모두 가진 기본 행(상품주문번호 `2026111000001`, 상품명 `국산 햇양파 3kg`, 옵션정보 `특품`, 수량 `2`, 수취인명 `김철수`, 수취인연락처1 `010-1234-5678`, 수취인연락처2 ``, 우편번호 `06234`, 배송지 `서울특별시 강남구 테헤란로 1`, 배송메세지 `문 앞`)에 overrides를 덮어쓴다.

```js
test('송장: 열 매핑과 품목명 결합', () => {
  const r = convertOrders([order()], []);
  assert.equal(r.ok, true);
  assert.deepEqual(r.labels[0], {
    받는분성명: '김철수', 받는분전화번호: '010-1234-5678', 받는분기타연락처: '',
    받는분우편번호: '06234', 받는분주소: '서울특별시 강남구 테헤란로 1',
    품목명: '국산 햇양파 3kg / 특품', 내품수량: '2', 배송메세지: '문 앞', 고객주문번호: '2026111000001',
  });
});
test('송장: 옵션 없으면 상품명만', () => { /* 옵션정보 '' → 품목명 '국산 햇양파 3kg' */ });
test('열 순서가 달라도 동작', () => { /* order()의 키를 역순으로 재구성한 행 → ok */ });
test('필수 열 누락 시 빠진 열 이름', () => {
  const row = order(); delete row.수량; delete row.배송지;
  const r = convertOrders([row], []);
  assert.equal(r.ok, false);
  assert.match(r.error, /수량/); assert.match(r.error, /배송지/);
});
test('열 이름 앞뒤 공백 무시', () => { /* 키 '수량 ' → ok, 내품수량 '2' */ });
test('앞자리 0 유지·복구', () => {
  // 문자열 '010…','06234' 유지 / 숫자로 들어온 1012345678 → '01012345678', 6234 → '06234'
});
test('수량 오류는 주문번호와 함께 실패', () => {
  for (const q of ['0', '-1', '1.5', '']) {
    const r = convertOrders([order({ 수량: q })], []);
    assert.equal(r.ok, false); assert.match(r.error, /2026111000001/);
  }
  assert.equal(convertOrders([order({ 수량: ' 2 ' })], []).labels[0].내품수량, '2');
});
test('빈 행 건너뜀', () => { /* [order(), {모든 키 ''}] → labels.length 1 */ });
```

- [ ] **Step 3: 실패 확인**

Run: `node --test`
Expected: FAIL (`Cannot find module './convert.js'`)

- [ ] **Step 4: `convert.js`에 Interfaces의 상수와 함수 구현**

`fixPhone`: 숫자만이고 길이 9~10, 첫 글자가 `0`이 아니면 `'0' + s`. `fixZip`: 숫자만이고 길이 < 5면 `padStart(5, '0')`. 수량 검증: `/^\d+$/` 이고 `> 0`. 오류 메시지 예: `필요한 열이 없어요: 수량, 배송지`, `수량이 올바르지 않은 주문이 있어요: 2026111000001`.

- [ ] **Step 5: 통과 확인**

Run: `node --test`
Expected: 모든 테스트 PASS

- [ ] **Step 6: Commit**

```bash
git add convert.js convert.test.js
git commit -m "주문 검증과 CJ 송장 변환"
```

---

### Task 2: 거래처별 발주서와 파일명

**Files:**
- Modify: `convert.js`
- Test: `convert.test.js`

**Interfaces:**
- Consumes: Task 1의 `convertOrders`, `cleanRows`, `missingColumns`
- Produces:
  - `convertOrders`가 `purchaseOrders`(거래처명 → `PO_COLUMNS` 키를 가진 행 배열)와 `unmatched`를 채운다. 매칭표에 `상품명`/`거래처` 열이 없으면 `{ok:false}`.
  - `safeFileName(name: string) -> string`
  - `purchaseOrderFiles(purchaseOrders: Object<string, object[]>, date: string) -> {fileName: string, rows: object[]}[]` — `발주서_<safe거래처>_<date>.xlsx`, 이름이 겹치면 두 번째부터 `발주서_<safe거래처> (2)_<date>.xlsx`

- [ ] **Step 1: 실패하는 테스트 작성**

```js
const MAP = [{ 상품명: '국산 햇양파 3kg', 거래처: '무안농산' }, { 상품명: '제주 감귤 5kg', 거래처: '제주팜' }];
test('발주서: 거래처별 분리와 열', () => {
  const r = convertOrders([order(), order({ 상품주문번호: '2', 상품명: '제주 감귤 5kg', 옵션정보: '' })], MAP);
  assert.deepEqual(Object.keys(r.purchaseOrders).sort(), ['무안농산', '제주팜']);
  assert.deepEqual(r.purchaseOrders['무안농산'][0], {
    상품주문번호: '2026111000001', 상품명: '국산 햇양파 3kg', 옵션: '특품', 수량: '2', 수취인명: '김철수',
    연락처: '010-1234-5678', 우편번호: '06234', 주소: '서울특별시 강남구 테헤란로 1', 배송메세지: '문 앞',
  });
});
test('미매칭: 송장 포함, 발주서 제외, 경고', () => {
  const r = convertOrders([order({ 상품명: '없는 상품' })], MAP);
  assert.equal(r.labels.length, 1);
  assert.deepEqual(r.purchaseOrders, {});
  assert.deepEqual(r.unmatched, [{ orderNo: '2026111000001', productName: '없는 상품' }]);
});
test('상품명 앞뒤 공백 무시', () => { /* 주문 ' 국산 햇양파 3kg ' + 매칭표 '국산 햇양파 3kg  ' → 무안농산 */ });
test('매칭표 중복은 첫 행, 거래처 빈칸은 미매칭', () => { /* [{양파,A},{양파,B}] → A / [{양파,''}] → unmatched */ });
test('매칭표 열 누락 시 실패', () => { /* [{이름:'x'}] → ok false, error에 '거래처' */ });
test('파일명 금지 문자 치환', () => {
  assert.equal(safeFileName('A/B:C*?"<>|\\'), 'A_B_C_______');
});
test('파일명 충돌 시 (2)', () => {
  const files = purchaseOrderFiles({ 'A/B': [{}], 'A_B': [{}] }, '2026-11-10');
  assert.deepEqual(files.map(f => f.fileName), ['발주서_A_B_2026-11-10.xlsx', '발주서_A_B (2)_2026-11-10.xlsx']);
});
```

- [ ] **Step 2: 실패 확인** — Run: `node --test` / Expected: 새 테스트 FAIL

- [ ] **Step 3: Interfaces대로 구현**

- [ ] **Step 4: 통과 확인** — Run: `node --test` / Expected: 전부 PASS

- [ ] **Step 5: Commit** — `git commit -m "거래처별 발주서와 파일명 처리"`

---

### Task 3: 송장번호 되돌리기

**Files:**
- Modify: `convert.js`
- Test: `convert.test.js`

**Interfaces:**
- Consumes: `cleanRows`, `missingColumns`, `TRACKING_COLUMNS`, `DISPATCH_COLUMNS`, `DISPATCH_FIXED`
- Produces: `convertTracking(rows: object[]) -> {ok: true, rows: object[], missing: string[], duplicates: string[]} | {ok: false, error: string}` — `rows`는 `DISPATCH_COLUMNS` 키, `missing`은 운송장번호가 빈 주문번호, `duplicates`는 두 번 이상 나온 주문번호(첫 번째만 사용)

- [ ] **Step 1: 실패하는 테스트 작성**

```js
test('일괄발송: 변환과 고정값, 앞자리 0 유지', () => {
  const r = convertTracking([{ 고객주문번호: '2026111000001', 운송장번호: '0123456789012' }]);
  assert.deepEqual(r.rows, [{ 상품주문번호: '2026111000001', 배송방법: '택배,등기,소포', 택배사: 'CJ대한통운', 송장번호: '0123456789012' }]);
});
test('운송장번호 빈 행 제외 + 경고', () => { /* {1,''} → rows [], missing ['1'] */ });
test('중복 주문번호는 첫 번째만 + 경고', () => { /* {1,'A'},{1,'B'} → rows 1개 송장번호 'A', duplicates ['1'] */ });
test('송장번호 결과 필수 열 누락', () => { /* [{고객주문번호:'1'}] → ok false, error에 '운송장번호' */ });
```

- [ ] **Step 2: 실패 확인** — `node --test` / 새 테스트 FAIL
- [ ] **Step 3: 구현**
- [ ] **Step 4: 통과 확인** — `node --test` / 전부 PASS
- [ ] **Step 5: Commit** — `git commit -m "송장번호를 스마트스토어 일괄발송 양식으로 변환"`

---

### Task 4: 샘플 데이터

**Files:**
- Create: `sample.js`
- Test: `convert.test.js`

**Interfaces:**
- Consumes: `convertOrders`, `convertTracking`
- Produces: `SAMPLE_ORDERS: object[]`(주문 20건, `ORDER_COLUMNS` 키 + 실제 스마트스토어 파일처럼 `주문번호`, `구매자명` 등 무시되는 열 2~3개), `SAMPLE_MAPPING: object[]`(거래처 3곳: 무안농산·제주팜·강원한우, 상품 6종), `sampleTracking(labels: object[]) -> object[]`(각 송장에 `6` + 11자리 가짜 운송장번호, 마지막 1건은 빈칸). 내보내기 방식은 `convert.js`와 같게(`window.Sample`).
- 가상 데이터: 이름·주소·전화번호는 실존하지 않는 값(전화번호 `010-0000-xxxx`). 주문 중 1건은 매칭표에 없는 상품(`수제 딸기잼 500g`)으로 미매칭 경고를 시연한다.

- [ ] **Step 1: 실패하는 테스트 작성**

```js
test('샘플: 20건, 거래처 3곳, 미매칭 1건', () => {
  const r = convertOrders(SAMPLE_ORDERS, SAMPLE_MAPPING);
  assert.equal(r.ok, true); assert.equal(r.labels.length, 20);
  assert.equal(Object.keys(r.purchaseOrders).length, 3);
  assert.equal(r.unmatched.length, 1);
  const t = convertTracking(sampleTracking(r.labels));
  assert.equal(t.rows.length, 19); assert.equal(t.missing.length, 1);
});
```

- [ ] **Step 2: 실패 확인** — `node --test` / FAIL
- [ ] **Step 3: `sample.js` 작성**
- [ ] **Step 4: 통과 확인** — `node --test` / 전부 PASS
- [ ] **Step 5: Commit** — `git commit -m "샘플 주문·매칭표·송장번호 데이터"`

---

### Task 5: 웹페이지

**Files:**
- Create: `index.html`

**Interfaces:**
- Consumes: `window.Convert`(Task 1~3 전부), `window.Sample`(Task 4)
- Produces: 완성된 페이지

화면 구성(위에서 아래로):
1. 제목, 한 줄 설명, "파일은 이 컴퓨터 밖으로 나가지 않아요" 안내, 「샘플로 해보기」·「샘플 엑셀 내려받기」 버튼
2. **1단계 송장·발주서**: 주문 엑셀 / 매칭표 파일 선택 → 「변환」 → 요약(주문 N건, 거래처 N곳), 미매칭 경고(빨강, 상품명·주문번호 목록), 「송장 내려받기」·「발주서 zip 내려받기」
3. **2단계 송장번호 등록**: 송장번호 결과 파일 선택 → 「변환」 → 요약, 빈 송장번호·중복 경고, 「일괄발송 파일 내려받기」
4. 비밀번호 걸린 파일 안내 문구(1단계 근처)

구현 메모:
- `<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js">`, `.../jszip/3.10.1/jszip.min.js`, 그다음 `convert.js`, `sample.js`
- 읽기: `XLSX.read(arrayBuffer, {type:'array'})` → 첫 시트 → `XLSX.utils.sheet_to_json(ws, {defval:'', raw:false})`. 예외 시 `파일을 열 수 없어요. 스마트스토어에서 비밀번호 없이 다시 내려받아 주세요.`
- 쓰기: `XLSX.utils.json_to_sheet(rows, {header: <해당 COLUMNS>})`. 송장은 `XLSX.writeFile`, 발주서는 `purchaseOrderFiles` 결과를 `XLSX.write(wb, {type:'array', bookType:'xlsx'})`로 JSZip에 넣고 `generateAsync({type:'blob'})` → `<a download>`
- 날짜: 로컬 기준 `YYYY-MM-DD`
- 샘플 엑셀 내려받기: `SAMPLE_ORDERS`, `SAMPLE_MAPPING`, `sampleTracking(...)`을 각각 xlsx로 저장(`샘플_주문.xlsx`, `샘플_매칭표.xlsx`, `샘플_송장번호결과.xlsx`)
- 「샘플로 해보기」: 파일 선택 없이 샘플 데이터로 1·2단계 결과를 바로 표시
- 휴대폰 폭에서도 깨지지 않게 단일 열, 버튼 크게

- [ ] **Step 1: `index.html` 작성**

- [ ] **Step 2: 브라우저 수동 확인** — `index.html`을 브라우저로 열고:
  - 「샘플로 해보기」 → 주문 20건, 거래처 3곳, 미매칭 1건(수제 딸기잼 500g), 2단계 19건 + 빈 송장번호 1건 경고
  - 「샘플 엑셀 내려받기」로 받은 3개 파일을 직접 올려서 같은 결과
  - 송장 xlsx를 열어 전화번호 `010-…`, 우편번호 `0…` 유지 확인
  - zip을 풀어 거래처 파일 3개 확인
  - 엑셀이 아닌 파일(.txt)을 올려 오류 문구 확인
  - 브라우저 폭을 휴대폰 크기로 줄여 가로 스크롤 없음 확인

- [ ] **Step 3: `node --test` 재실행** — Expected: 전부 PASS

- [ ] **Step 4: Commit** — `git commit -m "변환기 웹페이지"`
