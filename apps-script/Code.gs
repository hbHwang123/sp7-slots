/**
 * SP7 음색 슬롯 예약 — Google Apps Script 백엔드
 *
 * 시트 "slots" 에 칸 하나가 한 줄로 저장됩니다.
 * 열: kb | slot | name | purpose | start | end | pinHash | updatedAt | batch
 *   kb    : new(신기체) / old(구기체)
 *   slot  : A1 ~ D16
 *   start/end : YYYY-MM-DD (텍스트)
 *   pinHash   : PIN의 해시값 (PIN 원문은 저장하지 않음)
 *   batch     : 한 번에 함께 예약한 칸들이 공유하는 묶음 번호
 *
 * 관리자 PIN: [프로젝트 설정 → 스크립트 속성]에 ADMIN_PIN 을 넣으면
 *            그 PIN으로 모든 예약을 수정·취소할 수 있습니다.
 */

const SHEET = 'slots';
const HEAD = ['kb', 'slot', 'name', 'purpose', 'start', 'end', 'pinHash', 'updatedAt', 'batch'];
const C = { kb: 0, slot: 1, name: 2, purpose: 3, start: 4, end: 5, pin: 6, at: 7, batch: 8 };
const KBS = ['new', 'old'];
const SLOT_RE = /^[ABCD]([1-9]|1[0-6])$/;
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const PIN_RE = /^\d{4}$/;

/* ---------- 진입점 ---------- */

function doGet() {
  return out({ ok: true, rows: listRows() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: '요청 형식이 잘못됐어요.' }); }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return out({ ok: false, error: '사람이 몰려 있어요. 잠시 후 다시 시도해 주세요.' });
  try {
    const slots = cleanSlots(req.slots);
    if (!slots) return out({ ok: false, error: '잘못된 칸이 포함돼 있어요.', rows: listRows() });
    switch (req.action) {
      case 'reserve': return out(reserve(req, slots));
      case 'update':  return out(update(req, slots));
      case 'cancel':  return out(cancel(req, slots));
      case 'clear':   return out(clearExpired(slots));
      default:        return out({ ok: false, error: '알 수 없는 요청이에요.' });
    }
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 동작 (여러 칸을 한 번에, 전부 성공하거나 전부 취소) ---------- */

function reserve(r, slots) {
  const v = validate(r, true);
  if (v) return { ok: false, error: v };
  const index = rowIndex();
  const taken = slots.filter(s => index[s.kb + '_' + s.slot]);
  if (taken.length) return { ok: false, error: '이미 예약된 칸이 있어요: ' + names(taken), rows: listRows() };
  const batch = String(r.batch || Utilities.getUuid()).replace(/[^\w-]/g, '').slice(0, 24);
  const now = new Date().toISOString(), h = hash(r.pin);
  const data = slots.map(s => [s.kb, s.slot, clean(r.name), clean(r.purpose), r.start, r.end, h, now, batch]);
  const sh = sheet();
  sh.getRange(sh.getLastRow() + 1, 1, data.length, HEAD.length).setValues(data);
  return { ok: true, rows: listRows() };
}

function update(r, slots) {
  const v = validate(r, false);
  if (v) return { ok: false, error: v };
  const found = findAll(slots);
  if (found.missing.length) return { ok: false, error: '예약이 없어진 칸이 있어요: ' + names(found.missing), rows: listRows() };
  if (found.rows.some(x => !pinOk(r.pin, x.values[C.pin]))) return { ok: false, error: 'PIN이 맞지 않아요.' };
  const sh = sheet(), now = new Date().toISOString();
  found.rows.forEach(x => sh.getRange(x.index, C.name + 1, 1, 4).setValues([[clean(r.name), clean(r.purpose), r.start, r.end]]));
  found.rows.forEach(x => sh.getRange(x.index, C.at + 1).setValue(now));
  return { ok: true, rows: listRows() };
}

function cancel(r, slots) {
  const found = findAll(slots);
  if (found.rows.some(x => !pinOk(r.pin, x.values[C.pin]))) return { ok: false, error: 'PIN이 맞지 않아요.' };
  deleteRows(found.rows);
  return { ok: true, rows: listRows() };
}

/** 기간이 지난 예약은 PIN 없이 누구나 비울 수 있음 */
function clearExpired(slots) {
  const found = findAll(slots), t = today();
  const live = found.rows.filter(x => !(String(x.values[C.end]) < t));
  if (live.length) return { ok: false, error: '아직 기간이 끝나지 않은 칸이 있어요: ' + names(live.map(x => ({ kb: x.values[C.kb], slot: x.values[C.slot] }))) };
  deleteRows(found.rows);
  return { ok: true, rows: listRows() };
}

/* ---------- 도우미 ---------- */

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange('A:I').setNumberFormat('@'); // 날짜가 자동 변환되지 않도록 텍스트로
    sh.setFrozenRows(1);
  }
  if (sh.getRange(1, HEAD.length).getValue() !== HEAD[HEAD.length - 1]) { // 이전 버전 시트면 batch 열 추가
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.getRange('I:I').setNumberFormat('@');
  }
  return sh;
}

function allValues() {
  const sh = sheet();
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, HEAD.length).getDisplayValues();
}

function listRows() {
  return allValues().filter(v => v[C.kb] && v[C.slot]).map(v => ({
    kb: v[C.kb], slot: v[C.slot], name: v[C.name], purpose: v[C.purpose], start: v[C.start], end: v[C.end], batch: v[C.batch]
  }));
}

function rowIndex() {
  const idx = {};
  allValues().forEach((v, i) => { if (v[C.kb]) idx[v[C.kb] + '_' + v[C.slot]] = { index: i + 2, values: v }; });
  return idx;
}

function findAll(slots) {
  const idx = rowIndex(), rows = [], missing = [];
  slots.forEach(s => { const x = idx[s.kb + '_' + s.slot]; if (x) rows.push(x); else missing.push(s); });
  return { rows, missing };
}

function deleteRows(found) {
  const sh = sheet();
  found.map(x => x.index).sort((a, b) => b - a).forEach(i => sh.deleteRow(i)); // 아래쪽부터 지워야 번호가 안 밀림
}

function cleanSlots(list) {
  if (!Array.isArray(list) || !list.length || list.length > 128) return null;
  const seen = {}, out = [];
  for (const s of list) {
    if (!s || KBS.indexOf(s.kb) < 0 || !SLOT_RE.test(s.slot || '')) return null;
    const k = s.kb + '_' + s.slot;
    if (!seen[k]) { seen[k] = 1; out.push({ kb: s.kb, slot: s.slot }); }
  }
  return out;
}

function names(slots) {
  const label = { new: '신기체', old: '구기체' };
  return slots.map(s => label[s.kb] + ' ' + s.slot).join(', ');
}

function validate(r, needPin) {
  if (!clean(r.name)) return '사용자 이름을 입력해 주세요.';
  if (!clean(r.purpose)) return '사용 목적을 입력해 주세요.';
  if (!DATE_RE.test(r.start || '') || !DATE_RE.test(r.end || '')) return '날짜 형식이 잘못됐어요.';
  if (r.end < r.start) return '종료일이 시작일보다 빠를 수 없어요.';
  if (!PIN_RE.test(r.pin || '')) return needPin ? 'PIN은 숫자 4자리로 정해 주세요.' : 'PIN을 입력해 주세요.';
  return '';
}

function clean(s) {
  // 시트 수식으로 해석되지 않도록 앞의 = + - @ 제거, 길이 제한
  return String(s || '').trim().replace(/^[=+\-@]+/, '').slice(0, 40);
}

function pinOk(pin, stored) {
  if (!PIN_RE.test(pin || '')) return false;
  const admin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (admin && pin === admin) return true;
  return hash(pin) === stored;
}

function hash(pin) {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('SALT');
  if (!salt) { salt = Utilities.getUuid(); props.setProperty('SALT', salt); }
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin);
  return bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function today() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 처음 한 번 실행: 시트를 만들고 권한을 승인합니다. */
function setup() {
  sheet();
  hash('0000');
  Logger.log('준비 완료: "' + SHEET + '" 시트가 생성됐어요.');
}
