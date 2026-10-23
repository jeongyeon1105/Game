// 레크레이션 가사 읽어주기: 구글 시트('곡')에 저장해서 링크로 접속한 모든 기기가 같은 데이터를 본다.
const HEAD = ['id', '구분', '제목', '가수', '가사'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('레크레이션 퀴즈')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 시트가 없으면 만든다. '@' 서식으로 두어 숫자·날짜 변환과 수식 실행을 막는다.
function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('곡');
  if (!s) {
    s = ss.insertSheet('곡');
    s.getRange('A:E').setNumberFormat('@');
    s.appendRow(HEAD);
    s.setFrozenRows(1);
  }
  return s;
}

function read_(s) {
  return s.getDataRange().getValues().slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return {id: r[0], type: r[1], title: r[2], singer: r[3], lyrics: r[4]};
  });
}

// 모든 쓰기의 공통 흐름: 잠금 → 작업 → 최신 목록 반환
function write_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sheet_();
    fn(s);
    return read_(s);
  } finally {
    lock.releaseLock();
  }
}

function getData() {
  return read_(sheet_());
}

function rowOf_(s, id) {
  return s.getDataRange().getValues().findIndex(function (r, i) { return i > 0 && r[0] === id; }) + 1;
}

// song.id가 있으면 수정, 없으면 추가
function saveSong(song) {
  const type = song.type === 'dance' ? 'dance' : 'lyrics';
  const title = String(song.title || '').trim().slice(0, 100);
  const singer = String(song.singer || '').trim().slice(0, 100);
  const lyrics = String(song.lyrics || '').trim().slice(0, 20000);
  if (!title || (type === 'lyrics' && !lyrics)) throw new Error('입력값을 확인해 주세요.');
  return write_(function (s) {
    const row = song.id ? rowOf_(s, song.id) : 0;
    if (song.id && !row) throw new Error('이미 삭제된 문제입니다.');
    s.getRange(row || s.getLastRow() + 1, 1, 1, 5).setValues([[song.id || Utilities.getUuid(), type, title, singer, lyrics]]);
  });
}

function deleteSong(id) {
  return write_(function (s) {
    const row = rowOf_(s, id);
    if (row) s.deleteRow(row);
  });
}
