/**
 * 6학년 영어 3단원 평가 결과 수신용 Apps Script
 *
 * 사용 방법:
 * 1) 결과를 받을 구글 스프레드시트를 새로 만든다.
 * 2) 확장 프로그램 → Apps Script 메뉴를 연다.
 * 3) 이 파일(code.gs)의 내용을 모두 복사해서 붙여넣고 저장한다.
 * 4) 우측 상단 "배포" → "새 배포" → 유형 "웹 앱" 선택.
 *    - 실행할 사용자: "나"
 *    - 액세스 권한: "모든 사용자"
 * 5) 배포 완료 후 표시되는 "웹 앱 URL"을 복사한다.
 * 6) index.html 파일의 GAS_URL = ""  안에 그 URL을 붙여넣고 저장 → 푸시.
 *
 * 학생들이 시험을 제출하면 스프레드시트의 "응답" 시트에 한 줄씩 자동 기록됩니다.
 */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20 * 1000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('응답');
    if (!sheet) sheet = ss.insertSheet('응답');

    // payload 파싱 (form-encoded 또는 raw JSON 모두 지원)
    let data = null;
    if (e && e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch (err) { data = null; }
    }
    if (!data) return jsonOut({ ok: false, error: 'No payload' });

    // 헤더 자동 작성 (시트가 비어 있을 때만)
    if (sheet.getLastRow() === 0) {
      const headers = [
        '전송 일시', '학년', '반', '번호', '이름',
        '총점', '이해', '표현', '소요(초)'
      ];
      for (let i = 1; i <= 20; i++) headers.push(i + '번');
      // 표현 영역의 단답형 문항(11~15, 17, 18, 20번) 원문 답안도 기록
      headers.push('11번 답안', '12번 답안', '13번 답안', '14번 답안', '15번 답안',
                   '17번 답안', '18번 답안', '20번 답안');
      headers.push('시작 시각', '종료 시각');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#fff3cd');
    }

    const marks = data.marks || [];
    const sa = data.shortAnswers || {};
    const row = [
      new Date(),
      data.grade || '',
      data.cls || '',
      data.num || '',
      data.name || '',
      data.score, data.recScore, data.repScore,
      data.durationSec
    ];
    for (let i = 0; i < 20; i++) row.push(marks[i] || '');
    row.push(sa['11'] || '');
    row.push(sa['12'] || '');
    row.push(sa['13'] || '');
    row.push(sa['14'] || '');
    row.push(sa['15'] || '');
    row.push(sa['17'] || '');
    row.push(sa['18'] || '');
    row.push(sa['20'] || '');
    row.push(data.startTime ? new Date(data.startTime) : '');
    row.push(data.endTime ? new Date(data.endTime) : '');

    sheet.appendRow(row);
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return jsonOut({ ok: true, message: '6학년 영어 3단원 평가 결과 수신 서버 (POST 전용)' });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
