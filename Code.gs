// ============================================================
//  졸업앨범 음성 녹음 서버
//  시트1 'recordings': timestamp | school | name | class | fileName | fileId | mimeType
//  시트2 'passwords' : school | password | downloadEnabled | startDate | endDate
// ============================================================

var SPREADSHEET_ID   = '14nc7vEO8aYd146HQJyT3D0TEkqU_tC0c9i0hO6VDwkQ';
var ROOT_FOLDER_NAME = '졸업앨범_음성';
var SHEET_NAME       = 'recordings';
var PW_SHEET_NAME    = 'passwords';

// ------------------------------------------------------------
//  POST
// ------------------------------------------------------------
function doPost(e) {
  try {
    var raw    = e.postData.contents;
    var data   = JSON.parse(raw);
    var action = data.action || 'upload';

    if (action === 'setPassword')        return setPassword(data);
    if (action === 'deletePassword')     return deletePasswordRow(data.school);
    if (action === 'setDownloadEnabled') return setDownloadEnabled(data.school, data.enabled);
    if (action === 'deleteRecording')    return deleteRecording(data.fileId);

    // ── 음성 업로드 ──
    var school    = data.school;
    var name      = data.name;
    var className = data.className;
    var audioB64  = data.audioBase64;
    var mimeType  = data.mimeType;

    if (!school || !name || !className || !audioB64)
      return makeJson({ success: false, error: '필수 항목 누락' });

    var sheet   = getSheet();
    var allRows = sheet.getDataRange().getValues();
    for (var i = 1; i < allRows.length; i++) {
      if (allRows[i][1] === school && allRows[i][2] === name && allRows[i][3] === className)
        return makeJson({ success: false, error: 'ALREADY_SUBMITTED' });
    }

    var rootFolder   = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
    var schoolFolder = getOrCreateFolder(rootFolder, school);

    var ext = 'webm';
    if (mimeType.indexOf('mp4') !== -1) ext = 'm4a';
    if (mimeType.indexOf('ogg') !== -1) ext = 'ogg';

    var fileName = name + '_' + className + '.' + ext;
    var bytes    = Utilities.base64Decode(audioB64);
    var blob     = Utilities.newBlob(bytes, mimeType, fileName);
    var file     = schoolFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    sheet.appendRow([new Date(), school, name, className, fileName, file.getId(), mimeType]);
    return makeJson({ success: true, fileId: file.getId() });

  } catch (err) {
    return makeJson({ success: false, error: err.message });
  }
}

// ------------------------------------------------------------
//  GET
// ------------------------------------------------------------
function doGet(e) {
  var action = e.parameter.action || '';

  if (action === 'schools') {
    var values  = getSheet().getDataRange().getValues();
    var schools = [];
    for (var i = 1; i < values.length; i++) {
      var s = values[i][1];
      if (s && schools.indexOf(s) === -1) schools.push(s);
    }
    schools.sort();
    return makeJson({ schools: schools });
  }

  if (action === 'checkPassword') {
    var school = e.parameter.school || '';
    var pw     = e.parameter.pw     || '';
    if (!school || !pw) return makeJson({ ok: false });
    var pwRows = getPwSheet().getDataRange().getValues();
    for (var k = 1; k < pwRows.length; k++) {
      if (String(pwRows[k][0]).trim() === school)
        return makeJson({ ok: String(pwRows[k][1]).trim() === pw.trim() });
    }
    return makeJson({ ok: true, noPassword: true });
  }

  if (action === 'list') {
    var school2 = e.parameter.school || '';
    var all     = getSheet().getDataRange().getValues();
    var students = [];
    for (var j = 1; j < all.length; j++) {
      if (all[j][1] !== school2) continue;
      students.push({
        school: all[j][1], name: all[j][2], className: all[j][3],
        fileName: all[j][4], fileId: all[j][5],
        timestamp: all[j][0] ? all[j][0].toString() : ''
      });
    }
    return makeJson({ students: students });
  }

  // ── 빠른 메타데이터 (Drive 접근 없음, ~0.5초) ──
  if (action === 'meta') {
    try {
      var fid   = e.parameter.f || '';
      var rows3 = getSheet().getDataRange().getValues();
      var found = null;
      for (var n = 1; n < rows3.length; n++) {
        if (rows3[n][5] === fid) { found = rows3[n]; break; }
      }
      if (!found) return makeJson({ ok: false, error: '녹음을 찾을 수 없습니다' });
      var pwRows4 = getPwSheet().getDataRange().getValues();
      var dlEnabled = true;
      for (var q = 1; q < pwRows4.length; q++) {
        if (String(pwRows4[q][0]).trim() === String(found[1]).trim()) {
          var v = pwRows4[q][2];
          if (v === false || String(v).toLowerCase() === 'false') dlEnabled = false;
          break;
        }
      }
      return makeJson({
        ok: true,
        school: found[1], name: found[2], className: found[3],
        downloadEnabled: dlEnabled
      });
    } catch(emeta) {
      return makeJson({ ok: false, error: emeta.message });
    }
  }

  // ── 오디오 데이터 (b64, ~3~10초) — 재생 + 다운로드 공용 ──
  if (action === 'audio') {
    try {
      var fid2  = e.parameter.f || '';
      var rows5 = getSheet().getDataRange().getValues();
      var found2 = null;
      for (var r2 = 1; r2 < rows5.length; r2++) {
        if (rows5[r2][5] === fid2) { found2 = rows5[r2]; break; }
      }
      if (!found2) return makeJson({ ok: false, error: '녹음을 찾을 수 없습니다' });
      var f5  = DriveApp.getFileById(fid2);
      var bl5 = f5.getBlob();
      // downloadEnabled 확인 (meta 실패 시 폴백으로 사용)
      var pwRowsAu = getPwSheet().getDataRange().getValues();
      var dlAu = true;
      for (var qa = 1; qa < pwRowsAu.length; qa++) {
        if (String(pwRowsAu[qa][0]).trim() === String(found2[1]).trim()) {
          var va = pwRowsAu[qa][2];
          if (va === false || String(va).toLowerCase() === 'false') dlAu = false;
          break;
        }
      }
      return makeJson({
        ok: true,
        b64:  Utilities.base64Encode(bl5.getBytes()),
        mime: bl5.getContentType(),
        school: found2[1], name: found2[2], className: found2[3],
        downloadEnabled: dlAu
      });
    } catch(ex) {
      return makeJson({ ok: false, error: ex.message });
    }
  }

  if (action === 'getPasswords') {
    var pwAll = getPwSheet().getDataRange().getValues();
    var list  = [];
    for (var p = 1; p < pwAll.length; p++) {
      if (!pwAll[p][0]) continue;
      var dv2 = pwAll[p][2];
      list.push({
        school:          String(pwAll[p][0]).trim(),
        password:        String(pwAll[p][1]).trim(),
        downloadEnabled: !(dv2 === false || String(dv2).toLowerCase() === 'false'),
        startDate:       pwAll[p][3] ? String(pwAll[p][3]).trim() : '',
        endDate:         pwAll[p][4] ? String(pwAll[p][4]).trim() : ''
      });
    }
    return makeJson({ passwords: list });
  }

  return makeJson({ error: '알 수 없는 요청' });
}

// ------------------------------------------------------------
//  비밀번호 + 기간 저장 (passwords 시트 A~E열)
// ------------------------------------------------------------
function setPassword(data) {
  var school    = (data.school    || '').trim();
  var password  = (data.password  || '').trim();
  var startDate = (data.startDate || '').trim();
  var endDate   = (data.endDate   || '').trim();

  if (!school || !password)
    return makeJson({ success: false, error: '학교명과 비밀번호를 입력하세요' });

  var sheet = getPwSheet();
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === school) {
      sheet.getRange(i + 1, 2).setValue(password);
      sheet.getRange(i + 1, 4).setValue(startDate);
      sheet.getRange(i + 1, 5).setValue(endDate);
      return makeJson({ success: true, updated: true });
    }
  }
  sheet.appendRow([school, password, true, startDate, endDate]);
  return makeJson({ success: true, created: true });
}

// ------------------------------------------------------------
//  다운로드 토글
// ------------------------------------------------------------
function setDownloadEnabled(school, enabled) {
  if (!school) return makeJson({ success: false, error: '학교명이 없습니다' });
  var sheet = getPwSheet();
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(school).trim()) {
      sheet.getRange(i + 1, 3).setValue(enabled === true || enabled === 'true');
      return makeJson({ success: true });
    }
  }
  return makeJson({ success: false, error: '해당 학교를 찾을 수 없습니다' });
}

// ------------------------------------------------------------
//  비밀번호 행 삭제
// ------------------------------------------------------------
function deletePasswordRow(school) {
  if (!school) return makeJson({ success: false, error: '학교명이 없습니다' });
  var sheet = getPwSheet();
  var rows  = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() === String(school).trim()) {
      sheet.deleteRow(i + 1);
      return makeJson({ success: true });
    }
  }
  return makeJson({ success: false, error: '해당 학교를 찾을 수 없습니다' });
}

// ------------------------------------------------------------
//  녹음 삭제 (Drive 파일 휴지통 + recordings 행 삭제)
// ------------------------------------------------------------
function deleteRecording(fileId) {
  if (!fileId) return makeJson({ success: false, error: 'fileId 없음' });
  try {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch(fe) { /* 이미 없음 */ }
    var sheet = getSheet();
    var rows  = sheet.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][5]).trim() === String(fileId).trim()) {
        sheet.deleteRow(i + 1);
        return makeJson({ success: true });
      }
    }
    return makeJson({ success: true });
  } catch(err) {
    return makeJson({ success: false, error: err.message });
  }
}

// ------------------------------------------------------------
//  헬퍼
// ------------------------------------------------------------
function getOrCreateFolder(parent, name) {
  var iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['timestamp','school','name','class','fileName','fileId','mimeType']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getPwSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(PW_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PW_SHEET_NAME);
    sh.appendRow(['school','password','downloadEnabled','startDate','endDate']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function makeJson(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
