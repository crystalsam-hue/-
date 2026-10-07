/**
 * 알림 발송: 이메일(Gmail) + 카카오 알림톡(솔라피 API)
 *
 * 솔라피 API Key/Secret 은 시트가 아니라 스크립트 속성에 저장합니다.
 * (메뉴 [근태관리 > 알림톡 API 키 등록])
 */

var SOLAPI_URL = 'https://api.solapi.com/messages/v4/send';

/** 출근/외출/복귀/퇴근 1건 → 센터장(및 선택적으로 회장)에게 즉시 알림 */
function notifyEvent_(record, center, settings) {
  var title = '[' + record.centerName + '] ' + record.name + ' ' + record.action +
    (record.tags.length ? ' (' + record.tags.join(', ') + ')' : '');
  var lines = [
    '센터: ' + record.centerName,
    '직원: ' + record.name + ' (' + record.employeeId + ')',
    '구분: ' + record.action,
    '시각: ' + record.date + ' ' + record.time,
    '비고: ' + ([record.note].concat(record.tags).filter(String).join(' / ') || '-')
  ];
  var html = '<div style="font-family:sans-serif;font-size:14px">' +
    lines.map(escapeHtml_).join('<br>') +
    '<br><br><a href="' + ss_().getUrl() + '">실시간 현황 보기</a></div>';

  var variables = {
    '#{센터}': record.centerName,
    '#{이름}': record.name,
    '#{구분}': record.action,
    '#{시각}': record.time.slice(0, 5),
    '#{비고}': [record.note].concat(record.tags).filter(String).join(' / ') || '-'
  };

  var emails = [];
  var phones = [];
  if (center.managerEmail) emails.push(center.managerEmail);
  if (center.managerPhone) phones.push(center.managerPhone);
  if (settings.chairmanRealtime) {
    if (settings.chairmanEmail) emails.push(settings.chairmanEmail);
    if (settings.chairmanPhone) phones.push(settings.chairmanPhone);
  }

  if (settings.emailOn && emails.length) {
    MailApp.sendEmail({ to: emails.join(','), subject: title, htmlBody: html, body: lines.join('\n') });
  }
  if (settings.alimtalkOn && settings.eventTemplateId) {
    phones.forEach(function (p) {
      sendAlimtalk_(settings, p, settings.eventTemplateId, variables, title + '\n' + lines.join('\n'));
    });
  }
}

/** 솔라피 알림톡 1건 발송 (실패 시 솔라피가 SMS로 대체 발송) */
function sendAlimtalk_(settings, to, templateId, variables, fallbackText) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('SOLAPI_API_KEY');
  var apiSecret = props.getProperty('SOLAPI_API_SECRET');
  if (!apiKey || !apiSecret || !settings.pfId || !settings.senderPhone) {
    console.warn('알림톡 설정 누락 (API 키 / pfId / 발신번호)');
    return false;
  }

  var payload = {
    message: {
      to: to,
      from: settings.senderPhone,
      text: fallbackText,
      kakaoOptions: {
        pfId: settings.pfId,
        templateId: templateId,
        variables: variables,
        disableSms: false
      }
    }
  };

  var res = UrlFetchApp.fetch(SOLAPI_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: solapiAuthHeader_(apiKey, apiSecret) },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    console.error('알림톡 발송 실패', res.getResponseCode(), res.getContentText());
    return false;
  }
  return true;
}

function solapiAuthHeader_(apiKey, apiSecret) {
  var date = new Date().toISOString();
  var salt = Utilities.getUuid().replace(/-/g, '');
  var bytes = Utilities.computeHmacSha256Signature(date + salt, apiSecret);
  var signature = bytes.map(function (b) {
    var h = (b & 0xff).toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
  return 'HMAC-SHA256 apiKey=' + apiKey + ', date=' + date + ', salt=' + salt + ', signature=' + signature;
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}
