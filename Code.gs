/**
 * PetConnect SOS - Google Apps Script Backend
 * 
 * [구글 시트 설정 방법]
 * 1. 구글 시트를 새로 만듭니다.
 * 2. 첫 번째 시트의 이름을 'DB'로 변경합니다. (또는 코드 아래의 SHEET_NAME을 변경하세요)
 * 3. 1행에 다음 헤더를 차례로 입력합니다:
 *    A1: 제품ID
 *    B1: 등록여부
 *    C1: 보호자이메일
 *    D1: 보호자연락처
 *    E1: 강아지이름
 *    F1: 특이사항
 *    G1: 비밀번호
 * 
 * 4. 2행부터 제품ID를 미리 입력해둡니다. (예: H00001, H00002 ...)
 *    이때 등록여부는 'N'으로 입력해둡니다.
 * 
 * [배포 방법]
 * 1. 메뉴에서 "확장 프로그램 > Apps Script"를 클릭합니다.
 * 2. 이 코드를 복사하여 붙여넣고 저장합니다.
 * 3. 우측 상단 "배포 > 새 배포"를 클릭합니다.
 * 4. 유형: "웹 앱", 실행하는 사용자: "나", 액세스할 수 있는 사용자: "모든 사용자"로 설정하고 배포합니다.
 * 5. 생성된 "웹 앱 URL"을 복사하여 HTML 파일들(index.html, register.html, find.html)의 WEB_APP_URL 변수에 넣습니다.
 */

const SHEET_NAME = 'DB';

function doGet(e) {
  // CORS 처리를 위해 JSONP 형태나 단순 JSON 반환
  // GET 요청은 주로 기기 등록 여부 확인용입니다.
  const action = e.parameter.action;
  
  if (action === 'check') {
    return handleCheckDevice(e.parameter.id);
  }
  
  return createJsonResponse({ status: 'error', message: 'Unknown GET action' });
}

function doPost(e) {
  // POST 요청은 JSON 형태로 body를 통해 들어온다고 가정 (fetch API)
  let params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch(error) {
    return createJsonResponse({ status: 'error', message: 'Invalid JSON payload' });
  }

  const action = params.action;

  if (action === 'register') {
    return handleRegister(params);
  } else if (action === 'sos') {
    return handleSOS(params);
  }

  return createJsonResponse({ status: 'error', message: 'Unknown POST action' });
}

// 1. 기기 상태 확인 (없으면 자동 생성)
function handleCheckDevice(id) {
  if (!id) return createJsonResponse({ status: 'error', message: 'No ID provided' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // 1행은 헤더이므로 2행부터 검색
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const isRegistered = data[i][1]; // B열: 등록여부
      
      if (isRegistered === 'Y') {
        return createJsonResponse({
          status: 'success',
          data: {
            isRegistered: 'Y',
            petName: data[i][4], // E열
            petDesc: data[i][5]  // F열
          }
        });
      } else {
        return createJsonResponse({
          status: 'success',
          data: {
            isRegistered: 'N'
          }
        });
      }
    }
  }

  // DB에 없는 새로운 ID가 들어왔을 경우: 시트 맨 아래에 N 상태로 자동 추가
  sheet.appendRow([id, 'N', '', '', '', '', '']);
  
  return createJsonResponse({
    status: 'success',
    data: {
      isRegistered: 'N'
    }
  });
}

// 2. 보호자 정보 등록 (없으면 자동 생성 후 등록)
function handleRegister(params) {
  const { id, email, phone, petName, petDesc, password } = params;
  
  if(!id || !email || !phone || !password) {
    return createJsonResponse({ status: 'error', message: 'Missing required fields' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      if (data[i][1] === 'Y') {
        return createJsonResponse({ status: 'error', message: 'Device already registered' });
      }
      
      // 데이터 업데이트 (행 번호는 i+1)
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, 2).setValue('Y');          // 등록여부
      sheet.getRange(rowIndex, 3).setValue(email);        // 이메일
      sheet.getRange(rowIndex, 4).setValue(phone);        // 연락처
      sheet.getRange(rowIndex, 5).setValue(petName);      // 강아지이름
      sheet.getRange(rowIndex, 6).setValue(petDesc);      // 특이사항
      sheet.getRange(rowIndex, 7).setValue(password);     // 비밀번호

      return createJsonResponse({ status: 'success', message: 'Registered successfully' });
    }
  }
  
  // DB에 없는 ID인데 바로 가입 시도가 온 경우: 바로 Y 상태로 자동 추가
  sheet.appendRow([id, 'Y', email, phone, petName, petDesc, password]);
  
  return createJsonResponse({ status: 'success', message: 'Registered successfully' });
}

// 3. SOS 위치/연락처 전송
function handleSOS(params) {
  const { id, finderPhone, lat, lng } = params;
  
  if(!id || !finderPhone) {
    return createJsonResponse({ status: 'error', message: 'Missing finder information' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id && data[i][1] === 'Y') {
      const ownerEmail = data[i][2];
      const petName = data[i][4];
      
      // 이메일 내용 구성
      const subject = `🚨 [PetConnect SOS] '${petName}'(을)를 보호 중인 제보자가 있습니다!`;
      
      let body = `보호자님, 누군가 반려동물 '${petName}'(을)를 발견하고 SOS 신호를 보냈습니다!\n\n`;
      body += `📞 제보자 연락처: ${finderPhone}\n`;
      
      if (lat && lng) {
        body += `📍 발견 위치 (구글 지도): https://www.google.com/maps?q=${lat},${lng}\n\n`;
      } else {
        body += `📍 발견 위치: 제보자 기기에서 GPS 수집이 거부되어 위치를 알 수 없습니다. 제보자에게 전화로 문의해주세요.\n\n`;
      }
      
      body += `최대한 빨리 제보자에게 연락해 보시기 바랍니다.\n- PetConnect SOS 시스템`;

      // 구글 메일(Gmail) 앱을 통해 무료 이메일 전송
      try {
        GmailApp.sendEmail(ownerEmail, subject, body);
        return createJsonResponse({ status: 'success', message: 'SOS signal sent' });
      } catch (err) {
        return createJsonResponse({ status: 'error', message: 'Failed to send email: ' + err.toString() });
      }
    }
  }

  return createJsonResponse({ status: 'error', message: 'Device not found or not registered' });
}

// JSON 응답을 만들어주는 헬퍼 함수
function createJsonResponse(responseObject) {
  return ContentService
    .createTextOutput(JSON.stringify(responseObject))
    .setMimeType(ContentService.MimeType.JSON);
}
