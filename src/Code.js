function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('AIRBUS WORKSHEET PORTAL')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 이미지 및 폴더 링크를 HTML 태그로 변환하는 함수
function getImageUrlAsHtml(rawUrl) {
  if (!rawUrl) return "";
  
  const urls = rawUrl.split(/[\n,\s]+/).filter(u => u.includes("http"));
  if (urls.length === 0) return "";

  let html = "";
  urls.forEach(url => {
    let driveMatch = url.match(/[-\w]{25,}/); 
    
    if (url.includes("folders") || url.includes("folderview")) {
      html += `<br><a href="${url}" target="_blank" style="display:inline-block; margin-top:10px; background:#f0fdf4; color:#166534; padding:8px 14px; border-radius:4px; font-weight:bold; text-decoration:none; border:1px solid #bbf7d0;">📁 관련 이미지 폴더 열기 (새창)</a>`;
    }
    else if (url.includes("drive.google.com") && driveMatch) {
      html += `<br><img src="https://drive.google.com/uc?export=view&id=${driveMatch[0]}" style="max-width:100%; height:auto; margin-top:15px; border-radius:4px; border: 1px solid #ccc; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">`;
    } 
    else if (url.match(/\.(jpeg|jpg|gif|png|webp|bmp)/i)) {
      html += `<br><img src="${url}" style="max-width:100%; height:auto; margin-top:15px; border-radius:4px; border: 1px solid #ccc; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">`;
    } 
    else {
      html += `<br><a href="${url}" target="_blank" style="display:inline-block; margin-top:10px; background:#e0f2fe; color:#005587; padding:6px 12px; border-radius:4px; font-weight:bold; text-decoration:none; border:1px solid #bae6fd;">🔗 관련 링크 열기</a>`;
    }
  });
  return html;
}

// 시트의 모든 데이터를 하나로 취합하는 함수
function getAllPortalData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const db = { jascRaw: [], cabinRaw: [], cabinDesc: [], fml: [], gii: [], giiDesc: [], nef: [], nefDesc: [] };

    // 일반 데이터 로드 함수
    const getSheetData = (name) => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error(`'${name}' 시트를 찾을 수 없습니다.`);
      return sheet.getDataRange().getDisplayValues(); 
    };

    // 정렬(Alignment) 데이터까지 함께 로드하는 함수
    const getSheetDataWithAlign = (name) => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error(`'${name}' 시트를 찾을 수 없습니다.`);
      const range = sheet.getDataRange();
      return {
        values: range.getDisplayValues(),
        aligns: range.getHorizontalAlignments()
      };
    };

    // 1. JASC CODE
    db.jascRaw = getSheetData("JASC_CODE");

    // 2. CABIN CODE
    const cabinData = getSheetData("CABIN_CODE");
    db.cabinRaw = [["Problem Code", "Problem Meaning"]];
    for (let i = 1; i < cabinData.length; i++) {
      if (cabinData[i][0]) {
        db.cabinRaw.push([ cabinData[i][0].trim(), cabinData[i][1] ? cabinData[i][1].trim() : "" ]);
      }
    }

    const cabinDescData = getSheetData("CABIN_DESC");
    for (let i = 1; i < cabinDescData.length; i++) {
      if (cabinDescData[i][0] !== "") { 
        db.cabinDesc.push({ 
          stt: String(cabinDescData[i][0]).trim(), 
          category: String(cabinDescData[i][1]).trim(),
          reporter: String(cabinDescData[i][2]).trim(),
          classification: String(cabinDescData[i][3]).trim()
        });
      }
    }

    // 3. FML LIST
    const fmlData = getSheetData("FML_LIST");
    for (let i = 1; i < fmlData.length; i++) {
      const taskType = fmlData[i][0] ? String(fmlData[i][0]).trim() : "";
      const desc = fmlData[i][1] ? String(fmlData[i][1]).trim() : "";
      const logStatus = fmlData[i][2] ? String(fmlData[i][2]).trim() : "-";
      const remark = fmlData[i][3] ? String(fmlData[i][3]).trim() : "-";

      if (taskType !== "" || desc !== "") {
        db.fml.push({ label: `${taskType} : ${desc}`, log: logStatus, note: remark });
      }
    }

    // 4. GII_RII_LIST 
    const giiData = getSheetData("GII_RII_LIST");
    for (let i = 1; i < giiData.length; i++) {
      if (giiData[i][0] !== "") { 
        db.gii.push({
          fleet: giiData[i][0] ? String(giiData[i][0]).trim() : "",
          ata: giiData[i][1] ? String(giiData[i][1]).trim() : "",
          type: giiData[i][2] ? String(giiData[i][2]).trim() : "",
          subject: giiData[i][3] ? String(giiData[i][3]).trim() : "", 
          note: giiData[i][4] ? String(giiData[i][4]).trim() : "",    
          giiRii: giiData[i][5] ? String(giiData[i][5]).trim() : ""   
        });
      }
    }

    // GII_DESC (정렬 데이터 포함)
    const giiDescObj = getSheetDataWithAlign("GII_DESC");
    for (let i = 0; i < giiDescObj.values.length; i++) {
      if (giiDescObj.values[i][0]) {
        let title = String(giiDescObj.values[i][0]).trim();
        let content = giiDescObj.values[i][1] ? String(giiDescObj.values[i][1]).trim().replace(/\r?\n/g, '<br>') : "";
        let imgUrl = giiDescObj.values[i][2] ? String(giiDescObj.values[i][2]).trim() : "";
        let imgHtml = getImageUrlAsHtml(imgUrl);
        
        // 구글 시트 정렬 값 적용 (general은 left로 변환)
        let tAlign = giiDescObj.aligns[i][0] === 'general' ? 'left' : giiDescObj.aligns[i][0];
        let cAlign = giiDescObj.aligns[i][1] === 'general' ? 'left' : giiDescObj.aligns[i][1];

        let finalHtml = `<div style="margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px dashed #ccc;">`;
        finalHtml += `<div style="text-align: ${tAlign}; font-weight: bold; margin-bottom: 5px;">${title}</div>`;
        if (content) finalHtml += `<div style="text-align: ${cAlign};">${content}</div>`;
        if (imgHtml) finalHtml += `<div style="text-align: center;">${imgHtml}</div>`;
        finalHtml += `</div>`;
        
        db.giiDesc.push(finalHtml);
      }
    }

    // 5. NEF LIST
    const nefData = getSheetData("NEF_LIST");
    for (let i = 1; i < nefData.length; i++) {
      if (nefData[i][1]) { 
        db.nef.push({
          head: nefData[i][0] ? String(nefData[i][0]).trim() : "",
          mainSub: nefData[i][1] ? String(nefData[i][1]).trim() : "",
          number: nefData[i][2] ? String(nefData[i][2]).trim() : "",
          discrepancy: nefData[i][3] ? String(nefData[i][3]).trim() : "",
          note: nefData[i][4] ? String(nefData[i][4]).trim() : ""
        });
      }
    }

    // NEF_DESC (정렬 데이터 포함)
    const nefDescObj = getSheetDataWithAlign("NEF_DESC");
    for (let i = 0; i < nefDescObj.values.length; i++) {
      if (nefDescObj.values[i][0]) {
        let title = String(nefDescObj.values[i][0]).trim();
        let content = nefDescObj.values[i][1] ? String(nefDescObj.values[i][1]).trim().replace(/\r?\n/g, '<br>') : "";
        let imgUrl = nefDescObj.values[i][2] ? String(nefDescObj.values[i][2]).trim() : "";
        let imgHtml = getImageUrlAsHtml(imgUrl);
        
        // 구글 시트 정렬 값 적용
        let tAlign = nefDescObj.aligns[i][0] === 'general' ? 'left' : nefDescObj.aligns[i][0];
        let cAlign = nefDescObj.aligns[i][1] === 'general' ? 'left' : nefDescObj.aligns[i][1];

        let finalHtml = `<div style="margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px dashed #ccc;">`;
        finalHtml += `<div style="text-align: ${tAlign}; font-weight: bold; margin-bottom: 5px;">${title}</div>`;
        if (content) finalHtml += `<div style="text-align: ${cAlign};">${content}</div>`;
        if (imgHtml) finalHtml += `<div style="text-align: center;">${imgHtml}</div>`;
        finalHtml += `</div>`;

        db.nefDesc.push(finalHtml);
      }
    }

    return db;
  } catch (error) {
    throw new Error(error.message);
  }
}

// 오프라인 파일(단일 HTML) 생성기
function downloadOfflinePortal() {
  try {
    const dbData = getAllPortalData();
    let htmlContent = HtmlService.createHtmlOutputFromFile('Index').getContent();
    
    // HTML 내부의 변수에 DB 데이터를 통째로 주입 (Embedding)
    const jsonString = JSON.stringify(dbData);
    htmlContent = htmlContent.replace('const OFFLINE_DATA = null;', `const OFFLINE_DATA = ${jsonString};`);
    
    // 한글 깨짐 방지를 위한 Base64 인코딩
    const encodedHtml = Utilities.base64Encode(Utilities.newBlob(htmlContent).getBytes());

    const dialogHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: 'Malgun Gothic', sans-serif; text-align: center; padding: 20px; background: #f8fafc; margin: 0; }
          .btn { display: inline-block; padding: 14px 24px; background-color: #005587; color: white; text-decoration: none; border-radius: 6px; font-weight: bold; margin-top: 15px; cursor: pointer; border: none; font-size: 15px; width: 100%; transition: 0.2s; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
          .btn:hover { background-color: #00b5e2; transform: translateY(-2px); }
        </style>
      </head>
      <body>
        <h3 style="color: #00205B; margin-top: 0;">✅ 파일 준비 완료!</h3>
        <p style="color: #444; font-size: 14px; word-break: keep-all;">오프라인 작동을 위한 모든 데이터가 내장되었습니다.<br>아래 버튼을 눌러 HTML 파일을 저장하세요.</p>
        <button class="btn" onclick="downloadFile()">📥 내 컴퓨터로 다운로드</button>
        <script>
          function downloadFile() {
            const b64Data = "${encodedHtml}";
            const byteCharacters = atob(b64Data);
            const byteNumbers = new Array(byteCharacters.length);
            for (let i = 0; i < byteCharacters.length; i++) {
              byteNumbers[i] = byteCharacters.charCodeAt(i);
            }
            const byteArray = new Uint8Array(byteNumbers);
            const blob = new Blob([byteArray], {type: 'text/html;charset=utf-8'});
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = "Airbus_Worksheet_Offline.html";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            setTimeout(function() { google.script.host.close(); }, 1500);
          }
        </script>
      </body>
      </html>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(dialogHtml).setWidth(350).setHeight(220);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '✈️ 오프라인 파일 다운로드');

  } catch (error) {
    SpreadsheetApp.getUi().alert("오류 발생: " + error.message);
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ 정비 포털 관리')
      .addItem('📥 오프라인 파일 바로 다운로드', 'downloadOfflinePortal')
      .addToUi();
}