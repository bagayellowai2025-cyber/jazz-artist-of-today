// 這是 Web App 的專屬入口大門
function doGet(e) {
  // 告訴系統去讀取叫做 'Index' 的 HTML 檔案，並把它變成網頁顯示出來
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Jazz Artist of Today') // 設定瀏覽器分頁標題
      .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // 確保手機版畫面比例正確
}

// 下面保留你原本的 getTodayArtist() 程式碼...

function getTodayArtist() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Artists");
    const data = sheet.getDataRange().getValues();
    
    // 取得今天的日期並格式化為 YYYY-MM-DD
    const today = new Date();
    const timeZone = Session.getScriptTimeZone();
    const todayString = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");
    
    let targetRow = null;
    
    // 尋找符合今天日期的資料
    for (let i = 1; i < data.length; i++) {
      let rowDate = data[i][1];
      let rowDateString = "";
      if (rowDate instanceof Date) {
         rowDateString = Utilities.formatDate(rowDate, timeZone, "yyyy-MM-dd");
      } else {
         rowDateString = String(rowDate); 
      }
      
      if (rowDateString === todayString) {
        targetRow = data[i];
        break;
      }
    }
    
    // 防呆機制：如果找不到今天的人物，預設抓取第一筆
    if (!targetRow && data.length > 1) {
      targetRow = data[1]; 
    }
    
    if (!targetRow) {
      return { error: "目前資料庫中沒有人物資料" };
    }
    
    // 【關鍵修復點】：將所有欄位強制加上 String() 轉換為純文字。
    // 這能防止 Google Sheets 隱藏的 Date 或 Object 格式導致打包崩潰。
    const artistData = {
      artistId: String(targetRow[0] || ""),
      displayDate: String(targetRow[1] || ""), 
      fullName: String(targetRow[2] || ""),
      nickname: String(targetRow[3] || ""),
      lifespan: String(targetRow[4] || ""),
      introduction: String(targetRow[5] || ""),
      photoUrl: String(targetRow[6] || ""),
      mediaLink: String(targetRow[7] || "")
    };
    
    return artistData;
    
  } catch (error) {
    // 如果後端真的壞掉了，我們把真正的錯誤原因變成字串傳給前端，而不是傳 null
    return { error: "後端抓取資料失敗：" + error.toString() };
  }
}

// 核心零件：專門負責向 Gemini API 發送問題並接收回答
function callGeminiAPI(promptText) {
  // 1. 從保險箱拿出 API Key
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return "找不到 API Key，請檢查專案設定。";

  // 2. 設定 API 端點 (使用 gemini-1.5-flash 模型)
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=" + apiKey;

  // 3. 準備要傳給 AI 的資料 (依照 Google API 規定的 JSON 格式)
  const payload = {
    "contents": [{
      "parts": [{"text": promptText}]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true // 讓我們能看到詳細的錯誤訊息
  };

  // 4. 發送請求並解析結果
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    // 檢查 API 是否回傳錯誤 (例如額度用盡或 Key 無效)
    if (json.error) {
      Logger.log("API 錯誤：" + json.error.message);
      return "抱歉，歷史學家目前需要休息一下 (API Error)。";
    }

    // 成功的話，從複雜的 JSON 結構中剝取出純文字回答
    const answer = json.candidates[0].content.parts[0].text;
    return answer;

  } catch (error) {
    Logger.log("連線錯誤：" + error.toString());
    return "連線發生錯誤，請稍後再試。";
  }
}

// 測試用零件：讓我們不用開網頁，直接在 GAS 裡測試 AI 是否活著
function testGemini() {
  const testQuestion = "請用一句話形容 Lindy Hop 這項舞蹈。";
  Logger.log("正在詢問 AI：" + testQuestion);
  
  const response = callGeminiAPI(testQuestion);
  Logger.log("AI 的回答是：" + response);
}

// ====================================================
// 【模組二】任務 1：抓取指定人物讚數最高的前 3 名問答
// ====================================================
function getTopQnA(artistId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QnA_History");
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 

    let filteredList = [];

    for (let i = 1; i < data.length; i++) {
      // 【防呆 1】：強制轉字串並用 trim() 去除頭尾可能隱藏的空白字元
      let rowArtistId = String(data[i][1]).trim();
      let targetArtistId = String(artistId).trim();
      
      if (rowArtistId === targetArtistId) {
        filteredList.push({
          qaId: String(data[i][0]).trim(),
          artistId: rowArtistId,
          userId: String(data[i][2]),
          question: String(data[i][3]),
          aiAnswer: String(data[i][4]),
          likesCount: Number(data[i][5]) || 0, 
          // 【防呆 2】：安全轉換時間格式，如果沒填寫時間就給 0
          timestamp: data[i][6] ? new Date(data[i][6]).getTime() : 0 
        });
      }
    }

    // 優先比讚數（大到小），如果讚數一樣，比時間（新到舊）
    filteredList.sort((a, b) => {
      if (b.likesCount !== a.likesCount) {
        return b.likesCount - a.likesCount;
      }
      return b.timestamp - a.timestamp;
    });

    return filteredList.slice(0, 3);

  } catch (error) {
    // 【防呆 3】：如果真的壞掉，前端會印出錯誤原因卡片，而不是默默變空白
    return [{ 
      qaId: "error", 
      question: "系統讀取通知", 
      aiAnswer: "無法讀取熱門問答，原因：" + error.toString(), 
      likesCount: 0 
    }];
  }
}

// ====================================================
// 【模組二】任務 2：更新指定 QA_ID 的按讚數 (+1)
// ====================================================
function updateLikeCount(qaId, action) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QnA_History");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(qaId).trim()) {
        const currentLikes = Number(data[i][5]) || 0;
        let newLikes = currentLikes;
        
        // 判斷前端傳來的指令是按讚還是收回讚
        if (action === "unlike") {
          newLikes = Math.max(0, currentLikes - 1); // 確保就算出錯也不會扣到變負數
        } else {
          newLikes = currentLikes + 1; // 預設為 +1
        }
        
        sheet.getRange(i + 1, 6).setValue(newLikes); 
        return { success: true, newLikes: newLikes };
      }
    }
    return { success: false, error: "找不到該筆問答紀錄" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ====================================================
// 【加強調整】修改 askHistorian，使其回傳物件（包含新生成的 QA_ID）
// ====================================================
function askHistorian(artistId, userQuestion) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Artists");
    const data = sheet.getDataRange().getValues();
    
    let artistName = "";
    let artistBio = "";
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(artistId)) {
        artistName = String(data[i][2]);
        artistBio = String(data[i][5]);
        break;
      }
    }
    
    const systemPrompt = `
      你是一位熱愛 Lindy Hop 且精通 1920-1950 年代爵士樂文化的歷史學家。
      現在有一位舞者正在詢問關於「${artistName}」的問題。
      
      【人物背景資料參考】：
      ${artistBio}
      
      【你的任務】：
      請以繁體中文回答舞者的問題。可以分享一些這位人物的fun fact，或是他對於爵士樂發展的影響。
      請針對問題核心回答，避免長篇大論，字數盡量控制在 150 字以內。
      
      舞者的問題是：${userQuestion}
    `;
    
    const aiAnswer = callGeminiAPI(systemPrompt);
    
    const qnaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QnA_History");
    const qaId = "QA_" + new Date().getTime(); 
    const userId = "Anonymous"; 
    const likesCount = 0; 
    const timestamp = new Date(); 
    
    qnaSheet.appendRow([qaId, artistId, userId, userQuestion, aiAnswer, likesCount, timestamp]);
    
    // 【修改重點】：不要只回傳字串，改回傳物件，讓前端拿到 QA_ID 才能對新卡片按讚！
    return {
      qaId: qaId,
      answer: aiAnswer
    };
    
  } catch (error) {
    return { error: "歷史學家正在查閱文獻，請稍後再試。(" + error.toString() + ")" };
  }
}