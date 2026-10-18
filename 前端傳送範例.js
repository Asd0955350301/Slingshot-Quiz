// 將 Apps Script 部署完成後取得的 /exec 網址貼在這裡。
const ANSWER_API_URL = 'https://script.google.com/macros/s/AKfycbyLBkfFD2A4axX9-EHKY4OEhu38F4X43f_1vev_RFK-7owf3e-O_WzRUZ2gamTJmG6vHg/exec';

/**
 * 每次石頭擊中任一答案後呼叫此函式。
 * 使用 URLSearchParams 可避免 application/json 產生跨網域預檢請求。
 */
async function saveAnswerToGoogleSheets(answer) {
  const form = new URLSearchParams({
    sessionId: answer.sessionId,
    playerName: answer.playerName,
    questionNumber: String(answer.questionNumber),
    questionText: answer.questionText,
    selectedOption: answer.selectedOption,
    selectedText: answer.selectedText,
    correctAnswer: answer.correctAnswer,
    isCorrect: String(answer.isCorrect),
    scoreEarned: String(answer.scoreEarned),
    heartsRemaining: String(answer.heartsRemaining),
    gameResult: answer.gameResult, // 進行中、成功或失敗
    deviceType: getDeviceType()
  });

  // Apps Script Web App 的跨網域回應通常為 opaque；成功送出不代表
  // 瀏覽器能讀取回傳 JSON，因此前端應以不阻塞遊戲的方式儲存。
  await fetch(ANSWER_API_URL, {
    method: 'POST',
    mode: 'no-cors',
    body: form,
    keepalive: true
  });
}

function getDeviceType() {
  const width = window.innerWidth;
  if (width < 600) return '手機';
  if (width < 1024) return '平板';
  return '電腦';
}

// 呼叫範例：
// saveAnswerToGoogleSheets({
//   sessionId: crypto.randomUUID(),
//   playerName: '小安',
//   questionNumber: 1,
//   questionText: '哪一項屬於移動污染源？',
//   selectedOption: 'C',
//   selectedText: '汽車與機車',
//   correctAnswer: 'C',
//   isCorrect: true,
//   scoreEarned: 10,
//   heartsRemaining: 5,
//   gameResult: '進行中'
// }).catch(console.error);
