// ⚠️ 記得把這裡換回你專屬的 Google API 網址 ⚠️
const API_URL = "https://script.google.com/macros/s/AKfycbwfysuXE77bR2x8YuK2bRHL9aPKNyxj2zug2MXBO6Km0mZbamXbAzaF_oqZwEOCjzEdvQ/exec";

var accuracyChartInstance = null;
var profChartInstance = null;

var allVocabData = []; 
var vocabData = [];
var globalProgressData = []; 
var currentBrowseIndex = 0;
var currentQuizItem = null;
var synth = window.speechSynthesis;
var currentSheetName = ''; 
var currentUser = '';
var userRole = 'free'; 
var isCurrentWordAnswered = false; 

// --------------------------------------------------------
// 💡 萬能 AI 文字助理引擎
// --------------------------------------------------------
function getAITextAssistant(taskType) {
    if (userRole !== 'premium' && userRole !== 'admin') {
        showToast('🔒 此為 Premium 專屬功能，請升級解鎖！', 'warn');
        return;
    }

    let word = currentQuizItem.word;
    let lang = currentSheetName;
    
    if (taskType === 'sentence') {
        let btn = document.getElementById('aiSentenceBtn');
        let originalText = btn.innerText;
        btn.innerText = "⏳ 思考中...";
        btn.disabled = true;

        apiCall('getAIAssistant', { taskType: 'sentence', word: word, lang: lang, userRole: userRole }, function(res) {
            btn.innerText = originalText;
            btn.disabled = false;
            if(res.success) {
                let reply = res.reply;
                let exMatch = reply.match(/【例句】(.*?)(?=【翻譯】|$)/s);
                let transMatch = reply.match(/【翻譯】(.*)/s);
                
                if(exMatch) document.getElementById('browseExample').innerText = "✨ " + exMatch[1].trim();
                if(transMatch) {
                    document.getElementById('browseExampleTrans').innerText = transMatch[1].trim();
                    document.getElementById('browseExampleTrans').classList.remove('revealed');
                }
                
                currentQuizItem.aiSentence = reply;
                updateGlobalProgressLocallyText(word, 'sentence', reply);
                apiCall('saveAIText', {sheetName: currentSheetName, word: word, username: currentUser, textType: 'sentence', textData: reply});
                
                showToast("✨ AI 已為你生成專屬例句！", "success");
            } else { showToast(res.message, "danger"); }
        }, function(err) {
            btn.innerText = originalText; btn.disabled = false; showToast("連線失敗", "danger");
        });
    } 
    else if (taskType === 'mnemonic' || taskType === 'spelling_mnemonic') {
        let boxId = taskType === 'mnemonic' ? 'aiMnemonicBox' : 'spellingMnemonicBox';
        let box = document.getElementById(boxId);
        box.style.display = 'block';
        box.innerHTML = '<div class="spinner" style="display:inline-block; width:15px; height:15px; border-top-color:var(--premium);"></div> AI 正在發功找記憶法...';
        
        if (taskType === 'spelling_mnemonic') document.getElementById('spellingMnemonicBtn').style.display = 'none';
        if (taskType === 'mnemonic') document.getElementById('aiMnemonicBtn').style.display = 'none';

        apiCall('getAIAssistant', { taskType: 'mnemonic', word: word, lang: lang, userRole: userRole }, function(res) {
            if(res.success) {
                box.innerHTML = `💡 <b>專屬 AI 記憶法：</b><br>${marked.parse(res.reply)}`;
                currentQuizItem.aiMnemonic = res.reply;
                updateGlobalProgressLocallyText(word, 'mnemonic', res.reply);
                apiCall('saveAIText', {sheetName: currentSheetName, word: word, username: currentUser, textType: 'mnemonic', textData: res.reply});
            } else box.innerHTML = `❌ ${res.message}`;
        }, function(err) { box.innerHTML = `❌ 連線失敗`; });
    }
}

function updateGlobalProgressLocallyText(word, type, text) {
    let p = globalProgressData.find(item => item.word === word && item.lang === currentSheetName);
    if (!p) {
        p = { lang: currentSheetName, word: word, cCorrect: 0, cIncorrect: 0, sCorrect: 0, sIncorrect: 0, prof: '', speakSimilarity: '', speakNextReview: '', aiWordFeedback: '', aiExampleFeedback: '', aiWordAudioUrl: '', aiExampleAudioUrl: '', aiSentence: '', aiMnemonic: '' };
        globalProgressData.push(p);
    }
    if(type === 'sentence') p.aiSentence = text;
    if(type === 'mnemonic') p.aiMnemonic = text;
}

// --------------------------------------------------------
// 💡 終極 AudioContext 音頻解碼引擎
// --------------------------------------------------------
let audioCtx = null;
let currentAudioSource = null;
let currentPlaybackAudio = null;
let currentPlaybackBtn = null;
let originalBtnText = "";

function playAudio(btn, type, data) {
    if (!audioCtx) {
        audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    }
    if (audioCtx.state === 'suspended') {
        audioCtx.resume();
    }

    if (currentPlaybackAudio) {
        currentPlaybackAudio.pause();
        currentPlaybackAudio = null;
    }
    if (currentAudioSource) {
        try { currentAudioSource.stop(); } catch(e){}
        currentAudioSource = null;
    }

    if (currentPlaybackBtn) {
        currentPlaybackBtn.innerText = originalBtnText;
        currentPlaybackBtn.disabled = false;
    }

    if (currentPlaybackBtn === btn) {
        currentPlaybackBtn = null;
        return;
    }

    currentPlaybackBtn = btn;
    originalBtnText = btn.innerText;

    if (type === 'local') {
        btn.innerText = "🔊 播放中...";
        currentPlaybackAudio = new Audio(data);
        currentPlaybackAudio.onended = () => { 
            if(currentPlaybackBtn === btn) { btn.innerText = originalBtnText; currentPlaybackBtn = null; }
        };
        currentPlaybackAudio.play().catch(e => {
            showToast("播放失敗，請檢查權限", "danger");
            btn.innerText = originalBtnText;
        });
    } else if (type === 'cloud') {
        btn.innerText = "⏳ 讀取中...";
        btn.disabled = true;
        
        apiCall('getAudio', { audioUrl: data }, function(res) {
            if (res.success) {
                if (currentPlaybackBtn !== btn) return; 
                
                btn.innerText = "⏳ 轉碼中...";
                try {
                    const binaryString = atob(res.base64);
                    const len = binaryString.length;
                    const bytes = new Uint8Array(len);
                    for (let i = 0; i < len; i++) {
                        bytes[i] = binaryString.charCodeAt(i);
                    }
                    
                    audioCtx.decodeAudioData(bytes.buffer, function(buffer) {
                        if (currentPlaybackBtn !== btn) return; 

                        currentAudioSource = audioCtx.createBufferSource();
                        currentAudioSource.buffer = buffer;
                        currentAudioSource.connect(audioCtx.destination);
                        
                        currentAudioSource.onended = () => {
                            if(currentPlaybackBtn === btn) { 
                                btn.innerText = originalBtnText; 
                                currentPlaybackBtn = null; 
                                btn.disabled = false;
                            }
                        };
                        
                        btn.innerText = "🔊 播放中...";
                        btn.disabled = false;
                        currentAudioSource.start(0);
                        
                    }, function(err) {
                        showToast("音軌解碼失敗", "danger");
                        btn.innerText = originalBtnText;
                        btn.disabled = false;
                    });
                } catch(err) {
                    showToast("音檔解析錯誤", "danger");
                    btn.innerText = originalBtnText;
                    btn.disabled = false;
                }
            } else {
                btn.innerText = "❌ 讀取失敗"; btn.disabled = false;
                setTimeout(() => { if(currentPlaybackBtn === btn) btn.innerText = originalBtnText; }, 2000);
            }
        }, function(err) {
            btn.innerText = "❌ 連線失敗"; btn.disabled = false;
            setTimeout(() => { if(currentPlaybackBtn === btn) btn.innerText = originalBtnText; }, 2000);
        });
    }
}

function showToast(message, type = 'info') {
  let container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerText = message;
  container.appendChild(toast);
  setTimeout(() => { toast.remove(); }, 3000);
}

function levenshteinDistance(a, b) {
    if (a.length === 0) return b.length;
    if (b.length === 0) return a.length;
    var matrix = [];
    for (var i = 0; i <= b.length; i++) { matrix[i] = [i]; }
    for (var j = 0; j <= a.length; j++) { matrix[0][j] = j; }
    for (var i = 1; i <= b.length; i++) {
        for (var j = 1; j <= a.length; j++) {
            if (b.charAt(i - 1) === a.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(matrix[i - 1][j - 1] + 1, Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1));
            }
        }
    }
    return matrix[b.length][a.length];
}

function getSimilarity(a, b) {
    var dist = levenshteinDistance(a, b);
    var maxLen = Math.max(a.length, b.length);
    if (maxLen === 0) return 100;
    return (1 - dist / maxLen) * 100;
}

function apiCall(action, payload, onSuccess, onError) {
  if (API_URL.includes("請把你的")) {
     showToast("請先在原始碼 script.js 第 2 行替換你的 API_URL！", "danger");
     return;
  }
  payload.action = action;
  fetch(API_URL, {
    redirect: "follow",
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload)
  })
  .then(response => response.json())
  .then(data => { if(onSuccess) onSuccess(data); })
  .catch(error => { console.error('API 錯誤:', error); if(onError) onError(error); });
}

function toggleAuthMode(mode) {
  if(mode === 'login') {
    document.getElementById('loginBox').style.display = 'flex';
    document.getElementById('registerBox').style.display = 'none';
  } else {
    document.getElementById('loginBox').style.display = 'none';
    document.getElementById('registerBox').style.display = 'flex';
  }
}

function doRegister() {
  var u = document.getElementById('regUser').value.trim();
  var p = document.getElementById('regPass').value.trim();
  var err = document.getElementById('regError');
  if(!u || !p) { err.innerText = "帳號與密碼不能為空！"; err.style.display='block'; return; }
  
  document.getElementById('regLoader').style.display = 'block';
  err.style.display = 'none';
  
  apiCall('registerUser', {username: u, password: p}, function(res) {
    document.getElementById('regLoader').style.display = 'none';
    if(res.success) {
       showToast(res.message, 'success');
       toggleAuthMode('login');
       document.getElementById('loginUser').value = u;
    } else {
       err.innerText = res.message;
       err.style.display = 'block';
    }
  });
}

function doLogin() {
  var u = document.getElementById('loginUser').value.trim();
  var p = document.getElementById('loginPass').value.trim();
  var err = document.getElementById('loginError');
  if(!u || !p) { err.innerText = "帳號與密碼不能為空！"; err.style.display='block'; return; }
  
  document.getElementById('loginLoader').style.display = 'block';
  err.style.display = 'none';
  
  apiCall('loginUser', {username: u, password: p}, function(res) {
    document.getElementById('loginLoader').style.display = 'none';
    if(res.success) {
       currentUser = u;
       userRole = res.role; 
       showToast(`歡迎回來，${u}！`, 'success');
       
       document.getElementById('userBadge').innerText = "👤 " + u;
       document.getElementById('loginScreen').style.display = 'none';
       document.getElementById('welcomeScreen').style.display = 'flex'; 
    } else {
       err.innerText = res.message;
       err.style.display = 'block';
    }
  });
}

function doLogout() {
  currentUser = '';
  userRole = 'free';
  allVocabData = [];
  vocabData = [];
  globalProgressData = [];
  document.getElementById('loginPass').value = '';
  
  document.getElementById('startAppBtn').style.display = 'block';
  document.getElementById('welcomeLoader').style.display = 'none';
  document.getElementById('welcomeError').style.display = 'none';

  document.getElementById('appContainer').style.display = 'none';
  document.getElementById('welcomeScreen').style.display = 'none';
  document.getElementById('loginScreen').style.display = 'flex';
  showToast("已登出系統", "info");
}

function toggleFullScreen() {
  var doc = window.document;
  var docEl = doc.documentElement;
  var requestFullScreen = docEl.requestFullscreen || docEl.mozRequestFullScreen || docEl.webkitRequestFullScreen || docEl.msRequestFullscreen;
  var cancelFullScreen = doc.exitFullscreen || doc.mozCancelFullScreen || doc.webkitExitFullscreen || doc.msExitFullscreen;
  if(!doc.fullscreenElement && !doc.mozFullScreenElement && !doc.webkitFullscreenElement && !doc.msFullscreenElement) {
    if(requestFullScreen) requestFullScreen.call(docEl);
  } else {
    if(cancelFullScreen) cancelFullScreen.call(doc);
  }
}

let aiMediaRecorder;
let aiAudioChunks = [];
let isRecordingAI = false;
let currentAITarget = ''; 

async function toggleAIRecording(targetType) {
    if (!isRecordingAI) {
        try {
            let stream = await navigator.mediaDevices.getUserMedia({ audio: true });
            aiMediaRecorder = new MediaRecorder(stream);
            aiAudioChunks = [];
            currentAITarget = targetType;
            
            aiMediaRecorder.ondataavailable = e => aiAudioChunks.push(e.data);
            aiMediaRecorder.onstop = processAIAudio;
            aiMediaRecorder.start();
            
            isRecordingAI = true;
            let btnId = targetType === 'word' ? 'aiMicWordBtn' : 'aiMicExampleBtn';
            let fbId = targetType === 'word' ? 'aiWordFeedback' : 'aiExampleFeedback';
            
            document.getElementById(btnId).classList.add('recording');
            document.getElementById(btnId).innerText = "⏹️ 停止錄音";
            
            let fbBox = document.getElementById(fbId);
            fbBox.style.display = 'block';
            fbBox.innerHTML = '<span style="color:var(--danger); animation: pulse 1s infinite;">🎙️ 錄音中... 唸完請點擊停止</span>';
        } catch (err) {
            showToast("無法存取麥克風，請確認權限！", "danger");
        }
    } else {
        aiMediaRecorder.stop();
        isRecordingAI = false;
        
        let btnId = currentAITarget === 'word' ? 'aiMicWordBtn' : 'aiMicExampleBtn';
        let fbId = currentAITarget === 'word' ? 'aiWordFeedback' : 'aiExampleFeedback';
        let btnText = currentAITarget === 'word' ? '🎙️ 測驗單字' : '🎙️ 測驗例句';
        
        document.getElementById(btnId).classList.remove('recording');
        document.getElementById(btnId).innerText = btnText;
        
        document.getElementById(fbId).innerHTML = '<div class="spinner" style="display:inline-block; border-top-color:var(--premium); width:20px; height:20px;"></div> AI 上傳分析中...';
    }
}

function processAIAudio() {
    let audioBlob = new Blob(aiAudioChunks, { type: aiMediaRecorder.mimeType });
    let localBlobUrl = URL.createObjectURL(audioBlob); 
    
    let reader = new FileReader();
    reader.readAsDataURL(audioBlob);
    let activeTarget = currentAITarget; 
    
    reader.onloadend = function() {
        let base64data = reader.result.split(',')[1];
        
        apiCall('evaluateAudio', {
            base64Audio: base64data,
            mimeType: audioBlob.type,
            word: currentQuizItem.word,
            example: document.getElementById('aiExample').innerText,
            lang: currentSheetName,
            userRole: userRole,
            username: currentUser, 
            targetType: activeTarget
        }, function(res) {
            let fbId = activeTarget === 'word' ? 'aiWordFeedback' : 'aiExampleFeedback';
            if (res.success) {
                let parsedHtml = marked.parse(res.reply);
                let myAudioBtn = `<button class="btn-skip" style="font-size:0.85rem; padding:0.2rem 0.5rem; margin-top:0.5rem;" onclick="playAudio(this, 'local', '${localBlobUrl}')">🎧 播放我的錄音</button>`;
                document.getElementById(fbId).innerHTML = parsedHtml + myAudioBtn;
                
                if(activeTarget === 'word') {
                    currentQuizItem.aiWordFeedback = res.reply;
                    currentQuizItem.aiWordAudioUrl = res.audioUrl;
                } else {
                    currentQuizItem.aiExampleFeedback = res.reply;
                    currentQuizItem.aiExampleAudioUrl = res.audioUrl;
                }
                
                let p = globalProgressData.find(v => v.word === currentQuizItem.word && v.lang === currentSheetName);
                if(p) {
                    if(activeTarget === 'word') { p.aiWordFeedback = res.reply; p.aiWordAudioUrl = res.audioUrl; }
                    else { p.aiExampleFeedback = res.reply; p.aiExampleAudioUrl = res.audioUrl; }
                }
                
                apiCall('saveAIFeedback', { sheetName: currentSheetName, word: currentQuizItem.word, username: currentUser, targetType: activeTarget, feedback: res.reply, audioUrl: res.audioUrl });
                
            } else {
                document.getElementById(fbId).innerHTML = `<span style="color:var(--danger);">❌ ${res.message}</span>`;
            }
        }, function(err) {
            let fbId = activeTarget === 'word' ? 'aiWordFeedback' : 'aiExampleFeedback';
            document.getElementById(fbId).innerHTML = `<span style="color:var(--danger);">❌ 連線失敗，請重試</span>`;
        });
    }
}

function loadNextAITutor() {
  window.scrollTo({ top: 0, behavior: 'smooth' });
  if(vocabData.length === 0) return;
  currentQuizItem = vocabData[currentBrowseIndex]; 
  updateProgressUI(); 

  let authBox = document.getElementById('aiAuthContainer');
  let mainBox = document.getElementById('aiMainContainer');
  
  if (userRole === 'premium' || userRole === 'admin') {
      authBox.style.display = 'none';
      mainBox.style.display = 'block';
      
      document.getElementById('aiWord').innerText = formatWordDisplay(currentQuizItem);
      document.getElementById('aiTranslation').innerText = currentQuizItem.translation;
      
      let displayExample = currentQuizItem.example ? currentQuizItem.example.replace(/\[|\]/g, '') : '無例句';
      document.getElementById('aiExample').innerText = displayExample;
      document.getElementById('aiExampleTrans').innerText = currentQuizItem.exampleTranslation || '';
      
      let wFbBox = document.getElementById('aiWordFeedback');
      if(currentQuizItem.aiWordFeedback) {
          wFbBox.style.display = 'block';
          let myAudioBtn = currentQuizItem.aiWordAudioUrl ? `<button class="btn-skip" style="font-size:0.85rem; padding:0.2rem 0.5rem; margin-top:0.5rem;" onclick="playAudio(this, 'cloud', '${currentQuizItem.aiWordAudioUrl}')">🎧 播放我的錄音</button>` : '';
          wFbBox.innerHTML = marked.parse(currentQuizItem.aiWordFeedback) + myAudioBtn;
      } else { wFbBox.style.display = 'none'; wFbBox.innerHTML = ''; }
      
      let eFbBox = document.getElementById('aiExampleFeedback');
      if(currentQuizItem.aiExampleFeedback) {
          eFbBox.style.display = 'block';
          let myAudioBtn = currentQuizItem.aiExampleAudioUrl ? `<button class="btn-skip" style="font-size:0.85rem; padding:0.2rem 0.5rem; margin-top:0.5rem;" onclick="playAudio(this, 'cloud', '${currentQuizItem.aiExampleAudioUrl}')">🎧 播放我的錄音</button>` : '';
          eFbBox.innerHTML = marked.parse(currentQuizItem.aiExampleFeedback) + myAudioBtn;
      } else { eFbBox.style.display = 'none'; eFbBox.innerHTML = ''; }

  } else {
      authBox.style.display = 'block';
      mainBox.style.display = 'none';
      authBox.innerHTML = `
         <div class="ai-lock-screen">
            <h1 style="font-size: 4rem; margin: 0;">🔒</h1>
            <h2 style="color: var(--text);">專屬 AI 發音特診室</h2>
            <p style="margin-bottom: 2rem;">升級解鎖 Gemini 語音點評，自動記錄你的發音弱點！</p>
            <button class="premium-btn" onclick="showToast('請聯絡管理員升級為 Premium 帳號！🚀', 'warn')">升級 Premium 解鎖</button>
         </div>
      `;
  }
}

function applySpeakingProficiency(item, sim) {
  let nextDate = new Date();
  if (sim < 60) nextDate.setMinutes(nextDate.getMinutes() + 10);
  else if (sim < 85) nextDate.setDate(nextDate.getDate() + 1);
  else nextDate.setDate(nextDate.getDate() + 5);
  item.speakNextReview = nextDate.toISOString();
  return item.speakNextReview;
}

function skipSpeaking() {
  if(!currentQuizItem) return;
  var farFuture = "2099-12-31T00:00:00.000Z";
  currentQuizItem.speakNextReview = farFuture;
  let p = globalProgressData.find(item => item.word === currentQuizItem.word && item.lang === currentSheetName);
  if(p) p.speakNextReview = farFuture;

  updateBackend(currentQuizItem.word, true, null, 'speaking', null, farFuture, null);
  showToast(`已將「${currentQuizItem.word}」永久排除口說！\n(可於紀錄頁面恢復)`, "warn");
  nextQuizWord('speaking');
}

function restoreBlacklistWord(word) {
    let item = allVocabData.find(v => v.word === word);
    if(item) item.speakNextReview = '';
    let p = globalProgressData.find(v => v.word === word && v.lang === currentSheetName);
    if(p) p.speakNextReview = '';

    renderBlacklist();
    apiCall('restoreSpeaking', {sheetName: currentSheetName, word: word, username: currentUser});
    showToast(`已恢復「${word}」的口說測驗！`, "success");
}

function renderBlacklist() {
    let container = document.getElementById('blacklistContainer');
    container.innerHTML = '';
    let blacklistedWords = allVocabData.filter(v => v.speakNextReview && v.speakNextReview.includes('2099'));
    
    if(blacklistedWords.length === 0) {
        container.innerHTML = '<p style="text-align:center; color:var(--text-light); margin:0; font-size: 0.9rem;">目前沒有被略過的單字</p>';
        return;
    }

    blacklistedWords.forEach(v => {
        let div = document.createElement('div');
        div.style.display = 'flex';
        div.style.justifyContent = 'space-between';
        div.style.alignItems = 'center';
        div.style.borderBottom = '1px solid #f3f4f6';
        div.style.padding = '0.5rem 0';
        div.innerHTML = `<span style="font-size: 0.95rem;"><b>${v.word}</b> (${v.translation})</span>
                         <button class="btn-skip" style="font-size:0.8rem; padding:0.3rem 0.6rem;" onclick="restoreBlacklistWord('${v.word}')">恢復</button>`;
        container.appendChild(div);
    });
}

function renderAIHistory() {
    let container = document.getElementById('aiHistoryContainer');
    container.innerHTML = '';
    let aiWords = allVocabData.filter(v => v.aiWordFeedback || v.aiExampleFeedback);
    
    if(aiWords.length === 0) {
        container.innerHTML = '<p style="text-align:center; color:var(--text-light); margin:0; font-size: 0.9rem;">目前還沒有 AI 診斷紀錄喔！快去 AI 特訓班試試看吧！</p>';
        return;
    }

    aiWords.forEach((v, index) => {
        let div = document.createElement('div');
        div.className = 'ai-hist-item';
        
        let header = document.createElement('div');
        header.className = 'ai-hist-header';
        header.onclick = function() {
            let content = document.getElementById('ai-hist-content-' + index);
            content.style.display = content.style.display === 'none' ? 'block' : 'none';
        };
        header.innerHTML = `<span style="font-size: 1rem; color:var(--primary); font-weight:bold;">${v.word} <span style="font-size:0.85rem; color:var(--text-light); font-weight:normal;">(${v.translation})</span></span>
                         <span style="font-size:0.8rem; background:var(--premium); color:white; padding:0.2rem 0.5rem; border-radius:99px;">查看報告 ▼</span>`;
        
        let content = document.createElement('div');
        content.id = 'ai-hist-content-' + index;
        content.className = 'ai-hist-content';
        
        let html = '';
        if(v.aiWordFeedback) {
            html += `<div style="margin-bottom:1rem;">
                        <div style="margin-bottom: 0.5rem;"><strong>單字診斷：</strong>
                            <button class="speaker-btn" style="font-size:1rem;" onclick="speakText('${v.word}')">🔊 標準音</button>
                            ${v.aiWordAudioUrl ? `<button class="btn-skip" style="font-size:0.8rem; padding:0.2rem 0.5rem; margin-left:0.5rem;" onclick="playAudio(this, 'cloud', '${v.aiWordAudioUrl}')">🎧 我的錄音</button>` : ''}
                        </div>
                        ${marked.parse(v.aiWordFeedback)}
                     </div>`;
        }
        if(v.aiExampleFeedback) {
            let safeEx = (v.example ? v.example.replace(/\[|\]/g, '') : '').replace(/"/g, '&quot;');
            html += `<div>
                        <div style="margin-bottom: 0.5rem;"><strong>例句診斷：</strong>
                            <button class="speaker-btn" style="font-size:1rem;" onclick="speakText(this.getAttribute('data-text'))" data-text="${safeEx}">🔊 標準音</button>
                            ${v.aiExampleAudioUrl ? `<button class="btn-skip" style="font-size:0.8rem; padding:0.2rem 0.5rem; margin-left:0.5rem;" onclick="playAudio(this, 'cloud', '${v.aiExampleAudioUrl}')">🎧 我的錄音</button>` : ''}
                        </div>
                        ${marked.parse(v.aiExampleFeedback)}
                     </div>`;
        }
        
        content.innerHTML = html;
        div.appendChild(header);
        div.appendChild(content);
        container.appendChild(div);
    });
}

const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
var recognition = null;
var isRecording = false;
if (SpeechRecognition) {
  recognition = new SpeechRecognition();
  recognition.continuous = false;
  recognition.interimResults = false; 
  recognition.onstart = function() {
    isRecording = true;
    document.getElementById('micBtn').classList.add('recording');
    document.getElementById('recordingStatus').style.display = 'block';
    document.getElementById('speakingFeedback').innerText = '';
    document.getElementById('recognizedText').innerText = '正在聆聽... (請說話)';
  };
  
  recognition.onresult = function(event) {
    isCurrentWordAnswered = true; 
    var transcript = event.results[0][0].transcript.trim().toLowerCase();
    var targetWord = currentQuizItem.word.trim().toLowerCase();
    
    var cleanTranscript = transcript.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()\s]/g,"");
    var cleanTarget = targetWord.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()\s]/g,"");
    var cleanKanji = currentQuizItem.kanji ? currentQuizItem.kanji.trim().toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()\s]/g,"") : "";

    var simWord = getSimilarity(cleanTranscript, cleanTarget);
    var simKanji = cleanKanji ? getSimilarity(cleanTranscript, cleanKanji) : 0;
    var bestSim = Math.round(Math.max(simWord, simKanji));

    var isCorrect = (cleanTranscript === cleanTarget || cleanTranscript.includes(cleanTarget) ||
        (cleanKanji !== "" && (cleanTranscript === cleanKanji || cleanTranscript.includes(cleanKanji))) ||
        bestSim >= 75);

    var feedback = document.getElementById('speakingFeedback');
    let speakNextRev = applySpeakingProficiency(currentQuizItem, bestSim);

    if (isCorrect) {
      if (bestSim >= 75 && bestSim < 100 && !cleanTranscript.includes(cleanTarget) && (!cleanKanji || !cleanTranscript.includes(cleanKanji))) {
         document.getElementById('recognizedText').innerText = "系統聽到: 「" + transcript + "」";
         feedback.innerText = '發音過關！👍 (相似度 ' + bestSim + '%)';
      } else {
         document.getElementById('recognizedText').innerText = "系統聽到: 「" + transcript + "」";
         feedback.innerText = '發音標準！💯';
         bestSim = 100; 
      }
      feedback.className = 'feedback success';
    } else {
      document.getElementById('recognizedText').innerText = "系統聽到: 「" + transcript + "」";
      feedback.innerText = '發音不太對喔 (相似度 ' + bestSim + '%)，請再按一次麥克風重試！🤔';
      feedback.className = 'feedback danger';
    }
    
    updateBackend(currentQuizItem.word, isCorrect, null, 'speaking', null, speakNextRev, bestSim);
    currentQuizItem.speakSimilarity = bestSim; 
    updateGlobalProgressLocally(currentQuizItem.word, 'speaking', isCorrect, null, bestSim);
    
    document.getElementById('speakingPrevSim').innerText = "💡 剛剛發音相似度：" + bestSim + "%";
    document.getElementById('speakingNextBtn').style.display = 'block';
    setTimeout(function() { document.getElementById('speakingNextBtn').scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }, 100);
  };
  
  recognition.onerror = function(event) {
     var feedback = document.getElementById('speakingFeedback');
     if(event.error === 'no-speech') {
         feedback.innerText = "沒清楚收音辨識，請再按一次麥克風。";
     } else {
         feedback.innerText = "麥克風錯誤 (" + event.error + ")，請重試！";
     }
     feedback.className = 'feedback danger';
     document.getElementById('recognizedText').innerText = '辨識失敗';
  };
  
  recognition.onend = function() {
    isRecording = false;
    document.getElementById('micBtn').classList.remove('recording');
    document.getElementById('recordingStatus').style.display = 'none';
    
    var recText = document.getElementById('recognizedText');
    if (recText.innerText === '正在聆聽... (請說話)') {
        recText.innerText = '';
        var feedback = document.getElementById('speakingFeedback');
        if(!feedback.innerText) {
            feedback.innerText = "沒清楚收音辨識，請再按一次麥克風。";
            feedback.className = 'feedback danger';
        }
    }
  };
}

function toggleRecording() {
  if (!SpeechRecognition) { showToast("你的瀏覽器不支援語音辨識功能！", "danger"); return; }
  if (isRecording) { recognition.stop(); } 
  else { recognition.lang = document.getElementById('langSelect').options[document.getElementById('langSelect').selectedIndex].getAttribute('data-tts'); recognition.start(); }
}

function startApp() {
  var select = document.getElementById('startLangSelect');
  currentSheetName = select.value;
  document.getElementById('langSelect').value = currentSheetName;
  document.getElementById('startAppBtn').style.display = 'none';
  document.getElementById('welcomeLoader').style.display = 'block';
  document.getElementById('welcomeError').style.display = 'none';
  fetchDataFromSheet('welcome');
}

window.onload = function() {
  var inputEl = document.getElementById('spellingInput');
  if(inputEl) {
    inputEl.addEventListener('keypress', function(e) {
      if (e.key === 'Enter') {
        var nextBtn = document.getElementById('spellingNextBtn');
        if(nextBtn.style.display === 'block') {
            nextQuizWord('spelling'); 
        }
        else checkSpelling();
      }
    });
    inputEl.addEventListener('input', function() { this.classList.remove('error-input'); });
  }
  document.getElementById('loginPass').addEventListener('keypress', function(e){ if(e.key === 'Enter') doLogin(); });
  document.getElementById('regPass').addEventListener('keypress', function(e){ if(e.key === 'Enter') doRegister(); });
};

document.addEventListener('keydown', function(e) {
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
  var browseActive = document.getElementById('browse').classList.contains('active');
  if (browseActive && document.getElementById('appContainer').style.display !== 'none') {
    if (e.key === ' ') {
      e.preventDefault(); 
      var transEl = document.getElementById('browseTranslation');
      var exTransEl = document.getElementById('browseExampleTrans');
      if (!transEl.classList.contains('revealed')) {
        transEl.classList.add('revealed');
        exTransEl.classList.add('revealed');
      } else {
        speakText(document.getElementById('browseWord').getAttribute('data-speak'));
      }
    } else if (e.key === 'ArrowRight') nextWord();
    else if (e.key === 'ArrowLeft') prevWord();
    else if (e.key === '1') setProficiency('完全忘記');
    else if (e.key === '2') setProficiency('模糊');
    else if (e.key === '3') setProficiency('熟練');
  }
});

function revealText(el) { el.classList.add('revealed'); }

function changeLanguage() {
  var select = document.getElementById('langSelect');
  currentSheetName = select.value;
  document.querySelectorAll('.section').forEach(function(el) { el.classList.remove('active'); });
  document.getElementById('mainLoader').style.display = 'block';
  document.getElementById('globalControls').style.display = 'none';
  document.getElementById('globalProgress').style.display = 'none';
  document.getElementById('mainLoader').innerText = "載入 " + currentSheetName + " 題庫中...";
  
  document.getElementById('sortSelect').value = 'sm2'; 
  if(document.getElementById('levelSelect')) document.getElementById('levelSelect').value = 'all';
  fetchDataFromSheet('main');
}

function handleError(errorMsg, source) {
    if(source === 'welcome') {
      document.getElementById('welcomeLoader').style.display = 'none';
      document.getElementById('startAppBtn').style.display = 'block';
      var errEl = document.getElementById('welcomeError');
      errEl.innerText = "載入失敗：" + errorMsg;
      errEl.style.display = 'block';
    } else {
      document.getElementById('mainLoader').innerText = "載入失敗：" + errorMsg;
    }
}

function fetchDataFromSheet(source) {
  apiCall('getVocabularyData', {sheetName: currentSheetName, username: currentUser}, function(res) {
    if(res.success) {
        allVocabData = res.data.vocabList;
        globalProgressData = res.data.globalProgress;
        initApp();
    } else {
        handleError(res.message, source);
    }
  }, function(error) {
     handleError(error, source);
  });
}

function populateLevelDropdown() {
  var levelRow = document.getElementById('levelFilterRow');
  levelRow.style.display = 'flex';
  var select = document.getElementById('levelSelect');
  var currentVal = select.value;
  
  select.innerHTML = '<option value="all">全部等級</option>';
  var uniqueLevels = [];
  for(var i=0; i<allVocabData.length; i++) {
    var lvl = allVocabData[i].level;
    if(lvl && uniqueLevels.indexOf(lvl) === -1) uniqueLevels.push(lvl);
  }
  
  if(uniqueLevels.length === 0) {
     levelRow.style.display = 'none';
     return;
  }
  
  uniqueLevels.sort();
  for(var j=0; j<uniqueLevels.length; j++) {
    var opt = document.createElement('option');
    opt.value = uniqueLevels[j]; opt.innerText = uniqueLevels[j];
    select.appendChild(opt);
  }
  if (uniqueLevels.indexOf(currentVal) !== -1) select.value = currentVal;
}

function initApp() {
  document.getElementById('welcomeScreen').style.display = 'none';
  document.getElementById('appContainer').style.display = 'flex';
  document.getElementById('mainLoader').style.display = 'none';
  
  if (allVocabData.length === 0) {
    document.getElementById('browse').innerHTML = "<p style='text-align:center; padding: 2rem;'>[" + currentSheetName + "] 題庫為空，請在試算表新增單字！</p>";
    document.getElementById('browse').style.display = 'block';
    return;
  }
  
  document.getElementById('globalControls').style.display = 'flex';
  
  populateLevelDropdown(); 
  applySort(); 
  switchTab('browse');
}

function renderDashboard() {
  let totalLearned = 0;
  let totalCorrect = 0, totalWrong = 0;
  let profCounts = { '熟練':0, '模糊':0, '完全忘記':0 };
  
  let langStats = {
      '韓文': { count: 0, correct: 0, wrong: 0, prof3: 0, prof2: 0, prof1: 0, speakSimSum: 0, speakSimCount: 0 },
      '英文': { count: 0, correct: 0, wrong: 0, prof3: 0, prof2: 0, prof1: 0, speakSimSum: 0, speakSimCount: 0 },
      '日文': { count: 0, correct: 0, wrong: 0, prof3: 0, prof2: 0, prof1: 0, speakSimSum: 0, speakSimCount: 0 },
      '泰文': { count: 0, correct: 0, wrong: 0, prof3: 0, prof2: 0, prof1: 0, speakSimSum: 0, speakSimCount: 0 }
  };

  try {
      globalProgressData.forEach(p => {
          let hasInteracted = p.cCorrect>0 || p.cIncorrect>0 || p.sCorrect>0 || p.sIncorrect>0 || p.prof || (p.speakSimilarity !== undefined && p.speakSimilarity !== null && p.speakSimilarity !== '');
          if(hasInteracted) {
              totalLearned++;
              totalCorrect += (p.cCorrect + p.sCorrect);
              totalWrong += (p.cIncorrect + p.sIncorrect);

              if(p.prof === '熟練') profCounts['熟練']++;
              else if(p.prof === '模糊') profCounts['模糊']++;
              else if(p.prof === '完全忘記') profCounts['完全忘記']++;

              let l = p.lang || '未分類';
              if(!langStats[l]) langStats[l] = { count: 0, correct: 0, wrong: 0, prof3: 0, prof2: 0, prof1: 0, speakSimSum: 0, speakSimCount: 0 };
              
              langStats[l].count++;
              langStats[l].correct += (p.cCorrect + p.sCorrect);
              langStats[l].wrong += (p.cIncorrect + p.sIncorrect);
              if(p.prof === '熟練') langStats[l].prof3++;
              else if(p.prof === '模糊') langStats[l].prof2++;
              else if(p.prof === '完全忘記') langStats[l].prof1++;
              
              if (p.speakSimilarity !== undefined && p.speakSimilarity !== null && p.speakSimilarity !== '') {
                  let sim = Number(p.speakSimilarity);
                  if(!isNaN(sim)) {
                     langStats[l].speakSimSum += sim;
                     langStats[l].speakSimCount++;
                  }
              }
          }
      });

      document.getElementById('dashTotalWords').innerText = totalLearned;

      let tbody = document.getElementById('dashLangTable');
      tbody.innerHTML = '';
      for (let lang in langStats) {
         let st = langStats[lang];
         let totalQs = st.correct + st.wrong;
         let acc = totalQs > 0 ? Math.round((st.correct / totalQs) * 100) + '%' : '-';
         let avgSpeak = st.speakSimCount > 0 ? Math.round(st.speakSimSum / st.speakSimCount) + '%' : '-';

         let tr = document.createElement('tr');
         tr.innerHTML = `
            <td>${lang}</td>
            <td style="font-weight:bold;">${st.count}</td>
            <td>${acc}</td>
            <td style="color:var(--primary); font-weight:bold;">${avgSpeak}</td>
            <td style="color:var(--success); font-weight:bold;">${st.prof3}</td>
            <td style="color:var(--warn); font-weight:bold;">${st.prof2}</td>
            <td style="color:var(--danger); font-weight:bold;">${st.prof1}</td>
         `;
         tbody.appendChild(tr);
      }

      document.getElementById('accContainer').innerHTML = '<canvas id="accuracyChart"></canvas>';
      let ctxAcc = document.getElementById('accuracyChart').getContext('2d');
      
      let accData = (totalCorrect === 0 && totalWrong === 0) ? [1] : [totalCorrect, totalWrong];
      let accColors = (totalCorrect === 0 && totalWrong === 0) ? ['#e5e7eb'] : ['#10b981', '#ef4444'];
      let accLabels = (totalCorrect === 0 && totalWrong === 0) ? ['無測驗紀錄'] : ['答對總數', '答錯總數'];
      
      new Chart(ctxAcc, {
          type: 'doughnut',
          data: {
              labels: accLabels,
              datasets: [{ data: accData, backgroundColor: accColors }]
          },
          options: { maintainAspectRatio: false, responsive: true, layout: { padding: 10 }, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, padding: 15 } } } }
      });

      document.getElementById('profContainer').innerHTML = '<canvas id="profChart"></canvas>';
      let ctxProf = document.getElementById('profChart').getContext('2d');
      
      let pData = (profCounts['熟練']===0 && profCounts['模糊']===0 && profCounts['完全忘記']===0) ? [1] : [profCounts['熟練'], profCounts['模糊'], profCounts['完全忘記']];
      let pColors = (profCounts['熟練']===0 && profCounts['模糊']===0 && profCounts['完全忘記']===0) ? ['#e5e7eb'] : ['#10b981', '#f59e0b', '#ef4444'];
      let pLabels = (profCounts['熟練']===0 && profCounts['模糊']===0 && profCounts['完全忘記']===0) ? ['無熟練度紀錄'] : ['熟練', '模糊', '忘記'];

      new Chart(ctxProf, {
          type: 'pie',
          data: {
              labels: pLabels,
              datasets: [{ data: pData, backgroundColor: pColors }]
          },
          options: { maintainAspectRatio: false, responsive: true, layout: { padding: 10 }, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, padding: 15 } } } }
      });
  } catch (e) { 
      console.error("Dashboard Render Error:", e); 
      alert("圖表載入發生錯誤：" + e.message); 
  }
}

function applyQuizProficiency(item, level) {
  item.proficiency = level;
  let nextDate = new Date();
  if (level === '完全忘記') nextDate.setMinutes(nextDate.getMinutes() + 10);
  else if (level === '模糊') nextDate.setDate(nextDate.getDate() + 1);
  else if (level === '熟練') nextDate.setDate(nextDate.getDate() + 5);

  item.nextReview = nextDate.toISOString();
  updateGlobalProgressLocally(item.word, 'prof', null, level);
  return item.nextReview;
}

function applySort() {
  if (allVocabData.length === 0) return;
  var sortType = document.getElementById('sortSelect').value;
  var isDaily30 = document.getElementById('daily30Check').checked;
  var selectedLevel = document.getElementById('levelSelect') ? document.getElementById('levelSelect').value : 'all';

  var sortedData = allVocabData.slice();

  if (selectedLevel !== 'all') {
    sortedData = sortedData.filter(item => item.level === selectedLevel);
  }

  if (sortedData.length === 0) {
     showToast("此等級目前沒有單字喔！為您切換回全部等級。", "warn");
     document.getElementById('levelSelect').value = 'all';
     applySort(); 
     return;
  }

  if (sortType === 'sm2') {
    let now = new Date().getTime();
    sortedData = sortedData.filter(item => {
       if (!item.nextReview) return true; 
       return new Date(item.nextReview).getTime() <= now; 
    });
    sortedData.sort(function(a, b) {
      let aTime = a.nextReview ? new Date(a.nextReview).getTime() : 0;
      let bTime = b.nextReview ? new Date(b.nextReview).getTime() : 0;
      return aTime - bTime;
    });
  } 
  else if (sortType === 'speakSm2') {
    let now = new Date().getTime();
    sortedData = sortedData.filter(item => {
       if (item.speakNextReview && item.speakNextReview.includes('2099')) return false; 
       if (!item.speakNextReview) return true; 
       return new Date(item.speakNextReview).getTime() <= now; 
    });
    sortedData.sort(function(a, b) {
      let aTime = a.speakNextReview ? new Date(a.speakNextReview).getTime() : 0;
      let bTime = b.speakNextReview ? new Date(b.speakNextReview).getTime() : 0;
      return aTime - bTime;
    });
  }
  else if (sortType === 'random') {
    for (var i = sortedData.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var temp = sortedData[i];
      sortedData[i] = sortedData[j];
      sortedData[j] = temp;
    }
  } else if (sortType === 'proficiency') {
    var getProfValue = function(p) {
      if (p === '完全忘記') return 1;
      if (p === '模糊') return 2;
      if (p === '熟練') return 3;
      return 0; 
    };
    sortedData.sort((a, b) => getProfValue(a.proficiency) - getProfValue(b.proficiency));
  } else if (sortType === 'error') {
    sortedData.sort(function(a, b) {
      var scoreA = (a.choiceIncorrect + a.spellIncorrect) - (a.choiceCorrect + a.spellCorrect);
      var scoreB = (b.choiceIncorrect + b.spellIncorrect) - (b.choiceCorrect + b.spellCorrect);
      return scoreB - scoreA; 
    });
  }
  
  if (isDaily30) {
    vocabData = sortedData.slice(0, 30);
  } else {
    vocabData = sortedData;
  }
  
  currentBrowseIndex = 0; 
  isCurrentWordAnswered = false; 
  
  if(document.getElementById('browse').classList.contains('active')) loadBrowseCard();
  else if(document.getElementById('choice').classList.contains('active')) loadNextChoice();
  else if(document.getElementById('spelling').classList.contains('active')) loadNextSpelling();
  else if(document.getElementById('speaking').classList.contains('active')) loadNextSpeaking();
  else if(document.getElementById('ai_tutor').classList.contains('active')) loadNextAITutor();
}

function updateProgressUI() {
  if(vocabData.length === 0) {
     document.getElementById('globalProgress').style.display = 'none';
     return;
  }
  document.getElementById('globalProgress').style.display = 'block';
  var current = currentBrowseIndex + 1;
  var total = vocabData.length;
  document.getElementById('progressText').innerText = "目前進度：" + current + " / " + total;
  document.getElementById('progressFill').style.width = ((current / total) * 100) + "%";
}

function setProficiency(level) {
  if(vocabData.length === 0) return;
  var item = vocabData[currentBrowseIndex];
  let nextReviewISO = applyQuizProficiency(item, level);
  updateProficiencyUI(level); 
  
  apiCall('updateProficiency', {
     sheetName: currentSheetName,
     word: item.word,
     profValue: level,
     username: currentUser,
     nextReviewISO: nextReviewISO
  });
}

function updateProficiencyUI(level) {
  document.getElementById('prof-btn-1').className = 'prof-btn';
  document.getElementById('prof-btn-2').className = 'prof-btn';
  document.getElementById('prof-btn-3').className = 'prof-btn';
  if (level === '完全忘記') document.getElementById('prof-btn-1').classList.add('active-1');
  if (level === '模糊') document.getElementById('prof-btn-2').classList.add('active-2');
  if (level === '熟練') document.getElementById('prof-btn-3').classList.add('active-3');
}

function switchTab(tabId) {
  if (vocabData.length === 0 && tabId !== 'dashboard') return; 
  var sections = document.querySelectorAll('.section');
  for (var i = 0; i < sections.length; i++) sections[i].classList.remove('active');
  var tabs = document.querySelectorAll('.tab-btn');
  for (var j = 0; j < tabs.length; j++) tabs[j].classList.remove('active');
  
  document.getElementById(tabId).classList.add('active');
  var activeBtn = document.querySelector('.tab-btn[onclick="switchTab(\'' + tabId + '\')"]');
  if (activeBtn) activeBtn.classList.add('active');

  if(tabId === 'dashboard') {
     document.getElementById('globalControls').style.display = 'none';
     document.getElementById('globalProgress').style.display = 'none'; 
     setTimeout(() => { renderDashboard(); }, 150);
     return;
  } 
  else if (tabId === 'records') {
     document.getElementById('globalControls').style.display = 'none';
     document.getElementById('globalProgress').style.display = 'none'; 
     renderAIHistory();
     renderBlacklist();
     return;
  } else {
     document.getElementById('globalControls').style.display = 'flex';
  }

  let sortSelect = document.getElementById('sortSelect');
  let needsResort = false;

  if (tabId === 'speaking' || tabId === 'ai_tutor') {
      if (sortSelect.value !== 'speakSm2') {
          sortSelect.value = 'speakSm2';
          needsResort = true;
      }
  } else if (tabId === 'browse' || tabId === 'choice' || tabId === 'spelling') {
      if (sortSelect.value === 'speakSm2') {
          sortSelect.value = 'sm2';
          needsResort = true;
      }
  }

  if (needsResort) {
      applySort(); 
      return; 
  }

  if (isCurrentWordAnswered) {
      currentBrowseIndex = (currentBrowseIndex + 1) % vocabData.length;
      isCurrentWordAnswered = false; 
  }

  if (tabId === 'browse') loadBrowseCard();
  if (tabId === 'choice') loadNextChoice();
  if (tabId === 'spelling') loadNextSpelling();
  if (tabId === 'speaking') loadNextSpeaking(); 
  if (tabId === 'ai_tutor') loadNextAITutor();
}

function updateBackend(word, isCorrect, userInput, testType, autoProfLevel, nextReviewISO, speakSimilarity) {
  apiCall('updateTestResult', {
     sheetName: currentSheetName,
     word: word,
     isCorrect: isCorrect,
     userInput: userInput,
     testType: testType,
     username: currentUser,
     autoProfLevel: autoProfLevel,
     nextReviewISO: nextReviewISO,
     speakSimilarity: speakSimilarity
  });
}