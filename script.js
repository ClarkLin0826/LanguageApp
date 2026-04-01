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
// 🌙 深色模式切換邏輯 (新加入)
// --------------------------------------------------------
function initTheme() {
    const themeToggleBtn = document.getElementById('theme-toggle');
    const rootElement = document.documentElement;
    const savedTheme = localStorage.getItem('theme');
    
    // 如果之前存的是 dark，就加上 class 並換圖示
    if (savedTheme === 'dark') {
        rootElement.classList.add('dark-theme');
        if(themeToggleBtn) themeToggleBtn.innerText = '☀️';
    }

    if(themeToggleBtn) {
        themeToggleBtn.addEventListener('click', () => {
            rootElement.classList.toggle('dark-theme');
            if (rootElement.classList.contains('dark-theme')) {
                localStorage.setItem('theme', 'dark');
                themeToggleBtn.innerText = '☀️';
            } else {
                localStorage.setItem('theme', 'light');
                themeToggleBtn.innerText = '🌙';
            }
        });
    }
}

// --------------------------------------------------------
// 💡 萬能 AI 文字助理引擎 (包含即時錯字分析與記憶法)
// --------------------------------------------------------
function getAITextAssistant(taskType, userInput = '') {
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
    // 💡 即時拼字錯誤分析 (不存檔)
    else if (taskType === 'spelling_analysis') {
        let btn = document.getElementById('spellingAnalysisBtn');
        btn.innerText = "⏳ 老師分析中...";
        btn.disabled = true;

        let box = document.getElementById('spellingAnalysisBox');
        box.style.display = 'block';
        box.innerHTML = '<div class="spinner" style="display:inline-block; width:15px; height:15px; border-top-color:var(--danger);"></div> 正在抓出你的拼字盲點...';

        apiCall('getAIAssistant', { taskType: 'spelling', word: word, lang: lang, extraInfo: userInput, userRole: userRole }, function(res) {
            btn.style.display = 'none'; 
            if(res.success) {
                box.innerHTML = `🕵️‍♂️ <b>AI 抓漏：</b><br>${marked.parse(res.reply)}`;
            } else box.innerHTML = `❌ ${res.message}`;
        }, function(err) { 
            box.innerHTML = `❌ 連線失敗`; 
            btn.innerText = "🕵️‍♂️ 分析我為什麼拼錯"; 
            btn.disabled = false; 
        });
    }
    // 💡 記憶法 (會存入 R 欄)
    else if (taskType === 'mnemonic' || taskType === 'spelling_mnemonic') {
        let box = taskType === 'mnemonic' ? document.getElementById('aiMnemonicBox') : document.getElementById('spellingMnemonicBox');
        
        box.style.display = 'block';
        if (taskType === 'spelling_mnemonic') {
            box.innerHTML += '<div id="tempMnemonicLoader"><div class="spinner" style="display:inline-block; width:15px; height:15px; border-top-color:var(--premium);"></div> AI 正在發功找記憶法...</div>';
        } else {
            box.innerHTML = '<div class="spinner" style="display:inline-block; width:15px; height:15px; border-top-color:var(--premium);"></div> AI 正在發功找記憶法...';
        }
        
        if (taskType === 'mnemonic') {
            let btn = document.getElementById('aiMnemonicBtn');
            if(btn) btn.style.display = 'none';
        } else {
            let btn = document.getElementById('spellingMnemonicBtn');
            if(btn) btn.style.display = 'none';
        }

        apiCall('getAIAssistant', { taskType: 'mnemonic', word: word, lang: lang, userRole: userRole }, function(res) {
            let loader = document.getElementById('tempMnemonicLoader');
            if(loader) loader.remove();

            if(res.success) {
                let mnemonicHtml = `<div style="margin-top: 1rem;">💡 <b>專屬 AI 記憶法：</b><br>${marked.parse(res.reply)}</div>`;
                if (taskType === 'spelling_mnemonic') {
                    box.innerHTML += mnemonicHtml;
                } else {
                    box.innerHTML = mnemonicHtml;
                }
                
                currentQuizItem.aiMnemonic = res.reply;
                updateGlobalProgressLocallyText(word, 'mnemonic', res.reply);
                apiCall('saveAIText', {sheetName: currentSheetName, word: word, username: currentUser, textType: 'mnemonic', textData: res.reply});
            } else {
                if (taskType === 'spelling_mnemonic') box.innerHTML += `<div>❌ ${res.message}</div>`;
                else box.innerHTML = `❌ ${res.message}`;
            }
        }, function(err) { 
            let loader = document.getElementById('tempMnemonicLoader');
            if(loader) loader.remove();
            
            if (taskType === 'spelling_mnemonic') box.innerHTML += `<div>❌ 連線失敗</div>`;
            else box.innerHTML = `❌ 連線失敗`;
        });
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
  
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    document.body.appendChild(container);
  }
  
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
     showToast("請先替換你的 API_URL！", "danger");
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

// --------------------------------------------------------
// 💡 帳號登入與記憶功能
// --------------------------------------------------------
function checkAutoLogin() {
  const savedUser = localStorage.getItem('vocab_user');
  const savedRole = localStorage.getItem('vocab_role');
  const loginTime = localStorage.getItem('vocab_time');

  if (savedUser && savedRole && loginTime) {
    const now = new Date().getTime();
    const oneWeek = 7 * 24 * 60 * 60 * 1000; 

    if (now - parseInt(loginTime) < oneWeek) {
      currentUser = savedUser;
      userRole = savedRole;
      
      document.getElementById('userBadge').innerText = "👤 " + currentUser;
      document.getElementById('loginScreen').style.display = 'none';
      document.getElementById('welcomeScreen').style.display = 'flex';
      
      showToast(`歡迎回來，${currentUser}！`, 'success');
      return true;
    } else {
      clearAuthData();
    }
  }
  return false;
}

function clearAuthData() {
  localStorage.removeItem('vocab_user');
  localStorage.removeItem('vocab_role');
  localStorage.removeItem('vocab_time');
}

function toggleAuthMode(mode) {
  document.getElementById('loginBox').style.display = 'none';
  document.getElementById('registerBox').style.display = 'none';
  document.getElementById('forgotBox').style.display = 'none';

  if(mode === 'login') document.getElementById('loginBox').style.display = 'flex';
  else if(mode === 'register') document.getElementById('registerBox').style.display = 'flex';
  else if(mode === 'forgot') document.getElementById('forgotBox').style.display = 'flex';
}

function doRegister() {
  var u = document.getElementById('regUser').value.trim();
  var p = document.getElementById('regPass').value.trim();
  var e = document.getElementById('regEmail').value.trim(); 
  var err = document.getElementById('regError');
  if(!u || !p || !e) { err.innerText = "帳號、密碼與信箱都不能為空！"; err.style.display='block'; return; }
  
  document.getElementById('regLoader').style.display = 'block';
  err.style.display = 'none';
  
  apiCall('registerUser', {username: u, password: p, email: e}, function(res) {
    document.getElementById('regLoader').style.display = 'none';
    if(res.success) {
       showToast(res.message, 'success');
       toggleAuthMode('login');
       document.getElementById('loginUser').value = u;
    } else {
       err.innerText = res.message;
       err.style.display = 'block';
    }
  }, function(error) {
    document.getElementById('regLoader').style.display = 'none';
    err.innerText = "伺服器連線錯誤，請確認 API 網址是否正確。";
    err.style.display = 'block';
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
       
       localStorage.setItem('vocab_user', currentUser);
       localStorage.setItem('vocab_role', userRole);
       localStorage.setItem('vocab_time', new Date().getTime().toString());

       showToast(`歡迎回來，${u}！`, 'success');
       
       document.getElementById('userBadge').innerText = "👤 " + u;
       document.getElementById('loginScreen').style.display = 'none';
       document.getElementById('welcomeScreen').style.display = 'flex'; 
    } else {
       err.innerText = res.message;
       err.style.display = 'block';
    }
  }, function(error) {
    document.getElementById('loginLoader').style.display = 'none';
    err.innerText = "伺服器 500 錯誤：請確認後端是否發布為「新版本」。";
    err.style.display = 'block';
  });
}

function doForgotPassword() {
  var e = document.getElementById('forgotEmail').value.trim();
  var err = document.getElementById('forgotError');
  if(!e) { err.innerText = "請輸入註冊時的信箱！"; err.style.display='block'; return; }

  document.getElementById('forgotLoader').style.display = 'block';
  err.style.display = 'none';

  apiCall('forgotPassword', {email: e}, function(res) {
    document.getElementById('forgotLoader').style.display = 'none';
    if(res.success) {
       showToast(res.message, 'success');
       document.getElementById('forgotEmail').value = '';
       toggleAuthMode('login');
    } else {
       err.innerText = res.message;
       err.style.display = 'block';
    }
  }, function(error) { 
    document.getElementById('forgotLoader').style.display = 'none';
    err.innerText = "連線失敗，請檢查後端發布狀態。";
    err.style.display = 'block';
  });
}

function doLogout() {
  currentUser = '';
  userRole = 'free';
  allVocabData = [];
  vocabData = [];
  globalProgressData = [];
  document.getElementById('loginPass').value = '';
  
  clearAuthData();
  
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
        div.style.borderBottom = '1px solid var(--border)';
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

// --------------------------------------------------------
// 🚀 網頁載入時啟動深色模式
// --------------------------------------------------------
window.onload = function() {
  initTheme(); // 💡 觸發深色模式偵測
  checkAutoLogin(); 

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
      '泰文': { count: 0, correct: 0, wrong: 0, prof3: 0, prof2: 0, prof1: 0, speakSimSum: 0, speakSimCount: 0 },
      '越南文': { count: 0, correct: 0, wrong: 0, prof3: 0, prof2: 0, prof1: 0, speakSimSum: 0, speakSimCount: 0 }
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

function updateGlobalProgressLocally(word, type, isCorrect, profLevel, speakSim) {
  let p = globalProgressData.find(item => item.word === word && item.lang === currentSheetName);
  if (!p) {
      p = { lang: currentSheetName, word: word, cCorrect: 0, cIncorrect: 0, sCorrect: 0, sIncorrect: 0, prof: '', speakSimilarity: '', speakNextReview: '', aiWordFeedback: '', aiExampleFeedback: '', aiWordAudioUrl: '', aiExampleAudioUrl: '', aiSentence: '', aiMnemonic: '' };
      globalProgressData.push(p);
  }
  if (type === 'choice') {
      if (isCorrect) p.cCorrect++; else p.cIncorrect++;
  } else if (type === 'spelling') {
      if (isCorrect) p.sCorrect++; else p.sIncorrect++;
  } else if (type === 'prof') {
      p.prof = profLevel;
  }
  
  if (speakSim !== undefined && speakSim !== null) {
      p.speakSimilarity = speakSim;
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

if (speechSynthesis.onvoiceschanged !== undefined) {
  speechSynthesis.onvoiceschanged = function() { synth.getVoices(); };
}

function speakText(text) {
  if (!text) return;
  synth.cancel(); 
  var select = document.getElementById('langSelect');
  var lang = select.options[select.selectedIndex].getAttribute('data-tts');
  var rate = document.getElementById('rateSelect').value;
  var utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = lang;
  utterance.rate = parseFloat(rate);
  var voices = synth.getVoices();
  var bestVoice = null;
  for (var i = 0; i < voices.length; i++) {
    var voiceLang = voices[i].lang.replace('_', '-');
    if (voiceLang.includes(lang) || lang.includes(voiceLang)) {
      bestVoice = voices[i];
      if (voices[i].name.includes('Google')) break;
    }
  }
  if (bestVoice) utterance.voice = bestVoice;
  synth.speak(utterance);
}

function renderPhoneticAndPos(phonetic, pos) {
  var html = '<span>' + (phonetic ? phonetic : '') + '</span>';
  if (pos) html += '<span class="pos-badge">' + pos + '</span>';
  return html;
}

function formatWordDisplay(item) {
  var txt = item.word;
  if (item.kanji) txt += ' (' + item.kanji + ')';
  return txt;
}

function nextQuizWord(tab) {
  if(vocabData.length === 0) return;
  currentBrowseIndex = (currentBrowseIndex + 1) % vocabData.length;
  isCurrentWordAnswered = false; 
  if (tab === 'choice') loadNextChoice();
  if (tab === 'spelling') loadNextSpelling();
  if (tab === 'speaking') loadNextSpeaking();
  if (tab === 'ai_tutor') loadNextAITutor();
}

function loadBrowseCard() {
  isCurrentWordAnswered = false; 
  
  if(vocabData.length === 0) {
     return;
  }

  currentQuizItem = vocabData[currentBrowseIndex]; 

  document.getElementById('browseWord').innerText = formatWordDisplay(currentQuizItem);
  document.getElementById('browseWord').setAttribute('data-speak', currentQuizItem.word);
  document.getElementById('browsePhonetic').innerHTML = renderPhoneticAndPos(currentQuizItem.phonetic, currentQuizItem.pos);
  
  document.getElementById('browseTranslation').innerText = currentQuizItem.translation;
  document.getElementById('browseTranslation').classList.remove('revealed');
  
  if (currentQuizItem.aiSentence) {
      let exMatch = currentQuizItem.aiSentence.match(/【例句】(.*?)(?=【翻譯】|$)/s);
      let transMatch = currentQuizItem.aiSentence.match(/【翻譯】(.*)/s);
      if(exMatch) document.getElementById('browseExample').innerText = "✨ " + exMatch[1].trim();
      else document.getElementById('browseExample').innerText = "✨ " + currentQuizItem.aiSentence;
      
      if(transMatch) {
          document.getElementById('browseExampleTrans').innerText = transMatch[1].trim();
          document.getElementById('browseExampleTrans').classList.remove('revealed');
      }
  } else {
      var displayExample = currentQuizItem.example ? currentQuizItem.example.replace(/\[|\]/g, '') : '';
      document.getElementById('browseExample').innerText = displayExample;
      document.getElementById('browseExampleTrans').innerText = currentQuizItem.exampleTranslation || '';
      if(currentQuizItem.exampleTranslation) document.getElementById('browseExampleTrans').classList.remove('revealed');
      else document.getElementById('browseExampleTrans').classList.add('revealed'); 
  }
  
  let aiSentenceBtn = document.getElementById('aiSentenceBtn');
  let aiMnemonicBtn = document.getElementById('aiMnemonicBtn');

  if (userRole === 'premium' || userRole === 'admin') {
      aiSentenceBtn.innerText = "🔄 AI 換情境";
      aiSentenceBtn.onclick = function() { getAITextAssistant('sentence'); };
      aiSentenceBtn.style.opacity = "1";
      
      aiMnemonicBtn.innerText = "🆘 AI 記憶法";
      aiMnemonicBtn.onclick = function() { getAITextAssistant('mnemonic'); };
      aiMnemonicBtn.style.opacity = "1";
  } else {
      aiSentenceBtn.innerText = "🔒 AI 換情境";
      aiSentenceBtn.onclick = function() { showToast('請升級 Premium 解鎖 AI 助理', 'warn'); };
      aiSentenceBtn.style.opacity = "0.6";

      aiMnemonicBtn.innerText = "🔒 AI 記憶法";
      aiMnemonicBtn.onclick = function() { showToast('請升級 Premium 解鎖 AI 助理', 'warn'); };
      aiMnemonicBtn.style.opacity = "0.6";
  }
  
  let mBox = document.getElementById('aiMnemonicBox');
  if (currentQuizItem.aiMnemonic) {
      mBox.style.display = 'block';
      mBox.innerHTML = `💡 <b>專屬 AI 記憶法：</b><br>${marked.parse(currentQuizItem.aiMnemonic)}`;
  } else {
      mBox.style.display = 'none'; mBox.innerHTML = '';
  }
  
  updateProficiencyUI(currentQuizItem.proficiency);
  updateProgressUI(); 

  if(document.getElementById('autoAudioCheck').checked) speakText(currentQuizItem.word);

  var imgEl = document.getElementById('browseImg');
  imgEl.onerror = null; 
  if (currentQuizItem.imageUrl) {
    imgEl.src = currentQuizItem.imageUrl; imgEl.style.display = 'block';
  } else {
    var query = currentQuizItem.word; 
    if (currentQuizItem.example) {
      var words = currentQuizItem.example.match(/\b[a-zA-Z]{4,}\b/g);
      if (words && words.length > 0) query = words.join(',');
    }
    imgEl.src = "https://loremflickr.com/400/300/" + encodeURIComponent(query) + "/all?lock=" + Date.now();
    imgEl.style.display = 'block';
    imgEl.onerror = function() { this.style.display = 'none'; };
  }
}

function prevWord() {
  if(vocabData.length === 0) return;
  currentBrowseIndex = (currentBrowseIndex - 1 + vocabData.length) % vocabData.length;
  loadBrowseCard();
}

function nextWord() {
  if(vocabData.length === 0) return;
  currentBrowseIndex = (currentBrowseIndex + 1) % vocabData.length;
  loadBrowseCard();
}

function loadNextChoice() {
  window.scrollTo({ top: 0, behavior: 'smooth' });
  if(vocabData.length === 0) return;

  isCurrentWordAnswered = false; 
  currentQuizItem = vocabData[currentBrowseIndex]; 
  updateProgressUI(); 

  document.getElementById('choiceWord').innerText = formatWordDisplay(currentQuizItem);
  document.getElementById('choiceWord').setAttribute('data-speak', currentQuizItem.word);
  document.getElementById('choicePhonetic').innerHTML = renderPhoneticAndPos(currentQuizItem.phonetic, currentQuizItem.pos);
  document.getElementById('choiceFeedback').innerText = '';
  document.getElementById('choiceNextBtn').style.display = 'none';

  if(document.getElementById('autoAudioCheck').checked) speakText(currentQuizItem.word);

  var options = [currentQuizItem.translation];
  var otherVocabs = vocabData.filter(function(v) { return v.word !== currentQuizItem.word; });
  otherVocabs.sort(function() { return 0.5 - Math.random(); });
  for (var i = 0; i < Math.min(3, otherVocabs.length); options.push(otherVocabs[i].translation), i++);
  options.sort(function() { return 0.5 - Math.random(); });

  var optionsContainer = document.getElementById('choiceOptions');
  optionsContainer.innerHTML = '';
  for (var k = 0; k < options.length; k++) {
    (function(opt) {
      var btn = document.createElement('button');
      btn.className = 'option-btn'; btn.innerText = opt;
      btn.onclick = function() { checkChoice(btn, opt === currentQuizItem.translation); };
      optionsContainer.appendChild(btn);
    })(options[k]);
  }
  setTimeout(function() { document.getElementById('choiceWordDisplay').scrollIntoView({ behavior: 'smooth', block: 'center' }); }, 50);
}

function checkChoice(btnElement, isCorrect) {
  var buttons = document.getElementById('choiceOptions').querySelectorAll('.option-btn');
  for (var i = 0; i < buttons.length; i++) {
    buttons[i].disabled = true;
    if (buttons[i].innerText === currentQuizItem.translation) buttons[i].classList.add('correct');
  }
  
  isCurrentWordAnswered = true; 
  var feedback = document.getElementById('choiceFeedback');
  let nextRev = '';

  try {
      if (isCorrect) {
        btnElement.classList.add('correct');
        feedback.innerText = '答對了！'; feedback.className = 'feedback success';
        currentQuizItem.choiceCorrect++;
        updateGlobalProgressLocally(currentQuizItem.word, 'choice', true, null);
        
        nextRev = applyQuizProficiency(currentQuizItem, '熟練');
        updateBackend(currentQuizItem.word, true, null, 'choice', '熟練', nextRev); 
      } else {
        btnElement.classList.add('wrong');
        feedback.innerText = '答錯了！正確答案是：' + currentQuizItem.translation; feedback.className = 'feedback danger';
        currentQuizItem.choiceIncorrect++;
        updateGlobalProgressLocally(currentQuizItem.word, 'choice', false, null);
        
        nextRev = applyQuizProficiency(currentQuizItem, '完全忘記');
        updateBackend(currentQuizItem.word, false, null, 'choice', '完全忘記', nextRev);
      }
  } catch(e) {
      console.error(e);
  }

  document.getElementById('choiceNextBtn').style.display = 'block';
  setTimeout(function() { document.getElementById('choiceNextBtn').scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }, 50);
}

function loadNextSpelling() {
  window.scrollTo({ top: 0, behavior: 'smooth' });
  if(vocabData.length === 0) return;

  isCurrentWordAnswered = false; 
  currentQuizItem = vocabData[currentBrowseIndex]; 
  updateProgressUI(); 

  document.getElementById('spellingTranslation').innerText = currentQuizItem.translation;
  document.getElementById('spellingPhonetic').innerHTML = renderPhoneticAndPos(currentQuizItem.phonetic, currentQuizItem.pos);
  
  var exampleHtml = currentQuizItem.example;
  var ttsText = currentQuizItem.example; 
  
  if (exampleHtml) {
    if (exampleHtml.indexOf('[') !== -1 && exampleHtml.indexOf(']') !== -1) {
      ttsText = exampleHtml.replace(/\[|\]/g, '');
      exampleHtml = exampleHtml.replace(/\[.*?\]/g, '________');
    } else {
      var regex = new RegExp(currentQuizItem.word, 'gi');
      exampleHtml = exampleHtml.replace(regex, '________');
    }
  }
  
  var exampleEl = document.getElementById('spellingExample');
  exampleEl.innerText = exampleHtml;
  exampleEl.setAttribute('data-fulltext', ttsText);
  document.getElementById('spellingExampleTrans').innerText = currentQuizItem.exampleTranslation || '';

  var inputEl = document.getElementById('spellingInput');
  inputEl.value = ''; inputEl.disabled = false; inputEl.classList.remove('error-input');
  
  if (window.innerWidth > 768) inputEl.focus(); 
  else inputEl.blur(); 

  let spellMbox = document.getElementById('spellingMnemonicBox');
  if(spellMbox) {
      spellMbox.style.display = 'none';
      spellMbox.innerHTML = '';
  }
  
  document.getElementById('spellingFeedback').innerText = '';
  document.getElementById('spellingSubmitBtn').style.display = 'block';
  document.getElementById('spellingNextBtn').style.display = 'none';
  
  var hintEl = document.getElementById('spellingHint');
  if (currentSheetName === '日文') hintEl.innerText = "💡 提示：可輸入假名 或 漢字";
  else hintEl.innerText = "💡 提示：請填寫單字原型 (字典型態)";
  
  if(document.getElementById('autoAudioCheck').checked) speakText(ttsText ? ttsText : currentQuizItem.word);
  setTimeout(function() { document.getElementById('spellingTranslation').scrollIntoView({ behavior: 'smooth', block: 'center' }); }, 50);
}

function checkSpelling() {
  var inputEl = document.getElementById('spellingInput');
  var userInput = inputEl.value.trim().toLowerCase();
  if (!userInput) return;

  inputEl.disabled = true;
  document.getElementById('spellingSubmitBtn').style.display = 'none';

  isCurrentWordAnswered = true; 
  var isCorrect = false;
  var target1 = currentQuizItem.word.toLowerCase();
  var target2 = currentQuizItem.kanji ? currentQuizItem.kanji.toLowerCase() : '';
  if (userInput === target1 || (target2 !== '' && userInput === target2)) isCorrect = true;

  var feedback = document.getElementById('spellingFeedback');
  let nextRev = '';

  try {
      if (isCorrect) {
        feedback.innerText = '拼字正確！'; feedback.className = 'feedback success';
        inputEl.classList.remove('error-input');
        currentQuizItem.spellCorrect++;
        updateGlobalProgressLocally(currentQuizItem.word, 'spelling', true, null);
        
        nextRev = applyQuizProficiency(currentQuizItem, '熟練');
        updateBackend(currentQuizItem.word, true, null, 'spelling', '熟練', nextRev); 
      } else {
        inputEl.classList.add('shake', 'error-input');
        setTimeout(function() { inputEl.classList.remove('shake'); }, 400); 
        
        var errorMsg = '拼字錯誤！正確拼法為：' + currentQuizItem.word;
        if (currentQuizItem.kanji) errorMsg += ' 或 ' + currentQuizItem.kanji;
        if (currentQuizItem.errorLog) {
          var pastErrors = currentQuizItem.errorLog.split(',');
          var lastError = pastErrors[pastErrors.length - 1].trim();
          if (lastError) errorMsg += '\n(上次錯拼為：' + lastError + ')';
        }
        feedback.innerText = errorMsg; feedback.className = 'feedback danger';
        currentQuizItem.errorLog = currentQuizItem.errorLog ? currentQuizItem.errorLog + ', ' + userInput : userInput;
        
        let spellMbox = document.getElementById('spellingMnemonicBox');
        if (!spellMbox) {
            spellMbox = document.createElement('div');
            spellMbox.id = 'spellingMnemonicBox';
            spellMbox.className = 'ai-feedback-box';
            spellMbox.style.marginBottom = '1.5rem';
            spellMbox.style.borderRadius = '0.5rem';
            document.getElementById('spellingFeedback').parentNode.insertBefore(spellMbox, document.getElementById('spellingNextBtn'));
        }
        
        let safeUserInput = userInput.replace(/'/g, "\\'"); 
        
        let analysisHtml = `
            <div style="text-align: right; margin-bottom: 0.5rem;">
                <button id="spellingAnalysisBtn" class="btn-skip" style="font-size: 0.85rem; color: var(--danger); border-color: var(--danger);" onclick="getAITextAssistant('spelling_analysis', '${safeUserInput}')">🕵️‍♂️ 分析我為什麼拼錯</button>
            </div>
            <div id="spellingAnalysisBox" class="ai-feedback-box" style="margin-bottom: 1rem; border-left-color: var(--danger); color: #7f1d1d; background: #fef2f2; display:none;"></div>
        `;

        let mnemonicBtnHtml = '';
        if (currentQuizItem.aiMnemonic) {
            spellMbox.innerHTML = analysisHtml + `<div style="margin-top: 1rem;">💡 <b>專屬 AI 記憶法：</b><br>${marked.parse(currentQuizItem.aiMnemonic)}</div>`;
            spellMbox.style.display = 'block';
        } else {
            if (userRole === 'premium' || userRole === 'admin') {
                mnemonicBtnHtml = `<button id="spellingMnemonicBtn" class="btn-skip" style="font-size: 0.85rem; color: var(--premium); border-color: var(--premium);" onclick="getAITextAssistant('spelling_mnemonic')">🆘 產生 AI 記憶法</button>`;
            } else {
                mnemonicBtnHtml = `<button id="spellingMnemonicBtn" class="btn-skip" style="font-size: 0.85rem; color: var(--premium); border-color: var(--premium); opacity: 0.6;" onclick="showToast('請升級 Premium 解鎖 AI 助理', 'warn')">🔒 產生 AI 記憶法</button>`;
            }
            spellMbox.innerHTML = analysisHtml + `<div style="text-align: right; margin-bottom: 1rem;">${mnemonicBtnHtml}</div>`;
            spellMbox.style.display = 'block';
        }
        
        currentQuizItem.spellIncorrect++;
        updateGlobalProgressLocally(currentQuizItem.word, 'spelling', false, null);
        
        nextRev = applyQuizProficiency(currentQuizItem, '完全忘記');
        updateBackend(currentQuizItem.word, false, userInput, 'spelling', '完全忘記', nextRev); 
      }
  } catch (e) {
      console.error(e);
  }

  document.getElementById('spellingNextBtn').style.display = 'block';
  if (window.innerWidth > 768) document.getElementById('spellingNextBtn').focus(); 
  setTimeout(function() { document.getElementById('spellingNextBtn').scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }, 50);
}

function loadNextSpeaking() {
  window.scrollTo({ top: 0, behavior: 'smooth' });
  if(vocabData.length === 0) return;

  isCurrentWordAnswered = false; 
  currentQuizItem = vocabData[currentBrowseIndex]; 
  updateProgressUI(); 
  
  document.getElementById('speakingTranslation').innerText = currentQuizItem.translation;
  document.getElementById('speakingPhonetic').innerHTML = renderPhoneticAndPos(currentQuizItem.phonetic, currentQuizItem.pos);
  document.getElementById('speakingWord').innerText = formatWordDisplay(currentQuizItem);
  document.getElementById('speakingWord').setAttribute('data-speak', currentQuizItem.word);
  
  var prevSimEl = document.getElementById('speakingPrevSim');
  if (currentQuizItem.speakSimilarity !== undefined && currentQuizItem.speakSimilarity !== null && currentQuizItem.speakSimilarity !== '') {
     prevSimEl.innerText = "💡 上次發音相似度：" + currentQuizItem.speakSimilarity + "%";
  } else {
     prevSimEl.innerText = "";
  }

  document.getElementById('speakingFeedback').innerText = '';
  document.getElementById('recognizedText').innerText = '';
  document.getElementById('speakingNextBtn').style.display = 'none';
  document.getElementById('skipSpeakingBtn').style.display = 'inline-block'; 
  
  if (isRecording && recognition) recognition.stop();
  if(document.getElementById('autoAudioCheck').checked) speakText(currentQuizItem.word);
  setTimeout(function() { document.getElementById('speakingWord').scrollIntoView({ behavior: 'smooth', block: 'center' }); }, 50);
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
// ==========================================
// 🎯 智能分級測驗 (Placement Test) 核心邏輯 (完整版)
// ==========================================
let placementWords = [];
let placementIndex = 0;
let placementResults = {}; 

function startPlacementTest() {
  let uniqueLevels = [...new Set(allVocabData.map(w => w.level).filter(l => l))].sort();
  if(uniqueLevels.length === 0) {
      showToast("此語系的題庫尚未設定「等級(Level)」，無法進行分級測驗喔！", "warn");
      return;
  }

  placementWords = [];
  placementResults = {};
  
  uniqueLevels.forEach(lvl => {
      let wordsInLevel = allVocabData.filter(w => w.level === lvl);
      wordsInLevel.sort(() => 0.5 - Math.random());
      
      // 每個等級改抽 5 題，增加容錯率 (如果該等級不到5題就全拿)
      let selectedWords = wordsInLevel.slice(0, 5);
      placementWords.push(...selectedWords);
      
      placementResults[lvl] = { total: selectedWords.length, correct: 0 };
  });

  if(placementWords.length === 0) return;

  placementWords.sort(() => 0.5 - Math.random());
  placementIndex = 0;

  document.querySelectorAll('.section').forEach(el => el.classList.remove('active'));
  document.getElementById('globalControls').style.display = 'none';
  document.getElementById('globalProgress').style.display = 'none';
  document.getElementById('placementTestSection').classList.add('active');
  
  document.getElementById('placementQuizArea').style.display = 'block';
  document.getElementById('placementResultArea').style.display = 'none';

  loadPlacementQuestion();
}

// 💡 補回遺失的函數：負責把題目渲染到畫面上
function loadPlacementQuestion() {
  if (placementIndex >= placementWords.length) {
      finishPlacementTest();
      return;
  }

  let currentItem = placementWords[placementIndex];
  
  // 更新進度條
  let current = placementIndex + 1;
  let total = placementWords.length;
  document.getElementById('placementProgressText').innerText = `測驗進度：${current} / ${total}`;
  document.getElementById('placementProgressFill').style.width = ((current / total) * 100) + "%";

  // 顯示題目
  document.getElementById('placementWord').innerText = formatWordDisplay(currentItem);

  // 產生選項 (1個正確 + 3個隨機錯誤)
  let options = [currentItem.translation];
  let otherVocabs = allVocabData.filter(v => v.word !== currentItem.word);
  otherVocabs.sort(() => 0.5 - Math.random());
  for (let i = 0; i < Math.min(3, otherVocabs.length); i++) {
      options.push(otherVocabs[i].translation);
  }
  options.sort(() => 0.5 - Math.random());

  // 渲染選項按鈕
  let optionsContainer = document.getElementById('placementOptions');
  optionsContainer.innerHTML = '';
  options.forEach(opt => {
      let btn = document.createElement('button');
      btn.className = 'option-btn'; 
      btn.innerText = opt;
      btn.onclick = () => checkPlacementAnswer(opt === currentItem.translation, currentItem.level);
      optionsContainer.appendChild(btn);
  });
}

// 💡 補回遺失的函數：檢查答案對錯
function checkPlacementAnswer(isCorrect, level) {
  if (isCorrect) {
      placementResults[level].correct++;
  }
  placementIndex++;
  loadPlacementQuestion();
}

// 智能等級權重計算器，確保難度是由簡單到難排序
function getLevelWeight(lvl) {
  let str = String(lvl).toLowerCase();
  
  if (str.includes('n5')) return 10;
  if (str.includes('n4')) return 20;
  if (str.includes('n3')) return 30;
  if (str.includes('n2')) return 40;
  if (str.includes('n1')) return 50;

  if (str.includes('1급') || str.includes('1級')) return 10;
  if (str.includes('2급') || str.includes('2級')) return 20;
  if (str.includes('3급') || str.includes('3級')) return 30;
  if (str.includes('4급') || str.includes('4級')) return 40;
  if (str.includes('5급') || str.includes('5級')) return 50;
  if (str.includes('6급') || str.includes('6級')) return 60;

  let weight = 50; 
  if (str.includes('初級') || str.includes('入門') || str.includes('基礎')) weight = 15;
  if (str.includes('初中級')) weight = 25;
  if (str.includes('中級') && !str.includes('初中級') && !str.includes('中高級')) weight = 35;
  if (str.includes('中高級')) weight = 45;
  if (str.includes('高級') && !str.includes('中高級')) weight = 55;
  if (str.includes('進階')) weight = 65;
  
  return weight;
}

function finishPlacementTest() {
  document.getElementById('placementQuizArea').style.display = 'none';
  document.getElementById('placementResultArea').style.display = 'block';
  
  let detailsHtml = "";
  let dropLevelFound = false;

  let sortedLevels = Object.keys(placementResults).sort((a, b) => getLevelWeight(a) - getLevelWeight(b));
  let recommendedLevel = sortedLevels[0] || "全部等級"; 

  sortedLevels.forEach(lvl => {
      let res = placementResults[lvl];
      let accuracy = Math.round((res.correct / res.total) * 100) || 0;
      
      let accColor = accuracy >= 70 ? "var(--success)" : (accuracy >= 40 ? "var(--warn)" : "var(--danger)");
      detailsHtml += `<div style="display: flex; justify-content: space-between; border-bottom: 1px solid var(--border); padding: 0.5rem 0;">
                        <span>等級 <b>${lvl}</b></span>
                        <span style="color: ${accColor}; font-weight: bold;">${accuracy}% (${res.correct}/${res.total})</span>
                      </div>`;
      
      if (!dropLevelFound && accuracy < 70) {
          recommendedLevel = lvl;
          dropLevelFound = true;
      }
  });

  if (!dropLevelFound && sortedLevels.length > 0) {
      recommendedLevel = sortedLevels[sortedLevels.length - 1];
  }

  document.getElementById('placementDetails').innerHTML = detailsHtml;
  document.getElementById('placementRecommendation').innerHTML = `
      根據測驗結果，系統強烈建議您從<br>
      <span style="font-size: 2rem; color: var(--premium); font-weight: bold; display: inline-block; margin: 0.5rem 0;">${recommendedLevel}</span><br>
      開始學習！🚀
  `;

  // 💡 追加功能：自動將看板上方的等級過濾器切換到推薦的等級
  let levelSelect = document.getElementById('levelSelect');
  if(levelSelect && Array.from(levelSelect.options).some(opt => opt.value === recommendedLevel)) {
      levelSelect.value = recommendedLevel;
      applySort(); // 觸發系統重新整理單字堆
      showToast(`已為您自動載入「${recommendedLevel}」的單字！`, 'success');
  }
}