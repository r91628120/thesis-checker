
/*
 * 高中職小論文格式自檢系統
 * 主程式檔案：main.js
 * 版本：v2025-10-07a
 * 說明：
 *   - 第12項：僅檢查「陸／六、參考文獻」之後的文獻筆數（分類書籍/網路；展示前3筆）
 *   - 第13項：圖表需有標號/標題與來源，子行顯示、嚴格比對作者＋年份（資料來源／圖片來源）
 *   - 保留 UI、GAS、人次統計、樣式不變
 */


  const APP_VER = '2025-10-07a';
  (function () {
    try {
      const url = new URL(location.href);
      if (url.searchParams.get('v') !== APP_VER) {
        url.searchParams.set('v', APP_VER);
        location.replace(url.toString());
      }
    } catch (e) {}
  })();
  console.log("載入版本:", (new URL(location.href)).searchParams.get('v') || APP_VER);



  window.loadPdfjs = async () => {
    const base = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.2.67/";
    const { GlobalWorkerOptions, getDocument } = await import(base + "pdf.min.mjs");
    GlobalWorkerOptions.workerSrc = base + "pdf.worker.min.mjs";
    return { getDocument };
  };


document.write(new URL(location.href).searchParams.get('v')||APP_VER)


/* ===== 設定 ===== */
const GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzFsAnAyYm_gR9NTPqjECDhCjFOYZIDnhJ6NlHxe2Qr81Myb1J6_6ostG7PaCVAjcM/exec";
const ENABLE_COUNTER = true;
const CLOUD_RUN_BASE = "https://font-check-api-1009467346209.asia-east1.run.app";
const CLOUD_RUN_ENDPOINT = CLOUD_RUN_BASE + "/check";

/* ===== UI 變數 ===== */
const consent=document.getElementById('consent'),fileInput=document.getElementById('file'),drop=document.getElementById('drop');
const startBtn=document.getElementById('start'),clearBtn=document.getElementById('clear'),fontCheckBtn=document.getElementById('fontCheckBtn');
const statusEl=document.getElementById('status'),resultCard=document.getElementById('resultCard'),tbody=document.getElementById('tbody');
const downloadHtmlBtn=document.getElementById('downloadHtml'),copyBtn=document.getElementById('copyJson'),downloadJsonBtn=document.getElementById('downloadJson');
const jsonPreview=document.getElementById('jsonPreview'),counterBox=document.getElementById('counterBox'),counterEl=document.getElementById('counter'),fileTagEl=document.getElementById('fileTag');

/* ===== 狀態 ===== */
let arrayBuffer=null,fileName=null,fileBase=null,fileExt=null,lastSummaryJSON=null,_lastPageCount=0;
let _ruleCounter=0; const _rowsBuffer=[];
let _fontCheckMirror=null;

/* ===== 規則常數 ===== */
const RULES={
  minPages:4,maxPages:10,maxSizeMB:5,
  sixSections:['前言','文獻探討','研究方法','研究分析與結果','研究結論與建議','參考文獻'],
  quoteMaxChars:50,minQuoteCharsForDirect:10,marginCM:2.0
};
const CM_TO_PT=cm=>cm*72/2.54, MARGIN_PT=CM_TO_PT(RULES.marginCM);
const PUNCT=/[，。、．；：、！？—─\-…‧（）()〔〕【】《》〈〉「」『』“”"'\[\]{}<>‧·~`^_=+\|\\/:，,.;:!\?\s]/g;
const SCHOOL_PATTERNS=[/(國立|市立|縣立)[\u4e00-\u9fa5]{0,8}(高級中學|高中|高職|高工|家商|商工|農工|高農|職業學校|中學|國中|小學|大學|科技大學|科大|師範大學|大學附中|實驗高中|高中部)/g];

/* ===== 可用性 ===== */
function setUploadEnabled(ok){fileInput.disabled=!ok;const has=!!arrayBuffer;startBtn.disabled=!ok||!has;clearBtn.disabled=!ok||!has;fontCheckBtn.disabled=!ok||!has||fileExt!=='pdf'}
consent.addEventListener('change',()=>{setUploadEnabled(consent.checked);statusEl.textContent=consent.checked?'已同意，可上傳。':'請先同意政策。'});

/* DnD */
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();e.stopPropagation();if(!fileInput.disabled)drop.classList.add('drag');}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();e.stopPropagation();drop.classList.remove('drag');}));
drop.addEventListener('drop',e=>{if(fileInput.disabled)return;const f=e.dataTransfer.files?.[0];if(f)handleFile(f);});
fileInput.addEventListener('change',e=>{const f=e.target.files?.[0];if(f)handleFile(f);});
function baseName(n){return(n||'').replace(/\.[^.]+$/,'').trim();}
async function handleFile(f){
  const ext=(f.name.split('.').pop()||'').toLowerCase(); if(!['pdf','docx'].includes(ext)){statusEl.textContent='僅支援 PDF 或 DOCX';return;}
  if(f.size>RULES.maxSizeMB*1024*1024){statusEl.textContent=`檔案超過 ${RULES.maxSizeMB}MB 上限。`;return;}
  fileName=f.name;fileBase=baseName(f.name);fileExt=ext;arrayBuffer=await f.arrayBuffer();
  statusEl.textContent=`已選擇：${f.name}（${(f.size/1024/1024).toFixed(2)} MB）`;setUploadEnabled(consent.checked);
}
clearBtn.addEventListener('click',()=>{fileInput.value='';arrayBuffer=null;fileName=fileBase=fileExt=null;startBtn.disabled=clearBtn.disabled=fontCheckBtn.disabled=true;resultCard.style.display='none';tbody.innerHTML='';fileTagEl.textContent='';statusEl.textContent='已清除。';jsonPreview.textContent='';lastSummaryJSON=null;_rowsBuffer.length=0;_fontCheckMirror=null;});

/* ===== 主流程 ===== */
startBtn.addEventListener('click',async()=>{
  if(!arrayBuffer)return;tbody.innerHTML='';resultCard.style.display='none';statusEl.textContent='解析中…';_rowsBuffer.length=0;_ruleCounter=0;_fontCheckMirror=null;
  try{
    if(fileExt==='pdf'){await runPdfChecks(arrayBuffer);} else {await runDocxChecks(arrayBuffer);}
    resultCard.style.display='block';fileTagEl.textContent=`此次檢查檔案：${fileName||'(未命名)'}`;statusEl.textContent='完成檢查。';
    buildAndShowSummaryJSON({fileName,pages:_lastPageCount||0,passFailRows:_rowsBuffer.slice()});
    if(ENABLE_COUNTER){await bumpCounterEveryTime();await fetchCounter();}
  }catch(e){console.error(e);statusEl.textContent='解析失敗：請確認檔案是否有效。'}
});

/* ===== 進階字體檢查 ===== */
fontCheckBtn.addEventListener('click',async()=>{
  if(!arrayBuffer||fileExt!=='pdf'){alert('請先選擇 PDF 檔。');return;}
  try{
    statusEl.textContent='上傳至 Cloud Run 進行字體檢查…';
    const pdfBlob=new Blob([arrayBuffer],{type:"application/pdf"});
    const data=await fetchCloudRun(pdfBlob,fileName||'upload.pdf');
    if(data){
      const {pass,suggestion}=normalizeApiResult(data); const ev=buildEvidenceString(data);
      _fontCheckMirror={pass:Boolean(pass),evidence:ev,suggestion:suggestion||''};
      addRow('Cloud Run 字體字型檢查',!!pass,ev,suggestion||'請依規範修正字體。');
      syncBasicFontRuleMirror();
      buildAndShowSummaryJSON({fileName,pages:_lastPageCount||0,passFailRows:_rowsBuffer.slice()});
      resultCard.style.display='block';fileTagEl.textContent=`此次檢查檔案：${fileName||'(未命名)'}`;statusEl.textContent='Cloud Run 檢查完成。';
    }
  }catch(e){console.error(e);statusEl.textContent='Cloud Run 檢查失敗。'}
});
async function fetchCloudRun(fileBlob,nameForServer){
  const formData=new FormData(); formData.append("file",fileBlob,nameForServer||"upload.pdf");
  try{
    const res=await fetch(CLOUD_RUN_ENDPOINT,{method:"POST",body:formData,headers:{"Cache-Control":"no-cache"}});
    if(!res.ok){const t=await res.text().catch(()=> "");throw new Error(`API ${res.status} ${res.statusText} ${t}`)}
    return await res.json();
  }catch(err){
    console.error("Cloud Run 失敗：",err);
    addRow('Cloud Run 字體字型檢查',false,'CORS/快取問題：請確認 OPTIONS 與 Access-Control-Allow-Origin 設定。','請檢查 Cloud Run CORS 與公開存取設定。');return null;
  }
}
function normalizeApiResult(data){const pass=typeof data.pass==='boolean'?data.pass:(data.summary?.pass??data.ok??false);const suggestion=data.suggestion||data.summary?.suggestion||'';return {pass,suggestion}}
function buildEvidenceString(data){
  try{const parts=[]; if(data.summary?.pages)parts.push(`頁數：${data.summary.pages}`);
    if(data.summary?.dominant_fonts)parts.push(`主要字體：${data.summary.dominant_fonts.join('、')}`);
    if(Array.isArray(data.top_fonts))parts.push('常見字體：'+data.top_fonts.slice(0,6).join('、'));
    if(data.layout?.per_page_margins?.length){const f=data.layout.per_page_margins[0];parts.push(`第1頁邊界量測：左${f.left} / 右${f.right} / 上${f.top} / 下${f.bottom} pt`);}
    return parts.join('｜')||JSON.stringify(data).slice(0,800);
  }catch{return typeof data==='object'?JSON.stringify(data).slice(0,800):String(data)}
}

/* ===== PDF 檢查 ===== */
async function runPdfChecks(buf){
  const {getDocument}=await window.loadPdfjs(); const pdf=await getDocument({data:buf}).promise;
  const numPages=pdf.numPages; _lastPageCount=numPages;

  const pageTexts=[];
  addRow('頁數 4–10',(numPages>=RULES.minPages&&numPages<=RULES.maxPages),`共 ${numPages} 頁`,'若少於4或多於10頁，請修訂篇幅。');

  let pageNumOkPages=[];
  const headerTitleTexts=[];
  let minLeft=Infinity,minRight=Infinity,minTop=Infinity,minBottom=Infinity;
  let firstHeaderOK=false, firstFooterOK=false;

  for(let p=1;p<=numPages;p++){
    const page=await pdf.getPage(p); const viewport=page.getViewport({scale:1}); const W=viewport.width,H=viewport.height;
    const tc=await page.getTextContent();

    const topBandMinY = H - MARGIN_PT;
    const bottomBandMaxY = MARGIN_PT;
    const centerTol = Math.max(W*0.15,72);

    let hasCenterPageNum=false;
    const pageBuf=[];
    let _lastXY=null;

    for(const it of tc.items){
      const tr=it.transform; const x=tr[4], y=tr[5]; const w=(typeof it.width==='number')?it.width:0;
      const txt=(it.str||'').trim();

      if(txt){
        if(_lastXY){
          const newLine =
            (Math.abs(y - _lastXY.y) > 6 && x < _lastXY.x + 20) ||
            (y < _lastXY.y - 2 && x <= _lastXY.x);
          if(newLine) pageBuf.push('\n');
        }
        pageBuf.push(txt);
        _lastXY = {x,y};
      }

      if(p===1 && txt){
        const mid = x + w/2;
        const centeredByMid = (mid >= (W/2 - centerTol) && mid <= (W/2 + centerTol));
        const inTop2cm = (y >= topBandMinY);
        const looksLikeTitle = !/^\d+$/.test(txt) && txt.replace(/\s/g,'').length >= 2;
        if(inTop2cm && centeredByMid && looksLikeTitle){ firstHeaderOK = true; }
      }

      const centeredBottomByMid = (()=>{const mid=x+w/2;return (mid >= (W/2 - centerTol) && mid <= (W/2 + centerTol));})();
      const inBottom2cm = (y <= bottomBandMaxY);
      const m = txt.match(/^\s*(?:第)?\s*[-—]*\s*(\d{1,4})\s*[-—]*(?:\s*頁)?\s*$/);
      if(inBottom2cm && centeredBottomByMid && m){
        const num=parseInt(m[1],10);
        hasCenterPageNum=true;
        if(num!==p) {/* 忽略是否連號 */}
        if(p===1) firstFooterOK = true;
      }

      if(y>=topBandMinY && txt){ headerTitleTexts.push(txt); }

      const inBody=(y>MARGIN_PT)&&(y<(H-MARGIN_PT));
      if(inBody){
        const distLeft=x, distRight=W-(x+w), distBottom=y, distTop=H-y;
        if(distLeft<minLeft)minLeft=distLeft; if(distRight<minRight)minRight=distRight; if(distBottom<minBottom)minBottom=distBottom; if(distTop<minTop)minTop=distTop;
      }
    }
    if(hasCenterPageNum)pageNumOkPages.push(p);
    pageTexts[p]=pageBuf.join('');
  }

  addRow('首頁篇名需置中且落在距上緣 2 公分內', firstHeaderOK,
    firstHeaderOK?'已於首頁偵測到置中之篇名/標題':'未偵測到置中篇名或位置不在 2 公分內',
    '請將篇名置中於頁首，且位置落在距上緣 2 公分以內。');

  addRow('首頁頁碼需置中且落在距下緣 2 公分內', firstFooterOK,
    firstFooterOK?'已於首頁偵測到置中頁碼':'未偵測到置中頁碼或位置未落在距下緣 2 公分內',
    '請將頁碼置中於頁尾，並保持距下緣 2 公分以內。');

  const headerTitleGuess=(()=>{const n=s=>(s||'').replace(/\s/g,'').toLowerCase();const filtered=headerTitleTexts.filter(s=>n(s).length>=4);const freq=new Map();for(const s of filtered){const k=n(s);freq.set(k,(freq.get(k)||0)+1)}let bestKey='',bestCount=0;for(const [k,c] of freq){if(c>bestCount){bestKey=k;bestCount=c}}const raw=filtered.find(s=>n(s)===bestKey)||'';return{norm:bestKey,raw}})();
  const fileNorm=(fileBase||'').replace(/\s/g,'').toLowerCase(), titleNorm=headerTitleGuess.norm;
  const titleDetected=!!titleNorm, nameTitleConsistent = titleDetected && (titleNorm.includes(fileNorm)||fileNorm.includes(titleNorm));
  addRow('檔案名稱與頁首篇名一致（不符判定淘汰）',nameTitleConsistent,
    titleDetected?`檔名「${fileBase}」；頁首篇名「${headerTitleGuess.raw}」`:'未偵測到頁首篇名',
    '請使檔名與頁首篇名一致。');

  const pageNumPass=(pageNumOkPages.length===_lastPageCount);
  addRow('每頁需有置中頁碼（底端帶）',pageNumPass,
    pageNumPass?'所有頁底中央皆有頁碼':`缺少頁碼頁次：${missingRange(pageNumOkPages,_lastPageCount)}`,
    '請在頁尾置中放置阿拉伯數字頁碼。');

  const MPT=MARGIN_PT.toFixed(1);
  const marginPass=(isFinite(minLeft)&&minLeft>=MARGIN_PT)&&(isFinite(minRight)&&minRight>=MARGIN_PT)&&(isFinite(minBottom)&&minBottom>=MARGIN_PT)&&(isFinite(minTop)&&minTop>=MARGIN_PT);
  addRow('四周邊界 ≥ 2 公分（正文區推估）',marginPass,
    `估計最小邊界（pt）：左 ${isFinite(minLeft)?minLeft.toFixed(1):'—'}、右 ${isFinite(minRight)?minRight.toFixed(1):'—'}、上 ${isFinite(minTop)?minTop.toFixed(1):'—'}、下 ${isFinite(minBottom)?minBottom.toFixed(1):'—'}（門檻≈${MPT}）`,
    '請確認版面設定至少 2 公分。');

  const fullText = pageTexts.slice(1).join('\n\n[[PAGE_SPLIT]]\n\n');
  const firstText=pageTexts[1]||''; const coverHits=['封面','學校','指導老師','學生','班級','學號'].filter(k=>firstText.includes(k));
  addRow('不得有封面頁（4–10頁）',coverHits.length===0, coverHits.length?`第1頁疑似封面關鍵字：${coverHits.join('、')}`:'未發現','4–10頁之小論文不應包含封面資訊。');

  runTextRules(pageTexts);
}

/* ===== DOCX ===== */
async function runDocxChecks(buf){
  const {value:html}=await window.mammoth.convertToHtml({arrayBuffer:buf});
  const text=stripHtml(html); _lastPageCount=0;
  addRow('頁數 4–10（DOCX）',true,'DOCX 無法精確估頁，建議匯出 PDF 再檢查空間規則。','請匯出 PDF。');
  runTextRules([text]);
}
function stripHtml(html){const t=document.createElement('div');t.innerHTML=html;return t.textContent||t.innerText||''}

/* ===== 參考文獻工具（僅小幅修正） ===== */

// 取「最後一個」參考文獻標題位置
function _findRefStart(fullText) {
  const text = fullText || '';
  const pat = /(^|\n)[^\n]{0,6}\s*(?:陸[、，.]?\s*)?(參考文獻|參考資料|References)\s*$/gmi;
  let m, lastIdx = -1;
  while ((m = pat.exec(text)) !== null) lastIdx = m.index + (m[1] ? m[1].length : 0);
  if (lastIdx < 0) {
    const cands = ['參考文獻','參考資料','References']
      .map(t => text.lastIndexOf(t)).filter(i => i >= 0);
    if (cands.length) lastIdx = Math.max(...cands);
  }
  return lastIdx;
}

// 小標偵測：類別標題（不算一條文獻）
function _isCategoryHeader(line) {
  const s = (line || '').replace(/^[\s\u3000]+|[\s\u3000]+$/g, '');
  if (/^(陸[、，.]?)?\s*(參考文獻|參考資料|References)\s*$/i.test(s)) return true;
  if (/^([（(]?\s*[一二三四五六七八九十\d]+\s*[)）]?\s*[\.．、]?)\s*[\u4e00-\u9fa5A-Za-z]{0,12}(書籍|專書|期刊|期刊論文|論文|博碩士|會議|報紙|電子報|網站|網路|網路相關資源|資料|文獻|法規|政府|機構)\s*(類|資料|資源)?\s*$/.test(s)) return true;
  if (/^[【〈《]\s*[\u4e00-\u9fa5A-Za-z]{1,12}\s*(類|資料|資源)?\s*[】〉》]$/.test(s)) return true;
  return false;
}

// 參考文獻前處理（新增短網址換行合併）
function _preCleanRefText(s) {
  return (s || '')
    .replace(/\[\[PAGE_SPLIT\]\]/g, '\n')
    .replace(/^\s*\d{1,4}\s*$/gm, '')
    .replace(/([A-Za-z0-9])-\s*\n\s*([A-Za-z0-9])/g, '$1$2')
    .replace(/\n\s*(https?:\/\/\S+)/g, ' $1')
    .replace(/\n\s*(www\.\S+)/g, ' $1')
    .replace(/\n\s*(doi\.org\/\S+)/gi, ' $1')
    .replace(/\n\s*((?:reurl\.cc|bit\.ly|tinyurl\.com|t\.ly|t\.co|goo\.gl|is\.gd|ow\.ly|lihi\d?\.cc|ppt\.cc|forms\.gle|youtu\.be|github\.io|medium\.com|shorturl\.at|url\.cn)\S*)/gi, ' $1') /* 新增 */
    .replace(/([一二三四五六七八九十百千])\s+([一二三四五六七八九十百千])/g, '$1$2')
    .replace(/(?!^)([　\s]*)(?=[（(]?\s*(\d+|[一二三四五六七八九十百千]{1,4})\s*[)）]?\s*[\.．、])/g, '\n')
    .replace(/([^\.;。！？）”」])\n\s*([^\n])/g, '$1 $2')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{2,}/g, '\n')
    .trim();
}

// 文獻切條
function _parseReferenceEntries(refText) {
  const lines = (refText || '').split(/\n+/).map(s => s.trim()).filter(Boolean);
  const entries = [];
  let cur = '';

  const SHORT_DOMAINS = /(reurl\.cc|bit\.ly|tinyurl\.com|t\.ly|t\.co|goo\.gl|is\.gd|ow\.ly|lihi\d?\.cc|ppt\.cc|forms\.gle|youtu\.be|github\.io|medium\.com|shorturl\.at|url\.cn)/i;

  const isNewItemByIndex = s =>
    /^[\s\u3000]*\d+\s*[\.．、]/.test(s) ||
    /^[\s\u3000]*[一二三四五六七八九十百千]{1,4}\s*[\.．、]/.test(s) ||
    /^[\s\u3000]*[（(]\s*(\d+|[一二三四五六七八九十百千]{1,4})\s*[)）]\s*[\.．、]/.test(s);

  const isNewItemByAuthorYear = s =>
    /^[\s\u3000]*[\u4e00-\u9fa5]{1,12}.*（(\d{4}|無日期)）/.test(s) ||
    /^[\s\u3000]*[A-Za-z][A-Za-z .-]{0,30}\(\s*(\d{4}|n\.?d\.?)\s*\)/.test(s);

  const isContinuation = s =>
    /^\s{2,}/.test(s) || /^[\u3000]+/.test(s) ||
    /^(https?:\/\/|www\.|doi\.org|doi:|取自|Available at|網址)/i.test(s) ||
    /^[-–•．·●]/.test(s) ||
    /^([a-z0-9-]+\.)+[a-z]{2,}(?:[\/?#:].*)?$/i.test(s) ||
    /^\/[^\s]/.test(s) ||
    SHORT_DOMAINS.test(s);

  for (const raw of lines) {
    const line0 = raw;
    if (_isCategoryHeader(line0)) continue;

    const looksNew = isNewItemByIndex(line0) || isNewItemByAuthorYear(line0);
    const cont     = isContinuation(line0);

    const base = line0
      .replace(/^[\s\u3000]*[（(]?\s*(\d+|[一二三四五六七八九十百千]{1,4})\s*[)）]?\s*[\.．、]\s*/, '')
      .trim();

    if (!cur) {
      if (cont && !looksNew) continue;
      cur = base;
      continue;
    }

    if (looksNew) { entries.push(cur.trim()); cur = base; continue; }

    if (cont) { cur += ' ' + line0.trim(); }
    else { cur += ' ' + base; }
  }

  if (cur) entries.push(cur.trim());
  return entries.filter(s => s.replace(/\s/g, '').length >= 6);
}

function _normRefKey(s) {
  return s.replace(/\s+/g, ' ')
    .replace(/https?:\/\/(www\.)?/gi, '')
    .replace(/[，、；。]+/g, ' ')
    .trim()
    .toLowerCase();
}
function _dedupReferences(list) {
  const map = new Map();
  for (const e of list) { const k = _normRefKey(e); if (!map.has(k)) map.set(k, e); }
  return Array.from(map.values());
}
function _hasYear(text) {
  return /（\d{4}|無日期）|\(\d{4}\)|\b\d{4}\b年|\b(19|20)\d{2}\b|\(n\.?d\.?\)/i.test(text);
}
function _textWithoutUrls(s) { return s.replace(/https?:\/\/\S+/g, '').replace(/\s+/g, ' ').trim(); }
function _isQualifiedRef(item) {
  const s = item;
  const hasPublisher = /(出版社|Press\b|Publishing|出版|台北市|新北市|高雄市).{0,10}$/.test(s);
  const hasJournalMeta =
    /(\bVol\.?\s*\d+|\bNo\.?\s*\d+|\bpp?\.?\s*\d+[-–]\d+|第\s*\d+\s*卷|第\s*\d+\s*期|頁\s*\d+[-–]\d+)/i.test(s) ||
    /(Journal|期刊|學報|通訊|評論|會刊).{0,40}\d{4}/.test(s) ||
    /,\s*\d{1,4}\s*,\s*\d{1,4}\s*[-–]\s*\d{1,4}/.test(s) ||
    /,\s*\d{1,4}\b(?!\s*頁)/.test(s);
  const hasThesisOrConf = /(碩士論文|博士論文|學位論文|Master'?s|Ph\.?D\.?|dissertation|thesis|Proceedings|研討會|會議論文)/i.test(s);
  const hasLawOrGov     = /(法規|條例|辦法|函釋|教育部|衛福部|主計總處|行政院|內政部|各縣市政府)/.test(s);
  const hasNewspaper    = /(日報|電子報|聯合報|自由時報|中國時報|經濟日報|蘋果日報|天下雜誌|今周刊)/.test(s);
  const hasDOI          = /(doi\.org|doi:)/i.test(s);
  const hasAcademicPlatform =
    /(ScienceDirect|Wiley|Springer|Elsevier|Taylor & Francis|SAGE|ACM Digital Library|IEEE Xplore|Nature|Cell Press|Oxford Academic|Cambridge Core)/i.test(s);

  if (hasPublisher || hasJournalMeta || hasThesisOrConf || hasLawOrGov || hasNewspaper || hasDOI || hasAcademicPlatform) return true;
  if (_hasYear(s) && _textWithoutUrls(s).replace(/[，、；。.,]/g, '').length >= 6) return true;
  return false;
}
function _classifyReference(item) {
  const s = item;
  const hasPublisher = /(出版社|Press\b|Publishing|出版|台北市|新北市|高雄市).{0,10}$/.test(s);
  const hasJournalMeta =
    /(\bVol\.?\s*\d+|\bNo\.?\s*\d+|\bpp?\.?\s*\d+[-–]\d+)/i.test(s) ||
    /(Journal|期刊|學報|通訊|評論|會刊).{0,40}\d{4}/.test(s) ||
    /,\s*\d{1,4}\s*,\s*\d{1,4}\s*[-–]\s*\d{1,4}/.test(s) ||
    /,\s*\d{1,4}\b(?!\s*頁)/.test(s);
  const hasThesisOrConf = /(碩士論文|博士論文|學位論文|Master'?s|Ph\.?D\.?|dissertation|thesis|Proceedings|研討會|會議論文)/i.test(s);
  const hasLawOrGov     = /(法規|條例|辦法|函釋|教育部|衛福部|主計總處|行政院|內政部|各縣市政府)/.test(s);
  const hasNewspaper    = /(日報|電子報|聯合報|自由時報|中國時報|經濟日報|蘋果日報|天下雜誌|今周刊)/.test(s);
  const hasDOI          = /(doi\.org|doi:)/i.test(s);
  const hasAcademicPlatform =
    /(ScienceDirect|Wiley|Springer|Elsevier|Taylor & Francis|SAGE|ACM Digital Library|IEEE Xplore|Nature|Cell Press|Oxford Academic|Cambridge Core)/i.test(s);
  if (hasPublisher || hasJournalMeta || hasThesisOrConf || hasLawOrGov || hasNewspaper || hasDOI || hasAcademicPlatform) return 'nonweb';
  const isPureWeb = /(YouTube|Facebook|Instagram|IG|X\.com|Twitter|PTT|Dcard|blogspot|wordpress|medium\.com|wiki|維基|論壇|部落格)/i.test(s) || /(網站|網頁|線上文章|粉專)/.test(s);
  const hasLink = /(https?:\/\/|取自|Available at|網址)/i.test(s);
  if (isPureWeb) return 'web';
  if (hasLink)    return 'web';
  return 'nonweb';
}

/* ===== 圖表檢查升級：工具常數與函式（新增） ===== */

// 支援中文數字標號（圖一、表二…）
const CN_NUM = '[一二三四五六七八九十百千]+';   // 中文數字樣式
const AR_NUM = '\\d+';                            // 阿拉伯數字
const FIGTAB_LABEL = new RegExp(
  '(^|\\n|\\s)(圖|表)\\s*(' + AR_NUM + '|' + CN_NUM + ')[\\s．\\.、：:]', 'g'
);

// 來源詞彙（擴充）
const SOURCE_LABEL = /(資料來源[:：]|圖片來源[:：]|出處[:：]|來源[:：])/;

// 圖/表附近視窗（跨頁容錯拉大）
const WINDOW_BEFORE = 0;       // 嚴格只看下方
const WINDOW_AFTER  = 400;     // 向後 400 字
const WINDOW_EXTRA  = 200;     // 再向後延伸（避免換行/跨頁）

// 自製/自繪等自我來源（免比對參考文獻）
const SELF_MADE = /(作者自(製|繪|攝|整理)|自製|自行(繪製|整理|攝影)|研究者自製|自行製作)/;

/** 取出「就近的來源行」：只在圖/表 *下方* 視窗內找來源詞彙 */
function _extractNearbySource(fullText, startIdx) {
  const from = Math.max(0, startIdx - WINDOW_BEFORE);
  const to   = Math.min(fullText.length, startIdx + WINDOW_AFTER + WINDOW_EXTRA);
  const slice = fullText.slice(from, to);

  // 只取「圖表之後」的內容
  const after = slice.slice(Math.min(slice.length, startIdx - from));

  // 找第一個來源詞彙
  const m = after.match(SOURCE_LABEL);
  if (!m) return null;

  // 來源詞彙起點在 after 的哪裡？
  const pos = after.search(SOURCE_LABEL);
  if (pos < 0) return null;

  // 抓該行到行尾（或下一個斷行）
  const tail = after.slice(pos);
  const line = tail.split(/\r?\n/)[0].trim();

  // 拿掉標籤（資料來源：）保留內容
  const content = line.replace(SOURCE_LABEL, '').trim();

  return {
    label: line,              // 含「資料來源：」整行
    content,                  // 去掉標籤的內容
    hasSelfMade: SELF_MADE.test(line)
  };
}

/** 檢查「就近片段」是否存在作者–年份制（視為內嵌來源） */
function _hasInlineCiteNearby(fullText, startIdx) {
  const from = Math.max(0, startIdx - 80);
  const to   = Math.min(fullText.length, startIdx + 200);
  const slice = fullText.slice(from, to);
  // （王小明，2023）或 王小明（2023）或 無日期/n.d.
  const INLINE = /(（[^，\n]{1,12}，\s*(\d{4}|無日期)）|[^\s，、]{1,12}（\s*(\d{4}|n\.?d\.?)）)/;
  return INLINE.test(slice);
}

/** 判斷來源行看起來像「完整參考文獻」 */
function _looksLikeFullReference(line) {
  const s = line || '';
  const hasYear = _hasYear(s);
  const hasMeta = /(出版社|Publishing|Press|第\s*\d+\s*卷|第\s*\d+\s*期|頁\s*\d+|Vol\.?|No\.?|pp?\.?|doi\.org|https?:\/\/)/i.test(s);
  return hasYear && hasMeta;
}

/** 嘗試把來源行（已去除「資料來源：」）對應到「陸、參考文獻」清單 */
function _matchSourceToRefs(srcContent, refEntries) {
  if (!srcContent) return { matched: false, hit: null };
  const skey = _normRefKey(srcContent);
  for (const r of refEntries) {
    const rkey = _normRefKey(r);
    // 互為包含即可視為對上（避免極嚴格完全一致）
    if (skey.includes(rkey) || rkey.includes(skey)) {
      return { matched: true, hit: r };
    }
  }
  return { matched: false, hit: null };
}

/* ===== 文字規則（只有第11、12項含微調） ===== */

/* ===== 升級工具：參考文獻解析/分類 + 圖表來源嚴格比對 ===== */
function _parseReferenceEntries(fullText) {
  const refIdx = (typeof _findRefStart==='function') ? _findRefStart(fullText) : 0;
  const refTextRaw = refIdx >= 0 ? fullText.slice(refIdx).replace(/^(?:[（(]?(六|陸|6|VI|Ⅵ|第\s*六)[)）]?\s*[\.．、]?\s*)?(參考文獻|參考資料|References?)[:：]?\s*/i, '') : '';
  const text = refTextRaw.replace(/\u3000/g,' ').trim();
  const lines = text.split(/\r?\n/);
  const refs = [];
  let cur = '';
  const startRe = /^(?:[（(]?[一二三四五六七八九十]+[)）]?[、.．]?\s*|\d+[、.．]\s*|[A-Za-z一-龥].{0,20}?\(\d{4}/;
  for (let raw of lines) {
    const line = raw.trim();
    if (!line) continue;
    if (startRe.test(line)) {
      if (cur) refs.push(cur.trim());
      cur = line;
    } else if (/https?:\/\/|doi\.org|頁\s*\d+|pp?\./i.test(line)) {
      cur += ' ' + line;
    } else {
      cur += ' ' + line;
    }
  }
  if (cur) refs.push(cur.trim());
  const uniq = [];
  const seen = new Set();
  for (const r of refs) {
    const k = r.replace(/\s+/g,' ');
    if (!seen.has(k)) { seen.add(k); uniq.push(r); }
  }
  return uniq;
}

function _splitRefByType(refs) {
  const book = [], web = [];
  for (const r of refs) {
    if (/https?:\/\/|doi\.org/i.test(r) || /(網站|Blog|Pixnet|網址|Facebook|YouTube)/i.test(r)) web.push(r);
    else if (/(出版社|有限公司|Press|Publishing|學報|期刊)/.test(r)) book.push(r);
    else (r.length > 60 ? book : web).push(r);
  }
  return {book, web};
}

function _matchSourceToRefs(srcContent, refEntries) {
  if (!srcContent) return { matched: false, hit: null };
  const src = srcContent.replace(/^(資料來源|圖片來源)[:：]?/,'').trim();
  const y = (src.match(/(\d{4})/)||[])[1] || null;
  const key = (typeof _normRefKey==='function') ? _normRefKey(src) : src.toLowerCase();
  for (const r of refEntries) {
    const ry = (r.match(/(\d{4})/)||[])[1] || null;
    const rkey = (typeof _normRefKey==='function') ? _normRefKey(r) : r.toLowerCase();
    if (y && ry && y === ry && (rkey.includes(key) || key.includes(rkey))) {
      return { matched: true, hit: r };
    }
  }
  return { matched: false, hit: null };
}

/* 來源詞彙擴充（供第13項使用） */
const SOURCE_LABEL = /(資料來源[:：]|圖片來源[:：]|出處[:：]|來源[:：])/;
const CN_NUM = '[一二三四五六七八九十百千]+';
const AR_NUM = '\\d+';
const FIGTAB_LABEL = new RegExp('(^|\\n|\\s)(圖|表)\\s*(' + AR_NUM + '|' + CN_NUM + ')[\\s．\\.、：:]', 'g');
const SELF_MADE = /(作者自(製|繪|攝|整理)|自製|自行(繪製|整理|攝影)|研究者自製|自行製作)/;

function runTextRules(pageTexts){
  const pages=pageTexts.length-1||1;
  const fullText = pageTexts.slice(1).join('\n\n[[PAGE_SPLIT]]\n\n');

  // 六段落
  const sec=RULES.sixSections;
  const idx=sec.map(s=>fullText.indexOf(s)); const hasAll=idx.every(i=>i>=0); const inOrder=hasAll && idx.join(',')===idx.slice().sort((a,b)=>a-b).join(',');
  addRow('六大段落依序',hasAll&&inOrder,hasAll?(inOrder?'順序正確':'找到各段落但順序錯誤'):'有段落缺失','請補齊並按規定順序排列：'+sec.join(' → '));

  // 作者-年份
  const citeA=/（[^，\n]{1,12}，\s*\d{4}）/g, citeB=/[^\s，。、]{1,12}（\s*\d{4}）/g;
  const hasCite=(fullText.match(citeA)?.length||0)+(fullText.match(citeB)?.length||0);
  addRow('內文引註採 作者-年份',hasCite>0,hasCite?`檢出 ${hasCite} 處`:'未檢出引註樣式','請使用（姓名，年份）或 姓名（年份）。');

  // 直接引文長度
  const quotesRaw=[...(fullText.match(/「[^」]{1,200}」/g)||[]),...(fullText.match(/“[^”]{1,200}”/g)||[])];
  const clean=s=>s.replace(PUNCT,''); const isDirect=q=>clean(q).length>=RULES.minQuoteCharsForDirect;
  const directQuotes=quotesRaw.filter(q=>isDirect(q)), over=directQuotes.filter(q=>clean(q).length>RULES.quoteMaxChars);
  addRow(`直接引文 ≤ ${RULES.quoteMaxChars} 字（不含標點與空白）`,over.length===0,over.length?`超標處：${over.length} 例`:(directQuotes.length?`OK（偵測 ${directQuotes.length} 例）`:'OK（未偵測引號片段）'),'引號內淨字數≥10才視為直接引文。');

  // 參考文獻錨點（最後一次）
  const refIdx = _findRefStart(fullText);

  // 8) 不得含校名/姓名 —— ★修正：跨頁時只檢查錨點前半段
  const schoolHits=[]; const schoolPages=new Set();
  for(let p=1;p<=pages;p++){
    let t=pageTexts[p]||'';
    const pageStartOffset = fullText.indexOf(pageTexts[p]);
    if(refIdx>=0 && pageStartOffset>=refIdx) continue;
    if(refIdx>=0 && pageStartOffset < refIdx && (pageStartOffset + t.length) > refIdx){
      t = t.slice(0, refIdx - pageStartOffset);
    }
    for(const re of SCHOOL_PATTERNS){
      const m=t.match(re);
      if(m&&m.length){schoolHits.push(...m); schoolPages.add(p);}
    }
  }
  addRow('不得含校名/姓名',schoolHits.length===0,
    schoolHits.length?`偵測到疑似校名：${[...new Set(schoolHits)].join('、')}｜頁次：${[...schoolPages].join(', ')}`:'未發現',
    '請去識別化（以 XXX 取代完整校名）。');

  // 12) 參考文獻切條 + 四層統計 —— ★修正：移除標題行
  const refTextRaw   = refIdx >= 0 ? fullText.slice(refIdx).replace(/^(?:陸[、，.]?\s*)?(?:參考文獻|參考資料|References)\s*[\r\n]+/i, '') : '';
  const refTextFull  = _preCleanRefText(refTextRaw);
  let refEntries     = refIdx >= 0 ? _parseReferenceEntries(refTextFull) : [];
  refEntries         = _dedupReferences(refEntries);

  const A_total = refEntries.length;
  const B_hasYear = refEntries.filter(s => _hasYear(s)).length;
  const C_hasLink = refEntries.filter(s => /(https?:\/\/|doi\.org|doi:|取自|Available at|網址)/i.test(s)).length;

  const classified = refEntries.map(s => ({ text: s, kind: _classifyReference(s) }));
  const nonwebCount = classified.filter(x => x.kind === 'nonweb').length;
  const webCount    = classified.filter(x => x.kind === 'web').length;

  const qualifiedCount   = refEntries.filter(_isQualifiedRef).length;
  const unqualifiedCount = Math.max(0, A_total - qualifiedCount);

  const evidence13 = [
     `檢查總數量：${A_total}`,
     `合格數量：${qualifiedCount}`,
     `不合格數量：${unqualifiedCount}`,
     `網路來源數量：${webCount}`,
     `（A.總數：${A_total}｜B.含年份：${B_hasYear}｜C.含網址/DOI：${C_hasLink}）`
  ].join('｜');
  
(() => {
  const refEntries = _parseReferenceEntries(fullText);
  const { book, web } = _splitRefByType(refEntries);
  const totalRef = refEntries.length;
  const pass = totalRef >= 3;
  const examples = refEntries.slice(0,3).map((r,i)=>`(${i+1}) ${r.slice(0,60)}…`).join('<br/>');
  addRow(
    '參考文獻 ≥ 3',
    pass,
    `檢出文獻：${totalRef} 筆（書籍 ${book.length}｜網路 ${web.length}）<br/>` + examples,
    pass ? 'OK' : '參考文獻少於 3 筆，請補足。'
  );
})();
// 圖表
  /* ===== 圖表需有標號/標題與來源 —— 升級版（中文數字、內嵌引注、對位比對、跨頁容錯） ===== */
(() => {
  const entries = []; // 每一個圖/表的檢查結果

  // 逐一找「圖/表 + 編號」（支援阿拉伯數字與中文數字）
  let m;
  while ((m = FIGTAB_LABEL.exec(fullText)) !== null) {
    const whole = m[0];
    const kind  = m[2];        // 圖 or 表
    const idx   = m.index;     // 在全文中的位置

    // 1) 來源行（只取下方就近的第一個來源標籤）
    const near = _extractNearbySource(fullText, idx);  // { label, content, hasSelfMade } | null

    // 2) 內嵌引注（就近片段內出現 (作者, 年份) 或 作者(年份)）
    const hasInline = _hasInlineCiteNearby(fullText, idx);

    // 3) 判斷是否視為「已提供來源」
    const hasSourceLine = !!near;
    const isSelfMade    = !!(near && near.hasSelfMade);
    // 表格：沿用舊有 tableInlineOK 概念；圖表新增 figureInlineOK
    const inlineOK      = hasInline; // 無論圖或表，只要偵測到就當作 OK

    // 4) 若來源行看起來像完整參考文獻，嘗試對到「陸、參考文獻」
    let looksFullRef = false, refMatched = false, refHit = null;
    if (hasSourceLine && !isSelfMade) {
      looksFullRef = _looksLikeFullReference(near.content);
      if (looksFullRef) {
        const mres = _matchSourceToRefs(near.content, refEntries);
        refMatched = mres.matched; refHit = mres.hit;
      }
    }

    entries.push({
      kind,
      label: whole.trim(),
      index: idx,
      hasSourceLine,
      sourceLine: near?.label || '',
      sourceContent: near?.content || '',
      isSelfMade,
      inlineOK,
      looksFullRef,
      refMatched,
      refHit
    });
  }

  const totalFT = entries.length;

  // 「來源是否足夠」的新判定：逐一對位
  // 規則：每一個圖/表只要具備「來源行」或「內嵌引注」或「自製」三者之一，即視為這個項目有來源。
  const withAnySource = entries.filter(it => it.hasSourceLine || it.inlineOK || it.isSelfMade).length;
  const lackSource = totalFT > 0 && (withAnySource < totalFT);

  // 缺標號的新判定：有「來源」但完全找不到圖/表標號（極端情況）
  // → 若 totalFT === 0 但全文仍出現來源詞彙，則視為「缺標號」
  const anySourceTokens = (fullText.match(SOURCE_LABEL) || []).length > 0;
  const lackLabel = (totalFT === 0 && anySourceTokens);

  // 新版通過條件
  const figPass = totalFT === 0 ? true : (!lackSource && !lackLabel);

  
  // 統計細節
  const tableInlineOK = entries.filter(it => it.kind === '表' && it.inlineOK).length;
  const figureInlineOK = entries.filter(it => it.kind === '圖' && it.inlineOK).length;
  const sourceLineCount = entries.filter(it => it.hasSourceLine).length;
  const selfMadeCount   = entries.filter(it => it.isSelfMade).length;
  const fullRefCount    = entries.filter(it => it.looksFullRef).length;
  const fullRefMatched  = entries.filter(it => it.looksFullRef && it.refMatched).length;
  const fullRefUnmatch  = entries.filter(it => it.looksFullRef && !it.refMatched).length;

  // 新增：分開統計「圖」與「表」，以及「表內圖片引注」
  const figureCount = entries.filter(it => it.kind === '圖').length;
  const tableCount  = entries.filter(it => it.kind === '表').length;
  const tableImageInline = entries.filter(it => it.kind === '表' && it.hasSourceLine && /圖片來源[:：]/.test(it.sourceLine)).length;


  // 準備證據字串（維持你原頁面風格）
  const parts = [
    `檢出圖/表：${totalFT} 項`,
    `就近來源行：${sourceLineCount} 項`,
    `表內已引注：${tableInlineOK} 項`,
    `圖內已引注：${figureInlineOK} 項`,
    `自製/自繪：${selfMadeCount} 項`,
    `判定缺來源：${lackSource ? '是' : '否'}`,
    `判定缺標號：${lackLabel ? '是' : '否'}`
  ];

  // 若有完整參考文獻的來源行，再補充比對狀態
  if (fullRefCount > 0) {
    parts.push(`來源含完整文獻：${fullRefCount} 項`);
    parts.push(`其中已對到「陸、參考文獻」：${fullRefMatched} 項`);
    parts.push(`未對到者：${fullRefUnmatch} 項`);
  }

  // 若存在未對到的完整文獻來源，列舉前幾筆（避免過長）
  let suggest = '圖或表上方置左標號+標題；下方標示「資料來源：…」。若於圖/表內已（作者，年份）引注，可視為已提供來源。';
  if (fullRefUnmatch > 0) {
    const samples = entries
      .filter(it => it.looksFullRef && !it.refMatched)
      .slice(0, 3)
      .map(it => `「${it.sourceContent.slice(0, 60)}…」`)
      .join('；');
    suggest += ` 檢測到有完整參考文獻卻未出現在「陸、參考文獻」：${samples}。請把這些文獻補進「陸、參考文獻」。`;
  }

  addRow(
    '圖表需有標號/標題與來源',
    figPass,
    parts.join('｜') + '<br/>' +
      (`圖：${figureCount} 項｜表：${tableCount} 項`) + '<br/>' +
      (`表內圖片引注：${tableImageInline} 項`),
    suggest
  );})();

  // 一致性
  const refJoined = refEntries.join('\n');
  const citations=(fullText.match(/（[^，\n]{1,12}，\s*\d{4}）/g)||[]).map(s=>s.replace(/[（）\s]/g,'').replace(/，/g,''));
  let linked=0; for(const c of citations){ if(refJoined.includes(c)) linked++; }
  const unlinkedInText = Math.max(0,citations.length-linked);
  const refAuthors=(refJoined.match(/[\u4e00-\u9fa5A-Za-z]{1,12}（\d{4}/g)||[]).map(s=>s.replace(/（\d{4}.*/,'')); 
  const reverseUnlinked = refAuthors.filter(a=>!citations.some(c=>c.startsWith(a))).length;
  addRow('參考文獻與內文引註一致性',unlinkedInText===0 && reverseUnlinked<=1,
    `引註共 ${citations.length}；對應 ${linked}；未對應（內文→文獻）${unlinkedInText}；疑似未被引用（文獻→內文）${reverseUnlinked}`,
    '內文出現之引用需在參考文獻出現；未引用者不得列入參考文獻。');

  // 來源不可全為網路
  const nonAllWebPass = (A_total > 0) && (webCount < A_total);
  const evidence16 = [
    `分類 → 非網路 ${nonwebCount}｜純網路 ${webCount}｜總數 ${A_total}`,
    `（檢查總數量：${A_total}｜合格：${qualifiedCount}｜不合格：${unqualifiedCount}）`
  ].join('｜');
  addRow('參考文獻來源不可全為網路', nonAllWebPass, evidence16, '至少保留 1 筆非純網路來源（書籍/期刊/論文/報紙/會議/法規等）。');

  // 字體 mirror
  syncBasicFontRuleMirror();

  // 外部檢索助手
  const extEvidence=buildExternalSearchEvidence(fileBase||'小論文篇名');
  addRow('作品是否已於校外發表/得獎（需人工/GPT）',true,extEvidence,'使用下方「外部檢索助手」或「複製查核提示」至 AI 小助手進行確認。');
  setTimeout(()=>initExternalSearchAssist(fileBase||'小論文篇名'),0);
}

/* ===== 外部檢索助手（保留原樣） ===== */
function buildExternalSearchEvidence(baseTitle){
  const t=(baseTitle||'小論文篇名').trim(), mk=s=>encodeURIComponent(s);
  const items=[
    {label:'Google Scholar（繁中）',href:`https://scholar.google.com/scholar?hl=zh-TW&q=${mk(t)}`,q:`${t}`},
    {label:'Google：題名 + 小論文 得獎',href:`https://www.google.com/search?q=${mk(t+' 小論文 得獎')}`,q:`${t} 小論文 得獎`},
    {label:'Google：題名 + 專題製作 比賽 歷屆 作品',href:`https://www.google.com/search?q=${mk(t+' 專題製作 比賽 歷屆 作品')}`,q:`${t} 專題製作 比賽 歷屆 作品`},
    {label:'Google：題名 + site:shs.edu.tw',href:`https://www.google.com/search?q=${mk(t+' site:shs.edu.tw')}`,q:`${t} site:shs.edu.tw`},
    {label:'Google：題名 + site:edu.tw 小論文',href:`https://www.google.com/search?q=${mk(t+' site:edu.tw 小論文')}`,q:`${t} site:edu.tw 小論文`}
  ];
  const links=items.map((it,i)=>`<li class="mono">#${i+1} ${it.label}：<a data-extlink href="${it.href}" target="_blank" rel="noopener">${it.q}</a></li>`).join('');
  return `<div id="extSearchPanel" class="extsearch"><div class="row"><button id="btnOpenAllSearch">一鍵開啟所有搜尋</button><button id="btnCopyGptPrompt" class="primary">複製查核提示（給 AI小助手）</button></div><div class="hint" style="margin-top:6px">以下為自動生成的查詢（即將檢索的內容）：</div><ul style="margin-top:6px">${links}</ul></div>`;
}
function initExternalSearchAssist(baseTitle){
  const panel=document.getElementById('extSearchPanel'); if(!panel)return;
  const btnOpen=panel.querySelector('#btnOpenAllSearch'), btnCopy=panel.querySelector('#btnCopyGptPrompt');
  btnOpen?.addEventListener('click',()=>{panel.querySelectorAll('a[data-extlink]').forEach(a=>window.open(a.href,'_blank','noopener'));});
  btnCopy?.addEventListener('click',async()=>{
    const t=(baseTitle||'小論文篇名').trim();
    const qsEls=panel.querySelectorAll('a[data-extlink]'); const qs=Array.from(qsEls).map((a,i)=>`(${i+1}) ${a.textContent} -> ${a.href}`).join('\n');
    const summary=(typeof lastSummaryJSON==='object'&&lastSummaryJSON)?JSON.stringify(lastSummaryJSON,null,2):(document.getElementById('jsonPreview')?.textContent||'');
    const prompt=`請協助進行「校外發表／得獎」外部檢索確認：
題名/檔名：${t}

請到以下來源檢索並彙整前 10 筆結果（標題、來源、年份、連結），同時判斷是否疑似相同題名之公開發表或得獎紀錄：
1) Google Scholar（繁中）
2) Google 關鍵字：題名 +「小論文 得獎」、題名 +「專題製作 比賽 歷屆 作品」
3) Google 限站：site:shs.edu.tw、site:edu.tw 小論文

請用表格＋條列輸出；最後給「是否查到疑似相同題名」結論與理由，並附上實際使用的查詢字串清單。

（附：本次檢查摘要 JSON）
${summary||'(若空白可忽略)'}

（查詢字串與連結預覽）
${qs}`;
    try{await navigator.clipboard.writeText(prompt);alert('已複製查核提示，請點「打開AI助手」並貼上。');}catch{alert('複製失敗，請重試。')}
  });
}

/* ===== 工具 & 報告 ===== */
function missingRange(okPages,total){const ok=new Set(okPages);const miss=[];for(let i=1;i<=total;i++){if(!ok.has(i))miss.push(i)}return miss.length?miss.join(', '):'—'}
function addRow(rule,pass,evidence,suggestion){
  const tr=document.createElement('tr'); const badge=pass?'<span class="ok">通過</span>':'<span class="bad">未通過</span>';
  tr.innerHTML=`<td><span class="num">${++_ruleCounter}</span>${rule}</td><td>${badge}</td><td>${evidence||''}</td><td>${suggestion||''}</td>`; tbody.appendChild(tr);
  const map={
    '頁數 4–10':{rule_id:'file.pages',category:'檔案規格',severity:'critical'},
    '首頁篇名需置中且落在距上緣 2 公分內':{rule_id:'header.topmargin_firstpage',category:'頁首',severity:'critical'},
    '首頁頁碼需置中且落在距下緣 2 公分內':{rule_id:'footer.bottommargin_firstpage',category:'頁碼',severity:'critical'},
    '檔案名稱與頁首篇名一致（不符判定淘汰）':{rule_id:'filename.title_consistency',category:'命名',severity:'critical'},
    '每頁需有置中頁碼（底端帶）':{rule_id:'footer.pagenum',category:'頁碼',severity:'critical'},
    '四周邊界 ≥ 2 公分（正文區推估）':{rule_id:'layout.margins',category:'版面',severity:'major'},
    '不得有封面頁（4–10頁）':{rule_id:'file.coverpage',category:'檔案規格',severity:'critical'},
    '不得含校名/姓名':{rule_id:'privacy.no_school_or_name',category:'隱私',severity:'critical'},
    '六大段落依序':{rule_id:'structure.sections_order',category:'結構',severity:'critical'},
    '內文引註採 作者-年份':{rule_id:'references.citation_style',category:'參考文獻',severity:'major'},
    [`直接引文 ≤ ${RULES.quoteMaxChars} 字（不含標點與空白）`]:{rule_id:'quotes.limit',category:'引文',severity:'major'},
    '圖表需有標號/標題與來源':{rule_id:'figures.tables',category:'圖表',severity:'major'},
    '參考文獻 ≥ 3':{rule_id:'references.count',category:'參考文獻',severity:'critical'},
    '參考文獻與內文引註一致性':{rule_id:'references.consistency',category:'參考文獻',severity:'critical'},
    '參考文獻來源不可全為網路':{rule_id:'references.non_all_web',category:'參考文獻',severity:'critical'},
    '字體/字型規範（需人工確認）':{rule_id:'style.font',category:'版面',severity:'major'},
    '作品是否已於校外發表/得獎（需人工/GPT）':{rule_id:'external.duplicate_check',category:'人工檢查',severity:'minor'},
    '頁數 4–10（DOCX）':{rule_id:'file.pages.docx',category:'檔案規格',severity:'minor'},
    'Cloud Run 字體字型檢查':{rule_id:'fontcheck.cloudrun',category:'字體字型',severity:'major'}
  };
  const meta=map[rule]||{rule_id:'misc',category:'其他',severity:'minor'};
  _rowsBuffer.push({rule_id:meta.rule_id,category:meta.category,severity:meta.severity,pass,evidence:typeof evidence==='string'?{note:evidence}:(evidence||null),suggestion:suggestion||null});
}
function buildAndShowSummaryJSON({fileName='paper',pages=0,passFailRows=[]}={}){
  const weight=x=>x.severity==='critical'?2:(x.severity==='major'?1:0.5);
  const max=passFailRows.reduce((a,x)=>a+weight(x),0)||1; const score=Math.round(passFailRows.reduce((a,x)=>a+(x.pass?weight(x):0),0)/max*100);
  const hasCriticalFail=passFailRows.some(x=>!x.pass && x.severity==='critical'); const status=hasCriticalFail?'revise':'pass';
  const summary={meta:{report_id:(crypto.randomUUID?.()||String(Math.random()).slice(2)),generated_at:new Date().toISOString(),file:{name:fileName,pages,type:fileExt},rules_version:'v1.4'},
    summary:{score,totals:{checks:passFailRows.length,pass:passFailRows.filter(x=>x.pass).length,fail:passFailRows.filter(x=>!x.pass).length},status},
    checks:passFailRows.map(x=>({rule_id:x.rule_id,category:x.category,severity:x.severity,pass:x.pass,evidence:x.evidence||null,suggestion:x.suggestion||null}))};
  lastSummaryJSON=summary; jsonPreview.textContent=JSON.stringify(summary,null,2);
}
function syncBasicFontRuleMirror(){
  if(_fontCheckMirror){
    addRow('字體/字型規範（需人工確認）',_fontCheckMirror.pass, `（同步進階檢查）${_fontCheckMirror.evidence||''}`, _fontCheckMirror.suggestion||'');
  }else{
    addRow('字體/字型規範（需人工確認）',true,'建議：中文新細明體 12pt；英文 Times New Roman 12pt；黑色字體。','若需精確檢查請使用「字體字型進階檢查」。');
  }
}

/* ===== 下載/複製 ===== */
downloadHtmlBtn.addEventListener('click',()=>{
  const html='<!doctype html><meta charset="utf-8"><title>小論文檢查報告</title>'+document.querySelector('#resultCard').outerHTML;
  const blob=new Blob([html],{type:'text/html;charset=utf-8'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='paper-check-report.html';a.click();URL.revokeObjectURL(a.href);
});
copyBtn.addEventListener('click',async()=>{if(!lastSummaryJSON){alert('尚未產生檢查摘要。');return;}await navigator.clipboard.writeText(JSON.stringify(lastSummaryJSON,null,2));alert('已複製檢查摘要 JSON，前往 AI助手 後直接貼上即可。')});
downloadJsonBtn.addEventListener('click',()=>{if(!lastSummaryJSON){alert('尚未產生檢查摘要。');return;}const blob=new Blob([JSON.stringify(lastSummaryJSON,null,2)],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download((lastSummaryJSON.meta?.file?.name||'paper')+'.summary.json');a.click();URL.revokeObjectURL(a.href);});

/* ===== GAS 計數 ===== */
async function bumpCounterEveryTime(){
  try{
    await fetch(GAS_WEB_APP_URL,{method:'POST',headers:{'Content-Type':'application/json','Cache-Control':'no-cache'},body:JSON.stringify({action:'increment',site:'essayguard'})});
  }catch(e){}
}
async function fetchCounter(){
  try{
    const res=await fetch(GAS_WEB_APP_URL+(GAS_WEB_APP_URL.includes('?')?'&':'?')+'action=get&site=essayguard',{headers:{'Accept':'application/json','Cache-Control':'no-cache'}});
    const text=await res.text();let data;try{data=JSON.parse(text);}catch{data={};}
    if(typeof data.count==='number'){counterEl.textContent=data.count.toLocaleString();}else{counterEl.textContent='0';}
  }catch(e){counterEl.textContent='0';}
}
if(ENABLE_COUNTER)fetchCounter();

/* ===== 快取清理（避免 SW 舊檔） ===== */
(async () => {
  try {
    if ('serviceWorker' in navigator) {
      const regs = await navigator.serviceWorker.getRegistrations();
      for (const r of regs) { await r.unregister(); }
    }
    if (window.caches?.keys) {
      const keys = await caches.keys();
      for (const k of keys) { await caches.delete(k); }
    }
  } catch (e) {}
})();

console.log('小論文格式自檢系統 v2025-10-07a loaded.');
