/**
 * スプレッドシート側の「拡張機能 > Apps Script」に貼り付けてください。
 * 編集可能な元スプレッドシートの ID と、対象タブ名を必ず設定します。
 * 「ウェブアプリとしてデプロイ」: 実行ユーザー = 自分 / アクセス = 全員。
 * 公開POSTを許す仕組みのため、シートのバックアップと定期的な確認を推奨します。
 */
const SPREADSHEET_ID = '1YMACN6m-5tE3TSY4jNxVRdQieDCTwh7Zhted-z8zTl4';
const SHEET_NAME = '全曲'; // タブ名が分かれば入力。空欄なら「曲名」と「楽曲ID」があるシートを自動検出。

/** 新規投稿では原曲のChordWiki楽曲ページだけを受け付ける（Apps Script V8互換）。 */
function isOriginalChordWikiUrl_(value) {
  const input = String(value || '').trim();
  if (!input || input.length > 500 || /[\s\x00-\x1f\x7f]/.test(input)) return false;
  // クエリのない /wiki/個別ページ。タグ・検索・外部ドメインを許可しない。
  const direct = input.match(/^https:\/\/ja\.chordwiki\.org\/wiki\/([^/?#]+)$/i);
  if (direct) {
    try {
      const title = decodeURIComponent(direct[1].replace(/\+/g, ' '));
      return !!title.trim() && !/[\x00-\x1f\x7f/]/.test(title);
    } catch (error) { return false; }
  }
  // 原曲キーが省略か0の c=view URLは許可。それ以外の wiki.cgi は拒否。
  const cgi = input.match(/^https:\/\/ja\.chordwiki\.org\/wiki\.cgi\?([^#]+)$/i);
  if (!cgi) return false;
  const params = {};
  try {
    cgi[1].split('&').forEach(pair => {
      const i = pair.indexOf('=');
      const key = decodeURIComponent((i < 0 ? pair : pair.slice(0, i)).replace(/\+/g, ' '));
      const val = decodeURIComponent((i < 0 ? '' : pair.slice(i + 1)).replace(/\+/g, ' '));
      if (Object.prototype.hasOwnProperty.call(params, key)) throw Error('duplicate query');
      params[key] = val;
    });
  } catch (error) { return false; }
  return params.c === 'view' && !!String(params.t || '').trim() &&
    !/[\x00-\x1f\x7f]/.test(params.t) && (params.key === undefined || params.key === '0');
}

/** タブ名が未指定なら見出しから特定。曖昧な場合は書き込まない。 */
function getSongSheet_() {
  const book = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (SHEET_NAME) {
    const sheet = book.getSheetByName(SHEET_NAME);
    if (!sheet) throw Error('tab not found');
    return sheet;
  }
  const matches = book.getSheets().filter(sheet => {
    const rows = sheet.getDataRange().getDisplayValues();
    return rows.some(row => row.includes('曲名') && row.includes('楽曲ID'));
  });
  if (matches.length !== 1) throw Error('song sheet not uniquely identified: ' + matches.length);
  return matches[0];
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    // 一括登録のみGoogle管理者認証を要求。通常の1曲投稿はこれまでどおりゲスト利用可能。
    if (data.action === 'bulkImport') return bulkImport_(data);
    if (data.action === 'bulkMyListImport') return bulkMyListImport_(data);
    // 自動入力 bot 向けのハニーポット
    if (data.website) return reply_('rejected');
    const artist = clean_(data.artist, 150);
    const title = clean_(data.title, 150);
    const codeUrl = String(data.codeUrl || '').trim();
    const genre1 = clean_(data.genre1, 80);
    const genre2 = clean_(data.genre2, 80);
    const bpmValue = String(data.bpm || '').trim();
    if (!artist || !title || !genre1 || !isOriginalChordWikiUrl_(codeUrl)) throw Error('invalid input or unsupported ChordWiki URL');
    if (genre2 && genre1 === genre2) throw Error('duplicate genres');
    if (bpmValue && (!/^\d+$/.test(bpmValue) || +bpmValue < 1 || +bpmValue > 400)) throw Error('invalid bpm');

    lock.waitLock(10000);
    const sheet = getSongSheet_();
    if (!sheet) throw Error('tab not found');
    // index.html と同様に「曲名」を含む行をヘッダーとみなします。
    const values = sheet.getDataRange().getDisplayValues();
    const headerIndex = values.findIndex(row => row.includes('曲名'));
    if (headerIndex < 0) throw Error('header not found');
    const headers = values[headerIndex].map(s => String(s).trim());
    const names = {
      artist: ['アーティスト'], title: ['曲名'], codeUrl: ['コードwiki', 'コードWiki'],
      bpm: ['BPM'], genre1: ['ジャンル①', 'ジャンル'], genre2: ['ジャンル②']
    };
    const col = key => headers.findIndex(h => names[key].includes(h));
    if (Object.keys(names).some(key => col(key) < 0)) throw Error('required header missing');

    // 既存のジャンルと一致するものだけ受け付け、任意のジャンル注入を防ぎます。
    const allowed = new Set();
    values.slice(headerIndex + 1).forEach(row => {
      [col('genre1'), col('genre2')].forEach(i => {
        if (String(row[i] || '').trim()) allowed.add(String(row[i]).trim());
      });
    });
    if (!allowed.has(genre1) || (genre2 && !allowed.has(genre2))) throw Error('unknown genre');

    // 重複送信（同一アーティスト・曲名）を抑止。
    const normalized = x => String(x || '').trim().toLocaleLowerCase();
    if (values.slice(headerIndex + 1).some(row => normalized(row[col('artist')]) === normalized(artist) && normalized(row[col('title')]) === normalized(title))) {
      return reply_('duplicate');
    }
    const songIdColumn = headers.findIndex(h => h === '楽曲ID');
    if (songIdColumn < 0) throw Error('楽曲ID column missing');
    const row = new Array(sheet.getLastColumn()).fill('');
    row[songIdColumn] = Utilities.getUuid();
    row[col('artist')] = artist;
    row[col('title')] = title;
    row[col('codeUrl')] = codeUrl;
    row[col('bpm')] = bpmValue ? Number(bpmValue) : '';
    row[col('genre1')] = genre1;
    row[col('genre2')] = genre2;
    const nextRow = sheet.getLastRow() + 1;
    // 値として登録し、セル内の = から始まる入力が数式に変換されないよう保護します。
    const range = sheet.getRange(nextRow, 1, 1, row.length);
    range.setNumberFormat('@');
    range.setValues([row]);
    SpreadsheetApp.flush();
    return reply_('ok');
  } catch (err) {
    console.error(err);
    return reply_('error');
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function clean_(value, max) {
  return String(value == null ? '' : value).trim().slice(0, max);
}
function reply_(status) {
  return ContentService.createTextOutput(JSON.stringify({ status }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 一度だけ手動実行。空欄の楽曲IDにだけUUIDを発行。既存の値は変えません。 */
function assignMissingSongIds() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSongSheet_();
    if (!sheet) throw Error('tab not found');
    const values = sheet.getDataRange().getDisplayValues();
    const headerIndex = values.findIndex(row => row.includes('曲名'));
    if (headerIndex < 0) throw Error('header not found');
    const headers = values[headerIndex].map(h => String(h).trim());
    const idColumn = headers.indexOf('楽曲ID');
    const titleColumn = headers.indexOf('曲名');
    if (idColumn < 0) throw Error('楽曲ID column missing');
    const used = new Set();
    // 重複IDは自動変更しません。既存の個人データと紐づく可能性があるためエラーにします。
    values.slice(headerIndex + 1).forEach(row => {
      const id = String(row[idColumn] || '').trim();
      if (id && used.has(id)) throw Error('duplicate 楽曲ID: ' + id);
      if (id) used.add(id);
    });
    let added = 0;
    for (let i = headerIndex + 1; i < values.length; i++) {
      if (!String(values[i][titleColumn] || '').trim()) continue;
      if (!String(values[i][idColumn] || '').trim()) {
        let id;
        do { id = Utilities.getUuid(); } while (used.has(id));
        sheet.getRange(i + 1, idColumn + 1).setValue(id);
        used.add(id);
        added++;
      }
    }
    console.log('新規発行した楽曲ID: ' + added + '件');
  } finally {
    lock.releaseLock();
  }
}


// ===== CSV一括マスター登録（管理者専用） =====
// Apps Script > プロジェクトの設定 > スクリプト プロパティに以下を設定：
// BULK_ADMIN_UID = 管理者の Firebase Authentication UID（Googleログインユーザー）
// FIREBASE_WEB_API_KEY = Firebase プロジェクトの Web API Key（公開用のキー）
// 重要：管理者UIDや管理者パスワードを index.html に埋め込まないこと。
function bulkVerifyAdmin_(idToken) {
  const props = PropertiesService.getScriptProperties();
  const expected = String(props.getProperty('BULK_ADMIN_UID') || '').trim();
  const apiKey = String(props.getProperty('FIREBASE_WEB_API_KEY') || '').trim();
  const token = String(idToken || '');
  if (!expected || !apiKey || !token || token.length > 8192) return false;
  try {
    const result = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + encodeURIComponent(apiKey), {
      method:'post', contentType:'application/json', payload:JSON.stringify({idToken:token}),muteHttpExceptions:true
    });
    if (result.getResponseCode() !== 200) return false;
    const user = (JSON.parse(result.getContentText()).users || [])[0];
    return !!(user && !user.disabled && user.localId === expected &&
      (user.providerUserInfo || []).some(provider => provider.providerId === 'google.com'));
  } catch(err) { console.error('BULK_AUTH_ERROR',err); return false; }
}
function bulkNorm_(value) {
  return String(value == null ? '' : value).normalize('NFKC').trim().replace(/[\s\u3000]+/g,' ').toLocaleLowerCase();
}
function bulkUrl_(value) {
  const raw = String(value || '').trim();
  const direct = raw.match(/^https:\/\/ja\.chordwiki\.org\/wiki\/([^/?#]+)$/i);
  if (direct) {
    try {return 'chordwiki:' + bulkNorm_(decodeURIComponent(direct[1].replace(/\+/g,' ')));}catch(_){}
  }
  const cgi=raw.match(/^https:\/\/ja\.chordwiki\.org\/wiki\.cgi\?([^#]+)$/i);
  if(cgi){
    let title='';let command='';
    cgi[1].split('&').forEach(pair=>{
      const i=pair.indexOf('=');
      try {
        const name=decodeURIComponent((i<0?pair:pair.slice(0,i)).replace(/\+/g,' '));
        const val=decodeURIComponent((i<0?'':pair.slice(i+1)).replace(/\+/g,' '));
        if(name==='t')title=val;
        if(name==='c')command=val;
      }catch(_){}
    });
    if(command==='view'&&title)return 'chordwiki:'+bulkNorm_(title);
  }
  return raw.toLowerCase();
}
function bulkStatusKey_(id) { return 'bulk_csv_' + String(id || ''); }
function bulkSaveStatus_(id,value) {
  if (id) CacheService.getScriptCache().put(bulkStatusKey_(id),JSON.stringify(value),21600);
}
// GASの通常のPOSTはブラウザからCORSで応答を読めないので、一時的な処理結果だけJSONPで問い合わせる。
// requestIdは他人が推測できないランダムUUID。結果に楽曲名・個人情報を含めない。
function doGet(e) {
  const p=(e && e.parameter) || {};
  const callback=String(p.callback || '');
  const requestId=String(p.requestId || '');
  if (p.action!=='bulkStatus' || !/^[A-Za-z_$][\w$]{0,90}$/.test(callback) ||
    !/^[a-zA-Z0-9_-]{12,100}$/.test(requestId)) {
    return ContentService.createTextOutput('invalid request').setMimeType(ContentService.MimeType.TEXT);
  }
  const stored=CacheService.getScriptCache().get(bulkStatusKey_(requestId));
  const result=stored ? JSON.parse(stored) : {status:'pending'};
  return ContentService.createTextOutput(callback+'('+JSON.stringify(result)+');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function bulkImport_(data) {
  const requestId=String(data.requestId || '');
  if(!/^[a-zA-Z0-9_-]{12,100}$/.test(requestId)) return reply_('invalid_request');
  if(!bulkVerifyAdmin_(data.idToken)){
    bulkSaveStatus_(requestId,{status:'unauthorized'});
    return reply_('unauthorized');
  }
  const list=data.songs;
  if (!Array.isArray(list)||list.length<1||list.length>80) {
    bulkSaveStatus_(requestId,{status:'error',message:'1〜80曲ずつ送信してください。'});
    return reply_('invalid_batch');
  }
  const lock=LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const cached=CacheService.getScriptCache().get(bulkStatusKey_(requestId));
    if(cached) return reply_('already_processed');
    // 同一処理IDの再送を安全に停止（途中結果が不明なときも二重書き込みしない）。
    bulkSaveStatus_(requestId,{status:'processing'});
    const sheet=getSongSheet_();
    const raw=sheet.getDataRange().getDisplayValues();
    const hi=raw.findIndex(row=>row.includes('曲名') && row.includes('楽曲ID'));
    if(hi<0)throw Error('master header not found');
    const header=raw[hi].map(v=>String(v||'').trim());
    const aliases={
      title:['曲名'],artist:['アーティスト'],codeUrl:['コードwiki','コードWiki','コードWiki URL'],
      bpm:['BPM'],genre1:['ジャンル','ジャンル①'],genre2:['ジャンル②'],
      tieup:['タイアップ'],search:['検索'],lyrics:['作詞'],composition:['作曲'],arrangement:['編曲']
    };
    const col={};
    Object.keys(aliases).forEach(key=>col[key]=header.findIndex(h=>aliases[key].includes(h)));
    const idCol=header.indexOf('楽曲ID');
    if(['title','artist','codeUrl','genre1'].some(k=>col[k]<0)||idCol<0)throw Error('required master columns missing');
    const knownPairs=new Set(),knownUrls=new Set(),genreSet=new Set();
    raw.slice(hi+1).forEach(row=>{
      const name=bulkNorm_(row[col.title]);const artist=bulkNorm_(row[col.artist]);
      if(name&&artist)knownPairs.add(name+'\u001f'+artist);
      const original=bulkUrl_(row[col.codeUrl]);if(String(row[col.codeUrl]||'').trim())knownUrls.add(original);
      if(String(row[col.genre1]||'').trim())genreSet.add(String(row[col.genre1]).trim());
      if(col.genre2>=0 && String(row[col.genre2]||'').trim())genreSet.add(String(row[col.genre2]).trim());
    });
    let added=0,duplicates=0,invalid=0;
    const append=[];
    list.forEach(song=>{
      const title=clean_(song.title,150),artist=clean_(song.artist,150),url=String(song.codeUrl||'').trim();
      const genre1=clean_(song.genre1,80),genre2=clean_(song.genre2,80),bpm=String(song.bpm||'').trim();
      if(!title||!artist||!genre1||!isOriginalChordWikiUrl_(url)||
        (genre2&&genre1===genre2)||(bpm&&(!/^\d+$/.test(bpm)||+bpm<1||+bpm>400))||
        !genreSet.has(genre1)||(genre2&&!genreSet.has(genre2))){invalid++;return;}
      const pair=bulkNorm_(title)+'\u001f'+bulkNorm_(artist),urlKey=bulkUrl_(url);
      if(knownPairs.has(pair)||knownUrls.has(urlKey)){duplicates++;return;}
      const row=new Array(sheet.getLastColumn()).fill('');
      row[idCol]=Utilities.getUuid();
      const values={title,artist,codeUrl:url,bpm,genre1,genre2,
        tieup:clean_(song.tieup,250),search:clean_(song.search,800),
        lyrics:clean_(song.lyrics,150),composition:clean_(song.composition,150),arrangement:clean_(song.arrangement,150)};
      Object.keys(values).forEach(key=>{if(col[key]>=0)row[col[key]]=values[key];});
      append.push(row);added++;
      knownPairs.add(pair);knownUrls.add(urlKey);
    });
    if(append.length){
      const output=sheet.getRange(sheet.getLastRow()+1,1,append.length,sheet.getLastColumn());
      output.setNumberFormat('@');
      output.setValues(append);
      SpreadsheetApp.flush();
    }
    const result={status:'done',added,duplicates,invalid};
    bulkSaveStatus_(requestId,result);
    console.log('BULK_DONE',JSON.stringify(result));
    return reply_('done');
  }catch(err){
    console.error('BULK_FAILED',err);
    bulkSaveStatus_(requestId,{status:'error',message:'GAS実行ログを確認してください。'});
    return reply_('error');
  }finally{if(lock.hasLock())lock.releaseLock();}
}

// ===== 個人CSV → 原曲マスターの照合・追加 → 個人マイリスト用UUID/Key応答 =====
// Googleログインした本人のみ利用可能。Firebase個人データ自体への保存はブラウザから
// users/{auth.uid}/songs/{songUuid} へ行う（既存Firebaseルールが本人だけ許可）。
// 一括マスター登録の旧管理者用 bulkImport は従来どおり残す。
function bulkVerifyGoogleUser_(idToken, expectedUid) {
  const apiKey = String(PropertiesService.getScriptProperties().getProperty('FIREBASE_WEB_API_KEY') || '').trim();
  const token = String(idToken || ''), uid = String(expectedUid || '');
  if (!apiKey || !token || token.length > 8192 || !uid || uid.length > 150) return false;
  try {
    const response=UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key='+encodeURIComponent(apiKey),{
      method:'post',contentType:'application/json',payload:JSON.stringify({idToken:token}),muteHttpExceptions:true
    });
    if(response.getResponseCode()!==200)return false;
    const account=(JSON.parse(response.getContentText()).users||[])[0];
    return !!(account&&!account.disabled&&account.localId===uid&&
      (account.providerUserInfo||[]).some(p=>p.providerId==='google.com'));
  }catch(err){console.error('MYLIST_CSV_AUTH_ERROR',err);return false;}
}
// Source URL can contain key=N. Derive canonical master URL independently of the client.
function bulkParsePersonalUrl_(rawValue){
  const raw=String(rawValue||'').trim();
  if(!raw||raw.length>500||/[\s\x00-\x1f\x7f]/.test(raw))return null;
  let title='',key=0;
  const direct=raw.match(/^https:\/\/ja\.chordwiki\.org\/wiki\/([^/?#]+)$/i);
  if(direct){
    try{title=decodeURIComponent(direct[1].replace(/\+/g,' '));}catch(_){return null;}
  }else{
    const cgi=raw.match(/^https:\/\/ja\.chordwiki\.org\/wiki\.cgi\?([^#]+)$/i);
    if(!cgi)return null;
    const p={},pieces=cgi[1].split('&');
    try{
      pieces.forEach(piece=>{
        const idx=piece.indexOf('=');
        const k=decodeURIComponent((idx<0?piece:piece.slice(0,idx)).replace(/\+/g,' '));
        const v=decodeURIComponent((idx<0?'':piece.slice(idx+1)).replace(/\+/g,' '));
        if(Object.prototype.hasOwnProperty.call(p,k))throw Error('duplicate query');
        p[k]=v;
      });
    }catch(_){return null;}
    if(p.c!=='view'||!p.t)return null;
    title=p.t;
    if(p.key!==undefined&&p.key!==''){
      const n=String(p.key).trim();
      if(!/^[+-]?\d+$/.test(n))return null;
      key=Number(n);
    }
  }
  if(!title.trim()||/[\x00-\x1f\x7f/]/.test(title)||!Number.isInteger(key)||key< -5||key>6)return null;
  const originalUrl='https://ja.chordwiki.org/wiki/'+encodeURIComponent(title);
  if(!isOriginalChordWikiUrl_(originalUrl))return null;
  return {key,title,originalUrl,urlKey:'chordwiki:'+bulkNorm_(title)};
}
function bulkMyListImport_(data){
  const requestId=String(data.requestId||''),uid=String(data.userUid||'');
  if(!/^[a-zA-Z0-9_-]{12,100}$/.test(requestId))return reply_('invalid_request');
  if(!bulkVerifyGoogleUser_(data.idToken,uid)){
    bulkSaveStatus_(requestId,{status:'unauthorized'});return reply_('unauthorized');
  }
  const songs=data.songs;
  if(!Array.isArray(songs)||songs.length<1||songs.length>60){
    bulkSaveStatus_(requestId,{status:'error',message:'1〜60曲の範囲で送信してください。'});return reply_('invalid_batch');
  }
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(30000);
    if(CacheService.getScriptCache().get(bulkStatusKey_(requestId)))return reply_('already_processed');
    bulkSaveStatus_(requestId,{status:'processing'});
    const sheet=getSongSheet_(),raw=sheet.getDataRange().getDisplayValues();
    const hi=raw.findIndex(row=>row.includes('曲名')&&row.includes('楽曲ID'));
    if(hi<0)throw Error('master header not found');
    const header=raw[hi].map(v=>String(v||'').trim()),aliases={
      title:['曲名'],artist:['アーティスト'],codeUrl:['コードwiki','コードWiki','コードWiki URL'],
      bpm:['BPM'],genre1:['ジャンル','ジャンル①'],genre2:['ジャンル②'],tieup:['タイアップ'],
      search:['検索'],lyrics:['作詞'],composition:['作曲'],arrangement:['編曲']
    },col={};
    Object.keys(aliases).forEach(field=>col[field]=header.findIndex(h=>aliases[field].includes(h)));
    const idCol=header.indexOf('楽曲ID');
    if(idCol<0||['title','artist','codeUrl','genre1'].some(k=>col[k]<0))throw Error('required master columns missing');
    const byPair=new Map(),byUrl=new Map(),allowed=new Set();
    const register=(pair,urlKey,id)=>{
      if(pair){if(!byPair.has(pair))byPair.set(pair,new Set());byPair.get(pair).add(id);}
      if(urlKey){if(!byUrl.has(urlKey))byUrl.set(urlKey,new Set());byUrl.get(urlKey).add(id);}
    };
    raw.slice(hi+1).forEach(row=>{
      const id=String(row[idCol]||'').trim(),title=bulkNorm_(row[col.title]),artist=bulkNorm_(row[col.artist]);
      if(id)register(title&&artist?title+'\u001f'+artist:'',bulkUrl_(row[col.codeUrl]),id);
      if(String(row[col.genre1]||'').trim())allowed.add(String(row[col.genre1]).trim());
      if(col.genre2>=0&&String(row[col.genre2]||'').trim())allowed.add(String(row[col.genre2]).trim());
    });
    const appended=[],responses=[];
    songs.forEach(song=>{
      const sourceIndex=Number(song.sourceIndex),reply={sourceIndex,result:'invalid'};
      responses.push(reply);
      if(!Number.isInteger(sourceIndex)||sourceIndex<0||sourceIndex>250)return;
      const title=clean_(song.title,150),artist=clean_(song.artist,150),parsed=bulkParsePersonalUrl_(song.sourceUrl||song.codeUrl);
      const genre1=clean_(song.genre1,80),genre2=clean_(song.genre2,80),bpm=String(song.bpm||'').trim();
      if(!title||!artist||!parsed||(bpm&&(!/^\d+$/.test(bpm)||+bpm<1||+bpm>400))||(genre2&&genre1===genre2))return;
      const pair=bulkNorm_(title)+'\u001f'+bulkNorm_(artist),ids=new Set([
        ...(byPair.get(pair)||[]),...(byUrl.get(parsed.urlKey)||[])
      ]);
      if(ids.size>1){reply.result='ambiguous';return;}
      let songId='';
      if(ids.size===1){
        songId=Array.from(ids)[0];reply.result='existing';
      }else{
        if(!genre1||!allowed.has(genre1)||(genre2&&!allowed.has(genre2)))return;
        songId=Utilities.getUuid();reply.result='added';
        const row=new Array(sheet.getLastColumn()).fill('');row[idCol]=songId;
        const fields={title,artist,codeUrl:parsed.originalUrl,bpm,genre1,genre2,
          tieup:clean_(song.tieup,250),search:clean_(song.search,800),lyrics:clean_(song.lyrics,150),
          composition:clean_(song.composition,150),arrangement:clean_(song.arrangement,150)};
        Object.keys(fields).forEach(k=>{if(col[k]>=0)row[col[k]]=fields[k];});
        appended.push(row);
      }
      register(pair,parsed.urlKey,songId);
      reply.songId=songId;reply.personalKey=parsed.key;
    });
    if(appended.length){
      const target=sheet.getRange(sheet.getLastRow()+1,1,appended.length,sheet.getLastColumn());
      target.setNumberFormat('@');target.setValues(appended);SpreadsheetApp.flush();
    }
    const response={status:'done',rows:responses};
    bulkSaveStatus_(requestId,response);
    console.log('MYLIST_CSV_DONE uid='+uid+' added='+appended.length+' returned='+responses.length);
    return reply_('done');
  }catch(err){
    console.error('MYLIST_CSV_FAILED',err);
    bulkSaveStatus_(requestId,{status:'error',message:'GAS実行ログを確認してください。'});
    return reply_('error');
  }finally{if(lock.hasLock())lock.releaseLock();}
}
