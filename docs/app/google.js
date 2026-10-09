(function(global){'use strict';
const VERSION='2.0.0',MIME='application/vnd.google-apps.spreadsheet';
const SCOPE='openid email profile https://www.googleapis.com/auth/drive.file';
const {Book,changes}=ChoirModel;
const specs=Object.values(ChoirDomain.TABLES),widths=Object.fromEntries(specs.map(s=>[s[0],s[1].length]));
let config,configPromise,accessToken='',expires=0,profile=null,registry=[],selected='',snapshot=null,queue=Promise.resolve(),epoch=0;
const memory=new Map(),picked=new Set();
function error(message,status=400){return Object.assign(new Error(message),{status});}
function storageKey(){return 'choir-web-connections:'+profile.sub;}
function saveRegistry(){try{localStorage.setItem(storageKey(),JSON.stringify(registry));}catch(e){}}
function getRegistry(){try{const r=JSON.parse(localStorage.getItem(storageKey())||'[]');return Array.isArray(r)?r.filter(c=>/^[\w-]{20,}$/.test(c.sheetId)&&typeof c.name==='string'):[];}catch(e){return [];}}
function requireAuth(){if(!profile||!accessToken||Date.now()>=expires){if(profile)global.dispatchEvent(new Event('choir-auth-expired'));throw error('Google 연결이 만료되었습니다. 다시 연결한 뒤 저장 버튼을 눌러 주세요.',401);}}
async function settings(){if(!configPromise)configPromise=fetch('config.json',{cache:'no-store'}).then(r=>{if(!r.ok)throw error('로그인 설정을 읽지 못했습니다. 잠시 후 다시 시도해 주세요.');return r.json();}).then(c=>config=c);return configPromise;}
function configured(c){return !!(c&&/\.apps\.googleusercontent\.com$/.test(c.clientId||'')&&c.apiKey&&/^\d+$/.test(c.projectNumber||''));}
async function init(){const c=await settings();return {configured:configured(c)};}
function login({renew=false}={}){
 if(!configured(config))return Promise.reject(error('Google 로그인 연결을 준비 중입니다. 운영자 설정이 완료되면 사용할 수 있습니다.'));
 if(!global.google?.accounts?.oauth2)return Promise.reject(error('Google 로그인 창을 준비 중입니다. 잠시 후 다시 눌러 주세요.'));
 const prior=renew?profile?.sub:null;
 return new Promise((resolve,reject)=>{
  let done=false;const fail=e=>{if(!done){done=true;reject(e);}};
  const client=google.accounts.oauth2.initTokenClient({client_id:config.clientId,scope:SCOPE,include_granted_scopes:true,
   error_callback:e=>fail(error(e.type==='popup_closed'?'로그인 창이 닫혔습니다.':'로그인 창을 열지 못했습니다. 팝업 허용 후 다시 눌러 주세요.')),
   callback:async result=>{if(done)return;if(result.error){fail(error('Google 연결을 완료하지 못했습니다. 필요한 권한에 동의해 주세요.'));return;}
    try{
     if(!google.accounts.oauth2.hasGrantedAllScopes(result,'https://www.googleapis.com/auth/drive.file'))throw error('선택한 Google 시트를 읽고 저장하는 권한이 필요합니다.');
     const response=await fetch('https://www.googleapis.com/oauth2/v3/userinfo',{headers:{Authorization:'Bearer '+result.access_token}});
     if(!response.ok)throw error('Google 계정 정보를 확인하지 못했습니다.');const user=await response.json();if(!user.sub||!user.email)throw error('Google 계정을 확인하지 못했습니다.');
     if(prior&&user.sub!==prior)throw error('입력 중인 자료를 보호하려면 이전과 같은 Google 계정으로 연결해 주세요.');
     const same=profile?.sub===user.sub;accessToken=result.access_token;expires=Date.now()+Number(result.expires_in||3600)*1000-60000;profile=user;
     if(!same){epoch++;registry=getRegistry();selected='';snapshot=null;picked.clear();memory.clear();}
     done=true;resolve({email:user.email});
    }catch(e){fail(e);}
   }});
  client.requestAccessToken({prompt:renew?'':'select_account',...(renew&&profile?.email?{hint:profile.email}:{})});
 });
}
async function request(url,{method='GET',body}={}){
 requireAuth();const controller=method==='GET'?new AbortController():null,timer=controller?setTimeout(()=>controller.abort(),30000):null;
 let response;
 try{response=await fetch(url,{method,headers:{Authorization:'Bearer '+accessToken,...(body?{'Content-Type':'application/json'}:{})},...(body?{body:JSON.stringify(body)}:{}),...(controller?{signal:controller.signal}:{})});}
 catch(e){throw error(method==='GET'?'Google에서 자료를 읽지 못했습니다. 인터넷 연결을 확인해 주세요.':'저장 결과를 확인하지 못했습니다. 자료를 다시 불러와 저장 여부를 확인해 주세요. 중복 방지를 위해 자동 재시도하지 않습니다.',503);}
 finally{clearTimeout(timer);}
 if(response.status===401){expires=0;global.dispatchEvent(new Event('choir-auth-expired'));throw error('Google 연결을 갱신한 뒤 다시 시도해 주세요.',401);}
 if(!response.ok){const detail=await response.json().catch(()=>({}));throw error(response.status===403?'이 Google 계정에 시트 편집 권한이 없거나 Google 연결 설정이 완료되지 않았습니다.':response.status===404?'시트를 열 수 없습니다. Google 시트 선택으로 다시 연결해 주세요.':response.status===429?'요청이 많아 Google 응답이 지연됩니다. 잠시 후 다시 시도해 주세요.':detail.error?.message||'Google 요청을 완료하지 못했습니다.',response.status);}
 if(response.status===204)return {};return response.json();
}
const drive=(path,options)=>request('https://www.googleapis.com/drive/v3/'+path,options);
const sheets=(id,path,options)=>request('https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(id)+path,options);
async function metadata(id){const m=await drive('files/'+encodeURIComponent(id)+'?fields=id,name,mimeType,version,modifiedTime,capabilities(canEdit,canShare)');if(m.mimeType!==MIME)throw error('Google 스프레드시트를 선택해 주세요.');return m;}
function revision(m){return String(m.version||m.modifiedTime||'');}
let dbPromise;
function db(){if(!global.indexedDB)return Promise.resolve(null);if(!dbPromise)dbPromise=new Promise(resolve=>{const r=indexedDB.open('choir-web-cache',1);r.onupgradeneeded=()=>r.result.createObjectStore('snapshots');r.onsuccess=()=>resolve(r.result);r.onerror=()=>resolve(null);});return dbPromise;}
async function cache(action,key,value){const d=await db();if(!d)return action==='get'?memory.get(key):memory.set(key,value);return new Promise(resolve=>{try{const tx=d.transaction('snapshots',action==='get'?'readonly':'readwrite'),s=tx.objectStore('snapshots'),r=action==='get'?s.get(key):action==='delete'?s.delete(key):s.put(value,key);r.onsuccess=()=>resolve(r.result);r.onerror=()=>resolve();}catch(e){resolve();}});}
function cacheKey(id){return profile.sub+':'+id;}
async function load(id,{force=false,meta}={}){
 const runEpoch=epoch,key=cacheKey(id),m=meta||await metadata(id);
 if(!m.capabilities?.canEdit)throw error('시트 편집자로 공유받은 Google 계정으로 연결해 주세요.',403);
 let cached=!force&&(snapshot?.id===id?snapshot:await cache('get',key));
 if(cached&&cached.version===revision(m)){snapshot={...cached,book:new Book(cached.tables),meta:m};return snapshot;}
 const info=await sheets(id,'?fields=spreadsheetId,sheets(properties(sheetId,title,gridProperties(rowCount)))');
 const tabs=info.sheets.map(x=>x.properties),available=specs.filter(s=>tabs.some(t=>t.title===s[0]));
 const query=available.map(s=>'ranges='+encodeURIComponent("'"+s[0]+"'!A:"+String.fromCharCode(64+s[1].length))).join('&');
 const result=available.length?await sheets(id,'/values:batchGet?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER&'+query):{valueRanges:[]};
 const after=await metadata(id);if(revision(m)!==revision(after))throw error('자료를 불러오는 동안 시트가 변경되었습니다. 다시 불러와 주세요.',409);
 const tables=available.map((s,i)=>{const p=tabs.find(t=>t.title===s[0]),values=result.valueRanges[i].values||[];if(s[0]==='회차')values.slice(1).forEach(r=>{if(typeof r[1]==='number')r[1]=new Date(Date.UTC(1899,11,30)+Math.floor(r[1])*86400000).toISOString().slice(0,10);});return {name:s[0],sheetId:p.sheetId,capacity:p.gridProperties.rowCount,values};});
 if(runEpoch!==epoch)throw error('Google 계정이 변경되었습니다. 다시 연결해 주세요.',401);
 const value={id,version:revision(after),tables};snapshot={...value,book:new Book(tables),meta:after};await cache('put',key,value);return snapshot;
}
function validate(book){specs.forEach(([name,header])=>{const sh=book.getSheetByName(name);if(sh?.getLastRow()&&header.some((h,i)=>sh.values[0]?.[i]!==h))throw error('“'+name+'” 탭이 출석부 형식과 다릅니다. 원본을 변경하지 않았습니다. 출석부 시트를 선택해 주세요.');});}
function fixedName(book){const rows=book.getSheetByName('설정')?.values||[],names=rows.filter(r=>r[0]==='_choir_name').map(r=>String(r[1]));if(new Set(names).size>1)throw error('시트의 합창단 이름 설정이 서로 다릅니다.');return names[0]||'';}
function context(book,id=selected){const c=registry.find(c=>c.sheetId===id);return {book,id,name:c?.name||fixedName(book),email:profile.email,choirs:registry,inviteUrl:new URL('./',location.href).href+'?sheet='+encodeURIComponent(id)+'&choir='+encodeURIComponent(c?.name||fixedName(book))};}
async function commit(before,after,baseline){const requests=changes(before,after,widths);if(!requests.length)return;
 const current=await metadata(baseline.id);if(!current.capabilities?.canEdit)throw error('시트 편집 권한이 없습니다.',403);
 if(revision(current)!==baseline.version){snapshot=null;await cache('delete',cacheKey(baseline.id));throw error('다른 담당자가 시트를 변경했습니다. 입력 내용은 화면에 유지됩니다. 자료를 다시 불러온 뒤 변경사항을 확인하고 저장해 주세요.',409);}
 try{await sheets(baseline.id,':batchUpdate',{method:'POST',body:{requests}});}
 finally{snapshot=null;await cache('delete',cacheKey(baseline.id));}
}
async function connect(id,name,sample=false){
 let base=await load(id,{force:true});validate(base.book);const existing=fixedName(base.book);name=String(name||existing).trim();if(!name||name.length>80)throw error('80자 이내의 합창단 이름을 입력해 주세요.');if(existing&&existing!==name)throw error('이 시트의 합창단 이름은 “'+existing+'”입니다. 등록 후 이름을 바꿀 수 없습니다.');
 const missing=specs.filter(s=>!base.book.getSheetByName(s[0])),empty=specs.filter(s=>!base.book.getSheetByName(s[0])?.getLastRow());
 if(missing.length||empty.length||!existing){
  const info=await sheets(id,'?fields=sheets(properties(sheetId))');let next=Math.max(-1,...info.sheets.map(s=>s.properties.sheetId))+1;
  const requests=[],ids={};missing.forEach(s=>{ids[s[0]]=next++;requests.push({addSheet:{properties:{sheetId:ids[s[0]],title:s[0],gridProperties:{rowCount:1000,columnCount:Math.max(8,s[1].length)}}}});});
  empty.forEach(s=>requests.push({updateCells:{start:{sheetId:ids[s[0]]??base.book.getSheetByName(s[0]).sheetId,rowIndex:0,columnIndex:0},rows:[{values:s[1].map(ChoirModel.cell)}],fields:'userEnteredValue'}}));
  if(!existing)requests.push({appendCells:{sheetId:ids['설정']??base.book.getSheetByName('설정').sheetId,rows:[{values:['_choir_name',name].map(ChoirModel.cell)}],fields:'userEnteredValue'}});
  const now=await metadata(id);if(revision(now)!==base.version)throw error('연결하는 동안 시트가 변경되었습니다. 다시 연결해 주세요.',409);
  try{await sheets(id,':batchUpdate',{method:'POST',body:{requests}});}finally{snapshot=null;await cache('delete',cacheKey(id));}
  base=await load(id,{force:true});
 }
 if(!registry.some(c=>c.sheetId===id)){registry.push({sheetId:id,name,sample});saveRegistry();}
 selected=id;return {token:id};
}
async function create(name,nonce,sample=false){
 name=String(name||'').trim();if(!name||name.length>80)throw error('80자 이내의 합창단 이름을 입력해 주세요.');if(!/^[\w-]{10,80}$/.test(nonce))throw error('생성 요청을 다시 시작해 주세요.');
 const q="trashed=false and appProperties has { key='choirRequest' and value='"+nonce+"' }";
 const found=await drive('files?q='+encodeURIComponent(q)+'&fields=files(id,appProperties)&pageSize=100');let file=found.files?.[0];
 if(file&&file.appProperties?.choirName!==name)throw error('진행 중인 합창단 이름이 다릅니다. 연결 창을 다시 열어 주세요.');
 if(!file)file=await drive('files?fields=id',{method:'POST',body:{name:'합창단 출석부 · '+name,mimeType:MIME,appProperties:{choirRequest:nonce,choirName:name,choirSample:sample?'1':'0'}}});
 await connect(file.id,name,sample);
 if(sample){const base=await load(file.id);if(!base.book.getSheetByName('단원').getLastRow()||base.book.getSheetByName('단원').getLastRow()===1){const next=base.book.clone();ChoirDomain.seed(context(next));await commit(base.book,next,base);}}
 return {token:file.id};
}
async function pick(){requireAuth();if(!configured(config))throw error('Google 시트 연결 준비 중입니다.');
 if(!global.gapi)throw error('Google 시트 선택 창을 준비 중입니다. 잠시 후 다시 눌러 주세요.');
 await new Promise((resolve,reject)=>gapi.load('picker',{callback:resolve,onerror:()=>reject(error('시트 선택 창을 열지 못했습니다.')),timeout:10000,ontimeout:()=>reject(error('시트 선택 창을 다시 열어 주세요.'))}));
 return new Promise((resolve,reject)=>{const view=new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS).setIncludeFolders(false).setMimeTypes(MIME);const picker=new google.picker.PickerBuilder().setDeveloperKey(config.apiKey).setAppId(config.projectNumber).setOAuthToken(accessToken).setOrigin(location.origin).setTitle('출석부 Google 시트를 선택하세요').addView(view).setCallback(async data=>{if(data.action===google.picker.Action.CANCEL)resolve(null);if(data.action===google.picker.Action.PICKED){try{const id=data.docs[0].id;picked.add(id);const base=await load(id,{force:true});validate(base.book);resolve({id,name:fixedName(base.book),url:'https://docs.google.com/spreadsheets/d/'+id+'/edit'});}catch(e){reject(e);}}}).build();picker.setVisible(true);});
}
async function sharing(){const result=await drive('files/'+encodeURIComponent(selected)+'/permissions?fields=permissions(emailAddress,role,type),nextPageToken&pageSize=100');let perms=result.permissions||[],next=result.nextPageToken;while(next){const page=await drive('files/'+encodeURIComponent(selected)+'/permissions?fields=permissions(emailAddress,role,type),nextPageToken&pageSize=100&pageToken='+encodeURIComponent(next));perms=perms.concat(page.permissions||[]);next=page.nextPageToken;}
 return {editors:perms.filter(p=>['writer','organizer','fileOrganizer'].includes(p.role)&&p.emailAddress).map(p=>p.emailAddress).sort(),inviteUrl:context(snapshot?.book||new Book()).inviteUrl};}
async function logout(){const who=profile?.sub;epoch++;accessToken='';expires=0;profile=null;selected='';snapshot=null;picked.clear();memory.clear();registry=[];
 const d=await db();if(d&&who)await new Promise(resolve=>{const tx=d.transaction('snapshots','readwrite'),cursor=tx.objectStore('snapshots').openCursor();cursor.onsuccess=()=>{const c=cursor.result;if(c){if(String(c.key).startsWith(who+':'))c.delete();c.continue();}};tx.oncomplete=tx.onerror=()=>resolve();});return {ok:true};}
async function dispatch(method,path,body={}){
 const route=path.split('?')[0];if(route==='/api/auth/status')return {choirs:profile?registry:[],googleEmail:profile?.email||'',signedIn:!!profile,version:VERSION,requiresApproval:false};
 if(route==='/api/auth/logout')return logout();requireAuth();
 if(['/api/choirs','/api/auth/register'].includes(route)&&method==='POST'){
  if(body.createNew)return create(body.choir,body.requestId);
  const m=String(body.sheetId||'').match(/\/spreadsheets\/d\/([\w-]{20,})/),id=m?m[1]:String(body.sheetId||'');if(!picked.has(id)&&!registry.some(c=>c.sheetId===id))throw error('Google 시트 선택 버튼으로 기존 시트를 선택해 주세요.');return connect(id,body.choir);
 }
 if(['/api/choirs/select','/api/auth/login'].includes(route)){const c=registry.find(c=>c.sheetId===body.sheetId);if(!c)throw error('연결된 합창단을 선택해 주세요.');const base=await load(c.sheetId,{force:true});validate(base.book);if(fixedName(base.book)!==c.name)throw error('시트의 합창단 이름이 등록 정보와 다릅니다.');if(specs.some(s=>!base.book.getSheetByName(s[0])))throw error('출석부 탭이 누락되었습니다. 기존 시트 연결로 확인해 주세요.');selected=c.sheetId;return {token:selected};}
 const unlink=route.match(/^\/api\/choirs\/([\w-]{20,})$/);if(unlink&&method==='DELETE'){registry=registry.filter(c=>c.sheetId!==unlink[1]);saveRegistry();await cache('delete',cacheKey(unlink[1]));if(selected===unlink[1]){selected='';snapshot=null;}return {ok:true,choirs:registry};}
 if(route==='/api/sample'){const prior=registry.find(c=>c.sample);if(prior)return dispatch('POST','/api/choirs/select',{sheetId:prior.sheetId});let nonce;try{nonce=localStorage.getItem('choir-sample-request:'+profile.sub);}catch(e){}if(!nonce){nonce=crypto.randomUUID();try{localStorage.setItem('choir-sample-request:'+profile.sub,nonce);}catch(e){}}return create('샘플 합창단',nonce,true);}
 if(!selected||body.token!==selected)throw error('사용할 합창단을 선택해 주세요.',401);
 if(route==='/api/share'){if(method==='GET')return sharing();const email=String(body.email||'').trim().toLowerCase();if(email.length>254||!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email))throw error('Google 이메일 한 개를 정확히 입력해 주세요.');const meta=await metadata(selected);if(!meta.capabilities.canShare)throw error('시트 공유 권한이 있는 계정으로 접속해 주세요.',403);const list=await sharing();if(!list.editors.includes(email))await drive('files/'+encodeURIComponent(selected)+'/permissions?sendNotificationEmail=false',{method:'POST',body:{type:'user',role:'writer',emailAddress:email}});return {ok:true,email,inviteUrl:list.inviteUrl};}
 const mutating=method!=='GET';let base;
 if(mutating&&snapshot?.id===selected)base=snapshot;else base=await load(selected);
 validate(base.book);if(fixedName(base.book)!==registry.find(c=>c.sheetId===selected)?.name)throw error('등록된 합창단 이름과 시트 설정이 다릅니다.');
 const work=mutating?base.book.clone():base.book,result=ChoirDomain.run(context(work),method,path,body);if(mutating)await commit(base.book,work,base);return result||{ok:true};
}
function api(method,path,body){const task=queue.then(()=>dispatch(method,path,body));queue=task.catch(()=>{});return task;}
const params=new URLSearchParams(location.search),id=params.get('sheet')||'',name=(params.get('choir')||'').trim();global.choirInvitation=/^[\w-]{20,}$/.test(id)&&name&&name.length<=80?{sheetId:id,choir:name}:{};
global.ChoirWeb={init,login,pick,api};
})(globalThis);
