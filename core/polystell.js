/* =====================================================================
   POLYSTELL APPS — núcleo compartilhado (v1)
   <script src="/core/polystell.js?v=1"></script>
   Fornece: PS.auth (MSAL com recuperação interativa), PS.graph (retry
   correto, sem repetir gravação incerta), PS.sp (lista/itens/paginação
   completa), PS.shell (menu lateral, cabeçalho, conta, voltar ao hub),
   PS.ui (toast, diálogo, ícones) e telas padrão (entrada/carga/erro).
   Sem segredos: clientId/tenantId são configuração pública de SPA.
   Esconder botão é usabilidade; autorização real = permissão da lista.
   ===================================================================== */
(function(){
'use strict';
const CFG=Object.freeze({
 clientId:'c066630d-9dc2-4c4a-8c60-8f0056e9e5e2',
 tenantId:'fdb7cb21-cc9d-4f32-bb70-e71cad1eb164',
 siteHost:'polystellcombr.sharepoint.com', sitePath:'/sites/Apps',
 msalUrl:'https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js',
 hubUrl:'/', logo:'/logo-polystell.png',
 supportUrl:'https://suporte.polystell.com.br/portal/pt-br/home',
 baseScopes:['Sites.ReadWrite.All','User.Read'],
 groupScopes:['GroupMember.Read.All']
});
/* ---------------- utilitários ---------------- */
const $=(s,r=document)=>r.querySelector(s);
const $$=(s,r=document)=>[...r.querySelectorAll(s)];
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=s=>String(s??'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().trim();
const val=v=>{if(v==null)return'';if(typeof v==='object'&&'Value'in v)return String(v.Value??'');return String(v);};
const pad=n=>String(n).padStart(2,'0');
const todayKey=()=>{const d=new Date();return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());};
/* Coluna "somente data" do SharePoint chega como meia-noite local em UTC
   (ex.: 2026-10-08T03:00:00Z). Hora UTC ≤ 05 → vale a parte de data. */
function dateKey(v){if(!v)return'';const s=String(v).trim();if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;const m=s.match(/^(\d{4}-\d{2}-\d{2})T(\d{2})/);if(m&&+m[2]<=5)return m[1];const d=new Date(s);if(isNaN(d))return s.slice(0,10);return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());}
const fmt={
 hm:v=>v?new Date(v).toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}):'—',
 dm:v=>v?new Date(v).toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit'}):'—',
 dmhm:v=>v?new Date(v).toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit'})+' '+new Date(v).toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}):'—',
 data:k=>k?String(k).slice(0,10).split('-').reverse().join('/'):'—',
 dur:ms=>{if(!(ms>0))return'—';const m=Math.floor(ms/6e4);if(m<60)return m+' min';const h=Math.floor(m/60);if(h<48)return h+'h'+(m%60?pad(m%60):'');return Math.floor(h/24)+' dias';}
};
/* ---------------- ícones (mesma família do hub) ---------------- */
const IC={
 home:'<path d="m3 10 9-7 9 7v10a1 1 0 0 1-1 1h-5v-7H9v7H4a1 1 0 0 1-1-1Z"/>',
 grid:'<rect x="3.5" y="3.5" width="6" height="6" rx="1.3"/><rect x="14.5" y="3.5" width="6" height="6" rx="1.3"/><rect x="3.5" y="14.5" width="6" height="6" rx="1.3"/><rect x="14.5" y="14.5" width="6" height="6" rx="1.3"/>',
 users:'<circle cx="9" cy="8" r="3"/><path d="M3 21v-3a6 6 0 0 1 12 0v3M16 5a3 3 0 0 1 0 6m3 9v-3a5 5 0 0 0-3-4"/>',
 user:'<circle cx="12" cy="8" r="4"/><path d="M4 21v-1a8 8 0 0 1 16 0v1"/>',
 truck:'<path d="M2 6h11v10H2zM13 9h4l4 4v3h-8z"/><circle cx="6" cy="18" r="2"/><circle cx="17" cy="18" r="2"/>',
 box:'<path d="m3 7 9-4 9 4v10l-9 4-9-4Z"/><path d="m3 7 9 4 9-4M12 11v10"/>',
 alert:'<path d="m12 3 10 18H2L12 3Z"/><path d="M12 9v5m0 3h.01"/>',
 shield:'<path d="m12 3 8 3v6c0 5-8 9-8 9s-8-4-8-9V6l8-3Z"/><path d="m8 12 3 3 5-6"/>',
 gear:'<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 0 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 0 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 0 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 0 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1Z"/>',
 swap:'<path d="M7 4 3 8l4 4M3 8h14M17 20l4-4-4-4m4 4H7"/>',
 calendar:'<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4m10-4v4M3 11h18"/>',
 clock:'<circle cx="12" cy="12" r="8.5"/><path d="M12 7v5l3.5 2"/>',
 in:'<path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4M10 17l5-5-5-5M15 12H3"/>',
 out:'<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/>',
 search:'<circle cx="10.5" cy="10.5" r="6.8"/><path d="m16 16 4.5 4.5"/>',
 help:'<circle cx="12" cy="12" r="9"/><path d="M9.5 8.5a2.5 2.5 0 1 1 4.4 1.6L12 12v1m0 4h.01"/>',
 back:'<path d="M20 12H5m6-6-6 6 6 6"/>',
 chevR:'<path d="m9 5 7 7-7 7"/>', chevD:'<path d="m6 9 6 6 6-6"/>',
 menu:'<path d="M4 6h16M4 12h16M4 18h16"/>', x:'<path d="m6 6 12 12M6 18 18 6"/>',
 check:'<path d="m5 12 4 4L19 6"/>', plus:'<path d="M12 5v14M5 12h14"/>',
 refresh:'<path d="M20 11a8 8 0 0 0-14.9-3.9L3 9m0-5v5h5M4 13a8 8 0 0 0 14.9 3.9L21 15m0 5v-5h-5"/>',
 cloudoff:'<path d="m3 3 18 18M8.5 8.6A5 5 0 0 0 6 18h11m3.3-1.6A4 4 0 0 0 17 10h-1.3A7 7 0 0 0 10 5.3"/>',
 qr:'<path d="M4 8V4h4m8 0h4v4m0 8v4h-4m-8 0H4v-4"/><rect x="8" y="8" width="3" height="3"/><path d="M14 8h2v3m-8 3h3v2m3-2h2v2"/>',
 camera:'<path d="M3 8a2 2 0 0 1 2-2h2l2-2h6l2 2h2a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/><circle cx="12" cy="13" r="4"/>',
 logout:'<path d="M9 3H4v18h5M9 12h12m-5-5 5 5-5 5"/>',
 lock:'<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
 car:'<path d="M5 17h14M5 17a2 2 0 1 0 4 0m6 0a2 2 0 1 0 4 0M3 17v-4l2-5h14l2 5v4"/>',
 leaf:'<path d="M20 3c-8 0-16 2-16 10a7 7 0 0 0 7 7c8 0 9-9 9-17ZM4 21l11-11"/>'
};
const icon=(n,cls='')=>'<svg class="i '+cls+'" viewBox="0 0 24 24" aria-hidden="true">'+(IC[n]||IC.grid)+'</svg>';
/* ---------------- toast / diálogo ---------------- */
let toastT=null;
let dlgCapture=null;
function toast(msg,kind){if(dlgCapture){dlgCapture.push([msg,kind]);return;}const dl=$('#ps-dlg');if(dl&&dl.open){const m=dl.querySelector('.dlg-msg');if(m){m.className='dlg-msg banner '+(kind==='err'?'err':kind==='warn'?'warn':'ok');m.textContent=msg;m.hidden=false;return;}}
 let t=$('#ps-toast');if(!t){t=document.createElement('div');t.id='ps-toast';t.className='ps-toast';t.setAttribute('role','status');t.setAttribute('aria-live','polite');document.body.appendChild(t);}
 clearTimeout(toastT);t.textContent=msg;t.className='ps-toast show'+(kind?' '+kind:'');toastT=setTimeout(()=>t.classList.remove('show'),kind==='err'?6000:3800);}
let dlgBusy=false;
function dialog({title,html='',actions=[{label:'Fechar',value:null}]}){
 return new Promise(resolve=>{
  let d=$('#ps-dlg');if(!d){d=document.createElement('dialog');d.id='ps-dlg';d.className='ps-dlg';document.body.appendChild(d);}
  if(d.open)d.close();
  d.setAttribute('aria-labelledby','ps-dlg-t');
  d.innerHTML='<div class="body"><h3 id="ps-dlg-t">'+esc(title)+'</h3>'+html+'<div class="dlg-msg" role="alert" hidden></div></div><div class="acts">'+actions.map((a,i)=>'<button type="button" class="btn '+(a.kind||'')+'" data-a="'+i+'">'+esc(a.label)+'</button>').join('')+'</div>';
  let settled=false;const finish=v=>{if(settled)return;settled=true;if(d.open)d.close();resolve(v);};
  d.querySelectorAll('[data-a]').forEach(b=>b.onclick=async()=>{const a=actions[+b.dataset.a];
   if(!a.onClick){finish(a.value);return;}
   dlgBusy=true;d.querySelectorAll('.acts button').forEach(x=>x.disabled=true);
   /* mensagens geradas durante a ação: se o diálogo continua aberto, aparecem
      dentro dele (camada superior); se ele fecha, viram aviso normal */
   dlgCapture=[];let r;try{r=await a.onClick(d);}catch(e){r=false;dlgCapture.push([e.message||'Erro','err']);}
   const msgs=dlgCapture;dlgCapture=null;
   dlgBusy=false;d.querySelectorAll('.acts button').forEach(x=>x.disabled=false);
   if(r===false){const m=msgs[msgs.length-1];if(m)toast(m[0],m[1]);return;}
   finish(r===undefined?a.value:r);const m=msgs[msgs.length-1];if(m)toast(m[0],m[1]);});
  d.oncancel=e=>{if(dlgBusy){e.preventDefault();return;}finish(undefined);};
  d.showModal();
  setTimeout(()=>{const f=d.querySelector('.body input,.body select,.body textarea');if(f)f.focus();},40);
 });
}
/* ---------------- autenticação ---------------- */
let msalApp=null,account=null;
function loadScript(src,ms=15000){return new Promise((ok,no)=>{const s=document.createElement('script');const t=setTimeout(()=>{s.remove();no(new Error('Tempo esgotado ao carregar '+src.split('/').pop()));},ms);s.src=src;s.onload=()=>{clearTimeout(t);ok();};s.onerror=()=>{clearTimeout(t);no(new Error('Não foi possível carregar '+src.split('/').pop()));};document.head.appendChild(s);});}
const isInteraction=e=>!!e&&(e.name==='InteractionRequiredAuthError'||/interaction_required|consent_required|login_required|no_tokens_found|token_renewal/i.test(String(e.errorCode||'')+' '+String(e.message||'')));
const isPopupBlocked=e=>/popup_window_error|empty_window_error|popup/i.test(String(e?.errorCode||'')+' '+String(e?.message||''));
/* Recuperação de sessão:
   • na abertura (nada digitado ainda) → redirect;
   • depois de aberto, só com gesto do usuário (clique) → popup;
   • em segundo plano (timer) → NUNCA interativo: mostra a barra
     "Sessão expirada — Entrar novamente" e preserva a tela. */
let mode='boot',lastGesture=0,popupP=null;
function popupOnce(req){if(!popupP)popupP=msalApp.acquireTokenPopup(req).finally(()=>{popupP=null;});return popupP;}
['pointerdown','keydown'].forEach(ev=>document.addEventListener(ev,()=>{lastGesture=Date.now();},true));
const hasGesture=()=>navigator.userActivation?navigator.userActivation.isActive:Date.now()-lastGesture<4000;
function sessionBar(){if($('#ps-sess'))return;const b=document.createElement('div');b.id='ps-sess';b.className='ps-sess';b.setAttribute('role','alert');
 b.innerHTML=icon('lock')+'<span><b>Sua sessão expirou.</b> O que você digitou continua na tela.</span><button type="button" class="btn sm primary">Entrar novamente</button>';
 b.querySelector('button').onclick=async()=>{try{const r=await popupOnce({scopes:CFG.baseScopes,account});account=r.account||account;b.remove();window.dispatchEvent(new Event('ps:session'));}
  catch(p){if(isPopupBlocked(p))await msalApp.acquireTokenRedirect({scopes:CFG.baseScopes,account});else toast('Não foi possível entrar: '+(p.message||'erro'),'err');}};
 document.body.appendChild(b);}
const auth={
 async init(){
  if(location.protocol!=='https:'&&!['localhost','127.0.0.1'].includes(location.hostname))throw Object.assign(new Error('Acesse pelo endereço oficial (HTTPS): app.polystell.com.br'),{code:'https'});
  await loadScript(CFG.msalUrl);
  msalApp=new msal.PublicClientApplication({auth:{clientId:CFG.clientId,authority:'https://login.microsoftonline.com/'+CFG.tenantId,redirectUri:location.origin+location.pathname,navigateToLoginRequestUrl:false},cache:{cacheLocation:'sessionStorage'}});
  if(typeof msalApp.initialize==='function')await msalApp.initialize();
  const r=await msalApp.handleRedirectPromise();
  account=r?.account||msalApp.getActiveAccount()||msalApp.getAllAccounts()[0]||null;
  if(account)msalApp.setActiveAccount(account);
  return account;
 },
 ready(){mode='app';},
 login(scopes=CFG.baseScopes){return msalApp.loginRedirect({scopes});},
 logout(){try{sessionStorage.clear();}catch{}return msalApp.logoutRedirect({account,postLogoutRedirectUri:location.origin+location.pathname});},
 get account(){return account;},
 sessionLost(){sessionBar();return Object.assign(new Error('Sessão expirada — use "Entrar novamente" no topo da tela.'),{code:'session'});},
 async token(scopes=CFG.baseScopes,{optional=false,force=false}={}){
  const req={scopes,account,forceRefresh:force};
  try{return (await msalApp.acquireTokenSilent(req)).accessToken;}
  catch(e){
   if(!isInteraction(e))throw e;
   if(optional)throw Object.assign(new Error('Permissão opcional indisponível.'),{code:'optional'});
   if(mode==='boot'){await msalApp.acquireTokenRedirect(req);throw new Error('Redirecionando para entrar novamente…');}
   if(hasGesture()){try{const r=await popupOnce(req);account=r.account||account;const sb=$('#ps-sess');if(sb)sb.remove();return r.accessToken;}
    catch(p){if(isPopupBlocked(p)){sessionBar();throw Object.assign(new Error('Sessão expirada — use "Entrar novamente" no topo da tela.'),{code:'session'});}throw p;}}
   sessionBar();throw Object.assign(new Error('Sessão expirada — use "Entrar novamente" no topo da tela.'),{code:'session'});
  }
 }
};
/* ---------------- Microsoft Graph ---------------- */
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
async function graph(path,opt={}){
 const url=new URL(path.startsWith('https://')?path:'https://graph.microsoft.com/v1.0'+path);
 if(url.origin!=='https://graph.microsoft.com')throw new Error('Destino de API não autorizado.');
 const method=(opt.method||'GET').toUpperCase(),isRead=method==='GET';
 let tok=await auth.token(opt.scopes||CFG.baseScopes,{optional:!!opt.optional});
 let waited=0,renewed=false;
 for(let a=0;a<4;a++){
  const ctl=new AbortController(),tm=setTimeout(()=>ctl.abort(),opt.timeout||25000);let r;
  try{r=await fetch(url,{method,headers:{Authorization:'Bearer '+tok,'Content-Type':'application/json',...(opt.headers||{})},body:opt.body!=null?JSON.stringify(opt.body):undefined,signal:ctl.signal});}
  catch(e){clearTimeout(tm);
   /* Falha de rede: leitura pode repetir; gravação NÃO (pode ter sido aplicada) */
   if(isRead&&a<2){await sleep(1500*(a+1));continue;}
   throw Object.assign(new Error(isRead?'Sem conexão com o Microsoft 365.':'Sem confirmação do servidor — a gravação pode ou não ter sido aplicada.'),{status:0,uncertain:!isRead});}
  clearTimeout(tm);
  if(r.status===401&&!renewed){renewed=true;tok=await auth.token(opt.scopes||CFG.baseScopes,{optional:!!opt.optional,force:true});a--;continue;}
  if(r.status===401){if(opt.optional)throw Object.assign(new Error('Não autorizado (401).'),{status:401});throw auth.sessionLost();}
  /* 429 = não processado → pode repetir sempre. 503/504 → só leitura. */
  const retriable=r.status===429||(isRead&&[503,504].includes(r.status));
  if(retriable&&a<3){const ra=Number(r.headers.get('Retry-After'))||(a+1)*2;if(waited+ra>60)throw Object.assign(new Error('O Microsoft 365 pediu para aguardar '+ra+'s. Tente novamente em instantes.'),{status:r.status});waited+=ra;await sleep(ra*1000);continue;}
  if(r.status===204)return null;
  if(!r.ok){let d='',code='';try{const j=await r.json();d=j?.error?.message||'';code=j?.error?.code||'';}catch{}
   const msg=r.status===403?'Sem permissão para esta operação (403).':r.status===404?'Recurso não encontrado (404).':r.status===412?'O registro foi alterado em outro dispositivo (412).':'Microsoft Graph '+r.status+(d?' — '+d:'');
   throw Object.assign(new Error(msg),{status:r.status,code,detail:d,uncertain:!isRead&&r.status>=500});}
  try{return await r.json();}catch(e){if(isRead)throw new Error('Resposta inválida do Microsoft 365.');throw Object.assign(new Error('O servidor respondeu sem confirmação legível — a gravação pode ter sido aplicada.'),{status:r.status,uncertain:true});}
 }
 throw new Error('Falha ao consultar o Microsoft Graph.');
}
/* ---------------- SharePoint (site Apps) ---------------- */
let SITE=null,LISTS=null;
const sp={
 async site(){if(!SITE)SITE=(await graph('/sites/'+CFG.siteHost+':'+CFG.sitePath)).id;return SITE;},
 async lists(){if(LISTS)return LISTS;const s=await sp.site();const r=await graph('/sites/'+s+'/lists?$select=id,name,displayName&$top=999');LISTS={};(r.value||[]).forEach(l=>{LISTS[l.displayName]=l.id;});return LISTS;},
 async columns(listId){const s=await sp.site();const r=await graph('/sites/'+s+'/lists/'+listId+'/columns?$top=500');const m=new Map();(r.value||[]).forEach(c=>m.set(c.name,c));return m;},
 /* paginação completa; para em max e marca .truncated (nunca corta em silêncio) */
 async items(listId,query,max=5000){const s=await sp.site();let out=[],next='/sites/'+s+'/lists/'+listId+'/items?'+query,trunc=false;
  while(next){const p=await graph(next,{headers:{Prefer:'HonorNonIndexedQueriesWarningMayFailRandomly'}});out.push(...(p.value||[]));next=p['@odata.nextLink']||null;if(next&&out.length>=max){trunc=true;break;}}
  const rows=out.map(it=>({id:String(it.id),_etag:it['@odata.etag']||it.eTag||'',_by:it.createdBy?.user?.displayName||'',_byEmail:String(it.createdBy?.user?.email||'').toLowerCase(),_byId:String(it.createdBy?.user?.id||''),...(it.fields||{})}));rows.truncated=trunc;return rows;},
 async add(listId,fields){const s=await sp.site();const r=await graph('/sites/'+s+'/lists/'+listId+'/items',{method:'POST',body:{fields}});return {id:String(r.id),...(r.fields||{})};},
 async patch(listId,id,fields,{etag}={}){const s=await sp.site();return graph('/sites/'+s+'/lists/'+listId+'/items/'+id+'/fields',{method:'PATCH',body:fields,headers:etag?{'If-Match':etag}:{}});},
 async get(listId,id){const s=await sp.site();const r=await graph('/sites/'+s+'/lists/'+listId+'/items/'+id+'?$expand=fields');return {id:String(r.id),_etag:r['@odata.etag']||r.eTag||'',_by:r.createdBy?.user?.displayName||'',_byEmail:String(r.createdBy?.user?.email||'').toLowerCase(),_byId:String(r.createdBy?.user?.id||''),...(r.fields||{})};}
};
async function groups(){try{let next='/me/transitiveMemberOf/microsoft.graph.group?$select=id,displayName&$top=999';const out=new Set();
 while(next){const p=await graph(next,{scopes:CFG.groupScopes,optional:true,headers:{ConsistencyLevel:'eventual'}});(p.value||[]).forEach(g=>{if(g.id)out.add(String(g.id).toLowerCase());if(g.displayName)out.add(String(g.displayName).toLowerCase());});next=p['@odata.nextLink']||null;}
 return out;}catch(e){console.warn('Grupos indisponíveis:',e.message);return null;}}
async function photo(){try{const t=await auth.token(['User.Read'],{optional:true});const r=await fetch('https://graph.microsoft.com/v1.0/me/photos/96x96/$value',{headers:{Authorization:'Bearer '+t}});if(!r.ok)return null;return URL.createObjectURL(await r.blob());}catch{return null;}}
/* ---------------- telas padrão ---------------- */
function brandHtml(){return '<div class="ps-brand"><img src="'+CFG.logo+'" alt="Polystell" onerror="this.hidden=true;this.nextElementSibling.hidden=false"><span class="fb" hidden>polystell<i>.</i></span></div>';}
const screens={
 loading(text){document.body.innerHTML='<div class="ps-screen" aria-live="polite"><div class="box"><div class="ps-spin"></div><p>'+esc(text||'Carregando…')+'</p></div></div>';},
 error(title,msg,{retry=true,details=null,action=null}={}){
  document.body.innerHTML='<div class="ps-screen"><div class="box"><h1>'+esc(title)+'</h1><p>'+esc(msg)+'</p>'+(details&&details.length?'<ul>'+details.map(d=>'<li>'+esc(d)+'</li>').join('')+'</ul>':'')+
   '<div style="display:flex;gap:9px;justify-content:center;flex-wrap:wrap">'+(retry?'<button class="btn primary" id="ps-retry">'+icon('refresh')+'Tentar novamente</button>':'')+(action?'<button class="btn" id="ps-act">'+esc(action.label)+'</button>':'')+'<a class="btn" href="'+CFG.hubUrl+'">'+icon('back')+'Todos os aplicativos</a></div></div></div>';
  const r=$('#ps-retry');if(r)r.onclick=()=>location.reload();const a=$('#ps-act');if(a)a.onclick=action.onClick;},
 login({app,title='Bom ter você aqui.',text,onLogin}){
  document.body.innerHTML='<section class="ps-login"><div class="story">'+brandHtml()+'<div class="msg"><div class="eyebrow">'+esc(app)+'</div><h2>Um ambiente.<br>Mais possibilidades.</h2><p>Seus aplicativos e processos reunidos para um dia a dia mais simples e conectado.</p></div><div class="foot">Polystell Apps<br>Tecnologia a serviço das pessoas.</div></div>'+
   '<div class="panel2"><div class="form"><div class="eyebrow">'+esc(app)+'</div><h1>'+esc(title)+'</h1><p>'+esc(text||'Entre com a sua conta corporativa Microsoft.')+'</p>'+
   '<button class="btn primary ms" id="ps-login"><svg viewBox="0 0 21 21" width="18" height="18" aria-hidden="true"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>Entrar com Microsoft</button>'+
   '<div class="sec">'+icon('shield')+'<span>Use a mesma identidade do Microsoft 365.<br>O portal não solicita nem armazena sua senha.</span></div>'+
   '<div class="help">Precisa de ajuda? <a href="'+CFG.supportUrl+'" target="_blank" rel="noopener">Fale com a TI</a></div></div></div></section>';
  $('#ps-login').onclick=onLogin;}
};
/* ---------------- shell (menu + cabeçalho) ---------------- */
function mount({app,nav,current,onNav,account:acc}){
 const name=acc?.name||acc?.username||'Colaborador';const parts=name.split(/\s+/).filter(Boolean);
 const ini=(parts[0]?.[0]||'')+(parts.length>1?parts[parts.length-1][0]:'');
 document.body.innerHTML='<button class="ps-backdrop" id="ps-bd" hidden aria-label="Fechar menu"></button>'+
 '<aside class="ps-side" id="ps-side" aria-label="Navegação">'+brandHtml()+
 '<a class="ps-back" href="'+CFG.hubUrl+'">'+icon('back')+'Todos os aplicativos</a>'+
 '<nav class="ps-nav" id="ps-nav"></nav><div class="ps-sidefoot"><div id="ps-ctx"></div><span style="display:flex;gap:7px;align-items:center;padding:0 4px">'+icon('shield')+'Identidade Microsoft 365</span></div></aside>'+
 '<div class="ps-main"><header class="ps-top"><button class="ps-iconbtn ps-menubtn" id="ps-mb" aria-label="Abrir menu" aria-controls="ps-side" aria-expanded="false">'+icon('menu')+'</button>'+
 '<div class="ps-crumb"><a href="'+CFG.hubUrl+'" class="hide-sm">Polystell Apps</a><span class="hide-sm">'+icon('chevR')+'</span><span>'+esc(app)+'</span>'+icon('chevR')+'<b id="ps-cur"></b></div>'+
 '<div class="ps-topact"><span class="ps-sync" id="ps-sync"></span><a class="ps-iconbtn" href="'+CFG.supportUrl+'" target="_blank" rel="noopener" aria-label="Central de ajuda" title="Central de ajuda">'+icon('help')+'</a>'+
 '<button class="ps-user" id="ps-ub" aria-expanded="false" aria-controls="ps-um" aria-label="Conta"><span class="ps-av" id="ps-av"><span>'+esc(ini.toUpperCase())+'</span></span>'+icon('chevD')+'</button></div>'+
 '<div class="ps-umenu" id="ps-um" hidden><strong>'+esc(name)+'</strong><small>'+esc(acc?.username||'')+'</small><a href="'+CFG.hubUrl+'">'+icon('grid')+'Todos os aplicativos</a><button id="ps-out">'+icon('logout')+'Sair da conta</button></div></header>'+
 '<main class="ps-page" id="ps-page" tabindex="-1"></main></div>';
 let cur=current,counts={};
 const drawNav=()=>{let html='',sec=null;nav.filter(n=>!n.hidden).forEach(n=>{if(n.section&&n.section!==sec){sec=n.section;html+='<div class="lbl">'+esc(sec)+'</div>';}
  const c=counts[n.id];const cn=c==null?'':typeof c==='object'?(c.n?'<span class="n'+(c.alert?' alert':'')+'">'+c.n+'</span>':''):(c?'<span class="n">'+c+'</span>':'');
  html+='<button data-nav="'+n.id+'"'+(n.id===cur?' aria-current="page"':'')+'>'+icon(n.icon)+'<span class="t">'+esc(n.label)+'</span>'+cn+'</button>';});$('#ps-nav').innerHTML=html;};
 const mq=window.matchMedia('(max-width:820px)');
 const syncSide=()=>{const side=$('#ps-side'),open=side.classList.contains('open');side.inert=mq.matches&&!open;$('#ps-mb').setAttribute('aria-expanded',String(open));};
 const closeMenu=()=>{$('#ps-side').classList.remove('open');$('#ps-bd').hidden=true;syncSide();};
 const closeUser=()=>{$('#ps-um').hidden=true;$('#ps-ub').setAttribute('aria-expanded','false');};
 $('#ps-nav').onclick=e=>{const b=e.target.closest('[data-nav]');if(!b)return;closeMenu();onNav(b.dataset.nav);};
 $('#ps-mb').onclick=()=>{$('#ps-side').classList.add('open');$('#ps-bd').hidden=false;syncSide();const f=$('#ps-side a,#ps-side button');if(f)f.focus();};
 (mq.addEventListener?mq.addEventListener('change',syncSide):mq.addListener(syncSide));
 $('#ps-bd').onclick=closeMenu;
 $('#ps-ub').onclick=()=>{const m=$('#ps-um');m.hidden=!m.hidden;$('#ps-ub').setAttribute('aria-expanded',String(!m.hidden));};
 document.addEventListener('click',e=>{if(!e.target.closest('#ps-ub')&&!e.target.closest('#ps-um'))closeUser();});
 document.addEventListener('keydown',e=>{if(e.key==='Escape'){closeUser();if($('#ps-side').classList.contains('open')){closeMenu();$('#ps-mb').focus();}}});
 $('#ps-out').onclick=()=>auth.logout();
 photo().then(u=>{if(u)$('#ps-av').innerHTML='<img src="'+u+'" alt="">';});
 const api={
  page:$('#ps-page'),
  setCurrent(id,title){cur=id;$('#ps-cur').textContent=title||'';drawNav();},
  setCounts(c){counts=c||{};drawNav();},
  setSync(state,text){const s=$('#ps-sync');s.className='ps-sync'+(state==='bad'?' bad':'');s.innerHTML=icon(state==='bad'?'cloudoff':'refresh')+'<span>'+esc(text)+'</span>';s.title=text;},
  setCtx(html){$('#ps-ctx').innerHTML=html||'';},
  setNav(n){nav=n;drawNav();}
 };
 drawNav();syncSide();return api;
}
window.PS=Object.freeze({cfg:CFG,$,$$,esc,norm,val,fmt,todayKey,dateKey,icon,toast,dialog,auth,graph,sp,groups,photo,screens,mount,loadScript,sleep});
})();
