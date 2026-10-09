/* =====================================================================
   PORTARIA — regras e dados compartilhados entre o PC (index.html) e o
   celular (leitura.html). Depende de /core/polystell.js.  (v2.1)
   Garantias desta versão:
   • presença = movimentos com STATUS_MOV "Aberto" (nunca some por idade);
   • UMA entrada aberta por pessoa/convite/placa é garantida pelo SERVIDOR:
     a coluna CHAVE_ABERTA tem "valores exclusivos" — duas gravações
     simultâneas não conseguem gravar a mesma chave; na saída a chave é
     liberada (recebe "#<id>");
   • saídas usam If-Match (versão do item) — a segunda saída simultânea
     não sobrescreve a primeira;
   • situação do colaborador e validade do convite são reconferidas no
     servidor no momento da gravação;
   • leitura incompleta nunca é usada para concluir que alguém está fora;
   • a portaria NÃO lê o cadastro completo de colaboradores (salário/CPF):
     usa 00_Diretorio_Colaboradores.
   ===================================================================== */
(function(){
'use strict';
const {sp,norm,val,dateKey,todayKey}=PS;
const LISTAS={movP:'09_Mov_Pessoas',agend:'09_Agendamentos',prevV:'09_Prev_Veiculos',movV:'09_Mov_Veiculos',merc:'09_Mercadorias',ocor:'09_Ocorrencias',turnos:'09_Turnos',dir:'00_Diretorio_Colaboradores',params:'00_Parametros_Apps',transp:'00_Cadastro_Transportador'};
const OPCIONAIS=new Set(['transp']);
/* ---------- contrato com o SharePoint (conferido na abertura) ---------- */
const ESTRUTURA={
 movP:['TIPO','NOME','MATRICULA','EMPRESA','DOC_MASC','MOTIVO','MODAL','SEM_CRACHA','CRACHA','ENTRADA','SAIDA','PORTEIRO_ENT','PORTEIRO_SAI','AGENDAMENTO','STATUS_MOV','ANFITRIAO_NOME','ANFITRIAO_MATRICULA','ORIGEM','CHAVE_ABERTA'],
 movV:['PLACA','MOTORISTA','EMPRESA','FINALIDADE','NF','PREVISAO','ENTRADA','SAIDA','AUT_EXPEDICAO','AUT_POR','STATUS_MOV','PORTEIRO_ENT','PORTEIRO_SAI','CHAVE_ABERTA'],
 agend:['CODIGO','TIPO','NOME','EMPRESA','DOC_MASC','DATA','HORA','MOTIVO','STATUS','ANFITRIAO_NOME','ANFITRIAO_MATRICULA'],
 prevV:['TIPO','DATA','HORA','TRANSPORTADORA','REMETENTE','PLACA','MOTORISTA','NF','ORDEM','AREA','STATUS'],
 merc:['SETOR','REMETENTE','DESCRICAO','VOLUMES','CHEGADA','STATUS','RETIRADO_POR','RETIRADO_IDENT','RETIRADA_PROPRIA','DATA_RETIRADA','PORTEIRO_ENTREGA','PORTEIRO_RECEB','DESTINATARIO_NOME','DESTINATARIO_MATRICULA'],
 ocor:['TIPO','GRAVIDADE','DESCRICAO','STATUS','REGISTRADO_POR'],
 turnos:['DE','PARA','RESUMO','RECADOS','CIENCIA_EM'],
 dir:['NOME','MATRICULA','DEPARTAMENTO','SITUACAO','EMAIL'],
 params:['APP','MODULO','CHAVE','VALOR','ATIVO','DESCRICAO']
};
const INDICES={movP:['STATUS_MOV','ENTRADA','CHAVE_ABERTA'],movV:['STATUS_MOV','ENTRADA','CHAVE_ABERTA'],agend:['DATA','CODIGO'],prevV:['DATA'],merc:['STATUS'],ocor:['STATUS'],dir:['MATRICULA'],turnos:['CIENCIA_EM']};
const UNICOS={movP:['CHAVE_ABERTA'],movV:['CHAVE_ABERTA']};
const TIPOS={dir:{MATRICULA:'text'},movP:{MATRICULA:'text',STATUS_MOV:'choice',CHAVE_ABERTA:'text'},movV:{STATUS_MOV:'choice',CHAVE_ABERTA:'text'}};
const TIPO_NOME={text:'Texto (uma linha)',choice:'Escolha'};
const OPCOES={
 movP:{STATUS_MOV:['Aberto','Encerrado'],TIPO:['Colaborador','Diretoria','Visitante','Prestador','Terceiro'],MODAL:['Carro','Moto','Ônibus/Fretado','Bicicleta','A pé','Carona','Outro'],ORIGEM:['PC','Celular (QR)','Celular (digitado)']},
 movV:{STATUS_MOV:['Aberto','Encerrado'],FINALIDADE:['Coleta','Entrega','Serviço','Visita','Outro'],AUT_EXPEDICAO:['Pendente','Autorizado','Exceção']},
 agend:{STATUS:['Agendado','Na portaria','Entrou','Saiu','Não veio','Cancelado']},
 prevV:{TIPO:['Coleta','Entrega'],STATUS:['Previsto','Cancelado']},
 merc:{STATUS:['Aguardando retirada','Entregue']},
 ocor:{TIPO:['Segurança','Acesso indevido','Veículo','Mercadoria','Estrutura','Outro'],GRAVIDADE:['Baixa','Média','Alta'],STATUS:['Aberta','Em análise','Encerrada']},
 dir:{SITUACAO:['Ativo','Desligado']},
 params:{APP:['Portaria']}
};
const PESQUISAS={movP:{AGENDAMENTO:'agend'},movV:{PREVISAO:'prevV'},prevV:{TRANSPORTADORA:'transp'}};
/* situações do diretório que permitem ENTRAR (lista explícita) */
const SITUACAO_ENTRA=['Ativo'];
/* status administrativos do convite que a reconciliação nunca sobrescreve */
const ADMIN_FINAL=['Cancelado','Não veio'];
/* parâmetros editáveis na tela Segurança (TI); fase2 = depende de fluxo ainda não implantado */
const PARAM_DEF=[
 {mod:'Funções',k:'MODAL_OBRIGATORIO',d:'Exigir o modal de transporte na entrada de colaborador (ESG)',t:'bool',v:'Sim'},
 {mod:'Funções',k:'ENVIA_QR_WA',d:'Enviar o QR do convite por WhatsApp ao visitante',t:'bool',v:'Não',fase2:true},
 {mod:'Funções',k:'AVISO_MERCADORIA',d:'Avisar chegada e entrega de mercadoria no WhatsApp',t:'bool',v:'Não',fase2:true},
 {mod:'Funções',k:'LIBERACAO_WHATSAPP',d:'Anfitrião libera a entrada respondendo 1/2 no WhatsApp',t:'bool',v:'Não',fase2:true},
 {mod:'Prazos',k:'ALERTA_PERMANENCIA',d:'Pessoa com entrada aberta há mais de (horas) vira pendência',t:'num',v:14},
 {mod:'Prazos',k:'ALERTA_PATIO',d:'Veículo no pátio há mais de (horas) vira pendência',t:'num',v:4},
 {mod:'Prazos',k:'ALERTA_MERCADORIA',d:'Mercadoria aguardando há mais de (dias) vira pendência',t:'num',v:2},
 {mod:'Prazos',k:'ESPERA_ANFITRIAO',d:'Espera do anfitrião antes de ligar (minutos)',t:'num',v:5,fase2:true},
 {mod:'Prazos',k:'RETENCAO_VISITANTES',d:'Retenção de registros de visitantes (meses)',t:'num',v:12,fase2:true}
];
const TURNO_EXPIRA_H=13;
let L={},PARAMS={},PORTEIROS=[],PARAMS_AT=0,DIR=null,DIR_AT=0,TRANSP=null;
const now=()=>new Date().toISOString();
const isTrue=v=>['sim','1','true','s','yes'].includes(norm(v));
const guid=v=>String(v||'').replace(/[{}]/g,'').toLowerCase();
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const erroIncompleto=()=>Object.assign(new Error('Consulta de presença incompleta — operação interrompida para evitar registro duplicado. Avise a TI.'),{code:'incompleta'});
/* ---------- abertura ---------- */
/* listas que cada papel precisa enxergar (a permissão real é a da lista no SharePoint) */
const LISTAS_POSTO=['movP','agend','prevV','movV','merc','ocor','turnos','dir','params','transp'];
const LISTAS_LOG=['prevV','movV','params','transp'];
let COLS={};
async function boot({posto=true,log=false}={}){
 const all=await sp.lists();
 const precisa=new Set([...(posto?LISTAS_POSTO:[]),...(log?LISTAS_LOG:[])]);
 const faltando=[];
 for(const[k,n]of Object.entries(LISTAS)){L[k]=all[n];if(!L[k]&&precisa.has(k)&&!OPCIONAIS.has(k))faltando.push('Lista '+n+' não encontrada (ou sem permissão de leitura)');}
 if(faltando.length)return {ok:false,faltando};
 const checks=await Promise.all(Object.entries(ESTRUTURA).filter(([k])=>precisa.has(k)&&L[k]).map(async([k,cols])=>{try{const have=await sp.columns(L[k]);COLS[k]=have;const out=[],nm=LISTAS[k];
  cols.filter(c=>!have.has(c)).forEach(c=>out.push(nm+': falta a coluna '+c));
  (INDICES[k]||[]).filter(c=>have.has(c)&&!have.get(c).indexed).forEach(c=>out.push(nm+': crie o índice na coluna '+c));
  (UNICOS[k]||[]).filter(c=>have.has(c)&&!have.get(c).enforceUniqueValues).forEach(c=>out.push(nm+': ative "Exigir valores exclusivos" na coluna '+c));
  Object.entries(TIPOS[k]||{}).forEach(([c,t])=>{if(have.has(c)&&!have.get(c)[t])out.push(nm+': a coluna '+c+' precisa ser do tipo '+TIPO_NOME[t]);});
  Object.entries(OPCOES[k]||{}).forEach(([c,ops])=>{const col=have.get(c);if(!col||!col.choice)return;const tem=(col.choice.choices||[]);ops.filter(o=>!tem.includes(o)).forEach(o=>out.push(nm+': a coluna '+c+' precisa da opção "'+o+'"'));});
  Object.entries(PESQUISAS[k]||{}).forEach(([c,alvo])=>{const col=have.get(c);if(!col||!L[alvo])return;if(!col.lookup)out.push(nm+': a coluna '+c+' precisa ser do tipo Pesquisa');else if(guid(col.lookup.listId)!==guid(L[alvo]))out.push(nm+': a coluna '+c+' deve pesquisar na lista '+LISTAS[alvo]);});
  return out;}catch(e){return [LISTAS[k]+': '+e.message];}}));
 checks.flat().forEach(x=>faltando.push(x));
 if(faltando.length)return {ok:false,faltando};
 await loadParams();
 return {ok:true};
}
/* ---------- parâmetros e porteiros (00_Parametros_Apps, APP = Portaria) ---------- */
async function loadParams(){
 const rows=await sp.items(L.params,"$expand=fields&$filter=fields/APP eq 'Portaria'&$top=999");
 const P2={},PT=[];
 rows.filter(r=>norm(val(r.APP))==='portaria').forEach(r=>{
  const ativo=r.ATIVO!==false&&norm(val(r.ATIVO))!=='nao';
  if(norm(r.CHAVE)==='porteiro'){PT.push({id:r.id,nome:String(r.VALOR||'').trim(),obs:r.DESCRICAO||'',ativo});return;}
  P2[r.CHAVE]={id:r.id,valor:r.VALOR,ativo,mod:r.MODULO};});
 PARAMS=P2;PORTEIROS=PT.sort((a,b)=>a.nome.localeCompare(b.nome,'pt-BR'));PARAMS_AT=Date.now();
}
function param(k){const def=PARAM_DEF.find(p=>p.k===k);const r=PARAMS[k];const v=(r&&r.ativo&&r.valor!=null&&r.valor!=='')?r.valor:def?.v;return def?.t==='num'?Number(v)||Number(def.v):v;}
const on=k=>isTrue(param(k));
async function saveParam(k,valor){const def=PARAM_DEF.find(p=>p.k===k);const r=PARAMS[k];const v=String(valor);
 if(r)await sp.patch(L.params,r.id,{VALOR:v,ATIVO:true});
 else await sp.add(L.params,{APP:'Portaria',MODULO:def?.mod||'Geral',CHAVE:k,VALOR:v,ATIVO:true,DESCRICAO:def?.d||''});}
async function addPorteiro(nome,obs){return sp.add(L.params,{APP:'Portaria',MODULO:'Porteiros',CHAVE:'PORTEIRO',VALOR:nome,DESCRICAO:obs||'',ATIVO:true});}
async function setPorteiroAtivo(id,ativo){return sp.patch(L.params,id,{ATIVO:!!ativo});}
/* ---------- diretório de colaboradores (sem dados sensíveis) ---------- */
const mapDir=r=>({id:r.id,nome:String(r.NOME||'').trim(),mat:String(r.MATRICULA??'').trim(),dep:val(r.DEPARTAMENTO),sit:val(r.SITUACAO),email:r.EMAIL||''});
async function dir(force){if(!force&&DIR&&Date.now()-DIR_AT<20*60e3)return DIR;
 DIR=(await sp.items(L.dir,'$expand=fields($select=NOME,MATRICULA,DEPARTAMENTO,SITUACAO,EMAIL)&$top=999')).map(mapDir).filter(c=>c.nome);DIR_AT=Date.now();return DIR;}
async function colabPorMatricula(mat){const m=String(mat).replace(/'/g,'').trim();
 const rs=await sp.items(L.dir,"$expand=fields($select=NOME,MATRICULA,DEPARTAMENTO,SITUACAO,EMAIL)&$filter=fields/MATRICULA eq '"+encodeURIComponent(m)+"'&$top=5");
 return rs.map(mapDir)[0]||null;}
const podeEntrar=c=>!!c&&SITUACAO_ENTRA.some(s=>norm(s)===norm(c.sit));
const desligado=c=>!!c&&!podeEntrar(c);
function buscaColab(lista,q){const n=norm(q);if(n.length<2)return[];return lista.filter(c=>norm(c.nome).includes(n)||c.mat.includes(n)||norm(c.dep).includes(n)).slice(0,8);}
/* ---------- transportadoras (pesquisa → nome) ---------- */
let TRANSP_L=[];
async function transp(){if(TRANSP)return TRANSP;TRANSP={};if(!L.transp)return TRANSP;
 try{TRANSP_L=(await sp.items(L.transp,'$expand=fields&$top=999')).map(t=>({id:t.id,nome:String(t.RAZAO_SOCIAL||'').trim(),cnpj:String(t.CNPJ||'').trim()})).filter(t=>t.nome).sort((a,b)=>a.nome.localeCompare(b.nome,'pt-BR'));
  TRANSP_L.forEach(t=>TRANSP[t.id]=t.nome);}catch(e){console.warn('Transportadoras indisponíveis',e.message);}return TRANSP;}
/* ---------- turno / porteiro de plantão (compartilhado PC + celular) ---------- */
/* validade calculada NA HORA (não um booleano guardado): existe, tem porteiro e tem < 13 h */
const turnoValido=t=>!!t&&!!String(t.PARA||'').trim()&&Date.now()-new Date(t.CIENCIA_EM||0).getTime()<TURNO_EXPIRA_H*3600e3;
async function turnoAtual(){const rs=await sp.items(L.turnos,'$expand=fields&$orderby=fields/CIENCIA_EM desc&$top=1',1);const t=rs[0]||null;return {turno:t,ativo:turnoValido(t),lidoEm:Date.now()};}
async function ultimosTurnos(n=12){return sp.items(L.turnos,'$expand=fields&$orderby=fields/CIENCIA_EM desc&$top='+n,n);}
async function abrirTurno({de,para,recados,resumo}){return sp.add(L.turnos,{DE:de||'',PARA:para,RECADOS:recados||'',RESUMO:resumo||'',CIENCIA_EM:now()});}
/* ---------- consultas operacionais ---------- */
const keyShift=(k,d)=>{const x=new Date(k+'T12:00:00Z');x.setUTCDate(x.getUTCDate()+d);return x.toISOString().slice(0,10);};
async function abertosP(){return sp.items(L.movP,"$expand=fields&$filter=fields/STATUS_MOV eq 'Aberto'&$top=500");}
async function abertosV(){return sp.items(L.movV,"$expand=fields&$filter=fields/STATUS_MOV eq 'Aberto'&$top=500");}
const desdeHoje=()=>keyShift(todayKey(),-1)+'T03:00:00Z';
async function pessoasHoje(){const k=todayKey();const rs=await sp.items(L.movP,"$expand=fields&$filter=fields/ENTRADA ge '"+desdeHoje()+"'&$top=999");return rs.filter(x=>new Date(x.ENTRADA).toDateString()===new Date().toDateString()||(x.SAIDA&&new Date(x.SAIDA).toDateString()===new Date().toDateString()));}
async function veicHoje(){const rs=await sp.items(L.movV,"$expand=fields&$filter=fields/ENTRADA ge '"+desdeHoje()+"'&$top=999");return rs.filter(x=>new Date(x.ENTRADA).toDateString()===new Date().toDateString()||(x.SAIDA&&new Date(x.SAIDA).toDateString()===new Date().toDateString()));}
async function agendaEntre(dIni,dFim){const k=todayKey();const a=keyShift(k,dIni-1)+'T00:00:00Z',b=keyShift(k,dFim+1)+'T23:59:59Z';
 const rs=await sp.items(L.agend,"$expand=fields&$filter=fields/DATA ge '"+a+"' and fields/DATA le '"+b+"'&$top=999");
 const lo=keyShift(k,dIni),hi=keyShift(k,dFim);return rs.filter(x=>{const d=dateKey(x.DATA);return d>=lo&&d<=hi;});}
const agendaHoje=()=>agendaEntre(0,0);
const agendaRecente=()=>agendaEntre(-3,0);
async function agendaPorCodigo(cod){const c=String(cod).replace(/'/g,'').trim();return (await sp.items(L.agend,"$expand=fields&$filter=fields/CODIGO eq '"+encodeURIComponent(c)+"'&$top=3"))[0]||null;}
async function prevEntre(dIni,dFim){const k=todayKey();const a=keyShift(k,dIni-1)+'T00:00:00Z',b=keyShift(k,dFim+1)+'T23:59:59Z';
 const rs=await sp.items(L.prevV,"$expand=fields&$filter=fields/DATA ge '"+a+"' and fields/DATA le '"+b+"'&$top=999");
 const lo=keyShift(k,dIni),hi=keyShift(k,dFim);return rs.filter(x=>{const d=dateKey(x.DATA);return d>=lo&&d<=hi;});}
const prevHoje=()=>prevEntre(0,0);
async function mercAguardando(){return sp.items(L.merc,"$expand=fields&$filter=fields/STATUS eq 'Aguardando retirada'&$top=500");}
async function ocorAbertas(){return sp.items(L.ocor,"$expand=fields&$filter=fields/STATUS eq 'Aberta' or fields/STATUS eq 'Em análise'&$top=300");}
const getItem=(k,id)=>sp.get(L[k],id);
/* ---------- regras ---------- */
const lid=v=>v==null||v===''?'':String(v);
const abertoDoAgend=(abertos,agId)=>abertos.find(m=>lid(m.AGENDAMENTOLookupId)===String(agId));
const abertoDaMatricula=(abertos,mat)=>abertos.find(m=>String(m.MATRICULA||'').trim()===String(mat).trim()&&String(mat).trim()!=='');
const abertoDoNome=(abertos,nome)=>abertos.find(m=>norm(m.NOME)===norm(nome));
const abertoDaPlaca=(abertos,placa)=>abertos.find(m=>normPlaca(m.PLACA)===normPlaca(placa));
const abertoDaChave=(abertos,chave)=>abertos.find(m=>norm(m.CHAVE_ABERTA)===norm(chave));
/* convite: vale o dia todo; Cancelado / Não veio / outra data = bloqueado */
function regraConvite(a,aberto){
 const st=val(a.STATUS)||'Agendado';
 if(aberto)return {acao:'saida',st};
 if(st==='Cancelado')return {acao:'bloq',st,motivo:'Convite cancelado pelo anfitrião. Não libere; se a visita for confirmada, registre como visitante no computador.'};
 if(st==='Não veio')return {acao:'bloq',st,motivo:'Convite encerrado como "Não veio". Registre como visitante no computador, se for o caso.'};
 if(dateKey(a.DATA)!==todayKey())return {acao:'bloq',st,motivo:'Convite válido para '+PS.fmt.data(dateKey(a.DATA))+' — não é para hoje. Registre como visitante no computador, se a visita for confirmada.'};
 if(st==='Saiu'||st==='Entrou')return {acao:'entrada',st,reentrada:true};
 if(st==='Agendado'||st==='Na portaria')return {acao:'entrada',st};
 return {acao:'bloq',st,motivo:'Situação do convite não reconhecida ('+st+'). Confira com o anfitrião.'};
}
/* situação de uma previsão calculada pelo movimento (a portaria não edita a previsão da Expedição) */
function situacaoPrev(p,movsHoje){const ms=movsHoje.filter(m=>lid(m.PREVISAOLookupId)===String(p.id));
 if(ms.some(m=>val(m.STATUS_MOV)==='Aberto'))return 'Na portaria';if(ms.length)return 'Concluído';return val(p.STATUS)||'Previsto';}
function resumoTurno(d){
 const p=d.openP||[],v=d.openV||[],m=d.merc||[],o=d.ocor||[];
 const nomes=p.slice(0,12).map(x=>x.NOME).join(', ')+(p.length>12?'…':'');
 return ['Na planta: '+p.length+(p.length?' ('+nomes+')':''),'Veículos no pátio: '+v.length+(v.length?' ('+v.map(x=>x.PLACA).join(', ')+')':''),'Mercadorias aguardando: '+m.length,'Ocorrências abertas: '+o.length].join('\n');}
/* ---------- gravações ---------- */
const isDupErr=e=>e&&(e.status===409||/duplica|unique|exclusiv|already exists|valores duplicados/i.test(String(e.detail||'')+' '+String(e.message||'')));
/* cria o movimento; se o servidor recusar ou não confirmar, relê pela chave:
   – recusa por duplicidade → outro aparelho gravou antes → {dup}
   – sem confirmação e a chave está lá → a gravação valeu → {confirmado} */
async function criarMovimento(k,fields,chave,relerAbertos){
 try{return {novo:await sp.add(L[k],fields)};}
 catch(e){let fresh;try{fresh=await relerAbertos();}catch{throw e;}
  const m=abertoDaChave(fresh,chave);if(!m)throw e;
  return isDupErr(e)?{dup:m}:{novo:m,confirmado:true};}
}
async function statusConvite(agId,to){
 for(let t=0;t<2;t++){try{const cur=await getItem('agend',agId);const st=val(cur.STATUS);
   if(st===to||ADMIN_FINAL.includes(st))return true;
   await sp.patch(L.agend,agId,{STATUS:to});return true;}catch(e){if(t===0)await sleep(1200);}}
 return false;
}
/* PATCH com versão; 412 → relê: se já encerrado, foi o outro aparelho */
async function patchVersao(k,id,fields,etag,aindaValido){
 try{await sp.patch(L[k],id,fields,{etag});return {ok:true};}
 catch(e){if(e.status!==412)throw e;const at=await getItem(k,id);if(!aindaValido(at))return {conflito:true,mov:at};
  await sp.patch(L[k],id,fields,{etag:at._etag});return {ok:true};}
}
async function entradaPessoa(fields,{agId=null,mat=null,nome=null,permitirHomonimo=false}={}){
 const abertos=await abertosP();if(abertos.truncated)throw erroIncompleto();
 if(mat){const c=await colabPorMatricula(mat);
  if(!c)return {bloq:'Matrícula '+mat+' não encontrada no diretório. Não libere; confira com o RH.'};
  if(!podeEntrar(c))return {bloq:'Situação no diretório: '+(c.sit||'não informada')+'. Entrada não permitida — comunique a Segurança.'};}
 if(agId){const cur=await getItem('agend',agId);const r=regraConvite(cur,abertoDoAgend(abertos,agId));if(r.acao==='bloq')return {bloq:r.motivo};}
 const dup=agId?abertoDoAgend(abertos,agId):mat?abertoDaMatricula(abertos,mat):(!permitirHomonimo&&nome?abertoDoNome(abertos,nome):null);
 if(dup)return {dup};
 const chave=agId?'A:'+agId:mat?'M:'+mat:'N:'+norm(nome||fields.NOME)+'|'+norm(fields.EMPRESA)+(permitirHomonimo?'#'+Date.now().toString(36):'');
 const res=await criarMovimento('movP',{...fields,STATUS_MOV:'Aberto',ENTRADA:now(),CHAVE_ABERTA:chave},chave,abertosP);
 if(res.dup)return {dup:res.dup};
 let parcial=null;
 if(agId&&!(await statusConvite(agId,'Entrou')))parcial='Entrada registrada. A atualização do convite ficou pendente — o computador da portaria corrige automaticamente.';
 return {ok:true,novo:res.novo,confirmado:!!res.confirmado,parcial};
}
async function saidaPessoa(mov,porteiro){
 const atual=await getItem('movP',mov.id);
 if(val(atual.STATUS_MOV)!=='Aberto')return {jaSaiu:true,mov:atual};
 const r=await patchVersao('movP',mov.id,{SAIDA:now(),STATUS_MOV:'Encerrado',PORTEIRO_SAI:porteiro,CHAVE_ABERTA:(atual.CHAVE_ABERTA||'')+'#'+mov.id},atual._etag,at=>val(at.STATUS_MOV)==='Aberto');
 if(r.conflito)return {jaSaiu:true,mov:r.mov};
 let parcial=null;const ag=lid(atual.AGENDAMENTOLookupId);
 if(ag&&!(await statusConvite(ag,'Saiu')))parcial='Saída registrada. A atualização do convite ficou pendente — o computador da portaria corrige automaticamente.';
 return {ok:true,parcial};
}
async function entradaVeiculo(fields,{porteiro}){
 const abertos=await abertosV();if(abertos.truncated)throw erroIncompleto();
 const placa=normPlaca(fields.PLACA);const dup=abertoDaPlaca(abertos,placa);if(dup)return {dup};
 const chave='P:'+placa;
 const res=await criarMovimento('movV',{...fields,PLACA:placa,PORTEIRO_ENT:porteiro,STATUS_MOV:'Aberto',ENTRADA:now(),AUT_EXPEDICAO:'Pendente',CHAVE_ABERTA:chave},chave,abertosV);
 if(res.dup)return {dup:res.dup};return {ok:true,novo:res.novo,confirmado:!!res.confirmado};}
async function autorizarVeiculo(id,quem,porteiro){const at=await getItem('movV',id);if(val(at.STATUS_MOV)!=='Aberto')return {jaSaiu:true};
 const r=await patchVersao('movV',id,{AUT_EXPEDICAO:'Autorizado',AUT_POR:(quem+' (informado à portaria · '+porteiro+')').slice(0,250)},at._etag,x=>val(x.STATUS_MOV)==='Aberto');
 return r.conflito?{jaSaiu:true}:{ok:true};}
async function saidaVeiculo(v,{porteiro,excecao=null}){
 const atual=await getItem('movV',v.id);
 if(val(atual.STATUS_MOV)!=='Aberto')return {jaSaiu:true,mov:atual};
 const coleta=val(atual.FINALIDADE)==='Coleta',aut=val(atual.AUT_EXPEDICAO)==='Autorizado';
 const base={SAIDA:now(),STATUS_MOV:'Encerrado',PORTEIRO_SAI:porteiro,CHAVE_ABERTA:(atual.CHAVE_ABERTA||'')+'#'+v.id};
 const aberto=x=>val(x.STATUS_MOV)==='Aberto';
 if(coleta&&!aut){
  if(!excecao||excecao.trim().length<10)return {precisaExcecao:true};
  /* a saída já carrega a marca de exceção (rastro garantido) e não se repete;
     a ocorrência vem em seguida — se falhar, o aviso pede o registro manual */
  const r=await patchVersao('movV',v.id,{...base,AUT_EXPEDICAO:'Exceção',AUT_POR:('Exceção ('+porteiro+'): '+excecao.trim()).slice(0,250)},atual._etag,aberto);
  if(r.conflito)return {jaSaiu:true,mov:r.mov};
  try{await sp.add(L.ocor,{TIPO:'Veículo',GRAVIDADE:'Média',STATUS:'Aberta',REGISTRADO_POR:porteiro,DESCRICAO:'Coleta liberada SEM autorização da Expedição registrada. Responsável pela liberação: '+porteiro+' (porteiro de plantão).\nPlaca: '+atual.PLACA+' · NF: '+(atual.NF||'—')+' · Motorista: '+(atual.MOTORISTA||'—')+' · Empresa: '+(atual.EMPRESA||'—')+'\nMotivo informado: '+excecao.trim()});}
  catch(e){return {ok:true,excecao:true,parcial:'A saída foi registrada como exceção, mas a ocorrência não foi aberta. Registre-a em Ocorrências.'};}
  return {ok:true,excecao:true};
 }
 const r=await patchVersao('movV',v.id,base,atual._etag,aberto);if(r.conflito)return {jaSaiu:true,mov:r.mov};
 return {ok:true};
}
/* ---------- logística (Expedição) ---------- */
/* data só-dia gravada ao meio-dia local: não muda de dia por fuso */
const diaParaGravar=k=>{const [y,m,d]=k.split('-').map(Number);return new Date(y,m-1,d,12,0,0).toISOString();};
const opcoes=(k,c)=>{const col=COLS[k]&&COLS[k].get(c);return col&&col.choice?(col.choice.choices||[]):null;};
function validarPrev(f){const t=f.TIPO;
 if(t!=='Coleta'&&t!=='Entrega')return 'Escolha Coleta ou Entrega.';
 if(!/^\d{4}-\d{2}-\d{2}$/.test(f.dia||''))return 'Informe a data.';
 if(f.dia<todayKey())return 'A data não pode ser anterior a hoje.';
 if(f.HORA&&!/^([01]\d|2[0-3]):[0-5]\d$/.test(f.HORA))return 'Hora no formato 00:00.';
 if(t==='Coleta'&&!f.transp)return 'Na coleta, escolha a transportadora do cadastro.';
 if(t==='Entrega'&&!String(f.REMETENTE||'').trim())return 'Na entrega, informe o remetente (quem envia).';
 if(f.PLACA&&!placaValida(f.PLACA))return 'Placa fora do padrão ABC1D23 / ABC1234.';
 return null;}
async function salvarPrev(f,{id=null,etag=null}={}){const erro=validarPrev(f);if(erro)throw Object.assign(new Error(erro),{code:'validacao'});
 const fields={TIPO:f.TIPO,DATA:diaParaGravar(f.dia),HORA:f.HORA||'',
  PLACA:normPlaca(f.PLACA),MOTORISTA:String(f.MOTORISTA||'').trim(),REMETENTE:f.TIPO==='Entrega'?String(f.REMETENTE||'').trim():''};
 if(String(f.AREA||'').trim())fields.AREA=String(f.AREA).trim();
 /* NF e ORDEM respeitam o tipo real da coluna (texto ou número) */
 for(const c of ['NF','ORDEM']){const v=String(f[c]||'').trim();const num=COLS.prevV&&COLS.prevV.get(c)&&COLS.prevV.get(c).number;
  if(num){if(v!==''){if(!/^\d+([.,]\d+)?$/.test(v))throw Object.assign(new Error((c==='NF'?'Nota fiscal':'Ordem')+': use só números.'),{code:'validacao'});fields[c]=Number(v.replace(',','.'));}}else fields[c]=v;}
 if(f.TIPO==='Coleta')fields.TRANSPORTADORALookupId=Number(f.transp);
 if(!id){fields.STATUS='Previsto';return {novo:await sp.add(L.prevV,fields)};}
 const cur=await getItem('prevV',id);if(val(cur.STATUS)==='Cancelado')return {cancelada:true};
 if(val(cur.TIPO)!==f.TIPO)throw Object.assign(new Error('Para trocar entre coleta e entrega, cancele esta previsão e lance outra.'),{code:'validacao'});
 try{await sp.patch(L.prevV,id,fields,{etag:etag||cur._etag});}catch(e){if(e.status===412)return {conflito:true};throw e;}
 return {ok:true};}
async function cancelarPrev(id,etag){const cur=await getItem('prevV',id);if(val(cur.STATUS)==='Cancelado')return {ok:true};
 try{await sp.patch(L.prevV,id,{STATUS:'Cancelado'},{etag:etag||cur._etag});}catch(e){if(e.status===412)return {conflito:true};throw e;}return {ok:true};}
/* autorização dada pela própria Expedição, com o login dela */
async function autorizarPelaExpedicao(id,quem,obs){const at=await getItem('movV',id);
 if(val(at.STATUS_MOV)!=='Aberto')return {jaSaiu:true};if(val(at.AUT_EXPEDICAO)==='Autorizado')return {jaAut:true,mov:at};
 const hm=new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'});
 const r=await patchVersao('movV',id,{AUT_EXPEDICAO:'Autorizado',AUT_POR:('Expedição (no app): '+quem+' às '+hm+(obs?' — '+obs:'')).slice(0,250)},at._etag,x=>val(x.STATUS_MOV)==='Aberto'&&val(x.AUT_EXPEDICAO)!=='Autorizado');
 return r.conflito?(val(r.mov.STATUS_MOV)!=='Aberto'?{jaSaiu:true}:{jaAut:true,mov:r.mov}):{ok:true};}
/* corrige o status de PRESENÇA dos convites cuja 2ª gravação falhou (últimos 3 dias):
   movimento aberto e convite ≠ "Entrou" → "Entrou";  convite "Entrou" sem aberto → "Saiu".
   Relê no servidor antes de gravar; nunca sobrescreve Cancelado / Não veio (situação administrativa). */
async function reconciliar(agenda,abertos){
 const cand=[];
 agenda.forEach(a=>{const st=val(a.STATUS);if(ADMIN_FINAL.includes(st))return;const ab=abertoDoAgend(abertos,a.id);
  if(ab&&st!=='Entrou')cand.push({id:a.id,to:'Entrou'});else if(!ab&&st==='Entrou')cand.push({id:a.id,to:'Saiu'});});
 if(!cand.length)return 0;
 const fresh=await abertosP();if(fresh.truncated)return 0;let n=0;
 for(const c of cand){if((c.to==='Entrou')!==!!abertoDoAgend(fresh,c.id))continue;
  try{const cur=await getItem('agend',c.id);const st=val(cur.STATUS);if(st===c.to||ADMIN_FINAL.includes(st))continue;await sp.patch(L.agend,c.id,{STATUS:c.to},{etag:cur._etag});n++;}catch(e){console.warn('Reconciliação adiada',c.id,e.message);}}
 return n;
}
/* ---------- formatação / validação ---------- */
function maskDoc(v){const d=String(v||'').replace(/\D/g,'');if(!d)return String(v||'').trim();if(d.length===11)return '***.'+d.slice(3,6)+'.'+d.slice(6,9)+'-**';if(d.length>=7)return '*'.repeat(d.length-4)+d.slice(-4);return d;}
function normPlaca(p){return String(p||'').toUpperCase().replace(/[^A-Z0-9]/g,'');}
const placaValida=p=>/^[A-Z]{3}\d[A-Z0-9]\d{2}$/.test(normPlaca(p));
window.PORTARIA=Object.freeze({LISTAS,PARAM_DEF,TURNO_EXPIRA_H,boot,loadParams,param,on,saveParam,addPorteiro,setPorteiroAtivo,get porteiros(){return PORTEIROS;},get paramsEm(){return PARAMS_AT;},
 dir,colabPorMatricula,podeEntrar,desligado,buscaColab,transp,turnoAtual,turnoValido,ultimosTurnos,abrirTurno,resumoTurno,
 abertosP,abertosV,pessoasHoje,veicHoje,agendaHoje,agendaRecente,agendaPorCodigo,prevHoje,prevEntre,mercAguardando,ocorAbertas,
 get transportadoras(){return TRANSP_L;},diaMais:keyShift,opcoes,validarPrev,salvarPrev,cancelarPrev,autorizarPelaExpedicao,
 abertoDoAgend,abertoDaMatricula,abertoDoNome,abertoDaPlaca,regraConvite,situacaoPrev,
 entradaPessoa,saidaPessoa,entradaVeiculo,autorizarVeiculo,saidaVeiculo,reconciliar,
 add:(k,f)=>sp.add(L[k],f),get:getItem,patch:(k,id,f,o)=>sp.patch(L[k],id,f,o),maskDoc,normPlaca,placaValida,lid});
})();
