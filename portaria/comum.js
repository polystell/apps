/* =====================================================================
   PORTARIA — regras e dados compartilhados entre o PC (index.html) e o
   celular (leitura.html). Depende de /core/polystell.js.
   Princípios desta versão:
   • presença = movimentos com STATUS_MOV "Aberto" (nunca some por idade);
   • antes de gravar entrada, confere no servidor se já existe uma aberta;
   • gravações em duas etapas são reconciliadas e a mensagem diz o que
     realmente foi salvo;
   • a portaria NÃO lê o cadastro completo de colaboradores (salário/CPF):
     usa 00_Diretorio_Colaboradores (nome, matrícula, depto, situação).
   ===================================================================== */
(function(){
'use strict';
const {sp,norm,val,dateKey,todayKey}=PS;
const LISTAS={movP:'09_Mov_Pessoas',agend:'09_Agendamentos',prevV:'09_Prev_Veiculos',movV:'09_Mov_Veiculos',merc:'09_Mercadorias',ocor:'09_Ocorrencias',turnos:'09_Turnos',dir:'00_Diretorio_Colaboradores',params:'00_Parametros_Apps',transp:'00_Cadastro_Transportador'};
const OPCIONAIS=new Set(['transp']);
/* colunas que o app grava/lê — se faltar alguma, a abertura avisa qual */
const ESTRUTURA={
 movP:['TIPO','NOME','MATRICULA','EMPRESA','DOC_MASC','MOTIVO','MODAL','SEM_CRACHA','CRACHA','ENTRADA','SAIDA','PORTEIRO_ENT','PORTEIRO_SAI','AGENDAMENTO','STATUS_MOV','ANFITRIAO_NOME','ANFITRIAO_MATRICULA'],
 movV:['PLACA','MOTORISTA','EMPRESA','FINALIDADE','NF','PREVISAO','ENTRADA','SAIDA','AUT_EXPEDICAO','AUT_POR','STATUS_MOV'],
 agend:['CODIGO','TIPO','NOME','EMPRESA','DOC_MASC','DATA','HORA','MOTIVO','STATUS','ANFITRIAO_NOME','ANFITRIAO_MATRICULA'],
 prevV:['TIPO','DATA','HORA','TRANSPORTADORA','REMETENTE','NF','AREA','STATUS'],
 merc:['SETOR','REMETENTE','DESCRICAO','VOLUMES','CHEGADA','STATUS','RETIRADO_POR','RETIRADO_IDENT','RETIRADA_PROPRIA','DATA_RETIRADA','PORTEIRO_ENTREGA','DESTINATARIO_NOME','DESTINATARIO_MATRICULA'],
 ocor:['TIPO','GRAVIDADE','DESCRICAO','STATUS','REGISTRADO_POR'],
 turnos:['DE','PARA','RESUMO','RECADOS','CIENCIA_EM'],
 dir:['NOME','MATRICULA','DEPARTAMENTO','SITUACAO'],
 params:['APP','MODULO','CHAVE','VALOR','ATIVO','DESCRICAO']
};
/* índices obrigatórios: sem eles as consultas param ao passar de 5.000 itens */
const INDICES={movP:['STATUS_MOV','ENTRADA'],movV:['STATUS_MOV','ENTRADA'],agend:['DATA','CODIGO'],prevV:['DATA'],merc:['STATUS'],ocor:['STATUS'],dir:['MATRICULA'],turnos:['CIENCIA_EM']};
/* tipo exigido onde o filtro depende dele */
const TIPOS={dir:{MATRICULA:'text'},movP:{MATRICULA:'text',STATUS_MOV:'choice'},movV:{STATUS_MOV:'choice'}};
const TIPO_NOME={text:'Texto (uma linha)',choice:'Escolha'};
/* parâmetros editáveis na tela Segurança (TI) — padrão quando não há linha */
const PARAM_DEF=[
 {mod:'Funções',k:'MODAL_OBRIGATORIO',d:'Exigir o modal de transporte na entrada de colaborador (ESG)',t:'bool',v:'Sim'},
 {mod:'Funções',k:'ENVIA_QR_WA',d:'Enviar o QR do convite por WhatsApp ao visitante',t:'bool',v:'Sim'},
 {mod:'Funções',k:'AVISO_MERCADORIA',d:'Avisar chegada e entrega de mercadoria no WhatsApp',t:'bool',v:'Sim'},
 {mod:'Funções',k:'LIBERACAO_WHATSAPP',d:'Anfitrião libera a entrada respondendo 1/2 no WhatsApp',t:'bool',v:'Sim'},
 {mod:'Prazos',k:'ALERTA_PERMANENCIA',d:'Pessoa com entrada aberta há mais de (horas) vira pendência',t:'num',v:14},
 {mod:'Prazos',k:'ALERTA_PATIO',d:'Veículo no pátio há mais de (horas) vira pendência',t:'num',v:4},
 {mod:'Prazos',k:'ALERTA_MERCADORIA',d:'Mercadoria aguardando há mais de (dias) vira pendência',t:'num',v:2},
 {mod:'Prazos',k:'ESPERA_ANFITRIAO',d:'Espera do anfitrião antes de ligar (minutos)',t:'num',v:5},
 {mod:'Prazos',k:'RETENCAO_VISITANTES',d:'Retenção de registros de visitantes (meses)',t:'num',v:12}
];
const TURNO_EXPIRA_H=13;
let L={},PARAMS={},PORTEIROS=[],DIR=null,DIR_AT=0,TRANSP=null;
const now=()=>new Date().toISOString();
const isTrue=v=>['sim','1','true','s','yes'].includes(norm(v));
/* ---------- abertura ---------- */
async function boot(){
 const all=await sp.lists();
 const faltando=[];
 for(const[k,n]of Object.entries(LISTAS)){L[k]=all[n];if(!L[k]&&!OPCIONAIS.has(k))faltando.push('Lista '+n+' não encontrada (ou sem permissão de leitura)');}
 if(faltando.length)return {ok:false,faltando};
 const checks=await Promise.all(Object.entries(ESTRUTURA).map(async([k,cols])=>{try{const have=await sp.columns(L[k]);const out=[];
  cols.filter(c=>!have.has(c)).forEach(c=>out.push(LISTAS[k]+': falta a coluna '+c));
  (INDICES[k]||[]).filter(c=>have.has(c)&&!have.get(c).indexed).forEach(c=>out.push(LISTAS[k]+': crie o índice na coluna '+c));
  Object.entries(TIPOS[k]||{}).forEach(([c,t])=>{if(have.has(c)&&!have.get(c)[t])out.push(LISTAS[k]+': a coluna '+c+' precisa ser do tipo '+TIPO_NOME[t]);});
  return out;}catch(e){return [LISTAS[k]+': '+e.message];}}));
 checks.flat().forEach(x=>faltando.push(x));
 if(faltando.length)return {ok:false,faltando};
 await loadParams();
 return {ok:true};
}
/* ---------- parâmetros e porteiros (00_Parametros_Apps, APP = Portaria) ---------- */
async function loadParams(){
 const rows=await sp.items(L.params,"$expand=fields&$filter=fields/APP eq 'Portaria'&$top=999");
 PARAMS={};PORTEIROS=[];
 rows.filter(r=>norm(val(r.APP))==='portaria').forEach(r=>{
  const ativo=r.ATIVO!==false&&norm(val(r.ATIVO))!=='nao';
  if(norm(r.CHAVE)==='porteiro'){PORTEIROS.push({id:r.id,nome:String(r.VALOR||'').trim(),obs:r.DESCRICAO||'',ativo});return;}
  PARAMS[r.CHAVE]={id:r.id,valor:r.VALOR,ativo,mod:r.MODULO};
 });
 PORTEIROS.sort((a,b)=>a.nome.localeCompare(b.nome,'pt-BR'));
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
const desligado=c=>norm(c?.sit).includes('deslig');
function buscaColab(lista,q){const n=norm(q);if(n.length<2)return[];return lista.filter(c=>norm(c.nome).includes(n)||c.mat.includes(n)||norm(c.dep).includes(n)).slice(0,8);}
/* ---------- transportadoras (pesquisa → nome) ---------- */
async function transp(){if(TRANSP)return TRANSP;TRANSP={};if(!L.transp)return TRANSP;
 try{(await sp.items(L.transp,'$expand=fields($select=RAZAO_SOCIAL)&$top=999')).forEach(t=>TRANSP[t.id]=t.RAZAO_SOCIAL||'');}catch(e){console.warn('Transportadoras indisponíveis',e.message);}return TRANSP;}
/* ---------- turno / porteiro de plantão (compartilhado PC + celular) ---------- */
async function turnoAtual(){const rs=await sp.items(L.turnos,'$expand=fields&$orderby=fields/CIENCIA_EM desc&$top=1',1);const t=rs[0]||null;
 if(!t)return {turno:null,ativo:false};
 const idade=Date.now()-new Date(t.CIENCIA_EM||0).getTime();
 return {turno:t,ativo:!!t.PARA&&idade<TURNO_EXPIRA_H*3600e3};}
async function ultimosTurnos(n=12){return sp.items(L.turnos,'$expand=fields&$orderby=fields/CIENCIA_EM desc&$top='+n,n);}
async function abrirTurno({de,para,recados,resumo}){return sp.add(L.turnos,{DE:de||'',PARA:para,RECADOS:recados||'',RESUMO:resumo||'',CIENCIA_EM:now()});}
/* ---------- consultas operacionais ---------- */
const dayRange=k=>{const d=new Date(k+'T12:00:00');const a=new Date(d.getTime()-864e5).toISOString().slice(0,10),b=new Date(d.getTime()+864e5).toISOString().slice(0,10);return [a+'T00:00:00Z',b+'T23:59:59Z'];};
async function abertosP(){return sp.items(L.movP,"$expand=fields&$filter=fields/STATUS_MOV eq 'Aberto'&$top=500");}
async function abertosV(){return sp.items(L.movV,"$expand=fields&$filter=fields/STATUS_MOV eq 'Aberto'&$top=500");}
async function pessoasHoje(){const k=todayKey();return sp.items(L.movP,"$expand=fields&$filter=fields/ENTRADA ge '"+k+"T00:00:00Z'&$top=999");}
async function veicHoje(){const k=todayKey();return sp.items(L.movV,"$expand=fields&$filter=fields/ENTRADA ge '"+k+"T00:00:00Z'&$top=500");}
async function agendaHoje(){const k=todayKey(),[a,b]=dayRange(k);const rs=await sp.items(L.agend,"$expand=fields&$filter=fields/DATA ge '"+a+"' and fields/DATA le '"+b+"'&$top=500");return rs.filter(x=>dateKey(x.DATA)===k);}
async function agendaPorCodigo(cod){const c=String(cod).replace(/'/g,'').trim();return (await sp.items(L.agend,"$expand=fields&$filter=fields/CODIGO eq '"+encodeURIComponent(c)+"'&$top=3"))[0]||null;}
async function prevHoje(){const k=todayKey(),[a,b]=dayRange(k);const rs=await sp.items(L.prevV,"$expand=fields&$filter=fields/DATA ge '"+a+"' and fields/DATA le '"+b+"'&$top=500");return rs.filter(x=>dateKey(x.DATA)===k);}
async function mercAguardando(){return sp.items(L.merc,"$expand=fields&$filter=fields/STATUS eq 'Aguardando retirada'&$top=500");}
async function ocorAbertas(){return sp.items(L.ocor,"$expand=fields&$filter=fields/STATUS eq 'Aberta' or fields/STATUS eq 'Em análise'&$top=300");}
async function getItem(k,id){const s=await sp.site();const r=await PS.graph('/sites/'+s+'/lists/'+L[k]+'/items/'+id+'?$expand=fields');return {id:String(r.id),...(r.fields||{})};}
async function movP(id){const s=await sp.site();const r=await PS.graph('/sites/'+s+'/lists/'+L.movP+'/items/'+id+'?$expand=fields');return {id:String(r.id),...(r.fields||{})};}
async function movV(id){const s=await sp.site();const r=await PS.graph('/sites/'+s+'/lists/'+L.movV+'/items/'+id+'?$expand=fields');return {id:String(r.id),...(r.fields||{})};}
/* ---------- regras ---------- */
const lid=v=>v==null||v===''?'':String(v);
const abertoDoAgend=(abertos,agId)=>abertos.find(m=>lid(m.AGENDAMENTOLookupId)===String(agId));
const abertoDaMatricula=(abertos,mat)=>abertos.find(m=>String(m.MATRICULA||'').trim()===String(mat).trim()&&String(mat).trim()!=='');
const abertoDoNome=(abertos,nome)=>abertos.find(m=>norm(m.NOME)===norm(nome));
const abertoDaPlaca=(abertos,placa)=>abertos.find(m=>normPlaca(m.PLACA)===normPlaca(placa));
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
/* ---------- gravações ---------- */
function resumoTurno(d){
 const p=d.openP||[],v=d.openV||[],m=d.merc||[],o=d.ocor||[];
 const nomes=p.slice(0,12).map(x=>x.NOME).join(', ')+(p.length>12?'…':'');
 return ['Na planta: '+p.length+(p.length?' ('+nomes+')':''),'Veículos no pátio: '+v.length+(v.length?' ('+v.map(x=>x.PLACA).join(', ')+')':''),'Mercadorias aguardando: '+m.length,'Ocorrências abertas: '+o.length].join('\n');}
async function entradaPessoa(fields,{agId=null,mat=null,nome=null,permitirHomonimo=false}={}){
 const abertos=await abertosP();
 if(agId){const cur=await getItem('agend',agId);const r=regraConvite(cur,abertoDoAgend(abertos,agId));
  if(r.acao==='bloq')return {bloq:r.motivo};}
 const dup=agId?abertoDoAgend(abertos,agId):mat?abertoDaMatricula(abertos,mat):(!permitirHomonimo&&nome?abertoDoNome(abertos,nome):null);
 if(dup)return {dup};
 const novo=await sp.add(L.movP,{...fields,STATUS_MOV:'Aberto',ENTRADA:now()});
 let parcial=null;
 if(agId){try{await sp.patch(L.agend,agId,{STATUS:'Entrou'});}catch(e){parcial='A entrada foi registrada, mas o status do convite não foi atualizado — o sistema corrige na próxima atualização.';}}
 return {ok:true,novo,parcial};
}
async function saidaPessoa(mov,porteiro){
 const atual=await movP(mov.id);
 if(val(atual.STATUS_MOV)!=='Aberto')return {jaSaiu:true,mov:atual};
 await sp.patch(L.movP,mov.id,{SAIDA:now(),STATUS_MOV:'Encerrado',PORTEIRO_SAI:porteiro});
 let parcial=null;const ag=lid(atual.AGENDAMENTOLookupId);
 if(ag){try{await sp.patch(L.agend,ag,{STATUS:'Saiu'});}catch(e){parcial='A saída foi registrada, mas o status do convite não foi atualizado — o sistema corrige na próxima atualização.';}}
 return {ok:true,parcial};
}
async function entradaVeiculo(fields){const abertos=await abertosV();const dup=abertoDaPlaca(abertos,fields.PLACA);if(dup)return {dup};
 const novo=await sp.add(L.movV,{...fields,STATUS_MOV:'Aberto',ENTRADA:now(),AUT_EXPEDICAO:'Pendente'});return {ok:true,novo};}
async function autorizarVeiculo(id,quem,porteiro){const at=await movV(id);if(val(at.STATUS_MOV)!=='Aberto')return {jaSaiu:true};
 await sp.patch(L.movV,id,{AUT_EXPEDICAO:'Autorizado',AUT_POR:(quem+' (informado à portaria · '+porteiro+')').slice(0,250)});return {ok:true};}
async function saidaVeiculo(v,{porteiro,excecao=null}){
 const atual=await movV(v.id);
 if(val(atual.STATUS_MOV)!=='Aberto')return {jaSaiu:true,mov:atual};
 const coleta=val(atual.FINALIDADE)==='Coleta',aut=val(atual.AUT_EXPEDICAO)==='Autorizado';
 if(coleta&&!aut){
  if(!excecao||excecao.trim().length<10)return {precisaExcecao:true};
  /* a saída já carrega a marca de exceção (rastro garantido) e não se repete;
     a ocorrência vem em seguida — se falhar, o aviso pede o registro manual */
  await sp.patch(L.movV,v.id,{SAIDA:now(),STATUS_MOV:'Encerrado',AUT_EXPEDICAO:'Exceção',AUT_POR:('Exceção ('+porteiro+'): '+excecao.trim()).slice(0,250)});
  try{await sp.add(L.ocor,{TIPO:'Veículo',GRAVIDADE:'Média',STATUS:'Aberta',REGISTRADO_POR:porteiro,DESCRICAO:'Coleta liberada SEM autorização da Expedição registrada.\nPlaca: '+atual.PLACA+' · NF: '+(atual.NF||'—')+' · Motorista: '+(atual.MOTORISTA||'—')+' · Empresa: '+(atual.EMPRESA||'—')+'\nMotivo informado pela portaria: '+excecao.trim()});}
  catch(e){return {ok:true,excecao:true,parcial:'A saída foi registrada como exceção, mas a ocorrência não foi aberta. Registre-a em Ocorrências.'};}
  return {ok:true,excecao:true};
 }
 await sp.patch(L.movV,v.id,{SAIDA:now(),STATUS_MOV:'Encerrado'});return {ok:true};
}
/* corrige convites cuja segunda gravação falhou, nos dois sentidos:
   movimento aberto e convite ≠ "Entrou"  → "Entrou"
   convite "Entrou", sem aberto e com saída hoje → "Saiu"
   Antes de gravar, relê no servidor (evita corrigir durante a gravação do outro aparelho). */
async function reconciliar(agenda,abertos,movHoje){
 const cand=[];
 agenda.forEach(a=>{const st=val(a.STATUS);const ab=abertoDoAgend(abertos,a.id);
  if(ab&&st!=='Entrou')cand.push({id:a.id,to:'Entrou'});
  else if(!ab&&st==='Entrou'&&(movHoje||[]).some(m=>lid(m.AGENDAMENTOLookupId)===String(a.id)&&val(m.STATUS_MOV)==='Encerrado'))cand.push({id:a.id,to:'Saiu'});});
 if(!cand.length)return 0;
 const fresh=await abertosP();let n=0;
 for(const c of cand){const ab=!!abertoDoAgend(fresh,c.id);if((c.to==='Entrou')!==ab)continue;
  try{const cur=await getItem('agend',c.id);if(val(cur.STATUS)===c.to)continue;await sp.patch(L.agend,c.id,{STATUS:c.to});n++;}catch(e){console.warn('Reconciliação falhou',c.id,e.message);}}
 return n;
}
/* ---------- formatação / validação ---------- */
function maskDoc(v){const d=String(v||'').replace(/\D/g,'');if(!d)return String(v||'').trim();if(d.length===11)return '***.'+d.slice(3,6)+'.'+d.slice(6,9)+'-**';if(d.length>=7)return '*'.repeat(d.length-4)+d.slice(-4);return d;}
function normPlaca(p){return String(p||'').toUpperCase().replace(/[^A-Z0-9]/g,'');}
const placaValida=p=>/^[A-Z]{3}\d[A-Z0-9]\d{2}$/.test(normPlaca(p));
window.PORTARIA=Object.freeze({LISTAS,PARAM_DEF,boot,loadParams,param,on,saveParam,addPorteiro,setPorteiroAtivo,get porteiros(){return PORTEIROS;},
 dir,colabPorMatricula,desligado,buscaColab,transp,turnoAtual,ultimosTurnos,abrirTurno,resumoTurno,
 abertosP,abertosV,pessoasHoje,veicHoje,agendaHoje,agendaPorCodigo,prevHoje,mercAguardando,ocorAbertas,
 abertoDoAgend,abertoDaMatricula,abertoDoNome,abertoDaPlaca,regraConvite,situacaoPrev,
 entradaPessoa,saidaPessoa,entradaVeiculo,autorizarVeiculo,saidaVeiculo,reconciliar,
 add:(k,f)=>sp.add(L[k],f),get:getItem,patch:(k,id,f)=>sp.patch(L[k],id,f),maskDoc,normPlaca,placaValida,lid});
})();
