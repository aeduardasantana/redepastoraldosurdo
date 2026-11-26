// Agenda Pastoral — homologação restrita à conta institucional.
// A agenda existente é configurada exclusivamente nas propriedades privadas.
const AGENDA_CATEGORIAS_=["CELEBRACAO","FORMACAO","REUNIAO","RETIRO","ROMARIA","ENCONTRO","ASSEMBLEIA","VISITA","CONFRATERNIZACAO","DATA_COMEMORATIVA","OUTRO"];
const AGENDA_ACESSIBILIDADE_=["LIBRAS_DIRETA","INTERPRETACAO_LIBRAS","AMBOS","A_CONFIRMAR","NAO_INFORMADA"];
const AGENDA_SITUACOES_=["PLANEJAMENTO","A_CONFIRMAR","CONFIRMADO","ADIADO","CANCELADO"];
function agendaContexto_(){
 const auth=autorizacaoTerritorial_();
 if(!auth)throw Error("Acesso restrito à conta institucional autorizada.");
 const props=PropertiesService.getScriptProperties();
 const id=String(props.getProperty("AGENDA_CALENDAR_ID")||"").trim();
 if(props.getProperty("AGENDA_HOMOLOGACAO")!=="SIM"||!id)throw Error("Integração da Agenda Pastoral ainda não habilitada.");
 return {id};
}
function agendaDataValida_(s){
 return typeof s==="string"&&/^\d{4}-\d{2}-\d{2}$/.test(s)&&!isNaN(Date.parse(s+"T00:00:00Z"))&&new Date(s+"T00:00:00Z").toISOString().slice(0,10)===s;
}
function agendaDiaSeguinte_(s){return new Date(Date.parse(s+"T00:00:00Z")+86400000).toISOString().slice(0,10);}
function agendaValidar_(v,unidades){
 if(!v||typeof v!=="object")throw Error("Solicitação inválida.");
 const clean=(s,max,req=false)=>{if(typeof s!=="string")throw Error("Campo inválido.");s=s.trim();if((req&&!s)||s.length>max||/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(s)||/[<>]/.test(s))throw Error("Preencha texto simples sem HTML.");return s;};
 const titulo=clean(v.titulo,180,true),local=clean(v.local||"",300),descricao=clean(v.descricao||"",2000),unidade=clean(v.unidade,100,true);
 if(!unidades.some(x=>x.id===unidade))throw Error("Unidade territorial não autorizada.");
 if(!AGENDA_CATEGORIAS_.includes(v.categoria)||!AGENDA_ACESSIBILIDADE_.includes(v.acessibilidade)||!AGENDA_SITUACOES_.includes(v.situacao))throw Error("Classificação inválida.");
 if(!agendaDataValida_(v.inicio)||!agendaDataValida_(v.fim)||v.fim<v.inicio)throw Error("Período inválido.");
 let start,end;
 if(v.diaInteiro===true){start={date:v.inicio};end={date:agendaDiaSeguinte_(v.fim)};}
 else{
  const hora=s=>typeof s==="string"&&/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(s);
  if(!hora(v.horaInicio)||!hora(v.horaFim))throw Error("Informe os dois horários.");
  const a=v.inicio+"T"+v.horaInicio+":00-03:00",b=v.fim+"T"+v.horaFim+":00-03:00";
  if(b<=a)throw Error("Horário final deve ser posterior ao inicial.");
  start={dateTime:a,timeZone:"America/Sao_Paulo"};end={dateTime:b,timeZone:"America/Sao_Paulo"};
 }
 // Não publicar automaticamente: todos os registros criados no módulo são privados.
 // Situação pastoral é separada da situação técnica do Google Calendar.
 return {summary:titulo,location:local,description:descricao,start,end,visibility:"private",transparency:"transparent",extendedProperties:{private:{redeAgenda:"1",unidade,categoria:v.categoria,acessibilidade:v.acessibilidade,situacao:v.situacao,publicacao:"NAO_PUBLICAR"}}};
}
function agendaUnidades_(){
 const r=consultarTerritorio();if(!r?.ok)throw Error("Cadastro territorial indisponível.");
 return [{id:"CO",nome:"Regional Centro-Oeste"},...r.data.dioceses.filter(x=>x.id!=="D13").map(x=>({id:x.id,nome:x.nome})),...r.data.paroquias.map(x=>({id:x.id,nome:x.nome+" — "+x.cidade})),...r.data.comunidades.map(x=>({id:x.id,nome:x.nome+" — "+x.cidade}))];
}
function agendaApi_(ctx,method,path,body,etag){
 const url="https://www.googleapis.com/calendar/v3/calendars/"+encodeURIComponent(ctx.id)+"/events"+path;
 const headers={Authorization:"Bearer "+ScriptApp.getOAuthToken()};
 if(etag)headers["If-Match"]=etag;
 const options={method,headers,muteHttpExceptions:true,followRedirects:false};
 if(body){options.contentType="application/json";options.payload=JSON.stringify(body);}
 const res=UrlFetchApp.fetch(url,options),code=res.getResponseCode();
 if(code===409||code===412)throw Error("Conflito ou evento já registrado. Atualize a consulta antes de repetir.");
 if(code<200||code>=300)throw Error("Google Agenda indisponível ou sem autorização de acesso.");
 return JSON.parse(res.getContentText());
}
function agendaMapear_(e){
 const p=e.extendedProperties?.private||{},managed=p.redeAgenda==="1";
 return {id:e.id,revisao:e.etag||"",titulo:e.summary||"(Sem título)",inicio:e.start?.date||e.start?.dateTime||"",fim:e.end?.date||e.end?.dateTime||"",diaInteiro:!!e.start?.date,local:e.location||"",gerenciado:managed,situacao:managed?p.situacao:"REFERENCIA_GOOGLE",unidade:managed?p.unidade:"",categoria:managed?p.categoria:"",acessibilidade:managed?p.acessibilidade:"",descricao:managed?e.description||"":"",recorrente:!!e.recurringEventId};
}
function consultarAgendaPastoral(inicio,fim){
 try{
  const ctx=agendaContexto_();
  if(!agendaDataValida_(inicio)||!agendaDataValida_(fim)||fim<inicio||(Date.parse(fim)-Date.parse(inicio))/86400000>92)throw Error("Selecione um período de até 93 dias.");
  const unidades=agendaUnidades_(),items=[];let token="",pages=0;
  do{
   if(++pages>20)throw Error("Limite de consulta excedido.");
   const query="?singleEvents=true&showDeleted=false&orderBy=startTime&maxResults=250&timeMin="+encodeURIComponent(inicio+"T00:00:00-03:00")+"&timeMax="+encodeURIComponent(agendaDiaSeguinte_(fim)+"T00:00:00-03:00")+(token?"&pageToken="+encodeURIComponent(token):"");
   const r=agendaApi_(ctx,"get",query);items.push(...(r.items||[]).filter(e=>e.status!=="cancelled").map(agendaMapear_));token=r.nextPageToken||"";
  }while(token);
  return {ok:true,eventos:items,unidades};
 }catch(e){return {ok:false,reason:"Consulta indisponível. Verifique autorização, configuração e período."};}
}
function salvarAgendaPastoral(input){
 let lock;
 try{
  const ctx=agendaContexto_(),unidades=agendaUnidades_(),body=agendaValidar_(input,unidades);
  lock=LockService.getScriptLock();if(!lock.tryLock(10000))return {ok:false,reason:"Outra operação em andamento. Atualize antes de repetir."};
  let r;
  if(input.id){
   if(typeof input.id!=="string"||!/^[a-zA-Z0-9_-]{1,1024}$/.test(input.id)||typeof input.revisao!=="string"||!input.revisao)throw Error("Identificador ou revisão inválida.");
   const path="/"+encodeURIComponent(input.id),old=agendaApi_(ctx,"get",path);
   if(old.extendedProperties?.private?.redeAgenda!=="1"||old.recurringEventId||old.recurrence?.length||old.etag!==input.revisao)throw Error("Evento externo, recorrente ou alterado. Atualize a consulta.");
   // PATCH preserva os demais campos não gerenciados e usa controle de concorrência.
   body.extendedProperties.private={...old.extendedProperties.private,...body.extendedProperties.private};
   r=agendaApi_(ctx,"patch",path+"?sendUpdates=none",body,old.etag);
  }else{
   if(body.extendedProperties.private.situacao==="CANCELADO")throw Error("Não crie evento já cancelado.");
   // ID determinístico para repetição da mesma solicitação; Calendar API aceita hex minúsculo.
   const hash=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(body),Utilities.Charset.UTF_8).map(x=>((x+256)%256).toString(16).padStart(2,"0")).join("");
   body.id="rp"+hash; // r,p e hexadecimal pertencem ao alfabeto base32hex.
   r=agendaApi_(ctx,"post","?sendUpdates=none",body);
  }
  if(!r.id||r.extendedProperties?.private?.redeAgenda!=="1")throw Error("Gravação não confirmada.");
  return {ok:true,id:r.id};
 }catch(e){return {ok:false,reason:"Gravação não confirmada. Confira os campos e atualize a agenda antes de repetir."};}
 finally{if(lock)try{lock.releaseLock();}catch(e){}}
}
