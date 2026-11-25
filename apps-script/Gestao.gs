// v0.7 — Gestão em homologação. Não habilita coleta pessoal nem urna real.
const NIVEIS_REGIMENTAIS_=["NACIONAL","REGIONAL","ARQUIDIOCESANO_DIOCESANO","PAROQUIAL_REPRESENTANTE"];
const COORD_HEADERS_=["ID_COORDENACAO","ID_UNIDADE","NOME_EQUIPE","NIVEL","TIPO_EQUIPE","STATUS","INICIO_ATIVIDADE","FONTE","VALIDACAO","CRIADO_EM","ATUALIZADO_EM"];
const QUEST_HEADERS_=["ID","BLOCO","PERGUNTA","TIPO","OPÇÕES / REGRA","OBRIGATORIEDADE PROPOSTA","QUEM PODE ACESSAR","FINALIDADE"];
function gestaoConfigAutorizada_(){
 const auth=autorizacaoTerritorial_();
 const cfg=typeof gestaoConfig_==="function"?gestaoConfig_():null;
 return auth&&cfg&&cfg.operacionalId&&cfg.questionarioId?cfg:null;
}
function validarEquipeRascunho_(input,data,existentes){
 const clean=(x,max)=>{if(typeof x!=="string"||!x.trim()||x.length>max||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(x.trim()))throw Error("Campo inválido.");return x.trim();};
 if(!input||typeof input!=="object")throw Error("Solicitação inválida.");
 const nivel=clean(input.nivel,40),unidade=clean(input.unidade,100),nome=clean(input.nome,180),fonte=clean(input.fonte,500);
 const valido=nivel==="REGIONAL"?unidade==="CO":nivel==="ARQUIDIOCESANO_DIOCESANO"?data.dioceses.some(x=>x.id===unidade&&x.id!=="D13"):nivel==="PAROQUIAL_REPRESENTANTE"?data.paroquias.some(x=>x.id===unidade):false;
 if(!valido)throw Error("Nível ou unidade fora do escopo regional de homologação.");
 if(existentes.some(r=>r[1]===unidade&&r[3]===nivel&&r[5]!=="ARQUIVADA"))throw Error("Já existe registro para este nível e unidade.");
 return {nivel,unidade,nome,fonte,tipo:nivel==="PAROQUIAL_REPRESENTANTE"?"REPRESENTACAO":"COORDENACAO"};
}
function consultarGestao(){
 try{
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId);
 const equipes=tabela_(ss,"COORDENACOES",COORD_HEADERS_).map(r=>({id:r[0],unidade:r[1],nome:r[2],nivel:r[3],tipo:r[4],status:r[5],fonte:r[7],validacao:r[8]}));
 const q=tabela_(SpreadsheetApp.openById(cfg.questionarioId),"PERGUNTAS",QUEST_HEADERS_).map(r=>({id:r[0],bloco:r[1],pergunta:r[2],tipo:r[3],opcoes:r[4],regra:r[5]}));
 const cargos=tabela_(ss,"CARGOS_REGIMENTAIS",["NIVEL_REFERENCIA","CARGO","ORIGEM","PROVIMENTO","OBSERVACAO"]).map(r=>({nivel:r[0],cargo:r[1],origem:r[2],provimento:r[3],obs:r[4]}));
 return {ok:true,equipes,questionario:q,cargos};
 }catch(e){return {ok:false,reason:"Estrutura de gestão indisponível."};}
}
function criarCoordenacaoRascunho(input){
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const territory=consultarTerritorio();if(!territory.ok)return territory;
 let lock;
 try{
 lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId),s=ss.getSheetByName("COORDENACOES");
 const existing=tabela_(ss,"COORDENACOES",COORD_HEADERS_);
 let v;try{v=validarEquipeRascunho_(input,territory.data,existing);}catch(e){return {ok:false,reason:e.message};}
 const id=Utilities.getUuid(),now=new Date().toISOString(),row=s.getLastRow()+1;
 if(row>1000)return {ok:false,reason:"Limite da homologação."};
 if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),Math.min(100,1000-s.getMaxRows()));
 s.getRange(row,1,1,COORD_HEADERS_.length).setNumberFormat("@").setValues([[id,v.unidade,v.nome,v.nivel,v.tipo,"RASCUNHO","",v.fonte,"PENDENTE DE VALIDACAO",now,now]]);
 SpreadsheetApp.flush();
 return {ok:true,id};
 }catch(e){return {ok:false,reason:"Não foi possível registrar o rascunho."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
