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

const PESSOAS_HEADERS_=["ID_PESSOA","NOME_COMPLETO","NOME_PREFERIDO","DATA_NASCIMENTO","MUNICIPIO_RA","UF","EMAIL","TELEFONE_WHATSAPP","CANAL_CONTATO","STATUS_CADASTRO","CRIADO_EM","ATUALIZADO_EM","PAIS","CEP","BAIRRO","LOGRADOURO","NUMERO","COMPLEMENTO","ID_UNIDADE_VINCULO_INICIAL","AUTORIZACAO_CADASTRO_EM","VERSAO_AVISO","ORIGEM_AUTORIZACAO"];
function validarCadastroBasico_(input,data,hoje){
 const clean=(v,max,required=false)=>{if(typeof v!=="string")throw Error("Campo inválido.");const s=v.trim();if((required&&!s)||s.length>max||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(s))throw Error("Confira os campos de cadastro.");return s;};
 if(!input||typeof input!=="object"||input.autorizado!==true)throw Error("É necessário confirmar autorização institucional e da pessoa para o cadastro.");
 const nome=clean(input.nome,180,true),preferido=clean(input.preferido,100),nascimento=clean(input.nascimento,10,true),cidade=clean(input.cidade,120,true),uf=clean(input.uf,2,true),email=clean(input.email,180),telefone=clean(input.telefone,25),canal=clean(input.canal,20),pais=clean(input.pais,80,true),cep=clean(input.cep,9),bairro=clean(input.bairro,120),logradouro=clean(input.logradouro,250),numero=clean(input.numero,30),complemento=clean(input.complemento,150),unidade=clean(input.unidade,100);
 const d=new Date(nascimento+"T00:00:00Z");
 if(!/^\d{4}-\d{2}-\d{2}$/.test(nascimento)||Number.isNaN(d.getTime())||d.toISOString().slice(0,10)!==nascimento||nascimento>hoje||nascimento<"1900-01-01")throw Error("Data de nascimento inválida.");
 let idade=Number(hoje.slice(0,4))-Number(nascimento.slice(0,4));if(hoje.slice(5)<nascimento.slice(5))idade--;
 if(idade<18)throw Error("O fluxo de cadastro de menores ainda não está habilitado.");
 if(pais!=="Brasil"||!["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"].includes(uf))throw Error("Nesta etapa, informe endereço brasileiro e UF válida.");
 if(cep&&!/^\d{5}-?\d{3}$/.test(cep))throw Error("CEP inválido.");
 if([cep,bairro,logradouro,numero,complemento].some(Boolean)&&![cep,bairro,logradouro,numero].every(Boolean))throw Error("Complete CEP, bairro, logradouro e número (ou s/n), ou deixe o endereço detalhado em branco.");
 if(email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw Error("E-mail inválido.");
 if(telefone&&!/^[0-9 ()-]{8,25}$/.test(telefone))throw Error("Telefone inválido.");
 if(!["","EMAIL","WHATSAPP","SEM_CONTATO"].includes(canal)||(canal==="EMAIL"&&!email)||(canal==="WHATSAPP"&&!telefone))throw Error("Confira o canal de contato.");
 if(unidade&&unidade!=="CO"&&![...data.dioceses,...data.paroquias,...data.comunidades].some(x=>x.id===unidade))throw Error("Unidade de vínculo não reconhecida.");
 return {nome,preferido,nascimento,cidade,uf,email,telefone,canal,pais,cep,bairro,logradouro,numero,complemento,unidade};
}
function salvarCadastroBasico(input){
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const t=consultarTerritorio();if(!t.ok)return t;
 let v;try{v=validarCadastroBasico_(input,t.data,Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd"));}catch(e){return {ok:false,reason:e.message};}
 let lock;try{
 lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Outra operação em andamento. Tente novamente."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId),s=ss.getSheetByName("PESSOAS");
 const rows=tabela_(ss,"PESSOAS",PESSOAS_HEADERS_);
 if(rows.some(r=>normalizarChave_(r[1])===normalizarChave_(v.nome)&&r[3]===v.nascimento))return {ok:false,reason:"Possível cadastro existente com mesmo nome e nascimento. Solicite revisão antes de cadastrar novamente."};
 const id=Utilities.getUuid(),now=new Date().toISOString(),row=s.getLastRow()+1;
 if(row>1000)return {ok:false,reason:"Limite do cadastro inicial."};
 if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),Math.min(100,1000-s.getMaxRows()));
 const values=[id,v.nome,v.preferido,v.nascimento,v.cidade,v.uf,v.email,v.telefone,v.canal,"CADASTRADO",now,now,v.pais,v.cep,v.bairro,v.logradouro,v.numero,v.complemento,v.unidade,now,"CAD-BASICO-0.10","DECLARACAO_DO_OPERADOR"];
 s.getRange(row,1,1,values.length).setNumberFormat("@").setValues([values]);SpreadsheetApp.flush();
 return {ok:true,id};
 }catch(e){return {ok:false,reason:"Não foi possível confirmar a gravação. Confira a planilha antes de tentar novamente."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
function consultarResumoCadastro(){
 try{
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId),r=tabela_(ss,"PESSOAS",PESSOAS_HEADERS_).filter(x=>x[9]==="CADASTRADO");
 const unidades=Object.create(null);for(const x of r){const key=x[18]||"SEM_VINCULO";unidades[key]=(unidades[key]||0)+1;}
 const equipes=tabela_(ss,"COORDENACOES",COORD_HEADERS_).filter(x=>x[5]!=="ARQUIVADA");
 return {ok:true,total:r.length,unidades:Object.entries(unidades).map(([id,total])=>({id,total})),equipes:equipes.length,rascunhos:equipes.filter(x=>x[5]==="RASCUNHO").length};
 }catch(e){return {ok:false,reason:"Indicadores de cadastro indisponíveis."};}
}
