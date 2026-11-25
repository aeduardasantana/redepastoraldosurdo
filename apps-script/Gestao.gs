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
 const referencias=tabela_(ss,"COMPOSICAO_REFERENCIAS",REF_COORD_HEADERS_).map(r=>({id:r[0],equipe:r[1],nome:r[2],setor:r[3],funcao:r[4],origem:r[5],vigencia:r[6],fonte:r[7],validacao:r[8],vinculado:!!r[9]}));
 return {ok:true,equipes,questionario:q,cargos,referencias};
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

const PESSOAS_HEADERS_=["ID_PESSOA","NOME_COMPLETO","NOME_PREFERIDO","DATA_NASCIMENTO","MUNICIPIO_RA","UF","EMAIL","TELEFONE_WHATSAPP","CANAL_CONTATO","STATUS_CADASTRO","CRIADO_EM","ATUALIZADO_EM","PAIS","CEP","BAIRRO","LOGRADOURO","NUMERO","COMPLEMENTO","ID_UNIDADE_VINCULO_INICIAL","AUTORIZACAO_CADASTRO_EM","VERSAO_AVISO","ORIGEM_AUTORIZACAO","CPF_TOKEN"];
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
 const cpf=input.cpf==null?"":normalizarCpf_(input.cpf);return {nome,preferido,nascimento,cidade,uf,email,telefone,canal,pais,cep,bairro,logradouro,numero,complemento,unidade,cpf};
}
function salvarCadastroBasico(input){
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const t=consultarTerritorio();if(!t.ok)return t;
 let v;try{v=validarCadastroBasico_(input,t.data,Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd"));}catch(e){return {ok:false,reason:e.message};}
 let lock;try{
 lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Outra operação em andamento. Tente novamente."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId),s=ss.getSheetByName("PESSOAS");
 const rows=tabela_(ss,"PESSOAS",PESSOAS_HEADERS_);
 const token=v.cpf?cpfToken_(v.cpf,rows):"";
 if(token&&rows.some(r=>r[22]===token))return {ok:false,reason:"Possível cadastro existente. Encaminhe para revisão administrativa."};
 if(rows.some(r=>normalizarChave_(r[1])===normalizarChave_(v.nome)&&r[3]===v.nascimento))return {ok:false,reason:"Possível cadastro existente com mesmo nome e nascimento. Solicite revisão antes de cadastrar novamente."};
 const id=Utilities.getUuid(),now=new Date().toISOString(),row=s.getLastRow()+1;
 if(row>1000)return {ok:false,reason:"Limite do cadastro inicial."};
 if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),Math.min(100,1000-s.getMaxRows()));
 const values=[id,v.nome,v.preferido,v.nascimento,v.cidade,v.uf,v.email,v.telefone,v.canal,"CADASTRADO",now,now,v.pais,v.cep,v.bairro,v.logradouro,v.numero,v.complemento,v.unidade,now,"CAD-BASICO-0.11","DECLARACAO_DO_OPERADOR",token];
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

function normalizarCpf_(value){
 if(typeof value!=="string")throw Error("CPF inválido.");
 const raw=value.trim();if(!raw)return "";
 if(!/^[0-9.\- ]{11,18}$/.test(raw))throw Error("CPF inválido.");
 const n=raw.replace(/\D/g,"");if(n.length!==11||/^(\d)\1{10}$/.test(n))throw Error("CPF inválido.");
 for(let k=9;k<11;k++){let sum=0;for(let i=0;i<k;i++)sum+=Number(n[i])*(k+1-i);const digit=(sum*10)%11;if(Number(n[k])!==(digit===10?0:digit))throw Error("CPF inválido.");}
 return n;
}
function cpfToken_(cpf,rows){
 const prop=PropertiesService.getScriptProperties();let key=prop.getProperty("CPF_HMAC_KEY");
 if(!key){if(rows.some(r=>r[22]))throw Error("Configuração de identificação indisponível.");key=Utilities.getUuid()+Utilities.getUuid();prop.setProperty("CPF_HMAC_KEY",key);}
 return Utilities.computeHmacSha256Signature(cpf,key).map(b=>((b+256)%256).toString(16).padStart(2,"0")).join("");
}
const PRE_HEADERS_=["ID_PRE_CADASTRO","NOME_CONFORME_FONTE","CONTATO_CONFORME_FONTE","FUNCAO_INFORMADA_NA_FONTE","LOCAL_INFORMADO","CIDADE_RA","SITUACAO_NA_FONTE","DATA_NA_FONTE","FONTE","REFERENCIA_LINHA","STATUS_REVISAO","ID_PESSOA_CONFIRMADO","NOME_PADRONIZADO","CONTATO_PADRONIZADO","DDD_REFERENCIA","ID_DIOCESE","ID_UNIDADE","ID_LOCAL_REFERENCIA","VALIDACAO_CONTATO"];
function consultarPreCadastros(){
 try{const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const rows=tabela_(SpreadsheetApp.openById(cfg.operacionalId),"PRE_CADASTROS",PRE_HEADERS_);
 return {ok:true,registros:rows.filter(r=>r[10]==="PENDENTE_IDENTIFICACAO").map(r=>({id:r[0],nome:r[12]||r[1],contato:r[13]||r[2],funcao:r[3],local:r[4],cidade:r[5],situacao:r[6],data:r[7],ddd:r[14],diocese:r[15],unidade:r[16],referencia:r[17],validacao:r[18]}))};
 }catch(e){return {ok:false,reason:"Pré-cadastros indisponíveis."};}
}
function salvarPreCadastro(input){
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};const t=consultarTerritorio();if(!t.ok)return t;
 let lock;try{
 const clean=(v,max)=>{if(typeof v!=="string"||v.length>max||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(v.trim()))throw Error("Campo inválido.");return v.trim();};
 if(input?.autorizado!==true)return {ok:false,reason:"Confirme autorização para registrar o contato."};
 const nome=clean(input.nome,180),contato=clean(input.contato,200),ref=clean(input.referencia||"",100);const linked=(t.data.locais||[]).find(x=>x.id===ref);const canon=[...t.data.paroquias,...t.data.comunidades].find(x=>x.id===ref);if(ref&&!linked&&!canon)return {ok:false,reason:"Local não reconhecido."};const local=linked?.nome||canon?.nome||clean(input.local,180),cidade=linked?.cidade||canon?.cidade||clean(input.cidade,120),ddd=linked?.ddd||(t.data.locais||[]).find(x=>x.unidade===canon?.id)?.ddd||"";const phone=padronizarContato_(contato,ddd);
 if(!nome&&!contato)return {ok:false,reason:"Informe pelo menos nome ou contato."};
 lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId),s=ss.getSheetByName("PRE_CADASTROS"),rows=tabela_(ss,"PRE_CADASTROS",PRE_HEADERS_);
 if(rows.some(r=>r[10]==="PENDENTE_IDENTIFICACAO"&&normalizarChave_(r[1])===normalizarChave_(nome)&&normalizarChave_(r[2])===normalizarChave_(contato)&&normalizarChave_(r[4])===normalizarChave_(local)))return {ok:false,reason:"Referência já cadastrada. Revise os registros existentes."};
 const id=Utilities.getUuid(),row=s.getLastRow()+1;if(row>1000)return {ok:false,reason:"Limite da base."};if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),Math.min(100,1000-s.getMaxRows()));
 s.getRange(row,1,1,PRE_HEADERS_.length).setNumberFormat("@").setValues([[id,nome,contato,"",local,cidade,"",new Date().toISOString(),"Registro administrativo autorizado","","PENDENTE_IDENTIFICACAO","",padronizarNome_(nome),phone.contato,ddd,linked?.diocese||canon?.diocese||"",linked?.unidade||canon?.id||"",linked?.id||"",phone.status]]);SpreadsheetApp.flush();return {ok:true,id};
 }catch(e){return {ok:false,reason:"Não foi possível confirmar o pré-cadastro. Confira a planilha."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}

const REF_COORD_HEADERS_=["ID_REFERENCIA","ID_COORDENACAO","NOME_NA_FONTE","SETOR_NA_FONTE","FUNCAO_NA_FONTE","ORIGEM_TERRITORIAL","VIGENCIA_INFORMADA","FONTE","VALIDACAO","ID_PESSOA_VINCULADA"];
const GRUPO_HEADERS_=["ID_GRUPO","ID_UNIDADE","NOME","OBJETIVO","STATUS","CRIADO_EM"];
const MEMBRO_HEADERS_=["ID_MEMBRO","ID_GRUPO","ID_PESSOA","FUNCAO","STATUS","CRIADO_EM"];
const ALERTA_HEADERS_=["ID_ALERTA","TIPO_DESTINO","ID_DESTINO","TITULO","MENSAGEM","PRIORIDADE","STATUS","CRIADO_EM","CANAL"];
function textoOperacional_(v,max,obrigatorio=true){
 if(typeof v!=="string")throw Error("Campo inválido.");
 const s=v.trim();if((obrigatorio&&!s)||s.length>max||/^[=+\-@]|[\u0000-\u0008\u000b-\u001f\u007f]/.test(s))throw Error("Confira os campos.");
 return s;
}
function unidadeRedeValida_(id,data){
 return id==="CO"||[...data.dioceses,...data.paroquias,...data.comunidades].some(x=>x.id===id&&x.id!=="D13");
}
function appendOperacional_(ss,nome,headers,values){
 tabela_(ss,nome,headers);
 const s=ss.getSheetByName(nome),row=s.getLastRow()+1;
 if(row>1000)throw Error("Limite da base inicial.");
 if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),Math.min(100,1000-s.getMaxRows()));
 s.getRange(row,1,1,headers.length).setNumberFormat("@").setValues([values]);SpreadsheetApp.flush();
}
function operacaoRede_(handler){
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 let lock;try{
 lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId);return handler(ss);
 }catch(e){return {ok:false,reason:"Não foi possível confirmar a operação. Confira os campos e a base antes de repetir."};}
 finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
function consultarTrabalho(){
 try{
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId);
 const grupos=tabela_(ss,"GRUPOS_TRABALHO",GRUPO_HEADERS_).map(r=>({id:r[0],unidade:r[1],nome:r[2],objetivo:r[3],status:r[4]}));
 const pessoas=tabela_(ss,"PESSOAS",PESSOAS_HEADERS_).filter(r=>r[9]==="CADASTRADO").map(r=>({id:r[0],nome:r[1],unidade:r[18]}));
 const membros=tabela_(ss,"MEMBROS_GRUPOS",MEMBRO_HEADERS_).map(r=>({id:r[0],grupo:r[1],pessoa:r[2],funcao:r[3],status:r[4]}));
 const alertas=tabela_(ss,"ALERTAS",ALERTA_HEADERS_).map(r=>({id:r[0],tipo:r[1],destino:r[2],titulo:r[3],mensagem:r[4],prioridade:r[5],status:r[6],data:r[7],canal:r[8]}));
 return {ok:true,grupos,pessoas,membros,alertas};
 }catch(e){return {ok:false,reason:"Estrutura de grupos e alertas indisponível."};}
}
function criarGrupoTrabalho(input){
 if(!gestaoConfigAutorizada_())return {ok:false,reason:"Acesso negado."};
 const t=consultarTerritorio();if(!t.ok)return t;
 return operacaoRede_(ss=>{
 const unidade=textoOperacional_(input?.unidade,100),nome=textoOperacional_(input?.nome,120),objetivo=textoOperacional_(input?.objetivo,500);
 if(!unidadeRedeValida_(unidade,t.data))return {ok:false,reason:"Unidade não reconhecida."};
 const rows=tabela_(ss,"GRUPOS_TRABALHO",GRUPO_HEADERS_);
 if(rows.some(r=>r[1]===unidade&&normalizarChave_(r[2])===normalizarChave_(nome)&&r[4]!=="ARQUIVADO"))return {ok:false,reason:"Grupo já existente nesse território."};
 const id=Utilities.getUuid();
 appendOperacional_(ss,"GRUPOS_TRABALHO",GRUPO_HEADERS_,[id,unidade,nome,objetivo,"ATIVO",new Date().toISOString()]);
 return {ok:true,id};
 });
}
function adicionarMembroGrupo(input){
 return operacaoRede_(ss=>{
 const grupo=textoOperacional_(input?.grupo,100),pessoa=textoOperacional_(input?.pessoa,100),funcao=textoOperacional_(input?.funcao,80);
 if(input?.autorizado!==true)return {ok:false,reason:"Confirme a participação autorizada da pessoa."};
 const grupos=tabela_(ss,"GRUPOS_TRABALHO",GRUPO_HEADERS_),pessoas=tabela_(ss,"PESSOAS",PESSOAS_HEADERS_),membros=tabela_(ss,"MEMBROS_GRUPOS",MEMBRO_HEADERS_);
 if(!grupos.some(r=>r[0]===grupo&&r[4]==="ATIVO")||!pessoas.some(r=>r[0]===pessoa&&r[9]==="CADASTRADO"))return {ok:false,reason:"Grupo ou pessoa não disponível."};
 if(membros.some(r=>r[1]===grupo&&r[2]===pessoa&&r[4]==="ATIVO"))return {ok:false,reason:"Pessoa já vinculada ao grupo."};
 const id=Utilities.getUuid();appendOperacional_(ss,"MEMBROS_GRUPOS",MEMBRO_HEADERS_,[id,grupo,pessoa,funcao,"ATIVO",new Date().toISOString()]);return {ok:true,id};
 });
}
function criarAlertaRede(input){
 if(!gestaoConfigAutorizada_())return {ok:false,reason:"Acesso negado."};
 const t=consultarTerritorio();if(!t.ok)return t;
 return operacaoRede_(ss=>{
 const tipo=textoOperacional_(input?.tipo,20),destino=textoOperacional_(input?.destino,100),titulo=textoOperacional_(input?.titulo,140),mensagem=textoOperacional_(input?.mensagem,1500),prioridade=textoOperacional_(input?.prioridade,20);
 const grupos=tabela_(ss,"GRUPOS_TRABALHO",GRUPO_HEADERS_);
 if(!["NORMAL","URGENTE"].includes(prioridade)||!(tipo==="TERRITORIO"?unidadeRedeValida_(destino,t.data):tipo==="GRUPO"&&grupos.some(r=>r[0]===destino&&r[4]==="ATIVO")))return {ok:false,reason:"Destino ou prioridade inválidos."};
 const id=Utilities.getUuid();
 appendOperacional_(ss,"ALERTAS",ALERTA_HEADERS_,[id,tipo,destino,titulo,mensagem,prioridade,"RASCUNHO",new Date().toISOString(),"MURAL_INTERNO"]);
 return {ok:true,id};
 });
}
function publicarAlertaRede(input){
 return operacaoRede_(ss=>{
 const id=textoOperacional_(input?.id,100);if(input?.confirmado!==true)return {ok:false,reason:"Confirme a publicação no mural interno."};
 const rows=tabela_(ss,"ALERTAS",ALERTA_HEADERS_),r=rows.find(x=>x[0]===id);
 if(!r||r[6]!=="RASCUNHO")return {ok:false,reason:"Alerta indisponível ou já publicado."};
 const sheet=ss.getSheetByName("ALERTAS"),raw=sheet.getRange(2,1,sheet.getLastRow()-1,1).getDisplayValues(),i=raw.findIndex(x=>x[0]===id);
 if(i<0)throw Error("Alerta não localizado.");
 sheet.getRange(i+2,7).setValue("PUBLICADO_INTERNO");SpreadsheetApp.flush();
 return {ok:true,id,entregas:0};
 });
}

function padronizarNome_(s){
 return String(s||"").replace(/\s*\(?\bLibras\b\)?\s*/gi," ").replace(/\s+/g," ").trim().split(" ").map((w,i)=>!i||!["de","da","do","dos","das","e"].includes(w.toLowerCase())?w.split(/([-’'])/).map(z=>/^[-’']$/.test(z)?z:z.charAt(0).toUpperCase()+z.slice(1).toLowerCase()).join(""):w.toLowerCase()).join(" ");
}
function padronizarContato_(s,ddd){
 const parts=String(s||"").trim().split(/[\/;\n,]+/).map(x=>x.trim()).filter(Boolean),out=[];
 for(const x of parts){
  if(/[a-z]/i.test(x))return {contato:"",status:"REVISAR MULTIPLOS/OBSERVACOES"};
  let d=x.replace(/\D/g,"");if((d.length===8||d.length===9)&&/^\d{2}$/.test(ddd))d=ddd+d;
  if((d.length===12||d.length===13)&&d.startsWith("55"))d=d.slice(2);
  if(d.length!==10&&d.length!==11)return {contato:"",status:"REVISAR NUMERO OU DDD"};
  out.push("("+d.slice(0,2)+") "+d.slice(2,-4)+"-"+d.slice(-4));
 }
 return {contato:out.join(" / "),status:out.length?"FORMATADO — confirmar titularidade":"SEM CONTATO"};
}
function consultarPessoasPorLocal(){
 try{
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId);
 return {ok:true,pessoas:tabela_(ss,"PESSOAS",PESSOAS_HEADERS_).filter(r=>r[9]==="CADASTRADO").map(r=>({id:r[0],nome:r[1],unidade:r[18]}))};
 }catch(e){return {ok:false,reason:"Consulta indisponível."};}
}
