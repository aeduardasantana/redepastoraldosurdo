// Cadastro institucional, somente leitura. Configuração privada no pacote de implantação.
function mapTerritorio_(d,p,c,m){
 const fields=(row,indices)=>indices.map(i=>String(row[i]??"").trim());
 const distinct=(rows)=>{const ids=new Set();for(const r of rows){if(!r[0]||ids.has(r[0]))throw new Error("Identificador ausente ou duplicado.");ids.add(r[0]);}return ids;};
 const dioceses=d.map(r=>{const [id,nome,provincia,uf,fonte,situacao,cidade,logradouro,bairro,cep,obs,fonteContato,data,validacao]=fields(r,[0,1,2,3,4,5,6,7,8,9,16,17,19,20]);return {id,nome,provincia,uf,tipo:id==="D13"?"Jurisdição pessoal complementar":nome.startsWith("Arquidiocese")?"Arquidiocese":"Diocese",cidade,logradouro,bairro,cep,situacao,fonte,fonteContato,data,validacao,obs};});
 const paroq=p.map(r=>{const [id,diocese,nome,tipo,cidade,bairro,logradouro,cep,fonte,fonteFicha,data,situacao,pastoral]=fields(r,[0,1,3,4,5,6,7,8,17,18,19,20,21]);return {id,diocese,nome,tipo,cidade,bairro,logradouro,cep,fonte,fonteFicha,data,situacao,pastoral};});
 const comunidades=c.map(r=>{const [id,paroquia,diocese,nome,cidade,bairro,logradouro,cep,fonte,data,situacao,obs]=fields(r,[0,1,2,3,4,5,6,7,10,11,12,13]);return {id,paroquia,diocese,nome,cidade,bairro,logradouro,cep,fonte,data,situacao,obs};});
 const missas=m.map(r=>{const [id,diocese,cidade,nome,logradouro,dia,horario,situacao,data,fonte,validacao,tipo,unidade,statusProgramacao]=fields(r,[0,1,2,3,4,5,6,7,8,9,10,11,12,13]);return {id,diocese,cidade,nome,logradouro,dia,horario,situacao,data,fonte,validacao,tipo,unidade,statusProgramacao};});
 distinct(d);distinct(p);distinct(c);distinct(m);
 const did=new Set(dioceses.map(x=>x.id));const pmap=new Map(paroq.map(x=>[x.id,x]));
 if(paroq.some(x=>!did.has(x.diocese))||comunidades.some(x=>!pmap.has(x.paroquia)||pmap.get(x.paroquia).diocese!==x.diocese))throw new Error("Vínculo territorial inconsistente.");
 return {dioceses,paroquias:paroq,comunidades,missas:missas.map(x=>({...x,diocese:did.has(x.diocese)?x.diocese:"",vinculoPendente:!did.has(x.diocese),unidade:(pmap.get(x.unidade)?.diocese===x.diocese||comunidades.some(c=>c.id===x.unidade&&c.diocese===x.diocese))?x.unidade:""}))};
}

function lerTerritorio_(ss,nome,cabecalho,largura,max){
 const s=ss.getSheetByName(nome);
 if(!s||s.getLastRow()<1||s.getLastRow()>max)throw new Error("Limite ou estrutura da fonte.");
 const v=s.getRange(1,1,s.getLastRow(),largura).getDisplayValues();
 if(!cabecalho.every((x,i)=>v[0][i]===x))throw new Error("Cabeçalho incompatível.");
 return v.slice(1).filter(r=>r.some(x=>String(x??"").trim()));
}
function autorizacaoTerritorial_(){
 try{const cfg=typeof territorialConfig_==="function"?territorialConfig_():null;const email=normalizarEmail_(Session.getActiveUser().getEmail());return cfg&&email&&email===normalizarEmail_(cfg.owner)?cfg:null;}catch(e){return null;}
}
function verificarSessao(){
 const cfg=autorizacaoTerritorial_();
 return cfg?{ok:true,perfil:"Administração institucional",escopo:"Regional Centro-Oeste",modo:"Homologação — acesso restrito ao titular"}:{ok:false,reason:"Conta não autorizada ou identidade Google indisponível."};
}
function consultarTerritorio(){
 try{
   const cfg=autorizacaoTerritorial_();
   if(!cfg)return {ok:false,reason:"Consulta restrita à conta institucional autorizada."};
   const ss=SpreadsheetApp.openById(cfg.sourceId);
   const data=mapTerritorio_(
     lerTerritorio_(ss,"DIOCESES",["ID","Jurisdição","Província","UF","Fonte","Status","Município da cúria","Logradouro e número","Bairro","CEP","Telefone principal","Telefone alternativo","WhatsApp publicado","E-mail principal","E-mail alternativo/pastoral","E-mail comunicação","Observações de validação","Fonte dos contatos","Diretório de paróquias / CDIC","Data de referência","Validação do contato"],21,1000),
     lerTerritorio_(ss,"PARÓQUIAS",["ID unidade","ID diocese","Jurisdição","Nome no diretório","Tipo","Município","Bairro/localidade","Logradouro e número","CEP","Telefone","WhatsApp","E-mail","Site próprio","Expediente","Clero / função","Fundação","Missas (matriz)","Fonte diretório","Fonte ficha","Consulta","Situação da coleta","Pastoral do Surdo","Pastoral da Criança","Pastoral da Pessoa Idosa","PASCOM","Próximo passo"],26,1000),
     lerTerritorio_(ss,"COMUNIDADES",["ID comunidade","ID unidade matriz","ID diocese","Nome","Município","Bairro/localidade","Logradouro e número","CEP","Telefone","Missas","Fonte oficial","Data consulta","Situação","Observação"],14,1000),
     lerTerritorio_(ss,"MISSAS LIBRAS",["ID referência","ID diocese (validar)","Cidade / região","Igreja / comunidade","Endereço público","Dia","Horário informado","Situação na fonte (não garantia atual)","Última referência","Fonte","Validação para visita","Tipo de celebração","ID unidade vinculada","Situação da programação"],14,1000));
   return {ok:true,data};
 }catch(e){return {ok:false,reason:"Não foi possível consultar o cadastro institucional. Verifique a fonte e a autorização."};}
}

const SOLICITACAO_HEADERS_=["ID_SOLICITACAO","DATA_UTC","TIPO","ID_SUPERIOR","NOME","MUNICIPIO","UF","BAIRRO","ENDERECO","CEP","FONTE","SITUACAO"];
function normalizarChave_(s){return String(s??"").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]+/g," ").trim();}
function validarSolicitacao_(input,data){
 const clean=(s,max,required=false)=>{if(typeof s!=="string")throw Error("Campo inválido.");const v=s.trim();if((required&&!v)||v.length>max||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(v))throw Error("Preencha os campos institucionais sem fórmulas.");return v;};
 if(!input||typeof input!=="object")throw Error("Solicitação inválida.");
 const tipo=clean(input.tipo,30,true),superior=clean(input.superior,80,true),nome=clean(input.nome,180,true),cidade=clean(input.cidade,100,true),uf=clean(input.uf,2,true),bairro=clean(input.bairro,120),endereco=clean(input.endereco,300),cep=clean(input.cep,9),fonte=clean(input.fonte,500,true);
 if(!["PAROQUIA","QUASE_PAROQUIA","SANTUARIO","COMUNIDADE","CAPELA","NUCLEO"].includes(tipo)||!["GO","DF"].includes(uf)||cep&&!/^\d{5}-?\d{3}$/.test(cep))throw Error("Tipo, UF ou CEP inválido.");
 const local=["COMUNIDADE","CAPELA","NUCLEO"].includes(tipo),parents=local?data.paroquias:data.dioceses;
 if(!parents.some(x=>x.id===superior))throw Error("Selecione a unidade superior válida.");
 const candidates=local?data.comunidades.filter(x=>x.paroquia===superior):data.paroquias.filter(x=>x.diocese===superior);
 const same=(x)=>normalizarChave_(x.nome)===normalizarChave_(nome)&&(!x.cidade||normalizarChave_(x.cidade)===normalizarChave_(cidade));
 if(candidates.some(same))throw Error("Já existe unidade com este nome no território. Consulte o diretório.");
 return {tipo,superior,nome,cidade,uf,bairro,endereco,cep,fonte};
}
function solicitarUnidade(input){
 try{
  // O mesmo controle de identidade precede qualquer leitura ou escrita.
  const data=consultarTerritorio();if(!data.ok)return data;
  let v;try{v=validarSolicitacao_(input,data.data);}catch(e){return {ok:false,reason:e.message};}
  const lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Outra solicitação está sendo processada. Tente novamente."};
  try{
   const s=planilha_().getSheetByName("SOLICITACOES_UNIDADES");
   if(!s||s.getLastRow()>1000||s.getLastRow()<1)throw Error("Fila indisponível.");
   const rows=s.getRange(1,1,s.getLastRow(),SOLICITACAO_HEADERS_.length).getDisplayValues();
   if(!SOLICITACAO_HEADERS_.every((h,i)=>rows[0][i]===h))throw Error("Estrutura inválida.");
   if(rows.slice(1).some(r=>r[11]==="PENDENTE"&&r[3]===v.superior&&normalizarChave_(r[4])===normalizarChave_(v.nome)&&normalizarChave_(r[5])===normalizarChave_(v.cidade)))return {ok:false,reason:"Já existe solicitação pendente para essa unidade."};
   const id=Utilities.getUuid(),values=[id,new Date().toISOString(),v.tipo,v.superior,v.nome,v.cidade,v.uf,v.bairro,v.endereco,v.cep,v.fonte,"PENDENTE"];
   const next=s.getLastRow()+1;if(next>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),Math.min(100,1001-s.getMaxRows()));
   s.getRange(next,1,1,values.length).setNumberFormat("@").setValues([values]);SpreadsheetApp.flush();
   return {ok:true,id};
  }finally{lock.releaseLock();}
 }catch(e){return {ok:false,reason:"Não foi possível registrar a solicitação. Nenhuma inclusão direta na fonte mestre."};}
}
