// Cadastro institucional, somente leitura. Configuração privada no pacote de implantação.
function mapTerritorio_(d,p,c,m){
 const fields=(row,indices)=>indices.map(i=>String(row[i]??"").trim());
 const distinct=(rows)=>{const ids=new Set();for(const r of rows){if(!r[0]||ids.has(r[0]))throw new Error("Identificador ausente ou duplicado.");ids.add(r[0]);}return ids;};
 const dioceses=d.map(r=>{const [id,nome,provincia,uf,fonte,situacao,cidade,logradouro,bairro,cep,obs,fonteContato,data,validacao]=fields(r,[0,1,2,3,4,5,6,7,8,9,16,17,19,20]);return {id,nome,provincia,uf,tipo:id==="D13"?"Jurisdição pessoal complementar":nome.startsWith("Arquidiocese")?"Arquidiocese":"Diocese",cidade,logradouro,bairro,cep,situacao,fonte,fonteContato,data,validacao,obs,telefone:String(r[10]??""),telefoneAlternativo:String(r[11]??""),whatsapp:String(r[12]??""),email:String(r[13]??""),emailAlternativo:String(r[14]??""),emailComunicacao:String(r[15]??"")};});
 const paroq=p.map(r=>{const [id,diocese,nome,tipo,cidade,bairro,logradouro,cep,fonte,fonteFicha,data,situacao,pastoral]=fields(r,[0,1,3,4,5,6,7,8,17,18,19,20,21]);return {id,diocese,nome,tipo,cidade,bairro,logradouro,cep,fonte,fonteFicha,data,situacao,pastoral,telefone:String(r[9]??""),whatsapp:String(r[10]??""),email:String(r[11]??""),site:String(r[12]??""),expediente:String(r[13]??""),cnpj:String(r[26]??""),razaoSocial:String(r[27]??""),nomeDivulgacao:String(r[28]??""),numero:String(r[29]??""),complemento:String(r[30]??""),redesSociais:String(r[31]??""),vicariato:String(r[33]??""),setorEclesiastico:String(r[34]??""),localidadeDiretorio:String(r[35]??"")};});
 const comunidades=c.map(r=>{const [id,paroquia,diocese,nome,cidade,bairro,logradouro,cep,fonte,data,situacao,obs]=fields(r,[0,1,2,3,4,5,6,7,10,11,12,13]);return {id,paroquia,diocese,nome,cidade,bairro,logradouro,cep,fonte,data,situacao,obs,telefone:String(r[8]??""),nomeDivulgacao:String(r[14]??""),numero:String(r[15]??""),complemento:String(r[16]??""),whatsapp:String(r[17]??""),email:String(r[18]??""),expediente:String(r[19]??"")};});
 const missas=m.map(r=>{const [id,diocese,cidade,nome,logradouro,dia,horario,situacao,data,fonte,validacao,tipo,unidade,statusProgramacao,pastoral,fontePastoral,programacao]=fields(r,[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16]);return {id,diocese,cidade,nome,logradouro,dia,horario,situacao,data,fonte,validacao,tipo,programacao:programacao||tipo,unidade,statusProgramacao,pastoral,fontePastoral,idPastoral:String(r[18]??""),libras:String(r[19]??""),dataEspecifica:String(r[20]??""),horarioFinal:String(r[21]??""),publicarPortal:String(r[22]??"")};});
 distinct(d);distinct(p);distinct(c);distinct(m);
 const did=new Set(dioceses.map(x=>x.id));const pmap=new Map(paroq.map(x=>[x.id,x]));
 if(paroq.some(x=>!did.has(x.diocese))||comunidades.some(x=>!pmap.has(x.paroquia)||pmap.get(x.paroquia).diocese!==x.diocese))throw new Error("Vínculo territorial inconsistente.");
 return {dioceses,paroquias:paroq,comunidades,missas:missas.map(x=>({...x,unidadeInformada:x.unidade,diocese:did.has(x.diocese)?x.diocese:"",vinculoPendente:!did.has(x.diocese),unidade:(pmap.get(x.unidade)?.diocese===x.diocese||comunidades.some(c=>c.id===x.unidade&&c.diocese===x.diocese))?x.unidade:""}))};
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
     lerTerritorio_(ss,"PARÓQUIAS",["ID unidade","ID diocese","Jurisdição","Nome no diretório","Tipo","Município","Bairro/localidade","Logradouro e número","CEP","Telefone","WhatsApp","E-mail","Site próprio","Expediente","Clero / função","Fundação","Missas (matriz)","Fonte diretório","Fonte ficha","Consulta","Situação da coleta","Pastoral do Surdo","Pastoral da Criança","Pastoral da Pessoa Idosa","PASCOM","Próximo passo","CNPJ","Razão social","Nome de divulgação","Número","Complemento","Redes sociais","Status cadastro","VICARIATO_OFICIAL","SETOR_ECLESIASTICO","LOCALIDADE_DIRETORIO_OFICIAL","DATA_CONFERENCIA_DIRETORIO"],37,5000).filter(r=>r[32]!=="ARQUIVADO"),
     lerTerritorio_(ss,"COMUNIDADES",["ID comunidade","ID unidade matriz","ID diocese","Nome","Município","Bairro/localidade","Logradouro e número","CEP","Telefone","Missas","Fonte oficial","Data consulta","Situação","Observação","Nome de divulgação","Número","Complemento","WhatsApp","E-mail","Expediente","Status cadastro"],21,5000).filter(r=>r[20]!=="ARQUIVADO"),
     lerTerritorio_(ss,"MISSAS LIBRAS",["ID referência","ID diocese (validar)","Cidade / região","Igreja / comunidade","Endereço público","Dia","Horário informado","Situação na fonte (não garantia atual)","Última referência","Fonte","Validação para visita","Tipo de celebração","ID unidade vinculada","Situação da programação","Pastoral do Surdo no local","Fonte da situação pastoral","Programação","Status cadastro","ID pastoral","Acessibilidade Libras","Data específica","Horário final","Publicar no portal"],23,5000).filter(r=>r[17]!=="ARQUIVADO"));
   data.locais=lerTerritorio_(ss,"LOCAIS_REFERENCIA",["ID_LOCAL","ID_DIOCESE","NOME","TIPO","CIDADE_RA","ENDERECO","ID_UNIDADE","DDD_REFERENCIA","FONTE","VALIDACAO","MUNICIPIO","REGIAO_ADMINISTRATIVA","BAIRRO_SETOR","UF"],14,5000).map(r=>({id:r[0],diocese:r[1],nome:r[2],tipo:r[3],cidade:r[10]||r[11]||"",municipio:r[10],regiaoAdministrativa:r[11],bairro:r[12],uf:r[13],logradouro:r[5],unidade:r[6],ddd:r[7],fonte:r[8],validacao:r[9]}));
   data.pastorais=lerTerritorio_(ss,"PASTORAIS",["ID pastoral","ID território","ID diocese","Nome da pastoral","Situação da pastoral","Início de atuação","Contato institucional","Histórico e validação","Fonte","Status cadastro"],10,5000).filter(r=>r[9]!=="ARQUIVADO").map(r=>({id:r[0],unidade:r[1],diocese:r[2],nome:r[3],situacaoPastoral:r[4],inicio:r[5],contato:r[6],historico:r[7],fonte:r[8]}));
   return {ok:true,data:typeof prog16Normalizar_==="function"?prog16Normalizar_(data):data};
 }catch(e){return {ok:false,reason:"Não foi possível consultar o cadastro institucional. Verifique a fonte e a autorização."};}
}

const SOLICITACAO_HEADERS_=["ID_SOLICITACAO","DATA_UTC","TIPO","ID_SUPERIOR","NOME","MUNICIPIO","UF","BAIRRO","ENDERECO","CEP","FONTE","SITUACAO"];
function normalizarChave_(s){return String(s??"").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]+/g," ").trim();}
function validarSolicitacao_(input,data){
 const clean=(s,max,required=false)=>{if(typeof s!=="string")throw Error("Campo inválido.");const v=s.trim();if((required&&!v)||v.length>max||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(v))throw Error("Preencha os campos institucionais sem fórmulas.");return v;};
 if(!input||typeof input!=="object")throw Error("Solicitação inválida.");
 const tipo=clean(input.tipo,30,true),superior=clean(input.superior,80,true),nome=clean(input.nome,180,true),cidade=clean(input.cidade,100,true),uf=clean(input.uf,2,true),bairro=clean(input.bairro,120),endereco=clean(input.endereco,300),cep=clean(input.cep,9),fonte=clean(input.fonte||"Cadastro administrativo",500);
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

// Administração cadastral: campos de negócio editáveis, metadados preservados.
const FICHAS_TERRITORIAIS_={
 pastorais:{aba:"PASTORAIS",status:9,width:10,fields:{unidade:1,nome:3,situacaoPastoral:4,inicio:5,contato:6}},
 paroquias:{aba:"PARÓQUIAS",status:32,width:33,fields:{diocese:1,nome:3,tipo:4,cidade:5,bairro:6,logradouro:7,cep:8,telefone:9,whatsapp:10,email:11,site:12,expediente:13,responsavel:14,fundacao:15,cnpj:26,razaoSocial:27,nomeDivulgacao:28,numero:29,complemento:30,redesSociais:31}},
 comunidades:{aba:"COMUNIDADES",status:20,width:21,fields:{paroquia:1,nome:3,cidade:4,bairro:5,logradouro:6,cep:7,telefone:8,nomeDivulgacao:14,numero:15,complemento:16,whatsapp:17,email:18,expediente:19}},
 missas:{aba:"MISSAS LIBRAS",status:17,width:23,fields:{diocese:1,cidade:2,nome:3,logradouro:4,dia:5,horario:6,tipo:11,unidade:12,statusProgramacao:13,programacao:16}}
};
function fichaTerritorialDados_(row,def){return Object.fromEntries(Object.entries(def.fields).map(([key,i])=>[key,String(row[i]??"")]));}
function fichaTerritorialRows_(ss,def){const sh=ss.getSheetByName(def.aba);if(!sh||sh.getLastRow()<1||sh.getLastRow()>5000||sh.getRange(1,def.status+1).getDisplayValue()!=="Status cadastro")throw Error("Estrutura cadastral incompatível.");return {sh,rows:sh.getRange(1,1,sh.getLastRow(),def.width).getDisplayValues()};}
function consultarFichaTerritorial(tipo,id){
 try{const cfg=autorizacaoTerritorial_();const def=Object.prototype.hasOwnProperty.call(FICHAS_TERRITORIAIS_,tipo)?FICHAS_TERRITORIAIS_[tipo]:null;if(!cfg||!def)return {ok:false,reason:"Acesso negado ou tipo inválido."};const {rows}=fichaTerritorialRows_(SpreadsheetApp.openById(cfg.sourceId),def);const r=rows.slice(1).find(x=>x[0]===id&&x[def.status]!=="ARQUIVADO");if(!r)return {ok:false,reason:"Cadastro indisponível."};const campos=fichaTerritorialDados_(r,def);return {ok:true,campos,revisao:JSON.stringify(campos)};}catch(e){return {ok:false,reason:"Ficha indisponível."};}
}
function salvarFichaTerritorial(input){
 const cfg=autorizacaoTerritorial_(),def=Object.prototype.hasOwnProperty.call(FICHAS_TERRITORIAIS_,input?.tipo)?FICHAS_TERRITORIAIS_[input.tipo]:null;if(!cfg||!def)return {ok:false,reason:"Acesso negado ou tipo inválido."};let lock;
 try{
  const clean=(v)=>{if(typeof v!=="string"||v.length>1000||/^[=]|[\u0000-\u001f\u007f]/.test(v.trim()))throw Error("Campo inválido.");return v.trim();};
  const values={};for(const key of Object.keys(def.fields))values[key]=clean(input.campos?.[key]??"");
  if(values.email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(values.email))return {ok:false,reason:"E-mail inválido."};
  if(values.cnpj&&!/^[A-Z0-9]{12}\d{2}$/i.test(values.cnpj.replace(/[.\-\/]/g,"")))return {ok:false,reason:"Confira o formato do CNPJ (14 posições)."};
  if(values.cep&&!/^\d{5}-?\d{3}$/.test(values.cep))return {ok:false,reason:"CEP inválido."};
  lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};
  const t=consultarTerritorio();if(!t.ok)return t;const data=t.data,ss=SpreadsheetApp.openById(cfg.sourceId);const {sh,rows}=fichaTerritorialRows_(ss,def);
  const idx=input.id?rows.findIndex((r,i)=>i>0&&r[0]===input.id&&r[def.status]!=="ARQUIVADO"):-1;
  if(input.id&&idx<1)return {ok:false,reason:"Cadastro indisponível."};
  const old=idx>0?rows[idx]:null;if(old&&JSON.stringify(fichaTerritorialDados_(old,def))!==input.revisao)return {ok:false,reason:"A ficha mudou. Reabra antes de salvar."};
  let extra={};
  if(input.tipo==="paroquias"){
   const d=data.dioceses.find(x=>x.id===values.diocese);if(!d||!values.nome)return {ok:false,reason:"Informe nome e jurisdição."};if(old&&old[1]!==d.id)return {ok:false,reason:"Transferência de jurisdição exige revisão dos vínculos."};extra[2]=d.nome;
  }else if(input.tipo==="comunidades"){
   const p=data.paroquias.find(x=>x.id===values.paroquia);if(!p||!values.nome)return {ok:false,reason:"Informe nome e unidade matriz."};if(old&&old[1]!==p.id)return {ok:false,reason:"Transferência de unidade matriz exige revisão dos vínculos."};extra[2]=p.diocese;
  }else if(input.tipo==="pastorais"){
   const u=[{id:"CO",diocese:""},...data.dioceses.map(d=>({...d,diocese:d.id})),...data.paroquias,...data.comunidades,...data.locais].find(x=>x.id===values.unidade);if(!u||!values.nome)return {ok:false,reason:"Informe nome da pastoral e território."};if(old&&old[1]!==u.id)return {ok:false,reason:"Mudança de território exige revisão dos vínculos."};if(values.situacaoPastoral&&!["Ativa","Inativa","Constituída","Em implantação","Não informada"].includes(values.situacaoPastoral))return {ok:false,reason:"Situação pastoral inválida."};extra[2]=u.diocese||"";
  }else{
   const u=[...data.paroquias,...data.comunidades].find(x=>x.id===values.unidade);
   if(values.unidade&&!u)return {ok:false,reason:"Unidade inválida."};if(u){values.diocese=u.diocese;values.nome=u.nomeDivulgacao||u.nome;values.cidade=u.cidade||values.cidade;if(!values.logradouro)values.logradouro=u.logradouro||"";}
   if(!data.dioceses.some(x=>x.id===values.diocese)||!values.programacao||(!values.unidade&&!values.nome))return {ok:false,reason:"Informe jurisdição, local e programação."};
   if(!values.statusProgramacao)values.statusProgramacao="A CONFIRMAR";
   if(!["CONFIRMADA","RECORRENTE INFORMADA","A CONFIRMAR","DIVERGENTE"].includes(values.statusProgramacao))return {ok:false,reason:"Situação inválida."};
   if(!values.tipo)values.tipo=values.programacao;
   if(!old){extra[8]=new Date().toISOString().slice(0,10);extra[9]="Registro administrativo";}
  }
  const matchKeys=input.tipo==="paroquias"?["diocese","nome","cidade","bairro"]:input.tipo==="comunidades"?["paroquia","nome","cidade","bairro"]:input.tipo==="pastorais"?["unidade","nome"]:["diocese","unidade","nome","programacao","dia","horario"];if(rows.slice(1).some(r=>r[0]!==input.id&&r[def.status]!=="ARQUIVADO"&&matchKeys.every(k=>normalizarChave_(r[def.fields[k]])===normalizarChave_(values[k]))))return {ok:false,reason:"Já existe registro equivalente. Abra a ficha existente."};
  const id=old?old[0]:Utilities.getUuid(), target=old?idx+1:sh.getLastRow()+1;
  if(!old){if(target>5000)return {ok:false,reason:"Limite atual da base."};if(target>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),100);const r=Array(def.width).fill("");r[0]=id;r[def.status]="ATIVO";for(const [k,i] of Object.entries(def.fields))r[i]=values[k];for(const [i,v] of Object.entries(extra))r[Number(i)]=v;sh.getRange(target,1,1,def.width).setNumberFormat("@").setValues([r]);}
  else {for(const [k,i] of Object.entries(def.fields))if(values[k]!==old[i])sh.getRange(target,i+1).setNumberFormat("@").setValue(values[k]);for(const [i,v] of Object.entries(extra))if(v!==old[Number(i)])sh.getRange(target,Number(i)+1).setNumberFormat("@").setValue(v);}
  SpreadsheetApp.flush();return {ok:true,id};
 }catch(e){return {ok:false,reason:"Gravação não confirmada. Confira a ficha e tente novamente."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
function arquivarFichaTerritorial(tipo,id,revisao){
 const cfg=autorizacaoTerritorial_(),def=Object.prototype.hasOwnProperty.call(FICHAS_TERRITORIAIS_,tipo)?FICHAS_TERRITORIAIS_[tipo]:null;if(!cfg||!def)return {ok:false,reason:"Acesso negado."};let lock;
 try{lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};const ss=SpreadsheetApp.openById(cfg.sourceId),{sh,rows}=fichaTerritorialRows_(ss,def);const idx=rows.findIndex((r,i)=>i>0&&r[0]===id&&r[def.status]!=="ARQUIVADO");if(idx<1||JSON.stringify(fichaTerritorialDados_(rows[idx],def))!==revisao)return {ok:false,reason:"Reabra a ficha antes de excluir."};
 const op=typeof gestaoConfigAutorizada_==="function"?gestaoConfigAutorizada_():null;if(!op?.operacionalId)return {ok:false,reason:"Não foi possível verificar os vínculos da base operacional."};
 const books=[ss,SpreadsheetApp.openById(op.operacionalId)];for(const book of books)for(const tab of book.getSheets()){const refs=tab.createTextFinder(id).matchEntireCell(true).matchCase(true).findAll();if(refs.some(r=>!(book.getId()===ss.getId()&&tab.getSheetId()===sh.getSheetId()&&r.getRow()===idx+1)))return {ok:false,reason:"Registro possui vínculos ou referências. Revise-os antes de arquivar."};}
 sh.getRange(idx+1,def.status+1).setValue("ARQUIVADO");SpreadsheetApp.flush();return {ok:true,id};
 }catch(e){return {ok:false,reason:"Exclusão não confirmada."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
