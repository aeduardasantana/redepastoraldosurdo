// Módulo compartilhado de endereço institucional. EnderecoDiocesano.gs e Cep.gs são dependências.
const ER16_LOC_=["ID_LOCAL","ID_DIOCESE","NOME","TIPO","CIDADE_RA","ENDERECO","ID_UNIDADE","DDD_REFERENCIA","FONTE","VALIDACAO","MUNICIPIO","REGIAO_ADMINISTRATIVA","BAIRRO_SETOR","UF"];
const ER16_ORG_=["ID entidade","Tipo","Nome","Município/abrangência","UF","Telefone","WhatsApp","E-mail","Fonte / canal","Situação","Etapa","Observações / próxima ação","Data referência"];
const ER16_PAR_=["ID unidade","ID diocese","Jurisdição","Nome no diretório","Tipo","Município","Bairro/localidade","Logradouro e número","CEP","Telefone","WhatsApp","E-mail","Site próprio","Expediente","Clero / função","Fundação","Missas (matriz)","Fonte diretório","Fonte ficha","Consulta","Situação da coleta","Pastoral do Surdo","Pastoral da Criança","Pastoral da Pessoa Idosa","PASCOM","Próximo passo","CNPJ","Razão social","Nome de divulgação","Número","Complemento","Redes sociais","Status cadastro"];
const ER16_COM_=["ID comunidade","ID unidade matriz","ID diocese","Nome","Município","Bairro/localidade","Logradouro e número","CEP","Telefone","Missas","Fonte oficial","Data consulta","Situação","Observação","Nome de divulgação","Número","Complemento","WhatsApp","E-mail","Expediente","Status cadastro"];
const ER16_FIELDS_={pais:3,cep:4,logradouro:5,numero:6,complemento:7,bairro:8,cidade:9,uf:10,regiaoAdministrativa:23};
const ER16_LOCAL_FIELDS_=[4,5,10,11,12,13];
function er16Tabela_(ss,n,h,max){return endDioV015Tabela_(ss,n,h,max||5000);}
function er16Index_(t){const m=new Map();for(let i=1;i<t.rows.length;i++){const r=t.rows[i];if(!r.some(Boolean))continue;if(!r[0]||m.has(r[0]))throw Error("ID ausente ou duplicado.");m.set(r[0],{r,i});}return m;}
function er16Base_(cfg){
 const ss=SpreadsheetApp.openById(cfg.sourceId);
 const tabs={dioceses:er16Tabela_(ss,"DIOCESES",END_DIO_V015_DIO_HEADERS_,1000),paroquias:er16Tabela_(ss,"PARÓQUIAS",ER16_PAR_),comunidades:er16Tabela_(ss,"COMUNIDADES",ER16_COM_),locais:er16Tabela_(ss,"LOCAIS_REFERENCIA",ER16_LOC_),organizacoes:er16Tabela_(ss,"ORGANIZAÇÕES",ER16_ORG_),enderecos:er16Tabela_(ss,"ENDERECOS_INSTITUCIONAIS",END_DIO_V015_HEADERS_)};
 const idx=Object.fromEntries(Object.entries(tabs).map(([k,t])=>[k,er16Index_(t)]));
 return {ss,tabs,idx};
}
function er16Alvo_(b,tipo,id){
 const i=b.idx;let x,alvo,diocese="",nome="",categoria="",origem="",proj=null,ufFixa="",seed={};
 const get=(k,id)=>i[k].get(id);
 if(tipo==="regional"||(tipo==="organizacoes"&&id==="ORG001")){
  if(!["CO","ORG001"].includes(id))throw Error("Raiz inválida.");
  x=get("organizacoes","ORG001");if(!x||!get("enderecos","CI-CNBBCO-001"))throw Error("Referência regional incompleta.");
  alvo="CI-CNBBCO-001";nome=x.r[2];categoria="Organismo eclesial";origem="CONTATOS_INSTITUCIONAIS";ufFixa="GO";
 }else if(tipo==="locais"){
  x=get("locais",id);if(!x)throw Error("Local inexistente.");
  if(x.r[6]){
   const p=get("paroquias",x.r[6]),c=get("comunidades",x.r[6]);
   if(Number(!!p)+Number(!!c)!==1||((p?p.r[1]:c.r[2])!==x.r[1]))throw Error("Vínculo local divergente.");
   const a=er16Alvo_(b,p?"paroquias":"comunidades",x.r[6]);return {...a,referencia:id,nomeReferencia:x.r[2]};
  }
  alvo=id;diocese=x.r[1];nome=x.r[2];categoria=x.r[3];origem=["CI-INOSEB-DF-001","CI-SMALDONE-GO-001"].includes(id)?"CONTATOS_INSTITUCIONAIS":"LOCAIS_REFERENCIA";ufFixa=get("dioceses",diocese)?.r[3]||"";
  if(!ufFixa||x.r[13]&&x.r[13]!==ufFixa)throw Error("UF territorial divergente.");
  seed={logradouro:x.r[5],cidade:x.r[10],regiaoAdministrativa:x.r[11],bairro:x.r[12],uf:x.r[13]||ufFixa};
 }else if(tipo==="dioceses"){
  if(!/^D(?:0[1-9]|1[0-2])$/.test(id))throw Error("Jurisdição territorial inválida.");
  x=get(tipo,id);if(!x)throw Error("Jurisdição inexistente.");
  alvo=id;diocese=id;nome=x.r[1];categoria="Jurisdição";origem="DIOCESES";ufFixa=x.r[3];proj={tipo,x,cols:{cidade:6,bairro:8,cep:9},logradouro:7};
 }else if(tipo==="paroquias"||tipo==="comunidades"){
  x=get(tipo,id);if(!x||x.r[tipo==="paroquias"?32:20]==="ARQUIVADO")throw Error("Unidade indisponível.");
  const p=tipo==="paroquias";diocese=x.r[p?1:2];const d=get("dioceses",diocese);if(!d||diocese==="D13")throw Error("Jurisdição inválida.");
  if(!p){const matriz=get("paroquias",x.r[1]);if(!matriz||matriz.r[1]!==diocese||matriz.r[32]==="ARQUIVADO")throw Error("Unidade matriz divergente.");}
  alvo=id;nome=x.r[3];categoria=p?x.r[4]||"Paróquia":"Comunidade";origem=p?"PARÓQUIAS":"COMUNIDADES";ufFixa=d.r[3];proj={tipo,x,cols:p?{cidade:5,bairro:6,cep:8,numero:29,complemento:30}:{cidade:4,bairro:5,cep:7,numero:15,complemento:16},logradouro:p?7:6};
  seed=Object.fromEntries(Object.entries(proj.cols).map(([k,j])=>[k,x.r[j]]));seed.logradouro=x.r[proj.logradouro];seed.uf=ufFixa;
 }else if(tipo==="organizacoes"&&["CI-INOSEB-DF-001","CI-SMALDONE-GO-001"].includes(id)){
  return er16Alvo_(b,"locais",id);
 }else if(tipo==="organizacoes"){
  x=get(tipo,id);if(!x)throw Error("Organização inexistente.");
  alvo=id;nome=x.r[2];categoria=x.r[1];origem="ORGANIZAÇÕES";
  // Município/abrangência e UF de atuação NÃO são endereço físico.
 }else throw Error("Tipo inválido.");
 const e=get("enderecos",alvo);
 if(e&&(e.r[13]!==origem||String(e.r[19]||"")!==diocese||ufFixa&&e.r[10]&&e.r[10]!==ufFixa))throw Error("Endereço existente com referência divergente; revisão necessária.");
 const campos=Object.fromEntries(Object.entries(ER16_FIELDS_).map(([k,j])=>[k,String(e?e.r[j]||"":seed[k]||"")]));
 if(!campos.uf&&ufFixa)campos.uf=ufFixa;
 const locais=[...i.locais.values()].filter(l=>(l.r[6]||l.r[0])===alvo);
 if(locais.some(l=>l.r[1]!==diocese))throw Error("Local vinculado a jurisdição divergente.");
 const revisao=JSON.stringify([alvo,e?.r||null,x?.r||null,locais.map(l=>l.r)]);
 return {alvo,referencia:id,nome,categoria,origem,diocese,ufFixa,e,campos,proj,locais,revisao};
}
function listarEnderecosRedeV016(){
 const cfg=autorizacaoTerritorial_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 try{
  const b=er16Base_(cfg),itens=[{tipo:"regional",id:"CO",nome:"CNBB Regional Centro-Oeste",alvo:"CI-CNBBCO-001",temEndereco:b.idx.enderecos.has("CI-CNBBCO-001")}];
  for(const tipo of ["dioceses","paroquias","comunidades","locais","organizacoes"])for(const {r} of b.idx[tipo].values()){
   if(tipo==="dioceses"&&r[0]==="D13"||tipo==="organizacoes"&&r[0]==="ORG001"||tipo==="paroquias"&&r[32]==="ARQUIVADO"||tipo==="comunidades"&&r[20]==="ARQUIVADO")continue;
   const id=r[0],alvo=tipo==="locais"?r[6]||id:id,nome=r[tipo==="dioceses"?1:tipo==="locais"||tipo==="organizacoes"?2:3],e=b.idx.enderecos.get(alvo);
   itens.push({tipo,id,nome,alvo,temEndereco:!!e,localidade:e?[e.r[9],e.r[23],e.r[10]].filter(Boolean).join(" · "):"",abrangencia:tipo==="organizacoes"?r[3]||"":""});
  }
  for(const id of ["CI-INOSEB-DF-001","CI-SMALDONE-GO-001"]){const l=b.idx.locais.get(id),e=b.idx.enderecos.get(id);if(l&&e)itens.push({tipo:"organizacoes",id,nome:l.r[2],alvo:id,temEndereco:true,localidade:[e.r[9],e.r[23],e.r[10]].filter(Boolean).join(" · "),abrangencia:""});}
  return {ok:true,itens};
 }catch(e){return {ok:false,reason:e.message||"Catálogo indisponível."};}
}
function consultarEnderecoRedeV016(tipo,id){
 const cfg=autorizacaoTerritorial_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 try{const a=er16Alvo_(er16Base_(cfg),tipo,id);return {ok:true,alvo:a.alvo,nome:a.nome,nomeReferencia:a.nomeReferencia||"",campos:a.campos,revisao:a.revisao,ufFixa:a.ufFixa,existente:!!a.e,locais:a.locais.length};}
 catch(e){return {ok:false,reason:e.message||"Endereço indisponível."};}
}
function er16Plano_(a,input){
 const raw=input?.campos;if(!raw||typeof raw!=="object"||Array.isArray(raw)||Object.keys(raw).some(k=>!Object.prototype.hasOwnProperty.call(ER16_FIELDS_,k)))throw Error("Campos inválidos.");
 const v={},max={pais:80,cep:9,logradouro:500,numero:50,complemento:300,bairro:120,cidade:120,uf:2,regiaoAdministrativa:120};
 for(const k of Object.keys(ER16_FIELDS_)){const x=Object.prototype.hasOwnProperty.call(raw,k)?raw[k]:a.campos[k];if(typeof x!=="string"||x.length>max[k]||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(x.trim()))throw Error("Confira os campos de endereço.");v[k]=x.trim();}
 if(v.pais&&v.pais!=="Brasil")throw Error("A consulta postal atual é brasileira.");
 if(v.uf&&!/^(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)$/.test(v.uf))throw Error("UF inválida.");
 if(a.ufFixa&&v.uf!==a.ufFixa)throw Error("UF da jurisdição não pode ser alterada nesta ficha.");
 if(v.cep){if(!/^\d{5}-?\d{3}$/.test(v.cep))throw Error("CEP inválido.");v.cep=v.cep.replace(/-/g,"").replace(/^(\d{5})(\d{3})$/,"$1-$2");}
 const changed=Object.keys(v).filter(k=>v[k]!==a.campos[k]);
 if(["numero","complemento"].some(k=>changed.includes(k)&&a.campos[k]&&endDioV015Contem_(v.logradouro,a.campos[k])&&!endDioV015Contem_(v.logradouro,v[k])))throw Error("Separe do logradouro o número ou complemento anterior antes de salvar.");
 const fisico=[v.cep,v.logradouro,v.numero,v.complemento,v.bairro,v.cidade,v.regiaoAdministrativa].some(Boolean);
 if(!a.e&&(!fisico||input.confirmarEndereco!==true))throw Error("Informe um endereço físico e confirme sua inclusão. Abrangência não é sede.");
 return {v,changed,composto:endDioV015Compor_(v)};
}
function er16Batch_(id,req){
 if(!req.length)return;
 const resp=UrlFetchApp.fetch("https://sheets.googleapis.com/v4/spreadsheets/"+encodeURIComponent(id)+":batchUpdate",{method:"post",contentType:"application/json",headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()},payload:JSON.stringify({requests:req}),muteHttpExceptions:true,followRedirects:false});
 if(resp.getResponseCode()!==200)throw Error("Gravação não confirmada; confira API Sheets, permissões e dados antes de repetir.");
}
function er16Append_(tab,row){return {appendCells:{sheetId:tab.sh.getSheetId(),rows:[{values:row.map(v=>({userEnteredValue:{stringValue:String(v||"")}}))}],fields:"userEnteredValue"}};}
function er16Set_(req,tab,x,j,v){if(String(x.r[j]||"")!==String(v||""))req.push(endDioV015Update_(tab,x.i,j,String(v||"")));}
function er16Espelho_(b,a){
 if(!a.locais.length)return null;
 const g=gestaoConfigAutorizada_();if(!g?.operacionalId)throw Error("Espelho operacional indisponível.");
 const ss=SpreadsheetApp.openById(g.operacionalId),tab=er16Tabela_(ss,"LOCAIS_VINCULADOS",ER16_LOC_),idx=er16Index_(tab),req=[];
 for(const l of a.locais){const m=idx.get(l.r[0]);if(!m||l.r.some((v,j)=>!ER16_LOCAL_FIELDS_.includes(j)&&String(v||"")!==String(m.r[j]||"")))throw Error("Identidade do espelho divergente: "+l.r[0]+". Nenhuma sobrescrita automática.");
  for(const j of ER16_LOCAL_FIELDS_)er16Set_(req,tab,m,j,l.r[j]);
 }
 return {id:g.operacionalId,req};
}
function er16AtualizarPostal_(b,a,p,req){
 if(!p.v.cep)return;
 const t=er16Tabela_(b.ss,"CEP_REFERENCIA",END_DIO_V015_CEP_HEADERS_,10000),norm=x=>String(x||"").replace(/\D/g,""),hits=t.rows.map((r,i)=>({r,i})).filter(x=>x.i&&norm(x.r[0])===norm(p.v.cep));
 if(hits.length>1)throw Error("CEP duplicado na base postal.");
 const ref=a.origem+":"+a.alvo;
 if(!hits.length)req.push(er16Append_(t,[p.v.cep,"","","","","","A consultar","Cadastro administrativo",ref,"",""]));
 else {const h=hits[0],refs=String(h.r[8]||"").split(/\s*\|\s*/).filter(Boolean);if(!refs.includes(ref))er16Set_(req,t,h,8,[...refs,ref].join(" | "));}
}
function er16PlanoGravacao_(b,a,p){
 const req=[],v=p.v,changed=p.changed,novo=!a.e,afeta=k=>novo||changed.includes(k),predial=["logradouro","numero","complemento"].some(afeta),postal=["cep","logradouro","bairro","cidade","uf"].some(afeta);
 const today=Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd");
 if(novo){
  const r=Array(24).fill("");Object.assign(r,{0:a.alvo,1:a.categoria,2:a.nome,12:p.composto,13:a.origem,14:"Cadastro administrativo — endereço declarado",15:a.locais.map(l=>l.r[0]).filter(id=>id!==a.alvo).join(" | "),16:"A validar",17:"Endereço informado no CRM em "+today,18:today,19:a.diocese,20:v.cep?"A consultar":"Sem CEP"});
  for(const [k,j] of Object.entries(ER16_FIELDS_))r[j]=v[k];req.push(er16Append_(b.tabs.enderecos,r));
 }else{
  for(const k of changed)er16Set_(req,b.tabs.enderecos,a.e,ER16_FIELDS_[k],v[k]);
  if(postal){er16Set_(req,b.tabs.enderecos,a.e,20,v.cep?"A consultar":"Sem CEP");if(["cep","cidade","uf"].some(afeta))er16Set_(req,b.tabs.enderecos,a.e,11,"");}
  if(changed.length){
   er16Set_(req,b.tabs.enderecos,a.e,16,"A validar");
   const nota=[a.e.r[17],"Endereço editado no CRM em "+today+"; campos: "+changed.join(", ")+". Fonte original preservada."].filter(Boolean).join(" | ");
   if(nota.length>45000)throw Error("Histórico excede o limite.");er16Set_(req,b.tabs.enderecos,a.e,17,nota);
  }
 }
 if(a.proj){const q=a.proj,t=b.tabs[q.tipo];for(const [k,j] of Object.entries(q.cols))if(afeta(k))er16Set_(req,t,q.x,j,v[k]);if(predial)er16Set_(req,t,q.x,q.logradouro,p.composto);}
 for(const l of a.locais){const t=b.tabs.locais;if(afeta("cidade")||afeta("regiaoAdministrativa"))er16Set_(req,t,l,4,v.regiaoAdministrativa||v.cidade);
  if(predial)er16Set_(req,t,l,5,p.composto);
  for(const [k,j] of Object.entries({cidade:10,regiaoAdministrativa:11,bairro:12,uf:13}))if(afeta(k))er16Set_(req,t,l,j,v[k]);}
 er16AtualizarPostal_(b,a,p,req);return req;
}
function salvarEnderecoRedeV016(input){
 const cfg=autorizacaoTerritorial_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 let lock;try{
  lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};
  const b=er16Base_(cfg),a=er16Alvo_(b,input?.tipo,input?.id);
  if(a.revisao!==input.revisao)return {ok:false,reason:"Cadastro alterado. Reabra a ficha."};
  const p=er16Plano_(a,input);
  // Validação de identidade do espelho antes de qualquer gravação.
  er16Espelho_(b,a);
  const req=er16PlanoGravacao_(b,a,p);
  er16Batch_(cfg.sourceId,req);
  // A API é atômica por planilha, não entre planilhas. O mestre é a fonte.
  try{const bb=er16Base_(cfg),aa=er16Alvo_(bb,input.tipo,input.id),esp=er16Espelho_(bb,aa);if(esp)er16Batch_(esp.id,esp.req);}
  catch(e){return {ok:true,id:a.alvo,espelhoPendente:true,warning:"Mestre gravado; espelho operacional pendente. Use Sincronizar espelho e confira a ficha. "+e.message};}
  return {ok:true,id:a.alvo,espelhoPendente:false};
 }catch(e){return {ok:false,reason:e.message||"Gravação não confirmada."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
function sincronizarEspelhoEnderecoV016(tipo,id){
 const cfg=autorizacaoTerritorial_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 let lock;try{lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};const b=er16Base_(cfg),a=er16Alvo_(b,tipo,id),esp=er16Espelho_(b,a);if(esp)er16Batch_(esp.id,esp.req);return {ok:true,atualizados:esp?.req.length||0};}
 catch(e){return {ok:false,reason:e.message||"Sincronização indisponível."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
