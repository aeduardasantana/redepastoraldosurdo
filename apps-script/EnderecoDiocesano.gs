// Endereço diocesano — módulo complementar independente; não redeclarar os endpoints existentes.
const END_DIO_V015_CAMPOS_={cep:4,logradouro:5,numero:6,complemento:7,bairro:8,cidade:9,regiaoAdministrativa:23};
const END_DIO_V015_HEADERS_=["ID_REFERENCIA","TIPO","NOME","PAIS","CEP","LOGRADOURO","NUMERO","COMPLEMENTO","LOCALIDADE","MUNICIPIO","UF","CODIGO_IBGE","ENDERECO_ORIGINAL","ABA_ORIGEM","FONTE","IDS_LOCAIS_VINCULADOS","STATUS_ENDERECO","OBSERVACAO_VALIDACAO","DATA_ESTRUTURACAO","ID_DIOCESE","SITUACAO_POSTAL","COBERTURA_CADASTRAL","REVISAO_INSTITUCIONAL","REGIAO_ADMINISTRATIVA"];
const END_DIO_V015_DIO_HEADERS_=["ID","Jurisdição","Província","UF","Fonte","Status","Município da cúria","Logradouro e número","Bairro","CEP"];
const END_DIO_V015_CEP_HEADERS_=["CEP","LOGRADOURO_POSTAL","BAIRRO_POSTAL","MUNICIPIO_POSTAL","UF","CODIGO_IBGE","STATUS_VALIDACAO","FONTE_ORIGEM","REFERENCIAS_TERRITORIAIS","FONTE_CONSULTA","DATA_CONSULTA"];
function endDioV015Tabela_(ss,nome,headers,max){
 const sh=ss.getSheetByName(nome);if(!sh||sh.getLastRow()<1||sh.getLastRow()>max)throw Error("Estrutura indisponível.");
 const rows=sh.getRange(1,1,sh.getLastRow(),headers.length).getDisplayValues();
 if(!headers.every((h,i)=>rows[0][i]===h))throw Error("Cabeçalho incompatível.");
 return {sh,rows};
}
function endDioV015Unico_(tab,id){
 const hits=tab.rows.map((r,i)=>({r,i})).filter(x=>x.i>0&&x.r[0]===id);
 if(hits.length!==1)throw Error("ID ausente ou duplicado.");
 return hits[0];
}
function endDioV015Contexto_(cfg,id){
 if(!/^D(?:0[1-9]|1[0-2])$/.test(id))throw Error("Jurisdição territorial inválida.");
 const ss=SpreadsheetApp.openById(cfg.sourceId),dio=endDioV015Tabela_(ss,"DIOCESES",END_DIO_V015_DIO_HEADERS_,1000),end=endDioV015Tabela_(ss,"ENDERECOS_INSTITUCIONAIS",END_DIO_V015_HEADERS_,5000);
 const d=endDioV015Unico_(dio,id),e=endDioV015Unico_(end,id);
 if(e.r[13]!=="DIOCESES"||e.r[19]!==id||!["GO","DF"].includes(d.r[3])||e.r[10]!==d.r[3])throw Error("Referência territorial divergente. Solicite revisão.");
 const campos=Object.fromEntries(Object.entries(END_DIO_V015_CAMPOS_).map(([k,i])=>[k,e.r[i]||""]));
 const revisao=JSON.stringify([d.r,e.r]);
 return {ss,dio,end,d,e,campos,revisao,uf:d.r[3],nome:d.r[1]};
}
function consultarEnderecoDiocesanoV015(id){
 const cfg=autorizacaoTerritorial_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 try{const c=endDioV015Contexto_(cfg,String(id||""));return {ok:true,campos:c.campos,uf:c.uf,nome:c.nome,revisao:c.revisao};}
 catch(e){return {ok:false,reason:e.message||"Endereço indisponível."};}
}
function endDioV015Normal_(s){return String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9]+/g," ").trim();}
function endDioV015Contem_(a,b){const n=endDioV015Normal_(b);return !!n&&(" "+endDioV015Normal_(a)+" ").includes(" "+n+" ");}
function endDioV015Compor_(v){
 let out=v.logradouro;
 for(const x of [v.numero,v.complemento])if(x&&!endDioV015Contem_(out,x))out=[out,x].filter(Boolean).join(", ");
 return out;
}
function endDioV015Planejar_(c,input){
 const dados=input?.campos;if(!dados||typeof dados!=="object"||Array.isArray(dados)||Object.keys(dados).some(k=>!Object.prototype.hasOwnProperty.call(END_DIO_V015_CAMPOS_,k)))throw Error("Campos inválidos.");
 const v={};const limites={cep:9,logradouro:500,numero:50,complemento:300,bairro:120,cidade:120,regiaoAdministrativa:120};
 for(const k of Object.keys(END_DIO_V015_CAMPOS_)){
  const raw=Object.prototype.hasOwnProperty.call(dados,k)?dados[k]:c.campos[k];
  if(typeof raw!=="string"||raw.length>limites[k]||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(raw.trim()))throw Error("Confira os campos de endereço.");
  v[k]=raw.trim();
 }
 if(v.cep){if(!/^\d{5}-?\d{3}$/.test(v.cep))throw Error("CEP inválido.");v.cep=v.cep.replace(/-/g,"").replace(/^(\d{5})(\d{3})$/,"$1-$2");}
 const changed=Object.keys(v).filter(k=>v[k]!==c.campos[k]);
 if(["numero","complemento"].some(k=>changed.includes(k)&&c.campos[k]&&endDioV015Contem_(v.logradouro,c.campos[k])&&!endDioV015Contem_(v.logradouro,v[k])))throw Error("O logradouro ainda contém o número ou complemento anterior. Separe esses campos antes de salvar.");
 const endereco=changed.some(k=>["logradouro","numero","complemento"].includes(k))?endDioV015Compor_(v):c.d.r[7];
 return {v,changed,endereco};
}
function endDioV015Update_(tab,i,j,value){
 if(tab.sh.getRange(i+1,j+1).getFormula())throw Error("Campo calculado protegido.");
 return {updateCells:{range:{sheetId:tab.sh.getSheetId(),startRowIndex:i,endRowIndex:i+1,startColumnIndex:j,endColumnIndex:j+1},rows:[{values:[{userEnteredValue:{stringValue:value}}]}],fields:"userEnteredValue"}};
}
function salvarEnderecoDiocesanoV015(input){
 const cfg=autorizacaoTerritorial_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 let lock;try{
  lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};
  const id=String(input?.id||""),c=endDioV015Contexto_(cfg,id);
  if(c.revisao!==input.revisao)return {ok:false,reason:"O cadastro mudou. Reabra antes de salvar."};
  const p=endDioV015Planejar_(c,input);if(!p.changed.length)return {ok:true,id,semAlteracao:true};
  const req=p.changed.map(k=>endDioV015Update_(c.end,c.e.i,END_DIO_V015_CAMPOS_[k],p.v[k]));
  const espelho={cidade:6,bairro:8,cep:9};
  for(const [k,j] of Object.entries(espelho))if(p.changed.includes(k)&&p.v[k]!==c.d.r[j])req.push(endDioV015Update_(c.dio,c.d.i,j,p.v[k]));
  if(p.endereco!==c.d.r[7])req.push(endDioV015Update_(c.dio,c.d.i,7,p.endereco));
  // Dados editados não equivalem a validação postal. Originais e fontes anteriores permanecem intactos.
  if(p.changed.some(k=>["cep","cidade","bairro","logradouro"].includes(k))){
   if(c.e.r[20]!==(p.v.cep?"A consultar":"Sem CEP"))req.push(endDioV015Update_(c.end,c.e.i,20,p.v.cep?"A consultar":"Sem CEP"));
   if(c.e.r[11]&&p.changed.some(k=>["cep","cidade"].includes(k)))req.push(endDioV015Update_(c.end,c.e.i,11,""));
  }
  if(c.e.r[16]!=="A validar")req.push(endDioV015Update_(c.end,c.e.i,16,"A validar"));
  const today=Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd");
  const nota=[c.e.r[17],"Endereço editado no CRM em "+today+"; campos: "+p.changed.join(", ")+". Fonte original preservada."].filter(Boolean).join(" | ");
  if(nota.length>45000)throw Error("Histórico de endereço excedeu o limite; solicite revisão.");
  req.push(endDioV015Update_(c.end,c.e.i,17,nota));
  if(p.v.cep){
   const cep=endDioV015Tabela_(c.ss,"CEP_REFERENCIA",END_DIO_V015_CEP_HEADERS_,10000),norm=x=>String(x||"").replace(/\D/g,"");
   const hits=cep.rows.map((r,i)=>({r,i})).filter(x=>x.i>0&&norm(x.r[0])===norm(p.v.cep));
   if(hits.length>1)throw Error("CEP duplicado na base postal.");
   const ref="DIOCESES:"+id;
   if(!hits.length)req.push({appendCells:{sheetId:cep.sh.getSheetId(),rows:[{values:[p.v.cep,"","","","","","A consultar","Cadastro administrativo",ref,"",""].map(x=>({userEnteredValue:{stringValue:x}}))}],fields:"userEnteredValue"}});
   else {const h=hits[0],refs=String(h.r[8]||"").split(/\s*\|\s*/).filter(Boolean);if(!refs.includes(ref))req.push(endDioV015Update_(cep,h.i,8,[...refs,ref].join(" | ")));}
  }
  // Uma única operação atômica da API Sheets; nenhuma gravação parcial entre tabelas.
  const url="https://sheets.googleapis.com/v4/spreadsheets/"+encodeURIComponent(cfg.sourceId)+":batchUpdate";
  const resp=UrlFetchApp.fetch(url,{method:"post",contentType:"application/json",headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()},payload:JSON.stringify({requests:req}),muteHttpExceptions:true,followRedirects:false});
  if(resp.getResponseCode()!==200)return {ok:false,reason:"Gravação não confirmada. Confira a autorização da API Google Sheets e reabra a ficha."};
  return {ok:true,id};
 }catch(e){return {ok:false,reason:e.message||"Gravação não confirmada."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
