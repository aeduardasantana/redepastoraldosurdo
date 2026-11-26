// Programações — normalização de vínculos e gravação compatíveis com projeto consolidado.
const PROG16_HEADERS_=["ID referência","ID diocese (validar)","Cidade / região","Igreja / comunidade","Endereço público","Dia","Horário informado","Situação na fonte (não garantia atual)","Última referência","Fonte","Validação para visita","Tipo de celebração","ID unidade vinculada","Situação da programação","Pastoral do Surdo no local","Fonte da situação pastoral","Programação","Status cadastro","ID pastoral","Acessibilidade Libras","Data específica","Horário final","Publicar no portal"];
function prog16Destino_(data,id){
 if(!id)return null;
 const all=[...(data.paroquias||[]),...(data.comunidades||[]),...(data.locais||[])].filter(x=>x.id===id);
 if(all.length!==1)return null;
 const u=all[0];if(!(data.dioceses||[]).some(d=>d.id===u.diocese&&d.id!=="D13"))return null;
 if(u.unidade&&!(data.paroquias||[]).concat(data.comunidades||[]).some(p=>p.id===u.unidade&&p.diocese===u.diocese))return null;
 return u;
}
function prog16Normalizar_(data){
 const ds=new Set((data.dioceses||[]).filter(d=>d.id!=="D13").map(d=>d.id));
 return {...data,missas:(data.missas||[]).map(m=>{
  const raw=String(m.unidadeInformada??m.unidade??""),u=prog16Destino_(data,raw),valido=!!u&&u.diocese===m.diocese;
  return {...m,unidadeInformada:raw,unidade:valido?raw:"",vinculoPendente:!ds.has(m.diocese)||!!raw&&!valido,nivelVinculo:valido?"LOCAL":ds.has(m.diocese)&&!raw?"DIOCESANO":"A_REVISAR"};
 })};
}
function consultarTerritorioV016(){
 const r=consultarTerritorio();return r?.ok?{...r,data:prog16Normalizar_(r.data)}:r;
}
function prog16Validar_(data,values,old){
 const dioc=(data.dioceses||[]).find(d=>d.id===values.diocese&&d.id!=="D13");
 if(!dioc||!values.programacao)return "Informe jurisdição e programação.";
 const u=prog16Destino_(data,values.unidade);
 if(values.unidade&&(!u||u.diocese!==dioc.id))return "Unidade inexistente ou de outra jurisdição.";
 if(!values.nome&&u)values.nome=u.nomeDivulgacao||u.nome;
 if(!values.nome&&!values.unidade)return "Informe a referência da programação.";
 if(u){if(!values.cidade)values.cidade=u.cidade||u.municipio||"";if(!values.logradouro)values.logradouro=u.logradouro||"";}
 if(!values.statusProgramacao)values.statusProgramacao="A CONFIRMAR";
 if(!["CONFIRMADA","RECORRENTE INFORMADA","A CONFIRMAR","DIVERGENTE"].includes(values.statusProgramacao))return "Situação inválida.";
 if(!values.tipo)values.tipo=values.programacao;
 // Uma alteração de vínculo não pode deslocar uma pastoral já associada sem revisão.
 if(old?.[18]){
  const ps=(data.pastorais||[]).filter(p=>p.id===old[18]);
  if(ps.length!==1||ps[0].diocese!==dioc.id||ps[0].unidade!==values.unidade)return "Revise antes o vínculo pastoral existente; nenhuma alteração foi gravada.";
 }
 return "";
}
function salvarProgramacaoV016(input){
 const cfg=autorizacaoTerritorial_();if(!cfg||input?.tipo!=="missas")return {ok:false,reason:"Acesso negado ou tipo inválido."};
 let lock;try{
  const def=FICHAS_TERRITORIAIS_.missas,campos=input.campos;
  if(!campos||typeof campos!=="object"||Array.isArray(campos)||Object.keys(campos).some(k=>!Object.prototype.hasOwnProperty.call(def.fields,k)))return {ok:false,reason:"Campos inválidos."};
  const values={};for(const k of Object.keys(def.fields)){const v=campos[k]??"";if(typeof v!=="string"||v.length>1000||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(v.trim()))return {ok:false,reason:"Campo inválido."};values[k]=v.trim();}
  lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};
  const t=consultarTerritorioV016();if(!t?.ok)return t;
  const {sh,rows}=fichaTerritorialRows_(SpreadsheetApp.openById(cfg.sourceId),def);
  if(!PROG16_HEADERS_.every((h,j)=>rows[0][j]===h))return {ok:false,reason:"Cabeçalho incompatível."};
  const hits=input.id?rows.map((r,i)=>({r,i})).filter(x=>x.i>0&&x.r[0]===input.id):[];
  if(input.id&&(hits.length!==1||hits[0].r[17]==="ARQUIVADO"))return {ok:false,reason:"Registro ausente ou duplicado."};
  const old=hits[0]?.r,idx=hits[0]?.i;
  if(old&&JSON.stringify(fichaTerritorialDados_(old,def))!==input.revisao)return {ok:false,reason:"Ficha alterada; reabra antes de salvar."};
  const err=prog16Validar_(t.data,values,old);if(err)return {ok:false,reason:err};
  const keys=["diocese","unidade","nome","programacao","dia","horario"];
  if(rows.slice(1).some(r=>r[0]!==input.id&&r[17]!=="ARQUIVADO"&&keys.every(k=>normalizarChave_(r[def.fields[k]])===normalizarChave_(values[k]))))return {ok:false,reason:"Programação equivalente já cadastrada."};
  const id=old?old[0]:Utilities.getUuid();
  if(!old){const r=Array(23).fill("");r[0]=id;r[17]="ATIVO";r[8]=new Date().toISOString().slice(0,10);r[9]="Registro administrativo";for(const [k,j] of Object.entries(def.fields))r[j]=values[k];const target=sh.getLastRow()+1;if(target>5000)return {ok:false,reason:"Limite da base."};if(target>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),100);sh.getRange(target,1,1,23).setNumberFormat("@").setValues([r]);}
  else {
   const req=[];for(const [k,j] of Object.entries(def.fields))if(values[k]!==old[j]){if(sh.getRange(idx+1,j+1).getFormula())return {ok:false,reason:"Campo calculado protegido."};req.push({updateCells:{range:{sheetId:sh.getSheetId(),startRowIndex:idx,endRowIndex:idx+1,startColumnIndex:j,endColumnIndex:j+1},rows:[{values:[{userEnteredValue:{stringValue:values[k]}}]}],fields:"userEnteredValue"}});}
   if(req.length){const resp=UrlFetchApp.fetch("https://sheets.googleapis.com/v4/spreadsheets/"+encodeURIComponent(cfg.sourceId)+":batchUpdate",{method:"post",contentType:"application/json",headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()},payload:JSON.stringify({requests:req}),muteHttpExceptions:true,followRedirects:false});if(resp.getResponseCode()!==200)return {ok:false,reason:"Gravação não confirmada. Verifique API Sheets e reabra a ficha."};}
  }
  SpreadsheetApp.flush();return {ok:true,id};
 }catch(e){return {ok:false,reason:"Gravação não confirmada. Confira a ficha."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
