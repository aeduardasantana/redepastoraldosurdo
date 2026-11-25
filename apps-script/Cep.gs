// CEP — v2026.11.25. Módulo complementar independente; não repetir Código.gs/Gestao.gs.
function redeNormalizarCep_(value){
  const s=String(value||"").trim();
  return /^\d{5}-?\d{3}$/.test(s)?s.replace(/-/g,""):"";
}
function redeCepPostal_(raw,cep){
  if(!raw||raw.erro===true||raw.erro==="true"||redeNormalizarCep_(raw.cep)!==cep)return null;
  const s=(v,max)=>String(v||"").trim().slice(0,max);
  const uf=s(raw.uf,2);
  if(!/^(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|SC|SP|SE|TO)$/.test(uf)||!s(raw.localidade,120))return null;
  return {cep:cep.slice(0,5)+"-"+cep.slice(5),logradouro:s(raw.logradouro,250),localidade:s(raw.bairro,120),municipio:s(raw.localidade,120),uf,ibge:/^\d{7}$/.test(String(raw.ibge||""))?String(raw.ibge):"",pais:"Brasil"};
}
function consultarCepRede(value){
  const cfg=autorizacaoTerritorial_();
  if(!cfg)return {ok:false,reason:"Consulta restrita."};
  const cep=redeNormalizarCep_(value);
  if(!cep)return {ok:false,reason:"Informe um CEP com oito dígitos."};
  const formatado=cep.slice(0,5)+"-"+cep.slice(5);
  const headers=["CEP","LOGRADOURO_POSTAL","BAIRRO_POSTAL","MUNICIPIO_POSTAL","UF","CODIGO_IBGE","STATUS_VALIDACAO","FONTE_ORIGEM","REFERENCIAS_TERRITORIAIS","FONTE_CONSULTA","DATA_CONSULTA"];
  try{
    const ss=SpreadsheetApp.openById(cfg.sourceId),sh=ss.getSheetByName("CEP_REFERENCIA");
    if(!sh||sh.getLastRow()<1||!headers.every((v,i)=>sh.getRange(1,1,1,headers.length).getDisplayValues()[0][i]===v))return {ok:false,reason:"Base postal indisponível."};
    const read=()=>sh.getRange(1,1,sh.getLastRow(),headers.length).getDisplayValues();
    let rows=read(),hits=rows.slice(1).filter(r=>redeNormalizarCep_(r[0])===cep);
    if(hits.length>1)return {ok:false,reason:"CEP duplicado na base. Solicite revisão."};
    const old=hits[0],today=Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd");
    const validRecent=r=>r&&r[6]==="Confirmado"&&r[9]&&/^\d{4}-\d{2}-\d{2}$/.test(r[10])&&(Date.parse(today)-Date.parse(r[10]))>=0&&(Date.parse(today)-Date.parse(r[10]))<=180*86400000;
    if(validRecent(old)){const data=redeCepPostal_({cep:old[0],logradouro:old[1],bairro:old[2],localidade:old[3],uf:old[4],ibge:old[5]},cep);if(data)return {ok:true,data,source:old[9],cached:true};}
    // Uma consulta manual por solicitação; sem consulta em lote e sem transmitir outros dados.
    const url="https://viacep.com.br/ws/"+cep+"/json/";
    const resp=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:false});
    if(resp.getResponseCode()!==200)return {ok:false,reason:"Serviço postal indisponível. Preencha manualmente."};
    const raw=JSON.parse(resp.getContentText()),data=redeCepPostal_(raw,cep);
    if(!data)return {ok:false,reason:"CEP não localizado ou resposta incompleta. Confira a informação."};
    const lock=LockService.getScriptLock();if(!lock.tryLock(3000))return {ok:true,data,source:url,cached:false,warning:"Consulta realizada; base local ocupada."};
    try{
      rows=read();const match=rows.map((r,i)=>({r,i})).filter(x=>x.i>0&&redeNormalizarCep_(x.r[0])===cep);
      if(match.length>1)return {ok:true,data,source:url,cached:false,warning:"CEP duplicado na base; não gravado."};
      const item=match[0];
      if(item&&item.r[6]==="Divergente")return {ok:true,data,source:url,cached:false,warning:"Divergência anterior preservada para revisão."};
      if(item&&validRecent(item.r)&&item.r[9]===url)return {ok:true,data,source:url,cached:false};
      // Registrar exclusivamente dados postais públicos, sem ID de pessoa ou endereço predial.
      const row=[formatado,data.logradouro,data.localidade,data.municipio,data.uf,data.ibge,"Confirmado",item?item.r[7]:"Consulta pontual",item?item.r[8]:"",url,today];
      const next=item?item.i+1:sh.getLastRow()+1;
      if(next>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),20);
      sh.getRange(next,1,1,headers.length).setNumberFormat("@").setValues([row]);SpreadsheetApp.flush();
      return {ok:true,data,source:url,cached:false};
    }finally{lock.releaseLock();}
  }catch(e){return {ok:false,reason:"Não foi possível consultar o CEP. Confira autorização e conexão ou preencha manualmente."};}
}
