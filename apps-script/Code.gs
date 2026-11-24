/**
 * POC somente leitura, dados fictícios. Vincular à planilha criada para o teste.
 * Não publicar para produção: identidade externa precisa de homologação.
 */
function configurarPoc_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss || ss.getSheetByName("LEIA_ME")?.getRange("B2").getDisplayValue()!=="POC FICTÍCIA — não cadastrar pessoas reais") throw new Error("Planilha POC não reconhecida.");
  PropertiesService.getScriptProperties().setProperty("POC_SHEET_ID",ss.getId());
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index").setTitle("Rede Pastoral do Surdo | POC");
}
function planilha_(){
  const id=PropertiesService.getScriptProperties().getProperty("POC_SHEET_ID");
  if(!id)throw new Error("Configuração pendente.");
  return SpreadsheetApp.openById(id);
}
function tabela_(ss,nome,colunas){
  const tab=ss.getSheetByName(nome);
  if(!tab)throw new Error("Estrutura indisponível.");
  const last=tab.getLastRow();
  if(last>1000)throw new Error("Limite da POC atingido.");
  if(last<1)throw new Error("Cabeçalho ausente.");
  const rows=tab.getRange(1,1,last,colunas.length).getDisplayValues();
  if(!colunas.every((v,i)=>rows[0][i]===v))throw new Error("Cabeçalho inválido.");
  return rows.slice(1).filter(r=>r.some(Boolean));
}
function normalizarEmail_(email){return String(email||"").trim().toLowerCase();}
function dadosPermitidos_(email,units,access,records,date){
  email=normalizarEmail_(email);
  if(!email || email.endsWith("@example.invalid") && date!=="TEST_ONLY") {
    // As identidades de fixture só são aceitas no teste unitário com marcador explícito.
    return {ok:false};
  }
  const hoje=date==="TEST_ONLY"?"2026-11-24":date;
  const validDate=s=>/^\d{4}-\d{2}-\d{2}$/.test(s);
  const byId=new Map();
  for(const u of units){
    if(!u[0]||byId.has(u[0]))return {ok:false};
    byId.set(u[0],u);
  }
  const parents={REGIONAL:"",DIOCESE:"REGIONAL",PAROQUIA:"DIOCESE"};
  function lineage(id){
    const path=[],seen=new Set();
    while(id){
      const u=byId.get(id);
      if(!u||u[4]!=="SIM"||seen.has(id)||!Object.prototype.hasOwnProperty.call(parents,u[2]))return [];
      seen.add(id);path.push(id);
      if(u[2]==="REGIONAL")return u[1]?[]:path;
      if(!u[1]||byId.get(u[1])?.[2]!==parents[u[2]])return [];
      id=u[1];
    }
    return [];
  }
  const allowed=new Set();
  for(const a of access){
    if(normalizarEmail_(a[0])!==email||a[3]!=="SIM"||!["COORDENACAO","SECRETARIA","AGENTE"].includes(a[2])||!validDate(a[4])||!validDate(a[5])||a[4]>hoje||a[5]<hoje||!lineage(a[1]).length)continue;
    // Agentes veem somente a própria unidade; coordenação/secretaria veem descendentes.
    for(const u of units){const path=lineage(u[0]);if(path.length&&(a[2]==="AGENTE"?u[0]===a[1]:path.includes(a[1])))allowed.add(u[0]);}
  }
  if(!allowed.size)return {ok:false};
  return {ok:true,units:units.filter(u=>allowed.has(u[0])).map(u=>({id:u[0],name:u[3],type:u[2]})),records:records.filter(r=>allowed.has(r[1])).map(r=>({id:r[0],unit:r[1],title:r[2],description:r[3]}))};
}
function consultarPainel(){
  // Não aceitar e-mail, unidade ou papel do cliente. Identidade sempre obtida do servidor.
  const email=normalizarEmail_(Session.getActiveUser().getEmail());
  if(!email)return {ok:false,reason:"Identidade Google não confirmada. Acesso negado."};
  try{
    const ss=planilha_();
    const units=tabela_(ss,"UNIDADES",["ID","PAI_ID","TIPO","NOME","ATIVA"]);
    const access=tabela_(ss,"ACESSOS",["EMAIL","UNIDADE_ID","FUNCAO","ATIVO","INICIO_ISO","FIM_ISO"]);
    const records=tabela_(ss,"REGISTROS_TESTE",["ID","UNIDADE_ID","TITULO","DESCRICAO"]);
    const date=Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd");
    const result=dadosPermitidos_(email,units,access,records,date);
    return result.ok?result:{ok:false,reason:"Conta não autorizada para a POC."};
  }catch(e){return {ok:false,reason:"Serviço de teste indisponível."};}
}
