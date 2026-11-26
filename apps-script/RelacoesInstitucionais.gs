// Relações institucionais — fonte única: Cadastro Mestre Territorial.
// Acesso restrito ao administrador na homologação; não altera hierarquia nem permissões.
const RI_HEADERS_=["ID_RELACAO","ID_ENTIDADE_ORIGEM","TIPO_ENTIDADE_ORIGEM","TIPO_RELACAO","ID_ENTIDADE_DESTINO","TIPO_ENTIDADE_DESTINO","ID_DIOCESE_ORIGEM","ID_DIOCESE_DESTINO","INICIO_INFORMADO","FIM_INFORMADO","STATUS","DATA_REFERENCIA","FONTE","OBSERVACAO"];
const RI_TIPOS_=["CUIDADO_PASTORAL_CONFIADO_A","VINCULO_CONGREGACIONAL","RESPONSABILIDADE_ADMINISTRATIVA","COOPERACAO_INSTITUCIONAL","COOPERACAO_PASTORAL","OUTRA_RELACAO_INSTITUCIONAL"];
const RI_STATUS_=["DOCUMENTADO","A_CONFIRMAR","VIGENCIA_CONFIRMADA","ENCERRADO","ARQUIVADO"];
function riDataValida_(s){return /^\d{4}-\d{2}-\d{2}$/.test(s)&&!isNaN(Date.parse(s+"T00:00:00Z"))&&new Date(s+"T00:00:00Z").toISOString().slice(0,10)===s;}
function riEntidades_(ss){
 const mapa=new Map(),dioceses=new Map();
 function incluir(tipo,id,nome,diocese){
  if(!id||!nome||mapa.has(id))throw Error("Identificador institucional ausente ou duplicado.");
  if(diocese&&!dioceses.has(diocese))throw Error("Jurisdição não reconhecida.");
  mapa.set(id,{id,tipo,nome,diocese,nomeDiocese:dioceses.get(diocese)||""});
 }
 for(const r of lerTerritorio_(ss,"DIOCESES",["ID","Jurisdição"],2,1000)){if(dioceses.has(r[0]))throw Error("Jurisdição duplicada.");dioceses.set(r[0],r[1]);}
 for(const [id,nome] of dioceses)incluir("DIOCESE",id,nome,id);
 for(const r of lerTerritorio_(ss,"PARÓQUIAS",ER16_PAR_,33,5000))if(r[32]!=="ARQUIVADO")incluir("UNIDADE_PAROQUIAL",r[0],r[28]||r[3],r[1]);
 for(const r of lerTerritorio_(ss,"COMUNIDADES",ER16_COM_,21,5000))if(r[20]!=="ARQUIVADO")incluir("COMUNIDADE",r[0],r[14]||r[3],r[2]);
 for(const r of lerTerritorio_(ss,"LOCAIS_REFERENCIA",ER16_LOC_,14,5000))incluir("LOCAL_REFERENCIA",r[0],r[2],r[1]);
 for(const r of lerTerritorio_(ss,"ORGANIZAÇÕES",ER16_ORG_,13,5000))incluir("ORGANIZACAO",r[0],r[2],"");
 return mapa;
}
function riContexto_(cfg){
 const ss=SpreadsheetApp.openById(cfg.sourceId),tab=er16Tabela_(ss,"RELACOES_INSTITUCIONAIS",RI_HEADERS_,5000),idx=er16Index_(tab);
 return {ss,tab,idx,entidades:riEntidades_(ss)};
}
function riRevision_(row){return JSON.stringify(row.slice(0,14));}
function riMapear_(r,entidades){
 const origem=entidades.get(r[1]),destino=entidades.get(r[4]);
 const valida=!!origem&&!!destino&&origem.id!==destino.id&&origem.tipo===r[2]&&destino.tipo===r[5]&&origem.diocese===r[6]&&destino.diocese===r[7]&&RI_TIPOS_.includes(r[3])&&RI_STATUS_.includes(r[10]);
 return {id:r[0],origem:r[1],tipoOrigem:r[2],tipo:r[3],destino:r[4],tipoDestino:r[5],dioceseOrigem:r[6],dioceseDestino:r[7],inicio:r[8],fim:r[9],status:r[10],dataReferencia:r[11],fonte:r[12],observacao:r[13],nomeOrigem:origem?.nome||r[1],nomeDestino:destino?.nome||r[4],nomeDioceseOrigem:origem?.nomeDiocese||r[6],nomeDioceseDestino:destino?.nomeDiocese||r[7],valida,revisao:riRevision_(r)};
}
function riValidar_(v,entidades,existente){
 if(!v||typeof v!=="object")throw Error("Solicitação inválida.");
 const clean=(k,max,req=false)=>{const s=String(v[k]??"").trim();if((req&&!s)||s.length>max||/[\u0000-\u001f\u007f<>]/.test(s)||/^[=+@]/.test(s))throw Error("Campo inválido: "+k);return s;};
 const origem=entidades.get(clean("origem",100,true)),destino=entidades.get(clean("destino",100,true));
 if(!origem||!destino||origem.id===destino.id)throw Error("Selecione duas entidades distintas e válidas.");
 const tipo=clean("tipo",80,true),status=clean("status",80,true),inicio=clean("inicio",100),fim=clean("fim",100),data=clean("dataReferencia",10),fonte=clean("fonte",1000),obs=clean("observacao",45000);
 if(!RI_TIPOS_.includes(tipo)||!RI_STATUS_.includes(status))throw Error("Tipo ou situação inválida.");
 if(existente&&(existente[1]!==origem.id||existente[4]!==destino.id||existente[3]!==tipo))throw Error("Origem, destino e tipo são imutáveis. Arquive a relação incorreta e crie outra.");
 if(existente&&existente[10]==="ARQUIVADO")throw Error("Relação arquivada não pode ser alterada.");
 if(data&&!riDataValida_(data))throw Error("Data de referência inválida.");
 if(riDataValida_(inicio)&&riDataValida_(fim)&&fim<inicio)throw Error("Fim anterior ao início.");
 if(["DOCUMENTADO","VIGENCIA_CONFIRMADA","ENCERRADO"].includes(status)&&(!fonte||!data))throw Error("Informe fonte e data de referência.");
 if(status==="ENCERRADO"&&!fim)throw Error("Informe o fim do vínculo, ainda que aproximado.");
 if(["VIGENCIA_CONFIRMADA","ENCERRADO","ARQUIVADO"].includes(status)&&!obs)throw Error("Justifique a situação na observação.");
 return [existente?.[0]||"",origem.id,origem.tipo,tipo,destino.id,destino.tipo,origem.diocese,destino.diocese,inicio,fim,status,data,fonte,obs];
}
function listarRelacoesInstitucionais(){
 const cfg=autorizacaoTerritorial_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 try{const b=riContexto_(cfg);return {ok:true,entidades:[...b.entidades.values()],relacoes:[...b.idx.values()].map(x=>riMapear_(x.r,b.entidades)),tipos:RI_TIPOS_,situacoes:RI_STATUS_};}
 catch(e){return {ok:false,reason:"Cadastro de relações indisponível: "+e.message};}
}
function salvarRelacaoInstitucional(input){
 const cfg=autorizacaoTerritorial_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 let lock;try{
  lock=LockService.getScriptLock();if(!lock.tryLock(5000))throw Error("Cadastro ocupado. Tente novamente.");
  const b=riContexto_(cfg),id=String(input?.id||""),x=id?b.idx.get(id):null;
  if(id&&(!x||input.revisao!==riRevision_(x.r)))throw Error("Cadastro alterado ou inexistente. Atualize antes de editar.");
  const row=riValidar_(input,b.entidades,x?.r);
  if(!["ARQUIVADO","ENCERRADO"].includes(row[10])&&[...b.idx.values()].some(y=>y.r[0]!==row[0]&&y.r[10]!=="ARQUIVADO"&&y.r[10]!=="ENCERRADO"&&y.r[1]===row[1]&&y.r[3]===row[3]&&y.r[4]===row[4]))throw Error("Já existe uma relação não encerrada com esses identificadores.");
  if(!x){let n=0;for(const k of b.idx.keys()){const m=/^REL-INST-(\d+)$/.exec(k);if(m)n=Math.max(n,Number(m[1]));}row[0]="REL-INST-"+String(n+1).padStart(4,"0");}
  if(x&&riRevision_(row)===riRevision_(x.r))return {ok:true,id:row[0],alterado:false};
  if(x){
   const anterior=JSON.stringify(x.r.slice(8,13));
   const obsAnterior=x.r[13]&&!row[13].includes(x.r[13])?" | Observação anterior: "+x.r[13]:"";
   const nota="Revisão "+Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd HH:mm:ss")+"; valores anteriores (início, fim, status, data, fonte): "+anterior+obsAnterior;
   row[13]=[row[13],nota].filter(Boolean).join(" | ");
   if(row[13].length>45000)throw Error("Histórico excede o limite.");
  }
  const req=x?[{updateCells:{start:{sheetId:b.tab.sh.getSheetId(),rowIndex:x.i,columnIndex:0},rows:[{values:row.map(v=>({userEnteredValue:{stringValue:String(v)}}))}],fields:"userEnteredValue"}}]:[er16Append_(b.tab,row)];
  er16Batch_(cfg.sourceId,req);
  return {ok:true,id:row[0],alterado:true};
 }catch(e){return {ok:false,reason:e.message||"Gravação não confirmada."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
