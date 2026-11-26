// Complemento v0.15 para Código.gs CONSOLIDADO. Não contém configurações privadas.
// Substituir somente Gestao.gs; manter Código.gs, Cep.gs, Agenda.gs e manifesto existentes.
function resumoCoordenacoesAtivasV015_(equipes,comp,hoje){
 const dt=s=>!s||/^\d{4}-\d{2}-\d{2}$/.test(s)&&!isNaN(Date.parse(s+"T00:00:00Z"))&&new Date(s+"T00:00:00Z").toISOString().slice(0,10)===s;
 const pessoas=new Set(),coordenacoesTerritoriais=[];
 for(const e of equipes){if(!e[0]||!e[1]||e[5]==="ARQUIVADA")continue;
 const membros=comp.filter(c=>c[1]===e[0]&&c[2]&&c[6]==="ATIVO"&&!c[9]&&dt(c[4])&&dt(c[5])&&(!c[4]||c[4]<=hoje)&&(!c[5]||c[5]>=hoje));
 const situacao=e[5]==="VACANTE"?(membros.length?"INDEFINIDA":"VACANTE"):["DOCUMENTADA","INFORMADA","ATIVA","RATIFICADA"].includes(e[5])&&membros.length?"ATIVA":"INDEFINIDA";
 if(situacao==="ATIVA")for(const m of membros)pessoas.add(m[2]);
 coordenacoesTerritoriais.push({unidade:e[1],situacao});
 }return {coordenacoesTerritoriais,pessoasCoordenacao:pessoas.size,equipes:coordenacoesTerritoriais.filter(e=>e.situacao==="ATIVA").length};
}

function consultarResumoCadastroV015(){
 try {
  const auth=autorizacaoTerritorial_(), cfg=gestaoConfigAutorizada_();
  if(!auth||!cfg)return {ok:false,reason:"Acesso negado."};
  const base=consultarResumoCadastro();if(!base?.ok)return base;
  const ss=SpreadsheetApp.openById(cfg.operacionalId);
  const equipes=tabela_(ss,"COORDENACOES",["ID_COORDENACAO","ID_UNIDADE","NOME_EQUIPE","NIVEL","TIPO_EQUIPE","STATUS"]);
  const comp=tabela_(ss,"COMPOSICAO",["ID_COMPOSICAO","ID_COORDENACAO","ID_PESSOA","CARGO_FUNCAO","INICIO_VIGENCIA","FIM_VIGENCIA","STATUS","ATO_INDICACAO","DATA_RATIFICACAO","ENCERRADO_EM"]);
  const resumo=resumoCoordenacoesAtivasV015_(equipes,comp,Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd"));
  const territorial=SpreadsheetApp.openById(auth.sourceId);
  const rows=lerTerritorio_(territorial,"MISSAS LIBRAS",["ID referência","ID diocese (validar)","Cidade / região","Igreja / comunidade","Endereço público","Dia","Horário informado","Situação na fonte (não garantia atual)","Última referência","Fonte","Validação para visita","Tipo de celebração","ID unidade vinculada"],13,5000);
  const programacaoUnidades=Object.create(null);for(const r of rows){if(!r[0])continue;if(Object.prototype.hasOwnProperty.call(programacaoUnidades,r[0]))throw Error("ID de programação duplicado.");programacaoUnidades[r[0]]=r[12]||"";}
  return {...base,...resumo,programacaoUnidades};
 }catch(e){return {ok:false,reason:"Indicadores de coordenação indisponíveis. Verifique a estrutura da base."};}
}
/** Consultas separadas para a ficha ampliada, restritas ao titular institucional. */
const FICHA_COMPLEMENTOS_V015_ = {
 familia: [
 ["RELACOES_PESSOAS",["ID_RELACAO","ID_PESSOA","ID_PESSOA_RELACIONADA","NOME_RELACIONADO_SEM_CADASTRO","TIPO_RELACAO","CONTATO_COMPARTILHADO","INICIO","FIM","STATUS","ORIGEM","OBSERVACAO_RESTRITA"],[2,3,4,5,6,7,8]],
 ["CONTATOS_EMERGENCIA",["ID_CONTATO","ID_PESSOA","ID_PESSOA_CONTATO","NOME_CONTATO","GRAU_RELACAO","TELEFONE","WHATSAPP","MEIO_COMUNICACAO","ORDEM_PREFERENCIA","AUTORIZACAO_OU_BASE_REGISTRADA","DATA_CONFIRMACAO","STATUS","OBSERVACAO_RESTRITA"],[2,3,4,5,6,7,8,10,11]]
 ],
 socioeconomico: [
 ["PERFIL_SOCIOECONOMICO",["ID_AVALIACAO","ID_PESSOA","DATA_REFERENCIA","SITUACAO_TRABALHO","VINCULO_TRABALHISTA","OCUPACAO","ESCOLARIDADE","PROCURA_TRABALHO","CADUNICO_SITUACAO","CADUNICO_DATA_ATUALIZACAO","RECEBE_BENEFICIO","RENDA_DOMICILIAR_MENSAL_BRL","NUM_MORADORES_DOMICILIO","RENDA_DOMICILIAR_PER_CAPITA_BRL","FAIXA_PER_CAPITA_SM","SALARIO_MINIMO_REFERENCIA_BRL","CONDICAO_OCUPACAO_DOMICILIO","ARRANJO_DOMICILIAR","NECESSIDADE_APOIO_AUTODECLARADA","BASE_LEGAL_OU_AUTORIZACAO","ORIGEM","OBSERVACAO_RESTRITA","MOTIVO_FORA_FORCA_TRABALHO","INGRESSO_SERVICO_PUBLICO","FORMA_CESSAO_DOMICILIO"],[2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,22,23,24]],
 ["BENEFICIOS_PESSOAS",["ID_BENEFICIO","ID_PESSOA","ID_AVALIACAO","BENEFICIO_DECLARADO","TITULARIDADE","SITUACAO","DATA_REFERENCIA","ORIGEM","OBSERVACAO_RESTRITA"],[3,4,5,6]]
 ],
 historico: [
 ["PARTICIPACOES_EVENTOS",["ID_PARTICIPACAO","ID_PESSOA","ID_EVENTO","SITUACAO_INSCRICAO","PRESENCA","REFERENCIAS_ORIGEM","ORIGEM_REGISTRO","DATA_REGISTRO"],[2,3,4,7]]
 ]
};
function consultarComplementosFichaV015(id,secao){
 try{
  const titular=autorizacaoTerritorial_(),cfg=gestaoConfigAutorizada_();
  if(!titular||!cfg||typeof id!=="string"||!Object.prototype.hasOwnProperty.call(FICHA_COMPLEMENTOS_V015_,secao))return {ok:false,reason:"Consulta restrita."};
  const ss=SpreadsheetApp.openById(cfg.operacionalId),pessoas=tabela_(ss,"PESSOAS",PESSOAS_HEADERS_);
  if(!pessoas.some(r=>r[0]===id&&["CADASTRADO","PARCIAL"].includes(r[9])))return {ok:false,reason:"Ficha indisponível."};
  const nomes=new Map(pessoas.map(r=>[r[0],r[1]]));
  const eventos=secao==="historico"?new Map(tabela_(ss,"EVENTOS",["ID_EVENTO","NOME","ANO","DATA_INICIO","DATA_FIM","SITUACAO","ORIGEM"]).map(r=>[r[0],r[1]])):new Map();
  const grupos=FICHA_COMPLEMENTOS_V015_[secao].map(([aba,headers,indices])=>{
   const rows=tabela_(ss,aba,headers).filter(r=>r[1]===id);
   return {aba,registros:rows.map(r=>indices.filter(i=>r[i]!==""&&r[i]!=null).map(i=>({campo:headers[i],valor:i===2&&aba==="PARTICIPACOES_EVENTOS"?(eventos.get(r[i])||r[i]):i===2&&["RELACOES_PESSOAS","CONTATOS_EMERGENCIA"].includes(aba)?(nomes.get(r[i])||r[i]):String(r[i])})))};
  });
  return {ok:true,grupos};
 }catch(e){return {ok:false,reason:"Consulta complementar indisponível."};}
}
