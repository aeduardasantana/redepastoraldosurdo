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
 const nivel=clean(input.nivel,40),unidade=clean(input.unidade,100),nome=clean(input.nome,180),fonte=clean(input.fonte||"Cadastro administrativo",500);
 const valido=nivel==="REGIONAL"?unidade==="CO":nivel==="ARQUIDIOCESANO_DIOCESANO"?data.dioceses.some(x=>x.id===unidade&&x.id!=="D13"):nivel==="PAROQUIAL_REPRESENTANTE"?data.paroquias.some(x=>x.id===unidade):false;
 if(!valido)throw Error("Nível ou unidade fora do escopo regional de homologação.");
 if(existentes.some(r=>r[1]===unidade&&r[3]===nivel&&r[5]!=="ARQUIVADA"))throw Error("Já existe registro para este nível e unidade.");
 return {nivel,unidade,nome,fonte,tipo:nivel==="PAROQUIAL_REPRESENTANTE"?"REPRESENTACAO":"COORDENACAO"};
}

const DOCUMENTOS_HEADERS_=["ID_DOCUMENTO","TITULO","CATEGORIA","URL","VISIBILIDADE","STATUS","ORDEM","ANO","MANDATO","SITUACAO_DOCUMENTAL"];
const DOCS_ACESSOS_HEADERS_=["ID_ACESSO","ID_PESSOA","IDENTIDADE_GOOGLE","ID_UNIDADE","PERFIL_AUTORIZADO","INICIO_VIGENCIA","FIM_VIGENCIA","STATUS","APROVADO_POR_ID","APROVADO_EM","REVOGADO_EM"];
const DOCS_COMPOSICAO_HEADERS_=["ID_COMPOSICAO","ID_COORDENACAO","ID_PESSOA","CARGO_FUNCAO","INICIO_VIGENCIA","FIM_VIGENCIA","STATUS","ATO_INDICACAO","DATA_RATIFICACAO","ENCERRADO_EM","MOTIVO_ADMINISTRATIVO_CODIFICADO"];
function membroRegionalDocumentos_(email,acessos,composicao,equipes,hoje){
 const dt=s=>/^\d{4}-\d{2}-\d{2}$/.test(s)&&!isNaN(Date.parse(s+"T00:00:00Z"))&&new Date(s+"T00:00:00Z").toISOString().slice(0,10)===s;
 const vig=(ini,fim)=>dt(ini)&&ini<=hoje&&(!fim||dt(fim)&&fim>=hoje);
 const e=normalizarEmail_(email);if(!e||!dt(hoje))return false;
 const validTeams=new Set(equipes.filter(r=>r[1]==="CO"&&r[3]==="REGIONAL"&&["DOCUMENTADA","ATIVA","RATIFICADA"].includes(r[5])).map(r=>r[0]));
 const grants=acessos.filter(r=>normalizarEmail_(r[2])===e&&r[1]&&r[3]==="CO"&&r[4]==="COORDENACAO_REGIONAL"&&r[7]==="ATIVO"&&r[8]&&r[9]&&!r[10]&&vig(r[5],r[6]));
 return grants.some(a=>composicao.some(c=>validTeams.has(c[1])&&c[2]===a[1]&&c[3]&&c[6]==="ATIVO"&&!c[9]&&vig(c[4],c[5])));
}
function consultarDocumentos(){
 const publico=[{id:"DOC-REG-2026",titulo:"Regimento Interno da Pastoral do Surdo — 2026",categoria:"Regimento",url:"https://drive.google.com/file/d/1cT0LzVCe0bwhqhtRbaoM_Y4_SzDxjY0z/view"}];
 const denied={ok:true,publico,regional:false,restritos:[]};
 try{
  const email=normalizarEmail_(Session.getActiveUser().getEmail());if(!email)return denied;
  const cfg=typeof gestaoConfig_==="function"?gestaoConfig_():null;if(!cfg?.operacionalId)return denied;
  const ss=SpreadsheetApp.openById(cfg.operacionalId);
  const admin=!!autorizacaoTerritorial_();
  if(!admin&&!membroRegionalDocumentos_(email,tabela_(ss,"ACESSOS",DOCS_ACESSOS_HEADERS_),tabela_(ss,"COMPOSICAO",DOCS_COMPOSICAO_HEADERS_),tabela_(ss,"COORDENACOES",COORD_HEADERS_),Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd")))return denied;
  const docs=tabela_(ss,"DOCUMENTOS",DOCUMENTOS_HEADERS_).filter(r=>r[4]==="COORDENACAO_REGIONAL"&&r[5]==="ATIVO"&&r[0]&&r[1]&&/^https:\/\/(?:docs\.google\.com\/(?:document|spreadsheets|presentation)\/d\/[A-Za-z0-9_-]+\/(?:edit|view)|drive\.google\.com\/file\/d\/[A-Za-z0-9_-]+\/view)$/.test(r[3])).sort((a,b)=>(Number(a[6])||0)-(Number(b[6])||0)).map(r=>({id:r[0],titulo:r[1],categoria:r[2],url:r[3],ano:r[7],mandato:r[8],situacao:r[9]}));
  return {ok:true,publico,regional:true,restritos:docs};
 }catch(e){return denied;}
}

function consultarGestao(){
 try{
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId);
 const equipes=tabela_(ss,"COORDENACOES",COORD_HEADERS_).map(r=>({id:r[0],unidade:r[1],nome:r[2],nivel:r[3],tipo:r[4],status:r[5],fonte:r[7],validacao:r[8]}));
 const q=tabela_(SpreadsheetApp.openById(cfg.questionarioId),"PERGUNTAS",QUEST_HEADERS_).map(r=>({id:r[0],bloco:r[1],pergunta:r[2],tipo:r[3],opcoes:r[4],regra:r[5]}));
 const cargos=tabela_(ss,"CARGOS_REGIMENTAIS",["NIVEL_REFERENCIA","CARGO","ORIGEM","PROVIMENTO","OBSERVACAO"]).map(r=>({nivel:r[0],cargo:r[1],origem:r[2],provimento:r[3],obs:r[4]}));
 const pessoas=new Map(tabela_(ss,"PESSOAS",PESSOAS_HEADERS_).filter(r=>["CADASTRADO","PARCIAL"].includes(r[9])).map(r=>[r[0],r]));const referencias=tabela_(ss,"COMPOSICAO_REFERENCIAS",REF_COORD_HEADERS_).map(r=>{const p=pessoas.get(r[9]);return {id:r[0],equipe:r[1],nome:p?.[1]||r[2],pessoa:p?.[0]||"",telefone:p&&p[8]!=="SEM_CONTATO"?p[7]:"",preferido:p?.[2]||"",setor:r[3],funcao:r[4],origem:r[5],vigencia:r[6],vinculado:!!p,instituicao:/institui[cç][aã]o/i.test(r[4])};});
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

const CAMPOS_FICHA_UNICA_=[{"id":"RA_RESIDENCIAL","bloco":"Endereço","label":"Região administrativa (DF)","tipo":"Texto","opcoes":[]},{"id":"DATA_BATISMO","bloco":"Sacramentos","label":"Data do Batismo","tipo":"Data","opcoes":[]},{"id":"LOCAL_BATISMO","bloco":"Sacramentos","label":"Igreja ou local do Batismo","tipo":"Texto","opcoes":[]},{"id":"DATA_PRIMEIRA_COMUNHAO","bloco":"Sacramentos","label":"Data da Primeira Comunhão","tipo":"Data","opcoes":[]},{"id":"DATA_CRISMA","bloco":"Sacramentos","label":"Data da Crisma","tipo":"Data","opcoes":[]},{"id":"DATA_MATRIMONIO","bloco":"Sacramentos","label":"Data do Matrimônio","tipo":"Data","opcoes":[]},{"id":"P08","bloco":"Acessibilidade","label":"Como você prefere se identificar, se desejar?","tipo":"Seleção","opcoes":["Pessoa surda","pessoa com deficiência auditiva","surdocega","ouvinte","outra","prefiro não informar"]},{"id":"P09","bloco":"Acessibilidade","label":"Qual forma de comunicação facilita sua participação?","tipo":"Múltipla escolha","opcoes":["Libras","português escrito","português oral","leitura labial","comunicação tátil","recursos visuais","outra","prefiro não informar"]},{"id":"P10","bloco":"Acessibilidade","label":"Como prefere receber os conteúdos pastorais?","tipo":"Múltipla escolha","opcoes":["Vídeo em Libras","texto simples","legendas","áudio","encontro presencial","outro"]},{"id":"P11","bloco":"Participação","label":"Você participa de celebrações ou atividades da Igreja atualmente?","tipo":"Seleção","opcoes":["Regularmente","às vezes","não atualmente","prefiro não informar"]},{"id":"P12","bloco":"Participação","label":"O que dificulta sua participação?","tipo":"Múltipla escolha","opcoes":["Falta de intérprete","horários","transporte","distância","comunicação não acessível","falta de informação","acolhimento","disponibilidade","outro","nenhuma","prefiro não informar"]},{"id":"P13","bloco":"Participação","label":"Você gostaria de participar de alguma atividade pastoral?","tipo":"Múltipla escolha","opcoes":["Missas/novena","catequese","formação","convivência","voluntariado","conhecer a pastoral","não agora"]},{"id":"P14","bloco":"Sacramentos","label":"Você já recebeu o Batismo?","tipo":"Seleção","opcoes":["Sim","não","não sei","prefiro não informar"]},{"id":"P15","bloco":"Sacramentos","label":"Você já recebeu a Primeira Comunhão?","tipo":"Seleção","opcoes":["Sim","não","não sei","prefiro não informar"]},{"id":"P16","bloco":"Sacramentos","label":"Você já recebeu a Crisma (Confirmação)?","tipo":"Seleção","opcoes":["Sim","não","não sei","prefiro não informar"]},{"id":"P17","bloco":"Sacramentos","label":"Você já recebeu pela primeira vez o sacramento da Reconciliação (Confissão)?","tipo":"Seleção","opcoes":["Sim","não","não sei","prefiro não informar"]},{"id":"P18","bloco":"Sacramentos","label":"Você já recebeu o sacramento do Matrimônio?","tipo":"Seleção","opcoes":["Sim","não","não se aplica","não sei","prefiro não informar"]},{"id":"P19","bloco":"Sacramentos","label":"Você já recebeu a Unção dos Enfermos?","tipo":"Seleção","opcoes":["Sim","não","não se aplica","não sei","prefiro não informar"]},{"id":"P20","bloco":"Sacramentos","label":"Você já recebeu o sacramento da Ordem?","tipo":"Seleção","opcoes":["Sim","não","não se aplica","prefiro não informar"]},{"id":"P21","bloco":"Barreiras","label":"Para cada sacramento de iniciação ainda não recebido, você gostaria de recebê-lo ou conhecer melhor a preparação?","tipo":"Seleção por sacramento","opcoes":["Sim","talvez","não neste momento","já em preparação","prefiro não informar"]},{"id":"P22","bloco":"Barreiras","label":"Alguma dificuldade de acessibilidade impediu ou atrasou sua catequese ou preparação para esse sacramento?","tipo":"Múltipla escolha por sacramento","opcoes":["Falta de intérprete de Libras","catequista sem comunicação acessível","materiais sem Libras/legendas","falta de turma acessível","distância/transporte","horários","desconhecia a oferta","outra","nenhuma","não sei","prefiro não informar"]},{"id":"P23","bloco":"Barreiras","label":"Você encontrou dificuldades de comunicação na celebração ou no acesso a um sacramento?","tipo":"Múltipla escolha","opcoes":["Falta de intérprete","ausência de sacerdote que se comunique em Libras","orientações não acessíveis","falta de informação","outra","nenhuma","prefiro não informar"]},{"id":"P24","bloco":"Barreiras","label":"Você gostaria de encontrar uma catequese ou preparação sacramental acessível?","tipo":"Seleção","opcoes":["Sim","talvez","não agora","já participo"]},{"id":"P25","bloco":"Barreiras","label":"Qual recurso seria mais útil na preparação e nas celebrações?","tipo":"Múltipla escolha","opcoes":["Catequese em Libras","intérprete","catequista surdo","materiais bilíngues","vídeos com legendas","sacerdote com Libras","turma acessível","outro"]},{"id":"P26","bloco":"Acompanhamento","label":"Você deseja receber informações sobre locais e horários acessíveis para a Confissão?","tipo":"Seleção","opcoes":["Sim","não","depois"]},{"id":"P30","bloco":"Acompanhamento","label":"Você deseja que uma pessoa da pastoral entre em contato para ajudá-lo(a) a encontrar atendimento acessível?","tipo":"Seleção","opcoes":["Sim","não","depois"]},{"id":"P31","bloco":"Acompanhamento","label":"Qual ajuda você solicita neste momento?","tipo":"Múltipla escolha","opcoes":["Catequese","informações sobre sacramentos","celebrações em Libras","transporte/horário","contato com equipe local","outra","nenhuma"]},{"id":"P32","bloco":"Voluntariado","label":"Você deseja colaborar com a pastoral?","tipo":"Múltipla escolha","opcoes":["Não agora","agente pastoral","catequese","interpretação","comunicação","apoio a eventos","formação","outro"]},{"id":"P33","bloco":"Voluntariado","label":"Quais disponibilidades ou competências deseja informar?","tipo":"Texto/seleção","opcoes":[]},{"id":"P40","bloco":"Território","label":"Você tem vínculo com uma paróquia católica?","tipo":"Seleção","opcoes":["Sim","não","não sei","prefiro não informar"]},{"id":"P41","bloco":"Contexto religioso","label":"Como você se identifica em relação à religião, se desejar informar?","tipo":"Seleção","opcoes":["Católica","outra denominação cristã","outra religião","sem religião","não sei","prefiro não informar"]},{"id":"P42","bloco":"Contexto religioso","label":"Se desejar, qual comunidade religiosa ou denominação frequenta?","tipo":"Texto condicional","opcoes":[]},{"id":"P43","bloco":"Sacramentos","label":"Em qual tradição ou Igreja ocorreu seu Batismo, se recebeu?","tipo":"Seleção","opcoes":["Igreja Católica","outra Igreja cristã","não sei","não recebi","prefiro não informar"]},{"id":"P44","bloco":"Comunicação","label":"Como avalia sua compreensão do português escrito (leitura)?","tipo":"Seleção","opcoes":["Nenhuma","básica","intermediária","avançada","fluente","não sei","prefiro não informar"]},{"id":"P45","bloco":"Comunicação","label":"Como avalia sua expressão em português escrito?","tipo":"Seleção","opcoes":["Nenhuma","básica","intermediária","avançada","fluente","não sei","prefiro não informar"]},{"id":"P46","bloco":"Comunicação","label":"Como avalia sua compreensão da Libras?","tipo":"Seleção","opcoes":["Nenhuma","básica","intermediária","avançada","fluente","não sei","prefiro não informar"]},{"id":"P47","bloco":"Comunicação","label":"Como avalia sua expressão em Libras?","tipo":"Seleção","opcoes":["Nenhuma","básica","intermediária","avançada","fluente","não sei","prefiro não informar"]},{"id":"P48","bloco":"Comunicação","label":"Como avalia sua compreensão do português oral?","tipo":"Seleção","opcoes":["Nenhuma","básica","intermediária","avançada","fluente","não se aplica","prefiro não informar"]},{"id":"P49","bloco":"Comunicação","label":"Você utiliza comunicação oral (oralização)?","tipo":"Seleção","opcoes":["Sim","parcialmente/em algumas situações","não","não se aplica","prefiro não informar"]},{"id":"P50","bloco":"Acessibilidade","label":"Se desejar informar, qual grau de perda auditiva lhe foi indicado?","tipo":"Seleção","opcoes":["Leve","moderada","severa","profunda","outra classificação","não sei","não se aplica","prefiro não informar"]},{"id":"P51","bloco":"Participação","label":"Qual sua relação com a Pastoral do Surdo atualmente?","tipo":"Seleção","opcoes":["Beneficiário(a)","interessado(a)","familiar","intérprete","membro voluntário","outra","nenhuma","prefiro não informar"]}];

const CAMPOS_ECLESIAIS_=[{"id":"CONDICAO_ECLESIAL","label":"Condição eclesial católica","opcoes":["Leigo","Ministro ordenado","Não se aplica","Prefiro não informar"],"bloco":"Contexto religioso","tipo":"Seleção"},{"id":"GRAU_ORDEM","label":"Grau da Ordem","opcoes":["Diácono","Presbítero","Bispo","Não se aplica","Prefiro não informar"],"bloco":"Contexto religioso","tipo":"Seleção"},{"id":"VIDA_CONSAGRADA","label":"Vida consagrada","opcoes":["Sim","Não","Não se aplica","Prefiro não informar"],"bloco":"Contexto religioso","tipo":"Seleção"},{"id":"FORMACAO_VOCACIONAL","label":"Etapa de formação vocacional","opcoes":["Seminarista","Postulante","Noviço","Outra","Não se aplica","Prefiro não informar"],"bloco":"Contexto religioso","tipo":"Seleção"},{"id":"INSTITUTO_CONGREGACAO","label":"Instituto ou congregação (nome informado)","opcoes":[],"bloco":"Contexto religioso","tipo":"Texto"},{"id":"FORMA_TRATAMENTO","label":"Forma de tratamento","opcoes":[],"bloco":"Contexto religioso","tipo":"Texto"}];
const TODOS_CAMPOS_FICHA_=[...CAMPOS_FICHA_UNICA_,...CAMPOS_ECLESIAIS_];
const META_FICHA_UNICA_=["ORIGEM_CADASTRO","CRIADO_POR","ATUALIZADO_POR","ORIGEM_ULTIMA_EDICAO","AUTORIZACAO_DADOS_OPCIONAIS_EM"];
function validarCamposFichaUnica_(input,old,hoje){
 const values=Object.create(null),dados=input?.camposFicha;
 if(dados!==undefined&&(!dados||typeof dados!=="object"||Array.isArray(dados)))throw Error("Campos inválidos.");
 if(dados&&Object.keys(dados).some(k=>!TODOS_CAMPOS_FICHA_.some(f=>f.id===k)))throw Error("Campo desconhecido.");
 let alteracao=false;
 for(const f of TODOS_CAMPOS_FICHA_){
  const anterior=String(old?.[PESSOAS_HEADERS_.indexOf(f.id)]||"");let v=dados&&Object.prototype.hasOwnProperty.call(dados,f.id)?dados[f.id]:anterior;
  if(typeof v!=="string"||v.length>2500||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(v.trim()))throw Error("Confira: "+f.label);
  v=v.trim();
  if(v&&dados&&Object.prototype.hasOwnProperty.call(dados,f.id)){
   if(f.tipo==="Data"){const d=new Date(v+"T00:00:00Z");if(!/^\d{4}-\d{2}-\d{2}$/.test(v)||isNaN(d.getTime())||d.toISOString().slice(0,10)!==v||v>hoje)throw Error("Data inválida: "+f.label);}
   else if(f.tipo.includes("por sacramento")){let o;try{o=JSON.parse(v)}catch(e){throw Error("Confira: "+f.label);}if(!o||typeof o!=="object"||Array.isArray(o)||Object.keys(o).some(k=>!["Batismo","Primeira Comunhão","Crisma"].includes(k)))throw Error("Confira: "+f.label);for(const val of Object.values(o)){if(typeof val!=="string"||val.split(" | ").some(x=>x&&!f.opcoes.includes(x)))throw Error("Confira: "+f.label);}}
   else if(f.opcoes.length&&v.split(" | ").some(x=>!f.opcoes.includes(x)))throw Error("Confira: "+f.label);
  }
  if(v&&v!==anterior&&f.id!=="RA_RESIDENCIAL")alteracao=true;values[f.id]=v;
 }
 if(alteracao&&input.autorizacaoOpcionais!==true)throw Error("Confirme a autorização específica para registrar as informações opcionais de acompanhamento.");
 return {values,alteracao};
}
function metadadosFichaUnica_(r){const i=23+CAMPOS_FICHA_UNICA_.length;return {origem:r[i]||(/^IMPORTACAO_/.test(r[21]||"")?"ACAO_SISTEMICA":""),criadoPor:r[i+1]||"",atualizadoPor:r[i+2]||"",origemEdicao:r[i+3]||"",autorizacaoOpcionais:r[i+4]||""};}
function coberturaFichaUnica_(r){const idx=[1,2,3,4,5,6,7,8,12,13,14,15,16,17,22];return idx.filter(i=>!!r[i]).length+TODOS_CAMPOS_FICHA_.filter(f=>!!r[PESSOAS_HEADERS_.indexOf(f.id)]).length;}

const PESSOAS_HEADERS_=["ID_PESSOA","NOME_COMPLETO","NOME_PREFERIDO","DATA_NASCIMENTO","MUNICIPIO_RA","UF","EMAIL","TELEFONE_WHATSAPP","CANAL_CONTATO","STATUS_CADASTRO","CRIADO_EM","ATUALIZADO_EM","PAIS","CEP","BAIRRO","LOGRADOURO","NUMERO","COMPLEMENTO","ID_UNIDADE_VINCULO_INICIAL","AUTORIZACAO_CADASTRO_EM","VERSAO_AVISO","ORIGEM_AUTORIZACAO","CPF_TOKEN",...CAMPOS_FICHA_UNICA_.map(f=>f.id),...META_FICHA_UNICA_,...CAMPOS_ECLESIAIS_.map(f=>f.id)];
function validarCadastroBasico_(input,data,hoje){
 const clean=(v,max,required=false)=>{if(typeof v!=="string")throw Error("Campo inválido.");const s=v.trim();if((required&&!s)||s.length>max||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(s))throw Error("Confira os campos de cadastro.");return s;};
 if(!input||typeof input!=="object"||input.autorizado!==true)throw Error("É necessário confirmar autorização institucional e da pessoa para o cadastro.");
 const nome=clean(input.nome,180),preferido=clean(input.preferido,100),nascimento=clean(input.nascimento,10),cidade=clean(input.cidade,120),uf=clean(input.uf,2),email=clean(input.email,180),telefone=clean(input.telefone,25),canal=clean(input.canal,20),pais=clean(input.pais,80),cep=clean(input.cep,9),bairro=clean(input.bairro,120),logradouro=clean(input.logradouro,250),numero=clean(input.numero,30),complemento=clean(input.complemento,150),unidade=clean(input.unidade,100);
 if(nascimento){const d=new Date(nascimento+"T00:00:00Z");
 if(!/^\d{4}-\d{2}-\d{2}$/.test(nascimento)||Number.isNaN(d.getTime())||d.toISOString().slice(0,10)!==nascimento||nascimento>hoje||nascimento<"1900-01-01")throw Error("Data de nascimento inválida.");
 let idade=Number(hoje.slice(0,4))-Number(nascimento.slice(0,4));if(hoje.slice(5)<nascimento.slice(5))idade--;
 if(idade<18)throw Error("O fluxo de cadastro de menores ainda não está habilitado.");}
 if((pais&&pais!=="Brasil")||(uf&&!["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"].includes(uf)))throw Error("Informe país ou UF válidos, quando conhecidos.");
 if(cep&&!/^\d{5}-?\d{3}$/.test(cep))throw Error("CEP inválido.");

 if(email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw Error("E-mail inválido.");
 if(telefone&&!/^[0-9 ()-]{8,25}$/.test(telefone))throw Error("Telefone inválido.");
 if(!["","EMAIL","WHATSAPP","OUTRO","SEM_CONTATO"].includes(canal))throw Error("Confira o canal de contato.");
 if(unidade&&unidade!=="CO"&&![...data.dioceses,...data.paroquias,...data.comunidades,...(data.locais||[]).filter(x=>!x.unidade)].some(x=>x.id===unidade))throw Error("Unidade de vínculo não reconhecida.");
 const cpf=input.cpf==null?"":normalizarCpf_(input.cpf);if(![nome,preferido,email,telefone,cpf].some(Boolean))throw Error("Informe ao menos uma referência de identificação ou contato.");return {nome,preferido,nascimento,cidade,uf,email,telefone,canal,pais,cep,bairro,logradouro,numero,complemento,unidade,cpf};
}
function salvarCadastroBasico(input){
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const t=consultarTerritorio();if(!t.ok)return t;
 let v;try{v=validarCadastroBasico_(input,t.data,Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd"));}catch(e){return {ok:false,reason:e.message};}
 let lock;try{
 lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Outra operação em andamento. Tente novamente."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId),s=ss.getSheetByName("PESSOAS");
 const rows=tabela_(ss,"PESSOAS",PESSOAS_HEADERS_);
 const editId=typeof input.id==="string"?input.id.trim():"";const old=editId?rows.find(r=>r[0]===editId):null;if(editId&&(!old||!["CADASTRADO","PARCIAL"].includes(old[9])))return {ok:false,reason:"Registro indisponível."};if(old&&input.revisao!==old[11])return {ok:false,reason:"O cadastro foi atualizado. Abra novamente a ficha."};const token=v.cpf?cpfToken_(v.cpf,rows):(old?.[22]||"");
 if(token&&rows.some(r=>r[0]!==editId&&r[22]===token))return {ok:false,reason:"Possível cadastro existente. Encaminhe para revisão administrativa."};
 if(v.nome&&v.nascimento&&rows.some(r=>r[0]!==editId&&normalizarChave_(r[1])===normalizarChave_(v.nome)&&r[3]===v.nascimento))return {ok:false,reason:"Possível cadastro existente com mesmo nome e nascimento. Solicite revisão antes de cadastrar novamente."};
 if(v.telefone&&rows.some(r=>r[0]!==editId&&r[7]&&r[7].replace(/\D/g,"")===v.telefone.replace(/\D/g,"")&&normalizarChave_(r[1])===normalizarChave_(v.nome)))return {ok:false,reason:"Já existe cadastro com este nome e telefone. Abra a ficha existente."};
 const id=old?editId:Utilities.getUuid(),now=new Date().toISOString(),row=old?s.getRange(2,1,s.getLastRow()-1,1).getDisplayValues().findIndex(r=>r[0]===editId)+2:s.getLastRow()+1;if(old&&row<2)throw Error("Registro não localizado.");
 if(row>1000)return {ok:false,reason:"Limite do cadastro inicial."};
 if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),Math.min(100,1000-s.getMaxRows()));
 const hoje=Utilities.formatDate(new Date(),"America/Sao_Paulo","yyyy-MM-dd");let extras;try{extras=validarCamposFichaUnica_(input,old,hoje);}catch(e){return {ok:false,reason:e.message};}const meta=metadadosFichaUnica_(old||[]),ator=normalizarEmail_(Session.getActiveUser().getEmail());const status="CADASTRADO";const values=[id,v.nome,v.preferido,v.nascimento,v.cidade,v.uf,v.email,v.telefone,v.canal,status,old?.[10]||now,now,v.pais,v.cep,v.bairro,v.logradouro,v.numero,v.complemento,old?.[18]||v.unidade,old?.[19]||now,"CAD-UNICO-1.0",old?.[21]||"DECLARACAO_DO_OPERADOR",token,...CAMPOS_FICHA_UNICA_.map(f=>extras.values[f.id]),meta.origem||(!old?"MANUAL_USUARIO":""),meta.criadoPor||(!old?ator:""),ator,"MANUAL_USUARIO",extras.alteracao?now:meta.autorizacaoOpcionais,...CAMPOS_ECLESIAIS_.map(f=>extras.values[f.id])];
 s.getRange(row,1,1,values.length).setNumberFormat("@").setValues([values]);SpreadsheetApp.flush();
 return {ok:true,id,status};
 }catch(e){return {ok:false,reason:"Não foi possível confirmar a gravação. Confira a planilha antes de tentar novamente."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}
function consultarResumoCadastro(){
 try{
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId),r=tabela_(ss,"PESSOAS",PESSOAS_HEADERS_).filter(x=>["CADASTRADO","PARCIAL"].includes(x[9]));
 const unidades=Object.create(null),links=vinculosPessoas_(ss);for(const x of r){const keys=[...new Set(links.filter(v=>v.pessoa===x[0]&&v.tipo==="Território").map(v=>v.destino))];for(const key of keys.length?keys:["SEM_VINCULO"])unidades[key]=(unidades[key]||0)+1;}
 const equipes=tabela_(ss,"COORDENACOES",COORD_HEADERS_).filter(x=>x[5]!=="ARQUIVADA");
 return {ok:true,total:r.length,parciais:r.filter(x=>x[9]==="PARCIAL").length,unidades:Object.entries(unidades).map(([id,total])=>({id,total})),equipes:equipes.length,rascunhos:equipes.filter(x=>x[5]==="RASCUNHO").length};
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
 return {ok:true,registros:rows.filter(r=>r[10]==="PENDENTE_IDENTIFICACAO").map(r=>({id:r[0],nome:r[12],contato:r[13],funcao:r[3],local:r[4],cidade:r[5],situacao:r[6],data:r[7],ddd:r[14],diocese:r[15],unidade:r[16],referencia:r[17],validacao:r[18]}))};
 }catch(e){return {ok:false,reason:"Pré-cadastros indisponíveis."};}
}
function salvarPreCadastro(input){return {ok:false,reason:"Utilize o cadastro único de pessoas."};}
function salvarPreCadastroLegado_(input){
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};const t=consultarTerritorio();if(!t.ok)return t;
 let lock;try{
 const clean=(v,max)=>{if(typeof v!=="string"||v.length>max||/^[=+\-@]|[\u0000-\u001f\u007f]/.test(v.trim()))throw Error("Campo inválido.");return v.trim();};
 if(input?.autorizado!==true)return {ok:false,reason:"Confirme autorização para registrar o contato."};
 const idEdicao=clean(input.id||"",100),nome=clean(input.nome,180),contato=clean(input.contato,200),funcao=clean(input.funcao||"",120),ref=clean(input.referencia||"",100);if(/[\/,;]|\s*-\s*Vice$/i.test(nome))return {ok:false,reason:"Informe um único nome; registre o cargo em campo separado."};const linked=(t.data.locais||[]).find(x=>x.id===ref);const canon=[...t.data.paroquias,...t.data.comunidades].find(x=>x.id===ref);if(ref&&!linked&&!canon)return {ok:false,reason:"Local não reconhecido."};const local=linked?.nome||canon?.nome||clean(input.local,180),cidade=linked?.cidade||canon?.cidade||clean(input.cidade,120),ddd=linked?.ddd||(t.data.locais||[]).find(x=>x.unidade===canon?.id)?.ddd||"";const phone=padronizarContato_(contato,ddd);
 if(!nome&&!contato)return {ok:false,reason:"Informe pelo menos nome ou contato."};
 lock=LockService.getScriptLock();if(!lock.tryLock(5000))return {ok:false,reason:"Tente novamente."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId),s=ss.getSheetByName("PRE_CADASTROS"),rows=tabela_(ss,"PRE_CADASTROS",PRE_HEADERS_);
 if(rows.some(r=>r[0]!==idEdicao&&r[10]==="PENDENTE_IDENTIFICACAO"&&normalizarChave_(r[12])===normalizarChave_(nome)&&normalizarChave_(r[13])===normalizarChave_(phone.contato)&&normalizarChave_(r[4])===normalizarChave_(local)))return {ok:false,reason:"Referência já cadastrada. Revise os registros existentes."};
 if(contato&&!phone.contato)return {ok:false,reason:"Confira o telefone e o DDD."};
 if(idEdicao){const raw=s.getRange(2,1,Math.max(1,s.getLastRow()-1),PRE_HEADERS_.length).getDisplayValues();const idx=raw.findIndex(r=>r[0]===idEdicao&&r[10]==="PENDENTE_IDENTIFICACAO");if(idx<0)return {ok:false,reason:"Cadastro indisponível."};const old=raw[idx];if(old[11])return {ok:false,reason:"Registro já vinculado a uma ficha de pessoa."};old[3]=funcao;old[4]=local;old[5]=cidade;old[12]=padronizarNome_(nome);old[13]=phone.contato;old[14]=ddd;old[15]=linked?.diocese||canon?.diocese||"";old[16]=linked?.unidade||canon?.id||"";old[17]=linked?.id||"";old[18]=phone.status;s.getRange(idx+2,1,1,PRE_HEADERS_.length).setNumberFormat("@").setValues([old]);SpreadsheetApp.flush();return {ok:true,id:idEdicao};}
 const id=Utilities.getUuid(),row=s.getLastRow()+1;if(row>1000)return {ok:false,reason:"Limite da base."};if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),Math.min(100,1000-s.getMaxRows()));
 s.getRange(row,1,1,PRE_HEADERS_.length).setNumberFormat("@").setValues([[id,nome,contato,funcao,local,cidade,"",new Date().toISOString(),"Registro administrativo autorizado","","PENDENTE_IDENTIFICACAO","",padronizarNome_(nome),phone.contato,ddd,linked?.diocese||canon?.diocese||"",linked?.unidade||canon?.id||"",linked?.id||"",phone.status]]);SpreadsheetApp.flush();return {ok:true,id};
 }catch(e){return {ok:false,reason:"Não foi possível confirmar o pré-cadastro. Confira a planilha."};}finally{if(lock&&lock.hasLock())lock.releaseLock();}
}

function arquivarPreCadastro(id){return {ok:false,reason:"Utilize o cadastro único de pessoas."};}
function arquivarPreCadastroLegado_(id){return operacaoRede_(ss=>{const key=textoOperacional_(id,100);tabela_(ss,"PRE_CADASTROS",PRE_HEADERS_);const s=ss.getSheetByName("PRE_CADASTROS");const raw=s.getRange(2,1,Math.max(1,s.getLastRow()-1),PRE_HEADERS_.length).getDisplayValues();const idx=raw.findIndex(r=>r[0]===key&&r[10]==="PENDENTE_IDENTIFICACAO");if(idx<0)return {ok:false,reason:"Cadastro indisponível."};if(raw[idx][11])return {ok:false,reason:"Registro vinculado a uma ficha; revise o vínculo antes de excluir."};s.getRange(idx+2,11).setValue("ARQUIVADO");SpreadsheetApp.flush();return {ok:true,id:key};});}

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
 return {contato:out.join(" / "),status:out.length?"FORMATADO":"SEM CONTATO"};
}

const VINC_PESSOA_HEADERS_=["ID_VINCULO","ID_PESSOA","ID_UNIDADE","TIPO_VINCULO","PRINCIPAL","INICIO","FIM","STATUS","FONTE_VALIDACAO","CARGO_FUNCAO","SETOR"];
const MEMBROS_PASTORAIS_HEADERS_=["ID vínculo pastoral","ID pessoa","ID pastoral","Tipo de participação","Setor","Início","Fim","Status cadastro","Observações restritas"];
function vinculosPessoas_(ss){
 const territoriais=tabela_(ss,"VINCULOS_TERRITORIAIS",VINC_PESSOA_HEADERS_).filter(r=>r[7]==="ATIVO").map(r=>({id:r[0],pessoa:r[1],destino:r[2],tipo:"Território",cargo:r[9],setor:r[10]}));
 const pastorais=tabela_(ss,"MEMBROS_PASTORAIS",MEMBROS_PASTORAIS_HEADERS_).filter(r=>r[7]==="ATIVO").map(r=>({id:r[0],pessoa:r[1],destino:r[2],tipo:"Pastoral",cargo:r[3],setor:r[4]}));
 const refs=tabela_(ss,"COMPOSICAO_REFERENCIAS",REF_COORD_HEADERS_).filter(r=>r[9]).map(r=>({id:r[0],pessoa:r[9],destino:r[1],tipo:"Coordenação",cargo:r[4],setor:r[3],vigencia:r[6]}));
 return [...territoriais,...pastorais,...refs];
}
function salvarVinculoPessoa(input){
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};const t=consultarTerritorio();if(!t.ok)return t;
 return operacaoRede_(ss=>{
 const clean=(x,max)=>{const s=String(x??"").trim();if(s.length>max||/^[=+\-@]|[\u0000-\u001f]/.test(s))throw Error("Campo inválido.");return s;};
 const pessoa=clean(input.pessoa,100),tipo=clean(input.tipo,25),destino=clean(input.destino,100),cargo=clean(input.cargo,120),setor=clean(input.setor,80);
 if(!tabela_(ss,"PESSOAS",PESSOAS_HEADERS_).some(r=>r[0]===pessoa&&["CADASTRADO","PARCIAL"].includes(r[9])))return {ok:false,reason:"Pessoa indisponível."};
 const defs={"Território":["VINCULOS_TERRITORIAIS",VINC_PESSOA_HEADERS_],"Pastoral":["MEMBROS_PASTORAIS",MEMBROS_PASTORAIS_HEADERS_]};
 if(!Object.prototype.hasOwnProperty.call(defs,tipo))return {ok:false,reason:"Tipo inválido."};
 if(tipo==="Território"&&!unidadeRedeValida_(destino,t.data)&&!(t.data.locais||[]).some(x=>x.id===destino))return {ok:false,reason:"Selecione um território."};
 if(tipo==="Pastoral"&&!(t.data.pastorais||[]).some(x=>x.id===destino))return {ok:false,reason:"Selecione uma pastoral."};
 if(vinculosPessoas_(ss).some(r=>r.pessoa===pessoa&&r.tipo===tipo&&r.destino===destino&&r.cargo===cargo&&r.setor===setor))return {ok:false,reason:"Vínculo já cadastrado."};
 const [aba,headers]=defs[tipo];tabela_(ss,aba,headers);const sh=ss.getSheetByName(aba),id=Utilities.getUuid(),n=sh.getLastRow()+1;if(n>1000)return {ok:false,reason:"Limite atual da base."};if(n>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),100);
 const r=tipo==="Território"?[id,pessoa,destino,"VINCULO_PASTORAL","","","","ATIVO","Cadastro administrativo",cargo,setor]:[id,pessoa,destino,cargo,setor,"","","ATIVO",""];
 sh.getRange(n,1,1,r.length).setNumberFormat("@").setValues([r]);SpreadsheetApp.flush();return {ok:true,id};
 });
}
function encerrarVinculoPessoa(pessoa,tipo,id){
 return operacaoRede_(ss=>{const defs={"Território":["VINCULOS_TERRITORIAIS",VINC_PESSOA_HEADERS_],"Pastoral":["MEMBROS_PASTORAIS",MEMBROS_PASTORAIS_HEADERS_]};if(!Object.prototype.hasOwnProperty.call(defs,tipo))return {ok:false,reason:"Gerencie este vínculo em Coordenações."};const [aba,h]=defs[tipo];tabela_(ss,aba,h);const sh=ss.getSheetByName(aba),rows=sh.getRange(1,1,sh.getLastRow(),h.length).getDisplayValues(),i=rows.findIndex((r,j)=>j>0&&r[0]===id&&r[1]===pessoa&&r[7]==="ATIVO");if(i<1)return {ok:false,reason:"Vínculo indisponível."};sh.getRange(i+1,8).setValue("ENCERRADO");SpreadsheetApp.flush();return {ok:true};});
}
function vincularReferenciaCoordenacao(ref,pessoa){
 return operacaoRede_(ss=>{if(!tabela_(ss,"PESSOAS",PESSOAS_HEADERS_).some(r=>r[0]===pessoa&&["CADASTRADO","PARCIAL"].includes(r[9])))return {ok:false,reason:"Pessoa indisponível."};tabela_(ss,"COMPOSICAO_REFERENCIAS",REF_COORD_HEADERS_);const sh=ss.getSheetByName("COMPOSICAO_REFERENCIAS"),rows=sh.getRange(1,1,sh.getLastRow(),10).getDisplayValues(),i=rows.findIndex((r,j)=>j>0&&r[0]===ref);if(i<1||rows[i][9])return {ok:false,reason:"Referência já vinculada ou indisponível."};if(/institui[cç][aã]o/i.test(rows[i][4]))return {ok:false,reason:"Referência institucional."};sh.getRange(i+1,10).setValue(pessoa);SpreadsheetApp.flush();return {ok:true};});
}

function consultarPessoasPorLocal(){
 try{
 const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};
 const ss=SpreadsheetApp.openById(cfg.operacionalId);
 const vinculos=vinculosPessoas_(ss);return {ok:true,pessoas:tabela_(ss,"PESSOAS",PESSOAS_HEADERS_).filter(r=>["CADASTRADO","PARCIAL"].includes(r[9])).map(r=>({id:r[0],nome:r[1]||r[2]||"Identificação parcial",preferido:r[2],telefone:r[8]==="SEM_CONTATO"?"":r[7],email:r[6],unidade:r[18],status:r[9],preenchidos:coberturaFichaUnica_(r),origem:metadadosFichaUnica_(r).origem,vinculos:vinculos.filter(v=>v.pessoa===r[0])}))};
 }catch(e){return {ok:false,reason:"Consulta indisponível."};}
}

function consultarFichaPessoa(id){
 try{const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};const ss=SpreadsheetApp.openById(cfg.operacionalId),r=tabela_(ss,"PESSOAS",PESSOAS_HEADERS_).find(r=>r[0]===id&&["CADASTRADO","PARCIAL"].includes(r[9]));if(!r)return {ok:false,reason:"Ficha indisponível."};
 return {ok:true,pessoa:{id:r[0],nome:r[1],preferido:r[2],nascimento:r[3],cidade:r[4],uf:r[5],email:r[6],telefone:r[7],canal:r[8],revisao:r[11],pais:r[12],cep:r[13],bairro:r[14],logradouro:r[15],numero:r[16],complemento:r[17],unidade:r[18],vinculos:vinculosPessoas_(ss).filter(v=>v.pessoa===id),cpfRegistrado:!!r[22],camposFicha:Object.fromEntries(TODOS_CAMPOS_FICHA_.map(f=>[f.id,r[PESSOAS_HEADERS_.indexOf(f.id)]||""])),metadados:metadadosFichaUnica_(r),preenchidos:coberturaFichaUnica_(r)}};
 }catch(e){return {ok:false,reason:"Ficha indisponível."};}
}
const CONTATOS_INST_HEADERS_=["ID_CONTATO","ID_UNIDADE","TIPO_CONTATO","SETOR_FUNCAO","NOME_REFERENCIA","TELEFONE","WHATSAPP","EMAIL","FONTE","VALIDACAO","ATUALIZADO_EM"];
function consultarContatosInstitucionais(){
 try{const cfg=gestaoConfigAutorizada_();if(!cfg)return {ok:false,reason:"Acesso negado."};return {ok:true,contatos:tabela_(SpreadsheetApp.openById(cfg.operacionalId),"CONTATOS_INSTITUCIONAIS",CONTATOS_INST_HEADERS_).map(r=>({id:r[0],unidade:r[1],tipo:r[2],setor:r[3],nome:r[4],telefone:r[5],whatsapp:r[6],email:r[7],fonte:r[8],validacao:r[9]}))};}catch(e){return {ok:false,reason:"Contatos indisponíveis."};}
}
function salvarContatoInstitucional(input){
 if(!gestaoConfigAutorizada_())return {ok:false,reason:"Acesso negado."};const t=consultarTerritorio();if(!t.ok)return t;
 return operacaoRede_(ss=>{
 const v=(k,max)=>textoOperacional_(input?.[k]||"",max,false);
 const unidade=v("unidade",100),tipo=v("tipo",40),setor=v("setor",120),nome=v("nome",180),telefone=v("telefone",50),whatsapp=v("whatsapp",50),email=v("email",180),fonte=v("fonte",500);
 if(!unidadeRedeValida_(unidade,t.data)&&!(t.data.locais||[]).some(x=>x.id===unidade))return {ok:false,reason:"Selecione o local de referência."};
 if(!["CURIA_DIOCESANA","SECRETARIA_PAROQUIAL","PASTORAL_DO_SURDO","OUTRA_PASTORAL","COMUNICACAO","OUTRO"].includes(tipo)||![nome,telefone,whatsapp,email].some(Boolean)||input?.autorizado!==true)return {ok:false,reason:"Informe tipo, pelo menos uma referência de contato e confirme autorização."};
 if(email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))return {ok:false,reason:"E-mail inválido."};
 const rows=tabela_(ss,"CONTATOS_INSTITUCIONAIS",CONTATOS_INST_HEADERS_);
 if(rows.some(r=>r[1]===unidade&&r[2]===tipo&&[setor,nome,telefone,whatsapp,email].every((v,i)=>normalizarChave_(v)===normalizarChave_(r[i+3]))))return {ok:false,reason:"Referência já registrada."};
 const id=Utilities.getUuid();appendOperacional_(ss,"CONTATOS_INSTITUCIONAIS",CONTATOS_INST_HEADERS_,[id,unidade,tipo,setor,nome,telefone,whatsapp,email,fonte,"INFORMADO — validar vigência",new Date().toISOString()]);return {ok:true,id};
 });
}
