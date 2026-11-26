import test from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import {readFile} from "node:fs/promises";
const context=vm.createContext({});vm.runInContext(await readFile(new URL("../apps-script/Code.gs",import.meta.url),"utf8"),context);vm.runInContext(await readFile(new URL("../apps-script/Territorial.gs",import.meta.url),"utf8"),context);
const run=(d,p,c,m)=>JSON.parse(JSON.stringify(context.mapTerritorio_(d,p,c,m)));
const d=[["D01","Arquidiocese Teste","Província Teste","GO"],["D02","Diocese Teste","Província Teste","GO"]];
const p=[["P01","D01","","Paróquia Teste","Paróquia","Cidade Teste","","Rua Teste","00000-000","","","","","","NOME NÃO EXPORTADO"]];
const c=[["C01","P01","D01","Comunidade Teste","Cidade Teste","","Rua Teste"]];
test("mapeia IDs, endereços e jurisdições, não exporta nomes do clero",()=>{const x=run(d,p,c,[]);assert.equal(x.paroquias[0].logradouro,"Rua Teste");assert.equal(x.comunidades[0].paroquia,"P01");assert.ok(!JSON.stringify(x).includes("NOME NÃO EXPORTADO"));});
test("recusa duplicatas e vínculo incompatível",()=>{assert.throws(()=>run([...d,d[0]],p,c,[]));assert.throws(()=>run(d,p,[["C01","P01","D02","Teste"]],[]));});
test("celebração sem jurisdição fica pendente",()=>{const x=run(d,p,c,[["M01","", "Cidade","Local"]]);assert.equal(x.missas[0].vinculoPendente,true);assert.equal(x.missas[0].diocese,"");});
test("diretório real nega antes de ler planilha quando visitante não autorizado",()=>{let reads=0;context.territorialConfig_=()=>({owner:"owner@example.test",sourceId:"test"});context.Session={getActiveUser:()=>({getEmail:()=>"visitor@example.test"})};context.SpreadsheetApp={openById:()=>{reads++;throw Error("não deveria ler")}};assert.equal(context.consultarTerritorio("owner@example.test").ok,false);assert.equal(reads,0);context.Session={getActiveUser:()=>({getEmail:()=>""})};assert.equal(context.consultarTerritorio().ok,false);assert.equal(reads,0);});
test("HTML preserva marca e consulta no servidor",async()=>{const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");assert.match(h,/GEB Tecnologia/);assert.match(h,/__REGIONAL_LOGO_ATTR__/);assert.match(h,/consultarTerritorio\(\)/);const s=h.match(/<script>([\s\S]*?)<\/script>/)[1];assert.doesNotThrow(()=>new vm.Script(s));});

const navHtml=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");
const nav=vm.createContext({});vm.runInContext(navHtml.split("// BEGIN_NAV_LOGIC")[1].split("// END_NAV_LOGIC")[0],nav);
const fixture={dioceses:[{id:"D01",nome:"Arquidiocese teste",provincia:"Goiânia"},{id:"D02",nome:"Diocese teste",provincia:"Brasília"}],paroquias:[{id:"P1",diocese:"D01",nome:"Paróquia A",cidade:"Cidade",logradouro:"Rua teste, 10",fonteFicha:"https://example.test/ficha-a/"},{id:"P2",diocese:"D02",nome:"Paróquia B"}],comunidades:[{id:"C1",paroquia:"P1",diocese:"D01",nome:"Comunidade A"}],missas:[{id:"M1",diocese:"D01",fonte:"https://example.test/ficha-a",nome:"Local A"},{id:"M2",diocese:"D01",nome:"Outra igreja",fonte:"https://example.test/outra"},{id:"M3",diocese:"D02",fonte:"https://example.test/ficha-a"}]};
const json=x=>JSON.parse(JSON.stringify(x));
test("selecionar paróquia define jurisdição, província e categoria automaticamente",()=>{assert.deepEqual(json(nav.resolverSelecao(fixture,{provincia:"",jurisdicao:"",paroquia:"P1",comunidade:"",categoria:"dioceses"},"paroquia")),{provincia:"Goiânia",jurisdicao:"D01",paroquia:"P1",comunidade:"",categoria:"paroquias"});});
test("selecionar comunidade define todos os ancestrais",()=>{const x=nav.resolverSelecao(fixture,{comunidade:"C1",categoria:"dioceses"},"comunidade");assert.equal(x.paroquia,"P1");assert.equal(x.jurisdicao,"D01");assert.equal(x.categoria,"comunidades");});
test("mudança de província limpa descendentes",()=>{const x=nav.resolverSelecao(fixture,{provincia:"Brasília",jurisdicao:"D01",paroquia:"P1",comunidade:"C1",categoria:"comunidades"},"provincia");assert.equal(x.jurisdicao,"");assert.equal(x.paroquia,"");assert.equal(x.comunidade,"");assert.equal(x.categoria,"dioceses");});
test("referências relacionadas por ficha específica e mesma diocese, sem extrapolação",()=>{assert.deepEqual(json(nav.referenciasRelacionadas(fixture,fixture.paroquias[0],"paroquia")).map(x=>x.id),["M1"]);assert.equal(nav.referenciasRelacionadas(fixture,fixture.comunidades[0],"comunidade").length,0);});
const request={tipo:"PAROQUIA",superior:"D01",nome:"Nova unidade fictícia",cidade:"Cidade",uf:"GO",bairro:"",endereco:"",cep:"",fonte:"https://example.test/fonte"};
test("solicitação válida fica normalizada; rejeita duplicidade, superior inválido e fórmula",()=>{assert.equal(context.validarSolicitacao_(request,fixture).superior,"D01");assert.throws(()=>context.validarSolicitacao_({...request,nome:"Paróquia A"},fixture));assert.throws(()=>context.validarSolicitacao_({...request,superior:"X"},fixture));assert.throws(()=>context.validarSolicitacao_({...request,tipo:"CAPELA",superior:"D01"},fixture));assert.throws(()=>context.validarSolicitacao_({...request,nome:"=HYPERLINK(1)"},fixture));assert.throws(()=>context.validarSolicitacao_({...request,cep:"123"},fixture));});
test("solicitação de unidade não autorizada não grava",()=>{context.Session={getActiveUser:()=>({getEmail:()=>""})};context.LockService={getScriptLock:()=>{throw Error("não deve chegar ao lock")}};assert.equal(context.solicitarUnidade(request).ok,false);});

test("programação em Libras preserva tipo, vínculo explícito e situação",()=>{const m=["M4","D01","Cidade","Local","Rua","Terça-feira","14h","Ativo","23/11/2026","Fonte","Validar","Novena","P01","CONFIRMADA"];const x=run(d,p,c,[m]).missas[0];assert.equal(x.tipo,"Novena");assert.equal(x.unidade,"P01");assert.equal(x.statusProgramacao,"CONFIRMADA");assert.equal(run(d,p,c,[[...m.slice(0,12),"INEXISTENTE","CONFIRMADA"]]).missas[0].unidade,"");});
test("filtro de horários recorrentes e confirmados exclui divergências e ausência de horário",()=>{const x=[{id:"a",statusProgramacao:"CONFIRMADA",dia:"Domingo",horario:"10h"},{id:"b",statusProgramacao:"RECORRENTE INFORMADA",dia:"Sábado",horario:"9h"},{id:"c",statusProgramacao:"DIVERGENTE",dia:"Domingo",horario:"10h"},{id:"d",statusProgramacao:"A CONFIRMAR"},{id:"e",statusProgramacao:"CONFIRMADA",dia:"",horario:""}];assert.deepEqual(json(nav.filtrarAcessibilidade(x,"fixo")).map(r=>r.id),["a","b"]);assert.deepEqual(json(nav.filtrarAcessibilidade(x,"confirmada")).map(r=>r.id),["a"]);assert.deepEqual(json(nav.filtrarAcessibilidade(x,"pendente")).map(r=>r.id),["b","c","d"]);});
test("pastoral ativa exige evidência explícita independente do intérprete",()=>{const x=[{id:"a",pastoral:"Ativa — atividades informadas"},{id:"b",pastoral:"Não verificado",statusProgramacao:"CONFIRMADA"},{id:"c",pastoral:"Inativa"},{id:"d",pastoral:"A confirmar"}];assert.deepEqual(json(nav.filtrarAcessibilidade(x,"pastoral")).map(r=>r.id),["a"]);});
test("vínculo explícito prevalece sobre correspondência por fonte",()=>{const x={...fixture,missas:[{id:"M4",diocese:"D01",unidade:"P1"},{id:"M5",diocese:"D01",unidade:"OTHER",fonte:"https://example.test/ficha-a"}]};assert.deepEqual(json(nav.referenciasRelacionadas(x,fixture.paroquias[0],"paroquia")).map(r=>r.id),["M4"]);});
test("sessão real não lê registros fictícios e nega identidade ausente",()=>{context.territorialConfig_=()=>({owner:"owner@example.test",sourceId:"test"});context.Session={getActiveUser:()=>({getEmail:()=>"owner@example.test"})};context.SpreadsheetApp={openById:()=>{throw Error("sessão não deve ler planilhas")}};assert.equal(context.verificarSessao().ok,true);assert.equal(context.verificarSessao().escopo,"Regional Centro-Oeste");context.Session={getActiveUser:()=>({getEmail:()=>""})};assert.equal(context.verificarSessao().ok,false);assert.doesNotMatch(navHtml,/consultarPainel\(|Registros fictícios/);assert.match(navHtml,/verificarSessao\(\)/);});

test("pastorais ativas inclui locais de referência sem duplicar unidade vinculada",()=>{
 const data={...fixture,paroquias:[{...fixture.paroquias[0],pastoral:"Ativa — confirmação"}],comunidades:[],missas:[{id:"M1",diocese:"D01",unidade:"P1",pastoral:"Ativa — fonte",nome:"Matriz",cidade:"Cidade"},{id:"M2",diocese:"D01",unidade:"P1",pastoral:"Ativa — fonte",nome:"Matriz",cidade:"Cidade"},{id:"M3",diocese:"D02",pastoral:"Ativa — fonte",nome:"Local DF",cidade:"Cidade DF"},{id:"M4",diocese:"D02",pastoral:"Inativa na fonte"},{id:"M5",diocese:"D02",statusProgramacao:"CONFIRMADA"}]};
 assert.deepEqual(json(nav.locaisPastoraisAtivas(data,new Set(["D01","D02"]),"","","")).map(x=>x.id),["P1","M3"]);
 assert.deepEqual(json(nav.locaisPastoraisAtivas(data,new Set(["D02"]),"D02","","")).map(x=>x.id),["M3"]);
 assert.equal(nav.locaisPastoraisAtivas(data,new Set(["D02"]),"D02","P2","").length,0);
});
test("mapeamento conserva atuação pastoral e fonte independente da programação",()=>{
 const m=["M1","D01","Cidade","Local","Rua","Domingo","9h","Ativo","2026","Fonte","Validar","Missa","","RECORRENTE INFORMADA","Ativa — fonte","Fonte pastoral"];
 const x=run(d,p,c,[m]).missas[0];assert.equal(x.pastoral,"Ativa — fonte");assert.equal(x.fontePastoral,"Fonte pastoral");assert.equal(x.statusProgramacao,"RECORRENTE INFORMADA");
});

vm.runInContext(await readFile(new URL("../apps-script/Gestao.gs",import.meta.url),"utf8"),context);
test("níveis regimentais, representação paroquial e limites do escopo regional",()=>{
 const data={dioceses:[{id:"D01"},{id:"D13"}],paroquias:[{id:"P01"}]};
 const valid={nivel:"ARQUIDIOCESANO_DIOCESANO",unidade:"D01",nome:"Equipe teste",fonte:"Referência fictícia"};
 assert.equal(context.validarEquipeRascunho_(valid,data,[]).tipo,"COORDENACAO");
 assert.equal(context.validarEquipeRascunho_({...valid,nivel:"PAROQUIAL_REPRESENTANTE",unidade:"P01"},data,[]).tipo,"REPRESENTACAO");
 assert.throws(()=>context.validarEquipeRascunho_({...valid,nivel:"NACIONAL"},data,[]));
 assert.throws(()=>context.validarEquipeRascunho_({...valid,unidade:"D13"},data,[]));
 assert.throws(()=>context.validarEquipeRascunho_(valid,data,[["X","D01","","ARQUIDIOCESANO_DIOCESANO","","RASCUNHO"]]));
 assert.throws(()=>context.validarEquipeRascunho_({...valid,nome:"=FORMULA()"},data,[]));
});
test("gestão nega antes de leitura ou escrita sem identidade confirmada",()=>{
 context.Session={getActiveUser:()=>({getEmail:()=>""})};
 context.gestaoConfig_=()=>({operacionalId:"test",questionarioId:"test"});
 context.SpreadsheetApp={openById:()=>{throw Error("não deve ler")}};
 assert.equal(context.consultarGestao().ok,false);
 assert.equal(context.criarCoordenacaoRascunho({}).ok,false);
});
test("prévia sem envio de respostas pessoais, portal separado e sem acesso a planilhas",async()=>{
 const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");
 assert.match(h,/v0\.14/);assert.match(h,/sem campos de resposta nesta etapa/);
 assert.match(h,/Nenhum voto ou candidatura será registrado/);
 assert.doesNotMatch(h,/salvarPessoa\(|registrarVoto\(/);
 const portal=await readFile(new URL("../publico/index.html",import.meta.url),"utf8");
 assert.match(portal,/Entrar no sistema/);assert.match(portal,/pastoraldosurdocentro\.blogspot\.com/);
 assert.doesNotMatch(portal,/docs\.google\.com\/spreadsheets|google\.script\.run/);
});

test("cadastro básico inclui data, endereço e vínculo opcional",async()=>{const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");for(const id of ["p-nascimento","p-cep","p-uf","p-cidade","p-bairro","p-logradouro","p-numero","p-complemento","p-regional","p-diocese","p-paroquia","p-comunidade"])assert.ok(h.includes('id="'+id+'"'));assert.match(h,/salvarCadastroBasico\(input\)/);assert.match(h,/sem campos de resposta nesta etapa/);});

test("navegação por módulos inicia na home, mantém um painel visível e carrega cada consulta",async()=>{const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");const src=h.split("// BEGIN_PAGE_NAV")[1].split("// END_PAGE_NAV")[0];const ctx=vm.createContext({});vm.runInContext(src,ctx);assert.equal(ctx.paginaValida("cadastro-pessoas"),"cadastro-pessoas");assert.equal(ctx.paginaValida("inexistente"),"home");const ids=["home","territorio","inclusao","gestao","cadastro-pessoas","eleicoes","acessos-gestao","documentos","pagina-publica","acesso"];for(const id of ids){assert.ok(h.includes('data-page="'+id+'"'));assert.ok(h.includes('<section id="'+id+'"'+(id==="home"?">":" hidden>")));}assert.match(h,/T\(p\)\.hidden=p!==id/);assert.match(h,/\(id==="gestao"\|\|id==="cadastro-pessoas"\)&&!gestao\)carregarGestao\(\)/);assert.match(h,/abrirPagina\("home"\);/);assert.match(h,/questionario-status/);assert.match(h,/function atualizarHome\(/);assert.doesNotMatch(h,/<nav aria-label="Menu"><a/);});

const pessoa={nome:"Pessoa Teste",preferido:"",nascimento:"1990-03-04",cidade:"Cidade Teste",uf:"GO",email:"",telefone:"",canal:"SEM_CONTATO",pais:"Brasil",cep:"",bairro:"",logradouro:"",numero:"",complemento:"",unidade:"",autorizado:true};
test("cadastro básico aceita vínculo ausente ou unidade válida e recusa data inválida ou menor",()=>{
 const today="2026-11-24";
 assert.equal(context.validarCadastroBasico_(pessoa,fixture,today).unidade,"");
 for(const unidade of ["CO","D01","P1","C1"])assert.equal(context.validarCadastroBasico_({...pessoa,unidade},fixture,today).unidade,unidade);
 for(const x of [{unidade:"X"},{nascimento:"2000-02-30"},{nascimento:"2012-01-01"},{autorizado:false},{nome:"=FORMULA()"},{cep:"12345"}])assert.throws(()=>context.validarCadastroBasico_({...pessoa,...x},fixture,today));
});
test("cadastro básico e resumo não leem nem gravam sem autorização",()=>{
 context.Session={getActiveUser:()=>({getEmail:()=>""})};context.SpreadsheetApp={openById:()=>{throw Error("não deve abrir")}};
 assert.equal(context.salvarCadastroBasico(pessoa).ok,false);assert.equal(context.consultarResumoCadastro().ok,false);
});
test("home contém visão territorial e cadastro oficial sem opção nacional no menu eleitoral",async()=>{
 const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");
 assert.match(h,/home-territorial-table/);assert.match(h,/Pessoas vinculadas/);assert.match(h,/consultarResumoCadastro\(\)/);
 const sec=h.slice(h.indexOf('<section id="eleicoes"'),h.indexOf('<section id="acessos-gestao"'));
 assert.doesNotMatch(sec,/<option>Nacional<\/option>/);
});

test("gravação simulada preserva endereço, vínculo, autorização e evita duplicação",()=>{
 const headers=vm.runInContext("PESSOAS_HEADERS_",context);
 const rows=[Array.from(headers)];let writes=0,locked=false;
 const tab={getLastRow:()=>rows.length,getMaxRows:()=>200,getRange:(row,col,n,width)=>({getDisplayValues:()=>rows.slice(row-1,row-1+n).map(r=>r.slice(col-1,col-1+width)),setNumberFormat:()=>({setValues:v=>{writes++;rows.push(v[0]);}})})};
 context.Session={getActiveUser:()=>({getEmail:()=>"owner@example.test"})};
 context.territorialConfig_=()=>({owner:"owner@example.test",sourceId:"test"});
 context.gestaoConfig_=()=>({operacionalId:"test",questionarioId:"test"});
 const original=context.consultarTerritorio;
 context.consultarTerritorio=()=>({ok:true,data:fixture});
 context.SpreadsheetApp={openById:()=>({getSheetByName:()=>tab}),flush:()=>{}};
 context.Utilities={getUuid:()=>"TEST-001",formatDate:()=>"2026-11-24"};
 context.LockService={getScriptLock:()=>({tryLock:()=>{locked=true;return true;},hasLock:()=>locked,releaseLock:()=>{locked=false;}})};
 try{
 const p={...pessoa,unidade:"C1",cep:"74000-000",bairro:"Teste",logradouro:"Rua Teste",numero:"s/n"};
 assert.equal(context.salvarCadastroBasico(p).ok,true);assert.equal(writes,1);
 assert.equal(rows[1][3],"1990-03-04");assert.equal(rows[1][13],"74000-000");assert.equal(rows[1][18],"C1");assert.equal(rows[1][20],"CAD-UNICO-1.0");
 assert.equal(context.salvarCadastroBasico(p).ok,false);assert.equal(writes,1);
 }finally{context.consultarTerritorio=original;}
});

test("CPF normalizado e dígitos verificadores, sem exigir CPF para contato parcial",()=>{assert.equal(context.normalizarCpf_("529.982.247-25"),"52998224725");assert.equal(context.normalizarCpf_(""),"");assert.throws(()=>context.normalizarCpf_("11111111111"));assert.throws(()=>context.normalizarCpf_("52998224726"));});
test("pré-cadastro não autorizado não lê nem grava",()=>{context.Session={getActiveUser:()=>({getEmail:()=>""})};context.SpreadsheetApp={openById:()=>{throw Error("não deve abrir")}};assert.equal(context.consultarPreCadastros().ok,false);assert.equal(context.salvarPreCadastro({}).ok,false);});
test("ficha apresenta horários, link de localização institucional e detalhes recolhíveis",async()=>{const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");assert.match(h,/google\.com\/maps\/search\/\?api=1/);assert.match(h,/encodeURIComponent/);assert.doesNotMatch(h,/Identidade cadastral a validar|Composição conforme fontes institucionais/);assert.match(h,/consultarPreCadastros\(\)/);assert.match(h,/salvarPreCadastro\(/);});

test("grupos e alertas negam acesso antes de leitura",()=>{
 context.Session={getActiveUser:()=>({getEmail:()=>""})};context.SpreadsheetApp={openById:()=>{throw Error("não deve abrir")}};
 for(const method of ["consultarTrabalho","criarGrupoTrabalho","adicionarMembroGrupo","criarAlertaRede","publicarAlertaRede"])assert.equal(context[method]({}).ok,false);
});
test("validação territorial e campos de operação",()=>{
 assert.equal(context.unidadeRedeValida_("CO",fixture),true);assert.equal(context.unidadeRedeValida_("C1",fixture),true);assert.equal(context.unidadeRedeValida_("NACIONAL",fixture),false);assert.throws(()=>context.textoOperacional_("=FORMULA()",100));assert.equal(context.textoOperacional_("  Grupo teste  ",100),"Grupo teste");
});
test("menus de grupos e alertas e composição documental",async()=>{
 const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");
 for(const id of ["grupos","alertas"]) {assert.ok(h.includes('data-page="'+id+'"'));assert.ok(h.includes('<section id="'+id+'" hidden>'));}
 assert.match(h,/equipe-composicao/);assert.match(h,/criarGrupoTrabalho/);assert.match(h,/adicionarMembroGrupo/);assert.match(h,/criarAlertaRede/);assert.match(h,/publicarAlertaRede/);assert.match(h,/Nenhum disparo externo ocorre nesta versão/);
});

test("padronização de nomes e telefone conserva DDD informado e aplica referência somente quando falta",()=>{assert.equal(context.padronizarNome_("MARIA DE SOUZA LIBRAS"),"Maria de Souza");assert.equal(context.padronizarContato_("99999-0000","61").contato,"(61) 99999-0000");assert.equal(context.padronizarContato_("(62) 99999-0000","61").contato,"(62) 99999-0000");assert.equal(context.padronizarContato_("99999-0000","").contato,"");});
test("consulta de pessoas restrita e vínculos locais sem matriz",()=>{context.Session={getActiveUser:()=>({getEmail:()=>""})};context.SpreadsheetApp={openById:()=>{throw Error("não deve abrir")}};assert.equal(context.consultarPessoasPorLocal().ok,false);assert.equal(context.validarCadastroBasico_({...pessoa,unidade:"LOC-TEST"},{...fixture,locais:[{id:"LOC-TEST",unidade:""}]},"2026-11-24").unidade,"LOC-TEST");});
test("interface inclui filtro territorial, DDD e seleção de local",async()=>{const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");assert.match(h,/id="p-filtro-local"/);assert.match(h,/id="pre-ref"/);assert.match(h,/consultarPessoasPorLocal/);assert.match(h,/DDD sugerido/);});

test("ficha padrão aceita informação parcial sem data, cidade, UF, CPF ou endereço completo",()=>{
 const p={...pessoa,nascimento:"",cidade:"",uf:"",pais:"",cep:"74000-000"};
 assert.equal(context.validarCadastroBasico_(p,fixture,"2026-11-25").cep,"74000-000");
 assert.equal(context.validarCadastroBasico_({...p,nome:"",telefone:"999999999"},fixture,"2026-11-25").telefone,"999999999");
 assert.throws(()=>context.validarCadastroBasico_({...p,nome:""},fixture,"2026-11-25"));
});
test("ficha e contatos institucionais negam acesso sem identidade confirmada",()=>{
 context.Session={getActiveUser:()=>({getEmail:()=>""})};context.SpreadsheetApp={openById:()=>{throw Error("não deve ler")}};
 assert.equal(context.consultarFichaPessoa("X").ok,false);assert.equal(context.consultarContatosInstitucionais().ok,false);assert.equal(context.salvarContatoInstitucional({}).ok,false);
});
test("campos pessoais não obrigatórios, edição e diretório de contatos",async()=>{
 const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");
 for(const k of ["p-nome","p-nascimento","p-uf","p-cidade","p-pais"])assert.doesNotMatch(h,new RegExp('id="'+k+'"[^>]*\\brequired\\b'));
 assert.match(h,/consultarFichaPessoa\(id\)/);assert.match(h,/id="p-revisao"/);assert.match(h,/data-page="contatos"/);assert.match(h,/consultarContatosInstitucionais\(\)/);
});
test("gravação de ficha parcial e complementação mantêm o mesmo ID, com controle de revisão",()=>{
 const headers=Array.from(vm.runInContext("PESSOAS_HEADERS_",context)),rows=[headers];let locked=false;
 const tab={getLastRow:()=>rows.length,getMaxRows:()=>200,getRange:(row,col,n,width)=>({getDisplayValues:()=>rows.slice(row-1,row-1+n).map(r=>r.slice(col-1,col-1+width)),setNumberFormat:()=>({setValues:v=>{rows[row-1]=v[0];}})})};
 context.Session={getActiveUser:()=>({getEmail:()=>"owner@example.test"})};context.territorialConfig_=()=>({owner:"owner@example.test",sourceId:"test"});context.gestaoConfig_=()=>({operacionalId:"test",questionarioId:"test"});const original=context.consultarTerritorio;
 context.consultarTerritorio=()=>({ok:true,data:fixture});const linkTables={VINCULOS_TERRITORIAIS:"VINC_PESSOA_HEADERS_",MEMBROS_PASTORAIS:"MEMBROS_PASTORAIS_HEADERS_",COMPOSICAO_REFERENCIAS:"REF_COORD_HEADERS_",COMPOSICAO:"COMPOSICAO_GESTAO_HEADERS_"};context.SpreadsheetApp={openById:()=>({getSheetByName:name=>name==="PESSOAS"?tab:{getLastRow:()=>1,getRange:()=>({getDisplayValues:()=>[Array.from(vm.runInContext(linkTables[name],context))]})}}),flush:()=>{}};context.Utilities={getUuid:()=>"TEST-PARTIAL",formatDate:()=>"2026-11-25"};context.LockService={getScriptLock:()=>({tryLock:()=>{locked=true;return true;},hasLock:()=>locked,releaseLock:()=>{locked=false;}})};
 try{const p={...pessoa,nascimento:"",cidade:"",uf:"",pais:""};const first=context.salvarCadastroBasico(p);assert.equal(first.status,"CADASTRADO");assert.equal(rows.length,2);const detail=context.consultarFichaPessoa(first.id);assert.equal(detail.ok,true);assert.equal(detail.pessoa.id,first.id);assert.equal(context.salvarCadastroBasico({...p,id:first.id,revisao:"invalid"}).ok,false);const updated=context.salvarCadastroBasico({...p,nascimento:"1990-03-04",id:first.id,revisao:detail.pessoa.revisao});assert.equal(updated.status,"CADASTRADO");assert.equal(updated.id,first.id);assert.equal(rows.length,2);
 }finally{context.consultarTerritorio=original;}
});

test("programação nominal não altera tipo, dia ou horário",()=>{const row=["M30","D01","Cidade","Local","","Terça-feira","14h","","","","","Novena","P01","CONFIRMADA","","","Novena Perpétua de Nossa Senhora do Perpétuo Socorro"];const x=run(d,p,c,[row]).missas[0];assert.equal(x.programacao,row[16]);assert.equal(x.tipo,"Novena");assert.equal(x.horario,"14h");assert.equal(x.dia,"Terça-feira");assert.equal(run(d,p,c,[row.slice(0,16)]).missas[0].programacao,"Novena");});

test("cargo e setor são exibidos separadamente sem inventar titularidade",()=>{const ctx=vm.createContext({});vm.runInContext(navHtml.split("// BEGIN_CONTACT_LOGIC")[1].split("// END_CONTACT_LOGIC")[0],ctx);assert.equal(ctx.separarCargoPastoral("Coordenação/vice ouvinte").cargo,"Coordenação ou vice-coordenação");assert.equal(ctx.separarCargoPastoral("Coordenação/vice ouvinte").setor,"Ouvinte");assert.equal(ctx.separarCargoPastoral("Vice-coordenação surdo").cargo,"Vice-coordenação");assert.equal(ctx.separarCargoPastoral("Vice-coordenação surdo").setor,"Surdo");});
test("atalhos de contato usam número individual e saudação nominal",()=>{const ctx=vm.createContext({});vm.runInContext(navHtml.split("// BEGIN_CONTACT_LOGIC")[1].split("// END_CONTACT_LOGIC")[0],ctx);const a=ctx.linksContatoPastoral("(61) 99999-0000","Pessoa Teste");assert.equal(a.tel,"tel:+5561999990000");assert.equal(a.whatsapp,"https://wa.me/5561999990000?text="+encodeURIComponent("Olá, Pessoa Teste! Paz e bem!"));assert.equal(ctx.linksContatoPastoral("99999-0000","Pessoa Teste"),null);assert.equal(ctx.linksContatoPastoral("(61) 99999-0000 / (61) 98888-0000","Pessoa Teste"),null);}); 

test("ficha paroquial conserva campos institucionais e não exige dados ausentes",()=>{const row=Array(33).fill("");row[0]="P01";row[1]="D01";row[3]="Paróquia Teste";row[13]="Segunda a sexta";row[26]="00.000.000/E08G-12";row[27]="Entidade Teste";const result=run(d,[row],[],[]).paroquias[0];assert.equal(result.cnpj,row[26]);assert.equal(result.razaoSocial,row[27]);assert.equal(result.expediente,row[13]);assert.equal(result.whatsapp,"");});
test("manutenção territorial exige identidade autorizada",()=>{context.Session={getActiveUser:()=>({getEmail:()=>""})};context.SpreadsheetApp={openById:()=>{throw Error("não deve abrir")}};assert.equal(context.consultarFichaTerritorial("paroquias","P01").ok,false);assert.equal(context.salvarFichaTerritorial({tipo:"paroquias",campos:{nome:"Teste"}}).ok,false);assert.equal(context.arquivarFichaTerritorial("missas","M01","").ok,false);});

test("cadastro independente de pastorais prevalece sobre dados de programação",()=>{const data={dioceses:[{id:"D01"}],paroquias:[{id:"P1",diocese:"D01",pastoral:"Ativa"}],comunidades:[],missas:[{id:"M1",diocese:"D01",unidade:"P1",pastoral:"Ativa"}],pastorais:[{id:"PAS1",unidade:"P1",diocese:"D01",situacaoPastoral:"Inativa"}]};assert.equal(nav.locaisPastoraisAtivas(data,new Set(["D01"]),"D01","","").length,0);data.pastorais[0].situacaoPastoral="Ativa";assert.equal(nav.locaisPastoraisAtivas(data,new Set(["D01"]),"D01","","").length,1);});

test("documentos regionais exigem autorização vigente e composição ativa",()=>{const a=["A","P","p@example.test","CO","COORDENACAO_REGIONAL","2026-01-01","2027-12-31","ATIVO","ADM","2026-01-01",""],c=["C","E","P","Secretaria","2026-01-01","2027-12-31","ATIVO","","","",""],e=["E","CO","Equipe","REGIONAL","COORDENACAO","DOCUMENTADA"];const ok=(aa,cc,ee)=>context.membroRegionalDocumentos_("p@example.test",aa,cc,ee,"2026-11-25");assert.equal(ok([a],[c],[e]),true);assert.equal(ok([], [c],[e]),false);assert.equal(ok([a],[],[e]),false);const revoked=[...a];revoked[10]="2026-11-01";assert.equal(ok([revoked],[c],[e]),false);const other=[...c];other[2]="P2";assert.equal(ok([a],[other],[e]),false);});
test("visitante recebe somente regimento e não consulta base privada",()=>{context.Session={getActiveUser:()=>({getEmail:()=>""})};let reads=0;context.SpreadsheetApp={openById:()=>{reads++;throw Error("não permitido")}};const r=context.consultarDocumentos();assert.equal(r.regional,false);assert.equal(r.restritos.length,0);assert.equal(r.publico.length,1);assert.equal(reads,0);});

test("vínculos de pessoa usam o mesmo ID em tabelas distintas",()=>{const data={VINCULOS_TERRITORIAIS:[Array.from(vm.runInContext("VINC_PESSOA_HEADERS_",context)),["V1","P1","U1","VINCULO_PASTORAL","","","","ATIVO","","Coordenação","Ouvinte"],["V2","P1","U2","VINCULO_PASTORAL","","","","ENCERRADO","","",""]],MEMBROS_PASTORAIS:[Array.from(vm.runInContext("MEMBROS_PASTORAIS_HEADERS_",context)),["M1","P1","PAS1","Membro","Ouvinte","","","ATIVO",""]],COMPOSICAO_REFERENCIAS:[Array.from(vm.runInContext("REF_COORD_HEADERS_",context)),["R1","E1","Pessoa Teste","Ouvinte","Secretaria","","2026–2027","","","P1"]],COMPOSICAO:[Array.from(vm.runInContext("COMPOSICAO_GESTAO_HEADERS_",context)),["R1","E1","P1","Secretaria","","","ATIVO","","","","","Ouvinte","2026–2027","","IMPORTACAO_REFERENCIAS","",""]]};const ss={getSheetByName:name=>({getLastRow:()=>data[name].length,getRange:()=>({getDisplayValues:()=>data[name]})})};const previousUtilities=context.Utilities;context.Utilities={formatDate:()=>"2026-11-25"};const rows=context.vinculosPessoas_(ss);context.Utilities=previousUtilities;assert.equal(rows.length,3);assert.equal(rows.every(x=>x.pessoa==="P1"),true);assert.equal(rows.some(x=>x.tipo==="Coordenação"),true);});
test("fluxo legado não cria pessoa paralela",()=>{assert.equal(context.salvarPreCadastro({}).ok,false);assert.equal(context.arquivarPreCadastro("X").ok,false);});

test("ficha única permite dados opcionais vazios, valida datas e exige autorização específica",()=>{
 const schema=vm.runInContext("CAMPOS_FICHA_UNICA_",context),old=Array(23+schema.length+5).fill("");
 assert.equal(context.validarCamposFichaUnica_({},old,"2026-11-25").values.DATA_BATISMO,"");
 assert.throws(()=>context.validarCamposFichaUnica_({camposFicha:{DATA_BATISMO:"2000-05-01"}},old,"2026-11-25"));
 const v=context.validarCamposFichaUnica_({camposFicha:{DATA_BATISMO:"2000-05-01",P14:"Sim"},autorizacaoOpcionais:true},old,"2026-11-25");
 assert.equal(v.values.DATA_BATISMO,"2000-05-01");assert.equal(v.values.P14,"Sim");
 assert.throws(()=>context.validarCamposFichaUnica_({camposFicha:{DATA_BATISMO:"2027-01-01"},autorizacaoOpcionais:true},old,"2026-11-25"));
 assert.throws(()=>context.validarCamposFichaUnica_({camposFicha:{P27:"2000-01-01"},autorizacaoOpcionais:true},old,"2026-11-25"));
 const i=schema.findIndex(f=>f.id==="DATA_BATISMO");old[23+i]="2000-05-01";assert.equal(context.validarCamposFichaUnica_({},old,"2026-11-25").values.DATA_BATISMO,"2000-05-01");
 assert.equal(context.coberturaFichaUnica_(old),1);
});
test("metadados de origem distinguem importação e edição",()=>{const row=Array(23).fill("");row[21]="IMPORTACAO_PRE_CADASTROS";assert.equal(context.metadadosFichaUnica_(row).origem,"ACAO_SISTEMICA");});
test("formulário único contém sacramentos editáveis e biblioteca tem filtros",()=>{assert.match(navHtml,/id="p-campos-unicos"/);assert.match(navHtml,/DATA_BATISMO/);assert.match(navHtml,/camposFicha:lerCamposFichaUnica/);assert.match(navHtml,/id="docs-categoria"/);assert.match(navHtml,/id="docs-mandato"/);assert.match(navHtml,/id="docs-situacao"/);});

test("identidade eclesial opcional não desloca metadados nem cria cargos",()=>{const headers=vm.runInContext("PESSOAS_HEADERS_",context),r=Array(headers.length).fill("");assert.equal(headers.indexOf("ORIGEM_CADASTRO"),64);assert.equal(headers.indexOf("CONDICAO_ECLESIAL"),69);r[64]="ACAO_SISTEMICA";r[69]="Ministro ordenado";assert.equal(context.metadadosFichaUnica_(r).origem,"ACAO_SISTEMICA");assert.equal(context.validarCamposFichaUnica_({},r,"2026-11-25").values.CONDICAO_ECLESIAL,"Ministro ordenado");assert.throws(()=>context.validarCamposFichaUnica_({camposFicha:{CONDICAO_ECLESIAL:"Coordenador"},autorizacaoOpcionais:true},r,"2026-11-25"));});
test("ficha apresenta navegação temática e opções compactas",()=>{assert.match(navHtml,/data-ficha-aba="vinculos"/);assert.match(navHtml,/data-ficha-aba="sacramentos"/);assert.match(navHtml,/id="p-ra-slot"/);assert.match(navHtml,/ficha-opcoes/);assert.match(navHtml,/Salvar alterações/);assert.match(navHtml,/CONDICAO_ECLESIAL/);});

test("cargos normalizam setor e gênero sem duplicar os vínculos",()=>{assert.equal(context.chaveCargo_("Vice-Coordenadora","Surdos"),context.chaveCargo_("Vice-coordenador","Surdo"));assert.notEqual(context.chaveCargo_("Coordenadora","Surdo"),context.chaveCargo_("Coordenador","Ouvinte"));});
test("mandato aceita datas desconhecidas e não inventa início",()=>{assert.equal(context.validarPeriodoCargo_("","",""),true);assert.equal(context.validarPeriodoCargo_("2027-01-01","2026-01-01",""),false);assert.equal(context.validarPeriodoCargo_("2026-02-30","",""),false);const r=["A","E","P","Secretaria","","","ATIVO","","","","","Ouvinte","2026–2027"];assert.equal(context.situacaoCargo_(r,"2026-11-25"),"Em exercício");r[6]="ENCERRADO";assert.equal(context.situacaoCargo_(r,"2026-11-25"),"Encerrado");});
test("opções de cargo por nível preservam a composição local e não impõem a nacional",()=>{const e=[["E","CO","Regional","REGIONAL"],["D","D09","Diocese","ARQUIDIOCESANO_DIOCESANO"]],r=[["R","E","Nome","Surdo","Coordenadora Regional"],["S","D","Nome","Ouvintes","Secretária"]];const regional=context.opcoesCargosEquipe_("E",e,r,[]);assert.equal(regional.some(x=>x.cargo==="Assistente Episcopal"),true);assert.equal(regional.some(x=>x.cargo==="Secretária"),false);const diocese=context.opcoesCargosEquipe_("D",e,r,[]);assert.equal(diocese.some(x=>x.cargo==="Secretária"),true);assert.equal(diocese.some(x=>x.cargo==="Assistente Episcopal"),false);});
test("gestão de cargos está disponível na ficha de coordenação",()=>{assert.match(navHtml,/id="cargo-form"/);assert.match(navHtml,/salvarCargoPessoa\(/);assert.match(navHtml,/encerrarCargoPessoa\(/);assert.match(navHtml,/Gerenciar cargo/);});

test("Portal do Agente reúne atalhos existentes e recursos públicos",()=>{assert.match(navHtml,/<section id="portal-agente" hidden>/);assert.match(navHtml,/data-page="portal-agente"/);for(const id of ["territorio","gestao","cadastro-pessoas","grupos","alertas","documentos"])assert.ok(navHtml.includes('data-open="'+id+'"'));assert.match(navHtml,/pastoraldosurdocentro\.blogspot\.com/);assert.match(navHtml,/Regimento Interno/);});

test("ficha ampliada nega acesso externo antes da leitura e limita consulta ao ID",()=>{
 const previous={Session:context.Session,SpreadsheetApp:context.SpreadsheetApp,territorialConfig_:context.territorialConfig_,gestaoConfig_:context.gestaoConfig_,tabela_:context.tabela_};
 let reads=0;
 try{
 context.territorialConfig_=()=>({owner:"owner@example.test",sourceId:"test"});
 context.gestaoConfig_=()=>({operacionalId:"test",questionarioId:"test"});
 context.Session={getActiveUser:()=>({getEmail:()=>"visitor@example.test"})};
 context.SpreadsheetApp={openById:()=>{reads++;return {};}};
 assert.equal(context.consultarComplementosFicha("P1","historico").ok,false);assert.equal(reads,0);
 context.Session={getActiveUser:()=>({getEmail:()=>"owner@example.test"})};
 context.tabela_=(_ss,aba)=>aba==="PESSOAS"?[["P1","Pessoa teste","","","","","","","","PARCIAL"],["P2","Outra pessoa","","","","","","","","PARCIAL"]]:aba==="EVENTOS"?[["E1","Evento teste"]]:aba==="PARTICIPACOES_EVENTOS"?[["R1","P1","E1","Inscrito","Não verificada","FONTE INTERNA","ORIGEM INTERNA","2026-11-25"],["R2","P2","E1","OUTRA PESSOA","Presente"]]:[];
 const result=context.consultarComplementosFicha("P1","historico");assert.equal(result.ok,true);
 const json=JSON.stringify(result);assert.match(json,/Evento teste/);assert.match(json,/Não verificada/);assert.doesNotMatch(json,/OUTRA PESSOA|FONTE INTERNA|ORIGEM INTERNA/);
 assert.equal(context.consultarComplementosFicha("P3","historico").ok,false);
 assert.equal(context.consultarComplementosFicha("P1","__proto__").ok,false);
 }finally{Object.assign(context,previous);}
});
test("ficha ampliada disponibiliza seções de consulta e novo indicador",()=>{
 for(const id of ["familia","socioeconomico","historico"])assert.ok(navHtml.includes('data-ficha-aba="'+id+'"'));
 assert.match(navHtml,/Coordenações pastorais cadastradas/);assert.match(navHtml,/consultarComplementosFicha\(id,secao\)/);
});

 test("painel conta programações com dia e hora sem exigir confirmação; distingue locais e pastorais",()=>{const c=vm.createContext({});vm.runInContext(navHtml.split("// BEGIN_HOME_INDICATORS")[1].split("// END_HOME_INDICATORS")[0],c);const rows=[{id:"1",unidade:"P1",dia:"Domingo",horario:"10h",statusProgramacao:"CONFIRMADA"},{id:"2",unidade:"P1",dia:"Terça",horario:"14h",statusProgramacao:"RECORRENTE INFORMADA"},{id:"3",unidade:"P2",dia:"Domingo",horario:"9h",statusProgramacao:"DIVERGENTE"},{id:"4",unidade:"P3",dia:"",horario:"",statusProgramacao:"A CONFIRMAR"}];assert.equal(c.programacoesPainel(rows).length,3);assert.equal(c.locaisProgramacaoPainel(rows),2);assert.equal(c.locaisPastoraisPainel([{unidade:"CO"},{unidade:"D01",diocese:"D01"},{unidade:"P1",diocese:"D01"},{unidade:"P1",diocese:"D01"},{unidade:"P4",diocese:"D01"}]),2);});
