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
 assert.match(h,/v0\.9/);assert.match(h,/Não há gravação de respostas/);
 assert.match(h,/Nenhum voto ou candidatura será registrado/);
 assert.doesNotMatch(h,/salvarPessoa\(|registrarVoto\(/);
 const portal=await readFile(new URL("../publico/index.html",import.meta.url),"utf8");
 assert.match(portal,/Acessar área privada/);assert.match(portal,/pastoraldosurdocentro\.blogspot\.com/);
 assert.doesNotMatch(portal,/docs\.google\.com\/spreadsheets|google\.script\.run/);
});

test("cadastro apresenta data de nascimento e endereço estruturado sem enviar dados",async()=>{const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");assert.match(h,/Data de nascimento/);assert.match(h,/el.type="date"/);assert.match(h,/el.inputMode="numeric"/);assert.match(h,/P55","P38","P52","P04","P37","P39","P53","P54"/);assert.doesNotMatch(h,/salvarPessoa\(/);});

test("navegação por módulos inicia na home, mantém um painel visível e carrega cada consulta",async()=>{const h=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");const src=h.split("// BEGIN_PAGE_NAV")[1].split("// END_PAGE_NAV")[0];const ctx=vm.createContext({});vm.runInContext(src,ctx);assert.equal(ctx.paginaValida("cadastro-pessoas"),"cadastro-pessoas");assert.equal(ctx.paginaValida("inexistente"),"home");const ids=["home","territorio","inclusao","gestao","cadastro-pessoas","eleicoes","acessos-gestao","pagina-publica","acesso"];for(const id of ids){assert.ok(h.includes('data-page="'+id+'"'));assert.ok(h.includes('<section id="'+id+'"'+(id==="home"?">":" hidden>")));}assert.match(h,/T\(p\)\.hidden=p!==id/);assert.match(h,/\(id==="gestao"\|\|id==="cadastro-pessoas"\)&&!gestao\)carregarGestao\(\)/);assert.match(h,/abrirPagina\("home"\);/);assert.match(h,/questionario-status/);assert.match(h,/function atualizarHome\(/);assert.doesNotMatch(h,/<nav aria-label="Menu"><a/);});
