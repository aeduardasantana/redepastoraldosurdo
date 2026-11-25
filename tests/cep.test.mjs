import test from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import {readFile} from "node:fs/promises";
const src=await readFile(new URL("../apps-script/Cep.gs",import.meta.url),"utf8");
const ctx=vm.createContext({});vm.runInContext(src,ctx);
test("CEP: normalização preserva zeros e rejeita formato inválido",()=>{
 assert.equal(ctx.redeNormalizarCep_("01001-000"),"01001000");
 assert.equal(ctx.redeNormalizarCep_("01001000"),"01001000");
 for(const s of ["123","01001x000","01001 000","010010000"])assert.equal(ctx.redeNormalizarCep_(s),"");
});
test("CEP: valida resposta postal e preserva logradouro em quadra",()=>{
 const data=ctx.redeCepPostal_({cep:"73752-075",logradouro:"Quadra 5 MR 1",bairro:"Setor Leste",localidade:"Planaltina",uf:"GO",ibge:"5217609"},"73752075");
 assert.equal(data.logradouro,"Quadra 5 MR 1");assert.equal(data.localidade,"Setor Leste");
 assert.equal(ctx.redeCepPostal_({erro:true},"73752075"),null);
 assert.equal(ctx.redeCepPostal_({cep:"01001-000",localidade:"São Paulo",uf:"SP"},"73752075"),null);
});
test("CEP: visitante não autorizado não acessa a planilha nem serviço externo",()=>{
 ctx.autorizacaoTerritorial_=()=>null;
 ctx.SpreadsheetApp={openById:()=>{throw Error("não deve consultar")}};
 ctx.UrlFetchApp={fetch:()=>{throw Error("não deve consultar")}};
 assert.equal(ctx.consultarCepRede("01001-000").reason,"Consulta restrita.");
});
test("CEP: cache confirmado recente não consulta API e não envia identificadores pessoais",()=>{
 ctx.autorizacaoTerritorial_=()=>({sourceId:"TEST"});
 const headers=["CEP","LOGRADOURO_POSTAL","BAIRRO_POSTAL","MUNICIPIO_POSTAL","UF","CODIGO_IBGE","STATUS_VALIDACAO","FONTE_ORIGEM","REFERENCIAS_TERRITORIAIS","FONTE_CONSULTA","DATA_CONSULTA"];
 const rows=[headers,["01001-000","Praça da Sé","Sé","São Paulo","SP","3550308","Confirmado","Origem","ID-INSTITUCIONAL","Fonte postal","2026-11-25"]];
 ctx.SpreadsheetApp={openById:()=>({getSheetByName:()=>({getLastRow:()=>2,getRange:(r,c,n,w)=>({getDisplayValues:()=>rows.slice(r-1,r-1+n).map(x=>x.slice(c-1,c-1+w))})})})};
 ctx.Utilities={formatDate:()=>"2026-11-25"};
 ctx.UrlFetchApp={fetch:()=>{throw Error("cache deve evitar API")}};
 const out=ctx.consultarCepRede("01001-000");assert.equal(out.ok,true);assert.equal(out.cached,true);assert.doesNotMatch(JSON.stringify(out),/ID-INSTITUCIONAL/);
});
test("CEP: interface possui campo único e preenchimento explícito",async()=>{
 const html=await readFile(new URL("../apps-script/Index.html",import.meta.url),"utf8");
 assert.match(html,/Bairro \/ Setor \/ Região Administrativa/);
 assert.match(html,/id="p-ra-slot" hidden/);
 assert.match(html,/consultarCepRede\(cep\)/);
 assert.match(html,/id="p-cep-aplicar"/);
});
