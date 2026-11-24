import test from "node:test";
import assert from "node:assert/strict";
import {readFile} from "node:fs/promises";
import {modules,resolveModule} from "../web/b03/modules.mjs";
test("doze módulos, identificadores únicos e conteúdo",()=>{
  assert.equal(modules.length,12);
  assert.equal(new Set(modules.map(m=>m.id)).size,12);
  for(const m of modules){assert.match(m.id,/^[a-z]+$/);assert.ok(m.title&&m.summary);assert.ok(m.features.length);}
});
test("navegação válida e fallback",()=>{
  for(const m of modules) assert.equal(resolveModule("#"+m.id),m);
  assert.equal(resolveModule("#desconhecido"),modules[0]);
  assert.equal(resolveModule(""),modules[0]);
});
test("entrada estática referencia recursos locais e alerta de teste",async()=>{
  const html=await readFile(new URL("../web/b03/index.html",import.meta.url),"utf8");
  assert.match(html,/lang="pt-BR"/);assert.match(html,/Não cadastrar dados reais/);
  assert.match(html,/src="\.\/app\.mjs"/);assert.match(html,/href="\.\/styles\.css"/);
  assert.doesNotMatch(html,/(?:src|href)="https?:\/\//);
});
