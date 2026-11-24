import assert from "node:assert/strict";
import {projectUrl,publishableKey} from "../web/b03/supabase-config.mjs";
const headers={apikey:publishableKey,Accept:"application/json"};
const health=await fetch(projectUrl+"/auth/v1/health",{headers,signal:AbortSignal.timeout(15000)});
assert.equal(health.status,200,"Auth não respondeu com HTTP 200");
console.log("Auth: endpoint de saúde respondeu HTTP 200.");
for(const table of ["b03_test_profiles","b03_test_notes"]){
  const res=await fetch(projectUrl+"/rest/v1/"+table+"?select=*&limit=1",{headers,signal:AbortSignal.timeout(15000)});
  assert.ok([401,403].includes(res.status),"Acesso anônimo não negado como esperado: "+res.status);
  console.log("Data API: "+table+" acesso anônimo bloqueado HTTP "+res.status);
}
console.log("Teste externo sem cadastro, autenticação positiva ou escrita.");
