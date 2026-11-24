import test from "node:test";
import assert from "node:assert/strict";
import {createAuthClient} from "../web/b03/auth.mjs";
const id="00000000-0000-4000-8000-000000000031";
const fakeToken="token-ficticio-apenas-teste";
const reply=(data,status=200)=>({ok:status>=200&&status<300,status,json:async()=>data});
function mock(authorized=true){
  const calls=[];
  const client=createAuthClient(async(url,opts)=>{
    const path=new URL(url).pathname;calls.push({path,opts});
    if(path.endsWith("/token"))return reply({access_token:fakeToken,expires_in:3600});
    if(path.endsWith("/user"))return reply({id,is_anonymous:false});
    if(path.endsWith("/b03_test_profiles"))return reply(authorized?[{user_id:id,display_name:"Teste"}]:[]);
    if(path.endsWith("/b03_test_notes"))return opts.method==="POST"?reply(null,201):reply([{id:"teste",body:"Fictício"}]);
    if(path.endsWith("/logout"))return reply(null,204);
    throw new Error("Rota não prevista");
  });return{calls,client};
}
test("login simulado, perfil autorizado, nota e logout",async()=>{
  const {calls,client}=mock();
  assert.deepEqual(await client.signIn("teste@example.invalid","senha-ficticia"),{id,name:"Teste"});
  assert.equal((await client.notes()).length,1);await client.addNote("Nota fictícia");
  const profile=calls.find(c=>c.path.endsWith("/b03_test_profiles"));
  assert.equal(profile.opts.headers.Authorization,"Bearer "+fakeToken);
  assert.match(profile.opts.headers.apikey,/^sb_publishable_/);
  assert.equal(profile.opts.credentials,"omit");
  assert.equal(calls.find(c=>c.path.endsWith("/token")).opts.headers.Authorization,undefined);
  await client.logout();assert.equal(client.isSignedIn(),false);
  await assert.rejects(client.notes(),/Sessão expirada/);
});
test("nega perfil não autorizado e encerra sessão",async()=>{
  const {client,calls}=mock(false);
  await assert.rejects(client.signIn("teste@example.invalid","senha-ficticia"),/conta de teste autorizada/);
  assert.equal(client.isSignedIn(),false);assert.ok(calls.some(c=>c.path.endsWith("/logout")));
});
test("nega acesso sem sessão e falha de login",async()=>{
  const client=createAuthClient(async()=>reply({error:"interno"},400));
  await assert.rejects(client.notes(),/Sessão expirada/);
  await assert.rejects(client.signIn("teste@example.invalid","senha-ficticia"),/conta de teste autorizada/);
});
