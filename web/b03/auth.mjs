import { projectUrl, publishableKey } from "./supabase-config.mjs";
// Cliente de desenvolvimento: HTTPS, sem dependências externas e sem persistência de tokens.
export function createAuthClient(request = globalThis.fetch, now = () => Date.now()) {
  let session = null;
  const valid = () => session && session.expiresAt > now() + 5000;
  async function call(path, {method = "GET", body, token, prefer} = {}) {
    const headers = {apikey:publishableKey,Accept:"application/json"};
    if(token) headers.Authorization = "Bearer " + token;
    if(body !== undefined) headers["Content-Type"]="application/json";
    if(prefer) headers.Prefer=prefer;
    const res = await request(projectUrl + path,{method,headers,body:body===undefined?undefined:JSON.stringify(body),cache:"no-store",credentials:"omit",redirect:"error",referrerPolicy:"no-referrer"});
    if(!res.ok) throw new Error("Solicitação recusada ou indisponível.");
    if(res.status===204||res.status===205||prefer==="return=minimal"||path.startsWith("/auth/v1/logout"))return null;
    return res.json();
  }
  async function logout(){
    const token=session?.accessToken;
    session=null;
    if(token)await call("/auth/v1/logout?scope=local",{method:"POST",token});
  }
  function token(){
    if(!valid()){session=null;throw new Error("Sessão expirada. Entre novamente.");}
    return session.accessToken;
  }
  async function signIn(email,password){
    if(session)await logout();
    session=null;
    if(typeof email!=="string"||typeof password!=="string"||!email.trim()||!password)throw new Error("Informe e-mail e senha de teste.");
    try{
      const result=await call("/auth/v1/token?grant_type=password",{method:"POST",body:{email:email.trim(),password}});
      if(typeof result?.access_token!=="string"||!result.access_token||!Number.isFinite(Number(result.expires_in))||Number(result.expires_in)<=5)throw new Error("Resposta inválida.");
      session={accessToken:result.access_token,expiresAt:now()+Number(result.expires_in)*1000};
      const user=await call("/auth/v1/user",{token:token()});
      if(typeof user?.id!=="string"||!user.id||user.is_anonymous===true)throw new Error("Conta não autorizada.");
      const profile=await call("/rest/v1/b03_test_profiles?select=user_id,display_name&user_id=eq."+encodeURIComponent(user.id)+"&limit=1",{token:token()});
      if(!Array.isArray(profile)||profile.length!==1||profile[0].user_id!==user.id)throw new Error("Perfil não autorizado.");
      return {id:user.id,name:String(profile[0].display_name)};
    }catch(e){try{await logout();}catch{}throw new Error("Não foi possível entrar com uma conta de teste autorizada.");}
  }
  async function notes(){
    const rows=await call("/rest/v1/b03_test_notes?select=id,body,created_at&order=created_at.desc&limit=30",{token:token()});
    if(!Array.isArray(rows))throw new Error("Resposta inválida.");
    return rows;
  }
  async function addNote(body){
    if(typeof body!=="string"||!body.trim()||body.trim().length>500)throw new Error("Informe até 500 caracteres fictícios.");
    await call("/rest/v1/b03_test_notes",{method:"POST",token:token(),body:{body:body.trim()},prefer:"return=minimal"});
  }
  return Object.freeze({signIn,logout,notes,addNote,isSignedIn:()=>Boolean(valid())});
}
