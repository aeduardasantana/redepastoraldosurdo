import { createAuthClient } from "./auth.mjs";
import { modules, resolveModule } from "./modules.mjs";
const nav = document.querySelector("#nav");
const view = document.querySelector("#view");
function element(tag, text, className) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  if (className) node.className = className;
  return node;
}
for (const module of modules) {
  const link = element("a", module.title);
  link.href = "#" + module.id;
  nav.append(link);
}
function render() {
  const selected = resolveModule(location.hash);
  for (const link of nav.querySelectorAll("a")) {
    if (link.getAttribute("href") === "#" + selected.id) link.setAttribute("aria-current","page");
    else link.removeAttribute("aria-current");
  }
  const intro = element("div");
  intro.append(element("p","MÓDULO · PROTÓTIPO","kicker"),element("h1",selected.title),element("p",selected.summary,"muted"));
  const cards = element("div",undefined,"cards");
  for (const feature of selected.features) {
    const card=element("article",undefined,"card");
    card.append(element("h2",feature),element("p","Funcionalidade planejada. Disponível após implementação e homologação.","muted"));
    cards.append(card);
  }
  view.replaceChildren(intro,cards,element("p","Navegação de demonstração. Nenhuma operação grava dados.","muted"));
  document.title=selected.title+" | Rede Pastoral do Surdo — Desenvolvimento";
}
window.addEventListener("hashchange",render);
render();

const auth=createAuthClient();
const login=document.querySelector("#login-form"), panel=document.querySelector("#private-panel"), status=document.querySelector("#auth-status"), notesList=document.querySelector("#notes"), noteForm=document.querySelector("#note-form");
function showSignedOut(){panel.hidden=true;login.hidden=false;notesList.replaceChildren();document.querySelector("#profile-title").textContent="";document.querySelector("#note").value="";}
async function refreshNotes(){const rows=await auth.notes();notesList.replaceChildren(...rows.map(row=>element("li",String(row.body))));}
login.addEventListener("submit",async event=>{
  event.preventDefault();const button=login.querySelector("button");button.disabled=true;status.textContent="Validando conta de teste…";
  const password=document.querySelector("#password");let value=password.value;password.value="";
  try{const profile=await auth.signIn(document.querySelector("#email").value,value);value="";await refreshNotes();document.querySelector("#profile-title").textContent=profile.name;login.hidden=true;panel.hidden=false;status.textContent="Conta de teste autorizada.";}
  catch(e){try{await auth.logout();}catch{}showSignedOut();status.textContent="Não foi possível entrar. Verifique a conta de teste autorizada.";}
  finally{value="";button.disabled=false;}
});
document.querySelector("#logout").addEventListener("click",async()=>{const b=document.querySelector("#logout");b.disabled=true;try{await auth.logout();status.textContent="Sessão encerrada.";}catch{status.textContent="Sessão local encerrada; servidor não confirmou a saída.";}finally{showSignedOut();b.disabled=false;}});
noteForm.addEventListener("submit",async event=>{
 event.preventDefault();const b=noteForm.querySelector("button");b.disabled=true;
 try{await auth.addNote(document.querySelector("#note").value);document.querySelector("#note").value="";await refreshNotes();status.textContent="Nota fictícia gravada.";}
 catch{status.textContent="Não foi possível gravar ou carregar a nota.";if(!auth.isSignedIn())showSignedOut();}
 finally{b.disabled=false;}
});
setInterval(()=>{if(!panel.hidden&&!auth.isSignedIn()){showSignedOut();status.textContent="Sessão expirada. Entre novamente.";}},1000);
