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
