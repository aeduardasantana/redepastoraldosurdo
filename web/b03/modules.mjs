export const modules = Object.freeze([
  {id:"rede", title:"Minha rede", summary:"Comunicados e convivência pastoral.", features:["Mural de publicações","Comunicados regionais","Comentários e moderação"]},
  {id:"coordenacoes", title:"Coordenações", summary:"Organização das instâncias pastorais.", features:["Regional","(Arqui)dioceses","Paróquias"]},
  {id:"comunidade", title:"Comunidade", summary:"Pessoas, agentes e vínculos autorizados.", features:["Cadastro pastoral","Perfis","Preferências de comunicação"]},
  {id:"agenda", title:"Agenda", summary:"Compromissos, celebrações e formações.", features:["Calendário","Atividades","Lembretes"]},
  {id:"assembleias", title:"Assembleias", summary:"Decisões e encaminhamentos institucionais.", features:["Convocação e pauta","Atas","Deliberações"]},
  {id:"eventos", title:"Eventos e atividades", summary:"Planejamento, realização e acompanhamento.", features:["Inscrições quando autorizadas","Relatórios","Evidências"]},
  {id:"memoria", title:"Memória pastoral", summary:"Preservação do acervo institucional.", features:["Fotografias autorizadas","Documentos","Histórico"]},
  {id:"relatorios", title:"Relatórios", summary:"Fluxos de envio, revisão e aprovação.", features:["Rascunhos","Pendências","Consolidação"]},
  {id:"mandatos", title:"Mandatos", summary:"Cargos, vigência e transição.", features:["Composição","Histórico","Alertas de vigência"]},
  {id:"parceiros", title:"Rede de parceiros", summary:"Relacionamentos autorizados.", features:["Diretório","Aprovação","Vigência"]},
  {id:"secretaria", title:"Secretaria", summary:"Documentação e organização institucional.", features:["Atas e ofícios","Modelos","Integração documental"]},
  {id:"monitoramento", title:"Monitoramento", summary:"Indicadores para coordenações autorizadas.", features:["Pendências","Indicadores agregados","Acompanhamento territorial"]}
]);
export function resolveModule(hash) {
  const id = String(hash || "").replace(/^#/, "");
  return modules.find(module => module.id === id) ?? modules[0];
}
