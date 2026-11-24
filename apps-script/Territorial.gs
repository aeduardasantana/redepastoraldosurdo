// Cadastro institucional, somente leitura. Configuração privada no pacote de implantação.
function mapTerritorio_(d,p,c,m){
 const fields=(row,indices)=>indices.map(i=>String(row[i]??"").trim());
 const distinct=(rows)=>{const ids=new Set();for(const r of rows){if(!r[0]||ids.has(r[0]))throw new Error("Identificador ausente ou duplicado.");ids.add(r[0]);}return ids;};
 const dioceses=d.map(r=>{const [id,nome,provincia,uf,fonte,situacao,cidade,logradouro,bairro,cep,obs,fonteContato,data,validacao]=fields(r,[0,1,2,3,4,5,6,7,8,9,16,17,19,20]);return {id,nome,provincia,uf,tipo:id==="D13"?"Jurisdição pessoal complementar":nome.startsWith("Arquidiocese")?"Arquidiocese":"Diocese",cidade,logradouro,bairro,cep,situacao,fonte,fonteContato,data,validacao,obs};});
 const paroq=p.map(r=>{const [id,diocese,nome,tipo,cidade,bairro,logradouro,cep,fonte,fonteFicha,data,situacao,pastoral]=fields(r,[0,1,3,4,5,6,7,8,17,18,19,20,21]);return {id,diocese,nome,tipo,cidade,bairro,logradouro,cep,fonte,fonteFicha,data,situacao,pastoral};});
 const comunidades=c.map(r=>{const [id,paroquia,diocese,nome,cidade,bairro,logradouro,cep,fonte,data,situacao,obs]=fields(r,[0,1,2,3,4,5,6,7,10,11,12,13]);return {id,paroquia,diocese,nome,cidade,bairro,logradouro,cep,fonte,data,situacao,obs};});
 const missas=m.map(r=>{const [id,diocese,cidade,nome,logradouro,dia,horario,situacao,data,fonte,validacao]=fields(r,[0,1,2,3,4,5,6,7,8,9,10]);return {id,diocese,cidade,nome,logradouro,dia,horario,situacao,data,fonte,validacao};});
 distinct(d);distinct(p);distinct(c);distinct(m);
 const did=new Set(dioceses.map(x=>x.id));const pmap=new Map(paroq.map(x=>[x.id,x]));
 if(paroq.some(x=>!did.has(x.diocese))||comunidades.some(x=>!pmap.has(x.paroquia)||pmap.get(x.paroquia).diocese!==x.diocese))throw new Error("Vínculo territorial inconsistente.");
 return {dioceses,paroquias:paroq,comunidades,missas:missas.map(x=>({...x,diocese:did.has(x.diocese)?x.diocese:"",vinculoPendente:!did.has(x.diocese)}))};
}

function lerTerritorio_(ss,nome,cabecalho,largura,max){
 const s=ss.getSheetByName(nome);
 if(!s||s.getLastRow()<1||s.getLastRow()>max)throw new Error("Limite ou estrutura da fonte.");
 const v=s.getRange(1,1,s.getLastRow(),largura).getDisplayValues();
 if(!cabecalho.every((x,i)=>v[0][i]===x))throw new Error("Cabeçalho incompatível.");
 return v.slice(1).filter(r=>r.some(x=>String(x??"").trim()));
}
function consultarTerritorio(){
 try{
   const cfg=typeof territorialConfig_==="function"?territorialConfig_():null;
   const email=normalizarEmail_(Session.getActiveUser().getEmail());
   if(!cfg||!email||email!==normalizarEmail_(cfg.owner))return {ok:false,reason:"Consulta restrita à conta institucional de teste."};
   const ss=SpreadsheetApp.openById(cfg.sourceId);
   const data=mapTerritorio_(
     lerTerritorio_(ss,"DIOCESES",["ID","Jurisdição","Província"],21,1000),
     lerTerritorio_(ss,"PARÓQUIAS",["ID unidade","ID diocese","Jurisdição"],26,1000),
     lerTerritorio_(ss,"COMUNIDADES",["ID comunidade","ID unidade matriz","ID diocese"],14,1000),
     lerTerritorio_(ss,"MISSAS LIBRAS",["ID referência","ID diocese (validar)","Cidade / região"],11,1000));
   return {ok:true,data};
 }catch(e){return {ok:false,reason:"Não foi possível consultar o cadastro institucional. Verifique a fonte e a autorização."};}
}
