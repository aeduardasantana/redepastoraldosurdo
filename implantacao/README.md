# Implantação v0.15 — projeto consolidado existente
A estrutura real contém Código.gs consolidado (com configurações privadas), Index.html, Gestao.gs incremental, Cep.gs e Agenda.gs. Não copiar os módulos completos de apps-script/Gestao.gs e apps-script/Territorial.gs para esse projeto; causam declarações duplicadas.
1. Preservar Código.gs, Cep.gs, Agenda.gs, appsscript.json e propriedades privadas. Guardar cópia do Gestao.gs atual.
2. Substituir integralmente Gestao.gs por Gestao-complementar-v0.15.gs desta pasta (não acrescentar ao final).
3. Se Territorial.gs foi criado apenas durante a atualização e Código.gs já possui mapTerritorio_, retirar esse arquivo duplicado após guardar cópia. Não retirar arquivo anterior necessário.
4. Substituir Index.html por Index-v0.15.html da mesma pasta. Os dois arquivos devem ser da mesma revisão.
5. Salvar, testar /dev e confirmar cabeçalho v0.15, painel, ficha e Agenda Pastoral. Atualizar implantação existente somente após sucesso. Não modificar planilhas ou permissões.
O complemento usa nomes exclusivos v0.15; os endpoints existentes do arquivo consolidado permanecem intactos. Testes estáticos não comprovam implantação real.
