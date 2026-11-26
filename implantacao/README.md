# Implantação v0.15 — projeto consolidado existente
A estrutura real contém Código.gs consolidado (com configurações privadas), Index.html, Gestao.gs incremental, Cep.gs e Agenda.gs. Não copiar os módulos completos de apps-script/Gestao.gs e apps-script/Territorial.gs para esse projeto; causam declarações duplicadas.
1. Preservar Código.gs, Cep.gs, Agenda.gs, appsscript.json e propriedades privadas. Guardar cópia do Gestao.gs atual.
2. Substituir integralmente Gestao.gs por Gestao-complementar-v0.15.gs desta pasta (não acrescentar ao final).
3. Se Territorial.gs foi criado apenas durante a atualização e Código.gs já possui mapTerritorio_, retirar esse arquivo duplicado após guardar cópia. Não retirar arquivo anterior necessário.
4. Substituir Index.html por Index-v0.15.html da mesma pasta. Os dois arquivos devem ser da mesma revisão.
5. Salvar, testar /dev e confirmar cabeçalho v0.15, painel, ficha e Agenda Pastoral. Atualizar implantação existente somente após sucesso. Não modificar planilhas ou permissões.
O complemento usa nomes exclusivos v0.15; os endpoints existentes do arquivo consolidado permanecem intactos. Testes estáticos não comprovam implantação real.

## Endereço das cúrias
Criar o arquivo complementar `EnderecoDiocesano.gs` com o conteúdo de `apps-script/EnderecoDiocesano.gs` e substituir `Index.html` pelo pacote atualizado. Não alterar Código.gs nem copiar Territorial.gs inteiro. A gravação usa a API Google Sheets (escopo spreadsheets já previsto), que deve estar habilitada no projeto Google Cloud associado. O endpoint só edita endereços existentes D01–D12; mantém o ID e atualiza a ficha padronizada e a projeção em uma única operação. D13 não é jurisdição territorial. Homologar antes de atualizar /exec.
