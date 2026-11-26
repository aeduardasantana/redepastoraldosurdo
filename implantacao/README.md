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

## Endereço institucional integrado
Instale também `EnderecoRede.gs` (depende de `EnderecoDiocesano.gs` e `Cep.gs`) e o Index da mesma revisão. O catálogo oferece CO, dioceses, paróquias, comunidades, locais e organizações. Endereço físico de organização sem registro só é criado com confirmação explícita; abrangência não é usada como sede. CO e ORG001 apontam ao endereço CI-CNBBCO-001. Locais vinculados resolvem o ID da unidade; locais independentes usam ID próprio. O mestre é ENDERECOS_INSTITUCIONAIS. A API Sheets é atômica por arquivo, não entre arquivos: o espelho LOCAIS_VINCULADOS é reconciliado após gravar o mestre; falhas são exibidas e podem ser repetidas com Sincronizar espelho. Não editar o espelho isoladamente. Novas unidades mantêm o cadastro territorial legado e podem criar a ficha central pelo catálogo. Homologar os fluxos de leitura e gravação em /dev antes do /exec.

Os dois contatos/locais institucionais existentes (INOSEB e Centro Pastoral São Filippo Smaldone) também aparecem na categoria organizações, como outra visão dos mesmos IDs, sem criar endereços duplicados. As relações congregacionais e mantenedoras só deverão ser cadastradas com fonte própria; não são inferidas pela localização.
