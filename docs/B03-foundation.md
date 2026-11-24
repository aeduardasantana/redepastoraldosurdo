# B03 — Fundação de desenvolvimento (parcial)

A branch b03-desenvolvimento preserva a demonstração original em web/index.html. A nova base modular está em web/b03/. Não publicar, mesclar em main ou habilitar dados reais enquanto B02 e homologação não forem concluídos. Nenhuma instalação local necessária.

## Componentes e custo
HTML5, CSS, JavaScript ES Modules, APIs nativas do navegador e Node.js para testes na CI. Sem bibliotecas de terceiros adicionadas nesta etapa. Custo adicional zero. Antes de incorporar Supabase JS, verificar versão estável, licença, lockfile, build e chave publicável. Não armazenar tokens privilegiados no cliente.

## Banco e login
O projeto Supabase existente é o único previsto. Nesta entrega não existem migrações aplicadas, tabelas criadas, login real ou integração da interface com o banco. B03 continuará com migração versionada mínima, RLS desde a primeira tabela, grants mínimos e conta fictícia autorizada; B05 tratará a hierarquia e mandatos completos. Não confundir navegação demonstrativa com autorização. Não inserir dados reais.

## Teste e reversão
A CI da branch executa validações sem deploy. Preservar main e histórico; reverter commits desta branch em caso de defeito. Backup e restauração do banco ainda precisam de homologação. B02 continua pendente de publicação segura na Locaweb. B04–B12 não iniciados.
