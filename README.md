# Rede Pastoral do Surdo

Aplicação web de relacionamento e gestão pastoral. Regional Centro-Oeste.

## Estado atual
Etapa de teste de infraestrutura. A demonstração visual não possui autenticação nem banco de dados; não inserir dados reais.

## Publicação
Apenas arquivos estáticos em `web/` devem ser enviados ao diretório exclusivo da aplicação. Credenciais ficam nos GitHub Actions Secrets (HOST, USER, PASS), nunca no código. Exigir FTPS com certificado válido e HTTPS para o aplicativo. Não remover arquivos existentes no servidor.

## Próxima etapa
Tornar o repositório privado, validar autenticação e isolamento territorial em ambiente dedicado, aprovar o tratamento de dados e homologar o piloto.
