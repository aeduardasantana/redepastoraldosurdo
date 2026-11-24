# Rede Pastoral do Surdo

Aplicação web própria de relacionamento e gestão pastoral. Regional Centro-Oeste.

## Arquitetura aprovada — 24/11/2026
- Interface estática modular HTML, CSS e JavaScript; PWA opcional.
- Supabase Free (PostgreSQL e Auth); componentes open source gratuitos, com licenças verificadas e versões fixadas.
- GitHub Free e Actions; Locaweb existente para publicação segura; Google Drive para acervo institucional privado.
- Orçamento obrigatório: R$ 0,00 de custo adicional. Nenhum upgrade, contratação ou excedente pago autorizado.
- HumHub, MySQL/MariaDB e PHP/cron da Locaweb não são dependências da solução.

## Estado atual
B01 concluído; B02 em validação de infraestrutura. A demonstração visual não possui autenticação nem persistência. Não inserir dados reais. B03 ainda não iniciado.

## Publicação
Somente arquivos estáticos autorizados em `web/`. O workflow B02 realiza diagnóstico sem upload automático. Teste isolado de publicação exige ação manual do titular, identidade SSH verificada por chave confiável, SFTP criptografado, HTTPS com certificado válido, comparação de marcador e remoção exclusiva do teste. Nunca usar FTP simples ou ignorar validação de certificado. Credenciais ficam nos GitHub Actions Secrets, não no código. Não substituir arquivos existentes no servidor.

## Privacidade e próximas etapas
Repositório público temporariamente durante desenvolvimento com dados fictícios. Tornar privado antes da implantação definitiva ou uso de dados reais. Não publicar dados pessoais, documentos internos, tokens, chaves privilegiadas ou backups no GitHub. Usar apenas chaves publicáveis no cliente quando o B03 for autorizado; RLS e autorização territorial no banco, nunca apenas na interface. Não criar banco, tabelas ou usuários neste B02.
