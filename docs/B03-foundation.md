# B03 — Fundação de desenvolvimento (parcial)
A branch b03-desenvolvimento preserva web/index.html e main. A base modular está em web/b03/. Não publicar ou usar dados reais enquanto B02 e homologação não forem concluídos.

## Componentes e custo
HTML5, CSS, JavaScript ES Modules, Fetch API nativa e Node.js na CI. Sem bibliotecas adicionais, sem cobrança. Integração REST com Supabase Auth e Data API via HTTPS e chave publicável moderna; não armazenar senha, refresh token ou access token em código/storage do navegador. Sessão apenas em memória e nova entrada após recarga/expiração.

## Migração e testes
Migração 20261124230147_b03_test_foundation.sql aplicada ao projeto exclusivo Free e versionada. Tabelas b03_test_profiles e b03_test_notes, cinco políticas RLS, grants mínimos; perfil autorizado só por administrador e associado ao Auth. Nenhuma estrutura territorial definitiva (B05).
Teste transacional SQL com identidades simuladas passou: leitura/CRUD próprio, bloqueio cruzado, perfil inativo, sessão anônima, modificação de autorização/titularidade. ROLLBACK e contagens zero confirmados.
CI 45506592246 (commit 91d95714f58595dee031ec14897e8f820b20c067) aprovou seis testes de código, chamadas reais: saúde Auth 200, consulta anônima em ambas tabelas 401, senha fictícia inválida 400. Build estático sem deploy. Login positivo e logout simulados no teste unitário, não comprovados com conta real de teste.

## Pendências
O conector Supabase não oferece criação de usuário Auth pela API administrativa; não inserir manualmente em auth.users. Criar duas contas fictícias controladas pela responsável via painel Authentication > Users > Add user, sem enviar convite a terceiros. Autorizar cada UUID em b03_test_profiles exclusivamente por operação administrativa, então testar login positivo e acesso cruzado via API com credenciais temporárias, sem gravar senhas em GitHub/Drive/chat. Nenhuma conta criada até aqui. Não habilitar cadastro aberto na interface. Verificar no painel configuração de cadastro/convites e sessões antes do uso real.
Advisors: zero avisos de segurança retornados; um aviso de desempenho auth_rls_initplan no perfil (https://supabase.com/docs/guides/database/database-linter?lint=0003_auth_rls_initplan). Reavaliar antes de escala. Backup/restauração e revisão visual ainda pendentes. B02 precisa comprovar publicação segura. Não iniciar B04.
