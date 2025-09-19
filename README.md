Script_ad.py — Atualisação de atributos do Active Directory via XLSX por Pedro Costa

Este repositorio contém um script em Python que atualiza atributos de usuários no Active Directory a partir de uma planilha .xlsx. Eu o escrevi para ser simples de manter e flexível: basta ajustar os cabeçãlhos da planilha para alterar quais atributos serão atualizados. Este README também inclui o utilitario complementar busca_info.py, usado para exportar os usuários ativos antes de aplicar alterações em massa.


[1] busca_info.py — Exportação de usuários ATIVOS do AD para XLSX

Descrição
- Conecta ao Active Directory via LDAP/NTLM.
- Busca SOMENTE usuários ativos (exclui computadores e contas desabilitadas).
- Gera um arquivo XLSX com três colunas: Login (sAMAccountName), Mobile (mobile) e Title (title).
- Serve como PASSO INICIAL antes de rodar o Script_ad.py: use o XLSX gerado para revisar/ajustar valores e depois aplicar a atualização em massa.

Requisitos
- Python 3.8 ou superior
- Bibliotecas:
  pip install ldap3 openpyxl

Parâmetros
--server      Servidor LDAP (ex.: ldap://controlador.exemplo.local)
--user        Conta com permissão de leitura (formato DOMINIO\usuario)
--password    Senha da conta
--base        Base DN para a busca (ex.: "DC=exemplo,DC=local" ou "OU=Colaboradores,DC=exemplo,DC=local")
--output      Caminho do XLSX de saída (padrão: usuarios.xlsx)

Exemplos de uso (PowerShell)
# Exportar todos os usuários ativos do domínio:
python .\busca_info.py `
  --server ldap://controlador.exemplo.local `
  --user EXEMPLO\usuario `
  --password "SENHA_AQUI" `
  --base "DC=exemplo,DC=local" `
  --output usuarios.xlsx

# Exportar apenas de uma OU específica:
python .\busca_info.py `
  --server ldap://controlador.exemplo.local `
  --user EXEMPLO\usuario `
  --password "SENHA_AQUI" `
  --base "OU=Colaboradores,DC=exemplo,DC=local" `
  --output usuarios.xlsx

Recomendações
- Sempre gere e revise o usuarios.xlsx com o busca_info.py antes de atualizar atributos com o Script_ad.py.
- Se precisar de mais atributos (ex.: department, company), adicione-os na lista de atributos do script.


[2] Script_ad.py — Atualização de atributos via XLSX

O que o Script_ad.py faz
- Lê um .xlsx onde a 1ª coluna é o login do usuário (sAMAccountName) e as demais colunas representam atributos a serem atualizados.
- Localiza o DN do usuário via sAMAccountName.
- Atualiza atributos por MODIFY_REPLACE e, se uma célula estiver vazia, remove o atributo com MODIFY_DELETE.
- Aceita conexão LDAP simples, StartTLS (389) e LDAPS (636).
- Permite simular mudanças com --dry-run.

Requisitos
- Python 3.8 ou superior
- Bibliotecas:
  pip install ldap3 openpyxl

Mapeamento de apelidos (PT-BR → LDAP)
O script converte vários cabeçalhos comuns para os nomes de atributos LDAP. Se eu preferir, posso escrever diretamente os nomes de atributos LDAP e ignorar os apelidos.

login, sam, sAMAccountName → sAMAccountName (obrigatório na 1ª coluna)
telefone, celular, phone, mobile → mobile
cargo, titulo, título, title → title
departamento, department → department
empresa, company → company
gerencia → manager (espera o DN do gerente, ex.: CN=Fulano,OU=...,DC=...,DC=...)
descricao, description → description
cidade → l
estado → st
pais → co
rua → streetAddress
cep → postalCode

Se eu quiser suportar outros apelidos, basta editar o dicionário ALIAS no script.

Parâmetros e opções
--xlsx            Caminho do arquivo .xlsx
--server          Hostname/FQDN do Domain Controller (ex.: controlador.exemplo.local)
--base-dn         Base DN para a busca (ex.: "DC=exemplo,DC=local" ou "OU=Colaboradores,DC=exemplo,DC=local")
--user            Conta com permissão de escrita (UPN recomendado, ex.: usuario@exemplo.local)
--password        Senha da conta
--sheet           Nome da aba da planilha (opcional; se omitido, usa a aba ativa)
--skip-header     Se a primeira linha da planilha é cabeçalho
--security        ldap | starttls | ldaps
--insecure        Não valida o certificado TLS (apenas testes)
--dry-run         Simulação: não grava nada, apenas mostra o que faria

Recomendações práticas
- Eu sempre uso --user no formato UPN (usuario@dominio) para evitar problemas de sufixo.
- Ajusto --base-dn para a OU correta quando quero restringir a busca.
- Em produção, prefiro --security starttls ou --security ldaps sem --insecure.

1) Exemplos de uso (validação inicial)
Eu uso isso para validar cabeçalhos, escopo do base-dn e credenciais.

python .\Script_ad.py `
  --xlsx ".\usuarios.xlsx" `
  --server controlador.exemplo.local `
  --base-dn "DC=exemplo,DC=local" `
  --user "usuario@exemplo.local" `
  --password "SENHA_AQUI" `
  --skip-header `
  --security ldap `
  --dry-run

2) StartTLS (389) em produção
Teste primeiro sem validar o certificado:

python .\Script_ad.py `
  --xlsx ".\usuarios.xlsx" `
  --server controlador.exemplo.local `
  --base-dn "DC=exemplo,DC=local" `
  --user "usuario@exemplo.local" `
  --password "SENHA_AQUI" `
  --skip-header `
  --security starttls `
  --insecure `
  --dry-run

Aplicar de fato validando certificado (removo --insecure e --dry-run):

python .\Script_ad.py `
  --xlsx ".\usuarios.xlsx" `
  --server controlador.exemplo.local `
  --base-dn "DC=exemplo,DC=local" `
  --user "usuario@exemplo.local" `
  --password "SENHA_AQUI" `
  --skip-header `
  --security starttls

3) LDAPS (636) em produção
Teste sem validar certificado:

python .\Script_ad.py `
  --xlsx ".\usuarios.xlsx" `
  --server controlador.exemplo.local `
  --base-dn "DC=exemplo,DC=local" `
  --user "usuario@exemplo.local" `
  --password "SENHA_AQUI" `
  --skip-header `
  --security ldaps `
  --insecure `
  --dry-run

Aplicar validando certificado:

python .\Script_ad.py `
  --xlsx ".\usuarios.xlsx" `
  --server controlador.exemplo.local `
  --base-dn "DC=exemplo,DC=local" `
  --user "usuario@exemplo.local" `
  --password "SENHA_AQUI" `
  --skip-header `
  --security ldaps


TLS: requisitos no controlador de domínio
Para StartTLS/LDAPS funcionarem, o DC precisa ter um certificado com:
- EKU contendo Server Authentication
- Chave privada presente
- CN/SAN compatível com o FQDN que eu uso (ex.: controlador.exemplo.local)
- Instalado em Computador Local → Pessoal → Certificados

Eu valido portas com:
Test-NetConnection controlador.exemplo.local -Port 389
Test-NetConnection controlador.exemplo.local -Port 636


Dicas e cuidados
- Se eu usar --skip-header, a primeira linha da planilha é tratada como cabeçalho.
- Se eu deixar uma célula de atributo vazia, o script apaga aquele atributo no AD. Se não quero esse comportamento, eu removo a coluna do processamento.
- Se os usuários estiverem em OUs específicas, eu aponto --base-dn para a OU correta para acelerar e evitar ambiguidades.
- Para atributos que exigem DN (ex.: manager), eu devo fornecer o DN completo.
- Permissões: a conta informada em --user precisa ter permissão de escrita nos atributos que pretendo atualizar.

Troubleshooting
- Erro WinError 10054 (conexão encerrada pelo host remoto): geralmente LDAPS mal configurado no servidor. Eu valido primeiro com --security ldap --dry-run para testar bind e busca; depois ajusto certificado/portas e migro para StartTLS/LDAPS.
- startTLS failed - unavailable: o DC não está aceitando StartTLS. Verifico patches, políticas e certificado. Se necessário, uso LDAPS 636 ou corrijo o StartTLS no servidor.
- Falha de bind sem TLS e política “LDAP server signing requirements = Require”: eu uso StartTLS ou LDAPS.
- Usuário não encontrado: --base-dn fora do escopo correto. Eu aponto para a OU onde os usuários estão.
- Atributos não mudam: confirmar se eu não deixei --dry-run, se a conta tem permissão de escrita, e se os cabeçalhos correspondem aos atributos esperados.

Fluxo recomendado (resumo)
1) Rodar busca_info.py para extrair usuários ATIVOS → revisar/ajustar usuarios.xlsx.
2) Rodar Script_ad.py em --dry-run para validar mudanças.
3) Rodar Script_ad.py sem --dry-run (e com segurança TLS adequada) para aplicar.

Créditos: Pedro Costa
Linkedin: https://www.linkedin.com/in/pedro-costaf/
Github:   https://github.com/pdcfilho
