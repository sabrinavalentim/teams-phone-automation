# Configurar e Remover Ramais no Microsoft Teams com PowerShell GUI
 
Este repositório contém dois scripts PowerShell com interface gráfica para facilitar a **configuração e remoção de ramais (números Direct Routing)** no Microsoft Teams.
 
## Funcionalidades
 
- Interface intuitiva com campos de e-mail e número
- Validação de conexão com Teams
- Execução de comandos via `Set-CsPhoneNumberAssignment`
- Suporte a políticas de voz existentes
- Executável com ícone personalizado (usando ferramentas como PS2EXE)
 
## Requisitos
 
- PowerShell 5.1 ou superior
- Módulo MicrosoftTeams (`Install-Module MicrosoftTeams`)
- Permissões administrativas no Teams para modificar ramais
 
## Como usar
 
1. Execute o script `.ps1` com PowerShell (como administrador).
2. Preencha o e-mail e número.
3. Clique em **"Configurar Ramal"**.
 
## Autor
 
Este projeto foi criado por [Sabrina Santos] com o objetivo de automatizar e simplificar rotinas de administração do Microsoft Teams em ambientes corporativos.
