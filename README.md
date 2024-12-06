# contatosKomunic
Automação de Cadastro de Contatos no Komunic
![image](https://github.com/user-attachments/assets/d40f7aae-9787-4e44-9010-dd67efee1307)


Este script automatiza o processo de cadastro em massa de contatos no sistema Komunic, utilizando Selenium para interagir com a interface web.

A ferramenta foi criada para agilizar o preenchimento de dados repetitivos, como nome e telefone, a partir de uma planilha .xlsx. Ela também valida números de WhatsApp antes do cadastro.

Funcionalidades:
Login automatizado: Efetua login com credenciais fornecidas.
Processamento em lote: Lê dados de nomes e números a partir de um arquivo Excel.
Verificação de WhatsApp: Confirma se os números possuem WhatsApp antes de serem cadastrados.
Normalização de dados: Remove caracteres especiais e formata os dados para evitar erros no sistema.
Execução simples: Basta configurar um arquivo contatos.xlsx e executar o script.
Tecnologias utilizadas:
Python (bibliotecas: Selenium, OpenPyXL, Docx)
Selenium WebDriver para automação do navegador
Microsoft Edge como browser automatizado
Como contribuir:
Se você encontrou um bug ou deseja adicionar uma funcionalidade (como vinculação a organizações), contribuições são bem-vindas!
