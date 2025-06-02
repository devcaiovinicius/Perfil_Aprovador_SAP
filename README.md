👤 Script em Python para Criação de Perfil Aprovador - SAP  

📌 Objetivo  
Automatizar o processo de criação de perfis de aprovadores no SAP via SAP GUI, utilizando Python para automatizar todo o processo, processo executado via transações PA30 e SE38.

⚙️ Funcionalidades  
- Leitura de dados a partir de planilha Excel (.xlsm)  
- Execução da transação PA30 para cadastro de dados pessoais, organizacionais e endereço  
- Geração de perfil através do programa `/SHCM/RH_SYNC_BUPA_EMPL_SINGLE` na SE38  
- Geração em lote para múltiplos usuários  

🧠 Pré-requisitos  
- Windows com SAP GUI instalado e acesso ao sistema SAP  
- SAP GUI Scripting habilitado (Solicitar habilitação ao time Basis da sua empresa)  
- Python (qualquer versão compatível com a biblioteca pywin32)  
- Biblioteca `pywin32` instalada
- Planilha Excel com os campos: Nome, Sobrenome, Usuário, ID (matrícula), Sexo  

💡 Observações  
- O script utiliza caminhos e valores fixos (ex: `BR01`, `50000000`, datas, etc). Adapte conforme a realidade do seu ambiente.  
- Certifique-se de executar o script com o SAP GUI aberto e logado.  

📂 Estrutura esperada da planilha Excel (aba `PERFIL APROVADOR`):

| Nome | Sobrenome | Usuário | ID | Sexo |
|------|-----------|---------|----|------|
| João | Silva     | joaoslv | 12345678 | M |
| Maria| Souza     | mrsouza | 87654321 | F |

-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
👤 Python Script for Approver Profile Creation - SAP  

📌 Objective  
Automate the creation of approver profiles in SAP via SAP GUI, using Python to automate the process, process executed via transactions PA30 and SE38.


⚙️ Features  
- Excel (.xlsm) spreadsheet data reading  
- Executes PA30 transaction to register personal, organizational, and address data  
- Runs `/SHCM/RH_SYNC_BUPA_EMPL_SINGLE` program in SE38 to generate the profile  
- Batch creation for multiple users  

🧠 Prerequisites  
- Windows with SAP GUI installed and access to the SAP system  
- SAP GUI Scripting enabled (Request the Basis team of your company to enable this feature)  
- Python (any version compatible with the `pywin32` library)  
- `pywin32` library installed (`pip install pywin32`)  
- Excel file containing the fields: First Name, Last Name, Username, ID (Employee Number), Gender  

💡 Notes  
- The script uses hardcoded values (e.g., `BR01`, `50000000`, specific dates, etc.). Adjust according to your company’s setup.  
- Ensure SAP GUI is open and logged in before executing the script.
- 

📂 Expected Excel Spreadsheet Structure (`PERFIL APROVADOR` sheet):

| First Name | Last Name | Username | ID       | Gender |
|------------|-----------|----------|----------|--------|
| João       | Silva     | joaoslv  | 12345678 | M      |
| Maria      | Souza     | mrsouza  | 87654321 | F      |

