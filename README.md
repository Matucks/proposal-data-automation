Proposal Data Automation

Este repositório contém um script em Python projetado para automatizar o processamento de dados de propostas e gerar relatórios organizados em arquivos Excel. O script inclui a criação de abas para dados brutos, equipes (Team A e Team B), tabelas por filial e uma tabela dinâmica consolidada. Além disso, o script oferece a opção de envio automático por e-mail.

Funcionalidades

Processamento Automático de Arquivos Excel: Localiza automaticamente o arquivo mais recente na pasta de entrada.

Organização dos Dados:

Cria uma aba para os dados brutos (Raw Data).

Gera abas separadas para Team A e Team B com base em unidades predefinidas.

Cria abas separadas para cada filial encontrada nos dados.

Gera uma aba de tabela dinâmica consolidada (Consolidated Pivot).

Formatação de Tabelas: Todas as abas geradas são formatadas como tabelas dinâmicas do Excel para facilitar a análise.

Envio por E-mail: Permite enviar o arquivo Excel gerado para uma lista de destinatários configuráveis.

Requisitos

Python 3.8 ou superior.

Bibliotecas necessárias:

pandas

openpyxl

smtplib

Para instalar as dependências, execute:

pip install pandas openpyxl

Configuração

Certifique-se de que os arquivos Excel a serem processados estejam localizados na pasta configurada como entrada:

INPUT_FOLDER = 'C:\\AutomationProject\\Input'

Configure o caminho de saída para o arquivo gerado:

OUTPUT_PATH = 'C:\\AutomationProject\\Output\\Proposals.xlsx'

Configure o caminho para o arquivo de dados de funcionários ativos:

DATA_PATH = 'C:\\AutomationProject\\Data\\ActiveEmployees.xlsx'

Ajuste as unidades associadas a Team A e Team B conforme necessário:

team_a_units = ["UnitA", "UnitB", "UnitC", "UnitD", "UnitE"]
team_b_units = ["UnitF", "UnitG", "UnitH", "UnitI", "UnitJ"]

Configure as informações de e-mail para envio automático:

sender_email = 'automation@example.com'
sender_password = 'yourpassword'
recipients = ['recipient1@example.com', 'recipient2@example.com']
smtp_server = 'smtp.example.com'

Como Executar

Clone o repositório para sua máquina local:

git clone https://github.com/seu-usuario/proposal-data-automation.git
cd proposal-data-automation

Execute o script principal:

python main.py

O arquivo Excel gerado será salvo no diretório especificado no parâmetro OUTPUT_PATH.

Após a geração, você será solicitado a confirmar o envio do arquivo por e-mail.

Estrutura do Arquivo Excel

O arquivo gerado contém as seguintes abas:

Raw Data: Dados brutos com colunas reordenadas e duplicatas removidas.

Team A: Dados filtrados para unidades específicas predefinidas em team_a_units.

Team B: Dados filtrados para unidades específicas predefinidas em team_b_units.

Filiais: Abas individuais criadas para cada filial (baseadas na coluna Unit).

Consolidated Pivot: Uma tabela dinâmica consolidada que agrupa dados por filial (Unit) e status (Status).

Todas as abas são formatadas como tabelas dinâmicas para facilitar a navegação e a organização.

Contribuições

Contribuições são bem-vindas! Para relatar problemas, sugerir melhorias ou enviar pull requests, utilize a aba "Issues" no repositório.

Licença

Este projeto está licenciado sob a MIT License.

Autor: Gabriel Matuck

Contato: gabriel.matuck1@gmail.com
