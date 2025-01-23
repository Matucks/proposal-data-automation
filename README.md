# Proposal Data Automation

**Proposal Data Automation** é um projeto Python desenvolvido para automatizar o processamento de dados de propostas e gerar relatórios detalhados em arquivos Excel. O script organiza os dados em abas específicas, formata tabelas para análise, e inclui a opção de envio automático por e-mail.

---

## Funcionalidades Principais

### 1. Processamento Automático
- Identifica e processa automaticamente o arquivo mais recente na pasta de entrada.

### 2. Organização dos Dados
- **Raw Data**: Aba com os dados brutos reorganizados.
- **Team A e Team B**: Dados separados com base em unidades associadas a cada equipe.
- **Filiais**: Abas específicas criadas para cada filial encontrada nos dados.
- **Consolidated Pivot**: Tabela dinâmica consolidada agrupando os dados por filial e status.

### 3. Formatação e Relatórios
- Todas as abas são formatadas como tabelas dinâmicas do Excel para facilitar a navegação e a análise.

### 4. Envio Automático por E-mail
- O arquivo gerado pode ser enviado automaticamente para uma lista de destinatários configurável.

---

## Requisitos de Instalação

### 1. Versão do Python
- **Python 3.8 ou superior.**

### 2. Bibliotecas Necessárias
- `pandas`
- `openpyxl`
- `smtplib`

Instale as dependências executando o comando abaixo:
```bash
pip install pandas openpyxl
```

---

## Configuração Inicial

### 1. Diretórios de Trabalho
- **Pasta de Entrada**: Onde os arquivos Excel a serem processados devem estar localizados.
  ```python
  INPUT_FOLDER = 'C:\\AutomationProject\\Input'
  ```
- **Pasta de Saída**: Onde o arquivo Excel gerado será salvo.
  ```python
  OUTPUT_PATH = 'C:\\AutomationProject\\Output\\Proposals.xlsx'
  ```
- **Arquivo de Dados**: Caminho do arquivo que contém dados de funcionários ativos.
  ```python
  DATA_PATH = 'C:\\AutomationProject\\Data\\ActiveEmployees.xlsx'
  ```

### 2. Configurações de Equipes
- Ajuste as unidades associadas a cada equipe conforme necessário:
  ```python
  team_a_units = ["UnitA", "UnitB", "UnitC", "UnitD", "UnitE"]
  team_b_units = ["UnitF", "UnitG", "UnitH", "UnitI", "UnitJ"]
  ```

### 3. Configurações de E-mail
- Configure as credenciais e destinatários:
  ```python
  sender_email = 'automation@example.com'
  sender_password = 'yourpassword'
  recipients = ['recipient1@example.com', 'recipient2@example.com']
  smtp_server = 'smtp.example.com'
  ```

---

## Como Usar

1. **Clone o Repositório**
   ```bash
   git clone https://github.com/seu-usuario/proposal-data-automation.git
   cd proposal-data-automation
   ```

2. **Execute o Script Principal**
   ```bash
   python main.py
   ```

3. **Resultados**:
   - O arquivo Excel gerado será salvo no diretório configurado no `OUTPUT_PATH`.
   - Após a geração, você será solicitado a confirmar o envio do arquivo por e-mail.

---

## Estrutura do Arquivo Gerado

- **Raw Data**: Dados brutos reordenados e com duplicatas removidas.
- **Team A**: Dados filtrados com base nas unidades associadas a `team_a_units`.
- **Team B**: Dados filtrados com base nas unidades associadas a `team_b_units`.
- **Filiais**: Abas individuais criadas para cada filial (baseadas na coluna `Unit`).
- **Consolidated Pivot**: Tabela dinâmica consolidada agrupando os dados por filial e status.

Todas as abas são formatadas como tabelas dinâmicas para otimizar a análise.

---

## Contribuições

Contribuições são sempre bem-vindas! Para relatar problemas, sugerir melhorias ou enviar pull requests, utilize a aba ["Issues"](https://github.com/seu-usuario/proposal-data-automation/issues) no repositório.

---

## Licença

Este projeto está licenciado sob a [MIT License](https://opensource.org/licenses/MIT).

---

## Autor

- **Gabriel Matuck**  
  - **E-mail**: [gabriel.matuck1@gmail.com](mailto:gabriel.matuck1@gmail.com)

