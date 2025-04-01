# Sistema de Gestão Financeira - Ambiente de Testes

## Preparação do Ambiente

1. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```

2. Estrutura do Google Drive necessária:
   - Pasta "planilhas_base":
     - MODELO.xlsx
     - clientes.xlsx
     - fornecedores.xlsx
   - Pasta "clientes" (para arquivos individuais)

3. Gere o executável:
   ```
   pyinstaller sistema_principal.spec
   ```

4. Após a geração, copie para a pasta dist:
   - Pasta "planilhas_base" com os arquivos base
   - Pasta "clientes" vazia (será preenchida pelo sistema)

## Executando os Testes

1. Execute o arquivo Sistema_Gestao_Financeira.exe
2. Use os dados de teste fornecidos
3. Verifique os logs gerados em logs/sistema.log

## Observações

- Os arquivos na pasta planilhas_base são essenciais
- Mantenha backup dos dados de teste
- Logs são gerados na pasta logs/
