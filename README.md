# Análise Financeira de Notas Fiscais Eletrônicas (NF-e)

Este projeto realiza a leitura e análise de arquivos XML de Notas Fiscais Eletrônicas (NF-e), com foco na extração de informações financeiras, como vencimentos e valores, para uma análise mais detalhada e eficiente.

## Funcionalidades

- Leitura de arquivos XML de NF-e em uma pasta específica.
- Extração de informações financeiras, como parcelas, vencimentos e valores.
- Consolidação dos dados em um arquivo Excel (`notas.xlsx`) para facilitar a análise.
- Formatação de datas e valores monetários no padrão brasileiro.
- Evita duplicação de registros com base no número da nota e parcela.

## Como usar

1. **Pré-requisitos**:
   - Python 3.7 ou superior.
   - Bibliotecas necessárias: `pandas`, `openpyxl`.

   Para instalar as dependências, execute:
   ```bash
   pip install pandas openpyxl
   ```

2. **Estrutura do projeto**:
   - Certifique-se de que os arquivos XML das NF-e estão em uma pasta chamada `nfs` no mesmo diretório do script.
   - O script gera ou atualiza o arquivo `notas.xlsx` no mesmo diretório.

3. **Execução**:
   - Execute o script `NF_Vendas.py`:
     ```bash
     python NF_Vendas.py
     ```
   - O arquivo `notas.xlsx` será gerado ou atualizado com os dados extraídos.

## Observações

- **Foco financeiro**: O projeto prioriza a análise de vencimentos e valores das parcelas, permitindo uma visão detalhada das obrigações financeiras.
- **Namespace**: O script trata namespaces nos arquivos XML (NFe) automaticamente.
- **Formatação**: Datas são formatadas no padrão `DD/MM/AAAA` e valores monetários no formato brasileiro `1.234,56`.

## Estrutura do Arquivo Excel

O arquivo gerado (`notas.xlsx`) contém as seguintes colunas:

- **Numero NF**: Número da nota fiscal.
- **Data Emissao**: Data de emissão da nota.
- **Destinatario**: Nome do destinatário.
- **Parcela**: Número da parcela (se aplicável).
- **Vencimento**: Data de vencimento da parcela (se aplicável).
- **Valor parcela**: Valor da parcela (se aplicável).
- **Valor total NF**: Valor total da nota fiscal.
- **Forma pgto**: Forma de pagamento.
- **Emitente**: Nome do emitente.