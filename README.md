# NFe XML Data Extractor to Excel

Este script em Python automatiza a extração de informações críticas de arquivos **XML de Notas Fiscais Eletrônicas (NF-e)**, consolidando-as em uma planilha Excel (.xlsx) organizada por itens.

O diferencial deste extrator é a capacidade de percorrer todos os itens (`det`) de cada nota, além de utilizar expressões regulares para localizar números de pedidos de compra dentro do campo de informações complementares.

## 🚀 Funcionalidades

* **Processamento em Lote:** Lê todos os arquivos `.xml` de uma pasta específica.
* **Extração Detalhada por Item:** Se uma nota possui 10 itens, o script gera 10 linhas correspondentes, mantendo os dados do cabeçalho da nota em cada uma.
* **Inteligência via Regex:** Identifica automaticamente números de pedidos (padrão iniciado em `4500`) dentro das informações complementares (`infCpl`).
* **Output Organizado:** Gera um arquivo Excel com timestamp no nome (`dados_nfe_AAAAMMDD_HHMM.xlsx`) para evitar que dados antigos sejam sobrescritos.

## 📊 Dados Extraídos

O script organiza a planilha com as seguintes colunas:

1.  **data_hora_emissao**: Data e hora de emissão da nota.
2.  **numero_nfe**: Número do documento fiscal.
3.  **cfop**: Código Fiscal de Operações e Prestações do item.
4.  **natOp**: Natureza da operação.
5.  **ncm**: Nomenclatura Comum do Mercosul do item.
6.  **descricao**: Descrição completa do produto.
7.  **quantidade**: Quantidade comercializada.
8.  **valor**: Valor total do item (bruto).
9.  **estado_emitente**: UF do emissor.
10. **municipio_emitente**: Nome da cidade do emissor.
11. **numero_pedido**: Número do pedido capturado via Regex.
12. **numero_nfe_id**: Chave de acesso da nota (removendo o prefixo 'NFe').

## 🛠️ Configuração

Antes de executar, ajuste o caminho da pasta onde seus arquivos XML estão armazenados no script:

```python
pasta_xml = 'D:\\seu_caminho_aqui'
```

## 📝 Como usar

   * Coloque todos os arquivos XML que deseja processar na pasta configurada.
   * Execute o script Python.
   * Ao finalizar, o console exibirá o nome do arquivo Excel gerado no diretório raiz do script.
   * O arquivo estará pronto para análise, filtros e criação de tabelas dinâmicas.

## 🔍 Tratamento de Erros

    O script ignora arquivos que não possuem extensão .xml.
    Caso um arquivo esteja corrompido ou fora do padrão do Portal da NF-e, o script exibirá um erro de ParseError no console, mas continuará processando os demais arquivos da pasta.
