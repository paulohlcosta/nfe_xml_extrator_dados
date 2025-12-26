import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from datetime import datetime
import re

# Define o caminho da pasta onde estão os arquivos XML
pasta_xml = 'D:\\xml_vtin'

# Gera o timestamp no formato AAAAmmdd_HHMM
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
# Define o nome do arquivo de saída com o timestamp
nome_arquivo_xlsx = f'dados_nfe_{timestamp}.xlsx'

# Define os namespaces para as tags XML da NFE
namespaces = {
    'nfe': 'http://www.portalfiscal.inf.br/nfe'
}

# Cria um novo arquivo Excel
workbook = Workbook()
# Seleciona a folha ativa
sheet = workbook.active
# Define o cabeçalho da folha
sheet.title = "Dados NFE"
# Alteração na ordem das colunas do cabeçalho
cabecalho = ['data_hora_emissao', 'numero_nfe', 'cfop', 'natOp', 'ncm', 'descricao', 'quantidade', 'valor', 'estado_emitente', 'municipio_emitente', 'numero_pedido', 'numero_nfe_id']
sheet.append(cabecalho) # Adiciona o cabeçalho na primeira linha

# Itera sobre todos os arquivos na pasta
for arquivo in os.listdir(pasta_xml):
    # Verifica se o arquivo termina com .xml
    if arquivo.endswith('.xml'):
        caminho_completo = os.path.join(pasta_xml, arquivo)
        try:
            # Faz o parsing do arquivo XML
            tree = ET.parse(caminho_completo)
            root = tree.getroot()
            
            # Extrai os dados da NFe que são comuns a todos os itens
            numero_nfe_id = root.find('.//nfe:infNFe', namespaces).attrib['Id'][3:]
            numero_nfe = root.find('.//nfe:ide/nfe:nNF', namespaces).text if root.find('.//nfe:ide/nfe:nNF', namespaces) is not None else ''
            natOp = root.find('.//nfe:ide/nfe:natOp', namespaces).text if root.find('.//nfe:ide/nfe:natOp', namespaces) is not None else ''
            data_hora_emissao = root.find('.//nfe:ide/nfe:dhEmi', namespaces).text if root.find('.//nfe:ide/nfe:dhEmi', namespaces) is not None else ''
            
            # Extrai os dados do emitente
            estado_emitente = root.find('.//nfe:enderEmit/nfe:UF', namespaces).text if root.find('.//nfe:enderEmit/nfe:UF', namespaces) is not None else ''
            municipio_emitente = root.find('.//nfe:enderEmit/nfe:xMun', namespaces).text if root.find('.//nfe:enderEmit/nfe:xMun', namespaces) is not None else ''

            # Extrai o texto do campo de informações complementares, se existir
            infCpl_text = root.find('.//nfe:infCpl', namespaces).text if root.find('.//nfe:infCpl', namespaces) is not None else ''

            # Procura por um padrão de 10 dígitos que começa com '4500' no texto
            match = re.search(r'4500\d{6}', infCpl_text)
            if match:
                numero_pedido = match.group(0)
            else:
                numero_pedido = ''

            # Encontra todos os itens (<det>) dentro da NFE
            itens = root.findall('.//nfe:det', namespaces)

            # Percorre cada item e extrai as informações
            for item in itens:
                # Extrai os dados específicos de cada item
                cfop = item.find('.//nfe:CFOP', namespaces).text if item.find('.//nfe:CFOP', namespaces) is not None else ''
                ncm = item.find('.//nfe:NCM', namespaces).text if item.find('.//nfe:NCM', namespaces) is not None else ''
                descricao = item.find('.//nfe:xProd', namespaces).text if item.find('.//nfe:xProd', namespaces) is not None else ''

                # Extrai a quantidade e o valor, convertendo para float
                quantidade = item.find('.//nfe:qCom', namespaces).text
                if quantidade:
                    quantidade = float(quantidade)
                else:
                    quantidade = 0.0
                
                valor = item.find('.//nfe:vProd', namespaces).text
                if valor:
                    valor = float(valor)
                else:
                    valor = 0.0
                
                # Adiciona a linha de dados na planilha, respeitando a nova ordem do cabeçalho
                sheet.append([
                    data_hora_emissao,
                    numero_nfe,
                    cfop,
                    natOp,
                    ncm,
                    descricao,
                    quantidade,
                    valor,
                    estado_emitente,
                    municipio_emitente,
                    numero_pedido,
                    numero_nfe_id
                ])
        
        except ET.ParseError as e:
            print(f"Erro ao processar o arquivo {arquivo}: {e}")
        except Exception as e:
            print(f"Ocorreu um erro inesperado ao processar o arquivo {arquivo}: {e}")

# Salva o arquivo Excel
workbook.save(nome_arquivo_xlsx)
print(f"Processamento concluído. Os dados foram salvos no arquivo '{nome_arquivo_xlsx}'.")