import xmltodict
import os
import pandas as pd

def extrair_dados(nota_xml, dados_coletados):
    with open(f'nfs/{nota_xml}', "rb") as xml_file:
        xml_dict = xmltodict.parse(xml_file)

        if "NFe" in xml_dict:
            detalhes_nf = xml_dict["NFe"]['infNFe']
        else:
            detalhes_nf = xml_dict['nfeProc']["NFe"]['infNFe']

        numero = detalhes_nf["@Id"]
        empresa = detalhes_nf['emit']['xNome']
        cliente = detalhes_nf["dest"]["xNome"]
        endereco_cliente = detalhes_nf["dest"]["enderDest"]

        peso_total = detalhes_nf["transp"].get("vol", {}).get("pesoB", "Não informado")

        dados_coletados.append([numero, empresa, cliente, endereco_cliente, peso_total])

diretorio_saida = "Notas_Processadas"
os.makedirs(diretorio_saida, exist_ok=True)

arquivos_xml = os.listdir("nfs")

colunas_tabela = ["ID Nota", "Empresa Emissora", "Nome Cliente", "Endereço", "Peso (Kg)"]
dados_extraidos = []

for xml in arquivos_xml:
    extrair_dados(xml, dados_extraidos)

df_notas = pd.DataFrame(columns=colunas_tabela, data=dados_extraidos)
arquivo_saida = os.path.join(diretorio_saida, "NotasFiscais.xlsx")
df_notas.to_excel(arquivo_saida, index=False)
print("Processo concluído!")