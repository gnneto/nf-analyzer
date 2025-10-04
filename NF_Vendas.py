import xml.etree.ElementTree as ET
import pandas as pd
import os
from datetime import datetime

pasta = 'nfs'
arquivo_excel = 'notas.xlsx'
dados = []

def formata_data(data_xml: str) -> str:
    if not data_xml:
        return ""
    try:
        if "T" in data_xml:
            dt = datetime.fromisoformat(data_xml.split("-")[0] + "-" + data_xml.split("-")[1] + "-" + data_xml.split("-")[2][:2])
        else:
            dt = datetime.strptime(data_xml, "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return data_xml 

def formata_valor(valor_xml: str) -> str:
    if not valor_xml:
        return ""
    try:
        return f"{float(valor_xml):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return valor_xml

# lê todos os xmls da pasta
for arquivo in os.listdir(pasta):
    if arquivo.endswith(".xml"):
        caminho = os.path.join(pasta, arquivo)

        tree = ET.parse(caminho)
        root = tree.getroot()

        # trata namespace
        ns = {'ns': root.tag.split("}")[0].strip("{") if "}" in root.tag else {}}

        # campos principais
        nNF = root.find(".//ns:ide/ns:nNF", ns).text 
        data_emissao = root.find(".//ns:ide/ns:dhEmi", ns).text if root.find(".//ns:ide/ns:dhEmi", ns) is not None else ""
        emitente = root.find(".//ns:emit/ns:xNome", ns).text if root.find(".//ns:emit/ns:xNome", ns) is not None else ""
        destinatario = root.find(".//ns:dest/ns:xNome", ns).text if root.find(".//ns:dest/ns:xNome", ns) is not None else ""
        valor_total = root.find(".//ns:ICMSTot/ns:vNF", ns).text if root.find(".//ns:ICMSTot/ns:vNF", ns) is not None else ""
        form_pgto = root.find(".//ns:pag/ns:tPag", ns).text if root.find(".//ns:pag/ns:tPag", ns) is not None else ""

        # trata duplicatas
        duplicatas = root.findall(".//ns:dup", ns)
        if duplicatas:
            for dup in duplicatas:
                numero_dup = dup.find("ns:nDup", ns).text if dup.find("ns:nDup", ns) is not None else ""
                vencimento = dup.find("ns:dVenc", ns).text if dup.find("ns:dVenc", ns) is not None else ""
                valor_dup = dup.find("ns:vDup", ns).text if dup.find("ns:vDup", ns) is not None else ""

                dados.append({
                    "Numero NF": nNF,
                    "Data Emissao": formata_data(data_emissao),
                    "Destinatario": destinatario,
                    "Parcela": numero_dup,
                    "Vencimento": formata_data(vencimento),
                    "Valor parcela": formata_valor(valor_dup),
                    "Valor total NF": formata_valor(valor_total),
                    "Forma pgto": form_pgto,
                    "Emitente": emitente,
                })
        else:
            dados.append({
                "Numero NF": nNF,
                "Data Emissao": formata_data(data_emissao),
                "Destinatario": destinatario,
                "Parcela": "",
                "Vencimento": "",
                "Valor parcela": "",
                "Valor total NF": formata_valor(valor_total),
                "Forma pgto": form_pgto,
                "Emitente": emitente,
            })

# cria DataFrame com novos dados
df_novos = pd.DataFrame(dados)

if os.path.exists(arquivo_excel):
    df_existente = pd.read_excel(arquivo_excel, engine="openpyxl")
    df_final = pd.concat([df_existente, df_novos], ignore_index=True)
else:
    df_final = df_novos

# evita duplicadas (mesma NF e parcela)
df_final.drop_duplicates(subset=["Numero NF", "Parcela"], inplace=True)

# conversões de tipo numérico
def limpar_valores(coluna):
    return (
        coluna.astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .replace("", None)
        .astype(float)
    )

# trata duplicadas
df_final["Parcela"] = pd.to_numeric(df_final["Parcela"], errors="coerce")
df_final["Valor parcela"] = limpar_valores(df_final["Valor parcela"])
df_final["Valor total NF"] = limpar_valores(df_final["Valor total NF"])

df_final["Numero NF"] = pd.to_numeric(df_final["Numero NF"], errors="coerce").astype("Int64")

# exporta pro Excel
df_final.to_excel(arquivo_excel, index=False, engine="openpyxl")

print("Dados adicionados e salvos com sucesso!")
