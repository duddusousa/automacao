import os
from openpyxl import Workbook
import pdfplumber
import re

diretorio = 'pdfs'
arquivos = os.listdir(diretorio)

if len(arquivos) == 0:
    raise Exception("Nenhum arquivo encontrado no diretório")

wb = Workbook()
ws = wb.active
ws.title = 'Arquivo Teste'

# Cabeçalho
ws['A1'] = 'DATA'
ws['F1'] = 'ENTRADA'
ws['G1'] = 'SAIDA'

ultima_linha = 2  # começa depois do cabeçalho

for file in arquivos:
    with pdfplumber.open(f"{diretorio}/{file}") as pdf:
        for pagina in pdf.pages:

            texto_pdf = pagina.extract_text()

            padrao_data = r"\d{2}\s[A-Z]{3}\s\d{4}"
            padrao_valores = r"(Total de entradas|Total de saídas)\s([+-]\s?\d{1,3}(?:\.\d{3})*,\d{2})"

            datas = re.findall(padrao_data, texto_pdf)
            valores = re.findall(padrao_valores, texto_pdf)

            for i in range(len(valores)):
                tipo, valor_texto = valores[i]
                data = datas[i] if i < len(datas) else ""

                # limpar valor
                valor_limpo = valor_texto.replace("+", "").replace("-", "").replace(".", "").replace(",", ".")
                valor = float(valor_limpo)

                ws[f"A{ultima_linha}"] = data

                if "entradas" in tipo:
                    ws[f"F{ultima_linha}"] = valor
                else:
                    ws[f"G{ultima_linha}"] = valor

                ultima_linha += 1

wb.save("resultado.xlsx")
print("Excel gerado com sucesso 🔥")
