import os
import pandas as pd
import pdfplumber
import easyocr
import numpy as np
import re
from PIL import Image, ImageEnhance
from datetime import datetime
import time

# --- CONFIGURAÇÃO ---
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
pasta_input = os.path.join(BASE_DIR, "DATA", "input")
pasta_output = os.path.join(BASE_DIR, "DATA", "output")
caminho_banco = os.path.join(pasta_output, "Banco_Mestre_Brasul.xlsx")

regex_cod = re.compile(r'\d{1,2}[\.\, ]\d{2}[\.\, ]\d{2,3}')
unidades_lista = ['KG', 'M2', 'M3', 'UN', 'M', '%', 'CJ', 'PA', 'VB', 'M1', 'H', 'MES', 'VB']


def limpar_valor(t):
    res = re.sub(r'[^0-9\,\.]', '', t.replace('O', '0').replace('I', '1').replace('L', '1'))
    return res if res else "0"


def extrair_total_brasul():
    tempo_total = time.time()
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🚀 INICIANDO EXTRATOR DE DADOS BRASUL v1.0")

    if not os.path.exists(pasta_output): os.makedirs(pasta_output)

    # CARREGA A LISTA DE ARQUIVOS QUE JA FORAM ENVIADOS PARA A TRAVA DE SEGURANÇA
    arquivos_processados = []
    if os.path.exists(caminho_banco):
        try:
            df_existente = pd.read_excel(caminho_banco)
            arquivos_processados = df_existente['Obra_Arq'].unique().tolist()
            print(f"BANCO CARREGADO. {len(arquivos_processados)} ARQUIVOS JÁ ESTÃO NO COFRE.")
        except:
            print("ERRO AO LER BANCO EXISTENTE, PROCESSANDO COMO NOVO.")

    reader = easyocr.Reader(['pt'])
    arquivos_na_pasta = [f for f in os.listdir(pasta_input) if f.endswith('.pdf')]

    for i, arquivo in enumerate(arquivos_na_pasta):
        t_arq = time.time()

        # TRAVA DE SEGURANÇA: SE JA ESTIVER NO EXCEL O ARQUIVO SALVO, ESSE COMANDO FAZ ELE PULAR
        if arquivo in arquivos_processados:
            print(f"({i + 1}/{len(arquivos_na_pasta)}) PULANDO: {arquivo} (JÁ EXISTE NO SISTEMA)")
            continue

        print(f"({i + 1}/{len(arquivos_na_pasta)}) ANALISANDO: {arquivo}")
        dados_arquivo = []
        nome_obra = arquivo.replace('.pdf', '')

        try:
            with pdfplumber.open(os.path.join(pasta_input, arquivo)) as pdf:
                texto_capa = pdf.pages[0].extract_text()
                if texto_capa:
                    m = re.search(r'ESCOLA\s*[:\-]\s*(.*)', texto_capa, re.IGNORECASE)
                    if m: nome_obra = m.group(1).split('\n')[0].strip().upper()

                for n_pag, pagina in enumerate(pdf.pages):
                    img_raw = pagina.to_image(resolution=500).original
                    img = ImageEnhance.Contrast(img_raw.convert("L")).enhance(3.5)
                    resultado = reader.readtext(np.array(img.convert("RGB")))

                    texto_pag = " ".join([it[1].upper() for it in resultado])
                    tipo_p = "ACUMULADO" if any(
                        k in texto_pag for k in ["ACUMULADO", "MEDIÇÃO", "ANTERIOR", "PERIODO"]) else "QUANTITATIVA"

                    linhas_y = {}
                    for (bbox, texto, prob) in resultado:
                        if prob < 0.05: continue
                        y = (bbox[0][1] + bbox[2][1]) / 2
                        achou = False
                        for k in linhas_y.keys():
                            if abs(y - k) < 25:
                                linhas_y[k].append((bbox[0][0], texto))
                                achou = True
                                break
                        if not achou: linhas_y[y] = [(bbox[0][0], texto)]

                    for y in sorted(linhas_y.keys()):
                        itens = sorted(linhas_y[y], key=lambda x: x[0])
                        cod, un, desc_parts, vals = "", "", [], []

                        for j, (x, txt) in enumerate(itens):
                            t = txt.strip()
                            if regex_cod.search(t) and not cod:
                                cod = t.replace(',', '.')
                            elif t.upper() in unidades_lista:
                                un = t.upper()
                            elif any(c.isdigit() for c in t) and j > 1:
                                v = limpar_valor(t)
                                if v != "0": vals.append(v)
                            else:
                                if len(t) > 1: desc_parts.append(t)

                        desc_f = " ".join(desc_parts).upper().strip()

                        if cod or len(desc_f) > 3:
                            dados_arquivo.append({
                                'Obra': nome_obra, 'Obra_Arq': arquivo, 'Tipo': tipo_p,
                                'Cod': cod, 'Desc': desc_f, 'UN': un,
                                'Q_Orc': vals[0] if len(vals) > 0 else "0",
                                'Q_Acum': vals[1] if (len(vals) > 1 and tipo_p == "ACUMULADO") else "0"
                            })

            # SALVA NO EXCEL APÓS CADA ARQUIVO PROCESSADO (CHAVE DE SEGURANÇA CASO CAIA O SISTEMA OU A MAQUINA DESLIGUE)
            if dados_arquivo:
                df_n = pd.DataFrame(dados_arquivo)
                if os.path.exists(caminho_banco):
                    df_f = pd.concat([pd.read_excel(caminho_banco), df_n], ignore_index=True)
                else:
                    df_f = df_n
                cols = ['Obra', 'Obra_Arq', 'Tipo', 'Cod', 'Desc', 'UN', 'Q_Orc', 'Q_Acum']
                df_f[cols].to_excel(caminho_banco, index=False)
                print(f"{arquivo} SALVO COM SUCESSO NO COFRE.")

        except Exception as e:
            print(f"ERRO AO PROCESSAR {arquivo}: {e}")

    print(f"\n🏁 FINALIZADO! TEMPO TOTAL: {(time.time() - tempo_total) / 3600:.2f} HORAS.")


if __name__ == "__main__":
    extrair_total_brasul()