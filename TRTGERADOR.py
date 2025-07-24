import pandas as pd
import openpyxl
from datetime import datetime
import os

def colar_valor_ao_lado(ws, texto_alvo, valor, ocorrencia=1):
    cont = 0
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and texto_alvo.lower() in cell.value.lower():
                cont += 1
                if cont == ocorrencia:
                    col_destino = cell.column + 1
                    linha_destino = cell.row
                    ws.cell(row=linha_destino, column=col_destino).value = valor
                    return True
    return False
#PROCESSAR PLANILHAS O LOCAL DOS ARQUIVOS
def processar_planilhas():
    caminho_origem = r"C:\Users\wendel.ferreira\Desktop\TESTAT\RelatóriodeAudiências2025.xlsx"
    caminho_destino = r"C:\Users\wendel.ferreira\Desktop\TESTAT\TRT28.07a02.8.xlsx"

    try:
        if not os.path.exists(caminho_origem) or not os.path.exists(caminho_destino):
            print("Erro: Um dos arquivos não foi encontrado.")
            return

        df_origem = pd.read_excel(caminho_origem, sheet_name='Julho')
        mapeamento_colunas = {
            'DATA': 'DATA',
            'HORARIO': 'HORARIO', 
            'RECLAMANTE': 'RECLAMANTE',
            'RECLAMADO': 'RECLAMADO',
            'Nº DO PROCESSO': 'NUM_PROCESSO',
            'TURMA': 'TURMA',
            'RELATOR': 'RELATOR'
        }
        df_origem = df_origem.rename(columns=mapeamento_colunas)
        colunas_necessarias = ['DATA', 'HORARIO', 'RECLAMANTE', 'RECLAMADO', 'NUM_PROCESSO', 'TURMA', 'RELATOR']
        df_filtrado = df_origem[colunas_necessarias].copy()
        df_filtrado = df_filtrado.dropna(subset=['DATA', 'HORARIO'])
        df_filtrado['DATA'] = pd.to_datetime(df_filtrado['DATA'], errors='coerce')
        df_filtrado['HORARIO_TIME'] = pd.to_datetime(df_filtrado['HORARIO'], format='%H:%M', errors='coerce').dt.time
        df_filtrado = df_filtrado.sort_values(['DATA', 'HORARIO_TIME'])
        df_filtrado = df_filtrado.drop('HORARIO_TIME', axis=1)

        wb_destino = openpyxl.load_workbook(caminho_destino)
        ws_destino = wb_destino["Julho"] if "Julho" in wb_destino.sheetnames else wb_destino.active

        for i, row in df_filtrado.iterrows():
            colar_valor_ao_lado(ws_destino, "Nº processo:", str(row['NUM_PROCESSO']), ocorrencia=i+1)
            colar_valor_ao_lado(ws_destino, "Reclamante:", str(row['RECLAMANTE']), ocorrencia=i+1)
            colar_valor_ao_lado(ws_destino, "Horário:", str(row['HORARIO']), ocorrencia=i+1)
            colar_valor_ao_lado(ws_destino, "Relator:", str(row['RELATOR']), ocorrencia=i+1)
            colar_valor_ao_lado(ws_destino, "Reclamado:", str(row['RECLAMADO']), ocorrencia=i+1)

            # >>> AJUSTE FINO: separação TURMA / LOCAL
            turma_texto = str(row['TURMA']).strip()

            # Divide por espaço ou hífen
            partes = turma_texto.split()
            if len(partes) >= 2:
                numero_turma = partes[0]
                local_turma = " ".join(partes[1:])
            else:
                numero_turma = turma_texto
                local_turma = ""

            colar_valor_ao_lado(ws_destino, "Turma:", numero_turma, ocorrencia=i+1)
            colar_valor_ao_lado(ws_destino, "Local:", local_turma, ocorrencia=i+1)

        wb_destino.save(caminho_destino)
        wb_destino.close()
        print("Automação concluída com sucesso.")

    except Exception as e:
        print(f"Erro durante a execução: {e}")

if __name__ == "__main__":
    processar_planilhas()
