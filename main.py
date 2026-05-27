import pandas as pd
from pathlib import Path
import zipfile
import io

import streamlit as st
import zipfile
import tempfile
from pathlib import Path

NUMERO_REGRAS = 6
NUMERO_PARTIDAS_GRUPOS = 72
CONDICAO_ACERTOU_PLACAR = "Acertou placar"
CONDICAO_ACERTOU_VENCEDOR = "Acertou vencedor"
CONDICAO_ACERTOU_EMPATE = "Acertou empate"
MULTIPLICADOR_TERCEIRO_FINAL = "Multiplicador terceiro lugar e final"
MULTIPLICADOR_SEMIFINAIS = "Multiplicador semifinais"
MULTIPLICADOR_QUARTAS = "Multiplicador quartas"
EMPATE = "EMPATE"
regras = []
pontuacoes_participantes = []
pontos_acertou_placar = None
pontos_acertou_vencedor = None
pontos_acertou_empate = None
multiplicador_terceiro_final = None
multiplicador_semifinais = None
multiplicador_quartas = None

def exec_round(gabarito, diretorio_participantes, numero_partidas, multiplicador):
    # Ler arquivo de resultado oficial
    file = pd.read_excel(
        gabarito,
        engine="openpyxl"
    )
    # Extrair resultados oficiais
    partidas = []
    for i in range(numero_partidas):
        partidas.append({"equipe_1": file.iloc[i + 3, 3],
                         "equipe_2": file.iloc[i + 3, 4],
                         "score_1": -1 if pd.isna(file.iloc[i + 3, 5]) else file.iloc[i + 3, 5],
                         "score_2": -1 if pd.isna(file.iloc[i + 3, 6]) else file.iloc[i + 3, 6],
                         "vencedor": file.iloc[i + 3, 7]
                         })
    print("\tpartidas na rodada: ")
    for partida in partidas:
       print(f"\t{partida}")

    # Contar pontos de cada participante
    pasta = Path(diretorio_participantes)

    for arquivo in pasta.iterdir():
        if arquivo.suffix == ".xlsx":
            print(f"\t{arquivo}")
            pontuacao_total = 0
            file = pd.read_excel(
                arquivo,
                engine="openpyxl"
            )
            nome_participante = file.iloc[0, 2]
            print(f"\tContando pontos de {nome_participante}")
            for i in range(numero_partidas):
                previsao_participante = {
                    "equipe_1": file.iloc[i + 3, 3],
                    "equipe_2": file.iloc[i + 3, 4],
                    "score_1": -1 if pd.isna(file.iloc[i + 3, 5]) else file.iloc[i + 3, 5],
                    "score_2": -1 if pd.isna(file.iloc[i + 3, 6]) else file.iloc[i + 3, 6],
                    "vencedor": file.iloc[i + 3, 7]
                }
                # Valores para vencedor:
                # "?": resultados nao preenchidos corretamente
                # "EMPATE": empate
                # TIME: nome do time vencedor
                if partidas[i]["vencedor"] != "?":
                    if previsao_participante["vencedor"] != "?":
                        acertou = False
                        print(f"\tAnalisando {partidas[i]["equipe_1"]} x {partidas[i]["equipe_2"]}")

                        # Caso participante tenha acertado placar
                        if previsao_participante["score_1"] == partidas[i]["score_1"] and previsao_participante[
                            "score_2"] == partidas[i]["score_2"]:
                            pontuacao_total += pontos_acertou_placar * multiplicador
                            acertou = True
                            print("\t\tacertou placar")
                        # Caso participante tenha acertado vencedor
                        if previsao_participante["vencedor"] == partidas[i]["vencedor"] and partidas[i][
                            "vencedor"] != EMPATE:
                            pontuacao_total += pontos_acertou_vencedor * multiplicador
                            acertou = True
                            print("\t\tacertou vencedor")
                        # Caso participante tenha acertado empate
                        if previsao_participante["vencedor"] == partidas[i]["vencedor"] and partidas[i][
                            "vencedor"] == EMPATE:
                            pontuacao_total += pontos_acertou_empate * multiplicador
                            acertou = True
                            print("\t\tacertou empate")
                        if not acertou:
                            print("\t\tnao acertou")
                    else:
                        print(
                            f"\tParticipante nao preencheu {partidas[i]["equipe_1"]} x {partidas[i]["equipe_2"]} corretamente")
                else:
                    print(
                        f"\tResultado oficial de {partidas[i]["equipe_1"]} x {partidas[i]["equipe_2"]} nao preenchido corretamente")
            # Adiciona pontuacao a pontuacoes_participantes. Cria nova entrada caso nao exista
            achou_participante = False
            for participante in pontuacoes_participantes:
                if participante["nome_participante"] == nome_participante:
                    #print("Participante ja existe")
                    participante["pontuacao_total"] += pontuacao_total
                    achou_participante = True
            if not achou_participante:
                #print("Criando novo participante")
                pontuacoes_participantes.append({"nome_participante": nome_participante,
                                                "pontuacao_total": pontuacao_total})
        print()

    print("Pontuacoes apos essa rodada: ")
    print(pontuacoes_participantes)
    print()

st.title("ZIP Spreadsheet Reader")

uploaded_file = st.file_uploader(
    "Upload a ZIP file",
    type=["zip"]
)

if uploaded_file is not None:

    st.success(f"Uploaded: {uploaded_file.name}")

    with tempfile.TemporaryDirectory() as temp_dir:

        # Save uploaded ZIP
        zip_path = Path(temp_dir) / uploaded_file.name

        with open(zip_path, "wb") as f:
            f.write(uploaded_file.read())

        # Extract ZIP
        extract_path = Path(temp_dir) / "extracted"

        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(extract_path)

        # Search for xlsx files
        regras_file = None
        resultado_oficial_fase_grupos = None
        resultado_oficial_rodada_32 = None
        resultado_oficial_oitavas = None
        resultado_oficial_quartas = None
        resultado_oficial_semifinais = None
        resultado_oficial_terceiro_lugar_final = None

        # Search extracted files
        for file_obj in extract_path.rglob("*.xlsx"):

            match file_obj.name:

                case "regras_pontuacao.xlsx":
                    regras_file = file_obj

                case "resultado_oficial_fase_grupos.xlsx":
                    resultado_oficial_fase_grupos = file_obj

                case "resultado_oficial_rodada_32.xlsx":
                    resultado_oficial_rodada_32 = file_obj

                case "resultado_oficial_oitavas.xlsx":
                    resultado_oficial_oitavas = file_obj

                case "resultado_oficial_quartas.xlsx":
                    resultado_oficial_quartas = file_obj

                case "resultado_oficial_semifinais.xlsx":
                    resultado_oficial_semifinais = file_obj

                case "resultado_oficial_terceiro_lugar_final.xlsx":
                    resultado_oficial_terceiro_lugar_final = file_obj

        if regras_file is None:
            st.error("regras_pontuacao.xlsx was not found inside the ZIP.")

        else:
            st.success(f"Found file: {regras_file.name}")

            # Read spreadsheet with pandas
            df = pd.read_excel(
                regras_file,
                engine="openpyxl"
            )
            # Extrair regras
            for it in range(NUMERO_REGRAS):
                regras.append({"regra": df.iloc[it + 2, 0],
                               "pontos": df.iloc[it + 2, 1]})
            pontos_acertou_placar = None
            pontos_acertou_vencedor = None
            pontos_acertou_empate = None
            multiplicador_terceiro_final = None
            multiplicador_semifinais = None
            multiplicador_quartas = None
            for d in regras:
                if d["regra"] == CONDICAO_ACERTOU_PLACAR:
                    pontos_acertou_placar = d["pontos"]
                if d["regra"] == CONDICAO_ACERTOU_VENCEDOR:
                    pontos_acertou_vencedor = d["pontos"]
                if d["regra"] == CONDICAO_ACERTOU_EMPATE:
                    pontos_acertou_empate = d["pontos"]
                if d["regra"] == MULTIPLICADOR_TERCEIRO_FINAL:
                    multiplicador_terceiro_final = d["pontos"]
                if d["regra"] == MULTIPLICADOR_SEMIFINAIS:
                    multiplicador_semifinais = d["pontos"]
                if d["regra"] == MULTIPLICADOR_QUARTAS:
                    multiplicador_quartas = d["pontos"]
            st.write(pontos_acertou_placar)
            st.write(pontos_acertou_vencedor)
            st.write(pontos_acertou_empate)
            st.write(multiplicador_terceiro_final)
            st.write(multiplicador_semifinais)
            st.write(multiplicador_quartas)

            print("EXEC FASE DE GRUPOS")
            exec_round(resultado_oficial_fase_grupos, extract_path / "planilhas_participantes" / "fase_de_grupos",
                       NUMERO_PARTIDAS_GRUPOS, 1)








'''










print("pontuacoes participantes: ")
print(pontuacoes_participantes)
pontuacoes_participantes_ordenado = sorted(pontuacoes_participantes, key=lambda x: x["pontuacao_total"], reverse=True)
print("pontuacoes participantes ordenado: ")
print(pontuacoes_participantes_ordenado)

# Criar planilha com pontuacao final
df = pd.DataFrame(pontuacoes_participantes_ordenado)
df.to_excel(
    "pontuacao_final.xlsx",
    engine="openpyxl",
    index=False
)

print()
input("Press Enter to exit...")'''