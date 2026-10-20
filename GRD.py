# Criar Site
# Manter Site no Ar: streamlit run GRD.py
# Instalar Biblioteca: pip install  xxxxxx

import streamlit as st
import pandas as pd
import plotly.express as px
import os
from openpyxl.styles import Font
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side

# ==========================================
# CARREGAR GRDs COM CACHE
# ==========================================

@st.cache_data
def carregar_grds():

    wb = load_workbook(
        "BASE DE DADOS GRDs CADASTRADAS.xlsx",
        data_only=True
    )

    lista_df = []

    for aba in wb.sheetnames:

        ws = wb[aba]

        dados = list(ws.values)

        if len(dados) > 1:

            colunas = dados[0]

            df_temp = pd.DataFrame(
                dados[1:],
                columns=colunas
            )

            df_temp["Ano"] = aba


            # Cálculo das colunas V, W, X e Y

            hoje = pd.Timestamp.now().normalize()

            if "Data de Entrada" in df_temp.columns:

                df_temp["Data de Entrada"] = pd.to_datetime(
                    df_temp["Data de Entrada"],
                    dayfirst=True,
                    errors="coerce"
                )

                # V = Data Devolução
                df_temp["Data Devolução"] = df_temp["Data de Entrada"].apply(
                    lambda x: x + pd.Timedelta(days=30)
                    if pd.notnull(x)
                    else pd.NaT
                )

                # Converter Data de Saída
                if "Data de Saída" in df_temp.columns:
                    df_temp["Data de Saída"] = pd.to_datetime(
                        df_temp["Data de Saída"],
                        dayfirst=True,
                        errors="coerce"
                    )

                # X = Dias p/ Vencimento
                df_temp["Dias p/ Vencimento GRD"] = df_temp.apply(
                    lambda row:
                    ""
                    if pd.notnull(row.get("Data de Saída"))
                    else (
                        (row["Data Devolução"] - hoje).days
                        if pd.notnull(row["Data Devolução"])
                        and (row["Data Devolução"] - hoje).days >= 0
                        else ""
                    ),
                    axis=1
                )

                # Y = Dias Fora do Prazo
                df_temp["Dias Fora do Prazo GRD"] = df_temp.apply(
                    lambda row:
                    ""
                    if pd.notnull(row.get("Data de Saída"))
                    else (
                        (hoje - row["Data Devolução"]).days
                        if pd.notnull(row["Data Devolução"])
                        and (hoje - row["Data Devolução"]).days > 0
                        else ""
                    ),
                    axis=1
                )

                # W = Status GRD

                df_temp["Status GRD"] = df_temp.apply(
                    lambda row:
                    (
                        "NO PRAZO"
                        if pd.notnull(row.get("Data de Saída"))
                        and pd.notnull(row.get("Data Devolução"))
                        and row["Data de Saída"] <= row["Data Devolução"]
                        else (
                            "FORA DO PRAZO"
                            if pd.notnull(row.get("Data de Saída"))
                            and pd.notnull(row.get("Data Devolução"))
                            and row["Data de Saída"] > row["Data Devolução"]
                            else (
                                "NO PRAZO"
                                if pd.isnull(row.get("Data de Saída"))
                                and pd.notnull(row.get("Data Devolução"))
                                and hoje <= row["Data Devolução"]
                                else (
                                    "FORA DO PRAZO"
                                    if pd.isnull(row.get("Data de Saída"))
                                    and pd.notnull(row.get("Data Devolução"))
                                    and hoje > row["Data Devolução"]
                                    else ""
                                )
                            )
                        )
                    ),
                    axis=1
                )
           
            # Remove linhas vazias
            df_temp = df_temp.dropna(how="all")

            # Mantém somente registros de GRD
            if "Número GRD" in df_temp.columns:

                df_temp = df_temp[
                    df_temp["Número GRD"]
                    .astype(str)
                    .str.startswith("GRD_", na=False)
                ]

            lista_df.append(df_temp)

    if lista_df:

        return pd.concat(
            lista_df,
            ignore_index=True
        )

    return pd.DataFrame()

# Estilo de Borda - Base de Dados GRDs Cadastradas
borda_fina = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)


# Título: TRAMITAÇÃO DE DOCUMENTAÇÃO TÉCNICA
#                GERÊNCIA - EI/ET
#         GRD - GUIA DE REMESSA DE DOCUMENTOS 
st.set_page_config(layout="wide")
st.markdown("""
<style>
.block-container {
padding-top: 2.2rem;
}
</style>
""", unsafe_allow_html=True)
st.markdown("""
<div style="text-align:center;">

<div style="font-size:30px; font-weight:bold;">
TRAMITAÇÃO DE DOCUMENTAÇÃO TÉCNICA
</div>

<div style="font-size:24px; font-weight:bold; margin-top:-5px;">
GERÊNCIA - EI/ET
</div>

<hr style="width:60%; margin:5px auto;">

<div style="font-size:22px; font-weight:bold; margin-top:-5px;">
GRD - GUIA DE REMESSA DE DOCUMENTOS
</div>

</div>
""", unsafe_allow_html=True)

if "reset_form" not in st.session_state:
   st.session_state.reset_form = 0


# Botão Nova GRD - Apagar campos
if st.sidebar.button("🆕 Nova GRD", type="primary"):

    st.session_state.reset_form += 1


    st.rerun()

#Título
st.sidebar.write('## Cadastrar GRD')

# Ler Excel

def carregar_base_interna():
    return pd.read_excel("BASE INTERNA.xlsx")

df = carregar_base_interna()

# Filtrar Subestações
lista_subestacoes = [""] + (
df["SUBESTACAO"]
.dropna()
.astype(str)
.tolist()
)
subestacao = st.sidebar.selectbox(
"Subestação",
lista_subestacoes,
key=f"subestacao_{st.session_state.reset_form}"
)

# Filtrar Contratadas
lista_contratadas = [""] + (
df["CONTRATADA"]
.dropna()
.astype(str)
.tolist()
)
contratada = st.sidebar.selectbox(
"Contratada",
lista_contratadas,
key=f"contratada_{st.session_state.reset_form}"
)

# Filtrar contratos conforme a contratada
if contratada == "BRASIL REAL":
    contratos =(df["BRASIL REAL_CONTRATOS"]
    .dropna()
    .astype(int)
    .astype(str)
    .tolist())

elif contratada == "CTRN":
     contratos = (df["CTRN_CONTRATOS"]
     .dropna()
     .astype(int)
     .astype(str)
     .tolist())

elif contratada == "CYMI":
        contratos = (df["CYMI_CONTRATOS"]
        .dropna()
        .astype(int)
        .astype(str)
        .tolist())
    
elif contratada == "MPE":
    contratos = (df["MPE_CONTRATOS"]
    .dropna()
    .astype(int)
    .astype(str)
    .tolist())
    
elif contratada == "TECNOVA":
    contratos = (df["TECNOVA_CONTRATOS"]
    .dropna()
    .astype(int)
    .astype(str)
    .tolist())
    
elif contratada == "VISION":
    contratos = (df["VISION_CONTRATOS"]
     .dropna()
    .astype(int)
    .astype(str)
    .tolist())
    
elif contratada == "WEG":
    contratos = (df["WEG_CONTRATOS"]
    .dropna()
    .astype(int)
    .astype(str)
    .tolist())
    
else:
    contratos = []

# Selectbox dos contratos
lista_contratos = st.sidebar.selectbox(
"Contrato",
[""] + contratos,
key=f"contrato_{st.session_state.reset_form}"
)

# Filtrar empreendimentos conforme a contratada
if contratada == "BRASIL REAL":
    empreendimentos =(df["BRASIL REAL_EMPREENDIMENTOS"]
    .dropna()
    .astype(str)
    .tolist())

elif contratada == "CTRN":
     empreendimentos = (df["CTRN_EMPREENDIMENTOS"]
     .dropna()
     .astype(str)
     .tolist())

elif contratada == "CYMI":
        empreendimentos = (df["CYMI_EMPREENDIMENTOS"]
        .dropna()
        .astype(str)
        .tolist())

elif contratada == "MPE":
    empreendimentos = (df["MPE_EMPREENDIMENTOS"]
    .dropna()
    .astype(str)
    .tolist())
    
elif contratada == "TECNOVA":
    empreendimentos = (df["TECNOVA_EMPREENDIMENTOS"]
    .dropna()
    .astype(str)
    .tolist())

elif contratada == "VISION":
    empreendimentos = (df["VISION_EMPREENDIMENTOS"]
     .dropna()
    .astype(str)
    .tolist())

elif contratada == "WEG":
    empreendimentos = (df["WEG_EMPREENDIMENTOS"]
    .dropna()
    .astype(str)
    .tolist())

else:
    empreendimentos = []

# Selectbox dos empreendimentos
lista_empreendimento = st.sidebar.selectbox(
"Empreendimento",
[""] + empreendimentos,
key=f"empreendimento_{st.session_state.reset_form}"
)

# Filtrar responsável conforme a contratada
if contratada == "BRASIL REAL":
    responsável_contratadas = (df["BRASIL REAL_RESPONSAVEIS"]
    .dropna()
    .astype(str)
    .tolist())

elif contratada == "CTRN":
     responsável_contratadas = (df["CTRN_RESPONSAVEIS"]
     .dropna()
     .astype(str)
     .tolist())

elif contratada == "CYMI":
        responsável_contratadas = (df["CYMI_RESPONSAVEIS"]
        .dropna()
        .astype(str)
        .tolist())

elif contratada == "MPE":
    responsável_contratadas = (df["MPE_RESPONSAVEIS"]
    .dropna()
    .astype(str)
    .tolist())
    
elif contratada == "TECNOVA":
    responsável_contratadas = (df["TECNOVA_RESPONSAVEIS"]
    .dropna()
    .astype(str)
    .tolist())

elif contratada == "VISION":
    responsável_contratadas = (df["VISION_RESPONSAVEIS"]
     .dropna()
    .astype(str)
    .tolist())

elif contratada == "WEG":
    responsável_contratadas = (df["WEG_RESPONSAVEIS"]
    .dropna()
    .astype(str)
    .tolist())

else:
    responsável_contratadas = []

# Selectbox dos responsáveis contratadas
lista_responsavel_contratadas = st.sidebar.selectbox(
"Responsável Contratadas",
[""] + responsável_contratadas,
key=f"responsavel_contratada_{st.session_state.reset_form}"
)

# Filtrar Agrupamentos
lista_agrupamentos = [""] + (
df["AGRUPAMENTO"]
.dropna()
.astype(str)
.tolist()
)
agrupamento = st.sidebar.selectbox(
"Agrupamento",
lista_agrupamentos,
key=f"agrupamento_{st.session_state.reset_form}"
)

# Filtrar subagrupamentos conforme agrupamento
if agrupamento == "EQUIPAMENTOS - EQP":
    subagrupamentos = (df["EQP_SUBAGRUPAMENTO"]
    .dropna()
    .astype(str)
    .tolist())

elif agrupamento == "SUBESTAÇÕES - SE":
     subagrupamentos = (df["SE_SUBAGRUPAMENTO"]
     .dropna()
     .astype(str)
     .tolist())

elif agrupamento == "LINHAS DE TRANSMISSÃO - LT":
     subagrupamentos = (df["LT_SUBAGRUPAMENTO"]
     .dropna()
     .astype(str)
     .tolist())
     
else:
    subagrupamentos = []

    # Selectbox dos subagrupamentos
lista_subagrupamentos = st.sidebar.selectbox(
"Subagrupamento",
[""] + subagrupamentos,
key=f"subagrupamento_{st.session_state.reset_form}"
)

# Filtrar responsáveis CEMIG conforme agrupamento
if agrupamento == "EQUIPAMENTOS - EQP":
    responsável_cemig = (df["EQP_RESPONSAVEIS"]
    .dropna()
    .astype(str)
    .tolist())

elif agrupamento == "SUBESTAÇÕES - SE":
     responsável_cemig = (df["SE_RESPONSAVEIS"]
     .dropna()
     .astype(str)
     .tolist())

elif agrupamento == "LINHAS DE TRANSMISSÃO - LT":
     responsável_cemig = (df["LT_RESPONSAVEIS"]
     .dropna()
     .astype(str)
     .tolist())
     
else:
    responsável_cemig = []

    # Selectbox dos subagrupamentos
lista_responsavel_cemig = st.sidebar.selectbox(
"Responsável CEMIG - Recebimento",
[""] + responsável_cemig,
key=f"responsavel_cemig_{st.session_state.reset_form}"
)

    # Identificação Contratada
identificação_contratada  = st.sidebar.text_input("Identificação Contratada",
                                                  key=f"identificacao_contratada_{st.session_state.reset_form}")

# Data de Entrada
data = st.sidebar.date_input('Data de Entrada',value=None,
format="DD/MM/YYYY",
key=f"data_entrada_{st.session_state.reset_form}")

# Gerar Número da GRD
sigla_agrupamento = agrupamento.split(" - ")[-1]
sigla_subagrupamento = lista_subagrupamentos.split(" - ")[-1]

campos_preenchidos = all([
subestacao,
contratada,
lista_contratos,
lista_empreendimento,
lista_responsavel_contratadas,
agrupamento,
lista_subagrupamentos,
lista_responsavel_cemig,
identificação_contratada,
data
])

if campos_preenchidos:

    numero_grd = (
    f"GRD_"
    f"{sigla_agrupamento}_"
    f"{sigla_subagrupamento}_"
    f"{identificação_contratada}_"
    f"{contratada}_"
    f"{lista_contratos}_"
    f"{subestacao}_"
    f"{data.strftime('%d%m%y')}"
    )
    
    st.markdown(
    f"""
    <div style="
    text-align:center;
    margin-top:5px;
    padding:2.5px;
    border:0.3px solid #003087;
    border-radius:2.5px;
    background-color:#f8f9fa;
    ">
    <h5 style="color:#003087;">
    NÚMERO DE IDENTIFICAÇÃO DA GRD
    </h5>
    <h5 style="color:#d00000;">
    {numero_grd}
    </h6>
    </div>
    """,
    unsafe_allow_html=True
    )
    
else:
    st.warning("Preencha todos os campos do Formulário Cadastrar GRD para gerar o Número de Identificação da GRD")

    # Inclusão das Fichas

st.markdown("#### Relação de Fichas para Cadastro")

dados = pd.DataFrame(
{
"Item": [""] * 100,
"Identificação da Ficha": [""] * 100,
"Título da Ficha": [""] * 100,
"Qtd. Docs": [""] * 100,
"Qtd. Pág.": [""] * 100,
"Rev.": [""] * 100,
"Escopo(s)": [""] * 100,
}
)

tabela_fichas = st.data_editor(
dados,
key=f"tabela_fichas_{st.session_state.reset_form}",
use_container_width=True,
hide_index=True,
column_config={
"Item": st.column_config.NumberColumn(
"Item",
min_value=1,
step=1,
width="small",
),
"Identificação da Ficha": st.column_config.TextColumn(
"Identificação da Ficha",
width="medium"
),
"Título da Ficha": st.column_config.TextColumn(
"Título da Ficha",
width="large" # Coluna maior
),
"Qtd. Docs": st.column_config.NumberColumn(
"Qtd. Docs",
min_value=1,
step=1,
width="small"
),
"Qtd. Pág.": st.column_config.NumberColumn(
"Qtd. Pág.",
min_value=1,
step=1,
width="small"
),
"Revisão": st.column_config.TextColumn(
"Revisão",
width="small"
),
"Escopo(s)": st.column_config.TextColumn(
"Escopo(s)",
width="medium"
),
}
)

# Botão de Cadastro
if st.button("📋 Cadastrar GRD", type="primary"):

    campos_preenchidos = all([
    subestacao,
    contratada,
    lista_contratos,
    lista_empreendimento,
    lista_responsavel_contratadas,
    agrupamento,
    lista_subagrupamentos,
    lista_responsavel_cemig,
    identificação_contratada,
    data
    ])
    
    if not campos_preenchidos:
        st.error("Preencha todos os campos do formulário Cadastrar GRD")
    
    else:
        
        fichas_preenchidas = tabela_fichas.replace("", pd.NA).dropna(how="all")
        
        if fichas_preenchidas.empty:
            st.error("Preencha pelo menos uma linha da tabela Relação de Fichas para Cadastro")
        
        else:
        
            linhas_com_erro = []
        
            for idx, linha in fichas_preenchidas.iterrows():
        
             if linha.isna().any():
                linhas_com_erro.append(str(int(linha["Item"])))
        
            if linhas_com_erro:
                st.error(
                    f"Existem campos vazios no(s) item(s): {', '.join(linhas_com_erro)}"
                )
        
            else:

                # Dados da GRD
        # Montar uma linha para cada ficha
                
                dados_exportacao = []
            
                for _, ficha in fichas_preenchidas.iterrows():
                
                    dados_exportacao.append({

                    "Item": ficha["Item"],
                    "Número GRD": numero_grd,
                    "Subestação": subestacao,
                    "Contratada": contratada,
                    "Contrato": lista_contratos,
                    "Empreendimento": lista_empreendimento,
                    "Responsável Contratada": lista_responsavel_contratadas,
                    "Agrupamento": agrupamento,
                    "Subagrupamento": lista_subagrupamentos,
                    "Responsável CEMIG - Recebimento": lista_responsavel_cemig,
                    "Identificação Contratada": identificação_contratada,
                    "Data de Entrada": data.strftime("%d/%m/%Y"),
                    "Identificação da Ficha": ficha["Identificação da Ficha"],
                    "Título da Ficha": ficha["Título da Ficha"],
                    "Qtd. Docs": ficha["Qtd. Docs"],
                    "Qtd. Pág.": ficha["Qtd. Pág."],
                    "Rev.": ficha["Rev."],
                    "Escopo(s)": ficha["Escopo(s)"],
                    "Observações": " "
                })
                
                df_exportacao = pd.DataFrame(dados_exportacao)

                # ==========================================
                # ATUALIZAR BASE DE DADOS DAS GRDs
                # ==========================================

                arquivo_base = "BASE DE DADOS GRDs CADASTRADAS.xlsx"

                try:

                    wb_base = load_workbook(arquivo_base)
                    wb_base.calculation.fullCalcOnLoad = True
                    wb_base.calculation.forceFullCalc = True

                    ano_aba = str(data.year)

                    if ano_aba not in wb_base.sheetnames:
                        st.error(
                            f"A aba {ano_aba} não existe na Base de Dados GRDs Cadastradas."
                        )
                        st.stop()

                    ws_base = wb_base[ano_aba]

                    # Verificar se a GRD já existe na base
                    grd_existente = False

                    for linha in range(2, ws_base.max_row + 1):

                        valor_grd = ws_base.cell(linha, 1).value

                        if valor_grd is None:
                            continue

                        valor_grd = str(valor_grd).strip()

                        if valor_grd == "":
                            continue

                        if valor_grd == numero_grd:

                            grd_existente = True
                            break

                    if grd_existente:

                        st.error(
                            f"❌ A GRD {numero_grd} já está cadastrada na base de dados."
                        )

                        st.stop()
                    

                    # Procura a primeira linha vazia na coluna A
                    proxima_linha = None

                    for linha in range(2, ws_base.max_row + 1):

                        valor = ws_base.cell(linha, 1).value

                        if valor is None or str(valor).strip() == "":
                            proxima_linha = linha
                            break

                    if proxima_linha is None:
                        proxima_linha = ws_base.max_row + 1

                    for _, registro in df_exportacao.iterrows():

                        ws_base.cell(proxima_linha, 1).value = registro["Número GRD"]
                        ws_base.cell(proxima_linha, 2).value = registro["Subestação"]
                        ws_base.cell(proxima_linha, 3).value = registro["Contratada"]
                        ws_base.cell(proxima_linha, 4).value = registro["Contrato"]
                        ws_base.cell(proxima_linha, 5).value = registro["Empreendimento"]
                        ws_base.cell(proxima_linha, 6).value = registro["Responsável Contratada"]
                        ws_base.cell(proxima_linha, 7).value = registro["Agrupamento"]
                        ws_base.cell(proxima_linha, 8).value = registro["Subagrupamento"]
                        ws_base.cell(proxima_linha, 9).value = registro["Responsável CEMIG - Recebimento"]
                        ws_base.cell(proxima_linha, 10).value = registro["Identificação Contratada"]
                        ws_base.cell(proxima_linha, 11).value = registro["Data de Entrada"]
                        ws_base.cell(proxima_linha, 12).value = registro["Identificação da Ficha"]
                        ws_base.cell(proxima_linha, 13).value = registro["Título da Ficha"]
                        ws_base.cell(proxima_linha, 14).value = registro["Qtd. Docs"]
                        ws_base.cell(proxima_linha, 15).value = registro["Qtd. Pág."]
                        ws_base.cell(proxima_linha, 16).value = registro["Rev."]
                        ws_base.cell(proxima_linha, 17).value = registro["Escopo(s)"]

                        # Aplicar bordas em toda a linha cadastrada
                        for col in range(1, ws_base.max_column + 1):
                            ws_base.cell(proxima_linha, col).border = borda_fina
                            
                        # ==========================================
                        # COPIAR FORMATAÇÕES E FÓRMULAS COLUNAS V+
                        # ==========================================

                        # Utiliza a última linha válida como modelo
                        linha_modelo = None

                        for linha in range(proxima_linha - 1, 1, -1):

                            if ws_base.cell(linha, 22).value is not None:
                                linha_modelo = linha
                                break

                        # Se não encontrar, usa a linha 2
                        if linha_modelo is None:
                            linha_modelo = 2

                        for col in range(22, ws_base.max_column + 1):

                            origem = ws_base.cell(linha_modelo, col)
                            destino = ws_base.cell(proxima_linha, col)

                            # Copiar toda a formatação
                            destino.font = copy(origem.font)
                            destino.fill = copy(origem.fill)
                            destino.border = copy(origem.border)
                            destino.alignment = copy(origem.alignment)
                            destino.number_format = origem.number_format
                            destino.protection = copy(origem.protection)

                            # Copiar fórmula ajustando para a nova linha
                            if isinstance(origem.value, str) and origem.value.startswith("="):

                                destino.value = Translator(
                                    origem.value,
                                    origin=origem.coordinate
                                ).translate_formula(destino.coordinate)

                            else:
                                destino.value = origem.value                 

                        # CENTRALIZAR COLUNAS J EM DIANTE
                        for col in range(10, ws_base.max_column + 1):

                            ws_base.cell(
                                proxima_linha,
                                col
                            ).alignment = Alignment(
                                horizontal="center",
                                vertical="center"
                            )    

                        proxima_linha += 1

                        # Ajustar largura automática das colunas A até U
                        for coluna in ws_base.iter_cols(min_col=1, max_col=21):

                            maior_tamanho = 0
                            letra_coluna = get_column_letter(coluna[0].column)

                            for cell in coluna:

                                if cell.value is not None:

                                    tamanho_texto = len(str(cell.value))

                                    if tamanho_texto > maior_tamanho:
                                        maior_tamanho = tamanho_texto

                            ws_base.column_dimensions[letra_coluna].width = min(maior_tamanho + 5, 80)

                        # Larguras fixas das colunas de controle
                        ws_base.column_dimensions["V"].width = 15
                        ws_base.column_dimensions["W"].width = 18
                        ws_base.column_dimensions["X"].width = 22
                        ws_base.column_dimensions["Y"].width = 22

                    wb_base.save(arquivo_base)

                    wb_base.calculation.fullCalcOnLoad = True
                    wb_base.calculation.forceFullCalc = True

                    # Limpar cache da consulta de GRDs
                    carregar_grds.clear()

                except Exception as erro:

                    st.error(
                        f"Erro ao atualizar a Base de Dados: {erro}"
                    )

                    st.stop()
                #

                nome_arquivo = f"{numero_grd}.xlsx"

                df_exportacao.to_excel(nome_arquivo, index=False)

                # Abrir arquivo gerado
                wb = load_workbook(nome_arquivo)

                ws = wb.active

                # Renomear aba
                ws.title = "GRD"

                # Formatar cabeçalho
                for cell in ws[1]:
                    cell.font = Font(
                        name="Aptos Narrow",
                        bold=True,
                        italic=True
                    )

                # Ajustar largura das colunas
                for coluna in ws.columns:

                    maior_tamanho = 0
                    letra_coluna = get_column_letter(coluna[0].column)

                    for cell in coluna:

                        if cell.value is not None:

                            tamanho_texto = len(str(cell.value))

                            if tamanho_texto > maior_tamanho:
                                maior_tamanho = tamanho_texto

                    ws.column_dimensions[letra_coluna].width = min(maior_tamanho + 5, 60)

                # Aplicar bordas em todas as células utilizadas
                for linha in ws.iter_rows(
                    min_row=1,
                    max_row=ws.max_row,
                    min_col=1,
                    max_col=ws.max_column
                ):
                    for cell in linha:
                        cell.border = borda_fina

                wb.save(nome_arquivo)

                st.success("GRD Cadastrada com Sucesso!")

                with open(nome_arquivo, "rb") as arquivo:

                    st.download_button(
                        label="📥 Baixar Arquivo GRD",
                        data=arquivo,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                   
                    # FIM DO CADASTRO DA GRD
                    # ======================

# ==========================================
# CONSULTA GRDs CADASTRADAS
# ==========================================

st.markdown("---")
st.markdown("#### 📋 GRDs Cadastradas")

# ==========================================
# DOWNLOAD DA BASE
# ==========================================

arquivo_base = "BASE DE DADOS GRDs CADASTRADAS.xlsx"

if os.path.exists(arquivo_base):

    with open(arquivo_base, "rb") as arquivo:

        st.download_button(
            label="📥 Baixar Base de Dados",
            data=arquivo,
            file_name="BASE DE DADOS GRDs CADASTRADAS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            
        )

if st.button("🔄 Atualizar Base"):

    carregar_grds.clear()
    st.rerun()

try:

    df_consulta = carregar_grds()

    if not df_consulta.empty:

       
        # ==========================================
        # FILTROS
        # ==========================================

        col1, col2, col3, col4, col5, col6 = st.columns(6)

        # Botão para limpar filtros

        if st.button("🔄 Limpar Filtros"):
            
            st.session_state.filtro_ano = "Todos"
            st.session_state.filtro_subestacao = "Todas"
            st.session_state.filtro_contratada = "Todas"
            st.session_state.filtro_empreendimento = "Todos"
            st.session_state.filtro_agrupamento = "Todos"
            st.session_state.filtro_subagrupamento = "Todos"

            st.rerun()

        with col1:

            lista_anos = ["Todos"] + sorted(
                df_consulta["Ano"]
                .dropna()
                .astype(str)
                .unique()
                .tolist(),
                key=int
            )

            filtro_ano = st.selectbox(
                "Ano",
                lista_anos,
                key="filtro_ano"
            )

        with col2:

            lista_subestacoes = ["Todas"] + sorted(
                df_consulta["Subestação"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            filtro_subestacao = st.selectbox(
                "Subestação",
                lista_subestacoes,
                key="filtro_subestacao"
            )

        with col3:

            lista_contratadas = ["Todas"] + sorted(
                df_consulta["Contratada"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            filtro_contratada = st.selectbox(
                "Contratada",
                lista_contratadas,
                 key="filtro_contratada"
            )

        with col4:

            lista_empreendimentos = ["Todos"] + sorted(
                df_consulta["Empreendimento"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            filtro_empreendimento = st.selectbox(
                "Empreendimento",
                lista_empreendimentos,
                key="filtro_empreendimento"
            )

        with col5:

            lista_agrupamentos = ["Todos"] + sorted(
                df_consulta["Agrupamento"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            filtro_agrupamento = st.selectbox(
                "Agrupamento",
                lista_agrupamentos,
                key="filtro_agrupamento"
            )

        with col6:

            lista_subagrupamentos = ["Todos"] + sorted(
                df_consulta["Subagrupamento"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )

            filtro_subagrupamento = st.selectbox(
                "Subagrupamento",
                lista_subagrupamentos,
                 key="filtro_subagrupamento"
            )

        # ==========================================
        # APLICAR FILTROS
        # ==========================================

        df_filtrado = df_consulta.copy()

        if filtro_ano != "Todos":

            df_filtrado = df_filtrado[
                df_filtrado["Ano"].astype(str) == filtro_ano
            ]

        if filtro_subestacao != "Todas":

            df_filtrado = df_filtrado[
                df_filtrado["Subestação"].astype(str)
                == filtro_subestacao
            ]

        if filtro_contratada != "Todas":

            df_filtrado = df_filtrado[
                df_filtrado["Contratada"].astype(str)
                == filtro_contratada
            ]

        if filtro_empreendimento != "Todos":

            df_filtrado = df_filtrado[
                df_filtrado["Empreendimento"].astype(str)
                == filtro_empreendimento
            ]   

        if filtro_agrupamento != "Todos":

            df_filtrado = df_filtrado[
                df_filtrado["Agrupamento"].astype(str)
                == filtro_agrupamento
            ]

        if filtro_subagrupamento != "Todos":

            df_filtrado = df_filtrado[
                df_filtrado["Subagrupamento"].astype(str)
                == filtro_subagrupamento
            ]

        # ==========================================
        # COLUNAS EXIBIDAS
        # ==========================================
       
        # Base completa para Dashboard
        df_dashboard_base = df_consulta.copy()

        colunas_exibidas = [
            "Número GRD",
            "Data de Entrada",
            "Responsável Contratada",
            "Responsável CEMIG - Recebimento",
            "Identificação da Ficha",
            "Rev.",
            "Escopo(s)",
            "Status Carimbo",
            "Status Ficha",
            "Data de Saída",
            "Responsável CEMIG - Devolução",
            "Status GRD",
            "Dias p/ Vencimento GRD",
            "Dias Fora do Prazo GRD"
        ]

        colunas_exibidas = [
            col
            for col in colunas_exibidas
            if col in df_filtrado.columns
        ]

        df_filtrado = df_filtrado[colunas_exibidas]

        # Formatar datas para DD/MM/AAAA

        colunas_data = [
            "Data de Entrada",
            "Data de Saída",
            "Data Devolução"
        ]

        for coluna in colunas_data:

            if coluna in df_filtrado.columns:

                df_filtrado[coluna] = pd.to_datetime(
                    df_filtrado[coluna],
                    errors="coerce"
                ).dt.strftime("%d/%m/%Y")

        # ==========================================
        # TABELA
        # ==========================================

        st.data_editor(
            df_filtrado,
            use_container_width=True,
            hide_index=True,
            disabled=True,
            height=500
        )

        st.success(
            f"Total de registros encontrados: {len(df_filtrado)}"
        )

    else:

        st.warning(
            "Não existem registros cadastrados."
        )

except Exception as erro:

    st.error(
        f"Erro ao carregar a Base de Dados: {erro}"
    )

# ==========================================
# DEVOLVER GRD
# ==========================================

if "reset_devolucao" not in st.session_state:
    st.session_state.reset_devolucao = 0

st.markdown("---")
st.markdown("#### 📤 Devolução de GRD")

if st.button("🗑️ Limpar Campos Editáveis"):
    st.session_state.reset_devolucao += 1
    st.rerun()

df_grds = carregar_grds()

lista_grds = [""] + sorted(
    df_grds["Número GRD"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

with st.container(border=True):

    # ======================================
    # GRD
    # ======================================

    col_grd, col_vazio = st.columns([4, 4])

    with col_grd:

        grd_selecionada = st.selectbox(
            "GRD",
            lista_grds,
            key=f"grd_devolucao_{st.session_state.reset_devolucao}"
        )

    agrupamento_devolucao = ""
    responsaveis_devolucao = []
    df_grd = pd.DataFrame()

    if grd_selecionada:

        df_grd = df_grds[
            df_grds["Número GRD"] == grd_selecionada
        ].copy()

        if not df_grd.empty:

            agrupamento_devolucao = (
                df_grd["Agrupamento"]
                .iloc[0]
            )

    # ======================================
    # RESPONSÁVEIS POR AGRUPAMENTO
    # ======================================

    if agrupamento_devolucao == "EQUIPAMENTOS - EQP":

        responsaveis_devolucao = (
            df["EQP_RESPONSAVEIS"]
            .dropna()
            .astype(str)
            .tolist()
        )

    elif agrupamento_devolucao == "SUBESTAÇÕES - SE":

        responsaveis_devolucao = (
            df["SE_RESPONSAVEIS"]
            .dropna()
            .astype(str)
            .tolist()
        )

    elif agrupamento_devolucao == "LINHAS DE TRANSMISSÃO - LT":

        responsaveis_devolucao = (
            df["LT_RESPONSAVEIS"]
            .dropna()
            .astype(str)
            .tolist()
        )

    # ======================================
    # DADOS DA DEVOLUÇÃO
    # ======================================

    col1, col2, col3, col4 = st.columns([1.4, 2.2, 1, 2])

    with col1:

        st.text_input(
            "Agrupamento",
            value=agrupamento_devolucao,
            disabled=True
        )

    with col2:

        responsavel_devolucao = st.selectbox(
            "Responsável CEMIG - Devolução",
            [""] + responsaveis_devolucao,
            key=f"responsavel_devolucao_{st.session_state.reset_devolucao}"
        )

    with col3:

        data_saida = st.date_input(
            "Data de Saída",
            value=None,
            format="DD/MM/YYYY",
            key=f"data_saida_{st.session_state.reset_devolucao}"
        )

    # ======================================
    # TABELA DE DEVOLUÇÃO
    # ======================================

    st.markdown("#### Relação de Fichas para Devolução")

    dados_devolucao = pd.DataFrame(
    {
        "Item": [""] * 100,
        "Identificação da Ficha": [""] * 100,
        "Status Carimbo": [""] * 100,
        "Status Ficha": [""] * 100
    }
    )

    tabela_devolucao = st.data_editor(

        dados_devolucao,

        key=f"tabela_devolucao_{st.session_state.reset_devolucao}",

        use_container_width=True,

        hide_index=True,

        column_config={

            "Item": st.column_config.NumberColumn(
                "Item",
                min_value=1,
                step=1,
                width=60
            ),

            "Identificação da Ficha":
                st.column_config.TextColumn(
                    "Identificação da Ficha",
                    width="large"
                ),

            "Status Carimbo":
                st.column_config.SelectboxColumn(
                    "Status Carimbo",
                    options=[
                        "NÃO SE APLICA",
                        "VISTO",
                        "VISTO C/ NOTAS",
                        "REVISÃO",
                        "P/ CONHECIMENTO"
                    ],
                    width="medium"
                ),

            "Status Ficha":
                st.column_config.SelectboxColumn(
                    "Status Ficha",
                    options=[
                        "APROVADA",
                        "REPROVADA"
                    ],
                    width="medium"
                )

        }
    )

# ==========================================
# MENSAGEM DE SUCESSO DA DEVOLUÇÃO
# ==========================================

if "msg_devolucao" in st.session_state:

    st.success(
        st.session_state["msg_devolucao"]
    )

    st.info(
        f"""
    📧 Este arquivo deve ser enviado para a contratada.

    Assunto do e-mail:

    {grd_selecionada}
    """
    )

    if (
        "arquivo_devolucao_bytes"
        in st.session_state
    ):

        download_realizado = st.download_button(
            label="📥 Baixar Arquivo",
            data=st.session_state[
                "arquivo_devolucao_bytes"
            ],
            file_name=st.session_state[
                "arquivo_devolucao_nome"
            ],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_devolucao"
        )

        if download_realizado:

            del st.session_state[
                "arquivo_devolucao_bytes"
            ]

            del st.session_state[
                "arquivo_devolucao_nome"
            ]

            del st.session_state[
                "msg_devolucao"
            ]

            st.rerun()

# ======================================
# BOTÃO DEVOLVER GRD
# ======================================

if st.button(
    "📤 Devolver GRD",
    type="primary"
):

    fichas_devolvidas = (
        tabela_devolucao
        .replace("", pd.NA)
        .dropna(how="all")
    )

    # ==================================
    # VALIDAR CABEÇALHO
    # ==================================

    if not grd_selecionada:

        st.error(
            "Selecione uma GRD."
        )

        st.stop()

    if not responsavel_devolucao:

        st.error(
            "Selecione o Responsável CEMIG - Devolução."
        )

        st.stop()

    if data_saida is None:

        st.error(
            "Informe a Data de Saída."
        )

        st.stop()

    # ==================================
    # VALIDAR TABELA
    # ==================================

    if fichas_devolvidas.empty:

        st.error(
            "Preencha pelo menos uma linha da tabela Relação de Fichas para Devolução."
        )

        st.stop()

    linhas_com_erro = []

    for _, linha in fichas_devolvidas.iterrows():

        campos_item = [

            linha["Item"],

            linha["Identificação da Ficha"],

            linha["Status Carimbo"],

            linha["Status Ficha"]

        ]

        if any(
            pd.isna(valor)
            or str(valor).strip() == ""
            for valor in campos_item
        ):

            item = linha["Item"]

            if pd.isna(item):
                item = "Sem Item"

            linhas_com_erro.append(
                str(item)
            )

    if linhas_com_erro:

        st.error(
            f"Existem campos vazios no(s) item(s): {', '.join(linhas_com_erro)}"
        )

        st.stop()

    # ==================================
    # PROCESSAMENTO
    # ==================================

    arquivo_base = (
        "BASE DE DADOS GRDs CADASTRADAS.xlsx"
    )

    try:

        wb = load_workbook(arquivo_base)

        fichas_nao_encontradas = []

        fichas_ja_devolvidas = []

        registros_para_atualizar = []

        # ==================================
        # PROCURAR FICHAS
        # ==================================

        for _, ficha in fichas_devolvidas.iterrows():

            identificacao = (
                str(
                    ficha["Identificação da Ficha"]
                )
                .replace(".0", "")
                .strip()
            )

            encontrada = False

            for aba in wb.sheetnames:

                ws = wb[aba]

                for linha in range(
                    2,
                    ws.max_row + 1
                ):

                    numero_grd_base = str(
                        ws.cell(linha, 1).value
                    ).strip()

                    identificacao_base = (
                        str(
                            ws.cell(linha, 12).value
                        )
                        .replace(".0", "")
                        .strip()
                    )

                    if (
                        numero_grd_base == grd_selecionada
                        and
                        identificacao_base == identificacao
                    ):

                        encontrada = True

                        status_carimbo_existente = (
                            ws.cell(linha, 18).value
                        )

                        status_ficha_existente = (
                            ws.cell(linha, 19).value
                        )

                        responsavel_existente = (
                            ws.cell(linha, 20).value
                        )

                        data_saida_existente = (
                            ws.cell(linha, 21).value
                        )

                        if (
                            status_carimbo_existente not in [None, ""]
                            or status_ficha_existente not in [None, ""]
                            or responsavel_existente not in [None, ""]
                            or data_saida_existente not in [None, ""]
                        ):

                            fichas_ja_devolvidas.append(
                                identificacao
                            )

                        else:

                            registros_para_atualizar.append(
                                (
                                    ws,
                                    linha,
                                    ficha
                                )
                            )

                        break

                if encontrada:
                    break

            if not encontrada:

                fichas_nao_encontradas.append(
                    identificacao
                )

        # ==================================
        # FICHAS NÃO ENCONTRADAS
        # ==================================

        if fichas_nao_encontradas:

            st.error(
                "As seguintes fichas não foram encontradas:"
            )

            for ficha in fichas_nao_encontradas:

                st.write(
                    f"• {ficha}"
                )

            st.stop()

        # ==================================
        # FICHAS JÁ DEVOLVIDAS
        # ==================================

        if fichas_ja_devolvidas:

            st.error(
                "As seguintes fichas já foram devolvidas:"
            )

            for ficha in fichas_ja_devolvidas:

                st.write(
                    f"• {ficha}"
                )

            st.stop()

        # ==================================
        # ATUALIZAR BASE
        # ==================================

        for ws, linha, ficha in registros_para_atualizar:

            ws.cell(
                linha,
                18
            ).value = ficha[
                "Status Carimbo"
            ]

            ws.cell(
                linha,
                19
            ).value = ficha[
                "Status Ficha"
            ]

            ws.cell(
                linha,
                20
            ).value = (
                responsavel_devolucao
            )

            ws.cell(
                linha,
                21
            ).value = data_saida.strftime(
                "%d/%m/%Y"
            )

        wb.save(arquivo_base)

        carregar_grds.clear()
        st.cache_data.clear()

        # ==================================
        # GERAR ARQUIVO DAS FICHAS DEVOLVIDAS
        # ==================================

        from io import BytesIO

        colunas_exportacao = [
            "Item",
            "Número GRD",
            "Subestação",
            "Contratada",
            "Contrato",
            "Empreendimento",
            "Responsável Contratada",
            "Agrupamento",
            "Subagrupamento",
            "Responsável CEMIG - Recebimento",
            "Data de Entrada",
            "Identificação da Ficha",
            "Título da Ficha",
            "Qtd. Docs",
            "Qtd. Pág.",
            "Rev.",
            "Escopo(s)",
            "Status Carimbo",
            "Status Ficha",
            "Responsável CEMIG - Devolução",
            "Data de Saída",
            "Observações"
            
        ]

        dados_download = []

        # SOMENTE AS FICHAS DEVOLVIDAS
        for ws, linha_excel, ficha in registros_para_atualizar:

            registro = {}

            # Item informado na tabela de devolução
            registro["Item"] = ficha["Item"]

            for col in range(1, ws.max_column + 1):

                cabecalho = ws.cell(
                    1,
                    col
                ).value

                if (
                    cabecalho in colunas_exportacao
                    and cabecalho != "Item"
                ):

                    registro[cabecalho] = ws.cell(
                        linha_excel,
                        col
                    ).value

            dados_download.append(registro)

        df_download = pd.DataFrame(dados_download)
        
        df_download["Observações"] = ""

        ordem_colunas = [
            "Item",
            "Número GRD",
            "Subestação",
            "Contratada",
            "Contrato",
            "Empreendimento",
            "Responsável Contratada",
            "Agrupamento",
            "Subagrupamento",
            "Responsável CEMIG - Recebimento",
            "Data de Entrada",
            "Identificação da Ficha",
            "Título da Ficha",
            "Qtd. Docs",
            "Qtd. Pág.",
            "Rev.",
            "Escopo(s)",
            "Status Carimbo",
            "Status Ficha",
            "Responsável CEMIG - Devolução",
            "Data de Saída",
            "Observações"
        ]

        df_download = df_download[ordem_colunas]


        buffer = BytesIO()

        with pd.ExcelWriter(
            buffer,
            engine="openpyxl"
        ) as writer:

            df_download.to_excel(
                writer,
                sheet_name="DEVOLUCAO GRD",
                index=False
            )

            ws_export = writer.book["DEVOLUCAO GRD"]

            for coluna in ws_export.iter_cols():

                titulo = coluna[0].value

                if titulo in [
                    "Data de Entrada",
                    "Data de Saída"
                    
                ]:

                    for celula in coluna[1:]:

                        celula.number_format = "dd/mm/yyyy"

            # Formatar cabeçalho
            for cell in ws_export[1]:

                cell.font = Font(
                    name="Aptos Narrow",
                    size=12,
                    bold=True,
                    italic=True
                )

            # Aplicar bordas em todas as células utilizadas

            borda_fina = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin")
            )

            for linha in ws_export.iter_rows(
                min_row=1,
                max_row=ws_export.max_row,
                min_col=1,
                max_col=ws_export.max_column
            ):

                for cell in linha:

                    if cell.value is not None:

                        cell.border = borda_fina

            # Ajustar largura automática
            for coluna in ws_export.columns:

                maior_tamanho = 0

                letra_coluna = get_column_letter(
                    coluna[0].column
                )

                for cell in coluna:

                    if cell.value is not None:

                        maior_tamanho = max(
                            maior_tamanho,
                            len(str(cell.value))
                        )

                ws_export.column_dimensions[
                    letra_coluna
                ].width = min(
                    maior_tamanho + 5,
                    80
                )

        # Nome do arquivo
        
        arquivo_devolucao = (
        f"DEVOLUCAO_{grd_selecionada}.xlsx"
        )
        
        # Salvar na pasta raiz
        
        with open(
            arquivo_devolucao,
            "wb"
        ) as arquivo:
        
            arquivo.write(
                buffer.getvalue()
            )
        
        buffer.seek(0)
        
        # Disponibilizar para download
        
        nome_data = data_saida.strftime("%d%m%y")

        arquivo_devolucao = (
            f"DEVOLUCAO_{grd_selecionada}_{nome_data}.xlsx"
        )

        # Salvar na pasta raiz

        with open(
            arquivo_devolucao,
            "wb"
        ) as arquivo:

            arquivo.write(
                buffer.getvalue()
            )

        # Disponibilizar para download

        st.session_state[
            "arquivo_devolucao_bytes"
        ] = buffer.getvalue()

        st.session_state[
            "arquivo_devolucao_nome"
        ] = arquivo_devolucao

        carregar_grds.clear()
        st.cache_data.clear()

        st.session_state["msg_devolucao"] = (
            f"✅ {grd_selecionada} foi atualizada com sucesso!"
        )

        st.rerun()
        ####

    except Exception as erro:

        st.error(
            f"Erro ao processar devolução: {erro}"
        )


# ==========================================
# DASHBOARD
# ==========================================

st.markdown("---")
st.markdown("### 📊 Dashboard")

df_dashboard = df_dashboard_base.copy()

df_dashboard["Qtd. Docs"] = pd.to_numeric(
    df_dashboard["Qtd. Docs"],
    errors="coerce"
).fillna(0)

if not df_dashboard.empty:

    # ======================================
    # FILTROS DO DASHBOARD
    # ======================================

    colf1, colf2, colf3 = st.columns(3)

    with colf1:

        ano_dashboard = st.selectbox(
            "Ano",
            ["Todos"] + sorted(
                df_dashboard["Ano"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="dashboard_ano"
        )

    with colf2:

        contratada_dashboard = st.selectbox(
            "Contratada",
            ["Todas"] + sorted(
                df_dashboard["Contratada"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="dashboard_contratada"
        )

    with colf3:

        empreendimento_dashboard = st.selectbox(
            "Empreendimento",
            ["Todos"] + sorted(
                df_dashboard["Empreendimento"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="dashboard_empreendimento"
        )

    # ======================================
    # APLICAR FILTROS
    # ======================================

    if ano_dashboard != "Todos":

        df_dashboard = df_dashboard[
            df_dashboard["Ano"].astype(str)
            == ano_dashboard
        ]

    if contratada_dashboard != "Todas":

        df_dashboard = df_dashboard[
            df_dashboard["Contratada"]
            == contratada_dashboard
        ]

    if empreendimento_dashboard != "Todos":

        df_dashboard = df_dashboard[
            df_dashboard["Empreendimento"]
            == empreendimento_dashboard
        ]

    # ======================================
    # GARANTIR NUMÉRICO
    # ======================================

    df_dashboard["Qtd. Docs"] = pd.to_numeric(
        df_dashboard["Qtd. Docs"],
        errors="coerce"
    ).fillna(0)

    # ======================================
    # ETIQUETAS
    # ======================================

    total_docs = int(
    df_dashboard["Qtd. Docs"].sum()
    )

    docs_prazo = int(
        df_dashboard.loc[
            df_dashboard["Status GRD"]
            .astype(str)
            .str.upper()
            .str.strip()
            == "NO PRAZO",
            "Qtd. Docs"
        ].sum()
    )

    docs_fora = int(
        df_dashboard.loc[
            df_dashboard["Status GRD"]
            .astype(str)
            .str.upper()
            .str.strip()
            == "FORA DO PRAZO",
            "Qtd. Docs"
        ].sum()
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "📄 Documentos Tramitados",
        f"{int(total_docs):,}".replace(",", ".")
    )

    col2.metric(
        "✅ No Prazo",
        f"{int(docs_prazo):,}".replace(",", ".")
    )

    col3.metric(
        "⚠️ Fora do Prazo",
        f"{int(docs_fora):,}".replace(",", ".")
    )

    # ======================================
    # AGRUPAMENTOS
    # ======================================

    agrupamentos = [
        "EQUIPAMENTOS - EQP",
        "SUBESTAÇÕES - SE",
        "LINHAS DE TRANSMISSÃO - LT"
    ]

    dados_total = []
    dados_prazo = []
    dados_fora = []

    for agrupamento in agrupamentos:

        df_agr = df_dashboard[
            df_dashboard["Agrupamento"]
            == agrupamento
        ]

        total = df_agr["Qtd. Docs"].sum()

        prazo = df_agr[
            df_agr["Status GRD"]
            == "NO PRAZO"
        ]["Qtd. Docs"].sum()

        fora = df_agr[
            df_agr["Status GRD"]
            == "FORA DO PRAZO"
        ]["Qtd. Docs"].sum()

        dados_total.append(total)
        dados_prazo.append(prazo)
        dados_fora.append(fora)

    # ======================================
# 3 GRÁFICOS DE ROSCA
# ======================================

cores_agrupamentos = {
    "EQUIPAMENTOS - EQP": "#2E8B57",
    "SUBESTAÇÕES - SE": "#FF8C00",
    "LINHAS DE TRANSMISSÃO - LT": "#1F4E79"
}

g1, g2, g3 = st.columns(3)

# ======================================
# DOCUMENTOS TRAMITADOS
# ======================================

df_grafico_total = pd.DataFrame({
    "Agrupamento": agrupamentos,
    "Quantidade": dados_total
})

df_grafico_total = df_grafico_total[
    df_grafico_total["Quantidade"] > 0
]

with g1:

    fig1 = px.pie(
        df_grafico_total,
        names="Agrupamento",
        values="Quantidade",
        title="Documentos Tramitados (%)",
        hole=0.45,
        color="Agrupamento",
        color_discrete_map=cores_agrupamentos
    )

    fig1.update_traces(
         textinfo="percent",
         textposition="inside",
         textfont_size=16
    )

    fig1.update_layout(
        showlegend=True,
        legend=dict(
            orientation="v",
            font=dict(size=11)
        ),
        margin=dict(
            l=10,
            r=10,
            t=50,
            b=10
        ),
        height=400
    )

    st.plotly_chart(
        fig1,
        use_container_width=True
    )

# ======================================
# DOCUMENTOS NO PRAZO
# ======================================

df_grafico_prazo = pd.DataFrame({
    "Agrupamento": agrupamentos,
    "Quantidade": dados_prazo
})

df_grafico_prazo = df_grafico_prazo[
    df_grafico_prazo["Quantidade"] > 0
]

with g2:

    fig2 = px.pie(
        df_grafico_prazo,
        names="Agrupamento",
        values="Quantidade",
        title="Documentos No Prazo (%)",
        hole=0.45,
        color="Agrupamento",
        color_discrete_map=cores_agrupamentos
    )

    fig2.update_traces(
         textinfo="percent",
         textposition="inside",
         textfont_size=16
    )

    fig2.update_layout(
        showlegend=True,
        legend=dict(
            orientation="v",
            font=dict(size=11)
        ),
        margin=dict(
            l=10,
            r=10,
            t=50,
            b=10
        ),
        height=400
    )

    st.plotly_chart(
        fig2,
        use_container_width=True
    )

# ======================================
# DOCUMENTOS FORA DO PRAZO
# ======================================

df_grafico_fora = pd.DataFrame({
    "Agrupamento": agrupamentos,
    "Quantidade": dados_fora
})

# REMOVE AGRUPAMENTOS ZERADOS
df_grafico_fora = df_grafico_fora[
    df_grafico_fora["Quantidade"] > 0
]

with g3:

    if not df_grafico_fora.empty:

        fig3 = px.pie(
            df_grafico_fora,
            names="Agrupamento",
            values="Quantidade",
            title="Documentos Fora do Prazo (%)",
            hole=0.45,
            color="Agrupamento",
            color_discrete_map=cores_agrupamentos
        )

        fig3.update_traces(
             textinfo="percent",
             textposition="inside",
             textfont_size=16
        )

        fig3.update_layout(
            showlegend=True,
            legend=dict(
                orientation="v",
                font=dict(size=11)
            ),
            margin=dict(
                l=10,
                r=10,
                t=50,
                b=10
            ),
            height=400
        )

        st.plotly_chart(
            fig3,
            use_container_width=True
        )

    else:

        st.info(
            "Não existem documentos fora do prazo para os filtros selecionados."
        )

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# ==========================================
# ANÁLISE MENSAL DE TRAMITAÇÃO
# ==========================================

st.markdown("---")
st.markdown("### 📈 Documentos Tramitados por Mês")

df_indicador = df_dashboard_base.copy()

if not df_indicador.empty:

    # ======================================
    # TRATAMENTO DOS DADOS
    # ======================================

    df_indicador["Data de Entrada"] = pd.to_datetime(
        df_indicador["Data de Entrada"],
        dayfirst=True,
        errors="coerce"
    )

    df_indicador["Qtd. Docs"] = pd.to_numeric(
        df_indicador["Qtd. Docs"],
        errors="coerce"
    ).fillna(0)

    df_indicador = df_indicador.dropna(
        subset=["Data de Entrada"]
    )

    # ======================================
    # FILTROS
    # ======================================

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        filtro_ano = st.selectbox(
            "Ano",
            ["Todos"] + sorted(
                df_indicador["Ano"]
                .astype(str)
                .unique()
                .tolist()
            ),
            key="grafico_mes_ano"
        )

    with col2:

        filtro_agrupamento = st.selectbox(
            "Agrupamento",
            ["Todos"] + sorted(
                df_indicador["Agrupamento"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="grafico_mes_agrupamento"
        )

    with col3:

        filtro_contratada = st.selectbox(
            "Contratada",
            ["Todas"] + sorted(
                df_indicador["Contratada"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="grafico_mes_contratada"
        )

    with col4:

        filtro_empreendimento = st.selectbox(
            "Empreendimento",
            ["Todos"] + sorted(
                df_indicador["Empreendimento"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="grafico_mes_empreendimento"
        )

    # ======================================
    # APLICAR FILTROS
    # ======================================

    df_filtrado = df_indicador.copy()

    if filtro_ano != "Todos":

        df_filtrado = df_filtrado[
            df_filtrado["Ano"].astype(str)
            == filtro_ano
        ]

    if filtro_agrupamento != "Todos":

        df_filtrado = df_filtrado[
            df_filtrado["Agrupamento"].astype(str)
            == filtro_agrupamento
        ]

    if filtro_contratada != "Todas":

        df_filtrado = df_filtrado[
            df_filtrado["Contratada"].astype(str)
            == filtro_contratada
        ]

    if filtro_empreendimento != "Todos":

        df_filtrado = df_filtrado[
            df_filtrado["Empreendimento"].astype(str)
            == filtro_empreendimento
        ]

    # ======================================
    # CRIAR MÊS
    # ======================================

    df_filtrado["Mes"] = (
        df_filtrado["Data de Entrada"]
        .dt.month
    )

    # ======================================
    # SOMA DE QTD. DOCS POR MÊS E STATUS
    # ======================================

    resumo = (
        df_filtrado
        .groupby(
            ["Mes", "Status GRD"],
            as_index=False
        )["Qtd. Docs"]
        .sum()
    )

    # ======================================
    # TABELA AUXILIAR
    # ======================================

    meses = pd.DataFrame({
        "Mes": range(1, 13),
        "Mês": [
            "Jan", "Fev", "Mar", "Abr",
            "Mai", "Jun", "Jul", "Ago",
            "Set", "Out", "Nov", "Dez"
        ]
    })

    prazo = (
        resumo[
            resumo["Status GRD"]
            .astype(str)
            .str.upper()
            == "NO PRAZO"
        ][["Mes", "Qtd. Docs"]]
        .rename(
            columns={
                "Qtd. Docs": "No Prazo"
            }
        )
    )

    fora = (
        resumo[
            resumo["Status GRD"]
            .astype(str)
            .str.upper()
            == "FORA DO PRAZO"
        ][["Mes", "Qtd. Docs"]]
        .rename(
            columns={
                "Qtd. Docs": "Fora do Prazo"
            }
        )
    )

    dados_grafico = meses.merge(
        prazo,
        on="Mes",
        how="left"
    )

    dados_grafico = dados_grafico.merge(
        fora,
        on="Mes",
        how="left"
    )

    dados_grafico = dados_grafico.fillna(0)

    dados_grafico["Total"] = (
    dados_grafico["No Prazo"]
    + dados_grafico["Fora do Prazo"]
    )

    # ======================================
    # GRÁFICO
    # ======================================

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=dados_grafico["Mês"],
            y=dados_grafico["No Prazo"],
            name="No Prazo",
            marker_color="#00B050"
        )
    )

    fig.add_trace(
        go.Bar(
            x=dados_grafico["Mês"],
            y=dados_grafico["Fora do Prazo"],
            name="Fora do Prazo",
            marker_color="#FF0000"
        )
    )

    fig.add_trace(
        go.Scatter(
            x=dados_grafico["Mês"],
            y=dados_grafico["Total"],
            name="Total",
            mode="lines+markers",
            line=dict(
                color="#003366",
                width=2
            ),
            marker=dict(size=8),
            hovertemplate=
            "<b>%{x}</b><br>" +
            "Total: %{y}<extra></extra>"
        )
    )

    fig.update_layout(
        title="Documentos Tramitados por Mês de Entrada",
        barmode="stack",
        height=600,
        hovermode="x unified",
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_title="Mês",
        yaxis_title="Quantidade de Documentos",
        legend_title=""
    )

    fig.update_xaxes(
        showgrid=False
    )

    fig.update_yaxes(
        showgrid=True,
        gridcolor="#D9D9D9"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

# ==========================================
# INDICADOR IAPE
# ==========================================

st.markdown("---")
st.markdown("### 📊 IAPE - Indicador de Análise de Projetos de SEs, LTs e Documentos de Equipamentos")

df_iape = df_dashboard_base.copy()

if not df_iape.empty:

    # ======================================
    # FILTROS INDEPENDENTES
    # ======================================

    fi1, fi2, fi3, fi4 = st.columns(4)

    with fi1:

        filtro_ano_iape = st.selectbox(
            "Ano",
            ["Todos"] + sorted(
                df_iape["Ano"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="filtro_ano_iape"
        )

    with fi2:

        filtro_agrupamento_iape = st.selectbox(
            "Agrupamento",
            ["Todos"] + sorted(
                df_iape["Agrupamento"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="filtro_agrupamento_iape"
        )

    with fi3:

        filtro_contratada_iape = st.selectbox(
            "Contratada",
            ["Todas"] + sorted(
                df_iape["Contratada"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="filtro_contratada_iape"
        )

    with fi4:

        filtro_empreendimento_iape = st.selectbox(
            "Empreendimento",
            ["Todos"] + sorted(
                df_iape["Empreendimento"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="filtro_empreendimento_iape"
        )

    # ======================================
    # APLICAR FILTROS
    # ======================================

    if filtro_ano_iape != "Todos":

        df_iape = df_iape[
            df_iape["Ano"].astype(str)
            == filtro_ano_iape
        ]

    if filtro_agrupamento_iape != "Todos":

        df_iape = df_iape[
            df_iape["Agrupamento"].astype(str)
            == filtro_agrupamento_iape
        ]

    if filtro_contratada_iape != "Todas":

        df_iape = df_iape[
            df_iape["Contratada"].astype(str)
            == filtro_contratada_iape
        ]

    if filtro_empreendimento_iape != "Todos":

        df_iape = df_iape[
            df_iape["Empreendimento"].astype(str)
            == filtro_empreendimento_iape
        ]

    # ======================================
    # PREPARAÇÃO DOS DADOS
    # ======================================

    df_iape["Data de Entrada"] = pd.to_datetime(
        df_iape["Data de Entrada"],
        dayfirst=True,
        errors="coerce"
    )

    df_iape["Qtd. Docs"] = pd.to_numeric(
        df_iape["Qtd. Docs"],
        errors="coerce"
    ).fillna(0)

    df_iape = df_iape.dropna(
        subset=["Data de Entrada"]
    )

    df_iape["Mes"] = (
        df_iape["Data de Entrada"]
        .dt.month
    )

    # ======================================
    # CÁLCULO MENSAL
    # ======================================

    resultado = []

    for mes in range(1, 13):

        df_mes = df_iape[
            df_iape["Mes"] == mes
        ]

        total_docs = df_mes["Qtd. Docs"].sum()

        docs_prazo = df_mes[
            df_mes["Status GRD"]
            .astype(str)
            .str.upper()
            .str.strip()
            == "NO PRAZO"
        ]["Qtd. Docs"].sum()

        if total_docs == 0:

            resultado.append({
                "Mes": mes,
                "IAPE": 0
            })

            continue

        # ==================================
        # FATOR FV
        # ==================================

        if total_docs <= 400:
            fv = 1.00

        elif total_docs <= 600:
            fv = 1.05

        else:
            fv = 1.10

        iape = (
            docs_prazo /
            total_docs
        ) * fv * 100

        resultado.append({
            "Mes": mes,
            "IAPE": round(iape, 2)
        })

    df_resultado = pd.DataFrame(resultado)

    nomes_meses = {
        1: "Jan",
        2: "Fev",
        3: "Mar",
        4: "Abr",
        5: "Mai",
        6: "Jun",
        7: "Jul",
        8: "Ago",
        9: "Set",
        10: "Out",
        11: "Nov",
        12: "Dez"
    }

    df_resultado["Mês"] = (
        df_resultado["Mes"]
        .map(nomes_meses)
    )

    # ======================================
    # IAPE ACUMULADO
    # ======================================

    total_periodo = df_iape["Qtd. Docs"].sum()

    docs_prazo_periodo = df_iape[
        df_iape["Status GRD"]
        .astype(str)
        .str.upper()
        .str.strip()
        == "NO PRAZO"
    ]["Qtd. Docs"].sum()

    if total_periodo > 0:

        if total_periodo <= 400:
            fv_periodo = 1.00

        elif total_periodo <= 600:
            fv_periodo = 1.05

        else:
            fv_periodo = 1.10

        iape_acumulado = (
            docs_prazo_periodo /
            total_periodo
        ) * fv_periodo * 100

    else:

        iape_acumulado = 0

    col1, col2 = st.columns(2)

    with col1:
        st.metric(
            "IAPE Acumulado",
            f"{iape_acumulado:.0f}%"
        )

    with col2:
        st.metric(
            "Meta",
            "95%"
        )

    # ======================================
    # GRÁFICO
    # ======================================

    colors = []

    for valor in df_resultado["IAPE"]:

        if valor >= 95:
            colors.append("#00B050")  # Verde

        elif valor >= 90:
            colors.append("#FFC000")  # Amarelo

        else:
            colors.append("#C00000")  # Vermelho

    fig_iape = go.Figure()

    fig_iape.add_trace(
        go.Bar(
            x=df_resultado["Mês"],
            y=df_resultado["IAPE"],
            showlegend=False,
            marker_color=colors,
            text=[
                f"{valor:.1f}%"
                for valor in df_resultado["IAPE"]
            ],
            textposition="outside"
        )
    )

    fig_iape.add_trace(
        go.Scatter(
            x=df_resultado["Mês"],
            y=[95] * 12,
            mode="lines",
            name="Meta 95%",
            line=dict(
                color="red",
                width=3,
                dash="dash"
            )
        )
    )

    fig_iape.update_layout(
        title="Indicador IAPE por Mês",
        height=600,
        hovermode="x unified",
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_title="Mês",
        yaxis_title="IAPE (%)",
        yaxis=dict(
            range=[0, 110]
        ),
        legend=dict(
            orientation="h",
            y=1.05
        )
    )

    fig_iape.update_xaxes(
        showgrid=False
    )

    fig_iape.update_yaxes(
        showgrid=True,
        gridcolor="#D9D9D9"
    )

    st.plotly_chart(
        fig_iape,
        use_container_width=True
    )
else:

    st.info(
        "Não existem dados para exibição."
    )
        #FIM