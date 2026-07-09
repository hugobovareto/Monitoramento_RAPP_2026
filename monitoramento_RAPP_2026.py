'''
Objetivo: gerar 2 bases de dados em formato de planilha para serem usadas como fonte de dados para a aplicação em Google Apps Script para monitoramento dos estudantes em RAPP.
1ª base: Estudantes (base geral de todos componentes de todos estudantes em RAPP)
Colunas:
Matrícula (identificador único do estudante);
Nome;
Componente;
Cod_Inep;
Escola;
Série;
Direc;
Etapa_Ensino;
Enturmação;
Rendimento;
Tempo_Prova;
Situação.


2ª base: Informações inseridas pelo usuário (informação vai ser inserida por estudante e não por componente. Tavlez fosse bom dividir algumas informações para serem por componente, mas outras ficariam repetitivas, então fiz por estudante)
Colunas:
Matrícula (identificador único do estudante);
Aluno contatado;
Data_prevista;
Comentários.

São usadas as seguintes fontes de informação:
- Base geral de estudantes em RAPP (advinda do GPD e após tratamentos);
- Redash das avaliações de RAPP;
- Relatórios de Acompanhamento de Turmas e Progressão Parcial. 

1) Base geral de estudantes em RAPP:
Considerar somente com os componentes da Formação Geral Básica da BNCC (sem Ensino Religioso):
- Arte;
- Biologia;
- Ciências;
- Educação Física;
- Filosofia;
- Física;
- Geografia;
- História;
- Língua Espanhola;
- Língua Portuguesa;
- Língua Inglesa;
- Matemática;
- Química;
- Sociologia.

2) Redash das avaliações de RAPP:
Para informações de rendimento e tempo de prova, por componente para os estudantes.
O identificador único do estudante presente é o email, do qual será tirado a matrícula.
Exclusão de provas com tempo = 0.


3) Relatórios de Acompanhamento de Turmas e Progressão Parcial.
Para saber os estudantes enturmados e alguns que já estão aprovados (mesmo sem fazer a prova Plurall).


Tratamentos:
Estudante é considerado enturmado se aparece no Relatório de Acompanhamento de Turmas e Progressão Parcial.
Componente está aprovado se a nota (presenter no redash for >= 60).
Estudante com 'SITUAÇÃO FINAL' = APROVADO no Relatório de Acompanhamento de Turmas e Progressão Parcial é considerado aprovado, mesmo que não tenha feito a prova Plurall.
A nota desse estudante será a nota em 'MÉDIA FINAL'.


As colunas abaixo serão preenchidas pelo usuário da aplicação em Google Apps Script.
Aluno contatado;
Data_prevista;
Comentários.


'''
# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import openpyxl
import re

# Carregar os dados gerais dos estudantes em RAPP 
df_rapp = pd.read_excel(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026\20260625_GERAL_analises_RAPP.xlsx", sheet_name="Base RAPP")


# Considerar somente os componentes da Formação Geral Básica da BNCC (sem Ensino Religioso):
componentes_bncc = ['Arte',
                     'Biologia',
                     'Ciências',
                     'Educação Física',
                     'Filosofia',
                     'Física',
                     'Geografia',
                     'História',
                     'Língua Espanhola',
                     'Língua Portuguesa',
                     'Língua Inglesa',
                     'Matemática',
                     'Química',
                     'Sociologia']

df_rapp = df_rapp[df_rapp['COMPONENTE CURRICULAR'].isin(componentes_bncc)]

# Manunteção somente das colunas de interesse
df_rapp = df_rapp[['MATRÍCULA', 'NOME', 'COMPONENTE CURRICULAR', 'INEP ESCOLA', 'ESCOLA', 'SÉRIE', 'DIREC', 'ETAPA_RESUMIDA']]


# Carregar dados do Redash (para ter nota e tempo de prova do estudante)
df_redash = pd.read_csv(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026\SEC-RN_-_Rendimento_e_participação_dos_alunos_p_provas_-_RPP_-_Avaliações_em_andamento_2025_12_19.csv")

# Excluir valores que o tempo de prova foi 0 (zero) ou nulo
df_redash = df_redash[
    df_redash['tempo de prova'].notna() &
    (df_redash['tempo de prova'].str.strip() != '') &
    (df_redash['tempo de prova'] != '00:00:00')
]

# Criar a coluna 'MATRÍCULA' no df_redash, extraindo a matrícula do email do estudante
df_redash['MATRÍCULA'] = df_redash['email_aluno'].str.split('@').str[0]

# De acordo com a matrícula (extraída do email) e componente, juntar as informações de nota e tempo de prova do Redash com a base geral de estudantes em RAPP
# Selecionar as colunas de interesse do df_redash
df_redash_merge = df_redash[['MATRÍCULA', 'prova', 'rendimento (%)', 'tempo de prova']]

# Converter MATRÍCULA para string em ambos os dataframes para garantir que a junção funcione corretamente
df_rapp["MATRÍCULA"] = (
    df_rapp["MATRÍCULA"]
    .astype("string")
    .str.strip()
)

df_redash_merge["MATRÍCULA"] = (
    df_redash_merge["MATRÍCULA"]
    .astype("string")
    .str.strip()
)

# Merge
df_merged = df_rapp.merge(
    df_redash_merge,
    left_on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    right_on=['MATRÍCULA', 'prova'],
    how='left'
)

# Remover as colunas duplicadas vindas do Redash
df_merged = df_merged.drop(columns=['prova'])


# Adicionar coluna de Situação de acordo com a nota do estudante
# >=60: Aprovado;
# < 60: Reprovado;
# NaN: Não Avaliado.
df_merged['Situação'] = df_merged['rendimento (%)'].apply(
    lambda x: 'Não Avaliado' if pd.isna(x)
    else 'Aprovado' if x >= 60
    else 'Reprovado'
)

# Carregar dados dos Relatórios de Acompanhamento de Turmas e Progressão Parcial para saber enturmação e nota dos aprovados
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Enturmados_RAPP_2026"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 2 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=2)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_enturmados = pd.concat(dfs, ignore_index=True)

# Adicionar coluna 'Enturmação' no df_merged, com valor 'Sim' se o estudante e componente estiver no df_enturmados, caso contrário 'Não'
# Selecionar apenas as chaves do df_enturmados
df_enturmados_merge = (
    df_enturmados[['MATRÍCULA', 'COMPONENTE CURRICULAR']]
    .drop_duplicates()
    .assign(Enturmação='Sim')
)

# Converter MATRÍCULA para string em ambos os dataframes para garantir que a junção funcione corretamente
df_merged['MATRÍCULA'] = (
    df_merged['MATRÍCULA']
    .astype(str)
    .str.strip()
)

df_enturmados['MATRÍCULA'] = (
    df_enturmados['MATRÍCULA']
    .astype(str)
    .str.strip()
)

df_enturmados_merge['MATRÍCULA'] = (
    df_enturmados_merge['MATRÍCULA']
    .astype(str)
    .str.strip()
)

# Merge
df_merged = df_merged.merge(
    df_enturmados_merge,
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='left'
)

# Preencher os que não foram encontrados
df_merged['Enturmação'] = df_merged['Enturmação'].fillna('Não')


# Estudante com 'SITUAÇÃO FINAL' = APROVADO no df_enturmados é considerado aprovado, mesmo que não tenha feito a prova Plurall.
# A nota desse estudante será a nota que está em 'MÉDIA FINAL'.
# Trazer informações do df_enturmados
df_merged = df_merged.merge(
    df_enturmados[
        ['MATRÍCULA', 'COMPONENTE CURRICULAR', 'SITUAÇÃO FINAL', 'MÉDIA FINAL']
    ].drop_duplicates(),
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='left'
)

# Máscara dos aprovados
mask = df_merged['SITUAÇÃO FINAL'].eq('APROVADO')

# Converter 'MÉDIA FINAL' para numérico
df_merged['MÉDIA FINAL'] = (
    df_merged['MÉDIA FINAL']
    .str.replace(',', '.', regex=False)
)

df_merged['MÉDIA FINAL'] = pd.to_numeric(
    df_merged['MÉDIA FINAL'],
    errors='coerce'
)

# Multiplicar por 10 para ficar na mesma escala de 'rendimento (%)'
df_merged['MÉDIA FINAL'] = df_merged['MÉDIA FINAL'] * 10


# Atualizar nota e situação do df_merged para os estudantes aprovados no df_enturmados
df_merged.loc[mask, 'rendimento (%)'] = df_merged.loc[mask, 'MÉDIA FINAL']
df_merged.loc[mask, 'Situação'] = 'Aprovado'

# Excluir as colunas 'SITUAÇÃO FINAL' e 'MÉDIA FINAL' do df_merged
df_merged = df_merged.drop(columns=['SITUAÇÃO FINAL', 'MÉDIA FINAL'])

# Exportar a base final em Excel para usar na aplicação em Google Apps Script
df_merged.to_excel(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026\20260709_Monitoramento_RAPP.xlsx", index=False)


































