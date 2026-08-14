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


Excluir duplicatas do df_enturmados, considerando as colunas 'MATRÍCULA' e 'COMPONENTE CURRICULAR'.
Seguir o ordenamento de preferência:
# SITUAÇÃO FINAL = APROVADO;
# maior nota em MÉDIA FINAL.


Tratamentos:
Estudante é considerado enturmado se aparece no Relatório de Acompanhamento de Turmas e Progressão Parcial.
Componente está aprovado se a nota (presente no redash for >= 60).
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

df_enturmados = df_enturmados[df_enturmados['COMPONENTE CURRICULAR'].isin(componentes_bncc)]

# Criar a coluna 'ETAPA_RESUMIDA' no df_enturmados, com valores 'Ens. Fund. - Anos Finais' e 'Ensino Médio', de acordo com a Série
# Criar a coluna 'ETAPA_RESUMIDA' a partir da SÉRIE
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais',
    'TURMA I (6° E 7° ANOS)': 'Ens. Fund. - Anos Finais',
    'TURMA II (8° E 9° ANOS)': 'Ens. Fund. - Anos Finais',
    '1º Período (1ª Série)': 'Ensino Médio',
    '2º Período (2ª Série)': 'Ensino Médio',
    '3º Período (3ª Série)': 'Ensino Médio',
    '3° PERÍODO': 'Ensino Médio',
    '1º MÓDULO': 'Ensino Médio',
    '2º MODULO': 'Ensino Médio',
    '3º MÓDULO': 'Ensino Médio',
    'BLOCO B': 'Ensino Médio',
    'BLOCO C': 'Ensino Médio',
    '5º PERÍODO (8° E 9° ANO - ANOS FINAIS)': 'Ens. Fund. - Anos Finais',
    '4º PERÍODO (6° E 7° ANO - ANOS FINAIS)': 'Ens. Fund. - Anos Finais',
    '2º SEMESTRE': 'Ensino Médio',
    'UNICA': 'Ens. Fund. - Anos Finais'
}

df_enturmados['ETAPA_RESUMIDA'] = df_enturmados['SÉRIE'].map(mapeamento_etapa)


# Excluir duplicatas do df_enturmados, considerando as colunas 'MATRÍCULA' e 'COMPONENTE CURRICULAR'.
# Seguir o ordenamento de preferência:
# SITUAÇÃO FINAL = APROVADO;
# maior nota em MÉDIA FINAL.

# Garante que MÉDIA FINAL é numérica
df_enturmados['MÉDIA FINAL'] = pd.to_numeric(
    df_enturmados['MÉDIA FINAL'].str.replace(',', '.', regex=False),
    errors='coerce'
)

# Cria a prioridade (1 = APROVADO, 0 = demais)
df_enturmados['_prioridade'] = (
    df_enturmados['SITUAÇÃO FINAL']
    .eq('APROVADO')
    .astype(int)
)

# Ordena pelos critérios
df_enturmados = (
    df_enturmados
    .sort_values(
        by=['_prioridade', 'MÉDIA FINAL'],
        ascending=[False, False]
    )
    .drop_duplicates(
        subset=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
        keep='first'
    )
    .drop(columns='_prioridade')
    .reset_index(drop=True)
)


# Converter MATRÍCULA para string em ambos os dataframes para garantir que a junção funcione corretamente
df_rapp['MATRÍCULA'] = (
    df_rapp['MATRÍCULA']
    .astype(str)
    .str.strip()
)

df_enturmados['MATRÍCULA'] = (
    df_enturmados['MATRÍCULA']
    .astype(str)
    .str.strip()
)


# Merge entre o df_rapp e df_enturmados, considerando a chave 'MATRÍCULA' e 'COMPONENTE CURRICULAR'
# Merge externo (outer) para manter todos os registros de ambos os dataframes
df_final = df_rapp.merge(
    df_enturmados,
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='outer',
    suffixes=('', '_ent')
)

# Identifica as linhas que vieram somente do df_enturmados
novas_linhas = df_final['NOME'].isna()

# Preenche as colunas do df_rapp com as informações equivalentes do df_enturmados
df_final.loc[novas_linhas, 'NOME'] = df_final.loc[novas_linhas, 'ESTUDANTE']
df_final.loc[novas_linhas, 'INEP ESCOLA'] = df_final.loc[novas_linhas, 'INEP ESCOLA_ent']
df_final.loc[novas_linhas, 'ESCOLA'] = df_final.loc[novas_linhas, 'ESCOLA_ent']
df_final.loc[novas_linhas, 'SÉRIE'] = df_final.loc[novas_linhas, 'SÉRIE_ent']
df_final.loc[novas_linhas, 'DIREC'] = df_final.loc[novas_linhas, 'DIREC_ent']
df_final.loc[novas_linhas, 'ETAPA_RESUMIDA'] = df_final.loc[novas_linhas, 'ETAPA_RESUMIDA_ent']

# Situação para os novos registros
df_final.loc[novas_linhas, 'Situação'] = 'Não Avaliado'

# Mantém somente as colunas do df_rapp
df_final = df_final[df_rapp.columns.tolist() + ['Situação']]


# Adicionar coluna 'Enturmação' no df_final, com valor 'Sim' se o estudante e componente estiver no df_enturmados, caso contrário 'Não'
# Selecionar apenas as chaves do df_enturmados
df_enturmados_merge = (
    df_enturmados[['MATRÍCULA', 'COMPONENTE CURRICULAR']]
    .drop_duplicates()
    .assign(Enturmação='Sim')
)

# Converter MATRÍCULA para string em ambos os dataframes para garantir que a junção funcione corretamente
df_final['MATRÍCULA'] = (
    df_final['MATRÍCULA']
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
df_final = df_final.merge(
    df_enturmados_merge,
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='left'
)

# Preencher os que não foram encontrados
df_final['Enturmação'] = df_final['Enturmação'].fillna('Não')


# Estudante com 'SITUAÇÃO FINAL' = APROVADO no df_enturmados é considerado aprovado, mesmo que não tenha feito a prova Plurall.
# A nota desse estudante será a nota que está em 'MÉDIA FINAL'.
# Trazer informações do df_enturmados
df_final = df_final.merge(
    df_enturmados[
        ['MATRÍCULA', 'COMPONENTE CURRICULAR', 'SITUAÇÃO FINAL', 'MÉDIA FINAL']
    ].drop_duplicates(),
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='left'
)


# Máscara dos aprovados
mask = df_final['SITUAÇÃO FINAL'].eq('APROVADO')


# Converter 'MÉDIA FINAL' para numérico
df_final['MÉDIA FINAL'] = pd.to_numeric(
    df_final['MÉDIA FINAL'].astype(str).str.replace(',', '.', regex=False),
    errors='coerce'
)

# Multiplicar por 10 para ficar na mesma escala de 'rendimento (%)'
df_final['MÉDIA FINAL'] = df_final['MÉDIA FINAL'] * 10


# Atualizar nota e situação do df_final para os estudantes aprovados no df_enturmados
df_final.loc[mask, 'rendimento (%)'] = df_final.loc[mask, 'MÉDIA FINAL']
df_final.loc[mask, 'Situação'] = 'Aprovado'


# Excluir as colunas 'SITUAÇÃO FINAL' e 'MÉDIA FINAL' do df_merged
df_final = df_final.drop(columns=['SITUAÇÃO FINAL', 'MÉDIA FINAL'])


# Carregar dados do Redash (para ter nota e tempo de prova do estudante)
df_redash = pd.read_csv(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026\SEEC-RN_-_Rendimento_e_participação_dos_alunos_p_provas_-_RAPP_-_Avaliações_em_andamento_2026_08_13.csv")

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
df_final["MATRÍCULA"] = (
    df_final["MATRÍCULA"]
    .astype("string")
    .str.strip()
)

df_redash_merge["MATRÍCULA"] = (
    df_redash_merge["MATRÍCULA"]
    .astype("string")
    .str.strip()
)


# Merge trazendo o rendimento do Redash com um nome temporário
df_merged = df_final.merge(
    df_redash_merge[['MATRÍCULA', 'prova', 'rendimento (%)', 'tempo de prova']],
    left_on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    right_on=['MATRÍCULA', 'prova'],
    how='left',
    suffixes=('', '_redash')
)

# Substitui o rendimento do df_final pelo rendimento do Redash, somente quando houver valor no Redash
df_merged['rendimento (%)'] = df_merged['rendimento (%)_redash'].fillna(
    df_merged['rendimento (%)']
)

# Remove as colunas auxiliares
df_merged = df_merged.drop(
    columns=['prova', 'rendimento (%)_redash']
)


# Adiciona valores em Situação de acordo com o rendimento, exceto para os casos já aprovados vindos da Enturmação
condicoes = [
    df_merged['Situação'].eq('Aprovado'),
    df_merged['rendimento (%)'].isna(),
    df_merged['rendimento (%)'].ge(60)
]

valores = [
    'Aprovado',
    'Não Avaliado',
    'Aprovado'
]

df_merged['Situação'] = np.select(
    condicoes,
    valores,
    default='Não Aprovado'
)


# Reorganizar a ordem das colunas
df_merged = df_merged[
    ['MATRÍCULA',
     'NOME', 
     'COMPONENTE CURRICULAR',
     'INEP ESCOLA',
     'ESCOLA',
     'SÉRIE',
     'DIREC',
     'ETAPA_RESUMIDA',
     'rendimento (%)',
     'tempo de prova',
     'Situação',
     'Enturmação']
]

# Exportar a base final em Excel para usar na aplicação em Google Apps Script
df_merged.to_excel(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026\20260813_Monitoramento_RAPP.xlsx", index=False)











































##################################################################################
'''
MONITORAMENTO DE ENTURMAÇÃO DOS ESTUDANTES EM RAPP
Gerar uma planilha com quantitativo de estudantes (não considera granulação por componente).
Colunas:
DIREC;
Escolas iniciais;
Alunos iniciais;
Alunos enturmados;
Alunos aprovados;
Total alunos (inicial e enturmados);
% enturmados;
# faltam enturmar.

Para além do quadro acima, a planilha terá outras abas:
Total (lista com todos os componentes de estudantes em RAPP inicial + enturmados);
Inicial (lista com componentes iniciais de estudantes em RAPP, de acordo com base do GPD e tratamentos);
Enturmados (lista com componentes enturmados, de acordo com Relatórios de Acompanhamento de Turmas e Progressão Parcial);
Quadro (dados expostos acima)

Se o estudante aparece nos Relatórios de Acompanhamento de Turmas e Progressão Parcial, ele é considerado enturmado.
Os valores tidos como 'inicial' são os referentes à base geral de estudantes em RAPP (advinda do GPD e após tratamentos).

Tratamentos:
Estudante é considerado enturmado se aparece no Relatório de Acompanhamento de Turmas e Progressão Parcial.
Estudante com 'SITUAÇÃO FINAL' = APROVADO no Relatório de Acompanhamento de Turmas e Progressão Parcial é considerado aprovado, mesmo que não tenha feito a prova Plurall.
A nota desse estudante será a nota em 'MÉDIA FINAL'.

Excluir duplicatas do df_enturmados, considerando as colunas 'MATRÍCULA' e 'COMPONENTE CURRICULAR'.
Seguir o ordenamento de preferência:
# SITUAÇÃO FINAL = APROVADO;
# maior nota em MÉDIA FINAL.


Os dados incluem estudantes de regular e EPT.
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


# Manunteção somente das colunas de interesse
df_rapp = df_rapp[['MATRÍCULA', 'NOME', 'COMPONENTE CURRICULAR', 'INEP ESCOLA', 'ESCOLA', 'SÉRIE', 'DIREC', 'ETAPA ENSINO']]

# Carregar dados dos Relatórios de Acompanhamento de Turmas e Progressão Parcial para saber quantitativo de enturmados e aprovados
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


# Excluir duplicatas do df_enturmados, considerando as colunas 'MATRÍCULA' e 'COMPONENTE CURRICULAR'.
# Seguir o ordenamento de preferência:
# SITUAÇÃO FINAL = APROVADO;
# maior nota em MÉDIA FINAL.

# Criar uma cópia do dataframe original
df_enturmados_final = df_enturmados.copy()


# Garante que MÉDIA FINAL é numérica
df_enturmados_final['MÉDIA FINAL'] = pd.to_numeric(
    df_enturmados_final['MÉDIA FINAL'],
    errors='coerce'
)

# Cria a prioridade (1 = APROVADO, 0 = demais)
df_enturmados_final['_prioridade'] = (
    df_enturmados_final['SITUAÇÃO FINAL']
    .eq('APROVADO')
    .astype(int)
)

# Ordena pelos critérios
df_enturmados_final = (
    df_enturmados_final
    .sort_values(
        by=['_prioridade', 'MÉDIA FINAL'],
        ascending=[False, False]
    )
    .drop_duplicates(
        subset=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
        keep='first'
    )
    .drop(columns='_prioridade')
    .reset_index(drop=True)
)


# Converter MATRÍCULA para string em ambos os dataframes para garantir que a junção funcione corretamente
df_rapp['MATRÍCULA'] = (
    df_rapp['MATRÍCULA']
    .astype(str)
    .str.strip()
)

df_enturmados_final['MATRÍCULA'] = (
    df_enturmados_final['MATRÍCULA']
    .astype(str)
    .str.strip()
)


# Merge entre o df_rapp e df_enturmados_final, considerando a chave 'MATRÍCULA' e 'COMPONENTE CURRICULAR'
# Merge externo (outer) para manter todos os registros de ambos os dataframes
df_total = df_rapp.merge(
    df_enturmados_final,
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='outer',
    suffixes=('', '_ent')
)

# Identifica as linhas que vieram somente do df_enturmados
novas_linhas = df_total['NOME'].isna()

# Preenche as colunas do df_merged com as informações equivalentes do df_enturmados
df_total.loc[novas_linhas, 'NOME'] = df_total.loc[novas_linhas, 'ESTUDANTE']
df_total.loc[novas_linhas, 'INEP ESCOLA'] = df_total.loc[novas_linhas, 'INEP ESCOLA_ent']
df_total.loc[novas_linhas, 'ESCOLA'] = df_total.loc[novas_linhas, 'ESCOLA_ent']
df_total.loc[novas_linhas, 'SÉRIE'] = df_total.loc[novas_linhas, 'SÉRIE_ent']
df_total.loc[novas_linhas, 'DIREC'] = df_total.loc[novas_linhas, 'DIREC_ent']

# Situação para os novos registros
df_total.loc[novas_linhas, 'Situação'] = 'Não Avaliado'

# Mantém somente as colunas do df_rapp
df_total = df_total[df_rapp.columns]


# Adicionar coluna 'Enturmação' no df_total, com valor 'Sim' se o estudante e componente estiver no df_enturmados_final, caso contrário 'Não'
# Selecionar apenas as chaves do df_enturmados
df_enturmados_merge = (
    df_enturmados_final[['MATRÍCULA', 'COMPONENTE CURRICULAR']]
    .drop_duplicates()
    .assign(Enturmação='Sim')
)

# Converter MATRÍCULA para string em ambos os dataframes para garantir que a junção funcione corretamente
df_total['MATRÍCULA'] = (
    df_total['MATRÍCULA']
    .astype(str)
    .str.strip()
)

df_enturmados_final['MATRÍCULA'] = (
    df_enturmados_final['MATRÍCULA']
    .astype(str)
    .str.strip()
)

df_enturmados_merge['MATRÍCULA'] = (
    df_enturmados_merge['MATRÍCULA']
    .astype(str)
    .str.strip()
)

# Merge
df_total = df_total.merge(
    df_enturmados_merge,
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='left'
)

# Preencher os que não foram encontrados
df_total['Enturmação'] = df_total['Enturmação'].fillna('Não')


# Estudante com 'SITUAÇÃO FINAL' = APROVADO no df_enturmados_final é considerado aprovado, mesmo que não tenha feito a prova Plurall.
# A nota desse estudante será a nota que está em 'MÉDIA FINAL'.
# Trazer informações do df_enturmados
df_total = df_total.merge(
    df_enturmados_final[
        ['MATRÍCULA', 'COMPONENTE CURRICULAR', 'SITUAÇÃO FINAL', 'MÉDIA FINAL']
    ].drop_duplicates(),
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='left'
)

print(df_total['MÉDIA FINAL'].dtype)
# Máscara dos aprovados
mask = df_total['SITUAÇÃO FINAL'].eq('APROVADO')

# Converter 'MÉDIA FINAL' para numérico
df_total['MÉDIA FINAL'] = pd.to_numeric(
    df_total['MÉDIA FINAL'],
    errors='coerce'
)

# Multiplicar por 10 para ficar na mesma escala de 'rendimento (%)'
df_total['MÉDIA FINAL'] = df_total['MÉDIA FINAL'] * 10


# Atualizar nota e situação do df_final para os estudantes aprovados no df_enturmados
df_total.loc[mask, 'rendimento (%)'] = df_total.loc[mask, 'MÉDIA FINAL']
df_total.loc[mask, 'Situação'] = 'Aprovado'

# Excluir as colunas 'SITUAÇÃO FINAL' e 'MÉDIA FINAL' do df_merged
df_total = df_total.drop(columns=['SITUAÇÃO FINAL', 'MÉDIA FINAL'])


# Excluir coluna 'rendimento (%)' do df_total
df_total = df_total.drop(columns=['rendimento (%)'])


# Criar um arquivo em excel com as abas de df_total; df_rapp; df_enturmados e uma tabela com informações por DIREC
# Colunas da aba Tabela: DIREC; Escolas iniciais; Alunos iniciais; Alunos enturmados; Alunos aprovados; Total alunos (inicial e enturmados); % enturmados; # faltam enturmar.
'''
Cada coluna tem os valores:
DIREC: uma linha para cada valor presente na coluna df_total['DIREC']

Escolas Iniciais: total de valores únicos de 'INEP ESCOLA' para cada 'DIREC' no df_rapp;

Alunos Iniciais: total de valores únicos de 'MATRÍCULA' para cada 'DIREC' no df_rapp;

Alunos Enturmados: total de valores únicos de 'MATRÍCULA' para cada 'DIREC' no df_total que possuem a coluna 'Enturmação' = Sim;

Alunos Aprovados: total de valores únicos de 'MATRÍCULA' para cada 'DIREC' no df_total que possuem a coluna 'Situação' = Aprovado;

Total Alunos (inicial e enturmados): total de valores únicos de 'MATRÍCULA' para cada 'DIREC' no df_total 

% Enturmados: Alunos Enturmados / Total Alunos (inicial e enturmados)

Faltam Enturmar: Total Alunos (inicial e enturmados) menos Alunos Enturmados
'''

# Lista de DIRECs
resumo_direc = pd.DataFrame({
    'DIREC': sorted(df_total['DIREC'].dropna().unique())
})

# Escolas Iniciais (df_rapp)
escolas_iniciais = (
    df_rapp.groupby('DIREC')['INEP ESCOLA']
    .nunique()
    .rename('Escolas Iniciais')
)

# Alunos Iniciais (df_rapp)
alunos_iniciais = (
    df_rapp.groupby('DIREC')['MATRÍCULA']
    .nunique()
    .rename('Alunos Iniciais')
)

# Alunos Enturmados (df_total)
alunos_enturmados = (
    df_total[df_total['Enturmação'].eq('Sim')]
    .groupby('DIREC')['MATRÍCULA']
    .nunique()
    .rename('Alunos Enturmados')
)

# Alunos Aprovados (df_total)
alunos_aprovados = (
    df_total[df_total['Situação'].eq('Aprovado')]
    .groupby('DIREC')['MATRÍCULA']
    .nunique()
    .rename('Alunos Aprovados')
)

# Total de alunos (df_total)
total_alunos = (
    df_total.groupby('DIREC')['MATRÍCULA']
    .nunique()
    .rename('Total Alunos')
)

# Juntar tudo
resumo_direc = (
    resumo_direc
    .merge(escolas_iniciais, on='DIREC', how='left')
    .merge(alunos_iniciais, on='DIREC', how='left')
    .merge(alunos_enturmados, on='DIREC', how='left')
    .merge(alunos_aprovados, on='DIREC', how='left')
    .merge(total_alunos, on='DIREC', how='left')
)

# Preencher NaN com 0 nas colunas numéricas
cols = [
    'Escolas Iniciais',
    'Alunos Iniciais',
    'Alunos Enturmados',
    'Alunos Aprovados',
    'Total Alunos'
]

resumo_direc[cols] = resumo_direc[cols].fillna(0).astype(int)

# Indicadores finais
resumo_direc['% Enturmados'] = (
    resumo_direc['Alunos Enturmados']
    / resumo_direc['Total Alunos']
)

resumo_direc['Faltam Enturmar'] = (
    resumo_direc['Total Alunos']
    - resumo_direc['Alunos Enturmados']
)

# (Opcional) Formatar o percentual
resumo_direc['% Enturmados'] = (
    resumo_direc['% Enturmados'] * 100
).round(1)


# Salvar em Excel os dataframes e tabela com informações por DIREC
with pd.ExcelWriter(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026\20260723_Enturmação.xlsx") as writer:
    df_total.to_excel(writer, sheet_name='Total', index=False)
    df_rapp.to_excel(writer, sheet_name='Inicial', index=False)
    df_enturmados.to_excel(writer, sheet_name='Enturmados', index=False)
    resumo_direc.to_excel(writer, sheet_name='Tabela', index=False)















































###############################################################################
# BASE GERAL, SEM EXCLUSÃO DE NENHUM COMPONENTE
# ADIÇÃO DE INFORMAÇÕES DE 2026

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


# Manunteção somente das colunas de interesse
df_rapp = df_rapp[['MATRÍCULA', 'NOME', 'COMPONENTE CURRICULAR', 'INEP ESCOLA', 'ESCOLA', 'SÉRIE', 'DIREC', 'ETAPA DE ENSINO', 'ETAPA_RESUMIDA']]


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
# < 60: Não Aprovado;
# NaN: Não Avaliado.
df_merged['Situação'] = df_merged['rendimento (%)'].apply(
    lambda x: 'Não Avaliado' if pd.isna(x)
    else 'Aprovado' if x >= 60
    else 'Não Aprovado'
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


df_enturmados['SÉRIE'].value_counts()
# Criar a coluna 'ETAPA_RESUMIDA' no df_enturmados, com valores 'Ens. Fund. - Anos Finais' e 'Ensino Médio', de acordo com a Série
# Criar a coluna 'ETAPA_RESUMIDA' a partir da SÉRIE
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais',
    'TURMA I (6° E 7° ANOS)': 'Ens. Fund. - Anos Finais',
    'TURMA II (8° E 9° ANOS)': 'Ens. Fund. - Anos Finais',
    '1º Período (1ª Série)': 'Ensino Médio',
    '2º Período (2ª Série)': 'Ensino Médio',
    '3º Período (3ª Série)': 'Ensino Médio',
    '3° PERÍODO': 'Ensino Médio',
    '1º MÓDULO': 'Ensino Médio',
    '2º MODULO': 'Ensino Médio',
    '3º MÓDULO': 'Ensino Médio',
    'BLOCO B': 'Ensino Médio',
    'BLOCO C': 'Ensino Médio',
    'BLOCO D': 'Ensino Médio',
    '5º PERÍODO (8° E 9° ANO - ANOS FINAIS)': 'Ens. Fund. - Anos Finais',
    '4º PERÍODO (6° E 7° ANO - ANOS FINAIS)': 'Ens. Fund. - Anos Finais',
    '2º SEMESTRE': 'Ensino Médio',
    '3º SEMESTRE': 'Ensino Médio',
    'UNICA': 'Ens. Fund. - Anos Finais'
}

df_enturmados['ETAPA_RESUMIDA'] = df_enturmados['SÉRIE'].map(mapeamento_etapa)

# Excluir duplicatas do df_enturmados, considerando as colunas 'MATRÍCULA' e 'COMPONENTE CURRICULAR'.
# Seguir o ordenamento de preferência:
# SITUAÇÃO FINAL = APROVADO;
# maior nota em MÉDIA FINAL.

# Garante que MÉDIA FINAL é numérica
df_enturmados['MÉDIA FINAL'] = pd.to_numeric(
    df_enturmados['MÉDIA FINAL'],
    errors='coerce'
)

# Cria a prioridade (1 = APROVADO, 0 = demais)
df_enturmados['_prioridade'] = (
    df_enturmados['SITUAÇÃO FINAL']
    .eq('APROVADO')
    .astype(int)
)

# Ordena pelos critérios
df_enturmados = (
    df_enturmados
    .sort_values(
        by=['_prioridade', 'MÉDIA FINAL'],
        ascending=[False, False]
    )
    .drop_duplicates(
        subset=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
        keep='first'
    )
    .drop(columns='_prioridade')
    .reset_index(drop=True)
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



# Merge entre o df_merged e df_enturmados, considerando a chave 'MATRÍCULA' e 'COMPONENTE CURRICULAR'
# Merge externo (outer) para manter todos os registros de ambos os dataframes
df_final = df_merged.merge(
    df_enturmados,
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='outer',
    suffixes=('', '_ent')
)

# Identifica as linhas que vieram somente do df_enturmados
novas_linhas = df_final['NOME'].isna()

# Preenche as colunas do df_merged com as informações equivalentes do df_enturmados
df_final.loc[novas_linhas, 'NOME'] = df_final.loc[novas_linhas, 'ESTUDANTE']
df_final.loc[novas_linhas, 'INEP ESCOLA'] = df_final.loc[novas_linhas, 'INEP ESCOLA_ent']
df_final.loc[novas_linhas, 'ESCOLA'] = df_final.loc[novas_linhas, 'ESCOLA_ent']
df_final.loc[novas_linhas, 'SÉRIE'] = df_final.loc[novas_linhas, 'SÉRIE_ent']
df_final.loc[novas_linhas, 'DIREC'] = df_final.loc[novas_linhas, 'DIREC_ent']
df_final.loc[novas_linhas, 'ETAPA_RESUMIDA'] = df_final.loc[novas_linhas, 'ETAPA_RESUMIDA_ent']

# Situação para os novos registros
df_final.loc[novas_linhas, 'Situação'] = 'Não Avaliado'

# Mantém somente as colunas do df_merged
df_final = df_final[df_merged.columns]


# Adicionar coluna 'Enturmação' no df_final, com valor 'Sim' se o estudante e componente estiver no df_enturmados, caso contrário 'Não'
# Selecionar apenas as chaves do df_enturmados
df_enturmados_merge = (
    df_enturmados[['MATRÍCULA', 'COMPONENTE CURRICULAR']]
    .drop_duplicates()
    .assign(Enturmação='Sim')
)

# Converter MATRÍCULA para string em ambos os dataframes para garantir que a junção funcione corretamente
df_final['MATRÍCULA'] = (
    df_final['MATRÍCULA']
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
df_final = df_final.merge(
    df_enturmados_merge,
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='left'
)

# Preencher os que não foram encontrados
df_final['Enturmação'] = df_final['Enturmação'].fillna('Não')


# Estudante com 'SITUAÇÃO FINAL' = APROVADO no df_enturmados é considerado aprovado, mesmo que não tenha feito a prova Plurall.
# A nota desse estudante será a nota que está em 'MÉDIA FINAL'.
# Trazer informações do df_enturmados
df_final = df_final.merge(
    df_enturmados[
        ['MATRÍCULA', 'COMPONENTE CURRICULAR', 'SITUAÇÃO FINAL', 'MÉDIA FINAL']
    ].drop_duplicates(),
    on=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    how='left'
)


# Máscara dos aprovados
mask = df_final['SITUAÇÃO FINAL'].eq('APROVADO')

# Converter 'MÉDIA FINAL' para numérico
df_final['MÉDIA FINAL'] = pd.to_numeric(
    df_final['MÉDIA FINAL'],
    errors='coerce'
)

# Multiplicar por 10 para ficar na mesma escala de 'rendimento (%)'
df_final['MÉDIA FINAL'] = df_final['MÉDIA FINAL'] * 10


# Atualizar nota e situação do df_final para os estudantes aprovados no df_enturmados
df_final.loc[mask, 'rendimento (%)'] = df_final.loc[mask, 'MÉDIA FINAL']
df_final.loc[mask, 'Situação'] = 'Aprovado'

# Excluir as colunas 'SITUAÇÃO FINAL' e 'MÉDIA FINAL' do df_merged
df_final = df_final.drop(columns=['SITUAÇÃO FINAL', 'MÉDIA FINAL'])

# Inserção de informações de 2026
# Dataframe com os dados do Relatório Geral de Matrículas de 2026
df_geral_2026 = pd.read_excel(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026\20260729_2026_Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)

# Converter MATRÍCULA para string no df_geral_2026 para garantir que a junção funcione corretamente
df_geral_2026['MATRÍCULA'] = (
    df_geral_2026['MATRÍCULA']
    .astype(str)
    .str.strip()
)

# Converter 'DATA DA OPERAÇÃO' para datetime
df_geral_2026['DATA DA OPERAÇÃO'] = pd.to_datetime(
    df_geral_2026['DATA DA OPERAÇÃO'],
    dayfirst=True,
    errors='coerce'
)

# Ordenar pela data mais recente
df_geral_merge = (
    df_geral_2026
    .sort_values('DATA DA OPERAÇÃO', ascending=False)
    .drop_duplicates(subset='MATRÍCULA', keep='first')
    [
        [
            'MATRÍCULA',
            'DIREC',
            'CÓDIGO INEP ESCOLA',
            'ESCOLA',
            'ETAPA DE ENSINO',
            'SÉRIE',
            'SITUAÇÃO'
        ]
    ]
    .rename(columns={
        'DIREC': 'DIREC_2026',
        'CÓDIGO INEP ESCOLA': 'CÓDIGO INEP ESCOLA_2026',
        'ESCOLA': 'ESCOLA_2026',
        'ETAPA DE ENSINO': 'ETAPA DE ENSINO_2026',
        'SÉRIE': 'SÉRIE_2026',
        'SITUAÇÃO': 'SITUAÇÃO_2026'
    })
)

# Merge mantendo apenas as linhas de df_final
df_final = df_final.merge(
    df_geral_merge,
    on='MATRÍCULA',
    how='left'
)


# Exportar a base final em Excel
df_final.to_excel(r"D:\Scripts_Python\FGV\Monitoramento_RAPP_2026\20260729_Total_RAPP.xlsx", index=False)







