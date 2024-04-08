import pandas as pd
import plotly.express as px

df_principal = pd.read_excel("assets/Imersão Python.xlsx", sheet_name="Principal", engine='openpyxl')
print(df_principal)

df_total_de_acoes = pd.read_excel("assets/Imersão Python.xlsx", sheet_name="Total_de_acoes", engine='openpyxl')
print(df_total_de_acoes)

df_ticker = pd.read_excel("assets/Imersão Python.xlsx", sheet_name="Ticker", engine='openpyxl')
print(df_ticker)

df_chatgpt = pd.read_excel("assets/Imersão Python.xlsx", sheet_name="Chatgpt", engine='openpyxl')
print(df_chatgpt)

def_principal = df_principal[[ 'Ativo',	'Data',	'Último (R$)',	'Var. Dia (%)']].copy()
print(df_principal)

df_principal = df_principal.rename(columns={'Último (R$)':'valor_final', 'Var. Dia (%)':'var_dia_pct'}).copy()
print(df_principal)

df_principal['Var_pct'] = df_principal['var_dia_pct']/100
print(df_principal)

df_principal['valor_inicial'] = df_principal['valor_final']/ (df_principal['Var_pct'] + 1)
print(df_principal)

df_principal = df_principal.merge(df_total_de_acoes, left_on='Ativo', right_on='Código', how='left')
print(df_principal)

df_principal = df_principal.drop(columns=['Código'])
print(df_principal)

df_principal['Variacao_r$'] = (df_principal['valor_final'] - df_principal['valor_inicial'])* df_principal['Qtde. Teórica']
print(df_principal)

pd.options.display.float_format = '{:.2f}'.format

df_principal['Qtde. Teórica'] = df_principal['Qtde. Teórica'].astype(int)
print(df_principal)

df_principal = df_principal.rename(columns={'Qtde. Teórica':'Qtd_teorica'}).copy()
print(df_principal)

df_principal['Resultado'] = df_principal['Variacao_r$'].apply(lambda x: "Subiu" if x > 0 else ("Desceu" if x < 0 else "Estável") )
print(df_principal)

df_principal = df_principal.merge(df_ticker, left_on='Ativo', right_on='Ticker', how='left')
df_principal = df_principal.drop(columns=['Ticker'])
print(df_principal)

df_principal = df_principal.merge(df_chatgpt, left_on='Nome', right_on='Nome da Empresa', how='left')
df_principal = df_principal.drop(columns=['Nome da Empresa'])
print(df_principal)

df_principal['Cat_Idade'] = df_principal['Idade (anos)'].apply(lambda x: "Mais de 100" if x > 100 else ("Menor de 50" if x < 50 else "Entre 50 e 100") )
print(df_principal)

df_principal = df_principal.rename(columns={'Idade (anos)':'idade'}).copy()
print(df_principal)

# Calculando o maior valor
maior = df_principal['Variacao_r$'].max()

# Calculando o menor valor
menor = df_principal['Variacao_r$'].min()

# Calculando a média
media = df_principal['Variacao_r$'].mean()

# Calculando a média de quem subiu
media_subiu = df_principal[df_principal['Resultado'] == 'Subiu']['Variacao_r$'].mean()

# Calculando a média de quem desceu
media_desceu = df_principal[df_principal['Resultado'] == 'Desceu']['Variacao_r$'].mean()

# Exibindo os resultados
print(f"Maior\tR$ {maior:,.2f}")
print(f"Menor\tR$ {menor:,.2f}")
print(f"Média\tR$ {media:,.2f}")
print(f"Média de quem subiu\tR$ {media_subiu:,.2f}")
print(f"Média de quem desceu\tR$ {media_desceu:,.2f}")

df_principal_subiu = df_principal[df_principal['Resultado'] == 'Subiu']
print(df_principal_subiu)

df_analise_segmento = df_principal_subiu.groupby('Segmento')['Variacao_r$'].sum().reset_index()
print(df_analise_segmento)

df_analise_saldo = df_principal.groupby('Resultado')['Variacao_r$'].sum().reset_index()
print(df_analise_saldo)

# Calculando a soma da variação ($) por faixa etária
df_variacao_idade = df_principal.groupby('Cat_Idade')['Variacao_r$'].sum()

# Calculando a contagem de empresas por faixa etária
df_contagem_empresas = df_principal['Cat_Idade'].value_counts()

# Criando um DataFrame com os resultados
df_analise_etaria = pd.DataFrame({
    'Variacao_r$': df_variacao_idade,
    'Contagem_de_Empresas': df_contagem_empresas
})

# Formatando a coluna 'Soma da Variação ($)' para exibir valores em formato de moeda
df_analise_etaria['Variacao_r$'] = df_analise_etaria['Variacao_r$'].map(lambda x: f'R$ {x:,.2f}')
print(df_analise_etaria)

# Criando o gráfico de barras com orientação horizontal
fig = px.bar(df_analise_saldo, y='Resultado', x='Variacao_r$', text='Variacao_r$',
             title='Variação em Reais por Resultado', orientation='h')

# Exibindo o gráfico
fig.show()


# Agrupando os dados por segmento e somando as variações
df_analise_segmento = df_principal[df_principal['Resultado'] == 'Subiu'].groupby('Segmento')['Variacao_r$'].sum().reset_index()

# Criando o gráfico de pizza
fig = px.pie(df_analise_segmento,
             values='Variacao_r$',
             names='Segmento',
             title='Variação de Receita por Segmento (Segmentos com Crescimento)')

# Adicionando rótulos claros
fig.update_traces(textposition='inside', textinfo='percent+label')

# Ajustando o layout para melhor visualização
fig.update_layout(
    title_x=0.5,  # Centralizando o título
    legend_title='Segmento',  # Título da legenda
    margin=dict(t=50, b=50, l=50, r=50),  # Margens
    font=dict(family="Arial", size=12),  # Estilo de fonte
)

# Mostrando o gráfico
fig.show()


df_variacao_idade = df_variacao_idade.reset_index()

# Criando o gráfico de barras
fig = px.bar(df_variacao_idade, y='Cat_Idade', x='Variacao_r$',
             title='Soma da Variação ($) por Faixa Etária',
             labels={'Cat_Idade': 'Faixa Etária', 'Variacao_r$': 'Soma da Variação ($)'},
             orientation='h')  # orientação horizontal

# Exibindo o gráfico
fig.show()

fig = px.bar(x=df_contagem_empresas.values, y=df_contagem_empresas.index,
             labels={'y': 'Faixa Etária', 'x': 'Contagem de Empresas'},
             title='Contagem de Empresas por Faixa Etária',
             orientation='h')  # orientação horizontal

# Mostrando o gráfico
fig.show()