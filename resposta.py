import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# Carregar o arquivo Excel com múltiplas abas
file_path = 'Respostas ADM.xlsx'  # Ajuste o caminho se necessário
data = pd.read_excel(file_path, sheet_name=None)

# Remover espaços e normalizar os nomes das colunas para todas as abas
for sheet in data:
    data[sheet].columns = data[sheet].columns.str.strip()

# Aba Superintendente
superintendente_data = data['superintendente']
superintendente_data.columns = ['Superintendente', 'Resposta']

# Contar as frequências de respostas por superintendente usando a coluna 'Resposta'
superintendente_counts = superintendente_data.groupby(['Superintendente', 'Resposta']).size().unstack(fill_value=0)

# Aba Resposta
response_data = data['resposta']

# Contar as frequências de cada resposta para cada módulo, com separação de respostas múltiplas
modules = ["MÓDULO OBRAS", "MÓDULO FINANCEIRO", "MÓDULO SUPRIMENTOS", "OUTROS"]
response_counts = {}

for module in modules:
    # Dividir respostas por vírgula e criar uma lista única de respostas
    all_responses = response_data[module].dropna().str.split(',').sum()
    # Remover espaços em branco adicionais
    all_responses = [response.strip() for response in all_responses]
    # Contar a frequência de cada resposta
    response_counts[module] = pd.Series(all_responses).value_counts().to_frame(name='Contagem')

# Função para exibir o gráfico de respostas
def display_module_data(module, counts):
    st.markdown(f"<h2 style='text-align: center;'>{module}</h2>", unsafe_allow_html=True)
    
    fig, ax = plt.subplots(figsize=(16, 12))  # Aumentar bastante o tamanho do gráfico
    bars = counts[:-1].plot(kind='barh', ax=ax, color='#1f77b4', edgecolor='black')  # Excluir 'Total' do gráfico
    ax.set_xlabel('Número de marcações', fontsize=16)
    ax.tick_params(axis='both', labelsize=14)
    
    # Adicionar contagem de votos ao lado das barras
    for bar in bars.patches:
        ax.annotate(f'{int(bar.get_width())}', 
                    (bar.get_width(), bar.get_y() + bar.get_height() / 2),
                    ha='left', va='center', fontsize=14, color='black', xytext=(5, 0), textcoords='offset points')
    
    # Ajustes para layout mais compacto
    plt.tight_layout()
    st.pyplot(fig)

# Função para exibir as respostas por superintendente
def display_superintendente_data(superintendente_data):
    st.markdown("<h2 style='text-align: center;'>Respostas por Superintendente</h2>", unsafe_allow_html=True)
    
    fig, ax = plt.subplots(figsize=(10, 6))  # Aumentar o tamanho do gráfico
    bars = ax.bar(superintendente_data['Superintendente'], superintendente_data['Resposta'], color='#1f77b4', edgecolor='black')
    
    ax.set_xlabel('Superintendente', fontsize=14)

    # Adicionar o valor exato acima de cada barra
    for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, yval, int(yval), ha='center', va='bottom', fontsize=12, color='black')

    plt.tight_layout()
    st.pyplot(fig)

# Interface Streamlit com abas
st.title("Gráfico de respostas do formulário")
tab5, tab1, tab2, tab3, tab4 = st.tabs(["Superintendente"] + modules)

with tab5:
    display_superintendente_data(superintendente_data)

with tab1:
    display_module_data("MÓDULO OBRAS", response_counts["MÓDULO OBRAS"])

with tab2:
    display_module_data("MÓDULO FINANCEIRO", response_counts["MÓDULO FINANCEIRO"])

with tab3:
    display_module_data("MÓDULO SUPRIMENTOS", response_counts["MÓDULO SUPRIMENTOS"])

with tab4:
    display_module_data("OUTROS", response_counts["OUTROS"])
