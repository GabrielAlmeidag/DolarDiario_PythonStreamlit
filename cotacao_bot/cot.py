import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import plotly.express as px

# Configuração da página
st.set_page_config(
    page_title="Dashboard Cotações PTAX",
    page_icon="📈",
    layout="wide"
)

# Função para buscar cotações
def get_currency_data(currency, start_date, end_date):
    start_str = start_date.strftime('%m-%d-%Y')
    end_str = end_date.strftime('%m-%d-%Y')
    
    if currency == "dolar":
        url = f"https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarPeriodo(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)?@dataInicial='{start_str}'&@dataFinalCotacao='{end_str}'&$format=json"
    elif currency == "euro":
        url = f"https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoMoedaPeriodo(moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)?@moeda='EUR'&@dataInicial='{start_str}'&@dataFinalCotacao='{end_str}'&$format=json"
    
    response = requests.get(url)
    data = response.json().get("value", [])
    return pd.DataFrame(data)

# Barra lateral com filtros
with st.sidebar:
    st.header("Filtros")
    
    # Filtro de moedas
    moedas = st.multiselect(
        "Selecione as moedas",
        ["Dólar", "Euro"],
        default=["Dólar", "Euro"]
    )
    
    # Filtro de datas
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Data inicial", datetime.now().replace(day=1))
    with col2:
        end_date = st.date_input("Data final", datetime.now())
    
    # Filtro de tipo de cotação
    tipo_cotacao = st.radio(
        "Tipo de cotação",
        ["Compra", "Venda"],
        horizontal=True
    )

# Busca dados
df_dolar = pd.DataFrame()
df_euro = pd.DataFrame()

if "Dólar" in moedas:
    df_dolar = get_currency_data("dolar", start_date, end_date)
    if not df_dolar.empty:
        df_dolar["dataHoraCotacao"] = pd.to_datetime(df_dolar["dataHoraCotacao"])
        df_dolar["Moeda"] = "Dólar"

if "Euro" in moedas:
    df_euro = get_currency_data("euro", start_date, end_date)
    if not df_euro.empty:
        df_euro["dataHoraCotacao"] = pd.to_datetime(df_euro["dataHoraCotacao"])
        df_euro["Moeda"] = "Euro"

# Combine os dados
df = pd.concat([df_dolar, df_euro])

# Layout principal
st.title("Dashboard de Cotações PTAX")
st.markdown("---")

# Seção de KPIs
if not df.empty:
    st.header("Indicadores Chave")
    col1, col2, col3, col4 = st.columns(4)
    
    ultima_cotacao = df.sort_values("dataHoraCotacao").groupby("Moeda").last()
    
    with col1:
        st.metric(
            label="Cotação Atual Dólar",
            value=f"R$ {ultima_cotacao.loc['Dólar', 'cotacaoCompra']:.4f}" if "Dólar" in moedas else "-",
            delta=f"{((ultima_cotacao.loc['Dólar', 'cotacaoCompra'] - df_dolar['cotacaoCompra'].mean())/df_dolar['cotacaoCompra'].mean()*100):.2f}%" if "Dólar" in moedas else None
        )
    
    with col2:
        st.metric(
            label="Cotação Atual Euro",
            value=f"R$ {ultima_cotacao.loc['Euro', 'cotacaoCompra']:.4f}" if "Euro" in moedas else "-",
            delta=f"{((ultima_cotacao.loc['Euro', 'cotacaoCompra'] - df_euro['cotacaoCompra'].mean())/df_euro['cotacaoCompra'].mean()*100):.2f}%" if "Euro" in moedas else None
        )
    
    with col3:
        st.metric(
            label="Variação Dólar (período)",
            value=f"{(df_dolar['cotacaoCompra'].iloc[-1] - df_dolar['cotacaoCompra'].iloc[0]):.4f}" if "Dólar" in moedas else "-",
            delta=f"{((df_dolar['cotacaoCompra'].iloc[-1] - df_dolar['cotacaoCompra'].iloc[0])/df_dolar['cotacaoCompra'].iloc[0]*100):.2f}%" if "Dólar" in moedas else None
        )
    
    with col4:
        st.metric(
            label="Variação Euro (período)",
            value=f"{(df_euro['cotacaoCompra'].iloc[-1] - df_euro['cotacaoCompra'].iloc[0]):.4f}" if "Euro" in moedas else "-",
            delta=f"{((df_euro['cotacaoCompra'].iloc[-1] - df_euro['cotacaoCompra'].iloc[0])/df_euro['cotacaoCompra'].iloc[0]*100):.2f}%" if "Euro" in moedas else None
        )

# Gráficos
st.markdown("---")
st.header("Análise Temporal")

if not df.empty:
    fig = px.line(
        df,
        x="dataHoraCotacao",
        y=f"cotacao{tipo_cotacao}",
        color="Moeda",
        title=f"Evolução da Cotação de {tipo_cotacao}",
        labels={
            "dataHoraCotacao": "Data",
            f"cotacao{tipo_cotacao}": f"Valor (R$)"
        }
    )
    st.plotly_chart(fig, use_container_width=True)

# Tabela de dados
st.markdown("---")
st.header("Dados Detalhados")

if not df.empty:
    df_display = df[["Moeda", "dataHoraCotacao", "cotacaoCompra", "cotacaoVenda"]].copy()
    df_display = df_display.rename(columns={
        "dataHoraCotacao": "Data/Hora",
        "cotacaoCompra": "Compra (R$)",
        "cotacaoVenda": "Venda (R$)"
    })
    df_display["Data/Hora"] = df_display["Data/Hora"].dt.strftime("%d/%m/%Y %H:%M")
    
    st.dataframe(
        df_display.style.format({
            "Compra (R$)": "{:.4f}",
            "Venda (R$)": "{:.4f}"
        }),
        use_container_width=True
    )
else:
    st.warning("Nenhum dado disponível para os filtros selecionados")

# Estilo CSS adicional
st.markdown("""
<style>
    .stMetric {
        border: 1px solid #e6e6e6;
        border-radius: 5px;
        padding: 10px;
        background-color: #f9f9f9;
    }
    .stMetric label {
        font-size: 14px;
        color: #555;
    }
    .stMetric value {
        font-size: 24px;
    }
</style>
""", unsafe_allow_html=True)