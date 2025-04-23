
import requests
import pandas as pd
from datetime import datetime
import streamlit as st

def get_currency_data(currency):
    hoje = datetime.now()
    data_inicio = hoje.replace(day=1)
    data_inicio_str = data_inicio.strftime('%m-%d-%Y')
    data_fim_str = hoje.strftime('%m-%d-%Y')
    
    # Seleciona o endpoint correto para cada moeda
    if currency == "dolar":
        url = (
            f"https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
            f"CotacaoDolarPeriodo(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)?"
            f"@dataInicial='{data_inicio_str}'&@dataFinalCotacao='{data_fim_str}'"
            f"&$top=100&$format=json"
        )
    elif currency == "euro":
        url = (
            f"https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
            f"CotacaoMoedaPeriodo(moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)?"
            f"@moeda='EUR'&@dataInicial='{data_inicio_str}'&@dataFinalCotacao='{data_fim_str}'"
            f"&$top=100&$format=json"
        )
    
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()["value"]
    df = pd.DataFrame(data)
    df["dataHoraCotacao"] = pd.to_datetime(df["dataHoraCotacao"])
    df = df.sort_values("dataHoraCotacao")
    return df

# ====== Configuração do Streamlit ======
st.set_page_config(page_title="Cotações PTAX", layout="wide")
st.markdown("<h1 style='text-align: center; color: darkblue;'>💱 Dashboard de Cotações (PTAX)</h1>", unsafe_allow_html=True)
st.markdown("### 🔎 Dados do mês atual - Dólar e Euro em relação ao Real (BRL)")

# Adicionando seleção de período
st.sidebar.title("Opções")
data_inicio = st.sidebar.date_input("Data inicial", datetime.now().replace(day=1))
data_fim = st.sidebar.date_input("Data final", datetime.now())

st.divider()
col1, col2 = st.columns(2)

try:
    # ===== DÓLAR =====
    df_dolar = get_currency_data("dolar")
    ult_dolar = df_dolar.iloc[-1]
    penult_dolar = df_dolar.iloc[-2]
    var_dolar = ult_dolar["cotacaoCompra"] - penult_dolar["cotacaoCompra"]

    with col1:
        st.markdown("<h3 style='color: teal;'>🇺🇸 Dólar (USD/BRL)</h3>", unsafe_allow_html=True)
        st.metric("Compra hoje", f"R$ {ult_dolar['cotacaoCompra']:.4f}", f"{var_dolar:.4f}")
        st.line_chart(df_dolar.set_index("dataHoraCotacao")["cotacaoCompra"])
        
        # Adicionando comparação direta
        st.markdown("**Comparativo USD/EUR**")
        valor_eur_usd = ult_dolar["cotacaoCompra"] / ult_euro["cotacaoCompra"] if 'ult_euro' in locals() else 0
        st.write(f"1 EUR = {valor_eur_usd:.4f} USD")

    # ===== EURO =====
    df_euro = get_currency_data("euro")
    ult_euro = df_euro.iloc[-1]
    penult_euro = df_euro.iloc[-2]
    var_euro = ult_euro["cotacaoCompra"] - penult_euro["cotacaoCompra"]

    with col2:
        st.markdown("<h3 style='color: darkgreen;'>🇪🇺 Euro (EUR/BRL)</h3>", unsafe_allow_html=True)
        st.metric("Compra hoje", f"R$ {ult_euro['cotacaoCompra']:.4f}", f"{var_euro:.4f}")
        st.line_chart(df_euro.set_index("dataHoraCotacao")["cotacaoCompra"])
        
        # Adicionando histórico combinado
        st.markdown("**Histórico Combinado**")
        combined_df = pd.DataFrame({
            'Dólar': df_dolar.set_index("dataHoraCotacao")["cotacaoCompra"],
            'Euro': df_euro.set_index("dataHoraCotacao")["cotacaoCompra"]
        })
        st.line_chart(combined_df)

    # ===== TABELA COMPARATIVA =====
    st.divider()
    st.markdown("### 📊 Tabela Comparativa")
    
    comparativo = pd.DataFrame({
        "Moeda": ["Dólar (USD)", "Euro (EUR)"],
        "Compra (R$)": [ult_dolar["cotacaoCompra"], ult_euro["cotacaoCompra"]],
        "Variação (R$)": [var_dolar, var_euro],
        "Variação %": [
            (var_dolar/penult_dolar["cotacaoCompra"])*100,
            (var_euro/penult_euro["cotacaoCompra"])*100
        ]
    })
    st.dataframe(comparativo.style.format({
        "Compra (R$)": "{:.4f}",
        "Variação (R$)": "{:.4f}",
        "Variação %": "{:.2f}%"
    }))

except Exception as e:
    st.error("❌ Não foi possível acessar os dados da API. Verifique se há cotações disponíveis para este período.")
    st.exception(e)