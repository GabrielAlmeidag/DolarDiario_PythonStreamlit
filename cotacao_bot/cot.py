import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import plotly.express as px
import pythoncom
import win32com.client as win32

st.set_page_config(page_title="Dashboard Cotações PTAX", page_icon="📈", layout="wide")

def get_currency_data(currency: str, start_date, end_date) -> pd.DataFrame:
    date_format = '%m-%d-%Y'
    base_url = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
    params = {
        "@dataInicial": f"'{start_date.strftime(date_format)}'",
        "@dataFinalCotacao": f"'{end_date.strftime(date_format)}'",
        "$format": "json"
    }
    if currency.upper() == "USD":
        endpoint = "CotacaoDolarPeriodo"
    else:
        endpoint = "CotacaoMoedaPeriodo"
        params["@moeda"] = f"'{currency.upper()}'"
    if endpoint == "CotacaoDolarPeriodo":
        url = f"{base_url}{endpoint}(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)"
    else:
        url = f"{base_url}{endpoint}(moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)"
    url += "?" + "&".join(f"{k}={v}" for k, v in params.items())
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        return pd.DataFrame(response.json().get("value", []))
    except:
        return pd.DataFrame()

def send_email_via_outlook(dataframe: pd.DataFrame, to_email: str, subject: str) -> bool:
    if dataframe.empty:
        return False
    html_table = dataframe.to_html(
        index=False,
        header=True,
        border=0,
        justify="center",
        formatters={
            "Compra (R$)": lambda x: f"R$ {x:.4f}",
            "Venda (R$)": lambda x: f"R$ {x:.4f}"
        }
    )
    html = f"""
    <style>
      table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }}
      th {{ background-color: #0066cc; color: white; padding: 10px; text-align: center; }}
      td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
      tr:nth-child(even) {{ background-color: #f9f9f9; }}
    </style>
    <h1 style="color: #0066cc; font-family: Arial, sans-serif;">📊 Cotações PTAX</h1>
    {html_table}
    <p style="font-family: Arial, sans-serif; font-size:0.9em;">
      <em>Atualizado em: {datetime.now():%d/%m/%Y %H:%M}</em>
    </p>
    """
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to_email
        mail.Subject = subject
        mail.HTMLBody = html
        mail.Send()
        return True
    except:
        return False
    finally:
        pythoncom.CoUninitialize()

with st.sidebar:
    moedas = st.multiselect("Selecione as moedas", ["USD", "EUR", "GBP"], default=["USD"])
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Data inicial", datetime.now() - timedelta(days=7))
    with col2:
        end_date = st.date_input("Data final", datetime.now())
    tipo_cotacao = st.radio("Tipo de cotação", ["Compra", "Venda"], index=0, horizontal=True)
    st.markdown("---")
    email_to = st.text_input("Destinatário", "exemplo@email.com")
    email_subject = st.text_input("Assunto", f"Cotações PTAX - {datetime.now():%d/%m/%Y}")

@st.cache_data(ttl=3600, show_spinner="Buscando dados...")
def load_data(moedas, start_date, end_date):
    dfs = []
    for cur in moedas:
        df_cur = get_currency_data(cur, start_date, end_date)
        if not df_cur.empty:
            df_cur["Moeda"] = cur
            df_cur["dataHoraCotacao"] = pd.to_datetime(df_cur["dataHoraCotacao"])
            dfs.append(df_cur)
    return pd.concat(dfs) if dfs else pd.DataFrame()

df = load_data(moedas, start_date, end_date)
ultima = df.sort_values("dataHoraCotacao").groupby("Moeda").last()
df_last = ultima.reset_index()
df_last["Data/Hora"] = df_last["dataHoraCotacao"].dt.strftime("%d/%m/%Y %H:%M")
df_last = df_last[["Moeda", "Data/Hora", "cotacaoCompra", "cotacaoVenda"]].rename(
    columns={"cotacaoCompra": "Compra (R$)", "cotacaoVenda": "Venda (R$)"}
)

if st.sidebar.button("Enviar Relatório", type="primary"):
    if send_email_via_outlook(df_last, email_to, email_subject):
        st.success("Email enviado com sucesso via Outlook!")
        st.balloons()
    else:
        st.error("Falha ao enviar email")

st.title("📈 Dashboard Cotações PTAX")
st.markdown("---")

if not df.empty:
    coluna = "cotacaoCompra" if tipo_cotacao == "Compra" else "cotacaoVenda"
    cols_val = st.columns(len(moedas))
    for col, moeda in zip(cols_val, moedas):
        val = ultima.loc[moeda, coluna]
        mean = df[df["Moeda"] == moeda][coluna].mean()
        delta = ((val - mean) / mean * 100) if mean else 0
        with col:
            st.metric(f"{moeda} ({tipo_cotacao})", f"R$ {val:.4f}", f"{delta:.2f}%")
    cols_var = st.columns(len(moedas))
    for col, moeda in zip(cols_var, moedas):
        series = df[df["Moeda"] == moeda][coluna]
        var = series.iloc[-1] - series.iloc[0]
        pct = (var / series.iloc[0] * 100) if series.iloc[0] else 0
        with col:
            st.metric(f"Variação {moeda}", f"R$ {var:.4f}", f"{pct:.2f}%")
    st.markdown("---")
    fig = px.line(df, x="dataHoraCotacao", y=coluna, color="Moeda",
                  labels={coluna: "Valor (R$)", "dataHoraCotacao": "Data"})
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("---")
    st.markdown("### Última cotação disponível")
    st.dataframe(df_last.style.format({"Compra (R$)": "{:.4f}", "Venda (R$)": "{:.4f}"}), use_container_width=True)
else:
    st.warning("⚠️ Nenhum dado encontrado para os filtros selecionados")
    st.info("Dicas: use datas de dias úteis e verifique a API do BC")

st.markdown("""
<style>
  .stMetric { background: #f0f5ff; border-radius: 10px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 10px; }
  .stMetric label { font-weight: 600; color: #003366; }
  .stMetric value { font-size: 1.5rem; }
  .stButton>button { background-color: #0066cc; color: white; font-weight: bold; }
</style>
""", unsafe_allow_html=True)
