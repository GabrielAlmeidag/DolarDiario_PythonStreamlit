import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import plotly.express as px
import pythoncom
import win32com.client as win32  # pip install pywin32

# Configuração da página
st.set_page_config(
    page_title="Dashboard Cotações PTAX",
    page_icon="📈",
    layout="wide"
)

def get_currency_data(currency: str, start_date, end_date) -> pd.DataFrame:
    """
    Busca cotações PTAX para USD, EUR ou GBP no período indicado.
    currency: código ISO (USD, EUR, GBP)
    """
    date_format = '%m-%d-%Y'
    base_url = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
    params = {
        "@dataInicial": f"'{start_date.strftime(date_format)}'",
        "@dataFinalCotacao": f"'{end_date.strftime(date_format)}'",
        "$format": "json"
    }

    # Define endpoint e, se necessário, o parâmetro @moeda
    if currency.upper() == "USD":
        endpoint = "CotacaoDolarPeriodo"
    else:
        endpoint = "CotacaoMoedaPeriodo"
        params["@moeda"] = f"'{currency.upper()}'"

    # Monta URL
    if endpoint == "CotacaoDolarPeriodo":
        url = f"{base_url}{endpoint}(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)"
    else:
        url = f"{base_url}{endpoint}(moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)"
    url += "?" + "&".join(f"{k}={v}" for k, v in params.items())

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json().get("value", [])
        if not data:
            st.warning(f"Nenhum dado encontrado para {currency} no período selecionado")
        return pd.DataFrame(data)
    except requests.exceptions.RequestException as e:
        st.error(f"Erro HTTP ao buscar {currency}: {e}")
    except ValueError as e:
        st.error(f"Erro ao interpretar resposta JSON para {currency}: {e}")
    return pd.DataFrame()

def send_email_via_outlook(dataframe: pd.DataFrame, to_email: str, subject: str) -> bool:
    """
    Envia o DataFrame como tabela HTML via Outlook, usando DataFrame.to_html().
    """
    if dataframe.empty:
        st.warning("Nada a enviar - DataFrame vazio")
        return False

    # Prepara DataFrame
    df = dataframe.copy()
    df["Data/Hora"] = pd.to_datetime(df["dataHoraCotacao"]).dt.strftime("%d/%m/%Y %H:%M")
    df = df[["Moeda", "Data/Hora", "cotacaoCompra", "cotacaoVenda"]]
    df.columns = ["Moeda", "Data/Hora", "Compra (R$)", "Venda (R$)"]

    # Gera HTML sem índice e já formatado
    html_table = df.to_html(
        index=False,
        header=True,
        border=0,
        justify="center",
        formatters={
            "Compra (R$)": lambda x: f"R$ {x:.4f}",
            "Venda (R$)": lambda x: f"R$ {x:.4f}"
        }
    )

    # Monta o corpo do email com CSS inline
    html = f"""
    <style>
      table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }}
      th {{ background-color: #0066cc; color: white; padding: 10px; text-align: center; }}
      td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
      tr:nth-child(even) {{ background-color: #f9f9f9; }}
    </style>
    <h1 style="color: #0066cc; font-family: Arial, sans-serif;">📊 Cotações PTAX</h1>
    <p style="font-family: Arial, sans-serif;">
      <strong>Período:</strong> {df['Data/Hora'].min()} a {df['Data/Hora'].max()}
    </p>
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
    except Exception as e:
        st.error(f"Falha no envio via Outlook: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

# Sidebar com filtros
with st.sidebar:
    st.header("🔍 Filtros")
    moedas = st.multiselect("Selecione as moedas", ["USD", "EUR", "GBP"], default=["USD"])
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Data inicial", datetime.now() - timedelta(days=7))
    with col2:
        end_date = st.date_input("Data final", datetime.now())
    tipo_cotacao = st.radio("Tipo de cotação", ["Compra", "Venda"], index=0, horizontal=True)
    st.markdown("---")
    st.header("📤 Envio por Email")
    email_to = st.text_input("Destinatário", "exemplo@email.com")
    email_subject = st.text_input("Assunto", f"Cotações PTAX - {datetime.now():%d/%m/%Y}")

# Cache de dados
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

# Carrega e guarda na sessão
df = load_data(moedas, start_date, end_date)
st.session_state['df'] = df

# Botão de envio
if st.sidebar.button("Enviar Relatório", type="primary"):
    df_session = st.session_state.get('df', pd.DataFrame())
    if not df_session.empty:
        with st.spinner("Enviando email..."):
            if send_email_via_outlook(df_session, email_to, email_subject):
                st.success("Email enviado com sucesso via Outlook!")
                st.balloons()
            else:
                st.error("Falha ao enviar email")
    else:
        st.warning("Nenhum dado disponível para enviar")

# UI principal
st.title("📈 Dashboard Cotações PTAX")
st.markdown("---")

if not df.empty:
    # Métricas
    ultima = df.sort_values("dataHoraCotacao").groupby("Moeda").last()
    coluna = 'cotacaoCompra' if tipo_cotacao == "Compra" else 'cotacaoVenda'
    cols = st.columns(4)

    def make_metric(label, moeda):
        if moeda in moedas:
            val = ultima.loc[moeda, coluna]
            mean = df[df["Moeda"] == moeda][coluna].mean()
            delta = ((val - mean) / mean * 100) if mean else 0
            st.metric(f"{label} ({tipo_cotacao})", f"R$ {val:.4f}", f"{delta:.2f}%")
        else:
            st.metric(label, "-")

    with cols[0]: make_metric("USD", "USD")
    with cols[1]: make_metric("EUR", "EUR")
    with cols[2]:
        if "USD" in moedas:
            series = df[df["Moeda"] == "USD"][coluna]
            var = series.iloc[-1] - series.iloc[0]
            pct = (var / series.iloc[0] * 100) if series.iloc[0] else 0
            st.metric("Variação USD", f"R$ {var:.4f}", f"{pct:.2f}%")
    with cols[3]:
        if "EUR" in moedas:
            series = df[df["Moeda"] == "EUR"][coluna]
            var = series.iloc[-1] - series.iloc[0]
            pct = (var / series.iloc[0] * 100) if series.iloc[0] else 0
            st.metric("Variação EUR", f"R$ {var:.4f}", f"{pct:.2f}%")

    st.markdown("---")
    # Gráfico
    fig = px.line(df, x="dataHoraCotacao", y=coluna, color="Moeda",
                  labels={coluna: "Valor (R$)", "dataHoraCotacao": "Data"})
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    # Tabela interativa
    df_disp = (
        df[["Moeda", "dataHoraCotacao", "cotacaoCompra", "cotacaoVenda"]]
        .rename(columns={
            "dataHoraCotacao": "Data/Hora",
            "cotacaoCompra": "Compra (R$)",
            "cotacaoVenda": "Venda (R$)"
        })
    )
    df_disp["Data/Hora"] = df_disp["Data/Hora"].dt.strftime("%d/%m/%Y %H:%M")
    st.dataframe(
        df_disp.style.format({"Compra (R$)": "{:.4f}", "Venda (R$)": "{:.4f}"}),
        height=400, use_container_width=True
    )
else:
    st.warning("⚠️ Nenhum dado encontrado para os filtros selecionados")
    st.info("Dicas: use datas de dias úteis e verifique a API do BC")

# CSS personalizado
st.markdown("""
<style>
  .stMetric { background: white; border-radius: 10px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
  .stMetric label { font-weight: 600; color: #444; }
  .stMetric value { font-size: 1.5rem; }
  .stButton>button { background-color: #0066cc; color: white; font-weight: bold; }
</style>
""", unsafe_allow_html=True)
