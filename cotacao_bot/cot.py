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

# Função otimizada para buscar cotações
def get_currency_data(currency, start_date, end_date):
    date_format = '%m-%d-%Y'
    base_url = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
    params = {
        "@dataInicial": f"'{start_date.strftime(date_format)}'",
        "@dataFinalCotacao": f"'{end_date.strftime(date_format)}'",
        "$format": "json"
    }
    try:
        if currency == "dolar":
            url = (
                f"{base_url}"
                "CotacaoDolarPeriodo(dataInicial=@dataInicial,"
                "dataFinalCotacao=@dataFinalCotacao)"
            )
        else:
            url = (
                f"{base_url}"
                "CotacaoMoedaPeriodo(moeda=@moeda,"
                "dataInicial=@dataInicial,"
                "dataFinalCotacao=@dataFinalCotacao)"
            )
            params["@moeda"] = "'EUR'"
        url += "?" + "&".join(f"{k}={v}" for k, v in params.items())
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json().get("value", [])
        if not data:
            st.warning(f"Nenhum dado encontrado para {currency} no período selecionado")
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Erro ao buscar {currency}: {e}")
        return pd.DataFrame()

# Função para enviar e-mail via Outlook com tabela estilizada
def send_email_via_outlook(dataframe, to_email, subject):
    if dataframe.empty:
        st.warning("Nada a enviar - DataFrame vazio")
        return False

    # Preparar DataFrame para exibição
    df = dataframe.copy()
    df["Data/Hora"] = df["dataHoraCotacao"].dt.strftime("%d/%m/%Y %H:%M")
    df = df[["Moeda", "Data/Hora", "cotacaoCompra", "cotacaoVenda"]]
    df.columns = ["Moeda", "Data/Hora", "Compra (R$)", "Venda (R$)"]

    # Gerar linhas da tabela
    rows = ""
    for _, row in df.iterrows():
        rows += f"""
        <tr>
          <td style="padding:8px;text-align:center;">{row['Moeda']}</td>
          <td style="padding:8px;text-align:center;">{row['Data/Hora']}</td>
          <td style="padding:8px;text-align:right;">R$ {row['Compra (R$)']:.4f}</td>
          <td style="padding:8px;text-align:right;">R$ {row['Venda (R$)']:.4f}</td>
        </tr>
        """

    html = f"""
    <style>
      table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }}
      th {{ background-color: #0066cc; color: white; padding: 10px; text-align: center; }}
      td {{ border: 1px solid #ddd; }}
      tr:nth-child(even) {{ background-color: #f9f9f9; }}
    </style>
    <h1 style="color: #0066cc; font-family: Arial, sans-serif;">📊 Cotações PTAX</h1>
    <p style="font-family: Arial, sans-serif;">
      <strong>Período:</strong> {df['Data/Hora'].min()} a {df['Data/Hora'].max()}
    </p>
    <table>
      <thead>
        <tr>
          <th>Moeda</th>
          <th>Data/Hora</th>
          <th>Compra (R$)</th>
          <th>Venda (R$)</th>
        </tr>
      </thead>
      <tbody>
        {rows}
      </tbody>
    </table>
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

# Sidebar com filtros e configurações de e-mail
with st.sidebar:
    st.header("🔍 Filtros")
    moedas = st.multiselect("Selecione as moedas", ["Dólar", "Euro"], default=["Dólar"])
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

# Função de carregamento e cache
@st.cache_data(ttl=3600, show_spinner="Buscando dados...")
def load_data(moedas, start_date, end_date):
    dfs = []
    if "Dólar" in moedas:
        df_d = get_currency_data("dolar", start_date, end_date)
        if not df_d.empty:
            df_d["Moeda"] = "Dólar"
            df_d["dataHoraCotacao"] = pd.to_datetime(df_d["dataHoraCotacao"])
            dfs.append(df_d)
    if "Euro" in moedas:
        df_e = get_currency_data("euro", start_date, end_date)
        if not df_e.empty:
            df_e["Moeda"] = "Euro"
            df_e["dataHoraCotacao"] = pd.to_datetime(df_e["dataHoraCotacao"])
            dfs.append(df_e)
    return pd.concat(dfs) if dfs else pd.DataFrame()

# Carrega dados e armazena na sessão
df = load_data(moedas, start_date, end_date)
st.session_state['df'] = df

# Botão de envio
if st.sidebar.button("Enviar Relatório", type="primary"):
    df_session = st.session_state.get('df', pd.DataFrame())
    if not df_session.empty:
        with st.spinner("Enviando email..."):
            success = send_email_via_outlook(df_session, email_to, email_subject)
            if success:
                st.success("Email enviado com sucesso via Outlook!")
                st.balloons()
            else:
                st.error("Falha ao enviar email")
    else:
        st.warning("Nenhum dado disponível para enviar")

# Título e separador
st.title("📈 Dashboard Cotações PTAX")
st.markdown("---")

# Se houver dados, exibe métricas, gráfico e tabela
if not df.empty:
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

    with cols[0]: make_metric("Dólar", "Dólar")
    with cols[1]: make_metric("Euro", "Euro")
    with cols[2]:
        if "Dólar" in moedas:
            var = df[df["Moeda"]=="Dólar"][coluna].iloc[-1] - df[df["Moeda"]=="Dólar"][coluna].iloc[0]
            pct = (var / df[df["Moeda"]=="Dólar"][coluna].iloc[0] * 100) if df[df["Moeda"]=="Dólar"][coluna].iloc[0] else 0
            st.metric("Variação Dólar", f"R$ {var:.4f}", f"{pct:.2f}%")
    with cols[3]:
        if "Euro" in moedas:
            var = df[df["Moeda"]=="Euro"][coluna].iloc[-1] - df[df["Moeda"]=="Euro"][coluna].iloc[0]
            pct = (var / df[df["Moeda"]=="Euro"][coluna].iloc[0] * 100) if df[df["Moeda"]=="Euro"][coluna].iloc[0] else 0
            st.metric("Variação Euro", f"R$ {var:.4f}", f"{pct:.2f}%")

    st.markdown("---")
    fig = px.line(df, x="dataHoraCotacao", y=coluna, color="Moeda",
                  labels={coluna: "Valor (R$)", "dataHoraCotacao": "Data"})
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
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
    st.info("Dicas: Verifique se as datas são dias úteis e se a API do BC está acessível")

# Estilos CSS personalizados
st.markdown("""
<style>
  .stMetric { background: white; border-radius: 10px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
  .stMetric label { font-weight: 600; color: #444; }
  .stMetric value { font-size: 1.5rem; }
  .stButton>button { background-color: #0066cc; color: white; font-weight: bold; }
</style>
""", unsafe_allow_html=True)
