import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import plotly.express as px
import pythoncom
import win32com.client as win32

# — Page configuration —
st.set_page_config(
    page_title="Dashboard Cotações PTAX",
    page_icon="📈",
    layout="wide"
)

# — Custom CSS for light-green cards side by side with arrows —
st.markdown("""
<style>
  .metrics-container {
    display: flex;
    gap: 1rem;
    overflow-x: auto;
    padding: 1rem 0;
    justify-content: center;
  }
  .metric-card {
    background: linear-gradient(135deg, #e0f7e9 0%, #c8f0de 100%);
    color: #064e3b;
    border-radius: 1rem;
    padding: 1.5rem;
    min-width: 200px;
    flex: 0 0 auto;
    text-align: center;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    transition: transform 0.2s, box-shadow 0.2s;
  }
  .metric-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 6px 18px rgba(0,0,0,0.1);
  }
  .metric-card h3 {
    margin: 0.5rem 0 0.2rem;
    font-size: 1.1rem;
    font-weight: 600;
  }
  .metric-card .value {
    font-size: 2.25rem;
    font-weight: 700;
    margin: 0.3rem 0;
  }
  .metric-card .delta {
    font-size: 1rem;
    font-weight: 500;
  }
  .metric-card .delta.up::before {
    content: "▲ ";
    color: #16a34a;
  }
  .metric-card .delta.down::before {
    content: "▼ ";
    color: #dc2626;
  }
</style>
""", unsafe_allow_html=True)

# — Data fetching & email functions —
def get_currency_data(code, start_date, end_date):
    fmt = "%m-%d-%Y"
    base = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
    params = {
        "@dataInicial": f"'{start_date.strftime(fmt)}'",
        "@dataFinalCotacao": f"'{end_date.strftime(fmt)}'",
        "$format": "json"
    }
    endpoint = "CotacaoDolarPeriodo" if code == "USD" else "CotacaoMoedaPeriodo"
    if code != "USD":
        params["@moeda"] = f"'{code}'"
    url = f"{base}{endpoint}("
    if code == "USD":
        url += "dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao"
    else:
        url += "moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao"
    url += ")?" + "&".join(f"{k}={v}" for k, v in params.items())
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return pd.DataFrame(r.json().get("value", []))
    except:
        return pd.DataFrame()

@st.cache_data(ttl=1800)
def load_data(codes, start_date, end_date):
    frames = []
    for c in codes:
        dfc = get_currency_data(c, start_date, end_date)
        if not dfc.empty:
            dfc["Moeda"] = c
            dfc["dataHoraCotacao"] = pd.to_datetime(dfc["dataHoraCotacao"])
            frames.append(dfc)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

def send_email(df, to_email, subject):
    if df.empty:
        return False
    html = df.to_html(
        index=False, border=0, justify="center",
        formatters={
            "Compra (R$)": lambda x: f"R$ {x:.4f}",
            "Venda (R$)":   lambda x: f"R$ {x:.4f}"
        }
    )
    body = f"""
    <style>
      table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }}
      th {{ background: #c8f0de; color: #064e3b; padding: 8px; }}
      td {{ border: 1px solid #ddd; padding: 8px; }}
      tr:nth-child(even) {{ background: #f2f9f5; }}
    </style>
    <h2 style="color:#064e3b">📊 Cotações PTAX</h2>
    {html}
    <p style="font-size:0.8em"><em>Atualizado em {datetime.now():%d/%m/%Y %H:%M}</em></p>
    """
    try:
        pythoncom.CoInitialize()
        mail = win32.Dispatch("outlook.application").CreateItem(0)
        mail.To, mail.Subject, mail.HTMLBody = to_email, subject, body
        mail.Send()
        return True
    except:
        return False
    finally:
        pythoncom.CoUninitialize()

# — Sidebar controls —
st.sidebar.header("🔧 Configurações")
today = datetime.now().date()
only_today = st.sidebar.checkbox("Somente hoje", True)
if only_today:
    start_date = end_date = today
else:
    start_date, end_date = st.sidebar.date_input(
        "Selecionar período",
        [today - timedelta(days=7), today],
        min_value=today - timedelta(days=365),
        max_value=today
    )
codes = st.sidebar.multiselect("Moedas", ["USD","EUR","GBP"], ["USD"])
quote_type = st.sidebar.radio("Tipo", ["Compra","Venda"], horizontal=True)
st.sidebar.markdown("---")
email_to      = st.sidebar.text_input("Enviar para", "")
email_subject = st.sidebar.text_input("Assunto", f"PTAX {today:%d/%m/%Y}")
# (latest_df will be defined after loading data)
# — Load data & prepare —
df = load_data(codes, start_date, end_date)
if df.empty:
    st.warning("Nenhum dado disponível para o período selecionado.")
    st.stop()

latest_df = (
    df.sort_values("dataHoraCotacao")
      .groupby("Moeda")
      .last()
      .reset_index()
)
latest_df["Data/Hora"] = latest_df["dataHoraCotacao"].dt.strftime("%d/%m/%Y %H:%M")
latest_df = latest_df.rename(
    columns={"cotacaoCompra":"Compra (R$)", "cotacaoVenda":"Venda (R$)"}
)

# Now we can send the email using latest_df
if st.sidebar.button("📤 Enviar relatório"):
    subset = latest_df[["Moeda","Data/Hora","Compra (R$)","Venda (R$)"]]
    ok = send_email(subset, email_to, email_subject)
    st.sidebar.success("Enviado!" if ok else "Falha no envio")

# — Main UI —
st.title("📈 Dashboard Cotações PTAX")
st.caption(f"Período: {start_date:%d/%m/%Y} – {end_date:%d/%m/%Y}")

field = "cotacaoCompra" if quote_type == "Compra" else "cotacaoVenda"

tab1, tab2, tab3 = st.tabs(["📊 Métricas", "📈 Gráfico", "📋 Tabela"])

# Tab 1: cards side by side
with tab1:
    st.markdown('<div class="metrics-container">', unsafe_allow_html=True)
    for c in codes:
        val = float(latest_df.loc[latest_df["Moeda"] == c, 
                                  "Compra (R$)" if quote_type=="Compra" else "Venda (R$)"])
        avg = df[df["Moeda"] == c][field].mean()
        delta = (val - avg) / avg * 100 if avg else 0
        cls = "up" if delta >= 0 else "down"
        st.markdown(f"""
            <div class="metric-card">
              <h3>{c} ({quote_type})</h3>
              <div class="value">R$ {val:.4f}</div>
              <div class="delta {cls}">{abs(delta):.2f}% vs média</div>
            </div>
        """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# Tab 2: area chart
with tab2:
    fig = px.area(
        df, x="dataHoraCotacao", y=field, color="Moeda",
        labels={field:"Valor (R$)", "dataHoraCotacao":"Data"},
        line_shape="spline", template="plotly_white"
    )
    fig.update_traces(mode="lines+markers", marker=dict(size=5), opacity=0.6, fill='tozeroy')
    fig.update_layout(
        height=450,
        hovermode="x unified",
        legend_title_text="Moeda",
        xaxis=dict(rangeslider=dict(visible=True)),
        margin=dict(l=20,r=20,t=40,b=20)
    )
    st.plotly_chart(fig, use_container_width=True)

# Tab 3: last quotation table
with tab3:
    st.markdown("### Última Cotação Disponível")
    st.dataframe(latest_df[["Moeda","Data/Hora","Compra (R$)","Venda (R$)"]], use_container_width=True)
