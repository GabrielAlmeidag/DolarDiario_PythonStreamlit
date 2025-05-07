import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import plotly.express as px
import pythoncom
import win32com.client as win32

# — Page config —
st.set_page_config(
    page_title="Dashboard Cotações PTAX",
    page_icon="📈",
    layout="wide"
)

# — Custom CSS for polished UI —
st.markdown("""
<style>
body { background: #f0f2f5; color: #2b2b2b; }
.stSidebar { background: #ffffff; }
.metric-card {
  background: linear-gradient(135deg, #6c5ce7 0%, #a29bfe 100%);
  color: white;
  border-radius: 12px;
  padding: 1rem;
  box-shadow: 0 4px 12px rgba(0,0,0,0.1);
  text-align: center;
}
.metric-card h4 { margin: 0.5rem 0 0.2rem; font-size: 1.1rem; }
.metric-card .value { font-size: 2rem; font-weight: 700; margin: 0.2rem 0; }
.metric-card .delta { font-size: 0.9rem; }
.metric-card .delta.up { color: #00b894; }
.metric-card .delta.down { color: #d63031; }
.metrics-container {
  display: flex; gap: 1rem; flex-wrap: wrap; justify-content: center; margin-bottom: 2rem;
}
</style>
""", unsafe_allow_html=True)

# — Data functions —
def get_currency_data(code, start, end):
    url_base = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
    fmt = "%m-%d-%Y"
    params = {
        "@dataInicial": f"'{start.strftime(fmt)}'",
        "@dataFinalCotacao": f"'{end.strftime(fmt)}'",
        "$format": "json"
    }
    if code == "USD":
        ep = "CotacaoDolarPeriodo"
    else:
        ep = "CotacaoMoedaPeriodo"
        params["@moeda"] = f"'{code}'"
    url = f"{url_base}{ep}("
    if ep == "CotacaoDolarPeriodo":
        url += "dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao"
    else:
        url += "moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao"
    url += ")?" + "&".join(f"{k}={v}" for k,v in params.items())
    try:
        r = requests.get(url, timeout=10); r.raise_for_status()
        return pd.DataFrame(r.json().get("value", []))
    except:
        return pd.DataFrame()

@st.cache_data(ttl=1800)
def load_data(codes, start, end):
    frames = []
    for c in codes:
        d = get_currency_data(c, start, end)
        if not d.empty:
            d["Moeda"] = c
            d["dataHoraCotacao"] = pd.to_datetime(d["dataHoraCotacao"])
            frames.append(d)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

def send_email(df, to_email, subject):
    if df.empty: return False
    html = df.to_html(index=False, border=0, justify="center",
                      formatters={
                        "Compra (R$)": lambda x: f"R$ {x:.4f}",
                        "Venda (R$)": lambda x: f"R$ {x:.4f}"
                      })
    body = f"""
    <style>
      table {{ border-collapse: collapse; width: 100%; font-family: Arial,sans-serif; }}
      th {{ background: #6c5ce7; color: #fff; padding: 8px; }}
      td {{ border: 1px solid #ddd; padding: 8px; }}
    </style>
    <h2 style="color:#6c5ce7">📊 Cotações PTAX</h2>
    {html}
    <p><em>Atualizado em {datetime.now():%d/%m/%Y %H:%M}</em></p>
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

# — Sidebar —
st.sidebar.header("🔧 Configurações")
today = datetime.now().date()
only_today = st.sidebar.checkbox("Somente hoje", True)
if only_today:
    start_date = end_date = today
else:
    start_date, end_date = st.sidebar.date_input(
        "Período",
        [today - timedelta(days=7), today],
        min_value=today - timedelta(days=365),
        max_value=today
    )
codes = st.sidebar.multiselect("Escolha moedas", ["USD","EUR","GBP"], ["USD"])
quote_type = st.sidebar.radio("Tipo", ["Compra","Venda"], horizontal=True)
st.sidebar.markdown("---")
email_to = st.sidebar.text_input("Enviar para")
email_subj = st.sidebar.text_input("Assunto", f"PTAX {today:%d/%m/%Y}")
if st.sidebar.button("📤 Enviar relatório"):
    last = latest_df[["Moeda","Data/Hora","Compra (R$)","Venda (R$)"]]
    ok = send_email(last, email_to, email_subj)
    st.sidebar.success("Enviado!" if ok else "Falha no envio")

# — Load & prepare —
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

# — Header —
st.title("📈 Dashboard Cotações PTAX")
st.caption(f"Período: {start_date:%d/%m/%Y} – {end_date:%d/%m/%Y}")

field = "cotacaoCompra" if quote_type=="Compra" else "cotacaoVenda"

# — Tabs —
tab1, tab2, tab3 = st.tabs(["📊 Métricas", "📈 Gráfico", "📋 Tabela"])

with tab1:
    st.markdown('<div class="metrics-container">', unsafe_allow_html=True)
    for c in codes:
        val = float(latest_df.loc[latest_df["Moeda"]==c, f"{'Compra (R$)' if quote_type=='Compra' else 'Venda (R$)'}"])
        avg = df[df["Moeda"]==c][field].mean()
        d = (val - avg)/avg*100 if avg else 0
        cls = "up" if d>=0 else "down"
        st.markdown(f"""
          <div class="card metric-card">
            <h4>{c} ({quote_type})</h4>
            <div class="value">R$ {val:.4f}</div>
            <div class="delta {cls}">{d:+.2f}% vs média</div>
          </div>
        """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab2:
    fig = px.area(
        df, x="dataHoraCotacao", y=field, color="Moeda",
        labels={field:"Valor (R$)","dataHoraCotacao":"Data"},
        line_shape="spline", template="plotly_white"
    )
    fig.update_traces(mode="lines+markers", marker=dict(size=6), opacity=0.6, fill='tozeroy')
    fig.update_layout(
        height=450,
        hovermode="x unified",
        legend_title_text="Moeda",
        xaxis=dict(rangeslider=dict(visible=True)),
        margin=dict(l=20,r=20,t=40,b=20)
    )
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.markdown("### Última Cotação Disponível")
    st.dataframe(
        latest_df[["Moeda","Data/Hora","Compra (R$)","Venda (R$)"]],
        use_container_width=True
    )
