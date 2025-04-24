import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import plotly.express as px
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time

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
            url = f"{base_url}CotacaoDolarPeriodo(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)"
        elif currency == "euro":
            url = f"{base_url}CotacaoMoedaPeriodo(moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)"
            params["@moeda"] = "'EUR'"
        
        url += "?" + "&".join([f"{k}={v}" for k,v in params.items()])
        
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json().get("value", [])
        
        if not data:
            st.warning(f"Nenhum dado encontrado para {currency} no período selecionado")
        return pd.DataFrame(data)
        
    except Exception as e:
        st.error(f"Erro ao buscar {currency}: {str(e)}")
        return pd.DataFrame()

# Função robusta para enviar email
def send_email(dataframe, to_email, subject):
    if dataframe.empty:
        st.warning("Nada a enviar - DataFrame vazio")
        return False
    
    try:
        # Configurações SMTP (substitua com seus dados)
        smtp_config = {
            "server": "smtp.gmail.com",
            "port": 587,
            "user": "seu_email@gmail.com",  # Substitua
            "password": "sua_senha"        # Substitua
        }
        
        # Criar mensagem
        msg = MIMEMultipart()
        msg['From'] = smtp_config["user"]
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Corpo do email formatado
        html = f"""
        <h1 style="color: #0066cc;">📊 Cotações PTAX</h1>
        <p><strong>Período:</strong> {dataframe['dataHoraCotacao'].min()} a {dataframe['dataHoraCotacao'].max()}</p>
        {dataframe.to_html(index=False, border=1)}
        <p><em>Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}</em></p>
        """
        
        msg.attach(MIMEText(html, 'html'))
        
        # Envio seguro
        with smtplib.SMTP(smtp_config["server"], smtp_config["port"]) as server:
            server.starttls()
            server.login(smtp_config["user"], smtp_config["password"])
            server.send_message(msg)
        
        return True
        
    except Exception as e:
        st.error(f"Falha no envio: {str(e)}")
        return False

# Barra lateral com controles
with st.sidebar:
    st.header("🔍 Filtros")
    
    # Seleção de moedas
    moedas = st.multiselect(
        "Selecione as moedas",
        ["Dólar", "Euro"],
        default=["Dólar"]
    )
    
    # Seleção de período
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Data inicial", datetime.now() - timedelta(days=7))
    with col2:
        end_date = st.date_input("Data final", datetime.now())
    
    # Tipo de cotação
    tipo_cotacao = st.radio(
        "Tipo de cotação",
        ["Compra", "Venda"],
        index=0,
        horizontal=True
    )

    # Controle de envio de email
    st.header("📤 Envio por Email")
    email_to = st.text_input("Destinatário", "exemplo@email.com")
    email_subject = st.text_input("Assunto", f"Cotações PTAX - {datetime.now().strftime('%d/%m/%Y')}")
    
    if st.button("Enviar Relatório", type="primary"):
        if 'df' in globals() and not df.empty:
            with st.spinner("Enviando email..."):
                if send_email(df, email_to, email_subject):
                    st.success("Email enviado com sucesso!")
                    st.balloons()
                else:
                    st.error("Falha ao enviar email")
        else:
            st.warning("Nenhum dado disponível para enviar")

# Busca e processamento dos dados
@st.cache_data(ttl=3600, show_spinner="Buscando dados...")
def load_data(moedas, start_date, end_date):
    dfs = []
    
    if "Dólar" in moedas:
        df_dolar = get_currency_data("dolar", start_date, end_date)
        if not df_dolar.empty:
            df_dolar["Moeda"] = "Dólar"
            df_dolar["dataHoraCotacao"] = pd.to_datetime(df_dolar["dataHoraCotacao"])
            dfs.append(df_dolar)
    
    if "Euro" in moedas:
        df_euro = get_currency_data("euro", start_date, end_date)
        if not df_euro.empty:
            df_euro["Moeda"] = "Euro"
            df_euro["dataHoraCotacao"] = pd.to_datetime(df_euro["dataHoraCotacao"])
            dfs.append(df_euro)
    
    return pd.concat(dfs) if dfs else pd.DataFrame()

df = load_data(moedas, start_date, end_date)

# Visualização principal
st.title("📈 Dashboard Cotações PTAX")
st.markdown("---")

if not df.empty:
    # Seção de KPIs dinâmicos
    st.header("📊 Indicadores Chave")
    col1, col2, col3, col4 = st.columns(4)
    
    ultima_cotacao = df.sort_values("dataHoraCotacao").groupby("Moeda").last()
    coluna = 'cotacaoCompra' if tipo_cotacao == "Compra" else 'cotacaoVenda'
    
    # Função auxiliar para métricas
    def create_metric(label, currency):
        if currency in moedas:
            current = ultima_cotacao.loc[currency, coluna]
            mean = df[df["Moeda"]==currency][coluna].mean()
            delta = ((current - mean)/mean*100) if mean != 0 else 0
            st.metric(
                label=f"{label} ({tipo_cotacao})",
                value=f"R$ {current:.4f}",
                delta=f"{delta:.2f}%"
            )
        else:
            st.metric(label=label, value="-")
    
    with col1: create_metric("Dólar", "Dólar")
    with col2: create_metric("Euro", "Euro")
    with col3: 
        if "Dólar" in moedas:
            variation = df[df["Moeda"]=="Dólar"][coluna].iloc[-1] - df[df["Moeda"]=="Dólar"][coluna].iloc[0]
            percent = (variation/df[df["Moeda"]=="Dólar"][coluna].iloc[0]*100) if df[df["Moeda"]=="Dólar"][coluna].iloc[0] != 0 else 0
            st.metric("Variação Dólar", f"R$ {variation:.4f}", f"{percent:.2f}%")
        else:
            st.metric("Variação Dólar", "-")
    with col4:
        if "Euro" in moedas:
            variation = df[df["Moeda"]=="Euro"][coluna].iloc[-1] - df[df["Moeda"]=="Euro"][coluna].iloc[0]
            percent = (variation/df[df["Moeda"]=="Euro"][coluna].iloc[0]*100) if df[df["Moeda"]=="Euro"][coluna].iloc[0] != 0 else 0
            st.metric("Variação Euro", f"R$ {variation:.4f}", f"{percent:.2f}%")
        else:
            st.metric("Variação Euro", "-")

    # Gráfico interativo
    st.markdown("---")
    st.header("📅 Evolução Temporal")
    
    fig = px.line(
        df,
        x="dataHoraCotacao",
        y=coluna,
        color="Moeda",
        title=f"Cotação de {tipo_cotacao}",
        labels={coluna: "Valor (R$)", "dataHoraCotacao": "Data"},
        hover_data={"Moeda": True, coluna: ":.4f"}
    )
    st.plotly_chart(fig, use_container_width=True)

    # Tabela detalhada
    st.markdown("---")
    st.header("📋 Dados Completos")
    
    df_display = df[["Moeda", "dataHoraCotacao", "cotacaoCompra", "cotacaoVenda"]].copy()
    df_display.columns = ["Moeda", "Data/Hora", "Compra (R$)", "Venda (R$)"]
    df_display["Data/Hora"] = df_display["Data/Hora"].dt.strftime("%d/%m/%Y %H:%M")
    
    st.dataframe(
        df_display.style.format({
            "Compra (R$)": "{:.4f}",
            "Venda (R$)": "{:.4f}"
        }),
        height=400,
        use_container_width=True
    )
    
else:
    st.warning("⚠️ Nenhum dado encontrado para os filtros selecionados")
    st.info("Dicas: Verifique se as datas são dias úteis e se a API do BC está acessível")

# Estilos CSS melhorados
st.markdown("""
<style>
    .stMetric {
        background: white;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stMetric label {
        font-weight: 600;
        color: #444;
    }
    .stMetric value {
        font-size: 1.5rem;
    }
    .stButton>button {
        background-color: #0066cc;
        color: white;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)