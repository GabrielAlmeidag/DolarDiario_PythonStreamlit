import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import plotly.express as px
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import atexit
import time

# Configuração da página
st.set_page_config(
    page_title="Dashboard Cotações PTAX",
    page_icon="📈",
    layout="wide"
)

# Inicializa o scheduler
if 'scheduler' not in st.session_state:
    scheduler = BackgroundScheduler()
    scheduler.start()
    st.session_state.scheduler = scheduler
    atexit.register(lambda: scheduler.shutdown())

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

# Função para enviar email
def send_email(dataframe, to_email, subject):
    try:
        # Configurações do servidor SMTP (substitua com suas credenciais)
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        smtp_user = "seu_email@gmail.com"  # Substitua pelo seu email
        smtp_password = "sua_senha"       # Substitua pela sua senha/app password
        
        # Criar mensagem
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Criar corpo do email
        body = f"""
        <h1>Cotações PTAX - {datetime.now().strftime('%d/%m/%Y')}</h1>
        <p>Segue abaixo as cotações solicitadas:</p>
        {dataframe.to_html()}
        """
        
        msg.attach(MIMEText(body, 'html'))
        
        # Enviar email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)
        
        return True
    except Exception as e:
        st.error(f"Erro ao enviar email: {e}")
        return False

# Barra lateral com filtros
with st.sidebar:
    st.header("Filtros")
    
    # Filtro de moedas
    moedas = st.multiselect(
        "Selecione as moedas",
        ["Dólar", "Euro"],
        default=["Dólar"]
    )
    
    # Filtro de datas
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Data inicial", datetime.now() - timedelta(days=30))
    with col2:
        end_date = st.date_input("Data final", datetime.now())
    
    # Filtro de tipo de cotação
    tipo_cotacao = st.radio(
        "Tipo de cotação",
        ["Compra", "Venda"],
        index=0,
        horizontal=True
    )

    st.header("Envio de Email")
    email_enabled = st.checkbox("Ativar envio de emails")
    
    if email_enabled:
        email_to = st.text_input("Destinatário", "destinatario@email.com")
        email_subject = st.text_input("Assunto", f"Cotações PTAX - {datetime.now().strftime('%d/%m/%Y')}")
        
        frequency = st.selectbox(
            "Frequência de envio",
            ["Único", "Diário", "Semanal", "Mensal"],
            index=0
        )
        
        if frequency != "Único":
            send_time = st.time_input("Horário do envio", datetime.now().time())
            
            if frequency == "Semanal":
                day_of_week = st.selectbox("Dia da semana", ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"])
            elif frequency == "Mensal":
                day_of_month = st.number_input("Dia do mês", min_value=1, max_value=28, value=1)
        
        if st.button("Agendar Envio"):
            if 'df' in globals() and 'df' in locals() and not df.empty:
                if frequency == "Único":
                    if 'df' in globals() and 'df' in locals() and not df.empty and send_email(df, email_to, email_subject):
                        st.success("Email enviado com sucesso!")
                else:
                    # Configura o agendamento recorrente
                    if frequency == "Diário":
                        trigger = CronTrigger(
                            hour=send_time.hour,
                            minute=send_time.minute,
                            day_of_week='*'
                        )
                    elif frequency == "Semanal":
                        days_map = {"Segunda": "mon", "Terça": "tue", "Quarta": "wed", 
                                   "Quinta": "thu", "Sexta": "fri", "Sábado": "sat", "Domingo": "sun"}
                        trigger = CronTrigger(
                            hour=send_time.hour,
                            minute=send_time.minute,
                            day_of_week=days_map[day_of_week]
                        )
                    elif frequency == "Mensal":
                        trigger = CronTrigger(
                            hour=send_time.hour,
                            minute=send_time.minute,
                            day=day_of_month
                        )
                    
                    st.session_state.scheduler.add_job(
                        send_email,
                        trigger,
                        args=[df, email_to, email_subject]
                    )
                    st.success(f"Email agendado para envio {frequency.lower()} às {send_time.strftime('%H:%M')}")
            else:
                st.warning("Nenhum dado disponível para enviar")

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
st.title("📈 Dashboard de Cotações PTAX")
st.markdown("---")

# Seção de KPIs
if not df.empty:
    st.header("Indicadores Chave")
    col1, col2, col3, col4 = st.columns(4)
    
    ultima_cotacao = df.sort_values("dataHoraCotacao").groupby("Moeda").last()
    coluna = 'cotacaoCompra' if tipo_cotacao == "Compra" else 'cotacaoVenda'
    
    with col1:
        if "Dólar" in moedas:
            current = ultima_cotacao.loc['Dólar', coluna]
            mean = df_dolar[coluna].mean()
            delta = ((current - mean)/mean*100) if mean != 0 else 0
            st.metric(
                label=f"Cotação Atual Dólar ({tipo_cotacao})",
                value=f"R$ {current:.4f}",
                delta=f"{delta:.2f}%"
            )
        else:
            st.metric(label="Cotação Dólar", value="-")
    
    with col2:
        if "Euro" in moedas:
            current = ultima_cotacao.loc['Euro', coluna]
            mean = df_euro[coluna].mean()
            delta = ((current - mean)/mean*100) if mean != 0 else 0
            st.metric(
                label=f"Cotação Atual Euro ({tipo_cotacao})",
                value=f"R$ {current:.4f}",
                delta=f"{delta:.2f}%"
            )
        else:
            st.metric(label="Cotação Euro", value="-")
    
    with col3:
        if "Dólar" in moedas and len(df_dolar) > 1:
            variation = df_dolar[coluna].iloc[-1] - df_dolar[coluna].iloc[0]
            percent = (variation/df_dolar[coluna].iloc[0]*100) if df_dolar[coluna].iloc[0] != 0 else 0
            st.metric(
                label=f"Variação Dólar ({tipo_cotacao})",
                value=f"R$ {variation:.4f}",
                delta=f"{percent:.2f}%"
            )
        else:
            st.metric(label="Variação Dólar", value="-")
    
    with col4:
        if "Euro" in moedas and len(df_euro) > 1:
            variation = df_euro[coluna].iloc[-1] - df_euro[coluna].iloc[0]
            percent = (variation/df_euro[coluna].iloc[0]*100) if df_euro[coluna].iloc[0] != 0 else 0
            st.metric(
                label=f"Variação Euro ({tipo_cotacao})",
                value=f"R$ {variation:.4f}",
                delta=f"{percent:.2f}%"
            )
        else:
            st.metric(label="Variação Euro", value="-")

# Gráficos
st.markdown("---")
st.header("Análise Temporal")

if not df.empty:
    coluna = 'cotacaoCompra' if tipo_cotacao == "Compra" else 'cotacaoVenda'
    
    fig = px.line(
        df,
        x="dataHoraCotacao",
        y=coluna,
        color="Moeda",
        title=f"Evolução da Cotação de {tipo_cotacao}",
        labels={
            "dataHoraCotacao": "Data",
            coluna: f"Valor (R$)"
        }
    )
    fig.update_layout(
        hovermode="x unified",
        xaxis_title="Data",
        yaxis_title=f"Valor de {tipo_cotacao} (R$)",
        legend_title="Moeda"
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
        use_container_width=True,
        height=400
    )
else:
    st.warning("Nenhum dado disponível para os filtros selecionados")

# Estilo CSS adicional
st.markdown("""
<style>
    .stMetric {
        border: 1px solid #e6e6e6;
        border-radius: 8px;
        padding: 15px;
        background-color: #f9f9f9;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .stMetric label {
        font-size: 14px;
        color: #555;
        font-weight: bold;
    }
    .stMetric value {
        font-size: 24px;
        font-weight: bold;
    }
    .css-1aumxhk {
        background-color: #f0f2f6;
    }
</style>
""", unsafe_allow_html=True)