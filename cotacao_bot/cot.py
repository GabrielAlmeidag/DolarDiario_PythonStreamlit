import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import pythoncom
import win32com.client as win32

# — Page configuration —
st.set_page_config(
    page_title="Dashboard Avançado PTAX",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# — Custom CSS for enhanced styling —
st.markdown("""
<style>
  /* Main container styling */
  .main {
    background-color: #f8f9fa;
  }
  
  /* Enhanced metric cards */
  .metric-card {
    background: white;
    border-radius: 10px;
    padding: 1.5rem;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    transition: all 0.3s ease;
    border-left: 4px solid #4CAF50;
    margin-bottom: 1rem;
  }
  .metric-card:hover {
    transform: translateY(-5px);
    box-shadow: 0 10px 20px rgba(0,0,0,0.1);
  }
  .metric-card h3 {
    color: #555;
    font-size: 1rem;
    margin-bottom: 0.5rem;
  }
  .metric-card .value {
    font-size: 1.8rem;
    font-weight: 700;
    color: #333;
  }
  .metric-card .delta {
    font-size: 0.9rem;
    display: flex;
    align-items: center;
  }
  .delta.up {
    color: #28a745;
  }
  .delta.down {
    color: #dc3545;
  }
  
  /* Tabs styling */
  .stTabs [role="tablist"] {
    gap: 10px;
  }
  .stTabs [role="tab"] {
    border-radius: 8px 8px 0 0 !important;
    padding: 8px 16px !important;
    background: #e9ecef !important;
    border: 1px solid #dee2e6 !important;
  }
  .stTabs [role="tab"][aria-selected="true"] {
    background: #4CAF50 !important;
    color: white !important;
    border-color: #4CAF50 !important;
  }
  
  /* Sidebar styling */
  [data-testid="stSidebar"] {
    background: #f8f9fa !important;
    border-right: 1px solid #dee2e6;
  }
  
  /* Custom columns for metrics */
  .metric-column {
    padding: 0 10px;
  }
  
  /* Custom expander styling */
  .stExpander {
    border: 1px solid #dee2e6;
    border-radius: 8px;
  }
</style>
""", unsafe_allow_html=True)

# — Data fetching functions —
@st.cache_data(ttl=3600, show_spinner="Carregando dados do BCB...")
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
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        data = r.json().get("value", [])
        if data:
            df = pd.DataFrame(data)
            df["Moeda"] = code
            df["dataHoraCotacao"] = pd.to_datetime(df["dataHoraCotacao"])
            df["Dia"] = df["dataHoraCotacao"].dt.date
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao buscar dados para {code}: {str(e)}")
        return pd.DataFrame()

@st.cache_data(ttl=3600)
def load_data(codes, start_date, end_date):
    frames = []
    for c in codes:
        dfc = get_currency_data(c, start_date, end_date)
        if not dfc.empty:
            frames.append(dfc)
    if frames:
        df = pd.concat(frames, ignore_index=True)
        # Calculate daily statistics
        daily_stats = df.groupby(["Moeda", "Dia"]).agg({
            "cotacaoCompra": ["first", "last", "min", "max", "mean"],
            "cotacaoVenda": ["first", "last", "min", "max", "mean"]
        }).reset_index()
        daily_stats.columns = ["Moeda", "Dia", 
                              "Compra_Inicial", "Compra_Final", "Compra_Min", "Compra_Max", "Compra_Media",
                              "Venda_Inicial", "Venda_Final", "Venda_Min", "Venda_Max", "Venda_Media"]
        # Calculate variations
        for t in ["Compra", "Venda"]:
            daily_stats[f"{t}_Var_Dia"] = ((daily_stats[f"{t}_Final"] - daily_stats[f"{t}_Inicial"]) / 
                                         daily_stats[f"{t}_Inicial"]) * 100
            daily_stats[f"{t}_Var_Max"] = ((daily_stats[f"{t}_Max"] - daily_stats[f"{t}_Inicial"]) / 
                                         daily_stats[f"{t}_Inicial"]) * 100
            daily_stats[f"{t}_Var_Min"] = ((daily_stats[f"{t}_Min"] - daily_stats[f"{t}_Inicial"]) / 
                                         daily_stats[f"{t}_Inicial"]) * 100
        return df, daily_stats
    return pd.DataFrame(), pd.DataFrame()

def send_email(df, to_email, subject, analysis_text=""):
    if df.empty:
        return False
    
    # Create styled HTML table
    html = df.to_html(
        index=False, border=0, justify="center",
        classes="table table-striped table-bordered",
        formatters={
            "Compra (R$)": lambda x: f"R$ {x:.4f}",
            "Venda (R$)": lambda x: f"R$ {x:.4f}",
            "Variação (%)": lambda x: f"{x:.2f}%"
        }
    )
    
    body = f"""
    <html>
    <head>
    <style>
      body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
      .container {{ max-width: 900px; margin: 0 auto; padding: 20px; }}
      h1 {{ color: #2c3e50; border-bottom: 2px solid #4CAF50; padding-bottom: 10px; }}
      .table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
      .table th {{ background-color: #4CAF50; color: white; padding: 10px; text-align: left; }}
      .table td {{ padding: 8px; border: 1px solid #ddd; }}
      .table tr:nth-child(even) {{ background-color: #f2f2f2; }}
      .analysis {{ background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-top: 20px; }}
      .footer {{ margin-top: 30px; font-size: 0.8em; color: #777; text-align: center; }}
      .positive {{ color: #28a745; font-weight: bold; }}
      .negative {{ color: #dc3545; font-weight: bold; }}
    </style>
    </head>
    <body>
    <div class="container">
      <h1>📊 Relatório de Cotações PTAX</h1>
      <h2>{subject}</h2>
      {html}
      <div class="analysis">
        <h3>Análise</h3>
        <p>{analysis_text}</p>
      </div>
      <div class="footer">
        <p>Relatório gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
      </div>
    </div>
    </body>
    </html>
    """
    
    try:
        pythoncom.CoInitialize()
        mail = win32.Dispatch("outlook.application").CreateItem(0)
        mail.To = to_email
        mail.Subject = subject
        mail.HTMLBody = body
        mail.Send()
        return True
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {str(e)}")
        return False
    finally:
        pythoncom.CoUninitialize()

# — Sidebar controls —
st.sidebar.header("⚙️ Configurações")
today = datetime.now().date()

with st.sidebar.expander("🔍 Filtros de Período", expanded=True):
    date_range_type = st.radio(
        "Tipo de período",
        ["Hoje", "Últimos 7 dias", "Últimos 30 dias", "Personalizado"],
        index=0,
        horizontal=True
    )
    
    if date_range_type == "Hoje":
        start_date = end_date = today
    elif date_range_type == "Últimos 7 dias":
        start_date = today - timedelta(days=7)
        end_date = today
    elif date_range_type == "Últimos 30 dias":
        start_date = today - timedelta(days=30)
        end_date = today
    else:
        date_range = st.date_input(
            "Selecionar período",
            [today - timedelta(days=7), today],
            min_value=today - timedelta(days=365),
            max_value=today
        )
        if len(date_range) == 2:
            start_date, end_date = date_range
        else:
            start_date = end_date = date_range[0]

with st.sidebar.expander("💰 Moedas", expanded=True):
    available_currencies = ["USD", "EUR", "GBP", "JPY", "CHF", "AUD", "CAD"]
    codes = st.multiselect(
        "Selecione as moedas",
        available_currencies,
        ["USD", "EUR"],
        key="currency_select"
    )

with st.sidebar.expander("📊 Opções de Análise", expanded=True):
    quote_type = st.radio(
        "Tipo de cotação",
        ["Compra", "Venda"],
        index=0,
        horizontal=True
    )
    
    analysis_level = st.radio(
        "Nível de análise",
        ["Diário", "Intradiário"],
        index=0,
        help="Diário: análise por dia. Intradiário: análise por horário dentro do dia."
    )
    
    show_benchmark = st.checkbox(
        "Mostrar benchmark (USD)",
        True,
        help="Mostrar o dólar como referência em gráficos comparativos"
    )

# — Load data —
df, daily_stats = load_data(codes, start_date, end_date)

if df.empty:
    st.warning("⚠️ Nenhum dado disponível para o período selecionado.")
    st.stop()

# Prepare data for display
latest_df = df.sort_values("dataHoraCotacao").groupby("Moeda").last().reset_index()
latest_df["Data/Hora"] = latest_df["dataHoraCotacao"].dt.strftime("%d/%m/%Y %H:%M")
latest_df = latest_df.rename(columns={
    "cotacaoCompra": "Compra (R$)", 
    "cotacaoVenda": "Venda (R$)"
})

# Calculate metrics for display
metrics_data = []
for c in codes:
    subset = df[df["Moeda"] == c]
    if not subset.empty:
        latest = subset.iloc[-1]
        first = subset.iloc[0]
        
        compra_var = ((latest["cotacaoCompra"] - first["cotacaoCompra"]) / first["cotacaoCompra"]) * 100
        venda_var = ((latest["cotacaoVenda"] - first["cotacaoVenda"]) / first["cotacaoVenda"]) * 100
        
        metrics_data.append({
            "Moeda": c,
            "Compra_Inicial": first["cotacaoCompra"],
            "Compra_Final": latest["cotacaoCompra"],
            "Compra_Var": compra_var,
            "Venda_Inicial": first["cotacaoVenda"],
            "Venda_Final": latest["cotacaoVenda"],
            "Venda_Var": venda_var,
            "Compra_Min": subset["cotacaoCompra"].min(),
            "Compra_Max": subset["cotacaoCompra"].max(),
            "Venda_Min": subset["cotacaoVenda"].min(),
            "Venda_Max": subset["cotacaoVenda"].max(),
            "Data_Inicial": first["dataHoraCotacao"],
            "Data_Final": latest["dataHoraCotacao"]
        })

metrics_df = pd.DataFrame(metrics_data)

# — Main UI —
st.title("📊 Dashboard Avançado de Cotações PTAX")
st.caption(f"Período: {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}")

# Row 1: Key Metrics
st.subheader("📌 Principais Indicadores")
cols = st.columns(len(codes) + (1 if show_benchmark and "USD" not in codes else 0))

for i, c in enumerate(codes):
    with cols[i]:
        if c in metrics_df["Moeda"].values:
            data = metrics_df[metrics_df["Moeda"] == c].iloc[0]
            current_value = data[f"{quote_type}_Final"]
            variation = data[f"{quote_type}_Var"]
            
            st.markdown(f"""
            <div class="metric-card">
                <h3>{c} - {quote_type}</h3>
                <div class="value">R$ {current_value:.4f}</div>
                <div class="delta {'up' if variation >= 0 else 'down'}">
                    {'+' if variation >= 0 else ''}{variation:.2f}% 
                    <small>no período</small>
                </div>
                <div style="font-size: 0.8rem; margin-top: 8px; color: #666;">
                    Mín: R$ {data[f"{quote_type}_Min"]:.4f}<br>
                    Máx: R$ {data[f"{quote_type}_Max"]:.4f}
                </div>
            </div>
            """, unsafe_allow_html=True)

# Add USD benchmark if requested
if show_benchmark and "USD" not in codes:
    with cols[-1]:
        usd_data = get_currency_data("USD", start_date, end_date)
        if not usd_data.empty:
            usd_latest = usd_data.iloc[-1]
            usd_first = usd_data.iloc[0]
            usd_var = ((usd_latest["cotacaoCompra"] - usd_first["cotacaoCompra"]) / usd_first["cotacaoCompra"]) * 100
            
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ff9800;">
                <h3>USD - Benchmark</h3>
                <div class="value">R$ {usd_latest['cotacaoCompra']:.4f}</div>
                <div class="delta {'up' if usd_var >= 0 else 'down'}">
                    {'+' if usd_var >= 0 else ''}{usd_var:.2f}% 
                    <small>no período</small>
                </div>
            </div>
            """, unsafe_allow_html=True)

# Tabs for different views
tab1, tab2, tab3, tab4 = st.tabs(["📈 Análise Temporal", "🔄 Comparativo", "📋 Dados Detalhados", "📤 Exportar"])

with tab1:
    st.subheader(f"Análise Temporal - {quote_type}")
    
    if analysis_level == "Diário" and not daily_stats.empty:
        # Daily analysis
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        for c in codes:
            c_data = daily_stats[daily_stats["Moeda"] == c]
            fig.add_trace(
                go.Scatter(
                    x=c_data["Dia"],
                    y=c_data[f"{quote_type}_Media"],
                    name=f"{c} - Média",
                    mode="lines+markers",
                    line=dict(width=2),
                    marker=dict(size=8)
                ),
                secondary_y=False
            )
            
            # Add range (min-max)
            fig.add_trace(
                go.Scatter(
                    x=pd.concat([c_data["Dia"], c_data["Dia"][::-1]]),
                    y=pd.concat([c_data[f"{quote_type}_Max"], c_data[f"{quote_type}_Min"][::-1]]),
                    fill="toself",
                    fillcolor="rgba(0,100,80,0.2)",
                    line=dict(color="rgba(255,255,255,0)"),
                    hoverinfo="skip",
                    name=f"{c} - Variação",
                    showlegend=False
                ),
                secondary_y=False
            )
        
        fig.update_layout(
            height=500,
            template="plotly_white",
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=20, r=20, t=40, b=20),
            xaxis_title="Data",
            yaxis_title=f"Valor de {quote_type} (R$)"
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        # Intraday analysis
        fig = px.line(
            df, 
            x="dataHoraCotacao", 
            y=f"cotacao{quote_type}", 
            color="Moeda",
            labels={"dataHoraCotacao": "Data/Hora", f"cotacao{quote_type}": f"Valor {quote_type} (R$)"},
            template="plotly_white"
        )
        
        fig.update_traces(
            mode="lines+markers",
            marker=dict(size=5),
            line=dict(width=2)
        )
        
        fig.update_layout(
            height=500,
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=20, r=20, t=40, b=20),
            xaxis=dict(
                rangeslider=dict(visible=True),
                type="date"
            )
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    # Add statistics section
    st.subheader("📊 Estatísticas Descritivas")
    
    stats_cols = st.columns(2)
    
    with stats_cols[0]:
        st.markdown("**Variação Diária**")
        if not daily_stats.empty:
            fig = px.bar(
                daily_stats,
                x="Dia",
                y=f"{quote_type}_Var_Dia",
                color="Moeda",
                barmode="group",
                labels={f"{quote_type}_Var_Dia": "Variação (%)", "Dia": "Data"},
                template="plotly_white"
            )
            fig.update_layout(height=300)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Dados diários não disponíveis para o período selecionado.")
    
    with stats_cols[1]:
        st.markdown("**Volatilidade (Máxima e Mínima)**")
        if not daily_stats.empty:
            fig = go.Figure()
            
            for c in codes:
                c_data = daily_stats[daily_stats["Moeda"] == c]
                fig.add_trace(go.Bar(
                    x=c_data["Dia"],
                    y=c_data[f"{quote_type}_Var_Max"],
                    name=f"{c} - Máxima",
                    marker_color="green"
                ))
                fig.add_trace(go.Bar(
                    x=c_data["Dia"],
                    y=c_data[f"{quote_type}_Var_Min"],
                    name=f"{c} - Mínima",
                    marker_color="red"
                ))
            
            fig.update_layout(
                barmode="group",
                height=300,
                template="plotly_white",
                yaxis_title="Variação (%)"
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Dados diários não disponíveis para o período selecionado.")

with tab2:
    st.subheader("Análise Comparativa entre Moedas")
    
    if len(codes) > 1:
        comp_cols = st.columns(2)
        
        with comp_cols[0]:
            st.markdown("**Evolução Comparativa (Base 100)**")
            
            # Normalize to base 100 for comparison
            comparison_data = []
            for c in codes:
                c_df = df[df["Moeda"] == c].copy()
                if not c_df.empty:
                    base_value = c_df.iloc[0][f"cotacao{quote_type}"]
                    c_df["Normalized"] = (c_df[f"cotacao{quote_type}"] / base_value) * 100
                    comparison_data.append(c_df)
            
            if comparison_data:
                comparison_df = pd.concat(comparison_data)
                fig = px.line(
                    comparison_df,
                    x="dataHoraCotacao",
                    y="Normalized",
                    color="Moeda",
                    labels={"dataHoraCotacao": "Data/Hora", "Normalized": "Valor Normalizado (Base 100)"},
                    template="plotly_white"
                )
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
        
        with comp_cols[1]:
            st.markdown("**Correlação entre Moedas**")
            
            # Create pivot table for correlation
            pivot_df = df.pivot(index="dataHoraCotacao", columns="Moeda", values=f"cotacao{quote_type}")
            corr_matrix = pivot_df.corr()
            
            fig = go.Figure(data=go.Heatmap(  # ← Parênteses abertos corretamente
                z=corr_matrix.values,
                x=corr_matrix.columns,
                y=corr_matrix.index,
                colorscale="Viridis",
                zmin=-1,
                zmax=1,
                colorbar=dict(title="Correlação")
            ))  # ← Parênteses fechados corretamente
             
            fig.update_layout(
                height=400,
                xaxis_title="Moeda",
                yaxis_title="Moeda"
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Display correlation values
            st.markdown("**Valores de Correlação**")
            st.dataframe(corr_matrix.style.background_gradient(cmap="viridis", vmin=-1, vmax=1))
    
    else:
        st.info("Selecione pelo menos duas moedas para análise comparativa.")

with tab3:
    st.subheader("Dados Detalhados")
    
    if analysis_level == "Diário" and not daily_stats.empty:
        display_df = daily_stats.copy()
        display_df["Dia"] = display_df["Dia"].apply(lambda x: x.strftime("%d/%m/%Y"))
        
        # Format columns
        for col in ["Compra_Inicial", "Compra_Final", "Compra_Min", "Compra_Max", "Compra_Media",
                   "Venda_Inicial", "Venda_Final", "Venda_Min", "Venda_Max", "Venda_Media"]:
            display_df[col] = display_df[col].apply(lambda x: f"R$ {x:.4f}")
        
        for col in ["Compra_Var_Dia", "Compra_Var_Max", "Compra_Var_Min",
                   "Venda_Var_Dia", "Venda_Var_Max", "Venda_Var_Min"]:
            display_df[col] = display_df[col].apply(lambda x: f"{x:.2f}%")
        
        st.dataframe(display_df, use_container_width=True)
    else:
        display_df = df.copy()
        display_df["dataHoraCotacao"] = display_df["dataHoraCotacao"].dt.strftime("%d/%m/%Y %H:%M")
        display_df = display_df.rename(columns={
            "dataHoraCotacao": "Data/Hora",
            "cotacaoCompra": "Compra (R$)",
            "cotacaoVenda": "Venda (R$)"
        })
        st.dataframe(display_df[["Moeda", "Data/Hora", "Compra (R$)", "Venda (R$)"]], use_container_width=True)

with tab4:
    st.subheader("Exportar Dados e Análise")
    
    export_cols = st.columns(2)
    
    with export_cols[0]:
        st.markdown("**Exportar Dados**")
        
        export_format = st.radio(
            "Formato de exportação",
            ["CSV", "Excel"],
            index=0,
            horizontal=True
        )
        
        if st.button("⬇️ Baixar Dados Completos"):
            if export_format == "CSV":
                csv = df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="Clique para baixar",
                    data=csv,
                    file_name=f"ptax_data_{start_date}_{end_date}.csv",
                    mime="text/csv"
                )
            else:
                excel = df.to_excel(index=False)
                st.download_button(
                    label="Clique para baixar",
                    data=excel,
                    file_name=f"ptax_data_{start_date}_{end_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    with export_cols[1]:
        st.markdown("**Enviar por E-mail**")
        
        email_to = st.text_input("Destinatário", "")
        email_subject = st.text_input("Assunto", f"Análise PTAX {start_date} a {end_date}")
        
        analysis_text = st.text_area(
            "Adicionar análise personalizada",
            "Segue análise das cotações PTAX para o período selecionado. "
            "Destacam-se as seguintes observações:\n\n"
            "- Variação média no período\n"
            "- Comportamento comparativo entre moedas\n"
            "- Principais pontos de máxima e mínima"
        )
        
        if st.button("📤 Enviar Relatório Completo"):
            if not email_to:
                st.warning("Por favor, informe um destinatário")
            else:
                with st.spinner("Enviando e-mail..."):
                    # Prepare data for email
                    email_df = latest_df[["Moeda", "Data/Hora", "Compra (R$)", "Venda (R$)"]].copy()
                    email_df["Variação (%)"] = email_df.apply(
                        lambda row: ((row["Compra (R$)"] - metrics_df[metrics_df["Moeda"] == row["Moeda"]]["Compra_Inicial"].values[0]) / 
                                   metrics_df[metrics_df["Moeda"] == row["Moeda"]]["Compra_Inicial"].values[0]) * 100,
                        axis=1
                    )
                    
                    success = send_email(email_df, email_to, email_subject, analysis_text)
                    if success:
                        st.success("E-mail enviado com sucesso!")
                    else:
                        st.error("Falha ao enviar o e-mail")