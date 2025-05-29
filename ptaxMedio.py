import requests
from datetime import datetime, timedelta
from collections import OrderedDict
import win32com.client
import calendar
import locale

locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

def truncate(number, decimals=0):
    factor = 10 ** decimals
    return int(number * factor) / factor

def ultimo_dia_util(year, month):
    last_day = calendar.monthrange(year, month)[1]
    day = datetime(year, month, last_day)
    while day.weekday() >= 5:
        day -= timedelta(days=1)
    return day

today = datetime.now()
year = today.year
month = today.month
last_day_month = calendar.monthrange(year, month)[1]
ultimo_util = ultimo_dia_util(year, month)


if today.date() != ultimo_util.date():
    print("O script só pode rodar no último dia ÚTIL do mês.")
    exit()

data_inicio = datetime(year, month, 1)
data_fim = datetime(year, month, last_day_month)
inicio = data_inicio.strftime("%m-%d-%Y")
fim = data_fim.strftime("%m-%d-%Y")

if month == 12:
    next_month = 1
    next_year = year + 1
else:
    next_month = month + 1
    next_year = year
mes_seguinte = datetime(next_year, next_month, 1).strftime('%B').capitalize()
titulo_email = f"Dólar Médio {year} - {mes_seguinte}"

url = (
    "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
    "CotacaoDolarPeriodo(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)"
    f"?@dataInicial='{inicio}'"
    f"&@dataFinalCotacao='{fim}'"
    "&$orderby=dataHoraCotacao desc"
    "&$top=1000"
    "&$format=json"
)

resp = requests.get(url)
resp.raise_for_status()
cotacoes = resp.json().get('value', [])
unico_por_dia = OrderedDict()
for registro in cotacoes:
    dia = registro['dataHoraCotacao'][:10]
    if dia not in unico_por_dia:
        unico_por_dia[dia] = registro['cotacaoVenda']

unico_por_dia = dict(sorted(unico_por_dia.items(), key=lambda x: x[0]))
valores = list(unico_por_dia.values())
media = truncate(sum(valores) / len(valores), 4)

html = """
<table style="border-collapse:collapse;font-family:Arial;font-size:13px;">
  <tr style="background-color:#efe4c6;">
    <th style="padding:5px 12px;">Data</th>
    <th style="padding:5px 12px;">Venda</th>
  </tr>
"""
for dia, venda in unico_por_dia.items():
    html += f'<tr style="background-color:#faf7ee;"><td style="padding:5px 12px;">{datetime.strptime(dia, "%Y-%m-%d").strftime("%d/%m/%Y")}</td><td style="padding:5px 12px;">{venda:.4f}</td></tr>'
html += f"""<tr>
    <td style="padding:5px 12px;font-weight:bold;background:#fffbe4;text-align:right;" colspan="1"></td>
    <td style="padding:5px 12px;font-weight:bold;background:#fffbe4;color:#ffce1a;font-size:16px;">{media:.4f}</td>
  </tr>
</table>
"""

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = "destinatario@empresa.com"
mail.Subject = titulo_email
mail.HTMLBody = f"<h3>Cotações do Dólar PTAX - {today.strftime('%B').capitalize()}/{year}</h3>{html}"
mail.Send()
