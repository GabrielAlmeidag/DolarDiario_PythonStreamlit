# Dashboard Cotações PTAX

## Descrição
Este projeto oferece um dashboard desenvolvido em Streamlit para consulta, visualização e envio de cotações PTAX (dólar e euro) obtidas por meio da API do Banco Central do Brasil. Além dos gráficos e indicadores em tempo real, o sistema permite o envio de relatórios por e-mail de forma automatizada.

## Funcionalidades
- Seleção de moedas (dólar e euro)
- Definição de período para consulta de dados
- Escolha entre cotação de compra ou venda
- Indicadores chave (valor atual, média e variação percentual)
- Gráfico interativo de evolução temporal
- Tabela detalhada com histórico de cotações
- Envio de relatórios por e-mail, com corpo em HTML
- Log de execução (data e tempo de processamento)

## Requisitos
- Python 3.8 ou superior
- Streamlit
- Pandas
- Requests
- Plotly

## Instalação
1. Clone este repositório:
   ```bash
   git clone https://seu-repositorio.git
   cd seu-repositorio
   ```
2. Crie e ative um ambiente virtual (opcional, mas recomendado):
   ```bash
   python -m venv venv
   source venv/bin/activate  # Linux/macOS
   venv\Scripts\activate   # Windows
   ```
3. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

## Configuração
- Atualize as credenciais de e-mail no arquivo ou nas variáveis de ambiente:
  ```python
  smtp_config = {
      "server": "smtp.gmail.com",
      "port": 587,
      "user": "seu_email@gmail.com",
      "password": "sua_senha"
  }
  ```

## Uso
Para executar o dashboard, execute:
```bash
streamlit run app.py
```
Em seguida, acesse no navegador: `http://localhost:8501`.

### Envio de Relatório
- Preencha o destinatário e o assunto na barra lateral
- Clique em **Enviar Relatório**

## Estrutura de Arquivos
```
├── cot.py               # Código principal do Streamlit
├── requirements.txt     # Dependências do projeto
├── README.md            # Documentação deste projeto
├── ptaxMedio.py         # Calcula a média e envia por email
```

## Licença
Este projeto está disponível sob a licença MIT.
