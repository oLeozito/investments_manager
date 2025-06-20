import requests
from bs4 import BeautifulSoup

class FundoImobiliario:
    def __init__(self, codigo, dy_12m, dy_percent, preco_atual, segmento, tipo, val_patr, vacancia, qtdcotis, qtdcotas, cnpj, preco_teto, pvp, liquidez):
        self.codigo = codigo
        self.dy_12m = dy_12m
        self.dy_percent = dy_percent
        self.preco_atual = preco_atual
        self.segmento = segmento
        self.tipo = tipo
        self.val_patr = val_patr
        self.vacan = vacancia
        self.qtdcotis = qtdcotis
        self.qtdcotas = qtdcotas
        self.cnpj = cnpj
        self.pvp = pvp
        self.liquidez = liquidez

        self.preco_teto = preco_teto

# # Exemplo:
# fii = FundoImobiliario("XPML11", 11.04, 100.50, 92.00)
# fiis[fii.codigo] = fii


# Função busca codigos de Fundos Imobiliários.
def obter_fiis():
    letras = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    siglas = []
    
    url_fiis = "https://www.fundsexplorer.com.br/funds"
    headers = {'User-Agent': 'Mozilla/5.0'}

    resposta = requests.get(url_fiis, headers=headers)
    if resposta.status_code != 200:
        print(f"Erro ao acessar o site.")

    html_fiis = BeautifulSoup(resposta.text, 'html.parser')

    # Pega a lista geral de fiis dentro do HTML
    listar = html_fiis.find('div',class_='tickerFilter__results')

    # Pega as listas de cada letra
    for letra in letras:
        filtrar = listar.find('section',id=f"letter-id-{letra}")
        buscar = filtrar.find_all('div', attrs={'data-element': 'ticker-box-title'})
        for item in buscar:
            siglas.append(item.text.strip())
    
    return siglas

def obter_infos(fii):
    url = f'https://investidor10.com.br/fiis/{fii.lower()}/'
    headers = {'User-Agent': 'Mozilla/5.0'}

    # @Sigla / Segmento / Tipo(Tijolo/Papel) / @Preço da Cota / @DY Médio 12Meses % / @DY 12Meses em real / @Liquidez Diária / @P/VP / Valor Patrimônial / QTD Cotistas / Vacância /

    preco_cota = None
    dy12 = None
    pvp = None
    liquidez = None
    segmento = None
    tipo = None
    val_patri = None
    qtd_cotas = None
    qtd_cotis = None
    vacancia = None
    cnpj = None



    resposta = requests.get(url, headers=headers)
    if resposta.status_code != 200:
        print(f"Erro ao acessar o site para {fii}")
        return None, None

    soup = BeautifulSoup(resposta.text, 'html.parser')

    # --- DY 12M ---
    spans_dy = soup.find_all('span', class_='content--info--item--value amount')
    valores_rs = []

    for span in spans_dy:
        texto = span.text.strip()
        if 'R$' in texto:
            valor = texto.replace('R$', '').replace(',', '.').strip()
            try:
                valores_rs.append(float(valor))
            except ValueError:
                continue

    val_dy_12m = valores_rs[3] if len(valores_rs) >= 4 else None

    # --- Cota atual ---
    span_valores = soup.find_all('span', class_='value')
    for span in span_valores:
        texto = span.text.strip()
        if 'R$' in texto:
            valor = texto.replace('R$', '').replace(',', '.').strip()
            try:
                preco_cota = float(valor)
                break  # pegamos o primeiro valor R$ da cota
            except ValueError:
                continue

#   Buscando o Dividend Yeld 12 Meses em Porcentagem.
    card = soup.find('div', class_='_card dy')  # passo 1: acha o card
    if card:
        corpo = card.find('div', class_='_card-body')  # passo 2: entra no "_card-body"
        if corpo:
            span = corpo.find('span')  # passo 3: pega o primeiro span dentro do _card-body
            if span:
                dy12 = span.text.strip()
            else:
                print("Span não encontrado.")
        else:
            print("Div _card-body não encontrada.")
    else:
        print("Card _card dy não encontrado.")

    #Buscando P/VP
    card = soup.find('div', class_='_card vp')  # passo 1: acha o card
    if card:
        corpo = card.find('div', class_='_card-body')  # passo 2: entra no "_card-body"
        if corpo:
            span = corpo.find('span')  # passo 3: pega o primeiro span dentro do _card-body
            if span:
                pvp = span.text.strip()
            else:
                print("Span não encontrado.")
        else:
            print("Div _card-body não encontrada.")
    else:
        print("Card _card vp não encontrado.")

    #Buscando Liquidez Diária
    card = soup.find('div', class_='_card val')  # passo 1: acha o card
    if card:
        corpo = card.find('div', class_='_card-body')  # passo 2: entra no "_card-body"
        if corpo:
            span = corpo.find('span')  # passo 3: pega o primeiro span dentro do _card-body
            if span:
                liquidez = span.text.strip()
            else:
                print("Span não encontrado.")
        else:
            print("Div _card-body não encontrada.")
    else:
        print("Card _card val não encontrado.")

    #Varias Infos
    valores = []
    tabela = soup.find('div', id="table-indicators")
    if tabela:
        celulas = tabela.find_all('div', class_="cell")
        for cell in celulas:
            value = cell.find('div', class_="value")
            if value:
                valores.append(value.text.strip())
    else:
        print(f"Tabela de indicadores não encontrada para {fii}")


    def seguro(valores, i):
        return valores[i] if i < len(valores) else None

    cnpj       = seguro(valores, 1)
    segmento   = seguro(valores, 4)
    tipo       = seguro(valores, 5)
    vacancia   = seguro(valores, 9)
    qtd_cotis  = seguro(valores, 10)
    qtd_cotas  = seguro(valores, 11)
    val_patri  = seguro(valores, 13)

    
    rentabilidade_desejada = 12/100 #Rentabilidade de 12%a.a.
    if val_dy_12m == None:
        preco_teto = None
    else:
        preco_teto = val_dy_12m/rentabilidade_desejada

    fundo = FundoImobiliario(fii,val_dy_12m,dy12,preco_cota,segmento,tipo,val_patri,vacancia,qtd_cotis,qtd_cotas,cnpj,preco_teto,pvp,liquidez)

    return fundo

# --- Programa Principal ---
lista_fundos = {}
siglas = obter_fiis()
for sigla in siglas:
        print(sigla)
        lista_fundos[sigla] = obter_infos(sigla)
while True:   
    opcao = str(input("Qual fundo você deseja verificar o preço atual, e o DY?")).upper()
    print(f"Preço da cota: {lista_fundos[opcao].preco_atual}\nSegmento: {lista_fundos[opcao].segmento}\nDY 12 Meses: {lista_fundos[opcao].dy_12m}")
