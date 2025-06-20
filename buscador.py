import requests
from bs4 import BeautifulSoup
import json
import os
from datetime import datetime

class FundoImobiliario:
    def __init__(self, codigo, dy_12m, dy_percent, preco_atual, segmento, tipo, val_patr, vacancia, qtdcotis, qtdcotas, cnpj, preco_teto, pvp, liquidez):
        self.codigo = codigo
        self.dy_12m = dy_12m
        self.dy_percent = dy_percent
        self.preco_atual = preco_atual
        self.segmento = segmento
        self.tipo = tipo
        self.val_patr = val_patr
        self.vacancia = vacancia
        self.qtdcotis = qtdcotis
        self.qtdcotas = qtdcotas
        self.cnpj = cnpj
        self.pvp = pvp
        self.liquidez = liquidez
        self.preco_teto = preco_teto

    def to_dict(self):
        return self.__dict__

    @staticmethod
    def from_dict(dados):
        return FundoImobiliario(**dados)

class Data:
    def __init__(self, siglas=None, fundos=None, ultima_atualizacao=None):
        self.siglas = siglas if siglas else []
        self.fundos = fundos if fundos else {}  # dicionário com os dados dos fundos (em formato dict)
        self.ultima_atualizacao = ultima_atualizacao

    def salvar(self, arquivo="data.json"):
        dados = {
            "siglas": self.siglas,
            "fundos": self.fundos,
            "ultima_atualizacao": self.ultima_atualizacao
        }
        with open(arquivo, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=4)

    @staticmethod
    def carregar(arquivo="data.json"):
        if os.path.exists(arquivo):
            with open(arquivo, "r", encoding="utf-8") as f:
                dados = json.load(f)
                return Data(
                    siglas=dados.get("siglas", []),
                    fundos=dados.get("fundos", {}),
                    ultima_atualizacao=dados.get("ultima_atualizacao")
                )
        else:
            return Data()

def obter_fiis():
    letras = [chr(i) for i in range(ord('A'), ord('Z')+1)]
    siglas = []

    url_fiis = "https://www.fundsexplorer.com.br/funds"
    headers = {'User-Agent': 'Mozilla/5.0'}

    resposta = requests.get(url_fiis, headers=headers)
    if resposta.status_code != 200:
        print(f"Erro ao acessar o site.")

    html_fiis = BeautifulSoup(resposta.text, 'html.parser')
    listar = html_fiis.find('div', class_='tickerFilter__results')

    for letra in letras:
        filtrar = listar.find('section', id=f"letter-id-{letra}")
        buscar = filtrar.find_all('div', attrs={'data-element': 'ticker-box-title'})
        for item in buscar:
            siglas.append(item.text.strip())

    print(f"\nBusca de {len(siglas)} siglas completa.\n")
    return siglas

def obter_infos(fii):
    url = f'https://investidor10.com.br/fiis/{fii.lower()}/'
    headers = {'User-Agent': 'Mozilla/5.0'}

    resposta = requests.get(url, headers=headers)
    if resposta.status_code != 200:
        print(f"Erro ao acessar o site para {fii}")
        return None

    soup = BeautifulSoup(resposta.text, 'html.parser')

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

    preco_cota = None
    for span in soup.find_all('span', class_='value'):
        texto = span.text.strip()
        if 'R$' in texto:
            valor = texto.replace('R$', '').replace(',', '.').strip()
            try:
                preco_cota = float(valor)
                break
            except ValueError:
                continue

    def get_card_value(classe):
        card = soup.find('div', class_=f'_card {classe}')
        if card:
            corpo = card.find('div', class_='_card-body')
            if corpo:
                span = corpo.find('span')
                if span:
                    return span.text.strip()
        return None

    dy12 = get_card_value("dy")
    pvp = get_card_value("vp")
    liquidez = get_card_value("val")

    valores = {}
    tabela = soup.find('div', id="table-indicators")
    if tabela:
        celulas = tabela.find_all('div', class_="cell")
        for cell in celulas:
            nome = cell.find('span', class_="d-flex justify-content-between align-items-center name")
            value = cell.find('div', class_="value")
            if nome and value:
                valores[nome.text.strip()] = value.text.strip()

    cnpj       = valores.get("CNPJ")
    segmento   = valores.get("SEGMENTO")
    tipo       = valores.get("TIPO DE FUNDO")
    vacancia   = valores.get("VACÂNCIA")
    qtd_cotis  = valores.get("NUMERO DE COTISTAS")
    qtd_cotas  = valores.get("COTAS EMITIDAS")
    val_patri  = valores.get("VALOR PATRIMONIAL")

    preco_teto = round(val_dy_12m / 0.12, 2) if val_dy_12m else None

    return FundoImobiliario(
        fii, val_dy_12m, dy12, preco_cota, segmento, tipo,
        val_patri, vacancia, qtd_cotis, qtd_cotas, cnpj,
        preco_teto, pvp, liquidez
    )

# --- Programa Principal ---
data = Data.carregar()
siglas = data.siglas
lista_fundos = {
    codigo: FundoImobiliario.from_dict(info)
    for codigo, info in data.fundos.items()
}

if data.ultima_atualizacao:
    print(f"Última atualização: {data.ultima_atualizacao}")
else:
    print("Nenhuma atualização anterior encontrada.")

sair = 0
while sair != 1:
    opcao = str(input("\n##- MENU -##\n1 - Buscar siglas\n2 - Atualizar Informações\n3 - Acessar informações de um fundo\n0 - Sair\n\nO que você deseja fazer?: "))
    if opcao == "0":
        data.siglas = siglas
        data.fundos = {k: v.to_dict() for k, v in lista_fundos.items()}
        data.ultima_atualizacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        data.salvar()
        print(f"\nDados salvos em data.json.")
        sair = 1

    elif opcao == "1":
        siglas = obter_fiis()

    elif opcao == "2":
        for sigla in siglas:
            print(f"\nAtualizando {sigla}...")
            fundo = obter_infos(sigla)
            if fundo:
                lista_fundos[sigla] = fundo
        print("\nAtualização Completa.")

    elif opcao == "3":
        busca = str(input("Qual fundo você deseja verificar?: ")).upper()
        if busca in lista_fundos:
            fundo = lista_fundos[busca]
            print(f"\nPreço da cota: R${fundo.preco_atual}")
            print(f"Segmento: {fundo.segmento}")
            print(f"DY 12 Meses: R${fundo.dy_12m}")
            print(f"Preço Teto: R${fundo.preco_teto}")
        else:
            print("Fundo não encontrado.")

    else:
        print("\nDigite uma opção válida!")
