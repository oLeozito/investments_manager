import openpyxl

fundos = {}
setores = ['PAPÉIS','INDEFINIDO','FUNDO DE DESENVOLVIMENTO','IMÓVEIS INDUSTRIAIS E LOGÍSTICOS','MISTO','FUNDO DE FUNDOS','LAJES CORPORATIVAS','VAREJO','IMÓVEIS COMERCIAIS - OUTROS','SHOPPINGS','AGÊNCIAS DE BANCOS']

class Fundo:
    def __init__ (self):
        self.sigla = ""
        self.setor = ""
        self.preço = 0.0
        self.liquidez = ""
        self.pvp = 0.0
        self.dyac = 0.0
        self.dymedia = 0.0
        self.patrliquid = 0.0

plano = openpyxl.load_workbook("Filtro Fii.xlsx")
pagina = plano['Fundos']

for rows in pagina.iter_rows(min_row = 3, max_row = 79,max_col=8):
    fundo = Fundo()
    flag = 1
    for cell in rows:
        if flag == 1:
            fundo.sigla = str(cell.value)
        elif flag == 2:
            fundo.setor = str(cell.value)
        elif flag == 3:
            fundo.preço = float(str(cell.value).replace(",","."))
        elif flag == 4:
            fundo.liquidez = float(str(cell.value).replace(",","."))
        elif flag == 5:
            fundo.pvp = float(str(cell.value).replace(",","."))
        elif flag == 6:
            fundo.dyac = round(100*float(str(cell.value).replace(",",".")),2)
        elif flag == 7:
            fundo.dymedia = round(100*float(str(cell.value).replace(",",".")),2)
        elif flag == 8:
            fundo.patrliquid = float(str(cell.value).replace(",","."))
        flag+=1
    fundos[fundo.sigla] = fundo

def opc1(sigla):
    if sigla in fundos.keys():
                print(f"\nInformações sobre o {fundos[sigla].sigla}")
                print("Setor - ", fundos[sigla].setor)
                print("Preço - R$",fundos[sigla].preço)
                print("Liquidez - R$",fundos[sigla].liquidez)
                print("P/VP - ",fundos[sigla].pvp)
                print("DY - ",fundos[sigla].dyac,"%")
                print("DY 12M - ", fundos[sigla].dymedia,"%")
                print("Patrimônio - R$",fundos[sigla].patrliquid)
def opc2(setor,fundos,setores):
    match setor:
        case 1:
            setor = input("Qual setor você deseja analisar?: ").upper()
            if setor in setores:
                for cada in fundos.values():
                    if cada.setor == setor:
                        print(f'{cada.sigla} - Preço: R${cada.preço} - DY: {cada.dyac} - P/VP: {cada.pvp}')
            else:
                print("O setor não foi encontrado.")
        case 2:
            for each in fundos.values():
                if each.setor in setores:
                    print(f'{each.sigla} - Preço: R${each.preço} - DY: {each.dyac} - P/VP: {each.pvp} - Setor: {each.setor}')
        case 3:
            for y in fundos.values():
                if y.setor == "PAPÉIS" or y.setor == "MISTO":
                    print(f'{y.sigla} - Preço: R${y.preço} - DY: {y.dyac} - P/VP: {y.pvp} - Setor: {y.setor}')
def opc3(fundos):
    fundos_ordenados = sorted(fundos.values(), key=lambda fundo: fundo.preço)
    for cada in fundos_ordenados:
        print(f'{cada.sigla} - Preço: R${cada.preço} - DY: {cada.dyac} - P/VP: {cada.pvp}')
def opc4(fundos):
    fundos_ordenados = sorted(fundos.values(), key=lambda fundo: fundo.dyac, reverse= True)
    for cada in fundos_ordenados:
        print(f'{cada.sigla} - Preço: R${cada.preço} - DY: {cada.dyac} - P/VP: {cada.pvp}')
def menu(fundos):
    vari = ""
    while vari != "0":
        vari = str(input("""\nMenu
Analisar por:
1 - Sigla.
2 - Setor.
3 - Preços.
4 - Dividend Yeld.
0 - Sair.
Digite sua opção: """))

        match vari:
            case '1':
                sigla = str(input("\nDigite a sigla do fundo: ")).upper()
                if sigla in fundos.keys():
                    opc1(sigla)
                else:
                    print("Não encontramos o Fundo Imobiliário!")
            case '2':
                aux = int(input("""
1 - Analisar setor específico
2 - Filtrar apenas FIIs de Tijolo
3 - Filtrar apenas FIIs de Papel
Digite: """))
                opc2(aux,fundos,setores)
            case '3':
                opc3(fundos)
            case '4':
                opc4(fundos)
            case _:
                print("Por favor digite uma opção válida\n")

menu(fundos)