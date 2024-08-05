import os.path
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

#[ ATRIBUTOS ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#
# If modifying these scopes, delete the file token.json.
#.readonly (para deixar somente leitura)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1MvBziZf88oCOWaO0U5LmSN7WsKqqd-Bzq2OjQgg6bVQ"
SAMPLE_RANGE_NAME = "RESERVA_TECNICA!A1:Q"

navegador = webdriver.Chrome()
campo_login_inicial = '//*[@id="usuario"]'
campo_senha_inicial = '//*[@id="senha"]'
botao_entrar_inicial = '//*[@id="conteudoFull"]/div[1]/div[1]/div[8]/input'

#[ MÉTODOS ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

def logarSISREG():
    # Verificar se a página carregou por completo, se não, refresh()
    if navegador.find_elements(by=By.XPATH, value='//*[@id="usuario"]'):
        navegador.find_element(By.XPATH, campo_login_inicial).send_keys("5408989MATHEWSF")
        navegador.find_element(By.XPATH, campo_senha_inicial).send_keys('96894177~Leao')
        navegador.find_element(By.XPATH, botao_entrar_inicial).click()

def tokenGoogleSheetsAPI():
    creds = None
    # Verifica se o arquivo token.json existe
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # Se não houver credenciais válidas, realiza o fluxo de login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # Atualiza o token de acesso usando o token de atualização
            creds.refresh(Request())
        else:
            # Executa o fluxo de autenticação do OAuth 2.0
            flow = InstalledAppFlow.from_client_secrets_file("credencialgooglesheets.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Salva as credenciais para a próxima execução
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def getPlanilhaGeral():
    try:
        service = build("sheets", "v4", credentials=tokenGoogleSheetsAPI())
        # Ler informações das células do Google Sheets
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
        return result['values']
    except HttpError as err:
        print(err)

def setCelulaPlanilha(aba, celula, valor):
    try:
        service = build("sheets", "v4", credentials=tokenGoogleSheetsAPI())
        # Inserir / editar uma célula no Google Sheets
        sheet = service.spreadsheets()
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=aba + celula, valueInputOption="USER_ENTERED", body={'values': valor}).execute()
    except HttpError as err:
        print(err)

#[ MAIN ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
navegador.get("https://sisregiii.saude.gov.br/cgi-bin/index?logout=1")
time.sleep(5)
logarSISREG()

navegador.get("https://sisregiii.saude.gov.br/cgi-bin/autorizador")
time.sleep(5)

def extrair_tabela(Categoria, risco):
    Select(navegador.find_element(By.NAME, "st_visualizado")).select_by_visible_text(Categoria)
    Select(navegador.find_element(By.NAME, "co_risco")).select_by_visible_text(risco)    
    Select(navegador.find_element(By.NAME, "qtd_itens_pag")).select_by_index(3)
    time.sleep(5)

    total_de_paginas = navegador.find_element(By.XPATH, '//table/tbody/tr/td[contains(text(),"Exibindo Página")]')   
    print(int(total_de_paginas.text.split()[-1]))
    time.sleep(5)

    #for e extrair página por página
    #ir inserir tudo na planilha do google
    
    tabela = navegador.find_element(By.XPATH, "/html/body/form/table")
    lista_de_pacientes = tabela.find_elements(By.TAG_NAME, "tr")

    # Itera sobre as linhas e extrai os dados das células
    for row in lista_de_pacientes:
        cells = row.find_elements(By.TAG_NAME, 'td')
        cell_data = [cell.text.strip() for cell in cells]
        print(cell_data)    


extrair_tabela("Visualizadas", "Prioridade Zero - Emergência")

#Selecionar visualizados
#pegar todos vermelhos e setar como "Visualizado" e Risco "Vermelho"
#Pegar o quantitativo de páginas
#pegar a tabela da primeira página até a última
#jogar tudo na planilha

#pegar todos amarelo e setar como "Visualizado" e Risco "Amarelo"
#Pegar o quantitativo de páginas
#pegar a tabela da primeira página até a última
#jogar tudo na planilha

#pegar todos verdes e setar como "Visualizado" e Risco "Verde"
#Pegar o quantitativo de páginas
#pegar a tabela da primeira página até a última
#jogar tudo na planilha

#pegar todos azuls e setar como "Visualizado" e Risco "Azul"
#Pegar o quantitativo de páginas
#pegar a tabela da primeira página até a última
#jogar tudo na planilha

#Selecionar não visualizados
#pegar todos vermelhos e setar como "Não Visualizado" e Risco "Vermelho"
#Pegar o quantitativo de páginas
#pegar a tabela da primeira página até a última
#jogar tudo na planilha

#pegar todos amarelo e setar como "Não Visualizado" e Risco "Amarelo"
#Pegar o quantitativo de páginas
#pegar a tabela da primeira página até a última
#jogar tudo na planilha

#pegar todos verdes e setar como "Não Visualizado" e Risco "Verde"
#Pegar o quantitativo de páginas
#pegar a tabela da primeira página até a última
#jogar tudo na planilha

#pegar todos azuls e setar como "Não Visualizado" e Risco "Azul"
#Pegar o quantitativo de páginas
#pegar a tabela da primeira página até a última
#jogar tudo na planilha
 


#entrar chamado por chamado e pegar PROCEDIMENTO SOLICITADO e inserir na planilha
#inserir que o chamado foi conferido
#ir para próxima linha até o final



'''for i, chamadoID in enumerate(getPlanilhaGeral()):
    if i > 0:
        linha = 0
        print("Item[0]: " + chamadoID[0])
        # Continue imprimindo os itens conforme necessário...

        if chamadoID[16] == "SIM":
            print("Linha já conferida: " + chamadoID[16], i)
            continue
        elif chamadoID[16] == "NÃO":
            print("Linha sendo conferida: " + chamadoID[16], i)
            print("Entrando no chamado do SISREG de ID: " + chamadoID[2], i)
            navegador.get("https://sisregiii.saude.gov.br/cgi-bin/gerenciador_solicitacao?etapa=VISUALIZAR_FICHA&co_solicitacao=" + chamadoID[2] + "&co_seq_solicitacao=" + chamadoID[2] + "&ordenacao=2&pagina=0")
            time.sleep(10)

            validandoCNS = False
            tentativas = 0

            while not validandoCNS:
                elemento_codigo_solicitacao = navegador.find_element(By.XPATH, "//b[text()='Código da Solicitação:']")
                codigo_solicitacao_sisreg = elemento_codigo_solicitacao.find_element(By.XPATH, '../../following-sibling::tr[1]/td[0]').text
                print(codigo_solicitacao_sisreg)
                if str(codigo_solicitacao_sisreg) == str(chamadoID[6]):
                    print("Código da solicitação correto: " + chamadoID[6], codigo_solicitacao_sisreg)
                    validandoCNS = True
                else:
                    print("Código da solicitação diferente: " + chamadoID[6], codigo_solicitacao_sisreg)
                    navegador.refresh()
                    tentativas += 1
                    time.sleep(5)
                    if tentativas == 3:
                        break

            if navegador.find_elements(By.XPATH, value='//*[@id="fichaAmbulatorial"]/table/tbody[6]/tr[2]/td[1]/font'):
                valor = navegador.find_element(By.XPATH, '//*[@id="fichaAmbulatorial"]/table/tbody[6]/tr[2]/td[1]/font').text

                if valor.find(chamadoID[2]) != -1:
                    print('Entrou no chamado do SISREG id:' + chamadoID[2])
                    time.sleep(10)

                    if navegador.find_elements(By.XPATH, '/html/body/div[14]'):
                        print('Página carregada por completo!')
                        procedimento_solicitado = navegador.find_element(By.XPATH, '//*[@id="item-main"]/div/div[3]/div').find_element(By.NAME, 'date').get_attribute('value')
                        print(procedimento_solicitado)

                        historico_observacoes = navegador.find_element(By.XPATH, '//*[@id="heading-main-item"]/button/span[1]/i').get_attribute('data-bs-original-title')
                        print(historico_observacoes)

                        linha = i + 1
                        print("Dados inseridos na Aba: RESERVA_TÉCNICA! - Linha: K" + str(linha))
                        setCelulaPlanilha('RESERVA_TÉCNICA!', 'K' + str(linha), [[procedimento_solicitado]])
                        setCelulaPlanilha('RESERVA_TÉCNICA!', 'L' + str(linha), [[historico_observacoes]])
                        setCelulaPlanilha('RESERVA_TÉCNICA!', 'Q' + str(linha), 'SIM')
                    else:
                        print("Página não carregou por completo, atualizando-a")

                else:
                    print('Não entrou no chamado do SISREG id:' + chamadoID[2])

        else:
            print("Linha deu erro: " + chamadoID[9], i)
            setCelulaPlanilha('RESERVA_TÉCNICA!', 'Q' + str(linha), [["ERRO"]])
            continue
'''

print('Programa Finalizado')
