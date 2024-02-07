import os.path
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By

#[ ATRIBUTOS ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
#
# # If modifying these scopes, delete the file token.json.
#.readonly (para deixar somente leitura)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1MvBziZf88oCOWaO0U5LmSN7WsKqqd-Bzq2OjQgg6bVQ"
SAMPLE_RANGE_NAME = "RESERVA_TÉCNICA!A1:Q"

navegador = webdriver.Chrome()
campo_login_inicial = '//*[@id="usuario"]'
campo_senha_inicial = '//*[@id="senha"]'
botao_entrar_inicial = '//*[@id="conteudoFull"]/div[1]/div[1]/div[8]/input'

#[ MÉTODOS ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
def logarSISREG():
  #verificar se a página carregou por completo, se não, refresh()
  if navegador.find_elements(by=By.XPATH, value='//*[@id="usuario"]'):
    navegador.find_element(By.XPATH, campo_login_inicial).send_keys("5408989MATHEWSF")
    navegador.find_element(By.XPATH, campo_senha_inicial).send_keys('96894177~Leao')
    #time.sleep(50)
    #Retirar time.sleep e inserir biblioteca que passe do captcha
    navegador.find_element(By.XPATH, botao_entrar_inicial).click()

def tokenGoogleSheetsAPI():
  creds = None
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credencialgooglesheets.json", SCOPES
      )
      creds = flow.run_local_server(port=0)

    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())   
  return creds   


def getPlanilhaGeral():      
  try:    
    service = build("sheets", "v4", credentials=tokenGoogleSheetsAPI())

    #Ler informações [Células] o Google Sheets
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                range=SAMPLE_RANGE_NAME).execute()     
    return result['values']

  except HttpError as err:
    return print(err)
  
  
  
def setCelulaPlanilha(aba, celula, valor): 
  try:
    service = build("sheets", "v4", credentials=tokenGoogleSheetsAPI())
    #Inserir / editar uma informação [Célula] no Google Sheets
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                range=aba+celula, valueInputOption="USER_ENTERED", 
                                body={'values': valor}).execute()                

  except HttpError as err:
    print(err)  


#[ MAIN ] <-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
navegador.get("https://sisregiii.saude.gov.br/cgi-bin/index?logout=1")
time.sleep(10)
logarSISREG()

for i, chamadoID in enumerate(getPlanilhaGeral()):
  if(i > 0):
    linha = 0
    print("Item[0]: "+chamadoID[0])
    print("Item[1]: "+chamadoID[1])
    print("Item[2]: "+chamadoID[2])
    print("Item[3]: "+chamadoID[3])
    print("Item[4]: "+chamadoID[4])
    print("Item[5]: "+chamadoID[5])
    print("Item[6]: "+chamadoID[6])
    print("Item[7]: "+chamadoID[7])
    print("Item[8]: "+chamadoID[8])
    print("Item[9]: "+chamadoID[9])
    print("Item[10]: "+chamadoID[10])
    print("Item[11]: "+chamadoID[11])
    print("Item[12]: "+chamadoID[12])
    print("Item[13]: "+chamadoID[13])
    print("Item[14]: "+chamadoID[14])
    print("Item[15]: "+chamadoID[15])
    print("Item[16]: "+chamadoID[16])
    if(chamadoID[16] == "SIM"):
      print("Linha já conferida:"+ chamadoID[16], i)
      continue

    elif(chamadoID[16] == "NÃO"):
      print("Linha sendo conferida:"+ chamadoID[16], i)
      print("Entrando no chamado do SISREG de ID: "+chamadoID[2], i)
      navegador.get("https://sisregiii.saude.gov.br/cgi-bin/gerenciador_solicitacao?etapa=VISUALIZAR_FICHA&co_solicitacao="+chamadoID[2]+"&co_seq_solicitacao="+chamadoID[2]+"&ordenacao=2&pagina=0")

      time.sleep(10)

      validandoCNS = False
      elemento_codigo_solicitacao = 00000000000000
      codigo_solicitacao_sisreg = 00000000000000

      tentativas = 0
      while(validandoCNS == False):
        elemento_codigo_solicitacao = navegador.find_element(By.XPATH, "//b[text()='Código da Solicitação:']")
        codigo_solicitacao_sisreg = elemento_codigo_solicitacao.find_element(By.XPATH, '../../following-sibling::tr[1]/td[0]').text
        print(codigo_solicitacao_sisreg)
        if(str(codigo_solicitacao_sisreg) == str(chamadoID[6])):
          print("Código da solicitação correto: "+chamadoID[6], codigo_solicitacao_sisreg)
          validandoCNS = True
        elif(str(codigo_solicitacao_sisreg) != str(chamadoID[6])):
          print("Código da solicitação diferente: "+chamadoID[6], codigo_solicitacao_sisreg)          
          validandoCNS = False        
          navegador.refresh()
          tentativas+=1
          time.sleep(5)

          if(tentativas == 3):
            tentativas = 0 
            break

          

      if(navegador.find_elements(By.XPATH, value='//*[@id="fichaAmbulatorial"]/table/tbody[6]/tr[2]/td[1]/font')):
        valor = navegador.find_element(By.XPATH, '//*[@id="fichaAmbulatorial"]/table/tbody[6]/tr[2]/td[1]/font').text

        if(valor.find(chamadoID[2]) != -1):
          print('Entrou no chamado do SISREG id:'+chamadoID[2])
          time.sleep(10)

          if(navegador.find_elements(By.XPATH, '/html/body/div[14]')):
            print('Página carregada por completo!')
            procedimento_solicitado = navegador.find_element(By.XPATH, '//*[@id="item-main"]/div/div[3]/div').find_element(By.NAME, 'date').get_attribute('value')
            print(procedimento_solicitado)

            historico_observacoes = navegador.find_element(By.XPATH, '//*[@id="heading-main-item"]/button/span[1]/i').get_attribute('data-bs-original-title')
            print(historico_observacoes)  

            linha = i+1
            print("Dados inseridos na Aba: RESERVA_TÉCNICA! - Linha: K"+str(linha))
            setCelulaPlanilha('RESERVA_TÉCNICA!', 'K'+str(linha), [[procedimento_solicitado]])
            setCelulaPlanilha('RESERVA_TÉCNICA!', 'L'+str(linha), [[historico_observacoes]])
            setCelulaPlanilha('RESERVA_TÉCNICA!', 'Q'+str(linha), 'SIM')
          else:
            print("Página não carregou por completo, atualizando-a")      
            #Refresh e verificar
            
        else:
          print('Não entrou no chamado do SISREG id:'+chamadoID[2])
          #Atualizar para o mesmo link e voltar ao início do código
        print(valor)
    else:
      print("Linha deu erro:"+ chamadoID[9], i)
      setCelulaPlanilha('RESERVA_TÉCNICA!', 'Q'+str(linha), [["ERRO"]])
      continue

print('Programa Finalizado')