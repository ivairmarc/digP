from io import BytesIO
import os
import csv
import time
import numpy as np
import yaml
import mss
import datetime
import pyautogui
import pandas as pd
from random import random
from os import error, listdir
import PySimpleGUI as sg
from pathlib import  Path
from selenium import webdriver
from src.logger import logger, loggerMapClicked
from cv2 import cv2
from bs4 import BeautifulSoup
from random import randint, vonmisesvariate
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from tkinter import EXCEPTION, constants, filedialog
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.wait import WebDriverWait
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.expected_conditions import alert_is_present
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementNotSelectableException, NoAlertPresentException

#https://sites.google.com/chromium.org/driver/downloads
os.chdir('/home/anon/Documentos/PaulistaDigit/')

DirTerm = '/home/anon/Documentos/PaulistaDigit/termos/'

RG = '/home/anon/Documentos/PaulistaDigit/rg.jpg'
trmo = '/home/anon/Documentos/PaulistaDigit/img/Untitled1.pdf'

cfyml = 'config.yaml'
img_path = '/home/anon/Documentos/PaulistaDigit/img/'
resultado_csv = 'ResultadoConsultaCSV.csv'
qtd_uf = 'logs/uf.csv'

qtd_cons = '/home/anon/Documentos/PaulistaDigit/vl.csv'

cred = '/home/anon/Documentos/PaulistaDigit/sptx.json'

scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name(cred, scope)

gc = gspread.authorize(credentials)
wks = gc.open_by_key('1kINeC8H-0A7_VrsQtpthWS1gM6Kylb38nzDwhJf4nHI')

# Seleciona a primeira página da planilha
nb_abas = 0
worksheet = wks.get_worksheet(nb_abas)

cell = worksheet.findall('user')
credent = 'claudio'
cell = worksheet.find(credent)
colum = 2
val = worksheet.cell(cell.row, colum).value



stream = open(cfyml, 'r')
c = yaml.safe_load(stream)
ct = c['threshold']
pause = c['time_intervals']['interval_between_moviments']
pyautogui.PAUSE = pause

pasta = listdir(DirTerm)
for file in pasta:
    print(f'Termo: {file}')
    PthTerm = DirTerm + file
    try:
        if os.path.isfile(PthTerm):
            os.remove(PthTerm)
            print('Termo deletado')
    except OSError as er:
        print(er)

# LOGIN BANCOSEGURO
with open('user.csv', 'r') as csv_file:
    csv_dict = csv.DictReader(csv_file, delimiter=';')
    for lin in csv_dict:
        usuario = lin['nome']
        senha = lin['senha']
        API_KEY = lin['api']
        
# CELULAR
# SELECIONAR ARQUIVO
if os.path.isfile(qtd_cons):
    with open(qtd_cons, "r") as fa:
        lines = fa.readlines()
else:
    f = open(qtd_cons, 'w')
    f.write('0')
    f.close()
    with open(qtd_cons, "r") as fa:
        lines = fa.readlines()

        
# CELULAR
# SELECIONAR ARQUIVO
print(f'Total de Consultas realizadas: {lines[0]}')


Endereco = sg.popup_get_file('Selecione o arquivo csv')

if Endereco == None:
    print('Finalizando...')
    exit(69)
    
Caminho_Arquivo = Path(Endereco)

date = str(datetime.datetime.fromtimestamp(int(time.time())).strftime('%Y-%m-%d %H:%M:%S'))
print(date)

"""if os.path.isdir('C:\\ProgramData\\Wintrics'):
    print("...")
else:
    exit(69)"""

driver = webdriver.Firefox(executable_path='./geckodriver')

WbW = WebDriverWait(driver, 30)



# ABRIR ABA Banco


def Login():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'EUsuario_CAMPO')))
    driver.find_element(By.ID,'EUsuario_CAMPO').send_keys(usuario)
    time.sleep(1)
    driver.find_element(By.ID,'ESenha_CAMPO').click()
    time.sleep(1)
    driver.find_element(By.ID,'ESenha_CAMPO').send_keys(senha)
    time.sleep(1)

    # ENTRAR
    driver.find_element(By.ID,'lnkEntrar').click()
    # clicar no menu
    time.sleep(2)

def aba_liberar():
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="navbar-collapse-funcao"]/ul/li[5]')))
        driver.find_element(By.XPATH, '//*[@id="navbar-collapse-funcao"]/ul/li[5]').click()
        driver.find_element(By.XPATH, '//*[@id="WFP2010_MTDDSBENEF"]').click()
    except:
        driver.find_element(By.XPATH, '//*[@id="navbar-collapse-funcao"]/ul/li[5]').click()
        driver.find_element(By.XPATH, '//*[@id="WFP2010_MTDDSBENEF"]').click()

def nova_tab():
    winds = driver.window_handles
    driver.switch_to.window(winds[0])
    time.sleep(1)
    driver.execute_script("window.open('https://consignado.bancopaulista.com.br/WebAutorizador/Login/', '_blank')")
    driver.implicitly_wait(30)
    winds = driver.window_handles
    driver.switch_to.window(winds[1])

def print_time(delay, counter):

    while counter:
        time.sleep(delay)
        nova_tab()
        aba_liberar()
        counter -= 1

def Liberar_termo(CPF, Nome):

    id_CPF="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_AbaTermoAutorizacao_txtCpfCli_CAMPO"
    id_Nome="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_AbaTermoAutorizacao_txtNomeCli_CAMPO"
    id_DDD="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_AbaTermoAutorizacao_txtDDD_CAMPO"
    id_LOCAL="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_AbaTermoAutorizacao_txtLocalAssTermo_CAMPO"
    alet = randint(1,100)
    
    email = Nome.lower().replace(" ","_")+str(alet)+"@gmail.com.br"
    Local = ['Belém', 'Boa Vista', 'Macapá', 'Manaus', 'Palmas', 'Porto Velho', \
         'Rio Branco', 'Aracaju', 'Fortaleza', 'João Pessoa', 'Maceió', 'Natal',\
         'Recife', 'Salvador', 'São Luís', 'Teresina', 'Brasília','Campo Grande',\
         'Cuiabá', 'Goiânia', 'Belo Horizonte', 'Rio de Janeiro', 'São Paulo',\
         'Vitória', 'Curitiba', 'Florianópolis', 'Porto Alegre']

    WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.ID, id_CPF)))
    try:
        driver.find_element(By.ID,id_CPF).clear()
        
    except:
        None
    time.sleep(1)
    driver.find_element(By.XPATH, '//*[@id="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_AbaTermoAutorizacao_rdlTipoTermo_1"]').click()
    #CPF do Cliente:
    time.sleep(1)
    driver.find_element(By.ID,id_CPF).send_keys(CPF)
    time.sleep(1)
    driver.find_element(By.ID,id_Nome).click()
    alr = alerta_P()

    if 'CPF' in alr:
        driver.find_element(By.ID,id_CPF).send_keys(CPF)
        time.sleep(1)
        driver.find_element(By.ID,id_Nome).click()

    driver.implicitly_wait(10)
    time.sleep(3)
    try:
        driver.find_element(By.ID,id_Nome).clear()
    except:
        pass
    driver.find_element(By.XPATH, '//*[@id="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_AbaTermoAutorizacao_rdlTipoTermo_1"]').click()
    #Nome do Cliente:
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, id_Nome)))
    driver.find_element(By.ID,id_Nome).send_keys(Nome)
    time.sleep(1)
    driver.find_element(By.ID,id_DDD).click()
    time.sleep(1)

    #LOCAL
    WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.ID, id_LOCAL)))
    driver.find_element(By.ID,id_LOCAL).clear()
    driver.find_element(By.ID,id_LOCAL).send_keys(Local[randint(0,26)])
    time.sleep(1)
    
    

    time.sleep(1)
    
    # ENVIAR ASSINATURA ELETRONICA
    driver.find_element(By.ID,"btnImprimirTermo_txt").click()
    
    time.sleep(10)
    alerta_P()

def alerta_P():
    try:
        alerta = WebDriverWait(driver, 5).until(alert_is_present())
        time.sleep(1)
        texto_alerta = driver.switch_to.alert.text
        time.sleep(1)
        driver.switch_to.alert.accept()
        return texto_alerta
    except:
        texto_alerta = 'Termo de Autorização enviado para Assinatura Digital com sucesso!'
        return texto_alerta

def addRandomness(n, randomn_factor_size=None):
    """Returns n with randomness
    Parameters:
        n (int): A decimal integer
        randomn_factor_size (int): The maximum value+- of randomness that will be
            added to n

    Returns:
        int: n with randomness
    """

    if randomn_factor_size is None:
        randomness_percentage = 0.1
        randomn_factor_size = randomness_percentage * n

    random_factor = 2 * random() * randomn_factor_size
    if random_factor > 5:
        random_factor = 5
    without_average_random_factor = n - randomn_factor_size
    randomized_n = int(without_average_random_factor + random_factor)
    # logger('{} with randomness -> {}'.format(int(n), randomized_n))
    return int(randomized_n)

def moveToWithRandomness(x,y,t):
    pyautogui.moveTo(addRandomness(x,10),addRandomness(y,10),t+random()/2)

def remove_suffix(input_string, suffix):
    """Returns the input_string without the suffix"""

    if suffix and input_string.endswith(suffix):
        return input_string[:-len(suffix)]
    return input_string

def load_images(dir_path=img_path):
    """ Programatically loads all images of dir_path as a key:value where the
        key is the file name without the .png suffix

    Returns:
        dict: dictionary containing the loaded images as key:value pairs.
    """

    file_names = listdir(dir_path)
    targets = {}
    for file in file_names:
        path = dir_path + file
        # print(f'Aqui a Img: {file}')
        targets[remove_suffix(file, '.png')] = cv2.imread(path)

    return targets

def clickBtn(img, timeout=3, threshold = ct['default']):
    """Search for img in the scree, if found moves the cursor over it and clicks.
    Parameters:
        img: The image that will be used as an template to find where to click.
        timeout (int): Time in seconds that it will keep looking for the img before returning with fail
        threshold(float): How confident the bot needs to be to click the buttons (values from 0 to 1)
    """

    logger(None, progress_indicator=True)
    start = time.time()
    has_timed_out = False
    while(not has_timed_out):
        matches = positions(img, threshold=threshold)

        if(len(matches)==0):
            has_timed_out = time.time()-start > timeout
            continue

        x,y,w,h = matches[0]
        pos_click_x = x+w/2
        pos_click_y = y+h/2
        moveToWithRandomness(pos_click_x,pos_click_y,1)
        pyautogui.click()
        return True

    return False    
  
def printSreen():
    with mss.mss() as sct:
        monitor = sct.monitors[0]
        sct_img = np.array(sct.grab(monitor))
        # The screen part to capture
        # monitor = {"top": 160, "left": 160, "width": 1000, "height": 135}

        # Grab the data
        return sct_img[:,:,:3]

def positions(target, threshold=ct['default'],img = None):
    if img is None:
        img = printSreen()
    result = cv2.matchTemplate(img,target,cv2.TM_CCOEFF_NORMED)
    w = target.shape[1]
    h = target.shape[0]

    yloc, xloc = np.where(result >= threshold)


    rectangles = []
    for (x, y) in zip(xloc, yloc):
        rectangles.append([int(x), int(y), int(w), int(h)])
        rectangles.append([int(x), int(y), int(w), int(h)])

    rectangles, weights = cv2.groupRectangles(rectangles, 1, 0.2)
    return rectangles 


def Solver():
    page_url = driver.current_url
    print(page_url)
    import requests
    
    from twocaptcha import TwoCaptcha
    
    data_setekey= '6LeakfkdAAAAACcvNNC4gaiMY1rKOBn7M-8HOq2U'

    

    solver = TwoCaptcha(API_KEY)

    
    try:
       
        try:
            result = solver.recaptcha(
                sitekey=data_setekey,
                url=page_url,
                invisible=1)

        except Exception as e:
            print(f'\n Erro do Solver {e}')

        else:
            print('solved: ' + str(result))

            token = result['code']
            
            writ_tokon_js = f'document.getElementById("g-recaptcha-response").innerHTML="{token}";'
            #submit_js = 'document.getElementById("recaptcha-anchor").submit();'
            driver.execute_script(writ_tokon_js)
            time.sleep(3)
    except Exception as e:
        print(f'Erro do Script {e}')
    #driver.execute_script(submit_js)
    #time.sleep(3)

def ir_banco():
    
    # IR ABA BANCO
    """winds = driver.window_handles
    driver.switch_to.window(winds[1])"""

    time.sleep(1)

    # IR ABA SOLICITAR ASSINATURA
    
    def Aba_Solicitar():
        
        
        id_assinatura ="__tab_ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio"
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, id_assinatura)))

        driver.find_element(By.ID,id_assinatura).click()

        time.sleep(3)
        # Verifica se o input esta vaziu cpf
        cpf = 'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_txtCPFCliente_CAMPO'

        if driver.find_element(By.ID,cpf).get_attribute('value') == '':
            driver.find_element(By.ID,cpf).send_keys(CPF)
            
        
        id_chave="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_cboChaveTermo_CAMPO"
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, id_chave)))

        time.sleep(0.5)
        
        for h in range(4):
            
            # ESCOLHER CHAVE

            driver.find_element(By.ID,id_chave).click()
            time.sleep(1)

            select = Select(driver.find_element(By.ID,id_chave))
            time.sleep(1)
            
            try:
                select.select_by_index(1)
                time.sleep(0.5)
            except:
                continue
                
            if driver.find_element(By.ID,id_chave).get_attribute('value') != 'Selecione o Termo desejado':
                break


        #ESCOLHER RG
        try:
            id_tipo_doc = 'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_cboDocIdentifCli_CAMPO'
            element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, id_tipo_doc)))
        except:
            None

        drp = Select(element)
        time.sleep(1)
        drp.select_by_value("6")

        time.sleep(2)

        


    
    #CLICK SOLICITAR AUTORIZACAO
    def continua():
       
        try:
            time.sleep(8)
            WebDriverWait(driver, 70).until(EC.element_to_be_clickable((By.ID,'btnRealizarUpload_txt')))
            try:
                input_termo1 = 'FileUpGrd1'
                try:
                    driver.find_element(By.ID,input_termo1).send_keys(RG)
                except Exception as a:
                    print(a)
            except Exception as e:
                print(e)
                
            # esperar carregar o termo de consulta
            id_trC = 'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdDocIdentif_ctl02_lblArquivoGrd1'
            for i in range(25):
                tr = driver.find_element(By.ID,id_trC).text

                if tr != '':
                    break
                time.sleep(1)
            print(f'Termo de Consulta: {tr}')
            WebDriverWait(driver, 70).until(EC.element_to_be_clickable((By.ID,'btnSolicitarAutorizacao_txt')))
            id_bt = 'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_btnSolicitarAutorizacao_dvDBtn'
            try:

                driver.find_element(By.ID,id_bt).click()
                
            except:
                driver.find_element(By.ID,'btnSolicitarAutorizacao_txt').click()
                sg.popup_yes_no('não deu')
            try:
                
                try:
                    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID,"ctl00_Cph_PopReCaptcha_frameAjuda")))
                    """driver.switch_to.frame(driver.find_element(By.XPATH,'//*[@id="ctl00_cph_JN_ctl00_Recaptcha"]/div/div/iframe'))
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="recaptcha-anchor"]'))).click()"""

                    Solver()
                    driver.find_element(By.ID,'btnConfirmar_txt').click()
                    driver.switch_to.default_content()
                    #pega o retorno da posicao atual de x e y do mouse e passa o valor da tupla para as duas variaveis
                    
                except Exception as e:
                    print(f'Não cliquei no captcha: {e}')

                #clickBtn(images['captcha'], timeout=0.01)
            
            except Exception as e:
                print(f'\n Falha {e}')
            #sg.popup_yes_no('Olha eu aqui kkk')
            time.sleep(2)
            driver.find_element(By.ID,'btnSolicitarAutorizacao_txt').click()

            
    
            
            """id_Erro = '//*[@id="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdAutConsDadosBenif"]/tbody/tr[2]/td[2]'
                
            for k in range(30):

                time.sleep(1)

                try:
                    check = driver.find_elements(By.XPATH,id_Erro)
                except:
                    None

                if len(check) !=0:
                    
                    time.sleep(1)
                    break
        
            time.sleep(1)
            Error = driver.find_element(By.XPATH,id_Erro).text
        
            return Error"""
        except:
            try:
                time.sleep(8)
                
                WebDriverWait(driver, 70).until(EC.element_to_be_clickable((By.ID,'btnRealizarUpload_txt')))
                clickBtn(images['realizarupload'], timeout=1)
                time.sleep(9)
                WebDriverWait(driver, 70).until(EC.element_to_be_clickable((By.ID,'btnSolicitarAutorizacao_txt')))
                driver.find_element(By.ID,'btnSolicitarAutorizacao_txt').click()
            except:
                reiniciar_aba_ass()
                consulta.loc[ind_base, 'STATUS'] = "ERRO DE CONSULTA"
                consulta.to_csv(Caminho_Arquivo, sep=';', index=False)
   
    
    #CLICK ESCOLHER ARQUIVO RG
    def escolher_rg():
        
        id_escolher_arquivo ="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdArquivosUpload_ctl02_btnAnexarArquivoUpload"
       
        input_rg = 'FileMultArqUpload'
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID,id_escolher_arquivo)))
        try:
            driver.find_element(By.ID,input_rg).send_keys(RG)
        except Exception as a:
            print(f'{RG} \n')
            print(a)
        # GET RG
        try:            
            id_rg = 'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdArquivosUpload_ctl02_lblArquivoUpload'
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID,id_rg)))
            for i in range(40):
                rg = driver.find_element(By.ID,id_rg).text
                if rg != '':
                    break
                time.sleep(1)
            print(f'RG aqui: {rg}')
        except:
            pass

        

    # Selecionando o termo
    def eslcTerm():
        
        id_escolher_arquivo ='ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdTermoAutBenef_ctl02_AnexarGrd2'
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID,id_escolher_arquivo)))
        input_termo = 'FileUpGrd2'
        Pt =''
        try:
            pasta = listdir(DirTerm)
            for file in pasta:
                # print(f'Termo: {file}')
                Pt = DirTerm + file
        except:
            pass

        if Pt != '':
            PthTerm = Pt
        else:
            PthTerm = RG
        
        try:
            # GET Termo
            
            driver.find_element(By.ID,input_termo).send_keys(trmo)
            id_termo = 'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdTermoAutBenef_ctl02_lblArquivoGrd2'
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID,id_termo)))
            for i in range(40):
                termo = driver.find_element(By.ID,id_termo).text
                #print(f'Termo : {termo}')
                
                if termo != '':
                    break
                time.sleep(1)
            print(f'Termo aqui anexado: ')

        except Exception as e:
            print(f'ERRO 1{e}')

    # Inicio
    try:
        Aba_Solicitar()
    except:
        Aba_Solicitar()
        #sg.popup('Virifique se a pagina carregou e clique em YES')
    try:
        escolher_rg()
        eslcTerm()
        continua()
    except Exception as e:
        print(f'ERRO 2{e}')
        
           

    pasta = listdir(DirTerm)
    for file in pasta:
        print(f'\nTermo: {file}')
        PthTerm = DirTerm + file
        try:
            if os.path.isfile(PthTerm):
                os.remove(PthTerm)
                print('Termo deletado')
        except OSError as er:
            print(er)




def reiniciar_aba_ass():
    driver.get("https://consignado.bancopaulista.com.br/WebAutorizador")
    time.sleep(2)
    driver.implicitly_wait(30)
    WebDriverWait(driver, 80).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="navbar-collapse-funcao"]/ul/li[4]/a')))
    inss_button = driver.find_element(By.XPATH,'//*[@id="navbar-collapse-funcao"]/ul/li[5]/a')
    achains = ActionChains(driver)
    achains.move_to_element(inss_button).perform()
    time.sleep(1)
    driver.find_element(By.ID,'WFP2010_MTDDSBENEF').click()
    time.sleep(1)
    


def salvar_beneficio(consulta, ind_base):
    
    # PEGAR BENEFICIO
    # PRIMEIRO BENEFICIO

    try:
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli"]/tbody/tr[2]/td[1]/span')))
    except:
        None
    try:
        Beneficio = driver.find_element(By.XPATH,'//*[@id="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli"]/tbody/tr[2]/td[1]/span').text
        #consulta.loc[ind_base, 'MATRICULA'] = Beneficio
        print(f'Beneficío 1: {Beneficio}')

        Elegivel = driver.find_element(By.ID,'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli_ctl02_lblElegEmpre').text
        # consulta.loc[ind_base, 'ELEGÍVEL'] = Elegivel

        Bloqueado = driver.find_element(By.ID,'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli_ctl02_lblBloqEmpr').text
        #consulta.loc[ind_base, 'BLOQUEADO'] = Bloqueado

        Despacho = driver.find_element(By.ID,'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli_ctl02_lblDtDespBenef').text
        # consulta.loc[ind_base, 'DESPACHO'] = Despacho
    

        #consulta.loc[ind_base, 'ASSINATURA'] = "ASSINADO"

        if (Elegivel == 'Não') or (Bloqueado == 'Sim'):

            consulta.loc[ind_base, 'STATUS'] = "BLOQUEADO EMPRESTIMO"


        consulta.to_csv(Caminho_Arquivo, sep=';', index=False)

        qtd_benf = driver.find_elements(By.XPATH,'//*[@id="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli"]/tbody/tr')

        for j in range(2,len(qtd_benf)+1,1):

            if j > 2:

                ind_novo = len(consulta.index)

                Beneficio = driver.find_element(By.XPATH,f'//*[@id="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli"]/tbody/tr[{j}]/td[1]/span').text
                # consulta.loc[ind_novo, 'MATRICULA'] = Beneficio
                print(f'Beneficío {j-1}: {Beneficio}')

                Elegivel = driver.find_element(By.ID,f'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli_ctl0{j}_lblElegEmpre').text
                # consulta.loc[ind_novo, 'ELEGÍVEL'] = Elegivel

                Bloqueado = driver.find_element(By.ID,f'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli_ctl0{j}_lblBloqEmpr').text
                # consulta.loc[ind_novo, 'BLOQUEADO'] = Bloqueado

                Despacho = driver.find_element(By.ID,f'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli_ctl0{j}_lblDtDespBenef').text
                # consulta.loc[ind_novo, 'DESPACHO'] = Despacho

                #consulta.loc[ind_novo, 'ASSINATURA'] = "ASSINADO"

                if (Elegivel == 'Não') or (Bloqueado == 'Sim'):

                    consulta.loc[ind_novo, 'STATUS'] = "BLOQUEADO EMPRESTIMO"

                matricula = driver.find_element(By.XPATH,f'//*[@id="ctl00_Cph_jp1_pnlDadosBeneficiario_Container_ConsultaDadosBeneficio_grdListBenefCli"]/tbody/tr[{j}]/td[1]/span').text

                consulta.to_csv(Caminho_Arquivo, sep=';', index=False)
    except:
        consulta.loc[ind_base, 'STATUS'] = "BLOQUEADO EMPRESTIMO"
        consulta.to_csv(Caminho_Arquivo, sep=';', index=False )
        


    fim = time.time()
    print('')
    print(f'Tempo total de: {fim-inicio:.1f} segundos')


def Liberar_cli(ind_base, linha, aba_banco):      
    print('Liberando')
    print('#__________________')
        
    CPF = linha["CPF"]
    Nome = linha['Nome']
    
    
    driver.switch_to.window(aba_banco)
        
    # IR ABA TERMO DE AUTORIZACAO 'ctl00_Cph_jp1_pnlDadosBeneficiario_Container_AbaTermoAutorizacao_btnAssinaturaDigital'
    id_termo_autorizacao = "__tab_ctl00_Cph_jp1_pnlDadosBeneficiario_Container_AbaTermoAutorizacao"
    driver.find_element(By.ID,id_termo_autorizacao).click()
        
    time.sleep(1)
        
    texto_alerta= Liberar_termo(CPF, Nome)

    time.sleep(1)
    pst = 'sim'
        
    if texto_alerta == 'Já existe um Termo de Autorização de Consulta de Dados do Beneficiário emitido e válido para este cliente. Deseja realizar uma nova emissão?':

        print(texto_alerta)
            
        texto_alerta= Liberar_termo(CPF, Nome)
        
        if texto_alerta != 'Termo de Autorização enviado para Assinatura Digital com sucesso!':
            # ERRO AO LIBERAR
            pst = 'nao'
            reiniciar_aba_ass()
            consulta.loc[ind_base, 'STATUS'] = "ERRO DE CONSULTA"
            consulta.to_csv(Caminho_Arquivo, sep=';', index=False)

    if pst == 'sim':     
        try:
            Error = ir_banco()
        except:
            reiniciar_aba_ass()
            Error = ir_banco()
        
        if Error == 'Erro':
                
            
            consulta.loc[ind_base, 'STATUS'] = "BLOQUEADO EMPRESTIMO"
            print('Erro')
                
            fim = time.time()
            print('')
            print(f'Tempo total de: {fim-inicio:.1f} segundos')
            consulta.to_csv(Caminho_Arquivo, sep=';', index=False)
    
            
        consulta.loc[ind_base, 'STATUS'] = "LIBERADO"

        consulta.to_csv(Caminho_Arquivo, sep=';', index=False)
        print('Liberado')

def Selector(Id,Num,Text):
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, Id)))
    time.sleep(0.5)
    for h in range(4):
        driver.find_element(By.ID,Id).click()
        time.sleep(0.5)
        select = Select(driver.find_element(By.ID,Id))
        #time.sleep(0.5)
        try:
            select.select_by_index(Num)
            #time.sleep(0.5)
        except:
            continue
                    
        if driver.find_element(By.ID,Id).get_attribute('value') == Text:
            break   


def Selector_por_nome(id_parse,uf,Class):
    uff = driver.find_element(By.ID,id_parse)
    html = uff.get_attribute('outerHTML')
    uff = BeautifulSoup(html, 'html.parser')
    stat_table = uff.find_all('select', class_=Class)

    len(stat_table)
    stat_table = stat_table[0]
    qtd2 = 0
    #print(stat_table)
    try:
        os.remove(qtd_uf)
    except:
        print('nada aqui')
    f = open(qtd_uf, 'a+')
    f.write('name' + ';' + 'id' + '\n')
    for row in stat_table.find_all('option'):
        f = open(qtd_uf, 'a+')
        f.write(row.text + ';' + str(qtd2) + '\n')
        f.close()
        qtd2 = qtd2 + 1

    with open(qtd_uf, 'r') as rd_l:
        csv_dict = csv.DictReader(rd_l, delimiter=';')
        for lin in csv_dict:
            name = lin['name']
            id3 = lin['id']
            if name == uf:
                print(f'aqui .{name}')
                
                Selector(id_parse,id3,name)

                            

def Consultar_cli(ind_base, linha2):
    print('Digitando')
    print('_______')

    # IR ABA BANCO
    # driver.switch_to.window(aba_banco2)
    winds = driver.window_handles
    driver.switch_to.window(winds[1])
                

    CPF = linha2["CPF"]
    Nome = linha2['Nome']
    Tipo_d_Opera = linha2['Tipo_de_Operacao']
    GrupoConvenio =linha2['GrupoConvenio']
    Nasc = linha2['DataNascimento']
    Matricula = linha2['Matricula']
    Renda = linha2['RendaLiquida']
    Dt_contraCheq = linha2['Data_Contracheque']
    Naturalidade = linha2['Naturalidade']
    Sexo = linha2['Sexo']
    EstadoCivil = linha2['EstadoCivil']
    Tipo_de_Doc = linha2['Tipo_de_Doc']
    Rg = linha2['Numero_Documento']
    Emissor = linha2['Emissor']
    Uf = linha2['UF']
    Data_Emissao = linha2['DataEmissao']
    Mae = linha2['Mae']
    Ddd = linha2['DDD']
    Celular = linha2['Celular']
    Cep = linha2['CEP']
    Endereco = linha2['Endereco']
    Numero = linha2['Numero']
    Bairro = linha2['Bairro']
    Cidade = linha2['Cidade']
    Prazo = linha2['Prazo']
    Valor_Parcela = linha2['ValorParcela']
    
    

    print(f"2 do conculta Índice Base: {ind_base}")
    print(f"CPF do Cliente: {CPF}")
    print(f"Nome do Cliente: {Nome}")
    print(f"Nasc: {Nasc}")


                
    try:
        # tipo de operacao
        id_tipo_oper = 'ctl00_Cph_UcPrp_FIJN1_JnDadosIniciais_UcDIni_cboTipoOperacao_CAMPO'
        Selector(id_tipo_oper,1,'Proposta Nova')
        # Grupo de Convênio:
        time.sleep(0.5)
        id_convenio = 'ctl00_Cph_UcPrp_FIJN1_JnDadosIniciais_UcDIni_cboGrupoConvenio_CAMPO'
        Selector(id_convenio,2,'INSS')
        # orgao
        id_orgao = 'ctl00_Cph_UcPrp_FIJN1_JnDadosIniciais_UcDIni_cboOrigem5_CAMPO'
        Selector(id_orgao,1,'000001 - INSS - APOSENTADORIA')
        
        # cpf
        driver.find_element(By.ID,'ctl00_Cph_UcPrp_FIJN1_JnDadosIniciais_UcDIni_txtCPF_CAMPO').send_keys(CPF)

        
        # Nascimento
        id_nasc = 'ctl00_Cph_UcPrp_FIJN1_JnDadosIniciais_UcDIni_txtDataNascimento_CAMPO'

        WebDriverWait(driver, 4).until(EC.element_to_be_clickable((By.ID, id_nasc))).click()
        time.sleep(3)
        for i in range(3):
            nasc = driver.find_element(By.ID, id_nasc).text
            if nasc == '':
                driver.find_element(By.ID, id_nasc).send_keys(Nasc)
                break
            time.sleep(0.5)
        # Matriicula
        driver.find_element(By.ID,'ctl00_Cph_UcPrp_FIJN1_JnDadosIniciais_UcDIni_ucMatricula_txtMatricula_CAMPO').send_keys(Matricula)

        # renda
        driver.find_element(By.ID,'ctl00_Cph_UcPrp_FIJN1_JnDadosIniciais_UcDIni_txtRenda_CAMPO').send_keys(Renda)

        # contra cheque
        id_contra_cheq = 'ctl00_Cph_UcPrp_FIJN1_JnDadosIniciais_UcDIni_txtDataContraCheque_CAMPO'
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, id_contra_cheq))).click()
        time.sleep(3)
        driver.find_element(By.ID,id_contra_cheq).send_keys(Dt_contraCheq)
        # obter margem
        id_margem ='btnObterMargem_txt'
        driver.find_element(By.ID,id_margem).click()
        # vl_margem
        id_vl_margem = 'ctl00_Cph_UcPrp_FIJN1_JnDadosIniciais_UcDIni_txtValorMargem_CAMPO'
        for i in range(25):
            vl = driver.find_element(By.ID,id_vl_margem)
            Mrg = vl.get_attribute('value')
            if Mrg != '':
                print(f'Valor da Margem = {Mrg}')
                break
            time.sleep(1)
        

        # Nome
        time.sleep(2)
        id_nome = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_txtNome_CAMPO'
        for i in range(3):
            nm_txt = driver.find_element(By.ID,id_nome)
            Nm = nm_txt.get_attribute('value')
            if Nm != '':
                driver.find_element(By.ID,id_nome).clear()
                driver.find_element(By.ID,id_nome).send_keys(Nome)
                break
            time.sleep(1)

        # Naturalidade
        id_naturalidade = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_txtNatural_CAMPO'
        driver.find_element(By.ID,id_naturalidade).send_keys(Naturalidade)
        # sexo
        id_sexo = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_cbxSexo_CAMPO'
        if Sexo.lower() == 'f':
            Selector(id_sexo,2,'Feminino')
        else:
            Selector(id_sexo,1,'Masculino')

        # estado civil
        id_estado_civil = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_cbxEstadoCivil_CAMPO'
        tp = ['solteira','solteiro']
        if EstadoCivil.lower() in tp:
            print(f'Estado :{EstadoCivil.lower()}')
            Selector(id_estado_civil,1,'Solteiro')
        elif EstadoCivil.lower() == 'casado':
            Selector(id_estado_civil,2,'Casado')

        # rg
        id_rg = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_txtDocumento_CAMPO'
        driver.find_element(By.ID,id_rg).send_keys(Rg)

        # emissor
        id_emissor = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_txtEmissor_CAMPO'
        driver.find_element(By.ID,id_emissor).send_keys(Emissor)

        # uf
        id_uf = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_cbxUFDoc_CAMPO'
        Class = 'FICorObr brds'
        Selector_por_nome(id_uf,Uf, Class)

        # data emissao
        id_emissao = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_txtDataEmissao_CAMPO'
        driver.find_element(By.ID,id_emissao).send_keys(Data_Emissao)

        # mae
        id_mae ='ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_txtMae_CAMPO'
        driver.find_element(By.ID,id_mae).send_keys(Mae)

        # ddd
        id_dd = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_txtDddTelCelular_CAMPO'
        driver.find_element(By.ID,id_dd).send_keys(Ddd)
        # celular
        id_celular = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnC_txtTelCelular_CAMPO'
        driver.find_element(By.ID,id_celular).send_keys(Celular)
        # cep
        id_cep ='ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnCR_txtCEP_CAMPO'
        driver.find_element(By.ID,id_cep).send_keys(Cep)
        # endereco
        id_endereco ='ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnCR_txtEndereco_CAMPO'
        driver.find_element(By.ID,id_endereco).click()
        time.sleep(5)
        # numero
        id_numero ='ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnCR_txtNumero_CAMPO'
        try:
            for i in range(4):
                Vg = driver.find_element(By.ID,id_endereco)
                Ru = Vg.get_attribute('value')
                
                if Ru == '':
                    print(f'Endeco= {Ru}')
                    driver.find_element(By.ID,id_endereco).send_keys(Endereco)
                    
                    driver.find_element(By.ID,id_numero).send_keys(Numero)
                    # bairro 
                    id_bairro ='ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnCR_txtBairro_CAMPO'
                    driver.find_element(By.ID,id_bairro).send_keys(Bairro)
                    # cidade
                    id_cidade ='ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnCR_txtCidade_CAMPO'
                    driver.find_element(By.ID,id_cidade).send_keys(Cidade)
                    # uf
                    id_uf2 = 'ctl00_Cph_UcPrp_FIJN1_JnDadosCliente_UcDadosPessoaisClienteSnt_FIJN1_JnCR_cbxUF_CAMPO'
                    Class = 'FICorObr brds'
                    Selector_por_nome(id_uf2,Uf, Class)
                    break
                elif Ru != '':
                    print(f'Rua= {Ru}')
                    driver.find_element(By.ID,id_numero).clear()
                    driver.find_element(By.ID,id_numero).send_keys(Numero)
                    break
        except:
            pass
                
        try:
            # prazo
            id_prazo = 'ctl00_Cph_UcPrp_FIJN1_JnSimulacao_UcSimulacaoSnt_FIJanela1_FIJanelaPanel1_cbxPrazo_CAMPO'
            Class = 'brds'
            Selector_por_nome(id_prazo,Prazo, Class)
        except Exception as e:
            print(e)
        # valor da parcela
        id_parcela ='ctl00_Cph_UcPrp_FIJN1_JnSimulacao_UcSimulacaoSnt_FIJanela1_FIJanelaPanel1_txtVlrParcela_CAMPO'
        mrg = float(Mrg.replace(',', '.'))       
        
        vlf = int(mrg)
        driver.find_element(By.ID,id_parcela).send_keys(str(vlf))
        time.sleep(1)
        

        # referencia
        id_referencia = 'ctl00_Cph_UcPrp_FIJN1_JnClientes_UcDadosClienteSnt_FIJN1_FIJanelaPanel1_txtReferencia1_CAMPO'
        driver.find_element(By.ID,id_referencia).send_keys(Nome)
        # ddd
        id_dd2 = 'ctl00_Cph_UcPrp_FIJN1_JnClientes_UcDadosClienteSnt_FIJN1_FIJanelaPanel1_txtDDDReferencia1_CAMPO'
        driver.find_element(By.ID,id_dd2).send_keys(Ddd)
        # celular
        id_celular2 = 'ctl00_Cph_UcPrp_FIJN1_JnClientes_UcDadosClienteSnt_FIJN1_FIJanelaPanel1_txtTelReferencia1_CAMPO'
        driver.find_element(By.ID,id_celular2).send_keys(Celular)

        # calcular
        id_calcular = 'btnCalcular_txt'
        driver.find_element(By.ID,id_calcular).click()
        # tabela
        
        Div_table = 'ctl00_Cph_UcPrp_FIJN1_JnSimulacao_UcSimulacaoSnt_FIJanela1_FIJanelaPanel1_divScrollConvenio'

        tab = driver.find_element(By.ID,Div_table)
        html = tab.get_attribute('outerHTML')
        tab = BeautifulSoup(html, 'html.parser')
        stat_table = tab.find_all('table', class_='grid-view')

        len(stat_table)
        stat_table = stat_table[0]
        qtd2 = 2
        #print(stat_table)
        
        
        for row in stat_table.find_all('tr'):
            taxa = driver.find_element(By.XPATH,'//*[@id="ctl00_Cph_UcPrp_FIJN1_JnSimulacao_UcSimulacaoSnt_FIJanela1_FIJanelaPanel1_grdCondicoes"]/tbody/tr['+str(qtd2)+']/td[3]').text
            if taxa == 'INSS NOVO - TAXA 2,14':
                driver.find_element(By.XPATH,'//*[@id="ctl00_Cph_UcPrp_FIJN1_JnSimulacao_UcSimulacaoSnt_FIJanela1_FIJanelaPanel1_grdCondicoes"]/tbody/tr['+str(qtd2)+']/td[1]').click()
                consulta.loc[ind_base, 'DIGITADO'] = "SIM"
                consulta.to_csv(Caminho_Arquivo, sep=';', index=False)
                texto_alerta = 'passou'
                fim = time.time()
                print('Finalizado a consulta \n')
                print(f'Tempo total de: {fim-inicio:.1f} segundos')
                sg.popup_yes_no('Faz ae')
                #driver.find_element(By.XPATH,'//*[@id="ctl00_Cph_UcPrp_FIJN1_JnBotoes_UcBotoes_btnGravar_dvCBtn"]').click()
                
                break
            elif qtd2 >= 9:
                
                consulta.loc[ind_base, 'DIGITADO'] = "Reprovado"
                consulta.to_csv(Caminho_Arquivo, sep=';', index=False)
                texto_alerta = 'passou'
                fim = time.time()
                print('Finalizado a consulta \n')
                print(f'Tempo total de: {fim-inicio:.1f} segundos')
                driver.find_element(By.XPATH,'//*[@id="ctl00_Cph_UcPrp_FIJN1_JnBotoes_UcBotoes_btnCancelar_dvCBtn"]').click()
                break
            qtd2 = qtd2 + 1

        return texto_alerta
    except Exception as er:
        
        fim = time.time()
        print(f'Erro: \n {er}')
        print(f'Tempo total de: {fim-inicio:.1f} segundos')
        return texto_alerta
    
    
                

## Inicio
################# ASSINAR ################
if __name__==('__main__'):
    global images
    images = load_images()
    sts = {'', 'ativo'}
    if val ==  'bloqueado':
        sg.Popup('Hello!', 'Sua licença expirou!')
        exit(69)
    elif val in sts:
        # ABRIR ABA BancoSeguro

        driver.execute_script("window.open('https://consignado.bancopaulista.com.br/WebAutorizador/Login/', '_blank')")
        #driver.get('http://wss.credisim.com.br/BSGWEBSITES/WebAutorizador/Login/')
        driver.implicitly_wait(20)
        winds = driver.window_handles
        driver.switch_to.window(winds[-1])

        aba_banco = driver.current_window_handle

        time.sleep(2)
        ############## LOGAR BANCO
        Login()
        aba_liberar()
        # abrir 3 abas para liberar
        # print_time(1, 3)

        # ABRIR ABA BancoSeguro 2
        nova_tab()
        # ABRIR CONSULTA DE PROPOSTA
        time.sleep(5)
        WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.ID, 'navbar-collapse-funcao')))
        

        try:
            inss_button = driver.find_element(By.XPATH,'//*[@id="navbar-collapse-funcao"]/ul/li[1]/a')
            achains = ActionChains(driver)
            achains.move_to_element(inss_button).perform()
            time.sleep(1)
            driver.find_element(By.ID,'WFP2010_PWCDPRPS').click()
        except:
            inss_button = driver.find_element(By.XPATH,'//*[@id="navbar-collapse-funcao"]/ul/li[1]/a')
            achains = ActionChains(driver)
            achains.move_to_element(inss_button).perform()
            time.sleep(1)
            driver.find_element(By.ID,'WFP2010_PWCDPRPS').click()

        sg.popup_yes_no('Configure a pasta para dowload dos termos')

            
        inicio = time.time()
        consulta = pd.read_csv(Caminho_Arquivo, sep=';', dtype=str, encoding='latin1')
        consulta_livre = consulta[(consulta['STATUS'].isnull())]

            
        for ind_base, linha2 in consulta_livre.iterrows():
            ######## CONSULTA
            CPF = linha2["CPF"]
            Nome = linha2['Nome']
            

            texto1 = 'Não foi encontrado nenhum cliente para o CPF informado. Para que seja possível prosseguir, é necessário efetuar a impressão do termo de autorização de consulta de dados do beneficiário e, após, é necessário solicitar a autorização de consulta de dados do beneficiário junto a DataPrev através da rotina Autorização para Consulta de Dados do Beneficiário'
            try: 
                #Liberar_cli(ind_base, linha2, aba_banco)
                Consultar_cli(ind_base, linha2)
            except Exception as e:
                """Liberar_cli(ind_base, linha2, aba_banco)
                Consultar_cli(ind_base, linha2)"""
                

    