from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import os

pasta_atual = os.getcwd()

class ASSIST:
    def __init__(self, nav=None, headless=False) -> object:
        if nav == 'Chrome':
            from selenium.webdriver.chrome.options import Options
            options = Options()
            options.headless = headless
            self.browser = webdriver.Chrome(options=options)
        else:
            from selenium.webdriver.firefox.options import Options
            options = Options()
            options.headless = headless
            options.set_preference("browser.download.folderList", 2)
            options.set_preference("browser.download.manager.showWhenStarting", False)
            options.set_preference("browser.download.dir", f"{pasta_atual}\downloaded_report")
            options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/x-www-form-urlencoded")
            self.browser = webdriver.Firefox(options=options)
        self.wait = WebDriverWait(self.browser, 10)
        self.OS_para_avaliar = []
        self.id_do_tecnico = 42
        self.id_do_defeito = ''
        self.id_da_solucao = ''
        self.descricao_do_defeito_para_avaliar = ''
        self.cod_da_peca = ''
        self.defeitos_nao_liga = ['NAO CARREGA', 'NÃO ESTÁ CARREGANDO', 'NÃO CARREGA', 'NÃO LIGA',
                                  'NAO LIGA', 'NÃO ESTÁ LIGANDO', 'NAO ESTÁ LIGANDO', 'NÃO RECARREGA',
                                  'NAO RECARREGA', 'NAO MANTEM A CARGA', 'NÃO MANTEM A CARGA', 'PAROU DE CARREGAR',
                                  'NÃO ESTA LIGANDO', 'NAO ESTA LIGANDO']

        self.defeitos_botao_emperrado = ['EMPERRANDO', 'TRAVANDO', 'AGARRANDO', 'ENGANCHANDO', 'TRAVADO', 'EMPERRADO',
                                         'EMPURRANDO']

        self.defeitos_analogico = ['ANALÓGICO', 'DRIFTING', 'ANALOGICO', 'ANALGICO', 'ANALÓGICA', 'ANALIGCO', 'DRIFT',
                                   'DIRECIONAL', 'DIRECIONAIS', 'ALAVANCA', 'PUXANDO', 'MECHENDO SOZINHO',
                                   'MOVIMENTANDO SOZINHO', 'MEXE SOZINHO', 'PUXA', 'R3', 'L3', 'MANETE', 'ANALOGICOS',
                                   'ANALÓGICOS', 'JOYSTICK' ]

        self.defeitos_gatilho = ['GATILHO', 'GATILHOS', 'FROUXO', 'MOLE', 'R2', 'L2', 'RESISTÊNCIA', 'RESISTENCIA']

        self.defeitos_botao = ['BOTOES', 'BOTÕES', 'BOTAO', 'BOTÃO', 'QUADRADO', 'TRIANGULO', 'TRIÂNGULO',
                               'BOLA', 'BOLINHA', 'SETAS', 'SETA', 'SETINHA', 'SETINHAS', 'TOUCH PAD', 'TOUCH']
        self.modelo_do_produto = ''
        self.defeito_reclamado = ''
        self.saldo = ''
        self.lista_de_peca = \
            {'SONFIZCT1W - CFI-ZCT1W PS5 CONTROLE DUALSENSE':
                {'peça': 'SON0694', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 5,
                 'defeito_gatilho': 21, 'defeito_botao': 35, 'botao_emperrado': 7, 'nao_liga': 8},

             'SONCFIZCT1WRED - CFI-ZCT1W-RED':
                {'peça': 'SON0827', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 5,
                 'defeito_gatilho': 21, 'defeito_botao': 35, 'botao_emperrado': 7, 'nao_liga': 8},

             'SONCUHZCT2USTBLK - CUH-ZCT2U CONTROLE WIRELESS PS4 STEEL BLACK':
                 {'peça': 'SON0636', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 2,
                  'defeito_gatilho': 6, 'defeito_botao': 3, 'botao_emperrado': 3, 'nao_liga': 16},

             'SONCUH-ZCT2UBLACK - CUH-ZCT2U CONTROLE WIRELESS PS4 JET BLACK':
                {'peça': 'SON0598', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 2,
                 'defeito_gatilho': 6, 'defeito_botao': 3, 'botao_emperrado': 3, 'nao_liga': 16},

             'SONCUHZCT2U16XVRDCAM - CUH-ZCT2U CONTROLE WIRELESS PS4 16X CONTROLE DUALSHOCK VERDE CAMUFLADO':
                {'peça': 'SON0627', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 2,
                 'defeito_gatilho': 6, 'defeito_botao': 3, 'botao_emperrado': 3, 'nao_liga': 16},

             'SONCUH-ZCT2UMRED - CUH-ZCT2U CONTROLE WIRELESS PS4 MAGMA RED':
                {'peça': 'SON0594', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 2,
                 'defeito_gatilho': 6, 'defeito_botao': 3, 'botao_emperrado': 3, 'nao_liga': 16},

             'SONCUH-ZCT2U - CUH-ZCT2U CONTROLE WIRELESS PS4':
                {'peça': 'SON0598', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 2,
                 'defeito_gatilho': 6, 'defeito_botao': 3, 'botao_emperrado': 3, 'nao_liga': 16},

             'SONCFIZCT1WBLACK - CFI-ZCT1W-BLACK':
                {'peça': 'SON0837', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 5,
                 'defeito_gatilho': 21, 'defeito_botao': 35, 'botao_emperrado': 7, 'nao_liga': 8},

             'SONCUHZCT2U22XMNBL - CUH-ZCT2U CONTROLE WIRELESS PS4 22X MIDNIGHT BLUE':
                {'peça': 'SON0628', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 2,
                 'defeito_gatilho': 6, 'defeito_botao': 3, 'botao_emperrado': 3, 'nao_liga': 16},

             'SONCUHZCT2UGWH - CUH-ZCT2U CONTROLE WIRELESS PS4 GLACIER WHITE':
                {'peça': 'SON0637', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 2,
                 'defeito_gatilho': 6, 'defeito_botao': 3, 'botao_emperrado': 3, 'nao_liga': 16},

             'SONCUHZCT2USUNSETO - CUH-ZCT2U CONTROLE WIRELESS PS4 SUNSET ORANGE':
                {'peça': 'SON0651', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 2,
                 'defeito_gatilho': 6, 'defeito_botao': 3, 'botao_emperrado': 3, 'nao_liga': 16},

             'SONCFIZCPS5CONDSNP - CFI-ZCT1W-NOVA PINK':
                 {'peça': 'SON0892', 'peça_botao': 'SON0844', 'peça_gatilho': 'SON0836', 'defeito_analogico': 5,
                  'defeito_gatilho': 21, 'defeito_botao': 35, 'botao_emperrado': 7, 'nao_liga': 8}
             }

    def extrair_os(self, arquivo_com_as_os):
        arquivo = open(arquivo_com_as_os, "r")

        for linha in arquivo:
            if len(linha) > 3:
                linha = linha.strip()
                self.OS_para_avaliar.append(linha)

        arquivo.close()
        return self.OS_para_avaliar

    def acessar_assist(self, url=''):
        if url:
            self.browser.get(url)
        else:
            self.browser.get('https://pontonet.assistonline.com.br/bin/at/pesquisaOs.php')

    def login(self, login, password):
        usuario = self.browser.find_element("xpath", '//*[@id="usuario_login"]')
        senha = self.browser.find_element('xpath', '/html/body/table/tbody/tr[2]/td[2]/div[2]/div/form/'
                                                   'table/tbody/tr[2]/td/input')

        botao_entrar = self.browser.find_element('xpath', '/html/body/table/tbody/tr[2]/td[2]/div[2]/div/form/'
                                                          'table/tbody/tr[5]/td/input')

        usuario.send_keys(login)
        senha.send_keys(password)
        botao_entrar.click()
        sleep(0.5)
        get_url = self.browser.current_url
        if get_url == 'https://pontonet.assistonline.com.br/index.php?erro=1':
            raise RuntimeError('Usuário inválido')

    def acessar_os(self, os):
        if len(str(os)) == 6:
            url_ordem = f'https://pontonet.assistonline.com.br/bin/at/detalhes_os.php?osAbertura_id={os}'
            try:
                self.browser.get(url_ordem)
            except:
                raise RuntimeError(f'Erro ao tentar acessar a OS pelo link \n {url_ordem}')
        else:
            self.acessar_assist()
            campo_preencher_os = self.browser.find_element('xpath', '//*[@id="os"]')
            botao_find = self.browser.find_element('xpath', '/html/body/table[2]/tbody/tr/td[2]/table/'
                                                            'tbody/tr[2]/td/form/table/tbody/tr[14]/td/img')
            campo_preencher_os.send_keys(os)
            botao_find.click()
            try:
                entrar_na_os = self.wait.until(EC.presence_of_element_located(('xpath', '/html/body/table[2]/tbody/tr/'
                                                                                        'td[2]/table/tbody/tr[2]/td/div/'
                                                                                        'table/tbody/tr[3]/td[3]/a/img')))
                entrar_na_os.click()
            except:
                raise RuntimeError('Erro ao tentar acessar a OS')

    def verificar_status(self):
        reincidencia = 'Não é uma reincidência'
        status_os = 'Não verificado'
        try:
            lista = self.browser.find_elements('class name', 'link1')
            for n, elemento in enumerate(lista):
                if elemento.text.startswith('REINCIDÊNCIA:'):
                    reincidencia = elemento.text.strip(';')

                if elemento.text.startswith('Status:'):
                    status_os = elemento.text.removeprefix('Status:')
                    break

        except:
            raise RuntimeError('Erro ao verificar Status da OS')

        finally:
            return status_os.strip(), reincidencia.strip()

    def verificar_modelo_do_produto(self):
        modelo_do_produto = ''
        aba_detalhes_do_atendimento = self.browser.find_elements('tag name', 'table')

        for elemento in aba_detalhes_do_atendimento:
            if elemento.text == 'DETALHES DO ATENDIMENTO':
                elemento.click()
                break
        modelo = self.browser.find_element('id', 'atendimento').find_elements('tag name', 'tr')

        for elemento in modelo:
            if elemento.text.startswith('Descrição Modelo'):
                modelo_do_produto = elemento.text.removeprefix('Descrição Modelo')
                break
            else:
                modelo_do_produto = 'Não foi possível localizar o modelo do produto'

        self.modelo_do_produto = modelo_do_produto.strip()
        return modelo_do_produto.strip()

    def verificar_defeito_reclamado(self):

        aba_avaliacao = self.wait.until(EC.presence_of_element_located(('xpath', '/html/body/table[2]/tbody/tr/td[2]/table/'
                                                           'tbody/tr/td[2]/table[4]/tbody/tr/td[2]/a/span/b')))
        aba_avaliacao.click()
        defeito_reclamado = self.wait.until(EC.presence_of_element_located(('xpath', '/html/body/table[2]/tbody/tr/td[2]/table/tbody/tr/'
                                                               'td[2]/div[2]/table/tbody/tr/td[2]/table/tbody/tr[5]/td')))
        defeito = defeito_reclamado.text.upper()
        self.defeito_reclamado = defeito
        return defeito

    def apertar_botao_para_avaliar(self):
        botao = self.wait.until(EC.presence_of_element_located(('partial link text', f'Clique aqui para acessa-la.')))
        botao.click()

    def analise_do_defeito(self):
        if self.defeito_reclamado == '':
            raise RuntimeError('Defeito reclamado não foi verificado, usar função - verificar_defeito_reclamado()')

        if 'PROCON' in self.defeito_reclamado:
            raise RuntimeError('CASO DE PROCON')
        if 'LAG' in self.defeito_reclamado:
            raise RuntimeError('Caso de InputLag, não avaliado')

        for analogico in self.defeitos_analogico:
            if analogico in self.defeito_reclamado:
                return 'analogico'

        for nao_liga in self.defeitos_nao_liga:
            if nao_liga in self.defeito_reclamado:
                return 'não liga'

        for gatilho in self.defeitos_gatilho:
            if gatilho in self.defeito_reclamado:
                return 'gatilho'

        for botao_emperrado in self.defeitos_botao_emperrado:
            if botao_emperrado in self.defeito_reclamado:
                return 'botao emperrando'

        for botao in self.defeitos_botao:
            if botao in self.defeito_reclamado:
                return 'botao'

        return 'Defeito não encontrado!'

    def set_pendencia_validacao(self):
        confirmar = self.wait.until(EC.presence_of_element_located(('id', 'btn_confirmar')))
        confirmar.click()
        alert = Alert(self.browser)
        alert.accept()
        sleep(1)

    def set_avaliacao(self):
        self.limpar_avaliacao_anterior()
        if self.modelo_do_produto == '':
            raise RuntimeError('Não foi verificado o modelo do produto, usar função - verificar_modelo_do_produto()')

        qual_o_defeito = self.analise_do_defeito()
        self.__set_defeito(qual_o_defeito)

        self.apertar_botao_para_avaliar()

        botao_nome = self.wait.until(EC.presence_of_element_located(('css selector', f'#nomeTecnico > option:nth-child'
                                                                                     f'({self.id_do_tecnico})')))
        botao_nome.click()

        botao_defeito = self.browser.find_element('css selector', f'#td_defeito > select:nth-child(1) > '
                                                                  f'option:nth-child({self.id_do_defeito})')

        desc_do_defeito = self.browser.find_element('xpath', '//*[@id="descrDefeito"]')

        peca = self.browser.find_element('xpath', '/html/body/table[2]/tbody/tr/td[2]/form/table/tbody/tr/td/'
                                                  'table[3]/tbody/tr[1]/td/table[1]/tbody/tr[4]/td[1]/input')

        carregar_peca = self.browser.find_element('xpath', '//*[@id="id_qtdPeca"]')

        desc_do_defeito.send_keys(self.descricao_do_defeito_para_avaliar)
        botao_defeito.click()

        solucao = self.wait.until(EC.presence_of_element_located(('css selector', f'#td_solucao > select:nth-child(1) > '
                                                                                  f'option:nth-child({self.id_da_solucao})')))
        solucao.click()

        sleep(2)
        peca.send_keys(self.cod_da_peca)
        carregar_peca.send_keys('1')

        sleep(2)
        botao_adicionar_peca = self.wait.until(EC.presence_of_element_located(('xpath', '/html/body/table[2]/tbody/tr/td[2]/form/table/tbody'
                                                                                        '/tr/td/table[3]/tbody/tr[1]/td/table[1]/tbody/tr[2]/'
                                                                                        'td[3]/input[1]')))
        botao_adicionar_peca.click()

        botao_gravar_avaliacao = self.wait.until(EC.presence_of_element_located(('xpath', '//*[@id="button_1"]')))
        botao_gravar_avaliacao.click()
        sleep(1)

        if qual_o_defeito in ['gatilho', 'botao', 'botao emperrando'] or self.saldo == 'Saldo Salvage':
            try:
                alert = Alert(self.browser)
                alert.accept()
            except:
                print('Não há botão para apertar!')

        if 'PS4' in self.modelo_do_produto or qual_o_defeito == 'analogico' or qual_o_defeito == 'não liga':
            sleep(1)
            serial_do_controle = self.browser.find_element('xpath', '/html/body/table[2]/tbody/tr/td[2]/table[2]/'
                                                                    'tbody/tr/td/form/table[1]/tbody/tr/td[2]').text

            inserir_serial = self.browser.find_element('xpath', '/html/body/table[2]/tbody/tr/td[2]/table[2]/'
                                                                'tbody/tr/td/form/table[2]/tbody/tr[2]/td/table/'
                                                                'tbody/tr[2]/td/input')

            if len(serial_do_controle) > 3:
                inserir_serial.send_keys(serial_do_controle)
                concluir = self.browser.find_element('xpath','/html/body/table[2]/tbody/tr/td[2]/table[2]'
                                                             '/tbody/tr/td/form/table[2]/tbody/tr[3]/td/input')
                concluir.click()
            else:
                inserir_serial.send_keys('NT')
                concluir = self.browser.find_element('xpath', '/html/body/table[2]/tbody/tr/td[2]/'
                                                              'table[2]/tbody/tr/td/form/table[2]/tbody/tr[3]/td/input')
                concluir.click()

    def limpar_avaliacao_anterior(self):
        self.id_do_defeito = ''
        self.id_da_solucao = ''
        self.descricao_do_defeito_para_avaliar = ''
        self.cod_da_peca = ''

    def set_saldo(self, saldo):
        if saldo in ['Saldo em estoque', 'Saldo Salvage']:
            self.saldo = saldo
        else:
            raise KeyError(f'parametro saldo dever ser "Saldo em estoque" ou "Saldo Salvage": {saldo}')

    def __set_defeito(self, qual_defeito):
        peca_sem_salvage = ['SONCFIZCT1WRED - CFI-ZCT1W-RED',
            'SONCUHZCT2U16XVRDCAM - CUH-ZCT2U CONTROLE WIRELESS PS4 16X CONTROLE DUALSHOCK VERDE CAMUFLADO',
            'SONCUH-ZCT2UMRED - CUH-ZCT2U CONTROLE WIRELESS PS4 MAGMA RED',
            'SONCFIZCT1WBLACK - CFI-ZCT1W-BLACK',
            'SONCUHZCT2U22XMNBL - CUH-ZCT2U CONTROLE WIRELESS PS4 22X MIDNIGHT BLUE',
            'SONCUHZCT2UGWH - CUH-ZCT2U CONTROLE WIRELESS PS4 GLACIER WHITE',
            'SONCUHZCT2USUNSETO - CUH-ZCT2U CONTROLE WIRELESS PS4 SUNSET ORANGE',
            'SONCFIZCPS5CONDSNP - CFI-ZCT1W-NOVA PINK']

        if qual_defeito == 'analogico':
            if self.saldo == 'Saldo em estoque':
                self.__avaliar_analogico()
            elif self.saldo == 'Saldo Salvage':
                if self.modelo_do_produto in peca_sem_salvage:
                    self.__avaliar_analogico()
                else:
                    self.__avaliar_analogico_salvage()
            else:
                raise ValueError(f'Saldo utilizado não foi definido, utilize set_saldo()')

        elif qual_defeito == 'não liga':
            if self.saldo == 'Saldo em estoque':
                self.__avaliar_nao_liga()
            elif self.saldo == 'Saldo Salvage':
                if self.modelo_do_produto in peca_sem_salvage:
                    self.__avaliar_nao_liga()
                else:
                    self.__avaliar_nao_liga_salvage()
            else:
                raise ValueError(f'Saldo utilizado não foi definido, utilize set_saldo()')

        elif qual_defeito == 'gatilho':
            self.__avaliar_gatilho()

        elif qual_defeito == 'botao emperrando':
            self.__avaliar_botao_emperrado()

        elif qual_defeito == 'botao':
            self.__avaliar_botao()

        elif qual_defeito == 'Defeito não encontrado!':
            raise RuntimeError('Defeito não encontrado!')

    def __avaliar_nao_liga_salvage(self):
        self.id_do_defeito = self.lista_de_peca[self.modelo_do_produto]['nao_liga']
        if 'PS4' in self.modelo_do_produto:
            self.id_da_solucao = 10
        else:
            self.id_da_solucao = 9
        self.descricao_do_defeito_para_avaliar = 'Não liga, não carrega.'
        self.cod_da_peca = self.lista_de_peca[self.modelo_do_produto]['peça']

    def __avaliar_nao_liga(self):
        self.id_do_defeito = self.lista_de_peca[self.modelo_do_produto]['nao_liga']
        self.id_da_solucao = 2
        self.descricao_do_defeito_para_avaliar = 'Não liga, não carrega.'
        self.cod_da_peca = self.lista_de_peca[self.modelo_do_produto]['peça']

    def __avaliar_analogico_salvage(self):
        self.id_do_defeito = self.lista_de_peca[self.modelo_do_produto]['defeito_analogico']
        if 'PS4' in self.modelo_do_produto:
            self.id_da_solucao = 10
        else:
            self.id_da_solucao = 9
        self.descricao_do_defeito_para_avaliar = 'Falha no analógico.'
        self.cod_da_peca = self.lista_de_peca[self.modelo_do_produto]['peça']

    def __avaliar_analogico(self):
        self.id_do_defeito = self.lista_de_peca[self.modelo_do_produto]['defeito_analogico']
        self.id_da_solucao = 2
        self.descricao_do_defeito_para_avaliar = 'Falha no analógico.'
        self.cod_da_peca = self.lista_de_peca[self.modelo_do_produto]['peça']

    def __avaliar_gatilho(self):
        self.id_do_defeito = self.lista_de_peca[self.modelo_do_produto]['defeito_gatilho']
        if 'PS4' in self.modelo_do_produto:
            self.id_da_solucao = 10
        else:
            self.id_da_solucao = 9
        self.descricao_do_defeito_para_avaliar = 'Falha no gatilho.'
        self.cod_da_peca = self.lista_de_peca[self.modelo_do_produto]['peça_gatilho']

    def __avaliar_botao_emperrado(self):
        self.id_do_defeito = self.lista_de_peca[self.modelo_do_produto]['botao_emperrado']
        if 'PS4' in self.modelo_do_produto:
            self.id_da_solucao = 3
        else:
            self.id_da_solucao = 3
        self.descricao_do_defeito_para_avaliar = 'Botões travando.'
        self.cod_da_peca = self.lista_de_peca[self.modelo_do_produto]['peça_botao']

    def __avaliar_botao(self):
        self.id_do_defeito = self.lista_de_peca[self.modelo_do_produto]['defeito_botao']
        if 'PS4' in self.modelo_do_produto:
            self.id_da_solucao = 3
        else:
            self.id_da_solucao = 3
        self.descricao_do_defeito_para_avaliar = 'Falha nos botões.'
        self.cod_da_peca = self.lista_de_peca[self.modelo_do_produto]['peça_botao']

    def definir_tecnico(self, nome):
        lista_id = {'ALEF LOPES BARROS': 37,
                    'ALEXANDRE SILVA SANTOS': 54,
                    'ALEXSANDER LOPES LIMA': 28,
                    'ALINE DE ASSIS SANTOS CARIRI': 59,
                    'ANA CLAUDIA SOUZA': 22,
                    'ANA PAULA DA SILVA': 31,
                    'ANA PAULA DE OLIVEIRA': 24,
                    'ANDERSON BATISTA DOS SANTOS': 17,
                    'ANDRE DE PAULA BENACHIO': 14,
                    'ANGELO BORGES REIS': 51,
                    'BRUNO FREITAS DA SILVA DIAS': 20,
                    'CAIQUE DOS SANTOS RESENDE': 62,
                    'CIBIELE PATRICIA DE ALMEIDA': 34,
                    'Carlos Eduardo': 32,
                    'Caroline Moura Pardinho': 65,
                    'DANIEL JOAQUIM VITOR': 38,
                    'DEBORA CRISTINA SILVA SANTANA': 41,
                    'DENIS RAMOS FERRAZ': 7,
                    'DOUGLAS ARAUJO DA SILVA': 16,
                    'DOUGLAS FELINTO DOS SANTOS': 12,
                    'ELAINE DE MORAIS NEVES': 40,
                    'ERNESTO GRIMALDI': 6,
                    'FELIPE GOMES': 13,
                    'GIOVANA CESAR': 56,
                    'GUILHERME ASEVEDO ZACANTE': 45,
                    'GUILHERME SOARES': 53,
                    'HEBERT RODRIGUES': 60,
                    'HENRIQUE FERREIRA': 43,
                    'ITALO SANTOS': 25,
                    'JHONATAN BORGES TALHATTI': 15,
                    'JOAO RICARDO DE MORAIS': 10,
                    'JOSUE SILVA DOS SANTOS': 39,
                    'KAUAN BRUNO FERREIRA': 63,
                    'LAISA PEDROSO': 55,
                    'LEANDRO GAMA MACHADO': 36,
                    'LUCAS DE SOUZA ESTEVES': 18,
                    'LUCAS PATRICK': 44,
                    'LUCAS PEREIRA DOS SANTOS': 46,
                    'LUIZ CARLOS DE SOUZA': 21,
                    'MANOELA BARROS': 64,
                    'MARCELO ITALO': 9,
                    'MARILIA GABRIELA BARRETO DE FREITAS': 4,
                    'MATHEUS NERY SILVA': 29,
                    'Marcela Dionê Bezerra de Moraes': 58,
                    'NATHAN SOUZA ALVARENGA': 42,
                    'RAPHAEL YUKIO KOGA BARBOSA DOS SANTOS': 8,
                    'RENAN CUSTODIO DA SILVA': 30,
                    'RENAN SANTOS': 49,
                    'RENATO JOSEFINO PEREIRA': 23,
                    'RENATO TOLEDO DO AMORIM': 33,
                    'RODRIGO AUGUSTO SOARES NEGRI': 2,
                    'ROMÁRIO TRIGUEIRO LIMA': 48,
                    'RONNI NONATO SILVA': 11,
                    'RUBIA LOPES CORRADINI': 3,
                    'SANDRA PEREIRA DA SILVA': 61,
                    'SERGIO VINICIUS': 52,
                    'SILVANA ROBERTA ALMEIDA SANTOS MONARO': 27,
                    'SIMONE SOARES DE SOUZA': 26,
                    'TATIANE SILVA': 57,
                    'THIAGO CESAR AMANCIO': 5,
                    'TIAGO DE OLIVEIRA CHAVES': 19,
                    'VICTOR SANTANA BONIN': 35,
                    'WERMESON FELIPE SENA DA SILVA': 47,
                    'WERMESON SILVA': 50}
        id_tecnico = lista_id.get(nome, 'valor não encontrado')
        if id == 'valor não encontrado':
            raise RuntimeError('Nome do técnico não existe na lista')
        self.id_do_tecnico = id_tecnico

    def desfazer_avaliacao(self):
        self.apertar_botao_para_avaliar()
        botao_alterar = self.browser.find_element('xpath', '//*[@id="osConserto"]/table/tbody/tr/td/table[6]/tbody/tr/td[1]/input')

        botao_alterar.click()
        sleep(1)

        botao_desfazer = self.browser.find_element('xpath', '//*[@id="AVAL"]/table/tbody/tr/td/table[3]/tbody/tr[2]/td[1]/input')
        botao_desfazer.click()

    def get_sac(self):
        sac = self.browser.find_elements('class name', 'numOS')
        sac = sac[1].text[5:]
        return sac

    def get_os(self):
        model_list = (
            'CUH-ZCT2U CONTROLE WIRELESS PS4 GLACIER WHITE',
            'CFI-ZCT1W-RED',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 JET BLACK',
            'CFI-ZCT1W PS5 CONTROLE DUALSENSE',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 MAGMA RED',
            'CFI-ZCT1W-BLACK',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 16X CONTROLE DUALSHOCK VERDE CAMUFLADO',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 22X MIDNIGHT BLUE',
            'CUH-ZCT2U CONTROLE WIRELESS PS4',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 ELECTRIC PURPLE',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 SUNSET ORANGE',
            'CUH-ZCT2UZTX CONTROLE WIRELESS PS4 DUALSHOCK TLOU2',
            'CUH-ZCT1U CONTROLE WIRELESS PS4',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 TITANIUM BLUE',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 GOLD',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 ROSE GOLD',
            'CUH-ZCT2E CONTROLE WIRELESS PS4 PRETO',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 FORTNITE',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 STEEL BLACK',
            'CFI-ZCT1W-NOVA PINK',
            'CFI-ZCT1W-GALACTIC PURPLE',
            'CUH-ZCT2U CONTROLE DUALSHOCK VERDE CAMUFLADO',
            'CUH-ZCT2U CONTROLE WIRELESS PS4 CONTROLE DUALSHOCK VERMELHO CAMUFLADO - RED CMOU',
            'CFI-ZCT1W1-BLUE'
        )

        hoje = datetime.now().strftime('%d%m%Y')
        report_file = f'{os.getcwd()}/downloaded_report/Relatório_gerencial_zte_{hoje}.xlsx'
        if os.path.exists(report_file):
            os.remove(report_file)

        self.acessar_assist('https://pontonet.assistonline.com.br/bin/adm/relatorios_adm.php')
        x = self.browser.find_elements('tag name', 'b')
        for element in x:
            if element.text == 'RELATÓRIO GERENCIAL ZTE':
                x = element
                break
        x.click()

        print('Inserindo informações do relatório')
        form = self.browser.find_element('name', 'rel_adm_326')
        data_inicial = form.find_element('id', 'dataInicial')
        data_final = form.find_element('id', 'dataFinal')
        garantia = form.find_element('id', 'garantia_id')
        sony_garantia = garantia.find_elements('tag name', 'option')
        for element in sony_garantia:
            if element.text == 'SONY':
                element.click()

        data_inicial.send_keys('01122021')
        data_final.send_keys(hoje)
        form.find_element('tag name', 'img').click()
        print('Baixando relatório')

        sleep(30)

        relatorio = pd.read_excel(report_file)
        relatorio = relatorio.loc[relatorio['Modelo'].isin(model_list)]
        sem_avaliacao = relatorio.loc[relatorio['Status Atual'] == 'Pendência de Avaliação Técnica']

        os_sem_avaliacao = sem_avaliacao['Núm. OS'].tolist()
        return os_sem_avaliacao

    def fazer_qualidade_sony(self):
        self.apertar_botao_para_avaliar()
        passou = self.browser.find_element('name', 'analQualidade2')
        passou = passou.find_elements('tag name', 'option')
        passou[0].click()

        botao_concluir = self.browser.find_element('name', 'btnconcluir')
        botao_concluir.click()

        analise_qualidade_link = self.browser.find_element('partial link text', 'Clique aqui para analisá-la')
        analise_qualidade_link.click()

        nome_jogo = self.browser.find_element('name', 'analQualidade13')
        nome_jogo.send_keys('Não possui jogo')

        aprovado = self.browser.find_element('name', 'aprova')
        aprovado = aprovado.find_elements('tag name', 'option')
        aprovado[0].click()

        concluir = self.browser.find_element('name', 'grava')
        concluir.click()

    def screenshot(self):
        self.browser.save_screenshot('imagem.png')

    def close(self):
        self.browser.close()


def ordem_servico(acao):
    if not acao in ['fazer qualidade sony', 'desfazer avaliacao']:
        raise ValueError(f'"{acao}" não é uma opção válida.')
    login = input('Digite o seu usuário: ')
    senha = input('Digite a sua senha: ')

    assist = ASSIST(False)
    assist.acessar_assist()
    assist.login(login, senha)

    os_para_avaliar = assist.extrair_os(r'ordens.txt')
    for os in os_para_avaliar:
        try:
            assist.acessar_assist()
            assist.acessar_os(os)
            if acao.upper() == 'fazer qualidade sony':
                assist.fazer_qualidade_sony()
            elif acao.upper() == 'desfazer avaliacao':
                assist.desfazer_avaliacao()
                print(f'OS: {os} foi desfeita com sucesso!')
        except:
            print(f'Não foi possível avaliar a OS: {os}')

def teste():
    login = input('Digite o seu usuário: ')
    senha = input('Digite a sua senha: ')

    assist = ASSIST(True)
    assist.acessar_assist()
    assist.login(login, senha)
    assist.screenshot()
    assist.close()

if __name__ == '__main__':
    teste()