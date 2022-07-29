from pprint import pprint
from assist import ASSIST
from datetime import datetime
from time import sleep
from decouple import config
import os
import pandas as pd

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

def get_report_name():
    report_file = f'{os.getcwd()}/downloaded_report/Relatório_gerencial_zte_{hoje}.xlsx'
    return report_file

def erase_previous_report():
    report_file = get_report_name()
    if os.path.exists(report_file):
        os.remove(report_file)

def download_report(login, senha):
    nav = ASSIST(headless=True)
    nav.acessar_assist()
    print('logando')
    nav.login(login, senha)
    nav.acessar_assist('https://pontonet.assistonline.com.br/bin/adm/relatorios_adm.php')
    x = nav.browser.find_elements('tag name', 'b')
    for element in x:
        if element.text == 'RELATÓRIO GERENCIAL ZTE':
            x = element
            break
    x.click()
    print('Inserindo informações do relatório')
    form = nav.browser.find_element('name', 'rel_adm_326')
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
    nav.close()

class ReportFilter:
    def __init__(self, report_path):
        self.relatorio = pd.read_excel(report_path)
        self.relatorio = self.relatorio.loc[self.relatorio['Modelo'].isin(model_list)]

        self.sem_avaliacao = self.relatorio.loc[
            self.relatorio['Status Atual'] == 'Pendência de Avaliação Técnica'
        ]

        self.mecanico = self.relatorio.loc[
            (self.relatorio['Status Atual'] == 'Pendência de Conclusão do Conserto')
             & (self.relatorio['Descrição Peça 1'] == 'MECHANICAL CLEANING')]

        self.mola = self.relatorio.loc[
            (self.relatorio['Status Atual'] == 'Pendência de Conclusão do Conserto')
             & (self.relatorio['Descrição Peça 1'] == 'SPRING,TORTION(TRIGGER)')
        ]

        self.troca_salvage = self.relatorio.loc[
            (self.relatorio['Status Atual'] == 'Pendência de Conclusão do Conserto')
             & (self.relatorio['Reparo Efetuado 1'] == 'SUR (SALVAGE)')
             & (self.relatorio['Descrição Peça 1'] != 'SPRING,TORTION(TRIGGER)')
             & (self.relatorio['Descrição Peça 1'] != 'MECHANICAL CLEANING')]

        self.troca_estoque_atendido = self.relatorio.loc[
            (self.relatorio['Status Atual'] == 'Pendência de Conclusão do Conserto')
             & (self.relatorio['Reparo Efetuado 1'] == 'CONSOLE / UNIT EXCHANGE')
             & (self.relatorio['Descrição Peça 1'] != 'SPRING,TORTION(TRIGGER)')
             & (self.relatorio['Descrição Peça 1'] != 'MECHANICAL CLEANING')]

        self.falta_peca = self.relatorio.loc[
            self.relatorio['Status Atual'] == 'Pendência de Falta de Peças'
        ]
        self.validacao = self.relatorio.loc[
            self.relatorio['Status Atual'] == 'PRODUTO RECEBIDO EM PROCESSO DE VALIDAÇÃO'
        ]

    def convert_int_to_txt(self, lista: list) -> list:
        for i in range(len(lista)):
            lista[i] = str(lista[i])

    def defect_tolist(self):
        preos = [
        self.sem_avaliacao['Pré-OS'].tolist(),
        self.mecanico['Pré-OS'].tolist(),self.mola['Pré-OS'].tolist(),
        self.troca_salvage['Pré-OS'].tolist(),
        self.troca_estoque_atendido['Pré-OS'].tolist(),
        self.falta_peca['Pré-OS'].tolist(),
        self.validacao['Pré-OS'].tolist()
        ]
        for item in preos:
            self.convert_int_to_txt(item)

        os_sem_avaliacao = self.sem_avaliacao['Núm. OS'].tolist() + preos[0]
        os_mecanico = self.mecanico['Núm. OS'].tolist() + preos[1]
        os_mola = self.mola['Núm. OS'].tolist() + preos[2]
        os_troca_salvage = self.troca_salvage['Núm. OS'].tolist() + preos[3]
        os_troca_estoque_atendido = self.troca_estoque_atendido['Núm. OS'].tolist() + preos[4]
        os_falta_peca = self.falta_peca['Núm. OS'].tolist() + preos[5]
        os_validacao = self.validacao['Núm. OS'].tolist() + preos[6]


        order_list = {
            'sem avaliação': os_sem_avaliacao,
            'mecânico': os_mecanico,
            'mola': os_mola,
            'troca salvage': os_troca_salvage,
            'troca estoque atendido': os_troca_estoque_atendido,
            'falta peça': os_falta_peca,
            'processo de validação': os_validacao
        }

        return order_list

    def sem_avaliacao_to_txt(self):
        os_sem_avaliacao = self.validacao['Pré-OS'].tolist() + self.sem_avaliacao['Pré-OS'].tolist()
        with open('ordens_sem_avaliacao.txt', 'w') as arq:
            for ordem in os_sem_avaliacao:
                arq.write(f'{ordem}\n')

if __name__ == '__main__':
    report_file = get_report_name()
    erase_previous_report()
    download_report(config('USER'), config('PASSWORD'))
    report = ReportFilter(report_file)
    pprint(report.defect_tolist())
    report.sem_avaliacao_to_txt()
