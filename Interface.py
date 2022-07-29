import PySimpleGUI as Sg
import pandas as pd
import assist
import os as system
from threading import Thread
from datetime import datetime

finalizar_processo = False
avaliador = ''
data_de_hoje = datetime.now().strftime('%d-%m-%Y_%H-%M')

lista = ['ALEF LOPES BARROS', 'ALEXANDRE SILVA SANTOS', 'ALEXSANDER LOPES LIMA',
         'ALINE DE ASSIS SANTOS CARIRI', 'ANA CLAUDIA SOUZA', 'ANA PAULA DA SILVA',
         'ANA PAULA DE OLIVEIRA', 'ANDERSON BATISTA DOS SANTOS', 'ANDRE DE PAULA BENACHIO',
         'ANGELO BORGES REIS', 'BRUNO FREITAS DA SILVA DIAS', 'CAIQUE DOS SANTOS RESENDE',
         'CIBIELE PATRICIA DE ALMEIDA', 'Carlos Eduardo', 'Caroline Moura Pardinho',
         'DANIEL JOAQUIM VITOR', 'DEBORA CRISTINA SILVA SANTANA', 'DENIS RAMOS FERRAZ',
         'DOUGLAS ARAUJO DA SILVA', 'DOUGLAS FELINTO DOS SANTOS', 'ELAINE DE MORAIS NEVES',
         'ERNESTO GRIMALDI', 'FELIPE GOMES', 'GIOVANA CESAR', 'GUILHERME ASEVEDO ZACANTE',
         'GUILHERME SOARES', 'HEBERT RODRIGUES', 'HENRIQUE FERREIRA', 'ITALO SANTOS',
         'JHONATAN BORGES TALHATTI', 'JOAO RICARDO DE MORAIS', 'JOSUE SILVA DOS SANTOS',
         'KAUAN BRUNO FERREIRA', 'LAISA PEDROSO', 'LEANDRO GAMA MACHADO', 'LUCAS DE SOUZA ESTEVES',
         'LUCAS PATRICK', 'LUCAS PEREIRA DOS SANTOS', 'LUIZ CARLOS DE SOUZA', 'MANOELA BARROS',
         'MARCELO ITALO', 'MARILIA GABRIELA BARRETO DE FREITAS', 'MATHEUS NERY SILVA',
         'Marcela Dionê Bezerra de Moraes', 'NATHAN SOUZA ALVARENGA', 'RAPHAEL YUKIO KOGA BARBOSA DOS SANTOS',
         'RENAN CUSTODIO DA SILVA', 'RENAN SANTOS', 'RENATO JOSEFINO PEREIRA',
         'RENATO TOLEDO DO AMORIM', 'RODRIGO AUGUSTO SOARES NEGRI', 'ROMÁRIO TRIGUEIRO LIMA',
         'RONNI NONATO SILVA', 'RUBIA LOPES CORRADINI', 'SANDRA PEREIRA DA SILVA',
         'SERGIO VINICIUS', 'SILVANA ROBERTA ALMEIDA SANTOS MONARO', 'SIMONE SOARES DE SOUZA',
         'TATIANE SILVA', 'THIAGO CESAR AMANCIO', 'TIAGO DE OLIVEIRA CHAVES', 'VICTOR SANTANA BONIN',
         'WERMESON FELIPE SENA DA SILVA', 'WERMESON SILVA']
lista_de_erros = []
relatorio_os, relatorio_stats, relatorio_reincidencia, relatorio_model, relatorio_defect, \
relatorio_analise, relatorio_conclusao = [], [], [], [], [], [], []

Sg.theme('Reddit')
cor_botao = ('white', '#2b5c9c')
cor_azul = '#84a4c4'


def limpar_senha():
    with open('text/senha.txt', 'w') as arquivo:
        login = 'login:\n'
        senha = 'senha:\n'
        salvar = 'salvar:False'
        arquivo.write(f'{login}{senha}{salvar}')


def verificar_senha_salva():
    if not system.path.isdir('text'):
        system.mkdir('text')
        limpar_senha()

    with open('text/senha.txt', 'r') as arquivo:
        linhas = arquivo.readlines()
        for linha in linhas:
            linha = linha.strip()
            if 'login' in linha:
                login = linha[linha.find(':') + 1:]
            if 'senha' in linha:
                senha = linha[linha.find(':') + 1:]
            if 'salvar' in linha:
                salvar = linha[linha.find(':') + 1:]
    return login, senha, salvar


def salvar_senha(lg, password, save):
    with open('text/senha.txt', 'w') as arquivo:
        login = f'login:{lg}\n'
        senha = f'senha:{password}\n'
        salvar = f'salvar:{save}'
        arquivo.write(f'{login}{senha}{salvar}')


def fazer_janela1():
    layout_col1 = [
        [Sg.Text(f'Usuário {57 * " "}', font=('Helvetica', '12'), background_color=cor_azul)],
        [Sg.InputText(key='login', size=(41, 1), background_color='white')],
        [Sg.Text(f'Senha {60 * " "}', font=('Helvetica', '12'), background_color=cor_azul)],
        [Sg.InputText(key='senha', size=(41, 1), background_color='white', password_char='•')],
        [Sg.Checkbox(f'Salvar senha?', key='salvar_senha', background_color=cor_azul, checkbox_color='white',
                     font=('Helvetica', '12'), default=True)],
        [Sg.Checkbox(f'Esconder o navegador?', key='ver_navegador', background_color=cor_azul,
                     checkbox_color='white', default=False,
                     font=('Helvetica', '12'))],
        [Sg.Text(' ', background_color=cor_azul)],
        [Sg.Text(f'{17 * " "} ', background_color=cor_azul),
         Sg.Button('Login', size=(10, 1), button_color=cor_botao, font=('Helvetica', '15'), disabled=False)],
        [Sg.Text(' ', background_color=cor_azul)]
    ]

    layout = [
        [Sg.Image('Image/logo_assist.png'), Sg.Image('Image/logo_pontonet.png')],
        [Sg.Image('Image/menu.png')],
        [Sg.Column(layout_col1, element_justification='l', background_color=cor_azul)],
        [Sg.Text(' ')]
    ]
    return Sg.Window('Avaliador Automático', layout, element_justification='c', finalize=True, resizable=True)


def fazer_janela2():
    layout_col = [
        [Sg.Text('Selecione o técnico que irá avaliar:', font=('Helvetica', '11'))],
        [Sg.InputCombo(lista, button_background_color='#2b5c9c', size=(55, 1), key='tecnico')],
        [Sg.Text('Utilizar:', font=('Helvetica', '11'))],
        [Sg.InputCombo(['Saldo em estoque', 'Saldo Salvage'],button_background_color='#2b5c9c', size=(55, 1), key='saldo')],
        [Sg.Text('')],
        [Sg.Text('Escolha o arquivo com as ordens de serviço:                           ', font=('Helvetica', '11'))],
        [Sg.InputText(key='-arquivo-'), Sg.FileBrowse('Escolher...', button_color=cor_botao)],
    ]

    layout = [
        [Sg.Image('Image/logo_play.png')],
        [Sg.Column(layout_col, element_justification='l')],
        [Sg.Text('')],
        [Sg.Button('Fazer as avaliações', key='avaliar', size=(20, 3), button_color=cor_botao, font=('Helvetica', '12'))]
    ]
    return Sg.Window('Avaliador Automático', layout, element_justification='c', finalize=True, size=(700, 450),
                     resizable=True)


def fazer_janela3():
    layout_col1 = [
        [Sg.Text('Avaliando Ordens de Serviços', font=('arial', 18))],
        [Sg.Image('Image/gif.gif', key='gif')],
        [Sg.Text('Conectando...', key='texto_avaliador', font=('arial', 10))],
        [Sg.Button('Cancelar', size=(10,1), font=('arial', 10), button_color=cor_botao, key='sair')],
        [Sg.Text('')],
        [Sg.Output(size=(60, 10))]
    ]

    return Sg.Window('Avaliador', layout_col1, finalize=True, margins=(30, 30),
                     element_justification='c')


def avaliar(aval):
    lista_os = aval.extrair_os(values['-arquivo-'])
    for atual, os in enumerate(lista_os):
        stats, rein, model, defect, analise, conclusao = 'N/F', 'N/F', 'N/F', 'N/F', 'N/F', 'N/F'
        if finalizar_processo:
            break
        try:
            print('------------------------------------------------------------')
            print(f'Ordem de Serviço:{os} ')
            window3['texto_avaliador'].update(f'Acessando Assist...')
            aval.acessar_assist()

            window3['texto_avaliador'].update(f'Acessando OS: {os}...')
            aval.acessar_os(os)

            window3['texto_avaliador'].update(f'Verificando dados da Ordem de Serviço...')
            stats, rein = aval.verificar_status()

            print(f'Status: {stats}')
            model = aval.verificar_modelo_do_produto()
            print(f'Modelo: {model}')

            window3['texto_avaliador'].update(f'Analisando defeito...')
            defect = aval.verificar_defeito_reclamado()
            print(f'Defeito: {defect}')

            analise = aval.analise_do_defeito()
            print(f'Ordem de Serviço vai ser avaliada como {analise}')

            window3['texto_avaliador'].update(f'Realizando avaliação...')
            if stats == 'PRODUTO RECEBIDO EM PROCESSO DE VALIDAÇÃO':
                aval.apertar_botao_para_avaliar()
                aval.set_pendencia_validacao()
                aval.acessar_os(os)

            if not stats in ['PENDÊNCIA DE AVALIAÇÃO TÉCNICA', 'PRODUTO RECEBIDO EM PROCESSO DE VALIDAÇÃO']:
                raise RuntimeError('Não está em pendencia de avaliação')


            aval.set_avaliacao()
            print('Avaliado com Sucesso!!')
            conclusao = 'Avaliado com Sucesso!!'

        except Exception as err:
            conclusao = err
            print(conclusao)

        print(f'\nProgresso... {atual+1}/{len(lista_os)}')
        relatorio_os.append(os)
        relatorio_stats.append(stats.upper())
        relatorio_reincidencia.append(rein.upper())
        relatorio_model.append(model.upper())
        relatorio_defect.append(defect.upper())
        relatorio_analise.append(analise.upper())
        relatorio_conclusao.append(conclusao)

    print('\nProcesso Finalizado!')
    relatorio_final = {'OS': relatorio_os, 'Modelo': relatorio_model, 'Status': relatorio_stats,
                       'Reincidência:': relatorio_reincidencia, 'Defeito Reclamado': relatorio_defect,
                       'Analise do defeito': relatorio_analise, 'Conclusão': relatorio_conclusao}

    relatorio_excel = pd.DataFrame(relatorio_final)
    if not system.path.isdir('relatorios'):
        system.mkdir('relatorios')

    relatorio_excel.to_excel(f'relatorios/Relatorio_{data_de_hoje}.xlsx', index=False)
    window3['texto_avaliador'].update(f'Processo Finalizado!', font=('arial', 15))
    window3['sair'].update('Sair')

login, senha, salvar = verificar_senha_salva()
window1, window2, window3 = fazer_janela1(), None, None

if salvar == 'True':
    window1['login'].update(login)
    window1['senha'].update(senha)
x = False

while True:
    window, event, values = Sg.read_all_windows(timeout=60)

    if event == 'Cancelar':
        break

    if window3 != None:
        window3['gif'].update_animation('Image/gif.gif', time_between_frames=60)

    if event == 'sair':
        window3.close()
        break

    if event == 'avaliar':
        if values['-arquivo-'] == '':
            Sg.popup_ok('Por favor, selecione um arquivo!', button_color=cor_botao, title='Erro')
        elif '.txt' not in values['-arquivo-']:
            Sg.popup_ok('Arquivo escolhido não está correto!', button_color=cor_botao, title='Erro')
        else:
            avaliador.definir_tecnico(values['tecnico'])
            avaliador.set_saldo(values['saldo'])
            window2.hide()
            window3 = fazer_janela3()
            avalia = Thread(target=avaliar, args=(avaliador, ))
            avalia.start()

    if event == 'Login':
        if values['login'] == '' or values['senha'] == '':
            Sg.PopupOK('Usuário Inválido', button_color=cor_botao)
            window1['senha'].update('')
            window1['login'].update('')

        else:
            username = values['login'].strip()
            password = values['senha'].strip()
            avaliador = assist.ASSIST(headless=values['ver_navegador'])
            avaliador.acessar_assist()
            try:
                if values['salvar_senha'] == True:
                    salvar_senha(username, password, 'True')
                else:
                    limpar_senha()
                avaliador.login(username, password)
                window1.hide()
                window2 = fazer_janela2()
            except:
                avaliador.close()
                Sg.PopupOK('Usuário ou senha inválido', button_color=cor_botao, title='Erro', )
                window1['senha'].update('')
                window1['login'].update('')

    if event == Sg.WIN_CLOSED:
        if window3 != None:
            window3.close()
        break

if avaliador != '':
    avaliador.close()

window1.close()
finalizar_processo = True
relatorio = f'relatorios\Relatorio_{data_de_hoje}.xlsx'
system.system(relatorio)
