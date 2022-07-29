from threading import Thread
from gtts import gTTS
from decouple import config
import PySimpleGUI as Sg
import playsound
import os
import get_os


def baixar_relatorio(download_window):
    global baixou_relatorio
    global ordens
    global erro
    global erro_msg

    report_file = get_os.get_report_name()
    get_os.erase_previous_report()
    for attempt in range(3):
        try:
            get_os.download_report(config('USER'), config('PASSWORD'))
            report = get_os.ReportFilter(report_file)
            ordens = report.defect_tolist()
            baixou_relatorio = True
            break
        except Exception as err:
            if attempt == 2:
                erro = True
                erro_msg = err
                raise FileNotFoundError('Não foi possível fazer o download do relatório, reinicie a aplicação.')

            download_window['texto_download'].update(f'Erro ao baixar o relatório.\n  '
                                                     f' Realizando tentativa {attempt + 2}')
            print(f'Erro ao baixar o relatório, realizando tentativa {attempt + 2}')
            continue

def falar(texto):
    try:
        os.remove('fala.mp3')
    except Exception as err:
        print(f'Erro ao tentar remover arquivo mp3: {err}')
    sintese = gTTS(texto, lang='pt', slow=False)
    arq = os.path.join(os.getcwd(), 'fala.mp3')
    sintese.save(arq)
    playsound.playsound('fala.mp3')

def fazer_janela_main():
    Sg.theme('Reddit')
    layout = [
        [Sg.Button('Automático', size=(30, 4), font=25)],
        [Sg.Text(' ')],
        [Sg.Button('Personalizado', size=(30, 4), font=25)]
    ]
    return Sg.Window('Localizador de Status', layout, margins=(40, 40), finalize=True, element_justification='c')

def fazer_janela():
    Sg.theme('Reddit')
    col1 = [
        [Sg.Multiline(key='lista1', size=(30, 2))]
    ]

    col2 = [
        [Sg.Multiline(key='voz1', size=(30, 2))]
    ]

    layout = [
        [Sg.Frame('Lista de OS', layout=col1, key='cont1'),
         Sg.Frame('Frase que será falada', layout=col2, key='cont2')],
        [Sg.Button('Adicionar'), Sg.Button('Reiniciar'),
        Sg.Button('Ok'), Sg.Button('Voltar')]
    ]

    return Sg.Window('Localizador de Status', layout, margins=(20, 20), finalize=True, element_justification='c')


def fazer_janela2():
    Sg.theme('Reddit')
    layout = [
        [Sg.Text('Digite a OS:')],
        [Sg.InputText(key='ordem_servico')],
        [Sg.Button('Voltar')],
        [Sg.Button('Submit', visible=False, bind_return_key=True)]
    ]

    return Sg.Window('Localizador de Status', layout, finalize=True, margins=(50, 30))

def fazer_janela_downloading():
    Sg.theme('Reddit')
    layout = [
        [Sg.Image('Image/gif.gif', key='gif')],
        [Sg.Text('Fazendo o download do relatório', key='texto_download', font=('arial', 13))]
    ]

    return Sg.Window('Baixando relatorio',layout, finalize=True,
                    size=(400, 300), margins=(10, 60), element_justification='c')

def main():
    global ordens
    global baixou_relatorio
    global erro
    global erro_msg


    erro = False
    auto = False
    n = 1
    lista_os = {}
    vozes = {}
    lista_ordens = None
    baixou_relatorio = False
    window_menu, window, window2, window_download = fazer_janela_main(), None, None, None

    while True:
        win, event, values = Sg.read_all_windows(timeout=60)

        if erro:
            Sg.PopupError(f'Erro ao baixar o relatório, reinicie a aplicação\n\n{erro_msg}')
            break

        if window_download != None:
            window_download['gif'].update_animation('Image/gif.gif', time_between_frames=60)

        ###########################################################
        ########### Checks buttons from menu window ###############
        ###########################################################

        if event == 'Automático':
            baixou_relatorio = False
            auto = True
            window_menu.hide()
            window_download = fazer_janela_downloading()
            relatorio = Thread(target=baixar_relatorio, args=(window_download,))
            relatorio.start()

        if baixou_relatorio:
            lista_ordens = ordens
            baixou_relatorio = False
            window_download.close()
            if window2 is not None:
                window2.close()
            window2 = fazer_janela2()

        if event == 'Personalizado':
            window_menu.hide()

            if window is not None:
                window.close()
            window = fazer_janela()

        ###########################################################
        ############ Checks buttons from  window 1 ################
        ###########################################################

        if event == 'Adicionar':
            window.extend_layout(window['cont1'], [[Sg.Multiline(key=f'lista{str(n + 1)}', size=(30, 2))]])
            window.extend_layout(window['cont2'], [[Sg.Multiline(key=f'voz{str(n + 1)}', size=(30, 2))]])
            n += 1

        if event == 'Ok':
            auto = False
            for element in range(1, n+1):
                lista = values[f'lista{element}'].split('\n')
                voz = values[f'voz{element}']

                lista_os[f'lista{element}'] = lista
                vozes[f'voz{element}'] = voz

            window.hide()
            if window2 is not None:
                window2.close()
            window2 = fazer_janela2()

        if event == 'Reiniciar':
            n = 1
            window.close()
            window = fazer_janela()

        if event == 'Voltar' and win == window:
            window.close()
            window = None
            window_menu.un_hide()

        ###########################################################
        ############ Checks buttons from  window 2 ################
        ###########################################################

        if event == 'Voltar' and win == window2:
            window2.close()
            window2 = None
            window_menu.un_hide()
            lista_ordens = None
            baixou_relatorio = False

        if event == 'Submit':
            if auto:
                palavra_foi_achada = False
                status = lista_ordens.keys()
                for stat in status:
                    if values['ordem_servico'] in lista_ordens[stat]:
                        falar(stat)
                        palavra_foi_achada = True
                        break
                if not palavra_foi_achada:
                    falar('Não encontrado')

                window2['ordem_servico'].update('')
            else:
                palavra_foi_achada = False
                for numero in range(1, n+1):
                    if values['ordem_servico'] in lista_os[f'lista{numero}']:
                        falar(vozes[f'voz{numero}'])
                        palavra_foi_achada = True
                        break
                if not palavra_foi_achada:
                    falar('Não encontrado')

                window2['ordem_servico'].update('')

        if event == Sg.WINDOW_CLOSED:
            break

if __name__ == '__main__':
    main()
