import PySimpleGUI as sg
from pikepdf import Pdf, Page
import os
import zipfile
import requests
import shutil

def get_label_from_labelary(code): 
    url = 'http://api.labelary.com/v1/printers/8dpmm/labels/4x8/0/'
    files = {'file': code.decode('UTF-8')}
    headers = {
        'Accept': 'application/pdf',
        'X-Quality': 'Grayscale'        
    }
    response = requests.post(url, headers= headers, files= files, stream= True)
    if response.status_code == 200:
        response.raw.decode_content = True
        return response.raw
    else:
        pass

sg.theme("Default1")

LASER = 'EPSON L395'
TERMICA = "MFLABEL DT425A"

panel_izq = [
    [
        sg.In(size=(20, 1), enable_events=True, key='-CARPETA-'),
        sg.FileBrowse("Seleccionar ZIP",enable_events=True, size=(16, 1), file_types= (('Archivo ZIP', '*.zip'),), key="-ARCHIVO ZIP-")
    ],
    [
        sg.Text('Esperando archivo...', background_color="#770000", text_color="#FFFFFF", justification= 'center', s=(25, 1), key="-ESTADO-"),
        sg.Button("Historial", s=(11, 1) , key='-HISTORIAL-')
    ],
    [ sg.HSeparator(pad=(1, 10))],
    [
        sg.Button("Reporte", disabled= True, s=(18, 1), key='-R-'),
        sg.Button("Termica", disabled= True, s=(18, 1), key='-T-')
    ]
]

layout = [
    [
        panel_izq
    ]
]

window = sg.Window(title = "Generador de Etiquetas MELI", layout= layout, size=(300, 150), icon="EtiquetasZPL.ico")

while True:
    event, values = window.read()
    if event == '-HISTORIAL-':
        path = f'C:/EtiquetasMELI/historial/'
        os.system(f'start {os.path.realpath(path)}')

    if event == '-R-':
        with zipfile.ZipFile(values['-CARPETA-'], 'r') as zf:
            archivo = values['-CARPETA-'].split('/')[-1:][0].split('MercadoEnvios-')[1].replace(' ', '').split('.zip')[0]
            resumen = zf.extract('Control.pdf', f'C:/EtiquetasMELI/historial/{archivo}')
            os.system(f'PDFToPrinter.exe C:/EtiquetasMELI/historial/{archivo}/Control.pdf {LASER} /s')

    if event == '-T-':
        archivo = values['-CARPETA-'].split('/')[-1:][0].split('MercadoEnvios-')[1].replace(' ', '').split('.zip')[0]
        with zipfile.ZipFile(values['-CARPETA-'], 'r') as zf:
            zpl = zf.read('Etiqueta de envio.txt')
            os.makedirs(os.path.dirname(f'C:/EtiquetasMELI/historial/{archivo}/'), exist_ok= True)
            with open(f'C:/EtiquetasMELI/historial/{archivo}/etiqueta.pdf', 'wb') as out_file:
                shutil.copyfileobj(get_label_from_labelary(zpl), out_file)
            path = f'C:/EtiquetasMELI/historial/{archivo}'
            os.system(f'start C:/EtiquetasMELI/historial/{archivo}/etiqueta.pdf')
            
    if event == '-CARPETA-':
        try:
            with zipfile.ZipFile(values['-CARPETA-'], 'r') as zf:
                archivos = zf.namelist()
                if 'Etiqueta de envio.txt' in archivos and 'Control.pdf' in archivos:
                        window['-T-'].update(disabled= False)
                        window['-R-'].update(disabled= False)
                        window['-ESTADO-'].update('Archivo valido', background_color='#00CC00', text_color="#000000")
                else:
                    window['-T-'].update(disabled= True)
                    window['-R-'].update(disabled= True)
                    window['-ESTADO-'].update('El archivo no contiene etiquetas', background_color="#770000")
        except:
            pass

    if event == 'CERRAR' or event == sg.WIN_CLOSED:
        break
    

window.close


