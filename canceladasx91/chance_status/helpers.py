import os
import sys
import shutil
import win32com.client as win
from datetime import date, datetime, timedelta
import time, json

#Debe crear una copia de archivos
def copy_file(origin_path:str, destiny_path: str):
    exist_file = os.path.isfile(destiny_path)
    print('RPA_LOG_SCORE: Copiando:  ', origin_path, destiny_path)
    if not exist_file:
        shutil.copy(origin_path, destiny_path)


#Debe borrar el archivo
def delete_file(myfile: str):
    try:
        os.remove(myfile)
    except OSError as e:  ## if failed, report it back to the user ##
        s_message = str(e.filename) + str(e.strerror)
        put_log(s_message,"---","senderEmail/helpers/delete_file")


#Debe poder formatear los tiempos de minuitos y segundos de manera  01.
def format_time(minute):
    if minute < 10:
        minute = "0"+str(minute)
    return minute


#Deber crear un archivo de Log para los proceseos de automatización
def put_log(mensaje:str, marca:str, script:str, pat_log:str ="senderEmail.txt"):
    ruta = "E:\\Canceladasx91\\log_bot\\"+pat_log
    with open(ruta, 'a') as file:
        file.write(f'{datetime.now()};Script - {script}.py;{mensaje};Marca: {marca}\n')
        file.close()

#Debe limpiar la ruta que ingresa para sistemas operativos Windows
#Debe limpiar la ruta que ingresa para sistemas operativos Windows
def clear_folder_path(path: str):
    try:
        path = path.replace('\\\\', '/')
        path = path.replace('\\', '/')
        if path[-1] == '/':
            path = path[:-1]
        return path
    except IOError  as error:
        except_info = sys.exc_info()
        s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
        put_log(s_message,"--","senderEmail/helpers")


def hide_columns_from_vbscript(path):
    vbs = win.Dispatch("ScriptControl")
    vbs.Language = "vbscript"
    scode = """
    Function hideColumns(path)
        Set objExcel = CreateObject("Excel.Application")
        Set objWorkbook = objExcel.Workbooks.Open(path)
        Set objSheet = objExcel.ActiveWorkbook.Worksheets("Hoja2")
        objSheet.Columns("O:R").EntireColumn.Hidden = True
        objWorkbook.Save
        objWorkbook.Close
        objExcel.Quit
        End Function
    """
    vbs.AddCode(scode)
    vbs.Run("hideColumns", path)
    
#Funcion debe leer un archivo de json
def read_json(path_json:str):
    try:
        data = []
        with open(path_json) as json_file:
            data = json.load(json_file)
    
        return data
    except IOError as error:
        put_log(error, "Lectura", "senderEmail:helper.py")
    
    
    
def validate_date(date_text):
    try:
        datetime.strptime(date_text, '%Y-%m-%d')
        return True
    except ValueError:
        return False
 
def rest_two_date(date_load: str, date_input: str):
    d1 = datetime.strptime(date_load, "%Y-%m-%d")
    d2 = datetime.strptime(date_input, "%Y-%m-%d")
    return ((d2 - d1).days)
    
def validate_diferent(date_input):
    now_system = str(date.today())
    days_diferent = rest_two_date(date_input, now_system)
    is_valid = True
    
    if not (days_diferent >=0  and days_diferent <2):
        is_valid = False
                
    return is_valid


def rest_minute(date_complet_now, minor_minute):
    time = date_complet_now - timedelta(hours=0, minutes=minor_minute)
    hour_minor = time.strftime('%H:%M:%S')
    day_minor = time.strftime('%Y-%m-%d')
    
    return {
        "day_minor": day_minor,
        "hour_minor": hour_minor
    }
