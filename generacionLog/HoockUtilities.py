from datetime import date, datetime, timedelta
from munch import DefaultMunch
import os
import sys
import shutil
import json

class hoock_utilities():
    def __init__(self, config):
        self.config = config
        
    #Debe crear una copia de archivos
    def copy_file(self, origin_path:str, destiny_path: str):
        exist_file = os.path.isfile(destiny_path)
        if not exist_file:
            shutil.copy(origin_path, destiny_path)

    #Debe borrar el archivo
    def delete_file(self, myfile: str):
        try:
            os.remove(myfile)
        except OSError as e:
            s_message = str(e.filename) + str(e.strerror)
            self.put_log(s_message,"---","delete_file")

    #Debe poder formatear los tiempos de minuitos y segundos de manera  01.
    def format_time(self, minute):
        if minute < 10:
            minute = "0"+str(minute)
        return minute


    #Deber crear un archivo de Log para los proceseos de automatización
    def put_log(self, mensaje:str, marca:str, script:str, pat_log:str ="senderEmail.txt"):
        with open(pat_log, 'a') as file:
            file.write(f'{datetime.now()};Script - {script}.py;{mensaje};Marca: {marca}\n')
            file.close()

    #Debe limpiar la ruta que ingresa para sistemas operativos Windows
    def clear_folder_path(self, path: str):
        try:
            path = path.replace('\\\\', '/')
            path = path.replace('\\', '/')
            if path[-1] == '/':
                path = path[:-1]
            return path
        except IOError  as error:
            except_info = sys.exc_info()
            s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
            self.put_log(s_message,"--","clear_folder")
            
        
    #Funcion debe leer un archivo de json
    def read_json(self, path_json:str):
        try:
            data = []
            path_clear = self.clear_folder_path(path_json)
            with open(path_clear) as json_file:
                data = json.load(json_file)
        
            return data
        except IOError as error:
            self.put_log(error, "Lectura", "read_json")
    
    def get_data_json(self, path_json):
        data_json = self.read_json(path_json)
        return self.dictToObject(data_json) 
    
    def validate_date(self, date_text):
        try:
            datetime.strptime(date_text, '%Y-%m-%d')
            return True
        except ValueError:
            return False
    
    def rest_two_date(self, date_load: str, date_input: str):
        d1 = datetime.strptime(date_load, "%Y-%m-%d")
        d2 = datetime.strptime(date_input, "%Y-%m-%d")
        return ((d2 - d1).days)
        
    def validate_diferent(self, date_input):
        now_system = str(date.today())
        days_diferent = self.rest_two_date(date_input, now_system)
        is_valid = True

        if not (days_diferent >=0  and days_diferent <2):
            is_valid = False
        return is_valid

    def rest_minute(self, date_complete_now, minor_minute):
        time = date_complete_now - timedelta(hours=0, minutes=minor_minute) 
        hour_minor = time.strftime('%H:%M:%S')
        day_minor = time.strftime('%Y-%m-%d')    
        return {
            "day_minor": day_minor,
            "hour_minor": hour_minor
        }
    
    def sustract_day(self, date_now, day_minor):
        date_sustract = date_now - timedelta(days=day_minor)
        return date_sustract.strftime('%Y-%m-%d')



    #Debe obtener la lista de parametros de configutracion
    def params_config(self, ruta_file_json):
        path_json_config = ruta_file_json
        path_json = self.clear_folder_path(path_json_config)
        config_email = self.read_json(path_json)
        return self.dictToObject(config_email)

    #Should Dict convert to in Object
    def dictToObject(self, objectParms):
        return DefaultMunch.fromDict(objectParms)
        
    #obtener la fecha
    def get_date_complete(self):
        now_system = datetime.now()
        year = now_system.year
        mouth = self.format_time(now_system.month)
        day = self.format_time(now_system.day)
        date_complete = str(year)+"-"+str(mouth)+"-"+str(day)
        return date_complete
