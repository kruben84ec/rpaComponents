import xlwings as xw
from xlwings import Range, constants
from HoockUtilities import hoock_utilities 
import sys


class Report(hoock_utilities):
    def __init__(self, config):
        super(Report, self).__init__(config)
        self.configuration = self.get_data_json(config)
        self.log = str(self.configuration.log)
        self.results = {}
        self.ruta_final = self.copy_template(config)

    #Should read config from json
    def read_config(self, config_path):
        config = self.read_json(config_path)
        return config

    #Should copy File path principal to move other path
    def copy_template(self, path_config):
        try:
            date_execute = self.get_date_complete()
            path_config = self.configuration

            if path_config:
                template_report = str(path_config.principal)+"reporte_ejecucion.xlsx"
                path_final = str(path_config.ruta_reporte)+"reporte_ejecucion_"+date_execute+".xlsx"
                self.copy_file(template_report,path_final)
                return path_final
        except IOError as error:
            except_info = sys.exc_info()
            s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
            self.put_log(s_message,"--","Report", self.log+"/Report.txt")


    #Should change status of activate
    def create_report(self, params):
        config_report = self.dictToObject(params)
        
        path_config_brand =  config_report.path
        sheet_name = config_report.sheet_name
        path_config = self.clear_folder_path(path_config_brand)
        
        try:
            with xw.App(visible=False) as app:
                book = app.books.open(r''+path_config, editable=True)
                sheet_book = self.get_sheet_name(book, sheet_name)
                last_row = self.get_last_row_book(book, sheet_book)
                next_row_write = int(last_row+1)
                
                for index, record_data in enumerate(config_report.data, next_row_write):
                    position = str(index)
                    self.write_row(sheet_book,"A"+position, record_data.fecha_ejecucion)
                    self.write_row(sheet_book,"B"+str(index),record_data.month)
                    self.write_row(sheet_book,"C"+str(index), record_data.year)                               
                    self.write_row(sheet_book,"D"+str(index),record_data.marca)
                    self.write_row(sheet_book,"E"+str(index), record_data.hour_init)
                    self.write_row(sheet_book,"F"+str(index),record_data.hour_end)
                    self.write_row(sheet_book,"G"+str(index),record_data.hour_execute)
                    self.write_row(sheet_book,"H"+str(index), record_data.registros)
                    self.write_row(sheet_book,"I"+str(index), record_data.observations)       
                        
                self.save_report(book)
                self.close_report(book)
                
        except IOError as error:
            except_info = sys.exc_info()
            s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
            self.put_log(s_message,"--","Report", self.log+"/Report.txt")


    def get_last_row_book(self, book, sheet_book):
        ultima_fila = sheet_book.range('A' + str(book.sheets[0].cells.last_cell.row)).end('up').row
        return int(ultima_fila)

    def get_sheet_name(self,book, sheet_name):
        return book.sheets[sheet_name]
    
    def save_report(self, book):
        book.save()

    def close_report(self, book):
        book.close()
        
    def write_row(self, sheet_book, position, value_data):
        sheet_book[position].value = value_data

    def read_cell(self, sheet_book, position):
        return sheet_book[position].value

    #Cambio de estatus

    def chance_status(self, brand_search:str, status:str):
        try:
            path_config = self.configuration
            template_report = str(path_config.principal)+"configuracion.xlsx"
            path_config_file = self.clear_folder_path(template_report)
            
            with xw.App(visible=False) as app:
                book = app.books.open(r''+path_config_file, editable=True)
                sheet_book = self.get_sheet_name(book, "marcas")
                last_row = self.get_last_row_book(book, sheet_book)
                next_row_write = int(last_row+1)
                
                for index in range(2, next_row_write):
                    position = "D"+str(index)
                    brand = self.read_cell(sheet_book, position)
                    if brand == brand_search.upper():
                        position = "B"+str(index)    
                        self.write_row(sheet_book, position, status)
                        
                self.save_report(book)
                self.close_report(book)
        except IOError as error:
            except_info = sys.exc_info()
            s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
            self.put_log(s_message,"--","Report", self.log+"/Report.txt")

    def chance_status_all(self, status:str):
        try:
            path_config = self.configuration
            template_report = str(path_config.principal)+"configuracion.xlsx"
            path_config_file = self.clear_folder_path(template_report)
            
            with xw.App(visible=False) as app:
                book = app.books.open(r''+path_config_file, editable=True)
                sheet_book = self.get_sheet_name(book, "marcas")
                last_row = self.get_last_row_book(book, sheet_book)
                next_row_write = int(last_row+1)
                
                for index in range(2, next_row_write):
                    position = "B"+str(index)    
                    self.write_row(sheet_book, position, status)
                        
                self.save_report(book)
                self.close_report(book)
        except IOError as error:
            except_info = sys.exc_info()
            s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
            self.put_log(s_message,"--","Report", self.log+"/Report.txt")
