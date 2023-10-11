import xlwings as xw
import sys
import helpers

def chance_status_all(params):
    path_config_brand = params["path"] 
    status = params["status"]
    path_config = helpers.clear_folder_path(path_config_brand)
    sheet_name = params["sheet_name"]
    range_activate = params["range_activated"]
    cell_init = params["cell_init_write"]
    
    try:
        with xw.App(visible=False) as app:
            with app.books.open(r''+path_config, editable=True) as book:
                sheet_brand = book.sheets[sheet_name]
                sheet_records = sheet_brand.range(range_activate).value
                for index, value_cell in enumerate(sheet_records):
                    cell_brand = ""
                    cell_brand = str(cell_init)+str(index+2)
                    if sheet_brand[cell_brand].value:
                        sheet_brand[cell_brand].value = status
                        value_cell + " "  
                book.save()
                book.close()
            
    except IOError as error:
        except_info = sys.exc_info()
        s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
        helpers.put_log(s_message,"--","manager_brand", "change_status.txt")
        print('Giskard: ', 'existe un error al cambiar el status en el manager_brand')