import manager_brand
import sys
import helpers


try:
    if(len(sys.argv)>1):
        status_changed = sys.argv[1]
        path_config_brand = r"E:\\Canceladasx91\\config\\marcas.xlsx"
        params_manager_brand = {
            "path": path_config_brand,
            "status": status_changed,
            "sheet_name": "marcas",
            "range_activated": "A2:H12",
            "cell_init_write": "C"
        }
        manager_brand.chance_status_all(params_manager_brand)
except IOError as error:
    except_info = sys.exc_info()
    s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
    helpers.put_log(s_message,"--","change_status", "change_status.txt")
    print('Giskard: ', 'existe un error al cambiar el status en el change_status')