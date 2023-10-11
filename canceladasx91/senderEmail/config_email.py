import helpers
from munch import DefaultMunch

#Debe obtener la lista de parametros de configutracion
def params_config():
    path_json_config = "E:\\auxiliares\\config\\config_email.json"
    path_json = helpers.clear_folder_path(path_json_config)
    config_email = helpers.read_json(path_json)
    return DefaultMunch.fromDict(config_email)