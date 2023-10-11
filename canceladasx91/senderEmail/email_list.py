import helpers

#Debe obtener la lis de correos
def get_list_email():
    path_json_list_email = "E:\\auxiliares\\config\\email_list.json"
    path_json = helpers.clear_folder_path(path_json_list_email)
    email_list = helpers.read_json(path_json)
    email_receive = email_list["listEmail"]
    return email_receive