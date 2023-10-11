from HoockUtilities import hoock_utilities
import requests
import socket
import uuid
import time
import sys
from hooks.Email import Email


class ConnectApi(hoock_utilities):
    def __init__(self,config, query_parms):
        super().__init__(config)
        self.configuration = self.get_data_json(config)
        self.log = str(self.configuration.log)
        self.results = {}
        self.params_query_score = self.map_query(query_parms)
        self.senderEmail = Email(config)

        
    #Should map query
    def map_query(self, query_parms):
        query_parms = self.dictToObject(query_parms)
        params_query_score = {
            "fechaConsulta": query_parms.date_search,
            "horaInicio": query_parms.hour_init,
            "horaFin": query_parms.hour_end
        }
        return params_query_score
        
    #Should connect with microservice send body       
    def connect(self, end_point):
        params_query_score = self.params_query_score
        body_query = self.get_body(params_query_score)
        self.results = {}
        try:
            self.put_log(body_query,"Envio de parametros: "+str(end_point),"ConnectApi", self.log+"/connectApi.txt")
            response = requests.post(end_point, json=body_query)
            code_status = response.status_code
            if int(code_status) == 200:
                self.results = response.json()
                s_message = self.results
                self.put_log(s_message,"--","ConnectApi", self.log+"/connectApi.txt")
            else:
                email_sender = {
                        "subject": "Fallo en generación de información  Generar Información Manualmente desde la WEB",
                        "content": "<p>Error al generar información. Intentar generar la información manualmente desde la WEB. </p> <p>Comuníquese con el Área de Riesgos para notificar la novedad y que aseguren que los servicios estén operativos.</p> "
                    }
                self.senderEmail.sender_email(email_sender["subject"], email_sender["content"])
                self.results = {}
                
        except IOError as error:
            except_info = sys.exc_info()
            s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
            self.put_log(s_message,"--","ConnectApi", self.log+"/connectApi.txt")

    #Should return a structure of data  used in microservice
    def get_body(self, params_query):
        uuid_app = self.get_uuid()
        ip_host = self.get_ip()
        date_query = self.get_date()
        body_data = {
            "dinHeader": {
                "aplicacionId": "RPA",
                "canalId": "RPA",
                "uuid": uuid_app,
                "ip": ip_host,
                "horaTransaccion": date_query
            },
            "dinBody": params_query
        }
        s_message = body_data
        self.put_log(s_message,"--","ConnectApi", self.log+"/connectApi.txt")
        return body_data
    #Should get a uuid code
    def get_uuid(self):
        my_uuid = uuid.uuid4()
        return str(my_uuid)

    #Should ip del equipo
    def get_ip(self):
        host_name = socket.gethostname()
        ip = socket.gethostbyname(host_name)
        return ip

    #Should get la fecha de creacion
    def get_date(self):
        now = time.localtime()
        T_stamp = time.strftime("%Y-%m-%d %H:%M:%S", now) 
        return str(T_stamp)