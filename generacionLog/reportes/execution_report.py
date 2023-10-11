
from hooks.Report import Report
from hooks.Email import Email
class execution_report(Report):
    def __init__(self, config):
        super(execution_report, self).__init__(config)
        self.config = config
        self.configuration = self.get_data_json(config)
        self.log = str(self.configuration.log)
        self.results = {}
        
    #Should process data execution of each brand
    def proccess_data(self, query_reults,query_score, date_report):
        if "datos" in query_reults["dinBody"]:
            execution_records = query_reults["dinBody"]["datos"]
            for execution_record in execution_records:
                
                if int(execution_record["registros"]) == 0:
                    observation = "Se recomienda hacer un proceso Manual"
                    subject = "Notificación de Proceso Manual"
                    content ="Se recomienda realizar un proceso manual para la marca: "+execution_record["marca"]+" "
                    Email(self.config).sender_email(subject, content)
                    self.chance_status(execution_record["marca"], "desactivate")
                    
                else:
                    observation = query_reults["dinError"]["mensaje"]
                    self.chance_status(execution_record["marca"], "activate")
                    
                execution_record["fecha_ejecucion"] = str(query_score.date_search)
                execution_record["hour_init"] = query_score.hour_init
                execution_record["hour_end"] = query_score.hour_end
                execution_record["hour_execute"] = query_score.hour_execute
                execution_record["year"] = date_report.year
                execution_record["month"] = date_report.month_name
                execution_record["observations"] = observation


        self.results = {
            "path": self.ruta_final,
            "data": execution_records,
            "sheet_name": "report",
            "cell_init_write": "A"
        }
            
                