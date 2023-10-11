from api.ConnectApi import ConnectApi
from hooks.Hour import Hour
from hooks.Report import  Report
from hooks.Email import Email
import sys

try:
    if(len(sys.argv)>1):
        config = str(sys.argv[1])
        end_point = str(sys.argv[2])
        minute_load_data = int(sys.argv[3])

        date_object = Hour(config)
        report_object = Report(config)
        email_object = Email(config)
        date = date_object.get_date_calc()
        hour = date_object.get_execute_time()
        
        data_report = []
        

        ruta_final = report_object.copy_template(config)
        query_score = date_object.get_diference_hour(minute_load_data, hour.hour_init, date.date_calc)
        date_report = date_object.get_date_report()


        params_query_score = {
            "date_search": query_score.date_search,
            "hour_init": query_score.hour_init,
            "hour_end": query_score.hour_end
        }
      
        
            
        

        conecct_object = ConnectApi(config, params_query_score)
        results = conecct_object.connect(end_point)
        
        transaction_data = conecct_object.results
        
        if len(transaction_data):
            
            isNullTransaction = transaction_data["dinBody"] is None
            
            if not isNullTransaction and  "datos" in transaction_data["dinBody"]:
                execution_records = transaction_data["dinBody"]["datos"]
                
                data_report = []
                
                for execution_record in execution_records:
                    
                    if int(execution_record["registros"]) == 0 and  not (execution_record["marca"] == "TOTAL"):
                        observation = "Se recomienda hacer un proceso Manual"
                        email_sender = {
                            "subject": "Notificación de Proceso Manual",
                            "content": "Se recomienda realizar un proceso manual para la marca: "+execution_record["marca"]+" "
                        }
                        email_object.sender_email(email_sender["subject"], email_sender["content"])
                        report_object.chance_status(execution_record["marca"], "desactivate")
                        observation = "Sin Registros"
                        print('RPA-GENERACIÓN DE LOG SCORE: Envió de notificación: ', email_sender["content"])
                    else:
                        observation = "Satisfactorio"
                        report_object.chance_status(execution_record["marca"], "activate")
                        print('RPA-GENERACIÓN DE LOG SCORE, Marca : ', execution_record["marca"], " , registrar en el reporte")
                    
                    #Depuracion la marca Total no pude entrar y la 
                    if not (execution_record["marca"] == "TOTAL"):
                        execution_record["fecha_ejecucion"] = str(query_score.date_search)
                        execution_record["hour_init"] = query_score.hour_init
                        execution_record["hour_end"] = query_score.hour_end
                        execution_record["hour_execute"] = query_score.hour_execute
                        execution_record["year"] = date_report.year
                        execution_record["month"] = date_report.month_name
                        execution_record["observations"] = observation
                        data_report.append(execution_record)


                    
                    
                params_report = {
                    "path": ruta_final,
                    "data": data_report,
                    "sheet_name": "report",
                    "cell_init_write": "A"
                }


                report_object.create_report(params_report)
                print('RPA-GENERACIÓN DE LOG SCORE: Registrando Datos en el Reporte de ejecución: ', str(params_report))
                
                
        
            else:
                report_object.chance_status_all("desactivate")
                print('RPA-GENERACIÓN DE LOG SCORE: Envió de notificación: ')
                print('RPA-GENERACIÓN DE LOG SCORE: Se cambió el status del archivo de ejecuión de macro ')
        else:
            report_object.chance_status_all("desactivate")
            print('RPA-GENERACIÓN DE LOG SCORE: Se cambió el status del archivo de ejecuión de macro ')                    
            
    
    
        

except IOError as error:
    except_info = sys.exc_info()
    s_message = f'({except_info[2].tb_lineno}) {except_info[0]} {str(error)}'
    report_object.helpers.put_log(s_message,"--","main",  str(report_object.log)+"/main.txt")
