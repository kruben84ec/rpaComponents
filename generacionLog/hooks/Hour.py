from HoockUtilities import hoock_utilities
from datetime import datetime


class Hour(hoock_utilities):
    def __init__(self, config):
        super(Hour, self).__init__(config)
        self.configuration = self.get_data_json(config)
        self.log = str(self.configuration.log)
        self.results = {}
        self.date_report = ""
        
    #Should initilize query parms
    def get_query_params(self, minute_load_data):
        hour_query = self.get_execute_time()
        date_query = self.get_date_calc()
        self.results = self.get_diference_hour(minute_load_data, 
                                               hour_query.hour_init, 
                                               date_query.date_calc
                        )
        self.date_report = self.get_date_report()


    #Should to obtain params necesary that used in query microservice
    def get_diference_hour(self,minute_load_data, hour_init, date_calc):
        results = {}
        hour_now = self.get_hour()
        now_time = hour_init.strip().split(".")

        isLastTurn = (hour_init.strip().split(".")[0] == "00")  and (int(hour_init.strip().split(".")[1]) < 30)

        if isLastTurn:
            now_time[0] = 00
            now_time[1] = 00
            hour_init="24.00"

        #Should get Parameters to use in the Fraud API   
        time_search = self.get_time_search(date_calc, minute_load_data, now_time)

        hour_ = time_search["hour_minor"].replace(":", ".").split(".")

        
        
        results = {
            "hour_init": str(hour_[0]+"."+hour_[1]),
            "hour_end": str(hour_init.strip()),
            "date_search": str(time_search["day_minor"]),
            "hour_execute": str(hour_now)
        }
        
        return self.dictToObject(results)

    #Should rto response with hour of system
    def get_hour(self):
        now_system = datetime.now()
        hour_now = now_system.strftime("%H:%M:%S")
        return hour_now

    #Should  to obtain hour params to find in used microservice
    def get_time_search(self, date_calc, minute_load_data, now_time):
        year = int(date_calc.year)
        mouth = int(date_calc.month)
        day = int(date_calc.day)

        date_complet_now = datetime(year, mouth, day, int(now_time[0]), int(now_time[1]), 00)
        time_search = self.rest_minute(date_complet_now, minute_load_data)
        return time_search

    #Should return time and date
    def get_execute_time(self):
        time_search = datetime.now()
        #Delay de disponibilidad de información en la base de datos de Riesgos
        minutes = 4
        time_init = self.rest_minute(time_search, minutes)
        hour_time = time_init["hour_minor"].replace(":", ".").split(".")
        date_execute_time = {
            "hour_init":hour_time[0]+"."+hour_time[1]
        }
        return self.dictToObject(date_execute_time)
        
    #Should return date_calc
    def get_date_calc(self):
        date_calc = {
            "date_calc": datetime.now()
        }
        return self.dictToObject(date_calc)


    def get_name_mouth(self, number_mouth):
        mouth = [
            '', "ENERO", "FEBRERO", "MARZO",
            'ABRIL', "MAYO", "JUNIO", "JULIO",
            'AGOSTO', "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE",
            'DICIEMBRE'
        ]
        return mouth[int(number_mouth)]
        
    def get_date_report(self):
        now_system = datetime.now()
        year = now_system.year
        mouth = self.get_name_mouth(now_system.month)
        date_report = {
            "year": year,
            "month_name": mouth
        }
        return self.dictToObject(date_report)