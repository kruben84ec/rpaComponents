// //Obtener la fecha actual del sistema	

function getHourSystem  () { return new Date();}

function formatTime(minute)  {
    var minuteFormat = parseInt(minute, 10);
    if(minuteFormat <10) {
        return String("0")+String(minute)
    }
    return String(minute)
}

function getInitDate(minuteBetween) {
    //var delay_minute = 5;
    var delay_minute = 3;
    var todayInit = getHourSystem()
    var initDate = new Date(todayInit)
    var minuteStay = minuteBetween + delay_minute
    initDate.setMinutes(todayInit.getMinutes() - minuteStay)
    return initDate
}

function getSearchHour(){
    var hourseach = formatTime(getHourSystem().getHours());
    var minuteseach = formatTime(getHourSystem().getMinutes());
    return hourseach+"."+minuteseach;
}

function getInitHour(minuteBetween) {
    var initDate = getInitDate(minuteBetween);
    var initHour = formatTime(initDate.getHours());
    var initMinute = formatTime(initDate.getMinutes());
    return String(initHour)+"."+String(initMinute)
}

function getFinishHour(minuteBetween) {
    var initDate = getInitDate(minuteBetween);
    var finishDate = new Date(initDate)
    finishDate.setMinutes(initDate.getMinutes() + minuteBetween)
    var endHour = formatTime(finishDate.getHours());
    var endMinute = formatTime(finishDate.getMinutes());
    return String(endHour)+"."+String(endMinute)
}

function getInitFecha() {
    var fecha = getHourSystem();
    var dia = formatTime(fecha.getDate());
    var mes = formatTime(fecha.getMonth());
    var year = fecha.getFullYear();
    return String(year)+"-"+String(mes)+"-"+String(dia);
}

function getDateLog() {    
    var fecha = getHourSystem();

    var dia = formatTime(fecha.getDate());
    var mes_cal = fecha.getMonth()+1
    var mes = formatTime(mes_cal);
    var year = fecha.getFullYear();
    return String(year)+"-"+String(mes)+"-"+String(dia);
}

function getNameMouthLog() {    
    var fecha = getHourSystem();
    var mes_cal = fecha.getMonth()+1
    var nameMouth = ["",
        "ENERO", "FEBRERO", "MARZO",
        "ABRIL", "MAYO", "JUNIO",
        "JULIO", "AGOSTO", "SEPTIEMBRE",
        "OCTUBRE", "NOVIEMBRE", "DICIEMBRE" 
    ]
    return nameMouth[mes_cal];
}


function getYearLog() {    
    var fecha = getHourSystem();
    return String(fecha.getFullYear());
}



function getLastMount() {
    var dateSearch = new Date();
    var last3months = new Date(dateSearch.setMonth(dateSearch.getMonth()-3));
    var dateLastMonts = last3months.toISOString().split("T")[0];
    return dateLastMonts.split("-").join("");
}
