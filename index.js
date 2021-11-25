const xl = require('excel4node');
const Holidays = require('date-holidays')
const holidaysCalculator = new Holidays()

var year = 2022;
var month = 0;
var country = "IT";
var locale = "it-IT";

holidaysCalculator.init(country, "public|bank|observance");

var currentDate = new Date(year, month, 1, 0, 0, 0, 0);
var workBook = new xl.Workbook();
var currentWorkSheet = workBook.addWorksheet(getMonthName(currentDate, locale));
const holidayStyle = workBook.createStyle({
    font: {
        color: '#FF0000',
        size: 12,
    }
});

do{
    var holidayInfo = holidaysCalculator.isHoliday(currentDate);
    
    currentWorkSheet.cell(currentDate.getDate(), 1).number(currentDate.getDate());
    currentWorkSheet.cell(currentDate.getDate(), 2).string(getWeekDayName(currentDate, locale));

    if(holidayInfo || isSundayOrSaturday(currentDate))
    {   
        currentWorkSheet.cell(currentDate.getDate(), 1).style(holidayStyle);
        currentWorkSheet.cell(currentDate.getDate(), 2).style(holidayStyle);
        currentWorkSheet.cell(currentDate.getDate(), 3)
                    .string(holidayInfo ? holidayInfo[0].name : '') //actually holdayInfo is an array with one element
                    .style(holidayStyle);
    }

    if(currentDate.getMonth() != month) {
        month = currentDate.getMonth();
        currentWorkSheet = workBook.addWorksheet(getMonthName(currentDate, locale));
    }
    currentDate = new Date(currentDate.getTime()+1000*3600*24);
}
while(currentDate.getFullYear() == year);

workBook.write(`Calendar-${country}-${year}.xlsx`);


function getMonthName(date, locale) {
    var monthName = date.toLocaleString(locale, { month: "long" });
    return monthName;
}

function getWeekDayName(date, locale) {
    var weekdayName = date.toLocaleString(locale, { weekday: "long" }).toUpperCase().substring(0,3);
    return weekdayName;
}

function isSundayOrSaturday(date) {
    return date.getDay() == 0 || date.getDay() == 6;
}