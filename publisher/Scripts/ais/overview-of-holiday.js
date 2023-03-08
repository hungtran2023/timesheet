var date = new Date();
var currentMonth = date.getMonth();
var currentYear = date.getFullYear();
var minQuarter = 1;
var maxQuarter = 4;
var minYear = 2000;
var maxYear = currentYear + 1;
var currentQuarter = parseInt(date.getMonth() / 3) + 1;
var quarterRang = [[1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]]
var currentQuarterRang = quarterRang[currentQuarter - 1];
var daysOffList = [];
var firstCalendar = '#OverviewCalendarFirst';
var middleCalendar = '#OverviewCalendarMiddle';
var lastCalendar = '#OverviewCalendarLast';
var loginPageUrl = $("#loginPageUrl").val();

$(document).ready(function () {
    $("#spinner-loading").show();
    GetDaysOff(currentYear);
});

$("#NextThreeMonths").click(function () {
    $("#spinner-loading").show();
    currentQuarter = currentQuarter >= maxQuarter ? minQuarter : ++currentQuarter;
    currentQuarterRang = quarterRang[currentQuarter - 1];
    if (currentQuarter == minQuarter) {
        ++currentYear;
        if (currentYear > maxYear) {
            currentQuarter = maxQuarter;
            currentYear = maxYear;
            $("#spinner-loading").hide();
            return;
        }
        GetDaysOff(currentYear);
        return;
    }
    ReloadCalendar();
})

$("#PreviousThreeMonths").click(function () {
    $("#spinner-loading").show();
    currentQuarter = currentQuarter <= minQuarter ? maxQuarter : --currentQuarter;
    currentQuarterRang = quarterRang[currentQuarter - 1];
    if (currentQuarter == maxQuarter) {
        --currentYear;
        if (currentYear < minYear) {
            currentQuarter = minQuarter;
            currentYear = minYear;
            $("#spinner-loading").hide();
            return
        }
        GetDaysOff(currentYear);
        return;
    }
    ReloadCalendar();
})

function InitDatepicker(datepickerMarkup, currentYear, currentMonth, daysOff, date) {
    $(datepickerMarkup)
       .datepicker({
           startDate: GetDate(currentYear, currentMonth - 1, 1),
           endDate: GetDate(currentYear, currentMonth, 0),
           minDate: 0,
           daysOfWeekDisabled: [0, 6],
           beforeShowDay: function (date) {
               var flag = false;
               var temp;
               daysOff.filter(function (item) {
                   if (item.day == date.getDate()) {
                       flag = true;
                       temp = item;
                   }
               });
               if (flag == true && temp.isHoliday) {
                   return { classes: 'holiday' };
               } else if (flag == true) {
                   return { classes: 'authorised' };
               }
           }
       })
}

function GetDate(year, month, date) {
    return new Date(year, month, date);
}

function DestroyDatepicker(datepickerMarkup) {
    $(datepickerMarkup).datepicker('remove');
}

function GetDaysOff(currentYear) {
    var getDaysOffUrl = $('#GetDaysOffUrl').val();
    var getDaysOffUrl = getDaysOffUrl + "/" + currentYear;
    $.ajax({
        url: getDaysOffUrl,
        type: 'GET',
        dataType: "json",
        content: "application/json;charset=utf-8",
        success: function (result) {
            if (result == "SessionExpired") {
                window.location.replace(loginPageUrl);
                return;
            }
            daysOffList = result;
            ReloadCalendar();
        },
        error: function (result) {
            daysOffList = [];
            ReloadCalendar();
        }
    })
}

function GetDaysOffByMonth(daysOffList, currentMonth) {
    var result = [];
    daysOffList.filter(function (item) {
        if (item.month == currentMonth) {
            result.push(item);
        }
    })
    return result;
}

function ReloadCalendar() {
    DestroyDatepicker(firstCalendar);
    InitDatepicker(firstCalendar, currentYear, currentQuarterRang[0], GetDaysOffByMonth(daysOffList, currentQuarterRang[0]), date);
    DestroyDatepicker(middleCalendar);
    InitDatepicker(middleCalendar, currentYear, currentQuarterRang[1], GetDaysOffByMonth(daysOffList, currentQuarterRang[1]), date);
    DestroyDatepicker(lastCalendar);
    InitDatepicker(lastCalendar, currentYear, currentQuarterRang[2], GetDaysOffByMonth(daysOffList, currentQuarterRang[2]), date);
    $("#spinner-loading").hide();
    $(".datepicker-inline").find('.datepicker-switch').attr('colspan', 12);
}