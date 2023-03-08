var item = 0;
var currentIndex = 0;
var oldMonth, oldYear;
var rejectManyUrl = $("#RejectManyUrl").val();
var rejectUrl = $("#RejectUrl").val();
var approveManyUrl = $("#ApproveManyUrl").val();
var approveUrl = $("#ApproveUrl").val();
var teamCalendarDataUrl = $('#GetDataForTeamCalendar').val();
var loginPageUrl = $("#loginPageUrl").val();
var ActionExecute = {
    'click .authorizing__absence-requests--approve-inbox': function (e, value, row, index) {
        var id = [row.Id];
        $("#confirm-notificator__note").val("");
        $("#confirm-notificator__text").html("Are you sure you want to approve it?");
        $("#confirm-notificator").modal().one('click', "#confirm-notificator__Ok", function () {
            var note = $("#confirm-notificator__note").val();
            var data = {
                RequestIds: id,
                Note: note
            }
            DoPostForList(approveUrl, data);
        })
    },
    'click .authorizing__absence-requests--reject-inbox': function (e, value, row, index) {
        var id = [row.Id];
        $("#confirm-notificator__note").val("");
        $("#confirm-notificator__text").html("Are you sure you want to reject it?");
        $("#confirm-notificator").modal().one('click', "#confirm-notificator__Ok", function () {
            var note = $("#confirm-notificator__note").val();
            var data = {
                RequestIds: id,
                Note: note
            }
            DoPostForList(rejectUrl, data);
        })
    }
};
var Popovershow = {
    'mouseover .testpopover': function (e, value, row, index) {
        if (row.Note != null) {
            $(this).popover('show');
        }
    }
};

$(document).ready(function () {
    var currentMonth = new Date().getMonth() + 1;
    var currentYear = new Date().getFullYear();
    oldMonth = currentMonth;
    oldYear = currentYear;
    InitTeamCalendar(currentMonth, currentYear);
});

$("#authorizer-request-list").bootstrapTable({
    uniqueId: 'Id',
    formatShowingRows: function (pageFrom, pageTo, totalRows) {
        var pagesize = $("#authorizer-request-list").data("page-size");
        pagesize = parseInt(pagesize);
        var currentPage = parseInt(pageTo / pagesize);
        var totalPage = parseInt(totalRows / pagesize);
        if (totalRows == 0) {
            currentPage = 1;
            totalPage = 1;
        }
        else {
            if (totalRows % pagesize != 0) {
                totalPage++;
            }
            if (pageTo % pagesize != 0) {
                currentPage++;
            }
        }
        var pageData = "Page " + currentPage + "/" + totalPage;
        $("#select-page").val(currentPage);
        $("#page-data").html(pageData);
        return pageData;
    }
})

$(".authorizing__absence-requests--reject").click(function (e) {
    ResetAlert();
    var selectedList = $("#authorizer-request-list").bootstrapTable('getSelections');
    if (selectedList.length == 0) {
        ShowErrorAlert("Please select a request");
    }
    else {
        var listOfRequestId = [];
        $.each(selectedList, function (index, element) {
            listOfRequestId.push(element.Id);
        });
        $("#confirm-notificator__note").val("");
        $("#confirm-notificator__text").html("Are you sure you want to reject it?");
        $("#confirm-notificator").modal().one('click', "#confirm-notificator__Ok", function () {
            var note = $("#confirm-notificator__note").val();
            var data = {
                RequestIds: listOfRequestId,
                Note: note
            }
            DoPostForList(rejectManyUrl, data);
        })
    }
})

$(".authorizing__absence-requests--approve").click(function (e) {
    ResetAlert();
    var selectedList = $("#authorizer-request-list").bootstrapTable('getSelections');
    if (selectedList.length == 0) {
        ShowErrorAlert("Please select a request");
    }
    else {
        var listOfRequestId = [];
        $.each(selectedList, function (index, element) {
            listOfRequestId.push(element.Id);
        });
        $("#confirm-notificator__note").val("");
        $("#confirm-notificator__text").html("Are you sure you want to approve it?");
        $("#confirm-notificator").modal().one('click', "#confirm-notificator__Ok", function () {
            var note = $("#confirm-notificator__note").val();
            var data = {
                RequestIds: listOfRequestId,
                Note: note
            }
            DoPostForList(approveManyUrl, data);
        })
    }
})

function InitTeamCalendar(month, year) {
    $("#spinner-loading").show();
    $.ajax({
        url: teamCalendarDataUrl,
        type: 'POST',
        data: {month: month, year: year},
        dataType: "json",
        content: "application/json;charset=utf-8",
        success: function (result) {
            if (result == "SessionExpired") {
                window.location.replace(loginPageUrl);
                return;
            }
            flattenedData = jQuery.map(result, function (d) { return FlattenJson(d) });
            jQuery.each(result, FlattenJson);
            $('#team-calendar').bootstrapTable('destroy');
            $('#team-calendar').bootstrapTable({
                data: flattenedData,
                onPostBody: OnRefreshTable("#team-calendar", month, year)
            });
            $("#Months").val(month);
            $("#Years").val(year);
            $("#spinner-loading").hide();
        },
        error: function (result) {
            $("#spinner-loading").hide();
        }
    })
}

function OnRefreshTable($element, month, year) {
    $($element).find('.authorizing__column-day').attr('colspan', DaysInMonth(year, month));
    SetWeekenDaysInMonth($element, year, month);
    InitColumnTable($element, DaysInMonth(year, month));
}

function InitColumnTable(element, column) {
    $(element).find(".hidden").removeClass("hidden");
    var theadDayElemnet = $(element).find('thead tr:nth-child(2)');
    var numberOfChildElement = theadDayElemnet.children().length;
    while (numberOfChildElement > column) {
        numberOfChildElement--;
        var indexOfThead = numberOfChildElement;
        var indexOfTbody = numberOfChildElement + 1;
        var indexOfTheadElement = "th:eq(" + indexOfThead + ")";
        var indexOfTbodyElement = "td:eq(" + indexOfTbody + ")";
        theadDayElemnet.find(indexOfTheadElement).addClass("hidden");
        $(element).find('tbody tr').find(indexOfTbodyElement).addClass("hidden");
    }
}

function SetWeekenDaysInMonth(element, year, month) {
    $(element).find('thead tr:nth-child(2) th').removeClass("weekend-saturday");
    $(element).find('thead tr:nth-child(2) th').removeClass("weekend-sunday");
    var days = DaysInMonth(year, month);
    for (var i = 1; i <= days ; i++) {
        var newDate = new Date(year, month - 1, i)
        if (newDate.getDay() == 0) {
            $(element).find('thead tr:nth-child(2) th').filter(function () { return $(this).html() == i }).addClass("weekend-sunday");
        }
        if (newDate.getDay() == 6) {
            $(element).find('thead tr:nth-child(2) th').filter(function () { return $(this).html() == i }).addClass("weekend-saturday");
        }
    }
}

function DoPostForList(url, data) {
    $("#spinner-loading").show();
    ResetAlert();
    $.ajax({
        url: url,
        type: 'POST',
        data: data,
        dataType: "json",
        traditional: true,
        success: function (result) {
            if (result == "SessionExpired") {
                window.location.replace(loginPageUrl);
                return;
            }
            $('#authorizer-request-list').bootstrapTable('load', result.data);
            if (result.isSuccess == true) {
                ShowSuccessAlert(result.message);
            }
            else {
                ShowErrorAlert(result.message);
            }
            $("#spinner-loading").hide();
        },
        error: function (result) {
            ShowErrorAlert(ajaxErrorText);
            $("#spinner-loading").hide();
        }
    })
}

function ActionFormat(value, row, index) {
    return [
        '<i class="glyphicon glyphicon-ok authorizing__absence-requests--approve-icon"></i>',
        '<a class="authorizing__absence-requests--approve-inbox" href="javascript:void(0)" title="Approve">',
        'Approve',
        '</a>',
        '<i class="glyphicon glyphicon-remove authorizing__absence-requests--reject-icon"></i>',
        '<a class="authorizing__absence-requests--reject-inbox ml10" href="javascript:void(0)" title="Reject">',
        'Reject',
        '</a>'
    ].join('');
}

function PopupFormatter(value,row,index) {
    return '<div class="testpopover" data-toggle="popover"  data-placement="top" data-container="body" data-content="' + row.Note + '" data-original-title data-trigger="hover">' + value + '</div>'
}

function SearchForTeamCalendar() {

    var month = $("#Months").val();
    var year = $("#Years").val();
    currentIndex = 0;
    item = 0;
    if (month == oldMonth && year == oldYear) {
        oldMonth = month;
        oldYear = year;
        return;
    }
    oldMonth = month;
    oldYear = year;
    InitTeamCalendar(month, year);
}

function FlattenJson(data) {
    var result = {};
    function recurse(cur, prop) {
        if (Object(cur) !== cur) {
            result[prop] = cur;
        } else if (Array.isArray(cur)) {
            for (var i = 0, l = cur.length; i < l; i++)
                recurse(cur[i], prop + "[" + i + "]");
            if (l == 0)
                result[prop] = [];
        } else {
            var isEmpty = true;
            for (var p in cur) {
                isEmpty = false;
                recurse(cur[p], prop ? prop + "." + p : p);
            }
            if (isEmpty && prop)
                result[prop] = {};
        }
    }
    recurse(data, "");
    return result;
}

function DaysInMonth(year, month) {
    return new Date(year, month, 0).getDate();
}

function CellStyle(value, row, index) {
    if (index == currentIndex) {
        ++item;
    } else {
        item = 1;
        currentIndex++;
    }
    var rowIndex = item < 10 ? "0" + item.toString() : item.toString();
    rowIndex = "Day" + rowIndex + ".Status";
    switch (row[rowIndex]) {
        case "Taken":
            return {
                classes: "authorised"
            };
        case "Authorised":
            return {
                classes: "authorised"
            };
        case "New":
            return {
                classes: "in-progress"
            };
        case "In-Progress":
            return {
                classes: "in-progress"
            };
        case "Holiday":
            return {
                classes: "holiday"
            };
        default: {
            return {};
        }
    }
}
