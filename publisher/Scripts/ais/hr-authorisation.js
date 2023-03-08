var approveRequestUrl = $("#approveRequestUrl").val();
var rejectRequestUrl = $("#rejectRequestUrl").val();
var listOfRequestUrl = $("#listOfRequestUrl").val();
var loginPageUrl = $("#loginPageUrl").val();
var id = 0;
var statusNew = 0;
var statusInProgress = 1;
var statusUnAuthorised = 5;
var noError = false;
$(document).ready(function () {
    $('#StartDatePicker')
       .datepicker({
           autoclose: true,
           format: 'dd/mm/yyyy',
           todayHighlight: true,
           daysOfWeekDisabled: [0, 6]
       })
       .datepicker('setDate', new Date());
    $('#EndDatePicker')
        .datepicker({
            autoclose: true,
            format: 'dd/mm/yyyy',
            todayHighlight: true,
            daysOfWeekDisabled: [0, 6]
        })
        .datepicker('setDate', new Date());
    $("#hr-request-list").bootstrapTable({
        onPostBody: function () {
            var tableBody = $('#hr-request-list tbody');
            var tableRowsClass = $('#hr-request-list tbody tr');
            $('.search-sf').remove();
            $('#hr-request-list tbody tr').css('background', '#E7EBF5');
            tableRowsClass.each(function (i, val) {
                var dataRow = $('#hr-request-list').bootstrapTable('getRowByUniqueId', i);
                if (dataRow == null) {
                    return;
                }
                //if (dataRow.StatusId != statusNew && dataRow.StatusId != statusInProgress) {
                //    tableRowsClass.eq(i).hide();
                //}
                else {
                    $('.search-sf').remove();
                    tableRowsClass.eq(i).show();
                }
            });
            $('#hr-request-list tbody tr:visible:even').css('background', '#FFF2F2');
        },
        formatShowingRows: function (pageFrom, pageTo, totalRows) {
            var pagesize = $("#hr-request-list").data("page-size");
            pagesize = parseInt(pagesize);
            var currentPage = parseInt( pageTo / pagesize);
            var totalPage =parseInt( totalRows / pagesize);
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
        },
        onClickRow: function (row, $element) {
            if (row.Status == "Authorised" || row.Status == "Rejected" || row.Status == "Taken") {
                $('.notify-message').html("This request is " + row.Status + ". So you cannot edit it.");
                $('#hrAuthorisationNotifier').modal('show');
                return;
            }
            id = row.RequestId;
            $("#Id").val(row.RequestId);
            $("#Test").val(row.StartDate);
            $("#FullName").html(row.FullName);
            $("#AbsenceType").val(row.AbsenceType);
            $("#StartDate").val(row.StartDate);
            $("#EndDate").val(row.EndDate);
            $("#StartTime").val(row.StartTime);
            $('#StartDatePicker').datepicker('setDate', row.StartDate);
            $('#EndDatePicker').datepicker('setDate', row.EndDate);
            $("#EndTime").val(row.EndTime);
            $("#FirstAuthoriserId").val(row.FirstAuthoriserId);
            $("#SecondAuthoriserId").val(row.SecondAuthoriserId);
            $("#Note").val(row.Note);
            ResetAlert();
            $(".hr-authorisation").hide();
            $(".hr-authorisation-editor").show();
        }
    });

    $("#Status").change(function () {
        loadFilter();
    })

    $("#authorisation-form").formValidation({
        framework: 'bootstrap',
        fields: {
            HrNote: {
                validators: {
                    notEmpty: {
                        message: 'The Note is required'
                    }
                }
            }
        }
    })
    .on('success.field.fv', function (e, data) {
        var $parent = data.element.parents('.form-group');
        $parent.removeClass('has-success');
    })
    .on('success.form.fv', function (e, data) {
        var $form = $(e.target);
        var $button = $form.data('formValidation').getSubmitButton();
        var note = $("#HrNote").val();
        var type = $("#AbsenceType").val();
        var data = {
            id: id,
            note: note,
            type: type,
            status: $("#Status").val(),
            name: $("#Name").val(),
            department: $("Department").val()
        }
        if ($button.attr('id') == 'approve-request') {
            PostData(approveRequestUrl, data);
        }
        else {
            PostData(rejectRequestUrl, data);
        }
        e.preventDefault();
    });
    InitMenu();
});

$("#search").click(function () {
    loadFilter();
})

$("#show-all").click(function () {
    $.get(listOfRequestUrl, { },
            function (result) {
                $('#hr-request-list').bootstrapTable('load', result);
    });
})

$("#cancel-edit-request").click(function (e) {
    e.preventDefault();
    ResetAlert();
    ResetRequestForm();
    $(".hr-authorisation").show();
    $(".hr-authorisation-editor").hide();
})

$("#go-previous").click(function (e) {
    e.preventDefault();
    $("#hr-request-list").bootstrapTable("prevPage");
})

$("#go-next").click(function (e) {
    e.preventDefault();
    $("#hr-request-list").bootstrapTable("nextPage");
})

$("#select-page-btn").click(function (e) {
    e.preventDefault();
    var val = parseInt($("#select-page").val());
    $("#hr-request-list").bootstrapTable("selectPage", val);
})

function PostData(url, data) {
    $("#spinner-loading").show();
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
            if (result.isSuccess == true) {
                ShowSuccessAlert(result.message);
                ResetRequestForm();
            }
            else {
                ShowErrorAlert(result.message);
            }
            $(".hr-authorisation").show();
            $(".hr-authorisation-editor").hide();
            $('#hr-request-list').bootstrapTable('load', result.data);
            $("#spinner-loading").hide();
        },
        error: function (result) {
            ShowErrorAlert(ajaxErrorText);
            $("#spinner-loading").hide();
        }
    })
}

function loadFilter() {
    $.get(listOfRequestUrl, { status: $("#Status").val(), name: $("#Name").val(), department: $("#Department").val() },
            function (result) {
                $('#hr-request-list').bootstrapTable('load', result);
            });
}

function InitMenu() {
    $("ul[data-menu-toggle=Management-Console]").removeClass("hide");
    $("a[data-menu-toggle=Management-Console]").addClass("selected-menu");
    $("ul[data-menu-toggle=Annual-Leave]").removeClass("hide");
    $("a[data-menu-toggle=Annual-Leave]").addClass("selected-menu");
}

function ResetRequestForm() {
    $("#Id").val();
    $('#AbsenceType').val('4');
    $('#StartDatePicker').datepicker('setDate', new Date());
    $('#EndDatePicker').datepicker('setDate', new Date());
    $('#StartTime').val('08:00');
    $('#EndTime').val('17:30');
    $('#FirstAuthoriserId').val('0');
    $('#SecondAuthoriserId').val('0');
    $('#Note').val('');
    $('#HrNote').val('');
    $("#authorisation-form").data('formValidation').resetForm();
}



