var loginPageUrl = $("#loginPageUrl").val();
var maxYear = (new Date().getFullYear()) + 2;
var minYear = "01/01/2000";
var maxRangeOfDate = "01/01/" + maxYear;

$(document).ready(function () {
    var TIME_PATTERN = /^(08|09|8|9|1[0-7]{1}):[0-5]{1}[0-9]{1}$/;
    $('#StartDatePicker')
        .datepicker({
            autoclose: true,
            format: 'dd/mm/yyyy',
            todayHighlight: true,
            daysOfWeekDisabled: [0, 6]
        })
        .datepicker('setDate', new Date())
        .on('changeDate', function (e) {
            $('#request-form').formValidation('revalidateField', 'EndDate');
        });
    $('#EndDatePicker')
        .datepicker({
            autoclose: true,
            format: 'dd/mm/yyyy',
            todayHighlight: true,
            daysOfWeekDisabled: [0, 6]
        })
        .datepicker('setDate', new Date())
        .on('changeDate', function (e) {
            $('#request-form').formValidation('revalidateField', 'EndDate');
        });

    $('#request-form').formValidation({
        framework: 'bootstrap',
        fields: {
            AbsenceType: {
                validators: {
                    notEmpty: {
                        message: 'The Type of Absence is required'
                    }
                }
            },
            StartDate: {
                validators: {
                    notEmpty: {
                        message: 'The Start Date field is required'
                    },
                    date: {
                        message: 'Selected date should be after than 31/12/1999',
                        format: 'DD/MM/YYYY',
                        min: minYear
                    },
                    callback: {
                        message: 'The date must be before ' + maxRangeOfDate,
                        callback: function (value, validator) {
                            var m = new moment(value, 'DD/MM/YYYY', true);
                            var m1 = new moment(value, 'D/M/YYYY', true);
                            if (!m.isValid() && !m1.isValid) {
                                return false;
                            }
                            return m.isBefore(maxRangeOfDate) || m1.isBefore(maxRangeOfDate);
                        }
                    }
                },
                onSuccess: function (e, data) {
                    if (!data.fv.isValidField('EndDate')) {
                        data.fv.revalidateField('EndDate');
                    }
                }
            },
            EndDate: {
                validators: {
                    notEmpty: {
                        message: 'The End Date field is required'
                    },
                    date: {
                        format: 'DD/MM/YYYY',
                        min: 'StartDate',
                        message: 'Selected date should be after start date'
                    },
                    callback: {
                        message: 'The date must be before ' + maxRangeOfDate,
                        callback: function (value, validator) {
                            var m = new moment(value, 'DD/MM/YYYY', true);
                            var m1 = new moment(value, 'D/M/YYYY', true);
                            if (!m.isValid() && !m1.isValid) {
                                return false;
                            }
                            return m.isBefore(maxRangeOfDate) || m1.isBefore(maxRangeOfDate);
                        }
                    }
                },
                onSuccess: function (e, data) {
                    if (!data.fv.isValidField('StartDate')) {
                        data.fv.revalidateField('StartDate');
                    }
                }
            },
            StartTime: {
                verbose: false,
                validators: {
                    notEmpty: {
                        message: 'The Start time is required'
                    },
                    regexp: {
                        regexp: TIME_PATTERN,
                        message: 'The Start time must be between 08:00 and 17:00'
                    }
                }
            },
            EndTime: {
                verbose: false,
                validators: {
                    notEmpty: {
                        message: 'The End time is required'
                    },
                    regexp: {
                        regexp: TIME_PATTERN,
                        message: 'The End time must be between 08:00 and 17:00'
                    },
                    callback: {
                        message: 'The End time must be later then the Start one',
                        callback: function (value, validator, $field) {
                            var startTime = validator.getFieldElements('StartTime').val();
                            if (startTime == '' || !TIME_PATTERN.test(startTime)) {
                                return true;
                            }
                            var startHour = parseInt(startTime.split(':')[0], 10),
                                startMinutes = parseInt(startTime.split(':')[1], 10),
                                endHour = parseInt(value.split(':')[0], 10),
                                endMinutes = parseInt(value.split(':')[1], 10);

                            if (endHour > startHour || (endHour == startHour && endMinutes > startMinutes)) {
                                validator.updateStatus('StartTime', validator.STATUS_VALID, 'callback');
                                return true;
                            }
                            return false;
                        }
                    }
                }
            },
            FirstAuthoriserId: {
                validators: {
                    notEmpty: {
                        message: 'The Authoriser is required'
                    },
                }
            },
            SecondAuthoriserId:{
                validators: {
                    different: {
                        field: 'FirstAuthoriserId',
                        message: 'The second authoriser cannot be same as the first one'
                    }
                }
            },
            Note: {
                validators: {
                    stringLength: {
                        max: 500,
                        message: 'The note must be less than 500 characters long'
                    },
                }
            },
        }
    })
    .on('success.field.fv', function (e, data) {
        if (data.field === 'StartDate' && !data.fv.isValidField('EndDate')) {
            data.fv.revalidateField('EndDate');
        }
        if (data.field === 'EndDate' && !data.fv.isValidField('StartDate')) {
            data.fv.revalidateField('StartDate');
        }
        var $parent = data.element.parents('.form-group');
        $parent.removeClass('has-success');

    })
    .on('success.form.fv', function (e, data) {
        e.preventDefault();
        if ($("#submit-request").html() == "Update Request") {
            var updateRequestUrl = $("#updateRequestUrl").val();
            PostData(updateRequestUrl, $(this).serialize());
        }
        else {
            var addRequestUrl = $('#addRequestUrl').val();
            PostData(addRequestUrl, $(this).serialize());
        }
    });
    

    $('#request-list').bootstrapTable({
        onClickRow: function (row, $element) {
            if (row.Status == "New") {
                $("#Id").val(row.Id);
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
                $("#submit-request").html("Update Request");
                $("#cancel-request").show();
            }
        },
        formatShowingRows: function (pageFrom, pageTo, totalRows) {
            var pagesize = $("#request-list").data("page-size");
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
});

$(".holiday-booking__remove-btn").click(function (e) {
    ResetAlert();
    var selectedList = $("#request-list").bootstrapTable('getSelections');
    var listOfRequestId = "";
    selectedList.forEach(function (item, index) {
        if (item.Status == "New") {
            listOfRequestId += item.Id + ",";
        }
    })
    if (listOfRequestId.length == 0) {
        ShowErrorAlert("Please select a request");
    }
    else {
        var deleteUrl = $("#deleteRequestUrl").val();
        var data = { 'listOfRequestId': listOfRequestId };
        PostData(deleteUrl, data);
    }
})

$("#cancel-request").click(function (e) {
    e.preventDefault();
    ResetRequestForm();
})

function PostData(url, data) {
    ResetAlert();
    $("#spinner-loading").show();
    $.ajax({
        url: url,
        type: 'POST',
        data: data,
        dataType: "json",
        content: "application/json;charset=utf-8",
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
            $('#request-list').bootstrapTable('load', result.data);
            $("#spinner-loading").hide();
        },
        error: function (result) {
            ShowErrorAlert(ajaxErrorText);
            $("#spinner-loading").hide();
        }
    })
}

function ResetRequestForm() {
    var leaderId = $('#directLeaderId').val();
    $("#Id").val(''),
    $('#AbsenceType').val('4'),
    $('#StartDatePicker').datepicker('setDate', new Date()),
    $('#EndDatePicker').datepicker('setDate', new Date()),
    $('#StartTime').val('08:00'),
    $('#EndTime').val('17:00'),
    $('#FirstAuthoriserId').val(leaderId),
    $('#SecondAuthoriserId').val('0'),
    $('#Note').val(''),
    $("#submit-request").html('Submit Request');
    $("#cancel-request").hide();
    $("#request-form").data('formValidation').resetForm();
}

function RemoveFormatter(value, row, index) {
    if (row.Status !== "New" && index % 2 == 0) {
        return '<div class="hide-checkbox hide-checkbox-even"></div>';
    }
    if (row.Status !== "New" && index % 2 != 0) {
        return '<div class="hide-checkbox hide-checkbox-odd"></div>';
    }
}
