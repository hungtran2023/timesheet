var loginPageUrl = $("#loginPageUrl").val();
var timesheetListPageUrl = $("#timesheetListPageUrl").val();
var timesheetUrl = $("#timesheetPageUrl").val();
var workingHours = $("#workingHours").val();
var maxYear = (new Date().getFullYear()) + 2;
var minYear = '01/01/2000';
var maxRangeOfDate = "01/01/" + maxYear;

$(document).ready(function () {
    var TIME_PATTERN = /^(08|09|8|9|1[0-7]{1}):[0-5]{1}[0-9]{1}$/;
    $('#StartDatePicker')
        .datepicker({
            autoclose: true,
            format: 'dd/mm/yyyy',
            todayHighlight: true,
            daysOfWeekDisabled: [0, 6],
            constrainInput: false
        })
        .datepicker('setDate', new Date())
        .on('changeDate', function (e) {
            $('#hr-timesheet-form').formValidation('revalidateField', 'EndDate');
        });
    $('#EndDatePicker')
        .datepicker({
            autoclose: true,
            format: 'dd/mm/yyyy',
            todayHighlight: true,
            daysOfWeekDisabled: [0, 6],
            constrainInput: false,
        })
        .datepicker('setDate', new Date())
        .on('changeDate', function (e) {
            $('#hr-timesheet-form').formValidation('revalidateField', 'EndDate');
        });

    $('#hr-timesheet-form').formValidation({
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
            Hours: {
                validators: {
                    numeric: {
                        message: 'The value is not a number',
                        thousandsSeparator: '',
                        decimalSeparator: '.'
                    },
                    notEmpty: {
                        message: 'The Hours field is required'
                    },
                    between: {
                        min: 0.1,
                        max: parseFloat(workingHours),
                        massage: 'The Hours should be greater than 0 and less than or equal 8.5'
                    },

                }
            },
            Note: {
                validators: {
                    notEmpty: {
                        message: 'The Note field is required'
                    }
                }
            }
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
        var url = $("#addTimeSheetUrl").val();
        $("#spinner-loading").show();
        $.ajax({
            url: url,
            type: 'POST',
            data: $(this).serialize(),
            dataType: "json",
            content: "application/json;charset=utf-8",
            success: function (result) {
                if (result == "SessionExpired") {
                    window.location.replace(loginPageUrl);
                    return;
                }
                if (result.isSuccess == true) {
                    window.location.replace(timesheetUrl);
                }
                else {
                    $("#spinner-loading").hide();
                    ShowErrorAlert(result.message);
                }
            },
            error: function (result) {
                $("#spinner-loading").hide();
                $('.notify-message').html(result.message);
                $('#holidayBookingNotifier').modal('show');
                $("#spinner-loading").hide();
            }
        })
    })

    InitMenu();
});

$("#cancel-timesheet").click(function (e) {
    e.preventDefault();
    window.location.replace(timesheetListPageUrl)
})

function InitMenu() {
    $("ul[data-menu-toggle=Management-Console]").removeClass("hide");
    $("a[data-menu-toggle=Management-Console]").addClass("selected-menu");
    $("ul[data-menu-toggle=Timesheets]").removeClass("hide");
    $("a[data-menu-toggle=Timesheets]").addClass("selected-menu");
}
