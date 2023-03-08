var ajaxErrorText = "Errors occured please try another time.";

$(document).ready(function () {
    $("#Alert a.close").click(function () {
        ResetAlert();
    });
})

function ResetAlert() {
    $("#Alert").addClass("hide");
    $("#Alert").removeClass("alert-success");
    $("#Alert").removeClass("alert-danger");
}

function ShowErrorAlert(text) {
    $("#Alert").removeClass("hide");
    $("#Alert").addClass("alert-danger");
    $("#AlertMessageHeader").html("Error: ");
    $("#AlertMessage").html(text);
}

function ShowSuccessAlert(text) {
    $("#Alert").removeClass("hide");
    $("#Alert").addClass("alert-success");
    $("#AlertMessageHeader").html("Success: ");
    $("#AlertMessage").html(text);
}

