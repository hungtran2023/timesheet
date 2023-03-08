$(document).ready(function () {
    $("a[data-menu-toggle]").on('click', function (e) {
        $(this).parents('div.body-content__left-menu--content').find('ul[data-menu-toggle]').addClass("hide");
        $(this).parents('div.body-content__left-menu--content').find('a[data-menu-toggle]').removeClass("selected-menu");
        $(this).addClass("selected-menu");
        $(this).closest('ul[data-menu-toggle]').toggleClass("hide");
        var menuToShow = $(this).data('menu-toggle');
        var parentMenuToShow = $(this).closest('ul[data-menu-toggle]').data('menu-toggle');
        $("ul[data-menu-toggle=" + menuToShow + "]").toggleClass("hide");
        $("a[data-menu-toggle=" + parentMenuToShow + "]").addClass("selected-menu");
    });

    ActiveLink();
})

function ActiveLink() {
    $(".body-content__left-menu--content a[href]").each(function () {
        if (document.location.href.indexOf("HREnterTimeSheet") != -1) {
            if (this.href.indexOf("management/tms/tms_list_staff.asp") != -1) {
                $(this).addClass("menu-active");
                return
            }
        }
        if (document.location.href == this.href) {
            $(this).addClass("menu-active");
            return
        }
    })
}

function GetQueryString(param) {
    var query = location.search.substring(1);
    var splitQuery = query.split("&");
    var result = "";
    $(splitQuery).each(function (index, value) {
        var pair = value.split("=");
        if (pair[0] == param) {
            result = pair[1];
            return;
        }
    })
    return result;
}

