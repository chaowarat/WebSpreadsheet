
String.prototype.fileExists = function () {
    filename = this.trim();

    var response = jQuery.ajax({
        url: filename,
        type: 'HEAD',
        async: false
    }).status;

    return (response != "200") ? false : true;
}

$(function () {
    alert(window.location.href);
    alert(getAbsolutePath());
    alert(location.pathname);
    //alert('AAA' + "~/sheets/temporary.html".fileExists());
    //alert('BBB' + "css/page-style.css".fileExists());
    //alert('BBB' + "~/css/page-style.css".fileExists());
    //Here is where we initiate the sheets
    //every time sheet is created it creates a new jQuery.sheet.instance (array), to manipulate each sheet, the jQuery object is returned
    $('#jQuerySheet0').sheet({
        title: 'Benefit Admin',
        inlineMenu: inlineMenu($.sheet.instance),
        //urlGet: "~/sheets/temporary.html",
        autoFiller: true
    });

    function getAbsolutePath() {
        var loc = window.location;
        var pathName = loc.pathname.substring(0, loc.pathname.lastIndexOf('/') + 1);
        return loc.href.substring(0, loc.href.length - ((loc.pathname + loc.search + loc.hash).length - pathName.length));
    }

});

function inlineMenu(I) {
    I = (I ? I.length : 0);

    //we want to be able to edit the html for the menu to make them multi-instance
    var html = $('#inlineMenu').html().replace(/sheetInstance/g, "$.sheet.instance[" + I + "]");

    var menu = $(html);

    //The following is just so you get an idea of how to style cells
    menu.find('.colorPickerCell').colorPicker().change(function () {
        $.sheet.instance[I].cellChangeStyle('background-color', $(this).val());
    });

    menu.find('.colorPickerFont').colorPicker().change(function () {
        $.sheet.instance[I].cellChangeStyle('color', $(this).val());
    });

    menu.find('.colorPickers').children().eq(1).css('background-image', "url('images/palette.png')");
    menu.find('.colorPickers').children().eq(3).css('background-image', "url('images/palette_bg.png')");


    return menu;
}

function goToObj(s) {
    $('html, body').animate({
        scrollTop: $(s).offset().top
    }, 'slow');
    return true;
}  