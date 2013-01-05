﻿
var htmlData = '';

$(function () {
    $('#jQuerySheet0').sheet({
        inlineMenu: inlineMenu($.sheet.instance),
        //urlGet: "/sheets/temporary.html",
        fnSave: function () { $('#Button2').trigger('click'); },
        autoFiller: true,
        buildSheet: '0x0'
    });
    getHtmlData(window.location.pathname);
});

function getHtmlData(pathname) {
    $.ajax({
        type: "POST",
        url: pathname + "/getHTML",
        data: "{}",
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (msg) {
            htmlData = msg.d;
            $('#buttonHidden').trigger('click');
        },
        error: function (request, status, error) {
            alert(request.responseText);
        }
    });
}

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