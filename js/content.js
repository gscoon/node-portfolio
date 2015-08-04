var content = {
    createNewSubwindow: function(){
        var len = $('.sub_window').length;
        var sw = $('<div class="sub_window" id="sw_{0}"></div>'.format(len));
        $('#outer').append(sw);
        return sw;
    },

    displayTemplateTypeMenu: function(){

    },

    displaySheetMenu: function(selector, sheetArray){
        var html = '';

        sheetArray.forEach(function(sheet){
            html += '<a class="sheet_select_a" href="javascript:void">{0}</a>'.format(sheet.name);
        });

        $(selector).html(html);
    },

    displaySheetOptions: function(){

    },

    showSubWindow: function(selector){
        $('.sub_window').hide();
        $(selector).fadeIn(300);
    }
}

module.exports = content;
