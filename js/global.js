site = typeof site == 'object' ? site : {};

// First, checks if it isn't implemented yet.
if (!String.prototype.format) {
    String.prototype.format = function() {
        var args = arguments;
        return this.replace(/{(\d+)}/g, function(match, number) {
          return typeof args[number] != 'undefined'
            ? args[number]
            : match
          ;
        });
    };
}

String.prototype.isJSONString = function(){
    try {
        JSON.parse(this);
    } catch (e) {
        return false;
    }
    return true;
}

Date.prototype.formatMYSQL = function(includeTime){
    if(typeof includeTime != 'boolean') includeTime = true;
    var date = this;
    var retString = date.getFullYear() + "-" + (date.getMonth()+1) + "-" + date.getDate();

    if(includeTime){
        var hours = date.getHours();
        var minutes = date.getMinutes();
        minutes = minutes < 10 ? '0'+minutes : minutes;
        var secs = date.getSeconds();
        secs = secs < 10 ? '0'+secs : secs;
        retString += " " + hours + ':' + minutes + ':' + secs;
    }

    return retString;
}

// First, checks if it isn't implemented yet.
if (!String.prototype.format) {
    String.prototype.format = function() {
        var args = arguments;
        return this.replace(/{(\d+)}/g, function(match, number) {
          return typeof args[number] != 'undefined'
            ? args[number]
            : match
          ;
        });
    };
}
