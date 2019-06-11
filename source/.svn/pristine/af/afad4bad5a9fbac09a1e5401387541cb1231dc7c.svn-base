jQuery(document).ready(function ($) {
    globalJquery.extend($.ui.dialog.prototype, {
        'addbutton': function (buttonName, func) {
            var buttons = this.element.dialog('option', 'buttons');
            buttons[buttonName] = func;
            this.element.dialog('option', 'buttons', buttons);
        }
    });

    globalJquery.extend($.ui.dialog.prototype, {
        'removebutton': function (buttonName) {
            var buttons = this.element.dialog('option', 'buttons');
            delete buttons[buttonName];
            this.element.dialog('option', 'buttons', buttons);
        }
    });
}); 

function SetElement(id, str) {
    targetEl = document.getElementById(id);
    targetEl.value = str;
}

function SetText(id, str) {
    targetEl = document.getElementById(id);
    targetEl.innerHTML = str;
}

function ShowIt(id) {
    document.getElementById(id).style.visibility = "visible";
}

function HideIt(id) {
    document.getElementById(id).style.visibility = "hidden";
}

function SetEnable(id) {
    targetEl = document.getElementById(id);
    targetEl.disabled = false;
}

function IsNumeric(sText) {
    var ValidChars = "0123456789.";
    var IsNumber = true;
    var Char;


    for (i = 0; i < sText.length && IsNumber == true; i++) {
        Char = sText.charAt(i);
        if (ValidChars.indexOf(Char) == -1) {
            IsNumber = false;
        }
    }
    return IsNumber;
}

function isMaxText(aText, limit) {
    if (aText.length > limit)
        return true;
    else
        return false;
}

 