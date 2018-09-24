function abrirVentana(url, name, width, height) {
    abrirventana(url, name, width, height, null)
}
function abrirventana(url, name, width, height) {
    abrirventana(url, name, width, height, null)
}

function abrirVentana(url, name, width, height, opc) {
    abrirventana(url, name, width, height, opc)
}



function abrirventana(url, name, width, height, opc) {
 
    var str = "height=" + height + ",innerHeight=" + height;
    str += ",width=" + width + ",innerWidth=" + width;
    if (window.screen) {
        var ah = screen.availHeight - 30;
        var aw = screen.availWidth - 10;

        var xc = (aw - width) / 2;
        var yc = (ah - height) / 2;
        if (xc < 0)
            xc = 0
        if (yc < 0)
            yc = 0
        str += ",left=" + xc + ",screenX=" + xc;
        str += ",top=" + yc + ",screenY=" + yc;
    }

    str += ",resizable=yes,status=0,menubar=0,toolbar=0,location=0"
    if (opc != null)
        str += opc
    var auxi;
    auxi = url.substr(url.lastIndexOf('/') + 1, url.length);
    auxi = auxi.substr(0, auxi.indexOf(".asp"));

    window.open("../" + url, "", str);
    //window.open("../" + url, auxi, str);

}




function setScrollPosition(controlId, hiddenId) {

    var hidden = GetElement(hiddenId);

    
    if (hidden != null) {
        setTimeout("setScrollValue('" + controlId + "','" + hidden.value + "')", 0.1);       }
}

function setScrollValue(controlId, value) {
    var control = GetElement(controlId);

    if (control != null)
        control.scrollTop = value;
}

function saveScrollPosition(controlId, hiddenId) {

    var hidden = GetElement(hiddenId);
    var control = GetElement(controlId);
  
    if (hidden != null && control != null)
        hidden.value = control.scrollTop;

}
function ClearValue(id) {

    element = GetElement(id);

    if (element != null) {
        element.value = "";
  }
}
function GetElement(id) {

    element = document.getElementById(id);


    if (element == null && document.getElementsByName(id).length > 0)
        element = document.getElementsByName(id)[0];

    return element;
}

function popvideo(path, title, id) {
    
    //window.open("./Controls/PopVideo/popVideo/index.html?path=gti.flv" + "&title=" + title, "popVideo", "width=510,height=430,toolbar=0"); //FLV 480 x 320
    //window.open("./Controls/PopVideo/popVideo/index.html?path=gti.flv" + "&title=" + title, "popVideo", "width=350,height=350,toolbar=0"); //FLV 320 x 240
    window.open("./Controls/PopVideo/popVideo/index.html?path=./../../../../" + path + "&title=" + title, id, "width=830,height=635,toolbar=0"); //FLV 800 x 600

}

