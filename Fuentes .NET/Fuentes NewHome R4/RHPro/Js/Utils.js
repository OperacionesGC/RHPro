 


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
    window.open("./Controls/PopVideo/popVideo/index.html?path=./../../../../" + path + "&title=" + title, id, "width=830,height=635,toolbar=0"); //FLV 800 x 600

}

function AjustarIframe(id) {
    var altura;
    
    if (!window.opera && document.all && document.getElementById) {
        altura =(id.contentWindow.document.body.scrollHeight+ 30) + "px";;
    } else if (document.getElementById) {
        altura = ( id.contentDocument.body.scrollHeight + 30) + "px";
    }
    
    id.style.height = altura;
    
    
}
