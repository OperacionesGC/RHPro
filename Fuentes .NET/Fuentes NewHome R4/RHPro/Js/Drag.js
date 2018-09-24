 
 var objTomado = "";//Mantiene el objeto que se selecciona para moverse
 var posMouseX = 0;//Mantiene la posicion X del mouse
 var posMouseY = 0;//Mantiene la posicion Y del mouse
 var navegador = "";//Mantiene el navegador
 var TDdestinoFinal ="";
// Detecto el navegadoregador
if(navigator.userAgent.indexOf("MSIE")>=0) navegador=0; // IE
else navegador=1; // Otros
 

function color(td) {
	  TDdestinoFinal = td;
//Pone el color de fondo del TD en el cual esta parado el objeto seleccionado	
 if (objTomado!="")
   td.style.backgroundColor = "#CCCCCC";
}
 
function saleTD(td) {
//Actualiza el color de fondo del TD en el cual estaba parado cuando hay un objeto seleccionado	
 if (objTomado!="")
   td.style.backgroundColor = "#FFFFFF";
}
 
  
 
function posRaton(e) { 
//Retorna un arreglo con las coordenadas x e y del mouse
  var pos=new Array();
  var x;
  var y;
  if (navegador!=0) {
    pos["x"] = e.pageX;
    pos["y"] = e.pageY;
  }
  else {
    pos["x"] = event.clientX + document.documentElement.scrollLeft;
    pos["y"] = event.clientY + document.documentElement.scrollTop;
  }  
 
  return pos; 
} 

function noEventos(event){
//Detiene todo tipo de eventos	
	if(navegador==0) //IE
	{
		window.event.cancelBubble=true;
		window.event.returnValue=false;
	}

	if(navegador==1) event.preventDefault();

}


function Top(obj,e){
	
//  var leftTomado =  ( document.getElementById("contenedor").offsetLeft + obj.offsetLeft ) ;  
//  var topTomado = (document.getElementById("contenedor").offsetTop + obj.offsetTop  )  ;  
  
  
  var x;
  var y;	 
/*  if(navegador==1) {  
   // document.exf1.sv_x.value = window.pageXOffset;
   // document.exf1.sv_y.value = window.pageYOffset;
 
    _x = event.pageXOffset;
    _y = event.pageYOffset;
 
  }
  else { //IE
    _x = event.clientX + document.body.scrollLeft;
    _y = event.clientY +  document.body.scrollTop;
  }*/
  
  x = (navegador==0) ? e.pageX : event.clientX
  y = (navegador==0) ? e.pageY : event.clientY


  obj.innerHTML = "Left:"+x+" Top:"+y; 
}

function Posicion_Absoluta(obj){
var left, top;
var pos=new Array();
    left = top = 0;
    if (obj.offsetParent) {
        do {
            left += obj.offsetLeft;
            top  += obj.offsetTop;
        } while (obj = obj.offsetParent);
    }
    pos["x"] = left;
    pos["y"] = top;
    
    return  pos;
}


function Tomar(obj){
//Prepara el objeto tomado para el movimiento	
 /*var pos = posRaton(event); 	
 var id_TDorigen = obj.id.replace("drag_","");
 objTomado = obj;
 
  var leftMouse = pos["x"];
  var topMouse = pos["y"];
  
  objTomado.style.left = leftMouse + "px";
  objTomado.style.top = topMouse + "px";  
 
  posMouseX = leftMouse;
  posMouseY = topMouse; 
 
  document.getElementById("gadnro_"+id_TDorigen).style.height = "80px";
  noEventos(event); */
 var pos = posRaton(event); 	
 var id_TDorigen = obj.id.replace("drag_","");
 objTomado = obj;
 
  var leftMouse = pos["x"];   
  var topMouse = pos["y"];
  
 // objTomado.style.left = leftTomado + "px";  
 objTomado.style.left = Posicion_Absoluta(objTomado)["x"] + "px";
 objTomado.style.top = Posicion_Absoluta(objTomado)["y"] + "px";
 objTomado.style.width = objTomado.offsetWidth + "px";

 
  posMouseX = leftMouse;
  posMouseY = topMouse; 
 
 // document.getElementById("gadnro_"+id_TDorigen).style.height = "200px";
  document.getElementById("gadnro_"+id_TDorigen).style.height = objTomado.offsetHeight + "px";
  noEventos(event);  

}

function Mover(){
//Realiza el movimiento del objeto tomado  
 
 
  if (objTomado) {//Pone al objeto flotante y lo inicializa segun las coordenadas x=left e y=top del mouse
     var pos = posRaton(event); 
     objTomado.style.position = "absolute";     
     
	 objTomado.style.left = pos["x"]  + "px";
     objTomado.style.top = pos["y"] + "px";
   
     noEventos(event); 
  } 
 
   
}

function Soltar(TDdestino){ 

 TDdestino = TDdestinoFinal;
 //Suelta el objeto tomado en un TD determinado por TDdestino
 if ( (objTomado) && (TDdestino)) { 
 	 
	 var origen = objTomado.id.replace("drag_","");
	 var destino = TDdestino.id.replace("gadnro_","");
 
	 document.getElementById("drag_"+origen).style.position = "";	
	 if (origen!=destino) {
	   /* Intercambia contenido entre TDs y Divs */	 
	   IntercambiaContenido(origen,destino);	 	   	   
	   document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro1="+origen+"&gadnro2="+destino+"&sube=0&desactiva=0&activa=0";  
	 }
	 TDdestino.style.backgroundColor = "#FFFFFF"; 
 
	 objTomado = "";
 }
 noEventos(event);
}

function IntercambiaContenido(origen,destino){ 
//Intercambia el InnerHTML del TD origen (TD del objeto tomado) con el TD destino (en el cual se suelta el objeto tomado). Luego intercambia el id de los TD.	
	var TD_origen = document.getElementById("gadnro_"+origen);	    
	var TD_destino = document.getElementById("gadnro_"+destino);
	var drag_origen = document.getElementById("drag_"+origen);	   
	var drag_destino = document.getElementById("drag_"+destino);	   	 
	var Aux = TD_origen.innerHTML;	 
	//Intercambia innerHTML de los TD
	TD_origen.innerHTML =TD_destino.innerHTML;
	TD_destino.innerHTML = Aux;
    TD_origen.style.height = "";
	//Intercambia id de los TD		   
	Aux = TD_destino.id;
	TD_destino.id = TD_origen.id;
	TD_origen.id = Aux;
}
 
 