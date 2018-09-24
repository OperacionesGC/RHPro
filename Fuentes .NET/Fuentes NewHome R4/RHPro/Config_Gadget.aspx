<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Config_Gadget.aspx.cs" Inherits="RHPro.Config_Gadget" %>

<!-- JPB: Esta pagina sirve para ordenar los gadgets. Se le debe pasar los parametros a traves de variables por URL --> 
 


<%
    RHPro.Lenguaje Obj_Lenguaje;    
    Obj_Lenguaje = new RHPro.Lenguaje();    
    //Recupero el numero de gadget a procesar    
    int gadnro = Convert.ToInt32(HttpContext.Current.Request.QueryString["gadnro"]);    
    //Recupero el orden (subir o bajar)    
    int sube = Convert.ToInt32(HttpContext.Current.Request.QueryString["sube"]);    
    //Recupero si hay que desactivar      
    int desactiva = Convert.ToInt32(HttpContext.Current.Request.QueryString["desactiva"]);    
    //Recupero si hay que activar      
    int activa = Convert.ToInt32(HttpContext.Current.Request.QueryString["activa"]);
    //Recupero el  gadget origen    
    int gadnro1 = Convert.ToInt32(HttpContext.Current.Request.QueryString["gadnro1"]);    
    //Recupero el  gadget destino    
    int gadnro2 = Convert.ToInt32(HttpContext.Current.Request.QueryString["gadnro2"]);

    //Recupero el ancho actual del gadget
    int anchoActual = Convert.ToInt32(HttpContext.Current.Request.QueryString["anchoActual"]);
    //Recupero el alto actual del gadget
    int altoActual = Convert.ToInt32(HttpContext.Current.Request.QueryString["alturaActual"]);
    //Recupero las dimensiones del gadget
    int ModificoAlto = Convert.ToInt32(HttpContext.Current.Request.QueryString["ModificoAlto"]);
    int ModificoAncho = Convert.ToInt32(HttpContext.Current.Request.QueryString["ModificoAncho"]);

    //Recupero el tipo de gadget
    int gadtipo = Convert.ToInt32(HttpContext.Current.Request.QueryString["gadtipo"]);
 
    if ((ModificoAncho == -1) || (ModificoAlto==-1))
    { 
        if (ModificoAncho == -1)
        {
            if (AjustarAncho(gadnro, anchoActual, gadtipo))
                Response.Write("<script>window.parent.document.getElementById('ifrm2').src = '';  window.parent.window.top.location = window.parent.window.top.location; </script> ");
            else
                Response.Write("<script>window.parent.Cerrar_Progreso(); alert('" + Obj_Lenguaje.Label_Home("Error al ajustar el ancho.") + "'); window.parent.document.getElementById('ifrm2').src = '' ;  </script> ");
        }
        else
            if (ModificoAlto == -1)
            {
                if (AjustarAlto(gadnro, altoActual, gadtipo))
                    Response.Write("<script>window.parent.document.getElementById('ifrm2').src = '';  window.parent.window.top.location = window.parent.window.top.location; </script> ");
                else 
                    Response.Write("<script>window.parent.Cerrar_Progreso(); alert('" + Obj_Lenguaje.Label_Home("Error al ajustar la altura.") + "'); window.parent.document.getElementById('ifrm2').src = '' ;  </script> ");
            }
    }
    else
    {

        if (activa == 0)
        {//Si no activa        
            if (desactiva == 0)
            {//Si no desactiva

                if ((gadnro1 != -1) && (gadnro2 != -1))
                {

                    if (IntercambiarPosicion(gadnro1, gadnro2)) //Intercambia la posicion del gadget con el anterior gadget            
                    {
                        if (!IsPostBack) //Cuando finaliza de ordenar el gadget se refresca el contenedor principal
                            Response.Write("<script>window.parent.document.getElementById('ifrm2').src = '';  window.parent.window.top.location = window.parent.window.top.location; </script> ");
                    }
                    else Response.Write("<script>window.parent.Cerrar_Progreso(); alert('" + Obj_Lenguaje.Label_Home("Error al intercambiar posiciones.") + "'); window.parent.document.getElementById('ifrm2').src = '' ;  </script> ");
                }
                else
                {
                    Int32 posicion = get_Posicion(gadnro);
                    if (sube == -1)
                    {
                        if (posicion > get_Min_Posicion())
                        { //Verifico que no desborde si elige subir el gadget
                            //IntercambiarPosicion(gadnro, Anterior_Gadget(posicion)); //Intercambia la posicion del gadget con el anterior gadget 
                            if (IntercambiarPosicion(gadnro, Siguiente_Gadget(posicion))) //Intercambia la posicion del gadget con el anterior gadget            
                            {
                                if (!IsPostBack) //Cuando finaliza de ordenar el gadget se refresca el contenedor principal
                                    Response.Write("<script>window.parent.document.getElementById('ifrm2').src = ''; window.parent.window.top.location = window.parent.window.top.location; </script> ");
                            }
                            else Response.Write("<script>window.parent.Cerrar_Progreso(); alert('" + Obj_Lenguaje.Label_Home("Error al intercambiar posiciones.") + "');window.parent.document.getElementById('ifrm2').src = '' ;  </script> ");
                        }
                        else
                            Response.Write("<script>window.parent.Cerrar_Progreso(); window.parent.document.getElementById('ifrm2').src = '' ;  </script> ");
                    }
                    else
                    {
                        if (posicion < get_Max_Posicion())
                        {  //Verifico que no desborde si elige bajar el gadget 
                            //IntercambiarPosicion(gadnro, Siguiente_Gadget(posicion)); //Intercambia la posicion del gadget con el siguiente gadget 
                            if (IntercambiarPosicion(gadnro, Anterior_Gadget(posicion))) //Intercambia la posicion del gadget con el anterior gadget            
                            {
                                if (!IsPostBack) //Cuando finaliza de ordenar el gadget se refresca el contenedor principal
                                    Response.Write("<script>window.parent.document.getElementById('ifrm2').src = '' ; window.parent.window.top.location = window.parent.window.top.location;</script> ");
                                else Response.Write("<script>window.parent.Cerrar_Progreso(); alert('" + Obj_Lenguaje.Label_Home("Error al intercambiar posiciones.") + "');window.parent.document.getElementById('ifrm2').src = '' ;  </script> ");
                            }
                            else
                                Response.Write("<script>window.parent.Cerrar_Progreso(); window.parent.document.getElementById('ifrm2').src = '' ;  </script> ");
                        }
                        else
                            Response.Write("<script>window.parent.Cerrar_Progreso(); window.parent.document.getElementById('ifrm2').src = '' ;  </script> ");
                    }

                }
            }
            else
                if (desactiva == -1)
                { //Desactiva

                    if (Desactivar(gadnro))
                    {   //if (!IsPostBack) //Cuando finaliza de ordenar el gadget se refresca el contenedor principal
                        Response.Write("<script>window.parent.document.getElementById('ifrm2').src = '' ; window.parent.window.top.location.reload(); </script> ");
                    }
                    else
                    {
                        Response.Write("<script>window.parent.Cerrar_Progreso(); alert('Error al desactivar'); </script> ");
                    }
                }

        }
        else
            if (activa == -1)
            { //Activa
                Activar(gadnro);
                //if (!IsPostBack) //Cuando finaliza de ordenar el gadget se refresca el contenedor principal
                Response.Write("<script>window.parent.document.getElementById('ifrm2').src = '' ;window.parent.window.top.location.reload(); </script> ");
            }
            else
                Response.Write("<script>window.parent.Cerrar_Progreso(); window.parent.document.getElementById('ifrm2').src = '' ;  </script> ");

    }
 
%>
