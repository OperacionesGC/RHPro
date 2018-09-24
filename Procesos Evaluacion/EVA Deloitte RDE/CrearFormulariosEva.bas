Attribute VB_Name = "CrearFormulariosEva"
Option Explicit

'Version 1.00 - Deloitte RDE
' 17-05-2006 - LA - Agrego manejo de errores y que se escriba info en el log
' 02-02-2007 - LA - sacar las vistas
'                 - Inicializar Fecha Null para evadetevldor y incluir Comentarios para el log
'                 - Si no se borro la evacab en un proyecto No borrar el empleado del proyecto.


'Global Const Version = "1.01"
'Global Const FechaModificacion = "19-04-2007" ' Leticia Amadio
'Global Const UltimaModificacion = " " 'Secciones para Plan de Desarrollo - Deloitte

'Global Const Version = "1.02"
'Global Const FechaModificacion = "30-05-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " 'Asociarle la Etapa al empleado cuando se crea la evaluacion (y no solo cdo existe)

'Global Const Version = "1.03"
'Global Const FechaModificacion = "31-05-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " 'Borrar Secciones para CHILE (ref. Deloitte)


'Global Const Version = "1.04"
'Global Const FechaModificacion = "01-11-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Borrar Datos Seccion Evaluacion Areas

'Global Const Version = "1.05"
'Global Const FechaModificacion = "14-05-2008 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Borrar Datos Seccion Evaluacion General

'Global Const Version = "1.06"
'Global Const FechaModificacion = "17-04-2009 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Borrar Datos Seccion Comentarios de Competencias y borra del la ultima seccion a la primera

'Global Const Version = "1.07"
'Global Const FechaModificacion = "23-08-2011 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Deloitte no lo usa mas en el upgrade R3
                                      ' Integrar el proceso de DEL-Chile - agregar conexión encriptada - no dejar mensaje error inicio escriba en log.
                                      ' CUSTOM CHILE-DELOITTE -
                                      ' ( Version = 1.03.1 - 05-09-2007 - eliminar una evaluacion de un rol si el evaluador esta repetido )
                                      ' ( Version = 1.04 - 07-10-2009 -eliminar una evaluacion de Grupo Competencia y de Comptencias. )
                                      ' ( Version = 1.05 - 29-01-2010 - eliminar evaluaciones de nuevas secciones.)
    
Global Const Version = "1.08"
Global Const FechaModificacion = "30-01-2012 " ' Leticia Amadio
Global Const UltimaModificacion = " " ' Deloitte chile- agregar borrado de secciones objetivos.



' __________________________________________________________________________


Dim fs, f
'Global Flog

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global arrparam
Global evaevenro As Long
Global evaproynro As Long
Global perfil As String
Global modificar As String
Global listainicial As String

Global conceptos As String
Global acumuladores As String
Global procesos As String
Global idUser As String

    ' VERRRRRRRRLOOO
Const cautoevaluador = 1
Const cevaluador = 2
Const cgarante = 3
Const ctenroarea = 44 ' Division
Const ctenrogarante = 47


Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim Rsini As New ADODB.Recordset
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim ternro
Dim rsEmpl As New ADODB.Recordset

Dim tipoEvldor

Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros


    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProceso = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProceso = strCmdLine
                
            Else
                Exit Sub
            End If
        End If
    End If


    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
            
    Nombre_Arch = PathFLog & "CrearFormularioEvaDeloitte" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If


    On Error GoTo ME_Main

    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    cantRegistros = CInt(objRs!total)
    totalEmpleados = cantRegistros
    objRs.Close
       
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
   
    Flog.writeline "Inicio Proceso de Creación de Evaluaciones: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    Flog.writeline
    Flog.writeline "PID = " & PID


    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
        'Obtengo el nro de reporte
        parametros = objRs!bprcparam
       
        arrparam = Split(parametros, ".")
        evaevenro = arrparam(0)
        evaproynro = arrparam(1)
        modificar = arrparam(2)   ' define que se modifica-- borrar empls o generar form
        perfil = arrparam(3)      ' perfil del loguedo -- ver si se usa..
        
        listainicial = "0"
          'StrSql = "SELECT DISTINCT evacab.empleado "
          'StrSql = StrSql & " FROM  evacab "
          'StrSql = StrSql & " WHERE  evacab.evaevenro   = " & evaevenro
        StrSql = "SELECT DISTINCT evaproyemp.ternro "
        StrSql = StrSql & " FROM  evaproyemp "
        StrSql = StrSql & " WHERE  evaproyemp.evaproynro = " & evaproynro

                                        
        OpenRecordset StrSql, Rsini
        Do Until Rsini.EOF
            listainicial = listainicial & "," & Rsini("ternro")
            Rsini.MoveNext
        Loop
        Rsini.Close
        Set Rsini = Nothing
    
       'listainicial = arrparam(1)
       
       Flog.writeline " Parametro que entro: Evento: " & evaevenro
       Flog.writeline " Parametro que entro: proyecto: " & evaproynro
        Flog.writeline " Parametro que entro: modificar: " & modificar
       Flog.writeline " Lista inicial de empleados en el Evento: " & listainicial
       
     
       'EMPIEZA EL PROCESO

       'Obtengo los empleados sobre los que tengo que generar -
       Flog.writeline " Entra a cargar empleados de batch_empleado."
       CargarEmpleados NroProceso, rsEmpl
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
        objConn.Execute StrSql, , adExecuteNoRecords
                        
                If modificar = "ProyEmp" Then
                   Call borrarFormulario(evaproynro, evaevenro, listainicial)
                   Call borrarEmpProy(evaproynro, evaevenro, listainicial)
                   Call insertarEmpProy(NroProceso, evaproynro, evaevenro)
                    ' aca si es solo borrar empls --> ent tengo que borrar todos los empls del batch_Empl  ------???
                End If

              'Genero por cada empleado un registro
              Flog.writeline "   "
              Flog.writeline "   "
              Flog.writeline " Para cada empleado se GENERA su FORMULARIO EVALUACION."
                
               Do Until rsEmpl.EOF
                 EmpErrores = False
                 ternro = rsEmpl!ternro
          
                 'Genero los datos del empleado
                 'If modificar = "CrearFrm" Then
                        Flog.writeline "       "
                        Flog.writeline "      Generar Evaluacion para el empleado (ternro): " & ternro
                        Call generarFormulario(evaevenro, ternro)
                ' End If
                        
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
          
                cantRegistros = cantRegistros - 1
          
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
             
                objConn.Execute StrSql, , adExecuteNoRecords
          
                'Si se generaron todos los datos del empleado correctamente lo borro
                If Not EmpErrores Then
                StrSql = " DELETE FROM batch_empleado "
                StrSql = StrSql & " WHERE bpronro = " & NroProceso
                StrSql = StrSql & " AND ternro = " & ternro
                objConn.Execute StrSql, , adExecuteNoRecords
                End If
          
                rsEmpl.MoveNext
               Loop
             rsEmpl.Close
             Set rsEmpl = Nothing
                           
            objRs.Close
            Set objRs = Nothing
                            
    Else
        objRs.Close
        Set objRs = Nothing
        
        objConn.Close
        
        Exit Sub
    End If
    
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close
    
    objConn.Close
    
    Exit Sub
    
ME_Main:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub ' MAIN


'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    
    OpenRecordset StrEmpl, rsEmpl
End Sub



'__________________________________________________________________________
' 17-03-2006: L.A. no se elimine el evento
' 02-02-2006 - LA: - si no se borro la evacab No borrar el empleado.
'__________________________________________________________________________
Sub borrarEmpProy(evaproynro, evaevenro, listainicial)

Dim StrSql As String
Dim StrSql2 As String
Dim rs1 As New ADODB.Recordset
Dim ternro
Dim listborrar
Dim listnoborrar

Dim Empleado As Long
Dim tipsecprogdel
Dim rsBorrar As New ADODB.Recordset
Dim rsBorrar2 As New ADODB.Recordset

On Error GoTo ME_Proyecto

listnoborrar = ""

 ' saco from evacab --> evacab es si se le genero la evaluacion (puede haber empls e el proy y no se le genero form)
 StrSql = "SELECT evaproyemp.ternro, empleg, terape, ternom " ',evacab.evaevenro  ,evacab.empleado
 StrSql = StrSql & " FROM evaproyemp "
 StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = evaproyemp.ternro "
        ' evaevento --> se genera en generaren el asp... ---
 'StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evaevenro = evacab.evaevenro "
 'StrSql = StrSql & " WHERE evacab.empleado IN (" & listainicial & ")"
 StrSql = StrSql & " WHERE evaproyemp.ternro IN (" & listainicial & ")"
 StrSql = StrSql & " AND   evaproyemp.evaproynro  =" & evaproynro
 StrSql = StrSql & " AND   NOT EXISTS (SELECT * FROM batch_empleado WHERE "
  'StrSql = StrSql & " ternro = evacab.empleado"
 StrSql = StrSql & " ternro = evaproyemp.ternro "
 StrSql = StrSql & " and bpronro=" & NroProceso & ")"
 OpenRecordset StrSql, rsBorrar
 
 Do Until rsBorrar.EOF
     
    StrSql2 = "SELECT evacab.evacabnro "
    StrSql2 = StrSql2 & " FROM evacab "
    StrSql2 = StrSql2 & " WHERE empleado= " & rsBorrar("ternro")
    StrSql2 = StrSql2 & "   AND evacab.evaproynro =" & evaproynro
    OpenRecordset StrSql2, rsBorrar2
    
    If rsBorrar2.EOF Then
        listborrar = listborrar & "," & rsBorrar("ternro")
    Else
        listnoborrar = listnoborrar & " - " & rsBorrar("ternro")
    End If
    rsBorrar2.Close
    rsBorrar.MoveNext
 Loop
 
rsBorrar.Close
Set rsBorrar = Nothing
Set rsBorrar2 = Nothing

If Trim(listborrar) = "" Then
    listborrar = 0
Else
    listborrar = "0" & listborrar
End If
 

Flog.writeline "    Borrar Empleados del proyecto: " & listborrar
If listnoborrar <> "" Then
Flog.writeline "    Empleados que no se borraron porque no se borro su cabecera de evaluación (por algún inconveniente): " & listnoborrar
End If

StrSql = "DELETE FROM evaproyemp "
StrSql = StrSql & " WHERE ternro IN (" & listborrar & ") AND evaproynro=" & evaproynro

objConn.Execute StrSql, , adExecuteNoRecords
         

        
' ME FIJO SI NO HAY empleados en el equipo (en el proyecto), entonces elimino el EVENTO!!!!
'StrSql = "SELECT ternro FROM evaproyemp"
'StrSql = StrSql & " INNER JOIN evaproyecto ON evaproyecto.evaproynro = evaproyemp.evaproynro "
'StrSql = StrSql & " WHERE evaproyemp.evaproynro = " & evaproynro
'OpenRecordset StrSql, rs1
'If rs1.EOF Then
        'If Trim(evaevenro) <> "" Then
                'StrSql = "DELETE FROM evaevento "
                'StrSql = StrSql & " WHERE evaevenro=" & evaevenro
                'objConn.Execute StrSql, , adExecuteNoRecords
        'End If
'End If

'rs1.Close
'Set rs1 = Nothing


Exit Sub

ME_Proyecto:

    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    

End Sub
'-------------------------------------------------------------------------

'__________________________________________________________________________
'__________________________________________________________________________
Sub insertarEmpProy(NroProc, evaproynro, evaevenro)
Dim StrEmpl As String
Dim StrSql  As String
Dim rsEmpls As New ADODB.Recordset
Dim rs1     As New ADODB.Recordset
Dim ternro
On Error GoTo ME_Proyecto

Flog.writeline "    Inserta Empleados en el proyecto. "

StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
OpenRecordset StrEmpl, rsEmpls


Do Until rsEmpls.EOF
          'EmpErrores = False
   ternro = rsEmpls!ternro

        'buscar si ya existe ..........
        StrSql = "SELECT ternro "
        StrSql = StrSql & " FROM evaproyemp "
        StrSql = StrSql & " WHERE evaproynro = " & evaproynro
        StrSql = StrSql & "   AND  ternro     = " & ternro
        OpenRecordset StrSql, rs1

        If rs1.EOF Then
                StrSql = "INSERT INTO evaproyemp "
                StrSql = StrSql & " (evaproynro, ternro) "
                StrSql = StrSql & " VALUES (" & evaproynro & "," & ternro & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rs1.Close
        
    rsEmpls.MoveNext
    
Loop

Set rs1 = Nothing

rsEmpls.Close
Set rsEmpls = Nothing


Exit Sub

ME_Proyecto:

    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    

                  
End Sub


'__________________________________________________________________________
' 17-03-2006 - L.A. - se le aplico la función TRim a tipsecprogdel
'__________________________________________________________________________
Sub borrarFormulario(evaproynro, evaevenro, listainicial)

Dim StrSql As String
Dim rsBorrar As New ADODB.Recordset
Dim rsBorrar2 As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim ternro
        '????????
'Dim profecpago Dim EmprLogo Dim EmprLogoAlto Dim EmprLogoAncho Dim emprTer

Dim listborrar

Dim Empleado As Long

Dim tipsecprogdel

On Error GoTo ME_Borrar

 Flog.writeline "   "
 Flog.writeline " Entro a BORRAR evaluaciones de los empleados. "

 StrSql = "SELECT evacab.evacabnro, evacab.empleado, empleg, terape, ternom,evacab.evaevenro  "
 StrSql = StrSql & " FROM evacab "
 StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = evacab.empleado "
 StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evaevenro = evacab.evaevenro "
 StrSql = StrSql & " WHERE evacab.empleado IN (" & listainicial & ")"
 StrSql = StrSql & " AND   evaevento.evaproynro  =" & evaproynro
 StrSql = StrSql & " AND   NOT EXISTS (SELECT * FROM batch_empleado WHERE "
 StrSql = StrSql & " ternro = evacab.empleado"
 StrSql = StrSql & " and bpronro=" & NroProceso & ")"

 OpenRecordset StrSql, rsBorrar
 Do Until rsBorrar.EOF
    listborrar = listborrar & "," & rsBorrar("empleado")
                'armar cartel de aviso de los que se borraran
    rsBorrar.MoveNext
 Loop
rsBorrar.Close
Set rsBorrar = Nothing

If Trim(listborrar) = "" Then
    listborrar = 0
Else
    listborrar = "0" & listborrar
End If

Flog.writeline " Lista de empleados a borrar. " & listborrar
Flog.writeline "    "

StrSql = "SELECT evacab.empleado "
StrSql = StrSql & " FROM evacab "
StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evaevenro = evacab.evaevenro "
StrSql = StrSql & "         AND evaevento.evaproynro =" & evaproynro
StrSql = StrSql & " WHERE evacab.empleado IN (" & listborrar & ")"
'StrSql = StrSql & " AND   evacab.evaevenro =" & evaevenro
OpenRecordset StrSql, rsBorrar2

Do Until rsBorrar2.EOF
    Empleado = rsBorrar2("empleado")
    Flog.writeline "    Empleado a borrar (ternro): " & Empleado
    
        'borra_evaluacion_00.asp?llamadora=relacionar&empleado=<%=l_rs1("empleado")%>&evaevenro=<%=l_evaevenro%>','',50,50);
    StrSql = "SELECT evasecc.orden, evasecc.evaseccnro, evatiposecc.tipsecprogdel  "
    StrSql = StrSql & " FROM evadet "
    StrSql = StrSql & " INNER JOIN evasecc     ON evadet.evaseccnro=evasecc.evaseccnro "
    StrSql = StrSql & " INNER JOIN evatiposecc ON evasecc.tipsecnro=evatiposecc.tipsecnro "
    StrSql = StrSql & " INNER JOIN evacab      ON evacab.evacabnro=evadet.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado  = " & Empleado
    StrSql = StrSql & " GROUP BY evasecc.orden, evasecc.evaseccnro, evatiposecc.tipsecprogdel , evacab.evacabnro "
    StrSql = StrSql & " ORDER BY orden DESC "
    OpenRecordset StrSql, rs1
    Do Until rs1.EOF
        
        tipsecprogdel = Trim(rs1("tipsecprogdel"))
        Flog.writeline "    Programa de Borrado." & tipsecprogdel
        
        If Len(Trim(tipsecprogdel)) <> 0 Then
            Select Case tipsecprogdel
            
            Case "borra_areacom_eva_00.asp":
                Call Borra_areacom(evaevenro, Empleado)
            
            Case "borra_resultados_eva_00.asp":
                Call Borra_resultados(evaevenro, Empleado)
                
            Case "borra_planaccion_eva_00.asp":
                Call Borra_planaccion(evaevenro, Empleado)
                
            Case "borra_objetivos_00.asp":
                Call Borra_objetivos(evaevenro, Empleado)
            
            Case "borra_notas_eva_00.asp":
                Call Borra_notas(evaevenro, Empleado)
            
            Case "borra_vistos_eva_00.asp":
                Call Borra_vistos(evaevenro, Empleado)
            
            'por ahora para CODELCO unicamente ---------------------
            Case "borra_cierre_COD_eva_00.asp":
                Call Borra_cierre(evaevenro, Empleado)
            Case "borra_borrador_COD_eva_00.asp":
                Call Borra_borrador(evaevenro, Empleado)
            
            'por ahora para Deloitte unicamente ---------------------
            Case "borra_datosadm_eva_00.asp":
                Call Borra_datosadm(evaevenro, Empleado)
            Case "borra_calificobj_eva_00.asp":
                Call Borra_calificobj(evaevenro, Empleado)
                Call Borra_objetivos(evaevenro, Empleado) ' dado que la seccion tiene evaluacion de objs y calific gral de objs juntas
            Case "borra_objcom_eva_00.asp":
                Call Borra_objcom(evaevenro, Empleado)
            Case "borra_resultadosyarea_eva_00.asp":
                Call Borra_resultadosyarea(evaevenro, Empleado)
            Case "borra_areacom_eva_00.asp":
                Call Borra_areacom(evaevenro, Empleado) ' !!! HACERLO o ve rf con el otro???
            ' secciones Plan desarrollo - deloitte
            Case "borra_datospers_eva_00.asp":
                Call Borra_datospers(evaevenro, Empleado)
                Call Borra_estform(evaevenro, Empleado)
                Call Borra_trabant(evaevenro, Empleado)
                Call Borra_trabdoc(evaevenro, Empleado)
                Call Borra_trabfirm(evaevenro, Empleado)
            Case "borra_plandesa_eva_00.asp":
                Call Borra_plandesa(evaevenro, Empleado)
                Call Borra_notas(evaevenro, Empleado)
            Case "borra_idioma_eva_00.asp":
                Call Borra_idioma(evaevenro, Empleado)
                Call Borra_notas(evaevenro, Empleado)
                
            Case "borra_subcompxestr_eva_00.asp":
                Call Borra_subcompxestr(evaevenro, Empleado) ' evasubfresu
            Case "borra_compxestr_CH_eva_00.asp", "borra_compxestr_CH_eva_00.asp":
                Call Borra_areacom(evaevenro, Empleado)  ' borra los cometarios de area
                Call Borra_areaycompxestr(evaevenro, Empleado) ' evaarea  evaresultado evagruporesu
            ' Case "borra_compxestr_CH_eva_00.asp":
            '   Call Borra_grupocompxestr(evaevenro, Empleado) ' evagruporesu
           
            Case "borra_vistos_CH_eva_00.asp":
                Call Borra_vistosyCalifGral(evaevenro, Empleado) ' evavistos evavistoresu
                
            Case "borra_datosadmnotas_eva_00.asp":
                Call Borra_notas(evaevenro, Empleado)
                Call Borra_datosadm(evaevenro, Empleado)
            Case "borra_calificgralnota_eva_00.asp":
                Call Borra_calificgral(evaevenro, Empleado)
                Call Borra_notas(evaevenro, Empleado)
                
            Case "borra_areas_eva_00.asp":  ' evaarea y evaresultado
                Call Borra_areas(evaevenro, Empleado)
   
            Case "borra_evaluaciongral_eva_00.asp": ' evagralresu
                Call Borra_evaluaciongral(evaevenro, Empleado)
                Call Borra_notas(evaevenro, Empleado)
            
            Case "borra_objetivosII_eva_00.asp":
                Call Borra_objetivosII(evaevenro, Empleado)
                
            Case "borra_compcom_eva_00.asp":
                Call Borra_compcom(evaevenro, Empleado)
            End Select
            
        End If
        rs1.MoveNext
    Loop
    
    rs1.Close
    Set rs1 = Nothing
    
    Flog.writeline "    Borra la Cabecera de evaluacion."
    Call Borrar_cabecera(evaevenro, Empleado)
    
    rsBorrar2.MoveNext
Loop
rsBorrar2.Close
Set rsBorrar2 = Nothing

Exit Sub

ME_Borrar:
    Flog.writeline "    Error - Empleado: " & Empleado
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub



'_______________________________________________________________________
'--------------------------------------------------------------------
' GENERAR TODAS LAS TABLAS DE EVALUACION
'--------------------------------------------------------------------
Sub generarFormulario(ByVal evaevenro, ByVal ternro)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsSecc    As New ADODB.Recordset
Dim rs1   As New ADODB.Recordset
Dim rs2   As New ADODB.Recordset
Dim rs3   As New ADODB.Recordset
Dim rs4   As New ADODB.Recordset

Dim evacabnro  As Long
Dim evaseccnro As Long
Dim tipsecobj  As Integer
Dim evatevnro As Integer
Dim evarolaspdet As String

Dim evaetanro As Integer
Dim empreporta As Long

Dim habilitado As Integer
Dim evaseccmail As String
Dim nuevo As Integer

Dim evaluador
Dim fechahab
Dim horahab As String

Dim aprobada As Integer

Dim hora As String
Dim arrhr(2)
        
Dim evldrnro As Long
        
Dim crearEvldor As Integer

    
On Error GoTo MError
     ' en el fuente : generar_00
' falat DATOS PROY????????  ----
' pREG POR EL evento  Y SI NO EXISTE HAY QUE CREARLO!!!!!!!!!!!!
' buscar frorm asoc al depto y al perfil del empl?????????

' ACA supone que ya existe el evento ---------

'------------------------------------------------------------------
'Busco si exste ya la cabecera
'------------------------------------------------------------------

StrSql = "SELECT evacab.evacabnro, evaetanro, cabaprobada "
StrSql = StrSql & " FROM evacab "
StrSql = StrSql & " WHERE evacab.empleado = " & ternro
StrSql = StrSql & " AND   evacab.evaevenro = " & evaevenro
OpenRecordset StrSql, rsConsult

If rsConsult.EOF Then
    rsConsult.Close
    Set rsConsult = Nothing
    
    'INSERTAR CABECERA""""""""""""""""""""""""""""""""""""""""""""
    Flog.writeline "        Inserta Cabecera de Evaluacion."
    StrSql = "INSERT INTO evacab "
    StrSql = StrSql & " (evaevenro, empleado, cabevaluada, cabaprobada,cabobservacion, evaproynro,evaetanro) "
    StrSql = StrSql & " VALUES (" & evaevenro & ", " & ternro & ", 0, 0, null," & evaproynro & ",NULL)"
    
    objConn.Execute StrSql, , adExecuteNoRecords

    evacabnro = getLastIdentity(objConn, "evacab")
    
    
        Call setearEtapa("", evaevenro, evacabnro)
    
        'Buscar para crear evadet
        StrSql = " SELECT evaseccnro, tipsecobj "
        StrSql = StrSql & " FROM evasecc "
        StrSql = StrSql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro= evasecc.evatipnro"
        StrSql = StrSql & " INNER JOIN evaevento  ON evaevento.evatipnro = evatipoeva.evatipnro "
        StrSql = StrSql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro= evasecc.tipsecnro "
        StrSql = StrSql & " WHERE evaevenro = " & evaevenro
        OpenRecordset StrSql, rsSecc
        Do Until rsSecc.EOF
            evaseccnro = rsSecc("evaseccnro")
            tipsecobj = rsSecc("tipsecobj")
            
            Flog.writeline "        Crea los registros de Evaluacion para los evaluadores en la seccion: " & evaseccnro

       
            'INSERTAR EVADET """"""""""""""""""""""""""""""""""""""""""""
            StrSql = "INSERT INTO evadet "
            StrSql = StrSql & " (evacabnro , evaseccnro, detcargada) "
            StrSql = StrSql & " VALUES (" & evacabnro & ", " & evaseccnro & ", 0)"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'buscar evaluadores
            StrSql = "SELECT evaoblieva.evatevnro, evatevobli, evarolaspdet, afteranterior, evaobliorden , evaseccmail,evasecc.ultimasecc "
            StrSql = StrSql & " FROM evaoblieva "
            StrSql = StrSql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro= evaoblieva.evatevnro "
            StrSql = StrSql & " INNER JOIN evasecc ON evasecc.evaseccnro= evaoblieva.evaseccnro "
            StrSql = StrSql & " LEFT  JOIN evarolasp ON evarolasp.evarolnro = evatipevalua.evarolnro "
            StrSql = StrSql & " WHERE evaoblieva.evaseccnro = " & evaseccnro
            StrSql = StrSql & " ORDER BY evaobliorden "
            OpenRecordset StrSql, rsConsult
                        
            Do Until rsConsult.EOF
                crearEvldor = -1
                evaluador = ""
                evatevnro = rsConsult("evatevnro")
                If Not IsNull(rsConsult("evarolaspdet")) Then
                    evarolaspdet = Trim(rsConsult("evarolaspdet")) 'ASP que busca el ternro del evaluador
                Else
                    evarolaspdet = ""
                End If

                If rsConsult("ultimasecc") = -1 Then ' si es la ultima seccion no habilito ningun evadetevldor.
                   habilitado = 0
                Else
                    If rsConsult("afteranterior") = -1 Then
                        habilitado = 0
                    Else
                        habilitado = -1
                    End If
                End If
                evaseccmail = rsConsult("evaseccmail")
                
                
                'verificar que no exista ya evldrnro para otro objetivo de la misma evacab
                StrSql = "SELECT * "
                StrSql = StrSql & " FROM evadetevldor "
                StrSql = StrSql & " INNER JOIN evaluaobj ON evaluaobj.evldrnro = evadetevldor.evldrnro"
                StrSql = StrSql & "        AND evaluaobj.evaborrador = 0 "
                StrSql = StrSql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro"
                StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                StrSql = StrSql & " AND   evadetevldor.evatevnro = " & evatevnro
                OpenRecordset StrSql, rs1
                
                StrSql = "SELECT * "
                StrSql = StrSql & " FROM evadetevldor "
                StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                StrSql = StrSql & " AND   evadetevldor.evatevnro = " & evatevnro
                StrSql = StrSql & " AND   evadetevldor.evaseccnro = " & evaseccnro
                OpenRecordset StrSql, rs2
                nuevo = 0
                
                If (tipsecobj = -1 And rs1.EOF And rs2.EOF) Or (tipsecobj = 0 And rs2.EOF) Then
                    rs1.Close
                    Set rs1 = Nothing
                    rs2.Close
                    Set rs2 = Nothing
                
                    horahab = ""
                    fechahab = "NULL"

                   If habilitado = -1 Then
                        hora = Mid(Time, 1, 8)
                        hora = strto2(Left(hora, 2)) & Right(hora, 2)
                        fechahab = ConvFecha(Date)
                        horahab = hora
                    End If
                    
                    
                    Call buscarEvaluador(evarolaspdet, ternro, evaluador)
 
                     If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                        StrSql = "SELECT evaluador "
                        StrSql = StrSql & " FROM evadetevldor "
                        StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                        StrSql = StrSql & " AND   evadetevldor.evaseccnro = " & evaseccnro
                        StrSql = StrSql & " AND   evadetevldor.evatevnro <> " & evatevnro
                        StrSql = StrSql & " AND   evadetevldor.evaluador = " & evaluador
                        OpenRecordset StrSql, rs1
                        If Not rs1.EOF Then
                            crearEvldor = 0
                        End If
                        rs1.Close
                    End If
                                       
                    If crearEvldor = -1 Then
                        StrSql = "INSERT INTO evadetevldor "
                        StrSql = StrSql & "(evacabnro , evaseccnro, evatevnro, "
                        If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                            StrSql = StrSql & " evaluador,"
                        End If
                        StrSql = StrSql & " evldorcargada,habilitado,fechahab,horahab) "
                        StrSql = StrSql & " VALUES (" & evacabnro & ", "
                        StrSql = StrSql & evaseccnro & ", " & evatevnro
                        If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                            StrSql = StrSql & "," & evaluador
                        End If
                        StrSql = StrSql & ", 0,"
                        StrSql = StrSql & habilitado & ","
                        StrSql = StrSql & fechahab & ",'"
                        StrSql = StrSql & horahab & "')"
                                                                                    
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        nuevo = -1
                            
                        evldrnro = getLastIdentity(objConn, "evadetevldor")
                    End If
                    
                Else
                    evldrnro = rs1("evldrnro")
                    evaluador = rs1("evaluador")
                    rs1.Close
                    Set rs1 = Nothing
                    rs2.Close
                    Set rs2 = Nothing
                    
                End If
                
                'si se pierde el evldrnro... ha pasado
                StrSql = "SELECT evldrnro, habilitado "
                StrSql = StrSql & " FROM evadetevldor "
                StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                StrSql = StrSql & " AND   evadetevldor.evatevnro = " & evatevnro
                StrSql = StrSql & " AND   evadetevldor.evaseccnro = " & evaseccnro
                OpenRecordset StrSql, rs2
                If Not rs2.EOF Then
                    evldrnro = rs2("evldrnro")
                    habilitado = rs2("habilitado")
                End If
                rs2.Close
                Set rs2 = Nothing
                
                
                rsConsult.MoveNext
            Loop
            rsConsult.Close
            Set rsConsult = Nothing
            
            
            rsSecc.MoveNext
        Loop
        rsSecc.Close
        Set rsSecc = Nothing
Else
    ' ya existe el evacab...
     Flog.writeline "        El empleado tiene cabecera de Evaluacion."
     evacabnro = rsConsult("evacabnro")
     aprobada = rsConsult("cabaprobada")

     Call setearEtapa(rsConsult("evaetanro"), evaevenro, evacabnro)
     
    'creo evadet y evadetevldor para secciones nuevas....
     StrSql = " SELECT evaseccnro, tipsecobj "
     StrSql = StrSql & " FROM evasecc "
     StrSql = StrSql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro= evasecc.evatipnro"
     StrSql = StrSql & " INNER JOIN evaevento  ON evaevento.evatipnro = evatipoeva.evatipnro "
     StrSql = StrSql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro= evasecc.tipsecnro "
     StrSql = StrSql & " WHERE evaevenro = " & evaevenro
     OpenRecordset StrSql, rsSecc
     Do Until rsSecc.EOF
         evaseccnro = rsSecc("evaseccnro")
         tipsecobj = rsSecc("tipsecobj")
         Flog.writeline "        Crea los registros de Evaluacion para los evaluadores en la seccion " & evaseccnro

        StrSql = " SELECT evaseccnro "
        StrSql = StrSql & " FROM evadet "
        StrSql = StrSql & " WHERE evacabnro = " & evacabnro
        StrSql = StrSql & " AND  evaseccnro = " & evaseccnro
        OpenRecordset StrSql, rs2
        If rs2.EOF Then
            StrSql = "INSERT INTO evadet "
            StrSql = StrSql & " (evacabnro , evaseccnro, detcargada) "
            StrSql = StrSql & " VALUES (" & evacabnro & ", " & evaseccnro & ", 0)"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rs2.Close
        Set rs2 = Nothing
        
        ' Crear los evadetevldor ...............................................
        StrSql = "SELECT evaoblieva.evatevnro, evatevobli, evarolaspdet, afteranterior, evaobliorden , evaseccmail, evasecc.ultimasecc "
        StrSql = StrSql & " FROM evaoblieva "
        StrSql = StrSql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro= evaoblieva.evatevnro "
        StrSql = StrSql & " INNER JOIN evasecc ON evasecc.evaseccnro= evaoblieva.evaseccnro "
        StrSql = StrSql & " LEFT  JOIN evarolasp ON evarolasp.evarolnro = evatipevalua.evarolnro "
        StrSql = StrSql & " WHERE evaoblieva.evaseccnro = " & evaseccnro
        StrSql = StrSql & " ORDER BY evaobliorden "
        OpenRecordset StrSql, rsConsult
        Do Until rsConsult.EOF
               crearEvldor = -1
                evaluador = ""
                evatevnro = rsConsult("evatevnro")
                
               If Not IsNull(rsConsult("evarolaspdet")) Then
                    evarolaspdet = rsConsult("evarolaspdet") 'ASP que busca el ternro del evaluador
                Else
                    evarolaspdet = ""
                End If
                
                evaseccmail = rsConsult("evaseccmail")
                
                If aprobada = -1 Then ' si la cabecera esta aprobada no habilito ningun avedetevldor.
                    habilitado = 0
                Else
                    If rsConsult("ultimasecc") = -1 Then ' si es la ultima seccion no habilito ningun evadetevldor.
                        habilitado = 0
                    Else
                        If rsConsult("afteranterior") = -1 Then
                            habilitado = 0
                        Else
                            habilitado = -1
                        End If
                    End If
                End If
                
                'verificar que no exista ya evldrnro para otro objetivo de la misma evacab
                ' 14-01-2008 - AGREGAR que sea de la seccion para la que se quiere un nuevo rol!!
                '               VERIFICAR QUE ESTE BIEN ESTE NUEVO AGREGADO!
                StrSql = "SELECT * "
                StrSql = StrSql & " FROM evadetevldor "
                StrSql = StrSql & " INNER JOIN evaluaobj ON evaluaobj.evldrnro = evadetevldor.evldrnro"
                StrSql = StrSql & "        AND evaluaobj.evaborrador = 0 "
                StrSql = StrSql & " INNER JOIN evaobjetivo ON evaobjetivo.evaobjnro = evaluaobj.evaobjnro"
                StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                StrSql = StrSql & " AND   evadetevldor.evatevnro = " & evatevnro
                StrSql = StrSql & " AND   evadetevldor.evaseccnro = " & evaseccnro
                OpenRecordset StrSql, rs2
                
                StrSql = "SELECT * "
                StrSql = StrSql & " FROM evadetevldor "
                StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                StrSql = StrSql & " AND   evadetevldor.evatevnro = " & evatevnro
                StrSql = StrSql & " AND   evadetevldor.evaseccnro = " & evaseccnro
                OpenRecordset StrSql, rs3
                nuevo = 0
                If (tipsecobj = -1 And rs2.EOF And rs3.EOF) Or (tipsecobj = 0 And rs3.EOF) Then
                
                    rs2.Close
                    Set rs2 = Nothing
                    rs3.Close
                    Set rs3 = Nothing
                
                    horahab = ""
                    fechahab = "NULL"

                    If habilitado = -1 Then
                        hora = Mid(Time, 1, 8)
                        hora = strto2(Left(hora, 2)) & Right(hora, 2)
                        fechahab = ConvFecha(Date)
                        horahab = hora
                    End If
                    
                    
                   Call buscarEvaluador(evarolaspdet, ternro, evaluador)
                                
                    If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                        StrSql = "SELECT evaluador "
                        StrSql = StrSql & " FROM evadetevldor "
                        StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                        StrSql = StrSql & " AND   evadetevldor.evaseccnro = " & evaseccnro
                        StrSql = StrSql & " AND   evadetevldor.evatevnro <> " & evatevnro
                        StrSql = StrSql & " AND   evadetevldor.evaluador = " & evaluador
                        OpenRecordset StrSql, rs1
                        If Not rs1.EOF Then
                            crearEvldor = 0
                        End If
                        rs1.Close
                    End If
                                
                    If crearEvldor = -1 Then
                        StrSql = "INSERT INTO evadetevldor "
                        StrSql = StrSql & "(evacabnro , evaseccnro, evatevnro, "
                        If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                        StrSql = StrSql & " evaluador,"
                        End If
                        StrSql = StrSql & " evldorcargada,habilitado,fechahab,horahab) "
                        StrSql = StrSql & " VALUES (" & evacabnro & ", "
                        StrSql = StrSql & evaseccnro & ", " & evatevnro
                        If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                        StrSql = StrSql & "," & evaluador
                        End If
                        StrSql = StrSql & ", 0,"
                        StrSql = StrSql & habilitado & ","
                        StrSql = StrSql & fechahab & ",'"
                        StrSql = StrSql & horahab & "')"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        evldrnro = getLastIdentity(objConn, "evadetevldor")
                        
                        nuevo = -1
                    End If
                Else
                    If Not rs3.EOF Then
                        evldrnro = rs3("evldrnro")
                        evaluador = rs3("evaluador")
                    Else
                        If Not rs2.EOF Then
                            evldrnro = rs2("evldrnro")
                            evaluador = rs2("evaluador")
                        End If
                    End If
                    rs2.Close
                    Set rs2 = Nothing
                    rs3.Close
                    Set rs3 = Nothing
                End If
                    
            ' se pierde el evldrnro....
            StrSql = "SELECT evldrnro, habilitado "
            StrSql = StrSql & " FROM evadetevldor "
            StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
            StrSql = StrSql & " AND   evadetevldor.evatevnro = " & evatevnro
            StrSql = StrSql & " AND   evadetevldor.evaseccnro = " & evaseccnro
            OpenRecordset StrSql, rs3
            If Not rs3.EOF Then
                evldrnro = rs3("evldrnro")
                habilitado = rs3("habilitado")
            End If
            rs3.Close
            Set rs3 = Nothing
            
            rsConsult.MoveNext
            Loop
            
            rsConsult.Close
            Set rsConsult = Nothing
            
            rsSecc.MoveNext
        Loop ' rsSecc
        rsSecc.Close
        Set rsSecc = Nothing
            
End If ' de select evacab

    
Exit Sub

MError:
    Flog.writeline "       Error en el tercero " & ternro & " Error: " & Err.Description
    Flog.writeline "       Ultimo SQL Ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub


'____________________________________________________________________
' Setea a Etapa la priemr etapa al Evaluado, si no existe
'____________________________________________________________________
Sub setearEtapa(evaetanro, evaevenro, evacabnro)
Dim StrSql As String
Dim rs2 As New ADODB.Recordset

On Error GoTo ME_Etapa
    
    If Trim(evaetanro) = "" Or IsNull(evaetanro) Then
       'buscar la ETAPA
       StrSql = "SELECT evaforeta.evaetanro "
       StrSql = StrSql & " FROM evaforeta "
       StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evatipnro= evaforeta.evatipnro"
       StrSql = StrSql & " WHERE evaforeta.evadef = -1"
       StrSql = StrSql & " AND   evaevento.evaevenro = " & evaevenro
       OpenRecordset StrSql, rs2
       If Not rs2.EOF Then
           evaetanro = rs2("evaetanro")
       Else
           evaetanro = ""
       End If
       rs2.Close
       Set rs2 = Nothing
    
       If Len(Trim(evaetanro)) <> 0 And evaetanro <> 0 Then
           StrSql = "UPDATE evacab SET "
           StrSql = StrSql & " evaetanro= " & evaetanro
           StrSql = StrSql & " WHERE evacabnro = " & evacabnro
           objConn.Execute StrSql, , adExecuteNoRecords
       End If
       
    End If
    
Exit Sub

ME_Etapa:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub





' Call buscarEvaluador (evarolaspdet,ternro, evaluador)' byref evauado - emreporte
' ____________________________________________________________________
'  busca el Ternro de Evaluador en base al asp
' ____________________________________________________________________
Sub buscarEvaluador(evarolaspdet, ternro, ByRef evaluador)
Dim rs3   As New ADODB.Recordset
Dim rs4   As New ADODB.Recordset
Dim empreporta

On Error GoTo ME_Ev


' buscar empreporta
StrSql = "SELECT empreporta "
StrSql = StrSql & " FROM empleado"
StrSql = StrSql & " WHERE ternro= " & ternro
OpenRecordset StrSql, rs3
If Not rs3.EOF Then
    If rs3("empreporta") <> 0 And Not IsNull(rs3("empreporta")) Then
        empreporta = rs3("empreporta")
    Else
        empreporta = 0
    End If
End If
rs3.Close
Set rs3 = Nothing


If Len(Trim(evarolaspdet)) <> 0 Then

    Select Case evarolaspdet

        Case "buscar_auto_eva.asp":
             evaluador = ternro
                                 
        Case "buscar_revisor_eva.asp":
            If empreporta <> 0 Then
                evaluador = empreporta
            End If

        Case "buscar_garante_eva.asp":
            'hay que buscar por tipoestructura Garante
            StrSql = "SELECT estrnro "
            StrSql = StrSql & " FROM his_estructura "
            StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro
            StrSql = StrSql & "   AND his_estructura.htethasta IS NULL "
            StrSql = StrSql & "   AND his_estructura.tenro =" & ctenroarea
            OpenRecordset StrSql, rs3
            If Not rs3.EOF Then
                'Buscar un garante con esta area
                 StrSql = "SELECT his_estructura.ternro "
                 StrSql = StrSql & " FROM his_estructura "
                 StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = his_estructura.ternro "
                 StrSql = StrSql & " INNER JOIN his_estructura area ON his_estructura.ternro = area.ternro "
                 StrSql = StrSql & "        AND area.tenro   = " & ctenroarea
                 StrSql = StrSql & "        AND area.estrnro = " & rs3("estrnro")
                 StrSql = StrSql & " WHERE his_estructura.tenro = " & ctenrogarante
                 StrSql = StrSql & " AND   his_estructura.htethasta IS NULL "
                 OpenRecordset StrSql, rs4
                 If Not rs4.EOF Then
                     evaluador = rs4("ternro")
                     rs4.Close
                     Set rs4 = Nothing
                 End If
            End If
            rs3.Close


        Case "buscar_proyrevisor_eva.asp":
                            
             perfil = "Empleado"
             StrSql = "SELECT proysocio, proygerente, proyrevisor "
             StrSql = StrSql & " FROM evaproyecto "
             StrSql = StrSql & " WHERE evaproynro = " & evaproynro
             OpenRecordset StrSql, rs3
             If Not rs3.EOF Then
                If ternro = rs3("proygerente") Then
                  perfil = "GERENTE"
                Else
                    If ternro = rs3("proyrevisor") Then
                        perfil = "REVISOR"
                    Else
                        perfil = "Empleado"
                    End If
                End If
             End If
             rs3.Close
             Set rs3 = Nothing
             Call buscarEvaluadorProy(evaproynro, perfil, "revisor", evaluador)
            
       Case "buscar_proygerente_eva.asp":
               Call buscarEvaluadorProy(evaproynro, perfil, "gerente", evaluador)
       Case "buscar_proysocio_eva.asp":
                Call buscarEvaluadorProy(evaproynro, perfil, "socio", evaluador)
       End Select
       
 End If ' de evarolaspdet <> ""

Exit Sub

ME_Ev:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub



' ____________________________________________________________________
'  busca el Ternro de Evaluador segun la definicion del proyecto
' ____________________________________________________________________
Sub buscarEvaluadorProy(evaproynro, perfil, tipoEvldor, ByRef evaluador)

Dim StrSql As String
Dim rs5 As New ADODB.Recordset
                                        
StrSql = "SELECT proyrevisor "
                                        
Select Case tipoEvldor
        Case "revisor":
                'If Trim(UCase(perfil)) = "GERENTE" Then
                If Trim(perfil) = "GERENTE" Then
                        StrSql = "SELECT proysocio "
                Else
                    If Trim(perfil) = "REVISOR" Then
                        StrSql = "SELECT proygerente "
                    Else
                        StrSql = "SELECT proyrevisor "
                    End If
                        
                End If
        Case "gerente":
                StrSql = "SELECT proygerente "
        Case "socio":
                StrSql = "SELECT proysocio "
End Select
                                
StrSql = StrSql & " FROM evaproyecto "
StrSql = StrSql & " INNER JOIN evaproyemp ON evaproyecto.evaproynro=evaproyemp.evaproynro "
' StrSql = StrSql & "             AND evaproyemp.ternro = " & ternro
StrSql = StrSql & " WHERE evaproyecto.evaproynro = " & evaproynro
OpenRecordset StrSql, rs5
                                        
If Not rs5.EOF Then
   evaluador = rs5(0)
End If
                                        
rs5.Close
Set rs5 = Nothing

End Sub





'________________________________________________________________________________________
'________________________________________________________________________________________
' BORRA DATOS DE LAS SECCIONES Y EVADETELVDOR Y EVACAB
'________________________________________________________________________________________
'________________________________________________________________________________________

' _______________________________________________________
Sub Borrar_cabecera(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE  FROM evadetevldor WHERE evadetevldor.evacabnro IN "
    StrSql = StrSql & " (SELECT evacabnro FROM evacab WHERE "
    StrSql = StrSql & " evacab.evaevenro  = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = "DELETE FROM evadet WHERE evadet.evacabnro IN "
    StrSql = StrSql & " (SELECT evacabnro FROM evacab WHERE "
    StrSql = StrSql & " evacab.evaevenro  = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'BORRAR cabecera
    StrSql = "DELETE "
    StrSql = StrSql & " FROM evacab WHERE "
    StrSql = StrSql & " evaevenro= " & evaevenro
    StrSql = StrSql & " AND empleado = " & Empleado
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub
'__________________________________________________________________________
'__________________________________________________________________________

Sub Borra_areacom(evavenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
        
    StrSql = "DELETE FROM evaareacom  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaareacom.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
        
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_resultadosyarea(evavenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    'borra los datos de los resultados de las competencias
    StrSql = "DELETE FROM evaresultado"
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaresultado.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
       
    'borra los datos del area de las competencias
    StrSql = "DELETE FROM evaarea"
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaarea.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub
Sub Borra_resultados(evaevenro, Empleado)
    
    Dim StrSql As String
    Dim rs1 As New ADODB.Recordset
    Dim lista
    On Error GoTo ME_Borrar
    
    StrSql = " select evadetevldor.evldrnro from evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado = " & Empleado
    OpenRecordset StrSql, rs1
    lista = "0"
    Do Until rs1.EOF
        lista = lista & "," & rs1("evldrnro")
        rs1.MoveNext
     Loop
     rs1.Close
     Set rs1 = Nothing
     
     StrSql = "DELETE FROM evaresultado WHERE evaresultado.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
     
     objConn.Execute StrSql, , adExecuteNoRecords
    
     StrSql = "DELETE FROM evaarea WHERE evaarea.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
     
     objConn.Execute StrSql, , adExecuteNoRecords

    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub
Sub Borra_planaccion(evaevenro, Empleado)
    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evaplan WHERE evaplan.evldrnro IN "
    StrSql = StrSql & " (SELECT evadetevldor.evldrnro FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub
Sub Borra_objetivos(evaevenro, Empleado)
    
    Dim StrSql As String
    Dim rs1 As New ADODB.Recordset
    Dim lista
    On Error GoTo ME_Borrar
    
    StrSql = " select evaluaobj.evaobjnro from evadetevldor "
    StrSql = StrSql & " INNER JOIN evaluaobj ON evaluaobj.evldrnro= evadetevldor.evldrnro"
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado = " & Empleado
    OpenRecordset StrSql, rs1
    lista = "0"
    Do Until rs1.EOF
        lista = lista & "," & rs1("evaobjnro")
        rs1.MoveNext
     Loop
     rs1.Close
     Set rs1 = Nothing
     
    'borrar todos los resultados de objetivos (tiene un trnro asociado)
    StrSql = "DELETE FROM evaluaobj  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaluaobj.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
 
    'borrar todos los planes smart si hay alguno
    StrSql = "DELETE FROM evaplan WHERE evaplan.evldrnro IN "
    StrSql = StrSql & " (SELECT evadetevldor.evldrnro FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = "DELETE FROM evaplan WHERE evaplan.evaobjnro IN "
    StrSql = StrSql & " (" & lista & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Borrar todos los comentarios del objetivo que ya de ols EVADETEVLDOR
    StrSql = "DELETE FROM evaobjsgto  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaobjsgto.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Borrar todos los puntajes de la evaluacion, que son de objetivos obviamente.
    StrSql = "DELETE FROM evapuntaje  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evacabnro=evapuntaje.evacabnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Borrar el puntaje de objetivos General (Deloitte)
    StrSql = "DELETE FROM evagralobj  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evagralobj.evldrnro "
    StrSql = StrSql & " AND   evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado  = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
 
     ' borra los comentarios hechos acerca del objetivo
    StrSql = "DELETE FROM evaobjcom WHERE evaobjcom.evaobjnro IN "
    StrSql = StrSql & " (" & lista & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    'Borrar objetivos sin
    StrSql = "DELETE FROM evaobjetivo  WHERE evaobjetivo.evaobjnro IN "
    StrSql = StrSql & " (" & lista & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_objcom(evavenro, Empleado)

    'para Deloitte por ahora
    
    Dim StrSql As String
    On Error GoTo ME_Borrar
 
    StrSql = "DELETE FROM evaobjcom  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaobjcom.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub
Sub Borra_notas(evavenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
 
    StrSql = "DELETE FROM evanotas WHERE evanotas.evldrnro IN "
    StrSql = StrSql & " (SELECT evadetevldor.evldrnro FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub
Sub Borra_calificobj(evavenro, Empleado)
    
    'para Deloitte por ahora
    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evagralobj  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evagralobj.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub
Sub Borra_vistos(evavenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
 
    StrSql = "DELETE FROM evavistos WHERE evavistos.evldrnro IN "
    StrSql = StrSql & " (SELECT evadetevldor.evldrnro FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub
Sub Borra_datosadm(evavenro, Empleado)

    'Por ahora para deloitte unicamente
    
    Dim StrSql As String
    On Error GoTo ME_Borrar
 
    StrSql = "DELETE FROM evadatosadm  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evadatosadm.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado  = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
 
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub
Sub Borra_borrador(evavenro, Empleado)

    'Por ahora para CODELCO unicamente
    Dim StrSql As String
    Dim rs1 As New ADODB.Recordset
    Dim lista
    On Error GoTo ME_Borrar
 
    StrSql = " (select evaluaobjborr.evaobjborrnro from evadetevldor "
    StrSql = StrSql & " INNER JOIN evaluaobjborr ON evaluaobjborr.evldrnro= evadetevldor.evldrnro"
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado  = " & Empleado & ")"
    OpenRecordset StrSql, rs1
    lista = "0"
    Do Until rs1.EOF
        lista = lista & "," & rs1("evaobjborrnro")
        rs1.MoveNext
     Loop
     rs1.Close
     Set rs1 = Nothing
     
    'borrar todos los resultados de objetivos (tiene un trnro asociado)
    StrSql = "DELETE FROM evaluaobjborr  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaluaobjborr.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
 
   'Borrar objetivos borrador
    StrSql = "DELETE FROM evaobjborr WHERE evaobjborr.evaobjborrnro IN "
    StrSql = StrSql & " (" & lista & ")"

Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_cierre(evavenro, Empleado)

    'Por ahora para CODELCO unicamente
    Dim StrSql As String
    On Error GoTo ME_Borrar
 
    StrSql = "DELETE FROM evacierre WHERE evacierre.evldrnro IN "
    StrSql = StrSql & " (SELECT evadetevldor.evldrnro FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND   evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub



Sub Borra_datospers(evavenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evadatosper  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evacab  "
    StrSql = StrSql & " WHERE  evacab.evacabnro=evadatosper.evacabnro "
    StrSql = StrSql & "     AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


Sub Borra_estform(evavenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evaestform  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evacab  "
    StrSql = StrSql & " WHERE  evacab.evacabnro=evaestform.evacabnro "
    StrSql = StrSql & "     AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


Sub Borra_trabant(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evatrabant  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evacab  "
    StrSql = StrSql & " WHERE  evacab.evacabnro=evatrabant.evacabnro "
    StrSql = StrSql & "     AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


Sub Borra_trabdoc(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evatrabdoc  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evacab  "
    StrSql = StrSql & " WHERE  evacab.evacabnro=evatrabdoc.evacabnro "
    StrSql = StrSql & "     AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


Sub Borra_trabfirm(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evatrabfirma  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evacab  "
    StrSql = StrSql & " WHERE  evacab.evacabnro = evatrabfirma.evacabnro "
    StrSql = StrSql & "     AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


Sub Borra_plandesa(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evapldesaresu  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE  evadetevldor.evldrnro=evapldesaresu.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords


    StrSql = "DELETE FROM evaplandesa  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evacab  "
    StrSql = StrSql & " WHERE  evacab.evacabnro = evaplandesa.evacabnro "
    StrSql = StrSql & "     AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub



Sub Borra_idioma(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evaidiresu  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evacab  "
    StrSql = StrSql & " WHERE  evacab.evacabnro = evaidiresu.evacabnro "
    StrSql = StrSql & "     AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub



Sub Borra_subcompxestr(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    'borra los datos de los resultados de las sub-competencias
    StrSql = "DELETE FROM evasubfresu "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evasubfresu.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
       

Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


Sub Borra_areaycompxestr(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    'borra los datos de los resultados de las competencias
    StrSql = "DELETE FROM evaresultado"
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaresultado.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
       
    'borra los datos de los resultados de los grupos de competencias
    StrSql = "DELETE FROM evagruporesu "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evagruporesu.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
       
    'borra los datos del area de las competencias
    StrSql = "DELETE FROM evaarea"
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaarea.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub



Sub Borra_grupocompxestr(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    'borra los datos de los resultados de los grupos de competencias
    StrSql = "DELETE FROM evagruporesu "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evagruporesu.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
       
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_vistosyCalifGral(evaevenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    'borra los datos de los resultados de las aprobaciones
    StrSql = "DELETE FROM evavistoresu "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evavistoresu.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
       
    'borra los datos del area de las competencias
    StrSql = "DELETE FROM evavistos "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evavistos.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql


End Sub



Sub Borra_areas(evavenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    ' NO Deberia existir evaluaciones en evaresultado, igual se borra si existe alguna por cambio de seccion de Ev Comp y Areas a Ev. Areas
    ' borra los datos de los resultados de las competencias
    StrSql = "DELETE FROM evaresultado"
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaresultado.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
       
    'borra los datos del area de las competencias
    StrSql = "DELETE FROM evaarea"
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaarea.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub






Sub Borra_evaluaciongral(evavenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
    
      
    'borra los datos del area de las competencias
    StrSql = "DELETE FROM evagralresu "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evagralresu.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub
' _______________________________________________________

' _______________________________________________________
' ______________________________________________________
Sub Borra_objetivosII(evaevenro, Empleado)
    
    Dim StrSql As String
    Dim rs1 As New ADODB.Recordset
    Dim lista
    Dim listapredef
    On Error GoTo ME_BorrarII
    
     
    lista = "0"
    listapredef = "0"
  
    
    ' busca Objetivos definidos para el empleado (version nueva de Objetivos)
    StrSql = " SELECT DISTINCT evaobjetivo.evaobjnro "
    StrSql = StrSql & " FROM evaobjetivo "
    StrSql = StrSql & " INNER JOIN evaobjdet ON evaobjdet.evaobjnro = evaobjetivo.evaobjnro AND evaobjdet.evaobjpredef = 0 "
    StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaobjdet.evldrnro "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado
    OpenRecordset StrSql, rs1
    Do Until rs1.EOF
        lista = lista & "," & rs1("evaobjnro")
        rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
    ' buscar objetivos predefinidos
    StrSql = " SELECT DISTINCT evaobjetivo.evaobjnro "
    StrSql = StrSql & " FROM evaobjetivo "
    StrSql = StrSql & " INNER JOIN evaobjresu ON evaobjresu.evaobjnro = evaobjetivo.evaobjnro "
    StrSql = StrSql & " INNER JOIN evadetevldor ON evaobjresu.evldrnro = evadetevldor.evldrnro "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado
    StrSql = StrSql & "     AND evaobjetivo.evaobjnro NOT IN (" & lista & ")"
    OpenRecordset StrSql, rs1
    Do Until rs1.EOF
        listapredef = listapredef & "," & rs1("evaobjnro")
        rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
    
     
    'borrar todos los resultados de objetivos (tiene un ternro asociado)
    StrSql = "DELETE FROM evaobjresu  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " WHERE evadetevldor.evldrnro=evaobjresu.evldrnro "
    StrSql = StrSql & "     AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
 

    'Borrar todos los comentarios del objetivo que ya
    StrSql = "DELETE FROM evaobjsgto  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evaobjsgto.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = "DELETE FROM evaobjcom  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro= evaobjcom.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

 
     'Borrar los detalles de objetivos
    StrSql = "DELETE FROM evaobjdet  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro= evaobjdet.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
   
    'Borrar objetivos definidos por el usuario -  NO los predefinidos
    StrSql = "DELETE FROM evaobjetivo "
    StrSql = StrSql & " WHERE evaobjetivo.evaobjnro IN (" & lista & ")"
    StrSql = StrSql & "     AND evaobjetivo.evaobjnro NOT IN (" & listapredef & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    'borrar las evaluaciones generales de Objetivos (29-12-2009)
    StrSql = "DELETE FROM evagralobj  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro=evadetevldor.evacabnro"
    StrSql = StrSql & " WHERE evadetevldor.evldrnro=evagralobj.evldrnro "
    StrSql = StrSql & "     AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_BorrarII:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


Sub Borra_compcom(evavenro, Empleado)

    Dim StrSql As String
    On Error GoTo ME_Borrar
        
    StrSql = "DELETE FROM evafaccom  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evafaccom.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
        
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub



' _______________________________________________________

Sub Borra_calificgral(evavenro, Empleado)
    
    'para Deloitte por ahora
    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evagralrdp  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evagralrdp.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
 
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
 
End Sub





                                        

Function strto2(cad)
    If Trim(cad) <> "" Then
        If Len(cad) < 2 Then
            strto2 = "0" & cad
        Else
            strto2 = cad
        End If
    Else
        strto2 = "00"
    End If
End Function









