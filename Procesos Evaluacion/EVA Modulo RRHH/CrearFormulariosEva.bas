Attribute VB_Name = "CrearFormulariosEva"
Option Explicit

'Version 1.00

' 17-05-2006 - LA- Agrego manejo de errores y que se escriba info en el log
' 22-06-2006 - LA - Sacar las vistas de las consultas
' 02-02-2007 - LA - Inicializar Fecha Null para evadetevldor y incluir Comentarios para el log


'Global Const Version = "1.01"
'Global Const FechaModificacion = "19-04-2007" ' Leticia Amadio
'Global Const UltimaModificacion = " " 'Secciones para Plan de Desarrollo - Deloitte
                                      
'Global Const Version = "1.02"
'Global Const FechaModificacion = "30-05-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " 'Asociarle la Etapa al empleado cuando se crea la evaluacion (y no solo cdo existe)
                                      
'Global Const Version = "1.03"
'Global Const FechaModificacion = "31-05-2007 " ' Leticia Amadio - 06-06-2007
'Global Const UltimaModificacion = " " 'Borrar Secciones para CHILE (ref. Deloitte)
    
'Global Const Version = "1.04"
'Global Const FechaModificacion = "22-10-2007 " ' Breglia M
'Global Const UltimaModificacion = " " 'Nuevo rol por estructura para Bco Industrial
                              
'Global Const Version = "1.05"
'Global Const FechaModificacion = "01-11-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Borrar Datos Seccion Evaluacion Areas
                              
'Global Const Version = "1.06"
'Global Const FechaModificacion = "03-12-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Agregar prg de determinacion de Rol Supervisor

'Global Const Version = "1.07"
'Global Const FechaModificacion = "10-12-2007 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Borrar seccion Objetivos - solucionar que borre objetivos se evaluen o no objetivos (buscar objs desde version anterior y nueva sobre objetivos) - se agrego el el borrar de evagralobj

'Global Const Version = "1.08"
'Global Const FechaModificacion = "29-01-2008 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' La solucion anterior (borrar objetivos) es para ITAU - para estandar se comento


'Global Const Version = "1.09"
'Global Const FechaModificacion = "14-05-2008 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Borrar Datos Seccion Evaluacion General

'Global Const Version = "1.10"
'Global Const FechaModificacion = "2-07-2008 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Borrar Datos Seccion Evaluacion de Objetivos

'Global Const Version = "1.11"
'Global Const FechaModificacion = "17-11-2008 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Borrar Datos Seccion Plan Accion MultiVoice

'Global Const Version = "1.12"
'Global Const FechaModificacion = "17-04-2009 " ' Leticia Amadio
'Global Const UltimaModificacion = " " ' Borrar Datos Seccion Comentarios de Competencias y borra del la ultima seccion a la primera

'Global Const Version = "1.13"
'Global Const FechaModificacion = "07-10-2009 " ' Leticia Amadio
'Global Const UltimaModificacion = " " 'CUSTOM CHILE-DELOITTE - eliminar una evaluacion de Grupo Competencia y de Comptencias.

'Global Const Version = "1.14" ' Cesar Stankunas
'Global Const FechaModificacion = "25/11/2009"
'Global Const UltimaModificacion = ""    'Encriptacion de string connection

'Global Const Version = "1.15" ' Leticia A.
'Global Const FechaModificacion = "29/12/2009"
'Global Const UltimaModificacion = ""    'CUSTOM Santander Uruguay - eliminar evaluaciones de areas, objs, ev. grales


'Global Const Version = "1.16" ' FGZ
'Global Const FechaModificacion = "12/01/2010"
'Global Const UltimaModificacion = ""    'CUSTOM Santander Uruguay - Se ajustó el calculo del progreso del proceso


'Global Const Version = "1.17" ' Leti A.
'Global Const FechaModificacion = "19/02/2010"
'Global Const UltimaModificacion = ""   'Se  CUSTOM CHILE-DELOITTE - eliminar evaluaciones de evagral rdp + notas. -En el calculo del progreso se cambia ',' por '.' en el Incremento (por s viene asi)
                                        


'Global Const Version = "1.18" ' Leti A.
'Global Const FechaModificacion = "27/05/2010"
'Global Const UltimaModificacion = ""   ' se modifica la forma de cargar el evaluador, si en alguna seccion se cargo manualmente, para las nuevas secciones y/o roles se carga desde ahí.
                                       ' mejora en tiempos - que se muestre la barra de progreso cuando hay mas de 100 empleados (IncTotal)


'Global Const Version = "1.19" ' Leti A.
'Global Const FechaModificacion = "29/06/2011"
'Global Const UltimaModificacion = ""  ' se integra el proceso de Deloitte RDE - para usar un solo proceso
                                       ' se agregaron modificaciones a RDE (validador, etc.)
                                       ' agregar etapa a los roles - Tabla: evaroleta - evatipoform
                                       ' Agregar prg de asignacion de rol:consejero
                                        
                                        
'Global Const Version = "1.20" ' Leti A.
'Global Const FechaModificacion = "05/01/2012"
'Global Const UltimaModificacion = ""   ' cas-13764 - agregar prgama de borrado de Plan de Accion II
                                       ' modificación en la creacion de cabecera de ev. p que no de problema en getlastidentity
                                       ' se agrego control de versiones
                            
'Global Const Version = "1.21" ' Carmen Quintero.
'Global Const FechaModificacion = "18/06/2012"
'Global Const UltimaModificacion = ""   ' cas-13764 - Se modificó la función generar formulario, para que considere caso
                                       ' presentado al momento de asignar los evaluadores
                            
   
'Global Const Version = "1.22" ' Carmen Quintero.
'Global Const FechaModificacion = "26/06/2012"
'Global Const UltimaModificacion = ""   ' Caso - 14205 - Borrar Datos Seccion Puesto Banda Salarial

'Global Const Version = "1.23" ' Carmen Quintero.
'Global Const FechaModificacion = "17/12/2012"
'Global Const UltimaModificacion = ""   ' (CAS-16940 - AGD - GDD - Sección de Resultados) Borra Datos Seccion Resultados Integrales

'Global Const Version = "1.24" ' Carmen Quintero.
'Global Const FechaModificacion = "05/03/2013"
'Global Const UltimaModificacion = ""   ' (CAS 18630 - Heidt & Asoc - Bug Eliminacion de Evaluaciones) Se separaron las opciones de insertar y eliminar
                                       'los registros de los empleados evaluados.

'Global Const Version = "1.25" ' Carmen Quintero.
'Global Const FechaModificacion = "18/02/2014"
'Global Const UltimaModificacion = ""   ' (CAS-22072 - Raffo - Adecuaciones GDD - People Review) Borrar Datos Seccion People Review

'Global Const Version = "1.26" ' Carmen Quintero.
'Global Const FechaModificacion = "09/04/2014"
'Global Const UltimaModificacion = ""   ' (CAS-24639 - CAPUTO - Encuesta de Clima) Borrar Datos Seccion Encuesta de Clima

'Global Const Version = "1.27" ' Carmen Quintero.
'Global Const FechaModificacion = "15/12/2014"
'Global Const UltimaModificacion = ""   ' (CAS-27453 - Galicia - Nuevo Formulario de EDD 2014)
''                                       Borrar Datos Seccion Revisión Anual de Responsabilidades y Objetivos
''                                       Borrar Datos Seccion Revisión Anual de Responsabilidades y Competencias
''                                       Borrar Datos Seccion Planeamiento Anual de Desarrollo
''                                       Borrar Datos Seccion Evaluacion Global de Desempeño
''                                       Borrar Datos Seccion Síntesis

'Global Const Version = "1.28" ' Carmen Quintero.
'Global Const FechaModificacion = "14/04/2015"
'Global Const UltimaModificacion = ""   ' (CAS-29781 - LOJACK - Custom en GDD)
'                                       Borrar Datos Seccion Aprobación de Evaluación
'                                       Borrar Datos Seccion Opciones de Evaluación

Global Const Version = "1.29" ' Carmen Quintero.
Global Const FechaModificacion = "18/09/2015"
Global Const UltimaModificacion = ""   ' (CAS-28226 - Raffo - Adecuación GDD Sección Coaching de Obj -Competencias por niveles- Obj corporativo)
'                                       Borrar Datos Seccion Evaluación de Objetivos Raffo


' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' DELOITTE
'                 - Inicializar Fecha Null para evadetevldor y incluir Comentarios para el log
'                 - Si no se borro la evacab en un proyecto No borrar el empleado del proyecto.
' Global evaproynro As Long ' no esta
' Global modificar As String ' no esta
' Global listainicial As String
    ' VERRRRRRRRLOOO si poone ro no!!!!!!!!!





' __________________________________________________________________________


Dim fs, f
'Global Flog

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Version_Valida As Boolean '10-05-2012 - Leti -para control de veriones, x ahora se define aca...

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global arrparam
Global evaevenro As Long
Global listainicial As String
Global Opcion As Long

'Global conceptos As String ' va?
'Global acumuladores As String ' va?
Global procesos As String
Global idUser As String
Global Inc As Double
Global IncTotal

Const estrevaluador = 52 ' Tipo de Estructura de los Evaluadores
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
Dim rsBorrar As New ADODB.Recordset
Dim rsSecc As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim ternro
Dim rsEmpl As New ADODB.Recordset
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim EmpBorrar As Long
Dim EmpCrear As Long


Dim secciones
Dim roles


On Error GoTo CE

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
    
    
     Nombre_Arch = PathFLog & "CrearFormularioEva" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    Flog.writeline
   
    Flog.writeline "Inicio Proceso de Creación de Evaluaciones: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
       
    
    
    'Abro la conexion
    'OpenConnection strconexion, objConn

    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'On Error Resume Next
    'OpenConnection strconexion, objconnProgreso  ' esta bien aca??
    'OpenConnection strconexion, objConn
    'If Err.Number <> 0 Or Error_Encrypt Then
     '   Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
     '   Exit Sub
    'End If
    
    
    TiempoInicialProceso = GetTickCount
    
    HuboErrores = False
    

    ' Control de versiones --------------------------
    Version_Valida = ValidarV(Version, 69, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Fin
    End If
       
    
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    cantRegistros = CInt(objRs!total)
    totalEmpleados = cantRegistros
    objRs.Close
   


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
       
        arrparam = Split(parametros, "@")
        evaevenro = arrparam(0)
        Opcion = arrparam(1)
        
        ' deloitte
        'modificar = arrparam(2)  !!!! ' define que se modifica-- borrar empls o generar form
        

        listainicial = "0"
          ' deloitte
        'StrSql = "SELECT DISTINCT evaproyemp.ternro "
        'StrSql = StrSql & " FROM  evaproyemp "
        'StrSql = StrSql & " WHERE  evaproyemp.evaproynro = " & evaproynro

        StrSql = "SELECT DISTINCT evacab.empleado "
        StrSql = StrSql & " FROM  evacab "
        StrSql = StrSql & " WHERE  evacab.evaevenro   = " & evaevenro
        OpenRecordset StrSql, Rsini
        Do Until Rsini.EOF
            listainicial = listainicial & "," & Rsini("empleado")
            Rsini.MoveNext
        Loop
        Rsini.Close
        Set Rsini = Nothing
    
       'listainicial = arrparam(1)
       Flog.writeline " Parametro que entro: Evento: " & evaevenro
       Flog.writeline " Lista inicial de empleados en el Evento: " & listainicial
     
     
       'EMPIEZA EL PROCESO

       'Obtengo los empleados sobre los que tengo que generar -
       Flog.writeline " Entra a cargar empleados de batch_empleado."
       CargarEmpleados NroProceso, rsEmpl, Opcion
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       
        
        'FGZ - 12/01/2010 ------------------------------------------------------------
        'Calculo de incremento del progreso
'        StrSql = "SELECT evacab.evacabnro, evacab.empleado, empleg, terape, ternom "
'        StrSql = StrSql & " FROM evacab "
'        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = evacab.empleado "
'        StrSql = StrSql & " WHERE evacab.empleado IN (" & listainicial & ")"
'        StrSql = StrSql & " AND   evacab.evaevenro =" & evaevenro
'        StrSql = StrSql & " AND   NOT EXISTS (SELECT * FROM batch_empleado WHERE "
'        StrSql = StrSql & " ternro = evacab.empleado"
'        StrSql = StrSql & " and bpronro=" & NroProceso & ")"
'        OpenRecordset StrSql, rsBorrar
'        EmpBorrar = rsBorrar.RecordCount
'        EmpCrear = rsEmpl.RecordCount
'        Inc = (100 / (EmpBorrar + EmpCrear))
        'FGZ - 12/01/2010 ------------------------------------------------------------
        
       'Modificado 05/03/2013
       'Calculo de incremento del progreso
       StrSql = "SELECT evacab.evacabnro, evacab.empleado, empleg, terape, ternom "
       StrSql = StrSql & " FROM evacab "
       StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = evacab.empleado "
       StrSql = StrSql & " WHERE evacab.empleado IN (" & listainicial & ")"
       StrSql = StrSql & " AND   evacab.evaevenro =" & evaevenro
       StrSql = StrSql & " AND   EXISTS (SELECT * FROM batch_empleado WHERE "
       StrSql = StrSql & " ternro = evacab.empleado"
       StrSql = StrSql & " AND bpronro=" & NroProceso
       StrSql = StrSql & "  AND  beparam = 2"
       StrSql = StrSql & ")"
       OpenRecordset StrSql, rsBorrar
       EmpBorrar = rsBorrar.RecordCount
       EmpCrear = rsEmpl.RecordCount
       Inc = (100 / (EmpBorrar + EmpCrear))
       
       IncTotal = 0
            
       'Call buscarSeccionesyRoles(evaevenro, secciones, roles)
       
       If Opcion = 1 Then
       
            Call buscarSeccionesyRoles(evaevenro, secciones, roles)
    
           'Genero por cada empleado un registro
            Flog.writeline "   "
            Flog.writeline "   "
            Flog.writeline " Para cada empleado se GENERA su FORMULARIO EVALUACION."
            
           
           Do Until rsEmpl.EOF
           
              EmpErrores = False
              ternro = rsEmpl!ternro
              
              'Genero los datos del empleado
              Flog.writeline "       "
              Flog.writeline "      Generar Evaluacion para el empleado (ternro): " & ternro
              
              Call generarFormulario(evaevenro, ternro, secciones, roles)
                    
                     
              'Actualizo el estado del proceso
              TiempoAcumulado = GetTickCount
              
              cantRegistros = cantRegistros - 1
              
              'FGZ - 12/01/2010 ------------------------
              'Se cambió el update
              'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
              '         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
              '         ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
              '
              'objConn.Execute StrSql, , adExecuteNoRecords
              
              'StrSql = "UPDATE batch_proceso SET bprcprogreso = bprcprogreso + " & Inc & _

                
              IncTotal = CDbl(IncTotal) + Inc
            
              StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(IncTotal) & _
                       ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                       ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
              'objconnProgreso.Execute StrSql, , adExecuteNoRecords
              'FGZ - 12/01/2010 ------------------------
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
       End If
       
       If Opcion = 2 Then
       
          EmpErrores = False
       
          Call borrarFormulario(evaevenro, listainicial)
               
          'Actualizo el estado del proceso
          TiempoAcumulado = GetTickCount
              
          cantRegistros = cantRegistros - 1
                    
          IncTotal = CDbl(IncTotal) + Inc
            
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(IncTotal) & _
                       ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                       ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
          objConn.Execute StrSql, , adExecuteNoRecords
              
             
          'Si se procesó la opcion indicada correctamente borro los registros de batch_empleado
          If Not EmpErrores Then
                StrSql = " DELETE FROM batch_empleado "
                StrSql = StrSql & " WHERE bpronro = " & NroProceso
                StrSql = StrSql & "  AND  beparam = " & Opcion
                objConn.Execute StrSql, , adExecuteNoRecords
          End If
       End If
       
       
    Else
        objRs.Close
        Set objRs = Nothing
        
        objConn.Close
        Set objConn = Nothing
        
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
    Set objConn = Nothing
    
   
    
Fin:
    Flog.Close
    If objConn.State = adStateOpen Then objConn.Close
     
   
    Exit Sub
     
CE:
    HuboErrores = True
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description & " - " & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & " SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    

End Sub ' MAIN


'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset, Opcion)

Dim StrEmpl As String
    
    
    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc & " AND beparam=" & Opcion
    

    'StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    
    OpenRecordset StrEmpl, rsEmpl
End Sub





' ____________________________________________________________________________________
' consulta la definición del Formulario - Secciones y Roles
' ____________________________________________________________________________________
Sub buscarSeccionesyRoles(evaevenro, secciones, roles)

Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Dim secciones2
Dim roles2
Dim rol
Dim rolasp
Dim rolhab
    
On Error GoTo ME_SeccyRoles

 'Flog.writeline "   "
 'Flog.writeline " Entro a buscar Secciones y Roles de la Definición del Formulario. "
    
    
    'Secciones del Formulario de Evaluacion
    secciones2 = "0"
    roles2 = "0"
    
    StrSql = " SELECT evaseccnro "
    StrSql = StrSql & " FROM evasecc "
    StrSql = StrSql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro= evasecc.evatipnro"
    StrSql = StrSql & " INNER JOIN evaevento  ON evaevento.evatipnro = evatipoeva.evatipnro "
    StrSql = StrSql & " INNER JOIN evatiposecc ON evatiposecc.tipsecnro= evasecc.tipsecnro "
    StrSql = StrSql & " WHERE evaevenro = " & evaevenro
    OpenRecordset StrSql, rs2

    Do While Not rs2.EOF
        
        'buscar evaluadores
        StrSql = "SELECT evaoblieva.evatevnro, afteranterior, evasecc.ultimasecc,evarolaspdet, evaobliorden "
        StrSql = StrSql & " FROM evaoblieva " ' evatevobli,
        StrSql = StrSql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro= evaoblieva.evatevnro "
        StrSql = StrSql & " INNER JOIN evasecc ON evasecc.evaseccnro= evaoblieva.evaseccnro "
        StrSql = StrSql & " LEFT  JOIN evarolasp ON evarolasp.evarolnro = evatipevalua.evarolnro "
        StrSql = StrSql & " WHERE evaoblieva.evaseccnro = " & rs2("evaseccnro")
        StrSql = StrSql & " ORDER BY evaobliorden "
        OpenRecordset StrSql, rs1
        
        rol = rs2("evaseccnro")
        
        '*********************************
                StrSql = "SELECT evatevobli, "
        '****************************************
        
        Do While Not rs1.EOF
            
            If Not IsNull(rs1("evarolaspdet")) Then 'ASP que busca el ternro del evaluador
                rolasp = Trim(rs1("evarolaspdet"))
            Else
                rolasp = ""
            End If
            
            ' lo usa deloitte  y deloitte chile - lo saco p deloitte
            ' verrrrrrr deloitte ch
            'If rs1("ultimasecc") = -1 Then ' si es la ultima seccion no habilito ningun evadetevldor.
            '   rolhab = 0
            'Else
                If rs1("afteranterior") = -1 Then
                    rolhab = 0
                Else
                    rolhab = -1
                End If
            'End If
            
            ' XXXXXXXXXXXXXX deloittte
            'If aprobada = -1 Then ' si la cabecera esta aprobada no habilito ningun avedetevldor.
            '  habilitado = 0
            
            rol = rol & "--" & rs1("evatevnro") & "@" & rolasp & "@" & rolhab
            
        rs1.MoveNext
        Loop
        
        rs1.Close
        ' Set rs1 = Nothing
        
        
        secciones2 = secciones2 & "," & rs2("evaseccnro")
        roles2 = roles2 & "," & rol
       
       
    rs2.MoveNext
    Loop
    rs2.Close
    

    secciones = Split(secciones2, ",")
    roles = Split(roles2, ",")


Set rs1 = Nothing
Set rs2 = Nothing

Exit Sub

ME_SeccyRoles:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


Sub borrarFormulario(evaevenro, listainicial)

Dim StrSql As String
Dim rsBorrar As New ADODB.Recordset
Dim rsBorrar2 As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim ternro

Dim listborrar

Dim Empleado As Long

Dim tipsecprogdel

On Error GoTo ME_Borrar

 Flog.writeline "   "
 Flog.writeline " Entro a BORRAR evaluaciones de los empleados. "
 
 
 '____________________________________________________________
 ' PARA TESTEO
 '____________________________________________________________
    Dim rsBatchEmp As New ADODB.Recordset
    Dim listaBatchEmpl
    listaBatchEmpl = "0"
    StrSql = "SELECT * FROM batch_empleado WHERE "
    StrSql = StrSql & " bpronro=" & NroProceso
    OpenRecordset StrSql, rsBatchEmp
    Do Until rsBatchEmp.EOF
         listaBatchEmpl = listaBatchEmpl & "," & rsBatchEmp("ternro")
         'armar cartel de aviso de los que se borraran
         rsBatchEmp.MoveNext
    Loop
    rsBatchEmp.Close
    Set rsBatchEmp = Nothing

    Flog.writeline " Lista de empleados en Batch-Empls - Testeo.  " & listaBatchEmpl
    Flog.writeline "    "
 
   
 '____________________________________________________________
 '____________________________________________________________
 
' Comentado el 05/03/2013
' StrSql = "SELECT evacab.evacabnro, evacab.empleado, empleg, terape, ternom "
' StrSql = StrSql & " FROM evacab "
' StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = evacab.empleado "
' StrSql = StrSql & " WHERE evacab.empleado IN (" & listainicial & ")"
' StrSql = StrSql & " AND   evacab.evaevenro =" & evaevenro
' StrSql = StrSql & " AND   NOT EXISTS (SELECT * FROM batch_empleado WHERE "
' StrSql = StrSql & " ternro = evacab.empleado"
' StrSql = StrSql & " and bpronro=" & NroProceso & ")"
 
 StrSql = "SELECT evacab.evacabnro, evacab.empleado, empleg, terape, ternom "
 StrSql = StrSql & " FROM evacab "
 StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = evacab.empleado "
 StrSql = StrSql & " WHERE evacab.empleado IN (" & listainicial & ")"
 StrSql = StrSql & " AND   evacab.evaevenro =" & evaevenro
 StrSql = StrSql & " AND   EXISTS (SELECT * FROM batch_empleado WHERE "
 StrSql = StrSql & " ternro = evacab.empleado"
 StrSql = StrSql & " AND bpronro=" & NroProceso
 StrSql = StrSql & " AND  beparam=2"
 StrSql = StrSql & ")"


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
StrSql = StrSql & " WHERE evacab.empleado IN (" & listborrar & ")"
StrSql = StrSql & " AND   evacab.evaevenro =" & evaevenro
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
        
        If Len(tipsecprogdel) <> 0 Then
            Select Case tipsecprogdel
            
            Case "borra_areacom_eva_00.asp":
                Call Borra_areacom(evaevenro, Empleado)
            
            Case "borra_resultados_eva_00.asp":
                Call Borra_resultados(evaevenro, Empleado)
                
            Case "borra_planaccion_eva_00.asp":
                Call Borra_planaccion(evaevenro, Empleado)
                
            Case "borra_objetivos_00.asp":
                Call Borra_objetivos(evaevenro, Empleado)
            
            Case "borra_objetivos_plan_00.asp":
                Call Borra_objetivos_plan(evaevenro, Empleado)
                Call Borra_objetivos(evaevenro, Empleado)
            
            Case "borra_gralobj_eva_00.asp":
                Call Borra_calificobj(evaevenro, Empleado)
                
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
            Case "borra_datosadm_eva_00.asp", "borra_datosadmRDE_eva_00.asp":
                Call Borra_datosadm(evaevenro, Empleado)
            Case "borra_calificobj_eva_00.asp":
                Call Borra_calificobj(evaevenro, Empleado)
                Call Borra_objetivos(evaevenro, Empleado) ' dado que la seccion tiene evaluacion de objs y calific gral de objs juntas
            Case "borra_objcom_eva_00.asp":
                Call Borra_objcom(evaevenro, Empleado)
            Case "borra_resultadosyarea_eva_00.asp":
                Call Borra_resultadosyarea(evaevenro, Empleado)
            
            Case "borra_areacomRDP_00.asp":
                Call Borra_areacom(evaevenro, Empleado)
            Case "borra_datosadmRDP_eva_00.asp":
                Call Borra_datosadm(evaevenro, Empleado)
            Case "borra_calificobjRDP_eva_00.asp":
                Call Borra_calificobj(evaevenro, Empleado)
            Case "borra_calificcomp_eva_00.asp":
                Call Borra_resultadosyarea(evaevenro, Empleado)
            Case "borra_calificgralRDP_eva_00.asp":
                Call Borra_calificgral(evaevenro, Empleado)
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
            ' deloitte Chile
            Case "borra_subcompxestr_eva_00.asp":
                Call Borra_subcompxestr(evaevenro, Empleado) ' evasubfresu
            
            Case "borra_compxestr_CH_eva_00.asp", "borra_compxestr_CH_eva_00.asp":
                Call Borra_areacom(evaevenro, Empleado)  ' borra los cometarios de area
                Call Borra_areaycompxestr(evaevenro, Empleado) ' evaarea  evaresultado evagruporesu
            
            Case "borra_vistos_CH_eva_00.asp":
                Call Borra_vistosyCalifGral(evaevenro, Empleado) ' evavistos evavistoresu
            Case "borra_calificgralarea_eva_00.asp":
                Call Borra_calificgralarea(evaevenro, Empleado) 'evagralarea
                
            Case "borra_areas_eva_00.asp":  ' evaarea y evaresultado
                Call Borra_areas(evaevenro, Empleado)
                
                
                
            Case "borra_evaluaciongral_eva_00.asp": ' evagralresu
                Call Borra_evaluaciongral(evaevenro, Empleado)
                Call Borra_notas(evaevenro, Empleado)
                
            Case "borra_objetivosII_eva_00.asp":
                Call Borra_objetivosII(evaevenro, Empleado)
                
            Case "borra_planaccionII_eva_00.asp":
                Call Borra_planaccionII(evaevenro, Empleado)
                
            Case "borra_planaccion_MV_eva_00.asp":
                Call Borra_planaccionMV(evaevenro, Empleado)
                
            Case "borra_compcom_eva_00.asp":
                Call Borra_compcom(evaevenro, Empleado)
                
            Case "borra_resuyarea_uys_eva_00.asp":  ' evaarea - evaresultado - evagralarea
                Call Borra_areas_uys(evaevenro, Empleado)
                
            Case "borra_valoracgral_uys_eva_00.asp": ' evagralresu - evasalariob
                Call Borra_valoracgral_uys(evaevenro, Empleado)
            
            Case "borra_calificgralnota_eva_00.asp":
                Call Borra_calificgral(evaevenro, Empleado)
                Call Borra_notas(evaevenro, Empleado)
                
            Case "borra_planaccionII_eva_00.asp":
                Call Borra_planaccionII(evaevenro, Empleado)
                
            Case "borra_puestobanda_eva_00.asp":
                Call Borra_notas(evaevenro, Empleado)
                Call Borra_puestobanda(evaevenro, Empleado)
                
            ' sección de resultados integrales para AGD
            Case "borra_resultadosintegrales_eva_00.asp":
                Call Borra_resultadosintegrales(evaevenro, Empleado)
            
            ' sección de people review para RAFFO
            Case "borra_peoplereview_eva_00.asp":
                  Call Borra_peoplereview(evaevenro, Empleado)
                  
            ' sección de encuesta de clima para Caputo
            Case "borra_encuesta_caputo_eva_00.asp"
                 Call Borra_encuesta_caputo(evaevenro, Empleado)
                 
            ' sección de Revisión Anual de Responsabilidades y Objetivos para Galicia
            Case "borra_revisionobj_GAL_eva_00.asp":
                Call Borra_objetivosRevisionAnual(evaevenro, Empleado)
        
             ' sección de Revisión Anual de Responsabilidades y Competencias para Galicia
            Case "borra_revisioncomp_GAL_eva_00.asp":
                Call Borra_competenciasRevisionAnual(evaevenro, Empleado)
                
            ' sección de Plan Anual de Desarrollo para Galicia
            Case "borra_plananual_GAL_eva_00.asp":
                Call Borra_PlanAnualDesarrollo(evaevenro, Empleado)
                
            ' sección de Evaluacion Global de Desempeño para Galicia
            Case "borra_evaluacionglobal_GAL_eva_00.asp":
                Call Borra_evaluacionGlobal(evaevenro, Empleado)
                
            ' sección Síntesis para Galicia
            Case "borra_sintesis_GAL_eva_00.asp":
                Call Borra_Sintesis(evaevenro, Empleado)
                
            ' sección Aprobación de Evaluación para Lojack
            Case "borra_aprobacion_LOJ_eva_00.asp":
                Call Borra_aprobacion(evaevenro, Empleado)
                 
            ' sección Opciones de Evaluación para Lojack
            Case "borra_opcion_LOJ_eva_00.asp":
                Call Borra_opcion(evaevenro, Empleado)
                
            ' sección Evaluación de Objetivos Raffo
            Case "borra_objetivos_RAF_eva_00.asp"
                Call Borra_porcentaje(evaevenro, Empleado)
                
            End Select
            
        End If
        rs1.MoveNext
    Loop
    
    rs1.Close
    Set rs1 = Nothing
    
    Flog.writeline "    Borra la Cabecera de evaluacion."
    Call Borrar_cabecera(evaevenro, Empleado)
    
    
    
    TiempoAcumulado = GetTickCount
   
    IncTotal = CDbl(IncTotal) + Inc
       
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(IncTotal) & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProceso
    ' objconnProgreso.Execute StrSql, , adExecuteNoRecords
    objConn.Execute StrSql, , adExecuteNoRecords
    
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
' Generar todas las tablas de evaluacion
'--------------------------------------------------------------------
Sub generarFormulario(ByVal evaevenro, ByVal ternro, ByVal secciones, ByVal roles)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsSecc As New ADODB.Recordset
Dim rs1   As New ADODB.Recordset

Dim rs2   As New ADODB.Recordset
Dim rs3   As New ADODB.Recordset
Dim rs4   As New ADODB.Recordset

Dim evacabnro  As Long
Dim evaseccnro As Long
Dim evatevnro As Integer
Dim evarolaspdet As String
'XXXXXXXXXXXX Deloitte
' Dim tipsecobj  As Integer - en elestandar se saco todo lo refente a tipisecobj

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
        
'Dim evldrnro As Long
               
On Error GoTo MError


Dim I
Dim j
Dim rol
Dim rol2


' buscar empreporta
empreporta = 0

StrSql = "SELECT empreporta FROM empleado "
StrSql = StrSql & " WHERE ternro=" & ternro
OpenRecordset StrSql, rs2
If Not rs2.EOF Then
    If rs2("empreporta") <> 0 And Not IsNull(rs2("empreporta")) Then
        empreporta = rs2("empreporta")
    End If
End If
rs2.Close
            


'------------------------------------------------------------------
'Busco si existe ya la cabecera
'------------------------------------------------------------------

StrSql = "SELECT evacab.evacabnro, evaetanro, cabaprobada "
StrSql = StrSql & " FROM evacab "
StrSql = StrSql & " WHERE evacab.empleado = " & ternro
StrSql = StrSql & " AND   evacab.evaevenro = " & evaevenro
OpenRecordset StrSql, rsConsult

If rsConsult.EOF Then
        
        'EVACAB
        
        'objConn.BeginTrans
        
        Flog.writeline "        Inserta Cabecera de Evaluacion."
        
        ' xxx Deloittte
        '    StrSql = "INSERT INTO evacab "
        'StrSql = StrSql & " (evaevenro, empleado, cabevaluada, cabaprobada,cabobservacion, evaproynro,evaetanro) "
        'StrSql = StrSql & " VALUES (" & evaevenro & ", " & ternro & ", 0, 0, null," & evaproynro & ",NULL)"

        StrSql = "INSERT INTO evacab(evaevenro , empleado, cabevaluada, cabaprobada)"
        StrSql = StrSql & " VALUES (" & evaevenro & ", " & ternro & ", 0, 0)"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        ' comentar y volver a buscar el evacabnro ______________________________________________
        ' evacabnro = getLastIdentity(objConn, "evacab") 'NO ANDA - es segun la versión del motor de BD
        ' da error en getLastIdentity
        ' ______________________________________________________________________________________



        StrSql = "SELECT evacab.evacabnro, evaetanro, cabaprobada "
        StrSql = StrSql & " FROM evacab "
        StrSql = StrSql & " WHERE evacab.empleado = " & ternro
        StrSql = StrSql & " AND   evacab.evaevenro = " & evaevenro
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            evacabnro = rs2("evacabnro")
        Else
            evacabnro = -1
        End If
        rs2.Close
        
        
        If evacabnro = -1 Then
            
            Flog.writeline "       NO se pudo obtener el identificador de cabecera de evaluación - se tiene que volver a ejecutar el proceso."
            
        Else
        
            Call setearEtapa("", evaevenro, evacabnro)

            
            For I = 1 To UBound(secciones)
            
                evaseccnro = CInt(secciones(I)) 'rsSecc("evaseccnro")
    
                Flog.writeline "        Crea los registros de Evaluacion para los evaluadores en la seccion: " & evaseccnro
                
                
                'EVADET
                StrSql = "INSERT INTO evadet(evacabnro,evaseccnro,detcargada) "
                StrSql = StrSql & " VALUES (" & evacabnro & ", " & evaseccnro & ", 0)"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                
                rol = Split(roles(I), "--") ' secc.rol-rolasp-rolhab.
                
                
                If CInt(rol(0)) = evaseccnro Then
                
                    For j = 1 To UBound(rol)
                        rol2 = Split(rol(j), "@")
                      
                        evaluador = ""
                        
                        evatevnro = rol2(0)
                        evarolaspdet = rol2(1)  ' evarolaspdet = ""
                        habilitado = rol2(2)
                        
                        'VERRR esta consulta esta de mas... si se crea la cabecera de evaluacion entones NO existe evadetevldor
                        ' VERRRRR bien
                        StrSql = "SELECT * "
                        StrSql = StrSql & " FROM evadetevldor "
                        StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                        StrSql = StrSql & " AND   evadetevldor.evatevnro = " & evatevnro
                        StrSql = StrSql & " AND   evadetevldor.evaseccnro = " & evaseccnro
                        OpenRecordset StrSql, rs2
                        
                        If rs2.EOF Then
                            horahab = ""  ' XXXXX
                            fechahab = "NULL"
                        
                            If habilitado = -1 Then
                                hora = Mid(Time, 1, 8)
                                hora = strto2(Left(hora, 2)) & Right(hora, 2)
                                fechahab = ConvFecha(Date)
                                horahab = hora
                            End If
                            
                            Call buscarEvaluador(ternro, evarolaspdet, evaluador, empreporta)
                            
                        
                            StrSql = "INSERT INTO evadetevldor (evacabnro , evaseccnro, evatevnro,"
                            If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                            StrSql = StrSql & " evaluador,"
                            End If
                            StrSql = StrSql & " evldorcargada,habilitado,fechahab,horahab) "
                            StrSql = StrSql & " VALUES (" & evacabnro & ", " & evaseccnro & ", " & evatevnro
                            If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                            StrSql = StrSql & "," & evaluador
                            End If
                            StrSql = StrSql & ", 0," & habilitado & "," & fechahab & ",'" & horahab & "')"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                        Else
                            evaluador = rs2("evaluador")  ' VERRRR  para que lo quiero aca?111
                        End If
                        
                        rs2.Close
                        Set rs2 = Nothing
                        
                    Next
                
                End If
            
            Next
            
        End If  ' de evacabnro<>-1
Else
    
    Flog.writeline "        El empleado tiene cabecera de Evaluacion."
    evacabnro = rsConsult("evacabnro")
    aprobada = rsConsult("cabaprobada")
     
    Call setearEtapa(rsConsult("evaetanro"), evaevenro, evacabnro)
   
    'creo evadet y evadetevldor para secciones nuevas....
    For I = 1 To UBound(secciones)
        
        evaseccnro = CInt(secciones(I)) 'rsSecc("evaseccnro")
     
        Flog.writeline "        Crea los registros de Evaluacion para los evaluadores en la seccion " & evaseccnro
        
        StrSql = " SELECT evaseccnro "
        StrSql = StrSql & " FROM evadet "
        StrSql = StrSql & " WHERE evacabnro=" & evacabnro
        StrSql = StrSql & " AND  evaseccnro=" & evaseccnro
        OpenRecordset StrSql, rs2
        If rs2.EOF Then
            StrSql = "INSERT INTO evadet(evacabnro,evaseccnro,detcargada) "
            StrSql = StrSql & " VALUES (" & evacabnro & "," & evaseccnro & ",0)"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rs2.Close
        'Set rs2 = Nothing
        
        ' Crear los evadetevldor ..................
        rol = Split(roles(I), "--") ' secc.rol-rolasp-rolhab.
        
        If CInt(rol(0)) = evaseccnro Then
        
            For j = 1 To UBound(rol)
                rol2 = Split(rol(j), "@")
                
                evaluador = ""
                
                evatevnro = rol2(0)
                evarolaspdet = rol2(1)  ' evarolaspdet = ""
                
                If aprobada = -1 Then   ' si la cabecera esta aprobada no habilito ningun evadetevldor.
                    habilitado = 0
                Else
                    habilitado = rol2(2)
                End If
        
        
                StrSql = "SELECT * "
                StrSql = StrSql & " FROM evadetevldor "
                StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                StrSql = StrSql & " AND   evadetevldor.evatevnro = " & evatevnro
                StrSql = StrSql & " AND   evadetevldor.evaseccnro = " & evaseccnro
                OpenRecordset StrSql, rs3
                
                If rs3.EOF Then
                    rs3.Close
                    'Set rs3 = Nothing
                
                    horahab = ""
                    fechahab = "NULL"
                    
                    If habilitado = -1 Then
                        hora = Mid(Time, 1, 8)
                        hora = strto2(Left(hora, 2)) & Right(hora, 2)
                        fechahab = ConvFecha(Date)
                        horahab = hora
                    End If
               
          
                    ' PRIMERO BUSCAR - que un mismo evatevnro ya tenga cargado Manualmente el Evaluador
                    StrSql = "SELECT evaluador "
                    StrSql = StrSql & " FROM evadetevldor "
                    StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                    StrSql = StrSql & "     AND evadetevldor.evatevnro = " & evatevnro
                    StrSql = StrSql & "     AND NOT evadetevldor.evaluador is NULL  "
                    OpenRecordset StrSql, rs3
                    If Not rs3.EOF Then
                        evaluador = rs3("evaluador")
                    End If
                    rs3.Close
                    'Set rs3 = Nothing


                    If Trim(evaluador) = "" Then
                        Call buscarEvaluador(ternro, evarolaspdet, evaluador, empreporta)
                    End If
                   
                    
                   StrSql = "INSERT INTO evadetevldor(evacabnro , evaseccnro, evatevnro, "
                   If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                   StrSql = StrSql & " evaluador,"
                   End If
                   StrSql = StrSql & " evldorcargada,habilitado,fechahab,horahab) "
                   StrSql = StrSql & " VALUES (" & evacabnro & ", " & evaseccnro & ", " & evatevnro
                   If Trim(evaluador) <> "" And Not IsNull(evaluador) Then
                   StrSql = StrSql & "," & evaluador
                   End If
                   StrSql = StrSql & ", 0," & habilitado & "," & fechahab & ",'" & horahab & "')"
                   objConn.Execute StrSql, , adExecuteNoRecords
                    
                Else
                   'Comentado el 18/06/2012
                   'evaluador = rs3("evaluador")
                   'Fin
                   
                   'Agregado por Carmen Quintero 18/06/2012
                   ' Inicio
                   ' PRIMERO BUSCAR - que un mismo evatevnro ya tenga cargado Manualmente el Evaluador
                    StrSql = "SELECT evaluador"
                    StrSql = StrSql & " FROM evadetevldor "
                    StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & evacabnro
                    StrSql = StrSql & " AND evadetevldor.evatevnro = " & evatevnro
                    StrSql = StrSql & " AND NOT evadetevldor.evaluador is NULL  "
                    OpenRecordset StrSql, rs3
                    If Not rs3.EOF Then
                        evaluador = rs3("evaluador")
                    End If
                    rs3.Close
                                       
                    If Trim(evaluador) = "" Then
                        Call buscarEvaluador(ternro, evarolaspdet, evaluador, empreporta)
                    End If
                   
                    If Len(evaluador) > 0 Then
                        StrSql = "UPDATE evadetevldor SET "
                        StrSql = StrSql & " evaluador=" & evaluador
                        StrSql = StrSql & " WHERE evacabnro=" & evacabnro
                        StrSql = StrSql & " AND evatevnro=" & evatevnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    'Fin
                    'rs3.Close
                   'Set rs3 = Nothing
                End If
                    
            Next
            
        End If
            
    Next ' de secciones
    
   
End If ' de select evacab

rsConsult.Close
Set rsConsult = Nothing
        
    
Exit Sub

MError:
    Flog.writeline "       Error en el tercero " & ternro & " Error: " & Err.Description
    Flog.writeline "       Ultimo SQL Ejecutado: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub



'____________________________________________________________________
' busca el evaluador según el asp configurado
'____________________________________________________________________
Sub buscarEvaluador(ternro, evarolaspdet, ByRef evaluador, empreporta)
Dim rs2   As New ADODB.Recordset
Dim rs3   As New ADODB.Recordset
Dim rs4   As New ADODB.Recordset

Dim consejero ' para Deloitte


On Error GoTo ME_busqEvldor


evaluador = ""

    If Len(Trim(evarolaspdet)) <> 0 Then
        Select Case evarolaspdet
              
        Case "buscar_auto_eva.asp":
              evaluador = ternro
                                          
        Case "buscar_revisor_eva.asp":
           If empreporta <> 0 And Not IsNull(empreporta) Then
              evaluador = empreporta
            End If
            
        Case "buscar_supervisor_eva.asp":
            ' se busca el reporta A del reporta A
            StrSql = "SELECT super.ternro "
            StrSql = StrSql & " FROM empleado "
            StrSql = StrSql & " INNER JOIN empleado revisor ON revisor.ternro = empleado.empreporta "
            StrSql = StrSql & " INNER JOIN empleado super   ON super.ternro = revisor.empreporta "
            StrSql = StrSql & " WHERE empleado.ternro=" & ternro
            OpenRecordset StrSql, rs2
            If Not rs2.EOF Then
                evaluador = rs2("ternro")
            End If
            rs2.Close
        
            
        Case "buscar_garante_eva.asp":
            'hay que buscar por tipoestructura Garante - CODELCO
            StrSql = "SELECT estrnro "
            StrSql = StrSql & " FROM his_estructura "
            StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro
            StrSql = StrSql & "   AND his_estructura.htethasta IS NULL "
            StrSql = StrSql & "   AND his_estructura.tenro =" & ctenroarea
            OpenRecordset StrSql, rs2
             
            If Not rs2.EOF Then
                 'Buscar un garante de la division
                 StrSql = "SELECT his_estructura.ternro "
                 StrSql = StrSql & " FROM his_estructura "
                 StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = his_estructura.ternro "
                 StrSql = StrSql & " INNER JOIN his_estructura area ON his_estructura.ternro = area.ternro "
                 StrSql = StrSql & "    AND area.tenro=" & ctenroarea
                 StrSql = StrSql & "    AND area.estrnro=" & rs2("estrnro")
                 StrSql = StrSql & " WHERE his_estructura.tenro=" & ctenrogarante
                 StrSql = StrSql & "    AND his_estructura.htethasta IS NULL "
                 OpenRecordset StrSql, rs3
                 If Not rs3.EOF Then
                     evaluador = rs3("ternro")
                 End If
                 rs3.Close
            End If
            rs2.Close
        
        Case "buscar_por_estructura_eva.asp":
             'hay que buscar por estructura para BCO INDUSTRIAL
            StrSql = "SELECT estrcodext "
            StrSql = StrSql & " FROM estructura "
            StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro = estructura.estrnro "
            StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro
            StrSql = StrSql & "   AND his_estructura.htethasta IS NULL "
            StrSql = StrSql & "   AND his_estructura.tenro =" & estrevaluador
            OpenRecordset StrSql, rs2
            
            If Not rs2.EOF Then
               If Not EsNulo(rs2("estrcodext")) Then
                    'Buscar el empleado asociado al registro de estructura en el codigo externo
                    If IsNumeric(rs2("estrcodext")) Then
                        StrSql = "SELECT empleado.ternro "
                        StrSql = StrSql & " FROM empleado "
                        StrSql = StrSql & " WHERE empleado.empleg=" & CLng(Trim(rs2("estrcodext")))
                        OpenRecordset StrSql, rs3
                        If Not rs3.EOF Then
                        evaluador = rs3("ternro")
                        End If
                        rs3.Close
                    End If
               End If
            End If
            rs2.Close

            
        Case "buscar_proyrevisor_eva.asp":  ' DELOITTE RDE
            StrSql = "SELECT proysocio, proygerente, proyrevisor "
            StrSql = StrSql & " FROM evaproyecto "
            StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evaproynro = evaproyecto.evaproynro "
            StrSql = StrSql & " WHERE evaevenro=" & evaevenro
            OpenRecordset StrSql, rs2
            If Not rs2.EOF Then
                If (ternro <> rs2("proyrevisor") And ternro <> rs2("proygerente")) Then
                    evaluador = rs2("proyrevisor")
                Else
                    If ternro = rs2("proyrevisor") Then
                        evaluador = rs2("proygerente")  'Validador
                    Else
                        evaluador = rs2("proysocio")  'Aux. Validador
                    End If
                End If
            End If
            rs2.Close

            
        Case "Validador":  ' DELOITTE RDE
            StrSql = "SELECT proysocio, proygerente, proyrevisor "
            StrSql = StrSql & " FROM evaproyecto "
            StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evaproynro = evaproyecto.evaproynro "
            StrSql = StrSql & " WHERE evaevenro=" & evaevenro
            OpenRecordset StrSql, rs2
            
            If Not rs2.EOF Then
                If (ternro <> rs2("proyrevisor") And ternro <> rs2("proygerente")) Then
                    evaluador = rs2("proygerente")
                Else
                    If ternro = rs2("proyrevisor") Then
                        evaluador = rs2("proysocio")  'Aux. Validador
                    Else
                        evaluador = rs2("proysocio")  'Aux. Validador
                    End If
                End If
            End If
            rs2.Close

        Case "Tipo Estructura": ' DELOITTE - Consejero ..
            ' se busca el ternro asociado al tipo estuctura Consejero (GdD-Consejero)
            ' por ahora busco x nombre ...
            ' ver - legajo entero?, estructura 60 caracteres
            consejero = 0
            
            StrSql = "SELECT tenro  "
            StrSql = StrSql & " FROM tipoestructura "
            StrSql = StrSql & " WHERE tedabr LIKE '%Consejero%' "
            OpenRecordset StrSql, rs2
            If Not rs2.EOF Then
                consejero = rs2("tenro")
            End If
            rs2.Close
            
            
                        
            StrSql = "SELECT his_estructura.ternro, estrdabr, empleado.ternro evaluador "
            StrSql = StrSql & " FROM his_estructura "
            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
            StrSql = StrSql & " INNER JOIN empleado ON empleado.empleg = estructura.estrdabr "
            StrSql = StrSql & " WHERE his_estructura.ternro=" & ternro
            StrSql = StrSql & "     AND his_estructura.tenro=" & consejero
            StrSql = StrSql & " ORDER BY htetdesde DESC " ' ver si fijarme que este nulo htethsta o fecha en el periodo
            OpenRecordset StrSql, rs2
            If Not rs2.EOF Then
               evaluador = rs2("evaluador")
            End If
            rs2.Close
            

            
        End Select
        
        
    End If ' de evarolaspdet <> ""

                  


Set rs2 = Nothing
Set rs3 = Nothing

Exit Sub

ME_busqEvldor:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub



'____________________________________________________________________
' Setea a Etapa la primer etapa al Evaluado, si no existe
'____________________________________________________________________
Sub setearEtapa(evaetanro, evaevenro, evacabnro)
Dim StrSql As String
Dim rs2 As New ADODB.Recordset

On Error GoTo ME_Etapa
    
    ' crea todos los registros de evaroleta (se usen o no)
    StrSql = "SELECT DISTINCT evaoblieva.evatevnro "
    StrSql = StrSql & " FROM evaoblieva "
    StrSql = StrSql & " INNER JOIN evatipevalua ON evatipevalua.evatevnro= evaoblieva.evatevnro "
    StrSql = StrSql & " INNER JOIN evasecc ON evasecc.evaseccnro= evaoblieva.evaseccnro "
    StrSql = StrSql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro = evasecc.evatipnro "
    StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evatipnro = evatipoeva.evatipnro "
    StrSql = StrSql & " WHERE evaevento.evaevenro=" & evaevenro
    StrSql = StrSql & "   AND evaoblieva.evatevnro "
    StrSql = StrSql & "           NOT IN ( "
    StrSql = StrSql & "                    SELECT evatevnro FROM evaroleta "
    StrSql = StrSql & "                    WHERE evaroleta.evacabnro=" & evacabnro
    StrSql = StrSql & "                   ) "
    OpenRecordset StrSql, rs2
    
    Do While Not rs2.EOF
        StrSql = "INSERT INTO evaroleta(evacabnro,evatevnro,evaetanro) "
        StrSql = StrSql & " VALUES (" & evacabnro & "," & rs2("evatevnro") & ", 0)"
        objConn.Execute StrSql, , adExecuteNoRecords
    rs2.MoveNext
    Loop
    rs2.Close
    
    
    
    If Trim(evaetanro) = "" Or IsNull(evaetanro) Then
       'buscar la ETAPA
       StrSql = "SELECT evaforeta.evaetanro "
       StrSql = StrSql & " FROM evaforeta "
       StrSql = StrSql & " INNER JOIN evaevento ON evaevento.evatipnro= evaforeta.evatipnro "
       StrSql = StrSql & " WHERE evaforeta.evadef = -1"
       StrSql = StrSql & "  AND  evaevento.evaevenro = " & evaevenro
       OpenRecordset StrSql, rs2
       If Not rs2.EOF Then
           evaetanro = rs2("evaetanro")
       Else
           evaetanro = ""
       End If
       rs2.Close
       
    
    
       If Len(Trim(evaetanro)) <> 0 And evaetanro <> 0 Then
           
           StrSql = "UPDATE evacab SET "
           StrSql = StrSql & " evaetanro= " & evaetanro
           StrSql = StrSql & " WHERE evacabnro = " & evacabnro
           objConn.Execute StrSql, , adExecuteNoRecords
           
           
           StrSql = "SELECT * FROM evaroleta "
           StrSql = StrSql & " WHERE evaroleta.evacabnro=" & evacabnro
           OpenRecordset StrSql, rs2
           Do While Not rs2.EOF
                StrSql = "UPDATE evaroleta SET "
                StrSql = StrSql & " evaetanro=" & evaetanro
                StrSql = StrSql & " WHERE evacabnro=" & evacabnro
                StrSql = StrSql & "   AND evatevnro=" & rs2("evatevnro")
                objConn.Execute StrSql, , adExecuteNoRecords
           rs2.MoveNext
           Loop
           rs2.Close
           
       End If
       
    Else
       
    End If
    
    
Set rs2 = Nothing

Exit Sub

ME_Etapa:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

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
    
    StrSql = "DELETE  FROM evaroleta WHERE evaroleta.evacabnro IN "
    StrSql = StrSql & " (SELECT evacabnro FROM evacab WHERE "
    StrSql = StrSql & " evacab.evaevenro  = " & evaevenro
    StrSql = StrSql & "    AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
        
    StrSql = "DELETE FROM evadetevldor WHERE evadetevldor.evacabnro IN "
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
    
     
    lista = "0"
    
    ' busca Objetivos definidos para el empleado  (nueva version de seccion Objetivos)
    ' PARA ITAU - EN EL ESTANDAR SE COMENTA POR AHORA - hay que correr script en bd y a parte modiificar todas las secciones de objtivos (cambiar en consulta evldrnro, por evaluaobj.evldrnro o evaobjetivo.evldrnro)
    'StrSql = " SELECT DISTINCT evaobjetivo.evaobjnro "
    'StrSql = StrSql & " FROM evadetevldor "
    'StrSql = StrSql & " INNER JOIN evaobjetivo ON evaobjetivo.evldrnro = evadetevldor.evldrnro "
    'StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    'StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    'StrSql = StrSql & "     AND evacab.empleado = " & Empleado
    'OpenRecordset StrSql, rs1
    'Do Until rs1.EOF
    '    lista = lista & "," & rs1("evaobjnro")
    '    rs1.MoveNext
    'Loop
    'rs1.Close
    'Set rs1 = Nothing
    
    
    ' busca Objetivos definidos para el empleado (version anterior de seccion Objetivos)
    StrSql = " SELECT DISTINCT evaluaobj.evaobjnro "
    StrSql = StrSql & " FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evaluaobj ON evaluaobj.evldrnro = evadetevldor.evldrnro "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado
    StrSql = StrSql & "     AND evaluaobj.evaobjnro NOT IN (" & lista & ") "
    OpenRecordset StrSql, rs1
    Do Until rs1.EOF
        lista = lista & "," & rs1("evaobjnro")
        rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
        
    'StrSql = " SELECT DISTINCT evaluaobj.evaobjnro FROM evadetevldor "
    'StrSql = StrSql & " INNER JOIN evaluaobj ON evaluaobj.evldrnro= evadetevldor.evldrnro "
    'StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    'StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    'StrSql = StrSql & " AND   evacab.empleado = " & Empleado
    'OpenRecordset StrSql, rs1
   
    'Do Until rs1.EOF
    '   lista = lista & "," & rs1("evaobjnro")
    '    rs1.MoveNext
    'Loop
    'rs1.Close
    'Set rs1 = Nothing
     
     
    'borrar todos los resultados de objetivos (tiene un ternro asociado)
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
    
    ' los objetivos son unicos para cada empleado!
    StrSql = "DELETE FROM evaplan WHERE evaplan.evaobjnro IN "
    StrSql = StrSql & " (" & lista & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Borrar todos los comentarios del objetivo que ya de los EVADETEVLDOR
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
 
    'Borrar objetivos sin
    StrSql = "DELETE FROM evaobjetivo  WHERE evaobjetivo.evaobjnro IN "
    StrSql = StrSql & " (" & lista & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_objetivosRevisionAnual(evaevenro, Empleado)
    
    Dim StrSql As String
    Dim rs1 As New ADODB.Recordset
    Dim lista
    On Error GoTo ME_Borrar
    
    lista = "0"
    
    ' busca Objetivos definidos para el empleado
    StrSql = " SELECT DISTINCT evarevisionobj.evaobjnro "
    StrSql = StrSql & " FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evarevisionobj ON evarevisionobj.evldrnro = evadetevldor.evldrnro "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro = " & evaevenro
    StrSql = StrSql & "     AND evacab.empleado = " & Empleado
    StrSql = StrSql & "     AND evarevisionobj.evaobjnro NOT IN (" & lista & ") "
    OpenRecordset StrSql, rs1
    Do Until rs1.EOF
        lista = lista & "," & rs1("evaobjnro")
        rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
    
    'borrar todos los resultados de objetivos (tiene un ternro asociado)
    StrSql = "DELETE FROM evarevisionobj  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evarevisionobj.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
 
    
    'Borrar objetivos
    StrSql = "DELETE FROM evaobjetivo  WHERE evaobjetivo.evaobjnro IN "
    StrSql = StrSql & " (" & lista & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub






Sub Borra_objetivos_plan(evaevenro, Empleado)

    Dim StrSql As String
    Dim rs1 As New ADODB.Recordset
    Dim lista
    On Error GoTo ME_Borrar
    
    StrSql = "SELECT DISTINCT evaluaobj.evaobjnro FROM evadetevldor "
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
     
    StrSql = "DELETE FROM evaobjplan "
    StrSql = StrSql & " WHERE evaobjplan.evaobjnro IN (" & lista & ")"
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



Sub Borra_calificgralarea(evavenro, Empleado)
    
    'para Deloitte por ahora
    Dim StrSql As String
    On Error GoTo ME_Borrar
    
    StrSql = "DELETE FROM evagralarea  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro = evagralarea.evldrnro "
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
    
    
    'borrar las evaluaciones generales de Objetivos
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




' _______________________________________________________________________________
' _______________________________________________________________________________
Sub Borra_planaccionII(evaevenro, Empleado)

Dim StrSql As String
Dim rs1 As New ADODB.Recordset
Dim aspectos
Dim planes
    
On Error GoTo ME_Borrarpl

    planes = 0
    aspectos = 0
    
    StrSql = "SELECT DISTINCT evaaspecto.evaaspnro, evaplan.evaplnro "
    StrSql = StrSql & " FROM evaaspecto "
    StrSql = StrSql & " INNER JOIN evatipoaspecto ON evatipoaspecto.evataspnro = evaaspecto.evataspnro "
    StrSql = StrSql & " INNER JOIN evaaspectoplan ON evaaspectoplan.evaaspnro = evaaspecto.evaaspnro "
    StrSql = StrSql & " INNER JOIN evaplan ON evaplan.evaplnro = evaaspectoplan.evaplnro "
    StrSql = StrSql & " INNER JOIN evatipoplan ON evatipoplan.evatplnro = evaplan.evatipoplan "
    StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaaspectoplan.evldrnro "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro=" & evaevenro
    StrSql = StrSql & " AND evacab.empleado=" & Empleado
    OpenRecordset StrSql, rs1
        
    Do Until rs1.EOF
       aspectos = aspectos & "," & rs1("evaaspnro")
       planes = planes & "," & rs1("evaplnro")
    rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing


    StrSql = "DELETE FROM evaaspectoplan  "
    StrSql = StrSql & " WHERE evaaspnro IN (" & aspectos & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

    StrSql = "DELETE FROM evaplan  "
    StrSql = StrSql & " WHERE evaplnro IN (" & planes & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

    StrSql = "DELETE FROM evaaspecto "
    StrSql = StrSql & " WHERE evaaspnro IN (" & aspectos & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_Borrarpl:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql


End Sub

Sub Borra_puestobanda(evaevenro, Empleado)

    Dim StrSql As String
    
    On Error GoTo ME_Borrar
        
    StrSql = "DELETE FROM evapuestobanda  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evapuestobanda.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_resultadosintegrales(evaevenro, Empleado)

    Dim StrSql As String
    
    On Error GoTo ME_Borrar
        
    StrSql = "DELETE FROM evaresuintegrales  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " WHERE evadetevldor.evldrnro=evaresuintegrales.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_peoplereview(evaevenro, Empleado)

    Dim StrSql As String
    
    On Error GoTo ME_Borrar
        
    StrSql = "DELETE FROM evagenterevision  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " WHERE evadetevldor.evldrnro=evagenterevision.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_encuesta_caputo(evaevenro, Empleado)

    Dim StrSql As String
    
    On Error GoTo ME_Borrar
        
    StrSql = "DELETE FROM evaresultadoencuesta  "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor  "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " WHERE evadetevldor.evldrnro=evaresultadoencuesta.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


' _______________________________________________________________________________
' _______________________________________________________________________________
Sub Borra_planaccionMV(evaevenro, Empleado)


Dim StrSql As String
Dim rs1 As New ADODB.Recordset
Dim aspectos
Dim planes
    
On Error GoTo ME_Borrarpl

    planes = 0
    aspectos = 0
    
    StrSql = "SELECT DISTINCT evaaspecto.evaaspnro, evaplan.evaplnro "
    StrSql = StrSql & " FROM evaaspecto "
    StrSql = StrSql & " INNER JOIN evatipoaspecto ON evatipoaspecto.evataspnro = evaaspecto.evataspnro "
    StrSql = StrSql & " INNER JOIN evaaspectoplan ON evaaspectoplan.evaaspnro = evaaspecto.evaaspnro "
    StrSql = StrSql & " INNER JOIN evaplan ON evaplan.evaplnro = evaaspectoplan.evaplnro "
    StrSql = StrSql & " INNER JOIN evatipoplan ON evatipoplan.evatplnro = evaplan.evatipoplan "
    StrSql = StrSql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaaspectoplan.evldrnro "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro "
    StrSql = StrSql & " WHERE evacab.evaevenro=" & evaevenro
    StrSql = StrSql & " AND evacab.empleado=" & Empleado
    OpenRecordset StrSql, rs1
        
    Do Until rs1.EOF
       aspectos = aspectos & "," & rs1("evaaspnro")
       planes = planes & "," & rs1("evaplnro")
    rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing


    StrSql = "DELETE FROM evaaspectoplan  "
    StrSql = StrSql & " WHERE evaaspnro IN (" & aspectos & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

    StrSql = "DELETE FROM evaplan  "
    StrSql = StrSql & " WHERE evaplnro IN (" & planes & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

    StrSql = "DELETE FROM evaaspecto "
    StrSql = StrSql & " WHERE evaaspnro IN (" & aspectos & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_Borrarpl:
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


Sub Borra_areas_uys(evavenro, Empleado)

Dim StrSql As String
On Error GoTo ME_Borrar
    
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
    
    'borra los datos de evaluacion general de area
    StrSql = "DELETE FROM evagralarea "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evagralarea.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


Sub Borra_valoracgral_uys(evavenro, Empleado)
Dim StrSql As String
On Error GoTo ME_Borrar
    
      
    'borra los datos de la evaluacion gral de resultado
    StrSql = "DELETE FROM evagralresu "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evagralresu.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
     'borra los datos de la tabla evasalariob
    StrSql = "DELETE FROM evasalariob "
    StrSql = StrSql & " WHERE EXISTS (SELECT * FROM evadetevldor "
    StrSql = StrSql & " INNER JOIN evacab ON evacab.evacabnro =evadetevldor.evacabnro"
    StrSql = StrSql & " where evadetevldor.evldrnro=evasalariob.evldrnro "
    StrSql = StrSql & " AND evacab.evaevenro = " & evaevenro
    StrSql = StrSql & " AND evacab.empleado = " & Empleado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


Sub Borra_competenciasRevisionAnual(evaevenro, Empleado)
    
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
     
     StrSql = "DELETE FROM evarevisioncomp WHERE evarevisioncomp.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
     
     objConn.Execute StrSql, , adExecuteNoRecords
          
    Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


Sub Borra_PlanAnualDesarrollo(evaevenro, Empleado)
    
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
     
     StrSql = "DELETE FROM evaplananual WHERE evaplananual.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
     
     objConn.Execute StrSql, , adExecuteNoRecords
          
    Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_evaluacionGlobal(evaevenro, Empleado)
    
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
     
     StrSql = "DELETE FROM evaglobaldesemp WHERE evaglobaldesemp.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
     
     objConn.Execute StrSql, , adExecuteNoRecords
          
    Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_Sintesis(evaevenro, Empleado)
    
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
     
     StrSql = "DELETE FROM evasintesisobj WHERE evasintesisobj.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
     
     objConn.Execute StrSql, , adExecuteNoRecords
     
     StrSql = "DELETE FROM evasintesiscomp WHERE evasintesiscomp.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
     
     objConn.Execute StrSql, , adExecuteNoRecords
          
    Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub


Sub Borra_aprobacion(evaevenro, Empleado)
    
    Dim StrSql As String
    Dim rs1 As New ADODB.Recordset
    Dim lista
    On Error GoTo ME_Borrar
    
    StrSql = " SELECT evadetevldor.evldrnro FROM evadetevldor "
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
     
     StrSql = "DELETE FROM evaaprobacion WHERE evaaprobacion.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
     objConn.Execute StrSql, , adExecuteNoRecords
          
    Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_opcion(evaevenro, Empleado)
    
    Dim StrSql As String
    Dim rs1 As New ADODB.Recordset
    Dim lista
    On Error GoTo ME_Borrar
    
    StrSql = " SELECT evadetevldor.evldrnro FROM evadetevldor "
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
     
     StrSql = "DELETE FROM evaopcion WHERE evaopcion.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
     objConn.Execute StrSql, , adExecuteNoRecords
          
    Exit Sub

ME_Borrar:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql
    
End Sub

Sub Borra_porcentaje(evaevenro, Empleado)
    
    Dim StrSql As String
    Dim rs1 As New ADODB.Recordset
    Dim lista
    On Error GoTo ME_Borrar
    
    StrSql = " SELECT evadetevldor.evldrnro FROM evadetevldor "
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
     
     StrSql = "DELETE FROM evaporcentaje WHERE evaporcentaje.evldrnro IN "
     StrSql = StrSql & " (" & lista & ")"
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









