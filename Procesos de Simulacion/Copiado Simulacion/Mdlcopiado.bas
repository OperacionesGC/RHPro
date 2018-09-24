Attribute VB_Name = "Mdlcopiado"
Option Explicit
' Copiador de tablas de simulacion para el proceso de Simulacion de Bajas

Global Descripcion As String
Global Cantidad As Single


'----------------------------------------------------------

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial
' Autor      : Diego Rosso
' Fecha      : 05/07/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProcesoBatch = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProcesoBatch = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    'FGZ - 17/10/2012 -----------------------------------------------------------
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Nombre_Arch = PathFLog & "CopiadoSimulacion" & "-" & NroProcesoBatch & ".log"
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Nombre_Arch = PathFLog & "CopiadoSimulacion" & "-" & NroProcesoBatch & ".log"
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo ME_Gral

    'estaba generando los logs fuera de la carpeta de l usuario
    'Nombre_Arch = PathFLog & "CopiadoSimulacion" & "-" & NroProcesoBatch & ".log"
    'Set fs = CreateObject("Scripting.FileSystemObject")
    'Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 221 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    
    'Si no existe creo la carpeta donde voy a guardar el log
    If Not fs.FolderExists(PathFLog & rs_batch_proceso!iduser) Then fs.CreateFolder (PathFLog & rs_batch_proceso!iduser)
    Nombre_Arch = PathFLog & rs_batch_proceso!iduser & "\" & "CopiadoSimulacion" & "-" & NroProcesoBatch & ".log"
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    'FGZ - 17/10/2012 -----------------------------------------------------------
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline

'    On Error Resume Next
'    'Abro la conexion
'    OpenConnection strconexion, objConn
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
'    OpenConnection strconexion, objconnProgreso
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
'    On Error GoTo 0
    
    'FGZ - 11/11/2011 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 221, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Fin
    End If
    'FGZ - 11/11/2011 --------- Control de versiones ------
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'FGZ - 17/10/2012 ----------------------------------------------------------------------------
    'lo cambié mas arriba porque necesitaba recuperarlo para ver el usuario y crear el log
    ''Obtengo los datos del proceso
    'StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 221 AND bpronro =" & NroProcesoBatch
    'OpenRecordset StrSql, rs_batch_proceso
    'FGZ - 17/10/2012 ----------------------------------------------------------------------------
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Copiado(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    Flog.writeline "---------------------------------------------------"
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General: " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    GoTo Fin
ME_Gral:

End Sub


Public Sub Copiado(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que realiza el llenado de las tablas de simulaciòn (sim_)
' Autor      : Diego Rosso
' Fecha      : 05/08/2008
' Modificacion:
'   Fecha = "15/08/2010"
'   Autor = Diego Rosso
'   Se adapto el proceso de copiado ya que ahora hay 3 tipos de simulación.
' Modificacion:
'   Fecha = "04/11/2010"
'   Autor = Diego Rosso
'   Se Modifica la Busqueda de novedades historicas estaba filtrando por periodo y debia filtar por Fecha inicio
'   y fecha fin del proceso que la genero ya las hisnovemp no tienen nepliqdesde ni nepliqhasta
'   estos campos se usan si son retroacticas las novedades originale)

' --------------------------------------------------------------------------------------------


'*********************************************************************************************************************
'Fecha = "15/08/2010"
'Importante: Valores del campo protipoSim que indica que tipo de simulación se esta procesando
'  1 Simulacion normal
'  2 simulacion de baja
'  3 simulacion retroactivos
'*********************************************************************************************************************

Dim rs_emple As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_Consulta As New ADODB.Recordset
Dim rs_proceso As New ADODB.Recordset
Dim rs_proceso_retro As New ADODB.Recordset
Dim rs_Francos As New ADODB.Recordset
Dim rs_Comisiones As New ADODB.Recordset


Dim Carpeta
Dim totalEmpleados
Dim cantRegistros
Dim linea As String
Dim Sep As String

Dim Pronro As Long
Dim CantidadEmpleados As Long
Dim AniosAtras As Integer
Dim Apellido As String
Dim nombre As String
Dim EncontroHistoricos As Boolean
Dim ArrParam
Dim ArrProcesos
Dim Ind
Dim Conceptos As String
Dim lstConceptosEncontrados As String
On Error GoTo CE


TiempoAcumulado = GetTickCount

'----------------------------------------------------------------------------
' Levanto cada parametro por separado
'----------------------------------------------------------------------------

Flog.writeline "Levantando parametros. " & Parametros
If Not IsNull(Parametros) Then

    'Flog.writeline "If Len(Parametros) >= 1 Then"
    
    If Len(Parametros) >= 1 Then
        'Por ahora solo un parametro-->numero de proceso de liquidacion (simulacion)
        'Pronro = CLng(Parametros)
         
         'creo un array con todos los numeros de procesos que me van a servir para las novedades retroactivas
         ArrParam = Split(Parametros, ",")
         
         ' Para todo lo demas me sirve cualquier numero de proceso asi que dejo el primero (o el unico si no retroactivos)
         Pronro = CLng(ArrParam(0))
         
         If UBound(ArrParam) > 0 Then
            ArrProcesos = Split(ArrParam(1), "@")
         End If
       
       ' Flog.writeline "Parametro Numero de Proceso = " & Pronro
    End If

Else
    Flog.writeline "Parametros nulos"
    HuboError = True
    Exit Sub
End If
Flog.writeline "Terminó de levantar los parametros"

'Por default un año
AniosAtras = 1

'ACA BUSCAR EN CONFREP CANTIDAD DE AÑOS PARA ATRAS PARA EL COPIADO.
'Por ahora solo la uso en emp_lic
Flog.writeline "Levantando Configuracion del reporte"
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 232 "
OpenRecordset StrSql, rs_aux
Do While Not rs_aux.EOF
    Select Case rs_aux!confnrocol
        Case 1
            AniosAtras = rs_aux!confval
            Flog.writeline "Columna 1. Cantidad de Años para atras. Valor ingresado: " & rs_aux!confval
        Case Else
            Flog.writeline "Columna no reconocida "
    End Select
    rs_aux.MoveNext
Loop
rs_aux.Close

'LEVANTO LOS DATOS DEL SIM_PROCESO
Flog.writeline "Levantando los datos del proceso"
StrSql = "SELECT s.pronro, s.prodesc, s.pliqnro, s.tprocnro, s.profecpago, "
StrSql = StrSql & " s.profecini, s.profecfin, s.empnro, s.proaprob, s.proestdesc,"
StrSql = StrSql & " s.protipoSim, s.profecbaja, s.caunro, s.pliqnroreal, "
StrSql = StrSql & "  s.tprocnroreal, s.pronroreal, s.proconceptos, s.procajuretro "
StrSql = StrSql & " FROM sim_proceso s "
StrSql = StrSql & " WHERE s.pronro = " & Pronro
OpenRecordset StrSql, rs_proceso
If Not rs_proceso.EOF Then
   Flog.writeline Espacios(Tabulador * 1) & "Se levantaron los datos del proceso. Tipo de Simulación:  " & rs_proceso!protipoSim
Else
   Flog.writeline Espacios(Tabulador * 1) & "No se encontro el proceso. No se puede continuar."
   Exit Sub
End If


'Busco los empleados
Flog.writeline "Levantos los empleados a copiar"
'StrSql = "SELECT * FROM Sim_DatosBaja WHERE pronro = " & Pronro

'****************************************************************************************************
'AHORA LOS VOY A LEVANTAR DE SIM_CABLIQ YA QUE NO SE USA MAS LA TABLA SIM_DATOSBAJA PORQUE
'AHORA NO TODOS LOS PROCESOS SON DE BAJA.
'****************************************************************************************************

StrSql = "SELECT c.cliqnro, c.empleado ternro FROM sim_cabliq c "
'FGZ - 20/09/2011 ----------- se agruegó el join para que levante solo empleados y no clones
StrSql = StrSql & " INNER JOIN ter_tip on c.empleado = ter_tip.ternro AND ter_tip.tipnro = 1"
StrSql = StrSql & " WHERE Procesado = 0 AND c.pronro = " & Pronro
OpenRecordset StrSql, rs_emple
If Not rs_emple.EOF Then
   CantidadEmpleados = rs_emple.RecordCount
   cantRegistros = CantidadEmpleados
   Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Empleados para Procesar: " & CantidadEmpleados
Else
   Flog.writeline Espacios(Tabulador * 1) & "No se encontraron Empleados para Procesar."
End If



Flog.writeline "Empiezo a Recorrer los empleados seleccionados "


Do While Not rs_emple.EOF
    Flog.writeline "********************************************************"
    Flog.writeline "Empleado (ternro) = " & rs_emple!Ternro
    
   
    '----------------------------------------------------------------
    ' Buscar el apellido y nombre
    '----------------------------------------------------------------
    
    'StrSql = "SELECT * FROM tercero WHERE ternro = " & rs_emple!ternro
    StrSql = "SELECT * FROM empleado WHERE ternro = " & rs_emple!Ternro
     OpenRecordset StrSql, rs_Tercero
     If Not rs_Tercero.EOF Then
     
        Flog.writeline "Empleado (legajo) = " & rs_Tercero!empleg
       
        '----------------------------------------------------------------
        ' 2- Buscar el Apellido
        '----------------------------------------------------------------
     
        Apellido = rs_Tercero!terape & " " & rs_Tercero!terape2
        Flog.writeline "Apellido = " & Apellido
        
        '----------------------------------------------------------------
        ' 3 - Buscar el Nombre
        '----------------------------------------------------------------
        
        nombre = rs_Tercero!ternom & " " & rs_Tercero!ternom2
        Flog.writeline "Nombres = " & nombre
        
        
        
        Flog.writeline "Comienza transaccion"
        
        MyBeginTrans
        
    'a)  Novedades (novemp)  a sim_novemp
        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_novemp para el empleado"
        
        StrSql = "DELETE  FROM sim_novemp  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla Novedades (novemp)  a sim_novemp"
        
        EncontroHistoricos = False
        
        If rs_proceso!protipoSim = 3 Then
        
            If UBound(ArrParam) = 0 Then
               ArrProcesos = Split(rs_proceso!pronroreal, ",")
            End If
            
            'si es de retroactivos tengo que ciclar por cada proceso
            For Ind = 0 To UBound(ArrProcesos)
                 
                 lstConceptosEncontrados = "0"
                 
                 Flog.writeline "Levantando los datos del proceso"
                 'traigo los datos del proceso
                 StrSql = "SELECT p.pronro, p.prodesc, p.propend, p.profeccorr, p.profecplan, p.pliqnro, "
                 StrSql = StrSql & "  p.tprocnro, p.prosist, p.profecpago, p.profecini, p.profecfin, p.empnro,"
                 StrSql = StrSql & "  p.proaprob, p.proestdesc "
                 StrSql = StrSql & " FROM proceso p "
                 StrSql = StrSql & " WHERE p.pronro = " & ArrProcesos(Ind)
                 OpenRecordset StrSql, rs_proceso_retro
                 
                 If Not rs_proceso_retro.EOF Then
                     Flog.writeline "Se buscaran novedades historicas para el proceso real num: " & ArrProcesos(Ind)
                     'Busco en novedades historicas del proceso
                     StrSql = "SELECT h.* FROM hisnovemp h"
                     StrSql = StrSql & " Where h.empleado =  " & rs_emple!Ternro
                     StrSql = StrSql & " AND pronro =" & rs_proceso_retro!Pronro
                     'Tengo que dejar afuera los conceptos seleecionados
                     If (rs_proceso!proconceptos <> "0") And Not EsNulo(rs_proceso!proconceptos) Then
                        StrSql = StrSql & " AND h.concnro NOT IN (" & rs_proceso!proconceptos & " )"
                     End If
                     OpenRecordset StrSql, rs_Consulta
                     If rs_Consulta.RecordCount > 0 Then  '*1
                       EncontroHistoricos = True
                       Flog.writeline "Se encontraron:" & rs_Consulta.RecordCount & " novedades historicas"
                       Do While Not rs_Consulta.EOF
                            StrSql = "INSERT INTO sim_novemp (concnro,tpanro,empleado,nevalor,nevigencia,nedesde,nehasta,"
                            StrSql = StrSql & " neretro,nepliqdesde,nepliqhasta,pronro, netexto) Values "
                            StrSql = StrSql & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!Tpanro & ", " & rs_Consulta!Empleado & ", "
                            If EsNulo(rs_Consulta!nevalor) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!nevalor & ", "
                            End If
                            
                            'nevigencia siempre en true
                            StrSql = StrSql & "-1, "
                           
                            
                            'La fechas desde y hasta las traigo del proceso real y no de his_nov_emp
                            If EsNulo(rs_proceso_retro!profecini) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & ConvFecha(rs_proceso_retro!profecini) & ", "
                            End If
                            If EsNulo(rs_proceso_retro!profecfin) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & ConvFecha(rs_proceso_retro!profecfin) & ", "
                            End If
                            
                            If EsNulo(rs_Consulta!neretro) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & ConvFecha(rs_Consulta!neretro) & ", "
                            End If
                            If EsNulo(rs_Consulta!nepliqdesde) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!nepliqdesde & ", "
                            End If
                            If EsNulo(rs_Consulta!nepliqhasta) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!nepliqhasta & ", "
                            End If
                            If EsNulo(rs_Consulta!Pronro) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!Pronro & ", "
                            End If
                            
                            'En His_novemp no tengo el campo texto
                            StrSql = StrSql & "NULL" & ")"
                           
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            'armo list de conceptos
                            lstConceptosEncontrados = lstConceptosEncontrados & "," & rs_Consulta!ConcNro
                            
                            rs_Consulta.MoveNext
                       Loop
                     Else '*1
                        Flog.writeline "No se encontraron novedades historicas para el proceso."
                     End If '*1
                     Flog.writeline "Se buscaran novedades reales que tenga vigencia superpuesta con el proceso real"
                     'buscar las Nov_emp reales que tenga vigencia superpuesta con el desde/hasta del proceso real
                     'y que el concepto no haya sido traido antes
                     
                     '+++++Consulta ejemplo++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                     '  SELECT * FROM novemp n  Where n.Empleado = 101
                     '  AND n.nevigencia = -1
                     '  AND ((n.nedesde <= '01/10/2010' AND (n.nehasta >= '01/10/2010' OR n.nehasta IS NULL))
                     '  OR (n.nedesde >= '01/10/2010' AND n.nedesde <= '30/10/2010'))
                     '  AND n.concnro NOT IN ( 0,4246)
                     '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                     
                     StrSql = "SELECT * FROM novemp  n"
                     StrSql = StrSql & " Where n.empleado =  " & rs_emple!Ternro
                     StrSql = StrSql & " AND n.nevigencia = -1"
                     StrSql = StrSql & " AND (( n.nedesde <=  " & ConvFecha(rs_proceso_retro!profecini)
                     StrSql = StrSql & " AND  (n.nehasta >=  " & ConvFecha(rs_proceso_retro!profecini)
                     StrSql = StrSql & " OR n.nehasta IS NULL)) "
                     StrSql = StrSql & " OR "
                     StrSql = StrSql & " (n.nedesde >= " & ConvFecha(rs_proceso_retro!profecini)
                     StrSql = StrSql & " AND n.nedesde <= " & ConvFecha(rs_proceso_retro!profecfin) & " ))"
                     'Tengo que dejar afuera los conceptos ya copiados
                     StrSql = StrSql & " AND n.concnro NOT IN (" & lstConceptosEncontrados & " )"
                     
                     OpenRecordset StrSql, rs_Consulta

                     If Not rs_Consulta.EOF Then '*2
                        Flog.writeline "Se encontraron:" & rs_Consulta.RecordCount & " novedades reales con fecha de vigencia superpuesta con las del proceso real"
                        Do While Not rs_Consulta.EOF
                            StrSql = "INSERT INTO sim_novemp (concnro,tpanro,empleado,nevalor,nevigencia,nedesde,nehasta,"
                            StrSql = StrSql & " neretro,nepliqdesde,nepliqhasta,pronro, netexto) Values "
                            StrSql = StrSql & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!Tpanro & ", " & rs_Consulta!Empleado & ", "
                                
                            If EsNulo(rs_Consulta!nevalor) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!nevalor & ", "
                            End If
                            'nevigencia siempre en true
                            StrSql = StrSql & "-1, "
                                    
                            'La fechas desde y hasta las traigo del proceso real
                            If EsNulo(rs_proceso_retro!profecini) Then
                                StrSql = StrSql & "NULL" & ", " 'No deberia entrar nunca aca
                            Else
                                StrSql = StrSql & ConvFecha(rs_proceso_retro!profecini) & ", "
                            End If
                            If EsNulo(rs_proceso_retro!profecfin) Then
                                StrSql = StrSql & "NULL" & ", " 'No deberia entrar nunca aca
                            Else
                                StrSql = StrSql & ConvFecha(rs_proceso_retro!profecfin) & ", "
                            End If
                             
                            If EsNulo(rs_Consulta!neretro) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & ConvFecha(rs_Consulta!neretro) & ", "
                            End If
                            If EsNulo(rs_Consulta!nepliqdesde) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!nepliqdesde & ", "
                            End If
                            If EsNulo(rs_Consulta!nepliqhasta) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!nepliqhasta & ", "
                            End If
                            If EsNulo(rs_Consulta!Pronro) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!Pronro & ", "
                            End If
                             
                            If EsNulo(rs_Consulta!netexto) Then
                                StrSql = StrSql & "NULL" & ")"
                            Else
                                StrSql = StrSql & rs_Consulta!netexto & ")"
                            End If
                            
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            'armo list de conceptos
                            lstConceptosEncontrados = lstConceptosEncontrados & "," & rs_Consulta!ConcNro
                             
                            rs_Consulta.MoveNext

                        Loop
                     End If '*2
                     
                     Flog.writeline "Se buscaran novedades reales sin vigencia superpuesta que sean del proceso real"
                     StrSql = "SELECT * FROM novemp  "
                     StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
                     StrSql = StrSql & " AND nevigencia = 0 "
                     'Tengo que dejar afuera los conceptos ya copiados
                     StrSql = StrSql & " AND concnro NOT IN (" & lstConceptosEncontrados & " )"
                     
                     OpenRecordset StrSql, rs_Consulta
                     
                     If Not rs_Consulta.EOF Then '*3
                        Flog.writeline "Se encontraron:" & rs_Consulta.RecordCount & " novedades reales sin vigencia"
                        Do While Not rs_Consulta.EOF
                            StrSql = "INSERT INTO sim_novemp (concnro,tpanro,empleado,nevalor,nevigencia,nedesde,nehasta,"
                            'StrSql = StrSql & " neretro,nepliqdesde,nepliqhasta,pronro, nenro,netexto) Values "
                            StrSql = StrSql & " neretro,nepliqdesde,nepliqhasta,pronro, netexto) Values "
                            StrSql = StrSql & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!Tpanro & ", " & rs_Consulta!Empleado & ", "
                            If EsNulo(rs_Consulta!nevalor) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!nevalor & ", "
                            End If
                          
                            'nevigencia siempre en true
                            StrSql = StrSql & "-1, "
                          
                            If EsNulo(rs_proceso_retro!profecini) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & ConvFecha(rs_proceso_retro!profecini) & ", "
                            End If
                            If EsNulo(rs_proceso_retro!profecfin) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & ConvFecha(rs_proceso_retro!profecfin) & ", "
                            End If
                            If EsNulo(rs_Consulta!neretro) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & ConvFecha(rs_Consulta!neretro) & ", "
                            End If
                            If EsNulo(rs_Consulta!nepliqdesde) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!nepliqdesde & ", "
                            End If
                            If EsNulo(rs_Consulta!nepliqhasta) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!nepliqhasta & ", "
                            End If
                            If EsNulo(rs_Consulta!Pronro) Then
                                StrSql = StrSql & "NULL" & ", "
                            Else
                                StrSql = StrSql & rs_Consulta!Pronro & ", "
                            End If
                          
                            If EsNulo(rs_Consulta!netexto) Then
                                StrSql = StrSql & "NULL" & ")"
                            Else
                                StrSql = StrSql & rs_Consulta!netexto & ")"
                            End If
                        
                            objConn.Execute StrSql, , adExecuteNoRecords
                          
                            rs_Consulta.MoveNext
                        Loop
                     
                     Else '*3
                        Flog.writeline "No se encontraron novedades sin vigencia"
                     End If '*3
                     
                 Else
                     Flog.writeline "No se encontro el proceso real"
                 End If 'If Not rs_proceso_retro.EOF Then
                 
                 'reinicializo
                 EncontroHistoricos = False
                 

            Next
            
            
        Else 'If rs_proceso!protipoSim = 3 Then
            
            'Si es simulacion comun o baja saco todas las novedadesde nov_emp
            StrSql = "SELECT * FROM novemp  "
            StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
            OpenRecordset StrSql, rs_Consulta
            
            Do While Not rs_Consulta.EOF
                StrSql = "INSERT INTO sim_novemp (concnro,tpanro,empleado,nevalor,nevigencia,nedesde,nehasta,"
                'StrSql = StrSql & " neretro,nepliqdesde,nepliqhasta,pronro, nenro,netexto) Values "
                StrSql = StrSql & " neretro,nepliqdesde,nepliqhasta,pronro, netexto) Values "
                StrSql = StrSql & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!Tpanro & ", " & rs_Consulta!Empleado & ", "
                If EsNulo(rs_Consulta!nevalor) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & rs_Consulta!nevalor & ", "
                End If
                StrSql = StrSql & rs_Consulta!nevigencia & ", "
                
                If EsNulo(rs_Consulta!nedesde) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!nedesde) & ", "
                End If
                If EsNulo(rs_Consulta!nehasta) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!nehasta) & ", "
                End If
                If EsNulo(rs_Consulta!neretro) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!neretro) & ", "
                End If
                If EsNulo(rs_Consulta!nepliqdesde) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & rs_Consulta!nepliqdesde & ", "
                End If
                If EsNulo(rs_Consulta!nepliqhasta) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & rs_Consulta!nepliqhasta & ", "
                End If
                If EsNulo(rs_Consulta!Pronro) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & rs_Consulta!Pronro & ", "
                End If
                
                'If EsNulo(rs_Consulta!nenro) Then
                '    StrSql = StrSql & "NULL" & ", "
                'Else
                '   StrSql = StrSql & rs_Consulta!nenro & ", "
                'End If
                
                If EsNulo(rs_Consulta!netexto) Then
                    StrSql = StrSql & "NULL" & ")"
                Else
                    StrSql = StrSql & rs_Consulta!netexto & ")"
                End If
              
                objConn.Execute StrSql, , adExecuteNoRecords
                
                rs_Consulta.MoveNext
            Loop
        End If
        
        
        

'b)  Novedades de Ajuste (novaju) a sim_novaju

        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_novaju para el empleado"
        StrSql = "DELETE FROM sim_novaju  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla Novedades de Ajuste (novaju) a sim_novaju"
        StrSql = "SELECT * FROM novaju  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        OpenRecordset StrSql, rs_Consulta
        If Not rs_Consulta.EOF Then
        
            Flog.writeline "Cantidad de registros a Copiar: " & rs_Consulta.RecordCount
            'FGZ - 14/10/2011 -------------------------------------
            If TipoBD = 4 Then
                StrSql = "ALTER TRIGGER TRG_sim_novaju DISABLE"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                StrSql = "SET  IDENTITY_INSERT sim_novaju ON"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            'FGZ - 14/10/2011 -------------------------------------
            Do While Not rs_Consulta.EOF
                StrSql = "INSERT INTO sim_novaju (concnro,empleado,navalor,navigencia,nadesde,nahasta,"
                StrSql = StrSql & " naretro,naajuste,pronro,nanro,natexto,napliqdesde,napliqhasta) Values "
                StrSql = StrSql & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!Empleado & ", "
                If EsNulo(rs_Consulta!navalor) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & rs_Consulta!navalor & ", "
                End If
                StrSql = StrSql & rs_Consulta!navigencia & ", "
                
                If EsNulo(rs_Consulta!nadesde) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!nadesde) & ", "
                End If
                If EsNulo(rs_Consulta!nahasta) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!nahasta) & ", "
                End If
                If EsNulo(rs_Consulta!naretro) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!naretro) & ", "
                End If
                
                StrSql = StrSql & rs_Consulta!naajuste & ", "
                
                If EsNulo(rs_Consulta!Pronro) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & rs_Consulta!Pronro & ", "
                End If
                
                StrSql = StrSql & rs_Consulta!nanro & ", "
                
                If EsNulo(rs_Consulta!natexto) Then
                    StrSql = StrSql & "NULL" & ","
                Else
                    StrSql = StrSql & rs_Consulta!natexto & ","
                End If
               
                If EsNulo(rs_Consulta!napliqdesde) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & rs_Consulta!napliqdesde & ", "
                End If
                If EsNulo(rs_Consulta!napliqhasta) Then
                    StrSql = StrSql & "NULL" & ")"
                Else
                    StrSql = StrSql & rs_Consulta!napliqhasta & ")"
                End If
                
                objConn.Execute StrSql, , adExecuteNoRecords
                
                rs_Consulta.MoveNext
            Loop
            'FGZ - 14/10/2011 -------------------------------------
            If TipoBD = 4 Then
                StrSql = "ALTER TRIGGER TRG_sim_novaju ENABLE"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                StrSql = "SET  IDENTITY_INSERT sim_novaju OFF"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            'FGZ - 14/10/2011 -------------------------------------
            Flog.writeline "Fin de copia de NOVAJU: "
            Flog.writeline ""
        Else
            Flog.writeline "No se encontraro datos para el empleado en NOVAJU "
            Flog.writeline ""
        End If
            

'c)  Estado (empleado.empest) copiar en sim_empleado.empest (ver si simpre va inactivo por default)

 
 
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de tabla sim_empleado para el empleado"
        StrSql = "DELETE FROM sim_empleado  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de empleado copiar en sim_empleado"
        StrSql = "SELECT * FROM empleado  "
        StrSql = StrSql & " Where ternro= " & rs_emple!Ternro
        OpenRecordset StrSql, rs_Consulta
        If Not rs_Consulta.EOF Then
        
            Flog.writeline "Cantidad de registros a Copiar: " & rs_Consulta.RecordCount
            
                     
            Do While Not rs_Consulta.EOF
                StrSql = "INSERT INTO sim_empleado (empleg,empfecbaja,empfbajaprev,empest,empfaltagr,ternro,empremu"
                StrSql = StrSql & ") Values "
                StrSql = StrSql & "(" & rs_Consulta!empleg & ", "
                
                If rs_proceso!protipoSim = 2 Then 'Si es simulacion de baja pongo la fecha sino lo dejo en nulo
                    StrSql = StrSql & ConvFecha(rs_proceso!profecbaja) & ", " 'Lo traigo de la tabla SIM_DATOSBAJA
                
                    'DIEGO ROSSO MODIFIQUE ESTA LINEA. ESTABA TOMANDO MAL EL RECORDSET***************
                      'StrSql = StrSql & ConvFecha(rs_Consulta!bajfec) & ", "
                       StrSql = StrSql & ConvFecha(rs_proceso!profecbaja) & ", "
                    'DIEGO ROSSO MODIFIQUE ESTA LINEA. ESTABA TOMANDO MAL EL RECORDSET***************
                Else
                        If EsNulo(rs_Consulta!empfecbaja) Then
                            StrSql = StrSql & "NULL" & ", "
                        Else
                            StrSql = StrSql & ConvFecha(rs_Consulta!empfecbaja) & ", "
                        End If
                        If EsNulo(rs_Consulta!empfbajaprev) Then
                            StrSql = StrSql & "NULL" & ", "
                        Else
                            StrSql = StrSql & ConvFecha(rs_Consulta!empfbajaprev) & ", "
                        End If
                End If
                
                If rs_proceso!protipoSim = 2 Then 'Si es simulacion de baja va inactivo siempre
                
                    StrSql = StrSql & "0" & ", " 'Siempre Inactivo
                Else
                    StrSql = StrSql & rs_Consulta!empest & ", "
                End If
                If EsNulo(rs_Consulta!empfaltagr) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!empfaltagr) & ", "
                End If
                
                StrSql = StrSql & rs_Consulta!Ternro & ", "
                                
                If EsNulo(rs_Consulta!empremu) Then
                    StrSql = StrSql & "NULL" & ")"
                Else
                    StrSql = StrSql & rs_Consulta!empremu & ")"
                End If
                
                objConn.Execute StrSql, , adExecuteNoRecords
                
                StrSql = " SELECT ternro from ter_tip where ternro = " & rs_emple!Ternro & " AND tipnro = 26  "
                OpenRecordset StrSql, rs_aux
                If rs_aux.EOF Then
                'Licho - Loinserto de prepo, despues analizaremos.
                    StrSql = "INSERT INTO ter_tip (tipnro, ternro) VALUES (26," & rs_emple!Ternro & ") "
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                rs_aux.Close
                rs_Consulta.MoveNext
                
            Loop
            
        
            Flog.writeline "Fin de copia de Sim_Empleado: "
            Flog.writeline ""
        Else
            Flog.writeline "ERROR No se econtro el empleado en la tabla empleado"
            Flog.writeline ""
        End If



'd)  Fases ( fases) a sim_fases
       ' Copio
        'FGZ - 30/09/2014 -------------------------
        StrSql = "SELECT * FROM sim_fases  "
        StrSql = StrSql & " Where empleado= " & rs_emple!Ternro
        OpenRecordset StrSql, rs_Consulta
        Do While Not rs_Consulta.EOF
            StrSql = "DELETE FROM sim_fases_preaviso  "
            StrSql = StrSql & " Where fasnro =  " & rs_Consulta!fasnro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            rs_Consulta.MoveNext
        Loop
        'FGZ - 30/09/2014 -------------------------
        
        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de tabla sim_fases para el empleado"
        StrSql = "DELETE FROM sim_fases  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de Fases del empleado en sim_fases"
        StrSql = "INSERT INTO sim_fases  "
        StrSql = StrSql & "SELECT * FROM fases Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        If rs_proceso!protipoSim = 2 Then 'Si es simulacion de baja
            ' Cierro la Fase activa
            '*******************************************
            StrSql = "SELECT * FROM fases  "
            StrSql = StrSql & " Where empleado= " & rs_emple!Ternro
            StrSql = StrSql & " AND bajfec is null"
            OpenRecordset StrSql, rs_Consulta
            If Not rs_Consulta.EOF Then
                StrSql = "UPDATE sim_fases  "
                StrSql = StrSql & " SET bajfec =" & ConvFecha(rs_proceso!profecbaja)
                StrSql = StrSql & " , caunro =" & rs_proceso!caunro
                StrSql = StrSql & " , estado =0 "
                StrSql = StrSql & " Where empleado= " & rs_emple!Ternro
                StrSql = StrSql & " and fasnro= " & rs_Consulta!fasnro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
        Flog.writeline "Fin de copia de sim_fases: "
        Flog.writeline ""
        
        'FGZ - 30/09/2014 -------------------------------------------
        Flog.writeline "    Avisos de Fases del empleado en sim_fases"
        StrSql = "INSERT INTO sim_fases_preaviso  "
        StrSql = StrSql & "SELECT * FROM fases_preaviso Where fasnro IN (SELECT fasnro FROM fases Where empleado= " & rs_emple!Ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        'FGZ - 30/09/2014 -------------------------------------------
        

'e)  Estructuras (his_estructuras) copiar a sim_his_estructura (todas)


        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_his_estructura para el empleado"
        StrSql = "DELETE FROM sim_his_estructura  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla his_estructura a sim_his_estructura"
        'copiado
        StrSql = " SELECT * FROM his_estructura Where ternro =  " & rs_emple!Ternro
        OpenRecordset StrSql, rs_Consulta
        If Not rs_Consulta.EOF Then
            Flog.writeline "Cantidad de registros a Copiar: " & rs_Consulta.RecordCount
            
            Do While Not rs_Consulta.EOF
                StrSql = "INSERT INTO sim_his_estructura (tenro,ternro,estrnro,htetdesde,htethasta,hismotivo,"
                StrSql = StrSql & " tipmotnro) Values "
                StrSql = StrSql & "(" & rs_Consulta!tenro & ", " & rs_emple!Ternro & ", " & rs_Consulta!Estrnro & ", "
                StrSql = StrSql & ConvFecha(rs_Consulta!htetdesde) & ", "
                
                
                If EsNulo(rs_Consulta!htethasta) Then
                    If rs_proceso!protipoSim = 2 Then 'Si es simulacion de baja cierro con la fecha de baja de la simulacion
                        'FGZ - 17/10/2012 -----------------------------------------------------------
                        'Si la estructura es de situacion de revista ==> la cierro un dia antes
                        'StrSql = StrSql & ConvFecha(rs_proceso!profecbaja) & ", "
                        If rs_Consulta!tenro = 30 Then
                            StrSql = StrSql & ConvFecha(rs_proceso!profecbaja - 1) & ", "
                        Else
                            StrSql = StrSql & ConvFecha(rs_proceso!profecbaja) & ", "
                        End If
                        'FGZ - 17/10/2012 -----------------------------------------------------------
                    Else
                        'es nula y no es simulaciòn de baja por lo que la dejo abierta
                        StrSql = StrSql & "NULL" & ", "
                    End If
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!htethasta) & ", "
                End If
                
                If EsNulo(rs_Consulta!hismotivo) Then
                    If rs_proceso!protipoSim = 2 Then 'Si es simulacion de baja cierro con la fecha de baja de la simulacion
                        StrSql = StrSql & "'', "
                    Else
                        StrSql = StrSql & "NULL" & ", "
                    End If
                Else
                   StrSql = StrSql & "'" & rs_Consulta!hismotivo & "', "
                End If
                
                If EsNulo(rs_Consulta!tipmotnro) Then
                    StrSql = StrSql & "0) "
                Else
                   StrSql = StrSql & rs_Consulta!tipmotnro & ")"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                
                rs_Consulta.MoveNext
            Loop
        
            Flog.writeline "Fin de copia del historico de estructuras "
            Flog.writeline ""
        Else
            Flog.writeline "ERROR No se econtro el empleado en la tabla His_estructura"
            Flog.writeline ""
        End If

        'e- Anexo)  Situacion de revista
        'Si es una simulacion de baja ==> se le debe asignar al empleado la situacion de revisata asociada a la casuda de baja que se le asoció al proceso de simulacion
        If rs_proceso!protipoSim = 2 Then 'Si es simulacion de baja cierro con la fecha de baja de la simulacion
            'Busco la estructura asociada a la causa de baja
            StrSql = " SELECT estrnro FROM causa_sitrev WHERE caunro = " & rs_proceso!caunro
            OpenRecordset StrSql, rs_Consulta
            If rs_Consulta.EOF Then
                Flog.writeline "La Causa de baja asociada al proceso no tiene situacion de revista asociada. Se deberá asignar la situacion de revista manualmente."
                Flog.writeline ""
            Else
                StrSql = "INSERT INTO sim_his_estructura (tenro,ternro,estrnro,htetdesde,htethasta,hismotivo,"
                StrSql = StrSql & " tipmotnro) Values "
                StrSql = StrSql & "(30, " & rs_emple!Ternro & ", " & rs_Consulta!Estrnro & ", "
                StrSql = StrSql & ConvFecha(rs_proceso!profecbaja) & ", "
                StrSql = StrSql & "NULL" & ", "
                StrSql = StrSql & "'', "
                StrSql = StrSql & "0) "
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If



'f)  Licencias(emp_lic) a sim_emp_lic

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_emp_lic para el empleado"
        StrSql = "DELETE FROM sim_emp_lic  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla Emp_Lic  a sim_emp_lic. Usa Años atras"
        StrSql = "INSERT INTO sim_emp_lic  "
        StrSql = StrSql & " SELECT e.* " 'Licho 25/11/2011
        'StrSql = StrSql & " SELECT e.elfechadesde, e.elfechahasta, e.empleado, e.tdnro, e.emp_licnro, "
        'StrSql = StrSql & " e.eldiacompleto, e.elhoradesde, e.elhorahasta, e.thnro, e.elcantdias, "
        'StrSql = StrSql & " e.elcantdiashab, e.elcantdiasfer, e.elcanthrs, e.eltipo, e.elorden, "
        'StrSql = StrSql & " e.elmaxhoras, e.licnrosig, e.elfechacert, e.pronro, e.licestnro, e.elobs"
        StrSql = StrSql & " FROM emp_lic e"
        StrSql = StrSql & " Where e.empleado =  " & rs_emple!Ternro
        StrSql = StrSql & " and e.elfechadesde > " & ConvFecha(DateAdd("YYYY", -AniosAtras, Date))
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Fin de copia de sim_emp_lic "
        Flog.writeline ""
        


'g)  Pago Dto Vacaciones (vacpagdesc) a sim_vacpagdesc


        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_vacpagdesc para el empleado"
        StrSql = "DELETE FROM sim_vacpagdesc  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla vacpagdesc  a sim_vacpagdesc"
        StrSql = " INSERT INTO sim_vacpagdesc  "
        StrSql = StrSql & "SELECT v.* "
        'StrSql = StrSql & "SELECT v.vacpdnro, v.pronro, v.cantdias, v.pliqnro, v.tprocnro, v.[manual], "
        'StrSql = StrSql & "v.pago_dto, v.ternro, v.concnro, v.emp_licnro, v.vacnro "
        StrSql = StrSql & "FROM vacpagdesc v Where v.ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Fin de copia de sim_vacpagdesc: "
        Flog.writeline ""

'h)  Prestamos (prestamo y pre_cuota) a simprestamo sim_pre_cuota

        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_prestamo para el empleado"
        StrSql = "DELETE FROM sim_prestamo  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla prestamo  a sim_prestamo"
        StrSql = "INSERT INTO sim_prestamo "
        StrSql = StrSql & "SELECT p.* "
        'StrSql = StrSql & "SELECT p.prenro, p.predesc, p.preimp, p.preanio, p.precantcuo, p.estnro, "
        'StrSql = StrSql & " p.ternro, p.preimpcuo, p.premes, p.pretpotor, p.prequin, p.pretna,  "
        'StrSql = StrSql & " p.monnro, p.quincenal, p.lnprenro, p.perfecaut, p.iduser, "
        'StrSql = StrSql & " p.prefecotor, p.sucursal, p.precompr, p.pliqnro, p.prediavto, "
        'StrSql = StrSql & " p.preiva , p.preotrosgas, p.nroimp, p.fechaimp "
        StrSql = StrSql & " FROM prestamo p "
        StrSql = StrSql & " Where p.ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
              
        Flog.writeline "Fin de copia de sim_prestamo: "
        
        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_pre_cuota para el empleado"
        StrSql = "DELETE FROM sim_pre_cuota  "
        StrSql = StrSql & " Where prenro in (SELECT prenro FROM prestamo Where ternro =   " & rs_emple!Ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla pre_cuota  a sim_pre_cuota"
        StrSql = "INSERT INTO sim_pre_cuota "
        StrSql = StrSql & "SELECT spc.* "
        'StrSql = StrSql & "SELECT spc.prenro, spc.cuoimp, spc.cuonro, spc.pronro, spc.cuocancela, "
        'StrSql = StrSql & "spc.cuoano, spc.cuomes, spc.cuoquin, spc.cuofecvto,  "
        'StrSql = StrSql & "spc.cuonrocuo, spc.cuogastos, spc.cuoiva, spc.cuototal,  "
        'StrSql = StrSql & "spc.cuocapital, spc.cuointeres, spc.cuosaldo  "
        
        'sebastian stremel - 10/12/2012 - se comenta la linea de abajo y se reemplaza por FROM pre_cuota spc - CAS-16691- AGD- BUGS DEL SIMULADOR EN PRESTAMOS
        'StrSql = StrSql & "FROM sim_pre_cuota spc  "
        StrSql = StrSql & "FROM pre_cuota spc  "
        StrSql = StrSql & "INNER JOIN prestamo p ON p.prenro = spc.prenro "
        StrSql = StrSql & "and p.ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
              
        Flog.writeline "Fin de copia de sim_pre_cuota "
        Flog.writeline ""

'i)  Embargos (embargo, embcuota) a sim_embargo sim_embcuota

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_embargo para el empleado"
        StrSql = "DELETE FROM sim_embargo  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla embargo  a sim_embargo"
        StrSql = "INSERT INTO sim_embargo "
        StrSql = StrSql & "SELECT e.* "
        'StrSql = StrSql & "SELECT e.tpenro, e.ternro, e.embest, e.embprioridad, e.embdesext, "
        'StrSql = StrSql & "e.embimp, e.embcantcuo, e.embquincenal, e.embnro, e.embanioini,  "
        'StrSql = StrSql & "e.embmesini, e.embquinini, e.embaniofin, e.embmesfin, "
        'StrSql = StrSql & "e.embquifin, e.bennom, e.embfecaut, e.fpagnro, e.embimpfij,  "
        'StrSql = StrSql & "e.embimppor, e.embexp, e.embcar, e.embjuz, e.embcgoem, e.embiva, "
        'StrSql = StrSql & "e.embnroof, e.embsec, e.bencuit, e.bencuenta, e.benbanco,  "
        'StrSql = StrSql & "e.embdeuda, e.embimpmin, e.embfecest, e.monnro, e.retley  "
        StrSql = StrSql & "FROM embargo e  "
        StrSql = StrSql & " Where e.ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
              
        Flog.writeline "Fin de copia de sim_embargo: "
        
        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_embcuota para el empleado"
        StrSql = "DELETE FROM sim_embcuota  "
        StrSql = StrSql & " Where embnro in (SELECT embnro FROM sim_embargo Where ternro =   " & rs_emple!Ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla embcuota  a sim_embcuota"
        StrSql = "INSERT INTO sim_embcuota "
        StrSql = StrSql & "SELECT e.*  "
        'StrSql = StrSql & "SELECT e.embnro, e.embcimp, e.embcnro, e.pronro, e.embccancela,  "
        'StrSql = StrSql & "e.embcanio, e.embcmes, e.embcquin, e.embcimpreal, e.embcretro,  "
        'StrSql = StrSql & "e.embcaran, e.cliqnro "
        StrSql = StrSql & "FROM embcuota e "
        StrSql = StrSql & "inner join sim_embargo se on se.embnro = e.embnro "
        StrSql = StrSql & "and se.ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
              
        Flog.writeline "Fin de copia de sim_embcuota "
        Flog.writeline ""

'j)  Vales (vales) a sim_vales


        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_vales para el empleado"
        StrSql = "DELETE FROM sim_vales  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla vales  a sim_vales"
        StrSql = "INSERT INTO sim_vales "
        StrSql = StrSql & "SELECT v.valnro, v.empleado, v.ppagnro, v.monnro, v.valmonto, "
        StrSql = StrSql & "v.valfecped, v.valfecprev, v.pliqnro, v.valdesc, v.pliqdto, "
        StrSql = StrSql & "v.pronro, v.tvalenro, v.valrevis, v.valautoriz, v.val_estnro, "
        StrSql = StrSql & " v.nroimp, v.fechaimp, v.valusuario, v.valaprosup  "
        StrSql = StrSql & " FROM vales v "
        StrSql = StrSql & " Where v.empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
              
        Flog.writeline "Fin de copia de sim_vales "
        Flog.writeline ""
        
'k)  Tickets (ticket, emp_tick, ….) sim_emp_tik

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_emp_ticket para el empleado"
        StrSql = "DELETE FROM sim_emp_ticket  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla emp_ticket a sim_emp_ticket"
        StrSql = "INSERT INTO sim_emp_ticket "
        StrSql = StrSql & "SELECT e.etiknro, e.empleado, e.tiknro, e.tikpednro, e.etikfecha,  "
        StrSql = StrSql & "e.etikmonto, e.etikcant, e.etikhora, e.etikuser,  "
        StrSql = StrSql & "e.etikmanual, e.pronro "
        StrSql = StrSql & "FROM emp_ticket e  "
        StrSql = StrSql & " Where e.empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
              
        Flog.writeline "Fin de copia de sim_emp_ticket: "
        
        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_emp_tikdist para el empleado"
        StrSql = "DELETE FROM sim_emp_tikdist  "
        StrSql = StrSql & " Where etiknro in (SELECT etiknro FROM emp_ticket Where empleado = " & rs_emple!Ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla emp_tikdist  a sim_emp_tikdist"
        StrSql = "INSERT INTO sim_emp_tikdist "
        StrSql = StrSql & "SELECT et.etiknro, et.tikvalnro, et.tiknro, et.etikdmonto,  "
        StrSql = StrSql & "et.etikdmontouni, et.etikdcant  "
        StrSql = StrSql & "FROM emp_tikdist et  "
        StrSql = StrSql & "INNER JOIN emp_ticket et2 ON et2.etiknro = et.etiknro "
        StrSql = StrSql & "and et2.empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
              
        Flog.writeline "Fin de copia de sim_emp_tikdist: "
        Flog.writeline ""

'l)  Novedades de Gti (gti_acunov) sim_gti_acunov (VER)


        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_gti_acunov para el empleado"
        StrSql = "DELETE FROM sim_gti_acunov  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla de Novedades de Gti a sim_gti_acunov"
        StrSql = "INSERT INTO sim_gti_acunov "
        StrSql = StrSql & "SELECT ga.acnovnro, ga.concnro, ga.acnovvalor, ga.tpanro, "
        StrSql = StrSql & "ga.acnovhornro, ga.acnovfecaprob, ga.ternro, ga.gpanro, ga.pronro "
        StrSql = StrSql & "FROM gti_acunov ga  "
        StrSql = StrSql & "Where ga.ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
              
        Flog.writeline "Fin de copia de sim_gti_acunov "
        Flog.writeline ""

'm)  DDJJ (desmen) a sim_desmen

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_desmen para el empleado"
        StrSql = "DELETE FROM sim_desmen  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        Flog.writeline "Copiado de tabla de DDJJ a sim_desmen"
        StrSql = "SELECT * FROM desmen  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        OpenRecordset StrSql, rs_Consulta
        If Not rs_Consulta.EOF Then
            Flog.writeline "Cantidad de registros a Copiar: " & rs_Consulta.RecordCount
            
            Do While Not rs_Consulta.EOF
                StrSql = "INSERT INTO sim_desmen (itenro,empleado,desmondec,desmenprorra,desano,desfecdes,"
                StrSql = StrSql & " desfechas,descuit,desrazsoc,pronro) Values "
                StrSql = StrSql & "(" & rs_Consulta!itenro & ", " & rs_emple!Ternro & ", "
                
                If EsNulo(rs_Consulta!desmondec) Then
                     StrSql = StrSql & "NULL" & ", "
                Else
                   StrSql = StrSql & rs_Consulta!desmondec & ","
                End If
                
                StrSql = StrSql & rs_Consulta!desmenprorra & ", "
                StrSql = StrSql & rs_Consulta!desano & ", "
                
                If EsNulo(rs_Consulta!desfecdes) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!desfecdes) & ", "
                End If
                If EsNulo(rs_Consulta!desfechas) Then
                    StrSql = StrSql & "NULL" & ", "
                Else
                    StrSql = StrSql & ConvFecha(rs_Consulta!desfechas) & ", "
                End If
                
                If EsNulo(rs_Consulta!descuit) Then
                   StrSql = StrSql & "NULL" & ", "
                Else
                   StrSql = StrSql & "'" & rs_Consulta!descuit & "', "
                End If
                If EsNulo(rs_Consulta!desrazsoc) Then
                   StrSql = StrSql & "NULL" & ", "
                Else
                   StrSql = StrSql & "'" & rs_Consulta!desrazsoc & "', "
                End If
                If EsNulo(rs_Consulta!Pronro) Then
                     StrSql = StrSql & "NULL" & ") "
                Else
                   StrSql = StrSql & rs_Consulta!Pronro & ")"
                End If
                
                objConn.Execute StrSql, , adExecuteNoRecords
                
                rs_Consulta.MoveNext
            Loop
        
            Flog.writeline "Fin de copia de Desmen "
            Flog.writeline ""
        Else
            Flog.writeline "ERROR No se econtro el empleado en la tabla Desmen"
            Flog.writeline ""
        End If

        
'n)  Sueldo Remu (empleado.empremu) a sim_empleado.empremu
'esta en una copia anterior


'Ñ) Copiado de tabla ficharet a sim_ficharet  'Agregado el 29/08/2010
        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_ficharet para el empleado"
        StrSql = "DELETE FROM sim_ficharet "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla de ficharet a sim_ficharet"
        StrSql = "INSERT INTO sim_ficharet "
        StrSql = StrSql & " SELECT f.fecha, f.importe, f.pronro, f.liqsistema, f.empleado FROM ficharet f "
        StrSql = StrSql & " Where f.empleado = " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
              
        Flog.writeline "Fin de copia de sim_ficharet "
        Flog.writeline ""

'0) Copiado de tabla desliq a sim_desliq  'Agregado el 29/08/2010
        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_desliq para el empleado"
        StrSql = "DELETE FROM sim_desliq "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla de desliq a sim_desliq"
        StrSql = "INSERT INTO sim_desliq "
        StrSql = StrSql & " SELECT d.itenro, d.empleado, d.dlfecha, d.pronro, d.dlmonto, d.dlprorratea FROM desliq d "
        StrSql = StrSql & " Where d.empleado  = " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
              
        Flog.writeline "Fin de copia de sim_desliq "
        Flog.writeline ""

        '21-12-2010 - Diego Rosso - Se marca el empleado como procesado.
        StrSql = "Update sim_cabliq "
        StrSql = StrSql & "SET Procesado = -1 "
        StrSql = StrSql & " Where Pronro = " & Pronro
        StrSql = StrSql & " AND empleado = " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords

        
'FGZ - 09/11/2011 ----------------------------------------
'p)  Francos Compensatorios(emp_fr_comp) a sim_emp_fr_comp

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_emp_fr_comp para el empleado"
        StrSql = "DELETE FROM sim_emp_fr_comp  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 11/11/2011 -------------------------------------
        'If TipoBD = 4 Then
        '    StrSql = "ALTER TRIGGER sim_emp_fr_comp DISABLE"
        '    objConn.Execute StrSql, , adExecuteNoRecords
        'Else
        '    StrSql = "SET  IDENTITY_INSERT sim_emp_fr_comp ON"
        '    objConn.Execute StrSql, , adExecuteNoRecords
        'End If
        'FGZ - 11/11/2011 -------------------------------------
        Flog.writeline "Copiado de tabla emp_fr_comp  a sim_emp_fr_comp. Usa Años atras"
        
        'FGZ - 03/07/2012 ----------------------------------------------------------------------------------
        'Le agregué control para ver si existen registros porque si el select es vacio ==> falla el insert
        StrSql = " SELECT e.ternro, e.fecha, e.unidad, e.Cantidad, e.comentario, e.liq, e.pronro "
        StrSql = StrSql & " FROM emp_fr_comp e"
        StrSql = StrSql & " Where e.ternro =  " & rs_emple!Ternro
        StrSql = StrSql & " and e.fecha > " & ConvFecha(DateAdd("YYYY", -AniosAtras, Date))
        OpenRecordset StrSql, rs_Francos
        If Not rs_Francos.EOF Then
            StrSql = "INSERT INTO sim_emp_fr_comp  "
            'FAF - 05/12/2011 -------------------------------------
            'Se comenta el select de todos los campos.
            'StrSql = StrSql & " SELECT e.* "
            StrSql = StrSql & " SELECT e.frannro, e.ternro, e.fecha, e.unidad, e.Cantidad, e.comentario, e.liq, e.pronro "
            StrSql = StrSql & " FROM emp_fr_comp e"
            StrSql = StrSql & " Where e.ternro =  " & rs_emple!Ternro
            StrSql = StrSql & " and e.fecha > " & ConvFecha(DateAdd("YYYY", -AniosAtras, Date))
            'FAF - 05/12/2011 -------------------------------------
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        Flog.writeline "Fin de copia de sim_emp_fr_comp "
        Flog.writeline ""
        'FGZ - 09/11/2011 ----------------------------------------
        'FGZ - 11/11/2011 -------------------------------------
        'If TipoBD = 4 Then
        '    StrSql = "ALTER TRIGGER sim_emp_fr_comp ENABLE"
        '    objConn.Execute StrSql, , adExecuteNoRecords
        'Else
        '    StrSql = "SET  IDENTITY_INSERT sim_emp_fr_comp OFF"
        '    objConn.Execute StrSql, , adExecuteNoRecords
        'End If
        'FGZ - 11/11/2011 -------------------------------------
        'FGZ - 03/07/2012 ----------------------------------------------------------------------------------
        
        
        
'FGZ - 17/10/2012 ----------------------------------------
'q)  Comisiones(liq_comision) a sim_liq_comision

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_liq_comision para el empleado"
        StrSql = "DELETE FROM sim_liq_comision  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla liq_comision  a sim_liq_comision. Usa Años atras"
        
        StrSql = " SELECT lc.ternro, lc.fecha, lc.concnro, lc.tpanro, lc.thnro, lc.mpt, lc.tht, lc.th, lc.pronro "
        StrSql = StrSql & " FROM liq_comision lc"
        StrSql = StrSql & " Where lc.ternro =  " & rs_emple!Ternro
        StrSql = StrSql & " and lc.fecha > " & ConvFecha(DateAdd("YYYY", -AniosAtras, Date))
        OpenRecordset StrSql, rs_Comisiones
        If Not rs_Francos.EOF Then
            StrSql = "INSERT INTO sim_liq_comision  "
            StrSql = StrSql & " SELECT lc.ternro, lc.fecha, lc.concnro, lc.tpanro, lc.thnro, lc.mpt, lc.tht, lc.th, lc.pronro "
            StrSql = StrSql & " FROM liq_comision lc"
            StrSql = StrSql & " Where lc.ternro =  " & rs_emple!Ternro
            StrSql = StrSql & " and lc.fecha > " & ConvFecha(DateAdd("YYYY", -AniosAtras, Date))
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        Flog.writeline "Fin de copia de sim_liq_comision "
        Flog.writeline ""
        
        
'FGZ - 26/02/2013 ----------------------------------------
'r)  Gastos(gastos) a sim_gastos

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_gastos para el empleado"
        StrSql = "DELETE FROM sim_gastos  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla gastos a sim_gastos. Usa Años atras"
        StrSql = "INSERT INTO sim_gastos "
        StrSql = StrSql & " SELECT g.gasnro, g.gasdesabr,g.proyecnro,g.monnro,g.gasvalor,g.ternro,g.provnro,g.gasfechaida,g.gashoraida,g.gasfechavuelta,g.gashoravuelta,g.gasrevisadopor,g.gaspagacliente,g.gaspagado,g.tipgasnro,g.pronro,g.gasretro,g.pliqdesde,g.pliqhasta FROM gastos g "
        StrSql = StrSql & " Where g.ternro  = " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Fin de copia de sim_gastos "
        Flog.writeline ""
        


'EAM - 04/03/2013 ----------------------------------------
's)  Venta de Vacaciones (Sykes) a sim_vacvendidos

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_vacvendidos para el empleado"
        StrSql = "DELETE FROM sim_vacvendidos  WHERE ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla vacvendidos a sim_vacvendidos."
        StrSql = "INSERT INTO sim_vacvendidos " & _
                " SELECT v.vacvendidosnro,ternro,empleg,aprobado,fechapago,pronro,iduser,fechacarga,cantvacvendidos,vacnro,venc,automatico FROM vacvendidos v " & _
                " WHERE v.ternro  = " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Fin de copia de sim_vacvendidos "
        Flog.writeline ""
        
        
'FGZ - 26/02/2013 ----------------------------------------
't)     Acumuladores mensuales(acu_mes) a sim_acu_mes

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_acu_mes para el empleado"
        StrSql = "DELETE FROM sim_acu_mes  "
        StrSql = StrSql & " Where ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla acu_mes a sim_acu_mes. Usa Años atras"
        StrSql = "INSERT INTO sim_acu_mes "
        StrSql = StrSql & " SELECT ternro, acunro, amanio, ammonto, amcant, ammes, ammontoreal FROM acu_mes "
        StrSql = StrSql & " Where ternro  = " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Fin de copia de sim_acu_mes "
        Flog.writeline ""
                
                
        Flog.writeline "Commit trasaccion"
        MyCommitTrans

        
     Else
        Flog.writeline "No se encontró el tercero"
     End If
    
       'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
              
        cantRegistros = cantRegistros - 1
           
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((CantidadEmpleados - cantRegistros) * 100) / CantidadEmpleados) & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                 ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & bpronro
        objConn.Execute StrSql, , adExecuteNoRecords
    

    rs_emple.MoveNext
Loop

'Cierro y libero
If rs_Francos.State = adStateOpen Then rs_Francos.Close
Set rs_Francos = Nothing
If rs_Comisiones.State = adStateOpen Then rs_Comisiones.Close
Set rs_Comisiones = Nothing



Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error (sub Copiado): " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyRollbackTrans
    MyBeginTrans
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((CantidadEmpleados - cantRegistros) * 100) / CantidadEmpleados) & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub

