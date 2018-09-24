Attribute VB_Name = "MdlClonar"
Option Explicit

'Clona un empleado como un nuevo sim_empleado (NN) a partir de un empleado existente o a partir de un clon (NN)

'Const Version = "1.00"
'Const FechaVersion = "16/08/2011"
'Autor = FGZ

'Const Version = "1.01"
'Const FechaVersion = "20/09/2011"
''Ult 20/09/2011 - Gonzalez Nicolás - Se valida que las fecha desde y hasta no esten vacías

'Const Version = "1.02"
'Const FechaVersion = "17/11/2011"   'FGZ
'                   Se cambió la numeracion de legajos (ahora no se permiten los mismos nro de legajos entre clones y empleados)

'Const Version = "1.03"
'Const FechaVersion = "20/12/2011"   'Sebastian Stremel
'                   Se copian novedades, y las novedades de ajuste del empleado

'Const Version = "1.04"
'Const FechaVersion = "03/07/2014"   'Ruiz Miriam
'                   CAS-21864 - H&A - Bugs detectados en la primera vuelta de testeo de R4
'                   Se recompiló para nivelar los comentarios

'Const Version = "1.05"
'Const FechaVersion = "30/06/2015"   'Fernandez, Matias
                                    'CAS-30766 - MONRESA - Error en estructuras de simulación
                                    'si se selecciona historial, trae todas las estructuras, sino, trae las estructuras actuales a la fecha
                                    
'Const Version = "1.06"
'Const FechaVersion = "03/07/2015"   'Fernandez, Matias
                                    'CAS-30766 - MONRESA - Error en estructuras de simulación
                                    'Distinct en la consulta que trae las estructuras de his_estructuras.
                                    

                                    
Const Version = "1.07"
Const FechaVersion = "29/07/2015"   'Fernandez, Matias
                                    'CAS-30766 - MONRESA - Error en estructuras de simulación
                                    'Mira las estructuras a futuro cuando no se selecciona historico




'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

Global Descripcion As String
Global Cantidad As Single


'----------------------------------------------------------

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial
' Autor      : Fernando Zwenger
' Fecha      : 10/08/2011
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
'--------------------------------------------------- MDF INICIO
    'strCmdLine = Command()
    'ArrParametros = Split(strCmdLine, " ", -1)
    'If UBound(ArrParametros) > 0 Then
    '    If IsNumeric(ArrParametros(0)) Then
    '        NroProcesoBatch = ArrParametros(0)
    '        Etiqueta = ArrParametros(1)
    '    Else
    '        Exit Sub
    '    End If
    'Else
    '    If IsNumeric(strCmdLine) Then
    '        NroProcesoBatch = strCmdLine
    '    Else
    '        Exit Sub
    '    End If
    ' End If
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
    
'---------------------------------------------------MDFF FIN
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error Resume Next
    
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "ClonacionNN" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 306 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        'Call Copiado(NroProcesoBatch, bprcparam)
         Call Clonado(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso " & NroProcesoBatch
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
End Sub


Public Sub Copiado(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que realiza el llenado de las tablas de simulaciòn (sim_)
' Autor      : FGZ
' Fecha      : 10/08/2011
' Modificacion:

' --------------------------------------------------------------------------------------------

'*********************************************************************************************************************
'Fecha = "15/08/2010"
'Importante: Valores del campo protipoSim que indica que tipo de simulación se esta procesando
'  1 Simulacion normal
'  2 simulacion de baja
'  3 simulacion retroactivos
'  4 Gestion Presupuestaria con NNs
'*********************************************************************************************************************
Dim rs_emple As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_Consulta As New ADODB.Recordset
Dim rs_proceso As New ADODB.Recordset
Dim rs_proceso_retro As New ADODB.Recordset

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

StrSql = "SELECT c.cliqnro, c.empleado ternro FROM sim_cabliq c WHERE Procesado = 0 AND c.pronro = " & Pronro
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
                            StrSql = StrSql & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!tpanro & ", " & rs_Consulta!Empleado & ", "
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
                            StrSql = StrSql & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!tpanro & ", " & rs_Consulta!Empleado & ", "
                                
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
                            StrSql = StrSql & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!tpanro & ", " & rs_Consulta!Empleado & ", "
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
                StrSql = StrSql & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!tpanro & ", " & rs_Consulta!Empleado & ", "
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
            
            StrSql = "SET  IDENTITY_INSERT sim_novaju ON"
            objConn.Execute StrSql, , adExecuteNoRecords
            
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
            StrSql = "SET  IDENTITY_INSERT sim_novaju OFF"
            objConn.Execute StrSql, , adExecuteNoRecords
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
                        StrSql = StrSql & ConvFecha(rs_proceso!profecbaja) & ", "
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

'f)  Licencias(emp_lic) a sim_emp_lic

        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_emp_lic para el empleado"
        StrSql = "DELETE FROM sim_emp_lic  "
        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla Emp_Lic  a sim_emp_lic. Usa Años atras"
        StrSql = "INSERT INTO sim_emp_lic  "
        StrSql = StrSql & " SELECT e.elfechadesde, e.elfechahasta, e.empleado, e.tdnro, e.emp_licnro, "
        StrSql = StrSql & " e.eldiacompleto, e.elhoradesde, e.elhorahasta, e.thnro, e.elcantdias, "
        StrSql = StrSql & " e.elcantdiashab, e.elcantdiasfer, e.elcanthrs, e.eltipo, e.elorden, "
        StrSql = StrSql & " e.elmaxhoras, e.licnrosig, e.elfechacert, e.pronro, e.licestnro, e.elobs"
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
        StrSql = StrSql & "SELECT v.vacpdnro, v.pronro, v.cantdias, v.pliqnro, v.tprocnro, v.[manual], "
        StrSql = StrSql & "v.pago_dto, v.ternro, v.concnro, v.emp_licnro, v.vacnro "
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
        StrSql = StrSql & "SELECT p.prenro, p.predesc, p.preimp, p.preanio, p.precantcuo, p.estnro, "
        StrSql = StrSql & " p.ternro, p.preimpcuo, p.premes, p.pretpotor, p.prequin, p.pretna,  "
        StrSql = StrSql & " p.monnro, p.quincenal, p.lnprenro, p.perfecaut, p.iduser, "
        StrSql = StrSql & " p.prefecotor, p.sucursal, p.precompr, p.pliqnro, p.prediavto, "
        StrSql = StrSql & " p.preiva , p.preotrosgas, p.nroimp, p.fechaimp "
        StrSql = StrSql & " FROM prestamo p "
        StrSql = StrSql & " Where p.ternro =  " & rs_emple!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
              
        Flog.writeline "Fin de copia de sim_prestamo: "
        
        
        'Limpio la tabla para el empleado
        Flog.writeline "Borrado de sim_pre_cuota para el empleado"
        StrSql = "DELETE FROM sim_pre_cuota  "
        StrSql = StrSql & " Where prenro in (SELECT prenro FROM prestamo Where ternro =   " & rs_emple!Ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Copiado de tabla prestamo  a sim_prestamo"
        StrSql = "INSERT INTO sim_pre_cuota "
        StrSql = StrSql & "SELECT spc.prenro, spc.cuoimp, spc.cuonro, spc.pronro, spc.cuocancela, "
        StrSql = StrSql & "spc.cuoano, spc.cuomes, spc.cuoquin, spc.cuofecvto,  "
        StrSql = StrSql & "spc.cuonrocuo, spc.cuogastos, spc.cuoiva, spc.cuototal,  "
        StrSql = StrSql & "spc.cuocapital, spc.cuointeres, spc.cuosaldo  "
        StrSql = StrSql & "FROM sim_pre_cuota spc  "
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
        StrSql = StrSql & "SELECT e.tpenro, e.ternro, e.embest, e.embprioridad, e.embdesext, "
        StrSql = StrSql & "e.embimp, e.embcantcuo, e.embquincenal, e.embnro, e.embanioini,  "
        StrSql = StrSql & "e.embmesini, e.embquinini, e.embaniofin, e.embmesfin, "
        StrSql = StrSql & "e.embquifin, e.bennom, e.embfecaut, e.fpagnro, e.embimpfij,  "
        StrSql = StrSql & "e.embimppor, e.embexp, e.embcar, e.embjuz, e.embcgoem, e.embiva, "
        StrSql = StrSql & "e.embnroof, e.embsec, e.bencuit, e.bencuenta, e.benbanco,  "
        StrSql = StrSql & "e.embdeuda, e.embimpmin, e.embfecest, e.monnro, e.retley  "
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
        StrSql = StrSql & "SELECT e.embnro, e.embcimp, e.embcnro, e.pronro, e.embccancela,  "
        StrSql = StrSql & "e.embcanio, e.embcmes, e.embcquin, e.embcimpreal, e.embcretro,  "
        StrSql = StrSql & "e.embcaran, e.cliqnro "
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


Public Sub Clonado(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que realiza el llenado de las tablas de simulaciòn (sim_)
' Autor      : FGZ
' Fecha      : 10/08/2011
' Modificacion: 20/09/2011 - Gonzalez Nicolás - Se valida que las fecha desde y hasta no esten vacías

' --------------------------------------------------------------------------------------------
'*********************************************************************************************************************
'Fecha = "15/08/2010"
'Importante: Valores del campo protipoSim que indica que tipo de simulación se esta procesando
'  1 Simulacion normal
'  2 simulacion de baja
'  3 simulacion retroactivos
'  4 Gestion Presupuestaria con NNs
'*********************************************************************************************************************
Dim rs_emple As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_Consulta As New ADODB.Recordset
Dim rs_proceso As New ADODB.Recordset
Dim rs_proceso_retro As New ADODB.Recordset

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


'Definicion de variables
Dim Cantidad_Parametros As Long
Dim I As Long

Dim Modelo_Origen As Long
Dim Tercero_Modelo As Long
Dim Modelo_Nombre As String
Dim Modelo_Apellido As String
Dim Modelo_Legajo As String
Dim Cantidad_Clones As Long
Dim Fecha_Ingreso As Date
Dim Fecha_Egreso As Date
Dim Crea_Fases As Boolean
Dim Crea_His_Estructura As Boolean

Dim Clon_NroLeg As String
Dim Clon_Nombre  As String
Dim Clon_Apellido As String
Dim Tercero_Creado As Long
Dim Tipo_tercero_NN As Long
Dim fasnro As Long

Dim Aux_Fecha As Date

Dim Crea_Nov As Boolean
Dim StrSql1 As String
Dim nevalor
Dim nedesde
Dim nehasta
Dim neretro
Dim nepliqdesde
Dim nepliqhasta
'Dim Pronro
Dim netexto
Dim tipmotnro
Dim motivo
Dim pronro1 As String
Dim rs_ConsultaNov As New ADODB.Recordset
On Error GoTo CE


TiempoAcumulado = GetTickCount

'----------------------------------------------------------------------------
' Levanto los parametro del proceso
'----------------------------------------------------------------------------
'Modelo Origen
'Tercero Modelo
'Modelo de Nombre
'Modelo de Apellido
'Modelo de Nro de legajo
'Cantidad de Clones
'Fecha de Ingreso
'Fecha de Egreso
'Crea Fases
'Crea His_estructura


Flog.writeline Espacios(Tabulador * 0) & "Levantando parametros. " & Parametros
Cantidad_Parametros = 7
Tipo_tercero_NN = 26
'Defaults
    Modelo_Origen = 1  'Corresponde al tipo de tercero (clon o empleado) utilizado como modelo. Defualt = 1 (Empleado)
    Tercero_Modelo = 0  'Corresponde al nro de tercero (clon o empleado) utilizado como modelo
    Modelo_Nombre = "NN"
    Modelo_Apellido = "NN"
    Modelo_Legajo = ""
    Cantidad_Clones = 0
    Fecha_Ingreso = CDate("01/" & Month(Now()) & "/" & Year(Now()))
    Fecha_Egreso = CDate("31/12" & "/" & Year(Now()))
    Crea_Fases = True
    Crea_His_Estructura = True
    Crea_Nov = False

If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
         'creo un array con todos los numeros de procesos que me van a servir para las novedades retroactivas
         ArrParam = Split(Parametros, "@")
         
         If UBound(ArrParam) < Cantidad_Parametros Then
            Flog.writeline "Cantidad de parametros incorrectos"
            HuboError = True
            Exit Sub
         End If
         
        'Defaults
        Modelo_Origen = ArrParam(0)
        Tercero_Modelo = ArrParam(1)
        Modelo_Nombre = ArrParam(2)
        Modelo_Apellido = ArrParam(3)
        Modelo_Legajo = ArrParam(4)
        Cantidad_Clones = CLng(ArrParam(5))
        
        If IsDate(ArrParam(6)) = True Then
            Fecha_Ingreso = CDate(ArrParam(6))
        End If
        
        If IsDate(ArrParam(7)) = True Then
            Fecha_Egreso = CDate(ArrParam(7))
        End If
        
        '---------------MDF 30/06/2015
        'Crea_Fases = True MDF
        Crea_Fases = ArrParam(8)
        'Crea_His_Estructura = True
        Crea_His_Estructura = ArrParam(9)
        '---------------MDF 30/06/2015
        If UBound(ArrParam) > 9 Then 'levanta novedades
            If (ArrParam(10)) = -1 Then
                Crea_Nov = True
            End If
        End If
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "Parametros nulos"
    HuboError = True
    Exit Sub
End If

Flog.writeline Espacios(Tabulador * 0) & "Origen Modelo " & Modelo_Origen
Flog.writeline Espacios(Tabulador * 0) & "Tercero Modelo " & Tercero_Modelo
Flog.writeline Espacios(Tabulador * 0) & "Modelo de Nombre " & Modelo_Nombre
Flog.writeline Espacios(Tabulador * 0) & "Modelo de Apellido " & Modelo_Apellido
Flog.writeline Espacios(Tabulador * 0) & "Modelo de Nro de Legajo " & Modelo_Legajo
Flog.writeline Espacios(Tabulador * 0) & "Cantidad de clones " & Cantidad_Clones
Flog.writeline Espacios(Tabulador * 0) & "Fecha de Ingreso " & Fecha_Ingreso
Flog.writeline Espacios(Tabulador * 0) & "Fecha de Egreso " & Fecha_Egreso
Flog.writeline Espacios(Tabulador * 0) & "Crea Fases " & Crea_Fases
Flog.writeline Espacios(Tabulador * 0) & "Crea Estructuras " & Crea_His_Estructura

'Validaciones de parametros
If Cantidad_Clones < 0 Then
    Flog.writeline Espacios(Tabulador * 0) & "La Cantidad de clones no puede ser negativa"
End If


Flog.writeline "Terminó de levantar los parametros"

'--------------------------------------------------------------------
'Levanto configuraciones generales para todos los clones
''Por default un año
'AniosAtras = 1
'
''ACA BUSCAR EN CONFREP CANTIDAD DE AÑOS PARA ATRAS PARA EL COPIADO.
''Por ahora solo la uso en emp_lic
'Flog.writeline "Levantando Configuracion del reporte"
'StrSql = " SELECT * FROM confrep "
'StrSql = StrSql & " WHERE repnro = 232 "
'OpenRecordset StrSql, rs_aux
'Do While Not rs_aux.EOF
'    Select Case rs_aux!confnrocol
'        Case 1
'            AniosAtras = rs_aux!confval
'            Flog.writeline "Columna 1. Cantidad de Años para atras. Valor ingresado: " & rs_aux!confval
'        Case Else
'            Flog.writeline "Columna no reconocida "
'    End Select
'    rs_aux.MoveNext
'Loop
'rs_aux.Close

IncPorc = 100 / Cantidad_Clones

'Busco el empleado o clon modelo
StrSql = "SELECT * FROM tercero "
Select Case Modelo_Origen
Case 1:
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = tercero.ternro "
Case Else
    StrSql = StrSql & " INNER JOIN sim_empleado ON sim_empleado.ternro = tercero.ternro "
End Select
StrSql = StrSql & " WHERE tercero.ternro = " & Tercero_Modelo
OpenRecordset StrSql, rs_emple
If Not rs_emple.EOF Then
    For I = 1 To Cantidad_Clones
        Flog.writeline Espacios(Tabulador * 1) & "Creando clon " & I & " ..."
    
            MyBeginTrans
                'establezco el nro de legajo
                Clon_NroLeg = ""
                Call CalcularLegajo(Clon_NroLeg)
                'Clon_NroLeg = Modelo_Legajo & "_" & I
                
                Clon_Nombre = Modelo_Nombre & "_" & Clon_NroLeg
                Clon_Apellido = Modelo_Apellido & "_" & Clon_NroLeg
            
                Flog.writeline Espacios(Tabulador * 2) & "Legajo " & Clon_NroLeg
                Flog.writeline Espacios(Tabulador * 2) & "Nombre  " & Clon_Nombre
                Flog.writeline Espacios(Tabulador * 2) & "Apellido " & Clon_Apellido
                
        
                'a) creo el tercero
                StrSql = " INSERT INTO tercero(ternom,terape,terfecnac,tersex,nacionalnro,paisnro,estcivnro)"
                StrSql = StrSql & " VALUES('" & Clon_Nombre & "','" & Clon_Apellido & "'," & ConvFecha(rs_emple!terfecnac) & "," & rs_emple!tersex & ","
                If Not EsNulo(rs_emple!nacionalnro) Then
                  StrSql = StrSql & rs_emple!nacionalnro & ","
                Else
                  StrSql = StrSql & "Null,"
                End If
                If Not EsNulo(rs_emple!paisnro) Then
                  StrSql = StrSql & rs_emple!paisnro & ","
                Else
                  StrSql = StrSql & "Null,"
                End If
                StrSql = StrSql & rs_emple!estcivnro & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            
                Tercero_Creado = getLastIdentity(objConn, "tercero")
                
                'Inserto el Registro correspondiente en ter_tip
                StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & Tercero_Creado & "," & Tipo_tercero_NN & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
        
                'b)  'Creo el sim_empleado
                
                'StrSql = "INSERT INTO sim_empleado (empleg,empfecbaja,empfbajaprev,empest,empfaltagr,ternro,empremu"
                StrSql = "INSERT INTO sim_empleado (empleg,empfecbaja,empest,empfaltagr,ternro,empremu"
                StrSql = StrSql & ") VALUES "
                StrSql = StrSql & "(" & Clon_NroLeg & ", "
                StrSql = StrSql & ConvFecha(Fecha_Egreso) & ", "
                StrSql = StrSql & rs_emple!empest & ", "
                
                If Crea_Fases Then
                    StrSql = StrSql & ConvFecha(Fecha_Ingreso) & ", "
                Else
                    If EsNulo(rs_emple!empfaltagr) Then
                        StrSql = StrSql & "NULL" & ", "
                    Else
                        StrSql = StrSql & ConvFecha(Fecha_Ingreso) & ", "
                    End If
                End If
                StrSql = StrSql & Tercero_Creado & ", "
                                
                If EsNulo(rs_emple!empremu) Then
                    StrSql = StrSql & "NULL" & ")"
                Else
                    StrSql = StrSql & rs_emple!empremu & ")"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 2) & "sim_empleado creado"
        
        
                'c)  'Creo la fase
                If Crea_Fases Then
                
                    'StrSql = "DELETE FROM sim_fases  "
                    'StrSql = StrSql & " Where empleado =  " & rs_emple!ternro
                    'objConn.Execute StrSql, , adExecuteNoRecords
                
                    'Inserto las Fases
                    'StrSql = " INSERT INTO sim_fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                    'StrSql = StrSql & " VALUES( " & Tercero_Creado & "," & ConvFecha(Fecha_Ingreso) & "," & ConvFecha(Fecha_Egreso) & ","
                    ''If nro_causabaja <> 0 Then
                    ''    StrSql = StrSql & nro_causabaja
                    ''    StrSql = StrSql & ",0,-1,-1,-1,-1,-1)"  ' estado fase=0  - no mira ter_estado
                    ''Else
                    '    StrSql = StrSql & "null"
                    '    'StrSql = StrSql & "," & ter_estado & ",-1,-1,-1,-1,-1)"
                    '    'Pongo el estado siempre en -1
                    '    StrSql = StrSql & ",-1,-1,-1,-1,-1,-1)"
                    ''End If
                    
                    Call CalcularFase(fasnro)
                    StrSql = " INSERT INTO sim_fases(fasnro,empleado,altfec,bajfec,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                    StrSql = StrSql & " VALUES( " & fasnro & "," & Tercero_Creado & "," & ConvFecha(Fecha_Ingreso) & "," & ConvFecha(Fecha_Egreso)
                    StrSql = StrSql & ",-1,-1,-1,-1,-1,-1)"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 2) & "sim_fases creada"
                    
                    'If nro_causabaja <> 0 Then
                    '    Call AsignarSitRevistaAsoc(nro_causabaja, NroTercero, F_Baja)
                    'End If
                End If
        
        
                'd)  'Creo las estructuras
                
               ' If Crea_His_Estructura Then MDF 30/06/2015
                    
                    'Flog.writeline "Borrado de sim_his_estructura para el empleado"
                    'StrSql = "DELETE FROM sim_his_estructura  "
                    'StrSql = StrSql & " Where ternro =  " & rs_emple!ternro
                    'objConn.Execute StrSql, , adExecuteNoRecords
        
                    Aux_Fecha = CDate(Now)
                    
                    Select Case Modelo_Origen
                    Case 1: 'Empleado
                        StrSql = " SELECT * FROM his_estructura WHERE ternro =  " & rs_emple!Ternro
                    Case Else
                        StrSql = " SELECT distinct * FROM sim_his_estructura WHERE ternro =  " & rs_emple!Ternro
                    End Select
                    
                    If Not Crea_His_Estructura Then ' MDF 30/6/2015
                      StrSql = StrSql & " AND (( htetdesde <= " & ConvFecha(Aux_Fecha)
                      StrSql = StrSql & " AND (htethasta >= " & ConvFecha(Aux_Fecha) & " OR htethasta IS NULL))"
                      StrSql = StrSql & " or( htetdesde > " & ConvFecha(Aux_Fecha) & " and ( htethasta > " & ConvFecha(Aux_Fecha) & " or htethasta is null)))"
                    End If 'MDF 30/6/2015
                    StrSql = StrSql & "order by tenro, estrnro"
                    
                    OpenRecordset StrSql, rs_Consulta
                    If rs_Consulta.EOF Then
                        Flog.writeline Espacios(Tabulador * 2) & "ERROR No se econtro el empleado en la tabla His_estructura"
                    End If
                    Flog.writeline "---------------------------------------"
                    Flog.writeline "Estructuras seleccionadas:"
                    Flog.writeline StrSql
                    Flog.writeline "---------------------------------------"
                    
                    Do While Not rs_Consulta.EOF
                        StrSql = "INSERT INTO sim_his_estructura (tenro,ternro,estrnro,htetdesde,htethasta,hismotivo,"
                        StrSql = StrSql & " tipmotnro) Values "
                        StrSql = StrSql & "(" & rs_Consulta!tenro & ", " & Tercero_Creado & ", " & rs_Consulta!Estrnro & ", "
                       '---------MDF 01/07/2015
                        If Not Crea_His_Estructura Then
                         StrSql = StrSql & ConvFecha(Fecha_Ingreso) & ", "
                         StrSql = StrSql & ConvFecha(Fecha_Egreso) & ", "
                        Else
                          StrSql = StrSql & ConvFecha(rs_Consulta!htetdesde) & ", "
                          If Not IsNull(rs_Consulta!htethasta) Then
                            StrSql = StrSql & ConvFecha(rs_Consulta!htethasta) & ", "
                          Else
                            StrSql = StrSql & ConvFecha(Fecha_Egreso) & ", "
                          End If
                        End If
                        '---------MDF 01/07/2015
                        StrSql = StrSql & "'Clonacion', "
                        If EsNulo(rs_Consulta!tipmotnro) Then
                            StrSql = StrSql & "0) "
                        Else
                           StrSql = StrSql & rs_Consulta!tipmotnro & ")"
                        End If
                        objConn.Execute StrSql, , adExecuteNoRecords
                        rs_Consulta.MoveNext
                        Flog.writeline "Inserto: ------->" & StrSql
                    Loop
                    
               ' End If MDF 30/06/2015
               
                
                If Crea_Nov Then
                  Flog.writeline Espacios(Tabulador * 2) & "busca las novedades"
                  Select Case Modelo_Origen
                    Case 1: 'Empleado
                        StrSql = "SELECT * FROM novemp  "
                        StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
                        'StrSql = StrSql & " AND nevigencia = 0 "
                       
                    Case Else
                       StrSql = "SELECT * FROM sim_novemp  "
                       StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
                       'StrSql = StrSql & " AND nevigencia = 0 "
                    End Select
                    OpenRecordset StrSql, rs_Consulta
                    
                    If rs_Consulta.EOF Then
                        Flog.writeline Espacios(Tabulador * 2) & "No hay novedades para copiar"
                    End If
                    
                    Do While Not rs_Consulta.EOF
                        'controlo los que pueden ser nulos
                        If EsNulo(rs_Consulta!nevalor) Then
                            nevalor = "Null"
                        Else
                            nevalor = rs_Consulta!nevalor
                        End If
                    
                        If EsNulo(rs_Consulta!nedesde) Then
                            nedesde = "Null"
                        Else
                            nedesde = ConvFecha(rs_Consulta!nedesde)
                        End If
                            
                        If EsNulo(rs_Consulta!nehasta) Then
                            nehasta = "Null"
                        Else
                            nehasta = ConvFecha(rs_Consulta!nehasta)
                        End If
                            
                        If EsNulo(rs_Consulta!neretro) Then
                            neretro = "Null"
                        Else
                            neretro = ConvFecha(rs_Consulta!neretro)
                        End If
                            
                        If EsNulo(rs_Consulta!nepliqdesde) Then
                            nepliqdesde = "Null"
                        Else
                            nepliqdesde = rs_Consulta!nepliqdesde
                        End If
                            
                        If EsNulo(rs_Consulta!nepliqhasta) Then
                            nepliqhasta = "Null"
                        Else
                            nepliqhasta = rs_Consulta!nepliqhasta
                        End If
                        
                        If EsNulo(rs_Consulta!Pronro) Then
                            pronro1 = "Null"
                        Else
                            pronro1 = rs_Consulta!Pronro
                        End If
                        
                        If EsNulo(rs_Consulta!netexto) Then
                            netexto = "Null"
                        Else
                            netexto = "'" & rs_Consulta!netexto & "'"
                        End If
                            
                        If EsNulo(rs_Consulta!tipmotnro) Then
                            tipmotnro = "Null"
                        Else
                            tipmotnro = rs_Consulta!tipmotnro
                        End If
                            
                        If EsNulo(rs_Consulta!motivo) Then
                            motivo = "Null"
                        Else
                            motivo = "'" & rs_Consulta!motivo & "'"
                        End If
                        
                        StrSql1 = "INSERT INTO sim_novemp (concnro,tpanro,empleado,nevalor,nevigencia,nedesde,nehasta,neretro,nepliqdesde,nepliqhasta,"
                        StrSql1 = StrSql1 & " pronro,netexto,tipmotnro,motivo) Values "
                        StrSql1 = StrSql1 & "(" & rs_Consulta!ConcNro & ", " & rs_Consulta!tpanro & ", " & Tercero_Creado & ", " & nevalor & ", "
                        StrSql1 = StrSql1 & rs_Consulta!nevigencia & ", "
                        StrSql1 = StrSql1 & nedesde & ", "
                        StrSql1 = StrSql1 & nehasta & ", "
                        StrSql1 = StrSql1 & neretro & ", " & nepliqdesde & ", " & nepliqhasta & ", "
                        StrSql1 = StrSql1 & pronro1 & ", " & netexto & ", " & tipmotnro & ", "
                        StrSql1 = StrSql1 & motivo & ")"
                        objConn.Execute StrSql1, , adExecuteNoRecords
                        
                        Flog.writeline Espacios(Tabulador * 2) & "se inserto la novedad " & rs_Consulta!ConcNro
                        rs_Consulta.MoveNext
                    Loop
                        Flog.writeline Espacios(Tabulador * 2) & "se copiaron las novedades"
                        rs_Consulta.Close
                        
                    'novedades de ajustes
                    Flog.writeline Espacios(Tabulador * 2) & "busca las novedades de ajuste"
                    Select Case Modelo_Origen
                        Case 1: 'Empleado
                            StrSql = "SELECT * FROM novaju  "
                            StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
                            'StrSql = StrSql & " AND nevigencia = 0 "
                       
                        Case Else
                           StrSql = "SELECT * FROM sim_novaju  "
                           StrSql = StrSql & " Where empleado =  " & rs_emple!Ternro
                           'StrSql = StrSql & " AND nevigencia = 0 "
                        End Select
                        OpenRecordset StrSql, rs_Consulta
                        
                        If rs_Consulta.EOF Then
                            Flog.writeline Espacios(Tabulador * 2) & "No hay novedades de ajuste para copiar"
                        End If
                        Do While Not rs_Consulta.EOF
                            StrSql1 = "INSERT INTO sim_novaju (concnro,empleado,navalor,navigencia,nadesde,nahasta,naretro,naajuste,pronro,"
                            StrSql1 = StrSql1 & " natexto,napliqdesde,napliqhasta) Values "
                            StrSql1 = StrSql1 & "(" & rs_Consulta!ConcNro & ", " & Tercero_Creado & ", "
                            If EsNulo(rs_Consulta!navalor) Then
                                StrSql1 = StrSql1 & "NULL" & ", "
                            Else
                                StrSql1 = StrSql1 & rs_Consulta!navalor & ", "
                            End If
                            StrSql1 = StrSql1 & rs_Consulta!navigencia & ", "
                            If EsNulo(rs_Consulta!nadesde) Then
                                StrSql1 = StrSql1 & "NULL" & ", "
                            Else
                                StrSql1 = StrSql1 & ConvFecha(rs_Consulta!nadesde) & ", "
                            End If
                            
                            If EsNulo(rs_Consulta!nahasta) Then
                                StrSql1 = StrSql1 & "NULL" & ", "
                            Else
                                StrSql1 = StrSql1 & ConvFecha(rs_Consulta!nahasta) & ", "
                            End If
                            
                            If EsNulo(rs_Consulta!naretro) Then
                                StrSql1 = StrSql1 & "NULL" & ", "
                            Else
                                StrSql1 = StrSql1 & ConvFecha(rs_Consulta!naretro) & ", "
                            End If
                            
                            StrSql1 = StrSql1 & rs_Consulta!naajuste & ", "
                            If EsNulo(rs_Consulta!Pronro) Then
                                StrSql1 = StrSql1 & "NULL" & ", "
                            Else
                                StrSql1 = StrSql1 & rs_Consulta!Pronro & ", "
                            End If
                            
                            If EsNulo(rs_Consulta!natexto) Then
                                StrSql1 = StrSql1 & "NULL" & ", "
                            Else
                                StrSql1 = StrSql1 & "'" & rs_Consulta!natexto & "', "
                            End If
                            
                            If EsNulo(rs_Consulta!napliqdesde) Then
                                StrSql1 = StrSql1 & "NULL" & ", "
                            Else
                                StrSql1 = StrSql1 & rs_Consulta!napliqdesde & ", "
                            End If
                            
                            If EsNulo(rs_Consulta!napliqhasta) Then
                                StrSql1 = StrSql1 & "NULL" & ") "
                            Else
                                StrSql1 = StrSql1 & rs_Consulta!napliqhasta & ") "
                            End If
                            
                            objConn.Execute StrSql1, , adExecuteNoRecords
                            Flog.writeline Espacios(Tabulador * 2) & "se copio las novedad de ajuste " & rs_Consulta!ConcNro
                            rs_Consulta.MoveNext
                        Loop
                            Flog.writeline Espacios(Tabulador * 2) & "se copiaron las novedades de ajuste"
                            rs_Consulta.Close
                    End If
        
                
                'Revisar si es necesario el resto de los datos
                'Novedades
                'Novedades de ajuste
        
        
            MyCommitTrans
        
    
        '-----------------------------------------------------------------------------------------------
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
                  
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                     ", bprcempleados ='" & CStr(I) & "' WHERE bpronro = " & bpronro
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
        '-----------------------------------------------------------------------------------------------
        Flog.writeline Espacios(Tabulador * 1) & "Clon " & I & "creado. - " & "Legajo" & Clon_NroLeg & " " & Clon_Nombre & " " & Clon_Apellido
    Next I
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontraron el empleado modelo."
End If


'============================================================================================


'===============================================================================
FinClonado:
    If rs_emple.State = adStateOpen Then rs_emple.Close
    Set rs_emple = Nothing
Exit Sub


CE:
    HuboError = True
    Flog.writeline "=================================================================="
    Flog.writeline "Error (sub Clonado): " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyRollbackTrans
    
    TiempoAcumulado = GetTickCount
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((Cantidad_Clones - I) * 100) / Cantidad_Clones) & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                     ", bprcempleados ='" & CStr(I) & "' WHERE bpronro = " & bpronro
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    GoTo FinClonado
End Sub


Public Sub CalcularLegajo(ByRef Legajo As String)
Dim rs_leg As New ADODB.Recordset
Dim rs_emp As New ADODB.Recordset
Dim NroLegajo As Long
Dim Continuar As Boolean

    
        StrSql = "SELECT MAX(empleg) AS ProxLegajo FROM empleado"
        OpenRecordset StrSql, rs_leg
        If Not EsNulo(rs_leg!ProxLegajo) Then
            NroLegajo = rs_leg!ProxLegajo + 1
        Else
            NroLegajo = 1
        End If
        
        
        Continuar = True
                
        Do While Continuar
            StrSql = "SELECT empleg FROM sim_empleado WHERE empleg = " & NroLegajo
            OpenRecordset StrSql, rs_emp
        
            If rs_emp.EOF Then
                Continuar = False
            Else
                NroLegajo = NroLegajo + 1
            End If
        Loop
        Legajo = Str(NroLegajo)
        
If rs_leg.State = adStateOpen Then rs_leg.Close
If rs_emp.State = adStateOpen Then rs_emp.Close

Set rs_leg = Nothing
Set rs_emp = Nothing
End Sub


Public Sub CalcularLegajoSim(ByRef Legajo As String)
Dim rs_leg As New ADODB.Recordset
Dim rs_emp As New ADODB.Recordset
Dim NroLegajo As Long
Dim Continuar As Boolean

    
        StrSql = "SELECT MAX(empleg) AS ProxLegajo FROM sim_empleado"
        OpenRecordset StrSql, rs_leg
        If Not EsNulo(rs_leg!ProxLegajo) Then
            NroLegajo = rs_leg!ProxLegajo + 1
        Else
            NroLegajo = 1
        End If
        
        
        Continuar = True
                
        Do While Continuar
            StrSql = "SELECT empleg FROM sim_empleado WHERE empleg = " & NroLegajo
            OpenRecordset StrSql, rs_emp
        
            If rs_emp.EOF Then
                Continuar = False
            Else
                NroLegajo = NroLegajo + 1
            End If
        Loop
        Legajo = Str(NroLegajo)
        
If rs_leg.State = adStateOpen Then rs_leg.Close
If rs_emp.State = adStateOpen Then rs_emp.Close

Set rs_leg = Nothing
Set rs_emp = Nothing
End Sub


Public Sub CalcularFase(ByRef fasnro As Long)
Dim rs_Fases As New ADODB.Recordset
Dim rs_F As New ADODB.Recordset
Dim NroFase As Long
Dim Continuar As Boolean

    
        StrSql = "SELECT MAX(fasnro) AS Prox FROM sim_fases"
        OpenRecordset StrSql, rs_Fases
        If Not EsNulo(rs_Fases!Prox) Then
            NroFase = rs_Fases!Prox + 1
        Else
            NroFase = 1
        End If
        
        
        Continuar = True
        Do While Continuar
            StrSql = "SELECT fasnro FROM sim_fases WHERE fasnro = " & NroFase
            OpenRecordset StrSql, rs_F
        
            If rs_F.EOF Then
                Continuar = False
            Else
                NroFase = NroFase + 1
            End If
        Loop
        fasnro = NroFase
        
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_F.State = adStateOpen Then rs_F.Close

Set rs_Fases = Nothing
Set rs_F = Nothing
End Sub

