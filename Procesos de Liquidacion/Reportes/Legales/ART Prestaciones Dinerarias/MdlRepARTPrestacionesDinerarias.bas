Attribute VB_Name = "MdlRepARTPrestacionesDinerarias"
'Global Const Version = "1.01" ' Lisandro Moro
'Global Const FechaModificacion = "26/11/2008"
'Global Const UltimaModificacion = ""    'Se agrego la clase es feriado.
'                                        'Se agrego la captura de errores y la salida por flog.
'                                        'Se agregaron validaciones.

Global Const Version = "1.02" ' Cesar Stankunas
Global Const FechaModificacion = "04/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

Option Explicit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reportes.
' Autor      : FGZ
' Fecha      : 02/03/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
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
    
    Nombre_Arch = PathFLog & "Reporte_ART_Prestaciones_Dinerarias" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 43 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    Set rs_batch_proceso = Nothing
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        Flog.writeline " Proceso finalizado correctamente. "
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Set objConn = Nothing
    Flog.Close

End Sub


Public Sub ArtLpd02(ByVal bpronro As Long, ByVal EmpTer As Long, ByVal NroAcc As Long, ByVal FirmApenom As String, _
                    ByVal FirmTipoNroDoc As String, _
                    ByVal FirmCargo As String, _
                    ByVal FirmLugar As String, _
                    ByVal Empresa As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de ART Prestaciones Dinertarias
' Autor      : FGZ
' Fecha      : 07/05/2004
' Ult. Mod   : 26/11/2008 - Lisandro Moro - Se agrego le captura de error y la salida al flog
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim tdlicART       As Integer
Dim tdlicEmp       As Integer
Dim ConPorcRed     As Long
Dim ConJub         As Long
Dim ConAsigFam     As Long
Dim ConFondoNac    As Long
Dim ConINSSJP      As Long
Dim ConOS          As Long
Dim ConPrenatal    As Long
Dim ConHijo        As Long
Dim ConHijoDisc    As Long
Dim ConAyEsc       As Long
Dim ConMaternidad  As Long
Dim ConNacimiento  As Long
Dim ConAdopcion    As Long
Dim ConMatrimonio  As Long
Dim cantdias       As Integer
Dim AcuRem         As Long
Dim AcuDias        As Long
Dim ExisteAcuDias  As Boolean
Dim Des_Mes        As String

Dim tipo5 As String
Dim tipo6 As String
Dim tipo7 As String
Dim tipo8 As String
Dim tipo9 As String
Dim tipo10 As String
Dim tipo11 As String
Dim tipo12 As String
Dim tipo13 As String
Dim tipo14 As String
Dim tipo15 As String
Dim tipo16 As String
Dim tipo17 As String
Dim tipo18 As String


'auxiliares
Dim Aux_Empdor_Poliza As String
Dim Aux_Empdor_RazSoc As String
Dim Aux_Empdor_Cuit As String
Dim Aux_Empdor_Tel As String
Dim Aux_FirmApenom As String
Dim Aux_FirmTipoNroDoc As String
Dim Aux_FirmCargo As String
Dim Aux_FirmLugar As String

Dim Aux_Emp_Apeynom As String
Dim Aux_Emp_Domi As String
Dim Aux_Emp_Cuil As String
Dim Aux_Emp_CodPostal As String
Dim Aux_Emp_Localidad As String
Dim Aux_Emp_Prov As String
Dim Aux_Emp_Tel As String
Dim Aux_Emp_Modrellab As String

Dim Aux_Acc_FecIngreso As String
Dim Aux_Acc_Nro As String
Dim Aux_Acc_Fecha As String
Dim Aux_Acc_DiasBaja As Integer
Dim Aux_Acc_DiasART As Integer
Dim Aux_Acc_FecAlta As String
Dim Aux_Acc_FecReintDesde As String
Dim Aux_Acc_FecReintHasta As String

Dim Aux_PorcRed As Single
Dim Aux_PorcJub As Single
Dim Aux_PorcAsigFam As Single
Dim Aux_PorcFondoNac As Single
Dim Aux_PorcINSSJP As Single
Dim Aux_PorcOS As Single

Dim Aux_Det_Importe(12) As Single
Dim Aux_Det_Dias(12) As Integer
Dim Aux_Det_Anio(12) As String
Dim Aux_Det_Mes(12) As String

Dim Aux_MesesART As Integer
Dim Aux_DiasART As Integer

Dim Aux_Prenatal As Single
Dim Aux_Hijo As Single
Dim Aux_HijoDisc As Single
Dim Aux_AyEsc As Single
Dim Aux_Maternidad As Single
Dim Aux_Nacimiento As Single
Dim Aux_Adopcion As Single
Dim Aux_Matrimonio As Single

'Registro
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_Accidente As New ADODB.Recordset
Dim rs_acumulador As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Reporte As New ADODB.Recordset
Dim rs_Empresa As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_DetDom As New ADODB.Recordset
Dim rs_Telefono As New ADODB.Recordset
Dim rs_CUIT As New ADODB.Recordset
Dim rs_Localidad As New ADODB.Recordset
Dim rs_Provincia As New ADODB.Recordset
Dim rs_CUIL As New ADODB.Recordset
Dim rs_TipoContrato As New ADODB.Recordset
Dim rs_FormaLiq As New ADODB.Recordset
Dim rs_Accid_Visita As New ADODB.Recordset
Dim rs_VisitaMedica As New ADODB.Recordset
Dim rs_Lic_Accid As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_CabLiq As New ADODB.Recordset
Dim rs_Rep117 As New ADODB.Recordset

On Error GoTo CE:

StrSql = "Select * FROM reporte where reporte.repnro = 75"
OpenRecordset StrSql, rs_Reporte
If rs_Reporte.EOF Then
    Flog.writeln "El Reporte Numero 75 no ha sido Configurado"
    GoTo CE
End If
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 75 AND confnrocol = 1"
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "Falta configurar la columna 1 del reporte 75"
    GoTo CE
Else
    tdlicART = rs_Confrep!confval
End If
'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 75 AND confnrocol = 2"
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "Falta configurar la columna 2 del reporte 75"
    GoTo CE
Else
    tdlicEmp = rs_Confrep!confval
End If
'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 75 AND confnrocol = 3"
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "Falta configurar la columna 3 del reporte 75"
    GoTo CE
Else
    AcuRem = rs_Confrep!confval
End If

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 75 AND confnrocol = 4 "
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If Not rs_Confrep.EOF Then
    StrSql = "SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval
    OpenRecordset StrSql, rs_acumulador
    If Not rs_acumulador.EOF Then
        ExisteAcuDias = True
        AcuDias = rs_acumulador!acuNro
    Else
        AcuDias = -1
        ExisteAcuDias = False
    End If
Else
    ExisteAcuDias = False
End If

Call Columna(5, ConPorcRed, tipo5)
Call Columna(6, ConJub, tipo6)
Call Columna(7, ConAsigFam, tipo7)
Call Columna(8, ConFondoNac, tipo8)
Call Columna(9, ConINSSJP, tipo9)
Call Columna(10, ConOS, tipo10)
Call Columna(11, ConPrenatal, tipo11)
Call Columna(12, ConHijo, tipo12)
Call Columna(13, ConHijoDisc, tipo13)
Call Columna(14, ConAyEsc, tipo14)
Call Columna(15, ConMaternidad, tipo15)
Call Columna(16, ConNacimiento, tipo16)
Call Columna(17, ConAdopcion, tipo17)
Call Columna(18, ConMatrimonio, tipo18)

StrSql = "SELECT altfec FROM fases WHERE empleado =" & EmpTer & "AND ((fasrecofec = -1) "
StrSql = StrSql & "OR (fasrecofec = 0 AND bajfec is null)) ORDER BY fasrecofec"
OpenRecordset StrSql, rs_Fases
If Not rs_Fases.EOF Then
    Aux_Acc_FecIngreso = rs_Fases!altfec
Else
   Flog.writeline "No se encontró la fecha de alta del empleado"
   GoTo CE 'Exit Sub
End If

StrSql = "SELECT * FROM empleado WHERE ternro =" & EmpTer
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.writeline "No se encontró el empleado"
    GoTo CE 'Exit Sub
End If

Flog.writeline "    Empleado: " & rs_Empleado!empleg & " - " & rs_Empleado!terape & ", " & rs_Empleado!ternom
'If Not rs_Empleado.EOF Then
'    Aux_Acc_FecIngreso = rs_Empleado!empfaltagr
'Else
'    Flog.writeline "No se encontró el empleado"
'    Exit Sub
'End If

StrSql = "SELECT * FROM soaccidente WHERE accnro =" & NroAcc
OpenRecordset StrSql, rs_Accidente
If Not rs_Accidente.EOF Then
    Aux_Empdor_Poliza = rs_Accidente!accpoliza
    Fecha = rs_Accidente!accfecha
    'Aux_Acc_Nro = rs_Accidente!accdescext
    Aux_Acc_Nro = rs_Accidente!accnro
    Aux_Acc_Fecha = rs_Accidente!accfecha
Else
    Flog.writeline "No se encontró el accidente"
    GoTo CE 'Exit Sub
End If

' Comienzo la transaccion
MyBeginTrans


'seteo de las variables de progreso
Progreso = 0
IncPorc = (100 / 8)

'Depuracion del Temporario
StrSql = "DELETE FROM rep117 "
StrSql = StrSql & " WHERE ternro = " & EmpTer
StrSql = StrSql & " AND empresa = " & Empresa
objConn.Execute StrSql, , adExecuteNoRecords


'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DATOS DEL EMPLEADOR *****************************************
'**************************************************************************************************

StrSql = "SELECT * FROM empresa WHERE empnro =" & Empresa
OpenRecordset StrSql, rs_Empresa
If Not rs_Empresa.EOF Then
    StrSql = "SELECT * FROM tercero WHERE ternro =" & rs_Empresa!ternro
    OpenRecordset StrSql, rs_Tercero
    If Not rs_Tercero.EOF Then
        Aux_Empdor_RazSoc = rs_Tercero!terrazsoc
        
        StrSql = " SELECT * FROM detdom " & _
                 " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                 " WHERE cabdom.ternro = " & rs_Tercero!ternro '& " AND " & _
                 '" cabdom.tipnro = 12"
        OpenRecordset StrSql, rs_DetDom
       If Not rs_DetDom.EOF Then
           StrSql = "SELECT * FROM telefono WHERE domnro =" & rs_DetDom!domnro
           StrSql = StrSql & " AND telefono.teldefault = -1"
           OpenRecordset StrSql, rs_Telefono
    
           If Not rs_Telefono.EOF Then
               Aux_Empdor_Tel = rs_Telefono!telnro
           Else
               Flog.writeline "El Registro de Teléfono no está  disponible"
               Aux_Empdor_Tel = ""
               'UNDO, LEAVE.
           End If
            ' Buscar el CUIT
            StrSql = " SELECT * FROM ter_doc WHERE ternro = " & rs_Tercero!ternro
            StrSql = StrSql & " AND ter_doc.tidnro = 6"
            OpenRecordset StrSql, rs_CUIT
            If Not rs_CUIT.EOF Then
                Aux_Empdor_Cuit = rs_CUIT!nrodoc
            Else
                Flog.writeline "El Registro de C.U.I.T. no est  disponible"
                Aux_Empdor_Cuit = ""
            End If
      Else
           Flog.writeline "El Registro de Detalle de Domicilio no está disponible"
           Aux_Empdor_Cuit = ""
           Aux_Empdor_Tel = ""
           'UNDO, LEAVE.
      End If
    Else
        Flog.writeline "El Registro del Tercero no está disponible"
        Aux_Empdor_Cuit = ""
        Aux_Empdor_Tel = ""
    End If
Else
    Flog.writeline "El Registro de Empresa no est  disponible"
    Aux_Empdor_Cuit = ""
    Aux_Empdor_Tel = ""
    'UNDO, LEAVE.
End If

'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DATOS DEL FIRMANTE  *****************************************
'**************************************************************************************************

Aux_FirmApenom = FirmApenom
Aux_FirmTipoNroDoc = FirmTipoNroDoc
Aux_FirmCargo = FirmCargo
Aux_FirmLugar = FirmLugar & " - " + Format(Day(Now) & "/" & Month(Now) & "/" & Year(Now), "dd/mm/yyyy")

'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DATOS DEL TRABAJADOR ****************************************
'**************************************************************************************************

StrSql = "SELECT * FROM tercero WHERE ternro =" & rs_Empleado!ternro
OpenRecordset StrSql, rs_Tercero
If Not rs_Tercero.EOF Then
    Aux_Emp_Apeynom = rs_Tercero!terape + " " + rs_Tercero!ternom

    'Domicilio
    StrSql = " SELECT * FROM detdom "
    StrSql = StrSql & " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro "
    StrSql = StrSql & " WHERE cabdom.ternro = " & rs_Tercero!ternro
    StrSql = StrSql & " AND cabdom.tipnro = 1 AND cabdom.domdefault = -1"
    OpenRecordset StrSql, rs_DetDom
    If Not rs_DetDom.EOF Then
        Aux_Emp_Domi = rs_DetDom!calle & rs_DetDom!nro
        Aux_Emp_CodPostal = nullToString(rs_DetDom!codigopostal)
    Else
        Flog.writeline "El Registro del Domicilio del empleado no está disponible"
        Aux_Emp_Domi = ""
        Aux_Emp_CodPostal = ""
        'UNDO, LEAVE.
    End If

    ' Buscar la localidad
    StrSql = "SELECT * FROM localidad WHERE locnro =" & rs_DetDom!locnro
    OpenRecordset StrSql, rs_Localidad
    If Not rs_Localidad.EOF Then
        Aux_Emp_Localidad = nullToString(rs_Localidad!locdesc)
    Else
        Aux_Emp_Localidad = ""
    End If
    
    ' Buscar la provincia
    StrSql = "SELECT * FROM provincia WHERE provnro =" & rs_DetDom!provnro
    OpenRecordset StrSql, rs_Provincia
    If Not rs_Provincia.EOF Then
        Aux_Emp_Prov = nullToString(rs_Provincia!provdesc)
    Else
        Aux_Emp_Prov = ""
    End If
    
    ' Buscar el telefono
    StrSql = "SELECT * FROM telefono WHERE domnro =" & rs_DetDom!domnro
    StrSql = StrSql & " AND telefono.teldefault = -1"
    OpenRecordset StrSql, rs_Telefono
    If Not rs_Telefono.EOF Then
        Aux_Emp_Tel = nullToString(rs_Telefono!telnro)
    Else
        Aux_Emp_Tel = ""
    End If

    ' Buscar el CUIL
    StrSql = " SELECT * FROM ter_doc cuil WHERE ternro = " & rs_Tercero!ternro
    StrSql = StrSql & " AND cuil.tidnro = 10"
    OpenRecordset StrSql, rs_CUIL
    If Not rs_CUIL.EOF Then
        Aux_Emp_Cuil = nullToString(rs_CUIL!nrodoc)
    Else
        Aux_Emp_Cuil = ""
    End If

    ' Buscar el CONTRATO ACTUAL
    StrSql = " SELECT * FROM his_estructura " & _
             " INNER JOIN tipocont ON his_estructura.estrnro = tipocont.estrnro " & _
             " WHERE ternro = " & rs_Empleado!ternro & " AND " & _
             " tenro = 18 AND " & _
             " (htetdesde <= " & ConvFecha(Fecha) & ") AND " & _
             " ((" & ConvFecha(Fecha) & " <= htethasta) or (htethasta is null))"
    OpenRecordset StrSql, rs_TipoContrato
    
    ' Buscar la Forma de Liq
    StrSql = " SELECT * FROM his_estructura " & _
             " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " & _
             " WHERE ternro = " & rs_Empleado!ternro & " AND " & _
             " his_estructura.tenro = 22 AND " & _
             " (htetdesde <= " & ConvFecha(Fecha) & ") AND " & _
             " ((" & ConvFecha(Fecha) & " <= htethasta) or (htethasta is null))"
    OpenRecordset StrSql, rs_FormaLiq
    
    If Not rs_TipoContrato.EOF And Not rs_FormaLiq.EOF Then
        Aux_Emp_Modrellab = rs_FormaLiq!estrdabr & " - " + rs_TipoContrato!tcdabr
    Else
        Aux_Emp_Modrellab = ""
    End If

Else
    Flog.writeline "El Registro de Tercero no está  disponible"
    Aux_Emp_Apeynom = ""
    Aux_Emp_Domi = ""
    Aux_Emp_CodPostal = ""
    Aux_Emp_Cuil = ""
    Aux_Emp_Modrellab = ""
    'UNDO, LEAVE.
End If

'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DATOS DEL ACCIDENTE *****************************************
'**************************************************************************************************

StrSql = "SELECT * FROM soaccid_visita WHERE accnro = " & rs_Accidente!accnro
OpenRecordset StrSql, rs_Accid_Visita
If Not rs_Accid_Visita.EOF Then
    StrSql = "SELECT * FROM sovisitamedica WHERE vismednro = " & rs_Accid_Visita!visitamed
    OpenRecordset StrSql, rs_VisitaMedica
    If Not rs_VisitaMedica.EOF Then
      Aux_Acc_FecAlta = rs_VisitaMedica!vismedfecha
    Else
        Aux_Acc_FecAlta = ""
        Flog.writeline "El empleado no posee Accidentes  - Visita Medica"
        GoTo CE 'Exit Sub
    End If
Else
    Aux_Acc_FecAlta = ""
    Flog.writeline "El empleado no posee Accidentes - Visita Medica"
    GoTo CE 'Exit Sub
End If
cantdias = 0

StrSql = "SELECT * FROM lic_accid "
StrSql = StrSql & " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_accid.emp_licnro "
StrSql = StrSql & " WHERE accnro =" & rs_Accidente!accnro
StrSql = StrSql & " ORDER BY emp_lic.elfechadesde"
OpenRecordset StrSql, rs_Lic_Accid

If Not rs_Lic_Accid.EOF Then
    rs_Lic_Accid.MoveFirst
    Aux_Acc_FecReintDesde = rs_Lic_Accid!elfechadesde
    
    rs_Lic_Accid.MoveLast
    Aux_Acc_FecReintHasta = rs_Lic_Accid!elfechahasta
Else
    Flog.writeline " El accidente " & rs_Accidente!accnro & " no posee Licencia asociada."
    GoTo CE
End If
Do While Not rs_Lic_Accid.EOF
        
    cantdias = cantdias + rs_Lic_Accid!elcantdias
    
    rs_Lic_Accid.MoveNext
Loop
Aux_Acc_DiasBaja = cantdias
  

cantdias = 0
StrSql = "SELECT * FROM lic_accid "
StrSql = StrSql & " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_accid.emp_licnro "
StrSql = StrSql & " WHERE tdnro =" & tdlicART
StrSql = StrSql & " AND accnro =" & rs_Accidente!accnro
OpenRecordset StrSql, rs_Lic_Accid
Do While Not rs_Lic_Accid.EOF
        
    cantdias = cantdias + rs_Lic_Accid!elcantdias
    
    rs_Lic_Accid.MoveNext
Loop
Aux_Acc_DiasART = cantdias

'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DETALLE DE LAS REMUNERACIONES *******************************
'**************************************************************************************************

Dim I     As Integer
Dim mes   As Integer
Dim Anio  As Integer
Dim Total As Single
Dim Cantidad As Integer

    I = 1
    Total = 0
    Cantidad = 0
    
    Anio = IIf(Month(rs_Accidente!accfecha) = 1, Year(rs_Accidente!accfecha) - 1, Year(rs_Accidente!accfecha))
    mes = IIf(Month(rs_Accidente!accfecha) = 1, 12, Month(rs_Accidente!accfecha) - 1)
    
    Aux_PorcRed = 0
    Aux_PorcJub = 0
    Aux_PorcAsigFam = 0
    Aux_PorcFondoNac = 0
    Aux_PorcINSSJP = 0
    Aux_PorcOS = 0

    Do While I <= 12
        StrSql = "SELECT * FROM periodo "
        StrSql = StrSql & " WHERE pliqanio = " & Anio
        StrSql = StrSql & " AND pliqmes = " & I
        OpenRecordset StrSql, rs_Periodo
    
        If Not rs_Periodo.EOF Then
            'busco el importe para el acumulador AcumRem
            StrSql = "SELECT * FROM proceso "
            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
            StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " WHERE acu_liq.acunro = " & AcuRem
            StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
            StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
            OpenRecordset StrSql, rs_Acu_Liq
            
            Do While Not rs_Acu_Liq.EOF
               Total = Total + rs_Acu_Liq!almonto
               
               rs_Acu_Liq.MoveNext
             Loop
             
             'busco la cantidad para el acumulador AcumDias
            StrSql = "SELECT * FROM proceso "
            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
            StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " WHERE acu_liq.acunro = " & AcuDias
            StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
            StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
            OpenRecordset StrSql, rs_Acu_Liq
            
            Do While Not rs_Acu_Liq.EOF
               Cantidad = Cantidad + rs_Acu_Liq!alcant
               
               rs_Acu_Liq.MoveNext
             Loop
             
             ' busco el porcentaje de reduccion
             If Not IsNull(ConPorcRed) And Aux_PorcRed = 0 Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConPorcRed
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcRed = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
              
             'busco el porcentaje de Jubilacion
             If Not IsNull(ConJub) And Aux_PorcJub = 0 Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConJub
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcJub = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
             
             'busco el porcentaje de Asignaciones Familiares
             If Not IsNull(ConAsigFam) And Aux_PorcAsigFam = 0 Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConAsigFam
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcAsigFam = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
             
             'busco el porcentaje de Fondo Nacional de Desempleo
             If Not IsNull(ConFondoNac) And Aux_PorcFondoNac = 0 Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConFondoNac
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcFondoNac = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
             
             'busco el porcentaje de INSSJP
             If Not IsNull(ConINSSJP) And Aux_PorcINSSJP = 0 Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConINSSJP
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcINSSJP = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
             
            'busco el porcentaje de Obra Social
            If Not IsNull(ConOS) And Aux_PorcOS = 0 Then
                If tipo10 = "CO" Then
                    StrSql = "SELECT * FROM proceso "
                    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    StrSql = StrSql & " WHERE detliq.concnro = " & ConOS
                    StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                    OpenRecordset StrSql, rs_Detliq
                    
                    Do While Not rs_Detliq.EOF
                       Aux_PorcOS = rs_Detliq!dlicant
                       
                       rs_Detliq.MoveNext
                    Loop
                End If
                If tipo10 = "AC" Then
                    StrSql = "SELECT * FROM proceso "
                    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                    StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    StrSql = StrSql & " WHERE acu_liq.acunro = " & ConOS
                    StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                    OpenRecordset StrSql, rs_Acu_Liq
                    
                    Do While Not rs_Acu_Liq.EOF
                       Aux_PorcOS = rs_Acu_Liq!alcant
                       
                       rs_Acu_Liq.MoveNext
                    Loop
                End If
            End If
               
        End If 'rs_periodo.eof
        
        Call BusMes(mes, Des_Mes)
        
        Aux_Det_Importe(I) = Total
        Aux_Det_Dias(I) = IIf(ExisteAcuDias, Cantidad, 30)
        Aux_Det_Anio(I) = CStr(Anio)
        Aux_Det_Mes(I) = Des_Mes
        I = I + 1
        Total = 0
        Cantidad = 0
        Anio = IIf(mes = 1, Anio - 1, Anio)
        mes = IIf(mes = 1, 12, mes - 1)
    Loop


'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DETALLE DE LAS PRESTACIONES DINERARIAS **********************
'**************************************************************************************************
Dim fecdesde As Date
Dim fechasta As Date
Dim cantMesesART As Integer
Dim cantDiasART As Integer

cantDiasART = 0
cantMesesART = 0

StrSql = "SELECT * FROM lic_accid "
StrSql = StrSql & " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_accid.emp_licnro "
StrSql = StrSql & " WHERE tdnro =" & tdlicART
StrSql = StrSql & " AND accnro =" & rs_Accidente!accnro
OpenRecordset StrSql, rs_Lic_Accid
Do While Not rs_Lic_Accid.EOF
        
    If Year(rs_Lic_Accid!elfechadesde) = Year(rs_Lic_Accid!elfechahasta) And Month(rs_Lic_Accid!elfechadesde) = Month(rs_Lic_Accid!elfechahasta) Then
        fecdesde = CDate("01/" & Month(rs_Lic_Accid!elfechadesde) & "/" & Year(rs_Lic_Accid!elfechadesde))
        fechasta = CDate("01/" & Month(rs_Lic_Accid!elfechahasta) & "/" & Year(rs_Lic_Accid!elfechahasta))
        cantDiasART = cantDiasART + (rs_Lic_Accid!elfechahasta - rs_Lic_Accid!elfechadesde + 1)
    Else
        fecdesde = IIf(Month(rs_Lic_Accid!elfechadesde) = 12, CDate("01/01/" & Year(rs_Lic_Accid!elfechadesde) + 1), CDate("01/" & Month(rs_Lic_Accid!elfechadesde) + 1 & "/" & Year(rs_Lic_Accid!elfechadesde)))
        fechasta = IIf(Month(rs_Lic_Accid!elfechahasta) = 1, CDate("01/12/" & Year(rs_Lic_Accid!elfechahasta) - 1), CDate("01/" & Month(rs_Lic_Accid!elfechahasta) - 1 & "/" & Year(rs_Lic_Accid!elfechahasta)))
        cantDiasART = cantDiasART + ((Day(fecdesde - 1) - Day(rs_Lic_Accid!elfechadesde) + 1) + (Day(rs_Lic_Accid!elfechahasta) - Day(fechasta) + 1))
        cantMesesART = cantMesesART + (CInt((fechasta - fecdesde) / 30) + 1)
    End If
        
    rs_Lic_Accid.MoveNext
Loop
Aux_DiasART = cantDiasART
Aux_MesesART = cantMesesART

'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords
       
'**************************************************************************************************
'************************************ ASIGNACIONES FAMILIARES  ************************************
'**************************************************************************************************
'Dim Fecha As Date
Dim Monto  As Single
       
Aux_Prenatal = 0
Aux_Hijo = 0
Aux_HijoDisc = 0
Aux_AyEsc = 0
Aux_Maternidad = 0
Aux_Nacimiento = 0
Aux_Adopcion = 0
Aux_Matrimonio = 0
       
Fecha = CDate("01/" & Month(CDate(Aux_Acc_FecReintDesde)) & "/" & Year(CDate(Aux_Acc_FecReintDesde)))
Do While Fecha <= CDate(Aux_Acc_FecReintHasta)
   
    StrSql = "SELECT distinct cabliq.cliqnro FROM periodo "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " WHERE periodo.pliqanio =" & Year(Fecha)
    StrSql = StrSql & " AND periodo.pliqmes =" & Month(Fecha)
    StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
    OpenRecordset StrSql, rs_CabLiq
    
    Do While Not rs_CabLiq.EOF
        Call AsigFam(rs_CabLiq!cliqnro, ConPrenatal, Monto)
            Aux_Prenatal = Aux_Prenatal + Monto
        
        Call AsigFam(rs_CabLiq!cliqnro, ConHijo, Monto)
            Aux_Hijo = Aux_Hijo + Monto

        Call AsigFam(rs_CabLiq!cliqnro, ConHijoDisc, Monto)
            Aux_HijoDisc = Aux_HijoDisc + Monto

        Call AsigFam(rs_CabLiq!cliqnro, ConAyEsc, Monto)
            Aux_AyEsc = Aux_AyEsc + Monto

        Call AsigFam(rs_CabLiq!cliqnro, ConMaternidad, Monto)
            Aux_Maternidad = Aux_Maternidad + Monto
    
        Call AsigFam(rs_CabLiq!cliqnro, ConNacimiento, Monto)
            Aux_Nacimiento = Aux_Nacimiento + Monto

        Call AsigFam(rs_CabLiq!cliqnro, ConAdopcion, Monto)
            Aux_Adopcion = Aux_Adopcion + Monto

        Call AsigFam(rs_CabLiq!cliqnro, ConMatrimonio, Monto)
            Aux_Matrimonio = Aux_Matrimonio + Monto
       
       rs_CabLiq.MoveNext
    Loop
   
    Fecha = IIf(Month(Fecha) = 12, CDate("01/01/" & Year(Fecha) + 1), CDate("01/" & Month(Fecha) + 1 & "/" & Year(Fecha)))
Loop
       


'Inserto todos los campos
'Si no existe el rep117
StrSql = "SELECT * FROM rep117 "
StrSql = StrSql & " WHERE ternro = " & EmpTer
StrSql = StrSql & " AND bpronro = " & bpronro
StrSql = StrSql & " AND empresa = " & Empresa
OpenRecordset StrSql, rs_Rep117

If rs_Rep117.EOF Then
    'Inserto
    StrSql = "INSERT INTO rep117 (bpronro,empresa,iduser,fecha,hora,"
    StrSql = StrSql & "ternro,empdor_razsoc,empdor_cuit,empdor_poliza,empdor_tel, "
    StrSql = StrSql & "emp_apeynom,emp_cuil,emp_domi,emp_localidad,emp_codpostal,emp_prov,emp_tel,emp_modrellab, "
    StrSql = StrSql & "acc_nro, acc_fecha, acc_fecalta, acc_fecreintdesde, acc_fecreinthasta, acc_diasbaja, acc_diasart,acc_fecingreso, "
    StrSql = StrSql & "det_anio_1, det_anio_2, det_anio_3, det_anio_4, det_anio_5,det_anio_6, det_anio_7, det_anio_8, det_anio_9, det_anio_10, det_anio_11, det_anio_12, "
    StrSql = StrSql & "det_mes_1, det_mes_2, det_mes_3, det_mes_4, det_mes_5,det_mes_6, det_mes_7, det_mes_8, det_mes_9, det_mes_10, det_mes_11, det_mes_12, "
    StrSql = StrSql & "det_importe_1, det_importe_2, det_importe_3, det_importe_4, det_importe_5,det_importe_6, det_importe_7, det_importe_8, det_importe_9, det_importe_10, det_importe_11, det_importe_12, "
    StrSql = StrSql & "det_dias_1, det_dias_2, det_dias_3, det_dias_4, det_dias_5,det_dias_6, det_dias_7, det_dias_8, det_dias_9, det_dias_10, det_dias_11, det_dias_12, "
    StrSql = StrSql & "mesesart,diasart, "
    StrSql = StrSql & "porcred, porcjub, porcasigfam, porcfondonac, porcinssjp, porcos, "
    StrSql = StrSql & "prenatal, hijo, hijodisc, ayesc, maternidad, nacimiento, adopcion, matrimonio, "
    StrSql = StrSql & "firmapenom, firmtiponrodoc, firmcargo,firmlugar "
    
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & bpronro & ","
    StrSql = StrSql & Empresa & ","
    StrSql = StrSql & "'" & IdUser & "',"
    StrSql = StrSql & ConvFecha(Fecha) & ","
    StrSql = StrSql & "'" & Format(Hora, "hh:mm:ss") & "',"
    
    StrSql = StrSql & Left(EmpTer, 20) & ","
    StrSql = StrSql & "'" & Left(Aux_Empdor_RazSoc, 50) & "',"
    StrSql = StrSql & "'" & Left(Aux_Empdor_Cuit, 13) & "',"
    StrSql = StrSql & "'" & Left(Aux_Empdor_Poliza, 20) & "',"
    StrSql = StrSql & "'" & Left(Aux_Empdor_Tel, 30) & "',"
    
    StrSql = StrSql & "'" & Left(Aux_Emp_Apeynom, 50) & "',"
    StrSql = StrSql & "'" & Left(Aux_Emp_Cuil, 13) & "',"
    StrSql = StrSql & "'" & Left(Aux_Emp_Domi, 30) & "',"
    StrSql = StrSql & "'" & Left(Aux_Emp_Localidad, 30) & "',"
    StrSql = StrSql & "'" & Left(Aux_Emp_CodPostal, 10) & "',"
    StrSql = StrSql & "'" & Left(Aux_Emp_Prov, 25) & "',"
    StrSql = StrSql & "'" & Left(Aux_Emp_Tel, 20) & "',"
    StrSql = StrSql & "'" & Left(Aux_Emp_Modrellab, 25) & "',"
    
    StrSql = StrSql & "'" & Left(Aux_Acc_Nro, 30) & "',"
    StrSql = StrSql & ConvFecha(Aux_Acc_Fecha) & ","
    StrSql = StrSql & ConvFecha(Aux_Acc_FecAlta) & ","
    StrSql = StrSql & ConvFecha(Aux_Acc_FecReintDesde) & ","
    StrSql = StrSql & ConvFecha(Aux_Acc_FecReintHasta) & ","
    StrSql = StrSql & Aux_Acc_DiasBaja & ","
    StrSql = StrSql & Aux_Acc_DiasART & ","
    StrSql = StrSql & ConvFecha(Aux_Acc_FecIngreso) & ","
    
    For I = 1 To 12
        If IsNull(Aux_Det_Anio(I)) Then
            StrSql = StrSql & "'',"
        Else
            StrSql = StrSql & "'" & Left(Aux_Det_Anio(I), 4) & "',"
        End If
    Next I
    For I = 1 To 12
        If IsNull(Aux_Det_Mes(I)) Then
            StrSql = StrSql & "'',"
        Else
            StrSql = StrSql & "'" & Left(Aux_Det_Mes(I), 10) & "',"
        End If
    Next I
    For I = 1 To 12
        If IsNull(Aux_Det_Importe(I)) Then
            StrSql = StrSql & "null,"
        Else
            StrSql = StrSql & Aux_Det_Importe(I) & ","
        End If
    Next I
    For I = 1 To 12
        If IsNull(Aux_Det_Dias(I)) Then
            StrSql = StrSql & "null,"
        Else
            StrSql = StrSql & Aux_Det_Dias(I) & ","
        End If
    Next I
    
    StrSql = StrSql & Aux_MesesART & ","
    StrSql = StrSql & Aux_DiasART & ","
    
    StrSql = StrSql & Aux_PorcRed & ","
    StrSql = StrSql & Aux_PorcJub & ","
    StrSql = StrSql & Aux_PorcAsigFam & ","
    StrSql = StrSql & Aux_PorcFondoNac & ","
    StrSql = StrSql & Aux_PorcINSSJP & ","
    StrSql = StrSql & Aux_PorcOS & ","
    
    StrSql = StrSql & Aux_Prenatal & ","
    StrSql = StrSql & Aux_Hijo & ","
    StrSql = StrSql & Aux_HijoDisc & ","
    StrSql = StrSql & Aux_AyEsc & ","
    StrSql = StrSql & Aux_Maternidad & ","
    StrSql = StrSql & Aux_Nacimiento & ","
    StrSql = StrSql & Aux_Adopcion & ","
    StrSql = StrSql & Aux_Matrimonio & ","
    
    StrSql = StrSql & "'" & Left(Aux_FirmApenom, 50) & "',"
    StrSql = StrSql & "'" & Left(Aux_FirmTipoNroDoc, 50) & "',"
    StrSql = StrSql & "'" & Left(Aux_FirmCargo, 50) & "',"
    StrSql = StrSql & "'" & Left(Aux_FirmLugar, 50) & "'"
    
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
End If
            
'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords
            

'Fin de la transaccion
MyCommitTrans

If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Accidente.State = adStateOpen Then rs_Accidente.Close
If rs_acumulador.State = adStateOpen Then rs_acumulador.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_DetDom.State = adStateOpen Then rs_DetDom.Close
If rs_Telefono.State = adStateOpen Then rs_Telefono.Close
If rs_CUIT.State = adStateOpen Then rs_CUIT.Close
If rs_Localidad.State = adStateOpen Then rs_Localidad.Close
If rs_Provincia.State = adStateOpen Then rs_Provincia.Close
If rs_CUIL.State = adStateOpen Then rs_CUIL.Close
If rs_TipoContrato.State = adStateOpen Then rs_TipoContrato.Close
If rs_FormaLiq.State = adStateOpen Then rs_FormaLiq.Close
If rs_Accid_Visita.State = adStateOpen Then rs_Accid_Visita.Close
If rs_VisitaMedica.State = adStateOpen Then rs_VisitaMedica.Close
If rs_Lic_Accid.State = adStateOpen Then rs_Lic_Accid.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_CabLiq.State = adStateOpen Then rs_CabLiq.Close
If rs_Rep117.State = adStateOpen Then rs_Rep117.Close

Set rs_Empleado = Nothing
Set rs_Accidente = Nothing
Set rs_acumulador = Nothing
Set rs_Confrep = Nothing
Set rs_Reporte = Nothing
Set rs_Empresa = Nothing
Set rs_Tercero = Nothing
Set rs_DetDom = Nothing
Set rs_Telefono = Nothing
Set rs_CUIT = Nothing
Set rs_Localidad = Nothing
Set rs_Provincia = Nothing
Set rs_CUIL = Nothing
Set rs_TipoContrato = Nothing
Set rs_FormaLiq = Nothing
Set rs_Accid_Visita = Nothing
Set rs_VisitaMedica = Nothing
Set rs_Lic_Accid = Nothing
Set rs_Periodo = Nothing
Set rs_Acu_Liq = Nothing
Set rs_Detliq = Nothing
Set rs_CabLiq = Nothing
Set rs_Rep117 = Nothing

Exit Sub
CE:
    Flog.writeline " Error: " & Err.Description
    If Err.Number Then
        Flog.writeline "SQL ejecutada: " & StrSql
    End If
    
    HuboError = True
    MyRollbackTrans
    Exit Sub

End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer

Dim EmpTer As Long
Dim NroAcc As Long
Dim FirmApenom As String
Dim FirmTipoNroDoc As String
Dim FirmCargo As String
Dim FirmLugar As String
Dim Empresa As Long

Dim Separador As String

Separador = "@"
' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        EmpTer = CLng(Mid(parametros, pos1, pos2))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        NroAcc = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FirmApenom = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FirmTipoNroDoc = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FirmCargo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FirmLugar = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
    End If
End If


Call ArtLpd02(bpronro, EmpTer, NroAcc, FirmApenom, FirmTipoNroDoc, FirmCargo, FirmLugar, Empresa)

End Sub


Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para saber si es el ultimo empleado de la secuencia
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    
    rs.MoveNext
    If rs.EOF Then
        EsElUltimoEmpleado = True
    Else
        If rs!Empleado <> Anterior Then
            EsElUltimoEmpleado = True
        Else
            EsElUltimoEmpleado = False
        End If
    End If
    rs.MovePrevious
End Function


Public Sub Columna(ByVal NroCol As Integer, ByRef nroconc As Long, ByRef TipoNro As String)
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_acumulador As New ADODB.Recordset

    StrSql = "SELECT * FROM confrep WHERE repnro = 75 AND confnrocol = " & NroCol
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    If Not rs_Confrep.EOF Then
        Select Case rs_Confrep!conftipo
            Case "CO":
                StrSql = " SELECT * FROM concepto WHERE conccod = " & rs_Confrep!confval
                OpenRecordset StrSql, rs_Concepto
                If Not rs_Concepto.EOF Then
                    nroconc = rs_Concepto!concnro
                    TipoNro = "CO"
                End If
            Case "AC":
                StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval
                OpenRecordset StrSql, rs_acumulador
                If Not rs_acumulador.EOF Then
                    nroconc = rs_acumulador!acuNro
                    TipoNro = "AC"
                End If
         End Select
    End If
End Sub


Public Sub AsigFam(ByVal cliqnro As Long, ByVal nroconc As Integer, ByRef Valor As Single)
Dim rs_Detliq As New ADODB.Recordset

    Valor = 0
    
    StrSql = "SELECT * FROM detliq "
    StrSql = StrSql & " INNER JOIN cabliq ON detliq.cliqnro = cabliq.pronro "
    StrSql = StrSql & " WHERE detliq.concnro = " & nroconc
    StrSql = StrSql & " AND cabliq.cliqnro = " & cliqnro
    OpenRecordset StrSql, rs_Detliq
    
    If Not rs_Detliq.EOF Then
       Valor = rs_Detliq!dlicant
    End If
    
End Sub
Public Function nullToString(obj As Variant) As String
    If IsNull(obj) Then
        nullToString = ""
    Else
        nullToString = CStr(obj)
    End If
End Function
