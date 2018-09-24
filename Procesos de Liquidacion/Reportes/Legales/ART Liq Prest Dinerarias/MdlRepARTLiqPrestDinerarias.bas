Attribute VB_Name = "MdlRepARTPrestacionesDinerarias"
Option Explicit
'Global Const Version = "1.00"
'Global Const FechaModificacion = "25/09/2006"   'JC
'Global Const UltimaModificacion = " "   '

'Global Const Version = "1.01"
'Global Const FechaModificacion = "29/09/2006"   'MARIANO CAPRIZ
'Global Const UltimaModificacion = " "   'SE LE AGREGO LOG'S, NRO DE VERSION DEL PROCESO Y RUTINAS DE CONTROL DE ERROR

'Global Const Version = "1.02"
'Global Const FechaModificacion = "02/10/2006"   'MARIANO CAPRIZ
'Global Const UltimaModificacion = " "   'SE REALIZARON CORRECCIONES VARIAS EN EL CONTROL DE DATOS Y RUTINAS DE ERRORES


'Global Const Version = "1.03"
'Global Const FechaModificacion = "28/06/2007"   'FGZ
'Global Const UltimaModificacion = " "   'SE REALIZARON CORRECCIONES VARIAS DEL ESTANDAR Y
''                                       en el sub ArtLpd02 - Cuando busca los aguinaldos
''                                       antes los sacaba de la fecha de impresion del recibo
''                                           Aux_agu_mes_1 = Month(rs_Detliq!fechaimp)
''                                           Aux_agu_anio_1 = Year(rs_Detliq!fechaimp)
''                                       ahora lo saco de la fecha hasta del proceso

'Global Const Version = "1.04"
'Global Const FechaModificacion = "26/07/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Correcciones
''               Se cambió la forma de detectar si la caja de jubilacion es Reparto o Capitalicacion
''               Modificaciones varias en el calculo

'Global Const Version = "1.05"
'Global Const FechaModificacion = "17/12/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Correcciones
''               Se cambió la forma de calclular sac proporcionales
''               Modificaciones varias en el calculo
''

'Global Const Version = "1.06"
'Global Const FechaModificacion = "17/12/2008"   'Lisandro Moro
'Global Const UltimaModificacion = " "   'SE REALIZARON CORRECCIONES VARIAS DEL ESTANDAR Y
''                                       Se actualizaron las referencias a los archivos del proyecto
''                                       Se agregaron mas validaciones
''                                       Se valida primero si esta cargado el valor del confrep en confval2
''                                       Se modifico para que busque en todos los conceptos de SAC

'Global Const Version = "1.07"
'Global Const FechaModificacion = "06/01/2009"   'Lisandro Moro
'Global Const UltimaModificacion = " "   'SE busca la fecha de ingreso en la fase y no en la estructura empresa.

'Global Const Version = "1.08" ' Cesar Stankunas
'Global Const FechaModificacion = "04/08/2009"
'Global Const UltimaModificacion = ""    'Encriptacion de string connection


Global Const Version = "1.09" ' Leticia Amadio
Global Const FechaModificacion = "21/09/2010"
Global Const UltimaModificacion = ""    '  se agrego la busqueda del tipo de Contrato Laboral - se agrego la funciond de ConvFecha en ( fecha nac y ingr.)



' ____________________________________________________________________________________






Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String


Public Sub Main()

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
    
   ' carga las configuraciones basicas, formato de fecha, string de conexion,tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    Nombre_Arch = PathFLog & "Reporte_ART_Liq_Prestaciones_Dinerarias" & "-" & NroProcesoBatch & ".log"
        'Nombre_Arch = PathFLog & "Reporte_ART_Prestaciones_Dinerarias" & "-" & NroProcesoBatch & ".log"
        
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha        = " & FechaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 136 AND bpronro =" & NroProcesoBatch
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
    TiempoFinalProceso = GetTickCount
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.writeline "----------------------------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline "Cantidad de Lecturas en BD: " & Cantidad_de_OpenRecordset
    Flog.writeline "----------------------------------------------------------------------------------"
    Flog.Close
    If rs_batch_proceso.State = adStateOpen Then rs_batch_proceso.Close
    If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
    
    Set rs_batch_proceso = Nothing
    Set objConn = Nothing
    Set objconnProgreso = Nothing
Exit Sub

ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
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
    GoTo Fin:
End Sub


Public Sub ArtLpd02(ByVal bpronro As Long, ByVal EmpTer As Long, ByVal NroAcc As Long)

Dim tdlicART       As Integer
Dim tdlicEmp       As Integer
Dim ConPorcRed     As Long
Dim ConJub         As Long
Dim ConAsigFam     As Long
Dim ConFondoNac    As Long
Dim ConINSSJP      As Long
Dim ConOS          As String

Dim ConANSSal      As Long
Dim ConTC          As Long

Dim ConAgu1  As Long
Dim ConAgu2  As Long
Dim ConAguStr As String
Dim ConAgu2Str As String

Dim cantdias       As Integer
Dim AcuRem         As Long
Dim AcuDias        As Long
Dim ExisteAcuDias  As Boolean
Dim Des_Mes        As String

Dim tipo2 As String
Dim tipo3 As String

Dim tipo5 As String
Dim tipo6 As String
Dim tipo7 As String
Dim tipo8 As String
Dim tipo9 As String
Dim tipo10 As String
Dim tipo11 As String
Dim tipo12 As String

Dim aux_tot_importe As Double
Dim aux_tot_dias As Integer


'auxiliares

Dim Aux_Empdor_RazSoc As String
Dim Aux_Empdor_Cuit As String

Dim Aux_Emp_Apeynom As String
Dim Aux_Emp_Cuil As String
Dim Aux_Emp_Tel 'As String
Dim Aux_Emp_Contrato As String ' tipo de contrato

Dim Aux_Emp_fnac
Dim Aux_Emp_RegJub
Dim Aux_emp_AFJP
Dim Aux_emp_CodOS
Dim Aux_emp_OS
Dim Aux_Emp_Osoc
Dim Aux_Emp_fingreso

Dim Aux_Fecha_Desde As Date
Dim Aux_Fecha_Hasta As Date

Dim Aux_Acc_FecIngreso As String
Dim Aux_Acc_Nro As String
Dim Aux_Acc_nrosiniestro As String
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
Dim Aux_PorcANSSal As Single
Dim Aux_PorcTotalC

Dim Aux_agu_mes_1
Dim Aux_agu_anio_1
Dim Aux_agu_importe_1

Dim Aux_agu_mes_2
Dim Aux_agu_anio_2
Dim Aux_agu_importe_2
Dim AsigneA1 As Boolean

Dim EMPRESA

Dim Aux_Det_Importe(12) As Single
Dim Aux_Det_Dias(12) As Integer
Dim Aux_Det_Anio(12) As String
Dim Aux_Det_Mes(12) As String

Dim Aux_MesesART As Integer
Dim Aux_DiasART As Integer

'Registro
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Accidente As New ADODB.Recordset
Dim rs_acumulador As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Reporte As New ADODB.Recordset
Dim rs_Empresa As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
'Dim rs_Localidad As New ADODB.Recordset
'Dim rs_Provincia As New ADODB.Recordset
Dim rs_Accid_Visita As New ADODB.Recordset
Dim rs_VisitaMedica As New ADODB.Recordset
Dim rs_Lic_Accid As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_CabLiq As New ADODB.Recordset
Dim rs_Rep_prestdine As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_OS As New ADODB.Recordset
Dim fecAlta As Date
Dim fecBaja As Date
Dim StrSql2 As String
Dim Sac_Proporcional
Dim Aux_sac_prop_mes
Dim Aux_sac_prop_anio

Flog.writeline "Entro en ArtLpd02"

EMPRESA = 0

StrSql = "Select * FROM reporte where reporte.repnro = 177"
OpenRecordset StrSql, rs_Reporte
If rs_Reporte.EOF Then
    Flog.writeln "El Reporte Numero 177 no ha sido Configurado"
    Exit Sub
End If
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 177 AND confnrocol = 1"
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "Falta configurar la columna 1 del reporte 177" ' xxx cambio columna 3 por 1
    Exit Sub
Else
    If IsNull(rs_Confrep!confval) Then
        AcuRem = rs_Confrep!confval
    Else
        AcuRem = rs_Confrep!confval2
    End If
End If

'Configuracion del Reporte - dias
StrSql = "SELECT * FROM confrep WHERE repnro = 177 AND confnrocol = 4 "
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If Not rs_Confrep.EOF Then
    If IsNull(rs_Confrep!confval) Then
        StrSql = "SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval
    Else
        StrSql = "SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval2
    End If
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

On Error GoTo CE

Flog.writeline
Flog.writeline "Configuracion ----"
'Call Columna(2, ConAgu1, tipo2)
Call ColumnaArr(2, ConAguStr, tipo2)
'Call Columna(3, ConAgu2, tipo3)
Call ColumnaArr(3, ConAgu2Str, tipo3)
Call Columna(5, ConJub, tipo5)
Call Columna(6, ConINSSJP, tipo6)
Call Columna(7, ConFondoNac, tipo7)
Call Columna(8, ConAsigFam, tipo8)
Call Columna(9, ConANSSal, tipo9)
'Call Columna(10, ConOS, tipo10)
'-------------------------------------------------------------------
'Puede Haber Varios Conceptos de OS( todos en columnas 10)
tipo10 = "CO"
ConOS = "0"
StrSql = "SELECT * FROM confrep WHERE repnro = 177 AND confnrocol = 10"
OpenRecordset StrSql, rs_Confrep
Do While Not rs_Confrep.EOF
        tipo10 = rs_Confrep!conftipo
        Select Case rs_Confrep!conftipo
            Case "CO":
                If IsNull(rs_Confrep!confval2) Then
                    StrSql = " SELECT * FROM concepto WHERE conccod = " & rs_Confrep!confval
                Else
                    StrSql = " SELECT * FROM concepto WHERE conccod = '" & rs_Confrep!confval2 & "'"
                End If
                OpenRecordset StrSql, rs_OS
                If Not rs_OS.EOF Then
                    ConOS = ConOS & "," & rs_OS!ConcNro
                End If
            Case "AC":
                If IsNull(rs_Confrep!confval2) Then
                    StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval
                Else
                    StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval2
                End If
                OpenRecordset StrSql, rs_OS
                If Not rs_OS.EOF Then
                    ConOS = ConOS & "," & rs_OS!acuNro
                End If
        End Select
    
    rs_Confrep.MoveNext
Loop
'-------------------------------------------------------------------
Call Columna(11, ConTC, tipo11)
Call Columna(12, ConPorcRed, tipo12)
Flog.writeline

StrSql = "SELECT * FROM empleado WHERE ternro =" & EmpTer
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.writeline "No se encontró el empleado"
    Exit Sub
End If

StrSql = "SELECT * FROM soaccidente WHERE accnro =" & NroAcc
OpenRecordset StrSql, rs_Accidente
If Not rs_Accidente.EOF Then
    Flog.writeline "(rs_Accidente!accfecha - rs_Accidente!accnro - rs_Accidente!accfecha - rs_Accidente!accnrosiniestro) = (" & rs_Accidente!accfecha & " - " & rs_Accidente!accnro & " - " & rs_Accidente!accfecha & " - " & rs_Accidente!accnrosiniestro & ")"
    
    Fecha = rs_Accidente!accfecha
    Aux_Acc_Nro = rs_Accidente!accnro
    Aux_Acc_Fecha = rs_Accidente!accfecha
    
    If IsNull(rs_Accidente!accnrosiniestro) Then
        Flog.writeline "=== NO SE HA CARGADO EL NRO DE SINIESTRO PAR AEL ACCIDENTE NRO: " & NroAcc & " NO SE PUEDE CONTINUAR "
        Flog.writeline "PROCESO ABORTADO ===================================================================================="
        GoTo CE
    Else
        Aux_Acc_nrosiniestro = rs_Accidente!accnrosiniestro
    End If
Else
    Flog.writeline "No se encontró el accidente"
    Aux_Acc_Fecha = 0
    Exit Sub
End If

' Comienzo la transaccion
Flog.writeline "Inicio la Transaccion"
MyBeginTrans

'seteo de las variables de progreso
Progreso = 0
IncPorc = (100 / 8)

'Depuracion del Temporario
Flog.writeline "Depuracion del Temporario"
StrSql = "DELETE FROM Rep_prestdine "
StrSql = StrSql & " WHERE ternro = " & EmpTer
StrSql = StrSql & " and acc_nro = " & Aux_Acc_Nro
objConn.Execute StrSql, , adExecuteNoRecords


'Actualizo el progreso del Proceso
Flog.writeline "Actualizo el progreso del Proceso"
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords


'**************************************************************************************************
'************************************ DATOS DEL EMPLEADOR *****************************************
Flog.writeline "DATOS DEL EMPLEADOR"
'**************************************************************************************************
    StrSql = "select tercero.ternro,terrazsoc,nrodoc,htetdesde "
    StrSql = StrSql & " From his_estructura "
    StrSql = StrSql & " INNER JOIN empresa ON his_estructura.estrnro = empresa.estrnro and his_estructura.tenro = 10"
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empresa.ternro"
    StrSql = StrSql & " INNER JOIN ter_doc on ter_doc.ternro = tercero.ternro  AND tidnro = 6"
    StrSql = StrSql & " Where his_estructura.ternro=" & EmpTer
    
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Aux_Acc_Fecha)
    StrSql = StrSql & " AND (htethasta >= " & ConvFecha(Aux_Acc_Fecha) & " or htethasta is null)"
    
    OpenRecordset StrSql, rs_Empresa
    If Not rs_Empresa.EOF Then
        
       Aux_Empdor_RazSoc = rs_Empresa!terrazsoc
       Aux_Empdor_Cuit = rs_Empresa!nrodoc
       
       'Aux_Emp_fingreso = rs_Empresa!htetdesde 'lisandro moro - 06-01-2009
       EMPRESA = rs_Empresa!ternro
    Else
        Flog.writeline "El Registro de Empresa no est  disponible"
        Aux_Empdor_RazSoc = ""
        Aux_Empdor_Cuit = ""
     
        'UNDO, LEAVE.
    End If
    

'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
Flog.writeline "Actualizo el progreso del Proceso=" & Progreso
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DATOS DEL TRABAJADOR ****************************************
Flog.writeline "DATOS DEL TRABAJADOR"
'**************************************************************************************************

StrSql = "SELECT DISTINCT tercero.terape,tercero.ternom,telefono.telnro,tercero.terfecnac,estr_cod.nrocod"
'FGZ - 24/07/2007 - esto es para chequear el tipo de caja de jubilacion (No se debe chequear por codext)
'StrSql = StrSql & ",cuil.nrodoc,estructura.estrdext osocial,estructura2.estrcodext regjub"
StrSql = StrSql & ",cuil.nrodoc,estructura.estrdext osocial,cajjub.ticnro regjub,estructura2.estrcodext codextcaja"
'           ticnro = 1 --> Reparto
'                  <>1 --> Capitalizacion
StrSql = StrSql & ",estructura2.estrdabr afjp "

StrSql = StrSql & " FROM tercero "
StrSql = StrSql & " left join ter_doc cuil on cuil.ternro = tercero.ternro"
StrSql = StrSql & " AND cuil.tidnro = 10"

StrSql = StrSql & " left join cabdom on tercero.ternro = cabdom.ternro and cabdom.tidonro=2"
StrSql = StrSql & " left join telefono on telefono.domnro = cabdom.domnro AND telefono.teldefault = -1"
StrSql = StrSql & " left join his_estructura on his_estructura.ternro=tercero.ternro AND his_estructura.tenro=17 "
     'StrSql = StrSql & " and his_estructura.tenro=17 and his_estructura.htethasta is null"
     ' Leti A - 09-2010 - Busca la Obra Social elegida a la fecha del accidente
StrSql = StrSql & "  AND  his_estructura.htetdesde <=" & ConvFecha(Aux_Acc_Fecha)
StrSql = StrSql & "  AND (his_estructura.htethasta is NULL OR his_estructura.htethasta >= " & ConvFecha(Aux_Acc_Fecha) & " )"

StrSql = StrSql & " left join estructura on estructura.estrnro=his_estructura.estrnro"
StrSql = StrSql & " left join estr_cod on estr_cod.estrnro=estructura.estrnro and estr_cod.tcodnro=1"

StrSql = StrSql & " left join his_estructura his_estructura2 on his_estructura2.ternro=tercero.ternro"
StrSql = StrSql & " and his_estructura2.tenro=15 and his_estructura2.htethasta is null"
StrSql = StrSql & " left join estructura estructura2 on estructura2.estrnro=his_estructura2.estrnro"
'FGZ - 24/07/2007 - esto es para chequear el tipo de caja de jubilacion (No se debe chequear por codext)
StrSql = StrSql & " LEFT JOIN cajjub ON his_estructura2.estrnro = cajjub.estrnro "
StrSql = StrSql & " LEFT JOIN empleado on tercero.ternro = empleado.ternro "
StrSql = StrSql & " Where tercero.ternro = " & rs_Empleado!ternro
OpenRecordset StrSql, rs_Tercero

If Not rs_Tercero.EOF Then
    Aux_Emp_Apeynom = rs_Tercero!terape + " " + rs_Tercero!ternom
    Aux_Emp_Cuil = IIf(IsNull(rs_Tercero!nrodoc), "", rs_Tercero!nrodoc)
    Aux_Emp_Tel = IIf(IsNull(rs_Tercero!telnro), "", rs_Tercero!telnro)
    Aux_Emp_fnac = IIf(IsNull(rs_Tercero!terfecnac), "", rs_Tercero!terfecnac)
    Aux_emp_OS = IIf(IsNull(rs_Tercero!osocial), "", rs_Tercero!osocial)
    'Aux_Emp_RegJub = IIf(IsNull(rs_Tercero!regjub), "", rs_Tercero!regjub)
    'FGZ - 24/07/2007 - Le cambié esta categorizacion ----
    If EsNulo(rs_Tercero!regjub) Then
        Aux_Emp_RegJub = UCase(rs_Tercero!codextcaja)
    Else
        If rs_Tercero!regjub = 1 Then
            Aux_Emp_RegJub = "REP"
        Else
            Aux_Emp_RegJub = "CAP"
        End If
    End If
    'Aux_Emp_RegJub = rs_Tercero!regjub
    'FGZ - 24/07/2007 - Le cambié esta categorizacion ----
    Aux_emp_AFJP = IIf(IsNull(rs_Tercero!afjp), "", rs_Tercero!afjp)
    Aux_emp_CodOS = IIf(IsNull(rs_Tercero!nrocod), "", rs_Tercero!nrocod)
    
Else
    Flog.writeline "El Registro de Tercero no está  disponible"
    Aux_Emp_Apeynom = ""
    Aux_Emp_Cuil = ""
    'UNDO, LEAVE.
End If


' busco el tipo de contracto que tiene el empleado a la fecha del accidente
contratoActual Aux_Acc_Fecha, rs_Empleado!ternro, Aux_Emp_Contrato


StrSql = "SELECT altfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE empleado = " & rs_Empleado!ternro
StrSql = StrSql & " AND real = -1 "
'StrSql = StrSql & " AND not altfec is null "
'StrSql = StrSql & " AND altfec <= " & ConvFecha(rs_Accidente!accfecha)
StrSql = StrSql & " ORDER BY altfec "
OpenRecordset StrSql, rs_Fases
If rs_Fases.EOF Then
    rs_Fases.Close
    StrSql = "SELECT altfec "
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " WHERE empleado = " & rs_Empleado!ternro
    StrSql = StrSql & " AND altfec is null "
    StrSql = StrSql & " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    If rs_Fases.EOF Then
        rs_Fases.Close
        StrSql = "SELECT altfec "
        StrSql = StrSql & " FROM fases "
        StrSql = StrSql & " WHERE empleado = " & rs_Empleado!ternro
        StrSql = StrSql & " ORDER BY altfec "
        OpenRecordset StrSql, rs_Fases
        If rs_Fases.EOF Then
            Aux_Emp_fingreso = ""
        Else
            Aux_Emp_fingreso = rs_Fases!altfec
        End If
        rs_Fases.Close
    Else
        Aux_Emp_fingreso = rs_Fases!altfec
        rs_Fases.Close
    End If
Else
    Aux_Emp_fingreso = rs_Fases!altfec
    rs_Fases.Close
End If

'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DATOS DEL ACCIDENTE *****************************************
Flog.writeline "DATOS DEL ACCIDENTE"
'**************************************************************************************************

StrSql = "SELECT * FROM soaccid_visita WHERE accnro = " & rs_Accidente!accnro
OpenRecordset StrSql, rs_Accid_Visita
If Not rs_Accid_Visita.EOF Then
    StrSql = "SELECT * FROM sovisitamedica WHERE vismednro = " & rs_Accid_Visita!visitamed
    OpenRecordset StrSql, rs_VisitaMedica
    If Not rs_VisitaMedica.EOF Then
      Aux_Acc_FecAlta = rs_VisitaMedica!vismedfecha
    Else
        Aux_Acc_FecAlta = "01/01/1800"
    End If
Else
    Aux_Acc_FecAlta = "01/01/1800"
End If
cantdias = 0

Aux_Acc_FecReintHasta = "01/01/1800"
Aux_Acc_FecReintDesde = "01/01/1800"

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
objconnProgreso.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DETALLE DE LAS REMUNERACIONES *******************************
Flog.writeline "DETALLE DE LAS REMUNERACIONES"
'**************************************************************************************************

Dim I     As Integer
Dim Mes   As Integer
Dim Anio  As Integer
Dim Total As Single
Dim Cantidad As Integer
Dim v_fecha As Long
Dim Anio_hasta As Integer
Dim Mes_hasta As Integer


    I = 12
    Total = 0
    Cantidad = 0
    
    'FGZ - 24/07/2007 - cambié el orden ----
    'Anio = IIf(Month(rs_Accidente!accfecha) = 1, Year(rs_Accidente!accfecha) - 1, Year(rs_Accidente!accfecha))
    'mes = IIf(Month(rs_Accidente!accfecha) = 1, 12, Month(rs_Accidente!accfecha) - 1)
    Anio = Year(DateAdd("m", -12, rs_Accidente!accfecha))
    Mes = Month(DateAdd("m", -12, rs_Accidente!accfecha))
      
    
    Aux_PorcRed = 0
    Aux_PorcJub = 0
    Aux_PorcAsigFam = 0
    Aux_PorcFondoNac = 0
    Aux_PorcINSSJP = 0
    Aux_PorcOS = 0
    Aux_PorcTotalC = 0
    Aux_PorcANSSal = 0
    
    
    Aux_agu_mes_1 = 0
    Aux_agu_anio_1 = 0
    Aux_agu_importe_1 = 0
    Aux_agu_mes_2 = 0
    Aux_agu_anio_2 = 0
    Aux_agu_importe_2 = 0
    'FGZ - 24/07/2007 - le agregué esto -----
    aux_tot_importe = 0
    AsigneA1 = False
    'FGZ - 24/07/2007 - le agregué esto -----
    Flog.writeline
    Flog.writeline "Buscando detalles de liquidación de los últimos 12 meses ...."
    Flog.writeline
    Do While I >= 1
                  
        StrSql = "SELECT * FROM periodo "
        StrSql = StrSql & " WHERE pliqanio = " & Anio
        StrSql = StrSql & " AND pliqmes = " & Mes
        StrSql = StrSql & " ORDER BY pliqanio, pliqmes"
        OpenRecordset StrSql, rs_Periodo
    
        If Not rs_Periodo.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "-- Año: " & Anio
            Flog.writeline Espacios(Tabulador * 1) & "-- Mes: " & Mes
            
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
             If Not IsNull(ConPorcRed) Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConPorcRed
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                OpenRecordset StrSql, rs_Detliq
                Do While Not rs_Detliq.EOF
                   Aux_PorcRed = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
              
             'busco el porcentaje de Jubilacion
             If Not IsNull(ConJub) Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConJub
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcJub = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
             
             'busco el porcentaje de Asignaciones Familiares
             If Not IsNull(ConAsigFam) Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConAsigFam
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcAsigFam = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
             
             'busco el porcentaje de Fondo Nacional de Desempleo
             If Not IsNull(ConFondoNac) Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConFondoNac
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcFondoNac = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
             
             'busco el porcentaje de INSSJP
             If Not IsNull(ConINSSJP) Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConINSSJP
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcINSSJP = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
              
             'busco el porcentaje de ANSSAL
             If Not IsNull(ConANSSal) Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConANSSal
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcANSSal = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
              
              'busco el porcentaje de Tot_Contribuciones
             If Not IsNull(ConTC) Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " WHERE detliq.concnro = " & ConTC
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                OpenRecordset StrSql, rs_Detliq
                
                Do While Not rs_Detliq.EOF
                   Aux_PorcTotalC = rs_Detliq!dlicant
                   
                   rs_Detliq.MoveNext
                Loop
              End If
             
            'busco el porcentaje de Obra Social
            If Not IsNull(ConOS) Then
                If tipo10 = "CO" Then
                    StrSql = "SELECT * FROM proceso "
                    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    'StrSql = StrSql & " WHERE detliq.concnro = " & ConOS
                    StrSql = StrSql & " WHERE detliq.concnro IN (" & ConOS & ")"
                    StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                    StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
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
                    'StrSql = StrSql & " WHERE acu_liq.acunro = " & ConOS
                    StrSql = StrSql & " WHERE acu_liq.acunro IN ( " & ConOS & ")"
                    StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                    StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                    OpenRecordset StrSql, rs_Acu_Liq
                    
                    Do While Not rs_Acu_Liq.EOF
                       Aux_PorcOS = rs_Acu_Liq!alcant
                       
                       rs_Acu_Liq.MoveNext
                    Loop
                End If
            End If
            
            'Buscar SAC Proporcional -------------------------------
            'If Not IsNull(ConAgu2) Then
            If Not IsNull(ConAgu2Str) Then
                StrSql = "SELECT * FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                'StrSql = StrSql & " WHERE detliq.concnro = " & ConAgu2
                StrSql = StrSql & " WHERE detliq.concnro in (" & ConAgu2Str & ")"
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                OpenRecordset StrSql, rs_Detliq
                Do While Not rs_Detliq.EOF
                        Aux_sac_prop_mes = Month(rs_Detliq!profecfin)
                        Aux_sac_prop_anio = Year(rs_Detliq!profecfin)
                        Sac_Proporcional = Sac_Proporcional + rs_Detliq!dlimonto
                    rs_Detliq.MoveNext
                Loop
            End If
            
            'busco el aguinaldo 1
            If Mes = 6 Or Mes = 12 Then
                'If Not IsNull(ConAgu1) Then
                If Not IsNull(ConAguStr) Then
                    StrSql = "SELECT * FROM proceso "
                    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    'StrSql = StrSql & " WHERE detliq.concnro = " & ConAgu1
                    StrSql = StrSql & " WHERE detliq.concnro IN (" & ConAguStr & ")"
                    StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
                    StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
                    OpenRecordset StrSql, rs_Detliq
                    If rs_Detliq.EOF Then
                        If Sac_Proporcional <> 0 Then
                            If Not AsigneA1 Then
                                Aux_agu_mes_1 = Aux_sac_prop_mes
                                Aux_agu_anio_1 = Aux_sac_prop_anio
                                Aux_agu_importe_1 = Sac_Proporcional
                                aux_tot_importe = aux_tot_importe + Aux_agu_importe_1
                            Else
                                Aux_agu_mes_2 = Aux_sac_prop_mes
                                Aux_agu_anio_2 = Aux_sac_prop_anio
                                Aux_agu_importe_2 = Sac_Proporcional
                                aux_tot_importe = aux_tot_importe + Aux_agu_importe_2
                            End If
                        End If
                    End If
                    Do While Not rs_Detliq.EOF
                        If Not AsigneA1 Then
                            Aux_agu_mes_1 = Month(rs_Detliq!profecfin)
                            Aux_agu_anio_1 = Year(rs_Detliq!profecfin)
                            'Aux_agu_importe_1 = rs_Detliq!dlimonto
                            Aux_agu_importe_1 = rs_Detliq!dlimonto + Sac_Proporcional
                            aux_tot_importe = aux_tot_importe + Aux_agu_importe_1
                        Else
                            Aux_agu_mes_2 = Month(rs_Detliq!profecfin)
                            Aux_agu_anio_2 = Year(rs_Detliq!profecfin)
                            'Aux_agu_importe_2 = rs_Detliq!dlimonto
                            Aux_agu_importe_2 = rs_Detliq!dlimonto + Sac_Proporcional
                            aux_tot_importe = aux_tot_importe + Aux_agu_importe_2
                        End If
                        rs_Detliq.MoveNext
                    Loop
                    AsigneA1 = True
                  End If
                  Sac_Proporcional = 0
                  Aux_sac_prop_mes = 0
                  Aux_sac_prop_anio = 0
            End If
            'FGZ - 24/07/2007 - le saqué esta parte dado que el concepto de aguinaldo es el mismo
            '             'busco el aguinaldo 2
            '             If Not IsNull(ConAgu2) And Aux_agu_importe_2 = 0 Then
            '                StrSql = "SELECT * FROM proceso "
            '                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
            '                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
            '                StrSql = StrSql & " WHERE detliq.concnro = " & ConAgu2
            '                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleado!ternro
            '                StrSql = StrSql & " AND proceso.pliqnro =" & rs_Periodo!pliqnro
            '                OpenRecordset StrSql, rs_Detliq
            '
            '                Do While Not rs_Detliq.EOF
            '                    'FGZ - 28/06/2007 - saco los meses de las fechas del proceso
            '                    'Aux_agu_mes_2 = Month(rs_Detliq!fechaimp)
            '                    'Aux_agu_anio_2 = Year(rs_Detliq!fechaimp)
            '                    Aux_agu_mes_2 = Month(rs_Detliq!profecfin)
            '                    Aux_agu_anio_2 = Year(rs_Detliq!profecfin)
            '                    'FGZ - 28/06/2007 - saco los meses de las fechas del proceso
            '                   Aux_agu_importe_2 = rs_Detliq!dlimonto
            '                   'FGZ - 24/07/2007 - le agregué esto -----
            '                   aux_tot_importe = aux_tot_importe + Aux_agu_importe_2
            '                   'FGZ - 24/07/2007 - le agregué esto -----
            '                   rs_Detliq.MoveNext
            '                Loop
            '               End If
            'FGZ - 24/07/2007 - le saqué esta parte dado que el concepto de aguinaldo es el mismo
               
        End If 'rs_periodo.eof
        
        Aux_Det_Importe(I) = Total
        Aux_Det_Dias(I) = IIf(ExisteAcuDias, Cantidad, cantdias_mes(Mes))
        Aux_Det_Anio(I) = CStr(Anio)
        Aux_Det_Mes(I) = Mes
        'FGZ - 24/07/2007 - le agregué esto
        aux_tot_importe = aux_tot_importe + Total
        'FGZ - 24/07/2007 - le agregué esto
        I = I - 1
        Total = 0
        Cantidad = 0
        Anio = IIf(Mes = 12, Anio + 1, Anio)
        Mes = IIf(Mes = 12, 1, Mes + 1)
      
        'FGZ - 26/07/2007
        'Aux_Fecha_Hasta = rs_Periodo!pliqhasta  ' Leti A (4-10-2010) se comenta -no se usa
   Loop


'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords

'**************************************************************************************************
'************************************ DETALLE DE LAS PRESTACIONES DINERARIAS **********************
Flog.writeline "DETALLE DE LAS PRESTACIONES DINERARIAS"
'**************************************************************************************************
Dim fecdesde As Date
Dim fechasta As Date
Dim cantMesesART As Integer
Dim cantDiasART As Integer

'FGZ - 24/07/2007  - deshabilito esto
'aux_tot_importe = 0
'FGZ - 24/07/2007  - deshabilito esto
aux_tot_dias = 0


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
objconnProgreso.Execute StrSql, , adExecuteNoRecords
       

Dim Monto  As Single
Dim aux_BDC

aux_BDC = 0

StrSql = "SELECT * FROM Rep_prestdine "
StrSql = StrSql & " WHERE ternro = " & EmpTer
StrSql = StrSql & " AND bpronro = " & bpronro
OpenRecordset StrSql, rs_Rep_prestdine

If rs_Rep_prestdine.EOF Then
    
    StrSql = "INSERT INTO Rep_prestdine (bpronro,empresa,iduser,fecha,hora,"
    StrSql = StrSql & "ternro,empdor_razsoc,empdor_cuit,"
    StrSql = StrSql & "emp_apeynom,emp_cuil,emp_tel,emp_fechanac,"
    StrSql = StrSql & "emp_fechaingreso,emp_osocial,emp_osocialcodigo,"
    StrSql = StrSql & "emp_regjubilacion,emp_afjp, emp_contrato, "
    'StrSql = StrSql & "acc_nro, acc_fecha, acc_fecalta, acc_fecreintdesde, acc_fecreinthasta, acc_diasbaja, acc_diasart,acc_fecingreso,acc_nrosiniestro, "
    StrSql = StrSql & "acc_nro, acc_fecha, acc_fecalta, acc_fecreintdesde, acc_fecreinthasta, acc_diasbaja, acc_diasart,acc_nrosiniestro, "
    StrSql = StrSql & "det_anio_12, det_anio_11, det_anio_10, det_anio_9, det_anio_8,det_anio_7, det_anio_6, det_anio_5, det_anio_4, det_anio_3, det_anio_2, det_anio_1, "
    
    StrSql = StrSql & "det_mes_12, det_mes_11, det_mes_10, det_mes_9, det_mes_8,det_mes_7, det_mes_6, det_mes_5, det_mes_4, det_mes_3, det_mes_2, det_mes_1, "
    StrSql = StrSql & "det_importe_12, det_importe_11, det_importe_10, det_importe_9, det_importe_8,det_importe_7, det_importe_6, det_importe_5, det_importe_4, det_importe_3, det_importe_2, det_importe_1, "
    StrSql = StrSql & "det_dias_12, det_dias_11, det_dias_10, det_dias_9, det_dias_8,det_dias_7, det_dias_6, det_dias_5, det_dias_4, det_dias_3, det_dias_2, det_dias_1, "
        
    StrSql = StrSql & "agu_mes_1,agu_anio_1,agu_importe_1,agu_mes_2,agu_anio_2,agu_importe_2,"

    StrSql = StrSql & "tot_importe,tot_dias,bdc_ilt,"
    StrSql = StrSql & "porcred, porcjub, porcasigfam, porcfondonac, porcinssjp, porcos, "
    StrSql = StrSql & "porcanssal,porctotc"
      
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & bpronro & ","
    StrSql = StrSql & EMPRESA & ","
    StrSql = StrSql & "'" & IdUser & "',"
    StrSql = StrSql & ConvFecha(Fecha) & ","
    StrSql = StrSql & "'" & Format(Hora, "hh:mm:ss") & "',"
    
    StrSql = StrSql & EmpTer & ","
    StrSql = StrSql & "'" & Aux_Empdor_RazSoc & "',"
    StrSql = StrSql & "'" & Aux_Empdor_Cuit & "',"
    
    StrSql = StrSql & "'" & Aux_Emp_Apeynom & "',"
    StrSql = StrSql & "'" & Aux_Emp_Cuil & "',"
    
    StrSql = StrSql & "'" & Aux_Emp_Tel & "',"
    StrSql = StrSql & ConvFecha(Aux_Emp_fnac) & ","
    
    StrSql = StrSql & ConvFecha(Aux_Emp_fingreso) & ","
    
    StrSql = StrSql & "'" & Aux_emp_OS & "',"
    StrSql = StrSql & "'" & Aux_emp_CodOS & "',"
    StrSql = StrSql & "'" & Aux_Emp_RegJub & "',"
    StrSql = StrSql & "'" & Aux_emp_AFJP & "',"
    StrSql = StrSql & "'" & Aux_Emp_Contrato & "',"
    
         
    StrSql = StrSql & "'" & Aux_Acc_Nro & "',"
    StrSql = StrSql & ConvFecha(Aux_Acc_Fecha) & ","
    StrSql = StrSql & ConvFecha(Aux_Acc_FecAlta) & ","
    StrSql = StrSql & ConvFecha(Aux_Acc_FecReintDesde) & ","
    StrSql = StrSql & ConvFecha(Aux_Acc_FecReintHasta) & ","
    StrSql = StrSql & Aux_Acc_DiasBaja & ","
    StrSql = StrSql & Aux_Acc_DiasART & ","
    'StrSql = StrSql & ConvFecha(Aux_Acc_FecIngreso) & ","
    StrSql = StrSql & Aux_Acc_nrosiniestro & ","
    
    For I = 1 To 12
        If IsNull(Aux_Det_Anio(I)) Then
            StrSql = StrSql & "'',"
        Else
            StrSql = StrSql & "'" & Aux_Det_Anio(I) & "',"
        End If
    Next I
    For I = 1 To 12
        If IsNull(Aux_Det_Mes(I)) Then
            StrSql = StrSql & "'',"
        Else
            StrSql = StrSql & "'" & Aux_Det_Mes(I) & "',"
        End If
    Next I
    For I = 1 To 12
        If IsNull(Aux_Det_Importe(I)) Then
            StrSql = StrSql & "null,"
        Else
            StrSql = StrSql & Aux_Det_Importe(I) & ","
            'Aux_PorcTotalC = Aux_PorcTotalC + Aux_Det_Importe(I)
        End If
    Next I
    For I = 1 To 12
        If IsNull(Aux_Det_Dias(I)) Then
            StrSql = StrSql & "null,"
        Else
            StrSql = StrSql & Aux_Det_Dias(I) & ","
            aux_tot_dias = aux_tot_dias + IIf(Aux_Det_Importe(I) > 0, Aux_Det_Dias(I), 0)
        End If
        
    Next I
       
    'FGZ - 25/07/2007 - hago el recalculo de los dias corridos
    'Se deben calcular como la cantidad de dias corridos durante los ultimos 12 meses anteriores al accidente
    Aux_Fecha_Desde = DateAdd("m", -12, rs_Accidente!accfecha)
    If Aux_Emp_fingreso > Aux_Fecha_Desde Then
        Aux_Fecha_Desde = Aux_Emp_fingreso
    End If
    Aux_Fecha_Hasta = rs_Accidente!accfecha
    aux_tot_dias = DateDiff("d", Aux_Fecha_Desde, rs_Accidente!accfecha)
    'FGZ - 25/07/2007 - hago el recalculo de los dias corridos
    
    '--------------------------------------------------------------
    aux_tot_dias = 0
    StrSql2 = "SELECT * FROM fases WHERE empleado = " & rs_Accidente!Empleado & _
             " AND real = -1 " & _
             " AND not altfec is null " & _
             " AND altfec <= " & ConvFecha(rs_Accidente!accfecha)
    OpenRecordset StrSql2, rs_Fases
    Do While Not rs_Fases.EOF
        fecAlta = rs_Fases!altfec
        If Not EsNulo(rs_Fases!bajfec) Then
            fecBaja = rs_Fases!bajfec
        Else
            fecBaja = rs_Accidente!accfecha
        End If
    
        aux_tot_dias = aux_tot_dias + CantidadDeDias(Aux_Fecha_Desde, Aux_Fecha_Hasta, fecAlta, fecBaja)
    
        rs_Fases.MoveNext
    Loop
    If aux_tot_dias > 365 Then
        aux_tot_dias = 365
    End If
    '--------------------------------------------------------------
    
    
    If aux_tot_dias > 0 Then
        'aux_BDC = (Aux_PorcTotalC + Aux_agu_importe_1 + Aux_agu_importe_2) / aux_tot_dias
        'FGZ - 24/07/2007 - Lo reemplacé por esto
        aux_BDC = Round((aux_tot_importe) / aux_tot_dias, 2)
    End If
    
    StrSql = StrSql & Aux_agu_mes_1 & ","
    StrSql = StrSql & Aux_agu_anio_1 & ","
    StrSql = StrSql & Aux_agu_importe_1 & ","
    
    StrSql = StrSql & Aux_agu_mes_2 & ","
    StrSql = StrSql & Aux_agu_anio_2 & ","
    StrSql = StrSql & Aux_agu_importe_2 & ","
    
    StrSql = StrSql & aux_tot_importe & ","
    StrSql = StrSql & aux_tot_dias & ","
    StrSql = StrSql & aux_BDC & ","
    
    StrSql = StrSql & Aux_PorcRed & ","
    StrSql = StrSql & Aux_PorcJub & ","
    StrSql = StrSql & Aux_PorcAsigFam & ","
    StrSql = StrSql & Aux_PorcFondoNac & ","
    StrSql = StrSql & Aux_PorcINSSJP & ","
    StrSql = StrSql & Aux_PorcOS & ","
    
    StrSql = StrSql & Aux_PorcANSSal & ","
    StrSql = StrSql & Aux_PorcTotalC
           
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
End If
            
'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords
            

'Fin de la transaccion
MyCommitTrans

If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Accidente.State = adStateOpen Then rs_Accidente.Close
If rs_acumulador.State = adStateOpen Then rs_acumulador.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Accid_Visita.State = adStateOpen Then rs_Accid_Visita.Close
If rs_VisitaMedica.State = adStateOpen Then rs_VisitaMedica.Close
If rs_Lic_Accid.State = adStateOpen Then rs_Lic_Accid.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_CabLiq.State = adStateOpen Then rs_CabLiq.Close
If rs_Rep_prestdine.State = adStateOpen Then rs_Rep_prestdine.Close
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_OS.State = adStateOpen Then rs_OS.Close

Set rs_Empleado = Nothing
Set rs_Accidente = Nothing
Set rs_acumulador = Nothing
Set rs_Confrep = Nothing
Set rs_Reporte = Nothing
Set rs_Empresa = Nothing
Set rs_Tercero = Nothing

Set rs_Accid_Visita = Nothing
Set rs_VisitaMedica = Nothing
Set rs_Lic_Accid = Nothing
Set rs_Periodo = Nothing
Set rs_Acu_Liq = Nothing
Set rs_Detliq = Nothing
Set rs_CabLiq = Nothing
Set rs_Rep_prestdine = Nothing
Set rs_Fases = Nothing
Set rs_OS = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "==================================================================================="
    Flog.writeline Err.Description
    Flog.writeline "Ultimo SQL Ejecutado:"
    Flog.writeline StrSql
    Flog.writeline "==================================================================================="
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer

Dim EmpTer As Long
Dim NroAcc As Long
Dim EMPRESA As Long

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
        
    End If
End If

Call ArtLpd02(bpronro, EmpTer, NroAcc)

End Sub


Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
   
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

On Error GoTo CE

    StrSql = "SELECT * FROM confrep WHERE repnro = 177 AND confnrocol = " & NroCol
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    If Not rs_Confrep.EOF Then
        Flog.Write "Llamo Columna " & NroCol & " Tipo: " & rs_Confrep!conftipo & " = " & rs_Confrep!confval & "("
        Select Case rs_Confrep!conftipo
            Case "CO":
                If IsNull(rs_Confrep!confval2) Then
                    StrSql = " SELECT * FROM concepto WHERE conccod = " & rs_Confrep!confval
                Else
                    StrSql = " SELECT * FROM concepto WHERE conccod = '" & rs_Confrep!confval2 & "'"
                End If
                OpenRecordset StrSql, rs_Concepto
                If Not rs_Concepto.EOF Then
                    nroconc = rs_Concepto!ConcNro
                    Flog.Write nroconc
                    TipoNro = "CO"
                End If
            Case "AC":
                If IsNull(rs_Confrep!confval2) Then
                    StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval
                Else
                    StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval2
                End If
                OpenRecordset StrSql, rs_acumulador
                If Not rs_acumulador.EOF Then
                    nroconc = rs_acumulador!acuNro
                    Flog.Write nroconc
                    TipoNro = "AC"
                End If
         End Select
         Flog.writeline ")"
    End If

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "==================================================================================="
    Flog.writeline Err.Description
    Flog.writeline "Ultimo SQL Ejecutado:"
    Flog.writeline StrSql
    Flog.writeline "==================================================================================="
End Sub

Public Sub ColumnaArr(ByVal NroCol As Integer, ByRef nroconc As String, ByRef TipoNro As String)
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_acumulador As New ADODB.Recordset

On Error GoTo CE
    
    nroconc = 0
    
    StrSql = "SELECT * FROM confrep WHERE repnro = 177 AND confnrocol = " & NroCol
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    If Not rs_Confrep.EOF Then
        Flog.Write "Llamo Columna " & NroCol & " Tipo: " & rs_Confrep!conftipo & " = " & rs_Confrep!confval & "("
        Do While Not rs_Confrep.EOF
            Select Case rs_Confrep!conftipo
                Case "CO":
                    If IsNull(rs_Confrep!confval2) Then
                        StrSql = " SELECT * FROM concepto WHERE conccod = " & rs_Confrep!confval
                    Else
                        StrSql = " SELECT * FROM concepto WHERE conccod = '" & rs_Confrep!confval2 & "'"
                    End If
                    OpenRecordset StrSql, rs_Concepto
                    If Not rs_Concepto.EOF Then
                        nroconc = nroconc & "," & rs_Concepto!ConcNro
                        Flog.Write nroconc
                        TipoNro = "CO"
                    End If
                Case "AC":
                    If IsNull(rs_Confrep!confval2) Then
                        StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval
                    Else
                        StrSql = " SELECT * FROM acumulador WHERE acunro = " & rs_Confrep!confval2
                    End If
                    OpenRecordset StrSql, rs_acumulador
                    If Not rs_acumulador.EOF Then
                        nroconc = nroconc & "," & rs_acumulador!acuNro
                        Flog.Write nroconc
                        TipoNro = "AC"
                    End If
            End Select
            rs_Confrep.MoveNext
        Loop
        Flog.writeline ")"
    End If
    
Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "==================================================================================="
    Flog.writeline Err.Description
    Flog.writeline "Ultimo SQL Ejecutado:"
    Flog.writeline StrSql
    Flog.writeline "==================================================================================="
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
Public Function cantdias_mes(ByVal nromes)
Select Case nromes
    Case 1, 3, 5, 7, 8, 10, 12:
        cantdias_mes = 31
    Case 4, 6, 9, 11:
        cantdias_mes = 30
    Case 2:
        cantdias_mes = 28
End Select

End Function



' ________________________________________________________________________________
Sub contratoActual(ByVal Aux_Acc_Fecha, ternro, ByRef Aux_Emp_Contrato As String)

Dim rs As New ADODB.Recordset

On Error GoTo Cont

    ' Buscar el CONTRATO a la fecha del accidente
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN tipocont ON his_estructura.estrnro = tipocont.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro
    StrSql = StrSql & "     AND tenro = 18 "
    StrSql = StrSql & "     AND htetdesde <= " & ConvFecha(Aux_Acc_Fecha)
    StrSql = StrSql & "     AND (htethasta >= " & ConvFecha(Aux_Acc_Fecha) & " or htethasta is null)"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Aux_Emp_Contrato = rs!tcdabr
    Else
        Aux_Emp_Contrato = ""
    End If
    rs.Close
    
    Exit Sub
    
Cont:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "==================================================================================="
    Flog.writeline Err.Description
    Flog.writeline "Ultimo SQL Ejecutado:"
    Flog.writeline StrSql
    Flog.writeline "==================================================================================="
    
End Sub
