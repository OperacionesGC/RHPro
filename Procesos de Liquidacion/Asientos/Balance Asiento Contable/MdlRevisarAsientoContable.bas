Attribute VB_Name = "MdlAsientoRevisarBalance"
Option Explicit

Private Type TRegEstructura
    TE As Long
    Estructura As Long
    Porcentaje As Single
End Type

Global Inx             As Integer
Global Inxfin          As Integer
Global LI_1 As Integer
Global LI_2 As Integer
Global LI_3 As Integer

Global Inx_1 As Integer
Global Inx_2 As Integer
Global Inx_3 As Integer

'FGZ - 01/04/2005
Global vec_testr1(50)  As TRegEstructura
Global vec_testr2(50)  As TRegEstructura
Global vec_testr3(50)  As TRegEstructura
'FGZ - 01/04/2005

'Global vec_jor(50) As Single

Global Descripcion As String
Global Cantidad As Single
Global CatidadVueltas As Long

Global rs_Proc_Vol As New ADODB.Recordset
Global rs_Mod_Linea As New ADODB.Recordset
Global rs_Empleado As New ADODB.Recordset
Global rs_Mod_Asiento As New ADODB.Recordset

Global BUF_mod_linea As New ADODB.Recordset
Global BUF_temp As New ADODB.Recordset

Global CantidadEmpleados As Long
Global PrimeraVez As Boolean


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Asientos Contables.
' Autor      : FGZ
' Fecha      : 16/01/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    
    Nombre_Arch = PathFLog & "Balance_Asiento_Contable" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline "PID = " & PID
    
    TiempoInicialProceso = GetTickCount
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ",bprctiempo = 0 WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 104 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Revisar_Balance(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    objConn.Close
    objconnProgreso.Close
    Flog.Close

End Sub


Public Sub Revisar_Balance(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Programa que se revisa si el asiento balancea
'
' Autor      : FGZ
' Fecha      : 11/08/2005
' Ult. Mod   :
' --------------------------------------------------------------------------------------------
Dim Mantener_Liq As Boolean
Dim Analisis_Detallado As Boolean
Dim Todos As Boolean

Dim pos1 As Integer
Dim pos2 As Integer

Dim NroVol   As Long
Dim Fecha    As Date
Dim Total    As Long
Dim Balancea As Boolean
Dim Longitud As Integer

Dim TotalDebe    As Double
Dim TotalHaber    As Double
Dim Aux_TotalHaber As String
Dim Aux_TotalDebe    As String
Dim Aux_Monto As String
Dim Aux_Descripcion As String
Dim Aux_Cuenta As String

Dim Aux_TipoOrigen As String
Dim Aux_Origen As String

Dim rs_Proc As New ADODB.Recordset
Dim rs_Detalles As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset

' El formato del mismo es (pronro.mantener Liq Ant.Guardar Nov.Analisis Det.Todos)
' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        NroVol = CLng(Mid(Parametros, pos1, pos2))
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        HACE_TRAZA = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
    End If
End If

'Busco todos los legajos pertencientes al proceso de volcado
StrSql = "SELECT * FROM proc_vol_pl"
StrSql = StrSql & " INNER JOIN proc_vol_emp ON proc_vol_emp.pronro  = proc_vol_pl.pronro"
StrSql = StrSql & " WHERE proc_vol_pl.vol_cod =" & NroVol
StrSql = StrSql & " AND proc_vol_emp.vol_cod = " & NroVol
StrSql = StrSql & " ORDER BY proc_vol_emp.ternro"
OpenRecordset StrSql, rs_Proc
 
CantidadEmpleados = rs_Proc.RecordCount
Flog.Writeline "Cantidad de Empleados = " & CantidadEmpleados
If CantidadEmpleados = 0 Then
    CantidadEmpleados = 1
End If
IncPorc = 95 / CantidadEmpleados
Balancea = True
Do While Not rs_Proc.EOF ' (1)
    'Inicializo los acumuladores
    TotalDebe = 0
    TotalHaber = 0
    
    StrSql = "SELECT * FROM empleado where empleado.ternro = " & rs_Proc!ternro
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.Writeline "No se encontro el legajo"
        Exit Sub
    Else
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "----------------- Legajo " & rs_Empleado!empleg
        Longitud = 60
        Aux_Descripcion = "Descripcion"
        If Len(Aux_Descripcion) < Longitud Then
            Aux_Descripcion = String(Longitud - Len(Aux_Descripcion), " ") & Aux_Descripcion
        End If
        
        Longitud = 50
        Aux_Cuenta = "Cuenta"
        If Len(Aux_Cuenta) < Longitud Then
            Aux_Cuenta = String(Longitud - Len(Aux_Cuenta), " ") & Aux_Cuenta
        End If
        
        Longitud = 10
        Aux_Monto = "Monto"
        If Len(Aux_Monto) < Longitud Then
            Aux_Monto = String(Longitud - Len(Aux_Monto), " ") & Aux_Monto
        End If
        Flog.Writeline Espacios(Tabulador * 1) & Aux_Descripcion & Aux_Cuenta & Aux_Monto
    End If
    
    'Busco todas las lineas de asiento
    StrSql = "SELECT * FROM detalle_asi "
    StrSql = StrSql & " WHERE ternro = " & rs_Empleado!ternro
    StrSql = StrSql & " AND vol_cod = " & NroVol
    StrSql = StrSql & " ORDER BY masinro, cuenta"
    OpenRecordset StrSql, rs_Detalles
    Do While Not rs_Detalles.EOF
        'Muestro los datos
        Longitud = 60
        Aux_Descripcion = rs_Detalles!dldescripcion
        If Len(Aux_Descripcion) < Longitud Then
            Aux_Descripcion = String(Longitud - Len(Aux_Descripcion), " ") & Aux_Descripcion
        End If
        Longitud = 50
        Aux_Cuenta = rs_Detalles!Cuenta
        If Len(Aux_Cuenta) < Longitud Then
            Aux_Cuenta = String(Longitud - Len(Aux_Cuenta), " ") & Aux_Cuenta
        End If
        Longitud = 10
        Aux_Monto = CStr(Format(rs_Detalles!dlmonto, "####0.00"))
        If Len(Aux_Monto) < Longitud Then
            Aux_Monto = String(Longitud - Len(Aux_Monto), " ") & Aux_Monto
        End If
        
        'Determinar el origen de los datos
        If Not EsNulo(rs_Detalles!tipoOrigen) Then
            If rs_Detalles!tipoOrigen = 1 Then
                Aux_TipoOrigen = "Concepto "
                
                StrSql = "SELECT * FROM concepto "
                StrSql = StrSql & " WHERE concnro = " & rs_Detalles!Origen
                If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
                OpenRecordset StrSql, rs_Concepto
                If rs_Concepto.EOF Then
                    Aux_Origen = ""
                Else
                    Aux_Origen = CStr(rs_Concepto!Conccod)
                End If
            Else
                Aux_TipoOrigen = "Acumulador"
                Aux_Origen = CStr(rs_Detalles!Origen)
            End If
        Else
            Aux_TipoOrigen = "Desconocido"
            Aux_Origen = ""
        End If
        Flog.Writeline Espacios(Tabulador * 1) & Aux_Descripcion & Aux_Cuenta & Aux_Monto & Aux_TipoOrigen & Aux_Origen
        
        'Acumulo
        If rs_Detalles!dlmonto >= 0 Then
            TotalDebe = TotalDebe + rs_Detalles!dlmonto
        Else
            TotalHaber = TotalHaber + Abs(rs_Detalles!dlmonto)
        End If
    
        rs_Detalles.MoveNext
    Loop
    
    If TotalDebe <> TotalHaber Then
        Balancea = False
        Flog.Writeline " Legajo No balancea"
        Flog.Writeline
        Flog.Writeline "     Debe     " & "    Haber     "
        Longitud = 14
        Aux_TotalDebe = CStr(Format(TotalDebe, "####0.00"))
        If Len(Aux_TotalDebe) < Longitud Then
            Aux_TotalDebe = String(Longitud - Len(Aux_TotalDebe), " ") & Aux_TotalDebe
        End If
        Longitud = 14
        Aux_TotalHaber = CStr(Format(TotalHaber, "####0.00"))
        If Len(Aux_TotalHaber) < Longitud Then
            Aux_TotalHaber = String(Longitud - Len(Aux_TotalHaber), " ") & Aux_TotalHaber
        End If
        Flog.Writeline Aux_TotalDebe & Aux_TotalHaber
        Flog.Writeline "Diferencia " & Abs(TotalDebe - TotalHaber)
    End If
    
    rs_Proc.MoveNext
Loop

If rs_Proc.State = adStateOpen Then rs_Proc.Close
If rs_Detalles.State = adStateOpen Then rs_Detalles.Close
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close

Set rs_Proc = Nothing
Set rs_Detalles = Nothing
Set rs_Empleado = Nothing
Set rs_Concepto = Nothing
Exit Sub

CE:
    HuboError = True
    Flog.Writeline " Error: " & Err.Description
End Sub

