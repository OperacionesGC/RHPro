Attribute VB_Name = "ExpCashManagement"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "22/10/2010" 'Martin Ferraro
'Global Const UltimaModificacion = "Version Inicial"

Global Const Version = "1.01"
Global Const FechaModificacion = "16/11/2010" 'Martin Ferraro
Global Const UltimaModificacion = "Se quitaron campos, se cambiaron nulos por blancos, cambio del formato de la fecha"




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Exportacion de Cash Management.
' Autor      : Martin Ferraro
' Fecha      : 22/10/2010
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
    
    Nombre_Arch = PathFLog & "RHProExpCashMag-" & NroProcesoBatch & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 276 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call generaExp(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
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
End Sub


Public Function FormatoFecha(ByVal Fecha As Date) As String
Dim aux As String
    
    aux = Format(Year(Fecha), "0000")
    Select Case Month(Fecha)
        Case 1:
            aux = aux & "-ENE-"
        Case 2:
            aux = aux & "-FEB-"
        Case 3:
            aux = aux & "-MAR-"
        Case 4:
            aux = aux & "-ABR-"
        Case 5:
            aux = aux & "-MAY-"
        Case 6:
            aux = aux & "-JUN-"
        Case 7:
            aux = aux & "-JUL-"
        Case 8:
            aux = aux & "-AGO-"
        Case 9:
            aux = aux & "-SEP-"
        Case 10:
            aux = aux & "-OCT-"
        Case 11:
            aux = aux & "-NOV-"
        Case 12:
            aux = aux & "-DIC-"
    End Select
    aux = aux & Format(Month(Fecha), "00")
    
    
    FormatoFecha = aux
    
End Function


Public Sub generaExp(ByVal BproNro As Long, ByVal Parametros As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera Reporte General de Liquidacion.
' Autor      : Martin Ferraro
' Fecha      : 22/10/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Arreglo que contiene los parametros
Dim arrParam

'Parametros desde ASP
Dim ListaPPago As String


'RecordSet
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

'Variables
Dim fs, f
Dim Path As String
Dim NroModelo As Long
Dim Directorio As String
Dim Carpeta
Dim ArchDef As String
Dim Nombre_Arch As String
Dim Dir_Arch As String
Dim Linea As String
Dim Sep As String
Dim ArchExp
Dim usaEncab As Long
    
' Inicio codigo ejecutable
On Error GoTo CE

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
    
    Flog.writeline Espacios(Tabulador * 1) & "Parametros " & Parametros
    arrParam = Split(Parametros, "@")
    If UBound(arrParam) = 1 Then
    
        ListaPPago = arrParam(0)
        Flog.writeline Espacios(Tabulador * 1) & "Pedidos de Pago = " & ListaPPago
        
        Nombre_Arch = arrParam(1)
        Flog.writeline Espacios(Tabulador * 1) & "Archivo a generar = " & Nombre_Arch
        
    Else
        Flog.writeline Espacios(Tabulador * 0) & "ERROR. La cantidad de parametros no es la esperada."
        HuboError = True
        Exit Sub
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encuentran los paramentros."
    HuboError = True
    Exit Sub
End If
Flog.writeline


'---------------------------------------------------------------------------------------------------
'Datos del modelo
'---------------------------------------------------------------------------------------------------
NroModelo = 330
Flog.writeline Espacios(Tabulador * 0) & "Buscando datos del modelo " & NroModelo
StrSql = "SELECT sis_dirsalidas FROM sistema"
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
   Directorio = Trim(rs_Consult!sis_dirsalidas)
End If
rs_Consult.Close

StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
   Sep = IIf(Not IsNull(rs_Consult!modseparador), rs_Consult!modseparador, "")
   If Not IsNull(rs_Consult!modarchdefault) Then
      ArchDef = Trim(rs_Consult!modarchdefault)
   Else
      Flog.writeline Espacios(Tabulador * 1) & "ERROR. El modelo no tiene configurada la carpeta destino."
      HuboError = True
      Exit Sub
   End If
   usaEncab = IIf(Not IsNull(rs_Consult!modencab), rs_Consult!modencab, "0")
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el modelo " & NroModelo
    HuboError = True
    Exit Sub
End If
         
Directorio = Directorio & Trim(rs_Consult!modarchdefault)
         
'---------------------------------------------------------------------------------------------------
'Creacion del archivo
'---------------------------------------------------------------------------------------------------
Dir_Arch = Directorio & "\" & Nombre_Arch
Flog.writeline Espacios(Tabulador * 1) & "Generando Archivo " & Dir_Arch
On Error Resume Next
Set fs = CreateObject("Scripting.FileSystemObject")
Set ArchExp = fs.CreateTextFile(Dir_Arch, True)
If Err.Number <> 0 Then
     Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
     Set Carpeta = fs.CreateFolder(Directorio)
     Set ArchExp = fs.CreateTextFile(Dir_Arch, True)
End If
'desactivo el manejador de errores
On Error GoTo CE

'---------------------------------------------------------------------------------
'Imprimo Encabezado
'---------------------------------------------------------------------------------
If usaEncab = -1 Then
    'Linea = "ROWID" & Sep & "TRX_ID" & Sep & "BANK_ACCOUNT_ID" & Sep & "TRX_TYPE" & Sep & "TRX_TYPE_DSP" & Sep & "TRX_NUMBER" & Sep & "TRX_DATE" & Sep & "CURRENCY_CODE" & Sep & "STATUS" & Sep & "STATUS_DSP" & Sep & "EXCHANCE_RATE_TYPE" & Sep & "EXCHANGE_RATE_DATE" & Sep & "EXCHANGE_RATE" & Sep & "AMOUNT" & Sep & "CLEARED_AMOUNT" & Sep & "ERROR_AMOUNT" & Sep & "ACCTD_AMOUNT" & Sep & "ACCTD_CLEARED" & Sep & "ACCTD_CHARGES_AMOUNT" & Sep & "ACCTD_ERROR_AMOUNT" & Sep & "GL_DATE" & Sep & "CLEARED_DATE" & Sep & "CREATION_DATE" & Sep & "CREATED_BY" & Sep & "LAST_UPDATE_DATE" & Sep & "LAST_UPDATED_BY"
    Linea = "BANK_ACCOUNT_ID" & Sep & "TRX_TYPE" & Sep & "TRX_TYPE_DSP" & Sep & "TRX_NUMBER" & Sep & "TRX_DATE" & Sep & "CURRENCY_CODE" & Sep & "STATUS" & Sep & "STATUS_DSP" & Sep & "EXCHANCE_RATE_TYPE" & Sep & "EXCHANGE_RATE_DATE" & Sep & "EXCHANGE_RATE" & Sep & "AMOUNT" & Sep & "CLEARED_AMOUNT" & Sep & "ERROR_AMOUNT" & Sep & "ACCTD_AMOUNT" & Sep & "ACCTD_CLEARED" & Sep & "ACCTD_CHARGES_AMOUNT" & Sep & "ACCTD_ERROR_AMOUNT" & Sep & "GL_DATE" & Sep & "CLEARED_DATE"
    ArchExp.writeline Linea
End If

'---------------------------------------------------------------------------------
'Consulta Principal
'---------------------------------------------------------------------------------
StrSql = "SELECT pagomonto, ctabnro, pago.fpagnro, ppagnroped, ppagfecped"
StrSql = StrSql & " FROM pago"
StrSql = StrSql & " INNER JOIN pedidopago ON pedidopago.ppagnro = pago.ppagnro"
StrSql = StrSql & " WHERE pago.ppagnro IN (" & ListaPPago & ")"
StrSql = StrSql & " ORDER BY pago.ppagnro"
OpenRecordset StrSql, rs_Empleados

'---------------------------------------------------------------------------------
'seteo de las variables de progreso
'---------------------------------------------------------------------------------
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
   Flog.writeline "no hay empleados"
   CEmpleadosAProc = 1
End If
IncPorc = (99 / CEmpleadosAProc)
        
        
Flog.writeline
Flog.writeline
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el procesamiento de Registros"
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------"
        

'Comienzo a procesar los empleados
Do While Not rs_Empleados.EOF
    
    Linea = ""
    'ROW_ID
    'Linea = "NULL"
    'TRX_ID
    'Linea = Linea & Sep & "NULL"
    'BANK_ACCOUNT_ID
    Linea = Linea & Sep & IIf((rs_Empleados!fpagnro = "10") Or (rs_Empleados!fpagnro = "11"), rs_Empleados!ctabnro, " ")
    'TRX_TYPE
    Linea = Linea & Sep & "PAYMENT"
    'TRX_TYPE_DSP
    Linea = Linea & Sep & "PAYMENT"
    'TRX_TYPE_NUMBER
    Linea = Linea & Sep & IIf(EsNulo(rs_Empleados!ppagnroped), "", rs_Empleados!ppagnroped)
    'TRX_DATE
    Linea = Linea & Sep & IIf(EsNulo(rs_Empleados!ppagfecped), "", FormatoFecha(rs_Empleados!ppagfecped))
    'CURRENCY_CODE
    Linea = Linea & Sep & "ARS"
    'STATUS
    Linea = Linea & Sep & "Available"
    'STATUS_DSP
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & "Available"
    'EXCHANGE_RATE_TYPE
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'EXCHANGE_RATE_DATE
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'EXCHANGE_RATE
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'AMOUNT
    Linea = Linea & Sep & Replace(FormatNumber(rs_Empleados!pagomonto, 2), ",", "")
    'CLEARED_AMOUNT
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'ERROR_AMOUNT
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'ACCTD_AMOUNT
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'ACCTD_CLEARED
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'ACCTD_CHARGES_AMOUNT
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'ACCTD_ERROR_AMOUNT
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'GL_DATE
    Linea = Linea & Sep & FormatoFecha(Now)
    'CLEARED_DATE
    'Linea = Linea & Sep & "NULL"
    Linea = Linea & Sep & " "
    'CREATION_DATE
    'Linea = Linea & Sep & "NULL"
    'CREATED_BY
    'Linea = Linea & Sep & "NULL"
    'LAST_UPDATE_DATE
    'Linea = Linea & Sep & "NULL"
    'LAST_UPDATE_BY
    'Linea = Linea & Sep & "NULL"
       
    ArchExp.writeline Linea
    
    '---------------------------------------------------------------------------------
    'Actualizo el progreso
    '---------------------------------------------------------------------------------
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Paso a siguiente cabliq
    rs_Empleados.MoveNext
    
Loop

ArchExp.Close

If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close


Set rs_Empleados = Nothing
Set rs_Consult = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True

End Sub
