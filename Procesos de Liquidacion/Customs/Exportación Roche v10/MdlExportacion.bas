Attribute VB_Name = "MdlExportacion"
Option Explicit

'Const Version = "1.01"
'Const FechaVersion = "06/02/2007"   'Version Inicial.

'Const Version = "1.02"
'Const FechaVersion = "15/09/2008"   'Exportacion de Archivos de Prestamo y Detalle de prestamos (7 y 8)

Const Version = "1.03"
Const FechaVersion = "31/07/2009"   'Encriptacion de string connection

'-----------------------------------------------------------------------------------
Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global Fecha_Desde As Date
Global Fecha_Hasta As Date
Global Progreso As Double
Global StrSql2 As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : JMH
' Fecha      : 07/03/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

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
    'End If
    
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

    'Abro la conexion
    'OpenConnection strconexion, objConn
    'OpenConnection strconexion, objconnProgreso
    
        
    Nombre_Arch = PathFLog & "Exp_Roche" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 77 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        If Not IsNull(rs_batch_proceso!bprcfecdesde) Then
           Fecha_Desde = rs_batch_proceso!bprcfecdesde
        End If
        If Not IsNull(rs_batch_proceso!bprcfechasta) Then
           Fecha_Hasta = rs_batch_proceso!bprcfechasta
        End If
        
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' ,bprcprogreso=100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Close
    
    objconnProgreso.Close
    objConn.Close
End Sub

Public Sub Generacion1(ByVal Exportacion As Integer, ByVal Anio As Integer, ByVal mes As Integer, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : JMH
' Fecha      : 07/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim rs_Periodo As New ADODB.Recordset
Dim rs_Proceso As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String
Dim pos
Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Long
Dim cantidadProcesada As Long

Dim idLiq As String
Dim NroPeriodo As String
Dim FechaPago As Date
Dim HayInformacion As Boolean

On Error GoTo MError

HayInformacion = False
'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 251"
OpenRecordset StrSql, rs_Modelo

If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
'On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\PayrollPeriod_AR.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If

'desactivo el manejador de errores
'On Error GoTo 0

' Comienzo la transaccion
MyBeginTrans

Flog.writeline "Buscando los empleados del periodo."
StrSql = "SELECT empleg, cabliq.cliqnro, periodo.pliqnro,periodo.pliqmes, pliqdesde, pliqhasta, profecpago, periodo.pliqmes "
StrSql = StrSql & " FROM  periodo "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN cabliq   ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
StrSql = StrSql & " INNER JOIN detliq   ON detliq.cliqnro = cabliq.cliqnro "
StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
StrSql = StrSql & " AND his_estructura.tenro  = 32 AND (his_estructura.htetdesde <= periodo.pliqhasta) "
StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= periodo.pliqdesde) "
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
'--------------------------------------------------------------------------------------
'FGZ - 21/07/2005
'solo los conceptos imprimibles
StrSql = StrSql & " INNER JOIN concepto ON detliq.concnro = concepto.concnro AND concepto.concimp = -1"
'FGZ - 21/07/2005
'--------------------------------------------------------------------------------------
StrSql = StrSql & " WHERE periodo.pliqanio = " & Anio & " AND periodo.pliqmes = " & mes
StrSql = StrSql & " GROUP BY empleg, cabliq.cliqnro, periodo.pliqnro,periodo.pliqmes, pliqdesde, pliqhasta, profecpago, periodo.pliqmes "
StrSql = StrSql & " ORDER BY periodo.pliqnro, empleado.empleg "
OpenRecordset StrSql, rs_Periodo

Cantidad = rs_Periodo.RecordCount
cantidadProcesada = Cantidad
If Cantidad = 0 Then Cantidad = 1
IncPorc = 99 / Cantidad
Progreso = 0
Dim Error As Boolean

Do While Not rs_Periodo.EOF
   HayInformacion = True
               
   Error = False
   
   StrSql = "SELECT proceso.tprocnro, profecpago "
   StrSql = StrSql & " FROM proceso "
   StrSql = StrSql & " WHERE proceso.pliqnro = " & rs_Periodo!pliqnro
   OpenRecordset StrSql, rs_Proceso

   Do While Not rs_Proceso.EOF
   
      If rs_Proceso!tprocnro = 3 Then
         FechaPago = rs_Proceso!profecpago
      End If
      
      rs_Proceso.MoveNext
   Loop
   rs_Proceso.Close
   
'   idLiq = rs_Periodo!pliqnro & Day(FechaPago) & Month(FechaPago) & Year(FechaPago)
'   NroPeriodo = Format_StrNro(rs_Periodo!pliqmes, 2, True, "0")

   'FGZ - 31/05/2005
   NroPeriodo = Format_StrNro(rs_Periodo!pliqmes, 2, True, "0")
   idLiq = Format(FechaPago, "yyyy-MM-dd") & Chr(9) & NroPeriodo
   
   Cabecera = Format_StrNro(rs_Periodo!empleg, 7, False, "") & Chr(9) & idLiq & Chr(9)
   Cabecera = Cabecera & NroPeriodo & Chr(9) & Format(rs_Periodo!pliqdesde, "yyyy-MM-dd") & Chr(9)
   Cabecera = Cabecera & Format(rs_Periodo!pliqhasta, "yyyy-MM-dd") & Chr(9) & Format(FechaPago, "yyyy-MM-dd")
   
   Flog.writeline "Guardando los datos en el archivo."
        
   fExport.writeline Cabecera
        
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   'Progreso = Fix(((Cantidad - cantidadProcesada) * 100#) / Cantidad)
   Progreso = Progreso + IncPorc
        Flog.writeline "Actualizando el progreso del estado."
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                 ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
   rs_Periodo.MoveNext
Loop

rs_Periodo.Close
If Not HayInformacion Then
    fExport.writeline
End If
fExport.Close

MyCommitTrans

Set rs_Periodo = Nothing
Set rs_Proceso = Nothing

Exit Sub

MError:
    MyRollbackTrans
    Flog.writeline "Error: " & Err.Description

End Sub


Public Sub Generacion2(ByVal Exportacion As Integer, ByVal Periodo As Integer, ByVal Proceso As String, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : JMH
' Fecha      : 07/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String
Dim pos
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer

Dim idLiq As String
Dim NroPeriodo As String
Dim Percepciones As Double
Dim Deducciones As Double
Dim NetoPagar As Double
Dim Deduccion As Double
Dim Percepcion As String
Dim DeduccionStr As String
Dim PercepcionStr As String
Dim NetoPagarStr As String
Dim NetoPagarAux As String

Dim arrTipoConc(1000) As Integer
Dim I As Integer
Dim HayInformacion As Boolean

On Error GoTo MError

HayInformacion = False
'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 251"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
'On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\PayrollGeneral_AR.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If

'desactivo el manejador de errores
'On Error GoTo 0

' Comienzo la transaccion
MyBeginTrans


'Inicializo los tipos de conceptos
For I = 1 To 1000
    arrTipoConc(I) = 0
Next

'Busco el tipo de cada concepto
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 "
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
Do Until rs.EOF
    Select Case rs!conftipo
       'Remunerativo
       Case "RE"
          arrTipoConc(rs!confval) = 1
       'No Remunerativo
       Case "NR"
          arrTipoConc(rs!confval) = 2
       'Descuento
       Case "DS"
          arrTipoConc(rs!confval) = 3
    End Select
 
    rs.MoveNext
 Loop
 
Flog.writeline "Buscando el detalle de la liquidacion."
StrSql = " SELECT empleg, periodo.pliqnro,periodo.pliqmes, profecpago, estructura.estrdabr, cabliq.cliqnro "
StrSql = StrSql & " FROM  periodo "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pliqnro = periodo.pliqnro AND proceso.pronro IN (" & Proceso & ")"
StrSql = StrSql & " INNER JOIN cabliq   ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
StrSql = StrSql & " AND his_estructura.tenro  = 32 AND (his_estructura.htetdesde <= periodo.pliqhasta) "
StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= periodo.pliqdesde) "
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE periodo.pliqnro = " & Periodo
StrSql = StrSql & " GROUP BY empleg, periodo.pliqnro,periodo.pliqmes, profecpago, estructura.estrdabr, cabliq.cliqnro "
OpenRecordset StrSql, rs_Periodo

Cantidad = rs_Periodo.RecordCount
cantidadProcesada = Cantidad
If Cantidad = 0 Then Cantidad = 1
IncPorc = 99 / Cantidad
Progreso = 0
Dim Error As Boolean

Do While Not rs_Periodo.EOF
   HayInformacion = True
   Error = False
   
    Flog.writeline "Buscando el detalle de los conceptos para el empleado."
    StrSql = " SELECT detliq.concnro, detliq.dlimonto, concepto.tconnro "
    StrSql = StrSql & " FROM  detliq "
    'StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
    '--------------------------------------------------------------------------------------
    'FGZ - 21/07/2005
    'solo los conceptos imprimibles
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1"
    'FGZ - 21/07/2005
    '--------------------------------------------------------------------------------------
    StrSql = StrSql & " WHERE detliq.cliqnro = " & rs_Periodo!cliqnro
    OpenRecordset StrSql, rs_Detliq
    
    Percepciones = 0
    Deducciones = 0
   
    Do While Not rs_Detliq.EOF
        Flog.writeline "Controlando el tipo de concepto, percepcion-deduccion."
        StrSql = " SELECT * "
        StrSql = StrSql & " FROM confrep "
        StrSql = StrSql & " WHERE confrep.repnro = 122 AND confrep.confval = " & rs_Detliq!tconnro
        If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
        OpenRecordset StrSql, rs_Confrep
    
        If Not rs_Confrep.EOF Then
            If rs_Confrep!conftipo = "PER" Then
                Percepciones = Percepciones + rs_Detliq!dlimonto
            Else
                Deducciones = Deducciones + rs_Detliq!dlimonto
            End If
        End If
       
'        If arrTipoConc(rs_Detliq!tconnro) = 1 Or arrTipoConc(rs_Detliq!tconnro) = 2 Then
'           Percepciones = Percepciones + rs_Detliq!dlimonto
'        Else
'           Deducciones = Deducciones + rs_Detliq!dlimonto
'        End If
    
        rs_Detliq.MoveNext
    Loop
    rs_Detliq.Close
   
   
   'idLiq = rs_Periodo!pliqnro & Day(rs_Periodo!profecpago) & Month(rs_Periodo!profecpago) & Year(rs_Periodo!profecpago)
   'FGZ - 31/05/2005
   NroPeriodo = Format_StrNro(rs_Periodo!pliqmes, 2, True, "0")
   idLiq = Format(rs_Periodo!profecpago, "yyyy-MM-dd") & Chr(9) & NroPeriodo
   
   If Deducciones >= 0 Then
      NetoPagar = Percepciones - Deducciones
   Else: NetoPagar = Percepciones + Deducciones
   End If
   
   Deducciones = FormatNumber(Deducciones, 2)
   Percepciones = Round(Percepciones, 2)
   
   pos = InStr(1, Deducciones, ".")
   If pos = 0 Then
      Deduccion = 0
   Else
      Deduccion = Mid(Deducciones, pos + 1, Len(Deducciones) - pos)
   End If
   DeduccionStr = Fix(Deducciones) & "." & Deduccion
   
   pos = InStr(1, Percepciones, ".")
   If pos = 0 Then
      Percepcion = 0
   Else
      Percepcion = Mid(Percepciones, pos + 1, Len(Percepciones) - pos)
   End If
   PercepcionStr = Fix(Percepciones) & "." & Percepcion
      
   pos = InStr(1, NetoPagar, ".")
   If pos = 0 Then
      NetoPagarAux = 0
   Else
      NetoPagarAux = Mid(NetoPagar, pos + 1, Len(NetoPagar) - pos)
   End If
   NetoPagarStr = CStr(Fix(NetoPagar)) & "." & NetoPagarAux

      
   Cabecera = Format_StrNro(rs_Periodo!empleg, 7, False, "") & Chr(9) & idLiq & Chr(9)
   Cabecera = Cabecera & rs_Periodo!estrdabr & Chr(9) & Format(rs_Periodo!profecpago, "yyyy-MM-dd") & Chr(9)
   Cabecera = Cabecera & "" & Chr(9) & PercepcionStr & Chr(9) & DeduccionStr & Chr(9) & NetoPagarStr
   
   Flog.writeline "Escribiendo los datos en el archivo."
        
   fExport.writeline Cabecera
        
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = Progreso + IncPorc
   Flog.writeline "Actualizando el estado del proceso."
    
   StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
   objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
   rs_Periodo.MoveNext
Loop

rs_Periodo.Close
If Not HayInformacion Then
    fExport.writeline
End If
fExport.Close

MyCommitTrans

Set rs_Periodo = Nothing
Set rs_Detliq = Nothing
Set rs_Confrep = Nothing

Exit Sub
MError:
    MyRollbackTrans
    Flog.writeline "Error: " & Err.Description

End Sub


Public Sub Generacion3(ByVal Exportacion As Integer, ByVal Periodo As Integer, ByVal Proceso As String, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : JMH
' Fecha      : 07/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String
Dim pos
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Long
Dim cantidadProcesada As Long

Dim idLiq As String
Dim NroPeriodo As String
Dim PercepcionStr As String
Dim Percepcion As String
Dim HayInformacion As Boolean

On Error GoTo MError

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 251"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
'On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\PayrollPerceptions_AR.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If

'desactivo el manejador de errores
'On Error GoTo 0

HayInformacion = False
' Comienzo la transaccion
MyBeginTrans

StrSql = " SELECT empleg, periodo.pliqnro, periodo.pliqmes, profecpago, estructura.estrdabr, cabliq.cliqnro "
StrSql = StrSql & " FROM  periodo "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pliqnro = periodo.pliqnro AND proceso.pronro IN (" & Proceso & ")"
StrSql = StrSql & " INNER JOIN cabliq   ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
StrSql = StrSql & " AND his_estructura.tenro  = 32 AND (his_estructura.htetdesde <= periodo.pliqhasta) "
StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= periodo.pliqdesde) "
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE periodo.pliqnro = " & Periodo
StrSql = StrSql & " GROUP BY empleg, periodo.pliqnro, periodo.pliqmes, profecpago, estructura.estrdabr, cabliq.cliqnro "

Flog.writeline "Buscando el detalle de liquidacion."

OpenRecordset StrSql, rs_Periodo

Cantidad = rs_Periodo.RecordCount
cantidadProcesada = Cantidad

Dim Error As Boolean
If Cantidad = 0 Then Cantidad = 1
IncPorc = 99 / Cantidad
Progreso = 0
Do While Not rs_Periodo.EOF
           
   Error = False
   
   Flog.writeline "Buscando el detalle de liquidacion para el empleado."
   StrSql = " SELECT detliq.concnro, detliq.dlimonto, detliq.dlicant, concepto.tconnro, concepto.conccod, concepto.concabr "
   StrSql = StrSql & " FROM  detliq "
   'StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
   '--------------------------------------------------------------------------------------
   'FGZ - 21/07/2005
   'solo los conceptos imprimibles
   StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1"
   'FGZ - 21/07/2005
   '--------------------------------------------------------------------------------------
   StrSql = StrSql & " WHERE detliq.cliqnro = " & rs_Periodo!cliqnro
   OpenRecordset StrSql, rs_Detliq
    
   Do While Not rs_Detliq.EOF
       HayInformacion = True
       StrSql = " SELECT * "
       StrSql = StrSql & " FROM confrep "
       StrSql = StrSql & " WHERE confrep.repnro = 122 AND confrep.confval = " & rs_Detliq!tconnro
       OpenRecordset StrSql, rs_Confrep
       
       Flog.writeline "Controlando si el concepto es de percepcion-deduccion."
    
       If Not rs_Confrep.EOF Then
          If rs_Confrep!conftipo = "PER" Then
             
             'PercepcionStr = Fix(rs_Detliq!dlimonto) & "." & Abs(Round((rs_Detliq!dlimonto - Fix(rs_Detliq!dlimonto)) * 100))
             pos = InStr(1, rs_Detliq!dlimonto, ".")
             If pos = 0 Then
                Percepcion = 0
             Else
                Percepcion = Mid(rs_Detliq!dlimonto, pos + 1, Len(rs_Detliq!dlimonto) - pos)
             End If
             PercepcionStr = Fix(rs_Detliq!dlimonto) & "." & Percepcion
             
             'idLiq = rs_Periodo!pliqnro & Day(rs_Periodo!profecpago) & Month(rs_Periodo!profecpago) & Year(rs_Periodo!profecpago)
             'FGZ - 31/05/2005
             NroPeriodo = Format_StrNro(rs_Periodo!pliqmes, 2, True, "0")
             idLiq = Format(rs_Periodo!profecpago, "yyyy-MM-dd") & Chr(9) & NroPeriodo

             Cabecera = Format_StrNro(rs_Periodo!empleg, 7, False, "") & Chr(9) & idLiq & Chr(9)
             Cabecera = Cabecera & rs_Detliq!Conccod & Chr(9) & rs_Detliq!concabr & Chr(9)
             Cabecera = Cabecera & FormatNumber(rs_Detliq!dlicant, 2) & Chr(9) & PercepcionStr
             
             Flog.writeline "Escribiendo los datos en el archivo."
        
             fExport.writeline Cabecera
          End If
       End If
       rs_Confrep.Close
       
       rs_Detliq.MoveNext
   Loop
   rs_Detliq.Close
   
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = Progreso + IncPorc
    Flog.writeline "Actualizando el estado del proceso."
    
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
   rs_Periodo.MoveNext
Loop

rs_Periodo.Close
If Not HayInformacion Then
    fExport.writeline
End If
fExport.Close

MyCommitTrans

Set rs_Periodo = Nothing
Set rs_Detliq = Nothing
Set rs_Confrep = Nothing

Exit Sub
MError:
    MyRollbackTrans
    Flog.writeline "Error: " & Err.Description

End Sub


Public Sub Generacion4(ByVal Exportacion As Integer, ByVal Periodo As Integer, ByVal Proceso As String, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : JMH
' Fecha      : 07/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim rs_Periodo As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Long
Dim cantidadProcesada As Long

Dim idLiq As String
Dim NroPeriodo As String
Dim DeduccionStr As String
Dim Deduccion As String
Dim HayInformacion As Boolean
Dim pos

On Error GoTo MError

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 251"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
'On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\PayrollDeductions_AR.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If

'desactivo el manejador de errores
'On Error GoTo 0

HayInformacion = False
' Comienzo la transaccion
MyBeginTrans

StrSql = " SELECT empleg, periodo.pliqnro, periodo.pliqmes, profecpago, estructura.estrdabr, cabliq.cliqnro "
StrSql = StrSql & " FROM  periodo "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pliqnro = periodo.pliqnro AND proceso.pronro IN (" & Proceso & ")"
StrSql = StrSql & " INNER JOIN cabliq   ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
StrSql = StrSql & " AND his_estructura.tenro  = 32 AND (his_estructura.htetdesde <= periodo.pliqhasta) "
StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= periodo.pliqdesde) "
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE periodo.pliqnro = " & Periodo
StrSql = StrSql & " GROUP BY empleg, periodo.pliqnro, periodo.pliqmes, profecpago, estructura.estrdabr, cabliq.cliqnro "

Flog.writeline "Buscando el detalle de liquidacion."

OpenRecordset StrSql, rs_Periodo

Cantidad = rs_Periodo.RecordCount
cantidadProcesada = Cantidad

Dim Error As Boolean
If Cantidad = 0 Then Cantidad = 1
IncPorc = 99 / Cantidad
Progreso = 0
Do While Not rs_Periodo.EOF
           
   Error = False
   
   Flog.writeline "Buscando el detalle de liquidacion para el empleado."
   StrSql = " SELECT detliq.concnro, detliq.dlimonto, detliq.dlicant, concepto.tconnro, concepto.conccod, concepto.concabr "
   StrSql = StrSql & " FROM  detliq "
   'StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
   '--------------------------------------------------------------------------------------
   'FGZ - 21/07/2005
   'solo los conceptos imprimibles
   StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1"
   'FGZ - 21/07/2005
   '--------------------------------------------------------------------------------------
   StrSql = StrSql & " WHERE detliq.cliqnro = " & rs_Periodo!cliqnro
   OpenRecordset StrSql, rs_Detliq
    
   Do While Not rs_Detliq.EOF
       HayInformacion = True
       StrSql = " SELECT * "
       StrSql = StrSql & " FROM confrep "
       StrSql = StrSql & " WHERE confrep.repnro = 122 AND confrep.confval = " & rs_Detliq!tconnro
       OpenRecordset StrSql, rs_Confrep
    
       If Not rs_Confrep.EOF Then
          If rs_Confrep!conftipo = "DED" Then
                          
'             Deduccion = Round((rs_Detliq!dlimonto - Fix(rs_Detliq!dlimonto)) * 100)
'             If Deduccion < 0 Then
'                Deduccion = (-1) * Deduccion
'             End If
             
             pos = InStr(1, rs_Detliq!dlimonto, ".")
             If pos = 0 Then
                Deduccion = 0
             Else
                Deduccion = Mid(rs_Detliq!dlimonto, pos + 1, Len(rs_Detliq!dlimonto) - pos)
             End If
             'Deduccion = Round((rs_Detliq!dlimonto - Fix(rs_Detliq!dlimonto)) * 100)
             DeduccionStr = Fix(rs_Detliq!dlimonto) & "." & Deduccion
             
             'idLiq = rs_Periodo!pliqnro & Day(rs_Periodo!profecpago) & Month(rs_Periodo!profecpago) & Year(rs_Periodo!profecpago)
             'FGZ - 31/05/2005
             NroPeriodo = Format_StrNro(rs_Periodo!pliqmes, 2, True, "0")
             idLiq = Format(rs_Periodo!profecpago, "yyyy-MM-dd") & Chr(9) & NroPeriodo
                          
             Cabecera = Format_StrNro(rs_Periodo!empleg, 7, False, "") & Chr(9) & idLiq & Chr(9)
             Cabecera = Cabecera & rs_Detliq!Conccod & Chr(9) & rs_Detliq!concabr & Chr(9)
             Cabecera = Cabecera & FormatNumber(rs_Detliq!dlicant, 2) & Chr(9) & DeduccionStr
             
             Flog.writeline "Guardando los datos en el archivo."
        
             fExport.writeline Cabecera
          End If
       End If
       rs_Confrep.Close
       
       rs_Detliq.MoveNext
   Loop
   rs_Detliq.Close
   
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = Progreso + IncPorc
    Flog.writeline "Actualizando el estado del proceso."
    
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
   rs_Periodo.MoveNext
Loop

rs_Periodo.Close
If Not HayInformacion Then
    fExport.writeline
End If
fExport.Close

MyCommitTrans

Set rs_Periodo = Nothing
Set rs_Detliq = Nothing
Set rs_Confrep = Nothing

Exit Sub
MError:
    MyRollbackTrans
    Flog.writeline "Error: " & Err.Description

End Sub


Public Sub Generacion5(ByVal Exportacion As Integer, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : JMH
' Fecha      : 07/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String
Dim pos
Dim rs_Prestamo As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer

Dim idLiq As String
Dim NroPeriodo As String
Dim Monto As Double
Dim MontoAux As String
Dim HayInformacion As Boolean

On Error GoTo MError

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 251"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
'On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\Credits_AR.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If

'desactivo el manejador de errores
'On Error GoTo 0

HayInformacion = False
' Comienzo la transaccion
MyBeginTrans

StrSql = " SELECT DISTINCT empleg, prestamo.prenro, tipoprestamo.tpdesabr, estructura.estrdabr, prestamo.precantcuo, "
StrSql = StrSql & " prestamo.prefecotor, pre_cuota.cuofecvto, prestamo.preimp "
StrSql = StrSql & " FROM  prestamo "
StrSql = StrSql & " INNER JOIN pre_cuota ON pre_cuota.prenro = prestamo.prenro AND pre_cuota.cuonrocuo = 1"
StrSql = StrSql & " INNER JOIN pre_linea ON pre_linea.lnprenro = prestamo.lnprenro "
StrSql = StrSql & " INNER JOIN tipoprestamo ON tipoprestamo.tpnro = pre_linea.tpnro "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = prestamo.ternro "
StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
StrSql = StrSql & " AND his_estructura.tenro  = 32 AND (his_estructura.htetdesde <= " & ConvFecha(FechaHasta) & ") "
StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(FechaDesde) & ") "
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE prestamo.prefecotor >= " & ConvFecha(FechaDesde) & " AND prestamo.prefecotor <= " & ConvFecha(FechaHasta)
OpenRecordset StrSql, rs_Prestamo

Cantidad = rs_Prestamo.RecordCount
cantidadProcesada = Cantidad

Dim Error As Boolean
If Cantidad = 0 Then Cantidad = 1
IncPorc = 99 / Cantidad
Progreso = 0
Do While Not rs_Prestamo.EOF
    HayInformacion = True
   Error = False
   
   'Monto = Fix(rs_Prestamo!preimp) & "." & Abs(Round((rs_Prestamo!preimp - Fix(rs_Prestamo!preimp)) * 100))
    pos = InStr(1, rs_Prestamo!preimp, ".")
    If pos = 0 Then
       MontoAux = 0
    Else
       MontoAux = Mid(rs_Prestamo!preimp, pos + 1, Len(rs_Prestamo!preimp) - pos)
    End If
    Monto = Fix(rs_Prestamo!preimp) & "." & MontoAux
   
   Cabecera = Format_StrNro(rs_Prestamo!empleg, 7, False, "") & Chr(9) & rs_Prestamo!prenro & Chr(9)
   Cabecera = Cabecera & rs_Prestamo!estrdabr & Chr(9) & rs_Prestamo!tpdesabr & Chr(9) & " " & Chr(9)
   Cabecera = Cabecera & rs_Prestamo!precantcuo & Chr(9) & Format(rs_Prestamo!prefecotor, "yyyy-MM-dd") & Chr(9)
   Cabecera = Cabecera & Format(rs_Prestamo!cuofecvto, "yyyy-MM-dd") & Chr(9) & Monto
        
   fExport.writeline Cabecera
       
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
   rs_Prestamo.MoveNext
Loop

rs_Prestamo.Close
If Not HayInformacion Then
    fExport.writeline
End If
fExport.Close

MyCommitTrans

Set rs_Prestamo = Nothing

Exit Sub

MError:
    MyRollbackTrans
    Flog.writeline "Error: " & Err.Description
    
End Sub


Public Sub Generacion6(ByVal Exportacion As Integer, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : JMH
' Fecha      : 07/03/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String
Dim pos
Dim rs_Prestamo As New ADODB.Recordset
Dim rs_Cuota As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer

Dim idLiq As String
Dim NroPeriodo As String
Dim MontoRestante As Double
Dim MontoPago As Double
Dim MontoTotal As Double
Dim MontoRestanteStr As String
Dim MontoPagoStr As String
Dim MontoTotalStr As String
Dim HayInformacion As Boolean
Dim MontoAux As String

On Error GoTo MError

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 251"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
'On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\CreditsDetails_AR.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If

'desactivo el manejador de errores
'On Error GoTo 0

HayInformacion = False
' Comienzo la transaccion
MyBeginTrans

StrSql = " SELECT empleg, prestamo.prenro,  prestamo.prefecotor "
StrSql = StrSql & " FROM  prestamo "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = prestamo.ternro "
StrSql = StrSql & " WHERE prestamo.prefecotor >= " & ConvFecha(FechaDesde) & " AND prestamo.prefecotor <= " & ConvFecha(FechaHasta)
StrSql = StrSql & " ORDER BY empleado.empleg "

OpenRecordset StrSql, rs_Prestamo

Cantidad = rs_Prestamo.RecordCount
cantidadProcesada = Cantidad
Dim Error As Boolean
If Cantidad = 0 Then Cantidad = 1
IncPorc = 99 / Cantidad
Progreso = 0
Do While Not rs_Prestamo.EOF
           
   Error = False
   
   StrSql = " SELECT cuocancela, cuototal, cuofecvto, cuonrocuo, cuosaldo "
   StrSql = StrSql & " FROM  pre_cuota "
   StrSql = StrSql & " WHERE pre_cuota.prenro = " & rs_Prestamo!prenro
   StrSql = StrSql & " ORDER BY pre_cuota.cuonrocuo "

   OpenRecordset StrSql, rs_Cuota

   MontoTotal = 0
   MontoPago = 0
   MontoRestante = 0
   
   Do While Not rs_Cuota.EOF
        HayInformacion = True
      If rs_Cuota!cuonrocuo = 1 Then
         MontoTotal = rs_Cuota!cuototal + rs_Cuota!cuosaldo
      End If
      
      If rs_Cuota!cuocancela = -1 And rs_Cuota!cuofecvto <= FechaHasta And rs_Cuota!cuofecvto >= FechaDesde Then
         MontoPago = rs_Cuota!cuototal
         MontoRestante = rs_Cuota!cuosaldo
      Else: MontoRestante = MontoRestante
            MontoPago = 0
      End If
      
      pos = InStr(1, MontoPago, ".")
      If pos = 0 Then
         MontoAux = 0
      Else
         MontoAux = Mid(MontoPago, pos + 1, Len(MontoPago) - pos)
      End If
      'MontoPagoStr = Fix(MontoPago) & "." & Round((MontoPago - Fix(MontoPago)) * 100)
      MontoPagoStr = Fix(MontoPago) & "." & MontoAux
      
      pos = InStr(1, MontoRestante, ".")
      If pos = 0 Then
         MontoAux = 0
      Else
         MontoAux = Mid(MontoRestante, pos + 1, Len(MontoRestante) - pos)
      End If
      'MontoRestanteStr = Fix(MontoRestante) & "." & Round((MontoRestante - Fix(MontoRestante)) * 100)
      MontoRestanteStr = Fix(MontoRestante) & "." & MontoAux
      
      pos = InStr(1, MontoTotalStr, ".")
      If pos = 0 Then
         MontoAux = 0
      Else
         MontoAux = Mid(MontoTotalStr, pos + 1, Len(MontoTotalStr) - pos)
      End If
      'MontoTotalStr = Fix(MontoTotal) & "." & Round((MontoTotal - Fix(MontoTotal)) * 100)
      MontoTotalStr = Fix(MontoTotal) & "." & MontoAux
      
      Cabecera = Format_StrNro(rs_Prestamo!empleg, 7, False, "") & Chr(9) & rs_Prestamo!prenro & Chr(9)
      Cabecera = Cabecera & rs_Cuota!cuonrocuo & Chr(9) & Format(rs_Cuota!cuofecvto, "yyyy-MM-dd") & Chr(9)
      Cabecera = Cabecera & MontoPagoStr & Chr(9) & MontoRestanteStr & Chr(9)
      Cabecera = Cabecera & MontoTotalStr
      
      fExport.writeline Cabecera
      
      rs_Cuota.MoveNext
      
   Loop
   rs_Cuota.Close
   
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
   rs_Prestamo.MoveNext
Loop

rs_Prestamo.Close
If Not HayInformacion Then
    fExport.writeline
End If
fExport.Close

MyCommitTrans

Set rs_Prestamo = Nothing
Set rs_Cuota = Nothing
Exit Sub

MError:
    MyRollbackTrans
    Flog.writeline "Error: " & Err.Description

End Sub


Public Sub Generacion7(ByVal Exportacion As Integer, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : FGZ
' Fecha      : 07/03/2005
' Ult. Mod   : 16/09/2008
' Fecha      : solo de los prestamos activos (que tengan cuotas que vencen en el periodo)
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String
Dim pos
Dim rs_Prestamo As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer

Dim idLiq As String
Dim NroPeriodo As String
Dim Monto As Double
Dim MontoAux As String
Dim HayInformacion As Boolean

On Error GoTo MError

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 251"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
'On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\Credits_AR.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If

'desactivo el manejador de errores
'On Error GoTo 0

HayInformacion = False
' Comienzo la transaccion
MyBeginTrans

StrSql = " SELECT DISTINCT empleg, prestamo.prenro, tipoprestamo.tpdesabr, estructura.estrdabr, prestamo.precantcuo, "
StrSql = StrSql & " prestamo.prefecotor, pre_cuota.cuofecvto, prestamo.preimp "
StrSql = StrSql & " FROM  prestamo "
StrSql = StrSql & " INNER JOIN pre_cuota ON pre_cuota.prenro = prestamo.prenro "
StrSql = StrSql & " INNER JOIN pre_linea ON pre_linea.lnprenro = prestamo.lnprenro "
StrSql = StrSql & " INNER JOIN tipoprestamo ON tipoprestamo.tpnro = pre_linea.tpnro "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = prestamo.ternro "
StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
StrSql = StrSql & " AND his_estructura.tenro  = 32 AND (his_estructura.htetdesde <= " & ConvFecha(FechaHasta) & ") "
StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(FechaDesde) & ") "
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE prestamo.prenro IN (SELECT prenro FROM pre_cuota where cuofecvto >= " & ConvFecha(FechaDesde) & " AND cuofecvto <= " & ConvFecha(FechaHasta) & ")"
OpenRecordset StrSql, rs_Prestamo

Cantidad = rs_Prestamo.RecordCount
cantidadProcesada = Cantidad

Dim Error As Boolean
If Cantidad = 0 Then Cantidad = 1
IncPorc = 99 / Cantidad
Progreso = 0
Do While Not rs_Prestamo.EOF
    HayInformacion = True
   Error = False
   
   If (rs_Prestamo!cuofecvto >= FechaDesde) And (rs_Prestamo!cuofecvto <= FechaHasta) Then
        'Monto = Fix(rs_Prestamo!preimp) & "." & Abs(Round((rs_Prestamo!preimp - Fix(rs_Prestamo!preimp)) * 100))
         pos = InStr(1, rs_Prestamo!preimp, ".")
         If pos = 0 Then
            MontoAux = 0
         Else
            MontoAux = Mid(rs_Prestamo!preimp, pos + 1, Len(rs_Prestamo!preimp) - pos)
         End If
         Monto = Fix(rs_Prestamo!preimp) & "." & MontoAux
        
        Cabecera = Format_StrNro(rs_Prestamo!empleg, 7, False, "") & Chr(9) & rs_Prestamo!prenro & Chr(9)
        Cabecera = Cabecera & rs_Prestamo!estrdabr & Chr(9) & rs_Prestamo!tpdesabr & Chr(9) & " " & Chr(9)
        Cabecera = Cabecera & rs_Prestamo!precantcuo & Chr(9) & Format(rs_Prestamo!prefecotor, "yyyy-MM-dd") & Chr(9)
        Cabecera = Cabecera & Format(rs_Prestamo!cuofecvto, "yyyy-MM-dd") & Chr(9) & Monto
             
        fExport.writeline Cabecera
            
        TiempoAcumulado = GetTickCount
               
        cantidadProcesada = cantidadProcesada - 1
               
        Progreso = Progreso + IncPorc
         StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                  ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                  ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
         objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
   rs_Prestamo.MoveNext
Loop

rs_Prestamo.Close
If Not HayInformacion Then
    fExport.writeline
End If
fExport.Close

MyCommitTrans

Set rs_Prestamo = Nothing

Exit Sub

MError:
    MyRollbackTrans
    Flog.writeline "Error: " & Err.Description
    
End Sub


Public Sub Generacion8(ByVal Exportacion As Integer, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal bpronro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de la Exportacion de Roche
' Autor      : FGZ
' Fecha      : 16/09/2008
' Ult. Mod   : exporta solo las cuotas liquidades en el periodo
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String
Dim pos
Dim rs_Prestamo As New ADODB.Recordset
Dim rs_Cuota As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

Dim Cabecera As String

Const ForReading = 1
Const TristateFalse = 0
Dim fExport
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim Aux_str As String
Dim TipoCodEmpresa As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer

Dim idLiq As String
Dim NroPeriodo As String
Dim MontoRestante As Double
Dim MontoPago As Double
Dim MontoTotal As Double
Dim MontoRestanteStr As String
Dim MontoPagoStr As String
Dim MontoTotalStr As String
Dim HayInformacion As Boolean
Dim MontoAux As String

On Error GoTo MError

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 251"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.writeline "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'Activo el manejador de errores
'On Error Resume Next

'Archivo para la cabecera del Pedido de Pago
Archivo = Directorio & "\CreditsDetails_AR.txt"
Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)

If Err.Number <> 0 Then
    Flog.writeline "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport = fs.CreateTextFile(Archivo, True)
End If

'desactivo el manejador de errores
'On Error GoTo 0

HayInformacion = False
' Comienzo la transaccion
MyBeginTrans

StrSql = " SELECT empleg, prestamo.prenro,  prestamo.prefecotor "
StrSql = StrSql & " FROM  prestamo "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = prestamo.ternro "
'StrSql = StrSql & " WHERE prestamo.prefecotor >= " & ConvFecha(FechaDesde) & " AND prestamo.prefecotor <= " & ConvFecha(FechaHasta)
StrSql = StrSql & " WHERE prestamo.prenro IN (SELECT prenro FROM pre_cuota where cuofecvto >= " & ConvFecha(FechaDesde) & " AND cuofecvto <= " & ConvFecha(FechaHasta) & ")"
StrSql = StrSql & " ORDER BY empleado.empleg "
OpenRecordset StrSql, rs_Prestamo

Cantidad = rs_Prestamo.RecordCount
cantidadProcesada = Cantidad
Dim Error As Boolean
If Cantidad = 0 Then Cantidad = 1
IncPorc = 99 / Cantidad
Progreso = 0
Do While Not rs_Prestamo.EOF
           
   Error = False
   
   StrSql = " SELECT cuocancela, cuototal, cuofecvto, cuonrocuo, cuosaldo "
   StrSql = StrSql & " FROM  pre_cuota "
   StrSql = StrSql & " WHERE pre_cuota.prenro = " & rs_Prestamo!prenro
   StrSql = StrSql & " AND cuofecvto >= " & ConvFecha(FechaDesde) & " AND cuofecvto <= " & ConvFecha(FechaHasta)
   StrSql = StrSql & " AND cuocancela = -1 "
   StrSql = StrSql & " ORDER BY pre_cuota.cuonrocuo "
   OpenRecordset StrSql, rs_Cuota

   MontoTotal = 0
   MontoPago = 0
   MontoRestante = 0
   
   Do While Not rs_Cuota.EOF
        HayInformacion = True
      If rs_Cuota!cuonrocuo = 1 Then
         MontoTotal = rs_Cuota!cuototal + rs_Cuota!cuosaldo
      End If
      
      If rs_Cuota!cuocancela = -1 And rs_Cuota!cuofecvto <= FechaHasta And rs_Cuota!cuofecvto >= FechaDesde Then
         MontoPago = rs_Cuota!cuototal
         MontoRestante = rs_Cuota!cuosaldo
      Else: MontoRestante = MontoRestante
            MontoPago = 0
      End If
      
      pos = InStr(1, MontoPago, ".")
      If pos = 0 Then
         MontoAux = 0
      Else
         MontoAux = Mid(MontoPago, pos + 1, Len(MontoPago) - pos)
      End If
      'MontoPagoStr = Fix(MontoPago) & "." & Round((MontoPago - Fix(MontoPago)) * 100)
      MontoPagoStr = Fix(MontoPago) & "." & MontoAux
      
      pos = InStr(1, MontoRestante, ".")
      If pos = 0 Then
         MontoAux = 0
      Else
         MontoAux = Mid(MontoRestante, pos + 1, Len(MontoRestante) - pos)
      End If
      'MontoRestanteStr = Fix(MontoRestante) & "." & Round((MontoRestante - Fix(MontoRestante)) * 100)
      MontoRestanteStr = Fix(MontoRestante) & "." & MontoAux
      
      pos = InStr(1, MontoTotalStr, ".")
      If pos = 0 Then
         MontoAux = 0
      Else
         MontoAux = Mid(MontoTotalStr, pos + 1, Len(MontoTotalStr) - pos)
      End If
      'MontoTotalStr = Fix(MontoTotal) & "." & Round((MontoTotal - Fix(MontoTotal)) * 100)
      MontoTotalStr = Fix(MontoTotal) & "." & MontoAux
      
      Cabecera = Format_StrNro(rs_Prestamo!empleg, 7, False, "") & Chr(9) & rs_Prestamo!prenro & Chr(9)
      Cabecera = Cabecera & rs_Cuota!cuonrocuo & Chr(9) & Format(rs_Cuota!cuofecvto, "yyyy-MM-dd") & Chr(9)
      Cabecera = Cabecera & MontoPagoStr & Chr(9) & MontoRestanteStr & Chr(9)
      Cabecera = Cabecera & MontoTotalStr
      
      fExport.writeline Cabecera
      
      rs_Cuota.MoveNext
      
   Loop
   rs_Cuota.Close
   
   TiempoAcumulado = GetTickCount
          
   cantidadProcesada = cantidadProcesada - 1
          
   Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
   rs_Prestamo.MoveNext
Loop

rs_Prestamo.Close
If Not HayInformacion Then
    fExport.writeline
End If
fExport.Close

MyCommitTrans

Set rs_Prestamo = Nothing
Set rs_Cuota = Nothing
Exit Sub

MError:
    MyRollbackTrans
    Flog.writeline "Error: " & Err.Description
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : JMH
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim ArrParametros
Dim Exportacion As Integer
Dim Anio As Integer
Dim mes As Integer
Dim Periodo As Integer
Dim Proceso As String

'Orden de los parametros
'Pedido de Ticket

ArrParametros = Split(parametros, "@")
' Levanto cada parametro por separado
Exportacion = ArrParametros(0)

Select Case Exportacion
   
   Case 1:
        Anio = ArrParametros(1)
        mes = ArrParametros(2)
        Call Generacion1(Exportacion, Anio, mes, bpronro)
   Case 2:
        Periodo = ArrParametros(1)
        Proceso = ArrParametros(2)
        Call Generacion2(Exportacion, Periodo, Proceso, bpronro)
    Case 3:
        Periodo = ArrParametros(1)
        Proceso = ArrParametros(2)
        Call Generacion3(Exportacion, Periodo, Proceso, bpronro)
    Case 4:
        Periodo = ArrParametros(1)
        Proceso = ArrParametros(2)
        Call Generacion4(Exportacion, Periodo, Proceso, bpronro)
    Case 5:  'Prestamos otorgados en el periodo
        Call Generacion5(Exportacion, Fecha_Desde, Fecha_Hasta, bpronro)
    Case 6:  'Detalle de cuotas de Prestamo completo
        Call Generacion6(Exportacion, Fecha_Desde, Fecha_Hasta, bpronro)
    Case 7:  'Prestamos activos en el periodo
        Call Generacion7(Exportacion, Fecha_Desde, Fecha_Hasta, bpronro)
    Case 8  'Detalle cuotas de Prestamo liquidadas en el periodo
        Call Generacion8(Exportacion, Fecha_Desde, Fecha_Hasta, bpronro)
End Select

End Sub


