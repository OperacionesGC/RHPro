Attribute VB_Name = "MdlCertSueldo"
Option Explicit

'Const Version = "1.0"
'Const FechaVersion = "16-08-2007"
'Autor = Diego Rosso
'--------------------------------------------------------------------------------------------------
Const Version = "1.1"
Const FechaVersion = "05-11-2008"
'Autor = Diego Nuñez

Global CantEmplError
Global CantEmplSinError
Global Errores As Boolean


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte de Certificados sobre Sueldos chile.
' Autor      : Diego Rosso
' Fecha      : 16-08-2007
' Ultima Mod.:  Diego Nicolás Nuñez
' Fecha:        04-11-2008
' Descripcion:  El factor de actualización se trae de la tabla Periodos. El mismo es ingresado manualmente
'               por el usuario al momento de generar periodos de liquidación.
'               Por otro lado, todos aquellos valores almacenados en la tabla rep_cert_sueldo_det, son tomados
'               en valor absoluto.
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
    
    Nombre_Arch = PathFLog & "Certi_sueldo" & "-" & NroProcesoBatch & ".log"
    
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 193 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call CertiSueldo(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "**********************************************************"
    Flog.writeline
    Flog.writeline "Cantidad de Empleados Insertados: " & CantEmplSinError
    Flog.writeline "Cantidad de Empleados Con ERRORES: " & CantEmplError
    Flog.writeline
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline
    Flog.writeline "**********************************************************"
    If Not Errores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100  WHERE bpronro = " & NroProcesoBatch
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
    'MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans
End Sub


Public Sub CertiSueldo(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte Certificado Sueldo
' Autor      : Diego Rosso
' Fecha      : 16-08-2007
' --------------------------------------------------------------------------------------------


Dim Empresa As Long
Dim Lista_Mod As String
Dim anio As Integer
Dim Legdesde As Long
Dim LegHasta As Long
Dim Empest As Byte

'Sueldo Bruto
Dim EsSueldoConc As Boolean
Dim SueldoConf As Long
Dim Sueldo As Double

'Cotizaciones Previsionales
Dim EsCotizConc As Boolean
Dim CotizConf As Long
Dim Cotizacion As Double

'Renta Total Exenta
Dim EsRentaConc As Boolean
Dim RentaConf As Long
Dim Renta As Double

'Impuesto Unico
Dim EsImpuestoConc As Boolean
Dim ImpuestoConf As Long
Dim Impuesto As Double

'Mayor Retencion Solicitada
Dim EsRetencionConc As Boolean
Dim RetencionConf As Long
Dim Retencion As Double

'Renta Total Exenta
Dim EsRenTotExentaConc As Boolean
Dim RenTotExentaConf As Long
Dim RenTotExenta As Double

'Rebajas
Dim EsRebajasConc As Boolean
Dim RebajasConf As Long
Dim Rebajas As Double

'Factor de Actualizacion
Dim Factor As Double  'PCO
Dim FactorConf As Long

Dim I      As Integer
Dim pos1 As Integer
Dim pos2 As Integer
Dim UltimoEmpleado As Long
Dim Apellido As String
Dim Apellido2 As String
Dim NombreEmp As String
Dim NombreEmp2 As String
Dim RUT As String
Dim DV As String
Dim Num_linea
Dim Titulo As String
Dim Z As Byte
Dim meses(12) As Byte   '0 Falso 1 verdadero   para saber si el impuesto unico se liquido ese mes
Dim PeriodoDesde
Dim PeriodoHasta

'Recordsets
Dim rs_Empleados As New ADODB.Recordset
Dim rs_CantEmpleados As New ADODB.Recordset
Dim rs_Acu_liq As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Rut As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim objRs3 As New ADODB.Recordset


 ' Inicio codigo ejecutable
'    On Error GoTo CE

' El formato de los parametros pasados es
'  (titulo del reporte, Todos_los_modelos, empresa,pliqnro_desde, pliqnro_hasta)

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline "Levantando Parametros  "
Flog.writeline Espacios(Tabulador * 1) & Parametros
Flog.writeline
If Not IsNull(Parametros) Then
    
    If Len(Parametros) >= 1 Then
        
        'TITULO
        '-----------------------------------------------------------
        pos1 = 1
        pos2 = InStr(pos1, Parametros, "@") - 1
        Titulo = Mid(Parametros, pos1, pos2)
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Titulo = " & Titulo
        Flog.writeline
        '-------------------------------------------------------------
        
        'MODELOS DE LIQUIDACION
        '-------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, "@") - 1
        Lista_Mod = Mid(Parametros, pos1, pos2 - pos1 + 1)
   
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Lista_Mod = " & Lista_Mod
        Flog.writeline
        ' esta lista tiene los nro de procesos separados por comas
        '-------------------------------------------------------------
        
        
        'Empresa
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, "@") - 1
        Empresa = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Empresa = " & Empresa
        Flog.writeline
        '------------------------------------------------------------------------------------
        
        'AÑO
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, "@") - 1
        anio = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro AÑO = " & anio
        Flog.writeline
        '------------------------------------------------------------------------------------
        
        'Numero de Legajo desde
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, "@") - 1
        Legdesde = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro legajo desde = " & Legdesde
        Flog.writeline
        '------------------------------------------------------------------------------------
        
         'Numero de Legajo Hasta
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, "@") - 1
        LegHasta = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro legajo hasta = " & LegHasta
        Flog.writeline
        '------------------------------------------------------------------------------------
        
        'Estado del empleado
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        Empest = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Empest = " & Empest
        Flog.writeline
        '------------------------------------------------------------------------------------
        
    End If
Else
    Flog.writeline "ERROR..No se encontraron parametros para el proceso"
    Exit Sub
End If
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Terminó de levantar los parametros "
Flog.writeline


'Configuracion del Reporte
Flog.writeline "Levantando configuracion del Reporte"
StrSql = "SELECT * FROM confrep WHERE repnro = 208 "
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'Levanto la configuracion para el reporte

'Inicializo
SueldoConf = 0
CotizConf = 0
RentaConf = 0
ImpuestoConf = 0
RetencionConf = 0
RenTotExentaConf = 0
RebajasConf = 0
FactorConf = 0

'Sueldo Bruto
StrSql = "SELECT * FROM confrep WHERE repnro = 208 and confnrocol = 1 "
OpenRecordset StrSql, rs_Confrep
Flog.writeline "Levantando configuracion de columna 1"
If Not rs_Confrep.EOF Then
      If UCase(rs_Confrep!conftipo) = "CO" Then
           EsSueldoConc = True
                  
           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
               If Not EsNulo(rs_Confrep!confval2) Then
                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
               End If
           OpenRecordset StrSql, objRs3
           If Not objRs3.EOF Then
             SueldoConf = objRs3!concnro
           End If
           objRs3.Close
      Else
           EsSueldoConc = False
           SueldoConf = rs_Confrep!confval
      End If
      Flog.writeline "Se obtuvo la configuracion de columna 1"
Else
      Flog.writeline "ERROR no se configuro correctamente la columna 1"
End If

'Cotizaciones Previsionales
StrSql = "SELECT * FROM confrep WHERE repnro = 208 and confnrocol = 2 "
OpenRecordset StrSql, rs_Confrep
Flog.writeline "Levantando configuracion de columna 2"
If Not rs_Confrep.EOF Then
      If UCase(rs_Confrep!conftipo) = "CO" Then
           EsCotizConc = True
                  
           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
               If Not EsNulo(rs_Confrep!confval2) Then
                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
               End If
           OpenRecordset StrSql, objRs3
           If Not objRs3.EOF Then
             CotizConf = objRs3!concnro
           End If
           objRs3.Close
      Else
           EsCotizConc = False
           CotizConf = rs_Confrep!confval
      End If
      Flog.writeline "Se obtuvo la configuracion de columna 2"
Else
      Flog.writeline "ERROR no se configuro correctamente la columna 2"
End If


'Renta Total Exenta
StrSql = "SELECT * FROM confrep WHERE repnro = 208 and confnrocol = 3 "
OpenRecordset StrSql, rs_Confrep
Flog.writeline "Levantando configuracion de columna 3"
If Not rs_Confrep.EOF Then
      If UCase(rs_Confrep!conftipo) = "CO" Then
           EsRentaConc = True
                  
           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
               If Not EsNulo(rs_Confrep!confval2) Then
                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
               End If
           OpenRecordset StrSql, objRs3
           If Not objRs3.EOF Then
             RentaConf = objRs3!concnro
           End If
           objRs3.Close
      Else
           EsRentaConc = False
           RentaConf = rs_Confrep!confval
      End If
      Flog.writeline "Se obtuvo la configuracion de columna 3"
Else
      Flog.writeline "ERROR no se configuro correctamente la columna 3"
End If

'Impuesto Unico
StrSql = "SELECT * FROM confrep WHERE repnro = 208 and confnrocol = 4 "
OpenRecordset StrSql, rs_Confrep
Flog.writeline "Levantando configuracion de columna 4"
If Not rs_Confrep.EOF Then
      If UCase(rs_Confrep!conftipo) = "CO" Then
           EsImpuestoConc = True
                  
           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
               If Not EsNulo(rs_Confrep!confval2) Then
                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
               End If
           OpenRecordset StrSql, objRs3
           If Not objRs3.EOF Then
             ImpuestoConf = objRs3!concnro
           End If
           objRs3.Close
      Else
           EsImpuestoConc = False
           ImpuestoConf = rs_Confrep!confval
      End If
      Flog.writeline "Se obtuvo la configuracion de columna 4"
Else
      Flog.writeline "ERROR no se configuro correctamente la columna 4"
End If


'Mayor Retencion Solicitada
StrSql = "SELECT * FROM confrep WHERE repnro = 208 and confnrocol = 5 "
OpenRecordset StrSql, rs_Confrep
Flog.writeline "Levantando configuracion de columna 5"
If Not rs_Confrep.EOF Then
      If UCase(rs_Confrep!conftipo) = "CO" Then
           EsRetencionConc = True
                  
           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
               If Not EsNulo(rs_Confrep!confval2) Then
                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
               End If
           OpenRecordset StrSql, objRs3
           If Not objRs3.EOF Then
             RetencionConf = objRs3!concnro
           End If
           objRs3.Close
      Else
           EsRetencionConc = False
           RetencionConf = rs_Confrep!confval
      End If
      Flog.writeline "Se obtuvo la configuracion de columna 5"
Else
      Flog.writeline "ERROR no se configuro correctamente la columna 5"
End If


'Renta Total Exenta
StrSql = "SELECT * FROM confrep WHERE repnro = 208 and confnrocol = 6 "
OpenRecordset StrSql, rs_Confrep
Flog.writeline "Levantando configuracion de columna 6"
If Not rs_Confrep.EOF Then
      If UCase(rs_Confrep!conftipo) = "CO" Then
           EsRenTotExentaConc = True
                  
           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
               If Not EsNulo(rs_Confrep!confval2) Then
                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
               End If
           OpenRecordset StrSql, objRs3
           If Not objRs3.EOF Then
             RenTotExentaConf = objRs3!concnro
           End If
           objRs3.Close
      Else
           EsRenTotExentaConc = False
           RenTotExentaConf = rs_Confrep!confval
      End If
      Flog.writeline "Se obtuvo la configuracion de columna 6"
Else
      Flog.writeline "ERROR no se configuro correctamente la columna 6"
End If


'Rebajas
StrSql = "SELECT * FROM confrep WHERE repnro = 208 and confnrocol = 7 "
OpenRecordset StrSql, rs_Confrep
Flog.writeline "Levantando configuracion de columna 7"
If Not rs_Confrep.EOF Then
      If UCase(rs_Confrep!conftipo) = "CO" Then
           EsRebajasConc = True
                  
           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
               If Not EsNulo(rs_Confrep!confval2) Then
                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
               End If
           OpenRecordset StrSql, objRs3
           If Not objRs3.EOF Then
             RebajasConf = objRs3!concnro
           End If
           objRs3.Close
      Else
           EsRebajasConc = False
           RebajasConf = rs_Confrep!confval
      End If
      Flog.writeline "Se obtuvo la configuracion de columna 7"
Else
      Flog.writeline "ERROR no se configuro correctamente la columna 7"
End If
  
'Factor de Actualizacion
StrSql = "SELECT * FROM confrep WHERE repnro = 208 and confnrocol = 8 "
OpenRecordset StrSql, rs_Confrep
Flog.writeline "Levantando configuracion de columna 8"
If Not rs_Confrep.EOF Then
      If UCase(rs_Confrep!conftipo) = "PCO" Then
           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
               If Not EsNulo(rs_Confrep!confval2) Then
                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
               End If
           OpenRecordset StrSql, objRs3
           If Not objRs3.EOF Then
             FactorConf = objRs3!concnro
           End If
           objRs3.Close
      Else
           Flog.writeline "ERROR no se configuro correctamente la columna 8. Debe ser del tipo PCO"
      End If
      Flog.writeline "Se obtuvo la configuracion de columna 8"
Else
      Flog.writeline "ERROR no se configuro correctamente la columna 8"
End If

Flog.writeline "Se obtuvo la configuracion del Reporte"


'----------------------------------------
UltimoEmpleado = -1
Num_linea = 0
PeriodoDesde = "01/" & "01/" & anio
PeriodoHasta = "31/" & "12/" & anio

    StrSql = "SELECT distinct(empleado.ternro),empleado.empleg, cabliq.empleado FROM proceso "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " INNER JOIN  tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
    StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & Empresa
    StrSql = StrSql & " WHERE "
    If Lista_Mod <> "0" Then
        StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
    End If
    StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
    StrSql = StrSql & " AND empleado.empleg >= " & Legdesde
    StrSql = StrSql & " AND empleado.empleg <= " & LegHasta
    If Empest <> 2 Then
        StrSql = StrSql & " AND empleado.empest= " & Empest
    End If
    StrSql = StrSql & " ORDER BY empleado.ternro"
    OpenRecordset StrSql, rs_Empleados
    
    
    If rs_Empleados.State = adStateOpen Then
        Flog.writeline "Busco los empleados"
    Else
        Flog.writeline "Se supero el tiempo de espera "
        HuboError = True
    End If
    
If Not HuboError Then
        
    
        'seteo de las variables de progreso
        Progreso = 0
          
        'Cantidad de empleados
        CEmpleadosAProc = rs_Empleados.RecordCount
        
        If CEmpleadosAProc = 0 Then
           Flog.writeline ""
           Flog.writeline "NO hay empleados"
           Exit Sub
           CEmpleadosAProc = 1
        End If
        
        IncPorc = (99 / CEmpleadosAProc)
        Flog.writeline
        Flog.writeline
        
        'Inicializo la cantidad de empleados con errores a 0
        CantEmplError = 0
        CantEmplSinError = 0
        
'        RentaPagada = 0
'        RentaPagadaAnio = 0
'        RentAccEneAbr = 0
'        RentaGrabada = 0
'        RebajasTot = 0
'        TotalRemu = 0
    Do While Not rs_Empleados.EOF
        
          If rs_Empleados!ternro <> UltimoEmpleado Then  'Es el primero
                    
               UltimoEmpleado = rs_Empleados!ternro
               Flog.writeline "_______________________________________________________________________"
                 
                
                'Buscar el apellido y nombre
                    StrSql = "SELECT * FROM tercero WHERE ternro = " & rs_Empleados!ternro
                    OpenRecordset StrSql, rs_Tercero
                    If Not rs_Tercero.EOF Then
                    
                        If EsNulo(rs_Tercero!terape) Then Apellido = "" Else Apellido = Left(rs_Tercero!terape, 50)
                        If EsNulo(rs_Tercero!terape2) Then Apellido2 = "" Else Apellido2 = Left(rs_Tercero!terape2, 50)
                        If EsNulo(rs_Tercero!ternom) Then NombreEmp = "" Else NombreEmp = Left(rs_Tercero!ternom, 50)
                        If EsNulo(rs_Tercero!ternom2) Then NombreEmp2 = "" Else NombreEmp2 = Left(rs_Tercero!ternom2, 50)
                    
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "ERROR al obtener Apellido o Nombre del Empleado"
                        Exit Sub
                    End If
                    Flog.writeline
                    Flog.writeline "Empleado: ------------------->" & rs_Empleados!empleg & "  " & Apellido & "  " & NombreEmp
                    Flog.writeline
          End If
                            
               'Reviso si es el ultimo empleado
                If EsElUltimoEmpleado(rs_Empleados, UltimoEmpleado) Then
                    
                    'Inicializo
                        HuboError = False 'Para cada empleado
                        Errores = False 'En el proceso
                        
                                                   
                                        
                    ' ----------------------------------------------------------------
                    ' Buscar el Rut DEL EMPLEADO
                    Flog.writeline
                    Flog.writeline "Obteniendo el RUT y DV del empleado. "
                    StrSql = " SELECT nrodoc FROM tercero " & _
                             " INNER JOIN ter_doc  ON (tercero.ternro = ter_doc.ternro AND ter_doc.tidnro = 1) " & _
                             " WHERE tercero.ternro= " & rs_Empleados!ternro
                    OpenRecordset StrSql, rs_Rut
          
                    If Not rs_Rut.EOF Then
                        RUT = Mid(rs_Rut!nrodoc, 1, Len(rs_Rut!nrodoc) - 1)
                        RUT = Replace(RUT, "-", "")
                        DV = Right(rs_Rut!nrodoc, 1)
                        Flog.writeline "RUT y DV obtenidos"
                    Else
                        Flog.writeline "Error al obtener los datos del RUT"
                        RUT = ""
                        DV = ""
                        HuboError = True
                    End If
                    Flog.writeline
              
                    '*****************************************************************************************
                    'Grabar cabecera
                    '*****************************************************************************************
                     Flog.writeline "----------------------------------------"
                     Flog.writeline "Grabando Cabecera"
                    
                    'Inserto en rep_cert_sueldo
                    StrSql = "INSERT INTO rep_cert_sueldo (bpronro, ternro, orden, Titulo, empnro, Anio, rut, DV, empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ") VALUES ("
                    StrSql = StrSql & NroProcesoBatch & ","
                    StrSql = StrSql & rs_Empleados!ternro & ","
                    StrSql = StrSql & Num_linea & ","
                    StrSql = StrSql & "'" & Titulo & "',"
                    StrSql = StrSql & Empresa & ","
                    StrSql = StrSql & anio & ","
                    StrSql = StrSql & "'" & RUT & "',"
                    StrSql = StrSql & "'" & DV & "',"
                    StrSql = StrSql & rs_Empleados!empleg & ","
                    StrSql = StrSql & "'" & Apellido & "',"
                    StrSql = StrSql & "'" & Apellido2 & "',"
                    StrSql = StrSql & "'" & NombreEmp & "',"
                    StrSql = StrSql & "'" & NombreEmp2 & "'"
                    StrSql = StrSql & ")"
                    Flog.writeline
                    Flog.writeline "Insertando : " & StrSql
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline "Se grabo el registro"
                    Flog.writeline
                    'Sumo el numero de linea
                    Num_linea = Num_linea + 1
              
                    Flog.writeline "Se buscan los datos del periodo del empleado"
                    
                        'Busco todos los periodos que tiene liquidados el empleado en el año
                        StrSql = "SELECT distinct (periodo.pliqmes), periodo.pliqdesc,periodo.pliqnro,periodo.pliqipc FROM proceso "
                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & Empresa
                        StrSql = StrSql & " WHERE "
                         If Lista_Mod <> "0" Then
                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                        End If
                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                        StrSql = StrSql & " ORDER BY pliqmes asc "
                        OpenRecordset StrSql, rs_Periodo
                    
                        Do While Not rs_Periodo.EOF
                            'Inicializo variables
                            Sueldo = 0
                            Cotizacion = 0
                            Impuesto = 0
                            Renta = 0
                            Retencion = 0
                            RenTotExenta = 0
                            Rebajas = 0
                            Factor = 1
                            
                            'SUELDO
                            Flog.writeline "Obteniendo el sueldo para el empleado " & rs_Empleados!empleg & " Periodo: " & rs_Periodo!pliqmes
                            If EsSueldoConc Then
                                StrSql = "SELECT dlimonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND detliq.concnro = " & SueldoConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Sueldo = Sueldo + rs_Detliq!dlimonto
                                   rs_Detliq.MoveNext
                                Loop
                            Else
                                StrSql = "SELECT almonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND acu_liq.acunro = " & SueldoConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Sueldo = Sueldo + rs_Detliq!almonto
                                   rs_Detliq.MoveNext
                                Loop
                            End If
                            Flog.writeline "Se Obtuvo el sueldo bruto"
                            
                            'Cotizaciones Previsionales
                            Flog.writeline "Obteniendo las Cotizaciones Previsionales para el empleado " & rs_Empleados!empleg & " Periodo: " & rs_Periodo!pliqmes
                            If EsCotizConc Then
                                StrSql = "SELECT dlimonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND detliq.concnro = " & CotizConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Cotizacion = Cotizacion + rs_Detliq!dlimonto
                                   rs_Detliq.MoveNext
                                Loop
                            Else
                                StrSql = "SELECT almonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND acu_liq.acunro = " & CotizConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Cotizacion = Cotizacion + rs_Detliq!almonto
                                   rs_Detliq.MoveNext
                                Loop
                            End If
                            Flog.writeline "Se Obtuvo las Cotizaciones Previsionales "
                            
                            'Renta Total Exenta
                            Flog.writeline "Obteniendo la renta que afecta el impuesto unico para el empleado " & rs_Empleados!empleg & " Periodo: " & rs_Periodo!pliqmes
                            If EsRentaConc Then
                                StrSql = "SELECT dlimonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND detliq.concnro = " & RentaConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Renta = Renta + rs_Detliq!dlimonto
                                   rs_Detliq.MoveNext
                                Loop
                            Else
                                StrSql = "SELECT almonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND acu_liq.acunro = " & RentaConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Renta = Renta + rs_Detliq!almonto
                                   rs_Detliq.MoveNext
                                Loop
                            End If
                            Flog.writeline "Se Obtuvo la renta que afecta el impuesto unico  "
                                                        
                            'Impuesto Unico
                            Flog.writeline "Obteniendo el impuesto unico para el empleado " & rs_Empleados!empleg & " Periodo: " & rs_Periodo!pliqmes
                            If EsImpuestoConc Then
                                StrSql = "SELECT dlimonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND detliq.concnro = " & ImpuestoConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Impuesto = Impuesto + rs_Detliq!dlimonto
                                   rs_Detliq.MoveNext
                                Loop
                            Else
                                StrSql = "SELECT almonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND acu_liq.acunro = " & ImpuestoConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Impuesto = Impuesto + rs_Detliq!almonto
                                   rs_Detliq.MoveNext
                                Loop
                            End If
                            Flog.writeline "Se Obtuvo el impuesto unico  "
                            
                            'Mayor Retencion Solicitada
                            Flog.writeline "Obteniendo la Mayor Retencion Solicitada para el empleado " & rs_Empleados!empleg & " Periodo: " & rs_Periodo!pliqmes
                            If EsRetencionConc Then
                                StrSql = "SELECT dlimonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND detliq.concnro = " & RetencionConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Retencion = Retencion + rs_Detliq!dlimonto
                                   rs_Detliq.MoveNext
                                Loop
                            Else
                                StrSql = "SELECT almonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND acu_liq.acunro = " & RetencionConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Retencion = Retencion + rs_Detliq!almonto
                                   rs_Detliq.MoveNext
                                Loop
                            End If
                            Flog.writeline "Se Obtuvo la Mayor Retencion Solicitada "

                            'Renta Total Exenta
                            Flog.writeline "Obteniendo la Renta Total Exenta para el empleado " & rs_Empleados!empleg & " Periodo: " & rs_Periodo!pliqmes
                            If EsRenTotExentaConc Then
                                StrSql = "SELECT dlimonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND detliq.concnro = " & RenTotExentaConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   RenTotExenta = RenTotExenta + rs_Detliq!dlimonto
                                   rs_Detliq.MoveNext
                                Loop
                            Else
                                StrSql = "SELECT almonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND acu_liq.acunro = " & RenTotExentaConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   RenTotExenta = RenTotExenta + rs_Detliq!almonto
                                   rs_Detliq.MoveNext
                                Loop
                            End If
                            Flog.writeline "Se Obtuvo la Renta Total Exenta "

                            'Rebajas
                            Flog.writeline "Obteniendo las rebajas por zonas extremas para el empleado " & rs_Empleados!empleg & " Periodo: " & rs_Periodo!pliqmes
                            If EsRebajasConc Then
                                StrSql = "SELECT dlimonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND detliq.concnro = " & RebajasConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Rebajas = Rebajas + rs_Detliq!dlimonto
                                   rs_Detliq.MoveNext
                                Loop
                            Else
                                StrSql = "SELECT almonto FROM proceso "
                                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & Empresa
                                StrSql = StrSql & " WHERE "
                                 If Lista_Mod <> "0" Then
                                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                                End If
                                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                                StrSql = StrSql & " AND acu_liq.acunro = " & RebajasConf
                                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                                StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                                OpenRecordset StrSql, rs_Detliq
                                Do While Not rs_Detliq.EOF
                                   Rebajas = Rebajas + rs_Detliq!almonto
                                   rs_Detliq.MoveNext
                                Loop
                            End If
                            Flog.writeline "Se Obtuvo las rebajas por zonas extremas"
                            
                            'Factor de actualizacion
                            Flog.writeline "Obteniendo las rebajas por zonas extremas para el empleado " & rs_Empleados!empleg & " Periodo: " & rs_Periodo!pliqmes
                            StrSql = "SELECT dlicant FROM proceso "
                            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                            StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                            StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
                            StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
                            StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & Empresa
                            StrSql = StrSql & " WHERE "
                             If Lista_Mod <> "0" Then
                                StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
                            End If
                            StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
                            StrSql = StrSql & " AND detliq.concnro = " & FactorConf
                            StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                            StrSql = StrSql & " AND periodo.pliqmes=" & rs_Periodo!pliqmes
                            OpenRecordset StrSql, rs_Detliq
                            Do While Not rs_Detliq.EOF
                               Factor = rs_Detliq!dlicant
                               rs_Detliq.MoveNext
                            Loop
                            
                            
                        'Inserto en rep_ddjj_renta_det
                         StrSql = "INSERT INTO rep_cert_sueldo_det (bpronro, ternro, pliqnro, pliqdesc, pliqmes, SueldoBruto, Cotizaciones, RentaImpUni, ImpuestoUni, MayRetencion, RentaExenta, Rebajas, FactorAct "
                         StrSql = StrSql & ") VALUES ("
                         StrSql = StrSql & NroProcesoBatch & ","
                         StrSql = StrSql & rs_Empleados!ternro & ","
                         StrSql = StrSql & rs_Periodo!pliqnro & ","
                         StrSql = StrSql & "'" & Left(rs_Periodo!pliqdesc, 200) & "',"
                         StrSql = StrSql & rs_Periodo!pliqmes & ","
                         StrSql = StrSql & Abs(Sueldo) & ","
                         StrSql = StrSql & Abs(Cotizacion) & ","
                         StrSql = StrSql & Abs(Renta) & ","
                         StrSql = StrSql & Abs(Impuesto) & ","
                         StrSql = StrSql & Abs(Retencion) & ","
                         StrSql = StrSql & Abs(RenTotExenta) & ","
                         StrSql = StrSql & Abs(Rebajas) & ","
'                         StrSql = StrSql & Factor
                         StrSql = StrSql & rs_Periodo!pliqipc
                         StrSql = StrSql & ")"
                         Flog.writeline
                         Flog.writeline "Insertando : " & StrSql
                         objConn.Execute StrSql, , adExecuteNoRecords
                         Flog.writeline
                         Flog.writeline
                    
                            'Paso al siguiente periodo
                            rs_Periodo.MoveNext
                            
                        Loop ' rs_Periodo.EOF
                    
                    
                  
                  '-----------------------------------------------------------------------------------
                'Controlo errores en el empleado
                If Not HuboError Then
                    
                    CantEmplSinError = CantEmplSinError + 1
                    Flog.writeline
                    Flog.writeline "SE GRABO EL EMPLEADO "
                    Flog.writeline
                Else
                    
                    'Sumo 1 A la cantidad de errores
                    CantEmplError = CantEmplError + 1
                    Flog.writeline
                    Flog.writeline "SE DETECTARON ERRORES EN EL EMPLEADO "
                    Flog.writeline
                    Errores = True
                End If
                    
                    'Actualizo el progreso
                      Progreso = Progreso + IncPorc
                      TiempoAcumulado = GetTickCount
                    
                    If Errores = False Then
                      StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                       ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                       "' WHERE bpronro = " & NroProcesoBatch
                      objconnProgreso.Execute StrSql, , adExecuteNoRecords
                    Else
                      StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                       ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                       "',bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
                       objconnProgreso.Execute StrSql, , adExecuteNoRecords
                    End If
           
                    ' ----------------------------------------------------------------
                
                End If
                
                'Paso al siguiente Empleado
                rs_Empleados.MoveNext
            Loop
            
End If 'If Not HuboError


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_CantEmpleados.State = adStateOpen Then rs_CantEmpleados.Close
If rs_Acu_liq.State = adStateOpen Then rs_Acu_liq.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Rut.State = adStateOpen Then rs_Rut.Close
If objRs3.State = adStateOpen Then objRs3.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close

Set rs_Empleados = Nothing
Set rs_CantEmpleados = Nothing
Set rs_Acu_liq = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Tercero = Nothing
Set rs_Rut = Nothing
Set objRs3 = Nothing
Set rs_Periodo = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    'MyRollbackTrans
    'MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description


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


