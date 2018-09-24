Attribute VB_Name = "MdlDifAsientoContable"
Option Explicit

Global Const Version = "1.00"  'Juan A. Zamarbide - Diferencia entre Asientos Contables
Global Const FechaVersion = "29/02/2012" ' Custom Santander Chile Caso CAS-14650 - Santander Chile - Asiento Contable

'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------

Global rs_Empleado As New ADODB.Recordset
Global rs_Mod_Asiento As New ADODB.Recordset

Global Vol1 As Long
Global Vol2 As Long
Global CatidadVueltas As Long
Global Corte As Boolean
Global DispDet As String
Global NroVol As Long
Global vol_Fec_Asiento As Date




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Asientos Contables.
' Autor      : Martin Ferraro
' Fecha      : 07/08/2006
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
    
    
    Nombre_Arch = PathFLog & "Asiento_Contable" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline

    TiempoInicialProceso = GetTickCount
    
    On Error Resume Next
    'Abro la conexion
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
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "'"
    StrSql = StrSql & " , bprcfecinicioej = " & ConvFecha(Date)
    StrSql = StrSql & " , bprcestado = 'Procesando'"
    StrSql = StrSql & " , bprcpid = " & PID
    StrSql = StrSql & " , bprctiempo = 0 "
    StrSql = StrSql & " , bprcprogreso = 0 "
    StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 318 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call GenerarAsiento(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "NO SE ENCONTRO EL PROCESO " & NroProcesoBatch
    End If
    
    TiempoFinalProceso = GetTickCount
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords


Fin:
    Flog.Close
    If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error : " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
End Sub



Public Sub GenerarAsiento(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Programa que se ejecuta para generar Asiento Contable
'              Configurado en el tipo de proceso batch
' Autor      : Martin Ferraro
' Fecha      : 24/12/2006
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer

Dim total As Long
Dim NroAsientos As Long
Dim NroLineas As Long
Dim NroAsi As Long
Dim NroLin As Long
Dim Monto As Double
Dim montacum As Double
Dim dh As Integer
Dim canti As Double
Dim montoDebe As Double
Dim montoHaber As Double


Dim rs_ProcVol As New ADODB.Recordset
Dim rs_Proc_V_modasi As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_ProcDet As New ADODB.Recordset
Dim rs_asiento As New ADODB.Recordset

On Error GoTo ME_GenerarAsiento


' Levanto cada parametro por separado, el separador de parametros es "."
Flog.writeline Espacios(Tabulador * 0) & "Inicio del proceso de volcado."
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        NroVol = CLng(Mid(Parametros, pos1, pos2))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Vol1 = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Vol2 = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        DispDet = CStr(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        
                
    End If
End If

Flog.writeline Espacios(Tabulador * 0) & "Parametros: "
Flog.writeline Espacios(Tabulador * 0) & "            Numero de Proceso = " & NroVol
Flog.writeline Espacios(Tabulador * 0) & "            Analisis Detallado = " & HACE_TRAZA
'Flog.writeline Espacios(Tabulador * 0) & "            Corte Desbalance = " & corteDesbalance
Flog.writeline
Flog.writeline

Flog.writeline Espacios(Tabulador * 0) & "Buscando el proceso de volcado 1."
'Buscando el proceso de volcado
StrSql = "SELECT * FROM proc_vol WHERE proc_vol.vol_cod =" & Vol1
OpenRecordset StrSql, rs_ProcVol

If rs_ProcVol.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. Proceso de Volcado 1 no fué encontrado."
    Exit Sub
End If

Flog.writeline Espacios(Tabulador * 0) & "Buscando el proceso de volcado 2."
'Buscando el proceso de volcado
StrSql = "SELECT * FROM proc_vol WHERE proc_vol.vol_cod = " & Vol2
OpenRecordset StrSql, rs_ProcVol

If rs_ProcVol.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. Proceso de Volcado 2 no fué encontrado."
    Exit Sub
End If

Flog.writeline Espacios(Tabulador * 0) & "Buscando el nuevo proceso de volcado."
'Buscando el proceso de volcado
StrSql = "SELECT * FROM proc_vol WHERE proc_vol.vol_cod =" & NroVol
OpenRecordset StrSql, rs_ProcVol

If rs_ProcVol.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. El Nuevo Proceso de Volcado no encontrado."
    Exit Sub
End If


Flog.writeline Espacios(Tabulador * 0) & "Buscando modelos proceso de volcado."
'Buscando los modelos asociados al proceso
StrSql = "SELECT * FROM proc_v_modasi WHERE proc_v_modasi.vol_cod =" & NroVol
StrSql = StrSql & " ORDER BY asi_cod "
OpenRecordset StrSql, rs_Proc_V_modasi

If rs_Proc_V_modasi.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. El proceso de volcado no tiene modelos asociados."
    Exit Sub
End If

'seteo las variables iniciales
montoDebe = 0
montoHaber = 0

'seteo las variables de progreso
Progreso = 0
CatidadVueltas = rs_Proc_V_modasi.RecordCount

'variable global de fecha de asiento
vol_Fec_Asiento = rs_ProcVol!vol_Fec_Asiento


Flog.writeline Espacios(Tabulador * 0) & "Cantidad de modelos del proceso de volcado = " & CatidadVueltas
'Flog.writeline Espacios(Tabulador * 0) & "Cantidad de cabeceras a procesar del proceso de volcado = " & CantidadEmpleados


'Verifico que las tablas auxiliares no existan previamente, si existen, las elimino

StrSql = "IF EXISTS(SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'asiento1')" & _
          " DROP TABLE asiento1"
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "IF EXISTS(SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'asiento2')" & _
          " DROP TABLE asiento2"
objConn.Execute StrSql, , adExecuteNoRecords

'Preparar Consulta Asiento 1

StrSql = " SELECT linea_asi.linea, linea_asi.cuenta, linea_asi.desclinea, linea_asi.dh, linea_asi.monto " & _
            " INTO asiento1 FROM linea_asi " & _
            " WHERE linea_asi.vol_cod = " & Vol1
            
objConn.Execute StrSql, , adExecuteNoRecords

'Preparar Consulta Asiento 2

StrSql = " SELECT linea_asi.linea, linea_asi.cuenta, linea_asi.desclinea, linea_asi.dh, linea_asi.monto " & _
            " INTO asiento2 FROM linea_asi " & _
            " WHERE linea_asi.vol_cod = " & Vol2
            
objConn.Execute StrSql, , adExecuteNoRecords

'Preparar Consulta JOIN Linea

StrSql = " SELECT DISTINCT asiento1.cuenta cuenta1, asiento1.desclinea dlinea1, asiento1.dh dh1, asiento1.monto monto1, asiento1.linea linea1," & _
         " asiento2.cuenta cuenta2, asiento2.desclinea dlinea2, asiento2.dh dh2, asiento2.monto monto2, asiento2.linea linea2" & _
         " FROM asiento1 " & _
         " FULL OUTER JOIN asiento2 ON asiento1.cuenta = asiento2.cuenta "
         
OpenRecordset StrSql, rs_ProcVol

'Preparo Consulta JOIN Detalle Asiento

StrSql = "SELECT DISTINCT da1.cuenta c1, da2.cuenta c2, da1.dlmonto m1, da2.dlmonto m2, da1.detasinro detasinro1, da2.detasinro detasinro2," & _
         "da1.dlcantidad dlcantidad1, da2.dlcantidad dlcantidad2, da1.dlcosto1 dlcosto11, da2.dlcosto1 dlcosto12, da1.dlcosto2 dlcosto21," & _
         "da2.dlcosto2 dlcosto22, da1.dlcosto3 dlcosto31, da2.dlcosto3 dlcosto32, da1.dlcosto4 dlcosto41, da2.dlcosto4 dlcosto42," & _
         "da1.dldescripcion dldescripcion1, da2.dldescripcion dldescripcion2, da1.dlmontoacum dlmontoacum1, da2.dlmontoacum dlmontoacum2," & _
         "da1.dlporcentaje dlporcentaje1, da2.dlporcentaje dlporcentaje2, da1.empleg empleg1, da2.empleg empleg2, da1.lin_orden lin_orden1," & _
         "da2.lin_orden lin_orden2, da1.linaD_H linaD_H1, da2.linaD_H linaD_H2, da1.linadesc linadesc1, da2.linadesc linadesc2, da1.masinro masinro1," & _
         "da2.masinro masinro2, da1.origen origen1, da2.origen origen2, da1.terape terape1,da2.terape terape2, da1.ternro ternro1, da2.ternro ternro2," & _
         "da1.tipoorigen tipoorigen1, da2.tipoorigen tipoorigen2, da1.vol_cod vol_cod1, da2.vol_cod vol_cod2" & _
         " FROM detalle_asi da1 " & _
         " LEFT OUTER JOIN detalle_asi da2 ON da1.cuenta = da2.cuenta AND da1.origen = da2.origen AND da1.ternro = da2.ternro AND da1.origen = da2.origen " & _
         " WHERE da1.vol_cod = " & Vol1 & " AND da2.vol_cod = " & Vol2

Flog.writeline StrSql
OpenRecordset StrSql, rs_ProcDet


  
'Armar asiento
CatidadVueltas = rs_ProcVol.RecordCount + rs_ProcDet.RecordCount

IncPorc = 99 / CatidadVueltas

'Por cada Linea del Proceso de volcado genero una nueva del asiento nuevo
Do While Not rs_ProcVol.EOF
 'Haber = 0
 'Debe = -1
     If rs_ProcVol!cuenta1 <> "" Then
        If rs_ProcVol!cuenta2 <> "" Then
            'Generar Linea Diferencia
            Monto = rs_ProcVol!monto1 - rs_ProcVol!monto2  ' Verificar resta con respecto a el D o H
            If rs_ProcVol!dh1 = rs_ProcVol!dh2 Then
                If rs_ProcVol!dh1 = -1 Then
                    If rs_ProcVol!monto1 > rs_ProcVol!monto2 Then
                        dh = 0
                        montoHaber = montoHaber + Abs(Monto)
                    Else
                        dh = -1
                        montoDebe = montoDebe + Abs(Monto)
                    End If
                    'Flog.writeline "Genera Linea con la diferencia de montos, cuenta, etc"
                    If Monto <> 0 Then
                        Call GuardarLineaAsi(NroVol, rs_Proc_V_modasi!asi_cod, rs_ProcVol!linea1, Abs(Monto), rs_ProcVol!cuenta1, rs_ProcVol!dlinea1, dh)
                    End If
                Else
                    If rs_ProcVol!monto1 > rs_ProcVol!monto2 Then
                        dh = -1
                        montoDebe = montoDebe + Abs(Monto)
                    Else
                        dh = 0
                        montoHaber = montoHaber + Abs(Monto)
                    End If
                    If Monto <> 0 Then
                        Call GuardarLineaAsi(NroVol, rs_Proc_V_modasi!asi_cod, rs_ProcVol!linea1, Abs(Monto), rs_ProcVol!cuenta1, rs_ProcVol!dlinea1, dh)
                    End If
                End If
            Else
                If Monto > 0 Then
                    dh = 0
                    montoHaber = montoHaber + Abs(Monto)
                Else
                    dh = -1
                    montoDebe = montoDebe + Abs(Monto)
                End If
                If Monto <> 0 Then
                    Call GuardarLineaAsi(NroVol, rs_Proc_V_modasi!asi_cod, rs_ProcVol!linea1, Abs(Monto), rs_ProcVol!cuenta1, rs_ProcVol!dlinea1, dh)
                End If
            End If
        Else
            'Flog.writeline "Genera Linea cambiando debe x haber o viceversa"
            If rs_ProcVol!dh1 Then
                dh = 0
                montoHaber = montoHaber + Abs(rs_ProcVol!monto1)
            Else
                dh = -1
                montoDebe = montoDebe + Abs(rs_ProcVol!monto1)
            End If
            Call GuardarLineaAsi(NroVol, rs_Proc_V_modasi!asi_cod, rs_ProcVol!linea1, Abs(rs_ProcVol!monto1), rs_ProcVol!cuenta1, rs_ProcVol!dlinea1, dh)
        End If
    Else
        If rs_ProcVol!cuenta2 <> "" Then
            'Flog.writeline "Generar Linea como está (Sin Cambios en dh)"
            Call GuardarLineaAsi(NroVol, rs_Proc_V_modasi!asi_cod, rs_ProcVol!linea2, Abs(rs_ProcVol!monto2), rs_ProcVol!cuenta2, rs_ProcVol!dlinea2, rs_ProcVol!dh2)
            If rs_ProcVol!dh2 Then
                montoDebe = montoDebe + Abs(rs_ProcVol!monto2)
            Else
                montoHaber = montoHaber + Abs(rs_ProcVol!monto2)
            End If
        Else
            'ERRORRRRRRRRRRRR mensaje
            Flog.writeline "Error: Existen dos cuentas sin nombre"
        End If
    End If
    
    rs_ProcVol.MoveNext
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    
    'Grabo incremento del proceso en la base
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
Loop

Monto = 0
montacum = 0

Do While Not rs_ProcDet.EOF
 
    If rs_ProcDet!c1 <> "" Then
        If rs_ProcDet!c2 <> "" Then
            'Generar Linea Diferencia
            Monto = rs_ProcDet!m1 - rs_ProcDet!m2
            montacum = rs_ProcDet!dlmontoacum1 - rs_ProcDet!dlmontoacum2 'Revisar si tengo que sacar algun monto mas de diferencia
            canti = rs_ProcDet!dlcantidad1 - rs_ProcDet!dlcantidad2
            If rs_ProcDet!linaD_H1 = rs_ProcDet!linaD_H2 Then
                If rs_ProcDet!linaD_H1 <= 1 Then
                    If rs_ProcDet!dlcantidad1 > rs_ProcDet!dlcantidad2 Then
                        dh = 0
                    Else
                        dh = 1
                    End If
                Else
                    dh = 2
                End If
            Else
                If Monto > 0 Then
                    dh = 2
                Else
                    dh = 1
                End If
            End If
            ' Generar lo siguiente:
            '       -Linea con la diferencia de montos, cuenta, etc     OK
            '       -Lineas de diferencia del detalle del asiento
            Call GuardarDetalleAsi(rs_Proc_V_modasi!asi_cod, rs_ProcDet!c1, canti, rs_ProcDet!dlcosto11, rs_ProcDet!dlcosto21, rs_ProcDet!dlcosto31, rs_ProcDet!dlcosto41, rs_ProcDet!dldescripcion1, Monto, montacum, rs_ProcDet!dlporcentaje1, rs_ProcDet!Ternro1, rs_ProcDet!empleg1, rs_ProcDet!lin_orden1, rs_ProcDet!terape1, NroVol, rs_ProcDet!Origen1, rs_ProcDet!tipoorigen1, rs_ProcDet!linadesc1, dh)
        Else
            If rs_ProcDet!linaD_H1 = 0 Then
                dh = 0
            Else
                If rs_ProcDet!linaD_H1 = 1 Then
                    dh = 2
                Else
                    dh = 1
                End If
            End If
            'Generar Linea cambiando debe x haber o viceversa
            Call GuardarDetalleAsi(rs_Proc_V_modasi!asi_cod, rs_ProcDet!c1, rs_ProcDet!dlcantidad1, rs_ProcDet!dlcosto11, rs_ProcDet!dlcosto21, rs_ProcDet!dlcosto31, rs_ProcDet!dlcosto41, rs_ProcDet!dldescripcion1, rs_ProcDet!m1, rs_ProcDet!dlmontacum1, rs_ProcDet!dlporcentaje1, rs_ProcDet!Ternro1, rs_ProcDet!empleg1, rs_ProcDet!lin_orden1, rs_ProcDet!terape1, NroVol, rs_ProcDet!Origen1, rs_ProcDet!tipoorigen1, rs_ProcDet!linadesc1, dh)
        End If
    Else
        If rs_ProcDet!c2 <> "" Then
            'Generar Linea como está (Sin Cambios)
            Call GuardarDetalleAsi(rs_Proc_V_modasi!asi_cod, rs_ProcDet!c2, rs_ProcDet!dlcantidad2, rs_ProcDet!dlcosto12, rs_ProcDet!dlcosto22, rs_ProcDet!dlcosto32, rs_ProcDet!dlcosto42, rs_ProcDet!dldescripcion2, rs_ProcDet!m2, rs_ProcDet!dlmontacum2, rs_ProcDet!dlporcentaje2, rs_ProcDet!Ternro2, rs_ProcDet!empleg2, rs_ProcDet!lin_orden2, rs_ProcDet!terape2, NroVol, rs_ProcDet!Origen2, rs_ProcDet!tipoorigen2, rs_ProcDet!linadesc2, rs_ProcDet!linaD_H2)
        Else
            'ERRORRRRRRRRRRRR mensaje
            Flog.writeline "Error: NO Existe la cuentaaaaaaaa "
        End If
    End If
    
    'Paso al siguiente concepto/acumulador
    rs_ProcDet.MoveNext
    
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    
    'Grabo incremento del proceso en la base
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
Loop
'------------ Una vez finalizado esto, genero el asiento correspondiente (asiento) ----------------
'-------------------------------------------------------------------------------
'Creo el asiento
'-------------------------------------------------------------------------------
    StrSql = "SELECT * FROM asiento " & _
             " WHERE masinro = " & rs_Proc_V_modasi!asi_cod & _
             " AND vol_cod = " & NroVol
    OpenRecordset StrSql, rs_asiento
    
    If rs_asiento.EOF Then
        StrSql = "INSERT INTO asiento (masinro,asidebe,asihaber,vol_cod) " & _
                 " VALUES (" & rs_Proc_V_modasi!asi_cod & _
                 "," & Round(montoDebe, 4) & _
                 "," & Round(montoHaber, 4) & _
                 "," & NroVol & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE asiento SET asidebe = " & Round(montoDebe, 4) & _
                 ",asihaber =" & Round(montoHaber, 4) & _
                 " WHERE masinro = " & rs_Proc_V_modasi!asi_cod & _
                 " AND vol_cod =" & NroVol
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rs_asiento.Close



'Cuento la cantidad de lineas generadas
StrSql = "SELECT count(*) Lineas FROM linea_asi "
StrSql = StrSql & " WHERE vol_cod =" & NroVol
If rs_Aux.State = adStateOpen Then rs_Aux.Close
OpenRecordset StrSql, rs_Aux
If Not rs_Aux.EOF Then
    NroLin = rs_Aux!Lineas
End If

'Cuento la cantidad de asientos generados
StrSql = "SELECT COUNT(DISTINCT masinro) Asientos FROM linea_asi "
StrSql = StrSql & " WHERE vol_cod =" & NroVol
If rs_Aux.State = adStateOpen Then rs_Aux.Close
OpenRecordset StrSql, rs_Aux
If Not rs_Aux.EOF Then
    NroAsi = rs_Aux!Asientos
End If

StrSql = "UPDATE proc_vol SET vol_reg_cab = " & NroAsi & _
             ", vol_reg_det =" & NroLin & _
             " WHERE proc_vol.vol_cod =" & NroVol
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "IF EXISTS(SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'asiento1')" & _
          " DROP TABLE asiento1"
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "IF EXISTS(SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'asiento2')" & _
          " DROP TABLE asiento2"
objConn.Execute StrSql, , adExecuteNoRecords

If rs_ProcVol.State = adStateOpen Then rs_ProcVol.Close
If rs_Proc_V_modasi.State = adStateOpen Then rs_Proc_V_modasi.Close
If rs_ProcDet.State = adStateOpen Then rs_ProcDet.Close
If rs_Aux.State = adStateOpen Then rs_Aux.Close
If rs_asiento.State = adStateOpen Then rs_asiento.Close

Set rs_ProcVol = Nothing
Set rs_Proc_V_modasi = Nothing
Set rs_ProcDet = Nothing
Set rs_asiento = Nothing


Exit Sub

'Manejador de Errores del procedimiento
ME_GenerarAsiento:
    If rs_ProcVol.State = adStateOpen Then rs_ProcVol.Close
    If rs_Proc_V_modasi.State = adStateOpen Then rs_Proc_V_modasi.Close
    If rs_ProcDet.State = adStateOpen Then rs_ProcDet.Close
    Set rs_ProcVol = Nothing
    Set rs_Proc_V_modasi = Nothing
    Set rs_ProcDet = Nothing
   
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: GenerarAsiento"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
End Sub







Public Sub GuardarLineaAsi(ByVal vol_cod As Long, ByVal masinro As Long, ByVal linea As Integer, ByVal Monto As Double, ByVal cuenta As String, ByVal desclinea As String, ByVal dh As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Inserta las cuentas en la base de datos en linea_asi
' Autor      : Juan Zamarbide
' Fecha      : 10/01/2012
' --------------------------------------------------------------------------------------------
Dim indice As Long
Dim rs_Linea_asi As New ADODB.Recordset
    
On Error GoTo ME_GuardarLineaAsi

        'Miro si la linea ya esta en la base para el proceso y modelo
        StrSql = "SELECT * FROM linea_asi " & _
                 " WHERE linea_asi.vol_cod = " & vol_cod & _
                 " AND linea_asi.cuenta  = '" & cuenta & "'" & _
                 " AND linea_asi.masinro = " & masinro
        OpenRecordset StrSql, rs_Linea_asi
        
        If rs_Linea_asi.EOF Then
        
            'No existe una linea con esa cuenta, entonces la inserto
            StrSql = "INSERT INTO linea_asi (cuenta,vol_cod,masinro,linea,desclinea,monto,dh)" & _
                     " VALUES ('" & Mid(cuenta, 1, 50) & _
                     "'," & vol_cod & _
                     "," & masinro & _
                     "," & linea & _
                     ",'" & desclinea & _
                     "'," & Monto & _
                     "," & dh & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
        
            'la linea existe, debo modificar el monto
            StrSql = "UPDATE linea_asi SET monto = monto + " & Monto & _
                     " WHERE linea_asi.vol_cod =" & vol_cod & _
                     " AND linea_asi.cuenta  ='" & cuenta & "'" & _
                     " AND linea_asi.masinro =" & masinro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        rs_Linea_asi.Close
    
    
    



'cierro todo
If rs_Linea_asi.State = adStateOpen Then rs_Linea_asi.Close
Set rs_Linea_asi = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_GuardarLineaAsi:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: GuardarLineaAsi"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql


End Sub

Public Sub GuardarDetalleAsi(ByVal masinro As Long, ByRef cuenta As String, ByVal dlcantidad As Double, ByVal dlcosto1 As Integer, ByVal dlcosto2 As Integer, ByVal dlcosto3 As Integer, ByVal dlcosto4 As Integer, ByRef dldescripcion As String, ByVal dlmonto As Double, ByVal dlmontacum As Double, ByVal dlporc As Double, ByVal Ternro As Long, ByVal empleg As Long, ByVal lin_orden As Long, ByRef terape As String, ByVal vol_cod As Long, ByVal Origen As Long, ByVal tipoorigen As Long, ByVal linadesc As String, ByVal linaD_H As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Inserta el detalle de las cuentas en la base de datos en detalle_asi
' Autor      : Juan Zamarbide
' Fecha      : 11/01/2012
' --------------------------------------------------------------------------------------------

Dim indice As Long
Dim rs_detalle_asi As New ADODB.Recordset

On Error GoTo ME_GuardarDetalleAsi

               
            'Miro si el detalle ya esta en la base para el proceso, modelo, cuenta y empleado
            StrSql = "SELECT * FROM detalle_asi " & _
                     " WHERE detalle_asi.vol_cod = " & vol_cod & _
                     " AND detalle_asi.cuenta  = '" & Mid(cuenta, 1, 50) & "'" & _
                     " AND detalle_asi.masinro = " & masinro & _
                     " AND detalle_asi.Origen = " & Origen & _
                     " AND detalle_asi.tipoorigen = " & tipoorigen & _
                     " AND detalle_asi.dlcosto4 = " & dlcosto4 & _
                     " AND detalle_asi.ternro = " & Ternro
            OpenRecordset StrSql, rs_detalle_asi
            
            If rs_detalle_asi.EOF Then
            
                'No existe una detalle con esa cuenta y empleado, entonces lo inserto
                StrSql = "INSERT INTO detalle_asi (masinro, cuenta,dlcantidad,dlcosto1,dlcosto2,dlcosto3,dlcosto4,dldescripcion " & _
                         ",dlmonto,dlmontoacum,dlporcentaje,ternro,empleg,lin_orden,terape,vol_cod, origen, tipoorigen,linadesc,linaD_H)" & _
                         " VALUES (" & masinro & _
                         ",'" & cuenta & _
                         "'," & dlcantidad & _
                         "," & dlcosto1 & _
                         "," & dlcosto2 & _
                         "," & dlcosto3 & _
                         "," & dlcosto4 & _
                         ",'" & dldescripcion & _
                         "'," & Round(dlmonto, 4) & _
                         "," & Round(dlmontacum, 4) & _
                         "," & dlporc & _
                         "," & Ternro & _
                         "," & empleg & _
                         "," & lin_orden & _
                         ",'" & Mid(terape, 1, 50) & _
                         "'," & vol_cod & _
                         "," & Origen & _
                         "," & tipoorigen & _
                         ",'" & Mid(linadesc, 1, 40) & _
                         "'," & linaD_H & _
                         ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
            
                'el detalle existe, debo modificar los montos y porcentaje
                StrSql = "UPDATE detalle_asi SET dlmonto = dlmonto + " & Round(dlmonto, 4) & _
                         ",dlmontoacum = dlmontoacum + " & Round(dlmontacum, 4) & _
                         ",dlporcentaje = dlporcentaje + " & Round(dlporc, 4) & _
                         " WHERE detalle_asi.vol_cod =" & vol_cod & _
                         " AND detalle_asi.cuenta  ='" & cuenta & "'" & _
                         " AND detalle_asi.masinro =" & masinro & _
                         " AND detalle_asi.Origen = " & Origen & _
                         " AND detalle_asi.tipoorigen = " & tipoorigen & _
                         " AND detalle_asi.dlcosto4 = " & dlcosto4 & _
                         " AND detalle_asi.ternro = " & Ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            rs_detalle_asi.Close

    
    
    'Reseteo los indices - 11/05/2007 - Martin Ferraro

'cierro todo
If rs_detalle_asi.State = adStateOpen Then rs_detalle_asi.Close
Set rs_detalle_asi = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_GuardarDetalleAsi:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: GuardarDetalleAsi"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub




