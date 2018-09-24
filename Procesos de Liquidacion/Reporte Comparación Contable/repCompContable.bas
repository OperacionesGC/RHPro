Attribute VB_Name = "repCompContable"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = ""
'Global Const UltimaModificacion = "Inicial"

'Global Const Version = "1.01"
'Global Const FechaModificacion = "16/01/2009"
'Global Const UltimaModificacion = "" ' Lisandro Moro - Correxion en sqls y referencias.

'Global Const Version = "1.02" ' Cesar Stankunas
'Global Const FechaModificacion = "05/08/2009"
'Global Const UltimaModificacion = ""    'Encriptacion de string connection

'Global Const Version = "1.03" ' Miriam Ruiz - CAS-34374 - VISION - Error en reporte comparativo
'Global Const FechaModificacion = "04/12/2015"
'Global Const UltimaModificacion = ""    'Se agregan log de información

Global Const Version = "1.04" ' Borrelli Facundo - CAS-34374 - VISION - Error en reporte comparativo (CAS-15298) [Entrega 2]
Global Const FechaModificacion = "23/02/2016"
Global Const UltimaModificacion = "Se agrega informacion detallada al log"


Dim fs, f
'Global Flog

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global Proc_Vol_1 As Integer
Global Proc_Vol_2 As Integer
Global Mod_Asi As Integer

'DATOS DE LA TABLA batch_proceso
Global bpfecha As Date
Global bphora As String
Global bpusuario As String

Global repNro As Integer
Global conceptos As String
Global acumuladores As String
Global procesos As String
Global idUser As String

Private Sub Main()

Dim NombreArchivo As String
Dim directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim Ternro
Dim rsEmpl As New ADODB.Recordset
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim ArrParametros
Dim parametros As String
'Dim ArrParametros

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
    
    Nombre_Arch = PathFLog & "ReporteCompContable" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    'Flog.writeline "PID = " & PID
    
    'FB
        Flog.writeline "-------------------------------------------------"
        Flog.writeline "Version                  : " & Version
        Flog.writeline "Fecha Ultima Modificacion: " & FechaModificacion
        Flog.writeline "Ultima Modificacion      : " & UltimaModificacion
        Flog.writeline "PID                      : " & PID
        Flog.writeline "-------------------------------------------------"
        Flog.writeline "Inicio Proceso Reporte comparativo contable : " & Now
        Flog.writeline "-------------------------------------------------"
        Flog.writeline
    'FB
    
    TiempoInicialProceso = GetTickCount
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo CE
    HuboErrores = False
    
    'Flog.writeline "Inicio Proceso de Control Pagos : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 0, bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    Progreso = 0
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo el Proceso de Volcado 1
       Proc_Vol_1 = ArrParametros(0)
       
       'Obtengo el Proceso de Volcado 1
       Proc_Vol_2 = ArrParametros(1)
       
       'Obtengo el modelo de asiento
       Mod_Asi = ArrParametros(2)
       
       'EMPIEZA EL PROCESO
       Call generarCompContable
    Else
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

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function

'--------------------------------------------------------------------
' Se encarga de buscar las auditorias
'--------------------------------------------------------------------
Sub generarCompContable()

Dim StrSql As String
Dim rsConsult1 As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim rsConsult3 As New ADODB.Recordset
Dim Legajo As Integer
Dim Cantidad As Integer
Dim cantidadProcesada As Integer

On Error GoTo MError

Flog.writeline " Entra a generar el reporte "
 
StrSql = " SELECT masinro "
StrSql = StrSql & " FROM mod_asiento "
If Mod_Asi <> "-1" Then
   StrSql = StrSql & " WHERE mod_asiento.masinro = " & Mod_Asi
End If
Flog.writeline " Modelo de asiento: " & Mod_Asi
Flog.writeline

OpenRecordset StrSql, rsConsult1

Cantidad = rsConsult1.RecordCount
cantidadProcesada = Cantidad

Dim Cuenta As String
Dim Nomina As String
Dim Monto As Double

Do Until rsConsult1.EOF

    'StrSql = " SELECT DISTINCT proc_vol.vol_cod, detliq.concnro, linea_asi.linea, linea_asi.masinro, "
    'StrSql = StrSql & " proc_vol_pl.pronro, cuenta, desclinea, conccod, concabr, dh, sum(detliq.dlimonto) as monto "
    'StrSql = StrSql & " FROM proc_vol "
    'StrSql = StrSql & " INNER JOIN linea_asi ON linea_asi.masinro = " & rsConsult1!masinro & " AND linea_asi.vol_cod = proc_vol.vol_cod "
    'StrSql = StrSql & " INNER JOIN asi_con ON asi_con.linaorden = linea_asi.linea AND asi_con.masinro = linea_asi.masinro "
    'StrSql = StrSql & " INNER JOIN proc_vol_pl ON proc_vol_pl.vol_cod = proc_vol.vol_cod "
    'StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proc_vol_pl.pronro "
    'StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro AND asi_con.concnro = detliq.concnro "
    'StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
    'StrSql = StrSql & " WHERE proc_vol.vol_cod = " & Proc_Vol_1
    'StrSql = StrSql & " GROUP BY proc_vol.vol_cod, detliq.concnro, linea_asi.linea, linea_asi.masinro, proc_vol_pl.pronro, cuenta, desclinea, dh, "
    'StrSql = StrSql & " conccod , concabr "
    'StrSql = StrSql & " ORDER BY linea_asi.masinro, linea_asi.linea "
    
    StrSql = " SELECT DISTINCT proc_vol.vol_cod, concepto.concnro, linea_asi.linea, linea_asi.masinro,"
    StrSql = StrSql & " linea_asi.cuenta, desclinea, conccod, concabr,asi_con.signo, dh, sum(detalle_asi.dlmonto) as monto, mod_linea.linaD_H "
    StrSql = StrSql & " FROM proc_vol "
    StrSql = StrSql & " INNER JOIN linea_asi ON linea_asi.masinro = " & rsConsult1!masinro & " AND linea_asi.vol_cod = proc_vol.vol_cod "
    StrSql = StrSql & " INNER JOIN detalle_asi ON detalle_asi.lin_orden = linea_asi.linea AND detalle_asi.masinro = linea_asi.masinro AND "
    StrSql = StrSql & " detalle_asi.vol_cod = linea_asi.vol_cod AND detalle_asi.cuenta = linea_asi.cuenta AND detalle_asi.tipoorigen = 1 "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detalle_asi.origen "
    StrSql = StrSql & " INNER JOIN mod_linea ON mod_linea.masinro = linea_asi.masinro AND mod_linea.linaorden = linea_asi.linea "
    StrSql = StrSql & " INNER JOIN asi_con ON asi_con.linaorden = linea_asi.linea AND asi_con.masinro = linea_asi.masinro AND asi_con.concnro = detalle_asi.origen"
    StrSql = StrSql & " WHERE proc_vol.vol_cod = " & Proc_Vol_1
    StrSql = StrSql & " GROUP BY proc_vol.vol_cod, concepto.concnro, linea_asi.linea, linea_asi.masinro, linea_asi.cuenta, desclinea,asi_con.signo, dh, "
    StrSql = StrSql & " Conccod , concabr, mod_linea.linaD_H  "
    StrSql = StrSql & " ORDER BY linea_asi.masinro, linea_asi.linea "
    OpenRecordset StrSql, rsConsult2
    
    If rsConsult2.EOF Then
        Flog.writeline " Consulta 1 vacía, si el reporte no arroja resultados verificar que el asiento del proceso de volcado 1 haya sido generado con análisis detallado "
        Flog.writeline
    Else
        Flog.writeline " Consulta 1 arrojó resultados"
        Flog.writeline
    End If
    
    'Seteo el progreso
    If rsConsult2.RecordCount <> 0 Then
        Cantidad = rsConsult2.RecordCount
    Else
        Cantidad = 1
    End If
    
    IncPorc = 25 / Cantidad
    
    Do Until rsConsult2.EOF
      Cuenta = rsConsult2!Cuenta & " " & rsConsult2!desclinea
      Nomina = rsConsult2!ConcCod & " " & rsConsult2!concabr
      Monto = rsConsult2!Monto
      
      Flog.writeline " Cuenta/Nomina/Monto1: " & Cuenta & "/" & Nomina & "/" & Monto
      Flog.writeline
      
      If IsNull(Monto) Then
        Monto = 0
      End If
      
      Select Case rsConsult2!dh 'debe
      Case True:
          Select Case CInt(rsConsult2!signo)
             'Case 0 '+
             Case 1 '+
                StrSql = " INSERT INTO rep_comp_contable "
                StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & ",'" & Cuenta & "'"
                StrSql = StrSql & ",'" & Nomina & "'"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",1"
                StrSql = StrSql & "," & Proc_Vol_1
                StrSql = StrSql & ")"
            'Case 1  '-
            Case 2  '-
                StrSql = " INSERT INTO rep_comp_contable "
                StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & ",'" & Cuenta & "'"
                StrSql = StrSql & ",'" & Nomina & "'"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",1"
                StrSql = StrSql & "," & Proc_Vol_1
                StrSql = StrSql & ")"
            'Case 2  '+/-
            Case 3  '+/-
                If Monto < 0 Then
                    StrSql = " INSERT INTO rep_comp_contable "
                    StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & ",'" & Cuenta & "'"
                    StrSql = StrSql & ",'" & Nomina & "'"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",1"
                    StrSql = StrSql & "," & Proc_Vol_1
                    StrSql = StrSql & ")"
                Else
                    StrSql = " INSERT INTO rep_comp_contable "
                    StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & ",'" & Cuenta & "'"
                    StrSql = StrSql & ",'" & Nomina & "'"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",1"
                    StrSql = StrSql & "," & Proc_Vol_1
                    StrSql = StrSql & ")"
                End If
          End Select
        Case Else   'Haber
          Select Case CInt(rsConsult2!signo)
             'Case 0 '+
             Case 2 '+
                StrSql = " INSERT INTO rep_comp_contable "
                StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & ",'" & Cuenta & "'"
                StrSql = StrSql & ",'" & Nomina & "'"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",1"
                StrSql = StrSql & "," & Proc_Vol_1
                StrSql = StrSql & ")"
            'Case 1  '-
            Case 1  '-
    
                StrSql = " INSERT INTO rep_comp_contable "
                StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & ",'" & Cuenta & "'"
                StrSql = StrSql & ",'" & Nomina & "'"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",1"
                StrSql = StrSql & "," & Proc_Vol_1
                StrSql = StrSql & ")"
            'Case 2  '+/-
            Case 3  '+/-
    
                If Monto < 0 Then
                    StrSql = " INSERT INTO rep_comp_contable "
                    StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & ",'" & Cuenta & "'"
                    StrSql = StrSql & ",'" & Nomina & "'"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",1"
                    StrSql = StrSql & "," & Proc_Vol_1
                    StrSql = StrSql & ")"
                Else
                    StrSql = " INSERT INTO rep_comp_contable "
                    StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & ",'" & Cuenta & "'"
                    StrSql = StrSql & ",'" & Nomina & "'"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",1"
                    StrSql = StrSql & "," & Proc_Vol_1
                    StrSql = StrSql & ")"
                End If
          End Select
        End Select
        objConn.Execute StrSql, , adExecuteNoRecords
      
        'Actualizo el progreso
        TiempoAcumulado = GetTickCount
        Progreso = Progreso + IncPorc
        cantidadProcesada = cantidadProcesada - 1
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
      
      rsConsult2.MoveNext
    Loop
    
    rsConsult2.Close
    
    StrSql = " SELECT DISTINCT proc_vol.vol_cod, acumulador.acunro, linea_asi.linea, linea_asi.masinro,"
    StrSql = StrSql & " linea_asi.cuenta, desclinea, acudesabr,asi_acu.signo, dh, sum(detalle_asi.dlmonto) as monto, mod_linea.linaD_H  "
    StrSql = StrSql & " FROM proc_vol "
    StrSql = StrSql & " INNER JOIN linea_asi ON linea_asi.masinro = " & rsConsult1!masinro & " AND linea_asi.vol_cod = proc_vol.vol_cod "
    StrSql = StrSql & " INNER JOIN detalle_asi ON detalle_asi.lin_orden = linea_asi.linea AND detalle_asi.masinro = linea_asi.masinro AND "
    StrSql = StrSql & " detalle_asi.vol_cod = linea_asi.vol_cod AND detalle_asi.cuenta = linea_asi.cuenta AND detalle_asi.tipoorigen = 2 "
    StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = detalle_asi.origen "
    StrSql = StrSql & " INNER JOIN mod_linea ON mod_linea.masinro = linea_asi.masinro AND mod_linea.linaorden = linea_asi.linea "
    StrSql = StrSql & " INNER JOIN asi_acu ON asi_acu.linaorden = linea_asi.linea AND asi_acu.masinro = linea_asi.masinro AND asi_acu.acunro = detalle_asi.origen"
    StrSql = StrSql & " WHERE proc_vol.vol_cod = " & Proc_Vol_1
    StrSql = StrSql & " GROUP BY proc_vol.vol_cod, acumulador.acunro, linea_asi.linea, linea_asi.masinro, linea_asi.cuenta, desclinea,asi_acu.signo, dh, "
    StrSql = StrSql & " acudesabr, mod_linea.linaD_H  "
    StrSql = StrSql & " ORDER BY linea_asi.masinro, linea_asi.linea "
    OpenRecordset StrSql, rsConsult2
    
     If rsConsult2.EOF Then
        Flog.writeline " Consulta 2 vacía, si el reporte no arroja resultados verificar que el asiento del proceso de volcado 1 haya sido generado con análisis detallado "
        Flog.writeline
    Else
        Flog.writeline " Consulta 2 arrojó resultados"
        Flog.writeline
    End If
    
    'Seteo el progreso
    If rsConsult2.RecordCount <> 0 Then
        Cantidad = rsConsult2.RecordCount
    Else
        Cantidad = 1
    End If
    IncPorc = 25 / Cantidad
       
    Do Until rsConsult2.EOF
      Cuenta = rsConsult2!Cuenta & " " & rsConsult2!desclinea
      Nomina = rsConsult2!acuNro & " " & rsConsult2!acudesabr
      Monto = rsConsult2!Monto
      
      Flog.writeline " Cuenta/Nomina/Monto 2: " & Cuenta & "/" & Nomina & "/" & Monto
      Flog.writeline
      
      If IsNull(Monto) Then
        Monto = 0
      End If
      
        Select Case rsConsult2!dh 'debe
        Case True:
            Select Case CInt(rsConsult2!signo)
            Case 1
                StrSql = " INSERT INTO rep_comp_contable "
                StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & ",'" & Cuenta & "'"
                StrSql = StrSql & ",'" & Nomina & "'"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",1"
                StrSql = StrSql & "," & Proc_Vol_1
                StrSql = StrSql & ")"
            Case 2
                StrSql = " INSERT INTO rep_comp_contable "
                StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & ",'" & Cuenta & "'"
                StrSql = StrSql & ",'" & Nomina & "'"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",1"
                StrSql = StrSql & "," & Proc_Vol_1
                StrSql = StrSql & ")"
            Case 3
                If Monto < 0 Then
                    StrSql = " INSERT INTO rep_comp_contable "
                    StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & ",'" & Cuenta & "'"
                    StrSql = StrSql & ",'" & Nomina & "'"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",1"
                    StrSql = StrSql & "," & Proc_Vol_1
                    StrSql = StrSql & ")"
                Else
                    StrSql = " INSERT INTO rep_comp_contable "
                    StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & ",'" & Cuenta & "'"
                    StrSql = StrSql & ",'" & Nomina & "'"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",1"
                    StrSql = StrSql & "," & Proc_Vol_1
                    StrSql = StrSql & ")"
                End If
            End Select
        Case Else   'Haber
            Select Case CInt(rsConsult2!signo)
            Case 2
                StrSql = " INSERT INTO rep_comp_contable "
                StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & ",'" & Cuenta & "'"
                StrSql = StrSql & ",'" & Nomina & "'"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",1"
                StrSql = StrSql & "," & Proc_Vol_1
                StrSql = StrSql & ")"
            Case 1
                StrSql = " INSERT INTO rep_comp_contable "
                StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & ",'" & Cuenta & "'"
                StrSql = StrSql & ",'" & Nomina & "'"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Abs(Monto)
                StrSql = StrSql & ",1"
                StrSql = StrSql & "," & Proc_Vol_1
                StrSql = StrSql & ")"
            Case 3
                If Monto < 0 Then
                    StrSql = " INSERT INTO rep_comp_contable "
                    StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & ",'" & Cuenta & "'"
                    StrSql = StrSql & ",'" & Nomina & "'"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",1"
                    StrSql = StrSql & "," & Proc_Vol_1
                    StrSql = StrSql & ")"
                Else
                    StrSql = " INSERT INTO rep_comp_contable "
                    StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol1 )"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & ",'" & Cuenta & "'"
                    StrSql = StrSql & ",'" & Nomina & "'"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & "," & Abs(Monto)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",1"
                    StrSql = StrSql & "," & Proc_Vol_1
                    StrSql = StrSql & ")"
                End If
            End Select
      End Select
      objConn.Execute StrSql, , adExecuteNoRecords
      
        'Actualizo el progreso
        TiempoAcumulado = GetTickCount
        Progreso = Progreso + IncPorc
        cantidadProcesada = cantidadProcesada - 1
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
      
      rsConsult2.MoveNext
    Loop
    rsConsult2.Close
    
    'PARA EL PROCESO DE VOLCADO 2 "
       
    StrSql = " SELECT DISTINCT proc_vol.vol_cod, concepto.concnro, linea_asi.linea, linea_asi.masinro,"
    StrSql = StrSql & " linea_asi.cuenta, desclinea, conccod, concabr,asi_con.signo, dh, sum(detalle_asi.dlmonto) as monto, mod_linea.linaD_H  "
    StrSql = StrSql & " FROM proc_vol "
    StrSql = StrSql & " INNER JOIN linea_asi ON linea_asi.masinro = " & rsConsult1!masinro & " AND linea_asi.vol_cod = proc_vol.vol_cod "
    StrSql = StrSql & " INNER JOIN detalle_asi ON detalle_asi.lin_orden = linea_asi.linea AND detalle_asi.masinro = linea_asi.masinro AND "
    StrSql = StrSql & " detalle_asi.vol_cod = linea_asi.vol_cod AND detalle_asi.cuenta = linea_asi.cuenta AND detalle_asi.tipoorigen = 1 "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detalle_asi.origen "
    StrSql = StrSql & " INNER JOIN mod_linea ON mod_linea.masinro = linea_asi.masinro AND mod_linea.linaorden = linea_asi.linea "
    StrSql = StrSql & " INNER JOIN asi_con ON asi_con.linaorden = linea_asi.linea AND asi_con.masinro = linea_asi.masinro AND asi_con.concnro = detalle_asi.origen"
    StrSql = StrSql & " WHERE proc_vol.vol_cod = " & Proc_Vol_2
    StrSql = StrSql & " GROUP BY proc_vol.vol_cod, concepto.concnro, linea_asi.linea, linea_asi.masinro, linea_asi.cuenta, desclinea,asi_con.signo, dh, "
    StrSql = StrSql & " Conccod , concabr, mod_linea.linaD_H  "
    StrSql = StrSql & " ORDER BY linea_asi.masinro, linea_asi.linea "
    OpenRecordset StrSql, rsConsult2
    
      If rsConsult2.EOF Then
        Flog.writeline " Consulta 3 vacía, si el reporte no arroja resultados verificar que el asiento del proceso de volcado 2 haya sido generado con análisis detallado "
        Flog.writeline
    Else
        Flog.writeline " Consulta 3 arrojó resultados"
        Flog.writeline
    End If
    
    'Seteo el progreso
    If rsConsult2.RecordCount <> 0 Then
        Cantidad = rsConsult2.RecordCount
    Else
        Cantidad = 1
    End If
    IncPorc = 25 / Cantidad
       
    Do Until rsConsult2.EOF
    
      Cuenta = rsConsult2!Cuenta & " " & rsConsult2!desclinea
      Nomina = rsConsult2!ConcCod & " " & rsConsult2!concabr
      Monto = rsConsult2!Monto
      
      Flog.writeline " Cuenta/Nomina/Monto 3: " & Cuenta & "/" & Nomina & "/" & Monto
      Flog.writeline
      
      If IsNull(Monto) Then
        Monto = 0
      End If
      
      StrSql = " SELECT * FROM rep_comp_contable "
      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
      
      OpenRecordset StrSql, rsConsult3
      If Not rsConsult3.EOF Then
        Select Case rsConsult2!dh 'debe
        Case True:
            Select Case CInt(rsConsult2!signo)
               Case 1
                  StrSql = " UPDATE rep_comp_contable "
                  StrSql = StrSql & " SET debe2=" & Abs(Monto)
                  StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                  StrSql = StrSql & " ,difdebe=" & CDbl(rsConsult3!debe1) - CDbl(Abs(Monto))
                  StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
               Case 2
                  StrSql = " UPDATE rep_comp_contable "
                  StrSql = StrSql & " SET haber2=" & Abs(Monto)
                  StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                  StrSql = StrSql & " ,difhaber=" & CDbl(rsConsult3!haber1) - CDbl(Abs(Monto))
                  StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
               Case 3
                  If Monto < 0 Then
                      StrSql = " UPDATE rep_comp_contable "
                      StrSql = StrSql & " SET haber2=" & Abs(Monto)
                      StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                      StrSql = StrSql & " ,difhaber=" & CDbl(rsConsult3!haber1) - CDbl(Abs(Monto))
                      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
                  Else
                      StrSql = " UPDATE rep_comp_contable "
                      StrSql = StrSql & " SET debe2=" & Abs(Monto)
                      StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                      StrSql = StrSql & " ,difdebe=" & CDbl(rsConsult3!debe1) - CDbl(Abs(Monto))
                      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
                  End If
            End Select
        Case Else
            Select Case CInt(rsConsult2!signo)
               Case 2
                  StrSql = " UPDATE rep_comp_contable "
                  StrSql = StrSql & " SET debe2=" & Abs(Monto)
                  StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                  StrSql = StrSql & " ,difdebe=" & CDbl(rsConsult3!debe1) - CDbl(Abs(Monto))
                  StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
               Case 1
                  StrSql = " UPDATE rep_comp_contable "
                  StrSql = StrSql & " SET haber2=" & Abs(Monto)
                  StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                  StrSql = StrSql & " ,difhaber=" & CDbl(rsConsult3!haber1) - CDbl(Abs(Monto))
                  StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
               Case 3
                  If Monto < 0 Then
                      StrSql = " UPDATE rep_comp_contable "
                      StrSql = StrSql & " SET haber2=" & Abs(Monto)
                      StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                      StrSql = StrSql & " ,difhaber=" & CDbl(rsConsult3!haber1) - CDbl(Abs(Monto))
                      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
                  Else
                      StrSql = " UPDATE rep_comp_contable "
                      StrSql = StrSql & " SET debe2=" & Abs(Monto)
                      StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                      StrSql = StrSql & " ,difdebe=" & CDbl(rsConsult3!debe1) - CDbl(Abs(Monto))
                      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
                  End If
            End Select
        End Select
      Else
        Select Case rsConsult2!dh 'debe
        Case True:
            Select Case CInt(rsConsult2!signo)
               Case 1
                  StrSql = " INSERT INTO rep_comp_contable "
                  StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                  StrSql = StrSql & " VALUES "
                  StrSql = StrSql & "(" & NroProceso
                  StrSql = StrSql & ",'" & Cuenta & "'"
                  StrSql = StrSql & ",'" & Nomina & "'"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto) * (-1)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",1"
                  StrSql = StrSql & "," & Proc_Vol_2
                  StrSql = StrSql & ")"
              Case 2
                  StrSql = " INSERT INTO rep_comp_contable "
                  StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                  StrSql = StrSql & " VALUES "
                  StrSql = StrSql & "(" & NroProceso
                  StrSql = StrSql & ",'" & Cuenta & "'"
                  StrSql = StrSql & ",'" & Nomina & "'"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto) * (-1)
                  StrSql = StrSql & ",1"
                  StrSql = StrSql & "," & Proc_Vol_2
                  StrSql = StrSql & ")"
              Case 3
                  If Monto < 0 Then
                      StrSql = " INSERT INTO rep_comp_contable "
                      StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                      StrSql = StrSql & " VALUES "
                      StrSql = StrSql & "(" & NroProceso
                      StrSql = StrSql & ",'" & Cuenta & "'"
                      StrSql = StrSql & ",'" & Nomina & "'"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto) * (-1)
                      StrSql = StrSql & ",1"
                      StrSql = StrSql & "," & Proc_Vol_2
                      StrSql = StrSql & ")"
                  Else
                      StrSql = " INSERT INTO rep_comp_contable "
                      StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                      StrSql = StrSql & " VALUES "
                      StrSql = StrSql & "(" & NroProceso
                      StrSql = StrSql & ",'" & Cuenta & "'"
                      StrSql = StrSql & ",'" & Nomina & "'"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto) * (-1)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",1"
                      StrSql = StrSql & "," & Proc_Vol_2
                      StrSql = StrSql & ")"
                  End If
            End Select
        Case Else
            Select Case CInt(rsConsult2!signo)
               Case 2
                  StrSql = " INSERT INTO rep_comp_contable "
                  StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                  StrSql = StrSql & " VALUES "
                  StrSql = StrSql & "(" & NroProceso
                  StrSql = StrSql & ",'" & Cuenta & "'"
                  StrSql = StrSql & ",'" & Nomina & "'"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto) * (-1)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",1"
                  StrSql = StrSql & "," & Proc_Vol_2
                  StrSql = StrSql & ")"
              Case 1
                  StrSql = " INSERT INTO rep_comp_contable "
                  StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                  StrSql = StrSql & " VALUES "
                  StrSql = StrSql & "(" & NroProceso
                  StrSql = StrSql & ",'" & Cuenta & "'"
                  StrSql = StrSql & ",'" & Nomina & "'"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto) * (-1)
                  StrSql = StrSql & ",1"
                  StrSql = StrSql & "," & Proc_Vol_2
                  StrSql = StrSql & ")"
              Case 3
                  If Monto < 0 Then
                      StrSql = " INSERT INTO rep_comp_contable "
                      StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                      StrSql = StrSql & " VALUES "
                      StrSql = StrSql & "(" & NroProceso
                      StrSql = StrSql & ",'" & Cuenta & "'"
                      StrSql = StrSql & ",'" & Nomina & "'"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto) * (-1)
                      StrSql = StrSql & ",1"
                      StrSql = StrSql & "," & Proc_Vol_2
                      StrSql = StrSql & ")"
                  Else
                      StrSql = " INSERT INTO rep_comp_contable "
                      StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                      StrSql = StrSql & " VALUES "
                      StrSql = StrSql & "(" & NroProceso
                      StrSql = StrSql & ",'" & Cuenta & "'"
                      StrSql = StrSql & ",'" & Nomina & "'"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto) * (-1)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",1"
                      StrSql = StrSql & "," & Proc_Vol_2
                      StrSql = StrSql & ")"
                  End If
            End Select
        End Select
      End If
      rsConsult3.Close
      objConn.Execute StrSql, , adExecuteNoRecords
      
        'Actualizo el progreso
        TiempoAcumulado = GetTickCount
        Progreso = Progreso + IncPorc
        cantidadProcesada = cantidadProcesada - 1
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
      
      rsConsult2.MoveNext
    Loop
    
    rsConsult2.Close
    
    StrSql = " SELECT DISTINCT proc_vol.vol_cod, acumulador.acunro, linea_asi.linea, linea_asi.masinro,"
    StrSql = StrSql & " linea_asi.cuenta, desclinea, acudesabr,asi_acu.signo, dh, sum(detalle_asi.dlmonto) as monto, mod_linea.linaD_H  "
    StrSql = StrSql & " FROM proc_vol "
    StrSql = StrSql & " INNER JOIN linea_asi ON linea_asi.masinro = " & rsConsult1!masinro & " AND linea_asi.vol_cod = proc_vol.vol_cod "
    StrSql = StrSql & " INNER JOIN detalle_asi ON detalle_asi.lin_orden = linea_asi.linea AND detalle_asi.masinro = linea_asi.masinro AND "
    StrSql = StrSql & " detalle_asi.vol_cod = linea_asi.vol_cod AND detalle_asi.cuenta = linea_asi.cuenta AND detalle_asi.tipoorigen = 2 "
    StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = detalle_asi.origen "
    StrSql = StrSql & " INNER JOIN mod_linea ON mod_linea.masinro = linea_asi.masinro AND mod_linea.linaorden = linea_asi.linea "
    StrSql = StrSql & " INNER JOIN asi_acu ON asi_acu.linaorden = linea_asi.linea AND asi_acu.masinro = linea_asi.masinro AND asi_acu.acunro = detalle_asi.origen"
    StrSql = StrSql & " WHERE proc_vol.vol_cod = " & Proc_Vol_2
    StrSql = StrSql & " GROUP BY proc_vol.vol_cod, acumulador.acunro, linea_asi.linea, linea_asi.masinro, linea_asi.cuenta, desclinea,asi_acu.signo, dh, "
    StrSql = StrSql & " acudesabr, mod_linea.linaD_H  "
    StrSql = StrSql & " ORDER BY linea_asi.masinro, linea_asi.linea "
    OpenRecordset StrSql, rsConsult2
       
      If rsConsult2.EOF Then
        Flog.writeline " Consulta 4 vacía, si el reporte no arroja resultados verificar que el asiento del proceso de volcado 2 haya sido generado con análisis detallado "
        Flog.writeline
    Else
        Flog.writeline " Consulta 4 arrojó resultados"
        Flog.writeline
    End If
       
    'Seteo el progreso
    If rsConsult2.RecordCount <> 0 Then
        Cantidad = rsConsult2.RecordCount
    Else
        Cantidad = 1
    End If
    IncPorc = 24 / Cantidad
       
    Do Until rsConsult2.EOF
    
      Cuenta = rsConsult2!Cuenta & " " & rsConsult2!desclinea
      Nomina = rsConsult2!acuNro & " " & rsConsult2!acudesabr
      Monto = rsConsult2!Monto
      Flog.writeline " Cuenta/Nomina/Monto 3: " & Cuenta & "/" & Nomina & "/" & Monto
      Flog.writeline
      
      If IsNull(Monto) Then
        Monto = 0
      End If
      
      StrSql = " SELECT * FROM rep_comp_contable "
      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
      OpenRecordset StrSql, rsConsult3
      
      If Not rsConsult3.EOF Then
      
        Select Case rsConsult2!dh 'debe
        Case True:
            Select Case CInt(rsConsult2!signo)
               Case 1
                  StrSql = " UPDATE rep_comp_contable "
                  StrSql = StrSql & " SET debe2=" & Abs(Monto)
                  StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                  StrSql = StrSql & " ,difdebe=" & CDbl(rsConsult3!debe1) - CDbl(Abs(Monto))
                  StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
               Case 2
                  StrSql = " UPDATE rep_comp_contable "
                  StrSql = StrSql & " SET haber2=" & Abs(Monto)
                  StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                  StrSql = StrSql & " ,difhaber=" & CDbl(rsConsult3!haber1) - CDbl(Abs(Monto))
                  StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
               Case 3
                  If Monto < 0 Then
                      StrSql = " UPDATE rep_comp_contable "
                      StrSql = StrSql & " SET haber2=" & Abs(Monto)
                      StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                      StrSql = StrSql & " ,difhaber=" & CDbl(rsConsult3!haber1) - CDbl(Abs(Monto))
                      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
                  Else
                      StrSql = " UPDATE rep_comp_contable "
                      StrSql = StrSql & " SET debe2=" & Abs(Monto)
                      StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                      StrSql = StrSql & " ,difdebe=" & CDbl(rsConsult3!debe1) - CDbl(Abs(Monto))
                      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
                  End If
            End Select
        Case Else
            Select Case CInt(rsConsult2!signo)
               Case 2
                  StrSql = " UPDATE rep_comp_contable "
                  StrSql = StrSql & " SET debe2=" & Abs(Monto)
                  StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                  StrSql = StrSql & " ,difdebe=" & CDbl(rsConsult3!debe1) - CDbl(Abs(Monto))
                  StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
               Case 1
                  StrSql = " UPDATE rep_comp_contable "
                  StrSql = StrSql & " SET haber2=" & Abs(Monto)
                  StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                  StrSql = StrSql & " ,difhaber=" & CDbl(rsConsult3!haber1) - CDbl(Abs(Monto))
                  StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
               Case 3
                  If Monto < 0 Then
                      StrSql = " UPDATE rep_comp_contable "
                      StrSql = StrSql & " SET haber2=" & Abs(Monto)
                      StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                      StrSql = StrSql & " ,difhaber=" & CDbl(rsConsult3!haber1) - CDbl(Abs(Monto))
                      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
                  Else
                      StrSql = " UPDATE rep_comp_contable "
                      StrSql = StrSql & " SET debe2=" & Abs(Monto)
                      StrSql = StrSql & " ,procvol2=" & Proc_Vol_2
                      StrSql = StrSql & " ,difdebe=" & CDbl(rsConsult3!debe1) - CDbl(Abs(Monto))
                      StrSql = StrSql & " WHERE cuenta = '" & Cuenta & "' AND cc_nomina = '" & Nomina & "' AND bpronro = " & NroProceso
                  End If
            End Select
        End Select
      Else
        Select Case rsConsult2!dh 'debe
        Case True:
            Select Case CInt(rsConsult2!signo)
               Case 1
                  StrSql = " INSERT INTO rep_comp_contable "
                  StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                  StrSql = StrSql & " VALUES "
                  StrSql = StrSql & "(" & NroProceso
                  StrSql = StrSql & ",'" & Cuenta & "'"
                  StrSql = StrSql & ",'" & Nomina & "'"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto) * (-1)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",1"
                  StrSql = StrSql & "," & Proc_Vol_2
                  StrSql = StrSql & ")"
              Case 2
                  StrSql = " INSERT INTO rep_comp_contable "
                  StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                  StrSql = StrSql & " VALUES "
                  StrSql = StrSql & "(" & NroProceso
                  StrSql = StrSql & ",'" & Cuenta & "'"
                  StrSql = StrSql & ",'" & Nomina & "'"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto) * (-1)
                  StrSql = StrSql & ",1"
                  StrSql = StrSql & "," & Proc_Vol_2
                  StrSql = StrSql & ")"
              Case 3
                  If Monto < 0 Then
                      StrSql = " INSERT INTO rep_comp_contable "
                      StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                      StrSql = StrSql & " VALUES "
                      StrSql = StrSql & "(" & NroProceso
                      StrSql = StrSql & ",'" & Cuenta & "'"
                      StrSql = StrSql & ",'" & Nomina & "'"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto) * (-1)
                      StrSql = StrSql & ",1"
                      StrSql = StrSql & "," & Proc_Vol_2
                      StrSql = StrSql & ")"
                  Else
                      StrSql = " INSERT INTO rep_comp_contable "
                      StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                      StrSql = StrSql & " VALUES "
                      StrSql = StrSql & "(" & NroProceso
                      StrSql = StrSql & ",'" & Cuenta & "'"
                      StrSql = StrSql & ",'" & Nomina & "'"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto) * (-1)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",1"
                      StrSql = StrSql & "," & Proc_Vol_2
                      StrSql = StrSql & ")"
                  End If
            End Select
        Case Else
            Select Case CInt(rsConsult2!signo)
               Case 2
                  StrSql = " INSERT INTO rep_comp_contable "
                  StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                  StrSql = StrSql & " VALUES "
                  StrSql = StrSql & "(" & NroProceso
                  StrSql = StrSql & ",'" & Cuenta & "'"
                  StrSql = StrSql & ",'" & Nomina & "'"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto) * (-1)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",1"
                  StrSql = StrSql & "," & Proc_Vol_2
                  StrSql = StrSql & ")"
              Case 1
                  StrSql = " INSERT INTO rep_comp_contable "
                  StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                  StrSql = StrSql & " VALUES "
                  StrSql = StrSql & "(" & NroProceso
                  StrSql = StrSql & ",'" & Cuenta & "'"
                  StrSql = StrSql & ",'" & Nomina & "'"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto)
                  StrSql = StrSql & ",0"
                  StrSql = StrSql & "," & Abs(Monto) * (-1)
                  StrSql = StrSql & ",1"
                  StrSql = StrSql & "," & Proc_Vol_2
                  StrSql = StrSql & ")"
              Case 3
                  If Monto < 0 Then
                      StrSql = " INSERT INTO rep_comp_contable "
                      StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                      StrSql = StrSql & " VALUES "
                      StrSql = StrSql & "(" & NroProceso
                      StrSql = StrSql & ",'" & Cuenta & "'"
                      StrSql = StrSql & ",'" & Nomina & "'"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto) * (-1)
                      StrSql = StrSql & ",1"
                      StrSql = StrSql & "," & Proc_Vol_2
                      StrSql = StrSql & ")"
                  Else
                      StrSql = " INSERT INTO rep_comp_contable "
                      StrSql = StrSql & " (bpronro , cuenta, cc_nomina, debe1, haber1, debe2, haber2, difdebe, difhaber, empresa, procvol2 )"
                      StrSql = StrSql & " VALUES "
                      StrSql = StrSql & "(" & NroProceso
                      StrSql = StrSql & ",'" & Cuenta & "'"
                      StrSql = StrSql & ",'" & Nomina & "'"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & "," & Abs(Monto) * (-1)
                      StrSql = StrSql & ",0"
                      StrSql = StrSql & ",1"
                      StrSql = StrSql & "," & Proc_Vol_2
                      StrSql = StrSql & ")"
                  End If
            End Select
        End Select
      End If
      rsConsult3.Close
      objConn.Execute StrSql, , adExecuteNoRecords
      
        'Actualizo el progreso
        TiempoAcumulado = GetTickCount
        Progreso = Progreso + IncPorc
        cantidadProcesada = cantidadProcesada - 1
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
      
      rsConsult2.MoveNext
    Loop
    rsConsult2.Close
    
          
'    cantidadProcesada = cantidadProcesada - 1
'    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & (((Cantidad - cantidadProcesada) * 100) / Cantidad) & _
'             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
'             ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
'    objConn.Execute StrSql, , adExecuteNoRecords

    rsConsult1.MoveNext
Loop

rsConsult1.Close

Exit Sub

MError:
    Resume Next
    Flog.writeline "Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub


Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
  End If

End Function

Function sinDatos(Str)

  If IsNull(Str) Then
     sinDatos = True
  Else
     If Trim(Str) = "" Then
        sinDatos = True
     Else
        sinDatos = False
     End If
  End If

End Function


