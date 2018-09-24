Attribute VB_Name = "MdlSIIDDJJ"
Option Explicit

'Const Version = "1.0"
'Const FechaVersion = "09-08-2007"
'Autor = Diego Rosso

'Const Version = "1.1"
'Const FechaVersion = "31/07/2009" 'Martin Ferraro - Encriptacion de string connection
Const Version = "1.2"
Const FechaVersion = "13/02/2014" 'Gonzalez Nicolás -

Global CantEmplError
Global CantEmplSinError
Global Errores As Boolean


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte SII-DDJJ CHILE.
' Autor      : Diego Rosso
' Fecha      : 09-08-2007
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
    
    
'strCmdLine = Command()
'
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProcesoBatch = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'
'        If IsNumeric(strCmdLine) Then
'            NroProcesoBatch = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If

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

    Nombre_Arch = PathFLog & "SII_DDJJ_RENTAS" & "-" & NroProcesoBatch & ".log"
    
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
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 191 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call SIIDDJJ(NroProcesoBatch, bprcparam)
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


Public Sub SIIDDJJ(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte SII-DDJJ
' Autor      : Diego Rosso
' Fecha      : 09-08-2007
' --------------------------------------------------------------------------------------------


Dim empresa As Long
Dim Lista_Mod As String
Dim PeriodoDesde As Date
Dim PeriodoHasta As Date

'Renta Total Neta
Dim EsRenTotNetaConc As Boolean
Dim RenTotNetaConf As Long
Dim RenTotNeta As Double

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

'Totalizados
Dim EsRentaPagadaConc As Boolean
Dim RentaPagadaConf As Long
Dim RentaPagada As Double

Dim EsRentaPagadaAnioConc As Boolean
Dim RentaPagadaAnioConf As Long
Dim RentaPagadaAnio As Double

Dim EsRentAccEneAbrConc As Boolean
Dim RentAccEneAbrConf As Long
Dim RentAccEneAbr As Double

Dim EsRentaGrabadaConc As Boolean
Dim RentaGrabadaConf As Long
Dim RentaGrabada As Double

Dim EsRebajasTotConc As Boolean
Dim RebajasTotConf As Long
Dim RebajasTot As Double

Dim EsTotalRemuConc As Boolean
Dim TotalRemuConf As Long
Dim TotalRemu As Double

'Sueldo bruto (Certificado ESS)
Dim SueldoBrutoConf As Long
Dim SueldoBruto As Double

'Cotización previsional o de Salud de Cargo del Trabajador (Certificado ESS)
Dim CotPrevConf As Long
Dim CotPrev As Double

'Factor Acutalización (Certificado ESS)
Dim FactorActConf As Long
Dim FactorAct As Double



Dim I      As Integer
Dim pos1 As Integer
Dim pos2 As Integer
Dim UltimoEmpleado As Long
Dim Apellido As String
Dim Apellido2 As String
Dim NombreEmp As String
Dim NombreEmp2 As String
Dim Rut As String
Dim DV As String
Dim Num_linea
Dim Titulo As String
Dim Z As Byte
Dim Meses(12) As Byte   '0 Falso 1 verdadero   para saber si el impuesto unico se liquido ese mes

'Datos de la empresa
Dim razSoc As String
Dim RutEmpresa As String
Dim domicilio As String
Dim comuna As String
Dim email As String
Dim fax As String
Dim tel As String
Dim rutEmp As String
Dim mes As String
Dim tipoContrato As String
Dim domnro As String
Dim empNom As String

Dim tipoTel As Integer
Dim tipoFax As Integer
Dim tipoDocEmpresa As Integer
Dim AnioTributario As String
Dim Hsjornada As Integer
Dim DocReplegal As String
Dim Arrauxdeci(5, 5) As String
Dim AcumNumCert As Integer
Dim ArrRenTotNeta(11) As Long
Dim ArrImpuesto(11) As Long
Dim ArrRetencion(11) As Long
Dim ArrRenTotExenta(11) As Long
Dim ArrRebajas(11) As Long
Dim ArrSueldoBruto(11) As Long
Dim ArrCotPrev(11) As Long
Dim ArrFactorAct(11) As Long
'Recordsets
Dim rs_Empleados As New ADODB.Recordset
Dim rs_CantEmpleados As New ADODB.Recordset
Dim rs_Acu_liq As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Rut As New ADODB.Recordset
Dim objRs3 As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset


 ' Inicio codigo ejecutable
    On Error GoTo CE

For I = 0 To 5
    Arrauxdeci(I, 0) = 0
    Arrauxdeci(I, 1) = 0
Next

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
        empresa = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Empresa = " & empresa
        Flog.writeline
        '------------------------------------------------------------------------------------
        
        'Periodo Desde
        '------------------------------------------------------------------------------------
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, "@") - 1
'        PeriodoDesde = Mid(Parametros, pos1, pos2 - pos1 + 1)
'
'        Flog.writeline "Posicion 1 = " & pos1
'        Flog.writeline "Pos 2 = " & pos2
'        Flog.writeline "Parametro PeriodoDesde = " & PeriodoDesde
'        Flog.writeline
        '------------------------------------------------------------------------------------
        
        'Periodo Hasta
        '------------------------------------------------------------------------------------
        'pos1 = pos2 + 2
        'pos2 = Len(Parametros)
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, "@") - 1
'        PeriodoHasta = Mid(Parametros, pos1, pos2 - pos1 + 1)
'
'        Flog.writeline "Posicion 1 = " & pos1
'        Flog.writeline "Pos 2 = " & pos2
'        Flog.writeline "Parametro Periodohasta = " & PeriodoHasta
'        Flog.writeline
        '------------------------------------------------------------------------------------
                
        'Año Tributario
        '------------------------------------------------------------------------------------
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        AnioTributario = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Año Tributario = " & AnioTributario
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
'Flog.writeline "Levantando configuracion del Reporte"
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 "
'OpenRecordset StrSql, rs_Confrep
'If rs_Confrep.EOF Then
'    Flog.writeline "No se encontró la configuración del Reporte"
'    Exit Sub
'End If

'Inicializo
RenTotNetaConf = 0
ImpuestoConf = 0
RetencionConf = 0
RenTotExentaConf = 0
RebajasConf = 0
RentaPagadaConf = 0
RentaPagadaAnioConf = 0
RentAccEneAbrConf = 0
RentaGrabadaConf = 0
RebajasTotConf = 0
TotalRemuConf = 0


Dim ConfTexto As String
'-----PARÁMETROS INICIALES DEL CONFREP
StrSql = "SELECT confnrocol,conftipo,confval,confval2 FROM confrep "
StrSql = StrSql & " WHERE repnro = 206 "
StrSql = StrSql & " ORDER BY confnrocol DESC"
OpenRecordset StrSql, rs_Confrep
Flog.writeline "Levantando configuracion"
If Not rs_Confrep.EOF Then
    Do While Not rs_Confrep.EOF
'        Select Case UCase(rs_Confrep!conftipo)
'            Case "TEL":
'                tipoTel = rs_Confrep!confval
'                Flog.writeline "Se obtuvo la configuracion de columna 27"
'            Case "FAX":
'                tipoFax = rs_Confrep!confval
'                Flog.writeline "Se obtuvo la configuracion de columna 28"
'            Case "DOC":
'                If rs_Confrep!confnrocol = 24 Then
'                    tipoDocEmpresa = rs_Confrep!confval
'                    Flog.writeline "Se obtuvo la configuracion de columna 24 (Tipo Doc. Empresa) "
'                ElseIf rs_Confrep!confnrocol = 25 Then
'                    DocReplegal = rs_Confrep!confval2
'                    Flog.writeline "Se obtuvo la configuracion de columna 24 (Doc. Representante Legal) "
'                End If
'
'            Case "JOR"
'                Hsjornada = rs_Confrep!confval
'                Flog.writeline "Se obtuvo la configuracion de columna 26"
'            Case Else
'                Flog.writeline "ERROR no se configuro correctamente la columna " & rs_Confrep!confnrocol
'        End Select
        If UCase(rs_Confrep!conftipo) = "CO" Then
            Flog.writeline "ERROR al Verificar Columnas del reporte"
            Flog.writeline "Solo se permiten configurar AC, revise la configuración del reporte"
            Exit Sub
        End If
       
       'Busco los AC
       Select Case rs_Confrep!confnrocol
                Case 2: 'Renta Total Neta Pagada
                        'EsRenTotNetaConc = False
                        RenTotNetaConf = rs_Confrep!confval
                        ConfTexto = "Renta Total Neta Pagada"
                Case 3: 'Impuesto único Retenido
                        'EsImpuestoConc = False
                        ImpuestoConf = rs_Confrep!confval
                        ConfTexto = "Impuesto único Retenido"
                Case 4: 'Mayor Retencion Solicitada
                        'EsRetencionConc = False
                        RetencionConf = rs_Confrep!confval
                        ConfTexto = "Mayor Retencion Solicitada"
                Case 5: 'Renta Total Exenta y/o no gravada
                        'EsRenTotExentaConc = False
                        RenTotExentaConf = rs_Confrep!confval
                        ConfTexto = "Renta Total Exenta y/o no gravada"
                Case 6: 'Rebaja por zonas extremas (Franquicia DL. 889)
                        'EsRebajasConc = False
                        RebajasConf = rs_Confrep!confval
                        ConfTexto = "Rebaja por zonas extremas (Franquicia DL. 889)"
                Case 13: 'RENTA TOTAL NETA PAGADA
                    'auxdeci1
                    Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
                    Arrauxdeci(0, 0) = 0
                    'Arrauxdeci(0, 1) = GetConcAcu(rs_Confrep!conftipo, rs_Confrep!confval, rs_Confrep!confval2)
                    Arrauxdeci(0, 1) = rs_Confrep!confval
                Case 14: 'POR RENTA TOTAL NETA PAGADA DURANTEL AÑO
                    'auxdeci2
                    Arrauxdeci(1, 0) = 0
                    Arrauxdeci(1, 1) = rs_Confrep!confval
                    'Arrauxdeci(1, 1) = GetConcAcu(rs_Confrep!conftipo, rs_Confrep!confval, rs_Confrep!confval2)
                    'Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
                Case 15: 'POR RENTAS ACCESORIAS Y/O COMPLEMENTARIA PAGADA ENTRE ENE-ABR AÑO SGTE.
                    'auxdeci3
                    Arrauxdeci(2, 0) = 0
                    Arrauxdeci(2, 1) = rs_Confrep!confval
                    'Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
                Case 16: 'RENTA TOTAL EXENTA Y/O NO GRAVADA
                    'auxdeci4
                    Arrauxdeci(3, 0) = 0
                    Arrauxdeci(3, 1) = rs_Confrep!confval
                    'Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
                Case 17: 'REBAJA POR ZONAS EXTREMAS (FRANQUICIA D.L.889)
                    'auxdeci5
                    Arrauxdeci(4, 0) = 0
                    Arrauxdeci(4, 1) = rs_Confrep!confval
                    'Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
                Case 18:
                    'auxdeci6
                    Arrauxdeci(5, 0) = 0
                    Arrauxdeci(5, 1) = rs_Confrep!confval
                    'Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
                
                Case 19:
                        SueldoBrutoConf = rs_Confrep!confval
                        ConfTexto = "Sueldo Bruto"
                Case 20:
                        CotPrevConf = rs_Confrep!confval
                        ConfTexto = "Cotización previsional o de Salud de Cargo del Trabajador "
                Case 21:
                        FactorActConf = rs_Confrep!confval
                        ConfTexto = "Factor Acutalización (Certificado ESS)"
                        
                Case 24: 'Tipo Doc. Empresa
                        tipoDocEmpresa = rs_Confrep!confval
                        ConfTexto = "Tipo Doc. Empresa"
                Case 25: 'Doc. Representante Legal
                        DocReplegal = rs_Confrep!confval2
                        ConfTexto = "Doc. Representante Legal"
                Case 26: 'Jornada
                        Hsjornada = rs_Confrep!confval
                        ConfTexto = "Jornada"
                Case 27: 'Teléfono
                        tipoTel = rs_Confrep!confval
                        ConfTexto = "Teléfono"
                Case 28:
                        tipoFax = rs_Confrep!confval
                        ConfTexto = "Fax"
        End Select
        Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol & " " & ConfTexto
        rs_Confrep.MoveNext
    Loop
Else
    Flog.writeline "ERROR al Verificar Columnas del reporte"
    Exit Sub
End If



'----------------------------------------
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
domnro = "0"
razSoc = ""
RutEmpresa = ""
domicilio = ""
comuna = ""
Hsjornada = 0
'busco el rut de la empresa
StrSql = "select * FROM estructura "
StrSql = StrSql & " inner join empresa on empresa.estrnro = estructura.estrnro "
StrSql = StrSql & " inner join tercero on empresa.ternro = tercero.ternro "
StrSql = StrSql & " left join ter_doc on ter_doc.ternro= tercero.ternro "
StrSql = StrSql & " left join cabdom on cabdom.ternro = tercero.ternro "
StrSql = StrSql & " left join detdom on detdom.domnro = cabdom.domnro "
'StrSql = StrSql & " WHERE estructura.estrnro = " & Empresa
StrSql = StrSql & " WHERE tidnro = " & tipoDocEmpresa & " And cabdom.domdefault = -1 AND empresa.empnro = " & empresa

OpenRecordset StrSql, objRs2
If Not objRs2.EOF Then
    If objRs2!domnro <> "" Then
        domnro = objRs2!domnro
    End If
   
    If objRs2!terrazsoc <> "" Then
        razSoc = objRs2!terrazsoc
    End If
   
    If objRs2!nrodoc <> "" Then
        RutEmpresa = objRs2!nrodoc
    End If
    
    If objRs2!calle <> "" Then
        domicilio = objRs2!calle
    End If
    
    If objRs2!nro <> "" Then
        domicilio = domicilio & " " & objRs2!nro
    Else
        domicilio = domicilio
    End If
    
    If objRs2!sector <> "" Then
        domicilio = domicilio & " " & objRs2!sector
    Else
        domicilio = domicilio
    End If
    
    If objRs2!torre <> "" Then
        domicilio = domicilio & " " & objRs2!torre
    Else
        domicilio = domicilio
    End If
    
    If objRs2!piso <> "" Then
        domicilio = domicilio & " " & objRs2!piso
    Else
        domicilio = domicilio
    End If
    
    If objRs2!oficdepto <> "" Then
        domicilio = domicilio & " " & objRs2!oficdepto & " "
    Else
        domicilio = domicilio
    End If
    
    If objRs2!auxchr2 <> "" Then
        comuna = objRs2!auxchr2
    End If
    
    If objRs2!email <> "" Then
        email = objRs2!email
    Else
        email = ""
    End If
Else
    Flog.writeline "Error al levantar los datos de la empresa."

End If
 
domicilio = domicilio
 
objRs2.Close
 
 
'busco el fax y telefono
StrSql = " select tipoTel,telnro from telefono "
StrSql = StrSql & " inner join tipotel on telefono.tipotel = tipotel.titelnro "
StrSql = StrSql & " WHERE telefono.domnro =  " & domnro
OpenRecordset StrSql, objRs2
If Not objRs2.EOF Then
   Do While Not objRs2.EOF
      If objRs2!tipoTel = tipoFax Then
          fax = objRs2!telnro
          Flog.writeline "Se obtuvo N° de Fax :" & fax
      ElseIf objRs2!tipoTel = tipoTel Then
          tel = objRs2!telnro
          Flog.writeline "Se obtuvo N° de Teléfono :" & tel
      End If
   objRs2.MoveNext
   Loop
End If

'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

'Levanto la configuracion para el reporte

'Inicializo
'RenTotNetaConf = 0
'ImpuestoConf = 0
'RetencionConf = 0
'RenTotExentaConf = 0
'RebajasConf = 0
'RentaPagadaConf = 0
'RentaPagadaAnioConf = 0
'RentAccEneAbrConf = 0
'RentaGrabadaConf = 0
'RebajasTotConf = 0
'TotalRemuConf = 0

'Renta Total Neta Pagada
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 2 "
'OpenRecordset StrSql, rs_Confrep
'Flog.writeline "Levantando configuracion de columna 2"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsRenTotNetaConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             RenTotNetaConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsRenTotNetaConc = False
'           RenTotNetaConf = rs_Confrep!confval
'      End If
'      Flog.writeline "Se obtuvo la configuracion de columna 2"
'Else
'      Flog.writeline "ERROR no se configuro correctamente la columna 2"
'End If


'Impuesto Unico
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 3 "
'OpenRecordset StrSql, rs_Confrep
'Flog.writeline "Levantando configuracion de columna 3"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsImpuestoConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             ImpuestoConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsImpuestoConc = False
'           ImpuestoConf = rs_Confrep!confval
'      End If
'      Flog.writeline "Se obtuvo la configuracion de columna 3"
'Else
'      Flog.writeline "ERROR no se configuro correctamente la columna 3"
'End If


'Mayor Retencion Solicitada
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 4 "
'OpenRecordset StrSql, rs_Confrep
'Flog.writeline "Levantando configuracion de columna 4"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsRetencionConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             RetencionConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsRetencionConc = False
'           RetencionConf = rs_Confrep!confval
'      End If
'      Flog.writeline "Se obtuvo la configuracion de columna 4"
'Else
'      Flog.writeline "ERROR no se configuro correctamente la columna 4"
'End If


'Renta Total Exenta
'
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 5 "
'OpenRecordset StrSql, rs_Confrep
'Flog.writeline "Levantando configuracion de columna 5"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsRenTotExentaConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             RenTotExentaConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsRenTotExentaConc = False
'           RenTotExentaConf = rs_Confrep!confval
'      End If
'      Flog.writeline "Se obtuvo la configuracion de columna 5"
'Else
'      Flog.writeline "ERROR no se configuro correctamente la columna 5"
'End If


'Rebajas
'
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 6 "
'OpenRecordset StrSql, rs_Confrep
'Flog.writeline "Levantando configuracion de columna 6"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsRebajasConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             RebajasConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsRebajasConc = False
'           RebajasConf = rs_Confrep!confval
'      End If
'      Flog.writeline "Se obtuvo la configuracion de columna 6"
'Else
'      Flog.writeline "ERROR no se configuro correctamente la columna 6"
'End If
'
'Flog.writeline "Se obtuvo la configuracion del Reporte"
  
'Totalizados. Renta Total Neta Pagada
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 7 "
'OpenRecordset StrSql, rs_Confrep
'Flog.Writeline "Levantando configuracion de columna 7"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsRentaPagadaConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             RentaPagadaConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsRentaPagadaConc = False
'           RentaPagadaConf = rs_Confrep!confval
'      End If
'      Flog.Writeline "Se obtuvo la configuracion de columna 7"
'Else
'      Flog.Writeline "ERROR no se configuro correctamente la columna 7"
'End If

'Totalizados. Renta Total Neta Pagada Durante el año
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 8 "
'OpenRecordset StrSql, rs_Confrep
'Flog.Writeline "Levantando configuracion de columna 8"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsRentaPagadaAnioConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             RentaPagadaAnioConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsRentaPagadaAnioConc = False
'           RentaPagadaAnioConf = rs_Confrep!confval
'      End If
'      Flog.Writeline "Se obtuvo la configuracion de columna 8"
'Else
'      Flog.Writeline "ERROR no se configuro correctamente la columna 8"
'End If

'Totalizados. Renta accesorias entre Enero y Abril
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 9 "
'OpenRecordset StrSql, rs_Confrep
'Flog.Writeline "Levantando configuracion de columna 9"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsRentAccEneAbrConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             RentAccEneAbrConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsRentAccEneAbrConc = False
'           RentAccEneAbrConf = rs_Confrep!confval
'      End If
'      Flog.Writeline "Se obtuvo la configuracion de columna 9"
'Else
'      Flog.Writeline "ERROR no se configuro correctamente la columna 9"
'End If

'Totalizados. Renta Total exenta y/o no gravada
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 10 "
'OpenRecordset StrSql, rs_Confrep
'Flog.Writeline "Levantando configuracion de columna 10"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsRentaGrabadaConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             RentaGrabadaConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsRentaGrabadaConc = False
'           RentaGrabadaConf = rs_Confrep!confval
'      End If
'      Flog.Writeline "Se obtuvo la configuracion de columna 10"
'Else
'      Flog.Writeline "ERROR no se configuro correctamente la columna 10"
'End If

'Totalizados. Rebajas total
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 11 "
'OpenRecordset StrSql, rs_Confrep
'Flog.Writeline "Levantando configuracion de columna 11"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsRebajasTotConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             RebajasTotConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsRebajasTotConc = False
'           RebajasTotConf = rs_Confrep!confval
'      End If
'      Flog.Writeline "Se obtuvo la configuracion de columna 11"
'Else
'      Flog.Writeline "ERROR no se configuro correctamente la columna 11"
'End If

'Totalizados. Total Remuneracion imponible
'StrSql = "SELECT * FROM confrep WHERE repnro = 206 and confnrocol = 12 "
'OpenRecordset StrSql, rs_Confrep
'Flog.Writeline "Levantando configuracion de columna 12"
'If Not rs_Confrep.EOF Then
'      If UCase(rs_Confrep!conftipo) = "CO" Then
'           EsTotalRemuConc = True
'
'           StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!confval
'               If Not EsNulo(rs_Confrep!confval2) Then
'                    StrSql = StrSql & " OR conccod = '" & rs_Confrep!confval2 & "'"
'               End If
'           OpenRecordset StrSql, objRs3
'           If Not objRs3.EOF Then
'             TotalRemuConf = objRs3!concnro
'           End If
'           objRs3.Close
'      Else
'           EsTotalRemuConc = False
'           TotalRemuConf = rs_Confrep!confval
'      End If
'      Flog.Writeline "Se obtuvo la configuracion de columna 12"
'Else
'      Flog.Writeline "ERROR no se configuro correctamente la columna 12"
'End If

'--------------------
'CONFIGURACIONES PARA RECUADRO : TOTAL MONTOS ANUALES SIN ACTUALIZAR
'StrSql = "SELECT confnrocol,confval,confval2,conftipo FROM confrep WHERE repnro = 206 and confnrocol IN (13,14,15,16,17,18) "
'OpenRecordset StrSql, rs_Confrep
'If Not rs_Confrep.EOF Then
'    Do While Not rs_Confrep.EOF
'        Select Case rs_Confrep!confnrocol
'            Case 13:
'                'RENTA TOTAL NETA PAGADA
'                'auxdeci1
'                Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
'                Arrauxdeci(0, 0) = rs_Confrep!conftipo
'                Arrauxdeci(0, 1) = GetConcAcu(rs_Confrep!conftipo, rs_Confrep!confval, rs_Confrep!confval2)
'            Case 14:
'                'POR RENTA TOTAL NETA PAGADA DURANTEL AÑO
'                'auxdeci2
'                Arrauxdeci(1, 0) = rs_Confrep!conftipo
'                Arrauxdeci(1, 1) = GetConcAcu(rs_Confrep!conftipo, rs_Confrep!confval, rs_Confrep!confval2)
'                Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
'            Case 15:
'                'POR RENTAS ACCESORIAS Y/O COMPLEMENTARIA PAGADA ENTRE ENE-ABR AÑO SGTE.
'                'auxdeci3
'                Arrauxdeci(2, 0) = rs_Confrep!conftipo
'                Arrauxdeci(2, 1) = GetConcAcu(rs_Confrep!conftipo, rs_Confrep!confval, rs_Confrep!confval2)
'                Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
'            Case 16:
'                'RENTA TOTAL EXENTA Y/O NO GRAVADA
'                'auxdeci4
'                Arrauxdeci(3, 0) = rs_Confrep!conftipo
'                Arrauxdeci(3, 1) = GetConcAcu(rs_Confrep!conftipo, rs_Confrep!confval, rs_Confrep!confval2)
'                Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
'            Case 17:
'                'REBAJA POR ZONAS EXTREMAS (FRANQUICIA D.L.889)
'                'auxdeci5
'                Arrauxdeci(4, 0) = rs_Confrep!conftipo
'                Arrauxdeci(4, 1) = GetConcAcu(rs_Confrep!conftipo, rs_Confrep!confval, rs_Confrep!confval2)
'                Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
'            Case 18:
'                'auxdeci6
'                Arrauxdeci(5, 0) = rs_Confrep!conftipo
'                Arrauxdeci(5, 1) = GetConcAcu(rs_Confrep!conftipo, rs_Confrep!confval, rs_Confrep!confval2)
'                Flog.writeline "Se obtuvo la configuracion de columna " & rs_Confrep!confnrocol
'            Case Else
'                Flog.writeline "ERROR no se configuro correctamente la columna " & rs_Confrep!confnrocol
'                'TOTAL REMUNERACIÓN IMPONIBLE PARA EFECTOS PREVISIONALES ACTUALIZADA A TODOS LOS TRABAJADORES
'        End Select
'
'        rs_Confrep.MoveNext
'    Loop
'Else
'      Flog.writeline "ERROR no se enontró configuración para el cuadro TOTAL MONTOS ANUALES SIN ACTUALIZAR"
'      Flog.writeline "Configurar AC/CO en las columnas: 13,14,15,16,17 y 18"
'
'End If

'--------------------
UltimoEmpleado = -1
Num_linea = 1

StrSql = "SELECT distinct(empleado.ternro),empleado.empleg, cabliq.empleado FROM proceso "
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & " INNER JOIN  tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
StrSql = StrSql & " WHERE "
If Lista_Mod <> "0" Then
    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
End If
'    StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
StrSql = StrSql & " periodo.pliqanio = " & AnioTributario
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
    
    RentaPagada = 0
    RentaPagadaAnio = 0
    RentAccEneAbr = 0
    RentaGrabada = 0
    RebajasTot = 0
    TotalRemu = 0
    
    AcumNumCert = 0
    Do While Not rs_Empleados.EOF
          'Incremento para generar el número de certificado correspondiente
          'AcumNumCert = AcumNumCert + 1
          
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
                Rut = Mid(rs_Rut!nrodoc, 1, Len(rs_Rut!nrodoc) - 1)
                Rut = Replace(Rut, "-", "")
                DV = Right(rs_Rut!nrodoc, 1)
                'Flog.Writeline "RUT y DV obtenidos"
                Flog.writeline "RUT obtenido"
            Else
                Flog.writeline "Error al obtener los datos del RUT"
                Rut = ""
                DV = ""
                HuboError = True
            End If
            Flog.writeline ""
              
            '********************************************************************************
            'Busca los valores de los conceptos o acumuladores liquidados que se configuraron en
            'el confrep para el empleado en todos los procesos del tipo (modelo) seleccionado
            'en el filtro en el periodo de un año
            
            'Inicializo el arreglo en 0(Falso)
            For Z = 1 To 12
                Meses(Z) = 0
            Next Z
            'Inicializo variables
            RenTotNeta = 0
            Impuesto = 0
            Retencion = 0
            RenTotExenta = 0
            Rebajas = 0
                    
'            For Z = 0 To 11
'                ArrRenTotNeta(Z) = 0
'                ArrImpuesto(Z) = 0
'                ArrRetencion(Z) = 0
'                ArrRenTotExenta(Z) = 0
'                ArrRebajas(Z) = 0
'            Next
'            For I = 0 To 5
'                Arrauxdeci(Z, 0) = 0
'            Next
            'Flog.Writeline "Obteniendo la Renta Total Neta Pagada para el empleado "
            For Z = 0 To 11
                'lIMPIO ARREGLOS ANTES DE ASIGNARLES VALOR
                
                '---------------------------------------------

                '---------------------------------------------
                'Renta Total Neta Pagada.
                ArrRenTotNeta(Z) = buscarConceptoAcumPorEtiqueta("ACM", rs_Empleados!ternro, RenTotNetaConf, Z, AnioTributario)
                RenTotNeta = RenTotNeta + ArrRenTotNeta(Z)
                
                '---------------------------------------------
                'Impuesto unico retenido
                ArrImpuesto(Z) = buscarConceptoAcumPorEtiqueta("ACM", rs_Empleados!ternro, ImpuestoConf, Z, AnioTributario)
                Impuesto = Impuesto + ArrImpuesto(Z)
                
                '---------------------------------------------
                'Retención
                ArrRetencion(Z) = buscarConceptoAcumPorEtiqueta("ACM", rs_Empleados!ternro, RetencionConf, Z, AnioTributario)
                Retencion = Retencion + ArrRetencion(Z)
                '---------------------------------------------
                
                'Renta Total Exenta y/o no gravada
                ArrRenTotExenta(Z) = buscarConceptoAcumPorEtiqueta("ACM", rs_Empleados!ternro, RenTotExentaConf, Z, AnioTributario)
                RenTotExenta = RenTotExenta + ArrRenTotExenta(Z)
                
                '---------------------------------------------
                'Rebajas por zonas extremas
                ArrRebajas(Z) = buscarConceptoAcumPorEtiqueta("ACM", rs_Empleados!ternro, RebajasConf, Z, AnioTributario)
                Rebajas = Rebajas + ArrRebajas(Z)
                '---------------------------------------------
                
                'auxdeci_ 1 a 6
                For I = 0 To 5
                    Arrauxdeci(I, 0) = buscarConceptoAcumPorEtiqueta("ACM", rs_Empleados!ternro, Arrauxdeci(I, 1), Z, AnioTributario)
                Next
                
                '---------------------------------------------
                'Sueldo Bruto (Certificado ESS)
                ArrSueldoBruto(Z) = buscarConceptoAcumPorEtiqueta("ACM", rs_Empleados!ternro, SueldoBrutoConf, Z, AnioTributario)
                SueldoBruto = SueldoBruto + ArrSueldoBruto(Z)
                '---------------------------------------------
                '---------------------------------------------
                'Cotización previsional o de Salud de Cargo del Trabajador (Certificado ESS)
                ArrCotPrev(Z) = buscarConceptoAcumPorEtiqueta("ACM", rs_Empleados!ternro, CotPrevConf, Z, AnioTributario)
                CotPrev = CotPrev + ArrCotPrev(Z)
                '---------------------------------------------

                '---------------------------------------------
                'Factor Acutalización (Certificado ESS)
                ArrFactorAct(Z) = buscarConceptoAcumPorEtiqueta("ACM", rs_Empleados!ternro, FactorActConf, Z, AnioTributario)
                FactorAct = FactorAct + ArrFactorAct(Z)
                '---------------------------------------------


                

            Next
            'buscarConceptoAcumPorEtiquetaAnual("ACA", RenTotNetaConf, AnioTributario, "SEXO", ByVal estrnro As Long, ByVal fechaDesde As Date, ByVal fechaHasta As Date, ByVal empresa As Long, ByVal sucursal As Long)
            'ArrRentaTotNetaP = buscarConceptoAcumPorEtiqueta("AC", rs_Empleados!Ternro, RenTotNetaConf, 1, 2014)
            
'            Flog.Writeline "Obteniendo la Renta Total Neta Pagada para el empleado "
'            If EsRenTotNetaConc Then
'                StrSql = "SELECT dlimonto FROM proceso "
'                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  and  empresa.empnro = " & empresa
'                StrSql = StrSql & " WHERE "
'                 If Lista_Mod <> "0" Then
'                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                End If
'
'                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                StrSql = StrSql & " AND detliq.concnro = " & RenTotNetaConf
'                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                OpenRecordset StrSql, rs_Detliq
'                Do While Not rs_Detliq.EOF
'                   RenTotNeta = rs_Detliq!dlimonto
'                   rs_Detliq.MoveNext
'                Loop
'            Else
'                StrSql = "SELECT almonto FROM proceso "
'                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                StrSql = StrSql & " WHERE "
'                 If Lista_Mod <> "0" Then
'                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                End If
'
'                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                StrSql = StrSql & " AND acu_liq.acunro = " & RenTotNetaConf
'                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                OpenRecordset StrSql, rs_Detliq
'                Do While Not rs_Detliq.EOF
'                   RenTotNeta = rs_Detliq!almonto
'                   rs_Detliq.MoveNext
'                Loop
'            End If
'            Flog.Writeline "Se Obtuvo la Renta Total Neta Pagada "
            
'            Flog.Writeline "Obteniendo El impuesto unico retenido para el empleado "
            'Impuesto Unico
'            If EsImpuestoConc Then
'                StrSql = "SELECT dlimonto, profecini FROM proceso "
'                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                StrSql = StrSql & " WHERE "
'                If Lista_Mod <> "0" Then
'                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                End If
'                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & "  AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                StrSql = StrSql & " AND detliq.concnro = " & ImpuestoConf
'                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                OpenRecordset StrSql, rs_Detliq
'                Do While Not rs_Detliq.EOF
'                   Impuesto = rs_Detliq!dlimonto
'                   'Meses(Month(rs_Detliq!profecini)) = 1 'lo pongo en verdadero para saber que tiene liquidado ese concepto en ese mes
'                   Meses(Month(rs_Detliq!profecini)) = JornadaDelEmpleado(rs_Empleados!Ternro, Hsjornada, rs_Detliq!profecini)
'                   rs_Detliq.MoveNext
'                Loop
'            Else
'                StrSql = "SELECT almonto, profecini FROM proceso "
'                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                StrSql = StrSql & " WHERE "
'                If Lista_Mod <> "0" Then
'                    StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                End If
'                StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                StrSql = StrSql & " AND acu_liq.acunro = " & ImpuestoConf
'                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                OpenRecordset StrSql, rs_Detliq
'                Do While Not rs_Detliq.EOF
'                   Impuesto = rs_Detliq!almonto
'                   'Meses(Month(rs_Detliq!profecini)) = 1 'lo pongo en verdadero para saber que tiene liquidado ese acumulador en ese mes
'                   Meses(Month(rs_Detliq!profecini)) = JornadaDelEmpleado(rs_Empleados!Ternro, Hsjornada, rs_Detliq!profecini)
'                   rs_Detliq.MoveNext
'                Loop
'            End If
'            Flog.Writeline "Se obtuvo El impuesto unico retenido "
                    
'                    Flog.Writeline "Obteniendo la Mayor Retencion Solicitada para el empleado "
'                    'Mayor Retencion Solicitada
'                    If EsRetencionConc Then
'                        StrSql = "SELECT dlimonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND detliq.concnro = " & RetencionConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           Retencion = rs_Detliq!dlimonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    Else
'                        StrSql = "SELECT almonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND acu_liq.acunro = " & RetencionConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           Retencion = rs_Detliq!almonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    End If
'                    Flog.Writeline "Se obtuvo la Mayor Retencion Solicitada "
                    

'                    Flog.Writeline "Obteniendo la Renta Total Exenta y/o no gravada para el empleado "
'                    'Renta Total Exenta
'                    If EsRenTotExentaConc Then
'                        StrSql = "SELECT dlimonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                         If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND detliq.concnro = " & RenTotExentaConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RenTotExenta = rs_Detliq!dlimonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    Else
'                        StrSql = "SELECT almonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND acu_liq.acunro = " & RenTotExentaConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RenTotExenta = rs_Detliq!almonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    End If
'                    Flog.Writeline "Se obtuvo la Renta Total Exenta y/o no gravada "

'                    Flog.Writeline "Obteniendo las Rebajas por zonas extremas para el empleado "
'                    'Rebajas
'                    If EsRebajasConc Then
'                        StrSql = "SELECT dlimonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND detliq.concnro = " & RebajasConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           Rebajas = rs_Detliq!dlimonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    Else
'                        StrSql = "SELECT almonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND  proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND acu_liq.acunro = " & RebajasConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           Rebajas = rs_Detliq!almonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    End If
'                    Flog.Writeline "Se obtuvo las Rebajas por zonas extremas  "
                    
                    '********************************************************************************
                    
                    '********************************************************************************
                    'TOTALIZADOS. Se graban en la tabla de cabecera del reporte al finalizar de procesar los empleados
                    'Renta total neta pagada
'                    If EsRentaPagadaConc Then
'                        StrSql = "SELECT dlimonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND detliq.concnro = " & RentaPagadaConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RentaPagada = rs_Detliq!dlimonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    Else
'                        StrSql = "SELECT almonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND  proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND acu_liq.acunro = " & RebajasConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RentaPagada = rs_Detliq!almonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    End If
'
'                    'Renta total neta pagada durante el año
'                    If EsRentaPagadaAnioConc Then
'                        StrSql = "SELECT dlimonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND detliq.concnro = " & RentaPagadaAnioConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RentaPagadaAnio = rs_Detliq!dlimonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    Else
'                        StrSql = "SELECT almonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND  proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND acu_liq.acunro = " & RentaPagadaAnioConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RentaPagadaAnio = rs_Detliq!almonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    End If
'
'                    'Renta accesorias y/o complementaria pagada entre enero y abril año siguiente
'                    If EsRentAccEneAbrConc Then
'                        StrSql = "SELECT dlimonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND detliq.concnro = " & RentAccEneAbrConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RentAccEneAbr = rs_Detliq!dlimonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    Else
'                        StrSql = "SELECT almonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND  proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND acu_liq.acunro = " & RentAccEneAbrConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RentAccEneAbr = rs_Detliq!almonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    End If
'
'                    'Renta total exenta y/o no gravada
'                    If EsRentaGrabadaConc Then
'                        StrSql = "SELECT dlimonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND detliq.concnro = " & RentaGrabadaConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RentaGrabada = rs_Detliq!dlimonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    Else
'                        StrSql = "SELECT almonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND  proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND acu_liq.acunro = " & RentaGrabadaConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RentaGrabada = rs_Detliq!almonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    End If
'
'                    'Rebaja por zonas extremas
'                    If EsRebajasTotConc Then
'                        StrSql = "SELECT dlimonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND detliq.concnro = " & RebajasTotConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RebajasTot = rs_Detliq!dlimonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    Else
'                        StrSql = "SELECT almonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND  proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND acu_liq.acunro = " & RebajasTotConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           RebajasTot = rs_Detliq!almonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    End If
'
'                    'Total remuneracion imponible para efectos previsionales
'                    If EsTotalRemuConc Then
'                        StrSql = "SELECT dlimonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND detliq.concnro = " & TotalRemuConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           TotalRemu = rs_Detliq!dlimonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    Else
'                        StrSql = "SELECT almonto FROM proceso "
'                        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                        StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                        StrSql = StrSql & " WHERE "
'                        If Lista_Mod <> "0" Then
'                            StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                        End If
'                        StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND  proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                        StrSql = StrSql & " AND acu_liq.acunro = " & TotalRemuConf
'                        StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                        OpenRecordset StrSql, rs_Detliq
'                        Do While Not rs_Detliq.EOF
'                           TotalRemu = rs_Detliq!almonto
'                           rs_Detliq.MoveNext
'                        Loop
'                    End If
                
                    'auxdeci1 a auxdeci6
                    '********************************************************************************
'                    For I = 0 To 5
'                        If Arrauxdeci(I, 0) = "CO" Then
'                            StrSql = "SELECT dlimonto FROM proceso "
'                            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                            StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
'                            StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                            StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                            StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                            StrSql = StrSql & " WHERE "
'                            If Lista_Mod <> "0" Then
'                                StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                            End If
'
'                            StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                            StrSql = StrSql & " AND detliq.concnro = " & Arrauxdeci(I, 1)
'                            StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                            OpenRecordset StrSql, rs_Detliq
'                            Do While Not rs_Detliq.EOF
'                               Arrauxdeci(I, 1) = rs_Detliq!dlimonto
'                               rs_Detliq.MoveNext
'                            Loop
'
'                        ElseIf Arrauxdeci(I, 0) = "AC" Then
'                            StrSql = "SELECT almonto FROM proceso "
'                            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
'                            StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'                            StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
'                            StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
'                            StrSql = StrSql & " INNER JOIN empresa ON empresa.empnro = proceso.empnro  AND  empresa.empnro = " & empresa
'                            StrSql = StrSql & " WHERE "
'                            If Lista_Mod <> "0" Then
'                                StrSql = StrSql & " tipoproc.tprocnro IN (" & Lista_Mod & ") AND "
'                            End If
'                            StrSql = StrSql & " proceso.profecini >=" & ConvFecha(PeriodoDesde) & " AND  proceso.profecfin <=" & ConvFecha(PeriodoHasta)
'                            StrSql = StrSql & " AND acu_liq.acunro = " & Arrauxdeci(I, 1)
'                            StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!Ternro
'                            OpenRecordset StrSql, rs_Detliq
'                            Do While Not rs_Detliq.EOF
'                               Arrauxdeci(I, 1) = rs_Detliq!almonto
'                               rs_Detliq.MoveNext
'                            Loop
'
'                        End If
'                    Next
                    
                    
                    '********************************************************************************
                    '********************************************************************************
                  
                  '-----------------------------------------------------------------------------------
                'Controlo errores en el empleado
                If Not HuboError Then
                   'Inserto en rep_ddjj_renta_det
                    For I = 0 To 11
                        StrSql = "INSERT INTO rep_ddjj_renta_det (bpronro, ternro, orden, Titulo, Rut, DV, TerApe, TerApe2, TerNom, TerNom2, Empnro, periodo_desde, periodo_hasta, RentaPagada, ImpUnico, "
                        StrSql = StrSql & "MayReten, RentaGrabada, Rebajas, Mes_1, Mes_2, Mes_3, Mes_4, Mes_5, Mes_6, Mes_7, Mes_8, Mes_9, Mes_10, Mes_11, Mes_12 "
                        
                        StrSql = StrSql & ",auxdeci_1,auxdeci_2,auxdeci_3,auxdeci_4,auxdeci_5,auxdeci_6"
                        StrSql = StrSql & ",numcertif,mes"
                        StrSql = StrSql & ",sueldobruto,cotizprev,factoractualiz"
                        
                        StrSql = StrSql & ") VALUES ("
                        StrSql = StrSql & NroProcesoBatch & ","
                        StrSql = StrSql & rs_Empleados!ternro & ","
                        StrSql = StrSql & Num_linea & ","
                        StrSql = StrSql & "'" & Left(Titulo, 200) & "',"
                        StrSql = StrSql & "'" & Rut & "',"
                        StrSql = StrSql & "'" & DV & "',"
                        StrSql = StrSql & "'" & Apellido & "',"
                        StrSql = StrSql & "'" & Apellido2 & "',"
                        StrSql = StrSql & "'" & NombreEmp & "',"
                        StrSql = StrSql & "'" & NombreEmp2 & "',"
                        StrSql = StrSql & empresa & ","
                        'StrSql = StrSql & ConvFecha(PeriodoDesde) & ","
                        'StrSql = StrSql & ConvFecha(PeriodoHasta) & ","
'                        StrSql = StrSql & RenTotNeta & ","
'                        StrSql = StrSql & Impuesto & ","
'                        StrSql = StrSql & Retencion & ","
'                        StrSql = StrSql & RenTotExenta & ","
'                        StrSql = StrSql & Rebajas & ","
                        
                        StrSql = StrSql & ConvFecha("01/01/" & AnioTributario) & ","
                        StrSql = StrSql & ConvFecha("31/12/" & AnioTributario) & ","
                        StrSql = StrSql & ArrRenTotNeta(I) & ","
                        StrSql = StrSql & ArrImpuesto(I) & ","
                        StrSql = StrSql & ArrRetencion(I) & ","
                        StrSql = StrSql & ArrRenTotExenta(I) & ","
                        StrSql = StrSql & ArrRebajas(I) & ","

                        For Z = 1 To 11
                            StrSql = StrSql & Meses(Z) & ","
    '                        If Meses(Z) = 0 Then
    '                            StrSql = StrSql & 0 & ","
    '                        Else
    '                            StrSql = StrSql & -1 & ","
    '                        End If
                        Next Z
                        
                        StrSql = StrSql & Meses(Z) & ","
    '                    If Meses(12) = 0 Then   'Va sin coma
    '                        StrSql = StrSql & 0
    '                    Else
    '                        StrSql = StrSql & -1
    '                    End If
    
                        For Z = 0 To 5
                           StrSql = StrSql & Arrauxdeci(Z, 0) & ","
                        Next
                        'StrSql = StrSql & Arrauxdeci(5, 1)
                        
                        'N° Certificado del empleado
                        StrSql = StrSql & Num_linea
                        
                        'Mes
                        StrSql = StrSql & "," & I + 1 & ","
                        
                        
                        
                        StrSql = StrSql & ArrSueldoBruto(I) & ","
                        StrSql = StrSql & ArrCotPrev(I) & ","
                        StrSql = StrSql & ArrFactorAct(I)
                        
                        
                        
                        StrSql = StrSql & ")"
                        Flog.writeline
                        Flog.writeline "Insertando : " & StrSql
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline
                        Flog.writeline
                    
                    Next
                    
                    
                    'Sumo el numero de linea
                    Num_linea = Num_linea + 1
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
            
            'le resto uno a la cantidad de registros para q no quede defasado
                Num_linea = Num_linea - 1
            'Grabar totalizados
            Flog.writeline "----------------------------------------"
            Flog.writeline "Grabando totalizados..."
            'Inserto en rep_ddjj_renta
            StrSql = "INSERT INTO rep_ddjj_renta (bpronro, Empnro, periodo_desde, periodo_hasta, RentaPagada, RentaPagadaAnio,  "
            StrSql = StrSql & "RentAccEneAbr, RentaGrabada, Rebajas, TotalRemun,CantRegistros "
            StrSql = StrSql & ", anio,declarante_rut,declarante_razsoc,declarante_domicilio,declarante_comuna,declarante_email,declarante_fax,declarante_tel,RepLegal_rut"
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & NroProcesoBatch & ","
            StrSql = StrSql & empresa & ","
            StrSql = StrSql & ConvFecha(PeriodoDesde) & ","
            StrSql = StrSql & ConvFecha(PeriodoHasta) & ","
            StrSql = StrSql & RentaPagada & ","
            StrSql = StrSql & RentaPagadaAnio & ","
            StrSql = StrSql & RentAccEneAbr & ","
            StrSql = StrSql & RentaGrabada & ","
            StrSql = StrSql & RebajasTot & ","
            StrSql = StrSql & TotalRemu & ","
            StrSql = StrSql & Num_linea
            
            '--
            StrSql = StrSql & ",'" & AnioTributario & "'"
            StrSql = StrSql & ",'" & RutEmpresa & "'"
            StrSql = StrSql & ",'" & razSoc & "'"
            StrSql = StrSql & ",'" & domicilio & "'"
            StrSql = StrSql & ",'" & comuna & "'"
            StrSql = StrSql & ",'" & email & "'"
            StrSql = StrSql & ",'" & fax & "'"
            StrSql = StrSql & ",'" & tel & "'"
            StrSql = StrSql & ",'" & DocReplegal & "'" 'rut rep legal
            
            
            StrSql = StrSql & ")"
            Flog.writeline
            Flog.writeline "Insertando : " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Se grabo el registro"
            Flog.writeline
                   
            
End If 'If Not HuboError




If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_CantEmpleados.State = adStateOpen Then rs_CantEmpleados.Close
If rs_Acu_liq.State = adStateOpen Then rs_Acu_liq.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Rut.State = adStateOpen Then rs_Rut.Close
If objRs3.State = adStateOpen Then objRs3.Close

Set rs_Empleados = Nothing
Set rs_CantEmpleados = Nothing
Set rs_Acu_liq = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Tercero = Nothing
Set rs_Rut = Nothing
Set objRs3 = Nothing

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

Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal anterior As Long) As Boolean
    rs.MoveNext
    If rs.EOF Then
        EsElUltimoEmpleado = True
    Else
        If rs!Empleado <> anterior Then
            EsElUltimoEmpleado = True
        Else
            EsElUltimoEmpleado = False
        End If
    End If
    rs.MovePrevious
End Function

Public Function JornadaDelEmpleado(ByVal ternro As String, ByVal Canthoras As Integer, ByVal FechaCalculo As Date) As Integer
    Dim rs_jornada As New ADODB.Recordset
    ' 1 = "P"
    ' 2 = "C"
    StrSql = "SELECT HorasDia,his_estructura.estrnro FROM his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " INNER JOIN regimenHorario ON regimenHorario.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " WHERE  his_estructura.tenro = 21 "
    StrSql = StrSql & " AND ternro = " & ternro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(FechaCalculo) & " AND ( htethasta >= " & ConvFecha(FechaCalculo) & " OR htethasta IS Null )"
    
    OpenRecordset StrSql, rs_jornada
    If Not rs_jornada.EOF Then
        If EsNulo(rs_jornada!HorasDia) Then
            JornadaDelEmpleado = 2
            Flog.writeline "Se encontró Jornada Parcial"
        Else
            If (rs_jornada!HorasDia >= Canthoras) Then
                JornadaDelEmpleado = 1
                Flog.writeline "Se encontró Jornada Completa"
            Else
                JornadaDelEmpleado = 2
                Flog.writeline "Se encontró Jornada Parcial"
            End If
        End If
    Flog.writeline "Estructura: " & rs_jornada!estrnro
    Else
        JornadaDelEmpleado = 0
        Flog.writeline "Verificar Régimen horario para el ternro: " & ternro
    End If
    
    
   rs_jornada.Close
End Function

Public Function GetConcAcu(ByVal tipo As String, confval, confval2) As Long
    'si es CO devuelve concnro si es AC devuelve confval
    Dim objConcAcu As New ADODB.Recordset
    If UCase(tipo) = "CO" Then
         StrSql = "SELECT concnro FROM concepto WHERE conccod = '" & confval2 & "'"
         OpenRecordset StrSql, objConcAcu
         If Not objConcAcu.EOF Then
           GetConcAcu = objConcAcu!concnro
         End If
         objConcAcu.Close
    Else
         GetConcAcu = confval
    End If

End Function
