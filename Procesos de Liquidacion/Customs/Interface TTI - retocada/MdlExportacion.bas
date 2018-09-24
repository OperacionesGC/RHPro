Attribute VB_Name = "MdlExportacion"
Option Explicit

Global Const Version = "2.00"
Global Const FechaModificacion = "15/11/2005"
Global Const UltimaModificacion = " " 'Define si es un A o M en base a la auditoria

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

Private Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

Global IdUser As String
Global Fecha As Date
Global hora As String

'Adrián - Declaración de dos nuevos registros.
Global rs_Empresa As New ADODB.Recordset
Global rs_tipocod As New ADODB.Recordset

Global Desde As Date
Global Hasta As Date
Global StrSql2 As String
Global separadorDecimales As String
Global totalImporte As Double
Global Total As Single
Global UltimaLeyenda As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 07/09/2004
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
    
    Nombre_Arch = PathFLog & "Exportacion" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline "-----------------------------------------------------------------"
    Flog.Writeline "Version = " & Version
    Flog.Writeline "Modificacion = " & UltimaModificacion
    Flog.Writeline "Fecha = " & FechaModificacion
    Flog.Writeline "-----------------------------------------------------------------"
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 0) & "PID = " & PID
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 88 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.Writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
End Sub

Public Sub Generacion(ByVal bpronro As Long, ByVal fechadesde As Date, ByVal horadesde As String, ByVal fechahasta As Date, ByVal horahasta As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de archivos de Auditoria
' Autor      : Fernando Favre
' Fecha      : 11/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0

Dim fExport
Dim fAuxiliar
Dim fExport_aux
Dim fAuxiliar_aux
Dim Directorio As String
Dim Archivo As String
Dim Archivo_aux As String
Dim Intentos As Integer
Dim carpeta
Dim separadorCampos As String
Dim separadorDecimales As String
Dim seguir As Boolean
Dim valor_traducido As String
Dim valor_campo As String
Dim ternro As Integer
Dim fecha_auditoria As Date
Dim Valor As String
Dim primero As Boolean
Dim strLinea_valor As String
Dim ted_orden_aux As Integer
Dim ted_tcampo_ant As String
Dim ted_longitud_ant As Integer
Dim ternro_ant As Integer
Dim tabextdescabr_ant As String
Dim tabextorden_ant As Integer
Dim tabexthist_ant As Boolean
Dim empleg_ant As String
Dim tipnro_ant As Integer
Dim tabextnro_ant As Integer

Dim strLinea As String
Dim strLinea_aux As String
Dim Aux_Linea As String
Dim posini As Integer
Dim cant_reg As Long
Dim conv_vcond1
Dim conv_vcond2
Dim unaSolaSqlCondicion As Boolean
Dim StrSql_1 As String
Dim StrSql_2 As String
Dim strLineaTraduccion As String
Dim buscarDatosNoAuditoria As Boolean
Dim seguir_conv_campos As Boolean
Dim cond_aux As Boolean
Dim valor2_ant As Integer
Dim seConcateno As Boolean
Dim NroSecuencia

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Auditoria As New ADODB.Recordset
Dim rs_Conv_Campos As New ADODB.Recordset
Dim rs_Tablas_Ext_Def As New ADODB.Recordset
Dim rs_Tablas_Ext As New ADODB.Recordset
Dim rs_Tablas_Pre_Trad As New ADODB.Recordset
Dim rs_Tablas_Campos_Trad As New ADODB.Recordset
Dim rs_cond1 As New ADODB.Recordset
Dim rs_cond2 As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset


On Error GoTo CE

MyBeginTrans

'Borro la tabla uxiliar que utilizo
StrSql = "DELETE FROM tablas_pre_trad "
objconnProgreso.Execute StrSql, , adExecuteNoRecords


'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 257"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
    Else
        Flog.Writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
    separadorDecimales = rs_Modelo!modsepdec
    separadorCampos = rs_Modelo!modseparador
Else
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If


'Seteo el nombre del archivo auxiliar
Archivo_aux = Directorio & "\SDaux.txt"

Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport_aux = fs.CreateTextFile(Archivo_aux, True)
If Err.Number <> 0 Then
    Flog.Writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Set fExport_aux = fs.CreateTextFile(Archivo_aux, True)
End If
'desactivo el manejador de errores
On Error GoTo CE

'--------------------------------------------------------------------------------------
'Genero el archivo que contiene la definicion de estructuras de datos
'--------------------------------------------------------------------------------------
Flog.Writeline Espacios(Tabulador * 1) & "-------------------------------------"
Flog.Writeline Espacios(Tabulador * 1) & "Exportando el archivo de estructura de datos"
Flog.Writeline

StrSql = "SELECT * FROM tablas_ext ORDER BY tabextorden"
OpenRecordset StrSql, rs_Tablas_Ext
If rs_Tablas_Ext.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontraron definiciones de campos de Tablas."
    Exit Sub
End If

cant_reg = 0
Do While Not rs_Tablas_Ext.EOF

        StrSql = "SELECT * FROM tablas_ext_def WHERE tabextnro=" & rs_Tablas_Ext!tabextnro
        StrSql = StrSql & " ORDER BY ted_orden"
        OpenRecordset StrSql, rs_Tablas_Ext_Def
        
        cant_reg = cant_reg + rs_Tablas_Ext_Def.RecordCount
        posini = 1
        Do While Not rs_Tablas_Ext_Def.EOF
                strLinea = "M4T_" & Completar_Espacios(rs_Tablas_Ext!tabextdescabr, 25)
                strLinea = strLinea & Completar_Espacios(rs_Tablas_Ext_Def!campextdesc, 25)
                strLinea = strLinea & Completar_Espacios(rs_Tablas_Ext_Def!ted_tcampo, 1)
                strLinea = strLinea & Completar_Ceros(CStr(rs_Tablas_Ext_Def!ted_longitud), 3)
                strLinea = strLinea & Completar_Ceros(CStr(rs_Tablas_Ext_Def!ted_decimales), 2)
                strLinea = strLinea & Completar_Ceros(CStr(posini), 3)
                strLinea = strLinea & "M4T_" & Completar_Espacios(rs_Tablas_Ext!tabextdescabr, 36)
                strLinea = strLinea & Completar_Espacios(rs_Tablas_Ext_Def!campextdesc, 40)
                strLinea = strLinea & "X"
                
                posini = posini + CInt(rs_Tablas_Ext_Def!ted_longitud)
                
                fExport_aux.Writeline strLinea
                
                rs_Tablas_Ext_Def.MoveNext
        Loop
        rs_Tablas_Ext_Def.Close
        
        rs_Tablas_Ext.MoveNext
Loop
rs_Tablas_Ext.Close
fExport_aux.Close


On Error Resume Next
Intentos = 0
Err.Number = 1
Do Until Err.Number = 0 Or Intentos = 10
     Err.Number = 0
     Set fExport_aux = fs.getfile(Archivo_aux)
     If fExport_aux.Size = 0 Then
         Err.Number = 1
         Intentos = Intentos + 1
     End If
Loop
On Error GoTo CE


If Not Intentos = 10 Then
    'Seteo el nombre del archivo generado
    Archivo = Directorio & "\SD" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & ".txt"
    
    'Creo el archivo final
    Set fExport = fs.CreateTextFile(Archivo, True)
    'Inserto el encabezado en el archivo
    strLinea = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & Completar_Ceros(cant_reg, 8)
    strLinea = Completar_Espacios(strLinea, 144)
    fExport.Writeline strLinea
    
    'Abro el archivo auxiliar
    Set fExport_aux = fs.OpenTextFile(Archivo_aux, ForReading, TristateFalse)
 
    Do While Not fExport_aux.AtEndOfStream
         strLinea = fExport_aux.ReadLine
         fExport.Writeline strLinea
    Loop
    fExport_aux.Close
    fExport.Close
End If


'--------------------------------------------------------------------------------------
'Genero el archivo de datos a exportar
'--------------------------------------------------------------------------------------

'Seteo el nombre del archivo de datos auxiliar
Archivo_aux = Directorio & "\DDaux.txt"

Set fExport_aux = fs.CreateTextFile(Archivo_aux, True)

'Busco las modificaciones en auditoria
StrSql = "SELECT * FROM auditoria WHERE aud_fec>=" & ConvFecha(fechadesde) & " AND aud_fec <=" & ConvFecha(fechahasta)
OpenRecordset StrSql, rs_Auditoria
If rs_Auditoria.EOF Then
    Flog.Writeline Espacios(Tabulador * 1) & "No se encontraron auditorias. No esta activada o no se produjeron cambios."
    Exit Sub
End If

' Comienzo la transaccion
Do While Not rs_Auditoria.EOF

    fecha_auditoria = Format(CDate(rs_Auditoria!aud_fec & " " & rs_Auditoria!aud_hor), FormatoInternoFecha)
    
    If (fecha_auditoria >= Desde And fecha_auditoria <= Hasta) Then
        Flog.Writeline Espacios(Tabulador * 2) & "Auditoria - Tabla: " & rs_Auditoria!aud_tabla & " Campo: " & rs_Auditoria!aud_campo & " Fecha: " & rs_Auditoria!aud_fec & " Hora: " & rs_Auditoria!aud_hor
    
        StrSql = "SELECT tablas_conv_campos.*, tablas_ext_def.ted_longitud, ted_decimales, ted_tcampo, "
        StrSql = StrSql & "ted_orden, tablas_ext.tabextnro, tabextdescabr, tabextorden, tabexthist "
        StrSql = StrSql & "FROM tablas_conv_campos "
        StrSql = StrSql & "INNER JOIN tablas_ext_def ON tablas_conv_campos.tabextdefnro =tablas_ext_def.tabextdefnro "
        StrSql = StrSql & "INNER JOIN tablas_ext ON tablas_ext_def.tabextnro =tablas_ext.tabextnro "
        StrSql = StrSql & " WHERE tabladb='" & rs_Auditoria!aud_tabla & "' AND campodb='" & rs_Auditoria!aud_campo & "' "
        StrSql = StrSql & " AND tablas_conv_campos.conv_ver_aud = -1 ORDER BY tablas_ext.tabextnro,conv_secuencia"
        OpenRecordset StrSql, rs_Conv_Campos
        
        Do While Not rs_Conv_Campos.EOF
        
            NroSecuencia = rs_Conv_Campos!conv_secuencia
            
            If evaluarCondiciones(rs_Auditoria!aud_ternro, rs_Auditoria, rs_Conv_Campos) Then
                
                ' Busco la traduccion del campo
                Valor = traduccionCampo(rs_Auditoria!aud_tabla, rs_Auditoria!aud_campo, rs_Auditoria!aud_actual, rs_Auditoria!aud_ternro, rs_Conv_Campos)
                
                ' Verifico si ya existe una modificacion anterior de este campo
                StrSql = "SELECT * FROM tablas_pre_trad WHERE tabladb = '" & rs_Conv_Campos!tabladb & "' "
                StrSql = StrSql & "AND campodb = '" & rs_Conv_Campos!campodb & "' AND ternro = " & rs_Auditoria!aud_ternro
                StrSql = StrSql & " AND tabextdefnro = " & rs_Conv_Campos!tabextdefnro & " AND conv_secuencia = " & rs_Conv_Campos!conv_secuencia
                
                OpenRecordset StrSql, rs2
                If rs2.EOF Then
                    StrSql = "INSERT INTO tablas_pre_trad (tabladb, campodb, ted_orden, tabextnro, tabextdescabr, tabextorden, tabexthist, ternro, tabextdefnro, conv_secuencia, conv_tcampo, ted_tcampo, ted_longitud, valor, valor2 ) "
                    StrSql = StrSql & "VALUES ('" & rs_Conv_Campos!tabladb & "',"
                    StrSql = StrSql & "'" & rs_Conv_Campos!campodb & "',"
                    StrSql = StrSql & rs_Conv_Campos!ted_orden & ","
                    StrSql = StrSql & rs_Conv_Campos!tabextnro & ",'"
                    StrSql = StrSql & rs_Conv_Campos!tabextdescabr & "',"
                    StrSql = StrSql & rs_Conv_Campos!tabextorden & ","
                    StrSql = StrSql & rs_Auditoria!acnro & ","
                    StrSql = StrSql & rs_Auditoria!aud_ternro & ","
                    StrSql = StrSql & rs_Conv_Campos!tabextdefnro & ","
                    StrSql = StrSql & rs_Conv_Campos!conv_secuencia & ",'"
                    StrSql = StrSql & rs_Conv_Campos!conv_tcampo & "','"
                    StrSql = StrSql & rs_Conv_Campos!ted_tcampo & "',"
                    StrSql = StrSql & rs_Conv_Campos!ted_longitud & ",'"
                    StrSql = StrSql & Valor & "',"
                    StrSql = StrSql & rs_Auditoria!aud_empresa & ")"
                Else
                    StrSql = "UPDATE tablas_pre_trad SET "
                    StrSql = StrSql & " valor = '" & Valor & "',"
                    StrSql = StrSql & " valor2 = " & rs_Auditoria!aud_empresa & " "
                    StrSql = StrSql & "WHERE tabladb = '" & rs_Conv_Campos!tabladb & "' "
                    StrSql = StrSql & "AND campodb = '" & rs_Conv_Campos!campodb & "' AND ternro = " & rs_Auditoria!aud_ternro
                    StrSql = StrSql & " AND tabextdefnro = " & rs_Conv_Campos!tabextdefnro & " AND conv_secuencia = " & rs_Conv_Campos!conv_secuencia
                End If
                rs2.Close
                
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
                Call traduccionCamposNoVinculantes(rs_Auditoria, rs_Conv_Campos)
                
            End If
            
            rs_Conv_Campos.MoveNext
        Loop
        rs_Conv_Campos.Close
        
    End If
    'Siguiente registro de audiroria
    
    rs_Auditoria.MoveNext
Loop


'--------------------------------------------------------------------------------------------
' Realizo la segunda pasada para determinar las conversiones definitivas
'--------------------------------------------------------------------------------------------
StrSql = "SELECT tablas_pre_trad.*, ter_tip.tipnro "
StrSql = StrSql & " FROM tablas_pre_trad "
StrSql = StrSql & " INNER JOIN ter_tip ON tablas_pre_trad.ternro=ter_tip.ternro "
StrSql = StrSql & " ORDER BY tablas_pre_trad.ternro, tablas_pre_trad.tabextorden, tablas_pre_trad.ted_orden, tablas_pre_trad.conv_secuencia"
OpenRecordset StrSql, rs_Tablas_Pre_Trad
ternro = 0
primero = True
strLinea_aux = ""
strLinea = ""
ted_tcampo_ant = ""
strLinea_valor = ""
ted_longitud_ant = 0
cant_reg = 0
If Not rs_Tablas_Pre_Trad.EOF Then
    Do While Not rs_Tablas_Pre_Trad.EOF
        
        seConcateno = False
        If ted_orden_aux = rs_Tablas_Pre_Trad!ted_orden And ternro_ant = rs_Tablas_Pre_Trad!ternro And tabextnro_ant = rs_Tablas_Pre_Trad!tabextnro Then
            strLinea_aux = strLinea_aux & rs_Tablas_Pre_Trad!Valor
        Else
            If Not primero Then
                If ted_tcampo_ant = "N" Then
                    strLinea_valor = strLinea_valor & Completar_Ceros(strLinea_aux, ted_longitud_ant)
                Else
                    strLinea_valor = strLinea_valor & Completar_Espacios(strLinea_aux, ted_longitud_ant)
                End If
                
                If Not rs2.EOF Then
                    rs2.MoveNext
                End If
                
                seConcateno = True
            End If
            
            strLinea_aux = rs_Tablas_Pre_Trad!Valor
            
        End If
        

        '-----------------------------------------------------
        ' Armo el pre-encabezado
        '-----------------------------------------------------
        If (ternro_ant <> rs_Tablas_Pre_Trad!ternro Or tabextnro_ant <> rs_Tablas_Pre_Trad!tabextnro) And Not primero Then
            If Not seConcateno And Not rs2.EOF Then
                rs2.MoveNext
                If ted_tcampo_ant = "N" Then
                    strLinea_valor = strLinea_valor & Completar_Ceros(strLinea_aux, ted_longitud_ant)
                Else
                    strLinea_valor = strLinea_valor & Completar_Espacios(UCase(strLinea_aux), ted_longitud_ant)
                End If
            End If
            
            Do While Not rs2.EOF
                ' Busco la traduccion del campo
                strLineaTraduccion = buscarTraduccionCampo(rs2!tabextorden, rs2!tabextnro, rs2!tabextdefnro, ternro_ant, valor2_ant)
        
                If rs2!ted_tcampo = "N" Then
                    strLinea_valor = strLinea_valor & Completar_Ceros(strLineaTraduccion, rs2!ted_longitud)
                Else
                    strLinea_valor = strLinea_valor & Completar_Espacios(UCase(strLineaTraduccion), rs2!ted_longitud)
                End If
                    
                rs2.MoveNext
            Loop
            
            strLinea = Completar_Espacios(UCase(tabextdescabr_ant), 25)
            strLinea = strLinea & Completar_Ceros(tabextorden_ant, 4)
            If CInt(tabexthist_ant) = 1 Then
                strLinea = strLinea & "A"
            Else
                If CInt(tabexthist_ant) = 2 Then
                    strLinea = strLinea & "M"
                Else
                    strLinea = strLinea & " "
                End If
            End If
            
            ' Busco el legajo del empleado
            Select Case tipnro_ant
                Case 1:
                    StrSql = "SELECT empleg FROM empleado WHERE ternro=" & ternro_ant
                Case 3:
                    StrSql = "SELECT empleg FROM empleado "
                    StrSql = StrSql & " INNER JOIN familiar ON empleado.ternro=familiar.empleado "
                    StrSql = StrSql & "WHERE familiar.ternro = " & ternro_ant
            End Select
            OpenRecordset StrSql, rs_Empleados
            empleg_ant = ""
            If Not rs_Empleados.EOF Then
                empleg_ant = rs_Empleados!empleg
            End If
            rs_Empleados.Close
            
            strLinea = strLinea & Completar_Ceros(empleg_ant, 10)
            
            strLinea_valor = Completar_Espacios(strLinea_valor, 250)
            strLinea = strLinea & UCase(strLinea_valor)
            
            fExport_aux.Writeline strLinea
                
            cant_reg = cant_reg + 1
            strLinea_valor = ""
        End If
            
            
        '-----------------------------------------------------
        ' Si no existe auditoria, se ingresan los espacios
        '-----------------------------------------------------
        If tabextorden_ant <> rs_Tablas_Pre_Trad!tabextorden Or ternro_ant <> rs_Tablas_Pre_Trad!ternro Then
            StrSql = "SELECT tablas_ext_def.*, tablas_ext.tabextorden "
            StrSql = StrSql & "FROM tablas_ext_def "
            StrSql = StrSql & "INNER JOIN tablas_ext ON tablas_ext_def.tabextnro = tablas_ext.tabextnro "
            StrSql = StrSql & "WHERE tablas_ext.tabextdescabr = '" & rs_Tablas_Pre_Trad!tabextdescabr & "' "
            StrSql = StrSql & "ORDER BY ted_orden"
            If rs2.State = adStateOpen Then rs2.Close
            OpenRecordset StrSql, rs2
        End If
        
        '-----------------------------------------------------
        ' Guardo valores anteriores
        '-----------------------------------------------------
        ternro_ant = rs_Tablas_Pre_Trad!ternro
        tabextdescabr_ant = rs_Tablas_Pre_Trad!tabextdescabr
        tabextorden_ant = rs_Tablas_Pre_Trad!tabextorden
        tabexthist_ant = rs_Tablas_Pre_Trad!tabexthist
        tipnro_ant = rs_Tablas_Pre_Trad!tipnro
        valor2_ant = rs_Tablas_Pre_Trad!valor2
                
        ted_tcampo_ant = rs_Tablas_Pre_Trad!ted_tcampo
        ted_longitud_ant = rs_Tablas_Pre_Trad!ted_longitud
        ted_orden_aux = rs_Tablas_Pre_Trad!ted_orden
        tabextnro_ant = rs_Tablas_Pre_Trad!tabextnro
            
        rs_Tablas_Pre_Trad.MoveNext

        seguir_conv_campos = False
        
        If Not rs2.EOF Then
            If rs2!ted_orden < CInt(ted_orden_aux) Then
                seguir_conv_campos = True
            End If
            Do While seguir_conv_campos
                ' Busco la traduccion del campo
                strLineaTraduccion = buscarTraduccionCampo(rs2!tabextorden, rs2!tabextnro, rs2!tabextdefnro, ternro_ant, valor2_ant)
    
                If rs2!ted_tcampo = "N" Then
                    strLinea_valor = strLinea_valor & Completar_Ceros(strLineaTraduccion, rs2!ted_longitud)
                Else
                    strLinea_valor = strLinea_valor & Completar_Espacios(UCase(strLineaTraduccion), rs2!ted_longitud)
                End If
                
                rs2.MoveNext
                If rs2.EOF Then
                    seguir_conv_campos = False
                Else
                    If rs2!ted_orden >= CInt(ted_orden_aux) Then
                        seguir_conv_campos = False
                    End If
                End If
            Loop
        End If
        
        primero = False
        
    Loop
    
    strLinea = Completar_Espacios(UCase(tabextdescabr_ant), 25)
    strLinea = strLinea & Completar_Ceros(tabextorden_ant, 4)
    If CInt(tabexthist_ant) = 1 Then
        strLinea = strLinea & "A"
    Else
        If CInt(tabexthist_ant) = 2 Then
            strLinea = strLinea & "M"
        Else
            strLinea = strLinea & " "
        End If
    End If
    
    Select Case tipnro_ant
        Case 1:
            StrSql = "SELECT empleg FROM empleado WHERE ternro=" & ternro_ant
        Case 3:
            StrSql = "SELECT empleg FROM empleado "
            StrSql = StrSql & " INNER JOIN familiar ON empleado.ternro=familiar.empleado "
            StrSql = StrSql & "WHERE familiar.ternro = " & ternro_ant
    End Select
    OpenRecordset StrSql, rs_Empleados
    empleg_ant = ""
    If Not rs_Empleados.EOF Then
        empleg_ant = rs_Empleados!empleg
    End If
    rs_Empleados.Close
    
    strLinea = strLinea & Completar_Ceros(empleg_ant, 10)
    
    If ted_tcampo_ant = "N" Then
        strLinea_valor = strLinea_valor & Completar_Ceros(strLinea_aux, ted_longitud_ant)
    Else
        strLinea_valor = strLinea_valor & Completar_Espacios(UCase(strLinea_aux), ted_longitud_ant)
    End If
            
    '-----------------------------------------------------
    ' Si no existe auditoria, se ingresan los lugares definidos para cada campo
    '-----------------------------------------------------
    If Not rs2.EOF Then
        rs2.MoveNext
        Do Until rs2.EOF
            strLineaTraduccion = buscarTraduccionCampo(rs2!tabextorden, rs2!tabextnro, rs2!tabextdefnro, ternro_ant, valor2_ant)

            If rs2!ted_tcampo = "N" Then
                strLinea_valor = strLinea_valor & Completar_Ceros(strLineaTraduccion, rs2!ted_longitud)
            Else
                strLinea_valor = strLinea_valor & Completar_Espacios(UCase(strLineaTraduccion), rs2!ted_longitud)
            End If
                
            rs2.MoveNext
        Loop
    End If
    
    strLinea_valor = Completar_Espacios(strLinea_valor, 250)
    strLinea = strLinea & UCase(strLinea_valor)
    
    fExport_aux.Writeline strLinea
    
    cant_reg = cant_reg + 1
    
End If

rs_Tablas_Pre_Trad.Close


'Cierro el archivo creado
fExport_aux.Close


On Error Resume Next
Intentos = 0
Err.Number = 1
Do Until Err.Number = 0 Or Intentos = 10
     Err.Number = 0
     Set fExport_aux = fs.getfile(Archivo_aux)
     If fExport_aux.Size = 0 Then
         Err.Number = 1
         Intentos = Intentos + 1
     End If
Loop
On Error GoTo CE


If Not Intentos = 10 Then
    'Seteo el nombre del archivo generado
    Archivo = Directorio & "\DD" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & ".txt"
    
    'Creo el archivo final
    Set fExport = fs.CreateTextFile(Archivo, True)
    'Inserto el encabezado en el archivo
    strLinea = ""
    strLinea = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & Completar_Ceros(cant_reg, 8) & Completar_Espacios(strLinea, 9) & "0000"
    strLinea = Completar_Espacios(strLinea, 144)
    fExport.Writeline strLinea
    
    'Abro el archivo auxiliar
    Set fExport_aux = fs.OpenTextFile(Archivo_aux, ForReading, TristateFalse)
 
    Do While Not fExport_aux.AtEndOfStream
        strLinea = ""
        strLinea = fExport_aux.ReadLine
        fExport.Writeline strLinea
    Loop
    fExport_aux.Close
    fExport.Close
End If

'Fin de la transaccion
MyCommitTrans


If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_Auditoria.State = adStateOpen Then rs_Auditoria.Close
If rs_Conv_Campos.State = adStateOpen Then rs_Conv_Campos.Close
If rs_Tablas_Ext_Def.State = adStateOpen Then rs_Tablas_Ext_Def.Close
If rs_Tablas_Ext.State = adStateOpen Then rs_Tablas_Ext.Close
If rs_Tablas_Pre_Trad.State = adStateOpen Then rs_Tablas_Pre_Trad.Close
If rs_Tablas_Campos_Trad.State = adStateOpen Then rs_Tablas_Campos_Trad.Close
If rs2.State = adStateOpen Then rs2.Close

Set rs_Modelo = Nothing
Set rs_Auditoria = Nothing
Set rs_Conv_Campos = Nothing
Set rs_Tablas_Ext_Def = Nothing
Set rs_Tablas_Ext = Nothing
Set rs_Tablas_Pre_Trad = Nothing
Set rs_Tablas_Campos_Trad = Nothing
Set rs2 = Nothing


Exit Sub

CE:
    Flog.Writeline " Error: " & Err.Description
'    Resume Next
    HuboError = True
    MyRollbackTrans

    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    If rs_Auditoria.State = adStateOpen Then rs_Auditoria.Close
    If rs_Conv_Campos.State = adStateOpen Then rs_Conv_Campos.Close
    If rs_Tablas_Ext_Def.State = adStateOpen Then rs_Tablas_Ext_Def.Close
    If rs_Tablas_Ext.State = adStateOpen Then rs_Tablas_Ext.Close
    If rs_Tablas_Pre_Trad.State = adStateOpen Then rs_Tablas_Pre_Trad.Close
    If rs_Tablas_Campos_Trad.State = adStateOpen Then rs_Tablas_Campos_Trad.Close
    
    Set rs_Modelo = Nothing
    Set rs_Auditoria = Nothing
    Set rs_Conv_Campos = Nothing
    Set rs_Tablas_Ext_Def = Nothing
    Set rs_Tablas_Ext = Nothing
    Set rs_Tablas_Pre_Trad = Nothing
    Set rs_Tablas_Campos_Trad = Nothing
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
Dim Separador As String

Dim fechadesde As Date
Dim horadesde As String
Dim fechahasta As Date
Dim horahasta As String

'Orden de los parametros
'fechadesde
'horadesde
'fechahasta
'horahasta

    Separador = "@"
    ' Levanto cada parametro por separado
    If Not IsNull(parametros) Then
        If Len(parametros) >= 1 Then
            pos1 = 1
            pos2 = InStr(pos1, parametros, Separador) - 1
            fechadesde = Mid(parametros, pos1, pos2 - pos1 + 1)
            
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, Separador) - 1
            horadesde = Mid(parametros, pos1, pos2 - pos1 + 1)
            Desde = Format(CDate(fechadesde & " " & horadesde), FormatoInternoFecha)
    
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, Separador) - 1
            fechahasta = Mid(parametros, pos1, pos2 - pos1 + 1)
            
            pos1 = pos2 + 2
            horahasta = Mid(parametros, pos1)
            
            Hasta = Format(CDate(fechahasta & " " & horahasta), FormatoInternoFecha)
            Flog.Writeline Espacios(Tabulador) & "Se evaluaran las auditorias generadas entre el " & Desde & " y el " & Hasta & "."
        Else
            Flog.Writeline Espacios(Tabulador) & "Los parametros del proceso estan mal configurados. El separador es @."
        End If
    Else
        Flog.Writeline Espacios(Tabulador) & "No se encontraron los parametros del proceso."
    End If
    
    Call Generacion(bpronro, fechadesde, horadesde, fechahasta, horahasta)
End Sub
Private Sub traduccionCamposNoVinculantes(ByRef rs_Aud As Recordset, ByRef rs_Conv_Camp As Recordset)
'--------------------------------------------------------------------------------
'  Descripción: Traduce los campos que no son vinculantes a la tabla auditoria
'  Autor: Fernando Favre
'  Fecha: 10/09/2005
'-------------------------------------------------------------------------------
Dim Valor As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs_Condicion As New ADODB.Recordset
Dim Formato As String
Dim resultado As Boolean
Dim sql As String
Dim Tabla As String
Dim Sql_Ant As String
Dim seguir As Boolean
Dim Sqlnueva As String

StrSql = "SELECT tablas_ext.*, tablas_ext_def.* "
StrSql = StrSql & "FROM tablas_ext "
StrSql = StrSql & "INNER JOIN tablas_ext_def ON tablas_ext_def.tabextnro = tablas_ext.tabextnro "
StrSql = StrSql & "WHERE tablas_ext_def.tabextdefnro=" & rs_Conv_Camp!tabextdefnro & " AND tabextdefnro NOT IN (SELECT tabextdefnro FROM tablas_conv_campos)"
StrSql = StrSql & "ORDER BY tablas_ext_def.ted_orden"
OpenRecordset StrSql, rs
Do Until rs.EOF
    Valor = rs!ted_vdefecto
    
    If rs!ted_tcampo = "N" Then
       Valor = Completar_Ceros(Valor, rs!ted_longitud)
    Else
       Valor = Completar_Espacios(Valor, rs!ted_longitud)
    End If
    
    ' Verifico si ya existe una modificacion anterior de este campo
    StrSql = "SELECT * FROM tablas_pre_trad WHERE tabextnro = " & rs!tabextnro & " AND ternro = " & rs_Aud!aud_ternro & " "
    StrSql = StrSql & " AND tabextdefnro = " & rs!tabextdefnro
    OpenRecordset StrSql, rs2
    
    If rs2.EOF Then
        StrSql = "INSERT INTO tablas_pre_trad (tabladb, campodb, ted_orden, tabextnro, tabextdescabr, tabextorden, tabexthist, ternro, tabextdefnro, conv_secuencia, conv_tcampo, ted_tcampo, ted_longitud, valor ) "
        StrSql = StrSql & "VALUES ('" & rs_Conv_Camp!tabladb & "',"
        StrSql = StrSql & "'" & rs_Conv_Camp!campodb & "',"
        StrSql = StrSql & rs!ted_orden & ","
        StrSql = StrSql & rs!tabextnro & ",'"
        StrSql = StrSql & rs!tabextdescabr & "',"
        StrSql = StrSql & rs!tabextorden & ","
        StrSql = StrSql & rs!tabexthist & ","
        StrSql = StrSql & rs_Aud!aud_ternro & ","
        StrSql = StrSql & rs!tabextdefnro & ","
        StrSql = StrSql & "null,'"
        StrSql = StrSql & rs!ted_tcampo & "','"
        StrSql = StrSql & rs!ted_tcampo & "',"
        StrSql = StrSql & rs!ted_longitud & ",'"
        StrSql = StrSql & Valor & "')"
    Else
        StrSql = "UPDATE tablas_pre_trad SET "
        StrSql = StrSql & " valor = '" & Valor & "' "
        StrSql = StrSql & "WHERE tabextnro = " & rs!tabextnro & " AND ternro = " & rs_Aud!aud_ternro & " "
        StrSql = StrSql & " AND tabextdefnro = " & rs!tabextdefnro
    End If
    rs2.Close

    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs.MoveNext
Loop
rs.Close

StrSql = "SELECT tablas_ext.*, tablas_ext_def.*, tablas_conv_campos.* "
StrSql = StrSql & "FROM tablas_ext "
StrSql = StrSql & "INNER JOIN tablas_ext_def ON tablas_ext_def.tabextnro = tablas_ext.tabextnro "
StrSql = StrSql & "INNER JOIN tablas_conv_campos ON tablas_ext_def.tabextdefnro = tablas_conv_campos.tabextdefnro "
StrSql = StrSql & "WHERE tablas_ext_def.tabextdefnro=" & rs_Conv_Camp!tabextdefnro & " AND tablas_conv_campos.conv_ver_aud = 0 "
StrSql = StrSql & "ORDER BY tablas_ext_def.ted_orden, conv_secuencia"
OpenRecordset StrSql, rs
Do Until rs.EOF
    
    If evaluarCondiciones(rs_Aud!aud_ternro, rs_Aud, rs) Then
    
        resultado = False
        
        sql = "SELECT * "
        sql = sql & "FROM tablas_conv_tab "
        sql = sql & "INNER JOIN tablas_conv_campos ON tablas_conv_tab.tabextconvnro = tablas_conv_campos.tabextconvnro "
        sql = sql & "INNER JOIN tablas_ext_def ON tablas_ext_def.tabextdefnro = tablas_conv_campos.tabextdefnro "
        sql = sql & "WHERE tablas_ext_def.tabextnro=" & rs!tabextnro & " AND tablas_ext_def.ted_orden =" & rs!ted_orden & " ORDER BY ordencond"
                
        OpenRecordset sql, rs4
        Tabla = rs!tabladb
        Sql_Ant = ""
        sql = ""
        seguir = True
        resultado = True
        sql = "SELECT " & rs!campodb & " FROM " & rs!tabladb & " WHERE 1=1  "
        Do While seguir And Not rs4.EOF
               
            If CInt(rs4!basedato) = -1 Then
                On Error Resume Next
                Valor = rs_Aud.Fields("" & rs4!valorcond & "")
                If Err.Number <> 0 Then
                    Flog.Writeline Espacios(Tabulador * 3) & "Error en la referencia del campo: " & rs4!valorcond & " de la " & rs4!ordencond & " condicion para la tabla " & rs_Aud!aud_tabla & " y el campo " & rs_Aud!aud_campo
                    Flog.Writeline
                    
                    GoTo con_Error
                End If
                On Error GoTo con_Error
            Else
                Valor = rs4!valorcond
            End If
            
            Select Case rs4!tipovalor
                Case 0:
                    ' number
                    Valor = Valor
                Case 1:
                    ' date
                    Valor = ConvFecha(Valor)
                Case 2:
                    ' string
                    Valor = "'" & Valor & "'"
            End Select
            
            
            If Tabla = rs4!tabcond Then
                ' Si la tabla de la condicion anterior es sobre la misma tabla,
                ' entonces las condiciones se agrupan en el WHERE
                sql = sql & " " & rs4!conectorcond & " " & rs4!campcond & rs4!opcond & Valor
                Sqlnueva = False
            Else
                Flog.Writeline Espacios(Tabulador * 3) & "Error: La Tabla: " & rs_Aud!aud_tabla & " y el campo " & rs_Aud!aud_campo & " en las condiciones hacen referencia a otra tabla."
                Flog.Writeline
                
                GoTo con_Error
            End If
            
            Tabla = rs4!tabcond
            
            rs4.MoveNext
            Sql_Ant = sql
            If rs4.EOF Then
                seguir = False
            End If
        Loop
           
        If sql <> "" Then
            ' Desactivo el manejador de errores
            On Error Resume Next
            OpenRecordset sql, rs_Condicion
            If Err.Number <> 0 Then
                Flog.Writeline Espacios(Tabulador * 3) & "Error en la Condicion: " & sql & " para la tabla " & rs_Aud!aud_tabla & " y el campo " & rs_Aud!aud_campo
                Flog.Writeline
            
                GoTo con_Error
            Else
                Valor = rs_Condicion(0)
            End If
            On Error GoTo con_Error
        
        End If
        
        Valor = traduccionCampo(rs!tabladb, rs!campodb, Valor, rs_Aud!aud_ternro, rs)

        ' Verifico si ya existe una modificacion anterior de este campo
        StrSql = "SELECT * FROM tablas_pre_trad WHERE tabladb = '" & rs!tabladb & "' AND campodb = '" & rs!campodb & "' "
        StrSql = StrSql & " AND ternro = " & rs_Aud!aud_ternro & " AND tabextdefnro = " & rs!tabextdefnro & " AND conv_secuencia = " & rs!conv_secuencia
        OpenRecordset StrSql, rs2
        
        If rs2.EOF Then
            StrSql = "INSERT INTO tablas_pre_trad (tabladb, campodb, ted_orden, tabextnro, tabextdescabr, tabextorden, tabexthist, ternro, tabextdefnro, conv_secuencia, conv_tcampo, ted_tcampo, ted_longitud, valor ) "
            StrSql = StrSql & "VALUES ('" & rs!tabladb & "',"
            StrSql = StrSql & "'" & rs!campodb & "',"
            StrSql = StrSql & rs!ted_orden & ","
            StrSql = StrSql & rs!tabextnro & ",'"
            StrSql = StrSql & rs!tabextdescabr & "',"
            StrSql = StrSql & rs!tabextorden & ","
            StrSql = StrSql & rs!tabexthist & ","
            StrSql = StrSql & rs_Aud!aud_ternro & ","
            StrSql = StrSql & rs!tabextdefnro & ","
            StrSql = StrSql & "null,'"
            StrSql = StrSql & rs!ted_tcampo & "','"
            StrSql = StrSql & rs!ted_tcampo & "',"
            StrSql = StrSql & rs!ted_longitud & ",'"
            StrSql = StrSql & Valor & "')"
        Else
            StrSql = "UPDATE tablas_pre_trad SET "
            StrSql = StrSql & " valor = '" & Valor & "' "
            StrSql = StrSql & "WHERE tabladb = '" & rs!tabladb & "' AND campodb = '" & rs!campodb & "' "
            StrSql = StrSql & " AND ternro = " & rs_Aud!aud_ternro & " AND tabextdefnro = " & rs!tabextdefnro & " AND conv_secuencia = " & rs!conv_secuencia
        End If
        rs2.Close
        
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rs.MoveNext
    End If
    
Loop
rs.Close

Exit Sub
con_Error:
   Flog.Writeline " Error: " & Err.Description
'   Resume Next
   HuboError = True

End Sub

Private Function traduccionCampo(ByVal Tabla As String, ByVal campo As String, ByVal Valor As String, ByVal ternro As Integer, ByRef rs_Conv As Recordset)
'--------------------------------------------------------------------------------
'  Descripción: devuelve la traduccion del campo
'  Autor: Fernando Favre
'  Fecha: 13/05/2005
'-------------------------------------------------------------------------------
Dim Valor_aux As String
Dim sql As String
Dim rs As New ADODB.Recordset
Dim rs_Trad_Campos As New ADODB.Recordset
Dim Formato As String
Dim str_pos1 As Integer
Dim str_aux As String
Dim str_pos2 As Integer

    
    Valor_aux = Valor
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    ' CASOS PARTICULARES
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    Select Case rs_Conv!tabextorden
        Case "2010":
            'Caso particular de Escolaridad
            Select Case rs_Conv!ted_orden
            Case 7:
                ' Busco el cod. externo del nivel de estudio
                sql = "SELECT nivcodext FROM nivest "
                sql = sql & "INNER JOIN estudio_actual ON estudio_actual.nivnro = nivest.nivnro "
                sql = sql & "WHERE estudio_actual.ternro = " & ternro
                OpenRecordset sql, rs
                If Not rs.EOF Then
                    Valor_aux = rs(0)
                Else
                    Valor_aux = ""
                End If
                rs.Close
            End Select
        Case "10"
            Select Case rs_Conv!ted_orden
            Case 13:
                str_pos1 = InStr(Valor_aux, "-")
                If str_pos1 <> 0 Then
                    str_aux = Mid(Valor_aux, 1, str_pos1 - 1)
                    str_pos2 = InStr(4, Valor_aux, "-")
                    If str_pos2 <> 0 Then
                        str_aux = str_aux & Mid(Valor_aux, str_pos1 + 1, str_pos2 - str_pos1 - 1)
                        str_aux = str_aux & Mid(Valor_aux, str_pos2 + 1)
                    Else
                        str_aux = str_aux & Mid(Valor_aux, str_pos1 + 1)
                    End If
                    Valor_aux = str_aux
                End If
            Case 17:
                ' Busco los tres primeros caracteres
                If Valor = 14 Then
                    Valor_aux = "EUA"
                Else
                    sql = "SELECT paisdesc FROM pais "
                    sql = sql & "WHERE paisnro = " & Valor
                    OpenRecordset sql, rs
                    If Not rs.EOF Then
                        Valor_aux = rs(0)
                        Valor_aux = Mid(Valor_aux, 1, 3)
                    Else
                        Valor_aux = ""
                    End If
                End If
                rs.Close
            Case 20:
                ' Busco los tres primeros caracteres
                If Valor = 16 Then
                    Valor_aux = "EUA"
                Else
                    sql = "SELECT nacionaldes FROM nacionalidad "
                    sql = sql & "WHERE nacionalnro = " & Valor
                    OpenRecordset sql, rs
                    If Not rs.EOF Then
                        Valor_aux = rs(0)
                        Valor_aux = Mid(Valor_aux, 1, 3)
                    Else
                        Valor_aux = ""
                    End If
                End If
                rs.Close
            End Select
        Case "1050"
            Select Case rs_Conv!ted_orden
            Case 6:
                ' Busco los tres primeros caracteres
                sql = "SELECT estrcodext FROM estructura "
                sql = sql & "WHERE estrnro = " & Valor
                OpenRecordset sql, rs
                If Not rs.EOF Then
                    Valor_aux = rs(0)
                    If Not EsNulo(Valor_aux) Then
                        Valor_aux = CInt(Valor_aux)
                        If Valor_aux > 0 And Valor_aux < 10 Then
                            Valor_aux = "L00" & Valor_aux
                        ElseIf Valor_aux >= 10 And Valor_aux <= 51 Then
                            Valor_aux = "L0" & Valor_aux
                        ElseIf (Valor_aux >= 101 And Valor_aux <= 128) Then
                            Valor_aux = "A" & Valor_aux
                        ElseIf (Valor_aux >= 302 And Valor_aux <= 399) Then
                            Valor_aux = "F" & Valor_aux
                        Else
                            Valor_aux = ""
                        End If
                    Else
                        Valor_aux = ""
                    End If
                Else
                    Valor_aux = ""
                End If
                rs.Close
            End Select
    End Select

    Select Case rs_Conv!conv_tcampo
        Case "D":
'            If rs_Conv!conv_tconv = "FH" Then
'                Valor_aux = Format(Valor, "ddmmyyyy")
'            Else
                Valor_aux = Format(Valor, "mm/dd/yyyy")
'            End If
        Case "N":
            Formato = String(rs_Conv!ted_longitud - rs_Conv!ted_decimales, "#")
            If Not EsNulo(rs_Conv!ted_decimales) Then
                If CInt(rs_Conv!ted_decimales) > 0 Then
                    Formato = Formato + "." + String(rs_Conv!ted_decimales, "0")
                End If
            End If
            Valor_aux = Format(Valor, Formato)
    End Select
    
    Select Case rs_Conv!conv_tconv
        Case "UN", "UK":
            Valor_aux = Valor_Defecto(rs_Conv!tabextdefnro)
            
        Case "CV", "CK":
            StrSql = "SELECT ttd_traducido FROM tablas_trad_campos WHERE ttd_valor = '" & Valor_aux & "' "
            StrSql = StrSql & "AND tabladb='" & Tabla & "' AND campodb='" & campo & "'"
            OpenRecordset StrSql, rs_Trad_Campos
            If Not rs_Trad_Campos.EOF Then
                Valor_aux = rs_Trad_Campos!ttd_traducido
            Else
                Flog.Writeline Espacios(Tabulador * 4) & "Warning. No se encontro la traducción para el valor: " & Valor
                Flog.Writeline
                    
                Valor_aux = ""
            End If
            rs_Trad_Campos.Close
            
        Case "FH", "FE":
            Valor_aux = Format(Valor, "yyyymmdd")
'        Case Else
'            Valor_aux = Valor_Defecto(rs_Conv!tabextdefnro)
    End Select
    
    If Not EsNulo(rs_Conv!conv_bini) Then
        If Not EsNulo(rs_Conv!conv_bfin) Then
            Valor_aux = Mid(Valor, rs_Conv!conv_bini, rs_Conv!conv_bfin - rs_Conv!conv_bini)
        Else
            Valor_aux = Mid(Valor, rs_Conv!conv_bini)
        End If
    End If
                            
    traduccionCampo = Valor_aux
    
    Set rs_Trad_Campos = Nothing

End Function
Private Function Completar_Espacios(ByVal Str As String, ByVal cant As Long) As String
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve un string de longitud determinada en cant,
'               completado con espacios en blanco
'  Autor: Fernando Favre
'  Fecha: 13/05/2005
'-------------------------------------------------------------------------------
    Dim long_str As Long
    Dim str_aux As String
    
    long_str = Len(Str)
    str_aux = ""
    If long_str >= cant Then
        str_aux = Left(Str, cant)
    Else
        str_aux = Str & Space(cant - long_str)
    End If
    
    Completar_Espacios = str_aux
    
End Function

Private Function Completar_Ceros(ByVal num As String, ByVal cant As Integer) As String
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve un string de longitud determinada en cant,
'               completado con 0 adelante del num
'  Autor: Fernando Favre
'  Fecha: 13/05/2005
'-------------------------------------------------------------------------------
    Dim long_str As Long
    Dim str_aux As String
    Dim i As Integer
    
    long_str = Len(num)
    If long_str >= cant Then
        str_aux = Right(num, cant)
    Else
        For i = 1 To cant - long_str
            str_aux = str_aux & "0"
        Next
        str_aux = str_aux & num
    End If
    
    Completar_Ceros = str_aux
     
End Function
Private Function Valor_Defecto(ByVal tabextdefnro As Integer) As String
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve un string con el valor por defecto definido para ese campo
'  Autor: Fernando Favre
'  Fecha: 13/05/2005
'-------------------------------------------------------------------------------
    Dim Str As String
    Dim rs As New ADODB.Recordset
    
    Str = "SELECT tabextdescabr, campextdesc, ted_vdefecto "
    Str = Str & "FROM tablas_ext_def INNER JOIN tablas_ext ON tablas_ext_def.tabextnro = tablas_ext.tabextnro "
    Str = Str & "WHERE tabextdefnro=" & tabextdefnro
    OpenRecordset Str, rs
    If Not rs.EOF Then
        Valor_Defecto = rs!ted_vdefecto
    Else
        Valor_Defecto = ""
        Flog.Writeline Espacios(Tabulador * 2) & "----------------------------------------------------------------------"
        Flog.Writeline Espacios(Tabulador * 2) & "No se encontro el valor por defecto para el campo " & rs!campextdesc & " de la Tabla " & rs!tabextdescabr
        Flog.Writeline
    End If
    rs.Close

End Function

Private Function evaluarCondiciones(ByVal ternro As Integer, ByRef rs_Aud As Recordset, ByRef rs_Conv As Recordset)
            
Dim sql As String
Dim rs As New ADODB.Recordset
Dim rs_Condicion As New Recordset
Dim Valor
Dim Tabla As String
Dim Sqlnueva As Boolean
Dim Sql_Ant As String
Dim resultado As Boolean
Dim seguir As Boolean
Dim terceroValido As Boolean

On Error GoTo con_Error

terceroValido = True
sql = "SELECT tipnro FROM ter_tip WHERE ternro = " & ternro
OpenRecordset sql, rs
If Not rs.EOF Then
    ' Verifico que el tercero sea un empleado y la tabla no sea FAMILIARES
    If CInt(rs!tipnro) = 1 And CInt(rs_Conv!tabextorden) = "2000" Then
        terceroValido = False
    End If
    ' Verifico que el tercero sea un familiar y la tabla no sea Empleados
    If CInt(rs!tipnro) = 3 And CInt(rs_Conv!tabextorden) = "10" Then
        terceroValido = False
    End If
End If
rs.Close

resultado = False
If terceroValido Then
Select Case rs_Conv!tabextorden
    Case "2010":
        Select Case rs_Conv!ted_orden
        Case 7:
            terceroValido = False
            ' Busco el cod. externo del nivel de estudio
            sql = "SELECT familiar.ternro FROM estudio_actual "
            sql = sql & "INNER JOIN familiar ON estudio_actual.ternro=familiar.ternro "
            sql = sql & "WHERE estudio_actual.ternro = " & ternro & " AND familiar.parenro=2"
            OpenRecordset sql, rs
            resultado = Not rs.EOF
            rs.Close
        End Select
End Select
End If

If terceroValido Then
    sql = "SELECT * "
    sql = sql & "FROM tablas_conv_tab "
    sql = sql & "INNER JOIN tablas_conv_campos ON tablas_conv_tab.tabextconvnro = tablas_conv_campos.tabextconvnro "
    sql = sql & "INNER JOIN tablas_ext_def ON tablas_ext_def.tabextdefnro = tablas_conv_campos.tabextdefnro "
    sql = sql & "WHERE tablas_ext_def.tabextnro=" & rs_Conv!tabextnro & " AND tablas_ext_def.ted_orden =" & rs_Conv!ted_orden & " ORDER BY ordencond"
            
    OpenRecordset sql, rs
    Tabla = ""
    Sql_Ant = ""
    sql = ""
    seguir = True
    resultado = True
    Do While seguir And Not rs.EOF
        
        If CInt(rs!basedato) = -1 Then
            On Error Resume Next
            Valor = rs_Aud.Fields("" & rs!valorcond & "")
            If Err.Number <> 0 Then
                Flog.Writeline Espacios(Tabulador * 3) & "Error en la referencia del campo: " & rs!valorcond & " de la " & rs!ordencond & " condicion para la tabla " & rs_Aud!aud_tabla & " y el campo " & rs_Aud!aud_campo
                Flog.Writeline
                
                GoTo con_Error
            End If
            On Error GoTo con_Error
        Else
            Valor = rs!valorcond
        End If
        
        Select Case rs!tipovalor
            Case 0:
                ' number
                Valor = Valor
            Case 1:
                ' date
                Valor = ConvFecha(Valor)
            Case 2:
                ' string
                Valor = "'" & Valor & "'"
        End Select
        
        
        If Tabla = rs!tabcond Then
            ' Si la tabla de la condicion anterior es sobre la misma tabla,
            ' entonces las condiciones se agrupan en el WHERE
            sql = sql & " " & rs!conectorcond & " " & rs!campcond & rs!opcond & Valor
            Sqlnueva = False
        Else
            ' Si la tabla de la condicion anterior es distinta a la tabla,
            ' entonces creo una nueva SQL
            sql = "SELECT * FROM " & rs!tabcond & " WHERE " & rs!campcond & rs!opcond & Valor
            Sqlnueva = True
        End If
        
        ' Si es una sola sql, la ejecuto directamente
        If Sqlnueva Then
            If Sql_Ant <> "" Then
                ' Desactivo el manejador de errores
                On Error Resume Next
                OpenRecordset Sql_Ant, rs_Condicion
                If Err.Number <> 0 Then
                    Flog.Writeline Espacios(Tabulador * 3) & "Error en la Condicion: " & Sql_Ant & " para la tabla " & rs_Aud!aud_tabla & " y el campo " & rs_Aud!aud_campo
                    Flog.Writeline
                
                    GoTo con_Error
                End If
                On Error GoTo con_Error
            
                resultado = resultado And Not rs_Condicion.EOF
                
                If Not resultado Then
                    seguir = False
                End If
            End If
        End If
        
        Tabla = rs!tabcond
        
        rs.MoveNext
        Sql_Ant = sql
        If rs.EOF Then
            seguir = False
        End If
    Loop
    
    If sql <> "" Then
        ' Desactivo el manejador de errores
        On Error Resume Next
        OpenRecordset sql, rs_Condicion
        If Err.Number <> 0 Then
            Flog.Writeline Espacios(Tabulador * 3) & "Error en la Condicion: " & sql & " para la tabla " & rs_Aud!aud_tabla & " y el campo " & rs_Aud!aud_campo
            Flog.Writeline
        
            GoTo con_Error
        End If
        On Error GoTo con_Error
    
        resultado = resultado And Not rs_Condicion.EOF
    End If
End If

evaluarCondiciones = resultado

Exit Function

con_Error:
   Flog.Writeline " Error: " & Err.Description
   HuboError = True
End Function
Private Function buscarTraduccionCampo(ByVal tabextorden As String, ByVal tabextnro As Integer, ByVal tabextdefnro As Integer, ByVal aud_ternro As Integer, ByVal valor2 As Integer)
'--------------------------------------------------------------------------------
'  Descripción: devuelve la traduccion del campo
'  Autor: Fernando Favre
'  Fecha: 03/10/2005
'-------------------------------------------------------------------------------
Dim Valor_aux As String
Dim rs_Trad_Campos As New ADODB.Recordset
Dim Formato As String
            
Dim sql As String
Dim sql_e As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs_Condicion As New Recordset
Dim Valor
Dim Tabla As String
Dim Sqlnueva As Boolean
Dim Sql_Ant As String
Dim resultado As Boolean
Dim seguir As Boolean
Dim terceroValido As Boolean
Dim campcond_aux As String
Dim tabcond_aux As String
Dim noConsiderar
Dim Fin As Boolean
Dim cont As Integer
Dim str_pos1 As Integer
Dim str_aux As String
Dim str_pos2 As Integer
                
On Error GoTo con_Error

terceroValido = True
sql = "SELECT tipnro FROM ter_tip WHERE ternro = " & aud_ternro
OpenRecordset sql, rs
If Not rs.EOF Then
    ' Verifico que el tercero sea un empleado y la tabla no sea FAMILIARES
    If CInt(rs!tipnro) = 1 And CInt(tabextorden) = "2000" Then
        terceroValido = False
    End If
    ' Verifico que el tercero sea un familiar y la tabla no sea Empleados
    If CInt(rs!tipnro) = 3 And CInt(tabextorden) = "10" Then
        terceroValido = False
    End If
End If
rs.Close

resultado = False
If terceroValido Then
    sql = "SELECT * "
    sql = sql & "FROM tablas_conv_tab "
    sql = sql & "INNER JOIN tablas_conv_campos ON tablas_conv_tab.tabextconvnro = tablas_conv_campos.tabextconvnro "
    sql = sql & "INNER JOIN tablas_ext_def ON tablas_ext_def.tabextdefnro = tablas_conv_campos.tabextdefnro "
    sql = sql & "WHERE tablas_ext_def.tabextnro=" & tabextnro & " AND tablas_ext_def.tabextdefnro =" & tabextdefnro & " ORDER BY ordencond"
            
    OpenRecordset sql, rs
    sql = ""
    seguir = True
    resultado = True
    Do While seguir And Not rs.EOF
        
        noConsiderar = False
        If CInt(rs!basedato) = -1 Then
            Select Case UCase(CStr(rs!valorcond))
                Case "AUD_TERNRO":
                    Valor = aud_ternro
                Case "AUD_EMPRESA":
'                    If casosEspeciles(tabextorden, ted_orden, aud_ternro, resultado) Then
'                        Valor = resultado
                    'Cableado para encontrar el legajo del empleado en el caso de FAMILIARES
                    Select Case tabextorden
                    Case "2000":
                        Select Case rs!ted_orden
                        Case 2:
                            ' Busco el legajo del empleado del cual es familiar.
                            sql_e = "SELECT empleado.ternro FROM empleado "
                            sql_e = sql_e & "INNER JOIN familiar ON familiar.empleado = empleado.ternro "
                            sql_e = sql_e & "WHERE familiar.ternro = " & aud_ternro
                            OpenRecordset sql_e, rs1
                            If Not rs1.EOF Then
                                Valor = rs1(0)
                            End If
                            rs1.Close
                        Case Else
                            Valor = valor2
                        End Select
                    Case "2010"
                        Select Case rs!ted_orden
                        Case 2:
                            ' Busco el legajo del empleado del cual es familiar.
                            sql_e = "SELECT empleado.ternro FROM empleado "
                            sql_e = sql_e & "INNER JOIN familiar ON familiar.empleado = empleado.ternro "
                            sql_e = sql_e & "WHERE familiar.ternro = " & aud_ternro
                            OpenRecordset sql_e, rs1
                            If Not rs1.EOF Then
                                Valor = rs1(0)
                            End If
                            rs1.Close
                        Case Else
                            Valor = valor2
                        End Select
                    Case Else
                        Valor = valor2
                    End Select
                Case "AUD_ACTUAL":
                    ' Cuando estoy buscando el valor, no se deben considerar estos casos
                    noConsiderar = True
                Case Else:
                    Flog.Writeline Espacios(Tabulador * 3) & "Error. No esta definida la traduccion para el campo de la tabla " & rs!valorcond & " de la tabla auditoria "
                    Flog.Writeline
                    
                    GoTo con_Error
            End Select
        Else
            Valor = rs!valorcond
        End If
        
        If Not noConsiderar Then
            Select Case rs!tipovalor
                Case 0:
                    ' number
                    Valor = Valor
                Case 1:
                    ' date
                    Valor = ConvFecha(Valor)
                Case 2:
                    ' string
                    Valor = "'" & Valor & "'"
            End Select
            
            sql = sql & " " & rs!conectorcond & " " & rs!campcond & rs!opcond & Valor
            
            campcond_aux = rs!campodb
            tabcond_aux = rs!tabladb
        
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    
    
    If sql = "" Then
        Valor_aux = Valor_Defecto(tabextdefnro)
    Else
        Valor_aux = ""
        sql = "SELECT " & campcond_aux & " FROM " & tabcond_aux & " WHERE 1=1 " & sql
        ' Desactivo el manejador de errores
        On Error Resume Next
        OpenRecordset sql, rs_Condicion
        If Err.Number <> 0 Then
            Flog.Writeline Espacios(Tabulador * 3) & "Error en la Condicion: " & sql & " en buscar traduccion del campo sin auditoria"
            GoTo con_Error
        Else
            Valor_aux = rs_Condicion(0)
        End If
        On Error GoTo con_Error
    
    End If



    '----------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------
    ' CASOS PARTICULARES
    '----------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------
    Select Case tabextorden
        Case "2000":
            Select Case tabextdefnro
            Case 49:
                ' Se calcula el ordinal
                sql_e = "SELECT ternro FROM familiar "
                sql_e = sql_e & "WHERE empleado IN (SELECT empleado FROM familiar "
                sql_e = sql_e & "               WHERE familiar.ternro = " & aud_ternro & ")"
                OpenRecordset sql_e, rs1
                Fin = False
                cont = 1
                Do Until rs1.EOF Or Fin
                    If rs1(0) = aud_ternro Then
                        Fin = True
                    Else
                        cont = cont + 1
                    End If
                    rs1.MoveNext
                Loop
                Valor_aux = cont
                rs1.Close
            End Select
        '------------------------------------------------------------------------------
        Case "2010":
            Select Case tabextdefnro
            Case 39:
                ' Se calcula el ordinal
                sql_e = "SELECT ternro FROM familiar "
                sql_e = sql_e & "WHERE empleado IN (SELECT empleado FROM familiar "
                sql_e = sql_e & "               WHERE familiar.ternro = " & aud_ternro & ")"
                OpenRecordset sql_e, rs1
                Fin = False
                cont = 1
                Do Until rs1.EOF Or Fin
                    If rs1(0) = aud_ternro Then
                        Fin = True
                    Else
                        cont = cont + 1
                    End If
                    rs1.MoveNext
                Loop
                Valor_aux = cont
                rs1.Close
            End Select
        '------------------------------------------------------------------------------
        Case "10"
            Select Case tabextdefnro
            Case 27:
                str_pos1 = InStr(Valor_aux, "-")
                If str_pos1 <> 0 Then
                    str_aux = Mid(Valor_aux, 1, str_pos1 - 1)
                    str_pos2 = InStr(4, Valor_aux, "-")
                    If str_pos2 <> 0 Then
                        str_aux = str_aux & Mid(Valor_aux, str_pos1 + 1, str_pos2 - str_pos1 - 1)
                        str_aux = str_aux & Mid(Valor_aux, str_pos2 + 1)
                    Else
                        str_aux = str_aux & Mid(Valor_aux, str_pos1 + 1)
                    End If
                    Valor_aux = str_aux
                End If
            Case 31:
                ' Busco los tres primeros caracteres
                If Valor = 14 Then
                    Valor_aux = "EUA"
                Else
                    sql = "SELECT paisdesc FROM pais "
                    sql = sql & "WHERE paisnro = " & Valor_aux
                    OpenRecordset sql, rs
                    If Not rs.EOF Then
                        Valor_aux = rs(0)
                        Valor_aux = Mid(Valor_aux, 1, 3)
                    Else
                        Valor_aux = ""
                    End If
                End If
                rs.Close
            Case 34:
                ' Busco los tres primeros caracteres
                If Valor = 16 Then
                    Valor_aux = "EUA"
                Else
                    sql = "SELECT nacionaldes FROM nacionalidad "
                    sql = sql & "WHERE nacionalnro = " & Valor_aux
                    OpenRecordset sql, rs
                    If Not rs.EOF Then
                        Valor_aux = rs(0)
                        Valor_aux = Mid(Valor_aux, 1, 3)
                    Else
                        Valor_aux = ""
                    End If
                End If
                rs.Close
            End Select
        '------------------------------------------------------------------------------
        Case "1050"
            Select Case tabextdefnro
            Case 139:
                ' Busco los tres primeros caracteres
                sql = "SELECT estrdext FROM estructura "
                sql = sql & "WHERE estrnro = " & Valor_aux
                OpenRecordset sql, rs
                If Not rs.EOF Then
                    Valor_aux = rs(0)
                    If Not EsNulo(Valor_aux) Then
                        Valor_aux = CInt(Valor_aux)
                        If Valor_aux > 0 And Valor_aux < 10 Then
                            Valor_aux = "L00" & Valor_aux
                        End If
                        If Valor_aux >= 10 And Valor_aux <= 51 Then
                            Valor_aux = "L0" & Valor_aux
                        End If
                        If (Valor_aux >= 101 And Valor_aux <= 128) Or (Valor_aux >= 302 And Valor_aux <= 399) Then
                            Valor_aux = "L" & Valor_aux
                        Else
                            Valor_aux = ""
                        End If
                    Else
                        Valor_aux = ""
                    End If
                Else
                    Valor_aux = ""
                End If
                rs.Close
            End Select
        '------------------------------------------------------------------------------
        Case "1080"
            Select Case tabextdefnro
            Case 115:
                ' Busco los 6 primeros caracteres, despues de haber descartado 2 caracteres
                sql = "SELECT estrdabr FROM estructura "
                sql = sql & "WHERE estrnro = " & Valor_aux
                OpenRecordset sql, rs
                If Not rs.EOF Then
                    Valor_aux = rs(0)
                    Valor_aux = Mid(Valor_aux, 3, 6)
                    
                    Select Case Valor_aux
                    Case "000000", "000001": Valor_aux = "PRESID"
                    Case "000002", "000003", "000004", "000005": Valor_aux = "DIRGRAL"
                    Case "000006": Valor_aux = "SACME"
                    Case "000040": Valor_aux = "ASCOR"
                    Case Else:
                        If Not EsNulo(Valor_aux) Then
                            Valor_aux = CInt(Valor_aux)
                            If Valor_aux >= 5000 And Valor_aux <= 5999 Then
                                Valor_aux = "ASLEG1"
                            ElseIf Valor_aux >= 6000 And Valor_aux <= 6999 Then
                                Valor_aux = "AUDIT"
                            ElseIf Valor_aux >= 7000 And Valor_aux <= 7999 Then
                                Valor_aux = "RELINS"
                            ElseIf Valor_aux >= 8000 And Valor_aux <= 8999 Then
                                Valor_aux = "DIRGRAL"
                            ElseIf Valor_aux >= 10000 And Valor_aux <= 19999 Then
                                Valor_aux = "ADMIN"
                            ElseIf Valor_aux >= 20000 And Valor_aux <= 20999 Then
                                Valor_aux = "ASREG"
                            ElseIf Valor_aux >= 23000 And Valor_aux <= 23999 Then
                                Valor_aux = "DIRGRAL"
                            ElseIf Valor_aux >= 24000 And Valor_aux <= 24999 Then
                                Valor_aux = "DIRGRAL"
                            ElseIf Valor_aux >= 25000 And Valor_aux <= 25999 Then
                                Valor_aux = "DIRGRAL"
                            ElseIf Valor_aux >= 26000 And Valor_aux <= 26999 Then
                                Valor_aux = "AUDIT"
                            ElseIf Valor_aux >= 27000 And Valor_aux <= 27999 Then
                                Valor_aux = "RELINS"
                            ElseIf Valor_aux >= 29000 And Valor_aux <= 29999 Then
                                Valor_aux = "SACME"
                            ElseIf Valor_aux >= 30000 And Valor_aux <= 39999 Then
                                Valor_aux = "TECSYS1"
                            ElseIf Valor_aux >= 40000 And Valor_aux <= 49999 Then
                                Valor_aux = "CONTROL"
                            ElseIf Valor_aux >= 60000 And Valor_aux <= 69999 Then
                                Valor_aux = "RRHH"
                            ElseIf Valor_aux >= 70000 And Valor_aux <= 79999 Then
                                Valor_aux = "COMER1"
                            ElseIf Valor_aux >= 80000 And Valor_aux <= 89999 Then
                                Valor_aux = "DISTRI"
                            ElseIf Valor_aux >= 90000 And Valor_aux <= 99999 Then
                                Valor_aux = "TECNI"
                            Else
                                Valor_aux = ""
                            End If
                        Else
                            Valor_aux = ""
                        End If
                    End Select
                Else
                    Valor_aux = ""
                End If
                rs.Close
            End Select
    End Select




    sql = "SELECT * "
    sql = sql & "FROM tablas_conv_tab "
    sql = sql & "INNER JOIN tablas_conv_campos ON tablas_conv_tab.tabextconvnro = tablas_conv_campos.tabextconvnro "
    sql = sql & "INNER JOIN tablas_ext_def ON tablas_ext_def.tabextdefnro = tablas_conv_campos.tabextdefnro "
    sql = sql & "WHERE tablas_ext_def.tabextnro=" & tabextnro & " AND tablas_ext_def.tabextdefnro =" & tabextdefnro & " ORDER BY ordencond"
            
    OpenRecordset sql, rs
    
    If Not rs.EOF Then
        
        Select Case rs!conv_tcampo
            Case "D":
    '            If rs!conv_tconv = "FH" Then
    '                Valor_aux = Format(Valor_aux, "ddmmyyyy")
    '            Else
                    Valor_aux = Format(Valor_aux, "mm/dd/yyyy")
    '            End If
            Case "N":
                Formato = String(rs!ted_longitud - rs!ted_decimales, "#")
                If Not EsNulo(rs!ted_decimales) Then
                    If CInt(rs!ted_decimales) > 0 Then
                        Formato = Formato + "." + String(rs!ted_decimales, "0")
                    End If
                End If
                Valor_aux = Format(Valor_aux, Formato)
        End Select
        
        Select Case rs!conv_tconv
            Case "UN", "UK":
                Valor_aux = Valor_Defecto(rs!tabextdefnro)
                
            Case "CV", "CK":
                StrSql = "SELECT ttd_traducido FROM tablas_trad_campos "
                StrSql = StrSql & "INNER JOIN tablas_conv_campos ON tablas_trad_campos.tabladb = tablas_conv_campos.tabladb AND tablas_trad_campos.campodb=tablas_conv_campos.campodb "
                StrSql = StrSql & "WHERE ttd_valor = '" & Valor_aux & "' AND tablas_conv_campos.tabextdefnro =" & tabextdefnro
                OpenRecordset StrSql, rs_Trad_Campos
                If Not rs_Trad_Campos.EOF Then
                    Valor_aux = rs_Trad_Campos!ttd_traducido
                Else
                    Flog.Writeline Espacios(Tabulador * 4) & "Warning. No se encontro la traducción para el valor: " & Valor_aux
                    Flog.Writeline
                        
                    Valor_aux = ""
                End If
                rs_Trad_Campos.Close
                
            Case "FH", "FE":
                Valor_aux = Format(Valor_aux, "yyyymmdd")
'            Case Else
'                Valor_aux = Valor_Defecto(tabextdefnro)
        End Select
        
        If Not EsNulo(rs!conv_bini) Then
            If Not EsNulo(rs!conv_bfin) Then
                Valor_aux = Mid(Valor_aux, rs!conv_bini, rs!conv_bfin - rs!conv_bini)
            Else
                Valor_aux = Mid(Valor_aux, rs!conv_bini)
            End If
        End If
    End If
                            
    buscarTraduccionCampo = Valor_aux
    
End If

Exit Function

con_Error:
   Flog.Writeline " Error: " & Err.Description
'   Resume Next
   HuboError = True
End Function

'Function casosEspeciles(ByVal tabextorden As String, ByVal ted_orden As Integer, ByVal ternro As Integer, ByRef resultado As String)
' Dim sql As String
' Dim rs As New ADODB.Recordset
 
'    Select Case tabextorden
'        Case "2000":
            'FAMILIARES
            'Cableado para encontrar el legajo del empleado
'            If rs!ted_orden = 2 Then
'                ' Busco el legajo del empleado del cual es familiar.
'                sql = "SELECT empleado.ternro FROM empleado "
'                sql = sql & "INNER JOIN familiar ON familiar.empleado = empleado.ternro "
'                sql = sql & "WHERE familiar.ternro = " & ternro
'                OpenRecordset sql, rs
'                If Not rs.EOF Then
'                    resultado = rs(0)
'                End If
'                rs.Close
'        Case "30":
            'FASES_ALTA
'                sql = "SELECT alfec FROM fases WHERE empleado =" & ternro & " AND sueldo=-1"
'                OpenRecordset sql, rs
'                If Not rs.EOF Then
'                    resultado = rs(0)
'                End If
'                rs.Close
'    End Select
'End Function
