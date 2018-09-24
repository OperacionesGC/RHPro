Attribute VB_Name = "MdlABMAportes"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "12/11/2011"
'Global Const UltimaModificacion = "" 'Proceso para generar los Archivos de Alta/Modificación y Baja
                                     'de Declaraciones Juaradas de La Estrella
                                     
'Global Const Version = "1.01"
'Global Const FechaModificacion = "16/11/2011"
'Global Const UltimaModificacion = "" 'JAZ - Se agregó el manejo de errores

'Global Const Version = "1.02"
'Global Const FechaModificacion = "18/11/2011"
'Global Const UltimaModificacion = "" 'JAZ - Se corrigieron errores cuando realizaba el insert en la tabla rep_jub_mov

 'Global Const Version = "1.03"
 'Global Const FechaModificacion = "12/12/2011"
 'Global Const UltimaModificacion = "" 'JAZ - Se corrigieron errores en la búsqueda del valor del concepto
 
 'Global Const Version = "1.04"
 'Global Const FechaModificacion = "14/12/2011"
 'Global Const UltimaModificacion = "" 'JAZ - Se corrigieron errores varios en la búsqueda de las categorías
 
' Global Const Version = "1.05"
' Global Const FechaModificacion = "21/03/2012" ' Zamarbide Juan - Caso CAS - 15445 - Casa Humberto Lucaioli - Error en DDJJ La Estrella
' Global Const UltimaModificacion = "" 'Se cambió el btprcnro de 315 a 320 para que coincida con lo instalado en Lucaioli
 
' Global Const Version = "1.06"
' Global Const FechaModificacion = "04/05/2012" ' Dimatz, Rafael - Cas-15621- Vestiditos S.A. - DDJJ La Estrella - Error en DDJJ La Estrella
' Global Const UltimaModificacion = "" 'Se Agrego ConFecha en la Baja, Se cambio el nombre del registro bajafec por bajfec y se cambio en el filtro de la consulta del Alta la fecha <=
                                      'Se agrego en la consulta de Insert como varchar el nrodoc, hora e identificador
                                      'Se cambio el numero de proceso a 315 y luego se dejo en 320

 'Global Const Version = "1.07"
 'Global Const FechaModificacion = "07/06/2012" ' Deluchi, Ezequiel - Cas-15621
 'Global Const UltimaModificacion = "" 'Se corrigio insert cuando falta configurar categoria
                                      'Se cambio la recuperacion de datos del concepto configurado en el confrep a la columna confval2
                                      'Se agregro progreso

 'Global Const Version = "1.08"
 'Global Const FechaModificacion = "07/08/2012" ' Deluchi, Ezequiel - Cas-15621
 'Global Const UltimaModificacion = "" 'Se corrigio calculo de la fecha de nacimiento

 'Global Const Version = "1.09"
 'Global Const FechaModificacion = "17/09/2012" ' Deluchi, Ezequiel - CAS-16949 - Horwath Argentina - Bug reporte La Estrella
 'Global Const UltimaModificacion = "" 'Se cambio el tipo del parametro grabarDatos a long donde llega el legajo del empleado.

 'Global Const Version = "1.10"
 'Global Const FechaModificacion = "02/10/2012" ' Deluchi, Ezequiel - CAS-16949 - Horwath Arg - Reporte ABM La estrella. Error en fecha de ingreso
 'Global Const UltimaModificacion = "" 'Se informa la fecha de alta de la fase segun la fecha del periodo.

 'Global Const Version = "1.11"
 'Global Const FechaModificacion = "05/10/2012" ' Deluchi, Ezequiel - CAS-16949 - Horwath Arg - Reporte ABM La estrella. Error en fecha de ingreso
 'Global Const UltimaModificacion = "" 'Correcion en la construccion de la fecha de baja, al formar el dia y el mes.
 
 'Global Const Version = "1.12"
 'Global Const FechaModificacion = "18/10/2012" ' Deluchi, Ezequiel - CAS-17303 - NGA - OLX - Bug Reporte Altas La Estrella
 'Global Const UltimaModificacion = "" 'Cambio del tipo de la variable que guardaba el valor del aporte inicial, no guardaba los valores decimales.

 'Global Const Version = "1.13"
 'Global Const FechaModificacion = "11/12/2012" ' CAS-17789 - MIMO - ERROR EN REPORTE LA ESTRELLA
 'Global Const UltimaModificacion = "" 'Si no encuentra la fase activa en el periodo, no graba los datos (antes cargaba now() en la fecha).

 'Global Const Version = "1.14"
 'Global Const FechaModificacion = "14/12/2012" ' CAS-17789 - MIMO - ERROR EN REPORTE LA ESTRELLA
 'Global Const UltimaModificacion = "" 'Si no encuentra la fase activa en el periodo, no graba los datos (antes cargaba now() en la fecha), correccion en la busqueda de la fase.

' Global Const Version = "1.15"
' Global Const FechaModificacion = "14/05/2013" ' CAS-19169 - MARKET LINE - BUG EN ABM REPORTE LA ESTRELLA
' Global Const UltimaModificacion = "" 'Overflow - Ternro as Integer!
 
 
 
'Global Const Version = "1.16"
' Global Const FechaModificacion = "28/06/2013" ' Fernandez Matias - CAS-19169 - MARKET LINE - BUG EN ABM REPORTE LA ESTRELLA.
' Global Const UltimaModificacion = "" 'Overflow - Ternro as Integer!

'Global Const Version = "1.17"
'Global Const FechaModificacion = "08/07/2013" ' Fernandez Matias - CAS-19169 - MARKET LINE - BUG EN ABM REPORTE LA ESTRELLA.
'Global Const UltimaModificacion = "Recompilacion" 'Overflow - Ternro as Integer!

Global Const Version = "1.18"
Global Const FechaModificacion = "17/11/2014" ' Borrelli Facundo - CAS-27083 - VISION BASE MARCO - BUG EN ABM REPORTE LA ESTRELLA
Global Const UltimaModificacion = "Se encripta string de conexion" 'Se agrega el modulo configuracion
                                'Se corrigen errores en tipos de datos al hacer el insert
                                'Se agregan detalles al log (parametros).
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Type TipoAltaMod
    Nro_ID As String                'Numerico long 15 -
    Nom_Ape As String               'Alfanumérico 30 - Apellido y Nombre Empleado
    Sexo As String                  'Alfanumérico 1 - F / M
    Fec_Nac As String               'Alfanumérico 6 - DDMMAA
    Tipo_Doc As String              'Numerico long 1  - 1 - DNI / LC / LE y 4 - CI / PAS
    Nro_Doc As String               'Numerico long 8  -
    Fec_Ing As String               'Numérico 6 - DDMMAA
    Agrupacion As String            'Numerico 1 -
    Categoria As String             'Alfanumerico 1 -
    Apo_Inic As Double                'Numerico 15 -
End Type
Private Type TipoBaja
    Nro_ID As String                'Numerico long 15 -
    Nom_Ape As String               'Alfanumérico 30 - Apellido y Nombre Empleado
    Tipo_Doc As String              'Numerico long 1  - 1 - DNI / LC / LE y 4 - CI / PAS
    Nro_Doc As String               'Numerico long 8  -
    Fec_Baja As String              'Numerico long 6  - DDMMAA
    Cod_Baja As Single              'Numerico long 1 - Según
End Type
Dim Reg1 As TipoAltaMod
Dim Reg2 As TipoBaja
Global IdUser As String
Global Fecha As Date
Global hora As String

'Adrián - Declaración de dos nuevos registros.
Global rs_Empresa As New ADODB.Recordset
Global rs_tipocod As New ADODB.Recordset

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : Juan A. Zamarbide
' Fecha      : 14/10/2011
' Ultima Mod.:
' Descripcion: Procedimiento Inicial
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProcesoBatch = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'            EncriptStrconexion = CBool(ArrParametros(2))
'            c_seed = ArrParametros(2)
'        Else
'            Exit Sub
'        End If
'    Else
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

    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "Exp_LaEstrella_DJABM" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "Inicio                   : " & Format(Now, FormatoInternoFecha)
    
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
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 320 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    Flog.writeline "Primer paso " & StrSql
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        Flog.writeline "Segundo paso " & StrSql
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
    
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
Fin:
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
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
    HuboError = True
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    
End Sub
Public Sub Generacion(ByVal FiltroEmpleado As String, ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Agrupado As Boolean, _
    ByVal AgrupaTE1 As Boolean, ByVal Tenro1 As Long, Estrnro1 As Long, _
    ByVal AgrupaTE2 As Boolean, ByVal Tenro2 As Long, Estrnro2 As Long, _
    ByVal AgrupaTE3 As Boolean, ByVal Tenro3 As Long, Estrnro3 As Long, NroProc As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Generación de Altas, Bajas y Modificaciones de Declaraciones Juradas de Aportes
' Autor      : Juan A. Zamarbide
' Fecha      : 14/10/2011
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------



Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Nro_Reporte As Integer
Dim Conf_Ok As Boolean
Dim ConcNro As Long
Dim Nro_Concepto As Long

Dim concepto As String
Dim Acumulador2() As String
Dim n, I As Integer
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_Categoria As New ADODB.Recordset
Dim rs_Categoria2 As New ADODB.Recordset
Dim rs_Puesto As New ADODB.Recordset
Dim rs_Doc As New ADODB.Recordset
Dim rs_Alta As New ADODB.Recordset
Dim rs_Baja As New ADODB.Recordset
Dim rs_Fase As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Rep_jub_mov As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim Fec As Date

Const ForReading = 1
Const TristateFalse = 0
Dim fExport1, fExport2
Dim fAuxiliar
Dim directorio As String
Dim Archivo1 As String
Dim Archivo2 As String
Dim Intentos As Integer
Dim carpeta
Dim tieneFaseActiva As Boolean

Dim Aux_str As String
Dim TipoCodEmpresa As String
Dim NroEmpresa As Long

'Activo el manejador de errores
On Error GoTo CE

Flog.writeline "Ingresa a Generación..."
Fec = Date
'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If
Fecha_Inicio_periodo = rs_Periodo!pliqdesde
Fecha_Fin_Periodo = rs_Periodo!pliqhasta
n = 0
'Configuracion del Reporte
Nro_Reporte = 359
'Columna 1 - Aporte Inicial - AC - Utilizado por DJAYM
'Columna 2 - Renuncia/Despido - CB - Código de Baja
'Columna 3 - Fallecimiento - CB - Código de Baja
'Columna 4 - Invalidez - CB - Código de Baja
'Columna 5 - Jubilación Oficial - CB - Código de Baja

StrSql = "SELECT * FROM confrep WHERE repnro = " & Nro_Reporte
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
Else
    Do While Not rs_Confrep.EOF
        Select Case rs_Confrep!conftipo
        Case "CO":
            concepto = rs_Confrep!confval2
            Conf_Ok = True
        Case "AB":
            n = n + 1
            ReDim Preserve Acumulador2(n)
            Acumulador2(n - 1) = rs_Confrep!confval2
        End Select
        rs_Confrep.MoveNext
    Loop
End If
If Not Conf_Ok Then
    Flog.writeline "Columna 1. El concepto no esta configurado"
    Exit Sub
End If
Flog.writeline "ConfRep OK!!!"

'Busco los procesos a evaluar
StrSql = "SELECT DISTINCT empleado.empleg, empleado.ternro, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2  "
If AgrupaTE1 Then
    StrSql = StrSql & ", te1.tenro tenro1, te1.estrnro estrnro1"
End If
If AgrupaTE2 Then
    StrSql = StrSql & ", te2.tenro tenro2, te2.estrnro estrnro2"
End If
If AgrupaTE3 Then
    StrSql = StrSql & ", te3.tenro tenro3, te3.estrnro estrnro3"
End If
StrSql = StrSql & "  FROM  Empleado "
StrSql = StrSql & " INNER JOIN his_estructura  ON his_estructura.ternro = empleado.ternro and his_estructura.tenro = 10 "
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.estrnro =" & Empresa

If AgrupaTE1 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE1 ON te1.ternro = empleado.ternro "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE2 ON te2.ternro = empleado.ternro "
End If
'
StrSql = StrSql & " INNER JOIN cabliq on cabliq.empleado = empleado.ternro and pronro in (" & NroProc & ")"

If AgrupaTE3 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE3 ON te3.ternro = empleado.ternro "
End If
StrSql = StrSql & " WHERE " & FiltroEmpleado
StrSql = StrSql & " AND empresa.estrnro =" & Empresa
StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
If AgrupaTE1 Then
    StrSql = StrSql & " AND  te1.tenro = " & Tenro1 & " AND "
    If Estrnro1 <> 0 Then
        StrSql = StrSql & " te1.estrnro = " & Estrnro1 & " AND "
    End If
    StrSql = StrSql & " (te1.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te1.htethasta) or (te1.htethasta is null)) "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " AND te2.tenro = " & Tenro2 & " AND "
    If Estrnro2 <> 0 Then
        StrSql = StrSql & " te2.estrnro = " & Estrnro2 & " AND "
    End If
    StrSql = StrSql & " (te2.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te2.htethasta) or (te2.htethasta is null))  "
End If
If AgrupaTE3 Then
    StrSql = StrSql & " AND te3.tenro = " & Tenro3 & " AND "
    If Estrnro3 <> 0 Then
        StrSql = StrSql & " te3.estrnro = " & Estrnro3 & " AND "
    End If
    StrSql = StrSql & " (te3.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te3.htethasta) or (te3.htethasta is null))"
End If
StrSql = StrSql & " ORDER BY empleado.ternro"
OpenRecordset StrSql, rs_Procesos

Flog.writeline " SQL: " & StrSql

' Busco el estrnro de la empresa
StrSql = "SELECT empnro, estrnro FROM empresa WHERE empresa.estrnro = " & Empresa
OpenRecordset StrSql, rs_Empresa

' Busco el tipo de código La Estrella de la empresa.
If rs_Empresa.EOF Then
            Flog.writeline " No existe una estructura para esta Empresa"
Else
    NroEmpresa = rs_Empresa!Empnro
    
   StrSql = "SELECT nrocod"
   StrSql = StrSql & " FROM estr_cod"
   StrSql = StrSql & " INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro"
   StrSql = StrSql & " WHERE (tipocod.tcodnro = 32)"
   StrSql = StrSql & " AND estrnro = " & rs_Empresa!Estrnro
   OpenRecordset StrSql, rs_tipocod

    If rs_tipocod.EOF Then
        Flog.writeline " No existe número de La Estrella para esta Empresa"
        TipoCodEmpresa = String(15, "0")
    Else
        If Len(rs_tipocod!nrocod) < 15 Then
            TipoCodEmpresa = rs_tipocod!nrocod & String(15 - Len(rs_tipocod!nrocod), "0")
        Else
            TipoCodEmpresa = Left(rs_tipocod!nrocod, 15)
        End If
    End If
    Flog.writeline "Tipo código Empresa = " & TipoCodEmpresa
End If

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (99 / CConceptosAProc)
Flog.writeline "Cantidad de Emplados a Procesar= " & CConceptosAProc

'Adrián - Utilizo el nro de codigo Estrella para la empresa.
Reg1.Nro_ID = TipoCodEmpresa
Reg1.Apo_Inic = 0
Reg2.Nro_ID = TipoCodEmpresa

If Not rs_Procesos.EOF Then
    Do While Not rs_Procesos.EOF
        'Incremento el progreso
        Progreso = Progreso + IncPorc

        
        'Busco Alta de Empleados en el Periodo
        'LED 02/10/2012
        Flog.writeline "Busco Alta de Empleados en el Periodo"
        StrSql = "Select * from fases where estado = -1 and empleado = " & rs_Procesos!Ternro & _
                 " and (altfec >= " & ConvFecha(Fecha_Inicio_periodo) & "and altfec <= " & ConvFecha(Fecha_Fin_Periodo) & ") "
        
        OpenRecordset StrSql, rs_Alta
        I = rs_Alta.RecordCount
        'Busco cambios en la Categoría del empleado
        Flog.writeline "Busco cambios en la Categoría del empleado = " & rs_Procesos!empleg
        StrSql = "select * from  estructura es " & _
                 " inner join his_estructura he ON es.estrnro = he.estrnro " & _
                 " inner join estr_cod ec on ec.estrnro = es.estrnro " & _
                 " Where he.Tenro = 3 " & _
                 " and ((he.htetdesde >= " & ConvFecha(Fecha_Inicio_periodo) & ") " & _
                 " and (he.htethasta <= " & ConvFecha(Fecha_Fin_Periodo) & " or he.htethasta is null)) " & _
                 " and ternro = " & rs_Procesos!Ternro & _
                 " and tcodnro = 32"
                 
        OpenRecordset StrSql, rs_Categoria
        
        If (Not rs_Alta.EOF) Or ((Not rs_Categoria.EOF) And rs_Categoria.RecordCount > 0) Then
            
            Flog.writeline "Ingreso a Procesar el empleado Legajo = " & rs_Procesos!empleg
            ' Buscar el documento
            StrSql = " SELECT tercero.*, ter_doc.tidnro, ter_doc.nrodoc FROM tercero " & _
                 " INNER JOIN ter_doc ON tercero.ternro = ter_doc.ternro " & _
                 " WHERE tercero.ternro= " & rs_Procesos!Ternro & _
                 " ORDER BY ter_doc.tidnro "
            OpenRecordset StrSql, rs_Doc
            
            'Apellido y Nombre
            Reg1.Nom_Ape = Left(rs_Procesos!terape & " " & rs_Procesos!terape2 & " " & rs_Procesos!ternom & " " & rs_Procesos!ternom2, 30)
            'Sexo
            Reg1.Sexo = IIf(rs_Doc!tersex = -1, "M", "F")
            'Fecha de Nacimiento
            If Month(rs_Doc!terfecnac) <= 9 Then
                If Day(rs_Doc!terfecnac) <= 9 Then
                    Aux_str = "0" & Day(rs_Doc!terfecnac)
                Else
                    Aux_str = Day(rs_Doc!terfecnac)
                End If
                Aux_str = Aux_str & "0" & Month(rs_Doc!terfecnac) & Right(Year(rs_Doc!terfecnac), 2)
            Else
                If Day(rs_Doc!terfecnac) <= 9 Then
                    Aux_str = "0" & Day(rs_Doc!terfecnac)
                Else
                    Aux_str = Day(rs_Doc!terfecnac)
                End If
                Aux_str = Aux_str & Month(rs_Doc!terfecnac) & Right(Year(rs_Doc!terfecnac), 2)
            End If
            Reg1.Fec_Nac = Aux_str
            'N° Documento y Tipo
            If Not rs_Doc.EOF Then
                Select Case rs_Doc!tidnro
                    Case 1, 2, 3:
                        Reg1.Tipo_Doc = "1"
                    Case 4, 5:
                        Reg1.Tipo_Doc = "4"
                    Case Else
                        Reg1.Tipo_Doc = "1"
                    End Select
                    Reg1.Nro_Doc = Format_StrNro(Left(CStr(rs_Doc!NroDoc), 8), 8, True, "0")
            Else
                Flog.writeline "Error al obtener los datos del Documento"
                Reg1.Tipo_Doc = "1"
                Reg1.Nro_Doc = "00000000"
            End If
            'Busco la última Fase Activa
            Flog.writeline "Busco la fecha de Fase Activa en el periodo."
            StrSql = " SELECT * FROM fases " & _
                     " WHERE empleado = " & rs_Procesos!Ternro & _
                     " AND ((altfec <= " & ConvFecha(Fecha_Inicio_periodo) & " AND (bajfec is null or bajfec >= " & ConvFecha(Fecha_Fin_Periodo) & " or bajfec >= " & ConvFecha(Fecha_Inicio_periodo) & ")) OR " & _
                     " (altfec >= " & ConvFecha(Fecha_Inicio_periodo) & " AND (altfec <= " & ConvFecha(Fecha_Fin_Periodo) & ")) )"

                     '" AND ((altfec <= " & ConvFecha(Fecha_Inicio_periodo) & " and (bajfec >= " & ConvFecha(Fecha_Fin_Periodo) & "  OR bajfec is null )) OR " & _
                     '" (altfec >= " & ConvFecha(Fecha_Inicio_periodo) & " and (bajfec <= " & ConvFecha(Fecha_Fin_Periodo) & " OR bajfec is null)) ) "

            OpenRecordset StrSql, rs_Fase
            

            If Not rs_Fase.EOF Then
                tieneFaseActiva = True
                Reg1.Fec_Ing = rs_Fase!altfec ' Asigno la fecha de la fase en el rango del periodo
            Else
                tieneFaseActiva = False
                Reg1.Fec_Ing = Now()
                Flog.writeline "El empleado no tiene fecha de fase en el periodo."
            End If
            'Busco el Concepto informado en el Confrep
            Flog.writeline
            StrSql = "SELECT SUM(dlimonto) monto "
            StrSql = StrSql & "  FROM detliq "
            StrSql = StrSql & " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
            StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro "
            StrSql = StrSql & " INNER JOIN concepto ON detliq.concnro = concepto.concnro "
            StrSql = StrSql & " WHERE "
            StrSql = StrSql & " proceso.pronro IN (" & NroProc & ")"
            StrSql = StrSql & " AND empleado.ternro =" & rs_Procesos!Ternro
            StrSql = StrSql & " AND concepto.conccod = '" & concepto & "'"
            
            OpenRecordset StrSql, rs_Concepto
            If Not rs_Concepto.EOF Then
                Reg1.Apo_Inic = Format(IIf(Not IsNull(rs_Concepto!Monto), rs_Concepto!Monto, 0), "0000000000000.00")
            End If
            If Not rs_Alta.EOF Then
                Do While Not rs_Alta.EOF
                    'Buscar Categoria
                    StrSql = "select * from  estructura es " & _
                             " inner join his_estructura he ON es.estrnro = he.estrnro " & _
                             " inner join estr_cod ec on ec.estrnro = es.estrnro " & _
                             " Where he.Tenro = 3 " & _
                             " and ((he.htetdesde >= " & ConvFecha(Fecha_Inicio_periodo) & ") " & _
                             " and (" & ConvFecha(Fecha_Fin_Periodo) & " <= he.htetdesde)) " & _
                             " and ternro = " & rs_Procesos!Ternro & _
                             " and tcodnro = 32"
                    Flog.writeline "Busco la Categoría correspondiente al alta"
                    OpenRecordset StrSql, rs_Categoria2
                    If Not rs_Categoria2.EOF Then 'Si le asignaron categoría al empleado
                        Reg1.Agrupacion = Mid(rs_Categoria2!nrocod, 1, 1)
                        Reg1.Categoria = Mid(rs_Categoria2!nrocod, 2, 1)
                        Fec = rs_Categoria2!htetdesde
                    
                        If tieneFaseActiva Then
                            Flog.writeline "Ingresa a Grabar"
                            'Tengo todos los datos, Escribo en tabla
                            Call GrabarDatos(FiltroEmpleado, bpronro, Nroliq, Todos_Pro, 1, rs_Procesos!empleg, rs_Procesos!Ternro, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3, Empresa)
                        Else
                            Flog.writeline "No tiene fase activa para el perido, no se grabara el empleado"
                        End If
                    Else
                        Reg1.Agrupacion = "0"
                        Reg1.Categoria = "Z"
                        Fec = Date
                        Flog.writeline "No se encontró la categoría para el Empleado = " & CStr(rs_Procesos!Ternro) & " dado de alta el " & ConvFecha(rs_Alta!altfec)
                    End If
                    rs_Alta.MoveNext
                Loop
            End If
            If Not rs_Categoria.EOF Then
                Do While Not rs_Categoria.EOF
                    If rs_Categoria!htetdesde >= Fecha_Inicio_periodo And rs_Categoria!htetdesde <= Fecha_Fin_Periodo Then
                        Reg1.Agrupacion = Mid(rs_Categoria!nrocod, 1, 1)
                        Reg1.Categoria = Mid(rs_Categoria!nrocod, 2, 1)
                        'Tengo todos los datos, Escribo en tabla
                            Flog.writeline "Ingresa a Grabar"
                        If tieneFaseActiva Then
                            Call GrabarDatos(FiltroEmpleado, bpronro, Nroliq, Todos_Pro, 1, rs_Procesos!empleg, rs_Procesos!Ternro, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3, Empresa)
                        Else
                            Flog.writeline "No tiene fase activa para el perido, no se grabara el empleado"
                        End If
                    Else
                        Reg1.Agrupacion = "0"
                        Reg1.Categoria = "Z"
                    End If
                    rs_Categoria.MoveNext
                Loop
            End If
        End If
        'Busco si tiene bajas en el período
        Flog.writeline "Busco si tiene bajas en el período"
        StrSql = "Select * from fases where estado = 0 and empleado = " & rs_Procesos!Ternro & _
                 " and (bajfec >= " & ConvFecha(Fecha_Inicio_periodo) & " and bajfec <= " & ConvFecha(Fecha_Fin_Periodo) & ") "
                        
        OpenRecordset StrSql, rs_Baja
            
        If Not rs_Baja.EOF Then
            'Datos Empleado
            'Apellido y Nombre
            Reg2.Nom_Ape = Left(rs_Procesos!terape & " " & rs_Procesos!terape2 & " " & rs_Procesos!ternom & " " & rs_Procesos!ternom2, 30)
            ' Buscar el documento
            StrSql = " SELECT tercero.*, ter_doc.tidnro, ter_doc.nrodoc FROM tercero " & _
                 " INNER JOIN ter_doc ON tercero.ternro = ter_doc.ternro " & _
                 " WHERE tercero.ternro= " & rs_Procesos!Ternro & _
                 " ORDER BY ter_doc.tidnro "
            OpenRecordset StrSql, rs_Doc
            'N° Documento y Tipo
            If Not rs_Doc.EOF Then
                Select Case rs_Doc!tidnro
                    Case 1, 2, 3:
                        Reg2.Tipo_Doc = "1"
                    Case 4, 5:
                        Reg2.Tipo_Doc = "4"
                    Case Else
                        Reg2.Tipo_Doc = "1"
                    End Select
                    Reg2.Nro_Doc = Format_StrNro(Left(CStr(rs_Doc!NroDoc), 8), 8, True, "0")
            Else
                Flog.writeline "Error al obtener los datos del Documento"
                Reg2.Tipo_Doc = "1"
                Reg2.Nro_Doc = "00000000"
            End If
            Do While Not rs_Baja.EOF
                'Fecha de Baja
                If Month(rs_Baja!bajfec) <= 9 Then
                    If Day(rs_Baja!bajfec) <= 9 Then
                        Reg2.Fec_Baja = "0" & Day(rs_Baja!bajfec)
                    Else
                        Reg2.Fec_Baja = Day(rs_Baja!bajfec)
                    End If
                    Reg2.Fec_Baja = Reg2.Fec_Baja & "0" & Month(rs_Baja!bajfec) & Right(Year(rs_Baja!bajfec), 2)
                Else
                    If Day(rs_Baja!bajfec) <= 9 Then
                        Reg2.Fec_Baja = "0" & Day(rs_Baja!bajfec)
                    Else
                        Reg2.Fec_Baja = Day(rs_Baja!bajfec)
                    End If
                    Reg2.Fec_Baja = Reg2.Fec_Baja & Month(rs_Baja!bajfec) & Right(Year(rs_Baja!bajfec), 2)
                End If
                For I = 0 To UBound(Acumulador2)
                    If Dentro(Acumulador2(I), rs_Baja!caunro) Then 'Verifica que el código de Baja esté en cada linea del confrep
                        Reg2.Cod_Baja = I + 1 'Asigno el Nº de Código que coincide con el de la fila del Confrep
                        Exit For
                    End If
                Next
                Flog.writeline "Ingresa a Grabar"
                Call GrabarDatos(FiltroEmpleado, bpronro, Nroliq, Todos_Pro, 2, rs_Procesos!empleg, rs_Procesos!Ternro, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3, Empresa)
                rs_Baja.MoveNext
            Loop
        End If
        If rs_Doc.State = adStateOpen Then rs_Doc.Close
        If rs_Categoria.State = adStateOpen Then rs_Categoria.Close
        If rs_Baja.State = adStateOpen Then rs_Baja.Close
        If rs_Alta.State = adStateOpen Then rs_Alta.Close
        If rs_Fase.State = adStateOpen Then rs_Fase.Close
        If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
        If rs_Categoria2.State = adStateOpen Then rs_Categoria2.Close
        rs_Procesos.MoveNext
        
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & FormatNumber(Progreso, 2)
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(FormatNumber(IncPorc, 2)) & "' WHERE bpronro = " & bpronro
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Progreso = " & FormatNumber(Progreso, 2)
    Loop
Else
    Flog.writeline "No existen empleados a Procesar"
End If
    
'Fin de la transaccion

If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Doc.State = adStateOpen Then rs_Doc.Close
If rs_Categoria.State = adStateOpen Then rs_Categoria.Close
If rs_Categoria2.State = adStateOpen Then rs_Categoria2.Close
If rs_Baja.State = adStateOpen Then rs_Baja.Close
If rs_Alta.State = adStateOpen Then rs_Alta.Close
If rs_Fase.State = adStateOpen Then rs_Fase.Close
If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_tipocod.State = adStateOpen Then rs_tipocod.Close
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close

Set rs_Confrep = Nothing
Set rs_Concepto = Nothing
Set rs_Detliq = Nothing
Set rs_Doc = Nothing
Set rs_Categoria = Nothing
Set rs_Categoria2 = Nothing
Set rs_Baja = Nothing
Set rs_Alta = Nothing
Set rs_Fase = Nothing
Set rs_Procesos = Nothing
Set rs_Periodo = Nothing
Set rs_tipocod = Nothing
Set rs_Empresa = Nothing

Exit Sub

CE:
    HuboError = True
    MyRollbackTrans
    
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " Ultima sql ejecutada: " & StrSql

    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
    If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
    If rs_Doc.State = adStateOpen Then rs_Doc.Close
    If rs_Categoria.State = adStateOpen Then rs_Categoria.Close
    If rs_Categoria2.State = adStateOpen Then rs_Categoria2.Close
    If rs_Baja.State = adStateOpen Then rs_Baja.Close
    If rs_Alta.State = adStateOpen Then rs_Alta.Close
    If rs_Fase.State = adStateOpen Then rs_Fase.Close
    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If rs_tipocod.State = adStateOpen Then rs_tipocod.Close
    If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
    
    Set rs_Confrep = Nothing
    Set rs_Concepto = Nothing
    Set rs_Detliq = Nothing
    Set rs_Doc = Nothing
    Set rs_Categoria = Nothing
    Set rs_Categoria2 = Nothing
    Set rs_Baja = Nothing
    Set rs_Alta = Nothing
    Set rs_Fase = Nothing
    Set rs_Procesos = Nothing
    Set rs_Periodo = Nothing
    Set rs_tipocod = Nothing
    Set rs_Empresa = Nothing

End Sub
Public Sub GrabarDatos(ByVal FiltroEmpleado As String, ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal tipo As Integer, ByVal empleg As Long, ByVal Ternro As Long, _
    ByVal AgrupaTE1 As Boolean, ByVal Tenro1 As Long, Estrnro1 As Long, _
    ByVal AgrupaTE2 As Boolean, ByVal Tenro2 As Long, Estrnro2 As Long, _
    ByVal AgrupaTE3 As Boolean, ByVal Tenro3 As Long, Estrnro3 As Long, ByVal Empresa As Long)
 
Dim rs_Estructura As New ADODB.Recordset
On Error GoTo CE2

' JAZ - Reutilizé la Tabla rep_jub_mov del reporte de Declaraciones Juradas - La Estrella dado que contenía casi todos los campos que se utilizan en este reporte
    If tipo = 1 Then 'Tipo 1 = Alta/Modificación
    Flog.writeline "Preparo el SQL a grabar Tipo = 1"
        StrSql = "INSERT INTO rep_jub_mov (bpronro, pliqnro, pronro, iduser, Fecha, empleg, ternro,"
            StrSql = StrSql & " tiporegistro, Hora, nroidentificador, tidnro, nrodoc, importe,"
            StrSql = StrSql & " apeynom, fecnac, sexo, agrupacion, categoria,empresa,"
            StrSql = StrSql & " tenro1, estrnro1, tedesc1, estrdesc1, tenro2, estrnro2, tedesc2, estrdesc2, tenro3, estrnro3, tedesc3, estrdesc3 "
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & bpronro & ","
            StrSql = StrSql & Nroliq & ","
            If Not Todos_Pro Then
                StrSql = StrSql & "'" & Left(ListaNroProc, 200) & "',"
            Else
                StrSql = StrSql & "0" & ","
            End If
            StrSql = StrSql & "'" & IIf(Reg1.Nro_ID <> "", Reg1.Nro_ID, "1111") & "',"
            StrSql = StrSql & ConvFecha(Reg1.Fec_Ing) & "," 'Corresponde al campo Fecha (reutilizado)
            StrSql = StrSql & IIf(empleg <> 0, empleg, 0) & ","
            StrSql = StrSql & IIf(Ternro <> 0, Ternro, 0) & ","
            StrSql = StrSql & tipo & ","
            'FB - StrSql = StrSql & "'horaas',"
            StrSql = StrSql & ConvFecha(CDate(Now)) & ","
            'FB - StrSql = StrSql & "'identificador',"
            StrSql = StrSql & "'" & IIf(Reg1.Nro_ID <> "", Reg1.Nro_ID, "1111") & "',"
            StrSql = StrSql & Reg1.Tipo_Doc & ","
            StrSql = StrSql & IIf(Reg1.Nro_Doc <> 0, Reg1.Nro_Doc, 0) & ","
            StrSql = StrSql & IIf(Reg1.Apo_Inic <> 0, Reg1.Apo_Inic, 0) & ","
            StrSql = StrSql & "'" & IIf(Reg1.Nom_Ape <> "", Reg1.Nom_Ape, "Falta Nombre") & "',"
            StrSql = StrSql & "'" & IIf(Reg1.Fec_Nac <> "", Reg1.Fec_Nac, "11/11/2011") & "',"
            StrSql = StrSql & "'" & IIf(Reg1.Sexo <> "", Reg1.Sexo, "1") & "',"
            StrSql = StrSql & Reg1.Agrupacion & ","
            StrSql = StrSql & "'" & Reg1.Categoria & "',"
            StrSql = StrSql & IIf(Empresa <> 0, Empresa, 0) & ","
            'Estructuras
            If AgrupaTE1 Then
                StrSql = StrSql & Tenro1 & ","
            Else
                StrSql = StrSql & 0 & ","
            End If
            StrSql = StrSql & Estrnro1 & ","
            
            'Descripcion tipo estructura
            If AgrupaTE1 Then
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & Tenro1
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'StrSql = StrSql & "'" & rs_Estructura!tedabr & "'" & ","
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!tedabr, 25, True, False) & ","
                Else
                    'StrSql = StrSql & "' '" & ","
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & Estrnro1
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!estrdabr, 25, True, False) & ","
                Else
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
            Else
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
            End If
            
            If AgrupaTE2 Then
                StrSql = StrSql & Tenro2 & ","
            Else
                StrSql = StrSql & 0 & ","
            End If
            StrSql = StrSql & Estrnro2 & ","
            
            If AgrupaTE2 Then
                'Descripcion tipo estructura
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & Tenro2
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!tedabr, 25, True, False) & ","
                Else
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & Estrnro2
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!estrdabr, 25, True, False) & ","
                Else
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
            Else
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
            End If
            
            If AgrupaTE3 Then
                StrSql = StrSql & Tenro3 & ","
            Else
                StrSql = StrSql & 0 & ","
            End If
            StrSql = StrSql & Estrnro3 & ","
            
            'Descripcion tipo estructura
            If AgrupaTE3 Then
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & Tenro3
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'StrSql = StrSql & "'" & rs_Estructura!tedabr & "'" & ","
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!tedabr, 25, True, False) & ","
                Else
                    'StrSql = StrSql & "' '" & ","
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & Estrnro3
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!estrdabr, 25, True, False)
                Else
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False)
                End If
            Else
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False)
            End If
            StrSql = StrSql & ")"
    Else 'Tipo 2 = Baja
    Flog.writeline "Preparo el SQL a Grabar Tipo = 2"
        StrSql = "INSERT INTO rep_jub_mov (bpronro,pliqnro,pronro,iduser,empleg,ternro,"
        'StrSql = StrSql & "tiporegistro, fecha, hora, nroidentificador, tidnro, nrodoc,"
        StrSql = StrSql & "tiporegistro, Fecha, Hora, nroidentificador, tidnro, nrodoc,"
        StrSql = StrSql & "apeynom,bajafec,bajacod,empresa,"
        StrSql = StrSql & "tenro1,estrnro1,tedesc1,estrdesc1,tenro2,estrnro2,tedesc2,estrdesc2,tenro3,estrnro3,tedesc3,estrdesc3 "
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & bpronro & ","
        StrSql = StrSql & Nroliq & ","
        If Not Todos_Pro Then
            StrSql = StrSql & "'" & Left(ListaNroProc, 200) & "',"
        Else
            StrSql = StrSql & "0" & ","
        End If
        StrSql = StrSql & "'" & IIf(Reg2.Nro_ID <> "", Reg2.Nro_ID, "2222") & "',"
        StrSql = StrSql & IIf(empleg <> 0, empleg, 0) & ","
        StrSql = StrSql & IIf(Ternro <> 0, Ternro, 0) & ","
        StrSql = StrSql & tipo & ","
        StrSql = StrSql & ConvFecha(CDate(Now)) & ","
        'FB - StrSql = StrSql & "'hora'" & ","
        StrSql = StrSql & ConvFecha(CDate(Now)) & ","
        'FB - StrSql = StrSql & "'identificador'" & ","
        StrSql = StrSql & "'" & IIf(Reg2.Nro_ID <> "", Reg2.Nro_ID, "2222") & "',"
        StrSql = StrSql & Reg2.Tipo_Doc & ","
        'FB - StrSql = StrSql & IIf(CLng(Reg2.Nro_Doc) <> 0, "' & Reg2.Nro_Doc & '", 0) & ","
        StrSql = StrSql & "'" & IIf(Reg2.Nro_Doc <> "", Reg2.Nro_Doc, "0") & "',"
        StrSql = StrSql & "'" & IIf(Reg2.Nom_Ape <> "", Reg2.Nom_Ape, "Falta Nombre") & "',"
        StrSql = StrSql & "'" & IIf(Reg2.Fec_Baja <> "", Reg2.Fec_Baja, "11/11/2011") & "',"
        StrSql = StrSql & IIf(CInt(Reg2.Cod_Baja) <> 0, Reg2.Cod_Baja, 1) & ","
        StrSql = StrSql & IIf(Empresa <> 0, Empresa, 0) & ","
        'Estructuras
            If AgrupaTE1 Then
                StrSql = StrSql & Tenro1 & ","
            Else
                StrSql = StrSql & 0 & ","
            End If
            StrSql = StrSql & Estrnro1 & ","
            
            'Descripcion tipo estructura
            If AgrupaTE1 Then
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & Tenro1
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!tedabr, 25, True, False) & ","
                Else
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & Estrnro1
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'StrSql = StrSql & "'" & rs_Estructura!estrdabr & "'" & ","
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!estrdabr, 25, True, False) & ","
                Else
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
            Else
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
            End If
            
            If AgrupaTE2 Then
                StrSql = StrSql & Tenro2 & ","
            Else
                StrSql = StrSql & 0 & ","
            End If
            StrSql = StrSql & Estrnro2 & ","
            
            If AgrupaTE2 Then
                'Descripcion tipo estructura
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & Tenro2
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!tedabr, 25, True, False) & ","
                Else
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & Estrnro2
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!estrdabr, 25, True, False) & ","
                Else
                    'StrSql = StrSql & "' '" & ","
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
            Else
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
            End If
            
            If AgrupaTE3 Then
                StrSql = StrSql & Tenro3 & ","
            Else
                StrSql = StrSql & 0 & ","
            End If
            StrSql = StrSql & Estrnro3 & ","
            
            'Descripcion tipo estructura
            If AgrupaTE3 Then
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & Tenro3
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!tedabr, 25, True, False) & ","
                Else
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & Estrnro3
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & FormatearParaSql(rs_Estructura!estrdabr, 25, True, False)
                Else
                    StrSql = StrSql & FormatearParaSql(" ", 25, True, False)
                End If
            Else
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False) & ","
                StrSql = StrSql & FormatearParaSql(" ", 25, True, False)
            End If
            StrSql = StrSql & ")"
    End If
    
    'Flog.writeline "Consulta = " & StrSql
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
 Set rs_Estructura = Nothing
    
Exit Sub

CE2:
    HuboError = True
    MyRollbackTrans
    
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " Ultima sql ejecutada: " & StrSql
End Sub



Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : JAZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim pliqdesde As Long
Dim pliqhasta As Long
Dim Todos_Pro As Boolean
Dim Proc_Aprob As Integer
Dim Empresa As Long
Dim FiltroEmpleados As String

Dim Tenro1 As Long
Dim Tenro2 As Long
Dim Tenro3 As Long
Dim Estrnro1 As Long
Dim Estrnro2 As Long
Dim Estrnro3 As Long

Dim AgrupaTE1 As Boolean
Dim AgrupaTE2 As Boolean
Dim AgrupaTE3 As Boolean
Dim Agrupado As Boolean

Dim tipo As Integer
'FB
Dim todos_procesos As String
'--
On Error GoTo errorparam
'Inicializacion
Agrupado = False
Tenro1 = 0
Tenro2 = 0
Tenro3 = 0
AgrupaTE1 = False
AgrupaTE2 = False
AgrupaTE3 = False

'Orden de los parametros
'filtro de empleados
'pliqdesde
'pliqhasta
'fecha desde
'fecha hasta
'Proaprob
'Lista de procesos
'Tenro1
'estrnro1
'tenro2
'estrnro2
'tenro3
'estrnro3
'Empresa empnro
'no calienta

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        FiltroEmpleados = Mid(parametros, pos1, pos2 - pos1 + 1)
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        pliqdesde = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        pliqhasta = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FechaDesde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FechaHasta = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Proc_Aprob = Mid(parametros, pos1, pos2 - pos1 + 1)
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        NroProc = Mid(parametros, pos1, pos2 - pos1 + 1)
        'FB
        todos_procesos = Mid(NroProc, 4, 1)
        'Flog.writeline " Todos los Procesos: " & todos_procesos
        'If NroProc = "0" Then
        'FB - Se busca si se seleccionaron todos los procesos.
        If todos_procesos = "0" Then
            Todos_Pro = True
        Else
            Todos_Pro = False
        End If
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Tenro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Not Tenro1 = 0 Then
            Agrupado = True
            AgrupaTE1 = True
        End If
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estrnro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Tenro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Not Tenro2 = 0 Then
            AgrupaTE2 = True
        End If
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estrnro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Tenro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Not Tenro3 = 0 Then
            AgrupaTE3 = True
        End If
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estrnro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Empresa = 0 Then
            Flog.writeline " Error: Debe seleccionar una empresa en el filtro"
            Exit Sub
        End If
    End If
End If

Flog.writeline
Flog.writeline " Parametros: " & parametros
Flog.writeline Espacios(Tabulador * 1) & "Periodo desde: " & pliqdesde
Flog.writeline Espacios(Tabulador * 1) & "Periodo hasta: " & pliqhasta
Flog.writeline Espacios(Tabulador * 1) & "Fecha desde: " & FechaDesde
Flog.writeline Espacios(Tabulador * 1) & "Fecha hasta: " & FechaHasta
Flog.writeline Espacios(Tabulador * 1) & "Proceso Aprobado: " & Proc_Aprob
Flog.writeline Espacios(Tabulador * 1) & "Nro de Proceso: " & NroProc
Flog.writeline Espacios(Tabulador * 1) & "Todos los procesos: " & Todos_Pro
Flog.writeline Espacios(Tabulador * 1) & "Tipo Estructura 1: " & Tenro1
Flog.writeline Espacios(Tabulador * 1) & "Estructura 1: " & Estrnro1
Flog.writeline Espacios(Tabulador * 1) & "Tipo Estructura 2: " & Tenro2
Flog.writeline Espacios(Tabulador * 1) & "Estructura 2: " & Estrnro2
Flog.writeline Espacios(Tabulador * 1) & "Tipo Estructura 3: " & Tenro3
Flog.writeline Espacios(Tabulador * 1) & "Estructura 3: " & Estrnro3
Flog.writeline Espacios(Tabulador * 1) & "Empresa: " & Empresa
Flog.writeline

Call Generacion(FiltroEmpleados, bpronro, pliqdesde, Todos_Pro, Proc_Aprob, Empresa, Agrupado, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3, NroProc)

Exit Sub

errorparam:
    HuboError = True
    MyRollbackTrans
    
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " Ultima sql ejecutada: " & StrSql

End Sub


Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para saber si es el ultimo empleado de la secuencia
' Autor      : JAZ
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

Public Function Dentro(ByVal Linea As String, ByVal Valor As String)
Dim arrline() As String
Dim I As Integer
Dim bingo As Boolean
bingo = False
arrline = Split(Linea, ",")

For I = 0 To UBound(arrline)
    If Valor = arrline(I) Then
        bingo = True
        Exit For
    End If
Next
Dentro = bingo
End Function
