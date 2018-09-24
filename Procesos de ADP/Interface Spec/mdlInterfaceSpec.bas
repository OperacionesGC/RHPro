Attribute VB_Name = "mdlInterfaceSpec"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "22/02/2013" ' Sebastian Stremel - CAS-18151 - Akzo - Interfaces Spec

'Const Version = "1.1"
'Const FechaVersion = "17/04/2013" ' Sebastian Stremel - CAS-18151 - Akzo - Interfaces Spec -
'Se busca para nivel1, nivel2, y nivel3 la descripcion abreviada del tipo estr 6,5 y 8
'Se cambio descripcion de la version 1.0 ya que erra erronea.

'Const Version = "1.2"
'Const FechaVersion = "02/05/2013" ' Sebastian Stremel - CAS-18151 - Akzo - Interfaces Spec -
'Se corrige la busqueda de nivel1, nivel2, y nivel3 la descripcion abreviada del tipo estr 6,5 y 8

'Const Version = "1.3"
'Const FechaVersion = "20/05/2013" ' Sebastian Stremel - CAS-18151 - Akzo - Interfaces Spec -
'Si el nro de tarjeta es menor de 16 caracteres entonces se completa con ceros a la izquierda
'se graba el nro de legajo en el campo matricula

'Const Version = "1.4"
'Const FechaVersion = "26/11/2014" ' Sebastian Stremel - CAS-28267 - NGA - Citricus - Error Carga Registracion en SPEC MAnager
'Se cambia el formato de fecha segun sea oracle o sql la base del sistema externo a RHPro


'Const Version = "1.5"
'Const FechaVersion = "05/12/2014" ' Sebastian Stremel - CAS-28267 - NGA - Citricus - Error Carga Registracion en SPEC MAnager [Entrega 2]
'Se cambia el formato de fecha segun sea oracle o sql la base del sistema externo a RHPro

'Const Version = "1.6"
'Const FechaVersion = "15/04/2015" ' Sebastian Stremel - CAS-28267 - CAS-29507 - Citricos - SPEC - Modificaciones Interface SPEC
'Se agregan 3 estructuras configurables correspondientes a Grupo A,B,C

Const Version = "1.7"
Const FechaVersion = "15/05/2015" ' Sebastian Stremel - CAS-29507 - Citricos - SPEC - Modificaciones Interface SPEC [Entrega 2]
'Se modifica el proceso para que siempre inserte el numero de tarjeta en caso de existir y ademas inserta solo los empleados
'que tienen una estructura configurada por confrep.

'---------------------------------------------------------------------------------------------------------------------------------------------
Dim dirsalidas As String
Dim usuario As String
Dim Incompleto As Boolean
'-------------------------------------------------------------------------------------------------
'Conexion Externa
'-------------------------------------------------------------------------------------------------
Global ExtConn As New ADODB.Connection
Global ExtConnOra As New ADODB.Connection
Global ExtConnAccess As New ADODB.Connection
Global ExtConnAccess2 As New ADODB.Connection
Global ConnLE As New ADODB.Connection
Global Usa_LE As Boolean
Global Misma_BD As Boolean





Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de la interface SPEC.
' Autor      : Sebastian Stremel
' Fecha      : 19/02/2013
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim arrParametros


    strCmdLine = Command()
    arrParametros = Split(strCmdLine, " ", -1)
    If UBound(arrParametros) > 1 Then
        If IsNumeric(arrParametros(0)) Then
            NroProcesoBatch = arrParametros(0)
            Etiqueta = arrParametros(1)
            EncriptStrconexion = CBool(arrParametros(2))
            c_seed = arrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(arrParametros) > 0 Then
            If IsNumeric(arrParametros(0)) Then
                NroProcesoBatch = arrParametros(0)
                Etiqueta = arrParametros(1)
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

    Nombre_Arch = PathFLog & "InterfaceSpec-" & NroProcesoBatch & ".log"
    
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
    
   
    If App.PrevInstance Then
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Pendiente', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Hay una instancia previa del proceso ejecutando, se pone el proceso en estado pendiente."
        End
    End If

    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 389 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = IIf(EsNulo(rs_batch_proceso!bprcparam), "", rs_batch_proceso!bprcparam)
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call InterfaceSpec(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no se encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        If Incompleto Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    
Fin:
    Flog.Close
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

Public Sub InterfaceSpec(ByVal bpronro As Long, ByVal Parametros As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Interface Spec
' Autor      : Sebastian Stremel
' Fecha      : 18/02/2013
' Ultima Mod.: 17/04/2013 - Sebastian Stremel - ver comentario v1.1
' Descripcion:
' ---------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'RecordSets
'-------------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_access As New ADODB.Recordset
Dim strAccess As String


Dim ternro As String
Dim empleg As Long

Dim arrParametros

'variables para insert final
Dim sistema As String
Dim nroTarjeta As String
Dim estado As String
Dim apellido As String
Dim nombre As String
Dim nrodoc As String
Dim tarjdesde As String
Dim tarjhasta As String
Dim cod_pers As Long
Dim cod_cent As String
Dim nif_emp As String
Dim empresa As String
Dim empestrnro As Integer
Dim empternro As Integer
Dim Nivel_1 As String
Dim Nivel_2 As String
Dim Nivel_3 As String
Dim direccion As String
Dim poblacion As String
Dim provincia As String
Dim tel1 As String
Dim tel2 As String
Dim cpostal As String
Dim fecnac As String
Dim telTrabajo As String
Dim email As String
Dim sexo As String
Dim tersex As Integer
Dim puesto As String
Dim fechaAlta As String
Dim locnro As Integer
Dim provnro As Integer

Dim GrupoA As String
Dim GrupoB As String
Dim GrupoC As String
Dim GrupoD As String
Dim GrupoE As String
Dim GrupoF As String
Dim GrupoG As String
Dim GrupoH As String

'variables confrep
Dim table_name As String
Dim parcial As Integer
Dim WHEN_TS As Date
Dim nroConexion As Integer

Dim tipoBase As String
Dim formatoFecha As String

Dim teGrupoA As Integer
Dim teGrupoB As Integer
Dim teGrupoC As Integer
Dim teGrupoD As Integer
Dim teGrupoE As Integer
Dim teGrupoF As Integer
Dim teGrupoG As Integer
Dim teGrupoH As Integer

Dim teControlHorario As Integer
Dim estrnroControlHorario As Integer

teGrupoA = 0
teGrupoB = 0
teGrupoC = 0
teGrupoD = 0
teGrupoE = 0
teGrupoF = 0
teGrupoG = 0
teGrupoH = 0
teControlHorario = 0
estrnroControlHorario = 0

tipoBase = "ORA"
formatoFecha = "dd-mm-aaaa"

StrSql = "SELECT * FROM confrep WHERE repnro = 395 ORDER BY confnrocol ASC"
OpenRecordset StrSql, rs_Consult
If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte 395."
    HuboError = True
    Exit Sub
Else

    Do While Not rs_Consult.EOF
        
        Select Case UCase(rs_Consult!conftipo)
             Case "BDO":
                nroConexion = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
        End Select
        
        Select Case rs_Consult!confnrocol
            Case "2"
                table_name = IIf(EsNulo(rs_Consult!confval2), "", rs_Consult!confval2)
            Case "3"
                parcial = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "4"
                tipoBase = IIf(EsNulo(rs_Consult!conftipo), "ORA", rs_Consult!conftipo)
                formatoFecha = IIf(EsNulo(rs_Consult!confval2), "dd-mm-aaaa", rs_Consult!confval2)
            Case "5"
                teGrupoA = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "6"
                teGrupoB = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "7"
                teGrupoC = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "8"
                teGrupoD = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "9"
                teGrupoE = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "10"
                teGrupoF = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "11"
                teGrupoG = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "12"
                teGrupoH = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "13"
                teControlHorario = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
                estrnroControlHorario = IIf(EsNulo(rs_Consult!confval2), "0", rs_Consult!confval2)
        End Select
        rs_Consult.MoveNext
    Loop
End If

'-----------------------------------------------------------------------------------------------
'Busco los datos de la conexion de la empresa
'-----------------------------------------------------------------------------------------------
StrSql = " SELECT cnstring FROM conexion "
StrSql = StrSql & " WHERE cnnro = " & nroConexion
OpenRecordset StrSql, rs_aux

If Not rs_aux.EOF Then
    strconexion = rs_aux!cnstring
Else
    Flog.writeline Espacios(Tabulador * 0) & "Error en la configuracion de la conexion Oracle."
End If
If rs_aux.State = adStateOpen Then rs_aux.Close
'-----------------------------------------------------------------------------------------------

'busco el tipo de base

Dim u As Integer
Dim tipoOperacion As String
Dim query As String
Dim esDeSpec As Boolean

sistema = "PROPIO" 'fijo segun informo el cliente

arrParametros = Split(Parametros, "@")
    If UBound(arrParametros) > 2 Then
        ternro = arrParametros(0)
        nroTarjeta = arrParametros(1)
        
        If (Len(nroTarjeta) < 16) Then
            u = CInt(16) - Len(nroTarjeta)
            If u > 0 Then
                nroTarjeta = String(u, "0") & nroTarjeta
            End If
        End If
        estado = arrParametros(2)
        tarjdesde = arrParametros(3)
        tarjhasta = arrParametros(4)
        tarjdesde = CDate(tarjdesde)
        tarjdesde = Format(tarjdesde, "YYYYMMDD")
        tarjhasta = CDate(tarjhasta)
        tarjhasta = Format(tarjhasta, "YYYYMMDD")
    Else
       ternro = arrParametros(0)
       estado = arrParametros(1)
    End If
    
    If teControlHorario = 0 Then 'SI NO SE CONFIGURO LO TOMO COMO QUE ES DE SPEC
        esDeSpec = True
    Else
        'verifico si tiene la estructura control horario sino no lo inserto
        Flog.writeline "Buscando datos de la estructura control horario del tercero:" & ternro
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & teControlHorario
        StrSql = StrSql & " AND his_estructura.estrnro=" & estrnroControlHorario
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           esDeSpec = True
           Flog.writeline " El tercero tiene la estructura control horario por lo tanto pertence a spec"
        Else
           esDeSpec = False
           Flog.writeline " El tercero no tiene la estructura control horario por lo tanto no se procesa"
        End If
        rs_Consult.Close
    End If
    
    If esDeSpec Then
        'busco el legajo del empleado en rhpro
        StrSql = " SELECT terape2, terape, ternom2, ternom, empleg FROM empleado "
        StrSql = StrSql & " WHERE ternro=" & ternro
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            Flog.writeline " Se encontro el empleado, legajo:" & rs_Consult!empleg
            
            If Not EsNulo(rs_Consult!terape2) Then
                apellido = rs_Consult!terape2 & " " & rs_Consult!terape
            Else
                apellido = rs_Consult!terape
            End If
            
            If Not EsNulo(rs_Consult!ternom2) Then
                nombre = rs_Consult!ternom2 & " " & rs_Consult!ternom
            Else
                nombre = rs_Consult!ternom
            End If
            cod_pers = rs_Consult!empleg
        Else
            Flog.writeline " No se encontro el empleado para el nro de tercero:" & ternro
            Exit Sub
        End If
        rs_Consult.Close
        
        
        'BUSCO SI TIENE TARJETA EN CASO DE QUE NO VENGA POR EL PARAMETRO
        If EsNulo(nroTarjeta) Then
            StrSql = "SELECT hstjnrotar FROM gti_histarjeta "
            StrSql = StrSql & " WHERE ternro =" & ternro
            StrSql = StrSql & " AND (hstjfecdes<=" & ConvFecha(Date) & " AND (hstjfechas is null or hstjfechas>=" & ConvFecha(Date) & "))"
            OpenRecordset StrSql, rs_Consult
            If Not rs_Consult.EOF Then
                nroTarjeta = rs_Consult!hstjnrotar
               If (Len(nroTarjeta) < 16) Then
                    u = CInt(16) - Len(nroTarjeta)
                    If u > 0 Then
                        nroTarjeta = String(u, "0") & nroTarjeta
                    End If
                End If
            Else
                nroTarjeta = ""
            End If
            rs_Consult.Close
        End If
        
        
        'HASTA ACA
        
        'BUSCO EL NRO DE DOCUMENTO DEL EMPLEADO
        StrSql = " SELECT nrodoc  FROM ter_doc"
        StrSql = StrSql & " WHERE ternro=" & ternro
        StrSql = StrSql & " AND tidnro = 1"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            nrodoc = rs_Consult!nrodoc
        Else
            Flog.writeline "No se encontro el nro de documento del tipo 1 del tercero:" & ternro
            Exit Sub
        End If
        
        rs_Consult.Close
        
        cod_cent = "AKZ" 'fijo segun informo el cliente
        
        'BUSCO EL NIF DE LA EMPRESA
        Flog.writeline "Buscando datos empresa del empleado"
            
        StrSql = " SELECT empresa.ternro, estructura.estrnro, estrdabr, tedabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN empresa on empresa.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 10"
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           empresa = rs_Consult!estrdabr
           empestrnro = rs_Consult!estrnro
           empternro = rs_Consult!ternro
           Flog.writeline " Empresa de el empleado en la fecha: " & Date & " " & empresa
        Else
           Flog.writeline " No tiene empresa el empleado en la fecha: " & Date
           empestrnro = 0
        End If
        rs_Consult.Close
        
        'BUSCO EL NIF DE LA EMPRESA
        StrSql = " SELECT nrodoc FROM ter_doc "
        StrSql = StrSql & " WHERE ternro =" & empternro & " And tidnro = 28"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            nif_emp = rs_Consult!nrodoc
            Flog.writeline " Se encontro el NIF de la empresa "
        Else
            Flog.writeline " No se encontro el NIF de la empresa "
        End If
        rs_Consult.Close
        
        
        'BUSCO EL NIVEL 1 DEL EMPLEADO
        Flog.writeline "Buscando datos del nivel 1 del empleado"
            
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 6"
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           Nivel_1 = rs_Consult!estrdabr
           'empestrnro = rs_Consult!estrnro
           'empternro = rs_Consult!ternro
           Flog.writeline " nivel 1 de el empleado en la fecha: " & Date & " " & Nivel_1
        Else
           Flog.writeline " no tiene nivel 1 el empleado en la fecha: " & Date
           Nivel_1 = ""
        End If
        rs_Consult.Close
        
        'BUSCO EL NIVEL 2 DEL EMPLEADO
        Flog.writeline "Buscando datos del nivel 2 del empleado"
            
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 5"
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           Nivel_2 = rs_Consult!estrdabr
           'empestrnro = rs_Consult!estrnro
           'empternro = rs_Consult!ternro
           Flog.writeline " nivel 2 de el empleado en la fecha: " & Date & " " & Nivel_2
        Else
           Flog.writeline " no tiene nivel 2 el empleado en la fecha: " & Date
           Nivel_2 = ""
        End If
        rs_Consult.Close
        
        'BUSCO EL NIVEL 3 DEL EMPLEADO
        Flog.writeline "Buscando datos del nivel 3 del empleado"
            
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 8"
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           Nivel_3 = rs_Consult!estrdabr
           'empestrnro = rs_Consult!estrnro
           'empternro = rs_Consult!ternro
           Flog.writeline " nivel 3 de el empleado en la fecha: " & Date & " " & Nivel_3
        Else
           Flog.writeline " no tiene nivel 3 el empleado en la fecha: " & Date
           Nivel_3 = ""
        End If
        rs_Consult.Close
        
        'Nivel_1 = "sistemas"
        'Nivel_2 = "mesa de ayuda"
        
        
        'busco la direccion del empleado
        StrSql = " SELECT * FROM cabdom "
        StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro "
        StrSql = StrSql & " LEFT JOIN localidad on localidad.locnro= detdom.locnro "
        StrSql = StrSql & " LEFT JOIN provincia on provincia.provnro= detdom.provnro "
        StrSql = StrSql & " LEFT JOIN telefono ON telefono.domnro = cabdom.domnro and telefono.tipotel=1 "
        StrSql = StrSql & " WHERE cabdom.ternro =" & ternro & " And domdefault = -1"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            If Not EsNulo(rs_Consult!calle) Then
                direccion = rs_Consult!calle & " "
            End If
            
            If Not EsNulo(rs_Consult!nro) Then
                direccion = direccion & rs_Consult!nro & " "
            End If
            
            If Not EsNulo(rs_Consult!sector) Then
                direccion = direccion & rs_Consult!sector & " "
            End If
            
            If Not EsNulo(rs_Consult!torre) Then
                direccion = direccion & rs_Consult!torre & " "
            End If
            
            If Not EsNulo(rs_Consult!piso) Then
                direccion = direccion & rs_Consult!piso & " "
            End If
            
            If Not EsNulo(rs_Consult!locdesc) Then
                poblacion = rs_Consult!locdesc
            Else
                poblacion = ""
            End If
            
            If Not EsNulo(rs_Consult!provdesc) Then
                provincia = rs_Consult!provdesc
            Else
                provincia = ""
            End If
            
            If Not EsNulo(rs_Consult!telnro) Then
                tel1 = rs_Consult!telnro
            Else
                tel1 = ""
            End If
            
            If Not EsNulo(rs_Consult!codigopostal) Then
                cpostal = rs_Consult!codigopostal
            Else
                cpostal = ""
            End If
            
            Flog.writeline " Se encontro el domicilio del empleado "
        Else
            Flog.writeline " No se encontro la direccion del empleado "
        End If
        rs_Consult.Close
        
        'BUSCO LA FECHA DE NACIMIENTO DEL TERCERO
        StrSql = "SELECT terfecnac FROM tercero "
        StrSql = StrSql & " WHERE ternro=" & ternro
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            If Not EsNulo(rs_Consult!terfecnac) Then
                fecnac = Format(rs_Consult!terfecnac, "YYYYMMDD")
            Else
                fecnac = ""
            End If
        Else
            fecnac = ""
        End If
        rs_Consult.Close
        
        'BUSCO EL NRO DE TELEFONO CELULAR DEL EMPLEADO
        StrSql = " SELECT telnro FROM cabdom "
        StrSql = StrSql & " INNER JOIN telefono ON telefono.domnro = cabdom.domnro and telefono.tipotel=2 "
        StrSql = StrSql & " WHERE cabdom.ternro =" & ternro & " And domdefault = -1"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            If Not EsNulo(rs_Consult!telnro) Then
                tel2 = rs_Consult!telnro
            Else
                tel2 = ""
            End If
        Else
            tel2 = ""
        End If
        rs_Consult.Close
        
        'BUSCO EL NRO DE TELEFONO DE LA EMPRESA
        StrSql = " SELECT telnro FROM cabdom "
        StrSql = StrSql & " INNER JOIN telefono ON telefono.domnro = cabdom.domnro and telefono.tipotel=1 "
        StrSql = StrSql & " WHERE cabdom.ternro =" & empternro & " And domdefault = -1"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            If Not EsNulo(rs_Consult!telnro) Then
                telTrabajo = rs_Consult!telnro
            Else
                telTrabajo = ""
            End If
        Else
            telTrabajo = ""
        End If
        rs_Consult.Close
        
        'BUSCO EL MAIL DEL EMPLEADO
        StrSql = "SELECT empemail FROM empleado "
        StrSql = StrSql & " WHERE ternro=" & ternro
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            If Not EsNulo(rs_Consult!empemail) Then
                email = rs_Consult!empemail
            Else
                email = ""
            End If
        Else
            email = ""
        End If
        rs_Consult.Close
        
        'BUSCO EL SEXO DEL EMPLEADO
        StrSql = " SELECT tersex FROM tercero "
        StrSql = StrSql & "WHERE ternro=" & ternro
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            If Not EsNulo(rs_Consult!tersex) Then
                tersex = rs_Consult!tersex
            Else
                tersex = 0
            End If
        Else
            tersex = 0
        End If
        rs_Consult.Close
        
        If tersex = 0 Then
            sexo = "M"
        Else
            If tersex = -1 Then
             sexo = "H"
            End If
        End If
        
        'Busco el valor del grupo A
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & teGrupoA
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           GrupoA = rs_Consult!estrdabr
           Flog.writeline " Grupo A de el empleado en la fecha: " & Date & " " & GrupoA
        Else
           Flog.writeline " no tiene Grupo A el empleado en la fecha: " & Date
           GrupoA = ""
        End If
        rs_Consult.Close
        
        'Busco el valor del grupo B
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & teGrupoB
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           GrupoB = rs_Consult!estrdabr
           Flog.writeline " Grupo B de el empleado en la fecha: " & Date & " " & GrupoB
        Else
           Flog.writeline " no tiene Grupo B el empleado en la fecha: " & Date
           GrupoB = ""
        End If
        rs_Consult.Close
        
        'Busco el valor del grupo C
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & teGrupoC
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           GrupoC = rs_Consult!estrdabr
           Flog.writeline " Grupo C de el empleado en la fecha: " & Date & " " & GrupoC
        Else
           Flog.writeline " no tiene Grupo C el empleado en la fecha: " & Date
           GrupoC = ""
        End If
        rs_Consult.Close
        
        'Busco el valor del grupo D
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & teGrupoD
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           GrupoD = rs_Consult!estrdabr
           Flog.writeline " Grupo D de el empleado en la fecha: " & Date & " " & GrupoD
        Else
           Flog.writeline " no tiene Grupo D el empleado en la fecha: " & Date
           GrupoD = ""
        End If
        rs_Consult.Close
        
        'Busco el valor del grupo E
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & teGrupoE
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           GrupoE = rs_Consult!estrdabr
           Flog.writeline " Grupo E de el empleado en la fecha: " & Date & " " & GrupoE
        Else
           Flog.writeline " no tiene Grupo E el empleado en la fecha: " & Date
           GrupoE = ""
        End If
        rs_Consult.Close
        
        'Busco el valor del grupo F
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & teGrupoF
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           GrupoF = rs_Consult!estrdabr
           Flog.writeline " Grupo F de el empleado en la fecha: " & Date & " " & GrupoF
        Else
           Flog.writeline " no tiene Grupo F el empleado en la fecha: " & Date
           GrupoF = ""
        End If
        rs_Consult.Close
        
        'Busco el valor del grupo G
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & teGrupoG
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           GrupoG = rs_Consult!estrdabr
           Flog.writeline " Grupo G de el empleado en la fecha: " & Date & " " & GrupoG
        Else
           Flog.writeline " no tiene Grupo G el empleado en la fecha: " & Date
           GrupoG = ""
        End If
        rs_Consult.Close
        
        'Busco el valor del grupo H
        StrSql = " SELECT estructura.estrnro, estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & teGrupoH
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           GrupoH = rs_Consult!estrdabr
           Flog.writeline " Grupo H de el empleado en la fecha: " & Date & " " & GrupoH
        Else
           Flog.writeline " no tiene Grupo H el empleado en la fecha: " & Date
           GrupoH = ""
        End If
        rs_Consult.Close
        
        
        'BUSCO EL NOMBRE DE EMPLEO (LO SACO DEL NOMBRE DEL PUESTO)
        Flog.writeline "Buscando nombre del empleo"
            
        StrSql = " SELECT empresa.ternro, estructura.estrnro, estrdabr, tedabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN empresa on empresa.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = 4"
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(Date) & " AND (htethasta is null or htethasta>=" & ConvFecha(Date) & "))"
        StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
           puesto = rs_Consult!estrdabr
           Flog.writeline " puesto de el empleado en la fecha: " & Date & " " & puesto
        Else
           Flog.writeline " No tiene puesto el empleado en la fecha: " & Date
           puesto = ""
        End If
        rs_Consult.Close
        
        'BUSCO LA FECHA DE ALTA DEL EMPLEADO
        StrSql = " SELECT empfaltagr FROM empleado"
        StrSql = StrSql & " WHERE ternro=" & ternro
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            If Not EsNulo(rs_Consult!empfaltagr) Then
                fechaAlta = Format(rs_Consult!empfaltagr, "YYYYMMDD")
            Else
                fechaAlta = ""
            End If
        Else
            fechaAlta = ""
        End If
        rs_Consult.Close
        
        'Abro la conexion con las tablas de Oracle
        'strconexion = " Provider=OraOLEDB.Oracle.1;Persist Security Info=true; Data Source=rhoracle/SOSR3; user Id=SOSR3; password=SOSR3;"
        OpenConnExt strconexion, ExtConnOra
        'OpenConnection strconexion, ExtConn
        If Err.Number <> 0 Or Error_Encrypt Then
            Flog.writeline "Problemas en la conexion Oracle"
            Exit Sub
        End If
        
        Flog.writeline Espacios(Tabulador * 0) & "Conexion con Oracle establecida."
        
        'INSERTO LOS DATOS EN LA TABLA DE ORACLE
        Flog.writeline "Empieza la insercion en la tabla IMP_PERSONAL "
        StrSql = " INSERT INTO IMP_PERSONAL "
        StrSql = StrSql & "("
        StrSql = StrSql & " sistema ,"
        StrSql = StrSql & " estado ,"
        StrSql = StrSql & " cod_pers ,"
        StrSql = StrSql & " tarjeta ,"
        StrSql = StrSql & " ini_val ,"
        StrSql = StrSql & " fin_val ,"
        StrSql = StrSql & " apellidos ,"
        StrSql = StrSql & " nombre ,"
        StrSql = StrSql & " dni ,"
        StrSql = StrSql & " cod_cent ,"
        StrSql = StrSql & " nif_emp ,"
        StrSql = StrSql & " nivel_1 ,"
        StrSql = StrSql & " nivel_2 ,"
        StrSql = StrSql & " nivel_3 ,"
        StrSql = StrSql & " nivel_4 ,"
        StrSql = StrSql & " nivel_5 ,"
        StrSql = StrSql & " nivel_6 ,"
        StrSql = StrSql & " nivel_7 ,"
        StrSql = StrSql & " nivel_8 ,"
        StrSql = StrSql & " nivel_9 ,"
        StrSql = StrSql & " nivel_10 ,"
        StrSql = StrSql & " direccion ,"
        StrSql = StrSql & " poblacion ,"
        StrSql = StrSql & " provincia ,"
        StrSql = StrSql & " tlf_pers ,"
        StrSql = StrSql & " cod_post ,"
        StrSql = StrSql & " f_nacimiento ,"
        StrSql = StrSql & " tlf_empr ,"
        StrSql = StrSql & " extension ,"
        StrSql = StrSql & " movil ,"
        StrSql = StrSql & " edificio ,"
        StrSql = StrSql & " planta ,"
        StrSql = StrSql & " despacho ,"
        StrSql = StrSql & " vers_tarj ,"
        StrSql = StrSql & " num_ss ,"
        StrSql = StrSql & " matricula ,"
        StrSql = StrSql & " mail ,"
        StrSql = StrSql & " sexo ,"
        StrSql = StrSql & " grupo_a ,"
        StrSql = StrSql & " grupo_b ,"
        StrSql = StrSql & " grupo_c ,"
        StrSql = StrSql & " grupo_d ,"
        StrSql = StrSql & " grupo_e ,"
        StrSql = StrSql & " grupo_f ,"
        StrSql = StrSql & " grupo_g ,"
        StrSql = StrSql & " grupo_h ,"
        StrSql = StrSql & " dato_1 ,"
        StrSql = StrSql & " dato_2 ,"
        StrSql = StrSql & " dato_3 ,"
        StrSql = StrSql & " dato_4 ,"
        StrSql = StrSql & " dato_5 ,"
        StrSql = StrSql & " dato_6 ,"
        StrSql = StrSql & " dato_7 ,"
        StrSql = StrSql & " dato_8 ,"
        StrSql = StrSql & " nombre_empleo ,"
        StrSql = StrSql & " f_alta_empresa ,"
        StrSql = StrSql & " cal ,"
        StrSql = StrSql & " cal_fest)"
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & "'" & sistema & "',"
        StrSql = StrSql & "'" & estado & "',"
        StrSql = StrSql & "'" & cod_pers & "',"
        StrSql = StrSql & "'" & nroTarjeta & "',"
        StrSql = StrSql & "'" & tarjdesde & "',"
        StrSql = StrSql & "'" & tarjhasta & "',"
        StrSql = StrSql & "'" & apellido & "',"
        StrSql = StrSql & "'" & nombre & "',"
        StrSql = StrSql & "'" & nrodoc & "',"
        StrSql = StrSql & "'" & cod_cent & "',"
        StrSql = StrSql & "'" & nif_emp & "',"
        StrSql = StrSql & "'" & Nivel_1 & "',"
        StrSql = StrSql & "'" & Nivel_2 & "',"
        StrSql = StrSql & "'" & Nivel_3 & "',"
        StrSql = StrSql & "''," 'nivel 4
        StrSql = StrSql & "''," 'nivel 5
        StrSql = StrSql & "''," 'nivel 6
        StrSql = StrSql & "''," 'nivel 7
        StrSql = StrSql & "''," 'nivel 8
        StrSql = StrSql & "''," 'nivel 9
        StrSql = StrSql & "''," 'nivel 10
        StrSql = StrSql & "'" & direccion & "',"
        StrSql = StrSql & "'" & poblacion & "',"
        StrSql = StrSql & "'" & provincia & "',"
        StrSql = StrSql & "'" & tel1 & "',"
        StrSql = StrSql & "'" & cpostal & "',"
        StrSql = StrSql & "'" & fecnac & "',"
        StrSql = StrSql & "'" & telTrabajo & "',"
        StrSql = StrSql & "''," 'extension
        StrSql = StrSql & "'" & tel2 & "'," 'movil
        StrSql = StrSql & "''," 'edificio
        StrSql = StrSql & "''," 'planta
        StrSql = StrSql & "''," 'despacho
        StrSql = StrSql & "''," 'version tarjeta
        StrSql = StrSql & "''," 'numero de seguridad social
        StrSql = StrSql & "'" & cod_pers & "'," 'legajo del empleado
        StrSql = StrSql & "'" & email & "',"
        StrSql = StrSql & "'" & sexo & "',"
        StrSql = StrSql & "'" & GrupoA & "',"
        StrSql = StrSql & "'" & GrupoB & "',"
        StrSql = StrSql & "'" & GrupoC & "',"
        StrSql = StrSql & "'" & GrupoD & "',"
        StrSql = StrSql & "'" & GrupoE & "',"
        StrSql = StrSql & "'" & GrupoF & "',"
        StrSql = StrSql & "'" & GrupoG & "',"
        StrSql = StrSql & "'" & GrupoH & "',"
        StrSql = StrSql & "''," 'Dato 1
        StrSql = StrSql & "'',"
        StrSql = StrSql & "'',"
        StrSql = StrSql & "'',"
        StrSql = StrSql & "'',"
        StrSql = StrSql & "'',"
        StrSql = StrSql & "'',"
        StrSql = StrSql & "''," 'Dato 8
        StrSql = StrSql & "'" & puesto & "',"
        StrSql = StrSql & "'" & fechaAlta & "',"
        StrSql = StrSql & "''," ' cal
        StrSql = StrSql & "''" 'CAL_FEST
        StrSql = StrSql & ")"
        Flog.writeline "query de insercion:" & StrSql
        ExtConnOra.Execute StrSql, , adExecuteNoRecords
        Flog.writeline " Se inserto el registro correctamente"
        
        Flog.writeline "FORMATO DE FECHA:" & formatoFecha
        
        Flog.writeline "Inserto el flag para que SPEC comienze a importar "
        StrSql = "INSERT INTO DOWNCONF"
        StrSql = StrSql & "(WHEN_TS, TABLE_NAME, PARCIAL)"
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "("
        'StrSql = StrSql & " TO_DATE (" & ConvFecha(Now()) & ",'dd-mm-yyyy') , "
        If UCase(tipoBase) = "ORA" Then
            StrSql = StrSql & " TO_DATE (" & ConvFecha(Now()) & ",'" & formatoFecha & "' ) , "
        Else
            Select Case Trim(formatoFecha)
                Case "dd-mm-yyyy"
                    formatoFecha = 105
                    
                Case "dd-mm-yy"
                    formatoFecha = 5
                
                Case "dd/mm/yyyy"
                    formatoFecha = 103
                    
                Case "dd/mm/yy"
                    formatoFecha = 3
                    
                Case Else
                    formatoFecha = 111 ' aaaa/mm/dd
                    
            End Select
            StrSql = StrSql & " Convert(date, getdate()," & formatoFecha & "),"
        End If
        StrSql = StrSql & "'" & table_name & "', "
        StrSql = StrSql & parcial
        StrSql = StrSql & ")"
        Flog.writeline "Flag:" & StrSql
        ExtConnOra.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "CODIGO FORMATO DE FECHA:" & formatoFecha
        Flog.writeline " Se inserto el registro de flag correctamente"
    End If
   

End Sub

Public Sub OpenConnExt(strConnectionString As String, ByRef objConn As ADODB.Connection)
' ---------------------------------------------------------------------------------------------
' Descripcion: Abre conexion externa
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    If objConn.State <> adStateClosed Then objConn.Close
    objConn.CursorLocation = adUseClient
    
    'Indica que desde una transacción se pueden ver cambios que no se han producido en otras transacciones.
    objConn.IsolationLevel = adXactReadUncommitted
    
    objConn.CommandTimeout = 3600 'segundos
    objConn.ConnectionTimeout = 60 'segundos
    objConn.Open strConnectionString
End Sub
