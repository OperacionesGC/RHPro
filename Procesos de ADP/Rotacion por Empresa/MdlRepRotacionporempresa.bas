Attribute VB_Name = "MdlRepRotacionporempresa"
Option Explicit
'Global Const Version = "1.00"
'Global Const FechaModificacion = "20/11/2006"   'JV
'Global Const UltimaModificacion = " "

'Global Const Version = "1.01"
'Global Const FechaModificacion = "17/04/2007"   'FGZ
'Global Const UltimaModificacion = " "

'Global Const Version = "1.02"
'Global Const FechaModificacion = "17/04/2007"   'FGZ
'Global Const UltimaModificacion = " "   'le cambié v_empleado por empleado

'Global Const Version = "1.03"
'Global Const FechaModificacion = "19/10/2007"   'FAF
'Global Const UltimaModificacion = " "   'Se agrego el order by en todoas las consultas
                                        'Para el caso de las bajas, no consideraba que el empleado perteneciera a Contrato.
'Global Const Version = "1.04"
'Global Const FechaModificacion = "29/10/2007"   'FAF
'Global Const UltimaModificacion = " "   'Para el caso de las bajas, realizaba mal el calculo.

'Global Const Version = "1.05"
'Global Const FechaModificacion = "08/01/2009"   'Lisandro Moro
'Global Const UltimaModificacion = " "   'El reporte estaba incluyendo todas las bajas  hasta que hay en el sistema. Solo debe incluir las bajas que hubo en cada mes.

'Global Const Version = "1.06"
'Global Const FechaModificacion = "12/02/2009"   'Lisandro Moro
'Global Const UltimaModificacion = " "   'Se agregaron las fases para validar el estado del empleado.
'                                        'Se corrigio el estado del proceso ya que lo cerraba el appserver.
'                                        'Se mantiene la lista de empleados a traves de las consultas para no repetirlos.

Global Const Version = "1.07"
Global Const FechaModificacion = "31/07/2009"   'Martin Ferraro
Global Const UltimaModificacion = " "   'Encriptacion de string connection


'===========================================================================================================================
Global IdUser As String
Global Fecha As Date
Global Hora As String

Dim ternroAlta As String
Dim ternroBaja As String

Public Sub Main()

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
Dim ArrParametros

'    strCmdLine = Command()
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
    
On Error GoTo CE
   ' carga las configuraciones basicas, formato de fecha, string de conexion,tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    
    Nombre_Arch = PathFLog & "Reporte_Rotacion_x_empresa" & "-" & NroProcesoBatch & ".log"
       
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha        = " & FechaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo CE
    
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 1, bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 150 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        
        Flog.writeline "Llamo a LevantarParamteros"
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    Set rs_batch_proceso = Nothing
    
    TiempoFinalProceso = GetTickCount
    
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Set objConn = Nothing
    

    Flog.writeline "----------------------------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.Close
    
Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "==================================================================================="
    Flog.writeline Err.Description
    Flog.writeline "Ultimo SQL Ejecutado:"
    Flog.writeline StrSql
    Flog.writeline "==================================================================================="
End Sub

'Call rotacion_emp(bpronro, Finicio, Estruct, Subestr, Agencia, Agencianro, Orden, Ordenado)
'Public Sub rotacion_emp(ByVal bpronro As Long, ByVal Finicio As Long, ByVal Estruct As Long, ByVal Subestr As String, ByVal Agencia As Long, ByVal Agencianro As Long, ByVal Orden As String, ByVal Ordenado As String, ByVal Fecestr As Date)
Public Sub rotacion_emp(ByVal bpronro As Long, ByVal Finicio As Date, ByVal Estruct As Long, ByVal Subestr As String, ByVal Agencia As Long, ByVal Agencianro As Long, ByVal Orden As String, ByVal Ordenado As String, ByVal Fecestr As Date)

Dim Progreso
Dim IncPorc
Dim fecha_base
Dim Sbestr0
Dim Sbestr
Dim rs_sql As New ADODB.Recordset
Dim rs_Rep_rotacion As New ADODB.Recordset
Dim m_dota()
Dim m_baja()
Dim X, v_records
Dim pos1, pos2
Dim v_mes
Dim mes
Dim Anio
Dim l_fil_agen
Dim l_sql



On Error GoTo CE

fecha_base = CDate(Finicio)
mes = Month(fecha_base)
Anio = Year(fecha_base)
StrSql = "select * from estructura where ESTRUCTURA.TENRO=" & Estruct & " order by ESTRUCTURA.estrnro"
OpenRecordset StrSql, rs_sql
        pos1 = 2
        pos2 = InStr(pos1, Subestr, ",") - 2
        Sbestr0 = Mid(Subestr, pos1, pos2)
        Sbestr = Mid(Subestr, pos2 + 3, Len(Subestr) - (pos2 + 3))

v_records = rs_sql.RecordCount

Flog.writeline "filtro y ordenamiento"
'-----------------------------
'filtro y ordenamiento
'-----------------------------
l_fil_agen = "1=1 " ' cuando queremos todos los empleados
If Agencia = "-1" Then
    
    l_fil_agen = "  empleado.ternro not in (SELECT ternro from his_estructura agencia " & _
            " WHERE agencia.tenro=28 " & _
            " AND (agencia.htetdesde<=" & ConvFecha(Fecestr) & _
            " AND (agencia.htethasta is null or agencia.htethasta>=" & ConvFecha(Fecestr) & ")) )"
Else
    If Agencia = "-2" Then
        l_fil_agen = "  empleado.ternro in (SELECT ternro from his_estructura agencia " & _
            " WHERE agencia.tenro=28 " & _
                " AND (agencia.htetdesde<=" & ConvFecha(Fecestr) & _
                " AND (agencia.htethasta is null or agencia.htethasta>=" & ConvFecha(Fecestr) & ")) )"
    Else
        If Agencia <> "0" Then 'este caso se da cuando selecionamos una agencia determinada
            l_fil_agen = "  empleado.ternro in (SELECT ternro from his_estructura agencia " & _
                    " WHERE agencia.tenro=28 and agencia.estrnro=" & Agencianro & _
                    " AND (agencia.htetdesde<=" & ConvFecha(Fecestr) & _
                    " AND (agencia.htethasta is null or agencia.htethasta>=" & ConvFecha(Fecestr) & ")) )"
        End If
    End If
End If

l_sql = "SELECT DISTINCT empleado.ternro"
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " WHERE " & l_fil_agen

'-----------------------------
'fin filtro y ordenamiento
'-----------------------------

ReDim Preserve m_dota(v_records, 14)
ReDim Preserve m_baja(v_records, 14)

For X = 1 To v_records
    m_dota(X, 0) = rs_sql!estrnro
    m_baja(X, 0) = rs_sql!estrnro
    rs_sql.MoveNext
Next
 For X = 1 To mes  'For X = 1 To 12
    'cargar_mes x, anio, m_dota
    v_mes = Trim(Str(X))
    If Len(v_mes) < 2 Then
        v_mes = "0" & v_mes
    End If
    
    fecha_base = "01/" & v_mes & "/" & Anio
    ternroAlta = "0"
    ternroBaja = "0"

    Flog.writeline "Fecha de calculo --> " & fecha_base
    Flog.writeline " (A) DOTACION del MES = Dotacion INICIAL(1) + ALTAS CONTRATOS(2) + ALTAS SECTOR(4) - BAJAS CONTRATOS(3) - BAJAS SECTOR(B)"
    '-- (A) DOTACION del MES = Dotacion INICIAL(1) + ALTAS CONTRATOS(2) + ALTAS SECTOR(4) - BAJAS CONTRATOS(3) - BAJAS SECTOR(B)
        
'    StrSql = "select his_sectores.estrnro, count(*) cant from  his_estructura his_sectores"
'    StrSql = StrSql & " Where his_sectores.tenro = " & Estruct
'    StrSql = StrSql & " and his_sectores.htetdesde < '" & fecha_base & "'"
'    StrSql = StrSql & " and (his_sectores.htethasta is null or his_sectores.htethasta >= '" & fecha_base & "' ) "
'    StrSql = StrSql & " and exists (select 'x' from his_estructura his_contratos "
'    StrSql = StrSql & " Where his_sectores.ternro = his_contratos.ternro "
'    StrSql = StrSql & " and his_contratos.tenro=" & Sbestr0
'    StrSql = StrSql & " and his_contratos.estrnro IN (" & Sbestr & ")" ' -- FILTRO CONTRATOS
'    StrSql = StrSql & " and his_contratos.htetdesde < '" & fecha_base & "'"
'    StrSql = StrSql & " and (his_contratos.htethasta is null or his_contratos.htethasta >= '" & fecha_base & "' ))"
'    StrSql = StrSql & " and his_sectores.ternro in (" & l_sql & ")"
'    StrSql = StrSql & " group by his_sectores.estrnro "

''   FGZ - 17/04/2007 - faltaba todos los convfecha()
'    StrSql = "select his_sectores.estrnro, count(*) cant from  his_estructura his_sectores"
'    StrSql = StrSql & " Where his_sectores.tenro = " & Estruct
'    StrSql = StrSql & " and his_sectores.htetdesde < " & ConvFecha(fecha_base)
'    StrSql = StrSql & " and (his_sectores.htethasta is null or his_sectores.htethasta < " & ConvFecha(fecha_base) & " ) "
'    StrSql = StrSql & " and exists (select 'x' from his_estructura his_contratos "
'    StrSql = StrSql & " Where his_sectores.ternro = his_contratos.ternro "
'    StrSql = StrSql & " and his_contratos.tenro=" & Sbestr0
'    StrSql = StrSql & " and his_contratos.estrnro IN (" & Sbestr & ")" ' -- FILTRO CONTRATOS
'    StrSql = StrSql & " and his_contratos.htetdesde < " & ConvFecha(fecha_base)
'    StrSql = StrSql & " and (his_contratos.htethasta is null or his_contratos.htethasta < " & ConvFecha(fecha_base) & " ))"
'    StrSql = StrSql & " and his_sectores.ternro in (" & l_sql & ")"
'    StrSql = StrSql & " group by his_sectores.estrnro "
'    '-- FAF - 19/10/2007 - Se agrego el orden
'    StrSql = StrSql & " order by his_sectores.estrnro"
''   FGZ - 17/04/2007 - faltaba todos los convfecha()

    StrSql = "select his_sectores.estrnro, his_sectores.ternro "
    StrSql = StrSql & " FROM his_estructura his_sectores"
    StrSql = StrSql & " INNER JOIN fases ON his_sectores.ternro = fases.empleado AND (fases.altfec <= " & ConvFecha(fecha_base) & " AND (fases.bajfec is null OR fases.bajfec >= " & ConvFecha(fecha_base) & " ))"
    StrSql = StrSql & " Where his_sectores.tenro = " & Estruct
    StrSql = StrSql & " and his_sectores.htetdesde <= " & ConvFecha(fecha_base)
    StrSql = StrSql & " and (his_sectores.htethasta is null or his_sectores.htethasta >= " & ConvFecha(fecha_base) & " ) "
    StrSql = StrSql & " and exists (select 'x' from his_estructura his_contratos "
    StrSql = StrSql & " Where his_sectores.ternro = his_contratos.ternro "
    StrSql = StrSql & " and his_contratos.tenro=" & Sbestr0
    StrSql = StrSql & " and his_contratos.estrnro IN (" & Sbestr & ")" ' -- FILTRO CONTRATOS
    StrSql = StrSql & " and his_contratos.htetdesde < " & ConvFecha(fecha_base)
    StrSql = StrSql & " and (his_contratos.htethasta is null or his_contratos.htethasta >= " & ConvFecha(fecha_base) & " ))"
    StrSql = StrSql & " and his_sectores.ternro in (" & l_sql & ")"
    'StrSql = StrSql & " AND his_sectores.ternro NOT IN (" & ternroAlta & ")"
    'StrSql = StrSql & " group by his_sectores.estrnro "
    StrSql = StrSql & " order by his_sectores.estrnro"

    OpenRecordset StrSql, rs_sql

    Flog.writeline " Dotacion INICIAL(1)"
    Flog.writeline ""
    Flog.writeline " " & StrSql
    Flog.writeline ""

    If Not rs_sql.EOF Then
       cargar_mes v_mes, rs_sql, v_records, m_dota, 1, True
    End If
    rs_sql.Close
    
    
    '-- (2) ALTA CONTRATOS a SUMAR Dotacion
    StrSql = "select his_sectores.estrnro, his_contratos.ternro " 'count(*) cant "
    StrSql = StrSql & " from his_estructura his_contratos, his_estructura his_sectores "
    StrSql = StrSql & " Where his_contratos.tenro = " & Sbestr0
    StrSql = StrSql & " and his_contratos.ternro = his_sectores.ternro "
    StrSql = StrSql & " and his_contratos.estrnro IN (" & Sbestr & ")"  '-- FILTRO CONTRATOS
    StrSql = StrSql & " and month(his_contratos.htetdesde) = month('" & fecha_base & "')"
    StrSql = StrSql & " and year(his_contratos.htetdesde) = year( '" & fecha_base & "')"
    StrSql = StrSql & " AND his_sectores.tenro= " & Estruct
    StrSql = StrSql & " and his_sectores.htetdesde <=  his_contratos.htetdesde "
    StrSql = StrSql & " and (his_sectores.htethasta is null or his_sectores.htethasta > his_contratos.htetdesde) "
    StrSql = StrSql & " and his_contratos.ternro in (" & l_sql & ")"
    StrSql = StrSql & " AND his_contratos.ternro NOT IN (" & ternroAlta & ")"
    'StrSql = StrSql & " group by his_sectores.estrnro "
    '-- FAF - 19/10/2007 - Se agrego el orden
    StrSql = StrSql & " order by his_sectores.estrnro"

    OpenRecordset StrSql, rs_sql

    Flog.writeline " ALTAS CONTRATOS(2)"
    Flog.writeline ""
    Flog.writeline " " & StrSql
    Flog.writeline ""
    If Not rs_sql.EOF Then
       cargar_mes v_mes, rs_sql, v_records, m_dota, 1, True
    End If
    rs_sql.Close
    
    '-- (3) BAJAS de CONTRATOS a RESTAR Dotacion
    StrSql = "select his_sectores.estrnro, his_contratos.ternro " 'count(*) cant "
    StrSql = StrSql & " from his_estructura his_contratos, his_estructura his_sectores"
    StrSql = StrSql & " Where his_contratos.tenro =" & Sbestr0
    StrSql = StrSql & " and his_contratos.ternro = his_sectores.ternro"
    StrSql = StrSql & " and his_contratos.estrnro IN (" & Sbestr & ")" '-- FILTRO CONTRATOS
    StrSql = StrSql & " and month(his_contratos.htethasta) = month('" & fecha_base & "')"
    StrSql = StrSql & " and year(his_contratos.htethasta) = year('" & fecha_base & " ')"
    StrSql = StrSql & " and his_sectores.tenro=" & Estruct
    StrSql = StrSql & " and his_sectores.htetdesde < his_contratos.htethasta"
    StrSql = StrSql & " and (his_sectores.htethasta is null or his_sectores.htethasta >=  his_contratos.htethasta ) "
    StrSql = StrSql & " and his_contratos.ternro  in (" & l_sql & ")"
    'StrSql = StrSql & " AND his_contratos.ternro NOT IN (" & ternroAlta & ")"
    'StrSql = StrSql & " group by his_sectores.estrnro"
    '-- FAF - 19/10/2007 - Se agrego el orden
    StrSql = StrSql & " order by his_sectores.estrnro"
    
    OpenRecordset StrSql, rs_sql

    Flog.writeline " BAJAS CONTRATOS(3)"
    Flog.writeline ""
    Flog.writeline " " & StrSql
    Flog.writeline ""
    If Not rs_sql.EOF Then
       cargar_mes v_mes, rs_sql, v_records, m_dota, 0, True
    End If
    rs_sql.Close
    
    '-- (4) ALTAS de SECTOR del MES a SUMAR Dotacion
    StrSql = "select his_sectores.estrnro, his_sectores.ternro " 'count(*) cant "
    StrSql = StrSql & " from  his_estructura his_sectores"
    StrSql = StrSql & " Where his_sectores.tenro = " & Estruct
    StrSql = StrSql & " and month(his_sectores.htetdesde) = month('" & fecha_base & "')"
    StrSql = StrSql & " and year(his_sectores.htetdesde) = year('" & fecha_base & "')"
    StrSql = StrSql & " and exists (select 'x' from his_estructura his_contratos"
    StrSql = StrSql & " Where his_sectores.ternro = his_contratos.ternro"
    StrSql = StrSql & " and his_contratos.tenro= " & Sbestr0
    StrSql = StrSql & " and his_contratos.estrnro IN (" & Sbestr & ")" '(2911) -- FILTRO CONTRATOS
    StrSql = StrSql & " and his_contratos.htetdesde < his_sectores.htetdesde"
    StrSql = StrSql & " and (his_contratos.htethasta is null or his_contratos.htethasta > his_sectores.htetdesde ))"
    StrSql = StrSql & " and his_sectores.ternro in (" & l_sql & ")"
    StrSql = StrSql & " AND his_sectores.ternro NOT IN (" & ternroAlta & ")"
    'StrSql = StrSql & " group by his_sectores.estrnro"
    '-- FAF - 19/10/2007 - Se agrego el orden
    StrSql = StrSql & " order by his_sectores.estrnro"
    
    OpenRecordset StrSql, rs_sql

    Flog.writeline " ALTAS SECTOR(4)"
    Flog.writeline ""
    Flog.writeline " " & StrSql
    Flog.writeline ""
    If Not rs_sql.EOF Then
       cargar_mes v_mes, rs_sql, v_records, m_dota, 1, True
    End If
    rs_sql.Close
    
    Flog.writeline "(B) BAJAS DEL MES DE SECTORES = (1) BAJAS por cambio de sectores + (2) bajas empleados del sector"
    '-- (B) BAJAS DEL MES DE SECTORES = (1) BAJAS por cambio de sectores + (2) bajas empleados del sector
    Flog.writeline " (1) BAJAS por cambio de sectores"
    '--   (1) BAJAS por cambio de sectores
    StrSql = "select his_sectores.estrnro, his_sectores.ternro " 'count(*) cant"
    StrSql = StrSql & " from  his_estructura his_sectores"
    StrSql = StrSql & " Where his_sectores.tenro = " & Estruct
    StrSql = StrSql & " and month(his_sectores.htethasta) = month('" & fecha_base & "')"
    StrSql = StrSql & " and year(his_sectores.htethasta) = year('" & fecha_base & "')"
    StrSql = StrSql & " and exists (select 'x' from his_estructura his_contratos"
    StrSql = StrSql & " Where his_sectores.ternro = his_contratos.ternro"
    StrSql = StrSql & " and his_contratos.tenro=" & Sbestr0
    StrSql = StrSql & " and his_contratos.estrnro IN (" & Sbestr & ")"  '-- FILTRO CONTRATOS"
    StrSql = StrSql & " and his_contratos.htetdesde <= his_sectores.htethasta"
    StrSql = StrSql & " and (his_contratos.htethasta is null or his_contratos.htethasta >= his_sectores.htethasta) )"
    StrSql = StrSql & " and his_sectores.ternro in (" & l_sql & ")"
    'StrSql = StrSql & " group by his_sectores.estrnro"
    '-- FAF - 19/10/2007 - Se agrego el orden
    StrSql = StrSql & " order by his_sectores.estrnro"

    Flog.writeline ""
    Flog.writeline " " & StrSql
    Flog.writeline ""
    
    OpenRecordset StrSql, rs_sql

    If Not rs_sql.EOF Then
       cargar_mes v_mes, rs_sql, v_records, m_baja, 1, False
    End If
    rs_sql.Close

    Flog.writeline " (2) bajas empleados del sector"
    '-- (2) bajas empleados del sector
    StrSql = "select his_sectores.estrnro, his_sectores.ternro " 'count(*) cant"
    StrSql = StrSql & " from  his_estructura his_sectores, fases"
    StrSql = StrSql & " Where fases.Empleado = his_sectores.ternro"
    StrSql = StrSql & " and month(fases.bajfec) = month('" & fecha_base & "')"
    StrSql = StrSql & " and year(fases.bajfec) = year('" & fecha_base & "')"
    StrSql = StrSql & " AND his_sectores.tenro= " & Estruct
    StrSql = StrSql & " and his_sectores.htetdesde <= fases.bajfec"
    StrSql = StrSql & " and (his_sectores.htethasta is null or his_sectores.htethasta >= fases.bajfec )"
    '-- FAF - 19/10/2007 - Se agrego que pertenescan a Contratos, como el resto de las sql anteriores
    '-- FAF - 29/10/2007 - La pertenencia a Contratos debe ser a la fecha de baja de la fase.
    StrSql = StrSql & " and exists (select 'x' from his_estructura his_contratos"
    StrSql = StrSql & " Where his_sectores.ternro = his_contratos.ternro"
    StrSql = StrSql & " and his_contratos.tenro=" & Sbestr0
    StrSql = StrSql & " and his_contratos.estrnro IN (" & Sbestr & ")"  '-- FILTRO CONTRATOS"
    StrSql = StrSql & " and his_contratos.htetdesde <= fases.bajfec"
    StrSql = StrSql & " and (his_contratos.htethasta is null or his_contratos.htethasta >= fases.bajfec) )"
    '-- Fin modificacion
    StrSql = StrSql & " and his_sectores.ternro in (" & l_sql & ")"
    StrSql = StrSql & " AND his_sectores.ternro NOT IN (" & ternroBaja & ")"
    'StrSql = StrSql & " group by his_sectores.estrnro"
    '-- FAF - 19/10/2007 - Se agrego el orden
    StrSql = StrSql & " order by his_sectores.estrnro"

    Flog.writeline ""
    Flog.writeline StrSql
    Flog.writeline ""

    OpenRecordset StrSql, rs_sql

    If Not rs_sql.EOF Then
       cargar_mes v_mes, rs_sql, v_records, m_baja, 1, False
    End If
    rs_sql.Close
  
    Set rs_sql = Nothing
    
    IncPorc = X * 100 / mes
    
    ''Actualizo el progreso del Proceso
    Flog.writeline "Actualizo el progreso del Proceso"
    Progreso = IncPorc 'Progreso +
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords

Next

   
    ' Comienzo la transaccion
Flog.writeline "Inicio la Transaccion"
MyBeginTrans

'seteo de las variables de progreso
Progreso = 99
'IncPorc = (100 / 8)
IncPorc = 0


'Depuracion del Temporario
Flog.writeline "Depuracion del Temporario"
StrSql = "DELETE FROM Rep_rotacion "
StrSql = StrSql & " WHERE bpronro = " & bpronro

objConn.Execute StrSql, , adExecuteNoRecords


''Actualizo el progreso del Proceso
Flog.writeline "Actualizo el progreso del Proceso"
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords


'**************************************************************************************************
'************************************ DETALLE DE ROTACIONES ***************************************
Flog.writeline "DETALLE DE LAS ROTACIONES"
'**************************************************************************************************

'Progreso = Progreso + IncPorc
'TiempoAcumulado = GetTickCount
'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
'         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
'         "' WHERE bpronro = " & NroProcesoBatch
'objConn.Execute StrSql, , adExecuteNoRecords

''Actualizo el progreso del Proceso
'Progreso = Progreso + IncPorc
'TiempoAcumulado = GetTickCount
'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
'         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
'         "' WHERE bpronro = " & NroProcesoBatch
'objConn.Execute StrSql, , adExecuteNoRecords
       

Dim aux_TBaj
Dim aux_TDot
Dim j
Dim totales(13, 2)

aux_TBaj = 0
aux_TDot = 0

StrSql = "SELECT * FROM Rep_rotacion"
StrSql = StrSql & " WHERE "
StrSql = StrSql & " bpronro = " & bpronro
OpenRecordset StrSql, rs_sql
'rs_sql.EOF Or
If True Then
Flog.writeline "Entro en If True"
    For X = 1 To v_records
         StrSql = "INSERT INTO  Rep_rotacion (bpronro,Fecha,Hora,iduser,ESTRNRO1,ESTRNRO2,ESTRNRO3"
         StrSql = StrSql & ",MES_ACT,ANIO_ACT"
         StrSql = StrSql & ",MES_BAJA_1,MES_DOTA_1,MES_BAJA_2,MES_DOTA_2,MES_BAJA_3,MES_DOTA_3"
         StrSql = StrSql & ",MES_BAJA_4,MES_DOTA_4,MES_BAJA_5,MES_DOTA_5,MES_BAJA_6,MES_DOTA_6"
         StrSql = StrSql & ",MES_BAJA_7,MES_DOTA_7,MES_BAJA_8,MES_DOTA_8,MES_BAJA_9,MES_DOTA_9"
         StrSql = StrSql & ",MES_BAJA_10,MES_DOTA_10,MES_BAJA_11,MES_DOTA_11,MES_BAJA_12,MES_DOTA_12"
         StrSql = StrSql & ",TOT_BAJAS,TOT_DOTA,MES_ACTUAL"
         StrSql = StrSql & ") VALUES ("
         StrSql = StrSql & bpronro & ","
         StrSql = StrSql & ConvFecha(Fecha) & ","
         StrSql = StrSql & "'" & Format(Hora, "hh:mm:ss") & "',"
         StrSql = StrSql & "'" & IdUser & "',"
         StrSql = StrSql & Estruct & ","
         StrSql = StrSql & Sbestr0 & ","
         StrSql = StrSql & m_baja(X, 0) & ","
         
         StrSql = StrSql & mes & ","
         StrSql = StrSql & Anio & ","
         
         aux_TBaj = 0
         aux_TDot = 0
         
         For j = 1 To 12
            If IsNull(m_baja(X, j)) Or IsEmpty(m_baja(X, j)) Then
                StrSql = StrSql & "Null,"
             Else
                StrSql = StrSql & m_baja(X, j) & ","
                aux_TBaj = aux_TBaj + m_baja(X, j)
                totales(j, 1) = totales(j, 1) + m_baja(X, j)
            End If
            If IsNull(m_dota(X, j)) Or IsEmpty(m_dota(X, j)) Then
                StrSql = StrSql & "Null,"
             Else
                StrSql = StrSql & m_dota(X, j) & ","
                aux_TDot = aux_TDot + m_dota(X, j)
                totales(j, 2) = totales(j, 2) + m_dota(X, j)
            End If
         Next j
         
         If IsNull(aux_TBaj) Or IsEmpty(aux_TBaj) Then
             StrSql = StrSql & "Null,"
          Else
             StrSql = StrSql & aux_TBaj & ","
             totales(13, 1) = totales(13, 1) + aux_TBaj
         End If

         If IsNull(aux_TDot) Or IsEmpty(aux_TDot) Then
             StrSql = StrSql & "Null,"
          Else
             StrSql = StrSql & aux_TDot & ","
             totales(13, 2) = totales(13, 2) + aux_TDot
         End If
         
         StrSql = StrSql & mes & ")"
         
         'Flog.writeline "StrSql= " & StrSql
         
         objConn.Execute StrSql, , adExecuteNoRecords
                  
    Next X
         
         StrSql = "INSERT INTO  Rep_rotacion (bpronro,Fecha,Hora,iduser,ESTRNRO1,ESTRNRO2,ESTRNRO3"
         StrSql = StrSql & ",MES_ACT,ANIO_ACT"
         StrSql = StrSql & ",MES_BAJA_1,MES_DOTA_1,MES_BAJA_2,MES_DOTA_2,MES_BAJA_3,MES_DOTA_3"
         StrSql = StrSql & ",MES_BAJA_4,MES_DOTA_4,MES_BAJA_5,MES_DOTA_5,MES_BAJA_6,MES_DOTA_6"
         StrSql = StrSql & ",MES_BAJA_7,MES_DOTA_7,MES_BAJA_8,MES_DOTA_8,MES_BAJA_9,MES_DOTA_9"
         StrSql = StrSql & ",MES_BAJA_10,MES_DOTA_10,MES_BAJA_11,MES_DOTA_11,MES_BAJA_12,MES_DOTA_12"
         StrSql = StrSql & ",TOT_BAJAS,TOT_DOTA,MES_ACTUAL"
         StrSql = StrSql & ") VALUES ("
         StrSql = StrSql & bpronro & ","
         StrSql = StrSql & ConvFecha(Fecha) & ","
         StrSql = StrSql & "'" & Format(Hora, "hh:mm:ss") & "',"
         StrSql = StrSql & "'" & IdUser & "',"
         StrSql = StrSql & Estruct & ","
         StrSql = StrSql & Sbestr0 & ","
         StrSql = StrSql & "0,"
         
         StrSql = StrSql & mes & ","
         StrSql = StrSql & Anio & ","
         
         
         For j = 1 To 13
            If IsNull(totales(j, 1)) Or IsEmpty(totales(j, 1)) Then
                StrSql = StrSql & "Null,"
             Else
                StrSql = StrSql & totales(j, 1) & ","
            End If
            
            If IsNull(totales(j, 2)) Or IsEmpty(totales(j, 2)) Then
                StrSql = StrSql & "Null,"
             Else
                StrSql = StrSql & totales(j, 2) & ","
            End If

         Next j
         
         StrSql = StrSql & mes & ")"
         
         objConn.Execute StrSql, , adExecuteNoRecords
    
End If
           
'Actualizo el progreso del Proceso
Progreso = Progreso + IncPorc
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
         "' WHERE bpronro = " & NroProcesoBatch
objConn.Execute StrSql, , adExecuteNoRecords
            
Flog.writeline "Fin de la transaccion"
'Fin de la transaccion
MyCommitTrans

rs_sql.Close
Set rs_sql = Nothing


Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "==================================================================================="
    Flog.writeline Err.Description
    Flog.writeline "Ultimo SQL Ejecutado:"
    Flog.writeline StrSql
    Flog.writeline "==================================================================================="
End Sub


Public Sub cargar_mes(ByVal mes As Long, ByRef rs_sql As Recordset, ByVal v_records, ByRef m_rota, ByVal suma, ByVal boolDotacion As Boolean)
    Dim X
    rs_sql.MoveFirst
    For X = 1 To v_records
        rs_sql.MoveFirst
        Do While Not rs_sql.EOF
            If rs_sql!estrnro = m_rota(X, 0) Then
                
                If suma = 1 Then
                    m_rota(X, mes) = m_rota(X, mes) + 1
                Else
                    m_rota(X, mes) = m_rota(X, mes) - 1
                End If
                
                If boolDotacion Then
                    If suma = 1 Then 'Si es una baja por contratos no agrego el ternro ya que resta, pero es de dotacion.
                        ternroAlta = ternroAlta & "," & rs_sql!ternro
                    End If
                Else
                    ternroBaja = ternroBaja & "," & rs_sql!ternro
                End If
                
                'If rs_sql.EOF Then
                '    Exit For
                'End If
            End If
            rs_sql.MoveNext
        Loop
     Next

'    Dim X
'    rs_sql.MoveFirst
'    For X = 1 To v_records
'        If rs_sql!estrnro = m_rota(X, 0) Then
'
'            If suma = 1 Then
'                m_rota(X, mes) = m_rota(X, mes) + rs_sql!cant
'             Else
'                m_rota(X, mes) = m_rota(X, mes) - rs_sql!cant
'            End If
'
'            rs_sql.MoveNext
'
'            If rs_sql.EOF Then
'                Exit For
'            End If
'        End If
'     Next
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer

Dim Finicio As Date 'Long
Dim Estruct As Long
Dim Subestr As String
Dim Agencia As Long
Dim Agencianro As Long
Dim Orden As String
Dim Ordenado As String
Dim Festruct As Date

Dim Separador As String

On Error GoTo CE

Separador = "@"
' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Finicio = CDate(Mid(parametros, pos1, pos2))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Subestr = (Mid(parametros, pos1, pos2 - pos1 + 1))
                
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estruct = (Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Agencia = (Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Agencianro = (Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Orden = (Mid(parametros, pos1, pos2 - pos1 + 1))
 
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Ordenado = (Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Festruct = (Mid(parametros, pos1, pos2 - pos1 + 1))
        
    End If
End If

Flog.writeline "Llamo a rotacion_emp"
Call rotacion_emp(bpronro, Finicio, Estruct, Subestr, Agencia, Agencianro, Orden, Ordenado, Festruct)

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "==================================================================================="
    Flog.writeline Err.Description
    Flog.writeline "==================================================================================="
End Sub
