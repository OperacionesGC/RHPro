Attribute VB_Name = "MdlSICERE"
Option Explicit
'------------------------------------------------------------------------------------------
'Autor: Gonzalez Nicolás
'------------------------------------------------------------------------------------------
'Const Version = 1
'Const FechaVersion = "29/12/2010"
'Const Version = "1.01"
'Const FechaVersion = "23/05/2011"
'Modificado por Gonzalez Nicolás
'Const Version = "1.02"
'Const FechaVersion = "26/05/2011"
'Modificado por Gonzalez Nicolás
'Const Version = "1.03"
'Const FechaVersion = "27/05/2011"
'Modificado por Gonzalez Nicolás
'Const Version = "1.04"
'Const FechaVersion = "31/05/2011"
'Const Version = "1.05"
'Const FechaVersion = "27/06/2011"
'Const Version = "1.06"
'Const FechaVersion = "28/06/2011"
'Modificado por Gonzalez Nicolás
'Const Version = "1.07"
'Const FechaVersion = "18/07/2011"
'Const Version = "1.08"
'Const FechaVersion = "18/07/2011"
'Modificado por Gonzalez Nicolás
'Const Version = "1.09"
'Const FechaVersion = "20/03/2012"
'Modificado por Carmen Quintero (15426) - Se modifico la consulta que busca los empleados, que estan contenidos en el reporte

'Const Version = "1.10"
'Const FechaVersion = "26/09/2012" ' Se redefinió separa_lic (10) a (100) y separa_nov(10)  a (100)
'Modificado por Gonzalez Nicolás - CAS-17003 - Sykes - BUG SICERE

'Const Version = "1.11"
'Const FechaVersion = "01/10/2012" ' se corrigieron errores s/caso CAS-17003 - Sykes - Errores SICERE
'Modificado por Sebastian Stremel

'Const Version = "1.12"
'Const FechaVersion = "10/10/2012" ' se corrigieron errores s/caso CAS-17003 - Sykes - Errores SICERE
'Modificado por Sebastian Stremel

'Const Version = "1.13"
'Const FechaVersion = "22/10/2012" ' se limpio la variable nombre y nombre2  s/caso CAS-17003 - Sykes - Errores SICERE
'Modificado por Sebastian Stremel

'Const Version = "1.14"
'Const FechaVersion = "25/10/2012" ' se limpio la variable nombre y nombre2  s/caso CAS-17003 - Sykes - Errores SICERE
'Modificado por Sebastian Stremel

'Const Version = "1.15"
'Const FechaVersion = "26/10/2012" ' se limpio la variable nombre y nombre2  s/caso CAS-17003 - Sykes - Errores SICERE
'Modificado por Sebastian Stremel

'Const Version = "1.16"
'Const FechaVersion = "31/10/2012" ' se realizaron las modificaciones solicitadas en el s/caso CAS-17003 - Sykes - Errores SICERE
'Modificado por Sebastian Stremel

'Const Version = "1.17"
'Const FechaVersion = "01/11/2012" ' se realizaron las modificaciones solicitadas en el s/caso CAS-17003 - Sykes - Errores SICERE
'Modificado por Sebastian Stremel

'Const Version = "1.18"
'Const FechaVersion = "18/02/2013" ' CAS-17534 - Sykes Costa Rica - Cambio SICERE - Se agregó busqueda de empleados seleccionados desde el filtro.
'Modificado por Gonzalez Nicolás

'Const Version = "1.19"
'Const FechaVersion = "10/04/2013" ' CAS-18904 - SYKES COSTA RICA - Errores en exportacion SICERE - Se modificó la fecha que se utiliza para mostrar en el registro tipo 25 se toma mes y año de la fecha del periodo.
'Modificado por Gonzalez Nicolás

Const Version = "1.20"
Const FechaVersion = "10/06/2014" ' CAS-24569 - SYKES - ERROR EN REPORTE LEGAL SICERE (Costa Rica)
'                                   En caso de que no encuentre una exclusion, si el empleado se encuentra en un proceso
'                                   y ya fue dado de baja lo busco 2 periodos mas atras
'Modificado por NG - LM - FB

Public Type TipoRestriccion
    Estrnro As Long
    Valor As Double
End Type

Global nListaProc As String

'----------------------------------------------------------
Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte SICERE.
' Autor      : Gonzalez Nicolás
' Fecha      : 29/12/2010
' Descripcion:
' Ultima Mod.: 27/05/2011 - Gonzalez Nicolás - Se agregó progreso en batch_proceso
'                         - Se agrego un dato configurable por confrep
' Ultima Mod.: 19/07/2011 - Gonzalez Nicolás - Se agregó orden por apellido y nombre.
'                         - Cuando hay un tipo de jornada, si es 0 o NULL completa con espacio
' ---------------------------------------------------------------------------------------------
'Dim objconnMain As New ADODB.Connection
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

    On Error GoTo ME_Main
    Nombre_Arch = PathFLog & "Generacion_Reporte_SICERE" & "-" & NroProcesoBatch & ".log"
    
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
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'Flog.Writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 283 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    bprcparam = rs_batch_proceso!bprcparam
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call SICERE(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "No encontró el proceso"
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






Public Sub SICERE(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte del SICERE
' Autor      : Gonzalez Nicolás
' Fecha      : 29/12/2010
'            : 10/04/2013 - Gonzalez Nicolás - Se modificó la fecha que se utiliza para mostrar en el registro tipo 25 se toma mes y año de la fecha del periodo.
' --------------------------------------------------------------------------------------------

'Defino Variables
Dim StrSqlU As String
Dim Empresa As Long
Dim empresa_pat As String
Dim StrSqlSA As String

Dim pos1 As Integer
Dim pos2 As Integer
Dim Nroliq As Integer

Dim rs_patrono As New ADODB.Recordset
'Dim objconnProgreso As New ADODB.Connection
'OpenConnection strconexion, objconnProgreso

Dim Lista_Pro
'Dim Lista_Pro_F



'Guarda fecha del período
Dim aux_fecha As Date
Dim Fecha_Inicio_periodo As Date
Dim Fecha_Fin_Periodo As Date
Dim pliqmes As Integer


Dim aux_fecha_insert As String


'Variables de confrep
Dim sNeto As Long
Dim OcuCCSS As Integer
Dim SucCCSS As Integer
Dim JorCCSS As String
'Dim PermisoCCSS As String
'Dim incCCSS As String
'Dim icCCSS As String
'Dim PensionCCSS As String
'Dim excCCSS As String
'Dim classCCSS As String


Dim cod_estr1 As String

'Separa tipo de Lic si hay + de 1
Dim auxlic As String

Dim separa_lic(100) As String
Dim cont_lic As Integer
Dim a As Integer

Dim aux

'Separa tipo de Nov. si hay + de 1
Dim auxnov As String
Dim cont_nov As Integer
Dim separa_nov(100) As String

Dim cant_empresa As Integer
Dim aux_emp As String
Dim Empresa2(10) As Integer
Dim ax As Integer



Dim nro_patrono As String
Dim aux1Patrono As String
Dim aux2Patrono As String
Dim aux3Patrono As String
Dim aux4Patrono As String

Dim nro_empleado As String
'Dim idemp As Integer
Dim idemp As Long
Dim cont35 As Integer

Dim cont35aux As Integer

Dim cont25 As Integer
Dim nrodoc As String
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim cod_ocupacional As String
Dim codigo_ocupacional As String
Dim neto As String
Dim aux_neto
Dim con_sicere As String
Dim inex As String
Dim inex2 As String
Dim cj As String
Dim tipo_inex As String
Dim tipo_inex2 As String
Dim control_reg As String

Dim tc_oc As String
Dim aux_periodo
Dim pergtiaux
Dim perfecini
Dim perfechasta
'Lich0 10/6
Dim perfecini2
Dim perfechasta2

Dim tipodocext 'tipo documento del extranjero

Dim Sect As Integer
Dim NatPAt As Integer
Dim SegPlan As Integer

'Dim grabo_ok As String
Dim CEmpleadosAProc As Long
Dim IncPorc As Double
Dim Progreso As Double

Dim Exc_Forzada As String

aux_emp = 0
aux_fecha_insert = ""
    
'Inicio codigo ejecutable
'On Error GoTo ME_Main
Dim arr_lista
'Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline "levantando parametros " & parametros
Dim TiempoInicialProceso As Long
'TiempoInicialProceso = GetTickCount
OpenConnection strconexion, objConn

If Not IsNull(parametros) Then

    If Len(parametros) >= 1 Then
        
        arr_lista = Split(parametros, "@", -1, 1)
        Nroliq = arr_lista(0)
        Empresa = arr_lista(2)
        
        'Si pos2 = -1 hay un solo proceso
        pos1 = 1
        pos2 = InStr(pos1, arr_lista(1), ",") - 1
                
        'Si hay más de un proceso lo separa en un array
        If pos2 > -1 Then
            Lista_Pro = Split(arr_lista(1), ",", -1, 1)
        Else
            Lista_Pro = Split(arr_lista(1), ",", -1, 1)
        End If
        
        Flog.writeline "Parametro Lista_Pro = " & Lista_Pro(0)

        Flog.writeline "Parametro NroLiq = " & Nroliq
        nListaProc = arr_lista(1)
       
       
        'Asigno el valor de lista de proceso a la variable global para poder usar en el SICERE
        nListaProc = arr_lista(1)
               
    End If
Else
    Flog.writeline "Parametros nulos"
        
End If

Flog.writeline "Terminó de levantar los parametros"
Flog.writeline ""
'-------------------------------------------------------------------
'CARGO EL PERIODO
'-------------------------------------------------------------------
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_patrono
If rs_patrono.EOF Then
    Flog.writeline "No se encontró el Periodo asociado al proceso " & Lista_Pro(0)
    Exit Sub
End If
'Guardo fecha de inicio y fin del periodo
aux_fecha = rs_patrono!pliqhasta
pliqmes = rs_patrono!pliqmes
Fecha_Inicio_periodo = rs_patrono!pliqdesde
Fecha_Fin_Periodo = rs_patrono!pliqhasta

'--------------------------------------------------------------------
'TRAE DATOS FIJOS CONFIGURABLES CONFREP
'--------------------------------------------------------------------
StrSql = "SELECT * FROM  confrep WHERE repnro = 299 "
StrSql = StrSql & "ORDER BY confnrocol "
OpenRecordset StrSql, rs_patrono

If rs_patrono.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
Else
    Do While Not rs_patrono.EOF
        Select Case rs_patrono!conftipo
            Case "SA":
                    sNeto = Trim(rs_patrono!confval)
            Case "TC":
                    If rs_patrono!confval2 = "SUC" Then
                        SucCCSS = Trim(rs_patrono!confval)
                    ElseIf rs_patrono!confval2 = "OCU" Then
                        OcuCCSS = Trim(rs_patrono!confval)
                    End If
            Case "JO":
                'trae el numero de codigo
                JorCCSS = Trim(rs_patrono!confval)
            Case "DOC"
                If rs_patrono!confval2 = "NP" Then
                    NatPAt = Trim(rs_patrono!confval)
                ElseIf rs_patrono!confval2 = "SP" Then
                    SegPlan = Trim(rs_patrono!confval)
                ElseIf rs_patrono!confval2 = "SE" Then
                    Sect = Trim(rs_patrono!confval)
                End If
            Case "EXT"
                tipodocext = Trim(rs_patrono!confval)
        End Select
        rs_patrono.MoveNext
    Loop
    
    'Valida si algun concepto figura vacío
    'If IsEmpty(sNeto) Or IsEmpty(SucCCSS) Or IsEmpty(OcuCCSS) Or IsEmpty(JorCCSS) Or IsEmpty(PermisoCCSS) Or IsEmpty(incCCSS) Or IsEmpty(excCCSS) Or IsEmpty(classCCSS) Then
    '    Flog.writeline "Falta configurar por lo menos un concepto"
    'End If
End If


'------------------------
'Busca todas las empresas
If Empresa = 1 Then
   StrSql = "select count(*) cant_emp from empresa "
   OpenRecordset StrSql, rs_patrono
   cant_empresa = rs_patrono!cant_emp
        
   StrSql = "SELECT empnro,estrnro FROM empresa "
   OpenRecordset StrSql, rs_patrono
   
  'Guarda el N° de estructura de cada empresa
   Do While Not rs_patrono.EOF
        Empresa2(aux_emp) = rs_patrono!Estrnro
        aux_emp = aux_emp + 1
        rs_patrono.MoveNext
   Loop
Else
    Empresa2(0) = Empresa
    cant_empresa = 1
End If
'---------------

cont25 = 0

For ax = 0 To (cant_empresa - 1)
    Empresa = Empresa2(ax)
    'cont25 = cont25 + 1
    '--------------------------------------------------------------------
    'TRAE NÚMEROS DE DOCUMENTOS ASOCIADOS A LA EMPRESA SEGÙN TIPO
    '--------------------------------------------------------------------
    StrSql = "SELECT empresa.empnom,empresa.ternro,ter_doc.tidnro,ter_doc.nrodoc "
    StrSql = StrSql & "FROM Empresa "
    StrSql = StrSql & "INNER JOIN ter_doc on empresa.ternro = ter_doc.ternro "
    StrSql = StrSql & "WHERE (ter_doc.tidnro = " & NatPAt & " OR ter_doc.tidnro = 6 OR ter_doc.tidnro = " & SegPlan & " OR ter_doc.tidnro = " & Sect & ") "
    StrSql = StrSql & "AND empresa.estrnro =" & Empresa
    OpenRecordset StrSql, rs_patrono
    
       
    If rs_patrono.EOF Then
        'cont25 = cont25 - 1
        Flog.writeline "No se encuentra número patronal para la empresa " & Empresa
              
        'Exit Sub
    Else
        nro_patrono = ""
        
        Do While Not rs_patrono.EOF
            Select Case rs_patrono!tidnro
                Case NatPAt:
                     aux1Patrono = rs_patrono!nrodoc
                Case 6:
                    pos1 = InStr(rs_patrono!nrodoc, " ")
                    aux2Patrono = rs_patrono!nrodoc
                    If pos1 <> 0 Then
                        pos2 = Len(rs_patrono!nrodoc)
                        'Corta cadena que se antepone al N°
                        aux2Patrono = Mid(aux2Patrono, pos1, pos2)
                        aux2Patrono = Replace(aux2Patrono, "-", "")
                        aux2Patrono = Replace(aux2Patrono, "_", "")
                        aux2Patrono = Replace(aux2Patrono, ".", "")
                    End If
                        aux2Patrono = Replace(aux2Patrono, "-", "")
                        aux2Patrono = Replace(aux2Patrono, "_", "")
                        aux2Patrono = Replace(aux2Patrono, ".", "")
                    'Si la cantidad es inferior a 11 completar con 0 a la izq.
                    If Len(Trim(aux2Patrono)) <= 11 Then
                        aux2Patrono = Format_StrNro(Trim(aux2Patrono), 11, True, 0)
                    Else
                        Flog.writeline "No se encontró el número de documento tipo 6"
                    End If
                Case SegPlan:
                    aux3Patrono = Format_StrNro(rs_patrono!nrodoc, 3, True, 0)
                    
                Case Sect:
                    aux4Patrono = Format_StrNro(rs_patrono!nrodoc, 3, True, 0)
            End Select
            
            rs_patrono.MoveNext
        Loop
        
        'Registro de N° Patronal + N° Identificativo del patrono
        nro_patrono = "25" + aux1Patrono + aux2Patrono + aux3Patrono + aux4Patrono
        
        
        'TRAE EL COD. SUCURSAL + CODIGO OCUPACIONAL
        StrSql = "SELECT tcodnro,nrocod FROM estr_cod "
        StrSql = StrSql & "WHERE tcodnro = " & SucCCSS & " or tcodnro = " & OcuCCSS & " And Estrnro =" & Empresa & " ORDER BY  nrocod"
        OpenRecordset StrSql, rs_patrono
        If rs_patrono.EOF Then
            Flog.writeline "No se encuentra número de sucursal CCSS"
        Else
            Do While Not rs_patrono.EOF
                If rs_patrono!tcodnro = SucCCSS Then
                    'Agrega el N° de Sucursal de CCSS + fecha
                    'nro_patrono = nro_patrono + rs_patrono!nrocod + CStr(Year(Date)) + Format_StrNro(Month(Date), 2, True, "0")
                    nro_patrono = nro_patrono + rs_patrono!nrocod + CStr(Year(Fecha_Inicio_periodo)) + Format_StrNro(Month(Fecha_Inicio_periodo), 2, True, "0")
                ElseIf rs_patrono!tcodnro = OcuCCSS Then
                    'Guarda el n° de cod. ocupacional | Es igual para todos.
                    'cod_ocupacional = rs_patrono!nrocod
                    codigo_ocupacional = rs_patrono!nrocod
                End If
                rs_patrono.MoveNext
            Loop
        End If
          
        'Flog.writeline
        'Flog.writeline nro_patrono
        
        rs_patrono.Close
               
        
    '--------------------------------------------------------------------
    'GENERA DATOS DE LOS EMPLEADOS
    '--------------------------------------------------------------------
        a = 0

        StrSql = "SELECT DISTINCT Empleado.empleg , Empleado.Ternro, Empleado.terape, Empleado.terape2, Empleado.ternom, Empleado.ternom2 "
        StrSql = StrSql & ",("
        StrSql = StrSql & " SELECT CASE WHEN(ammonto IS NULL) THEN '0' "
        StrSql = StrSql & " ELSE ammonto END "
        StrSql = StrSql & " From acu_mes "
        StrSql = StrSql & " WHERE ammes = '" & pliqmes & "' AND amanio = '" & Year(Fecha_Inicio_periodo) & "' "
        StrSql = StrSql & " AND acunro = " & sNeto & "  AND ternro = cabliq.empleado "
        StrSql = StrSql & ") ammonto "
        StrSql = StrSql & " ,'' Exclusion_forzada "
        StrSql = StrSql & " From Empleado "
        StrSql = StrSql & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
        StrSql = StrSql & " INNER JOIN his_estructura ON cabliq.empleado  = his_estructura.ternro "
        StrSql = StrSql & " Where "
        'StrSql = StrSql & " cabliq.pronro = " & Lista_Pro(a) & ""
        StrSql = StrSql & " cabliq.pronro IN (" & nListaProc & ")"
        StrSql = StrSql & " AND his_estructura.estrnro = " & Empresa
        StrSql = StrSql & " AND his_estructura.htethasta is null "
        StrSql = StrSql & " AND Empleado.Ternro IN( "
        StrSql = StrSql & " SELECT DISTINCT(cabliq.empleado)from cabliq "
        
        'Comentado por Carmen Quintero 20/03/2012
        'StrSql = StrSql & " INNER JOIN proceso on proceso.pronro = cabliq.pronro WHERE  Month(proceso.profecini) = '" & pliqmes & "' and  Year(proceso.profecini) = '" & Year(Fecha_Inicio_periodo) & "' and proceso.proaprob <> 0 "
        
        StrSql = StrSql & " INNER JOIN proceso on proceso.pronro = cabliq.pronro WHERE proceso.pronro IN (" & nListaProc & ")  AND proceso.proaprob <> 0 "
        StrSql = StrSql & ") "
        
        
        'V1.18 - BUSCA EMPLEADOS SELECCIONADOS EN EL FILTRO SIN LIQ.
        StrSql = StrSql & " UNION "
        StrSql = StrSql & " SELECT Empleado.empleg , Empleado.Ternro, Empleado.terape, Empleado.terape2, Empleado.ternom, Empleado.ternom2 ,'0' AMMONTO"
        StrSql = StrSql & " ,'EF' Exclusion_forzada "
        StrSql = StrSql & " FROM batch_empleado"
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro"
        StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro  = his_estructura.ternro "
        StrSql = StrSql & " WHERE bpronro = " & bpronro
        StrSql = StrSql & " AND his_estructura.estrnro = " & Empresa
        
        StrSql = StrSql & " ORDER BY empleado.terape,empleado.terape2,empleado.ternom ASC"
                
        'If UBound(Lista_Pro) <= 1 Then
            'StrSql = StrSql & " ORDER BY empleado.empleg "
        '    StrSql = StrSql & " ORDER BY empleado.terape,empleado.terape2,empleado.ternom ASC"
        '    OpenRecordset StrSql, rs_patrono
        'Else
        '    For a = 1 To UBound(Lista_Pro)
        '        If a = 1 Then
        '            StrSqlU = "(" & StrSql & ")" & " UNION " & "(" & StrSql & ")"
                    'StrSqlU = StrSqlU & " ORDER BY empleado.terape,empleado.terape2,empleado.ternom ASC"
        '        Else
        '            StrSqlU = StrSqlU & " UNION " & "(" & StrSql & ")"
                    'StrSqlU = StrSqlU & " ORDER BY empleado.terape,empleado.terape2,empleado.ternom ASC"
        '        End If
        '        OpenRecordset StrSqlU, rs_patrono
        '    Next
        'End If
        OpenRecordset StrSql, rs_patrono
         Flog.writeline StrSql
      'Determino la proporcion de progreso
        Progreso = 0
        CEmpleadosAProc = rs_patrono.RecordCount
        If CEmpleadosAProc = 0 Then
           CEmpleadosAProc = 1
        End If
        IncPorc = (99 / CEmpleadosAProc)
        
        If rs_patrono.EOF Then
            Flog.writeline "No existen empleados asociados a una liquidación"
        Else

            
            'GRABA EL N° PATRONAL
            'StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,condicion,fecha) "
            StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
            StrSql = StrSql & "VALUES "
            StrSql = StrSql & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & Empresa & ",'',"
            StrSql = StrSql & "'" & nro_patrono & "',''"
            StrSql = StrSql & ",'','','',"
            'StrSql = StrSql & "'','','')"
            StrSql = StrSql & "'','','','','','',NULL,NULL)"
            objConn.Execute StrSql, , adExecuteNoRecords
            '-------------------------------------------
            Flog.writeline "Nº Patronal Insertado"
            cont25 = cont25 + 1
            empresa_pat = Empresa
            nro_empleado = "35"
            'idemp = rs_patrono!empleg
            
            
            'Cuenta la cantidad de registros con Nº35
            'cont35 = 0
            Do While Not rs_patrono.EOF
                'limpio las variables
                nombre = ""
                nombre2 = ""
                
                'SI EL EMPLEADO SE EXCLUYE FORZADAMENTE POR EL FILTRO SE LE ASIGNA -> EF
                Exc_Forzada = rs_patrono!Exclusion_forzada
                
                'Incremento el progreso
                Progreso = Progreso + IncPorc
                
                Flog.writeline "=================================================="
                Flog.writeline "Legajo: " & rs_patrono!empleg
                Flog.writeline "Apellido: " & rs_patrono!terape & " " & rs_patrono!terape2
                Flog.writeline "=================================================="
            
            
            
                idemp = rs_patrono!empleg
                cont35 = cont35 + 1
                
                'TRAE NUMERO DE DOC. + ID
                aux = Split(emp_nrodoc(rs_patrono!Ternro, tipodocext), "@")
                                               
                If aux(1) <> 2 And aux(1) <> tipodocext Then
                    'Si es nativo
                    nro_empleado = nro_empleado + "0"
                    
                ElseIf tipodocext = aux(1) Then
                    'Si es extranjero
                    nro_empleado = nro_empleado + "7"
                Else
                    nro_empleado = nro_empleado + "0"
                End If
                
                'Trae el tipo de codigo ocupacional
                cod_ocupacional = codidgo_asoc_emp(rs_patrono!Ternro, Fecha_Inicio_periodo, Fecha_Fin_Periodo, OcuCCSS)
                If cod_ocupacional = "" Then
                    cod_ocupacional = codigo_ocupacional
                End If
                
                              
                'TRAE N° DE DOC. DEL EMPLEADO
                nrodoc = aux(0)
                
                'Trae Apellido 1, Apellido 2 y Nombre
                apellido = Format_Str(rs_patrono!terape, 20, True, " ")
                If Trim(rs_patrono!terape2) <> "" Then
                    apellido2 = Format_Str(rs_patrono!terape2, 20, True, " ")
                Else
                    apellido2 = Format_Str("NOINDICAOTRO", 20, True, " ")
                End If
                
                'sebastian stremel trae segundo nombre 01/10/2012
                nombre = rs_patrono!ternom
                If Trim(rs_patrono!ternom2) <> "" Then
                    nombre2 = rs_patrono!ternom2
                    'nombre = rs_patrono!ternom & " " & nombre2
                    nombre = nombre & " " & nombre2
                End If
                
                
                nombre = Format_Str(nombre, 60, True, " ")
                'hasta aca
                
                'nombre = Format_Str(rs_patrono!ternom, 60, True, " ")
                
                'Formatea y Guarda el NETO
                If rs_patrono!ammonto = "" Or IsNull(rs_patrono!ammonto) = True Then
                    neto = 0
                    neto = Format_StrNro(neto, 15, True, "0")
                Else
                    aux_neto = FormatNumber(rs_patrono!ammonto, 2)
                    'aux_neto = CStr(aux_neto)
                    neto = Replace(aux_neto, ",", "")
                    aux_neto = Replace(neto, ".", "")
                    neto = Format_StrNro(aux_neto, 15, True, "0")
                End If
                                
                'Guarda los datos personales del empleado
                'nro_empleado = nro_empleado & nrodoc & UCase(apellido) & UCase(apellido2) & UCase(nombre) & cod_ocupacional
                                
                'DEVUELVE LA FECHA DEL PERIODO GTI P/ LICENCIAS/PERMISOS Y NOVEDADES 'sebastian stremel, lo pidio javier irastorza 10/10/2012
                aux_periodo = periodo_gti(rs_patrono!Ternro, Fecha_Inicio_periodo, Fecha_Fin_Periodo)
                If aux_periodo = False Then
                    Flog.writeline "No se encontró el Periodo GTI para el empleado: " & rs_patrono!terape & " " & rs_patrono!terape2 & " " & rs_patrono!ternom
                    
                Else
                    pergtiaux = Split(aux_periodo, "@")
                    perfecini = pergtiaux(0)
                    perfechasta = pergtiaux(1)
                End If
                
                'valida si es jubilado devuelve (A-C)
                'con_sicere = clase_seguro(rs_patrono!Ternro, Fecha_Inicio_periodo, Fecha_Fin_Periodo)
                con_sicere = clase_seguro(rs_patrono!Ternro, perfecini, perfechasta)
                              
                'Trae Inclusion (IN) o Exclusion (EX)
                'inex = TC_IN_EX(rs_patrono!Ternro, JorCCSS, Fecha_Inicio_periodo, Fecha_Fin_Periodo)
                inex = TC_IN_EX(rs_patrono!Ternro, JorCCSS, perfecini, perfechasta, "I")
                
                If Exc_Forzada = "EF" Then
                    'SE TOMA LA FECHA DE BAJA LA FECHA DESDE DEL INICIO DEL PERIODO. V1.18
                    inex2 = "EX" & "@" & Fecha_Inicio_periodo
                Else
                    inex2 = TC_IN_EX(rs_patrono!Ternro, JorCCSS, perfecini, perfechasta, "E")
                End If
                'cj = ""
                If (inex = "") And (inex2 = "") Then
                'Trae cambio de jornada del empleado
                    'cj = cambiojornada(rs_patrono!Ternro, JorCCSS, Fecha_Inicio_periodo, Fecha_Fin_Periodo)
                    cj = cambiojornada(rs_patrono!Ternro, JorCCSS, perfecini, perfechasta)
                End If
                '-------------------------------------------
                
                '-------------' 10/06/2014 - licho :S------------------------------
                'Si se encuentra en un proceso y ya fue dado de baja lo busco n periodo mas atras, osea 2
                If Exc_Forzada = "" And inex2 = "" Then
                    aux_periodo = periodo_gti(rs_patrono!Ternro, DateAdd("m", -1, Fecha_Inicio_periodo), DateAdd("m", -1, Fecha_Fin_Periodo))
                    If aux_periodo = False Then
                        Flog.writeline "No se encontró el Periodo GTI V2 para el empleado: " & rs_patrono!terape & " " & rs_patrono!terape2 & " " & rs_patrono!ternom
                    Else
                        pergtiaux = Split(aux_periodo, "@")
                        perfecini2 = pergtiaux(0)
                        perfechasta2 = pergtiaux(1)
                        
                        inex2 = TC_IN_EX(rs_patrono!Ternro, JorCCSS, perfecini2, perfechasta2, "E")
                        If inex2 = "" Then
                            Flog.writeline "No se encontro EXCLUSION para el mes anterior V2."
                        Else
                            Flog.writeline "YUJU encontro EXCLUSION para el mes anterior V2." & inex2
                        End If
                    End If
                End If
                '---------------------------------------------------
                
                
                '-------------------------------------------
                'Devuelve TC OC
                tc_oc = tc_ocupacion(rs_patrono!Ternro, Fecha_Inicio_periodo, Fecha_Fin_Periodo)
                
                'trae código asoc. a estructura
                'cod_estr1 = cod_estr(JorCCSS)
                
                'DEVUELVE LA FECHA DEL PERIODO GTI P/ LICENCIAS/PERMISOS Y NOVEDADES
                'aux_periodo = periodo_gti(rs_patrono!Ternro, Fecha_Inicio_periodo, Fecha_Fin_Periodo)
                If aux_periodo = False Then
                    'Flog.writeline "No se encontró el Periodo GTI para el empleado: " & rs_patrono!terape & " " & rs_patrono!terape2 & " " & rs_patrono!ternom
                    
                Else
                   ' pergtiaux = Split(aux_periodo, "@")
                   ' perfecini = pergtiaux(0)
                   ' perfechasta = pergtiaux(1)
                
                
                 'DEVUELVE SI TIENE LICENCIAS, PERMISOS Y PENSIONES ASOC.
                 'auxlic = licencia(rs_patrono!Ternro, Format_Fecha(perfecini, 1), Format_Fecha(perfechasta, 1))
                 auxlic = licencia(rs_patrono!Ternro, perfecini, perfechasta)
                 'SI HAY MAS DE UN TIPO LAS SEPARA
                 If auxlic <> "" Then
                     pos2 = 0
                     
                     If Len(auxlic) >= 1 Then
                         cont_lic = 0
                         Do Until pos2 = -1
                             pos1 = 1
                             pos2 = InStr(pos1, auxlic, "!") - 1
                             
                             If pos2 = -1 Then
                                 separa_lic(cont_lic) = auxlic
                                 Exit Do
                             End If
                             
                             separa_lic(cont_lic) = Mid(auxlic, pos1, pos2)
                             cont_lic = cont_lic + 1
                             
                             pos1 = pos2 + 2
                             
                             auxlic = Mid(auxlic, pos1, Len(auxlic))
                             pos2 = InStr(pos1, auxlic, "@") - 1
                             
                             If Len(auxlic) > 0 And pos2 = -1 Then
                                 cont_lic = cont_lic + 1
                                 separa_lic(cont_lic) = auxlic
                                 'auxlic = Mid(auxlic, pos1, Len(auxlic))
                             End If
                         Loop
                                           
                     End If
                 End If
                 
                '----------------------
                'DEVUELVE SI EL EMPLEADO TIENE ALGUN PERMISO
                auxnov = pe_novedad(rs_patrono!Ternro, perfecini, perfechasta)
                If auxnov <> "" Then
                     pos2 = 0
                     
                     If Len(auxnov) >= 1 Then
                         cont_nov = 0
                         Do Until pos2 = -1
                             cont_nov = cont_nov + 1
                             pos1 = 1
                             pos2 = InStr(pos1, auxnov, "!") - 1
                             
                             If pos2 = -1 Then
                                cont_nov = cont_nov - 1
                                 separa_nov(cont_nov) = auxnov
                                 Exit Do
                             End If
                             separa_nov(cont_nov) = Mid(auxnov, pos1, pos2)
                             
                             pos1 = pos2 + 2
                             
                             auxnov = Mid(auxnov, pos1, Len(auxnov))
                             pos2 = InStr(pos1, auxnov, "@") - 1
                             
                             If Len(auxnov) > 0 And pos2 = -1 Then
                                 cont_nov = cont_nov + 1
                                 separa_nov(cont_nov) = auxnov
                             End If
                         Loop
                                           
                     End If
                 End If
               End If ' cierra aux_periodo
               '----------------------
               'INSERT EN REPSICERE
               '---------------------
                tipo_inex = Mid(inex, 1, 2)
                tipo_inex2 = Mid(inex2, 1, 2)
                'Guarda en variable la consulta de salario
                StrSqlSA = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
                StrSqlSA = StrSqlSA & "VALUES "
                StrSqlSA = StrSqlSA & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & Empresa & ",'" & idemp & "',"
                StrSqlSA = StrSqlSA & "'" & nro_empleado & "',' " & nrodoc & "'"
                StrSqlSA = StrSqlSA & ",'" & apellido & "','" & apellido2 & "','" & nombre & "',"
                'StrSqlSA = StrSqlSA & "'" & cod_ocupacional & neto & "','" & con_sicere & "SA','" & Format_Fecha(Fecha_Inicio_periodo, 1) & "')"
                StrSqlSA = StrSqlSA & "'" & cod_ocupacional & neto & "'"
                'StrSqlSA = StrSqlSA & ",'" & con_sicere & inex & "','')"
                '------------------------------
                'ver si la fecha de inicio del periodo es correcta con la que hay que mostrar
                'StrSqlSA = StrSqlSA & ",'" & con_sicere & "','SA','','','','" & Format_Fecha(Fecha_Inicio_periodo, 1) & "',NULL)"
                If tipo_inex = "IC" Then
                    aux_fecha_insert = Mid(inex, InStr(inex, "@") + 1, 10)
                    If CDate(aux_fecha_insert) <= CDate(Fecha_Inicio_periodo) Then
                        StrSqlSA = StrSqlSA & ",'" & con_sicere & "','SA','','',''," & ConvFecha(Fecha_Inicio_periodo) & ",NULL)"
                    Else
                        StrSqlSA = StrSqlSA & ",'" & con_sicere & "','SA','','',''," & ConvFecha(aux_fecha_insert) & ",NULL)"
                    End If
                    'StrSqlSA = StrSqlSA & ",'" & con_sicere & "','SA','','',''," & ConvFecha(aux_fecha_insert) & ",NULL)"
                Else
                'seba 31/10
                    'aux_fecha_insert = Mid(inex, InStr(inex, "@") + 1, 10)
                    'If CDate(aux_fecha_insert) <= CDate(Fecha_Inicio_periodo) Then
                        StrSqlSA = StrSqlSA & ",'" & con_sicere & "','SA','','',''," & ConvFecha(Fecha_Inicio_periodo) & ",NULL)"
                    'Else
                    '    StrSqlSA = StrSqlSA & ",'" & con_sicere & "','SA','','',''," & ConvFecha(aux_fecha_insert) & ",NULL)"
                    'End If
                    'StrSqlSA = StrSqlSA & ",'" & con_sicere & "','SA','','',''," & ConvFecha(Fecha_Inicio_periodo) & ",NULL)"
                End If
                '----------------------
                
                If tipo_inex = "IC" Then
                    Flog.writeline "Se inserta IC"
                    'aux_fecha_insert = Left(Mid(inex, 12, 8), 4) & "/" & Mid(Mid(inex, 12, 8), 3, 2) & "/" & Right(Mid(inex, 12, 8), 2)
                    'aux_fecha_insert = Mid(inex, 12, 8)
                    aux_fecha_insert = Mid(inex, InStr(inex, "@") + 1, 10)
                    StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
                    '--------------------------------------
                    StrSql = StrSql & "VALUES "
                    StrSql = StrSql & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & Empresa & ",'" & idemp & "',"
                    StrSql = StrSql & "'" & nro_empleado & "',' " & nrodoc & "'"
                    StrSql = StrSql & ",'" & apellido & "','" & apellido2 & "','" & nombre & "',"
                    'StrSql = StrSql & "'" & cod_ocupacional & neto & "','" & con_sicere & inex & "','" & Format_Fecha(Fecha_Inicio_periodo, 1) & "')"
                    StrSql = StrSql & "'" & cod_ocupacional & neto & "'"
                    'StrSql = StrSql & ",'" & con_sicere & inex & "','')"
                    '------------------------------
                    'StrSql = StrSql & ",'" & con_sicere & "','" & Left(inex, 2) & "','" & Mid(inex, 3, 2) & "','" & Mid(inex, 5, 3) & "','" & Mid(inex, 8, 3) & "','" & Mid(inex, 12, 8) & "',NULL)"
                    If CDate(aux_fecha_insert) <= CDate(Fecha_Inicio_periodo) Then
                        StrSql = StrSql & ",'" & con_sicere & "','" & Left(inex, 2) & "','" & Mid(inex, 3, 2) & "','" & Mid(inex, 5, 3) & "','" & Mid(inex, 8, 3) & "'," & ConvFecha(Fecha_Inicio_periodo) & ",NULL)"
                    Else
                        StrSql = StrSql & ",'" & con_sicere & "','" & Left(inex, 2) & "','" & Mid(inex, 3, 2) & "','" & Mid(inex, 5, 3) & "','" & Mid(inex, 8, 3) & "'," & ConvFecha(aux_fecha_insert) & ",NULL)"
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'seba 31/10
                    cont35 = cont35 + 1
                    aux_fecha_insert = ""
                End If
                   
                If tipo_inex2 = "EX" Then
                    
                    
                    If tipo_inex = "IC" Then
                        'no graba registro salario
                    Else
                        'Graba registro de salario
                        Flog.writeline "Se inserta SA"
                        objConn.Execute StrSqlSA, , adExecuteNoRecords
                    End If
                    'Graba registro de salario
                    'Flog.writeline "Se inserta SA"
                    'objConn.Execute StrSqlSA, , adExecuteNoRecords
                    
                    Flog.writeline inex & " | " & Mid(inex, 4, Len(inex))
                    Flog.writeline "Se inserta EX"
                    'aux_fecha_insert = Left(Mid(inex, 4, Len(inex)), 4) & "/" & Mid(Mid(inex, 4, Len(inex)), 3, 2) & "/" & Right(Mid(inex, 4, Len(inex)), 2)
                    aux_fecha_insert = Mid(inex2, 4, Len(inex2))
                    StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
                    StrSql = StrSql & "VALUES "
                    StrSql = StrSql & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & Empresa & ",'" & idemp & "',"
                    StrSql = StrSql & "'" & nro_empleado & "',' " & nrodoc & "'"
                    StrSql = StrSql & ",'" & apellido & "','" & apellido2 & "','" & nombre & "',"
                    'StrSql = StrSql & "'" & cod_ocupacional & neto & "','" & con_sicere & inex & "','')"
                    
                    StrSql = StrSql & "'" & cod_ocupacional & neto & "'"
                    'StrSql = StrSql & ",'" & con_sicere & inex & "','')"
                    '------------------------------
                    'StrSql = StrSql & ",'" & con_sicere & "','" & Left(inex, 2) & "','','','','" & Mid(inex, 4, Len(inex)) & "',NULL)"
                    If CDate(aux_fecha_insert) <= CDate(Fecha_Inicio_periodo) Then
                        StrSql = StrSql & ",'" & con_sicere & "','" & Left(inex2, 2) & "','','',''," & ConvFecha(Fecha_Inicio_periodo) & ",NULL)"
                    Else
                        StrSql = StrSql & ",'" & con_sicere & "','" & Left(inex2, 2) & "','','',''," & ConvFecha(aux_fecha_insert) & ",NULL)"
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                    cont35 = cont35 + 1
                    aux_fecha_insert = ""
                Else
                    If tipo_inex <> "IC" Then
                        'Graba primer registro de salario
                         objConn.Execute StrSqlSA, , adExecuteNoRecords
                    End If
                    
                    If Not IsEmpty(separa_lic) And auxlic <> "" Then
                       
                       For a = 0 To (cont_lic)
                            If separa_lic(a) <> "" Then
                            
                                Flog.writeline "Se inserta Licencia " & Trim(separa_lic(a))
                                aux_fecha_insert = Left(Mid(separa_lic(a), 15, 8), 4) & "/" & Mid(Mid(separa_lic(a), 15, 8), 5, 2) & "/" & Right(Mid(separa_lic(a), 15, 8), 2)
                                'aux_fecha_insert = Mid(separa_lic(a), 15, 8)
                                StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
                                StrSql = StrSql & "VALUES "
                                StrSql = StrSql & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & Empresa & ",'" & idemp & "',"
                                StrSql = StrSql & "'" & nro_empleado & "',' " & nrodoc & "'"
                                StrSql = StrSql & ",'" & apellido & "','" & apellido2 & "','" & nombre & "',"
                                'StrSql = StrSql & "'" & cod_ocupacional & neto & "','" & con_sicere & separa_lic(a) & "','')"
                                StrSql = StrSql & "'" & cod_ocupacional & neto & "'"
                                'StrSql = StrSql & ",'" & con_sicere & inex & "','')"
                                '------------------------------
                                'StrSql = StrSql & ",'" & con_sicere & "','" & Left(separa_lic(a), 2) & "','','','" & Mid(separa_lic(a), 3, 3) & "','" & Mid(separa_lic(a), 7, 8) & "','" & Mid(separa_lic(a), 15, 8) & "')"
                                StrSql = StrSql & ",'" & con_sicere & "','" & Left(separa_lic(a), 2) & "','','','" & Mid(separa_lic(a), 3, 3) & "','" & Mid(separa_lic(a), 7, 8) & "'," & ConvFecha(aux_fecha_insert) & ")"
                                
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'Flog.writeline separa_lic(a) & " "
                                separa_lic(a) = Empty
                                cont35 = cont35 + 1
                                aux_fecha_insert = ""
                            'Flog.writeline "separa_lic " & StrSql
                            End If
                        Next
                                        
                    End If
                    If Not IsEmpty(separa_nov) And auxnov <> "" Then
                                            
                        For a = 0 To (cont_nov)
                            If separa_nov(a) <> "" Then
                                '
                                If a = 0 And tipo_inex = "IC" Then
                                'Graba registro de salario
                                                                   
                                    'objConn.Execute StrSqlSA, , adExecuteNoRecords
                                    cont35 = cont35 + 1
                                End If
                                
                                Flog.writeline "Se inserta Novedad " & Trim(separa_nov(a))
                                aux_fecha_insert = Left(Mid(separa_nov(a), 14, 8), 4) & "/" & Mid(Mid(separa_nov(a), 14, 8), 5, 2) & "/" & Right(Mid(separa_nov(a), 14, 8), 2)
                                'aux_fecha_insert = Mid(separa_nov(a), 14, 8)
                                StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
                                StrSql = StrSql & "VALUES "
                                StrSql = StrSql & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & Empresa & ",'" & idemp & "',"
                                StrSql = StrSql & "'" & nro_empleado & "',' " & nrodoc & "'"
                                StrSql = StrSql & ",'" & apellido & "','" & apellido2 & "','" & nombre & "',"
                                'StrSql = StrSql & "'" & cod_ocupacional & neto & "','" & con_sicere & separa_nov(a) & "','')"
                                StrSql = StrSql & "'" & cod_ocupacional & neto & "'"
                                'StrSql = StrSql & ",'" & con_sicere & inex & "','')"
                                '------------------------------
                                'StrSql = StrSql & ",'" & con_sicere & "','" & Left(separa_nov(a), 2) & "','','','" & Mid(separa_nov(a), 3, 1) & "','" & Mid(separa_nov(a), 5, 8) & "','" & Mid(separa_nov(a), 14, 8) & "')"
                                StrSql = StrSql & ",'" & con_sicere & "','" & Left(separa_nov(a), 2) & "','','','" & Mid(separa_nov(a), 3, 1) & "','" & Mid(separa_nov(a), 5, 8) & "'," & ConvFecha(aux_fecha_insert) & ")"
                                
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'Flog.writeline "separa_nov " & StrSql
                                separa_nov(a) = Empty
                                cont35 = cont35 + 1
                                aux_fecha_insert = ""
                            End If
                        Next
                    End If
                    
                    If cj <> "" And tipo_inex <> "IC" Then
                        Flog.writeline "Se inserta Cambio de jornada " & Trim(cj)
                        'aux_fecha_insert = Left(Mid(cj, 9, 8), 4) & "/" & Mid(Mid(cj, 9, 8), 3, 2) & "/" & Right(Mid(cj, 9, 8), 2)
                        'aux_fecha_insert = Left(Mid(cj, 7, 8), 4) & "/" & Mid(Mid(cj, 9, 8), 3, 2) & "/" & Right(Mid(cj, 9, 8), 2)
                        'aux_fecha_insert = Mid(cj, 9, 8)
                        aux_fecha_insert = Mid(cj, InStr(cj, "@") + 1, 10)
                        StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
                        StrSql = StrSql & "VALUES "
                        StrSql = StrSql & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & Empresa & ",'" & idemp & "',"
                        StrSql = StrSql & "'" & nro_empleado & "',' " & nrodoc & "'"
                        StrSql = StrSql & ",'" & apellido & "','" & apellido2 & "','" & nombre & "',"
                        'StrSql = StrSql & "'" & cod_ocupacional & neto & "','" & con_sicere & "','" & cj & "')"
                        StrSql = StrSql & "'" & cod_ocupacional & neto & "'"
                        '------------------------------
                        'StrSql = StrSql & ",'" & con_sicere & "','" & Left(cj, 2) & "','" & Mid(cj, 3, 2) & "','" & Mid(cj, 5, 3) & "','','" & Mid(cj, 9, 8) & "',NULL)"
                        StrSql = StrSql & ",'" & con_sicere & "','" & Left(cj, 2) & "','" & Mid(cj, 3, 2) & "','" & Mid(cj, 5, 3) & "',''," & ConvFecha(aux_fecha_insert) & ",NULL)"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        aux_fecha_insert = ""
                        cont35 = cont35 + 1
                   End If
                                   
                    If tc_oc <> "" And tipo_inex <> "IC" Then
                        Flog.writeline "Inserta Tipo de ocupacion " & Trim(tc_oc)
                        'aux_fecha_insert = Left(Mid(tc_oc, 4, 8), 4) & "/" & Mid(Mid(tc_oc, 4, 8), 3, 2) & "/" & Right(Mid(tc_oc, 4, 8), 2)
                        'aux_fecha_insert = Mid(tc_oc, 4, 8)
                        aux_fecha_insert = Mid(tc_oc, InStr(tc_oc, "@") + 1, 10)
                        StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
                        StrSql = StrSql & "VALUES "
                        StrSql = StrSql & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & Empresa & ",'" & idemp & "',"
                        StrSql = StrSql & "'" & nro_empleado & "',' " & nrodoc & "'"
                        StrSql = StrSql & ",'" & apellido & "','" & apellido2 & "','" & nombre & "',"
                        'StrSql = StrSql & "'" & cod_ocupacional & neto & "','" & con_sicere & "','" & tc_oc & "')"
                        
                        StrSql = StrSql & "'" & cod_ocupacional & neto & "'"
                        'StrSql = StrSql & ",'" & con_sicere & inex & "','')"
                        '------------------------------
                        'StrSql = StrSql & ",'" & con_sicere & "','" & Left(tc_oc, 2) & "','','','','" & Mid(tc_oc, 4, 8) & "',NULL)"
                        StrSql = StrSql & ",'" & con_sicere & "','" & Left(tc_oc, 2) & "','','',''," & ConvFecha(aux_fecha_insert) & ",NULL)"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        cont35 = cont35 + 1
                        aux_fecha_insert = ""
                   End If
                
                End If
    
                'Nro fijo se escribe en cada línea para todos los empleados
                nro_empleado = "35"
                           
                rs_patrono.MoveNext
                
                'Inserto progreso
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Int(Progreso)
                TiempoInicialProceso = GetTickCount
                MyBeginTrans
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
                    StrSql = StrSql & " , bprctiempo = " & TiempoInicialProceso
                    StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
                    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                MyCommitTrans
                
            Loop
        End If

        'rs_patrono.Close
    
    End If 'Cierra principal valida parametros
    aux1Patrono = ""
    aux2Patrono = ""
    aux3Patrono = ""
    aux4Patrono = ""
Next


'--------------------------------------------------------------------
    'CONTROL DEL ARCHIVO
    '--------------------------------------------------------------------
    '    If cant_empresa - 1 = ax And empresa_pat <> "" Then
            'control_reg = "15PAT" + Format_Fecha(Date, 1) + Format_StrNro(cont25, 10, True, 0) + Format_StrNro(cont35, 10, True, 0)
            
            'cantidad de lineas seba 01/11/2012
            StrSql = "SELECT COUNT(pat_reg) cant FROM repsicere WHERE pat_reg in('350','357') AND bpronro=" & bpronro
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                cont35aux = rs!Cant
            Else
                cont35aux = 0
            End If
            rs.Close
            
            'hasta aca
            
         control_reg = "15PAT" + Format_Fecha(Date, 1) + Format_StrNro(cont25, 10, True, 0) + Format_StrNro(cont35aux, 10, True, 0)
         
         
'         If Progreso <> 0 Then
'            'Graba fin de archivo
'            StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
'            StrSql = StrSql & "VALUES "
'            StrSql = StrSql & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & empresa_pat & ",'',"
'            StrSql = StrSql & "'" & control_reg & "',''"
'            StrSql = StrSql & ",'','','',"
'            StrSql = StrSql & "'','','','','','',NULL,NULL)"
'            objConn.Execute StrSql, , adExecuteNoRecords
'            Flog.writeline "Control de Archivo Insertado"
'        End If
        
        'V 1.18 - Elimino los empleados de batch_empleado
        StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & bpronro
        objConn.Execute StrSql, , adExecuteNoRecords


'Redondeo a 100%
If Int(Progreso) < 100 Then
    'Inserto progreso
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 100"
    StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Progreso = 100
End If


If Progreso <> 0 Then
    'Graba fin de archivo
    StrSql = "INSERT INTO repsicere (pronro,bpronro,empnro,empleg,pat_reg,nro_doc,apellido,apellido2,nombre,neto,cond_sic,tipo_cambio,num_jor,clase_jor,cod_cambio,fecha_desde,fecha_hasta) "
    StrSql = StrSql & "VALUES "
    StrSql = StrSql & "(" & Lista_Pro(0) & "," & NroProcesoBatch & "," & empresa_pat & ",'',"
    StrSql = StrSql & "'" & control_reg & "',''"
    StrSql = StrSql & ",'','','',"
    StrSql = StrSql & "'','','','','','',NULL,NULL)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Control de Archivo Insertado"
End If

Flog.writeline ""
Flog.writeline "El proceso se realizó con éxito"

End Sub





Public Function clase_seguro(ByVal Ternro, ByVal fecini, ByVal fecfin)
'Si elempleado es jubilado devuelve A si no C
Dim rs_valjul As New ADODB.Recordset

StrSql = "SELECT * FROM his_estructura "
StrSql = StrSql & "INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & "WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 31"
StrSql = StrSql & "AND(htetdesde >= " & ConvFecha(fecini) & " and htethasta <= " & ConvFecha(fecfin)
StrSql = StrSql & " OR (htetdesde <= " & ConvFecha(fecfin) & " and htethasta is null))"

'StrSql = StrSql & "AND(htetdesde >= '" & fecini & "' and htethasta <= '" & fecfin & "' "
'StrSql = StrSql & " OR (htetdesde <= '" & fecfin & "' and htethasta is null))"
OpenRecordset StrSql, rs_valjul

If rs_valjul.EOF Then
    clase_seguro = "C"
ElseIf rs_valjul!estrcodext = "-1" Then
    clase_seguro = "A"
Else
    clase_seguro = "C"
End If

End Function





Public Function tc_ocupacion(ByVal idemp, ByVal fecini, ByVal fecfin)
    'Devuelve OC si hay cambio de ocupacion
    Dim rs_ocup As New ADODB.Recordset
    
    StrSql = "SELECT * FROM his_estructura "
    'StrSql = "WHERE Tenro = 4 AND ternro = " & idemp
    'StrSql = StrSql & "WHERE htetdesde >= " & fecini & " AND (htethasta <= " & fecfin & " OR htethasta is null ) "
    StrSql = StrSql & "WHERE htetdesde >= " & ConvFecha(fecini) & " AND htetdesde <= " & ConvFecha(fecfin) & " AND (htethasta <= " & ConvFecha(fecfin) & " OR htethasta is null ) "
    StrSql = StrSql & "AND tenro = 4 AND ternro = " & idemp
    'StrSql = StrSql & "INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    'StrSql = StrSql & "WHERE his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = 31"
    
    OpenRecordset StrSql, rs_ocup
    
    If rs_ocup.EOF Then
        tc_ocupacion = ""
    Else
        'tc_ocupacion = "OC" & "@" & Format_Fecha(rs_ocup!htetdesde, 1)
        tc_ocupacion = "OC" & "@" & rs_ocup!htetdesde
    End If


End Function



Public Function periodo_gti(ByVal idemp, ByVal fecini, ByVal fecfin)
    'Devuelve OC si hay cambio de ocupacion
    Dim rs_gti As New ADODB.Recordset
    Dim Mes
    Dim Anio
    Mes = Month(fecini)
    Anio = Year(fecini)
    
    StrSql = "SELECT gti_per.pgtidesde, gti_per.pgtihasta"
    StrSql = StrSql & " FROM Empleado"
    StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
    StrSql = StrSql & " INNER JOIN alcance_tEstr ON his_estructura.tenro = alcance_tEstr.tenro And alcance_testr.tanro = 5 "
    StrSql = StrSql & " INNER JOIN gti_per_est ON gti_per_est.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " INNER JOIN gti_per ON gti_per_est.pgtinro = gti_per.pgtinro"
    StrSql = StrSql & " WHERE pgtimes = " & Mes & " And pgtianio = " & Anio
    StrSql = StrSql & " AND gti_per.pgtiestado = -1"
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(fecini) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(fecfin) & " <= his_estructura.htethasta)"
    
    'StrSql = StrSql & " AND (his_estructura.htetdesde <= '01/" & Mes & "/" & Anio & "')"
    'StrSql = StrSql & " AND (('01/" & Mes & "/" & Anio & "' <= his_estructura.htethasta)"
    
    StrSql = StrSql & " OR (his_estructura.htethasta is null))"
    StrSql = StrSql & " AND empleado.ternro = " & idemp
    StrSql = StrSql & " ORDER BY empleado.empleg"
    OpenRecordset StrSql, rs_gti
    
    If rs_gti.EOF Then
        periodo_gti = False
    Else
        periodo_gti = rs_gti!pgtidesde & "@" & rs_gti!pgtihasta
    End If


End Function

Public Function emp_nrodoc(ByVal idemp, ByVal tipodoc)
    Dim rs As New ADODB.Recordset
    Dim nrodoc
    Dim nrodoc_aux
    Dim tidnro
    Dim marca
    marca = 0
    StrSql = "SELECT nrodoc,tidnro FROM ter_doc "
    StrSql = StrSql & "WHERE "
    'StrSql = StrSql & "tidnro <=5 "
    StrSql = StrSql & " (tidnro <=5 OR tidnro = " & tipodoc & ")"
    StrSql = StrSql & " AND ternro = " & idemp
    StrSql = StrSql & " ORDER BY tidnro DESC "
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
        Flog.writeline "No se encuentra número de documento asociado al empleado " & idemp
        emp_nrodoc = Format_StrNro(0, 25, True, 0) & "@0"
    Else
        Do While Not rs.EOF
            If tipodoc = CStr(rs!tidnro) Then
                nrodoc_aux = rs!nrodoc
                tidnro = rs!tidnro
                marca = 1
            ElseIf marca = 0 Then
                nrodoc_aux = rs!nrodoc
                tidnro = rs!tidnro
            End If
            rs.MoveNext
        Loop
        'Trae N° doc y lo formatea // Agrega 0 y elimina  (-) (.)
         'nrodoc = Replace(rs!nrodoc, "-", "")
         nrodoc = Replace(nrodoc_aux, "-", "")
         nrodoc = Replace(nrodoc, ".", "")
         'emp_nrodoc = Format_StrNro(nrodoc, 25, True, 0)
         'emp_nrodoc = Format_StrNro(nrodoc, 25, True, "0") & "@" & rs!tidnro
         emp_nrodoc = Format_StrNro(nrodoc, 25, True, "0") & "@" & tidnro
    End If

End Function


Public Function codidgo_asoc_emp(ByVal idemp, ByVal fecini, ByVal fecfin, ByVal OcuCCSS)
    Dim rs_estructura As New ADODB.Recordset
    
    StrSql = "SELECT empleado.ternro,his_estructura.tenro,his_estructura.estrnro,estr_cod.nrocod "
    StrSql = StrSql & "FROM empleado "
    StrSql = StrSql & "INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & "INNER JOIN estr_cod on estr_cod.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "WHERE empleado.ternro = " & idemp & " and his_estructura.tenro = 4 "
    StrSql = StrSql & "AND estr_cod.tcodnro = " & OcuCCSS
    StrSql = StrSql & "AND his_estructura.htethasta is null "
    OpenRecordset StrSql, rs_estructura
    
    
    If rs_estructura.EOF Then
        codidgo_asoc_emp = ""
    Else
        codidgo_asoc_emp = rs_estructura!nrocod
    End If

End Function

Public Function licencia(ByVal idemp, ByVal fecini, ByVal fecfin)
'BUSCA LICENCIA ASOCIADA POR EMPLEADO X FECHA DESDE Y HASTA
'Formato de entrada de fecha AAAAMMDD
Dim rs_Lic As New ADODB.Recordset
Dim aux As Integer


'StrSql = "SELECT emp_lic.empleado,emp_lic.elfechadesde,emp_lic.elfechahasta"
StrSql = "SELECT emp_lic.empleado "
StrSql = StrSql & ",CASE WHEN (emp_lic.elfechadesde < " & ConvFecha(fecini) & ") THEN " & ConvFecha(fecini) & " ELSE emp_lic.elfechadesde END elfechadesde "
StrSql = StrSql & ",CASE WHEN (emp_lic.elfechahasta > " & ConvFecha(fecfin) & ") THEN " & ConvFecha(fecfin) & " ELSE emp_lic.elfechahasta END elfechahasta "
StrSql = StrSql & ",emp_lic.tdnro,confrep.conftipo,confrep.confetiq,confrep.confval2,confrep.repnro "
StrSql = StrSql & "FROM emp_lic "
StrSql = StrSql & "INNER JOIN lic_estado ON lic_estado.licestnro = emp_lic.licestnro "
StrSql = StrSql & "INNER JOIN confrep ON  confrep.confval = emp_lic.tdnro "
'StrSql = StrSql & "WHERE emp_lic.elfechadesde >= '" & fecini & "' AND emp_lic.elfechahasta <= '" & fecfin & "' "
'StrSql = StrSql & "WHERE (emp_lic.elfechadesde >= '" & fecini & "' AND emp_lic.elfechahasta <= '" & fecfin & "' "
'StrSql = StrSql & "WHERE (emp_lic.elfechadesde >= " & ConvFecha(fecini) & " AND emp_lic.elfechadesde <= " & ConvFecha(fecfin) & " OR (emp_lic.elfechahasta <= " & ConvFecha(fecfin) & " AND emp_lic.elfechahasta >= " & ConvFecha(fecini) & "))" 'query nico

StrSql = StrSql & " WHERE ((emp_lic.elfechadesde <= " & ConvFecha(fecini) & " AND emp_lic.elfechahasta >= " & ConvFecha(fecfin) & ")"
StrSql = StrSql & " OR (emp_lic.elfechadesde < " & ConvFecha(fecini) & ") AND ((emp_lic.elfechahasta >= " & ConvFecha(fecini) & ") AND (emp_lic.elfechahasta <= " & ConvFecha(fecfin) & "))"
StrSql = StrSql & " OR (emp_lic.elfechadesde >= " & ConvFecha(fecini) & ") AND ((emp_lic.elfechadesde <= " & ConvFecha(fecfin) & ") AND (emp_lic.elfechahasta >= " & ConvFecha(fecfin) & "))"
StrSql = StrSql & " OR (emp_lic.elfechadesde >= " & ConvFecha(fecini) & ") AND ((emp_lic.elfechadesde <= " & ConvFecha(fecfin) & ") AND (emp_lic.elfechahasta <= " & ConvFecha(fecfin) & ")))"
'StrSql = StrSql & " WHERE emp_lic.elfechadesde >= " & ConvFecha(fecini) & " AND emp_lic.elfechadesde <= " & ConvFecha(fecfin) & " AND (emp_lic.elfechahasta <= " & ConvFecha(fecfin) & " or emp_lic.elfechahasta is null)"


'StrSql = StrSql & "OR (Month(emp_lic.elfechahasta) = " & Month(fecfin) & " AND Year(emp_lic.elfechahasta) = " & Year(fecfin) & ") OR (Month(emp_lic.elfechadesde) = " & Month(fecfin) & " AND Year(emp_lic.elfechadesde) = " & Year(fecfin) & ")) "

StrSql = StrSql & "AND confrep.repnro = 299 "
StrSql = StrSql & "AND (confrep.conftipo = 'IN' or confrep.conftipo = 'PE' ) "
StrSql = StrSql & "AND emp_lic.empleado = " & idemp


OpenRecordset StrSql, rs_Lic
Flog.writeline "Query licencia:" & StrSql
If rs_Lic.EOF Then
    licencia = ""
    Flog.writeline "No se encuentra ninguna licencia asociada"
Else
    aux = 0
    Do While Not rs_Lic.EOF
        If aux = 0 Then
            licencia = rs_Lic!conftipo & Format_StrNro(Trim(rs_Lic!confval2), 3, True, " ") & "@" & Format_Fecha(rs_Lic!elfechadesde, 1) & Format_Fecha(rs_Lic!elfechahasta, 1)
            aux = aux + 1
        Else
            licencia = licencia & "!" & rs_Lic!conftipo & Format_StrNro(Trim(rs_Lic!confval2), 3, True, " ") & "@" & Format_Fecha(rs_Lic!elfechadesde, 1) & Format_Fecha(rs_Lic!elfechahasta, 1)
        End If
        rs_Lic.MoveNext
    Loop
    Flog.writeline "licencias ok" & licencia
    'cod_estr = Format_StrNro(rs_estructura!estrcodext & rs_estructura!nrocod, 8, True, " ")
End If

End Function


Public Function pe_novedad(ByVal idemp, ByVal fecini, ByVal fecfin)
'BUSCA PERMISO X NOVEDAD
'Formato de entrada de fecha AAAAMMDD
Dim rs_nov As New ADODB.Recordset
Dim aux As Integer



'StrSql = "SELECT confrep.conftipo,confrep.confval2,gti_novedad.gnovnro,gti_novedad.gnovotoa,gti_novedad.gnovdesde, "
'StrSql = StrSql & "gti_novedad.gnovhasta , gti_novedad.gtnovnro "
'StrSql = StrSql & "FROM gti_novedad "
'StrSql = StrSql & "INNER JOIN confrep ON confrep.confval =  gti_novedad.gtnovnro "
'StrSql = StrSql & "WHERE confrep.repnro = 299 "
'StrSql = StrSql & "AND confrep.conftipo = 'PE' "
'StrSql = StrSql & "AND gti_novedad.gnovotoa = " & idemp
'StrSql = StrSql & " AND gti_novedad.gnovdesde >= '" & fecini & "' AND gti_novedad.gnovdesde <= '" & fecfin & "' "
StrSql = "SELECT  confrep.conftipo,confrep.confval2"
StrSql = StrSql & ",gti_novedad.gnovotoa"
StrSql = StrSql & ",gti_novedad.gnovdesde, gti_novedad.gnovhasta, gti_novedad.gtnovnro, gti_novedad.gnovnro "
StrSql = StrSql & ",CASE WHEN (gti_novedad.gnovdesde < " & ConvFecha(fecini) & ") THEN " & ConvFecha(fecini) & " ELSE gti_novedad.gnovdesde END gnovdesde"
StrSql = StrSql & ",CASE WHEN (gti_novedad.gnovhasta > " & ConvFecha(fecfin) & ") THEN " & ConvFecha(fecfin) & " ELSE gti_novedad.gnovhasta END gnovhasta"
StrSql = StrSql & " From gti_novedad"
StrSql = StrSql & " INNER JOIN confrep ON confrep.confval =  gti_novedad.gtnovnro"
StrSql = StrSql & " WHERE confrep.repnro = 299 AND confrep.conftipo = 'PE' AND gti_novedad.gnovotoa = " & idemp
StrSql = StrSql & " AND (gti_novedad.gnovdesde >= " & ConvFecha(fecini) & " AND gti_novedad.gnovdesde <= " & ConvFecha(fecfin)
StrSql = StrSql & " OR (gti_novedad.gnovhasta <= " & ConvFecha(fecfin) & " AND gti_novedad.gnovhasta >= " & ConvFecha(fecini) & " ))"


OpenRecordset StrSql, rs_nov
Flog.writeline "pe_novedad: " & StrSql

If rs_nov.EOF Then
    pe_novedad = ""
    'Flog.writeline "No se encuentra ninguna licencia asociada"
Else
    aux = 0
    Do While Not rs_nov.EOF
        If aux = 0 Then
            pe_novedad = rs_nov!conftipo & rs_nov!confval2 & "@" & Format_Fecha(rs_nov!gnovdesde, 1) & "@" & Format_Fecha(rs_nov!gnovhasta, 1)
            aux = aux + 1
        Else
            pe_novedad = pe_novedad & "!" & rs_nov!conftipo & rs_nov!confval2 & "@" & Format_Fecha(rs_nov!gnovdesde, 1) & "@" & Format_Fecha(rs_nov!gnovhasta, 1)
        End If
        rs_nov.MoveNext
    Loop
        
End If

End Function
Public Function TC_IN_EX(ByVal idemp, ByVal tc, ByVal fecini, ByVal fecfin, ByVal tipo)
    Dim rs_tipcam As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim jornada
   
'----TRAE EL TIPO DE JORNADA ASOCIADA AL EMPLEADO-----
    StrSql = "SELECT estr_cod.nrocod,estructura.estrcodext "
    StrSql = StrSql & " FROM estr_cod "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = estr_cod.estrnro"
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro = estr_cod.estrnro "
    StrSql = StrSql & " WHERE estr_cod.tcodnro = " & tc
    StrSql = StrSql & " AND his_estructura.tenro = 8"
    'StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecfin) & " AND his_estructura.htethasta is null "
    'StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecfin) & " AND his_estructura.htethasta is null "
    StrSql = StrSql & " AND ((his_estructura.htetdesde <= " & ConvFecha(fecfin) & " )  AND ((his_estructura.htethasta is null) or (htethasta >=" & ConvFecha(fecini) & " ) ))"
    StrSql = StrSql & " AND ternro = " & idemp
    StrSql = StrSql & " ORDER BY htetdesde desc "
    OpenRecordset StrSql, rs
    If rs.EOF Then
        jornada = ""
     Else
        If rs!estrcodext = 0 Or rs!estrcodext = "" Then
            jornada = "  " & Format_StrNro(rs!nrocod, 3, True, 0)
        Else
            jornada = Format_StrNro(rs!estrcodext, 2, True, 0) & Format_StrNro(rs!nrocod, 3, True, 0)
        End If
        
    End If
    Flog.writeline "query que obtiene la jornada: " & StrSql
    Flog.writeline "Guardo jornada para IC: " & jornada
    Flog.writeline ""
    'Flog.writeline StrSql
'--------------------------------
'Guarda inclusion o exclusion de acuerdo a un periodo
'--------------------------------
    If tipo = "I" Then
        '--ALTA
        StrSql = "SELECT fasnro,empleado,altfec,bajfec,estado FROM fases "
        StrSql = StrSql & " WHERE Empleado = " & idemp
        'StrSql = StrSql & " AND altfec >= " & ConvFecha(fecini) & " AND altfec <= " & ConvFecha(fecfin) & " AND bajfec is null "
        StrSql = StrSql & " AND altfec >= " & ConvFecha(fecini) & " AND altfec <= " & ConvFecha(fecfin)
        OpenRecordset StrSql, rs_tipcam
        If rs_tipcam.EOF Then
            TC_IN_EX = ""
        Else
            'TC_IN_EX = "IC" & jornada & "GEN" & "@" & Format_Fecha(rs_tipcam!altfec, 1)
            TC_IN_EX = "IC" & jornada & "GEN" & "@" & rs_tipcam!altfec
            'Flog.Writeline "IN_EX " & StrSql
        End If
        rs_tipcam.Close
        
        Flog.writeline "Guardo Registro IC: " & TC_IN_EX
        Flog.writeline ""
        'Flog.writeline StrSql
    End If
    If tipo = "E" Then
        '--BAJA
        StrSql = "SELECT fasnro,empleado,altfec,bajfec,estado FROM fases "
        StrSql = StrSql & " WHERE Empleado = " & idemp
        StrSql = StrSql & " AND bajfec >= " & ConvFecha(fecini) & " AND bajfec <= " & ConvFecha(fecfin)
        OpenRecordset StrSql, rs_tipcam
        If rs_tipcam.EOF Then
            TC_IN_EX = TC_IN_EX & ""
        Else
            'TC_IN_EX = "EX" & "@" & Format_Fecha(rs_tipcam!bajfec, 1)
            TC_IN_EX = "EX" & "@" & rs_tipcam!bajfec
            Flog.writeline "IN_EX " & StrSql
        End If
    rs_tipcam.Close
    End If
    
End Function
Public Function cambiojornada(ByVal idemp, ByVal tc, ByVal fecini, ByVal fecfin)
'Recupera si existe un cambio de jornada
    Dim rs As New ADODB.Recordset
    Dim confval
    
    StrSql = "SELECT confval FROM confrep WHERE repnro = 299 "
    StrSql = StrSql & " AND conftipo = 'JH'"
    OpenRecordset StrSql, rs
    If rs.EOF Then
        cambiojornada = ""
        Flog.writeline "Falta configurar el tipo de Jornada"
        Exit Function
    Else
        confval = rs!confval
    End If
    rs.Close
    
    StrSql = "SELECT estr_cod.nrocod,estructura.estrcodext,his_estructura.htetdesde  "
    StrSql = StrSql & " FROM estr_cod "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = estr_cod.estrnro"
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro = estr_cod.estrnro "
    StrSql = StrSql & " WHERE estr_cod.tcodnro = " & tc
    StrSql = StrSql & " AND his_estructura.tenro = " & confval
    'StrSql = StrSql & " AND his_estructura.htetdesde >= " & fecini & "  AND his_estructura.htethasta is null"
    StrSql = StrSql & " AND his_estructura.htetdesde >= " & ConvFecha(fecini) & " AND his_estructura.htetdesde <= " & ConvFecha(fecfin) & " AND (his_estructura.htethasta <= " & ConvFecha(fecfin) & " or his_estructura.htethasta is null)"
    
    StrSql = StrSql & " AND ternro = " & idemp
    
    OpenRecordset StrSql, rs
    If rs.EOF Then
        cambiojornada = ""
        'Flog.writeline "No figura ningún tipo de cambio en este período"
        Exit Function
    Else
        If rs!estrcodext = 0 Or rs!estrcodext = "" Then
            cambiojornada = "JO" & "  " & Format_StrNro(rs!nrocod, 3, True, 0) & "@" & rs!htetdesde
        Else
            cambiojornada = "JO" & Format_StrNro(rs!estrcodext, 2, True, 0) & Format_StrNro(rs!nrocod, 3, True, 0) & "@" & rs!htetdesde
        End If
        'cambiojornada = "JO" & Format_StrNro(rs!estrcodext, 2, True, 0) & Format_StrNro(rs!nrocod, 3, True, 0) & "@" & Format_Fecha(rs!htetdesde, 1)
        
    End If
    
    rs.Close

End Function

