Attribute VB_Name = "MdlCambiosMasivos"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "22/03/2007"
                'Se agregaron los perfiles de los ampleados a autogestion.

'Const Version = 1.02
'Const FechaVersion = "20/06/2007"
'Const Modificacion = "" ' Gustavo Ring
                        ' Se agrego chequeo de estructura por tipo de estructura validas según
                        ' las estructuras que tenga el empleado

'Const Version = 1.03
'Const FechaVersion = "17/06/2008"
'Const Modificacion = "" ' Fernando Favre
                        ' Se agrego que imprima la ultima sql ejecuta cuando indica que ocurrio un error
                        ' Si ocurre un error, no se actualiza la barra de avance y no termina.

'Const Version = 1.04
'Const FechaVersion = "16/06/2008"
'Const Modificacion = "Dependencias entre estructuras" ' Lisandro Moro
'                        ' Se agrego el ciere entre las dependencias entre estructuras.

'Const Version = "1.05"
'Const FechaVersion = "13/02/2009"
'Const Modificacion = "Encriptacion de string de conexion" 'FGZ


'Const Version = 1.06
'Const FechaVersion = "17/11/2011" 'Margiotta, Emanuel
'Const Modificacion = "Vista de Empleado"
                    'Se agregó la funcion CreaVistaEmpleadoProceso para trabajar con la vista del usuario que dispara el proceso
                    'y se cambio la funcion cargar organizacion la sql que usaba la vista empleado.
                        
'Const Version = 1.07
'Const FechaVersion = "08/08/2012" 'Deluchi Ezequiel
'Const Modificacion = "Vista de Empleado, dependencias estructuras" 'CAS-16624 - La Caja - Bug Cambios masivos
                    'Se corrgio una consulta donde se llamaba a v_empleadoproc cuando debia llamar a empleado en Select_Datos
                    'Se cambio en la funcion depende la consulta que mira si hay dependencia de estructura, el where se hacia por el tipo, ahora se hace por estructura
                        
Const Version = 1.08
Const FechaVersion = "10/08/2012" 'Deluchi Ezequiel
Const Modificacion = "Correccion condiciones de dependencias de estructuras" 'CAS-16624 - La Caja - Bug Cambios masivos
                    'Se corrigio la condicion cuando existen varias dependencias
                        
Private Type ParesConceptoParametro
    concnro As Long
    tpanro As Long
End Type

Global idUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String
Global Pares() As ParesConceptoParametro
Global Porcentaje As Long


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 18/11/2004
' Ultima Mod.: 22/11/2005 - Mariano Capriz
' Descripcion:  Se contempla el caso de ke los parametros traigan tipo null en la tabla batch_proceso
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
'Dim HuboError As Boolean


Dim rs_batch_proceso As New ADODB.Recordset

Dim PID As String
Dim bprcparam As String
Dim ArrParametros
Dim id_user As String


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
    
    
    Nombre_Arch = PathFLog & "Cambios_Masivos" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    'Abro la conexion
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
    
    On Error GoTo MError:

    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now

    
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    Flog.writeline Espacios(Tabulador * 0) & "Buscando los parametros del proceso"
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 91 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    'Borro los errores del proceso en tabla batch_logs
    Flog.writeline Espacios(Tabulador * 0) & "Borrando los datos de la tabla de logs"
    bl_borrar (NroProcesoBatch)
    
    TiempoInicialProceso = GetTickCount
    Progreso = 0
    
    If Not rs_batch_proceso.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "Parametros"
        idUser = rs_batch_proceso!idUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        
        If IsNull(rs_batch_proceso!bprcparam) Then
            bprcparam = ""
        Else
            bprcparam = rs_batch_proceso!bprcparam
        End If
        
        id_user = rs_batch_proceso!idUser
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        On Error Resume Next
        
        Call LevantarParametros(NroProcesoBatch, id_user)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    
    
    
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado',bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " DELETE FROM cam_masivo "
        StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
        objConn.Execute StrSql, , adExecuteNoRecords
        
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto',bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
    
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
    Set objconnProgreso = Nothing
    Set objConn = Nothing
    
    Exit Sub

MError:
    Flog.writeline "Error en el Main: " & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en el Main", "Error en el Main: " & Replace(Err.Description, "'", ""))
    HuboError = True
    Exit Sub
    
End Sub


Function controlStrNull(Str)
  If IsNull(Str) Then
     controlStrNull = ""
  Else
     controlStrNull = Str
  End If
End Function

Public Sub LevantarParametros(ByVal bpronro As Long, ByVal idUser As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en la tabla cam_masivo
' Autor      : Martín Ferraro
' Fecha      : 24/05/2005
' Ult. Mod   : Lisandro Moro - Se agrego el perfil a los cambios masivos
' Fecha      : 22/03/2007
' --------------------------------------------------------------------------------------------

Dim rs_cam_masivo As New ADODB.Recordset

Dim l_estcivnro         As Long
Dim l_paisnro           As Long
Dim l_nacionalnro       As Long
Dim l_empfaltagr
Dim l_empdiscap         As Long
Dim l_empvivpropia      As Long
Dim l_emptarinsalubre   As Long
Dim l_tplatenro         As Long
Dim l_medicocab         As Long
Dim l_empreporta        As Long
Dim l_sqlEstr           As String
Dim l_sqlIdioma         As String
Dim l_sqlPCuerpo        As String
Dim l_sqlDocu           As String
Dim l_sqlReincor        As String
Dim l_sqlBaja           As String
Dim l_cambiafecha       As String
Dim l_estructuras       As String
Dim l_fechaestr
Dim l_opc_abrirestr     As Long
Dim l_todos_empl        As Long
Dim l_fecha_cierra
Dim l_tipofecestr       As Integer
Dim l_abrirEstrReinc      As Integer
Dim l_borrar            As Integer
Dim l_perfnro           As Integer

On Error GoTo MError
    
    Flog.writeline Espacios(Tabulador * 0) & "Fecha"
    StrSql = "SELECT * FROM cam_masivo WHERE bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_cam_masivo
    If Not rs_cam_masivo.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "Param 01"
        l_estcivnro = rs_cam_masivo!estcivnro
        Flog.writeline Espacios(Tabulador * 0) & "Param 02"
        l_paisnro = rs_cam_masivo!paisnro
        Flog.writeline Espacios(Tabulador * 0) & "Param 03"
        l_nacionalnro = rs_cam_masivo!nacionalnro
        
        Flog.writeline Espacios(Tabulador * 0) & "Param 04"
        If IsNull(rs_cam_masivo!empfaltagr) Then
            l_empfaltagr = ""
        Else
            l_empfaltagr = rs_cam_masivo!empfaltagr
        End If
        Flog.writeline Espacios(Tabulador * 0) & "Param 05"
        l_empdiscap = rs_cam_masivo!empdiscapval
        Flog.writeline Espacios(Tabulador * 0) & "Param 06"
        l_empvivpropia = rs_cam_masivo!empvivpropiaval
        Flog.writeline Espacios(Tabulador * 0) & "Param 07"
        l_emptarinsalubre = rs_cam_masivo!emptarinsalubreval
        Flog.writeline Espacios(Tabulador * 0) & "Param 08"
        l_tplatenro = rs_cam_masivo!tplatenro
        Flog.writeline Espacios(Tabulador * 0) & "Param 09"
        l_medicocab = rs_cam_masivo!medicocab
        Flog.writeline Espacios(Tabulador * 0) & "Param 10"
        l_empreporta = rs_cam_masivo!empreporta
        Flog.writeline Espacios(Tabulador * 0) & "Param 11"
        
        l_sqlEstr = controlStrNull(rs_cam_masivo!sqlEstr)
        Flog.writeline Espacios(Tabulador * 0) & "Param 12"
        l_sqlIdioma = controlStrNull(rs_cam_masivo!sqlIdioma)
        Flog.writeline Espacios(Tabulador * 0) & "Param 13"
        l_sqlPCuerpo = controlStrNull(rs_cam_masivo!sqlPCuerpo)
        Flog.writeline Espacios(Tabulador * 0) & "Param 14"
        l_sqlDocu = controlStrNull(rs_cam_masivo!sqlDocu)
        Flog.writeline Espacios(Tabulador * 0) & "Param 15"
        l_sqlReincor = controlStrNull(rs_cam_masivo!sqlReincor)
        Flog.writeline Espacios(Tabulador * 0) & "Param 16"
        l_abrirEstrReinc = rs_cam_masivo!abrirEstrReinc
        Flog.writeline Espacios(Tabulador * 0) & "Param 17"
        l_sqlBaja = controlStrNull(rs_cam_masivo!sqlBaja)
       
        Flog.writeline Espacios(Tabulador * 0) & "Param 18"
        l_cambiafecha = controlStrNull(rs_cam_masivo!cambiafecha)
        
        Flog.writeline Espacios(Tabulador * 0) & "Param 19"
        If IsNull(rs_cam_masivo!fechaestr) Then
            l_fechaestr = ""
        Else
            l_fechaestr = rs_cam_masivo!fechaestr
        End If
        Flog.writeline Espacios(Tabulador * 0) & "Param 20"
        l_tipofecestr = rs_cam_masivo!tipofecestr
        Flog.writeline Espacios(Tabulador * 0) & "Param 21"
        l_estructuras = controlStrNull(rs_cam_masivo!listaestructuras)
        Flog.writeline Espacios(Tabulador * 0) & "Param 22"
        l_opc_abrirestr = rs_cam_masivo!abrirestr
        
        Flog.writeline Espacios(Tabulador * 0) & "Param 23"
        l_todos_empl = rs_cam_masivo!todos_empl
        
        If l_todos_empl = -1 Then
            Call CreaVistaEmpleadoProceso("V_EMPLEADO", idUser)
        End If
        
        Flog.writeline Espacios(Tabulador * 0) & "Param 24"
        If IsNull(rs_cam_masivo!fecha_cierra) Then
            l_fecha_cierra = ""
        Else
            l_fecha_cierra = rs_cam_masivo!fecha_cierra
        End If
        
        Flog.writeline Espacios(Tabulador * 0) & "Param 25"
        l_borrar = rs_cam_masivo!borrar
        
        Flog.writeline Espacios(Tabulador * 0) & "Param 26"
        If IsNull(rs_cam_masivo!perfnro) Then
            l_perfnro = 0
        Else
            l_perfnro = rs_cam_masivo!perfnro
        End If
        
    
    Else
        'error
        Call bl_insertar(bpronro, 1, "No se encontraron los parametros en la tabla cam_masivo", "No se encontraron los parametros en la tabla cam_masivo")
        Exit Sub
    End If
    
    rs_cam_masivo.Close
    Set rs_cam_masivo = Nothing
    
    l_sqlIdioma = Mid(l_sqlIdioma, 2)
    l_sqlPCuerpo = Mid(l_sqlPCuerpo, 2)
    l_sqlDocu = Mid(l_sqlDocu, 2)
    
    Call Select_Datos(l_todos_empl, bpronro, l_estcivnro, l_paisnro, l_nacionalnro, l_empdiscap, l_empvivpropia, l_emptarinsalubre, l_medicocab, l_empreporta, l_empfaltagr, l_tplatenro, l_sqlIdioma, l_sqlPCuerpo, l_sqlDocu, l_sqlBaja, l_sqlReincor, l_cambiafecha, l_fecha_cierra, l_sqlEstr, l_estructuras, l_tipofecestr, l_fechaestr, l_opc_abrirestr, l_abrirEstrReinc, l_borrar, l_perfnro)
    
    Exit Sub

MError:
    Flog.writeline "Error en LevantarParametros: " & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en LevantarParametros", "Error en LevantarParametros: " & Replace(Err.Description, "'", ""))
    HuboError = True
    Exit Sub
    
End Sub


Public Sub Select_Datos(ByVal l_todos As Long, ByVal l_bpronro As Long, ByVal l_estcivnro As Long, ByVal l_paisnro As Long, ByVal l_nacionalnro As Long, ByVal l_empdiscap As Long, ByVal l_empvivpropia As Long, ByVal l_emptarinsalubre As Long, ByVal l_medicocab As Long, ByVal l_empreporta As Long, ByVal l_empfaltagr, ByVal l_tplatenro As Long, ByVal l_sqlIdioma As String, ByVal l_sqlPCuerpo As String, ByVal l_sqlDocu As String, ByVal l_sqlBaja As String, ByVal l_sqlReincor As String, ByVal l_cambiafecha As String, ByVal l_fecha_cierra, ByVal l_sqlEstr As String, ByVal l_estructuras As String, ByVal l_tipofecestr As Integer, ByVal l_fechaestr, ByVal l_opc_abrirestr As Long, ByVal l_abrirEstrReinc As Integer, ByVal l_borrar As Integer, ByVal l_perfnro As Integer)

' --------------------------------------------------------------------------------------------
' Descripcion: seleccionas los empleados segun sea todos o sel en batch_empleado
' Autor      : Martín Ferraro
' Fecha      : 26/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim rs_empleados As New ADODB.Recordset
Dim totalEmpleados
Dim cantEmpleados
 
On Error GoTo MError
    


    If l_todos = -1 Then
        StrSql = "SELECT ternro, empleg, terape, ternom FROM v_empleadoproc"
    Else
        StrSql = " SELECT empleado.ternro, empleg, terape, ternom FROM batch_empleado "
        StrSql = StrSql & " INNER JOIN empleado ON batch_empleado.ternro = empleado.ternro "
        StrSql = StrSql & " Where bpronro = " & l_bpronro
    End If
    OpenRecordset StrSql, rs_empleados
    
    totalEmpleados = rs_empleados.RecordCount
    cantEmpleados = 0
    
    Do While Not rs_empleados.EOF
        EmpleadoSinError = True
        
        Flog.writeline Espacios(Tabulador * 0) & "Procesando Empleado: " & rs_empleados!empleg & " - " & rs_empleados!terape & " ," & rs_empleados!ternom
        
        Call Cargar_Basicas(rs_empleados!ternro, l_estcivnro, l_paisnro, l_nacionalnro, l_empdiscap, l_empvivpropia, l_emptarinsalubre, l_medicocab, l_empreporta, l_empfaltagr)
        Call Cargar_Idiomas(rs_empleados!ternro, l_sqlIdioma)
        Call cargar_docu(rs_empleados!ternro, l_sqlDocu)
        Call cargar_p_cuerpo(rs_empleados!ternro, l_sqlPCuerpo)
        Call Baja(rs_empleados!ternro, l_sqlBaja)
        Call reincorporar(rs_empleados!ternro, l_sqlReincor, l_cambiafecha, l_abrirEstrReinc)
        Call Cargar_Organizacion(rs_empleados!ternro, l_tplatenro, l_fecha_cierra, l_sqlEstr)
        Call CargarEstructuras(rs_empleados!ternro, l_estructuras, l_tipofecestr, l_fechaestr, l_borrar)
        Call CargarAbrirEstr(rs_empleados!ternro, l_opc_abrirestr)
        Call CargarPerfil(rs_empleados!ternro, l_perfnro)
        
        If EmpleadoSinError Then
            StrSql = " DELETE FROM batch_empleado "
            StrSql = StrSql & " WHERE bpronro = " & l_bpronro
            StrSql = StrSql & " AND ternro = " & rs_empleados!ternro
    
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        cantEmpleados = cantEmpleados + 1
        
        Call ActualizarProgreso(l_bpronro, ((cantEmpleados * 100) / (totalEmpleados + 1)))
        
        rs_empleados.MoveNext
    Loop
    rs_empleados.Close
    
    Exit Sub

MError:
    Flog.writeline "Error en select_datos:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en select_datos", "Error en select_datos: " & Replace(Err.Description, "'", ""))
    HuboError = True
    Exit Sub
    
    
End Sub


Public Sub Cargar_Basicas(ByVal l_ternro As Long, ByVal l_estcivnro As Long, ByVal l_paisnro As Long, ByVal l_nacionalnro As Long, ByVal l_empdiscap As Long, ByVal l_empvivpropia As Long, ByVal l_emptarinsalubre As Long, ByVal l_medicocab As Long, ByVal l_empreporta As Long, ByVal l_empfaltagr)
' --------------------------------------------------------------------------------------------
' Descripcion: Cambia estcivil, pais, nac, fec alta, discapaciadad, vivienda, tar insalub
'              medico, reporta del empleado o tercero
' Autor      : Martín Ferraro
' Fecha      : 26/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    
Dim l_s As String

On Error GoTo MError

    Flog.writeline Espacios(Tabulador * 1) & "Verificando Básicas."
    'Actualizacion del tercero
    l_s = ""
    StrSql = " update tercero set "
    If l_estcivnro <> 0 Then
        StrSql = StrSql & l_s & " estcivnro=" & l_estcivnro
        l_s = ","
    End If
    If l_paisnro <> 0 Then
        StrSql = StrSql & l_s & " paisnro=" & l_paisnro
        l_s = ","
    End If
    If l_nacionalnro <> 0 Then
        StrSql = StrSql & l_s & " nacionalnro=" & l_nacionalnro
        l_s = ","
    End If
    If l_s = "," Then
        StrSql = StrSql & " WHERE ternro=" & l_ternro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

    'Actualizacion del empleado
    l_s = ""
    StrSql = " update empleado set "
    If l_empfaltagr <> "" Then
        StrSql = StrSql & l_s & " empfaltagr=" & ConvFecha(l_empfaltagr)
        l_s = ","
    End If
    If l_empdiscap <> 1 Then
        StrSql = StrSql & l_s & " empdiscap=" & l_empdiscap
        l_s = ","
    End If
    If l_empvivpropia <> 1 Then
        StrSql = StrSql & l_s & " empvivpropia=" & l_empvivpropia
        l_s = ","
    End If
    If l_emptarinsalubre <> 1 Then
        StrSql = StrSql & l_s & " emptarinsalubre=" & l_emptarinsalubre
        l_s = ","
    End If
    If l_medicocab <> 0 Then
        StrSql = StrSql & l_s & " medicocab=" & l_medicocab
        l_s = ","
    End If
    If l_empreporta <> 0 Then
        StrSql = StrSql & l_s & " empreporta=" & l_empreporta
        l_s = ","
    End If
    If l_s = "," Then
        StrSql = StrSql & " WHERE ternro=" & l_ternro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Fin de Básicas."
    
    Exit Sub

MError:
    Flog.writeline "Error en Cargar_Basicas:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en Cargar_Basicas", "Error en Cargar_Basicas: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
    
End Sub


Public Sub Cargar_Organizacion(ByVal l_ternro As Long, ByVal l_tplatenro As Long, ByVal l_fecha_cierra, ByVal l_sqlEstr As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Cambia el modelo de organizacion del empleado
' Autor      : Martín Ferraro
' Fecha      : 26/05/2005
' Ult. Mod   :
'        Margiotta, Emanuel - Cambie en la consulta el v_empleado por empleado
' --------------------------------------------------------------------------------------------
    
Dim l_tplatenroant As Long
Dim rs_vempleados As New ADODB.Recordset
Dim Fecha
Dim I As Long

On Error GoTo MError
    
    Flog.writeline Espacios(Tabulador * 1) & "Verificando Organizacion"
    
    If l_tplatenro <> 0 Then
        l_tplatenroant = 0
        'Busco el modelo de organizacion anterior
        StrSql = " SELECT tplatenro FROM empleado "
        StrSql = StrSql & " WHERE ternro = " & l_ternro
        OpenRecordset StrSql, rs_vempleados
        If Not rs_vempleados.EOF Then
            If IsNull(rs_vempleados!tplatenro) Then
                l_tplatenroant = 0
            Else
                l_tplatenroant = rs_vempleados!tplatenro
            End If
            
            If (l_tplatenro <> l_tplatenroant) Then
                'Cambiar el modelo
                If ChequerEmpleado(l_ternro, l_sqlEstr) Then
                    'Si cambio el debo modelo debo cerrar las estr del modelo anterior anteriones
                    Call CerrarEstrucModAnt(l_ternro, l_tplatenroant, l_fecha_cierra)
                    'Carga las estr del nuevo modelo
                    Call CargarEstr(l_ternro, l_sqlEstr)
                    'cambia modelo del empleado
                    Call CambiaModelo(l_ternro, l_tplatenro)
                Else
                    'Inserto comentario en batch_logs
                    Call bl_insertar_con_ternro(l_ternro, NroProcesoBatch, 3, "Cambio de Organización no efectuado por superposición de fechas en las estructuras.", "")
                End If
                
            Else
                Flog.writeline Espacios(Tabulador * 2) & "Error en Organizacion - El empleado ya pertenece al modelo. Imposible el cambio de modelo"
                Call bl_insertar(NroProcesoBatch, 1, "Error en Organizacion", "El empleado ya pertenece al modelo. Imposible el cambio de modelo.")
            End If
        End If
        rs_vempleados.Close
        
    End If
    
    Flog.writeline Espacios(Tabulador * 1) & "Fin Organizacion"
    
    Exit Sub

MError:
    Flog.writeline "Error en Cargar_Organizacion:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en Cargar_Organizacion", "Error en Cargar_Organizacion: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub


End Sub


Public Sub CambiaModelo(ByVal l_ternro As Long, ByVal l_tplatenro As Long)

On Error GoTo MError

    StrSql = " UPDATE empleado SET tplatenro = " & l_tplatenro
    StrSql = StrSql & " Where Empleado.ternro = " & l_ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Exit Sub

MError:
    Flog.writeline "Error en CambiaModelo:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en CambiaModelo", "Error en CambiaModelo: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub


Public Sub CargarEstr(ByVal l_ternro As Long, ByVal l_sqlEstr As String)
Dim arrLista
Dim arrEstr
Dim Fecha
Dim rs_estr As New ADODB.Recordset
Dim I As Long

On Error GoTo MError

    If l_sqlEstr <> "" Then
        arrLista = Split(l_sqlEstr, ";")
        Fecha = arrLista(0)
    
        For I = 1 To UBound(arrLista)
            arrEstr = Split(arrLista(I), ",")
            If Cerrar_Estr(arrEstr(0), l_ternro, Fecha) Then
                StrSql = "insert into his_estructura(tenro,estrnro,htetdesde,ternro) values("
                StrSql = StrSql & arrEstr(0) & "," & arrEstr(1) & "," & ConvFecha(Fecha)
                StrSql = StrSql & "," & l_ternro & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                'Inserto comentario en batch_logs
                Call bl_insertar_con_ternro(l_ternro, NroProcesoBatch, 3, "Cambio de Org, no se pudo asignar las estr " & arrEstr(0) & " porque ya la posee para fechas posteriores", "")
            End If
        Next
    End If
    
    Exit Sub

MError:
    Flog.writeline "Error en CargarEstr:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en CargarEstr", "Error en CargarEstr: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub


End Sub


Public Function Cerrar_Estr(ByVal tenro As Long, ByVal ternro As Long, ByVal fech)


Dim dia As Date
Dim rs_estr As New ADODB.Recordset
        
On Error GoTo MError
        
    dia = CDate(fech)
    StrSql = "select htetdesde from his_estructura"
    StrSql = StrSql & " where tenro = " & tenro
    StrSql = StrSql & " and ternro = " & ternro
    StrSql = StrSql & " and htethasta is null"
    StrSql = StrSql & " and htetdesde <= " & ConvFecha(dia - 1)
    OpenRecordset StrSql, rs_estr
    'Corto la estructura
    If Not rs_estr.EOF Then
        StrSql = "update his_estructura "
        StrSql = StrSql & "set htethasta = " & ConvFecha(dia - 1)
        StrSql = StrSql & " where tenro = " & tenro
        StrSql = StrSql & " and ternro = " & ternro
        StrSql = StrSql & " and htethasta is null"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rs_estr.Close
    
    StrSql = "select htetdesde from his_estructura"
    StrSql = StrSql & " where tenro = " & tenro
    StrSql = StrSql & " and ternro = " & ternro
    StrSql = StrSql & " and ((htethasta >=" & ConvFecha(dia)
    StrSql = StrSql & " and htetdesde <= " & ConvFecha(dia) & ")"
    StrSql = StrSql & " or (htethasta IS NULL and htetdesde >= " & ConvFecha(dia) & "))"
    
    OpenRecordset StrSql, rs_estr

    Cerrar_Estr = rs_estr.EOF
    
    rs_estr.Close

    Exit Function

MError:
    Flog.writeline "Error en Cerrar_Estr:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en Cerrar_Estr", "Error en Cerrar_Estr: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Function

End Function


Public Function ChequerEmpleado(ByVal l_ternro As Integer, ByVal l_sqlEstr As String) As Boolean

Dim arrEstr
Dim arrSqlEstr
Dim Fecha
Dim j As Long
Dim dia As Date
Dim rs_estr As New ADODB.Recordset
Dim chequa As Boolean

On Error GoTo MError

    If l_sqlEstr <> "" Then
        arrSqlEstr = Split(l_sqlEstr, ";")
        Fecha = arrSqlEstr(0)
        
        chequa = True
        'Por cada estructura de la lista
        For j = 1 To UBound(arrSqlEstr)
            arrEstr = Split(arrSqlEstr(j), ",")
            dia = CDate(Fecha)
            StrSql = "select htetdesde from his_estructura"
            StrSql = StrSql & " where tenro = " & arrEstr(0)
            StrSql = StrSql & " and ternro = " & l_ternro
            StrSql = StrSql & " and ((htethasta >=" & ConvFecha(dia)
            StrSql = StrSql & " and htetdesde <= " & ConvFecha(dia) & ")"
            StrSql = StrSql & " or (htethasta IS NULL and htetdesde >= " & ConvFecha(dia) & "))"
            OpenRecordset StrSql, rs_estr
            
            If Not rs_estr.EOF Then
               chequa = False
            End If
            rs_estr.Close
        Next 'Por cada estructura de la lista
        
        If Not chequa Then
            Flog.writeline "Se excluye al empleado porque tiene estructuras con fechas posteriores"
        End If
        
        ChequerEmpleado = chequa
    End If
    
    Exit Function

MError:
    Flog.writeline "Error en ChequearEmpleado:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en ChequearEmpleado", "Error en ChequearEmpleado: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Function

    
End Function


Public Sub CerrarEstrucModAnt(ByVal l_ternro As Long, ByVal l_tplatenro As Long, ByVal l_fecha)
' --------------------------------------------------------------------------------------------
' Descripcion: Cierra las estructuras del tercero l_ternro a la fecha l_fecha del modelo l_tplatenro
' Autor      : Martín Ferraro
' Fecha      : 30/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim rs_estr As New ADODB.Recordset

On Error GoTo MError

    StrSql = "SELECT  tenro FROM  adptte_estr WHERE adptte_estr.tplatenro =" & l_tplatenro
    OpenRecordset StrSql, rs_estr

    Do While Not rs_estr.EOF
    
        StrSql = " UPDATE his_estructura "
        StrSql = StrSql & " SET htethasta = " & ConvFecha(l_fecha)
        StrSql = StrSql & " WHERE his_estructura.htethasta is null "
        StrSql = StrSql & " AND his_estructura.ternro = " & l_ternro
        StrSql = StrSql & " AND his_estructura.tenro  = " & rs_estr!tenro
        StrSql = StrSql & " AND htetdesde < " & ConvFecha(l_fecha)
        objConn.Execute StrSql, , adExecuteNoRecords
    
        rs_estr.MoveNext
    Loop
    rs_estr.Close

    Exit Sub

MError:
    Flog.writeline "Error en CerrarEstrucModAnt:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en CerrarEstrucModAnt", "Error en CerrarEstrucModAnt: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub

End Sub


Public Sub Cargar_Idiomas(ByVal l_ternro As Long, ByVal l_sqlIdioma As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para cargar los idiomas al tercero
' Autor      : Martín Ferraro
' Fecha      : 24/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    
Dim arrLista
Dim arrIdioma
Dim I As Long

On Error GoTo MError
    
    Flog.writeline Espacios(Tabulador * 1) & "Verificando Idiomas."
    arrLista = Split(l_sqlIdioma, ";")
    If Trim(l_sqlIdioma) <> "" Then
        For I = 0 To UBound(arrLista)
            arrIdioma = Split(arrLista(I), ",")
            'borramos el seteo anterior
            StrSql = "DELETE FROM emp_idi "
            StrSql = StrSql & "WHERE idinro=" & arrIdioma(0)
            StrSql = StrSql & " and empleado = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
    
            'seteamos con los nuevos valores
            StrSql = "INSERT INTO emp_idi (idinro,empleado,empidlee,empidhabla,empidescr)"
            StrSql = StrSql & " VALUES (" & arrIdioma(0) & "," & l_ternro & "," & arrIdioma(1) & ","
            StrSql = StrSql & arrIdioma(2) & "," & arrIdioma(3) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Next
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Fin de Idiomas."
    
    Exit Sub

MError:
    Flog.writeline "Error en Cargar_idiomas: " & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en Cargar_idiomas", "Error en Cargar_idiomas: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub


Public Sub cargar_p_cuerpo(ByVal l_ternro As Long, ByVal l_sqlPCuerpo As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para cargar las partes del cuerpo al tercero
' Autor      : Martín Ferraro
' Fecha      : 24/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    
 Dim arrLista
 Dim arrPCuerpo
 Dim I As Long
    
 On Error GoTo MError
  
    Flog.writeline Espacios(Tabulador * 1) & "Verificando Partes del Cuerpo."
    If Trim(l_sqlPCuerpo) <> "" Then
        arrLista = Split(l_sqlPCuerpo, ";")
        For I = 0 To UBound(arrLista)
            arrPCuerpo = Split(arrLista(I), ",")
            'borramos el seteo anterior
            StrSql = "DELETE FROM poscaracfi "
            StrSql = StrSql & "WHERE ppartcnro=" & arrPCuerpo(0)
            StrSql = StrSql & " and ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'seteamos con los nuevos valores
            StrSql = "INSERT INTO poscaracfi (ppartcnro,ternro,pcfvalor,pcfdesabr)"
            StrSql = StrSql & " VALUES (" & arrPCuerpo(0) & "," & l_ternro & "," & arrPCuerpo(1)
            StrSql = StrSql & ",'" & arrPCuerpo(2) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
        Next
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Fin de Partes de Cuerpo."
    
    Exit Sub

MError:
    Flog.writeline "Error en cargar_p_cuerpo:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en cargar_p_cuerpo", "Error en cargar_p_cuerpo: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub


Public Sub cargar_docu(ByVal l_ternro As Long, ByVal l_sqlDocu As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para cargar los documentos al tercero
' Autor      : Martín Ferraro
' Fecha      : 24/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    
 Dim arrLista
 Dim arrDocu
 Dim I As Long
 
 On Error GoTo MError
 
    Flog.writeline Espacios(Tabulador * 1) & "Verificando Documentos"
    If Trim(l_sqlDocu) <> "" Then
        arrLista = Split(l_sqlDocu, ";")
        For I = 0 To UBound(arrLista)
            arrDocu = Split(arrLista(I), ",")
            'borramos el seteo anterior
            StrSql = "DELETE from ter_doc "
            StrSql = StrSql & " WHERE  tidnro = " & arrDocu(0)
            StrSql = StrSql & " and ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
    
            'seteamos con los nuevos valores
            StrSql = "INSERT INTO ter_doc (tidnro,ternro,nrodoc,fecvtodoc)"
            StrSql = StrSql & " VALUES (" & arrDocu(0) & "," & l_ternro & ",'" & arrDocu(1)
            If EsNulo(arrDocu(2)) Then
                StrSql = StrSql & "',Null)"
            Else
                StrSql = StrSql & "'," & ConvFecha(arrDocu(2)) & ")"
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        Next
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Fin de Documentos."
    
    Exit Sub

MError:
    Flog.writeline "Error en cargar_docu:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en cargar_docu", "Error en cargar_docu: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub


Public Sub reincorporar(ByVal l_ternro As Long, ByVal l_sqlReincor As String, ByVal l_cambiafecha As String, ByVal l_abrirEstrReinc As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Reincorporación del tercero
' Autor      : Martín Ferraro
' Fecha      : 27/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim rs_registros As New ADODB.Recordset
Dim reincorporar As Boolean
Dim arrDatos
Dim l_arr_fases
Dim l_estr_activo As Long
Dim l_lista_fases As String
Dim ind As Long

On Error GoTo MError

    Flog.writeline Espacios(Tabulador * 1) & "Verificando Reincorporacion."
    
    'Verifico el estado del empleado
    reincorporar = False
    If l_sqlReincor <> "" Then
        StrSql = " SELECT ternro , empest FROM empleado WHERE ternro = " & l_ternro
        OpenRecordset StrSql, rs_registros
        If Not rs_registros.EOF Then
            If rs_registros!empest = 0 Then
                'REINCORPORAR
                reincorporar = True
            Else
                'NO REINCORPORAR
                reincorporar = False
                Flog.writeline Espacios(Tabulador * 2) & "No Reincorporado, ya se encuentra en estado Activo."
                Call bl_insertar_con_ternro(l_ternro, NroProcesoBatch, 3, "Reincorporación No Efectuada, ya se encuentra en estado Activo.", "")
            End If
        End If
        rs_registros.Close
    End If 'if l_sqlReincor
    
    'Inicio de Reincorporacion
    If reincorporar = True Then
        arrDatos = Split(l_sqlReincor, ",")
                
        'Cierra todas las estructuras abiertas
        StrSql = "UPDATE his_estructura SET htethasta = " & ConvFecha(DateAdd("d", -1, CDate(arrDatos(1))))
        StrSql = StrSql & " WHERE ternro = " & l_ternro & " and htethasta is null"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Situacion de Revista----------------------------------------------------------------
        StrSql = " SELECT * FROM estructura WHERE tenro = 30  AND estrcodext = 1 "
        OpenRecordset StrSql, rs_registros
        l_estr_activo = -1
        If Not rs_registros.EOF Then
            l_estr_activo = rs_registros!estrnro
        End If
        rs_registros.Close
        
        If l_estr_activo <> -1 Then
            'Crear el his_estructura
            StrSql = "INSERT INTO his_estructura "
            StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
            StrSql = StrSql & " VALUES (30, " & l_ternro & ", "
            StrSql = StrSql & l_estr_activo & ", "
            StrSql = StrSql & ConvFecha(arrDatos(1)) & ", null) "
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'Fin Situacion de Revista----------------------------------------------------------------
        
        StrSql = "SELECT fasnro from fases where empleado=" & l_ternro
        StrSql = StrSql & " and (bajfec IS NULL or bajfec >= " & ConvFecha(arrDatos(1)) & ")"
        OpenRecordset StrSql, rs_registros
        If rs_registros.EOF Then
        
            rs_registros.Close
            'Si no respeta la antiguedad entonces las fases inactivas las pongo los campos sueldo, real, vac, indem en false
            If l_cambiafecha = "on" Then
                StrSql = "SELECT fasnro FROM fases WHERE estado=0 AND empleado=" & l_ternro
                OpenRecordset StrSql, rs_registros
                
                If Not rs_registros.EOF Then
                    l_lista_fases = rs_registros!fasnro
                    rs_registros.MoveNext
                    
                    Do While Not rs_registros.EOF
                        l_lista_fases = l_lista_fases & "," & rs_registros!fasnro
                        rs_registros.MoveNext
                    Loop
                    rs_registros.Close
                    
                    l_arr_fases = Split(l_lista_fases, ",")
                    For ind = 0 To UBound(l_arr_fases)
                        StrSql = "UPDATE fases SET fases.sueldo = 0, fases.vacaciones = 0, fases.indemnizacion = 0, fases.real = 0 WHERE fasnro = " & CInt(l_arr_fases(ind))
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Next
                    
                End If
            End If 'cambiafecha = "on"
        
            StrSql = "insert into fases (altfec,estado,sueldo,vacaciones,indemnizacion,real,empleado) "
            StrSql = StrSql & "values (" & ConvFecha(arrDatos(1)) & ",-1,"
            If l_cambiafecha = "off" Then
                StrSql = StrSql & "-1,-1,-1,-1,"
            Else
                StrSql = StrSql & "-1,-1,-1,-1,"
            End If
            StrSql = StrSql & l_ternro & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            rs_registros.Close
        End If
        
        StrSql = "UPDATE empleado SET empest=-1 WHERE ternro = " & l_ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Abrir las estructuras de los empleados reincorporados
        If l_abrirEstrReinc <> 0 Then
            Call AbrirEstrReincorporacion(l_ternro, ConvFecha(arrDatos(1)))
        End If

    End If 'Reincorporar
    
    Flog.writeline Espacios(Tabulador * 1) & "Fin Reincorporación."
    Set rs_registros = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline "Error en reincorporar:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en reincorporar", "Error en reincorporar: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub


Public Sub AbrirEstrReincorporacion(l_ternro, l_fechaestructuras)

Dim rs_reg As New ADODB.Recordset

On Error GoTo MError

    StrSql = " SELECT * FROM his_estructura "
    StrSql = StrSql & " where ternro = " & l_ternro
    StrSql = StrSql & " AND  his_estructura.htethasta = "
    StrSql = StrSql & " (SELECT MAX(he2.htethasta) FROM his_estructura he2 "
    StrSql = StrSql & " WHERE he2.ternro=his_estructura.ternro "
    StrSql = StrSql & " AND he2.tenro=his_estructura.tenro "
    StrSql = StrSql & " AND NOT he2.htethasta IS NULL) "
    StrSql = StrSql & " AND  his_estructura.htetdesde = "
    StrSql = StrSql & " (SELECT DISTINCT MAX(he3.htetdesde) FROM his_estructura he3 "
    StrSql = StrSql & " WHERE he3.ternro=his_estructura.ternro "
    StrSql = StrSql & " AND he3.tenro=his_estructura.tenro "
    StrSql = StrSql & " AND NOT he3.htethasta IS NULL) "
    StrSql = StrSql & " AND  NOT EXISTS "
    StrSql = StrSql & " (SELECT * FROM his_estructura he4 "
    StrSql = StrSql & " WHERE he4.ternro=his_estructura.ternro "
    StrSql = StrSql & " AND he4.tenro=his_estructura.tenro "
    StrSql = StrSql & " AND he4.htetdesde>his_estructura.htetdesde "
    StrSql = StrSql & " ) "
    OpenRecordset StrSql, rs_reg
    Do Until rs_reg.EOF
        StrSql = "INSERT INTO his_estructura "
        StrSql = StrSql & " (ternro,htetdesde,htethasta, "
        StrSql = StrSql & "  tenro,estrnro)  "
        StrSql = StrSql & " VALUES ( "
        StrSql = StrSql & l_ternro & ","
        StrSql = StrSql & l_fechaestructuras & ","
        StrSql = StrSql & "NULL ,"
        StrSql = StrSql & rs_reg!tenro & ","
        StrSql = StrSql & rs_reg!estrnro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        rs_reg.MoveNext
    Loop
    rs_reg.Close
    Set rs_reg = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline "Error en AbrirEstrReincorporaracion:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en AbrirEstrReincorporaracion", "Error en AbrirEstrReincorporaracion: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub


Public Sub Baja(ByVal l_ternro As Long, ByVal l_sqlBaja As String)
' --------------------------------------------------------------------------------------------
' Descripcion: baja de empleados
' Autor      : Martín Ferraro
' Fecha      : 27/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim rs_reg As New ADODB.Recordset
Dim rs_reg2 As New ADODB.Recordset
Dim Baja As Boolean
Dim arrDatos
Dim l_fecha_baja
Dim l_caunro
Dim l_estrnro

On Error GoTo MError

    Flog.writeline Espacios(Tabulador * 1) & "Verificando Baja."
            
    Baja = False
    If Trim(l_sqlBaja) <> "" Then
        StrSql = " SELECT ternro, empest FROM empleado WHERE ternro = " & l_ternro
        OpenRecordset StrSql, rs_reg
        If Not rs_reg.EOF Then
            If rs_reg!empest = -1 Then
                'BAJA
                Baja = True
            Else
                'NO BAJA
                Baja = False
                Flog.writeline Espacios(Tabulador * 2) & "Baja No Efectuada, ya se encuentra en estado Inactivo."
                'Inserto comentario en batch_logs
                Call bl_insertar_con_ternro(l_ternro, NroProcesoBatch, 3, "Baja No Efectuada, ya se encuentra en estado Inactivo.", "")
            End If
        End If
        rs_reg.Close
    End If 'if l_sqlBaja
    
    If Baja = True Then
        arrDatos = Split(l_sqlBaja, ",")
        l_fecha_baja = arrDatos(1)
        l_caunro = arrDatos(2)
    
        'Bajo las fases
        StrSql = " SELECT fasnro from fases where estado=-1 and empleado=" & l_ternro
        StrSql = StrSql & " and altfec<= " & ConvFecha(l_fecha_baja)
        OpenRecordset StrSql, rs_reg
        If Not rs_reg.EOF Then
            StrSql = " update fases set estado= 0, bajfec=" & ConvFecha(l_fecha_baja)
            StrSql = StrSql & ", caunro=" & l_caunro
            StrSql = StrSql & " Where fasnro =" & rs_reg!fasnro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rs_reg.Close
    
        StrSql = " SELECT caudesvin from causa WHERE caunro = " & l_caunro
        OpenRecordset StrSql, rs_reg
        
        If Not rs_reg.EOF Then
            'Si la causa de baja indicada tiene la marca de desvinculación en true: se debe colocar empest=false sino no hacer nada
            If rs_reg!caudesvin = -1 Then
                StrSql = "UPDATE empleado SET empest=0 WHERE ternro = " & l_ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
    
            'Verificar si tiene una sit. de Revista Relacionado al caunro. Si es asi, hay que crear una  HIS_ESTRUCTURA con esa sit. de Revista
            l_estrnro = ""
            StrSql = "SELECT estrnro, caunro FROM causa_sitrev WHERE caunro = " & l_caunro
            OpenRecordset StrSql, rs_reg2
            If Not rs_reg2.EOF Then
                l_estrnro = rs_reg2!estrnro
            End If
            rs_reg2.Close
            
            'Esta relacionado a una situacion de revista
            If Trim(l_estrnro) <> "" Then
                'Cierro cualquier estructura abierta
                StrSql = "UPDATE his_estructura SET"
                StrSql = StrSql & " htethasta = " & ConvFecha(DateAdd("d", -1, CDate(l_fecha_baja)))
                StrSql = StrSql & " WHERE tenro = 30 "
                StrSql = StrSql & " AND ternro = " & l_ternro
                StrSql = StrSql & " AND htethasta IS NULL "
                objConn.Execute StrSql, , adExecuteNoRecords

                'Crear el his_estructura
                StrSql = " INSERT INTO his_estructura "
                StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
                StrSql = StrSql & " VALUES (30, " & l_ternro & ", " & l_estrnro & ", "
                StrSql = StrSql & ConvFecha(l_fecha_baja) & ", NULL) "
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
        End If
        rs_reg.Close
        
    
    End If 'baja = True
    
    Set rs_reg = Nothing
    Flog.writeline Espacios(Tabulador * 1) & "Fin de Baja."
    
    Exit Sub
    
MError:
    Flog.writeline "Error en baja:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en baja", "Error en baja: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub


Public Sub CargarEstructuras(ByVal l_ternro As Long, ByVal l_estructuras As String, ByVal l_tipofecestr As Integer, ByVal l_fechaestr, ByVal l_borrar As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Cambio de estructuras
' Autor      : Martín Ferraro
' Fecha      : 27/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim arreglo
Dim arreglo2
Dim j As Long
Dim rs_reg As New ADODB.Recordset
Dim l_fec

On Error GoTo MError

    Flog.writeline Espacios(Tabulador * 1) & "Verificando Estructuras."
    If l_estructuras <> "" Then
        
        'Borrar Historial----------------------------------------------------------
        If l_borrar = -1 Then
            arreglo = Split(l_estructuras, ",")
            j = 1
            Do While (j <= UBound(arreglo) - 1)
                arreglo2 = Split(arreglo(j), "-")
                
                StrSql = " DELETE  his_estructura "
                StrSql = StrSql & " WHERE ternro = " & l_ternro
                StrSql = StrSql & " AND tenro = " & arreglo2(1)
                objConn.Execute StrSql, , adExecuteNoRecords
                j = j + 1
                
                'Borro las estructuras dependientes, pero se fija si es una estructura abiera
                cerrarEstructurasDependientes l_ternro, CLng(arreglo2(0)), CStr(Date), 0
            Loop
        End If
        '---------------------------------------------------------------------------
        
        'miro si el emplado esta activo
        StrSql = " SELECT ternro , empest FROM empleado WHERE "
        StrSql = StrSql & " ternro = " & l_ternro
        'StrSql = StrSql & " AND empest = -1 "
        OpenRecordset StrSql, rs_reg
        
        If Not rs_reg.EOF Then
            rs_reg.Close
            
            'Busco la fecha para la estr--------------------------------------------------
            l_fec = ""
            If l_tipofecestr = 0 Then
            'Fecha de alta reconocida
                StrSql = " SELECT fases.altfec, fases.fasrecofec "
                StrSql = StrSql & " FROM fases "
                StrSql = StrSql & " Where fasrecofec = -1 "
                StrSql = StrSql & " AND empleado = " & l_ternro
                StrSql = StrSql & " order by altfec Desc "
                OpenRecordset StrSql, rs_reg
                If Not rs_reg.EOF Then
                    l_fec = rs_reg!altfec
                End If
                rs_reg.Close
            End If
            If l_tipofecestr = 1 Then
                'Fecha real mas antigua
                StrSql = " SELECT fases.altfec "
                StrSql = StrSql & " FROM fases "
                StrSql = StrSql & " Where real = -1 "
                StrSql = StrSql & " AND empleado = " & l_ternro
                StrSql = StrSql & " order by altfec asc"
                OpenRecordset StrSql, rs_reg
                If Not rs_reg.EOF Then
                    l_fec = rs_reg!altfec
                End If
                rs_reg.Close
            End If
            If l_tipofecestr = 2 Then
                'Fecha real mas nueva
                StrSql = " SELECT fases.altfec "
                StrSql = StrSql & " FROM fases "
                StrSql = StrSql & " Where real = -1 "
                StrSql = StrSql & " AND empleado = " & l_ternro
                StrSql = StrSql & " order by altfec desc"
                OpenRecordset StrSql, rs_reg
                If Not rs_reg.EOF Then
                    l_fec = rs_reg!altfec
                End If
                rs_reg.Close
            End If
            If l_tipofecestr = 3 Then
                'Fecha real mas nueva
                l_fec = l_fechaestr
            End If
            '-------------------------------------------------------------------------------------
    
            If l_fec <> "" Then
                'Abro las estructuras seleccionadas
                arreglo = Split(l_estructuras, ",")
                j = 1
                Do While (j <= UBound(arreglo) - 1)
                    arreglo2 = Split(arreglo(j), "-")
                    Call abrirEstructura(arreglo2(1), arreglo2(0), l_ternro, l_fec)
                    j = j + 1
                Loop
            Else
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro fecha para abrir estructura, no se cambian las estructuras."
                'Inserto comentario en batch_logs
                Call bl_insertar_con_ternro(l_ternro, NroProcesoBatch, 3, "Cambio de Estr No Efectuado, no se encontro fecha para abrir estructuras.", "")
                
            End If
            
        Else
            rs_reg.Close
            Flog.writeline Espacios(Tabulador * 2) & "El empleado no esta activo, no se cambian las estructuras."
            'Inserto comentario en batch_logs
            Call bl_insertar_con_ternro(l_ternro, NroProcesoBatch, 3, "Cambio de Estr No Efectuado, porque el empleado esta inactivo.", "")
            
            
        End If
        
    End If
    
    Set rs_reg = Nothing
    Flog.writeline Espacios(Tabulador * 1) & "Fin de Estructuras."
    
    Exit Sub
    
MError:
    Flog.writeline "Error en CargarEstructuras:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en CargarEstructuras", "Error en CargarEstructuras: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub

Function depende(tenro2, estrnro2, ternro, fecestr) As Boolean
' Gustavo Ring. 20/06/2007. Función que devuelve verdadero si la estructura esta disponible
' para un empleado que tiene determinadas estructuras cargadas.

Dim rs_depende As New ADODB.Recordset
Dim rs_depende_estruc As New ADODB.Recordset
Dim cumpleRestriccion As Boolean
Dim tieneUnaEstructura As Boolean
Dim tenro1 As Integer
Dim entro As Boolean
Dim salir As Boolean


' Calculo todas las restricciones del tipo de estructura y la estructura especifica
StrSql = " SELECT tenro1,estrnro1,tenro2, estrnro2 FROM estruc_depende WHERE estrnro2= " & estrnro2 'tenro2
StrSql = StrSql & " order by tenro1 "

OpenRecordset StrSql, rs_depende
cumpleRestriccion = True

tieneUnaEstructura = True

If Not rs_depende.EOF Then
    tenro1 = rs_depende("tenro1")
End If


While Not rs_depende.EOF And cumpleRestriccion

   StrSql = " SELECT * FROM his_estructura"
   StrSql = StrSql & " WHERE his_estructura.tenro  = " & rs_depende("tenro1")
   StrSql = StrSql & " AND his_estructura.estrnro  = " & rs_depende("estrnro1")
   StrSql = StrSql & " AND (his_estructura.htetdesde <=" & ConvFecha(fecestr)
   StrSql = StrSql & " AND (his_estructura.htethasta is null or his_estructura.htethasta>=" & ConvFecha(fecestr) & "))"
   StrSql = StrSql & " AND his_estructura.ternro=" & ternro
   OpenRecordset StrSql, rs_depende_estruc
     
   
    If Not (rs_depende_estruc.EOF) Then
        tieneUnaEstructura = True
    Else
         tieneUnaEstructura = False
    End If
    
    If rs_depende("tenro1") <> tenro1 Then
            cumpleRestriccion = cumpleRestriccion And tieneUnaEstructura
            tieneUnaEstructura = False
            entro = True
           tenro1 = rs_depende("tenro1")
           rs_depende.MoveNext
    Else
        If tieneUnaEstructura Then
            salir = False
            While Not rs_depende.EOF And Not salir
                salir = rs_depende("tenro1") <> tenro1
                tenro1 = rs_depende("tenro1")
                rs_depende.MoveNext
                
            Wend
        Else
           tenro1 = rs_depende("tenro1")
           rs_depende.MoveNext
        End If
    End If

'    If Not (rs_depende_estruc.EOF) Then
'            tieneUnaEstructura = True
'    Else
'            tieneUnaEstructura = False
'    End If
   
   
'   tenro1 = rs_depende("tenro1")
'   rs_depende.MoveNext
   rs_depende_estruc.Close
   
Wend
If Not entro Then
    cumpleRestriccion = cumpleRestriccion And tieneUnaEstructura
End If

rs_depende.Close

depende = cumpleRestriccion
'depende = tieneUnaEstructura
End Function

Public Sub abrirEstructura(ByVal tenro, ByVal estrnro, ByVal ternro As Long, ByVal Fecha)

Dim rs_reg As New ADODB.Recordset

On Error GoTo MError
    
 If depende(tenro, estrnro, ternro, Fecha) Then
    'busco estr mayores a la fecha o abierta y cerrada entre fecha
    StrSql = "SELECT * FROM his_estructura"
    StrSql = StrSql & " WHERE tenro = " & tenro
    StrSql = StrSql & " AND ternro = " & ternro
    StrSql = StrSql & " AND (htetdesde >= " & ConvFecha(Fecha)
    StrSql = StrSql & " OR (htetdesde < " & ConvFecha(Fecha)
    StrSql = StrSql & " AND htethasta >= " & ConvFecha(Fecha) & " )) "
    OpenRecordset StrSql, rs_reg
    If rs_reg.EOF Then
    
        rs_reg.Close
        'busco estructuras abiertas para cortarlas
        StrSql = "SELECT * FROM his_estructura"
        StrSql = StrSql & " WHERE tenro = " & tenro
        StrSql = StrSql & " AND ternro = " & ternro
        StrSql = StrSql & " AND htethasta IS NULL "
        StrSql = StrSql & " AND htetdesde <=" & ConvFecha(Fecha)
        OpenRecordset StrSql, rs_reg
        If Not rs_reg.EOF Then
            'cierro la estructura abierta
            StrSql = "UPDATE his_estructura "
            StrSql = StrSql & " SET htethasta = " & ConvFecha(DateAdd("d", -1, CDate(Fecha)))
            StrSql = StrSql & " WHERE tenro = " & tenro
            StrSql = StrSql & " AND ternro = " & ternro
            StrSql = StrSql & " AND htethasta is null"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Cierro las estructuras dependientes
            cerrarEstructurasDependientes ternro, rs_reg("estrnro"), ConvFecha(DateAdd("d", -1, CDate(Fecha))), 0
            
        End If
        rs_reg.Close
        
        'inserto la estructura
        StrSql = "INSERT INTO his_estructura (estrnro,tenro,ternro,htetdesde,htethasta) "
        StrSql = StrSql & " VALUES (" & estrnro
        StrSql = StrSql & "        ," & tenro
        StrSql = StrSql & "        ," & ternro
        StrSql = StrSql & "        ," & ConvFecha(Fecha)
        StrSql = StrSql & "        ,null)"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        rs_reg.Close
        'Error
        Flog.writeline Espacios(Tabulador * 2) & "Superposición de Fechas, no se cambia el tipo estructura " & tenro & "."
        Call bl_insertar_con_ternro(ternro, NroProcesoBatch, 3, "Superposición de Fechas, no se cambia el tipo estructura " & tenro & ".", "")
    End If
 Else
        Flog.writeline Espacios(Tabulador * 2) & "La estructura:" & estrnro & " no esta disponible para el empleado:" & ternro
        Call bl_insertar_con_ternro(ternro, NroProcesoBatch, 3, "La estructura que se queria asignar no es válida para el empleado.", "")
 End If
 
 Set rs_reg = Nothing
 Exit Sub
    
MError:
    Flog.writeline "Error en AbrirEstructura:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en AbrirEstructura", "Error en AbrirEstructura: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub


Public Sub CargarAbrirEstr(ByVal l_ternro As Long, ByVal l_opc_abrirestr As Long)

Dim l_htetdesde
Dim l_htethasta
Dim rs_reg As New ADODB.Recordset

On Error GoTo MError

    If l_opc_abrirestr = 1 Then

        StrSql = "select tipoestructura.tedabr, estructura.estrdabr, "
        StrSql = StrSql & " his_estructura.tenro, his_estructura.estrnro, "
        StrSql = StrSql & " his_estructura.htetdesde, his_estructura.htethasta "
        StrSql = StrSql & " from his_estructura"
        StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " where ternro = " & l_ternro
        StrSql = StrSql & " AND  his_estructura.htethasta = "
        StrSql = StrSql & " (SELECT DISTINCT MAX(he2.htethasta) FROM his_estructura he2 "
        StrSql = StrSql & " WHERE he2.ternro=his_estructura.ternro "
        StrSql = StrSql & " AND he2.tenro=his_estructura.tenro "
        StrSql = StrSql & " AND NOT he2.htethasta IS NULL) "
        StrSql = StrSql & " AND  his_estructura.htetdesde = "
        StrSql = StrSql & " (SELECT DISTINCT MAX(he3.htetdesde) FROM his_estructura he3 "
        StrSql = StrSql & " WHERE he3.ternro=his_estructura.ternro "
        StrSql = StrSql & " AND he3.tenro=his_estructura.tenro "
        StrSql = StrSql & " AND NOT he3.htethasta IS NULL) "
        StrSql = StrSql & " AND  NOT EXISTS "
        StrSql = StrSql & " (SELECT * FROM his_estructura he4 "
        StrSql = StrSql & " WHERE he4.ternro=his_estructura.ternro "
        StrSql = StrSql & " AND he4.tenro=his_estructura.tenro "
        StrSql = StrSql & " AND he4.htetdesde>his_estructura.htetdesde "
        StrSql = StrSql & " ) "
        OpenRecordset StrSql, rs_reg
    
        Do While Not rs_reg.EOF
    
            If Trim(rs_reg!htetdesde) <> "" Then
                l_htetdesde = ConvFecha(rs_reg!htetdesde)
            Else
                l_htetdesde = "null"
            End If
    
            If Trim(rs_reg!htethasta) <> "" Then
                l_htethasta = ConvFecha(rs_reg!htethasta)
            Else
                l_htethasta = "null"
            End If
    
            StrSql = "UPDATE his_estructura SET htethasta = NULL "
            StrSql = StrSql & " WHERE ternro  = " & l_ternro
            StrSql = StrSql & " AND   tenro   = " & rs_reg!tenro
            StrSql = StrSql & " AND   estrnro = " & rs_reg!estrnro
            If IsNull(l_htetdesde) Or l_htetdesde = "null" Then
                StrSql = StrSql & " AND   htetdesde IS NULL "
            Else
                StrSql = StrSql & " AND   htetdesde = " & l_htetdesde
            End If
            If IsNull(l_htethasta) Or l_htethasta = "null" Then
                StrSql = StrSql & " AND   htethasta IS NULL "
            Else
                StrSql = StrSql & " AND   htethasta = " & l_htethasta
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
    
            rs_reg.MoveNext
        Loop
        rs_reg.Close
    
    End If
    Exit Sub
    
MError:
    Flog.writeline "Error en CargarAbrirEstr:" & Err.Description
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    Call bl_insertar(NroProcesoBatch, 1, "Error en CargarAbrirEstr", "Error en CargarAbrirEstr: " & Replace(Err.Description, "'", ""))
    HuboError = True
    EmpleadoSinError = False
    Exit Sub
    
End Sub

Public Sub CargarPerfil(ternro, perfnro)
' --------------------------------------------------------------------------------------------
' Descripcion: Actualizo el Perfil del empleado para Autogestion
' Autor      : Lisandro Moro
' Fecha      : 22/03/2007
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    On Error GoTo MError
        If perfnro <> 0 Then
            If perfnro = -1 Then
                StrSql = " UPDATE empleado SET perfnro = NULL "
            Else
                StrSql = " UPDATE empleado SET perfnro = " & perfnro
            End If
            StrSql = StrSql & " WHERE empleado.ternro = " & ternro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        Exit Sub
    
MError:
        Flog.writeline "Error en CargarPerfil:" & Err.Description
        Flog.writeline "Ultima SQL ejecutada: " & StrSql
        Call bl_insertar(NroProcesoBatch, 1, "Error en CargaPerfil", "Error en CargaPerfil: " & Replace(Err.Description, "'", ""))
        HuboError = True
        EmpleadoSinError = False
        Exit Sub
    
End Sub
Public Sub ActualizarProgreso(ByVal NroProceso As Long, ByVal Progreso As Single)
' --------------------------------------------------------------------------------------------
' Descripcion: Actualizo el progreso del Proceso
' Autor      : FGZ
' Fecha      : 19/11/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
End Sub

Private Sub cerrarEstructurasDependientes(ternro As Long, estrnro As Long, Fecha As String, Nivel As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Cierro las estructuras que dependan de la estructura que se este cerrando o borrando.
' Autor      : Lisandro Moro
' Fecha      : 26/06/2008
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

    Dim l_rs As New ADODB.Recordset

    StrSql = " SELECT te1.tedabr , e1.estrdabr , te2.tedabr , e2.estrdabr e2dabr, estrnro2 "
    StrSql = StrSql & " FROM estruc_depende "
    StrSql = StrSql & " INNER JOIN estructura e1 ON estruc_depende.estrnro1 = e1.estrnro "
    StrSql = StrSql & " INNER JOIN tipoestructura te1 ON estruc_depende.tenro1 = te1.tenro "
    StrSql = StrSql & " INNER JOIN estructura e2 ON estruc_depende.estrnro2 = e2.estrnro "
    StrSql = StrSql & " INNER JOIN tipoestructura te2 ON estruc_depende.tenro2 = te2.tenro "
    StrSql = StrSql & " INNER JOIN his_estructura h ON h.estrnro = estrnro2 "
    StrSql = StrSql & " WHERE ternro = " & ternro
    StrSql = StrSql & " AND estrnro1 = " & estrnro
    StrSql = StrSql & " AND h.htethasta is null "
    StrSql = StrSql & " ORDER BY te1.tedabr , e1.estrdabr , te2.tedabr , e2.estrdabr "

    OpenRecordset StrSql, l_rs
    If Not l_rs.EOF Then
        Do While Not l_rs.EOF
            
            '-------- CIERRO ------
            StrSql = "UPDATE his_estructura "
            StrSql = StrSql & "SET htethasta = " & Fecha
            StrSql = StrSql & " , hismotivo ='Cierre por dependencias proc. masivos. nivel " & Nivel & "'"
            StrSql = StrSql & " WHERE estrnro = " & CStr(CLng(l_rs(4)))
            StrSql = StrSql & " and ternro = " & ternro
            StrSql = StrSql & " and htethasta is null"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            '-----------------------------
        
            Flog.writeline Espacios(Tabulador * 3) & "Se cierra la estructura:" & l_rs("e2dabr") & "(" & estrnro & ")" & " para el empleado " & ternro & ", en proc. masivos, nivel " & Nivel
            Call bl_insertar_con_ternro(ternro, NroProcesoBatch, 3, "Dependencias entre etructuras.", "Se cierra la estructura:" & l_rs("e2dabr") & "(" & estrnro & ")" & " para el empleado " & ternro & ", en proc. masivos, nivel " & Nivel)
            
            cerrarEstructurasDependientes ternro, l_rs("estrnro2"), Fecha, Nivel + 1
            l_rs.MoveNext

        Loop
    Else
    
    End If
    l_rs.Close

End Sub
