Attribute VB_Name = "MdlCatalogarPostulantes"
Option Explicit

'Version: 1.01
'
Const Version = 1
Const FechaVersion = "22/07/2008"


Global IdUser As String
Global Fecha As Date
Global hora As String



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
    
    TiempoInicialProceso = GetTickCount
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    Nombre_Arch = PathFLog & "Postu_catalogo" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "Version                  : " & Version
    Flog.writeline Espacios(Tabulador * 0) & "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline Espacios(Tabulador * 0) & "PID                      : " & PID
    Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------"
    Flog.writeline
    
    'Abro la conexion
    On Error Resume Next
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
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 225 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParametros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
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
    GoTo Fin:
End Sub

Public Sub Generacion()
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la busqueda de postulantes
' Autor      :
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim cadenaSQL As String
Dim StrSql2 As String
Dim seguir As Boolean
Dim busempact As Boolean
Dim busempina As Boolean
Dim buspos As Boolean
Dim caracsql As String
Dim reqbusnro As Long
Dim progresoInc As Single

Dim busestposnro As String

'Registros
Dim rs_pos_busqueda As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_Pos_Terreqbus As New ADODB.Recordset
Dim rs_eliminar As New ADODB.Recordset


MyBeginTrans

    
'------------------------------------------------------------------------------------------
'Borrar la tabla de catalogo
 Flog.writeline Espacios(Tabulador * 1) & "Borrando Catalogo..."
 
 StrSql = "TRUNCATE TABLE pos_catalogo "
 objConn.Execute StrSql, , adExecuteNoRecords
 
 Flog.writeline Espacios(Tabulador * 1) & "Catalogo Elminado"
 
' Postulantes UNION Empleados ver estados del postulante y del empleado?
 StrSql = "SELECT tercero.ternro "
 StrSql = StrSql & "FROM tercero "
 StrSql = StrSql & "INNER JOIN pos_postulante ON tercero.ternro = pos_postulante.ternro " 'and tercero.ternro = 541"
 StrSql = StrSql & "INNER JOIN pos_estpos ON pos_estpos.estposnro = pos_postulante.estposnro " 'AND pos_estpos.estposnro IN (" & busestposnro & ")"
 StrSql = StrSql & " UNION "
 StrSql = StrSql & "SELECT tercero.ternro "
 StrSql = StrSql & "FROM tercero "
 StrSql = StrSql & "INNER JOIN empleado ON tercero.ternro = empleado.ternro " 'and empleado.ternro = 541"
 'If busempact And Not busempina Then
 '                StrSql = StrSql & "AND empest = -1 "
 '               ElseIf Not busempact And busempina Then
 '                   StrSql = StrSql & "AND empest = 0 "
 '               End If
 '           End If
 OpenRecordset StrSql, rs
            
            ' Cuento la cantidad de empleados y postulantes que se agregaran
            If Not rs.EOF Then
                CEmpleadosAProc = rs.RecordCount
            Else
                Flog.writeline Espacios(Tabulador * 1) & "No hay nada que procesar."
                CEmpleadosAProc = 1
            End If
            ' Defino un número para incrementar
            IncPorc = 95 / CEmpleadosAProc
            progresoInc = 0
            Progreso = 5
            Do Until rs.EOF
                
                'nombre y apellido
                cadenaSQL = "SELECT ternro, terape, ternom FROM tercero "
                cadenaSQL = cadenaSQL & "WHERE tercero.ternro =" & rs("ternro") & " "
                OpenRecordset cadenaSQL, rs_pos_busqueda
                If Not rs_pos_busqueda.EOF Then
                    
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("terape") & " " & rs_pos_busqueda("ternom") & "','Nombre y Apellido'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Nombre y Apellido: " & rs("ternro")
                End If
                If rs.State = adStateOpen Then rs_pos_busqueda.Close
            
                ' referencias ,possoc,hobby,posestpla
                cadenaSQL = "SELECT ternro, posref,possoc,poshobbdep,posestpla FROM pos_postulante "
                cadenaSQL = cadenaSQL & "WHERE pos_postulante.ternro =" & rs("ternro") & " "
                OpenRecordset cadenaSQL, rs_pos_busqueda
                If Not rs_pos_busqueda.EOF Then
                    If Not EsNulo(rs_pos_busqueda("posref")) Or Not EsNulo(rs_pos_busqueda("poshobbdep")) Then
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("posref") & " " & rs_pos_busqueda("poshobbdep") & "','Referencias-Hobby'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                     Flog.writeline Espacios(Tabulador * 1) & "Inserto Referencias-Hobby: " & rs("ternro")
                     End If
                End If
                If rs.State = adStateOpen Then rs_pos_busqueda.Close
            
                
                ' estcivil
                cadenaSQL = "SELECT ternro, estcivdesabr FROM tercero "
                cadenaSQL = cadenaSQL & "INNER JOIN estcivil ON estcivil.estcivnro=tercero.estcivnro "
                cadenaSQL = cadenaSQL & "WHERE tercero.ternro =" & rs("ternro") & " "
                OpenRecordset cadenaSQL, rs_pos_busqueda
                If Not rs_pos_busqueda.EOF Then
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("estcivdesabr") & "','EstadoCivil'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto EstadoCivil: " & rs("ternro")
          
                End If
                If rs.State = adStateOpen Then rs_pos_busqueda.Close


                'nacionalidad
                cadenaSQL = "SELECT ternro, nacionaldes FROM tercero "
                cadenaSQL = cadenaSQL & "INNER JOIN nacionalidad ON nacionalidad.nacionalnro=tercero.nacionalnro "
                cadenaSQL = cadenaSQL & "WHERE tercero.ternro =" & rs("ternro")
                OpenRecordset cadenaSQL, rs_pos_busqueda
                If Not rs_pos_busqueda.EOF Then
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("nacionaldes") & "','Nacionalidad'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto Nacionalidad: " & rs("ternro")
          
                End If
                If rs.State = adStateOpen Then rs_pos_busqueda.Close

                ' procedencia
                cadenaSQL = "SELECT ternro,prodesabr FROM v_postulante "
                cadenaSQL = cadenaSQL & "INNER JOIN pos_procedencia ON pos_procedencia.pronro = v_postulante.procnro "
                cadenaSQL = cadenaSQL & "WHERE v_postulante.ternro =" & rs("ternro")
                OpenRecordset cadenaSQL, rs_pos_busqueda
                If Not rs_pos_busqueda.EOF Then
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("prodesabr") & "','Procedencia'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto Procedencia: " & rs("ternro")
          
                End If
                If rs.State = adStateOpen Then rs_pos_busqueda.Close


                'calle domicilio
                cadenaSQL = "SELECT ternro, detdom.calle, detdom.nro, codigopostal, locdesc, provdesc,paisdesc FROM cabdom "
                cadenaSQL = cadenaSQL & "INNER JOIN detdom ON cabdom.domnro=detdom.domnro "
                cadenaSQL = cadenaSQL & "LEFT JOIN localidad ON localidad.locnro=detdom.locnro "
                cadenaSQL = cadenaSQL & "LEFT JOIN provincia ON provincia.provnro=detdom.provnro "
                cadenaSQL = cadenaSQL & "LEFT JOIN partido ON partido.partnro=detdom.partnro "
                cadenaSQL = cadenaSQL & "LEFT JOIN pais ON pais.paisnro=detdom.paisnro "
                cadenaSQL = cadenaSQL & "WHERE cabdom.ternro =" & rs("ternro")
                OpenRecordset cadenaSQL, rs_pos_busqueda
                Do Until rs_pos_busqueda.EOF
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("calle") & " " & rs_pos_busqueda("nro") & " " & rs_pos_busqueda("locdesc") & " " & rs_pos_busqueda("provdesc") & " " & rs_pos_busqueda("paisdesc") & "','Domicilio'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto Domicilio: " & rs("ternro")
          
                rs_pos_busqueda.MoveNext
                Loop
                If rs.State = adStateOpen Then rs_pos_busqueda.Close


                'carrera
                cadenaSQL = "SELECT ternro,carredudesabr,titdesabr,nivdesc FROM cap_estformal "
                cadenaSQL = cadenaSQL & "INNER JOIN cap_carr_edu ON cap_carr_edu.carredunro =cap_estformal.carredunro "
                cadenaSQL = cadenaSQL & "LEFT JOIN titulo ON titulo.titnro =cap_estformal.titnro "
                cadenaSQL = cadenaSQL & "LEFT JOIN nivest ON nivest.nivnro =cap_estformal.nivnro "
                cadenaSQL = cadenaSQL & "WHERE cap_estformal.ternro=" & rs("ternro")
                OpenRecordset cadenaSQL, rs_pos_busqueda
                Do Until rs_pos_busqueda.EOF
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("carredudesabr") & " " & rs_pos_busqueda("titdesabr") & " " & rs_pos_busqueda("nivdesc") & "','EstudioFormal'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto EstudioFormal: " & rs("ternro")
          
                rs_pos_busqueda.MoveNext
                Loop
                If rs.State = adStateOpen Then rs_pos_busqueda.Close



'                 habilidades
                cadenaSQL = "SELECT ternro, habdesabr FROM hab_ter "
                cadenaSQL = cadenaSQL & "INNER JOIN habilidad ON habilidad.habnro=hab_ter.habnro "
                cadenaSQL = cadenaSQL & "WHERE hab_ter.ternro =" & rs("ternro") & " "
                OpenRecordset cadenaSQL, rs_pos_busqueda
                Do Until rs_pos_busqueda.EOF
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("habdesabr") & "','Habilidades'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto Habilidades: " & rs("ternro")
          
                rs_pos_busqueda.MoveNext
                Loop
                If rs.State = adStateOpen Then rs_pos_busqueda.Close



'                especializacion
                cadenaSQL = "SELECT ternro, eltoana.eltanadesabr,espnivel.espnivdesabr  FROM especemp "
                cadenaSQL = cadenaSQL & "INNER JOIN eltoana ON eltoana.eltananro = especemp.eltananro "
                cadenaSQL = cadenaSQL & "INNER JOIN espnivel ON espnivel.espnivnro = especemp.espnivnro "
                cadenaSQL = cadenaSQL & "WHERE especemp.ternro =" & rs("ternro")
                OpenRecordset cadenaSQL, rs_pos_busqueda
                Do Until rs_pos_busqueda.EOF
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("eltanadesabr") & " " & rs_pos_busqueda("espnivdesabr") & "','Especialidades'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto Especilizacion: " & rs("ternro")
          
                rs_pos_busqueda.MoveNext
                Loop
                If rs.State = adStateOpen Then rs_pos_busqueda.Close

'                 cargo
                cadenaSQL = "SELECT empleado,cardesabr FROM empant "
                cadenaSQL = cadenaSQL & "INNER JOIN cargo ON cargo.carnro = empant.carnro "
                cadenaSQL = cadenaSQL & "WHERE empant.empleado =" & rs("ternro")
                OpenRecordset cadenaSQL, rs_pos_busqueda
                Do Until rs_pos_busqueda.EOF
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("empleado") & ",'" & rs_pos_busqueda("cardesabr") & "','Cargo'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto Cargo: " & rs("ternro")
          
                rs_pos_busqueda.MoveNext
                Loop
                If rs.State = adStateOpen Then rs_pos_busqueda.Close

'                 idiomas
                cadenaSQL = "SELECT empleado, ididesc, lee.idnivdesabr Nivel1, habla.idnivdesabr Nivel2, escr.idnivdesabr Nivel3 FROM emp_idi "
                cadenaSQL = cadenaSQL & " INNER JOIN idioma ON idioma.idinro = emp_idi.idinro "
                cadenaSQL = cadenaSQL & " INNER JOIN idinivel lee ON lee.idnivnro = emp_idi.empidlee "
                cadenaSQL = cadenaSQL & " INNER JOIN idinivel habla ON habla.idnivnro = emp_idi.empidhabla "
                cadenaSQL = cadenaSQL & " INNER JOIN idinivel escr ON escr.idnivnro = emp_idi.empidescr "
                cadenaSQL = cadenaSQL & " WHERE emp_idi.empleado=" & rs("ternro")
                OpenRecordset cadenaSQL, rs_pos_busqueda
                Do Until rs_pos_busqueda.EOF
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("empleado") & ",'" & rs_pos_busqueda("ididesc") & " Lee: " & rs_pos_busqueda("Nivel1") & " Habla: " & rs_pos_busqueda("Nivel2") & " Escribe: " & rs_pos_busqueda("Nivel3") & "','Idioma'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto Idiomas: " & rs("ternro")
          
                rs_pos_busqueda.MoveNext
                Loop
                If rs.State = adStateOpen Then rs_pos_busqueda.Close

'                 estructura
                cadenaSQL = " SELECT estruc_aplica.ternro, estructura.estrdabr FROM estruc_aplica "
                cadenaSQL = cadenaSQL & " INNER JOIN estructura ON estructura.estrnro = estruc_aplica.estrnro "
                cadenaSQL = cadenaSQL & "WHERE estruc_aplica.ternro =" & rs("ternro")
                OpenRecordset cadenaSQL, rs_pos_busqueda
                Do Until rs_pos_busqueda.EOF
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("estrdabr") & "','Estructura Aplicada'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto Estructura: " & rs("ternro")
          
                rs_pos_busqueda.MoveNext
                Loop
                If rs.State = adStateOpen Then rs_pos_busqueda.Close


'                 infgral
                cadenaSQL = "SELECT ternro,infdesabr FROM pos_infter "
                cadenaSQL = cadenaSQL & "INNER JOIN pos_infgral ON pos_infgral.infnro=pos_infter.infnro "
                cadenaSQL = cadenaSQL & "WHERE pos_infter.ternro =" & rs("ternro")
                OpenRecordset cadenaSQL, rs_pos_busqueda
                Do Until rs_pos_busqueda.EOF
                        StrSql = "INSERT INTO pos_catalogo "
                        StrSql = StrSql & "(ternro, caracteristica, tipo_caract, fec_actualizacion) "
                        StrSql = StrSql & "VALUES (" & rs_pos_busqueda("ternro") & ",'" & rs_pos_busqueda("infdesabr") & "','Informacion Gral'," & ConvFecha(Date) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto Informacion Gral: " & rs("ternro")
                 rs_pos_busqueda.MoveNext
                Loop
                If rs.State = adStateOpen Then rs_pos_busqueda.Close


                
                
                rs.MoveNext
                
                
                Flog.writeline Espacios(Tabulador * 1) & "Actualizo el progreso"
                Progreso = Progreso + IncPorc
                progresoInc = progresoInc + IncPorc
                If progresoInc >= 1 Then
                    TiempoAcumulado = GetTickCount
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                             " WHERE bpronro = " & NroProcesoBatch
                    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                    progresoInc = 0
                End If
            Loop
            If rs.State = adStateOpen Then rs.Close
         
    

MyCommitTrans

'Creo el catalogo de search
StrSql = "EXEC sp_fulltext_catalog 'Catalogo_Postulante','start_full' "
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline Espacios(Tabulador * 1) & "Creo el catalogo sp_fulltext_catalog "

If rs_pos_busqueda.State = adStateOpen Then rs_pos_busqueda.Close
If rs.State = adStateOpen Then rs.Close
If rs_Pos_Terreqbus.State = adStateOpen Then rs_Pos_Terreqbus.Close

Set rs_pos_busqueda = Nothing
Set rs = Nothing
Set rs_Pos_Terreqbus = Nothing

End Sub

Public Sub LevantarParametros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : Fernando Favre
' Fecha      : 17/05/2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim busnro As Long
Dim formal As Boolean
Dim reqpernro As Long
Dim agregar As Boolean

'Orden de los parametros
'busnro
'formal
'reqpernro
'agregar

Separador = "@"
' Levanto cada parametro por separado
'If Not IsNull(parametros) Then
'    If Len(parametros) >= 1 Then
'        pos1 = 1
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        busnro = Mid(parametros, pos1, pos2 - pos1 + 1)
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        formal = Mid(parametros, pos1, pos2 - pos1 + 1)
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        reqpernro = Mid(parametros, pos1, pos2 - pos1 + 1)
'
'        pos1 = pos2 + 2
'        agregar = Mid(parametros, pos1)
'
'    End If
'End If
Call Generacion
End Sub

