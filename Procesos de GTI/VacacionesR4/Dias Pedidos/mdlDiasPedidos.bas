Attribute VB_Name = "mdlDiasPedidos"
Option Explicit
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Global Ternro As Long
Global NroProceso As Long
Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Integer
Dim IncPorc As Single
Dim Progreso As Single

Global fec_proc As Integer ' 1 - Política Primer Reg.
                           ' 2 - Política Reg. del Turno
                           ' 3 - Política Ultima Reg.
Global Usa_Conv As Boolean

Dim objBTurno As New BuscarTurno
Dim objBDia As New BuscarDia
Dim objFeriado As New Feriado
Dim objFechasHoras As New FechasHoras

Global diatipo As Byte
Global ok As Boolean
Global fecha_desde As Date
Global fecha_hasta As Date
Global Tdias As Integer
Global Thoras As Integer
Global Tmin As Integer
Global Cod_justificacion1 As Long
Global Cod_justificacion2 As Long

Global Existe_Reg As Boolean

Global tiene_turno As Boolean
Global Nro_Turno As Long
Global Tipo_Turno As Integer

Global Tiene_Justif As Boolean
Global nro_justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global nro_grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global Trabaja     As Boolean ' Indica si trabaja para ese dia
Global Orden_Dia As Integer
Global Nro_Dia As Integer
Global Nro_Subturno As Integer
Global Dia_Libre As Boolean
Global dias_trabajados As Integer
Global Dias_laborables As Integer

Global aux_Tipohora As Integer
Global aux_TipoDia As Integer

Global E1 As String
Global E2 As String
Global E3 As String
Global S1 As String
Global S2 As String
Global S3 As String
Global FE1 As Date
Global FE2 As Date
Global FE3 As Date
Global FS1 As Date
Global FS2 As Date
Global FS3 As Date

Global fv1 As Date
Global fv2 As Date
Global fv3 As Date
Global fv4 As Date
Global fv5 As Date
Global fv6 As Date
Global fv7 As Date

Global v1 As String
Global v2 As String
Global v3 As String
Global v4 As String
Global v5 As String
Global v6 As String
Global v7 As String

Global Cant_emb As Integer
Global toltemp As String
Global toldto As String
Global acumula As Boolean
Global acumula_dto As Boolean
Global acumula_temp As Boolean
Global convenio As Long

Global tdias_oblig As Single
Global Tipo_Hora As Integer
Global HuboErrores As Boolean
Global SinError As Boolean
Global Todos_Posibles As Boolean
'---------------------------
'---------------------------
Global diascoract As Integer
Global DiasTom As Integer
Global diascorant As Integer
Global diasdebe As Integer
Global diastot As Integer
Global diasyaped As Integer
Global diaspend As Integer

Global nroTipvac As Long
Global hasta As Date
Global totferiados As Integer
Global tothabiles As Integer
Global totNohabiles As Integer
Global Aux_Fecha_Desde As Date
Global Aux_Fecha_Hasta As Date
Global Vac_Fecha_Desde As Date
Global Aux_Cant_dias As Integer
Global Vac_Cant_dias As Integer

Global objModVac As New ADODB.Recordset

'---------------------------
'---------------------------
Public Sub Main()

Dim fecha As Date
'Dim Ternro As Long
'Dim fecha_desde As Date
'Dim fecha_hasta As Date
Dim NroVac As Long
Dim Reproceso As Boolean
Dim parametros As String
Dim cantdias As Integer
Dim Columna As Integer
Dim Mensaje As String
Dim Genera As Boolean
Dim NroTPV As String

Dim pos1 As Integer
Dim pos2 As Integer

Dim objReg As New ADODB.Recordset

Dim strcmdLine As String
Dim objconnMain As New ADODB.Connection
Dim Archivo As String

Dim rs As New ADODB.Recordset
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset

Dim rs_tipovacac As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset

Dim rsDias As New ADODB.Recordset
Dim rsVac As New ADODB.Recordset
Dim rs_Periodos_Vac As New ADODB.Recordset
Dim l_TienePolAlcance As Boolean
'---------------------------------
'Dim diascoract As Integer
'Dim DiasTom As Integer
'Dim diascorant As Integer
'Dim diasdebe As Integer
'Dim diastot As Integer
'Dim diasyaped As Integer
'Dim diaspend As Integer
'
'Dim nroTipvac As Long
'Dim hasta As Date
'Dim totferiados As Integer
'Dim tothabiles As Integer
'Dim totNohabiles As Integer
'Dim Aux_Fecha_Desde As Date
'Dim Vac_Fecha_Desde As Date
'Dim Aux_Cant_dias As Integer
'Dim Vac_Cant_dias As Integer


'EAM
Dim dias_tranf_PAct As Integer
Dim estadoPeriodo As Integer

Dim PID As String
Dim ArrParametros


'NG
Dim usuario As String
Dim Texto As String
Dim modeloPais
Dim VersionPais
Dim Continua As Boolean

Dim vacModelo

    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
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
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If

    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    ' Creo el archivo de texto del desglose
    Archivo = PathFLog & "Vac_DiasPedidos" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas con la conexión "
        Exit Sub
    End If
    
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas con la conexión "
        Exit Sub
    End If
    On Error GoTo 0

    'Activo el manejador de errores
    On Error GoTo ce
    
    Dim arrPar
    
    '*******************************************************************************************************
    '--------------- VALIDO MODELOS SEGUN POLITICA 1515 | PUEDE TENER ALCANCE POR ESTRUCTURAS --------------
    '*******************************************************************************************************
    '____________________________________________________________________
    'VALIDO QUE LA POLITICA 1515 ESTE ACTIVA Y CONFIGURADA
    Version_Valida = True 'ValidaModeloyVersiones(Version, 11)
   'Version_Valida = ValidarV(Version, 11, TipoBD)
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Final
    End If
    'FGZ - 05/08/2009 --------- Control de versiones ------
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
    OpenRecordset StrSql, rs_Batch_Proceso
       
    If rs_Batch_Proceso.EOF Then Exit Sub
    parametros = rs_Batch_Proceso!bprcparam
    
    
    '____________________________________________________________
    'NG - VALIDA QUE ESTE ACTIVO LA TRADUCCION A MULTI IDIOMA
    usuario = rs_Batch_Proceso!iduser
    Call Valida_MultiIdiomaActivo(usuario)
    '------------------------------------------------------------
    '------------------------------------------------------------
    Flog.writeline "parametros: " & parametros
    If Not IsNull(parametros) Then
        If Len(parametros) >= 1 Then
            
            arrPar = Split(parametros, ".")
            'formato  parametros : reproceso.desde.hasta.autorizar.proximafirma.modelo
            If UBound(arrPar) < 5 Then
                Flog.writeline "ERROR, faltan parametros."
            Else
                Reproceso = arrPar(0)
                fecha_desde = arrPar(1)
                fecha_hasta = arrPar(2)
                Aux_Fecha_Desde = CDate(arrPar(3))
                Vac_Fecha_Desde = Aux_Fecha_Desde
                Aux_Cant_dias = arrPar(4)
                Vac_Cant_dias = Aux_Cant_dias
                vacModelo = arrPar(6)
                If Aux_Cant_dias = 0 Then
                    Todos_Posibles = True
                Else
                    Todos_Posibles = False
                End If
            End If
            
''''''            pos1 = 1
''''''            pos2 = InStr(pos1, parametros, ".") - 1
''''''            Reproceso = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
''''''
''''''            pos1 = pos2 + 2
''''''            pos2 = InStr(pos1, parametros, ".") - 1
''''''            fecha_desde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
''''''
''''''            pos1 = pos2 + 2
''''''            pos2 = InStr(pos1, parametros, ".") - 1
''''''            fecha_hasta = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
''''''
''''''            pos1 = pos2 + 2
''''''            pos2 = InStr(pos1, parametros, ".") - 1
''''''            Aux_Fecha_Desde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
''''''            Vac_Fecha_Desde = Aux_Fecha_Desde
''''''
''''''            pos1 = pos2 + 2
''''''            pos2 = InStr(pos1, parametros, ".") - 1
''''''            'pos2 = Len(parametros) - 1
''''''            If pos2 >= pos1 Then
''''''                Aux_Cant_dias = Mid(parametros, pos1, pos2 - pos1 + 1)
''''''                Vac_Cant_dias = Aux_Cant_dias
''''''                If Aux_Cant_dias = 0 Then
''''''                    Todos_Posibles = True
''''''                Else
''''''                    Todos_Posibles = False
''''''                End If
''''''            Else
''''''                Todos_Posibles = True
''''''            End If
''''''
''''''            pos1 = pos2 + 2
''''''            pos2 = InStr(pos1, parametros, ".")
''''''            vacModelo = Mid(parametros, Len(parametros), 2)
            
        End If
    End If
       
    
    Set objFechasHoras.Conexion = objConn
    
    
    '15/02/2016 - MDZ - levanto los datos del modelo de vacaciones
    StrSql = " SELECT * FROM vac_modelos WHERE modvacnro=" & vacModelo
    OpenRecordset StrSql, objModVac
    If objModVac.EOF Then
        Flog.writeline "no se encontro el modelo"
        GoTo Final
    Else
        Flog.writeline "Modelo de vacaciones :" & objModVac("modvacdesabr") & " (" & objModVac("modvacnro") & ")"
        
    End If
    
    
    StrSql = " SELECT * FROM batch_empleado " & _
             " WHERE batch_empleado.bpronro = " & NroProceso
    OpenRecordset StrSql, objReg
    
    SinError = True
    HuboErrores = False
    
'    StrSql = "SELECT * FROM alcance_testr WHERE tanro= 21"
'    OpenRecordset StrSql, rs
'    If Not rs.EOF Then
'        l_TienePolAlcance = True
'    Else
'        l_TienePolAlcance = False
'    End If

    
    Do While Not objReg.EOF
        
        Aux_Fecha_Desde = Vac_Fecha_Desde
        Aux_Cant_dias = Vac_Cant_dias
        
        Ternro = objReg!Ternro
        
        'Flog.writeline "Inicio Empleado:" & Ternro
        
        Flog.writeline ""
        Flog.writeline "========================================================================"
        Flog.writeline EscribeLogMI("Inicio Empleado") & ": " & Ternro
        
        
        '15/02/2016 - MDZ - levanto los periodos de vacaciones
         StrSql = "select vacmodelo, vacacion_detalle.vdetfdesde vacfecdesde, vacacion_detalle.vdetfhasta vacfechasta, vacacion.vacdesc, vacacion.vacnro  from vacacion left join vacacion_detalle on (vacacion.vacnro = vacacion_detalle.vacnro) " & _
                " Where vacacion.vacestado= -1 AND vacModelo =" & vacModelo & " AND vacacion_detalle.vdetfdesde Is Not Null "
        If objModVac("modvactipoperiodo") = 2 Then
             StrSql = StrSql & " AND vacacion_detalle.Ternro =" & Ternro
        End If
        StrSql = StrSql & " ORDER BY vacanio DESC "
        
''''''''
''''''''        If objModVac("modvactipoperiodo") = 1 Then
''''''''            Flog.writeline "periodo calendario"
''''''''
''''''''            StrSql = "SELECT " & _
''''''''                    " (select MIN(vdetfdesde) from vacacion_detalle vdd where vdd.vacnro = vacacion.vacnro and vdd.vdetfdesde<= " & ConvFecha(fecha_hasta) & " AND vdd.vdetfhasta>= " & ConvFecha(fecha_desde) & " and ternro= 0) vacfecdesde, " & _
''''''''                    " (select MIN(vdetfhasta) from vacacion_detalle vdh where vdh.vacnro = vacacion.vacnro and vdh.vdetfdesde<= " & ConvFecha(fecha_hasta) & " AND vdh.vdetfhasta>= " & ConvFecha(fecha_desde) & " and ternro= 0) vacfechasta " & _
''''''''                    " ,vacacion.vacdesc, vacacion.vacnro  " & _
''''''''                    " From vacacion " & _
''''''''                    " WHERE (select MIN(vdetfdesde) from vacacion_detalle vd where vd.vacnro = vacacion.vacnro and vd.vdetfdesde<= " & ConvFecha(fecha_hasta) & " AND vd.vdetfhasta>= " & ConvFecha(fecha_desde) & " and ternro= 0) is not null"
''''''''
''''''''        ElseIf objModVac("modvactipoperiodo") = 2 Then
''''''''            Flog.writeline "periodo aniversario"
''''''''
''''''''            StrSql = "SELECT " & _
''''''''                    " (select MIN(vdetfdesde) from vacacion_detalle vdd where vdd.vacnro = vacacion.vacnro and vdd.vdetfdesde<= " & ConvFecha(fecha_hasta) & " AND vdd.vdetfhasta>= " & ConvFecha(fecha_desde) & " and ternro= " & Ternro & " ) vacfecdesde, " & _
''''''''                    " (select MIN(vdetfhasta) from vacacion_detalle vdh where vdh.vacnro = vacacion.vacnro and vdh.vdetfdesde<= " & ConvFecha(fecha_hasta) & " AND vdh.vdetfhasta>= " & ConvFecha(fecha_desde) & " and ternro= " & Ternro & " ) vacfechasta " & _
''''''''                    " ,vacacion.vacdesc, vacacion.vacnro  " & _
''''''''                    " From vacacion " & _
''''''''                    " WHERE (select MIN(vdetfdesde) from vacacion_detalle vd where vd.vacnro = vacacion.vacnro and vd.vdetfdesde<= " & ConvFecha(fecha_hasta) & " AND vd.vdetfhasta>= " & ConvFecha(fecha_desde) & " and ternro= " & Ternro & " ) is not null"
''''''''
''''''''
''''''''        Else
''''''''            Flog.writeline "tipo de periodo desconocido"
''''''''            GoTo Final
''''''''        End If
''''''''
        'Flog.writeline StrSql
        OpenRecordset StrSql, rs_Periodos_Vac
        
            
        'Seteo el incremento del progreso
        CEmpleadosAProc = objReg.RecordCount
        If CEmpleadosAProc = 0 Then
            CEmpleadosAProc = 1
        End If
        CDiasAProc = rs_Periodos_Vac.RecordCount
        If CDiasAProc = 0 Then
            CDiasAProc = 1
        End If
        IncPorc = ((100 / CEmpleadosAProc) * (100 / CDiasAProc)) / 100
        
        'MDZ - 16/02/2016 - poolitica obsoleta, se resualve por tipo de dia utilizado en el periodo, campo tipodia
       'FGZ - 24/06/2009 -------
        'Diashabiles_LV = False
        'PoliticaOK = False
        'Call Politica(1510)
        'FGZ - 24/06/2009 -------
        
        
        
        
        'Aux_Fecha_Desde = fecha_desde
        Do While Not rs_Periodos_Vac.EOF
            Continua = False
            
            MyBeginTrans
            Flog.writeline EscribeLogMI("Periodo") & ": " & rs_Periodos_Vac!vacfecdesde & " - " & rs_Periodos_Vac!vacfechasta
                        
            'EAM- Obtiene el estado del periodo 05-10-2010
            StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & rs_Periodos_Vac!vacnro & " AND (venc = 1)"
            OpenRecordset StrSql, rs_vacdiascor
            
            'Estados (0 Cerrado | -1 abierto)
            estadoPeriodo = 0
            If rs_vacdiascor.EOF Then
                estadoPeriodo = -1
            End If
            rs_vacdiascor.Close
            
                        
            'Toma el primer período que se encuentre abierto
            If (estadoPeriodo = -1) Then
        
'                If objModVac("modvactipoperiodo") = 1 Then
                    'Valido que la fecha este dentro del período de vacaciones
                    If (Aux_Fecha_Desde >= rs_Periodos_Vac!vacfecdesde) And (Aux_Fecha_Desde <= rs_Periodos_Vac!vacfechasta) Then
                       Continua = True
                       Call GeneraPedido_ARG(Aux_Fecha_Desde, rs_Periodos_Vac!vacnro, rs_Periodos_Vac!vacdesc, alcannivel, Reproceso)
                    End If
                    
'                ElseIf objModVac("modvactipoperiodo") = 2 Then
'                    If (Aux_Fecha_Desde >= rs_Periodos_Vac!vacfechasta) And (Aux_Fecha_Desde <= DateAdd("m", 12, rs_Periodos_Vac!vacfechasta)) Then
'                       Continua = True
'                       Aux_Fecha_Hasta = DateAdd("m", 12, rs_Periodos_Vac!vacfechasta)
'                       'Call GeneraPedido_PY(Aux_Fecha_Desde, rs_Periodos_Vac!vacnro, rs_Periodos_Vac!vacdesc, alcannivel, Reproceso)
'                       Call GeneraPedido_ARG(Aux_Fecha_Desde, rs_Periodos_Vac!vacnro, rs_Periodos_Vac!vacdesc, alcannivel, Reproceso)
'                    End If
'                End If
                 
                           
            
               'Call GeneraPedido_ARG(Aux_Fecha_Desde, rs_Periodos_Vac!vacnro, rs_Periodos_Vac!vacdesc, alcannivel)
               'si la fecha en la que se va a generar los dias pedidos estan fuera del rengo de fechas del periodo
               'no se procesan
                If Continua = False Then
                    Flog.writeline "La fecha en la que se va a generar los dias pedidos estan fuera del rango de fechas del periodo " & Aux_Fecha_Desde
                End If
            Else
                Flog.writeline "El perido " & rs_Periodos_Vac!vacdesc & " (" & rs_Periodos_Vac!vacnro & ") se encuentra cerrado"
            End If
                
            MyCommitTrans
                
                
' ----------------------------------------------------------
siguiente:
            Progreso = Progreso + IncPorc
                
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
                
            If SinError Then
                 ' borro
                 StrSql = "DELETE FROM batch_empleado WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                 objConnProgreso.Execute StrSql, , adExecuteNoRecords
            Else
                 StrSql = "UPDATE batch_empleado SET estado = 'Error' WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                 objConnProgreso.Execute StrSql, , adExecuteNoRecords
            End If
        
        
        'FGZ - 27/09
            rs_Periodos_Vac.MoveNext
        Loop
        
        objReg.MoveNext
    Loop


'Deshabilito el manejador de errores
On Error GoTo 0

Final:
Flog.writeline EscribeLogMI("Fin") & " :" & Now
Flog.Close
   
    If HuboErrores Then
        ' actualizo el estado del proceso a Error
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        ' poner el bprcestado en procesado
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        
        ' -----------------------------------------------------------------------------------
        'FGZ - 22/09/2003
        'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso

        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
                 ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!iduser & "'"
        
        If Not IsNull(rs_Batch_Proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchora & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfechasta)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcprogreso
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecfin)
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!empnro
        End If
        If Not IsNull(rs_Batch_Proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcPid
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcUrgente
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
        ' FGZ - 22/09/2003
        ' -----------------------------------------------------------------------------------
    End If
        
    
Fin:
    objConn.Close
    objConnProgreso.Close
    Set objConn = Nothing
    Set objConnProgreso = Nothing
    If objReg.State = adStateOpen Then objReg.Close
    Set objReg = Nothing
    
    Exit Sub
    
    
ce:
    MyRollbackTrans
    HuboErrores = True
    SinError = False
    
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline EscribeLogMI("Error procesando Empleado") & ": " & Ternro & " " & fecha
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"
    GoTo siguiente
End Sub

Private Sub initVariablesTurno(ByRef T As BuscarTurno)
   p_turcomp = T.Compensa_Turno
   nro_grupo = T.Empleado_Grupo
   nro_justif = T.Justif_Numero
   justif_turno = T.justif_turno
   Tiene_Justif = T.Tiene_Justif
   Fecha_Inicio = T.FechaInicio
   Nro_fpgo = T.Numero_FPago
   Nro_Turno = T.Turno_Numero
   tiene_turno = T.tiene_turno
   Tipo_Turno = T.Turno_Tipo
   P_Asignacion = T.Tiene_PAsignacion
End Sub

Private Sub initVariablesDia(ByRef D As BuscarDia)
   Dia_Libre = D.Dia_Libre
   Nro_Dia = D.Numero_Dia
   Nro_Subturno = D.SubTurno_Numero
   Orden_Dia = D.Orden_Dia
   Trabaja = D.Trabaja
End Sub

Function activo(ByVal Tercero As Long, ByVal desde As Date, ByVal hasta As Date) As Boolean
Dim sql As String
Dim salida As Boolean
Dim rs_actv As New ADODB.Recordset

    'Busco una fase que contenga completamente el rango y este activa
    sql = " SELECT * "
    sql = sql & " FROM fases "
    sql = sql & " WHERE fases.empleado = " & Tercero
    sql = sql & " AND ( (altfec <= " & ConvFecha(desde) & ") AND ((bajfec >= " & ConvFecha(hasta) & ") OR (bajfec is null)) ) "
    sql = sql & " AND fases.estado = -1  "
    OpenRecordset sql, rs_actv
    
    If rs_actv.EOF Then
        salida = False
    Else
        salida = True
    End If
    rs_actv.Close
    
    activo = salida

End Function

'Private Sub DiasPedidos2(TipoVac As Long, FechaInicial As Date, FechaFinal As Date, Ternro As Long, ByRef CDias As Integer, ByRef cHabiles As Integer, ByRef cFeriados As Integer)
''Descripcion: Calcula la cantidad de días (LU-VI) entre dos fechas.
'
'Dim Fecha As Date
'Dim objFeriado As New Feriado
'Dim DHabiles(1 To 7) As Boolean
'Dim esFeriado As Boolean
'
'    CDias = 0
'    cHabiles = 0
'    cFeriados = 0
'
'    Fecha = FechaInicial
'
'    Set objFeriado.Conexion = objConn
'    Set objFeriado.ConexionTraza = objConn
'
'    StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & TipoVac
'    OpenRecordset StrSql, objRs
'    If objRs.EOF Then Exit Sub
'
'    DHabiles(1) = objRs!tpvhabiles__1
'    DHabiles(2) = objRs!tpvhabiles__2
'    DHabiles(3) = objRs!tpvhabiles__3
'    DHabiles(4) = objRs!tpvhabiles__4
'    DHabiles(5) = objRs!tpvhabiles__5
'    DHabiles(6) = objRs!tpvhabiles__6
'    DHabiles(7) = objRs!tpvhabiles__7
'
'
'    Do While Fecha <= FechaFinal
'
'        esFeriado = objFeriado.Feriado(Fecha, Ternro, False)
'
'        If (esFeriado) And (objRs!tpvferiado = 0) Then
'            cFeriados = cFeriados + 1
'        Else
'            If DHabiles(Weekday(Fecha)) Or ((esFeriado) And (objRs!tpvferiado = -1)) Then
'                CDias = CDias + 1
'            Else
'                cHabiles = cHabiles + 1
'            End If
'        End If
'
'        Fecha = DateAdd("d", 1, Fecha)
'    Loop
'
'    Set objFeriado = Nothing
'
'End Sub

'Private Sub DiasPedidos1(TipoVac As Long, FechaInicial As Date, ByRef Fecha As Date, Ternro As Long, cant As Integer, ByRef cHabiles As Integer, ByRef cFeriados As Integer)
''Calcula la fecha hasta a partir de la fecha desde, la cantidad de dias pedidos y el tipo
''de vacacion asociado a los dias correspòndientes, para el período
'
'Dim i As Integer
'Dim j As Integer
'Dim objFeriado As New Feriado
'Dim DHabiles(1 To 7) As Boolean
'Dim esFeriado As Boolean
'Dim objRs As New ADODB.Recordset
'
'    i = 0
'    j = 0
'    cHabiles = 0
'    cFeriados = 0
'
'    Fecha = FechaInicial
'
'    Set objFeriado.Conexion = objConn
'    Set objFeriado.ConexionTraza = objConn
'
'    StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & TipoVac
'    OpenRecordset StrSql, objRs
'    If objRs.EOF Then Exit Sub
'
'    DHabiles(1) = objRs!tpvhabiles__1
'    DHabiles(2) = objRs!tpvhabiles__2
'    DHabiles(3) = objRs!tpvhabiles__3
'    DHabiles(4) = objRs!tpvhabiles__4
'    DHabiles(5) = objRs!tpvhabiles__5
'    DHabiles(6) = objRs!tpvhabiles__6
'    DHabiles(7) = objRs!tpvhabiles__7
'
'
'    Do While i <= cant
'
'        esFeriado = objFeriado.Feriado(Fecha, Ternro, False)
'
'        If (esFeriado) And (objRs!tpvferiado = 0) Then
'            cFeriados = cFeriados + 1
'        Else
'            If DHabiles(Weekday(Fecha)) Or ((esFeriado) And (objRs!tpvferiado = -1)) Then
'                i = i + 1
'            Else
'                cHabiles = cHabiles + 1
'            End If
'        End If
'        If i < cant Then Fecha = DateAdd("d", 1, Fecha)
'    Loop
'
'    Set objFeriado = Nothing
'
'End Sub
'
'
'Private Sub DiasPedidos3(TipoVac As Long, FechaInicial As Date, FechaFinal As Date, Ternro As Long, ByRef CDias As Integer, ByRef cHabiles As Integer, ByRef cFeriados As Integer)
''Descripcion: Calcula la cantidad de días (LU-Sab) entre dos fechas.
'
'Dim Fecha As Date
'Dim objFeriado As New Feriado
'Dim DHabiles(1 To 7) As Boolean
'Dim esFeriado As Boolean
'
'    CDias = 0
'    cHabiles = 0
'    cFeriados = 0
'
'    Fecha = FechaInicial
'
'    Set objFeriado.Conexion = objConn
'    Set objFeriado.ConexionTraza = objConn
'
'    StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & TipoVac
'    OpenRecordset StrSql, objRs
'    If objRs.EOF Then Exit Sub
'
'    DHabiles(1) = objRs!tpvhabiles__1
'    DHabiles(2) = objRs!tpvhabiles__2
'    DHabiles(3) = objRs!tpvhabiles__3
'    DHabiles(4) = objRs!tpvhabiles__4
'    DHabiles(5) = objRs!tpvhabiles__5
'    DHabiles(6) = objRs!tpvhabiles__6
'    DHabiles(7) = objRs!tpvhabiles__7
'
'
'    Do While Fecha <= FechaFinal
'
'        esFeriado = objFeriado.Feriado(Fecha, Ternro, False)
'
'        If (esFeriado) And (objRs!tpvferiado = 0) Then
'            cFeriados = cFeriados + 1
'        Else
'            If (Weekday(Fecha) = 7) Or DHabiles(Weekday(Fecha)) Or ((esFeriado) And (objRs!tpvferiado = -1)) Then
'                CDias = CDias + 1
'            Else
'                cHabiles = cHabiles + 1
'            End If
'        End If
'
'        Fecha = DateAdd("d", 1, Fecha)
'    Loop
'
'    Set objFeriado = Nothing
'
'End Sub


Public Sub DiasPedidos_STD(ByVal tipoVac As Long, ByVal FechaInicial As Date, ByRef fecha As Date, ByVal Ternro As Long, ByRef cant As Integer, ByRef CHabiles As Integer, ByRef cNoHabiles As Integer, ByRef cFeriados As Integer)
'Calcula la fecha hasta a partir de la fecha desde, la cantidad de dias pedidos y el tipo
'de vacacion asociado a los dias correspòndientes, para el período
'se cambio de private a public

Dim i As Integer
Dim j As Integer
Dim objFeriado As New Feriado
Dim DHabiles(1 To 7) As Boolean
Dim EsFeriado As Boolean
Dim objRs As New ADODB.Recordset
Dim ExcluyeFeriados As Boolean


    StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & tipoVac
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        DHabiles(1) = objRs!tpvhabiles__1
        DHabiles(2) = objRs!tpvhabiles__2
        DHabiles(3) = objRs!tpvhabiles__3
        DHabiles(4) = objRs!tpvhabiles__4
        DHabiles(5) = objRs!tpvhabiles__5
        DHabiles(6) = objRs!tpvhabiles__6
        DHabiles(7) = objRs!tpvhabiles__7
    
        ExcluyeFeriados = CBool(objRs!tpvferiado)
        Flog.writeline EscribeLogMI("Tipo de Vacaciones") & ": " & objRs!tipvacdesabr
    Else
        Flog.writeline EscribeLogMI("No se encontro el tipo de Vacacion") & ": " & tipoVac
        Exit Sub
    End If

    
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = objConn
    
    i = 0
    j = 0
    CHabiles = 0
    cNoHabiles = 0
    cFeriados = 0
    
    fecha = FechaInicial
    
    Do While i < cant
    
        EsFeriado = objFeriado.Feriado(fecha, Ternro, False)
        
        If (EsFeriado) And Not ExcluyeFeriados Then
            cFeriados = cFeriados + 1
            'FGZ - 24/06/2009 -----------------------
            If PoliticaOK And Diashabiles_LV Then
                i = i + 1
            End If
            'FGZ - 24/06/2009 -----------------------
        Else
            If DHabiles(Weekday(fecha)) Or (EsFeriado And ExcluyeFeriados) Then
                i = i + 1
                If DHabiles(Weekday(fecha)) Then
                    CHabiles = CHabiles + 1
                End If
            Else
                cNoHabiles = cNoHabiles + 1
                'FGZ - 24/06/2009 -----------------------
                If PoliticaOK And Diashabiles_LV Then
                    i = i + 1
                End If
                'FGZ - 24/06/2009 -----------------------
            End If
'            If DHabiles(Weekday(Fecha)) Or ((esFeriado) And ExcluyeFeriados) Then
'                i = i + 1
'            Else
'                cHabiles = cHabiles + 1
'            End If
        End If
        
        If i < cant Then
            fecha = DateAdd("d", 1, fecha)
        Else
            i = i + 1
        End If
    Loop
    
    Set objFeriado = Nothing

End Sub


Public Function CalcularFechaHasta(ByVal fdesde As Date, ByVal dias As Integer, ByVal tipoVac As Integer, ByRef hab As Integer, ByRef nohab As Integer, ByRef feri As Integer)
            
        feri = 0
        hab = 0
        nohab = 0
        
        Dim objRs As New ADODB.Recordset
        Dim DHabiles(1 To 7) As Boolean
        Dim ExcluyeFeriados As Boolean
        Dim EsFeriado As Boolean
        
        StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & tipoVac
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            DHabiles(1) = objRs!tpvhabiles__1
            DHabiles(2) = objRs!tpvhabiles__2
            DHabiles(3) = objRs!tpvhabiles__3
            DHabiles(4) = objRs!tpvhabiles__4
            DHabiles(5) = objRs!tpvhabiles__5
            DHabiles(6) = objRs!tpvhabiles__6
            DHabiles(7) = objRs!tpvhabiles__7
        
            ExcluyeFeriados = CBool(objRs!tpvferiado)
            Flog.writeline EscribeLogMI("Tipo de Vacaciones") & ": " & objRs!tipvacdesabr
        Else
            Flog.writeline EscribeLogMI("No se encontro el tipo de Vacacion") & ": " & tipoVac
            Exit Function
        End If
        
        
        Dim fecha As Date
        fecha = fdesde
        Dim cant As Integer
        cant = 1
        
        Do

            EsFeriado = objFeriado.Feriado(fecha, Ternro, False)
            
            If EsFeriado Then
                
                If Not ExcluyeFeriados Then
                    feri = feri + 1 'sumo al contador de dias feriados
                    
                    cant = cant + 1
                End If
           
            Else
                If DHabiles(Weekday(fecha)) = -1 Then
                    '27/04/2015 - MDZ - si excluye francos verifico el turno
                    
                    hab = hab + 1 'sumo al contador de dias habiles
                    cant = cant + 1
                    
                Else
                    nohab = nohab + 1    'sumo al contador de dias NO habiles
                End If
            End If
            
            fecha = DateAdd("d", 1, fecha)
            
        Loop While cant <= CInt(dias)
        
        CalcularFechaHasta = DateAdd("d", -1, fecha)

End Function

