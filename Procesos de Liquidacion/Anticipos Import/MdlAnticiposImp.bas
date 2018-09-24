Attribute VB_Name = "MdlAnticiposImp"
Option Explicit

Const Version = 1.01 'Stankunas Cesar - Version Inicial
Const FechaVersion = "13/07/2010"


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial Anticipos.
' Autor      : Stankunas Cesar
' Fecha      : 26/04/2010
' Ultima Mod.: 07/07/2010 - Stankunas Cesar - Se agregó el calculo del Neto cuando es Tipo Remunerativo
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
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
    
    Nombre_Arch = PathFLog & "Importacion Anticipos" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 266 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Generar_anticipos(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Close
    objConn.Close
    objconnProgreso.Close
End Sub


Public Sub Generar_anticipos(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de anticipos
' Autor      : Cesar Stankunas
' Fecha      : 26/04/2010
' Ult. Mod   : 07/07/2010 - Stankunas Cesar - Se agregó el calculo del Neto cuando es Tipo Remunerativo
' --------------------------------------------------------------------------------------------
'Arreglo que contiene los parametros
Dim arrParam
Dim I As Long

'Parametros
Dim TanticipoNro As Long
Dim Lista_Pro As String
Dim PliqNro As Long
Dim PliqNroDto As Long
Dim Todos_Procesos As Boolean
Dim Todos_Empleados As Boolean
Dim Proaprob As Integer
Dim Usuario As String
Dim Remunerativo As String

'RecordSets
Dim rs_Consult As New ADODB.Recordset
Dim rs_Anticipos As New ADODB.Recordset
Dim rs_TipoAnticipo As New ADODB.Recordset

'Variables
Dim Lista_Conc As String
Dim Lista_Empleados As String
Dim Monto As Double
Dim AntNeto As Double
Dim Cantidad As Double
Dim Moneda As Long

On Error GoTo CE
    
' Levanto cada parametro por separado, el separador de parametros es "@"
'l_tanticiponro & "@" & l_pliqnro & "@" & l_todospro & "@" & l_pronro & "@" & l_todosemp & "@" & l_proaprob & "@" & l_usuario & "@" & l_pliqnrodto & "@" & l_remun
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    arrParam = Split(Parametros, "@")
    
    If UBound(arrParam) = 8 Then
        
        TanticipoNro = CLng(arrParam(0))
        PliqNro = CLng(arrParam(1))
        Todos_Procesos = CBool(arrParam(2))
        Lista_Pro = arrParam(3)
        Todos_Empleados = CBool(arrParam(4))
        Proaprob = CInt(arrParam(5))
        Usuario = arrParam(6)
        PliqNroDto = CLng(arrParam(7))
        Remunerativo = arrParam(8)
        If Remunerativo = "" Or IsNull(Remunerativo) Then
            Remunerativo = "0"
        End If
        
        Flog.writeline Espacios(Tabulador * 1) & "Tipo de anticipo = " & TanticipoNro
        Flog.writeline Espacios(Tabulador * 1) & "Periodo = " & PliqNro
        Flog.writeline Espacios(Tabulador * 1) & "Todos los Procesos = " & Todos_Procesos
        If Not Todos_Procesos Then
            Flog.writeline Espacios(Tabulador * 1) & "Lista de procesos = " & Lista_Pro
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Todos los empleados = " & Todos_Empleados
        Flog.writeline Espacios(Tabulador * 1) & "Procesos Aprobados = " & Proaprob
        Flog.writeline Espacios(Tabulador * 1) & "Usuario = " & Usuario
        Flog.writeline Espacios(Tabulador * 1) & "Periodo de Decuento = " & PliqNroDto
        Flog.writeline Espacios(Tabulador * 1) & "Remunerativo = " & Remunerativo
    Else
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La cantidad de parametros no es la esperada."
        Exit Sub
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encuentran los paramentros."
    Exit Sub
End If

Flog.writeline

'Comienzo la transaccion
MyBeginTrans

'--------------------------------------------------------------------------------------
'Si se selecciono Todos los Procesos los busco y armo la lista
'--------------------------------------------------------------------------------------
If Todos_Procesos Then
    Flog.writeline Espacios(Tabulador * 0) & "Buscando Todos los procesos del periodo " & PliqNro
    StrSql = "SELECT * FROM proceso"
    StrSql = StrSql & " WHERE pliqnro = " & PliqNro
    StrSql = StrSql & " AND proceso.proaprob = " & Proaprob
    
    OpenRecordset StrSql, rs_Consult
    
    If Not rs_Consult.EOF Then
        Lista_Pro = rs_Consult!pronro
        rs_Consult.MoveNext
        
        Do Until rs_Consult.EOF = True
            Lista_Pro = Lista_Pro & "," & rs_Consult!pronro
            rs_Consult.MoveNext
        Loop
        Flog.writeline Espacios(Tabulador * 1) & "Lista de procesos = " & Lista_Pro
        
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No hay procesos para el periodo."
        GoTo Fin
    End If
    
    rs_Consult.Close
    
End If

Flog.writeline

'--------------------------------------------------------------------------------------
'Busco la configuracion en el confrep
'--------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando los Conceptos del Confrep 281"
StrSql = " SELECT concepto.concnro, concepto.conccod"
StrSql = StrSql & " FROM confrep"
StrSql = StrSql & " INNER JOIN concepto ON concepto.conccod = confrep.confval2"
StrSql = StrSql & " WHERE repnro = 281 AND conftipo = 'CO' "
OpenRecordset StrSql, rs_Consult

If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentran los conceptos del confrep 281."
    GoTo Fin
Else
    Lista_Conc = rs_Consult!ConcNro
    Flog.writeline Espacios(Tabulador * 1) & "Concepto - " & rs_Consult!Conccod
    rs_Consult.MoveNext
    
    Do Until rs_Consult.EOF = True
        Lista_Conc = Lista_Conc & "," & rs_Consult!ConcNro
        Flog.writeline Espacios(Tabulador * 1) & "Concepto - " & rs_Consult!Conccod
        rs_Consult.MoveNext
    Loop
End If

rs_Consult.Close

Flog.writeline

'--------------------------------------------------------------------------------------
'Busco los Empleados
'--------------------------------------------------------------------------------------
If Not Todos_Empleados Then
    Flog.writeline Espacios(Tabulador * 0) & "Buscando empleados."
    StrSql = "SELECT ternro FROM batch_empleado "
    StrSql = StrSql & " WHERE bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_Consult
    
    If rs_Consult.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encuentran los Empleados."
        GoTo Fin
    Else
        Lista_Empleados = rs_Consult!Ternro
        rs_Consult.MoveNext
        
        Do Until rs_Consult.EOF = True
            Lista_Empleados = Lista_Empleados & "," & rs_Consult!Ternro
            rs_Consult.MoveNext
        Loop
    End If

    rs_Consult.Close
    
Else
    Lista_Empleados = ""
End If


Flog.writeline


'--------------------------------------------------------------------------------------
'Busco la moneda de origen del pais default
'--------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando la moneda de origen del pais default."
StrSql = " SELECT monnro, mondesabr, paisdesc, monorigen, monlocal, pais.paisnro"
StrSql = StrSql & " From moneda"
StrSql = StrSql & " INNER join pais ON moneda.paisnro = pais.paisnro"
StrSql = StrSql & " Where pais.paisdef = -1 And moneda.monorigen = -1"
OpenRecordset StrSql, rs_Consult

If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encuentran la moneda."
    GoTo Fin
Else
    Moneda = IIf(EsNulo(rs_Consult!monnro), 0, rs_Consult!monnro)
    Flog.writeline Espacios(Tabulador * 1) & "Pais: " & rs_Consult!paisnro & " - " & rs_Consult!paisdesc & " Moneda: " & rs_Consult!monnro & " - " & rs_Consult!mondesabr
    If Moneda = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "Moneda sin codigo."
        GoTo Fin
    End If
End If

rs_Consult.Close

Flog.writeline

Flog.writeline Espacios(Tabulador * 0) & "-----------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "CONSULTA PRINCIPAL - Buscando Detliq."
Flog.writeline Espacios(Tabulador * 0) & "-----------------------------------------------------"
Flog.writeline
'--------------------------------------------------------------------------------------
'CONSULTA PRINCIPAL
'Busco Todos los detliq de los conceptos de los empleados
'--------------------------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto, cabliq.pronro, proceso.prodesc, periodo.pliqhasta, periodo.pliqnro,"
StrSql = StrSql & " empleado.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2"
StrSql = StrSql & " FROM cabliq"
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro"
'Filtro de Procesos
StrSql = StrSql & " AND cabliq.pronro IN (" & Lista_Pro & ")"
StrSql = StrSql & " INNER JOIN periodo ON proceso.pliqnro = periodo.pliqnro"
StrSql = StrSql & " INNER JOIN detliq  ON cabliq.cliqnro = detliq.cliqnro"
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado"
'Filtro de empleados
If Not Todos_Empleados Then
    StrSql = StrSql & " AND empleado.ternro IN ( " & Lista_Empleados & " )"
End If
'Filtro de conceptos
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
StrSql = StrSql & " AND concepto.concnro IN ( " & Lista_Conc & " )"
StrSql = StrSql & " ORDER BY cabliq.pronro, concepto.conccod, empleado.empleg "
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    'Progreso
    Progreso = 0
    CEmpleadosAProc = rs_Consult.RecordCount
    If CEmpleadosAProc = 0 Then
       CEmpleadosAProc = 1
    End If
    IncPorc = CEmpleadosAProc / 100
    
    Do While Not rs_Consult.EOF
        
        Flog.writeline Espacios(Tabulador * 1) & "Procesando " & rs_Consult!empleg & " - " & rs_Consult!terape & " " & rs_Consult!ternom & " PROCESO: " & rs_Consult!pronro & " - " & rs_Consult!prodesc & " CONCEPTO: " & rs_Consult!Conccod & " - " & rs_Consult!concabr & " MONTO: " & rs_Consult!dlimonto & " CANTIDAD: " & rs_Consult!dlicant
        
        Monto = rs_Consult!dlimonto
        Cantidad = IIf(EsNulo(rs_Consult!dlicant), 1, IIf(rs_Consult!dlicant = 0, 1, rs_Consult!dlicant))
        
        If Monto <> 0 Then
            Flog.writeline
            'Creo Cantidad de anticipos con monto = monto/cantidad
            Flog.writeline Espacios(Tabulador * 2) & "Creando " & Fix(Cantidad) & " anticipos Por " & Monto
            For I = 1 To Fix(Cantidad)
                'Busco el Neto
                AntNeto = 0
                
                If Remunerativo = "-1" Then
                    StrSql = " SELECT tdporc FROM tipodescuento"
                    StrSql = StrSql & " INNER JOIN tipoanticipo ON tipodescuento.tdtipant = tipoanticipo.tanticiponro"
                    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.tenro = tipodescuento.tdtipestrnro AND his_estructura.estrnro = tipodescuento.tdestrnro"
                    StrSql = StrSql & " WHERE ternro = " & rs_Consult!Ternro
                    StrSql = StrSql & " AND tdtipant = " & TanticipoNro
                    StrSql = StrSql & " AND htethasta is NULL"
                    OpenRecordset StrSql, rs_Anticipos
                    If Not rs_Anticipos.EOF Then
                        AntNeto = CDbl(Monto) - (CDbl(Monto) * CDbl(rs_Anticipos!tdporc) / 100)
                    Else
                        AntNeto = CDbl(Monto)
                    End If
                Else
                    AntNeto = CDbl(Monto)
                End If
                
                'Inserto en anticipos
                StrSql = "INSERT INTO anticipos "
                StrSql = StrSql & "(empleado, ppagnro, monnro, antmonto, antfecped, pliqnro, antdesc,"
                StrSql = StrSql & " pliqdto, pronro, tanticiponro, antrevis, antusuario, antestado, antneto,pliqretro) "
                StrSql = StrSql & " VALUES ("
                StrSql = StrSql & rs_Consult!Ternro & ","
                StrSql = StrSql & "NULL" & ","
                StrSql = StrSql & Moneda & ","
                StrSql = StrSql & Monto & ","
                StrSql = StrSql & ConvFecha(rs_Consult!pliqhasta) & ","
                StrSql = StrSql & rs_Consult!PliqNro & ",'"
                StrSql = StrSql & IIf(EsNulo(rs_Consult!concabr), "", Mid(rs_Consult!concabr, 1, 30)) & "',"
                StrSql = StrSql & PliqNroDto & ","
                StrSql = StrSql & "NULL" & ","
                StrSql = StrSql & TanticipoNro & ","
                StrSql = StrSql & "0" & ",'"
                StrSql = StrSql & Usuario & "',"
                StrSql = StrSql & "1,"
                StrSql = StrSql & AntNeto & ","
                StrSql = StrSql & "0"
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Next
            Flog.writeline Espacios(Tabulador * 2) & "Creacion OK."
        End If
        
        Flog.writeline
        rs_Consult.MoveNext
        
        
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        
    Loop
    
    rs_Consult.Close
            
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontraron detliq para los conceptos."
End If

Fin:
'Fin de la transaccion
MyCommitTrans

'Cierro todo y libero
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing

Exit Sub

CE:
    MyRollbackTrans
    
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline
    Flog.writeline "**********************************************************"
    Flog.writeline " Empleado abortado: "
    Flog.writeline " Error: " & Err.Description
    Flog.writeline "**********************************************************"
    Flog.writeline
End Sub


