Attribute VB_Name = "MdlRHProExpLiq"
Option Explicit

'Global Const Version = "1.00" 'Exportacion de datos de Liquidacion
'Global Const FechaModificacion = "22/10/2008"
'Global Const UltimaModificacion = "" 'Martin Ferraro - Version Inicial


Global Const Version = "1.1" 'Exportacion de datos de Liquidacion
Global Const FechaModificacion = "12/09/2014"
Global Const UltimaModificacion = "" 'Sebastian Stremel - Se corrigio nombre de campo esfecha de la tabla escala
                                     ' Se agrega carpetas por usuarios - CAS-24538 - CCU - MEJORA EN SEGURIDAD EN IN-OUT


Global Seed As String 'Usado como clave de encriptacion/desencriptacion
Global encryptAct As Boolean
Global ArrSQL(31) As String
Global iduser





Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Exportacion.
' Autor      : Martin Ferraro
' Fecha      : 22/10/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim bprcFecha As Date
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
    
    Nombre_Arch = PathFLog & "Exp. Liquidacion RHPro" & " - " & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 230 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        bprcFecha = rs_batch_proceso!bprcFecha
        iduser = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call ExportLiq(NroProcesoBatch, bprcparam, bprcFecha)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
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


Public Sub ExportLiq(ByVal bpronro As Long, ByVal Parametros As String, ByVal bprcFecha As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Exportacion de datos de liquidacion
' Autor      : Martin Ferraro
' Fecha      : 22/10/2008
' --------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'Parametros
'-------------------------------------------------------------------------------------------------
Dim Auto As Boolean
Dim Periodo As Long
Dim Fecha As Date

'-------------------------------------------------------------------------------------------------
'Variables
'-------------------------------------------------------------------------------------------------
Dim ArrPar
Dim FecDesde As Date
Dim FecHasta As Date
Dim Mes As Integer
Dim Anio As Long
Dim directorio As String
Dim Sep As String
Dim SeparadorDecimal As String
Dim DescripcionModelo As String
Dim Archivo As String
Dim fExport
Dim carpeta
Dim cant As Long
Dim Linea As String
Dim CantReg As Long
Dim Corrimiento As Long

'-------------------------------------------------------------------------------------------------
'RecordSets
'-------------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset

'Inicio codigo ejecutable
On Error GoTo E_ExportLiq

'Valores default de encriptacion: Activa y semilla = 56238
Seed = "56238"
encryptAct = True
'-------------------------------------------------------------------------------------------------
'Levanto cada parametro por separado, el separador de parametros es "@"
'-------------------------------------------------------------------------------------------------
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    If Len(Parametros) >= 1 Then
    
        ArrPar = Split(Parametros, "@")
            
        If UBound(ArrPar) >= 0 Then
            Auto = CBool(ArrPar(0))
            
            If Not Auto Then
                Periodo = CLng(ArrPar(1))
                Flog.writeline Espacios(Tabulador * 0) & "Disparo de Exportacion Manual Periodo " & Periodo
            Else
                Fecha = bprcFecha
                Flog.writeline Espacios(Tabulador * 0) & "Disparo de Exportacion Planificada Fecha " & Fecha
                Mes = Month(Fecha)
                Anio = Year(Fecha)
            End If
        Else
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Faltan parametros."
            Exit Sub
        End If
    Else
        Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontraron los parametros."
        Exit Sub
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontraron los parametros."
    Exit Sub
End If
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Configuracion del Reporte
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuracion del reporte."
StrSql = "SELECT * FROM confrep WHERE repnro = 245 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult
If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte 245."
Else

    Do While Not rs_Consult.EOF

        Select Case rs_Consult!confnrocol
            Case 1:
                'Aca se configura el corrimiento de meses para atras a procesar cuando se corre en forma auto
                If Auto Then
                    Corrimiento = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
                    If Corrimiento <> 0 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Corrimiento de procesamiento " & Corrimiento & " mes/es."
                        Fecha = DateAdd("M", -1 * Corrimiento, Fecha)
                        Mes = Month(Fecha)
                        Anio = Year(Fecha)
                        Flog.writeline Espacios(Tabulador * 1) & "Nueva fecha con corrimiento " & Fecha
                    End If
                End If
            Case 999:
                If rs_Consult!conftipo = 0 Then
                    encryptAct = False
                Else
                    Seed = IIf(EsNulo(rs_Consult!confval2), "56238", rs_Consult!confval2)
                End If
        End Select

        rs_Consult.MoveNext
    Loop
End If
rs_Consult.Close
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Completo el dato que falta segun sea auto o no
'-------------------------------------------------------------------------------------------------
If Auto Then
    'Tengo una fecha, busco el periodo que esta en la fecha
    Flog.writeline Espacios(Tabulador * 0) & "Buscando Periodo a Mes " & Mes & " Año " & Anio
    StrSql = "SELECT pliqnro, pliqdesc, pliqdesde, pliqhasta, pliqmes, pliqanio"
    StrSql = StrSql & " FROM periodo"
    StrSql = StrSql & " WHERE pliqanio = " & Anio
    StrSql = StrSql & " AND pliqmes = " & Mes
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        Periodo = rs_Consult!pliqnro
        Flog.writeline Espacios(Tabulador * 1) & "Periodo encontrado " & Periodo & " - " & rs_Consult!pliqdesc
    Else
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el periodo a la fecha"
        HuboError = True
        Exit Sub
    End If
Else
    'Tengo un periodo, busco el mes y el año
    Flog.writeline Espacios(Tabulador * 0) & "Buscando Mes y Anio del Periodo " & Periodo
    StrSql = "SELECT pliqnro, pliqdesc, pliqdesde, pliqhasta, pliqmes, pliqanio"
    StrSql = StrSql & " FROM periodo"
    StrSql = StrSql & " WHERE pliqnro = " & Periodo
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        Mes = rs_Consult!pliqmes
        Anio = rs_Consult!pliqanio
        Flog.writeline Espacios(Tabulador * 1) & "Mes encontrado " & Mes & " Año " & Anio
    Else
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el periodo."
        HuboError = True
        Exit Sub
    End If
End If
    

'-------------------------------------------------------------------------------------------------
'Buscando Fecha desde y hasta
'-------------------------------------------------------------------------------------------------
FecDesde = PrimerDiaMes(Mes, Anio)
FecHasta = UltimoDiaMes(Mes, Anio)


'-------------------------------------------------------------------------------------------------
'Configuracion del Directorio de salida
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando directorio de salida."
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    directorio = Trim(rs_Consult!sis_dirsalidas)
    If "\" <> CStr(Right(directorio, 1)) Then
        directorio = directorio & "\"
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el registro de la tabla sistema nro 1"
    HuboError = True
    Exit Sub
End If
Flog.writeline
    
    
'-------------------------------------------------------------------------------------------------
'Configuracion del Modelo
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando Modelo Interface."
StrSql = "SELECT * FROM modelo WHERE modnro = 315"
OpenRecordset StrSql, rs_Consult
Sep = ""
If Not rs_Consult.EOF Then
    'directorio = Trim(rs_Consult!modarchdefault)
    Sep = IIf(Not IsNull(rs_Consult!modseparador), rs_Consult!modseparador, "")
    SeparadorDecimal = IIf(Not IsNull(rs_Consult!modsepdec), rs_Consult!modsepdec, ".")
    DescripcionModelo = rs_Consult!moddesc
    
    Flog.writeline Espacios(Tabulador * 1) & "Modelo 315 " & rs_Consult!moddesc
    Flog.writeline Espacios(Tabulador * 1) & "Directorio de Exportacion : " & directorio
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el modelo 315."
    HuboError = True
    Exit Sub
End If
Flog.writeline
     
'-----------------------------------------------------------------------------------------------
Set fs = CreateObject("Scripting.FileSystemObject")

If (Not fs.FolderExists(directorio & "PorUsr")) Then
     Set carpeta = fs.CreateFolder(directorio & "PorUsr")
End If
 
If (Not fs.FolderExists(directorio & "PorUsr\" & iduser)) Then
     Set carpeta = fs.CreateFolder(directorio & "PorUsr\" & iduser)
End If

If (Not fs.FolderExists(directorio & "PorUsr\" & iduser & Trim(rs_Consult!modarchdefault))) Then
     Set carpeta = fs.CreateFolder(directorio & "PorUsr\" & iduser & Trim(rs_Consult!modarchdefault))
End If
         
directorio = directorio & "PorUsr\" & iduser & Trim(rs_Consult!modarchdefault)

'Activo el manejador de errores
On Error Resume Next

'archivo
Archivo = directorio & "\ExpRHProLiq_" & Format(Mes & Anio, "000000") & "_" & bpronro & ".txt"

'Set fs = CreateObject("Scripting.FileSystemObject")
Set fExport = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
     Flog.writeline "La carpeta Destino no existe. Se creará."
     Set carpeta = fs.CreateFolder(directorio)
     Set fExport = fs.CreateTextFile(Archivo, True)
End If
On Error GoTo E_ExportLiq
Flog.writeline Espacios(Tabulador * 0) & "Archivo Creado: " & "Archivo"

'----------------------------------------------------------------------------------------------
        
        
'-------------------------------------------------------------------------------------------------
'Creacion del archivo
'-------------------------------------------------------------------------------------------------
'Seteo el nombre del archivo generado
'Archivo = directorio & "\ExpRHProLiq_" & Format(Mes & Anio, "000000") & "_" & bpronro & ".txt"
'Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
'On Error Resume Next
'Set fExport = fs.CreateTextFile(Archivo, True)
'If Err.Number <> 0 Then
'    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
'    Set carpeta = fs.CreateFolder(directorio)
'    Set fExport = fs.CreateTextFile(Archivo, True)
'End If
'On Error GoTo E_ExportLiq
'Flog.writeline Espacios(Tabulador * 0) & "Archivo Creado: " & "ExpRHProLiq_" & Format(Mes & Anio, "000000") & "_" & bpronro & ".txt"

'-------------------------------------------------------------------------------------------------
'Asigno las SQLs a un arreglo y las ejecuto para saber cuantos registros hay q exp p el avance
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "ARMADO DE SQLS PARA EXPORTAR DATOS"
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------------------"
CantReg = 0
Call cargarSQLs(Periodo, Mes, Anio, FecDesde, FecHasta, CantReg)

'-------------------------------------------------------------------------------------------------
'Seteo de las variables de progreso
'-------------------------------------------------------------------------------------------------
Flog.writeline
Progreso = 0
CEmpleadosAProc = CantReg
Flog.writeline
If CEmpleadosAProc = 0 Then
    Flog.writeline Espacios(Tabulador * 0) & "No hay Datos a Exportar"
    CEmpleadosAProc = 1
Else
    Flog.writeline Espacios(Tabulador * 0) & "Cantidad de Registros a Exportar: " & CEmpleadosAProc
End If
IncPorc = (100 / CEmpleadosAProc)
Flog.writeline
        

Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "INICIO DE ESCRITURA EN ARCHIVO"
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------------------"
'-------------------------------------------------------------------------------------------------
'Empleados para los mapeos (CODIGO 1)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Empleados"

OpenRecordset ArrSQL(1), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "1" & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!empleg)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & " , bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & " ' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline
        
        
'-------------------------------------------------------------------------------------------------
'Tabla Tipo de Conceptos (CODIGO 2)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla Tipo de Conceptos"

OpenRecordset ArrSQL(2), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "2" & Sep & CtrlNulo(rs_Consult!tconnro) & Sep & CtrlNulo(rs_Consult!tcondesc) & Sep & CtrlNulo(rs_Consult!sistema)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & " , bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & " ' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla Conceptos (CODIGO 3)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla Conceptos"

OpenRecordset ArrSQL(3), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "3" & Sep & CtrlNulo(rs_Consult!ConcNro) & Sep & CtrlNulo(rs_Consult!ConcCod) & Sep & CtrlNulo(rs_Consult!concabr) & Sep & CtrlNulo(rs_Consult!concorden) & Sep & CtrlNulo(rs_Consult!tconnro)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!concext) & Sep & CtrlNulo(rs_Consult!concvalid) & Sep & CtrlNulo(rs_Consult!concdesde) & Sep & CtrlNulo(rs_Consult!conchasta) & Sep & CtrlNulo(rs_Consult!concrepet)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!concretro) & Sep & CtrlNulo(rs_Consult!concniv) & Sep & CtrlNulo(rs_Consult!fornro) & Sep & CtrlNulo(rs_Consult!concimp) & Sep & CtrlNulo(rs_Consult!codseguridad)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!concusado) & Sep & CtrlNulo(rs_Consult!concpuente) & Sep & CtrlNulo(rs_Consult!Empnro) & Sep & CtrlNulo(rs_Consult!Conccantdec) & Sep & CtrlNulo(rs_Consult!Conctexto)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!concautor) & Sep & CtrlNulo(rs_Consult!concfecmodi) & Sep & CtrlNulo(rs_Consult!Concajuste) & Sep & CtrlNulo(rs_Consult!concapertura)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla Tipo Acumulador (CODIGO 4)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla Tipo Acumulador"

OpenRecordset ArrSQL(4), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "4" & Sep & CtrlNulo(rs_Consult!tacunro) & Sep & CtrlNulo(rs_Consult!tacudesc) & Sep & CtrlNulo(rs_Consult!sistema) & Sep & CtrlNulo(rs_Consult!tacudepu)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla Acumulador (CODIGO 5)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla Acumulador"

OpenRecordset ArrSQL(5), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "5" & Sep & CtrlNulo(rs_Consult!acuNro) & Sep & CtrlNulo(rs_Consult!acudesabr) & Sep & CtrlNulo(rs_Consult!acusist) & Sep & CtrlNulo(rs_Consult!acudesext) & Sep & CtrlNulo(rs_Consult!acumes)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!acutopea) & Sep & CtrlNulo(rs_Consult!acudesborde) & Sep & CtrlNulo(rs_Consult!acurecalculo) & Sep & CtrlNulo(rs_Consult!acuimponible) & Sep & CtrlNulo(rs_Consult!acuimpcont)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!acusel1) & Sep & CtrlNulo(rs_Consult!acusel2) & Sep & CtrlNulo(rs_Consult!acusel3) & Sep & CtrlNulo(rs_Consult!acuppag) & Sep & CtrlNulo(rs_Consult!acudepu)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!acuhist) & Sep & CtrlNulo(rs_Consult!acumanual) & Sep & CtrlNulo(rs_Consult!acuimpri) & Sep & CtrlNulo(rs_Consult!tacunro) & Sep & CtrlNulo(rs_Consult!Empnro)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!acuretro) & Sep & CtrlNulo(rs_Consult!acuorden) & Sep & CtrlNulo(rs_Consult!acunoneg) & Sep & CtrlNulo(rs_Consult!acuapertura)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla Item (CODIGO 6)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla Items"

OpenRecordset ArrSQL(6), rs_Consult

cant = 0
Do While Not rs_Consult.EOF


    Linea = "6" & Sep & CtrlNulo(rs_Consult!itenro) & Sep & CtrlNulo(rs_Consult!itenom) & Sep & CtrlNulo(rs_Consult!itesigno) & Sep & CtrlNulo(rs_Consult!iterenglon) & Sep & CtrlNulo(rs_Consult!itetipotope)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!iteporctope) & Sep & CtrlNulo(rs_Consult!iteitemstope) & Sep & CtrlNulo(rs_Consult!iteprorr)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla Tipo Proceso (CODIGO 7)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla Tipo Proceso"

OpenRecordset ArrSQL(7), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "7" & Sep & CtrlNulo(rs_Consult!tprocnro) & Sep & CtrlNulo(rs_Consult!tprocdesc) & Sep & CtrlNulo(rs_Consult!Empnro) & Sep & CtrlNulo(rs_Consult!tliqnro) & Sep & CtrlNulo(rs_Consult!final) & Sep & CtrlNulo(rs_Consult!ajugcias) & Sep & CtrlNulo(rs_Consult!tprocrecalculo)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla Proceso (CODIGO 8)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla Proceso"

OpenRecordset ArrSQL(8), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "8" & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!prodesc) & Sep & CtrlNulo(rs_Consult!propend) & Sep & CtrlNulo(rs_Consult!profeccorr) & Sep & CtrlNulo(rs_Consult!profecplan)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!pliqnro) & Sep & CtrlNulo(rs_Consult!tprocnro) & Sep & CtrlNulo(rs_Consult!prosist) & Sep & CtrlNulo(rs_Consult!profecpago) & Sep & CtrlNulo(rs_Consult!profecini)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!profecfin) & Sep & CtrlNulo(rs_Consult!Empnro) & Sep & CtrlNulo(rs_Consult!proaprob) & Sep & CtrlNulo(rs_Consult!proestdesc)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla Periodo (CODIGO 9)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla Periodo"

OpenRecordset ArrSQL(9), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "9" & Sep & CtrlNulo(rs_Consult!pliqnro) & Sep & CtrlNulo(rs_Consult!pliqdesc) & Sep & CtrlNulo(rs_Consult!pliqdesde) & Sep & CtrlNulo(rs_Consult!pliqhasta) & Sep & CtrlNulo(rs_Consult!pliqmes)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!pliqbackup) & Sep & CtrlNulo(rs_Consult!pliqdepurado) & Sep & CtrlNulo(rs_Consult!pliqbco) & Sep & CtrlNulo(rs_Consult!pliqsuc) & Sep & CtrlNulo(rs_Consult!pliqabierto)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!pliqultimo) & Sep & CtrlNulo(rs_Consult!pliqfecdep) & Sep & CtrlNulo(rs_Consult!pliqanio) & Sep & CtrlNulo(rs_Consult!pliqsist) & Sep & CtrlNulo(rs_Consult!pliqdepant)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!pliqtexto) & Sep & CtrlNulo(rs_Consult!Empnro)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla cabliq (CODIGO 10)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla cabliq"

OpenRecordset ArrSQL(10), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "10" & Sep & CtrlNulo(rs_Consult!cliqnro) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!Empleado) & Sep & CtrlNulo(rs_Consult!ppagnro) & Sep & CtrlNulo(rs_Consult!cliqtexto)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!cliqdesde) & Sep & CtrlNulo(rs_Consult!cliqhasta) & Sep & CtrlNulo(rs_Consult!nrorecibo) & Sep & CtrlNulo(rs_Consult!cliqnrocorr) & Sep & CtrlNulo(rs_Consult!nroimp)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!fechaimp) & Sep & CtrlNulo(rs_Consult!entregado) & Sep & CtrlNulo(rs_Consult!fechaentrega)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla detliq (CODIGO 11)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla detliq"

OpenRecordset ArrSQL(11), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "11" & Sep & CtrlNulo(rs_Consult!ConcNro) & Sep & CtrlNulo(rs_Consult!dlimonto) & Sep & CtrlNulo(rs_Consult!dlifec) & Sep & CtrlNulo(rs_Consult!cliqnro) & Sep & CtrlNulo(rs_Consult!dlicant)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!dlimonto_base) & Sep & CtrlNulo(rs_Consult!dliporcent) & Sep & CtrlNulo(rs_Consult!dlitexto) & Sep & CtrlNulo(rs_Consult!fornro) & Sep & CtrlNulo(rs_Consult!tconnro)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!dliretro) & Sep & CtrlNulo(rs_Consult!ajustado) & Sep & CtrlNulo(rs_Consult!dliqdesde) & Sep & CtrlNulo(rs_Consult!dliqhasta)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla acu_liq (CODIGO 12)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla acu_liq"

OpenRecordset ArrSQL(12), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "12" & Sep & CtrlNulo(rs_Consult!acuNro) & Sep & CtrlNulo(rs_Consult!cliqnro) & Sep & CtrlNulo(rs_Consult!almonto) & Sep & CtrlNulo(rs_Consult!alcant) & Sep & CtrlNulo(rs_Consult!alfecret)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!almontoreal)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla acu_mes (CODIGO 13)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla acu_mes"

OpenRecordset ArrSQL(13), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "13" & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!acuNro) & Sep & CtrlNulo(rs_Consult!amanio) & Sep & CtrlNulo(rs_Consult!ammonto) & Sep & CtrlNulo(rs_Consult!amcant)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!ammes) & Sep & CtrlNulo(rs_Consult!ammontoreal)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla desliq (CODIGO 14)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla desliq"

OpenRecordset ArrSQL(14), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "14" & Sep & CtrlNulo(rs_Consult!itenro) & Sep & CtrlNulo(rs_Consult!Empleado) & Sep & CtrlNulo(rs_Consult!dlfecha) & Sep & CtrlNulo(rs_Consult!pronro)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!dlmonto) & Sep & CtrlNulo(rs_Consult!dlprorratea)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla ficharet (CODIGO 15)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla ficharet"

OpenRecordset ArrSQL(15), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "15" & Sep & CtrlNulo(rs_Consult!Fecha) & Sep & CtrlNulo(rs_Consult!importe) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!liqsistema)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!Empleado)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla impproarg (CODIGO 16)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla impproarg"

OpenRecordset ArrSQL(16), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "16" & Sep & CtrlNulo(rs_Consult!acuNro) & Sep & CtrlNulo(rs_Consult!cliqnro) & Sep & CtrlNulo(rs_Consult!tconnro) & Sep & CtrlNulo(rs_Consult!ipacant) & Sep & CtrlNulo(rs_Consult!ipamonto)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla impmesarg (CODIGO 17)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla impmesarg"

OpenRecordset ArrSQL(17), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "17" & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!acuNro) & Sep & CtrlNulo(rs_Consult!tconnro) & Sep & CtrlNulo(rs_Consult!imaanio) & Sep & CtrlNulo(rs_Consult!imames)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!imacant) & Sep & CtrlNulo(rs_Consult!imamonto)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla traza_gan (CODIGO 18)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla traza_gan"

OpenRecordset ArrSQL(18), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "18" & Sep & CtrlNulo(rs_Consult!pliqnro) & Sep & CtrlNulo(rs_Consult!ConcNro) & Sep & CtrlNulo(rs_Consult!empresa) & Sep & CtrlNulo(rs_Consult!fecha_pago) & Sep & CtrlNulo(rs_Consult!Ternro)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!msr) & Sep & CtrlNulo(rs_Consult!nomsr) & Sep & CtrlNulo(rs_Consult!nogan) & Sep & CtrlNulo(rs_Consult!jubilacion) & Sep & CtrlNulo(rs_Consult!osocial)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!cuota_medico) & Sep & CtrlNulo(rs_Consult!prima_seguro) & Sep & CtrlNulo(rs_Consult!sepelio) & Sep & CtrlNulo(rs_Consult!estimados)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!otras) & Sep & CtrlNulo(rs_Consult!donacion) & Sep & CtrlNulo(rs_Consult!dedesp) & Sep & CtrlNulo(rs_Consult!noimpo) & Sep & CtrlNulo(rs_Consult!car_flia)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!conyuge) & Sep & CtrlNulo(rs_Consult!hijo) & Sep & CtrlNulo(rs_Consult!otras_cargas) & Sep & CtrlNulo(rs_Consult!retenciones) & Sep & CtrlNulo(rs_Consult!promo)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!saldo) & Sep & CtrlNulo(rs_Consult!sindicato) & Sep & CtrlNulo(rs_Consult!ret_mes) & Sep & CtrlNulo(rs_Consult!mon_conyuge) & Sep & CtrlNulo(rs_Consult!mon_hijo)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!mon_otras) & Sep & CtrlNulo(rs_Consult!viaticos) & Sep & CtrlNulo(rs_Consult!amortizacion) & Sep & CtrlNulo(rs_Consult!entidad1) & Sep & CtrlNulo(rs_Consult!entidad2)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!entidad3) & Sep & CtrlNulo(rs_Consult!entidad4) & Sep & CtrlNulo(rs_Consult!entidad5) & Sep & CtrlNulo(rs_Consult!entidad6) & Sep & CtrlNulo(rs_Consult!entidad7)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!entidad8) & Sep & CtrlNulo(rs_Consult!entidad9) & Sep & CtrlNulo(rs_Consult!entidad10) & Sep & CtrlNulo(rs_Consult!entidad11) & Sep & CtrlNulo(rs_Consult!entidad12)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!entidad13) & Sep & CtrlNulo(rs_Consult!entidad14) & Sep & CtrlNulo(rs_Consult!cuit_entidad1) & Sep & CtrlNulo(rs_Consult!cuit_entidad2)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!cuit_entidad3) & Sep & CtrlNulo(rs_Consult!cuit_entidad4) & Sep & CtrlNulo(rs_Consult!cuit_entidad5) & Sep & CtrlNulo(rs_Consult!cuit_entidad6)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!cuit_entidad7) & Sep & CtrlNulo(rs_Consult!cuit_entidad8) & Sep & CtrlNulo(rs_Consult!cuit_entidad9) & Sep & CtrlNulo(rs_Consult!cuit_entidad10)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!cuit_entidad11) & Sep & CtrlNulo(rs_Consult!cuit_entidad12) & Sep & CtrlNulo(rs_Consult!cuit_entidad13) & Sep & CtrlNulo(rs_Consult!cuit_entidad14)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!monto_entidad1) & Sep & CtrlNulo(rs_Consult!monto_entidad2) & Sep & CtrlNulo(rs_Consult!monto_entidad3) & Sep & CtrlNulo(rs_Consult!monto_entidad4)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!monto_entidad5) & Sep & CtrlNulo(rs_Consult!monto_entidad6) & Sep & CtrlNulo(rs_Consult!monto_entidad7) & Sep & CtrlNulo(rs_Consult!monto_entidad8)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!monto_entidad9) & Sep & CtrlNulo(rs_Consult!monto_entidad10) & Sep & CtrlNulo(rs_Consult!monto_entidad11) & Sep & CtrlNulo(rs_Consult!monto_entidad12)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!monto_entidad13) & Sep & CtrlNulo(rs_Consult!monto_entidad14) & Sep & CtrlNulo(rs_Consult!ganimpo) & Sep & CtrlNulo(rs_Consult!ganneta)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!total_entidad1) & Sep & CtrlNulo(rs_Consult!total_entidad2) & Sep & CtrlNulo(rs_Consult!total_entidad3) & Sep & CtrlNulo(rs_Consult!total_entidad4)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!total_entidad5) & Sep & CtrlNulo(rs_Consult!total_entidad6) & Sep & CtrlNulo(rs_Consult!total_entidad7) & Sep & CtrlNulo(rs_Consult!total_entidad8)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!total_entidad9) & Sep & CtrlNulo(rs_Consult!total_entidad10) & Sep & CtrlNulo(rs_Consult!total_entidad11) & Sep & CtrlNulo(rs_Consult!total_entidad12)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!total_entidad13) & Sep & CtrlNulo(rs_Consult!total_entidad14) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!imp_deter)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!eme_medicas) & Sep & CtrlNulo(rs_Consult!seguro_optativo) & Sep & CtrlNulo(rs_Consult!seguro_retiro) & Sep & CtrlNulo(rs_Consult!tope_os_priv)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!empleg) & Sep & CtrlNulo(rs_Consult!deducciones) & Sep & CtrlNulo(rs_Consult!art23) & Sep & CtrlNulo(rs_Consult!porcdeduc)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla traza_gan_item_top (CODIGO 19)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla traza_gan_item_top"

OpenRecordset ArrSQL(19), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "19" & Sep & CtrlNulo(rs_Consult!itenro) & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!empresa) & Sep & CtrlNulo(rs_Consult!Monto)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!ddjj) & Sep & CtrlNulo(rs_Consult!old_liq) & Sep & CtrlNulo(rs_Consult!liq) & Sep & CtrlNulo(rs_Consult!prorr)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla desmen (CODIGO 20)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla desmen"

OpenRecordset ArrSQL(20), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "20" & Sep & CtrlNulo(rs_Consult!itenro) & Sep & CtrlNulo(rs_Consult!Empleado) & Sep & CtrlNulo(rs_Consult!desmondec) & Sep & CtrlNulo(rs_Consult!desmenprorra) & Sep & CtrlNulo(rs_Consult!desano)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!desfecdes) & Sep & CtrlNulo(rs_Consult!desfechas) & Sep & CtrlNulo(rs_Consult!descuit) & Sep & CtrlNulo(rs_Consult!desrazsoc) & Sep & CtrlNulo(rs_Consult!pronro)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla batch_proceso (CODIGO 21)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla batch_proceso"

OpenRecordset ArrSQL(21), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "21" & Sep & CtrlNulo(rs_Consult!bpronro) & Sep & CtrlNulo(rs_Consult!btprcnro) & Sep & CtrlNulo(rs_Consult!bprcFecha) & Sep & CtrlNulo(rs_Consult!iduser) & Sep & CtrlNulo(rs_Consult!bprchora)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!bprcempleados) & Sep & CtrlNulo(rs_Consult!bprcfecdesde) & Sep & CtrlNulo(rs_Consult!bprcfechasta) & Sep & CtrlNulo(rs_Consult!bprcestado)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!bprcparam) & Sep & CtrlNulo(rs_Consult!bprcprogreso) & Sep & CtrlNulo(rs_Consult!bprcfecfin) & Sep & CtrlNulo(rs_Consult!bprchorafin)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!bprctiempo) & Sep & CtrlNulo(rs_Consult!Empnro) & Sep & CtrlNulo(rs_Consult!bprcurgente) & Sep & CtrlNulo(rs_Consult!bprcterminar)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!bprcConfirmado) & Sep & CtrlNulo(rs_Consult!bprcfecInicioEj) & Sep & CtrlNulo(rs_Consult!bprcFecFinEj) & Sep & CtrlNulo(rs_Consult!bprcPid)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!bprcHoraInicioEj) & Sep & CtrlNulo(rs_Consult!bprcHoraFinEj) & Sep & CtrlNulo(rs_Consult!bprctipomodelo)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla rep_recibo (CODIGO 22)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla rep_recibo"

OpenRecordset ArrSQL(22), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "22" & Sep & CtrlNulo(rs_Consult!bpronro) & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!Apellido) & Sep & CtrlNulo(rs_Consult!nombre)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!direccion) & Sep & CtrlNulo(rs_Consult!Legajo) & Sep & CtrlNulo(rs_Consult!pliqnro) & Sep & CtrlNulo(rs_Consult!pliqmes)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!pliqanio) & Sep & CtrlNulo(rs_Consult!pliqdepant) & Sep & CtrlNulo(rs_Consult!pliqfecdep) & Sep & CtrlNulo(rs_Consult!pliqbco)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!empfecalta) & Sep & CtrlNulo(rs_Consult!sueldo) & Sep & CtrlNulo(rs_Consult!categoria) & Sep & CtrlNulo(rs_Consult!centrocosto)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!Localidad) & Sep & CtrlNulo(rs_Consult!profecpago) & Sep & CtrlNulo(rs_Consult!formapago) & Sep & CtrlNulo(rs_Consult!empnombre)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!empdire) & Sep & CtrlNulo(rs_Consult!empcuit) & Sep & CtrlNulo(rs_Consult!emplogo) & Sep & CtrlNulo(rs_Consult!emplogoalto)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!Cuil) & Sep & CtrlNulo(rs_Consult!emplogoancho) & Sep & CtrlNulo(rs_Consult!empfirma) & Sep & CtrlNulo(rs_Consult!empfirmaalto)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!empfirmaancho) & Sep & CtrlNulo(rs_Consult!prodesc) & Sep & CtrlNulo(rs_Consult!descripcion) & Sep & CtrlNulo(rs_Consult!puesto)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!tenro1) & Sep & CtrlNulo(rs_Consult!tenro2) & Sep & CtrlNulo(rs_Consult!tenro3) & Sep & CtrlNulo(rs_Consult!estrnro1)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!estrnro2) & Sep & CtrlNulo(rs_Consult!estrnro3) & Sep & CtrlNulo(rs_Consult!Orden) & Sep & CtrlNulo(rs_Consult!Auxchar1)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!Auxchar2) & Sep & CtrlNulo(rs_Consult!Auxchar3) & Sep & CtrlNulo(rs_Consult!Auxchar4) & Sep & CtrlNulo(rs_Consult!Auxchar5)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!auxdeci1) & Sep & CtrlNulo(rs_Consult!auxdeci2) & Sep & CtrlNulo(rs_Consult!auxdeci3) & Sep & CtrlNulo(rs_Consult!auxdeci4)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!auxdeci5)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla rep_recibo_det (CODIGO 23)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla rep_recibo_det"

OpenRecordset ArrSQL(23), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "23" & Sep & CtrlNulo(rs_Consult!bpronro) & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!cliqnro) & Sep & CtrlNulo(rs_Consult!concabr)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!ConcCod) & Sep & CtrlNulo(rs_Consult!ConcNro) & Sep & CtrlNulo(rs_Consult!tconnro) & Sep & CtrlNulo(rs_Consult!concimp)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!dlicant) & Sep & CtrlNulo(rs_Consult!dlimonto) & Sep & CtrlNulo(rs_Consult!conctipo)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla rep_libroley (CODIGO 24)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla rep_libroley"

OpenRecordset ArrSQL(24), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "24" & Sep & CtrlNulo(rs_Consult!bpronro) & Sep & CtrlNulo(rs_Consult!Legajo) & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!Apellido)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!apellido2) & Sep & CtrlNulo(rs_Consult!nombre) & Sep & CtrlNulo(rs_Consult!nombre2) & Sep & CtrlNulo(rs_Consult!empresa)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!emprnro) & Sep & CtrlNulo(rs_Consult!pliqnro) & Sep & CtrlNulo(rs_Consult!fecpago) & Sep & CtrlNulo(rs_Consult!fecalta)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!fecbaja) & Sep & CtrlNulo(rs_Consult!contrato) & Sep & CtrlNulo(rs_Consult!categoria) & Sep & CtrlNulo(rs_Consult!direccion)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!puesto) & Sep & CtrlNulo(rs_Consult!documento) & Sep & CtrlNulo(rs_Consult!fecha_nac) & Sep & CtrlNulo(rs_Consult!est_civil)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!Cuil) & Sep & CtrlNulo(rs_Consult!estado) & Sep & CtrlNulo(rs_Consult!reg_prev) & Sep & CtrlNulo(rs_Consult!lug_trab)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!basico) & Sep & CtrlNulo(rs_Consult!neto) & Sep & CtrlNulo(rs_Consult!msr) & Sep & CtrlNulo(rs_Consult!asi_flia) & Sep & CtrlNulo(rs_Consult!dtos)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!bruto) & Sep & CtrlNulo(rs_Consult!prodesc) & Sep & CtrlNulo(rs_Consult!descripcion) & Sep & CtrlNulo(rs_Consult!pliqdesc)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!pliqmes) & Sep & CtrlNulo(rs_Consult!pliqanio) & Sep & CtrlNulo(rs_Consult!profecpago) & Sep & CtrlNulo(rs_Consult!pliqfecdep)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!pliqbco) & Sep & CtrlNulo(rs_Consult!ultima_pag_impr) & Sep & CtrlNulo(rs_Consult!Auxchar1) & Sep & CtrlNulo(rs_Consult!Auxchar2)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!Auxchar3) & Sep & CtrlNulo(rs_Consult!Auxchar4) & Sep & CtrlNulo(rs_Consult!Auxchar5) & Sep & CtrlNulo(rs_Consult!auxdeci1)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!auxdeci2) & Sep & CtrlNulo(rs_Consult!auxdeci3) & Sep & CtrlNulo(rs_Consult!auxdeci4) & Sep & CtrlNulo(rs_Consult!auxdeci5)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!Orden) & Sep & CtrlNulo(rs_Consult!tedabr1) & Sep & CtrlNulo(rs_Consult!tedabr2) & Sep & CtrlNulo(rs_Consult!tedabr3)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!estrdabr1) & Sep & CtrlNulo(rs_Consult!estrdabr2) & Sep & CtrlNulo(rs_Consult!estrdabr3) & Sep & CtrlNulo(rs_Consult!tipofam)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla rep_libroley_det (CODIGO 25)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla rep_libroley_det"

OpenRecordset ArrSQL(25), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "25" & Sep & CtrlNulo(rs_Consult!bpronro) & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!concabr) & Sep & CtrlNulo(rs_Consult!ConcCod)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!ConcNro) & Sep & CtrlNulo(rs_Consult!concimp) & Sep & CtrlNulo(rs_Consult!dlicant) & Sep & CtrlNulo(rs_Consult!dlimonto)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!conctipo)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla rep_libroley_fam (CODIGO 26)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla rep_libroley_fam"

OpenRecordset ArrSQL(26), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "26" & Sep & CtrlNulo(rs_Consult!bpronro) & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!NroDoc)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!sigladoc) & Sep & CtrlNulo(rs_Consult!ternrofam) & Sep & CtrlNulo(rs_Consult!terape) & Sep & CtrlNulo(rs_Consult!ternom)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!terfecnac) & Sep & CtrlNulo(rs_Consult!tersex) & Sep & CtrlNulo(rs_Consult!famest) & Sep & CtrlNulo(rs_Consult!famtrab)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!faminc) & Sep & CtrlNulo(rs_Consult!famsalario) & Sep & CtrlNulo(rs_Consult!famfecvto) & Sep & CtrlNulo(rs_Consult!famCargaDGI)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!famDGIdesde) & Sep & CtrlNulo(rs_Consult!famDGIhasta) & Sep & CtrlNulo(rs_Consult!famemergencia) & Sep & CtrlNulo(rs_Consult!paredesc)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla rep19 (CODIGO 27)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla rep19"

OpenRecordset ArrSQL(27), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "27" & Sep & CtrlNulo(rs_Consult!repnro) & Sep & CtrlNulo(rs_Consult!pliqnro) & Sep & CtrlNulo(rs_Consult!pronro) & Sep & CtrlNulo(rs_Consult!empresa)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!bpronro) & Sep & CtrlNulo(rs_Consult!Fecha) & Sep & CtrlNulo(rs_Consult!hora) & Sep & CtrlNulo(rs_Consult!iduser)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!empleg) & Sep & CtrlNulo(rs_Consult!msr) & Sep & CtrlNulo(rs_Consult!nomsr) & Sep & CtrlNulo(rs_Consult!nogan)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!jubilacion) & Sep & CtrlNulo(rs_Consult!osocial) & Sep & CtrlNulo(rs_Consult!cuota_medico)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!prima_seguro) & Sep & CtrlNulo(rs_Consult!sepelio) & Sep & CtrlNulo(rs_Consult!estimados) & Sep & CtrlNulo(rs_Consult!otras)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!donacion) & Sep & CtrlNulo(rs_Consult!dedesp) & Sep & CtrlNulo(rs_Consult!noimpo) & Sep & CtrlNulo(rs_Consult!car_flia)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!conyuge) & Sep & CtrlNulo(rs_Consult!hijo) & Sep & CtrlNulo(rs_Consult!otras_cargas) & Sep & CtrlNulo(rs_Consult!retenciones)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!promo) & Sep & CtrlNulo(rs_Consult!saldo) & Sep & CtrlNulo(rs_Consult!Desde) & Sep & CtrlNulo(rs_Consult!Hasta) & Sep & CtrlNulo(rs_Consult!Cuil)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!direccion) & Sep & CtrlNulo(rs_Consult!sindicato) & Sep & CtrlNulo(rs_Consult!ret_mes) & Sep & CtrlNulo(rs_Consult!dir_calle)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!dir_num) & Sep & CtrlNulo(rs_Consult!dir_piso) & Sep & CtrlNulo(rs_Consult!dir_dpto) & Sep & CtrlNulo(rs_Consult!dir_localidad)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!dir_pcia) & Sep & CtrlNulo(rs_Consult!dir_cp) & Sep & CtrlNulo(rs_Consult!cuit) & Sep & CtrlNulo(rs_Consult!monto_letras)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!emp_nombre) & Sep & CtrlNulo(rs_Consult!emp_cuit) & Sep & CtrlNulo(rs_Consult!ano) & Sep & CtrlNulo(rs_Consult!mon_conyuge)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!mon_hijo) & Sep & CtrlNulo(rs_Consult!mon_otras) & Sep & CtrlNulo(rs_Consult!viaticos) & Sep & CtrlNulo(rs_Consult!amortizacion)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!entidad1) & Sep & CtrlNulo(rs_Consult!entidad2) & Sep & CtrlNulo(rs_Consult!entidad3) & Sep & CtrlNulo(rs_Consult!entidad4)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!entidad5) & Sep & CtrlNulo(rs_Consult!entidad6) & Sep & CtrlNulo(rs_Consult!entidad7) & Sep & CtrlNulo(rs_Consult!entidad8)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!entidad9) & Sep & CtrlNulo(rs_Consult!entidad10) & Sep & CtrlNulo(rs_Consult!entidad11) & Sep & CtrlNulo(rs_Consult!entidad12)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!entidad13) & Sep & CtrlNulo(rs_Consult!entidad14) & Sep & CtrlNulo(rs_Consult!cuit_entidad1) & Sep & CtrlNulo(rs_Consult!cuit_entidad2)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!cuit_entidad3) & Sep & CtrlNulo(rs_Consult!cuit_entidad4) & Sep & CtrlNulo(rs_Consult!cuit_entidad5) & Sep & CtrlNulo(rs_Consult!cuit_entidad6)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!cuit_entidad7) & Sep & CtrlNulo(rs_Consult!cuit_entidad8) & Sep & CtrlNulo(rs_Consult!cuit_entidad9) & Sep & CtrlNulo(rs_Consult!cuit_entidad10)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!cuit_entidad11) & Sep & CtrlNulo(rs_Consult!cuit_entidad12) & Sep & CtrlNulo(rs_Consult!cuit_entidad13) & Sep & CtrlNulo(rs_Consult!cuit_entidad14)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!monto_entidad1) & Sep & CtrlNulo(rs_Consult!monto_entidad2) & Sep & CtrlNulo(rs_Consult!monto_entidad3) & Sep & CtrlNulo(rs_Consult!monto_entidad4)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!monto_entidad5) & Sep & CtrlNulo(rs_Consult!monto_entidad6) & Sep & CtrlNulo(rs_Consult!monto_entidad7) & Sep & CtrlNulo(rs_Consult!monto_entidad8)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!monto_entidad9) & Sep & CtrlNulo(rs_Consult!monto_entidad10) & Sep & CtrlNulo(rs_Consult!monto_entidad11) & Sep & CtrlNulo(rs_Consult!monto_entidad12)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!monto_entidad13) & Sep & CtrlNulo(rs_Consult!monto_entidad14) & Sep & CtrlNulo(rs_Consult!ganimpo) & Sep & CtrlNulo(rs_Consult!ganneta)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!total_entidad1) & Sep & CtrlNulo(rs_Consult!total_entidad2) & Sep & CtrlNulo(rs_Consult!total_entidad3) & Sep & CtrlNulo(rs_Consult!total_entidad4)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!total_entidad5) & Sep & CtrlNulo(rs_Consult!total_entidad6) & Sep & CtrlNulo(rs_Consult!total_entidad7) & Sep & CtrlNulo(rs_Consult!total_entidad8)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!total_entidad9) & Sep & CtrlNulo(rs_Consult!total_entidad10) & Sep & CtrlNulo(rs_Consult!total_entidad11) & Sep & CtrlNulo(rs_Consult!total_entidad12)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!total_entidad13) & Sep & CtrlNulo(rs_Consult!total_entidad14) & Sep & CtrlNulo(rs_Consult!imp_deter) & Sep & CtrlNulo(rs_Consult!eme_medicas)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!seguro_optativo) & Sep & CtrlNulo(rs_Consult!seguro_retiro) & Sep & CtrlNulo(rs_Consult!tope_os_priv) & Sep & CtrlNulo(rs_Consult!prorratea)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!suscribe) & Sep & CtrlNulo(rs_Consult!caracter) & Sep & CtrlNulo(rs_Consult!fecha_caracter) & Sep & CtrlNulo(rs_Consult!fecha_devolucion)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!dependencia_dgi) & Sep & CtrlNulo(rs_Consult!anual_final) & Sep & CtrlNulo(rs_Consult!Ternro) & Sep & CtrlNulo(rs_Consult!terape)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!ternom)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla escala_ded (CODIGO 28)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla escala_ded"

OpenRecordset ArrSQL(28), rs_Consult

cant = 0
Do While Not rs_Consult.EOF

    Linea = "28" & Sep & CtrlNulo(rs_Consult!esd_topeinf) & Sep & CtrlNulo(rs_Consult!esd_topesup) & Sep & CtrlNulo(rs_Consult!esd_porcentaje)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla confrep (CODIGO 29)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla confrep"

OpenRecordset ArrSQL(29), rs_Consult

cant = 0
Do While Not rs_Consult.EOF


    Linea = "29" & Sep & CtrlNulo(rs_Consult!repnro) & Sep & CtrlNulo(rs_Consult!confnrocol) & Sep & CtrlNulo(rs_Consult!confetiq)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!conftipo) & Sep & CtrlNulo(rs_Consult!confval) & Sep & CtrlNulo(rs_Consult!Empnro)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!confaccion) & Sep & CtrlNulo(rs_Consult!confval2)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Tabla escala (CODIGO 30)
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Exportando Tabla escala"

OpenRecordset ArrSQL(30), rs_Consult

cant = 0
Do While Not rs_Consult.EOF


StrSql = "SELECT , , , ,"
StrSql = StrSql & " , , , "

    Linea = "30" & Sep & CtrlNulo(rs_Consult!escnro) & Sep & CtrlNulo(rs_Consult!escinf) & Sep & CtrlNulo(rs_Consult!escsup)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!escporexe) & Sep & CtrlNulo(rs_Consult!esccuota) & Sep & CtrlNulo(rs_Consult!escano)
    Linea = Linea & Sep & CtrlNulo(rs_Consult!escmes) & Sep & CtrlNulo(rs_Consult!escfecha)
    fExport.writeline Encriptar(Linea)
    cant = cant + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

Flog.writeline Espacios(Tabulador * 1) & "Cantidad de registros exportados " & cant
Flog.writeline


fExport.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close

Set rs_Consult = Nothing
Set fExport = Nothing
Set fs = Nothing


Exit Sub

E_ExportLiq:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: E_ExportLiq"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub


Public Function Encriptar(ByVal Valor As String)
    
    If encryptAct Then
        Encriptar = Encrypt(Seed, Valor)
    Else
        Encriptar = Valor
    End If
    
End Function


Public Function CtrlNulo(ByVal Valor) As String

    If IsNull(Valor) Then
        CtrlNulo = "NULL"
    Else
        CtrlNulo = Replace(Valor, ",", ".")
    End If
    
End Function


Public Function PrimerDiaMes(ByVal Mes As Integer, ByVal Anio As Long) As Date

    PrimerDiaMes = CDate("01/" & CStr(Mes) & "/" & CStr(Anio))
    
End Function


Public Function UltimoDiaMes(ByVal Mes As Integer, ByVal Anio As Long) As Date
Dim aux As Date
    
    'Armo el primer dia del mes siguiente
    If Mes = 12 Then
        aux = CDate("01/01/" & CStr(Anio + 1))
    Else
        aux = CDate("01/" & CStr(Mes + 1) & "/" & CStr(Anio))
    End If
    
    'Le resto 1
    UltimoDiaMes = DateAdd("d", -1, aux)
    
End Function


Public Sub cargarSQLs(ByVal Periodo As Long, ByVal Mes As Long, ByVal Anio As Long, ByVal FecDesde As Date, ByVal FecHasta As Date, ByRef cantRecord As Long)

Dim rs_Consult As New ADODB.Recordset
Dim listaCabliq As String
Dim listaBatch As String
Dim listaRecibo As String
Dim listaLibro As String

On Error GoTo E_cargarSQLs


    'Tabla empleado
    StrSql = "SELECT empleg, ternro"
    StrSql = StrSql & " FROM empleado"
    StrSql = StrSql & " ORDER BY ternro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(1) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL empleados OK"
    
    'Tabla tipconcep
    StrSql = "SELECT tconnro, tcondesc, sistema"
    StrSql = StrSql & " FROM tipconcep"
    StrSql = StrSql & " ORDER BY tconnro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(2) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL tipconcep OK"
    
    'Tabla concepto
    StrSql = "SELECT concnro, conccod, concabr, concorden, tconnro, concext, concvalid, concdesde, conchasta, concrepet, concretro, concniv, fornro, concimp,"
    StrSql = StrSql & " codseguridad , concusado, concpuente, Empnro, Conccantdec, Conctexto, concautor, concfecmodi, Concajuste, concapertura"
    StrSql = StrSql & " FROM concepto"
    StrSql = StrSql & " ORDER BY concnro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(3) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL concepto OK"
    
    'Tabla tipacum
    StrSql = "SELECT tacunro, tacudesc, sistema, tacudepu"
    StrSql = StrSql & " FROM tipacum"
    StrSql = StrSql & " ORDER BY tacunro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(4) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL tipacum OK"
    
    'Tabla acumulador
    StrSql = "SELECT acunro, acudesabr, acusist, acudesext, acumes, acutopea, acudesborde, acurecalculo, acuimponible, acuimpcont, acusel1, acusel2,"
    StrSql = StrSql & " acusel3, acuppag, acudepu , acuhist, acumanual, acuimpri, tacunro, Empnro, acuretro, acuorden, acunoneg, acuapertura"
    StrSql = StrSql & " FROM acumulador"
    StrSql = StrSql & " ORDER BY acunro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(5) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL acumulador OK"
    
    'Tabla item
    StrSql = "SELECT itenro, itenom, itesigno, iterenglon, itetipotope, iteporctope, iteitemstope, iteprorr"
    StrSql = StrSql & " FROM item"
    StrSql = StrSql & " ORDER BY itenro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(6) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL item OK"
    
    'Tabla tipoproc
    StrSql = "SELECT tprocnro, tprocdesc, empnro, tliqnro, final, ajugcias, tprocrecalculo"
    StrSql = StrSql & " FROM tipoproc"
    StrSql = StrSql & " ORDER BY tprocnro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(7) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL tipconcep OK"
    
    'Tabla proceso
    StrSql = "SELECT pronro, prodesc, propend, profeccorr, profecplan, pliqnro, tprocnro,"
    StrSql = StrSql & " prosist, profecpago, profecini, profecfin, empnro, proaprob, proestdesc"
    StrSql = StrSql & " FROM proceso"
    StrSql = StrSql & " ORDER BY pronro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(8) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL proceso OK"
    
    'Tabla periodo
    StrSql = "SELECT pliqnro, pliqdesc, pliqdesde, pliqhasta, pliqmes, pliqbackup, pliqdepurado, pliqbco, pliqsuc,"
    StrSql = StrSql & " pliqabierto, pliqultimo, pliqfecdep, pliqanio, pliqsist,pliqdepant, pliqtexto, empnro"
    StrSql = StrSql & " FROM periodo"
    StrSql = StrSql & " ORDER BY pliqnro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(9) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL periodo OK"
    
    'Tabla cabliq
    StrSql = "SELECT cabliq.cliqnro, cabliq.pronro, cabliq.empleado, cabliq.ppagnro, cabliq.cliqtexto,"
    StrSql = StrSql & " cabliq.cliqdesde, cabliq.cliqhasta, cabliq.nrorecibo, cabliq.cliqnrocorr, cabliq.nroimp,"
    StrSql = StrSql & " cabliq.fechaimp, cabliq.entregado, cabliq.fechaentrega"
    StrSql = StrSql & " FROM cabliq"
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
    StrSql = StrSql & " AND periodo.pliqnro = " & Periodo
    StrSql = StrSql & " ORDER BY cliqnro"
    OpenRecordset StrSql, rs_Consult
    listaCabliq = "0"
    Do While Not rs_Consult.EOF
        listaCabliq = listaCabliq & "," & rs_Consult!cliqnro
        cantRecord = cantRecord + 1
        rs_Consult.MoveNext
    Loop
    ArrSQL(10) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL cabliq OK"
    
    'Tabla detliq
    StrSql = "SELECT concnro, dlimonto, dlifec, cliqnro, dlicant, dlimonto_base, dliporcent, dlitexto, fornro,"
    StrSql = StrSql & " tconnro, dliretro, ajustado, dliqdesde, dliqhasta"
    StrSql = StrSql & " FROM detliq"
    StrSql = StrSql & " WHERE cliqnro IN (" & listaCabliq & ")"
    StrSql = StrSql & " ORDER BY cliqnro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(11) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL detliq OK"
    
    'Tabla acu_liq
    StrSql = "SELECT acunro, cliqnro, almonto, alcant, alfecret, almontoreal"
    StrSql = StrSql & " FROM acu_liq"
    StrSql = StrSql & " WHERE cliqnro IN (" & listaCabliq & ")"
    StrSql = StrSql & " ORDER BY cliqnro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(12) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL acu_liq OK"
    
    'Tabla acu_mes
    StrSql = "SELECT ternro, acunro, amanio, ammonto, amcant, ammes, ammontoreal"
    StrSql = StrSql & " FROM acu_mes"
    StrSql = StrSql & " WHERE ammes = " & Mes
    StrSql = StrSql & " AND amanio = " & Anio
    StrSql = StrSql & " ORDER BY ternro, acunro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(13) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL acu_mes OK"
    
    'Tabla desliq
    StrSql = "SELECT desliq.itenro, desliq.empleado, desliq.dlfecha, desliq.pronro, desliq.dlmonto, desliq.dlprorratea"
    StrSql = StrSql & " FROM desliq"
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = desliq.pronro"
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
    StrSql = StrSql & " AND periodo.pliqnro = " & Periodo
    StrSql = StrSql & " ORDER BY desliq.pronro ,desliq.empleado, desliq.itenro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(14) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL desliq OK"
    
    'Tabla ficharet
    StrSql = "SELECT ficharet.fecha, ficharet.importe, ficharet.pronro, ficharet.liqsistema, ficharet.empleado"
    StrSql = StrSql & " FROM ficharet"
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = ficharet.pronro"
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
    StrSql = StrSql & " AND periodo.pliqnro = " & Periodo
    StrSql = StrSql & " ORDER BY ficharet.pronro, ficharet.empleado"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(15) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL ficharet OK"
    
    'Tabla impproarg
    StrSql = "SELECT acunro, cliqnro, tconnro, ipacant, ipamonto"
    StrSql = StrSql & " FROM impproarg"
    StrSql = StrSql & " WHERE cliqnro IN (" & listaCabliq & ")"
    StrSql = StrSql & " ORDER BY cliqnro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(16) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL tipconcep OK"
    
    'Tabla impmesarg
    StrSql = "SELECT ternro, acunro, tconnro, imaanio, imames, imacant, imamonto"
    StrSql = StrSql & " FROM impmesarg"
    StrSql = StrSql & " WHERE imames = " & Mes
    StrSql = StrSql & " AND imaanio = " & Anio
    StrSql = StrSql & " ORDER BY ternro, acunro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(17) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL impmesarg OK"
    
    'Tabla traza_gan
    StrSql = "SELECT pliqnro, concnro, empresa, fecha_pago, ternro, msr, nomsr, nogan, jubilacion,"
    StrSql = StrSql & " osocial, cuota_medico, prima_seguro, sepelio, estimados, otras,donacion, dedesp,"
    StrSql = StrSql & " noimpo, car_flia, conyuge, hijo, otras_cargas, retenciones, promo, saldo, sindicato, "
    StrSql = StrSql & " ret_mes, mon_conyuge, mon_hijo, mon_otras, viaticos, amortizacion, entidad1, entidad2,"
    StrSql = StrSql & " entidad3, entidad4, entidad5, entidad6, entidad7, entidad8, entidad9, entidad10, entidad11,"
    StrSql = StrSql & " entidad12, entidad13, entidad14, cuit_entidad1, cuit_entidad2, cuit_entidad3, cuit_entidad4,"
    StrSql = StrSql & " cuit_entidad5, cuit_entidad6, cuit_entidad7, cuit_entidad8, cuit_entidad9, cuit_entidad10,"
    StrSql = StrSql & " cuit_entidad11, cuit_entidad12, cuit_entidad13, cuit_entidad14, monto_entidad1, monto_entidad2,"
    StrSql = StrSql & " monto_entidad3, monto_entidad4, monto_entidad5, monto_entidad6, monto_entidad7, monto_entidad8,"
    StrSql = StrSql & " monto_entidad9, monto_entidad10, monto_entidad11, monto_entidad12, monto_entidad13,"
    StrSql = StrSql & " monto_entidad14, ganimpo, ganneta, total_entidad1, total_entidad2, total_entidad3,"
    StrSql = StrSql & " total_entidad4, total_entidad5, total_entidad6, total_entidad7, total_entidad8, total_entidad9,"
    StrSql = StrSql & " total_entidad10, total_entidad11, total_entidad12, total_entidad13, total_entidad14, pronro,"
    StrSql = StrSql & " imp_deter, eme_medicas, seguro_optativo, seguro_retiro, tope_os_priv, empleg, deducciones, art23,"
    StrSql = StrSql & " porcdeduc"
    StrSql = StrSql & " FROM  traza_gan"
    StrSql = StrSql & " WHERE pliqnro = " & Periodo
    StrSql = StrSql & " ORDER BY ternro, concnro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(18) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL traza_gan OK"
    
    'Tabla traza_gan_item_top
    StrSql = "SELECT  traza_gan_item_top.itenro, traza_gan_item_top.ternro, traza_gan_item_top.pronro, "
    StrSql = StrSql & " traza_gan_item_top.empresa, traza_gan_item_top.monto, traza_gan_item_top.ddjj, "
    StrSql = StrSql & " traza_gan_item_top.old_liq, traza_gan_item_top.liq, traza_gan_item_top.prorr"
    StrSql = StrSql & " FROM traza_gan_item_top"
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = traza_gan_item_top.pronro"
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
    StrSql = StrSql & " AND periodo.pliqnro = " & Periodo
    StrSql = StrSql & " ORDER BY traza_gan_item_top.pronro, traza_gan_item_top.ternro, traza_gan_item_top.itenro"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(19) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL traza_gan_item_top OK"
    
    'Tabla desmen
    StrSql = "SELECT itenro, empleado, desmondec, desmenprorra, desano, desfecdes,"
    StrSql = StrSql & " desfechas, descuit, desrazsoc, pronro"
    StrSql = StrSql & " FROM desmen"
    StrSql = StrSql & " ORDER BY empleado, desano"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(20) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL desmen OK"
    
    'Tabla batch_proceso (solo los de tipo recibo, F649 y libro ley, en las fechas del periodo y procesados)
    StrSql = "SELECT bpronro, btprcnro, bprcfecha, iduser, bprchora, bprcempleados, bprcfecdesde, bprcfechasta, bprcestado,"
    StrSql = StrSql & " bprcparam , bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcurgente, bprcterminar,"
    StrSql = StrSql & " bprcConfirmado, bprcfecInicioEj, bprcFecFinEj, bprcPid, bprcHoraInicioEj, bprcHoraFinEj, bprctipomodelo"
    StrSql = StrSql & " FROM batch_proceso"
    StrSql = StrSql & " WHERE btprcnro IN (26, 45, 50)"
    StrSql = StrSql & " AND bprcfecha >= " & ConvFecha(FecDesde)
    StrSql = StrSql & " AND bprcfecha <= " & ConvFecha(FecHasta)
    'StrSql = StrSql & " AND bprcestado = 'Procesado'"
    StrSql = StrSql & " ORDER BY bpronro"
    OpenRecordset StrSql, rs_Consult
    listaBatch = "0"
    Do While Not rs_Consult.EOF
        listaBatch = listaBatch & "," & rs_Consult!bpronro
        cantRecord = cantRecord + 1
        rs_Consult.MoveNext
    Loop
    ArrSQL(21) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL batch_proceso OK"
    
    'Tabla rep_recibo (solo los relacionados a los batch)
    StrSql = "SELECT bpronro, ternro, pronro, apellido, nombre, direccion, legajo, pliqnro, pliqmes,"
    StrSql = StrSql & " pliqanio, pliqdepant, pliqfecdep, pliqbco, empfecalta, sueldo, categoria, centrocosto,"
    StrSql = StrSql & " localidad, profecpago, formapago, empnombre, empdire, empcuit, emplogo, emplogoalto, cuil,"
    StrSql = StrSql & " emplogoancho, empfirma, empfirmaalto, empfirmaancho, prodesc, descripcion, puesto, tenro1, tenro2,"
    StrSql = StrSql & " tenro3, estrnro1, estrnro2, estrnro3, orden, auxchar1, auxchar2, auxchar3, auxchar4, auxchar5,"
    StrSql = StrSql & " auxdeci1, auxdeci2, auxdeci3, auxdeci4, auxdeci5"
    StrSql = StrSql & " FROM rep_recibo"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBatch & " )"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(22) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL rep_recibo OK"
    
    'Tabla rep_recibo_det (solo los relacionados a los batch)
    StrSql = "SELECT bpronro, ternro, pronro, cliqnro, concabr, conccod, concnro, tconnro, concimp, dlicant,"
    StrSql = StrSql & " dlimonto, conctipo"
    StrSql = StrSql & " FROM rep_recibo_det"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBatch & " )"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(23) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL rep_recibo_det OK"
    
    'Tabla rep_libroley (solo los relacionados a los batch)
    StrSql = "SELECT bpronro, legajo, ternro, pronro, apellido, apellido2, nombre, nombre2, empresa, emprnro,"
    StrSql = StrSql & " pliqnro, fecpago, fecalta, fecbaja, contrato, categoria, direccion, puesto, documento, fecha_nac,"
    StrSql = StrSql & " est_civil, cuil, estado, reg_prev, lug_trab, basico, neto, msr, asi_flia, dtos, bruto, prodesc,"
    StrSql = StrSql & " descripcion, pliqdesc, pliqmes, pliqanio, profecpago, pliqfecdep, pliqbco, ultima_pag_impr, auxchar1,"
    StrSql = StrSql & " auxchar2, auxchar3, auxchar4, auxchar5, auxdeci1, auxdeci2, auxdeci3, auxdeci4, auxdeci5, orden,"
    StrSql = StrSql & " tedabr1, tedabr2, tedabr3, estrdabr1, estrdabr2, estrdabr3, tipofam"
    StrSql = StrSql & " FROM rep_libroley"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBatch & " )"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(24) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL rep_libroley OK"
    
    'Tabla rep_libroley_det (solo los relacionados a los batch)
    StrSql = "SELECT bpronro, ternro, pronro, concabr, conccod, concnro, concimp, dlicant, dlimonto, conctipo"
    StrSql = StrSql & " FROM rep_libroley_det"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBatch & " )"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(25) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL rep_libroley_det OK"
    
    'Tabla rep_libroley_fam (solo los relacionados a los batch)
    StrSql = "SELECT bpronro, ternro, pronro, nrodoc, sigladoc, ternrofam, terape, ternom, terfecnac, tersex,"
    StrSql = StrSql & " famest, famtrab, faminc, famsalario, famfecvto, famCargaDGI, famDGIdesde, famDGIhasta,"
    StrSql = StrSql & " famemergencia, paredesc"
    StrSql = StrSql & " FROM rep_libroley_fam"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBatch & " )"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(26) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL rep_libroley_fam OK"
    
    'Tabla F649 (solo los relacionados a los batch)
    StrSql = "SELECT repnro, pliqnro, pronro, empresa, bpronro, Fecha, Hora, iduser, empleg, msr, nomsr,"
    StrSql = StrSql & " nogan, jubilacion, osocial, cuota_medico, prima_seguro, sepelio, estimados, otras, donacion,"
    StrSql = StrSql & " dedesp, noimpo, car_flia, conyuge, hijo, otras_cargas, retenciones, promo, saldo, desde, hasta,"
    StrSql = StrSql & " cuil, direccion, sindicato, ret_mes, dir_calle, dir_num, dir_piso, dir_dpto, dir_localidad,"
    StrSql = StrSql & " dir_pcia, dir_cp, cuit, monto_letras, emp_nombre, emp_cuit, ano, mon_conyuge, mon_hijo, mon_otras,"
    StrSql = StrSql & " viaticos, amortizacion, entidad1, entidad2, entidad3, entidad4, entidad5, entidad6, entidad7,"
    StrSql = StrSql & " entidad8, entidad9, entidad10, entidad11, entidad12, entidad13, entidad14, cuit_entidad1,"
    StrSql = StrSql & " cuit_entidad2, cuit_entidad3, cuit_entidad4, cuit_entidad5, cuit_entidad6, cuit_entidad7, cuit_entidad8,"
    StrSql = StrSql & " cuit_entidad9, cuit_entidad10, cuit_entidad11, cuit_entidad12, cuit_entidad13, cuit_entidad14,"
    StrSql = StrSql & " monto_entidad1, monto_entidad2, monto_entidad3, monto_entidad4, monto_entidad5, monto_entidad6,"
    StrSql = StrSql & " monto_entidad7, monto_entidad8, monto_entidad9, monto_entidad10, monto_entidad11, monto_entidad12,"
    StrSql = StrSql & " monto_entidad13, monto_entidad14, ganimpo, ganneta, total_entidad1, total_entidad2, total_entidad3, "
    StrSql = StrSql & " total_entidad4, total_entidad5, total_entidad6, total_entidad7, total_entidad8, total_entidad9,"
    StrSql = StrSql & " total_entidad10, total_entidad11, total_entidad12, total_entidad13, total_entidad14, imp_deter,"
    StrSql = StrSql & " eme_medicas, seguro_optativo, seguro_retiro, tope_os_priv, prorratea, suscribe, caracter, fecha_caracter,"
    StrSql = StrSql & " fecha_devolucion, dependencia_dgi, anual_final, ternro, terape, ternom"
    StrSql = StrSql & " FROM rep19"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBatch & " )"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(27) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL rep19"
    
    'Tabla escala_ded
    StrSql = "SELECT esd_topeinf, esd_topesup, esd_porcentaje"
    StrSql = StrSql & " FROM escala_ded"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(28) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL escala_ded OK"
    
    'Tabla confrep
    StrSql = "SELECT repnro, confnrocol, confetiq, conftipo, confval,"
    StrSql = StrSql & " empnro, confaccion, confval2"
    StrSql = StrSql & " FROM confrep"
    StrSql = StrSql & " WHERE repnro = 114"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(29) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL confrep OK"
    
    'Tabla escala
    StrSql = "SELECT escnro, escinf, escsup, escporexe,"
    StrSql = StrSql & " esccuota, escano, escmes, escfecha"
    StrSql = StrSql & " FROM escala"
    OpenRecordset StrSql, rs_Consult
        cantRecord = cantRecord + rs_Consult.RecordCount
    ArrSQL(30) = StrSql
    Flog.writeline Espacios(Tabulador * 1) & "SQL escala OK"

If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing

Exit Sub

E_cargarSQLs:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: cargarSQLs"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline " Error: " & Err.Description
    Flog.writeline "=================================================================="
    HuboError = True

End Sub
