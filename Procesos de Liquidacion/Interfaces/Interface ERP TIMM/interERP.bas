Attribute VB_Name = "interERP"
Option Explicit
 
' ---------------------------------------------------------------------------------------------
' Descripcion: Interfaz que genera el archivo .csv a partir de la tabla temporal
' temporal_ERP, la cual es cargada por el trigger
' Autor      : Manterola Maria Magdalena
' Fecha      : 23/05/2011
' ---------------------------------------------------------------------------------------------
 

'--------------------------------------------------------------------------------------
'Datos del proceso
'--------------------------------------------------------------------------------------
'Global Const NombreProceso = "ERPinterfaz"
'Global Const Ejecutable = "ERPinterfaz.exe"
'Global Const Version = "1.00"

Global Const NombreProceso = "ERPinterfaz"
Global Const Ejecutable = "ERPinterfaz.exe"
Global Const Version = "1.01"   'FGZ - 24/05/2011


'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------

Dim fs, f

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean
Global EmpleadosProcesados As Long
Global CantEmpModificados As Long
Global CantEmpInsertados As Long

Private Sub Main()
Dim strCmdLine As String
Dim Nombre_Arch As String
Dim Directorio

Dim StrSql As String
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Consulta As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim objRs As New ADODB.Recordset


Dim legActual As Long
Dim id As Integer
Dim fechagen As String
Dim sep As String
Dim tipDoc As String
Dim nroDoc As Long
Dim ape As String
Dim nombre As String
Dim emp As String
Dim dep As String
Dim suc As String
Dim pin As String
Dim Ternro
Dim i
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim primero As Boolean
Dim procesar As Boolean

    
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

    
    'NroProceso = NroProcesoBatch
    '------------------------------------------------------------------------
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas 'Modulo mdlDataAccess
    '------------------------------------------------------------------------
 
    
    TiempoInicialProceso = GetTickCount

    Nombre_Arch = PathFLog & "InterfazERP" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
   
    'Obtengo el Process ID
    PID = GetCurrentProcessId 'Modulo MdlGlobal
    Flog.writeline "-----------------------------------------------------------------"
    
    Flog.writeline "Nombre del Proceso = " & NombreProceso
    Flog.writeline "Nombre del Ejecutable = " & Ejecutable
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha = " & Date
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline

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

    HuboErrores = False
    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoInicialProceso = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 302 AND bpronro = " & NroProcesoBatch
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        'No existe el proceso
        Flog.writeline "No existe el proceso"
    Else
        
        'EMPIEZA EL PROCESO
        
        'Archivo de exportacion
        StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Directorio = Trim(rs!sis_dirsalidas)
        End If
        
        'Primero busco donde voy a ubicar el archivo que voy a crear
        StrSql = "SELECT modarchdefault FROM modelo WHERE modnro = 340"
        OpenRecordset StrSql, rs_Modelo
        If Not rs_Modelo.EOF Then
            If Not IsNull(rs_Modelo!modarchdefault) Then
                Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
            Else
                Flog.writeline "El modelo no tiene configurada la carpeta destino."
                GoTo Fin
            End If
        Else
            Flog.writeline "No se encontró el modelo."
            GoTo Fin
        End If
        
        'Ahora tengo que buscar en la tabla temporal los registros a cargar en el archivo
        Progreso = 0
        StrSql = "SELECT * FROM temporal_ERP "
        StrSql = StrSql & " INNER JOIN empleado ON empleado.empleg = temporal_ERP.leg "
        StrSql = StrSql & " ORDER BY fechagen DESC "
        OpenRecordset StrSql, rs_Consulta
        If rs_Consulta.EOF Then
            CEmpleadosAProc = 0
        Else
            CEmpleadosAProc = rs_Consulta.RecordCount
        End If
    
        If CEmpleadosAProc = 0 Then
            Flog.writeline "no hay empleados"
            CEmpleadosAProc = 1
            IncPorc = 100
        Else
            IncPorc = (100 / CEmpleadosAProc)
        End If
        
        primero = True
        procesar = False
        Do Until rs_Consulta.EOF
            If primero Then
                primero = False
                procesar = True
                Ternro = rs_Consulta!Ternro
            Else
                If Ternro <> rs_Consulta!Ternro Then
                    Ternro = rs_Consulta!Ternro
                    procesar = True
                Else
                    procesar = False
                End If
            End If
        
            If procesar Then
                legActual = rs_Consulta!leg
                id = rs_Consulta!id_abm
                fechagen = rs_Consulta!fechagen
                sep = rs_Consulta!separador
                tipDoc = rs_Consulta!tipoDoc
                nroDoc = rs_Consulta!nroDoc
                ape = rs_Consulta!Apellido
                nombre = rs_Consulta!nombre
                emp = IIf(EsNulo(rs_Consulta!empdesc), "", rs_Consulta!empdesc)
                dep = IIf(EsNulo(rs_Consulta!departamento), "", rs_Consulta!departamento)
                suc = IIf(EsNulo(rs_Consulta!sucdesc), "", rs_Consulta!sucdesc)
                pin = IIf(EsNulo(rs_Consulta!pin), "", rs_Consulta!pin)
                
                Call insertarEnCSV(Directorio, sep, id, fechagen, legActual, tipDoc, nroDoc, ape, nombre, emp, dep, suc, pin)
            End If
            
            'Actualizo el progreso
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & ", bprcempleados = bprcempleados - 1 WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
            rs_Consulta.MoveNext
        Loop
        
        'Ahora tengo que eliminar los registros de la tabla temporal
        StrSql = "DELETE FROM temporal_ERP "
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline "Fin."
    Flog.writeline "-----------------------------------------------------------------"
    
    
Fin:
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
    
If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing

If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
Set rs_Modelo = Nothing
End

ME_Main:
    HuboError = True
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
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    GoTo Fin
End Sub

'-------------------------------------------------------------------------------------
' Descripcion: Inserta los datos recuperados de la tabla temporal en un archivo .csv
' Autor      : Manterola Maria Magdalena
' Fecha      : 23/05/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Sub insertarEnCSV(ByVal ubicacion, ByVal sep, ByVal id, ByVal fechagen, ByVal leg, ByVal tipDoc, ByVal nroDoc, ByVal ape, ByVal nombre, ByVal emp, ByVal dep, ByVal suc, ByVal pin)
Dim nombreArch As String
Dim hora As String
Dim fs2
Dim ArchCSV
Dim Linea As String

    
    'Primero armo el nombre del archivo
    
    hora = Str(Hour(fechagen)) + ":" + Str(Minute(fechagen)) + ":" + Str(Second(fechagen))
    
    nombreArch = ubicacion & "\Leg" & leg & " " & Format(fechagen, "DD-mm-YYYY") & " " & Format(hora, "HH-mm-ss") & ".csv"
    
    Flog.writeline "El nombre del Archivo a Generar es: " & nombreArch
    
    Set fs2 = CreateObject("Scripting.FileSystemObject")
    Set ArchCSV = fs2.CreateTextFile(nombreArch, True)
    
    'Comienzo a Escribir en el archivo
    Flog.writeline "Comienzo a Escribir en el archivo"
    Linea = ""
    
    'ID (1 = Alta, 2 = Baja, 3 = Modifica).
    Flog.writeline "Ingreso ID (1 = Alta, 2 = Baja, 3 = Modifica): " & id
    Linea = Linea & id
    'ArchCSV.writeline id
    'ArchCSV.writeline sep
    
    'Fecha De ejecución.
    Flog.writeline "Fecha De Ejecución: " & Format(fechagen, "dd/mm/yyyy")
    Linea = Linea & sep & Format(fechagen, "dd/mm/yyyy")
    'ArchCSV.Write Format(fechagen, "dd/mm/yyyy")
    'ArchCSV.Write sep
    
    'Número de Legajo del Empleado.
    Flog.writeline "Número de Legajo del Empleado: " & leg
    Linea = Linea & sep & leg
    'ArchCSV.Write leg
    'ArchCSV.Write sep
    
    'Tipo de Documento.
    Flog.writeline "Tipo de Documento: " & tipDoc
    Linea = Linea & sep & tipDoc
    'ArchCSV.writeline tipDoc
    'ArchCSV.writeline sep
        
    'Número de Documento.
    Flog.writeline "Número de Documento: " & nroDoc
    Linea = Linea & sep & nroDoc
    'ArchCSV.writeline nroDoc
    'ArchCSV.writeline sep
    
    'Apellido del Empleado.
    Flog.writeline "Apellido del Empleado: " & ape
    Linea = Linea & sep & ape
    'ArchCSV.writeline ape
    'ArchCSV.writeline sep
    
    'Nombre del Empleado.
    Flog.writeline "Nombre del Empleado: " & nombre
    Linea = Linea & sep & nombre
    'ArchCSV.writeline nombre
    'ArchCSV.writeline sep
    
    'Empresa.
    Flog.writeline "Empresa: " & emp
    Linea = Linea & sep & emp
    'ArchCSV.writeline emp
    'ArchCSV.writeline sep
    
    'Departamento.
    Flog.writeline "Departamento: " & dep
    Linea = Linea & sep & dep
    'ArchCSV.writeline dep
    'ArchCSV.writeline sep
    
    'Sucursal.
    Flog.writeline "Sucursal: " & suc
    Linea = Linea & sep & suc
    'ArchCSV.writeline suc
    'ArchCSV.writeline sep
    
    'PIN.
    Flog.writeline "PIN: " & pin
    Linea = Linea & sep & pin
    'ArchCSV.writeline pin
    
    
    ArchCSV.writeline Linea
    ArchCSV.Close
End Sub
