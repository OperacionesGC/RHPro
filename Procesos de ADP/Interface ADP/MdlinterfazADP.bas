Attribute VB_Name = "MdlinterfazADP"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "03/12/2013" ' CAS-20730 - VSO - AGCO - Interfaces

'Const Version = "1.01"
'Const FechaVersion = "06/01/2014" ' CAS-20730 - VSO - AGCO - Interfaces [Entrega 3]
'       LED - Se obtienen los datos del empleado aunque no tenga reporta A.

'Const Version = "1.02"
'Const FechaVersion = "13/01/2014" ' CAS-20730 - VSO - AGCO - Interfaces [Entrega 5]
'       LED - Se obtiene el grado del puesto del empleado de la misma forma que desde la ventana de adp.

'Const Version = "1.03"
'Const FechaVersion = "19/06/2014" ' CAS-25459 - VSO - AGCO - Interfaz AGCO
'       LED - Cambios varios en todos los archivos de exportacion, ver excel adjunto a la tarea.

'Const Version = "1.04"
'Const FechaVersion = "03/07/2014" ' CAS-25459 - VSO - AGCO - Interfaz AGCO
'       LED - Cambios en el archivo Persona Fisica, campo 16 y campo 47

Const Version = "1.05"
Const FechaVersion = "29/04/2015" 'Sebastian Stremel - CAS-29209 - AGCO - Modificar nombres de archivos
'       Se cambian los nombres de los archivos se agrega lo siguiente 580_AAAAMMDD_HHmmSS.txt

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
Private Type ConexionEmpresa
    ConexionOracle As String    'Guarda la conexion de oracle para la empresa
    ConexionAcces As String     'Guarda la conexion de Acces para la empresa
    estrnroEmpresa As String    'Codigo de la estrucutra empresa configurada
End Type




Public Sub Main()

Dim objconnMain As New ADODB.Connection
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

    Nombre_Arch = PathFLog & "InterfazADP_" & NroProcesoBatch & ".log"
    
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
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Acutaliza el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 411 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = IIf(EsNulo(rs_batch_proceso!bprcparam), "", rs_batch_proceso!bprcparam)
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call interfazADP(NroProcesoBatch, bprcparam)
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


Public Sub interfazADP(ByVal bpronro As Long, ByVal Parametros As String)
 
 Dim directorio As String
 Dim Nombre_Arch As String
 Dim rs_consult As New ADODB.Recordset
 Dim rsEmpleados As New ADODB.Recordset
 Dim separador As String
 Dim separadorDecimal As String
 Dim strLineaExp As String
 Dim archSalida1    'Persona Fisica (ARPfisica.txt)
 Dim archSalida2    'Registro de Empleo (ARRegEmp.txt)
 Dim archSalida3    'Historico de Centro de Costos (ARHlotac.txt)
 Dim archSalida4    'Historico de Funciones (AR Hfuncao.txt)
 Dim archSalida5    'Historico Salarial (AR HSal.txt)
 Dim archSalida6    'Planilla de Licencias General (ARHafastg txt)
 Dim porc As Double
 Dim cantEmpleados As Integer
 Dim arrayArchivos
 Dim indice As Long
 Dim usaencabezado As String
 Dim teestablecimiento As String
 Dim tcodnro As String
    
    On Error GoTo CE
    
    arrayArchivos = Split(Parametros, "@")
    
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "Comienza la exportacion "
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------"
    Progreso = 0
    

    StrSql = " SELECT distinct e.ternro FROM empleado e ORDER BY e.ternro ASC "
    OpenRecordset StrSql, rsEmpleados
            
    cantEmpleados = rsEmpleados.RecordCount
    
    'porc = CLng(50 / cantEmpleados)
    If cantEmpleados = 0 Then
        Progreso = 100
        cantEmpleados = 1
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron empleados."
    End If
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline "No esta configurado el directorio de sistema."
        Exit Sub
    End If
    
    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "SELECT * FROM modelo WHERE modnro = 390 " '& NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = directorio & Trim(objRs!modarchdefault)
                
       separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
       separadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, "")
       usaencabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
       Flog.writeline "Directorio de exportacion: " & directorio
    Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Sub
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    'controlamos los archivos que se deben crear
    
    'obtengo el tipo de estructura configurado para obtener el dato establecimiento que es comun para todos los archivos
    StrSql = "SELECT confnrocol, confval FROM confrep WHERE repnro = 419 AND confnrocol in (4,7) "
    OpenRecordset StrSql, objRs
    teestablecimiento = 0
    tcodnro = 0
    teestablecimiento = 0
    Do While Not objRs.EOF
        Select Case objRs!confnrocol
            Case 4
                tcodnro = objRs!confval
            Case 7
                teestablecimiento = objRs!confval
        End Select
        objRs.MoveNext
        
    Loop
    
    If CStr(teestablecimiento) = "0" Then
        Flog.writeline "No se configuro la columna 7, tipo de estructura para informar el campo establecimiento de todos los archivos."
    End If
    
    If CStr(tcodnro) = "0" Then
        Flog.writeline "No se configuro la columna 9, tipo de codigo asociado a las estructuras para informar codigos asocidas a ellas."
    End If
    
    'Call borrarArchivo(directorio & "\ARPfisica.txt")
    'Call borrarArchivo(directorio & "\ARRegEmp.txt")
    'Call borrarArchivo(directorio & "\ARHlotac.txt")
    'Call borrarArchivo(directorio & "\AR Hfuncao.txt")
    'Call borrarArchivo(directorio & "\AR HSal.txt")
    'Call borrarArchivo(directorio & "\ARHafastg.txt")
    
    
    For indice = 0 To UBound(arrayArchivos)
        Select Case indice
            Case 0
                If arrayArchivos(indice) = -1 Then
                    'archSalida1 = directorio & "\ARPfisica.txt"
                    archSalida1 = directorio & "\AR_Pfisica_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                    Set archSalida1 = fs.CreateTextFile(archSalida1, True)
                Else
                    Flog.writeline "No se generara el archivo:  AR_Pfisica_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                End If
            Case 1
                If arrayArchivos(indice) = -1 Then
                    'archSalida2 = directorio & "\ARRegEmp.txt"
                    archSalida2 = directorio & "\AR_RegEmp_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                    Set archSalida2 = fs.CreateTextFile(archSalida2, True)
                Else
                    Flog.writeline "No se generara el archivo: AR_RegEmp_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                End If
                 
            Case 2
                If arrayArchivos(indice) = -1 Then
                    'archSalida3 = directorio & "\ARHlotac.txt"
                    archSalida3 = directorio & "\AR_Hlotac_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                    Set archSalida3 = fs.CreateTextFile(archSalida3, True)
                Else
                    Flog.writeline "No se generara el archivo: AR_Hlotac_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                End If
                
            Case 3
                If arrayArchivos(indice) = -1 Then
                    'archSalida4 = directorio & "\AR Hfuncao.txt"
                    archSalida4 = directorio & "\AR_Hfuncao_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                    Set archSalida4 = fs.CreateTextFile(archSalida4, True)
                Else
                    Flog.writeline "No se generara el archivo: AR_Hfuncao_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                End If
            
            Case 4
                If arrayArchivos(indice) = -1 Then
                    'archSalida5 = directorio & "\AR HSal.txt"
                    archSalida5 = directorio & "\AR_HSal_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                    Set archSalida5 = fs.CreateTextFile(archSalida5, True)
                Else
                    Flog.writeline "No se generara el archivo: AR_HSal_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                End If
            Case 5
                If arrayArchivos(indice) = -1 Then
                    'archSalida6 = directorio & "\ARHafastg.txt"
                    archSalida6 = directorio & "\AR_Hafastg_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                    Set archSalida6 = fs.CreateTextFile(archSalida6, True)
                Else
                    Flog.writeline "No se generara el archivo: AR_Hafastg_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                End If
        
        End Select
    Next
    
    porc = 100 / CLng(cantEmpleados)
    
    Do While Not rsEmpleados.EOF
        Progreso = Progreso + porc
        
        For indice = 0 To UBound(arrayArchivos)
            Select Case indice
                Case 0
                    If arrayArchivos(indice) = -1 Then
                        Call PersonaFísica(rsEmpleados!Ternro, archSalida1, separador, teestablecimiento, tcodnro)
                    End If
                Case 1
                    If arrayArchivos(indice) = -1 Then
                        Call registroEmpleo(rsEmpleados!Ternro, archSalida2, separador, teestablecimiento, tcodnro)
                    End If
                Case 2
                    'Historico de Centro de Costos
                    If arrayArchivos(indice) = -1 Then
                        Call historicoEstructura(rsEmpleados!Ternro, archSalida3, "5", separador, teestablecimiento, tcodnro)
                    End If
                Case 3
                    'Historico de Puestos
                    If arrayArchivos(indice) = -1 Then
                        Call historicoEstructura(rsEmpleados!Ternro, archSalida4, "4", separador, teestablecimiento, tcodnro)
                    End If
                Case 4
                    If arrayArchivos(indice) = -1 Then
                        Call historicoSalarial(rsEmpleados!Ternro, archSalida5, separador, separadorDecimal, teestablecimiento, tcodnro)
                    End If
                Case 5
                    If arrayArchivos(indice) = -1 Then
                        Call historicoLicencias(rsEmpleados!Ternro, archSalida6, separador, teestablecimiento, tcodnro)
                    End If
            End Select
        Next
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        

        rsEmpleados.MoveNext
    Loop
    
    Progreso = 100
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
  
    If rs_consult.State = adStateOpen Then rs_consult.Close
    If rsEmpleados.State = adStateOpen Then rsEmpleados.Close
    
GoTo Procesado
CE:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
Procesado:
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Los datos fueron Exportados Exitosamente."
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
End Sub




Public Function cambiaFecha(ByVal fecha As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea la fecha al formato de insercion de la base de datos.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


    If EsNulo(fecha) Then
        cambiaFecha = "NULL"
    Else
        cambiaFecha = ConvFecha(fecha)
    End If

End Function


'Obtiene el directorio configurado para el modelo
Public Function PathModelo(nroModelo)
 Dim directorio As String
 Dim rsAux As New ADODB.Recordset
 
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        directorio = Trim(rsAux!sis_direntradas)
    Else
        Flog.writeline "No esta configurado el directorio de sistema."
        Exit Function
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro= " & nroModelo
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        directorio = directorio & Trim(rsAux!modarchdefault)
        Flog.writeline "Directorio del modelo: " & directorio
     Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Function
    End If
        
    PathModelo = directorio
End Function


Public Function SeparadorModelo(nroModelo)
 Dim separador As String
 Dim rsAux As New ADODB.Recordset

    StrSql = "SELECT modseparador FROM modelo WHERE modnro= " & nroModelo
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        separador = Trim(rsAux!modseparador)
        Flog.writeline "Separador del modelo: " & separador
     Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Function
    End If
        
    SeparadorModelo = separador
End Function



Sub PersonaFísica(ByVal Ternro As String, ByRef archSalida, ByVal separador As String, ByVal teestablecimiento As String, ByVal tcodnro As String)
Dim strlinea As String
Dim rs_consult As New ADODB.Recordset
Dim email As String
Dim legajo As String
Dim nroIdentif As String
Dim fechaNacimiento As String
Dim nombre As String
Dim nrodni As String
Dim pasaporte As String
Dim barrio As String
Dim nroDomicilio As String
Dim sexo As String
    
    Flog.writeline "Entrando a importar el tercero: " & Ternro & ". Archivo: AR_Pfisica_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
    
    email = ""
    fechaNacimiento = ""
    nrodni = ""
    pasaporte = ""
    'busco datos basico del empleado
    StrSql = " SELECT empleado.terape, empleado.ternom, empleado.ternom2, empleado.terape, empleado.terape2, empleg, empemail, terfecnac " & _
             " , pas.nrodoc pasaporte, dni.nrodoc nrodni, tersex, detdom.barrio, cabdom.domnro, detdom.nro " & _
             " FROM empleado " & _
             " INNER JOIN tercero ON tercero.ternro = empleado.ternro " & _
             " LEFT JOIN cabdom ON cabdom.ternro = empleado.ternro AND cabdom.domdefault = -1 " & _
             " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
             " LEFT JOIN ter_doc pas ON pas.ternro = empleado.ternro AND pas.tidnro = 5 " & _
             " LEFT JOIN ter_doc dni ON dni.ternro = empleado.ternro AND dni.tidnro = 1 " & _
             " WHERE empleado.ternro = " & Ternro
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        If Not IsNull(rs_consult!empemail) Then
            email = rs_consult!empemail
            Flog.writeline "Mail obtenido para el tercero: " & Ternro
        End If
        legajo = CLng(88000000) + CLng(rs_consult!empleg)
        nroIdentif = CLng(26000000) + CLng(rs_consult!empleg)
        fechaNacimiento = rs_consult!terfecnac
        
        nombre = rs_consult!terape & ", " & rs_consult!ternom
        
        '----------------------------------------------------------------
        barrio = ""
        nroDomicilio = ""
        If Not IsNull(rs_consult!barrio) Then
            barrio = rs_consult!barrio
        End If
        If Not IsNull(rs_consult!nro) Then
            nroDomicilio = rs_consult!nro
        End If
        
        '----------------------------------------------------------------------
        If Not IsNull(rs_consult!pasaporte) Then
            pasaporte = Replace(rs_consult!pasaporte, ".", "")
            Flog.writeline "Pasaporte obtenido para el tercero: " & Ternro
        End If
        If Not IsNull(rs_consult!nrodni) Then
            nrodni = Replace(rs_consult!nrodni, ".", "")
            Flog.writeline "DNI obtenido para el tercero: " & Ternro
        End If
        If CLng(rs_consult!tersex) = -1 Then
            sexo = "YES"
        Else
            sexo = "NO"
        End If
    
        strlinea = "No"                                                         'pos 1
        strlinea = strlinea & separador & Left(barrio, 20)                      'pos 2
        strlinea = strlinea & separador & "No"                                  'pos 3
        Call lineaEnBlanco(strlinea, "", 2, separador)                          'pos 4-5
        strlinea = strlinea & separador & "90000000"                            'pos 6
        Call lineaEnBlanco(strlinea, "", 2, separador)                          'pos 7-8
        Call lineaEnBlanco(strlinea, "null", 1, separador)                      'pos 9
        strlinea = strlinea & separador & "1"                                   'pos 10
        Call lineaEnBlanco(strlinea, "null", 2, separador)                      'pos 11-12
        strlinea = strlinea & separador & fechaUltimaFase(Ternro)               'pos 13
        Call lineaEnBlanco(strlinea, "null", 1, separador)                      'pos 14
        strlinea = strlinea & separador & "54"                                  'pos 15
        strlinea = strlinea & separador & "54"                                  'pos 16
        strlinea = strlinea & separador & "No"                                  'pos 17
        Call lineaEnBlanco(strlinea, "", 1, separador)                          'pos 18
        strlinea = strlinea & separador & Left(email, 60)                       'pos 19
        Call lineaEnBlanco(strlinea, "null", 1, separador)                      'pos 20
        strlinea = strlinea & separador & "100"                                 'pos 21
        Call lineaEnBlanco(strlinea, "", 1, separador)                          'pos 22
        strlinea = strlinea & separador & obtenerCod(Ternro, teestablecimiento, Date, tcodnro, 0, 23, 8)   'pos 23
        strlinea = strlinea & separador & "7"                                   'pos 24
        Call lineaEnBlanco(strlinea, "null", 2, separador)                      'pos 25-26
        strlinea = strlinea & separador & "No"                                  'pos 27
        strlinea = strlinea & separador & "10"                                   'pos 28
        strlinea = strlinea & separador & Left(nrodni, 10)                      'pos 29
        Call lineaEnBlanco(strlinea, "", 2, separador)                          'pos 30-31
        strlinea = strlinea & separador & Left(legajo, 8)                       'pos 32
        Call lineaEnBlanco(strlinea, "", 1, separador)                          'pos 33
        strlinea = strlinea & separador & "YES"                                 'pos 34
        strlinea = strlinea & separador & fechaNacimiento                       'pos 35
        strlinea = strlinea & separador & "10002"                               'pos 36
        strlinea = strlinea & separador & "No"                                  'pos 37
        strlinea = strlinea & separador & Left(nombre, 60)                          'pos 38
        strlinea = strlinea & separador & Left(nombreFamiliar(Ternro, 10), 60)      'pos 39 - Nombre madre
        strlinea = strlinea & separador & Left(nombreFamiliar(Ternro, 9), 60)       'pos 40 - Nombre padre
        strlinea = strlinea & separador & Left(nroDomicilio, 6)                 'pos 41
        Call lineaEnBlanco(strlinea, "", 1, separador)                          'pos 42
        strlinea = strlinea & separador & Left(nroIdentif, 8)                   'pos 43
        Call lineaEnBlanco(strlinea, "", 2, separador)                          'pos 44-45
        strlinea = strlinea & separador & sexo                                  'pos 46
        strlinea = strlinea & separador & Left(telefono(rs_consult!domnro, 1), 15) 'pos 47
        Call lineaEnBlanco(strlinea, "", 1, separador)                          'pos 48
        Call lineaEnBlanco(strlinea, "null", 4, separador)                      'pos 49-52
        strlinea = strlinea & separador & "No"                                  'pos 53
        Call lineaEnBlanco(strlinea, "null", 1, separador)                      'pos 54
        strlinea = strlinea & separador & Left(nrodni, 10)                      'pos 55
        
        archSalida.writeline strlinea
        
        Flog.writeline "Linea para el el tercero: " & Ternro & " generada. Archivo: AR_Pfisica_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
        
        If rs_consult.State = adStateOpen Then rs_consult.Close
        Set rs_consult = Nothing
        
    End If
End Sub

Sub registroEmpleo(ByVal Ternro As String, ByRef archSalida, ByVal separador As String, ByVal teestablecimiento As String, ByVal tcodnro As String)
Dim strlinea As String
Dim fechaIngreso As String
Dim tenroCampo7 As String
Dim tenroCampo15 As String
Dim tenroCampo21 As String
Dim rs_consult As New ADODB.Recordset
Dim estado As String
Dim legajo As String
Dim nroId As String
Dim reportaA As String

    Flog.writeline "Entrando a importar el tercero: " & Ternro & ". Archivo: AR_RegEmp_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
    
    'Busco la configuracion del reporte
    StrSql = " SELECT confnrocol, confval FROM confrep WHERE repnro = 419 AND conftipo = 'TE' "
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        Do While Not rs_consult.EOF
            Select Case CLng(rs_consult!confnrocol)
                Case 1:
                    tenroCampo15 = rs_consult!confval
                    Flog.writeline "Se encontro tipo de Estructura para el campo 15 y 20. Archivo: AR_RegEmp_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                Case 2:
                    tenroCampo21 = rs_consult!confval
                    Flog.writeline "Se encontro tipo estructura para el campo 21. Archivo: AR_RegEmp_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
                Case 3:
                    tenroCampo7 = rs_consult!confval
                    Flog.writeline "Se encontro tipo de Estructura para el campo 7. Archivo: AR_RegEmp_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
            End Select
            rs_consult.MoveNext
        Loop
    Else
        tenroCampo15 = 0
        tenroCampo21 = 0
        tenroCampo7 = 0
        Flog.writeline "No se configuro los Tipos de estructuras para los campos 7, 15, 20, 21, 32 y/o 33. Archivo: AR_RegEmp_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
    End If
    
    'busco la fecha de ingreso por medio de la fase
    StrSql = " SELECT altfec FROM fases WHERE empleado = " & Ternro & " ORDER BY altfec ASC "
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        fechaIngreso = rs_consult!altfec
    End If
    
    'busco datos basico del empleado
    StrSql = " SELECT empleado.empest, empleado.empleg legajo, repA.empleg reportaA FROM empleado " & _
             " LEFT JOIN empleado repA ON repA.ternro = empleado.empreporta " & _
             " WHERE empleado.ternro = " & Ternro
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        If rs_consult!empest = -1 Then
            estado = "No"
        Else
            estado = "YES"
        End If
        legajo = CLng(88000000) + CLng(rs_consult!legajo)
        nroId = CLng(26000000) + CLng(rs_consult!legajo)
        reportaA = IIf(IsNull(rs_consult!reportaA), "", rs_consult!reportaA)
        If reportaA <> "" Then
            reportaA = CLng(88000000) + CLng(reportaA)
        End If
    End If
    
    
    strlinea = "No"                                                             'pos 1
    strlinea = strlinea & separador & fechaIngreso                              'pos 2
    Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 3
    Call lineaEnBlanco(strlinea, "", 1, separador)                              'pos 4
    strlinea = strlinea & separador & "No"                                      'pos 5
    Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 6
    strlinea = strlinea & separador & obtenerCod(Ternro, tenroCampo7, Date, tcodnro, 880000, 7, 6)   'pos 7
    strlinea = strlinea & separador & "18"                                      'pos 8
    strlinea = strlinea & separador & obtenerCod(Ternro, "5", Date, tcodnro, 0, 9, 15)   'pos 9
    Call lineaEnBlanco(strlinea, "", 1, separador)                              'pos 10
    strlinea = strlinea & separador & "No"                                      'pos 11
    Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 12
    strlinea = strlinea & separador & fechaUltimaFase(Ternro)                   'pos 13
    strlinea = strlinea & separador & "No"                                      'pos 14
    strlinea = strlinea & separador & obtenerCod(Ternro, tenroCampo15, Date, tcodnro, 0, 15, 50) 'pos 15
    Call lineaEnBlanco(strlinea, "", 1, separador)                              'pos 16
    strlinea = strlinea & separador & estado                                    'pos 17
    Call lineaEnBlanco(strlinea, "", 1, separador)                              'pos 18
    strlinea = strlinea & separador & "100"                                     'pos 19
    strlinea = strlinea & separador & "45"                                      'pos 20
    strlinea = strlinea & separador & obtenerCod(Ternro, teestablecimiento, Date, tcodnro, 0, 21, 8)    'pos 21
    strlinea = strlinea & separador & obtenerCod(Ternro, "4", Date, tcodnro, 880000, 22, 6)     'pos 22
    strlinea = strlinea & separador & Left(nroId, 8)                'pos 23
    strlinea = strlinea & separador & Left(legajo, 8)                           'pos 24
    strlinea = strlinea & separador & reportaA                                  'pos 25
    strlinea = strlinea & separador & "00"                                      'pos 26
    Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 27
    strlinea = strlinea & separador & Left(legajo, 10)                          'pos 28
    Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 29
    strlinea = strlinea & separador & "No"                                      'pos 30
    Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 31
    strlinea = strlinea & separador & "1"                                       'pos 32
    strlinea = strlinea & separador & "18"                                      'pos 33
        
    archSalida.writeline strlinea
    
    Flog.writeline "Linea para el el tercero: " & Ternro & " generada. Archivo: AR_RegEmp_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Sub

Sub historicoEstructura(ByVal Ternro As String, ByRef archSalida, ByVal tenro As String, ByVal separador As String, ByVal teestablecimiento As String, ByVal tcodnro As String)
Dim strlinea As String
Dim rs_consult As New ADODB.Recordset
Dim legajo As String
Dim archivoGenerar As String
Dim codFuncion As String

    'para generar el log correcto
    Select Case tenro
        Case "4"
            archivoGenerar = "AR_Hlotac_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
        Case "5"
            archivoGenerar = "AR_Hfuncao_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
    End Select
    
    Flog.writeline "Entrando a importar el tercero: " & Ternro & ". Archivo: " & archivoGenerar
        
    StrSql = " SELECT his_estructura.estrnro, estrcodext, empleg, htetdesde, nrocod FROM empleado " & _
             " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = " & tenro & _
             " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
             " LEFT JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro AND estr_cod.tcodnro = " & tcodnro & _
             " WHERE empleado.ternro = " & Ternro & _
             " ORDER BY htetdesde DESC "
    OpenRecordset StrSql, rs_consult
    Do While Not rs_consult.EOF
        legajo = CLng(88000000) + CLng(rs_consult!empleg)
        
        
        Select Case tenro
            Case "4" 'Archivo que genera los movimientos de puestos
                If Not IsNull(rs_consult!nrocod) Then
                    If rs_consult!nrocod <> "" Then
                        If IsNumeric(rs_consult!nrocod) Then
                            codFuncion = CLng(880000) + CLng(rs_consult!nrocod)
                        Else
                            Flog.writeline "El codigo AGCO asignado a la estructura: " & rs_consult!estrnro & " no es numerico. Archivo: " & archivoGenerar
                        End If
                    End If
                Else
                    codFuncion = ""
                End If
                
                strlinea = "100"                                                            'pos 1
                strlinea = strlinea & separador & obtenerCod(Ternro, teestablecimiento, rs_consult!htetdesde, tcodnro, 0, 2, 8) 'pos 2
                strlinea = strlinea & separador & Left(codFuncion, 6)                       'pos 3
                strlinea = strlinea & separador & Left(legajo, 8)                           'pos 4
                strlinea = strlinea & separador & rs_consult!htetdesde                      'pos 5
    
            Case "5" 'Archivo que genera los movimientos de centro de costo
                strlinea = IIf(IsNull(rs_consult!nrocod), "", rs_consult!nrocod)            'pos 1
                strlinea = strlinea & separador & "100"                                     'pos 2
                strlinea = strlinea & separador & obtenerCod(Ternro, teestablecimiento, rs_consult!htetdesde, tcodnro, 0, 3, 8) 'pos 3
                strlinea = strlinea & separador & Left(legajo, 8)                           'pos 4
                Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 5
                strlinea = strlinea & separador & rs_consult!htetdesde                      'pos 6
                
        End Select
        
        archSalida.writeline strlinea
        rs_consult.MoveNext
    Loop
    
    
    Flog.writeline "Linea para el el tercero: " & Ternro & " generada. Archivo: " & archivoGenerar
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Sub

Sub historicoSalarial(ByVal Ternro As String, ByRef archSalida, ByVal separador As String, ByVal separadorDecimal As String, ByVal teestablecimiento As String, ByVal tcodnro As String)
Dim strlinea As String
Dim rs_consult As New ADODB.Recordset
Dim rs_consultliq As New ADODB.Recordset
Dim legajo As String
Dim grado As String
Dim itemnro As String
Dim valorSalario As String
Dim codValorHoras As String
Dim valorHoras As String
Dim tipoValorHoras As String
Dim mes As String
Dim anio As String

    Flog.writeline "Entrando a importar el tercero: " & Ternro & ". Archivo: AR_HSal_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
    
    Flog.writeline "Buscando la configuracion del reporte para el archivo: AR_HSal_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
    
    StrSql = " SELECT confnrocol, conftipo, confval, confval2 FROM confrep WHERE repnro =  419 AND (confnrocol = 5 OR confnrocol = 6 ) "
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        Flog.writeline "Configuracion encontrada para el archivo: AR_HSal_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
        Do While Not rs_consult.EOF
            Select Case CLng(rs_consult!confnrocol)
                Case 5
                    itemnro = rs_consult!confval
                Case 6
                    codValorHoras = rs_consult!confval2
                    tipoValorHoras = rs_consult!conftipo
            End Select
            rs_consult.MoveNext
        Loop
    Else
        Flog.writeline "No se encontro la configuracion para el archivo: AR_HSal_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt, columna 5 y 6 del reporte."
    End If
    
    'busco primero si el empleado tiene algun grado asociado
    StrSql = " SELECT grado.gradesabr FROM empleado " & _
             " INNER JOIN grado ON grado.granro = empleado.granro " & _
             " WHERE empleado.ternro = " & Ternro
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        grado = rs_consult!gradesabr
    Else
        grado = ""
    End If
    'Busco el grado del empleado - Grado asociado al puesto vigente
    StrSql = " SELECT grado.granro, grado.gradesabr " & _
             " From grado " & _
             " INNER JOIN puesto_grado ON puesto_grado.granro = grado.granro " & _
             " INNER JOIN puesto       ON puesto.puenro = puesto_grado.puenro " & _
             " INNER JOIN estructura ON estructura.estrnro = puesto.estrnro " & _
             " INNER JOIN his_estructura ON his_estructura.estrnro = puesto.estrnro " & _
             " Where his_estructura.htethasta Is Null " & _
             " AND    his_estructura.ternro = " & Ternro
    OpenRecordset StrSql, rs_consult
    
    If Not rs_consult.EOF Then
        If (grado = "") Then
            'Si existe un solo grado para el puesto lo asigno, si el grado no esta cargado
            grado = rs_consult!gradesabr
        End If
        Flog.writeline "Grado del puesto actual encontrado para el tercero: " & Ternro & "."
    Else
        grado = ""
        Flog.writeline "El tercero: " & Ternro & " no tiene puesto o grado asociado al puesto actual."
    End If
    
    
    StrSql = " SELECT empleg, remdesabr, vpactado, remfecacuerdo FROM empleado " & _
             " INNER JOIN remu_emp on remu_emp.ternro = empleado.ternro AND remitenro = " & itemnro & _
             " INNER JOIN remu_per on remu_per.rempernro = remu_emp.remperiod " & _
             " WHERE Empleado.Ternro = " & Ternro
    OpenRecordset StrSql, rs_consult
    Do While Not rs_consult.EOF
        legajo = CLng(88000000) + CLng(rs_consult!empleg)
        
        'Busco el proceso de liquidacion segun la fecha de acuerdo del item
        StrSql = " SELECT DISTINCT pliqmes, pliqanio from proceso " & _
                 " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro " & _
                 " WHERE profecini <= " & cambiaFecha(rs_consult!remfecacuerdo) & " AND profecfin >= " & cambiaFecha(rs_consult!remfecacuerdo)
        OpenRecordset StrSql, rs_consultliq
        If Not rs_consultliq.EOF Then
            valorHoras = buscarConceptoAcumPorEtiqueta(tipoValorHoras, Ternro, codValorHoras, rs_consultliq!pliqmes, rs_consultliq!pliqanio)
            valorHoras = FormatNumber(CStr(valorHoras), 2)
            valorHoras = Replace(CStr(valorHoras), ".", "")
            valorHoras = Replace(CStr(valorHoras), ",", "")
            valorHoras = Left(CStr(valorHoras), Len(CStr(valorHoras)) - 2) & separadorDecimal & Right(CStr(valorHoras), 2)
        Else
            Flog.writeline "La fecha de acuerdo del item para el tercero: " & Ternro & " tiene una fecha de acuerdo fuera de todos los procesos de liquidacion."
        End If
        
        'Controlo que el valor pactado no sea nulo
        If IsNull(rs_consult!vpactado) Then
            valorSalario = 0
        Else
            If rs_consult!vpactado = "" Then
                valorSalario = 0
            Else
                valorSalario = rs_consult!vpactado
                valorSalario = FormatNumber(CStr(valorSalario), 2)
                valorSalario = Replace(CStr(valorSalario), ".", "")
                valorSalario = Replace(CStr(valorSalario), ",", "")
                valorSalario = Left(CStr(valorSalario), Len(CStr(valorSalario)) - 2) & separadorDecimal & Right(CStr(valorSalario), 2)
            
            End If
        End If
        
        
        strlinea = "100"                                                            'pos 1
        strlinea = strlinea & separador & obtenerCod(Ternro, teestablecimiento, rs_consult!remfecacuerdo, tcodnro, 0, 2, 8) 'pos 2
        Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 3
        strlinea = strlinea & separador & Left(legajo, 8)                           'pos 4
        strlinea = strlinea & separador & Left(rs_consult!remdesabr, 6)             'pos 5
        strlinea = strlinea & separador & Left(grado, 2)                            'pos 6
        strlinea = strlinea & separador & Right(valorHoras, 6)                      'pos 7
        strlinea = strlinea & separador & Right(valorSalario, 14)                   'pos 8
        Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 9
        strlinea = strlinea & separador & "M"                                       'pos 10
        strlinea = strlinea & separador & rs_consult!remfecacuerdo                  'pos 11
        archSalida.writeline strlinea
        rs_consult.MoveNext
        
    Loop
    
    Flog.writeline "Linea para el el tercero: " & Ternro & " generada. Archivo: AR_HSal_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
    
    If rs_consultliq.State = adStateOpen Then rs_consultliq.Close
    Set rs_consultliq = Nothing
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Sub
                     
Sub historicoLicencias(ByVal Ternro As String, ByRef archSalida, ByVal separador As String, ByVal teestablecimiento As String, ByVal tcodnro As String)
Dim strlinea As String
Dim rs_consult As New ADODB.Recordset
Dim listaLicencias As String

    Flog.writeline "Entrando a importar el tercero: " & Ternro & ". Archivo: AR_Hafastg_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
    
    Flog.writeline "Buscando la configuracion del reporte para el archivo: AR_Hafastg_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"

    StrSql = " SELECT confval FROM confrep WHERE repnro =  419 AND upper(conftipo) = 'LIC' "
    OpenRecordset StrSql, rs_consult
    listaLicencias = "0"
    If Not rs_consult.EOF Then
        Flog.writeline "Configuracion encontrada para el archivo: AR_Hafastg_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt"
        Do While Not rs_consult.EOF
            listaLicencias = listaLicencias & "," & rs_consult!confval
            rs_consult.MoveNext
        Loop
    Else
        Flog.writeline "No se encontro la configuracion para el archivo: AR_Hafastg_580_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_" & Right("00" & Hour(Time), 2) & Right("00" & Minute(Time), 2) & Right("00" & Second(Time), 2) & ".txt, columna tipo LIC."
    End If
    
    'Busco los datos del empleado junto con las licencias
    StrSql = " SELECT empleg, elfechadesde, elfechahasta, confval2 FROM empleado " & _
             " INNER JOIN emp_lic ON emp_lic.empleado = empleado.ternro " & _
             " INNER JOIN confrep ON confrep.confval = emp_lic.tdnro AND repnro = 419 AND UPPER(conftipo) = 'LIC' " & _
             " WHERE tdnro in (" & listaLicencias & ") AND empleado.ternro = " & Ternro & _
             " ORDER BY elfechadesde ASC "
    OpenRecordset StrSql, rs_consult
    Do While Not rs_consult.EOF
        legajo = CLng(88000000) + CLng(rs_consult!empleg)
        
        strlinea = rs_consult!elfechadesde                                          'pos 1
        strlinea = strlinea & separador & rs_consult!elfechahasta                   'pos 2
        strlinea = strlinea & separador & "100"                                     'pos 3
        strlinea = strlinea & separador & obtenerCod(Ternro, teestablecimiento, rs_consult!elfechadesde, tcodnro, 0, 4, 8)  'pos 4
        strlinea = strlinea & separador & Left(legajo, 8)                           'pos 5
        Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 6
        strlinea = strlinea & separador & rs_consult!confval2                       'pos 7
        Call lineaEnBlanco(strlinea, "null", 1, separador)                          'pos 8
        
        
        rs_consult.MoveNext
        archSalida.writeline strlinea
    Loop
    Flog.writeline "Entrando a importar el tercero: " & Ternro
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Sub


Sub lineaEnBlanco(ByRef strlinea As String, ByVal valor As String, ByVal Cantidad As Long, ByVal separador As String)
Dim i As Long
    
    For i = 1 To Cantidad
        If valor = "null" Then
            strlinea = strlinea & separador
        End If
        
        If valor = "" Then
            strlinea = strlinea & separador & " "
        End If
                
    Next
End Sub


Function fechaUltimaFase(ByVal Ternro As String)
Dim rs_consult As New ADODB.Recordset
Dim salida As String
    
    Flog.writeline "Obtenemos la ultima fase del ternro: " & Ternro
    salida = ""
    StrSql = " SELECT bajfec FROM fases where empleado = " & Ternro & " ORDER BY altfec DESC "
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        If Not IsNull(rs_consult!bajfec) Then
            salida = rs_consult!bajfec
            Flog.writeline "Fecha de baja de ultima fase, encontrada para el ternro: " & Ternro
        End If
    End If
    fechaUltimaFase = salida
        
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Function

Sub borrarArchivo(archivo)
Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(archivo) Then
        fs.deletefile archivo, True
    End If
End Sub


'Sub estructura(ByRef strlinea As String, ByVal ternro As String, ByVal tenro As Long, ByVal fecha As String, ByVal separador As String, ByVal columna As String, ByVal campo As Long, ByVal longitud As Long)
''Obtiene la descripcion o codigo externo de la estructura actual cargada al empleado, se corta la longitud de lo informado segun el campo, si el empleado no posee estructura dependiendo del campo devuelve null o blanco
'Dim rs_consult As New ADODB.Recordset
'
'    StrSql = " SELECT estructura.estrdabr, estructura.estrcodext FROM his_estructura " & _
'             " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro AND his_estructura.tenro = " & tenro & _
'             " WHERE ternro = " & ternro & " AND (htetdesde <= " & cambiaFecha(fecha) & " AND (htethasta >= " & cambiaFecha(fecha) & " OR htethasta is null))"
'    OpenRecordset StrSql, rs_consult
'
'    If Not rs_consult.EOF Then
'
'        Select Case columna
'            Case "estrcodext":
'                strlinea = strlinea & separador & Left(rs_consult!estrcodext, longitud)
'            Case "estrdabr":
'                strlinea = strlinea & separador & Left(rs_consult!estrdabr, longitud)
'        End Select
'    Else
'        'Si el empleado no tiene asociado tipo de estructura dependiendo el campo se informa null o blanco
'        Select Case CLng(campo)
'            Case 7: 'campo 7 - cargo - informa null (registro de empleo)
'                strlinea = strlinea & separador
'            Case 15: 'campo 15 - Descripción Escala - informa blanco
'                strlinea = strlinea & separador & " "
'            Case 20: 'campo 20 - Escala - informa null
'                strlinea = strlinea & separador
'            Case 21: 'campo 21 - Establecimiento  - informa null (todos los archivos)
'                strlinea = strlinea & separador
'            Case 22: 'campo 22 - Funcion  - informa null
'                strlinea = strlinea & separador
'            Case 23: 'campo 23 - Establecimiento  - informa null
'                strlinea = strlinea & separador
'            Case 32: 'campo 22 - Turma  - informa null
'                strlinea = strlinea & separador
'            Case 33: 'campo 22 - Vinculo  - informa null
'                strlinea = strlinea & separador
'
'        End Select
'
'    End If
'
'    If rs_consult.State = adStateOpen Then rs_consult.Close
'    Set rs_consult = Nothing
'
'End Sub

Function telefono(ByVal domnro As Integer, ByVal pos As Integer)
Dim salida As String
Dim rs_consult As New ADODB.Recordset
Dim StrSql As String
Dim posAux As Integer
    
    salida = " "
    StrSql = " SELECT telnro FROM telefono " & _
             " WHERE domnro = " & domnro & " ORDER BY telnro "
    OpenRecordset StrSql, rs_consult
    posAux = 1
    Do While Not rs_consult.EOF
        If posAux = pos Then
            salida = rs_consult!telnro
        End If
        posAux = posAux + 1
        rs_consult.MoveNext
    Loop
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing
    
    telefono = salida
End Function

Function nombreFamiliar(ByVal Ternro As Integer, ByVal codfam As Integer)
Dim salida As String
Dim rs_consult As New ADODB.Recordset
Dim StrSql As String

    salida = ""
    StrSql = " SELECT tercero.terape, tercero.ternom FROM familiar " & _
             " INNER JOIN tercero ON tercero.ternro = familiar.ternro " & _
             " WHERE familiar.parenro = " & codfam & " And familiar.Empleado = " & Ternro
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        salida = rs_consult!terape & ", " & rs_consult!ternom
    End If

    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing
    
    nombreFamiliar = salida
End Function

Function obtenerCod(ByVal Ternro As String, ByVal tenro As String, ByVal fecha As String, ByVal tcodnro As String, ByVal valor As String, ByVal campo As String, ByVal longitud As String)
Dim salida As String
Dim rs_consult As New ADODB.Recordset
Dim StrSql As String

    salida = ""
    StrSql = " SELECT his_estructura.estrnro, nrocod FROM his_estructura " & _
             " LEFT JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro AND estr_cod.tcodnro = " & tcodnro & _
             " Where his_estructura.ternro = " & Ternro & " And his_estructura.tenro = " & tenro & _
             " AND (his_estructura.htetdesde <= " & ConvFecha(fecha) & " AND (his_estructura.htethasta >= " & ConvFecha(fecha) & " OR his_estructura.htethasta is null)) "
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        If Not IsNull(rs_consult!nrocod) Then
            salida = Trim(rs_consult!nrocod)
            If IsNumeric(salida) Then
                If CLng(valor) > 0 Then
                    salida = CLng(salida) + CLng(valor)
                End If
            Else
                Flog.writeline "El codigo AGCO asignado a la estructura: " & rs_consult!estrnro & " no es numerico."
            End If
        Else
            Select Case campo
                Case 15 ' Descripción Escala - Registro Empleo.txt - Blanco
                    salida = " "
            End Select
        End If
    Else
        Select Case campo
            Case 15 ' Descripción Escala - Registro Empleo.txt - Blanco
                salida = " "
        End Select
    End If
    
    
    
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

    obtenerCod = Left(salida, longitud)
End Function

