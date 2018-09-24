Attribute VB_Name = "MdlinterfazGRDW"
Option Explicit

Const Version = "1.00"
Const FechaVersion = "30/09/2014" 'LED - CAS-26584 - GE - GRDW (Datawarehouse)
'LED - Version Inicial

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

    Nombre_Arch = PathFLog & "InterfazGRDW_" & NroProcesoBatch & ".log"
    
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 431 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = IIf(EsNulo(rs_batch_proceso!bprcparam), "", rs_batch_proceso!bprcparam)
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call interfazGRDW(NroProcesoBatch, bprcparam)
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


Public Sub interfazGRDW(ByVal bpronro As Long, ByVal Parametros As String)
 
 Dim directorio As String
 Dim Nombre_Arch As String
 Dim rs_consult As New ADODB.Recordset
 Dim rsEmpleados As New ADODB.Recordset
 Dim separador As String
 Dim separadorDecimal As String
 Dim strLineaExp As String
 Dim archSalida1    'dimension.txt
 Dim archSalida2    'payroll.txt
 Dim archSalida3    'pay_distribution.txt
 Dim archSalida4    'leave.txt
 Dim archSalida5    'paycodes.txt
 Dim archSalida6    'trailer.txt
 Dim porc As Double
 Dim cantEmpleados As Integer
 Dim arrayArchivos
 Dim indice As Long
 Dim usaencabezado As String
 Dim teestablecimiento As String
 Dim tcodnro As String
 Dim cantReg1 As Long
 Dim cantReg2 As Long
 Dim cantReg3 As Long
 Dim cantReg4 As Long
 Dim cantReg5 As Long
     
 Dim ostream1 As Object    'dimension.txt
 Dim ostream2 As Object    'payroll.txt
 Dim ostream3 As Object    'pay_distribution.txt
 Dim ostream4 As Object    'leave.txt
 Dim ostream5 As Object    'paycodes.txt
 Dim ostream6 As Object    'trailer.txt

'Dim bRet() As Byte
Set ostream1 = CreateObject("ADODB.Stream") 'dimension.txt
ostream1.Type = adTypeText
ostream1.Charset = "UTF-8" 'Indicate the charactor encoding
ostream1.Open
ostream1.Position = 0

Set ostream2 = CreateObject("ADODB.Stream") 'payroll.txt
ostream2.Type = adTypeText
ostream2.Charset = "UTF-8" 'Indicate the charactor encoding
ostream2.Open
ostream2.Position = 0
    
Set ostream3 = CreateObject("ADODB.Stream") 'pay_distribution.txt
ostream3.Type = adTypeText
ostream3.Charset = "UTF-8" 'Indicate the charactor encoding
ostream3.Open
ostream3.Position = 0
    
Set ostream4 = CreateObject("ADODB.Stream") 'leave.txt
ostream4.Type = adTypeText
ostream4.Charset = "UTF-8"
ostream4.Open
ostream4.Position = 0
    
Set ostream5 = CreateObject("ADODB.Stream") 'paycodes.txt
ostream5.Type = adTypeText
ostream5.Charset = "UTF-8"
ostream5.Open
ostream5.Position = 0
    
Set ostream6 = CreateObject("ADODB.Stream") 'trailer.txt
ostream6.Type = adTypeText
ostream6.Charset = "UTF-8"
ostream6.Open
ostream6.Position = 0
    
    
    On Error GoTo CE
    
    arrayArchivos = Split(Parametros, "@")
    
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "Comienza la exportacion "
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------"
    Progreso = 0
    
    'busco todos los empleados
    'StrSql = " SELECT e.ternro FROM empleado e ORDER BY e.empleg ASC "
    StrSql = " SELECT e.ternro FROM empleado e WHERE empleg <= 30 ORDER BY e.empleg ASC "
    OpenRecordset StrSql, rsEmpleados
            
    cantEmpleados = rsEmpleados.RecordCount
    
    'porc = CLng(50 / cantEmpleados)
    If cantEmpleados = 0 Then
        Progreso = 100
        cantEmpleados = 1
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron empleados."
        Exit Sub
    End If
    
    'Busco la carpeta in-out
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline "No esta configurado el directorio de sistema."
        Exit Sub
    End If
        
    'obtengo los datos del modelo
    StrSql = "SELECT * FROM modelo WHERE modnro = 393 "
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
    Call borrarArchivo(directorio & "\dimension.txt")
    Call borrarArchivo(directorio & "\payroll.txt")
    Call borrarArchivo(directorio & "\pay_distribution.txt")
    Call borrarArchivo(directorio & "\leave.txt")
    Call borrarArchivo(directorio & "\paycodes.txt")
    Call borrarArchivo(directorio & "\trailer.txt")
    
    cantReg1 = 0
    cantReg2 = 0
    cantReg3 = 0
    cantReg4 = 0
    cantReg5 = 0

    For indice = 0 To UBound(arrayArchivos)
        Select Case indice
            Case 0
                If arrayArchivos(indice) = -1 Then
                    archSalida1 = directorio & "\dimension.txt"
                    'Set archSalida1 = fs.CreateTextFile(archSalida1, True)
                    If usaencabezado Then
                        Call encabezado(indice, separador, ostream1)
                        'cantReg1 = cantReg1 + 1
                    End If

                Else
                    Flog.writeline "No se generara el archivo: dimension.txt"
                End If
            Case 1
                If arrayArchivos(indice) = -1 Then
                    archSalida2 = directorio & "\payroll.txt"
                    If usaencabezado Then
                        Call encabezado(indice, separador, ostream2)
                        'cantReg2 = cantReg2 + 1
                    End If
                Else
                    Flog.writeline "No se generara el archivo: payroll.txt"
                End If
                 
            Case 2
                If arrayArchivos(indice) = -1 Then
                    archSalida3 = directorio & "\pay_distribution.txt"
                    If usaencabezado Then
                        Call encabezado(indice, separador, ostream3)
                        'cantReg3 = cantReg3 + 1
                    End If
                Else
                    Flog.writeline "No se generara el archivo: pay_distribution.txt"
                End If
                
            Case 3
                If arrayArchivos(indice) = -1 Then
                    archSalida4 = directorio & "\leave.txt"
                    If usaencabezado Then
                        Call encabezado(indice, separador, ostream4)
                        'cantReg4 = cantReg4 + 1
                    End If
                Else
                    Flog.writeline "No se generara el archivo: leave.txt"
                End If
            
            Case 4
                If arrayArchivos(indice) = -1 Then
                    archSalida5 = directorio & "\paycodes.txt"
                    If usaencabezado Then
                        Call encabezado(indice, separador, ostream5)
                        'cantReg5 = cantReg5 + 1
                    End If
                Else
                    Flog.writeline "No se generara el archivo: paycodes.txt"
                End If
        
        End Select
    Next
    'este archivo se genera siempre
    archSalida6 = directorio & "\trailer.txt"
    
    porc = 80 / CLng(cantEmpleados)
    
    Do While Not rsEmpleados.EOF
        Progreso = Progreso + porc
        
        'el ultimo archivo no se realiza por empleado
        For indice = 0 To UBound(arrayArchivos) - 1
            Select Case indice
                Case 0  'dimension
                    If arrayArchivos(indice) = -1 Then
                        Call dimension(rsEmpleados!ternro, archSalida1, separador, separadorDecimal, ostream1, cantReg1)
                    End If
                Case 1  'payroll
                    If arrayArchivos(indice) = -1 Then
                        Call payroll(rsEmpleados!ternro, archSalida2, separador, separadorDecimal, usaencabezado, ostream2, cantReg2)
                    End If
                Case 2  'pay_distribution
                    If arrayArchivos(indice) = -1 Then
                        Call pay_distribution(rsEmpleados!ternro, archSalida3, separador, separadorDecimal, usaencabezado, ostream3, cantReg3)
                    End If
                Case 3
                    'Historico de Puestos
                    If arrayArchivos(indice) = -1 Then
                        Call leave(rsEmpleados!ternro, archSalida4, separador, separadorDecimal, usaencabezado, ostream4, cantReg4)
                    End If

            End Select
        Next
        
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        

        rsEmpleados.MoveNext
    Loop
    
'Este archivo no es por empleado por lo que no esta en el bucle principal
If arrayArchivos(4) = -1 Then
    Call paycodes(separador, usaencabezado, ostream5, cantReg5)
End If

Progreso = 90
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
objconnProgreso.Execute StrSql, , adExecuteNoRecords


Call trailer(separador, usaencabezado, cantReg1, cantReg2, cantReg3, cantReg4, cantReg5, ostream6)

Progreso = 95
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
objconnProgreso.Execute StrSql, , adExecuteNoRecords



For indice = 0 To UBound(arrayArchivos)
        Select Case indice
            Case 0
                If arrayArchivos(indice) = -1 Then
                    ostream1.Position = 0
                    ostream1.SaveToFile archSalida1 'Save the stream to a file
                End If
            Case 1
                If arrayArchivos(indice) = -1 Then
                    ostream2.Position = 0
                    ostream2.SaveToFile archSalida2 'Save the stream to a file
                End If
            Case 2
                If arrayArchivos(indice) = -1 Then
                    ostream3.Position = 0
                    ostream3.SaveToFile archSalida3 'Save the stream to a file
                End If
                
            Case 3
                If arrayArchivos(indice) = -1 Then
                    ostream4.Position = 0
                    ostream4.SaveToFile archSalida4 'Save the stream to a file
                End If
            
            Case 4
                If arrayArchivos(indice) = -1 Then
                    ostream5.Position = 0
                    ostream5.SaveToFile archSalida5 'Save the stream to a file
                End If
        
        End Select
    Next
    
    ostream6.Position = 0
    ostream6.SaveToFile archSalida6 'Save the stream to a file
    
    ostream1.Close
    Set ostream1 = Nothing
    ostream2.Close
    Set ostream2 = Nothing
    ostream3.Close
    Set ostream3 = Nothing
    ostream4.Close
    Set ostream4 = Nothing
    ostream5.Close
    Set ostream5 = Nothing
    ostream6.Close
    Set ostream6 = Nothing
    
    
    'archSalida1.Close
    'Set archSalida1 = Nothing
    
    Progreso = 100
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
  
    If objRs.State = adStateOpen Then objRs.Close
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




Public Function cambiaFecha(ByVal Fecha As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea la fecha al formato de insercion de la base de datos.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


    If EsNulo(Fecha) Then
        cambiaFecha = "NULL"
    Else
        cambiaFecha = ConvFecha(Fecha)
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



Sub dimension(ByVal ternro As String, ByRef archSalida, ByVal separador As String, ByVal separadorDecimal As String, ByRef ostream, ByRef cantreg As Long)
Dim strlinea As String
Dim rs_consult As New ADODB.Recordset
Dim nombre As String
Dim sexo As String
Dim conccod1Tipo As String
Dim conccod1 As String
Dim conccod2 As String
Dim acuAnual As String
Dim tpanro As String
Dim tenroNov As String
Dim tdnro As String
Dim teUnidadNegocio As String
    
    Flog.writeline "-----------------------------------------------------------------------------"
    Flog.writeline "Entrando a importar el tercero: " & ternro & ". Archivo: dimension.txt "
    Flog.writeline "-----------------------------------------------------------------------------"
    'busco la configuracion
    tdnro = "0"
    teUnidadNegocio = "0"
    StrSql = " SELECT confval2, confval, confnrocol, conftipo FROM confrep WHERE repnro = 457 AND confnrocol in (1,2,3,4,6)"
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        Do While Not rs_consult.EOF
            Select Case CLng(rs_consult!confnrocol)
                Case 1: 'concepto que se muestra en los campos Base Salary (Offer) y Actual Base Salary
                    conccod1Tipo = rs_consult!conftipo
                    conccod1 = rs_consult!confval2
                
                Case 2: 'tipo de estructura para obtener la novedad por estructura, campo Working Hours
                    If UCase(rs_consult!conftipo) = "TE" Then
                        tenroNov = rs_consult!confval2
                    End If
                    
                    'concepto y parametro para obtener la novedad por estructura, campo Working Hours
                    If UCase(rs_consult!conftipo) = "CP" Then
                        conccod2 = rs_consult!confval2
                        tpanro = rs_consult!confval
                    End If
                
                Case 3: 'concepto que se muestra en los campos Base Salary (Offer) y Actual Base Salary
                    tdnro = rs_consult!confval2
                
                Case 4: 'tipo de estructura grupo de negocio
                    acuAnual = rs_consult!confval
                
                Case 6: 'tipo de estructura grupo de negocio
                    teUnidadNegocio = rs_consult!confval
                                
            End Select
            
            rs_consult.MoveNext
        Loop
    Else
        Flog.writeline "No se encontro configuracion para el proceso, reporte numero 457. Se aborta el proceso."
        Exit Sub
    End If
    
    'busco datos basico del empleado
    StrSql = " SELECT empleado.empleg, empleado.terape2, empleado.ternom, empleado.ternom2, empleado.terape " & _
             " ,terfecnac, nacionalidad.nacionaldes, localidad.locdesc, detdom.codigopostal, tersex, detdom.calle, pais.paisdesc " & _
             " FROM empleado " & _
             " INNER JOIN tercero ON tercero.ternro = empleado.ternro " & _
             " INNER JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro " & _
             " LEFT JOIN cabdom ON cabdom.ternro = empleado.ternro AND cabdom.domdefault = -1 " & _
             " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
             " LEFT JOIN localidad ON localidad.locnro = detdom.locnro " & _
             " LEFT JOIN pais ON pais.paisnro = detdom.paisnro " & _
             " WHERE empleado.ternro = " & ternro
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then

        nombre = rs_consult!terape
        nombre = nombre & " " & rs_consult!ternom
        
        If CLng(rs_consult!tersex) = -1 Then
            sexo = "M"
        Else
            sexo = "F"
        End If
        
        
        strlinea = rs_consult!empleg
        strlinea = strlinea & separador & objConn.Properties(69)
        strlinea = strlinea & separador & rs_consult!empleg
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & rs_consult!empleg
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & Left(nombre, 40)
        strlinea = strlinea & separador & Left(nombre, 40)
        strlinea = strlinea & separador & Left(rs_consult!terape, 20)
        strlinea = strlinea & separador & Left(rs_consult!ternom, 20)
        strlinea = strlinea & separador & rs_consult!nacionaldes
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & Format(rs_consult!terfecnac, "YYYY-MM-DD")
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & sexo
        strlinea = strlinea & separador & rs_consult!nacionaldes
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & estructura(ternro, 19, Date)
        strlinea = strlinea & separador & "Mensual"
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & estructura(ternro, 4, Date)
        strlinea = strlinea & separador & fechaUltimaFase(ternro, "FECHAALTA")
        strlinea = strlinea & separador & fechaUltimaFase(ternro, "FECHAALTA")
        strlinea = strlinea & separador & fechaUltimaFase(ternro, "fecha")
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & Right(valorConAcu(ternro, conccod1Tipo, Date, conccod1, separadorDecimal), 15)
        strlinea = strlinea & separador & Right(valorConAcu(ternro, conccod1Tipo, Date, conccod1, separadorDecimal), 15)
        strlinea = strlinea & separador & estructura(ternro, 1, Date)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & empleadoActivo(ternro, Date)
        strlinea = strlinea & separador & estructura(ternro, 4, Date)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & Left(Replace(Replace(novestr(tenroNov, ternro, Date, conccod2, tpanro), ".", separadorDecimal), ",", separadorDecimal), 15)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & Format(licencia(ternro, Date, tdnro, "elfechadesde"), "YYYY-MM-DD")
        strlinea = strlinea & separador & Format(licencia(ternro, Date, tdnro, "elfechahasta"), "YYYY-MM-DD")
        strlinea = strlinea & separador & licencia(ternro, Date, tdnro, "tipdia.tddesc")
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & empleadoActivo(ternro, Date)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & BusMes(Month(Date))
        strlinea = strlinea & separador & datoProceso(ternro, Date, "PROFECPAGO")
        strlinea = strlinea & separador & Format("01/" & Month(Date) & "/" & Year(Date), "YYYY-MM-DD")
        strlinea = strlinea & separador & Format(DateAdd("d", -1, "01/" & Right("00" & Month(DateAdd("m", 1, Date)), 2) & "/" & Year(Date)), "YYYY-MM-DD")
        strlinea = strlinea & separador & datoProceso(ternro, Date, "PROFECINI")
        strlinea = strlinea & separador & datoProceso(ternro, Date, "prodesc")
        strlinea = strlinea & separador & "BDO"
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & Right(valorAcuAnual(ternro, Year(Date), acuAnual, separadorDecimal), 15)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & datoProceso(ternro, Date, "profecpago")
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & estructura(ternro, 10, Date)
        strlinea = strlinea & separador & codigoExtEstructura(ternro, "10", Date)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & domcilioEstructura(ternro, "1", Date, "ALL")
        strlinea = strlinea & separador & estructura(ternro, teUnidadNegocio, Date)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & empresaAbrev(ternro, Date)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & domcilioEstructura(ternro, "1", Date, "calle")
        strlinea = strlinea & separador & domcilioEstructura(ternro, "1", Date, "localidad")
        strlinea = strlinea & separador & domcilioEstructura(ternro, "1", Date, "cp")
        strlinea = strlinea & separador & domcilioEstructura(ternro, "1", Date, "pais")
        strlinea = strlinea & separador & ctabancaria(ternro, "10", Date)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & estructura(ternro, "5", Date)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & fechaUltimaFase(ternro, "motivo")
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & estructura(ternro, "5", Date)
        strlinea = strlinea & separador & "Argentina"
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & estructura(ternro, "3", Date)
        strlinea = strlinea & separador & estructura(ternro, "19", Date)
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & ""
        strlinea = strlinea & separador & estructura(ternro, "18", Date)
        strlinea = strlinea & separador & ""
        
        'guardo la linea en el archivo
        If Not EsNulo(strlinea) Then
            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
            cantreg = cantreg + 1
        End If
        Flog.writeline "-----------------------------------------------------------------------------"
        Flog.writeline "Linea para el el tercero: " & ternro & " generada. Archivo: dimension.txt"
        Flog.writeline "-----------------------------------------------------------------------------"
        
    End If
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Sub

Sub payroll(ByVal ternro As String, ByRef archSalida, ByVal separador As String, ByVal separadorDecimal As String, ByVal usaencabezado As String, ByRef ostream, ByRef cantreg As Long)
Dim strlinea As String  'pay_distribution
Dim rs_consult As New ADODB.Recordset
Dim tidnro As String
Dim pagos As String 'columna paycode
Dim tipPagos As String
Dim horas As String 'columna hours
Dim tipHoras As String
Dim montos As String 'columna amount
Dim tipMontos As String
Dim nrodoc As String

    Flog.writeline "-----------------------------------------------------------------------------"
    Flog.writeline "Entrando a importar el tercero: " & ternro & ". Archivo: payroll.txt"
    Flog.writeline "-----------------------------------------------------------------------------"
    
    Flog.writeline "Busco la configuracion del tipo de documento en el reporte"
    tidnro = "0"
    pagos = "'0'"
    horas = "'0'"
    montos = "'0'"
    StrSql = " SELECT confval, confnrocol, conftipo, confval2 FROM confrep WHERE repnro = 457 AND confnrocol in (7,8,9,10) "
    OpenRecordset StrSql, rs_consult
    Do While Not rs_consult.EOF
        Select Case CLng(rs_consult!confnrocol)
            Case 7
                tidnro = rs_consult!confval
            
            Case 8
                pagos = pagos & ",'" & rs_consult!confval2 & "'"
                tipPagos = rs_consult!conftipo
            
            Case 9
                horas = horas & ",'" & rs_consult!confval2 & "'"
                tipHoras = rs_consult!conftipo
        
            Case 10
                montos = montos & ",'" & rs_consult!confval2 & "'"
                tipMontos = rs_consult!conftipo
        End Select
        rs_consult.MoveNext
    Loop
    
    StrSql = " SELECT empleado.empleg, ter_doc.nrodoc FROM empleado " & _
             " LEFT JOIN ter_doc ON ter_doc.ternro = empleado.ternro AND tidnro = " & tidnro & _
             " WHERE empleado.ternro = " & ternro

    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        
                
        strlinea = "BDO"                                                            'pos 1
        strlinea = strlinea & separador & Left(rs_consult!empleg, 9)                'pos 2
        strlinea = strlinea & separador & rs_consult!nrodoc                         'pos 3
        strlinea = strlinea & separador & estructura(ternro, "10", Date)            'pos 4
        strlinea = strlinea & separador & Right(valorConAcuLista(ternro, tipPagos, Date, pagos, separadorDecimal), 15)    'pos 5
        strlinea = strlinea & separador & Right(valorConAcuLista(ternro, tipHoras, Date, horas, separadorDecimal), 15)    'pos 6
        strlinea = strlinea & separador & Right(valorConAcuLista(ternro, tipMontos, Date, montos, separadorDecimal), 15)  'pos 7
        strlinea = strlinea & separador & ""                                                'pos 8
        strlinea = strlinea & separador & ""                                                'pos 9
        strlinea = strlinea & separador & ""                                                'pos 10
        strlinea = strlinea & separador & ""                                                'pos 11
        strlinea = strlinea & separador & ""                                                'pos 12
        strlinea = strlinea & separador & ""                                                'pos 13
        strlinea = strlinea & separador & ""                                                'pos 14
        strlinea = strlinea & separador & Format(periodo(ternro, Date, "pliqdesde"), "YYYY-MM-DD")       'pos 15
        strlinea = strlinea & separador & Format(periodo(ternro, Date, "pliqhasta"), "YYYY-MM-DD")       'pos 16
        strlinea = strlinea & separador & "Argentina"                                       'pos 17
        
        
        'guardo la linea en el archivo
        If Not EsNulo(strlinea) Then
            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
            cantreg = cantreg + 1
        End If
        
        Flog.writeline "-----------------------------------------------------------------------------"
        Flog.writeline "Linea para el el tercero: " & ternro & " generada. Archivo: payroll.txt"
        Flog.writeline "-----------------------------------------------------------------------------"
    Else
        Flog.writeline "El tercero: " & ternro & " no es un empleado"
    End If
    
    
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Sub
Sub pay_distribution(ByVal ternro As String, ByRef archSalida, ByVal separador As String, ByVal separadorDecimal As String, ByVal usaencabezado As String, ByRef ostream, ByRef cantreg As Long)
Dim strlinea As String
Dim rs_consult As New ADODB.Recordset
Dim montos As String 'columna amount
Dim tipMontos As String
    Flog.writeline "-----------------------------------------------------------------------------"
    Flog.writeline "Entrando a importar el tercero: " & ternro & ". Archivo: pay_distribution.txt"
    Flog.writeline "-----------------------------------------------------------------------------"
    
    Flog.writeline "Busco la configuracion del tipo de documento en el reporte"
    montos = "'0'"
    StrSql = " SELECT confnrocol, conftipo, confval2 FROM confrep WHERE repnro = 457 AND confnrocol in (11) "
    OpenRecordset StrSql, rs_consult
    Do While Not rs_consult.EOF
        Select Case CLng(rs_consult!confnrocol)
            Case 11
                montos = montos & ",'" & rs_consult!confval2 & "'"
                tipMontos = rs_consult!conftipo
            
        End Select
        rs_consult.MoveNext
    Loop
    
    StrSql = " SELECT empleado.empleg FROM empleado WHERE empleado.ternro = " & ternro

    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        
                
        strlinea = "BDO"                                                            'pos 1
        strlinea = strlinea & separador & Left(rs_consult!empleg, 9)                'pos 2
        strlinea = strlinea & separador & estructura(ternro, "10", Date)            'pos 3
        strlinea = strlinea & separador & ""                                        'pos 4
        strlinea = strlinea & separador & ctabancariaEmpleado(ternro)               'pos 5
        strlinea = strlinea & separador & codigoExtEstructura(ternro, "41", Date)   'pos 6
        strlinea = strlinea & separador & "Citibank"                                'pos 7
        strlinea = strlinea & separador & "Electronico"                             'pos 8
        strlinea = strlinea & separador & estructura(ternro, "41", Date)            'pos 9
        strlinea = strlinea & separador & ""                                        'pos 10
        strlinea = strlinea & separador & "PESOS"                                   'pos 11
        strlinea = strlinea & separador & Right(valorConAcuLista(ternro, tipMontos, Date, montos, separadorDecimal), 15)      'pos 12
        strlinea = strlinea & separador & ""                                        'pos 12
        strlinea = strlinea & separador & ""                                        'pos 13
        
        'guardo la linea en el archivo
        If Not EsNulo(strlinea) Then
            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
            cantreg = cantreg + 1
        End If
        Flog.writeline "-----------------------------------------------------------------------------"
        Flog.writeline "Linea para el el tercero: " & ternro & " generada. Archivo: pay_distribution.txt"
        Flog.writeline "-----------------------------------------------------------------------------"
    Else
        Flog.writeline "El tercero: " & ternro & " no es un empleado"
    End If
    
    
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Sub

Sub leave(ByVal ternro As String, ByRef archSalida, ByVal separador As String, ByVal separadorDecimal As String, ByVal usaencabezado As String, ByRef ostream, ByRef cantreg As Long)
Dim strlinea As String
Dim rs_consult As New ADODB.Recordset

    Flog.writeline "-----------------------------------------------------------------------------"
    Flog.writeline "Entrando a importar el tercero: " & ternro & ". Archivo: leave.txt"
    Flog.writeline "-----------------------------------------------------------------------------"
    
    'por pedido del cliente no quiere informar nada en este archivo, se deja preparado para futuros cambios
    StrSql = " SELECT empleado.empleg FROM empleado WHERE empleado.ternro = " & ternro & " AND 1=2"

    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
                        
        strlinea = ""
        strlinea = strlinea & separador & ""                                        'pos 4
        strlinea = strlinea & separador & ""                                        'pos 5
        strlinea = strlinea & separador & ""                                        'pos 6
        strlinea = strlinea & separador & ""                                        'pos 7
        strlinea = strlinea & separador & ""                                        'pos 8
        strlinea = strlinea & separador & ""                                        'pos 9
        strlinea = strlinea & separador & ""                                        'pos 10
        strlinea = strlinea & separador & ""                                        'pos 11
        strlinea = strlinea & separador & ""                                        'pos 12
        strlinea = strlinea & separador & ""                                        'pos 13
        
        'guardo la linea en el archivo
        If Not EsNulo(strlinea) Then
            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
            cantreg = cantreg + 1
        End If
        
        Flog.writeline "-----------------------------------------------------------------------------"
        Flog.writeline "Linea para el el tercero: " & ternro & " generada. Archivo: leave.txt"
        Flog.writeline "-----------------------------------------------------------------------------"
    Else
        Flog.writeline "No hay empleado con datos para generar"
    End If
    
    
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Sub

Sub paycodes(ByVal separador As String, ByVal usaencabezado As String, ByRef ostream, ByRef cantreg As Long)
Dim strlinea As String
Dim rs_consult As New ADODB.Recordset
    
    Flog.writeline "-----------------------------------------------------------------------------"
    Flog.writeline "Entrando a importar archivo: paycodes.txt"
    Flog.writeline "-----------------------------------------------------------------------------"
    
    StrSql = " SELECT confval2, conftipo, concepto.conccod, concepto.concabr, acumulador.acunro, acumulador.acudesabr FROM confrep " & _
             " left join concepto on concepto.conccod = confrep.confval2 AND upper(confrep.conftipo) = 'CO' " & _
             " left join acumulador on acumulador.acunro = confrep.confval2 AND upper(confrep.conftipo) = 'AC' " & _
             " WHERE repnro = 457 AND confnrocol = 5 "
    OpenRecordset StrSql, rs_consult
    Do While Not rs_consult.EOF
        
        
        strlinea = "Argentina"                                                      'pos 1
        If UCase(rs_consult!conftipo) = "CO" Then
            strlinea = strlinea & separador & rs_consult!conccod                    'pos 2
            strlinea = strlinea & separador & rs_consult!concabr                    'pos 3
        End If
        If UCase(rs_consult!conftipo) = "AC" Then
            strlinea = strlinea & separador & rs_consult!acunro                     'pos 2
            strlinea = strlinea & separador & rs_consult!acudesabr                  'pos 3
        End If
        
        strlinea = strlinea & separador & ""                                        'pos 4
        strlinea = strlinea & separador & ""                                        'pos 5
        
        'guardo la linea en el archivo
        If Not EsNulo(strlinea) Then
            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
            cantreg = cantreg + 1
        End If
        
        rs_consult.MoveNext
        
    Loop
    
    Flog.writeline "-----------------------------------------------------------------------------"
    Flog.writeline "Archivo: paycodes.txt generado"
    Flog.writeline "-----------------------------------------------------------------------------"
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

End Sub
                     
Sub trailer(ByVal separador As String, ByVal usaencabezado As String, ByVal cantReg1 As Long, ByVal cantReg2 As Long, ByVal cantReg3 As Long, ByVal cantReg4 As Long, ByVal cantReg5 As Long, ByRef ostream)
Dim strlinea As String

    
    Flog.writeline "Entrando a importar archivo: trailer.txt"
    
    If usaencabezado Then
        strlinea = "dimension.txt"
        strlinea = strlinea & separador & "payroll.txt2"
        strlinea = strlinea & separador & "pay_distribution.txt"
        strlinea = strlinea & separador & "leave.txt"
        strlinea = strlinea & separador & "paycodes.txt"
        
        ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
    End If
        
    strlinea = CStr(cantReg1)                                   'pos 1
    strlinea = strlinea & separador & CStr(cantReg2)            'pos 2
    strlinea = strlinea & separador & CStr(cantReg3)            'pos 3
    strlinea = strlinea & separador & CStr(cantReg4)            'pos 4
    strlinea = strlinea & separador & CStr(cantReg5)            'pos 5
    
    'guardo la linea en el archivo
    If Not EsNulo(strlinea) Then
        ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
    End If
        
    
    Flog.writeline "Archivo: trailer.txt generado"
    
End Sub

Function fechaUltimaFase(ByVal ternro As String, ByVal campo As String)
Dim rs_consult As New ADODB.Recordset
Dim salida As String
    
    Flog.writeline "Obtenemos la ultima fase del ternro: " & ternro
    salida = ""
    StrSql = " SELECT fases.altfec, fases.bajfec, causa.caudes FROM fases " & _
             " LEFT JOIN causa on causa.caunro = fases.caunro " & _
             " WHERE empleado = " & ternro & " ORDER BY altfec DESC "
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        Flog.writeline "Fase encontrada para el ternro: " & ternro
        Select Case UCase(campo)
            Case "FECHAALTA"
                salida = Format(rs_consult!altfec, "YYYY-MM-DD")
                Flog.writeline "Fecha de Alta de ultima fase, encontrada para el ternro: " & ternro
        End Select
        If Not IsNull(rs_consult!bajfec) Then
            Flog.writeline "La fase posee fecha de baja"
            Select Case UCase(campo)
                Case "FECHA"
                    salida = Format(rs_consult!bajfec, "YYYY-MM-DD")
                    Flog.writeline "Fecha de baja de ultima fase, encontrada para el ternro: " & ternro
                Case "MOTIVO"
                    salida = rs_consult!caudes
                    Flog.writeline "Motivo de baja, encontrada para el ternro: " & ternro
            End Select
        Else
            Flog.writeline "La fase NO posee fecha de baja"
        End If
    Else
        Flog.writeline "El ternro: " & ternro & " no posee fases"
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



Function estructura(ByVal ternro As String, ByVal tenro As String, ByVal Fecha As String)
Dim salida As String
Dim rs_consult As New ADODB.Recordset
Dim StrSql As String

    salida = ""
    StrSql = " SELECT estructura.estrdabr FROM his_estructura " & _
             " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
             " WHERE his_estructura.ternro = " & ternro & " And his_estructura.tenro = " & tenro & _
             " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha) & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR his_estructura.htethasta is null)) "
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        salida = rs_consult!estrdabr
        Flog.writeline "Estructura tenro: " & tenro & " encontrada para el ternro: " & ternro & ", fecha: " & Fecha
    Else
        Flog.writeline "Estructura tenro: " & tenro & " NO encontrada para el ternro: " & ternro & ", fecha: " & Fecha
    End If
    
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

    estructura = salida
End Function


Sub encabezado(ByVal archivo As Long, ByVal separador As String, ByRef ostream)
Dim strlinea As String

    Select Case archivo
        Case 0: 'dimension.txt
            strlinea = "Legacy Employee ID"
            strlinea = strlinea & separador & "GLID"
            strlinea = strlinea & separador & "OHR ID"
            strlinea = strlinea & separador & "Employee Record"
            strlinea = strlinea & separador & "Name (Language)"
            strlinea = strlinea & separador & "Last Name"
            strlinea = strlinea & separador & "First Name"
            strlinea = strlinea & separador & "National ID"
            strlinea = strlinea & separador & "Passport Country"
            strlinea = strlinea & separador & "Passport Number"
            strlinea = strlinea & separador & "Date of birth"
            strlinea = strlinea & separador & "Degree of Disability"
            strlinea = strlinea & separador & "Gender"
            strlinea = strlinea & separador & "Nationality"
            strlinea = strlinea & separador & "Foreign Service Code"
            strlinea = strlinea & separador & "Foreign Service Date"
            strlinea = strlinea & separador & "Union Code"
            strlinea = strlinea & separador & "Pay Frequency"
            strlinea = strlinea & separador & "Assignment ID"
            strlinea = strlinea & separador & "Job Title"
            strlinea = strlinea & separador & "Termination Date"
            strlinea = strlinea & separador & "Grade"
            strlinea = strlinea & separador & "Base Salary(Offer)"
            strlinea = strlinea & separador & "Actual Base Salary"
            strlinea = strlinea & separador & "Work Location"
            strlinea = strlinea & separador & "Seniority Date"
            strlinea = strlinea & separador & "Assignment Status"
            strlinea = strlinea & separador & "Employment Type"
            strlinea = strlinea & separador & "Salary Effective Date"
            strlinea = strlinea & separador & "Last Increase Amount"
            strlinea = strlinea & separador & "Salary Increase Reason"
            strlinea = strlinea & separador & "Working Hours"
            strlinea = strlinea & separador & "Hours Change Date"
            strlinea = strlinea & separador & "Leave Start date"
            strlinea = strlinea & separador & "Leave Stop Date"
            strlinea = strlinea & separador & "Leave Reason Code"
            strlinea = strlinea & separador & "Leave Interval (.5/1 days)"
            strlinea = strlinea & separador & "Employment Status Code Date"
            strlinea = strlinea & separador & "Continuity of Service Date - PeSo"
            strlinea = strlinea & separador & "Legacy Payroll ID"
            strlinea = strlinea & separador & "Payroll ID"
            strlinea = strlinea & separador & "Payroll Name"
            strlinea = strlinea & separador & "Personal Account"
            strlinea = strlinea & separador & "Month"
            strlinea = strlinea & separador & "Pay Process Identifier"
            strlinea = strlinea & separador & "Data entry(Payrollteam)"
            strlinea = strlinea & separador & "Proportion"
            strlinea = strlinea & separador & "Salary Basis Type"
            strlinea = strlinea & separador & "Pay Factor"
            strlinea = strlinea & separador & "Gross Amount for whole year"
            strlinea = strlinea & separador & "Folder"
            strlinea = strlinea & separador & "Account"
            strlinea = strlinea & separador & "New Value"
            strlinea = strlinea & separador & "Account Distribution Number"
            strlinea = strlinea & separador & "Fiscal Year Fiscal Week"
            strlinea = strlinea & separador & "Fiscal Year"
            strlinea = strlinea & separador & "Fiscal Week"
            strlinea = strlinea & separador & "Pay End Date"
            strlinea = strlinea & separador & "Fiscal Week Pay End Date"
            strlinea = strlinea & separador & "Fiscal Year Fiscal Week"
            strlinea = strlinea & separador & "Legal Entity Start Date"
            strlinea = strlinea & separador & "Legal Entity End Date"
            strlinea = strlinea & separador & "Legal Entity Location"
            strlinea = strlinea & separador & "Business Name"
            strlinea = strlinea & separador & "Business Group"
            strlinea = strlinea & separador & "Sub Business"
            strlinea = strlinea & separador & "Company Name"
            strlinea = strlinea & separador & "Company Group"
            strlinea = strlinea & separador & "Street"
            strlinea = strlinea & separador & "City / Town"
            strlinea = strlinea & separador & "Zip / Postal Code"
            strlinea = strlinea & separador & "Country"
            strlinea = strlinea & separador & "Business Bank Account Number"
            strlinea = strlinea & separador & "Pension Unit Number"
            strlinea = strlinea & separador & "MBLF CITY1"
            strlinea = strlinea & separador & "MBLF CITY2"
            strlinea = strlinea & separador & "MBLF CITY3"
            strlinea = strlinea & separador & "Employee served"
            strlinea = strlinea & separador & "CLHNPSS HOME CURRENCY %"
            strlinea = strlinea & separador & "CLHNPSS HOST CURRENCY %"
            strlinea = strlinea & separador & "CLHNPSS HOST CURRENCY"
            strlinea = strlinea & separador & "CLHNPSS HOME CURRENCY"
            strlinea = strlinea & separador & "CLHNPMIS COVERAGE TYPE"
            strlinea = strlinea & separador & "Department Code"
            strlinea = strlinea & separador & "Reference Code"
            strlinea = strlinea & separador & "Termination Reason"
            strlinea = strlinea & separador & "Cost Center Name"
            strlinea = strlinea & separador & "CLHNPSS HOST CURRENCY"
            strlinea = strlinea & separador & "CLHNPSS HOME CURRENCY"
            strlinea = strlinea & separador & "CLHNPMIS COVERAGE TYPE"
            strlinea = strlinea & separador & "Department Code"
            strlinea = strlinea & separador & "Reference Code"
            strlinea = strlinea & separador & "Termination Reason"
            strlinea = strlinea & separador & "Cost Center Name"

            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
        Case 1: 'payroll.txt
            strlinea = "VENDOR_PAYROLL_ID"
            strlinea = strlinea & separador & "SSO_ID"
            strlinea = strlinea & separador & "ASSIGNMENT_ID"
            strlinea = strlinea & separador & "LEGAL_ENTITY"
            strlinea = strlinea & separador & "PAYCODE"
            strlinea = strlinea & separador & "HOURS"
            strlinea = strlinea & separador & "AMOUNT"
            strlinea = strlinea & separador & "CURRENCY"
            strlinea = strlinea & separador & "CONVERSION_FACTOR"
            strlinea = strlinea & separador & "AMOUNT_USD"
            strlinea = strlinea & separador & "CHECK ID"
            
            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
    
        Case 2: 'pay_distribution.txt
            strlinea = "VENDOR_PAYROLL_ID"
            strlinea = strlinea & separador & "SSO_ID"
            strlinea = strlinea & separador & "LEGAL_ENTITY"
            strlinea = strlinea & separador & "CHECK SEQ NUMBER"
            strlinea = strlinea & separador & "BANK AC NUMBER"
            strlinea = strlinea & separador & "BANK CODE"
            strlinea = strlinea & separador & "Bank Name"
            strlinea = strlinea & separador & "PAY DISTRIBUTION TYPE"
            strlinea = strlinea & separador & "PAYEE BANK"
            strlinea = strlinea & separador & "CHECK ID"
            strlinea = strlinea & separador & "CURRENCY"
            strlinea = strlinea & separador & "AMOUNT"
            strlinea = strlinea & separador & "CONVERSION_FACTOR"
            strlinea = strlinea & separador & "AMOUNT_USD"
            
            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
    
        Case 3: 'leave.txt
            strlinea = "VENDOR_PAYROLL_ID"
            strlinea = strlinea & separador & "SSO_ID"
            strlinea = strlinea & separador & "LEGAL_ENTITY"
            strlinea = strlinea & separador & "Days Differred OT Available"
            strlinea = strlinea & separador & "Days Holiday Reg Eligible"
            strlinea = strlinea & separador & "Days Holiday Reg Taken"
            strlinea = strlinea & separador & "Days Holiday Floating Eligible"
            strlinea = strlinea & separador & "Days Holiday Floating Taken"
            strlinea = strlinea & separador & "Days Holiday Personal Eligible"
            strlinea = strlinea & separador & "Days Holiday Personal Taken"
            strlinea = strlinea & separador & "Holiday Non-Standard Eligible"
            strlinea = strlinea & separador & "Holiday Non-Standard Taken"
            strlinea = strlinea & separador & "Hours Personal Illness Eligibility"
            strlinea = strlinea & separador & "Hours Personal Illness Remaining"
            strlinea = strlinea & separador & "Hours Personal Business Eligibility"
            strlinea = strlinea & separador & "Hours Personal Business Remaining"
            strlinea = strlinea & separador & "Vacation Hours Carry-over"
            strlinea = strlinea & separador & "Days Vacation Paid"
            strlinea = strlinea & separador & "Days Vacation Remaining"
            strlinea = strlinea & separador & "Days Vacation Taken"
            strlinea = strlinea & separador & "Hours Vacation Average"
            strlinea = strlinea & separador & "Days Vacation Eligible"
            
            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
    
        Case 4: 'leave.txt
            strlinea = "Payroll country"
            strlinea = strlinea & separador & "Pay code"
            strlinea = strlinea & separador & "Paycode type"
            strlinea = strlinea & separador & "Paycode ID"
            strlinea = strlinea & separador & "Paycode sequence"
            
            ostream.WriteText strlinea & Chr(13) & Chr(10) 'Write to the steam
    End Select
End Sub

Function valorConAcu(ByVal ternro As String, ByVal tipo As String, ByVal Fecha As String, ByVal codigo As String, ByVal separadorDecimal As String)
Dim rsValorLiq As New ADODB.Recordset
Dim salida As String

'busco la liquidacion de tipo modelo mensual (3)

Select Case UCase(tipo)
    Case "COM", "COC"
        StrSql = " SELECT dlimonto, dlicant FROM cabliq " & _
                 " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro " & _
                 " INNER JOIN concepto on concepto.concnro = detliq.concnro AND concepto.conccod = '" & codigo & "'" & _
                 " INNER JOIN proceso on proceso.pronro = cabliq.pronro AND proceso.tprocnro = 3 " & _
                 " INNER JOIN his_estructura on his_estructura.ternro = cabliq.empleado and tenro = 10 " & _
                 " INNER JOIN empresa on his_estructura.estrnro = empresa.estrnro and proceso.empnro = empresa.empnro " & _
                 " WHERE profecini <= " & ConvFecha(Fecha) & " AND profecfin >= " & ConvFecha(Fecha) & " and empleado = " & ternro & _
                 " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta >= " & ConvFecha(Fecha) & " or htethasta is null) "
        OpenRecordset StrSql, rsValorLiq
        If Not rsValorLiq.EOF Then
            If UCase(tipo) = "COM" Then
                salida = IIf(EsNulo(rsValorLiq!dlimonto), 0, rsValorLiq!dlimonto)
            End If
            
            If UCase(tipo) = "COC" Then
                salida = IIf(EsNulo(rsValorLiq!dlicant), 0, rsValorLiq!dlicant)
            End If
            Flog.writeline "Valor concepto : " & codigo & " encontrado para el ternro: " & ternro & ", fecha: " & Fecha
        Else
            salida = 0
            Flog.writeline "Valor concepto : " & codigo & " No encontrada para el ternro: " & ternro & ", fecha: " & Fecha
        End If
    Case "ALC", "ALM"
        StrSql = " SELECT acu_liq.alcant, acu_liq.almonto FROM cabliq " & _
                 " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro " & _
                 " INNER JOIN acumulador on acumulador.acunro = acu_liq.acunro AND acu_liq.acunro = " & codigo & _
                 " INNER JOIN proceso on proceso.pronro = cabliq.pronro AND proceso.tprocnro = 3 " & _
                 " INNER JOIN his_estructura on his_estructura.ternro = cabliq.empleado and tenro = 10 " & _
                 " INNER JOIN empresa on his_estructura.estrnro = empresa.estrnro and proceso.empnro = empresa.empnro " & _
                 " Where profecini <= " & ConvFecha(Fecha) & " And profecfin >= " & ConvFecha(Fecha) & " And Empleado = " & ternro & _
                 " And htetdesde <= " & ConvFecha(Fecha) & " And (htethasta >= " & ConvFecha(Fecha) & " Or htethasta Is Null) "
        OpenRecordset StrSql, rsValorLiq
        If Not rsValorLiq.EOF Then
            If UCase(tipo) = "ALM" Then
                salida = IIf(EsNulo(rsValorLiq!almonto), 0, rsValorLiq!almonto)
            End If
            
            If UCase(tipo) = "ALC" Then
                salida = IIf(EsNulo(rsValorLiq!alcant), 0, rsValorLiq!alcant)
            End If
            Flog.writeline "Valor Acumulador: " & codigo & " encontrad0 para el ternro: " & ternro & ", fecha: " & Fecha
        Else
            salida = 0
            Flog.writeline "Valor v : " & codigo & " No encontrad0 para el ternro: " & ternro & ", fecha: " & Fecha
        End If
    End Select
    
    If rsValorLiq.State = adStateOpen Then rsValorLiq.Close
    Set rsValorLiq = Nothing

    salida = Replace(FormatNumber(salida, 2), ",", "")
    salida = Replace(salida, ".", "")
    valorConAcu = Left(salida, Len(salida) - 2) & separadorDecimal & Right(salida, 2)

End Function

Function empleadoActivo(ByVal ternro As String, ByVal Fecha As String)
Dim salida As String
Dim rsAux As New ADODB.Recordset
    
    StrSql = " SELECT * FROM fases where empleado = " & ternro & _
             " AND altfec <= " & ConvFecha(Fecha) & " AND (bajfec >= " & ConvFecha(Fecha) & " or bajfec is null)"
    OpenRecordset StrSql, rsAux
    If Not rsAux.EOF Then
        salida = "Activo"
    Else
        salida = "Inactivo"
    End If
    
    If rsAux.State = adStateOpen Then rsAux.Close
    Set rsAux = Nothing

    empleadoActivo = salida

End Function

Function novestr(ByVal tenro As String, ByVal ternro As String, ByVal Fecha As String, ByVal conccod As String, ByVal tpanro As String)
Dim salida As String
Dim rsAux As New ADODB.Recordset

    StrSql = " SELECT ntevalor FROM novestr " & _
             " INNER JOIN his_estructura ON his_estructura.estrnro = novestr.estrnro AND his_estructura.tenro = " & tenro & _
             " INNER JOIN concepto ON concepto.concnro = novestr.concnro AND concepto.conccod = '" & conccod & "'" & _
             " WHERE novestr.tenro = " & tenro & " And his_estructura.ternro = " & ternro & _
             " AND (htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta >= " & ConvFecha(Fecha) & " or htethasta is null)) " & _
             " AND (ntedesde is null or ((ntedesde <= '22/08/2014' AND (ntehasta >= '22/08/2014' or ntehasta is null)) )) " & _
             " AND novestr.tpanro = " & tpanro
    OpenRecordset StrSql, rsAux
    If Not rsAux.EOF Then
        salida = rsAux!ntevalor
        Flog.writeline "Novedad por estructura encontrada para el ternro: " & ternro & ", fecha: " & Fecha & ", concepto: " & conccod & ", parametro: " & tpanro
    Else
        salida = 0
        Flog.writeline "Novedad por estructura NO encontrada para el ternro: " & ternro & ", fecha: " & Fecha & ", concepto: " & conccod & ", parametro: " & tpanro
    End If

    If rsAux.State = adStateOpen Then rsAux.Close
    Set rsAux = Nothing

    novestr = FormatNumber(salida, 2)

End Function

Function licencia(ByVal ternro As String, ByVal Fecha As String, ByVal tdnro As String, ByVal campo As String)
Dim salida As String
Dim rsAux As New ADODB.Recordset

    StrSql = " SELECT tipdia.tddesc, emp_lic.elfechadesde, emp_lic.elfechahasta FROM emp_lic " & _
             " INNER JOIN tipdia on tipdia.tdnro = emp_lic.tdnro " & _
             " WHERE elfechadesde <= " & ConvFecha(Fecha) & " And elfechahasta >= " & ConvFecha(Fecha) & " And Empleado = " & ternro & _
             " AND emp_lic.tdnro in (" & tdnro & ") "

    OpenRecordset StrSql, rsAux
    If Not rsAux.EOF Then
        Flog.writeline "Licencia encontrada para el ternro: " & ternro & ", fecha: " & Fecha
        Select Case UCase(campo)
            Case "TIPDIA.TDDESC":
                salida = rsAux!tddesc
            Case "ELFECHADESDE":
                salida = rsAux!elfechadesde
            Case "ELFECHAHASTA":
                salida = rsAux!elfechahasta
        
        End Select
    Else
        salida = ""
        Flog.writeline "Licencia NO encontrada para el ternro: " & ternro & ", fecha: " & Fecha
    End If

    If rsAux.State = adStateOpen Then rsAux.Close
    Set rsAux = Nothing

    licencia = salida

End Function

Function datoProceso(ByVal ternro As String, ByVal Fecha As String, ByVal campo As String)
Dim salida As String
Dim rsAux As New ADODB.Recordset

    StrSql = " SELECT proceso.prodesc, proceso.profecpago, proceso.profecini FROM cabliq " & _
             " INNER JOIN proceso on proceso.pronro = cabliq.pronro AND proceso.tprocnro = 3 " & _
             " INNER JOIN his_estructura on his_estructura.ternro = cabliq.empleado and tenro = 10 " & _
             " INNER JOIN empresa on his_estructura.estrnro = empresa.estrnro and proceso.empnro = empresa.empnro " & _
             " WHERE profecini <= " & ConvFecha(Fecha) & "  AND profecfin >= " & ConvFecha(Fecha) & "  and empleado = " & ternro & _
             " AND htetdesde <= " & ConvFecha(Fecha) & "  AND (htethasta >= " & ConvFecha(Fecha) & "  or htethasta is null)"
    OpenRecordset StrSql, rsAux
    If Not rsAux.EOF Then
        Flog.writeline "Proceso mensual encontrado para el ternro: " & ternro & ", fecha: " & Fecha
        Select Case UCase(campo)
            Case "PRODESC"
                salida = rsAux!prodesc
            Case "PROFECPAGO"
                salida = rsAux!profecpago
            Case "PROFECINI"
                salida = rsAux!PROFECINI
            
            Case Else
                salida = ""
        End Select
    Else
        salida = ""
        Flog.writeline "Proceso mensual NO encontrado para el ternro: " & ternro & ", fecha: " & Fecha
    End If
    If rsAux.State = adStateOpen Then rsAux.Close
    Set rsAux = Nothing

    datoProceso = salida

End Function

Function domcilioEstructura(ByVal ternro As String, ByVal tenro As String, ByVal Fecha As String, ByVal campo As String)
Dim salida As String
Dim rsAux As New ADODB.Recordset

    StrSql = " SELECT detdom.calle, detdom.nro, localidad.locdesc, detdom.codigopostal, pais.paisdesc  FROM his_estructura " & _
             " INNER JOIN sucursal ON sucursal.estrnro = his_estructura.estrnro " & _
             " INNER JOIN cabdom on cabdom.ternro = sucursal.ternro AND cabdom.domdefault = -1 " & _
             " INNER JOIN detdom on detdom.domnro = cabdom.domnro " & _
             " LEFT JOIN localidad on localidad.locnro = detdom.locnro " & _
             " LEFT JOIN pais on pais.paisnro = detdom.paisnro " & _
             " WHERE tenro = " & tenro & " And his_estructura.ternro = " & ternro & _
             " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta >= " & ConvFecha(Fecha) & " or htethasta is null) "
    OpenRecordset StrSql, rsAux
    If Not rsAux.EOF Then
        Flog.writeline "Domicilio del tipo de estructura: " & tenro & ", ternro: " & ternro & " encontrada, fecha: " & Fecha
        Select Case UCase(campo)
            Case "CALLE"
                salida = rsAux!calle & " " & rsAux!nro
            Case "LOCDESC"
                salida = IIf(IsNull(rsAux!locdesc), "", rsAux!locdesc)
            Case "CP"
                salida = IIf(IsNull(rsAux!codigopostal), "", rsAux!codigopostal)
            Case "PAIS"
                salida = IIf(IsNull(rsAux!paisdesc), "", rsAux!paisdesc)
            Case "ALL"
                salida = rsAux!calle
                salida = salida & IIf(IsNull(rsAux!locdesc), "", " " & rsAux!locdesc)
                salida = salida & IIf(IsNull(rsAux!codigopostal), "", " " & rsAux!codigopostal)
                salida = salida & IIf(IsNull(rsAux!paisdesc), "", " " & rsAux!paisdesc)
        End Select
    Else
        salida = ""
        Flog.writeline "Domicilio del tipo de estructura: " & tenro & ", ternro: " & ternro & " NO encontrada, fecha: " & Fecha
    End If
    If rsAux.State = adStateOpen Then rsAux.Close
    Set rsAux = Nothing

    domcilioEstructura = salida

End Function

Function ctabancaria(ByVal ternro As String, ByVal tenro As String, ByVal Fecha As String)
Dim salida As String
Dim rsAux As New ADODB.Recordset
    
    salida = ""
    StrSql = " SELECT ctabancaria.ctabnro FROM his_estructura  " & _
             " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro " & _
             " INNER JOIN ctabancaria on ctabancaria.ternro = empresa.ternro AND ctabancaria.ctabestado = -1 " & _
             " WHERE tenro = " & tenro & " And his_estructura.ternro = " & ternro & _
             " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta >= " & ConvFecha(Fecha) & " or htethasta is null)"
    OpenRecordset StrSql, rsAux
    If Not rsAux.EOF Then
        salida = rsAux!ctabnro
        Flog.writeline "Cuenta bancaria de la empresa del ternro: " & ternro & " encontrada, fecha: " & Fecha
    Else
        Flog.writeline "Cuenta bancaria de la empresa del ternro: " & ternro & " NO encontrada, fecha: " & Fecha
    End If
    
    If rsAux.State = adStateOpen Then rsAux.Close
    Set rsAux = Nothing

    ctabancaria = salida

End Function

Function valorConAcuLista(ByVal ternro As String, ByVal tipo As String, ByVal Fecha As String, ByVal codigo As String, ByVal separadorDecimal As String)
Dim rsValorLiq As New ADODB.Recordset
Dim salida As String

'busco la liquidacion de tipo modelo mensual (3)
salida = 0
Select Case UCase(tipo)
    Case "COM", "COC"
        StrSql = " SELECT sum(dlimonto) monto, sum(dlicant) cant FROM cabliq " & _
                 " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro " & _
                 " INNER JOIN concepto on concepto.concnro = detliq.concnro AND concepto.conccod in (" & codigo & ")" & _
                 " INNER JOIN proceso on proceso.pronro = cabliq.pronro AND proceso.tprocnro = 3 " & _
                 " INNER JOIN his_estructura on his_estructura.ternro = cabliq.empleado and tenro = 10 " & _
                 " INNER JOIN empresa on his_estructura.estrnro = empresa.estrnro and proceso.empnro = empresa.empnro " & _
                 " WHERE profecini <= " & ConvFecha(Fecha) & " AND profecfin >= " & ConvFecha(Fecha) & " and empleado = " & ternro & _
                 " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta >= " & ConvFecha(Fecha) & " or htethasta is null) "
        OpenRecordset StrSql, rsValorLiq
        If Not rsValorLiq.EOF Then
            If UCase(tipo) = "COM" Then
                salida = IIf(EsNulo(rsValorLiq!Monto), 0, rsValorLiq!Monto)
                Flog.writeline "Conceptos: " & codigo & " encontrado para el ternro: " & ternro & ", fecha: " & Fecha
            End If
            
            If UCase(tipo) = "COC" Then
                salida = IIf(EsNulo(rsValorLiq!cant), 0, rsValorLiq!cant)
                Flog.writeline "Conceptos: " & codigo & " encontrado para el ternro: " & ternro & ", fecha: " & Fecha
            End If
        Else
            Flog.writeline "Conceptos: " & codigo & " NO encontrado para el ternro: " & ternro & ", fecha: " & Fecha
        End If
    Case "ALC", "ALM"
        StrSql = " SELECT sum(acu_liq.alcant) cant, sum(acu_liq.almonto) monto FROM cabliq " & _
                 " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro " & _
                 " INNER JOIN acumulador on acumulador.acunro = acu_liq.acunro AND acu_liq.acunro in (" & codigo & ")" & _
                 " INNER JOIN proceso on proceso.pronro = cabliq.pronro AND proceso.tprocnro = 3 " & _
                 " INNER JOIN his_estructura on his_estructura.ternro = cabliq.empleado and tenro = 10 " & _
                 " INNER JOIN empresa on his_estructura.estrnro = empresa.estrnro and proceso.empnro = empresa.empnro " & _
                 " Where profecini <= " & ConvFecha(Fecha) & " And profecfin >= " & ConvFecha(Fecha) & " And Empleado = " & ternro & _
                 " And htetdesde <= " & ConvFecha(Fecha) & " And (htethasta >= " & ConvFecha(Fecha) & " Or htethasta Is Null) "
        OpenRecordset StrSql, rsValorLiq
        If Not rsValorLiq.EOF Then
            If UCase(tipo) = "ALM" Then
                salida = IIf(EsNulo(rsValorLiq!Monto), 0, rsValorLiq!Monto)
                Flog.writeline "Acumuladores: " & codigo & " encontrado para el ternro: " & ternro & ", fecha: " & Fecha
            End If
            
            If UCase(tipo) = "ALC" Then
                salida = IIf(EsNulo(rsValorLiq!cant), 0, rsValorLiq!cant)
                Flog.writeline "Acumuladores: " & codigo & " encontrado para el ternro: " & ternro & ", fecha: " & Fecha
            End If
        Else
            Flog.writeline "Acumuladores: " & codigo & " NO encontrado para el ternro: " & ternro & ", fecha: " & Fecha
        End If
        
    Case "AMM", "AMC"
        StrSql = " SELECT sum(amcant) cant, sum(ammonto) monto FROM acu_mes" & _
                 " WHERE ammes = " & Month(Fecha) & " AND amanio = " & Year(Fecha) & _
                 " AND ternro = " & ternro & " AND acu_mes.acunro in (" & codigo & ") "
        OpenRecordset StrSql, rsValorLiq
        If Not rsValorLiq.EOF Then
            If UCase(tipo) = "AMM" Then
                salida = IIf(EsNulo(rsValorLiq!Monto), 0, rsValorLiq!Monto)
                Flog.writeline "Acumuladores mensuales: " & codigo & " encontrado para el ternro: " & ternro & ", mes: " & Month(Fecha) & ", año: " & Year(Fecha)
            End If
            
            If UCase(tipo) = "AMC" Then
                salida = IIf(EsNulo(rsValorLiq!cant), 0, rsValorLiq!cant)
                Flog.writeline "Acumuladores mensuales: " & codigo & " encontrado para el ternro: " & ternro & ", mes: " & Month(Fecha) & ", año: " & Year(Fecha)
            End If
        Else
            Flog.writeline "Acumuladores mensual: " & codigo & " no encontrado para el ternro: " & ternro & ", mes: " & Month(Fecha) & ", año: " & Year(Fecha)
        End If
    End Select
    
    If rsValorLiq.State = adStateOpen Then rsValorLiq.Close
    Set rsValorLiq = Nothing
    
    salida = Replace(FormatNumber(salida, 2), ",", "")
    salida = Replace(salida, ".", "")
    valorConAcuLista = Left(salida, Len(salida) - 2) & separadorDecimal & Right(salida, 2)

End Function

Function ctabancariaEmpleado(ByVal ternro As String)
Dim rs_consult As New ADODB.Recordset
Dim salida As String
    
    salida = ""
    StrSql = " SELECT ctabnro FROM ctabancaria WHERE ctabestado = -1 AND ternro = " & ternro
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        salida = rs_consult!ctabnro
        Flog.writeline "Cuenta bancaria activa encontrada para el ternro: " & ternro
    Else
        Flog.writeline "Cuenta bancaria activa no encontrada para el ternro: " & ternro
    End If
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

    ctabancariaEmpleado = salida
End Function

Function codigoExtEstructura(ByVal ternro As String, ByVal tenro As String, ByVal Fecha As String)
Dim salida As String
Dim rs_consult As New ADODB.Recordset
Dim StrSql As String

    salida = ""
    StrSql = " SELECT estrcodext FROM his_estructura  " & _
             " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
             " WHERE his_estructura.ternro = " & ternro & " And his_estructura.tenro = " & tenro & _
             " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha) & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        If Not IsNull(rs_consult!estrcodext) Then
            salida = rs_consult!estrcodext
            Flog.writeline "Codigo externo de estructura encontrado para el ternro: " & ternro & ", tenro: " & tenro & ", fecha: " & Fecha
        Else
            Flog.writeline "Codigo externo nulo para el ternro: " & ternro & ", tenro: " & tenro & ", fecha: " & Fecha
        End If
    Else
        Flog.writeline "Codigo externo de estructura no encontrado para el ternro: " & ternro & ", tenro: " & tenro & ", fecha: " & Fecha
    End If
    
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

    codigoExtEstructura = salida
End Function

Function valorAcuAnual(ByVal ternro As String, ByVal anioLiq As String, ByVal acuAnual As String, ByVal separadorDecimal As String)
Dim salida As String
Dim rs_consult As New ADODB.Recordset
Dim StrSql As String
    
    salida = "0"
    
    StrSql = " SELECT sum(acu_liq.alcant) alcant, sum(acu_liq.almonto) almonto FROM periodo " & _
             " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
             " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
             " INNER JOIN tercero ON cabliq.empleado = tercero.ternro and ternro =  " & ternro & _
             " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
             " WHERE periodo.pliqanio = " & anioLiq & " And acu_liq.acunro = " & acuAnual
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        Flog.writeline "Acumulador anual encontrado para el ternro : " & ternro & ", anio: " & anioLiq
        salida = IIf(EsNulo(rs_consult!almonto), 0, rs_consult!almonto)
    End If
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

    salida = Replace(FormatNumber(salida, 2), ",", "")
    salida = Replace(salida, ".", "")
    valorAcuAnual = Left(salida, Len(salida) - 2) & separadorDecimal & Right(salida, 2)

End Function

Function periodo(ByVal ternro As String, ByVal Fecha As String, ByVal campo As String)
Dim salida As String
Dim rs_consult As New ADODB.Recordset
Dim StrSql As String

    salida = ""
    StrSql = " SELECT periodo.pliqdesde, periodo.pliqhasta FROM cabliq " & _
             " INNER JOIN proceso on proceso.pronro = cabliq.pronro AND proceso.tprocnro = 3 " & _
             " INNER JOIN his_estructura on his_estructura.ternro = cabliq.empleado and tenro = 10 " & _
             " INNER JOIN empresa on his_estructura.estrnro = empresa.estrnro and proceso.empnro = empresa.empnro " & _
             " INNER JOIN periodo on periodo.pliqnro = proceso.pliqnro " & _
             " WHERE profecini <= " & ConvFecha(Fecha) & " AND profecfin >= " & ConvFecha(Fecha) & " and empleado = " & ternro & _
             " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta >= " & ConvFecha(Fecha) & " or htethasta is null) "
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        Flog.writeline "Periodo encontrado para el ternro : " & ternro
        Select Case UCase(campo)
            Case UCase("pliqdesde")
                salida = rs_consult!pliqdesde
            Case UCase("pliqhasta")
                salida = rs_consult!pliqhasta
        End Select
    Else
        Flog.writeline "Periodo no encontrado para el ternro : " & ternro
    End If
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

    periodo = salida

End Function

Function empresaAbrev(ByVal ternro As String, ByVal Fecha As String)
Dim salida As String
Dim rs_consult As New ADODB.Recordset
Dim StrSql As String

    salida = ""
    StrSql = " SELECT terabr FROM his_estructura " & _
             " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
             " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro " & _
             " INNER JOIN tercero ON tercero.ternro = empresa.ternro " & _
             " WHERE his_estructura.ternro = " & ternro & " And his_estructura.tenro = 10 " & _
             " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha) & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR his_estructura.htethasta is null)) "
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        salida = rs_consult!terabr
        Flog.writeline "Campo Abreviado para empresa encontrada para el ternro: " & ternro & ", fecha: " & Fecha
    Else
        Flog.writeline "Campo Abreviado para empresa No encontrada para el ternro: " & ternro & ", fecha: " & Fecha
    End If
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing

    empresaAbrev = salida

End Function
