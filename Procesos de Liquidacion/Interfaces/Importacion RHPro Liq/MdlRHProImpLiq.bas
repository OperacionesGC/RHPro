Attribute VB_Name = "MdlRHProImpLiq"
Option Explicit

'Global Const Version = "1.00" 'Importacion de datos de Liquidacion
'Global Const FechaModificacion = "08/11/2008"
'Global Const UltimaModificacion = "" 'Martin Ferraro - Version Inicial

Global Const Version = "1.01" 'Importacion de datos de Liquidacion
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = "" ''MB - Encriptacion de string connection


Global Seed As String 'Usado como clave de encriptacion/desencriptacion
Global encryptAct As Boolean
Global Sep As String
Global CantRegErr As Long
Global ArchOpen As Boolean

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Importacion.
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
Dim bprcfecha As Date
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error Resume Next
    'Abro la conexion
'    OpenConnection strconexion, objConn
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
'    OpenConnection strconexion, objconnProgreso
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
'    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Imp. Liquidacion RHPro" & " - " & NroProcesoBatch & ".log"
    
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
    
    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 231 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        bprcfecha = rs_batch_proceso!bprcfecha
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call ImportLiq(NroProcesoBatch, bprcparam, bprcfecha)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100  WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
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


Public Sub ImportLiq(ByVal bpronro As Long, ByVal Parametros As String, ByVal bprcfecha As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Importacion de datos de liquidacion
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
Dim mes As Integer
Dim anio As Long
Dim Directorio As String
Dim SeparadorDecimal As String
Dim DescripcionModelo As String
Dim Archivo As String
Dim fImport
Dim Carpeta
Dim Linea As String
Dim ArrSQL(30) As String
Dim CantReg As Long
Dim strLineaArch As String
Dim strLinea As String
Dim tipoLinea As Integer
Dim tipoLineaAnt As Integer
Dim cambioTipo As Boolean
Dim Corrimiento As Long
Dim Folder
Dim File
Dim fileItem
Dim listaBatch As String
Dim Destino As String

'-------------------------------------------------------------------------------------------------
'RecordSets
'-------------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset

'Inicio codigo ejecutable
On Error GoTo E_ImportLiq

'Valores default de encriptacion: Activa y semilla = 56238
Seed = "56238"
encryptAct = True
HuboError = False
ArchOpen = False
'-------------------------------------------------------------------------------------------------
'Levanto cada parametro por separado, el separador de parametros es "@"
'-------------------------------------------------------------------------------------------------
Flog.writeline
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    If Len(Parametros) >= 1 Then
    
        ArrPar = Split(Parametros, "@")
        
        If UBound(ArrPar) >= 0 Then
            
            Auto = CBool(ArrPar(0))
            
            If Not Auto Then
                Archivo = ArrPar(1)
                Flog.writeline Espacios(Tabulador * 0) & "Disparo de Importacion Manual Archivo " & Archivo
            Else
                Flog.writeline Espacios(Tabulador * 0) & "Disparo de Importacion Planificada"
            End If
            
        Else
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Faltan parametros."
            HuboError = True
            Exit Sub
        End If
        
    Else
        Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontraron los parametros."
        HuboError = True
        Exit Sub
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontraron los parametros."
    HuboError = True
    Exit Sub
End If
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Configuracion del Reporte
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuracion del reporte."
StrSql = "SELECT * FROM confrep WHERE repnro = 246 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult
If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte 246."
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
                        mes = Month(Fecha)
                        anio = Year(Fecha)
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
'Configuracion del Directorio de salida
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando directorio de entrada."
StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    Directorio = Trim(rs_Consult!sis_direntradas)
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
StrSql = "SELECT * FROM modelo WHERE modnro = 316"
OpenRecordset StrSql, rs_Consult
Sep = ""
If Not rs_Consult.EOF Then
    Directorio = Directorio & Trim(rs_Consult!modarchdefault)
    Sep = IIf(Not IsNull(rs_Consult!modseparador), rs_Consult!modseparador, "")
    SeparadorDecimal = IIf(Not IsNull(rs_Consult!modsepdec), rs_Consult!modsepdec, ".")
    DescripcionModelo = rs_Consult!moddesc
    
    Flog.writeline Espacios(Tabulador * 1) & "Modelo 316 " & rs_Consult!moddesc
    Flog.writeline Espacios(Tabulador * 1) & "Directorio de Importacion : " & Directorio
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el modelo 316."
    HuboError = True
    Exit Sub
End If
Flog.writeline
     
        
'-------------------------------------------------------------------------------------------------
'Busqueda del archivo en caso de disparo automatico
'-------------------------------------------------------------------------------------------------
Set fs = CreateObject("Scripting.FileSystemObject")

If Auto Then
    Flog.writeline Espacios(Tabulador * 0) & "Buscando archivo a procesar."

    'Seteo el nombre del archivo generado
    On Error Resume Next
    
    'Busco Directorio
    Set Folder = fs.GetFolder(Directorio)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el directorio " & Directorio
        HuboError = True
        Exit Sub
    End If
    
    'Busco el primer archivo con extension txt
    Set File = Folder.Files
    Archivo = ""
    For Each fileItem In File
        If fs.GetExtensionName(Directorio & "\" & fileItem.Name) = "txt" Then
            Archivo = fileItem.Name
            Flog.writeline Espacios(Tabulador * 1) & "Archivo pendiente de procesamiento encontrado " & Archivo
            Exit For
        End If
    Next
    
    If Len(Archivo) = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró ningun archivo a procesar (.txt)"
        HuboError = True
        Exit Sub
    End If
    
    Set File = Nothing
    On Error GoTo E_ImportLiq
End If
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Renombro el archivo para que no lo tome otro proceso
'-------------------------------------------------------------------------------------------------
fs.MoveFile Directorio & "\" & Archivo, Directorio & "\" & Left(Archivo, Len(Archivo) - 4) & ".prc"
Archivo = Left(Archivo, Len(Archivo) - 4) & ".prc"


'-------------------------------------------------------------------------------------------------
'Apertura del archivo
'-------------------------------------------------------------------------------------------------
On Error Resume Next
Flog.writeline Espacios(Tabulador * 0) & "Buscando el archivo " & Archivo
'Archivo = Directorio & "\" & Archivo
Set fImport = fs.OpenTextFile(Directorio & "\" & Archivo, 1, 0)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "Error. No se encontro el archivo " & Archivo
    HuboError = True
    Exit Sub
End If
On Error GoTo E_ImportLiq
Flog.writeline

'Marco que abri el archivo
ArchOpen = True

'-------------------------------------------------------------------------------------------------
'Calculo la cantidad de lineas
'-------------------------------------------------------------------------------------------------
CantReg = 0
Do While Not fImport.AtEndOfStream
    strLineaArch = fImport.ReadLine
    CantReg = CantReg + 1
Loop
fImport.Close


'-------------------------------------------------------------------------------------------------
'Seteo de las variables de progreso
'-------------------------------------------------------------------------------------------------
Progreso = 0
CEmpleadosAProc = CantReg
Flog.writeline
If CEmpleadosAProc = 0 Then
    Flog.writeline Espacios(Tabulador * 0) & "No hay Datos a Importar"
    CEmpleadosAProc = 1
Else
    Flog.writeline Espacios(Tabulador * 0) & "Cantidad de Registros a Importar: " & CEmpleadosAProc
End If
IncPorc = (100 / CEmpleadosAProc)
Flog.writeline

        
Set fImport = fs.OpenTextFile(Directorio & "\" & Archivo, 1, 0)

Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "INICIO DE LECTURA DE LINEAS DEL ARCHIVO"
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------------------"
'-------------------------------------------------------------------------------------------------
'Empleados para los mapeos (CODIGO 1)
'-------------------------------------------------------------------------------------------------
tipoLineaAnt = 0
Do While Not fImport.AtEndOfStream
    
    'Leo la linea del archivo
    strLineaArch = fImport.ReadLine
    
    'Aplico desencriptacion
    strLinea = DesEncriptar(strLineaArch)
   'strLinea = strLineaArch
    
    'Busco el primer el elemento de la linea que determina de que tabla es
    tipoLinea = CInt(Left(strLinea, InStr(1, strLinea, Sep) - 1))
    
    'Verifico si cambio el tipo de linea
    cambioTipo = (tipoLinea <> tipoLineaAnt)
    If cambioTipo Then
        
        If tipoLineaAnt <> 0 Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "Cant. de Reg. Procesados = " & CantReg
            Flog.writeline Espacios(Tabulador * 0) & "Cant. de Reg. con Error  = " & CantRegErr
            Flog.writeline
        End If
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "PROCESANDO Reg de tipo " & tipoLinea & " - " & nombre(tipoLinea)
        Flog.writeline Espacios(Tabulador * 0) & "------------------------------------"
        
        'BORRAR TABLA
        Select Case tipoLinea
            Case 1 To 9:
                Call BorrarTabla(tipoLinea)
            Case 21:
                Call BorrarBatch(strLinea, listaBatch)
            Case 22:
                Call BorrarRecibo(listaBatch)
            Case 24:
                Call BorrarLibro(listaBatch)
            Case 27:
                Call BorrarF649(listaBatch)
            Case 28 To 30:
                Call BorrarTabla(tipoLinea)
        End Select
        
        tipoLineaAnt = tipoLinea
        CantReg = 0
        CantRegErr = 0
    End If
    
    
    'De acuerdo al tipo de linea llamo al correspondiente procedimiento
    Select Case tipoLinea
        Case 1:
            Call procEmpleados(strLinea)
        Case 2:
            Call procTipoConc(strLinea)
        Case 3:
            Call procConceptos(strLinea)
        Case 4:
            Call procTipoAcu(strLinea)
        Case 5:
            Call procAcumuladores(strLinea)
        Case 6:
            Call procItems(strLinea)
        Case 7:
            Call procTipoProc(strLinea)
        Case 8:
            Call procProcesos(strLinea)
        Case 9:
            Call procPeriodos(strLinea)
        Case 10:
            Call procCabLiq(strLinea)
        Case 11:
            Call procDetliq(strLinea)
        Case 12:
            Call procAcu_liq(strLinea)
        Case 13:
            Call procAcu_mes(strLinea)
        Case 14:
            Call procDesliq(strLinea)
        Case 15:
            Call procFicharet(strLinea)
        Case 16:
            Call procImpproarg(strLinea)
        Case 17:
            Call procImpmesarg(strLinea)
        Case 18:
            Call procTraza_gan(strLinea)
        Case 19:
            Call procTraza_gan_item_top(strLinea)
        Case 20:
            Call procDesmen(strLinea)
        Case 21:
            Call procBatch_proceso(strLinea)
        Case 22:
            Call procRep_recibo(strLinea)
        Case 23:
            Call procRep_recibo_det(strLinea)
        Case 24:
            Call procRep_libroley(strLinea)
        Case 25:
            Call procRep_libroley_det(strLinea)
        Case 26:
            Call procRep_libroley_fam(strLinea)
        Case 27:
            Call procF649(strLinea)
        Case 28:
            Call procEscala_ded(strLinea)
        Case 29:
            Call procConfrep(strLinea)
        Case 30:
            Call procEscala(strLinea)
    End Select
    
    CantReg = CantReg + 1
    
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & " , bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & " ' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

Loop

'Para el ultimo registro
If CantReg <> 0 Then
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Cant. de Reg. Procesados = " & CantReg
    Flog.writeline Espacios(Tabulador * 0) & "Cant. de Reg. con Error  = " & CantRegErr
    Flog.writeline
End If


fImport.Close
ArchOpen = False

'Muevo el archivo a la carpeta de backup o error dependiendo si lo proceso exitosamente o no
Set fImport = fs.getfile(Directorio & "\" & Archivo)

'Seteo el destino despendiendo si hubo algun error o no
If HuboError Then
    Destino = Directorio & "\Err\"
Else
    Destino = Directorio & "\bk\"
End If
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Moviendo el archivo a la carpeta " & Destino
'Muevo el archivo creando la carpeta respectiva si no existe
On Error Resume Next
    fImport.Move Destino & Archivo
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "La carpeta Destino no existe. Se creará."
        Set Carpeta = fs.CreateFolder(Destino)
        fImport.Move Destino & Archivo
    End If
On Error GoTo E_ImportLiq:
Flog.writeline

If rs_Consult.State = adStateOpen Then rs_Consult.Close

Set rs_Consult = Nothing
Set fImport = Nothing
Set fs = Nothing

Exit Sub

E_ImportLiq:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: ImportLiq"
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
    
'------------------------------------------------------------------
'Movimiento del archivo
'------------------------------------------------------------------
    If ArchOpen Then
        'Cierro el archivo
        fImport.Close
        
        'Muevo el archivo a la carpeta de backup o error dependiendo si lo proceso exitosamente o no
        Set fImport = fs.getfile(Directorio & "\" & Archivo)
        
        'Seteo el destino despendiendo si hubo algun error o no
        Destino = Directorio & "\Err\"
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "Moviendo el archivo a la carpeta " & Destino
        'Muevo el archivo creando la carpeta respectiva si no existe
        On Error Resume Next
            fImport.Move Destino & Archivo
            If Err.Number <> 0 Then
                Flog.writeline Espacios(Tabulador * 0) & "La carpeta Destino no existe. Se creará."
                Set Carpeta = fs.CreateFolder(Destino)
                fImport.Move Destino & Archivo
            End If
        On Error GoTo E_ImportLiq:
        Flog.writeline
    
    End If
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub


Public Function DesEncriptar(ByVal Valor As String)
    
    If encryptAct Then
        DesEncriptar = Decrypt(Seed, Valor)
    Else
        DesEncriptar = Valor
    End If
    
End Function


Public Function CtrlNuloTXT(ByVal Valor) As String

    If IsNull(Valor) Then
        CtrlNuloTXT = "NULL"
    Else
        If UCase(Valor) = "NULL" Then
            CtrlNuloTXT = "NULL"
        Else
            CtrlNuloTXT = "'" & Valor & "'"
        End If
    End If
    
End Function


Public Sub procEmpleados(ByVal Linea As String)
Dim arrLinea
Dim ternro As String
Dim empleg As String
    
On Error GoTo E_procEmpleados

    'Formato tipo, ternro, empleg
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 2 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    ternro = arrLinea(1)
    empleg = arrLinea(2)
    
    'Inserto Datos
    StrSql = "INSERT INTO tmp_empleado("
    StrSql = StrSql & " ternro, empleg)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & ternro
    StrSql = StrSql & " ," & empleg & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

E_procEmpleados:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procEmpleados"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procTipoConc(ByVal Linea As String)
Dim arrLinea
Dim tconnro As String
Dim Tcondesc As String
Dim sistema As String
    
On Error GoTo E_procTipoConc

    'Formato tipo, tconnro, tcondesc, sistema
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 3 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    tconnro = arrLinea(1)
    Tcondesc = arrLinea(2)
    sistema = arrLinea(3)
    
    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT tipconcep ON"
    StrSql = StrSql & " INSERT INTO tipconcep("
    StrSql = StrSql & " tconnro,tcondesc,sistema)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & tconnro
    StrSql = StrSql & " ," & CtrlNuloTXT(Tcondesc)
    StrSql = StrSql & " ," & sistema & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT tipconcep OFF"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procTipoConc:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procTipoConc"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procConceptos(ByVal Linea As String)
Dim arrLinea
Dim concnro As String
Dim Conccod As String
Dim concabr As String
Dim concorden As String
Dim tconnro As String
Dim concext As String
Dim concvalid As String
Dim concdesde As String
Dim conchasta As String
Dim concrepet As String
Dim concretro As String
Dim concniv As String
Dim fornro As String
Dim concimp As String
Dim codseguridad As String
Dim concusado As String
Dim concpuente As String
Dim Empnro As String
Dim Conccantdec As String
Dim Conctexto As String
Dim concautor As String
Dim concfecmodi As String
Dim Concajuste As String
Dim concapertura As String

On Error GoTo E_procConceptos
    
    'Formato tipo, concnro, conccod, concabr, concorden, tconnro, concext, concvalid, concdesde, conchasta, concrepet, concretro, concniv, fornro, concimp,"
    'codseguridad , concusado, concpuente, Empnro, Conccantdec, Conctexto, concautor, concfecmodi, Concajuste, concapertura"
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 24 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    concnro = arrLinea(1)
    Conccod = arrLinea(2)
    concabr = arrLinea(3)
    concorden = arrLinea(4)
    tconnro = arrLinea(5)
    concext = arrLinea(6)
    concvalid = arrLinea(7)
    concdesde = arrLinea(8)
    conchasta = arrLinea(9)
    concrepet = arrLinea(10)
    concretro = arrLinea(11)
    concniv = arrLinea(12)
    fornro = arrLinea(13)
    concimp = arrLinea(14)
    codseguridad = arrLinea(15)
    concusado = arrLinea(16)
    concpuente = arrLinea(17)
    Empnro = arrLinea(18)
    Conccantdec = arrLinea(19)
    Conctexto = arrLinea(20)
    concautor = arrLinea(21)
    concfecmodi = arrLinea(22)
    Concajuste = arrLinea(23)
    concapertura = arrLinea(24)
    
    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT concepto ON"
    StrSql = StrSql & " INSERT INTO concepto("
    StrSql = StrSql & " concnro, conccod, concabr, concorden, tconnro, concext, concvalid, concdesde, conchasta, concrepet, concretro, concniv, fornro, concimp,"
    StrSql = StrSql & " codseguridad , concusado, concpuente, Empnro, Conccantdec, Conctexto, concautor, concfecmodi, Concajuste, concapertura)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & concnro
    StrSql = StrSql & " ," & CtrlNuloTXT(Conccod)
    StrSql = StrSql & " ," & CtrlNuloTXT(concabr)
    StrSql = StrSql & " ," & concorden
    StrSql = StrSql & " ," & tconnro
    StrSql = StrSql & " ," & CtrlNuloTXT(concext)
    StrSql = StrSql & " ," & concvalid
    StrSql = StrSql & " ," & cambiaFecha(concdesde)
    StrSql = StrSql & " ," & cambiaFecha(conchasta)
    StrSql = StrSql & " ," & concrepet
    StrSql = StrSql & " ," & concretro
    StrSql = StrSql & " ," & concniv
    StrSql = StrSql & " ," & fornro
    StrSql = StrSql & " ," & concimp
    StrSql = StrSql & " ," & codseguridad
    StrSql = StrSql & " ," & concusado
    StrSql = StrSql & " ," & concpuente
    StrSql = StrSql & " ," & Empnro
    StrSql = StrSql & " ," & Conccantdec
    StrSql = StrSql & " ," & CtrlNuloTXT(Conctexto)
    StrSql = StrSql & " ," & CtrlNuloTXT(concautor)
    StrSql = StrSql & " ," & cambiaFecha(concfecmodi)
    StrSql = StrSql & " ," & Concajuste
    StrSql = StrSql & " ," & concapertura & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT concepto OFF"
    objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

E_procConceptos:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procConceptos"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procTipoAcu(ByVal Linea As String)

Dim arrLinea
Dim tacunro As String
Dim tacudesc As String
Dim sistema  As String
Dim tacudepu As String
    
On Error GoTo E_procTipoAcu

    'Formato tipo, tacunro, tacudesc, sistema, tacudepu
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 4 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    tacunro = arrLinea(1)
    tacudesc = arrLinea(2)
    sistema = arrLinea(3)
    tacudepu = arrLinea(4)
    
    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT tipacum ON"
    StrSql = StrSql & " INSERT INTO tipacum("
    StrSql = StrSql & " tacunro, tacudesc, sistema, tacudepu)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & tacunro
    StrSql = StrSql & " ," & CtrlNuloTXT(tacudesc)
    StrSql = StrSql & " ," & sistema
    StrSql = StrSql & " ," & tacudepu & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT tipacum OFF"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procTipoAcu:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procTipoAcu"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procAcumuladores(ByVal Linea As String)

Dim arrLinea
Dim acuNro As String
Dim acudesabr As String
Dim acusist As String
Dim acudesext As String
Dim acumes As String
Dim acutopea As String
Dim acudesborde As String
Dim acurecalculo As String
Dim acuimponible As String
Dim acuimpcont As String
Dim acusel1 As String
Dim acusel2 As String
Dim acusel3 As String
Dim acuppag As String
Dim acudepu As String
Dim acuhist As String
Dim acumanual As String
Dim acuimpri As String
Dim tacunro As String
Dim Empnro As String
Dim acuretro As String
Dim acuorden As String
Dim acunoneg As String
Dim acuapertura As String
    
On Error GoTo E_procAcumuladores

    'Formato tipo ,acunro, acudesabr, acusist, acudesext, acumes, acutopea, acudesborde, acurecalculo, acuimponible, acuimpcont, acusel1, acusel2,
    'acusel3, acuppag, acudepu , acuhist, acumanual, acuimpri, tacunro, Empnro, acuretro, acuorden, acunoneg, acuapertura
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 24 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    acuNro = arrLinea(1)
    acudesabr = arrLinea(2)
    acusist = arrLinea(3)
    acudesext = arrLinea(4)
    acumes = arrLinea(5)
    acutopea = arrLinea(6)
    acudesborde = arrLinea(7)
    acurecalculo = arrLinea(8)
    acuimponible = arrLinea(9)
    acuimpcont = arrLinea(10)
    acusel1 = arrLinea(11)
    acusel2 = arrLinea(12)
    acusel3 = arrLinea(13)
    acuppag = arrLinea(14)
    acudepu = arrLinea(15)
    acuhist = arrLinea(16)
    acumanual = arrLinea(17)
    acuimpri = arrLinea(18)
    tacunro = arrLinea(19)
    Empnro = arrLinea(20)
    acuretro = arrLinea(21)
    acuorden = arrLinea(22)
    acunoneg = arrLinea(23)
    acuapertura = arrLinea(24)

    
    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT acumulador ON"
    StrSql = StrSql & " INSERT INTO acumulador("
    StrSql = StrSql & " acunro, acudesabr, acusist, acudesext, acumes, acutopea, acudesborde, acurecalculo, acuimponible, acuimpcont, acusel1, acusel2,"
    StrSql = StrSql & " acusel3, acuppag, acudepu , acuhist, acumanual, acuimpri, tacunro, Empnro, acuretro, acuorden, acunoneg, acuapertura)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & acuNro
    StrSql = StrSql & " ," & CtrlNuloTXT(acudesabr)
    StrSql = StrSql & " ," & acusist
    StrSql = StrSql & " ," & CtrlNuloTXT(acudesext)
    StrSql = StrSql & " ," & acumes
    StrSql = StrSql & " ," & acutopea
    StrSql = StrSql & " ," & acudesborde
    StrSql = StrSql & " ," & acurecalculo
    StrSql = StrSql & " ," & acuimponible
    StrSql = StrSql & " ," & acuimpcont
    StrSql = StrSql & " ," & acusel1
    StrSql = StrSql & " ," & acusel2
    StrSql = StrSql & " ," & acusel3
    StrSql = StrSql & " ," & acuppag
    StrSql = StrSql & " ," & acudepu
    StrSql = StrSql & " ," & acuhist
    StrSql = StrSql & " ," & acumanual
    StrSql = StrSql & " ," & acuimpri
    StrSql = StrSql & " ," & tacunro
    StrSql = StrSql & " ," & Empnro
    StrSql = StrSql & " ," & acuretro
    StrSql = StrSql & " ," & acuorden
    StrSql = StrSql & " ," & acunoneg
    StrSql = StrSql & " ," & acuapertura & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT acumulador OFF"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procAcumuladores:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procAcumuladores"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procItems(ByVal Linea As String)

Dim arrLinea
Dim itenro As String
Dim itenom As String
Dim itesigno As String
Dim iterenglon As String
Dim itetipotope As String
Dim iteporctope As String
Dim iteitemstope As String
Dim iteprorr As String
    
On Error GoTo E_procItems
    
    'Formato tipo, itenro, itenom, itesigno, iterenglon, itetipotope, iteporctope, iteitemstope, iteprorr
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 8 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    itenro = arrLinea(1)
    itenom = arrLinea(2)
    itesigno = arrLinea(3)
    iterenglon = arrLinea(4)
    itetipotope = arrLinea(5)
    iteporctope = arrLinea(6)
    iteitemstope = arrLinea(7)
    iteprorr = arrLinea(8)

    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT item ON"
    StrSql = StrSql & " INSERT INTO item("
    StrSql = StrSql & " itenro, itenom, itesigno, iterenglon, itetipotope, iteporctope, iteitemstope, iteprorr)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & itenro
    StrSql = StrSql & " ," & CtrlNuloTXT(itenom)
    StrSql = StrSql & " ," & itesigno
    StrSql = StrSql & " ," & CtrlNuloTXT(iterenglon)
    StrSql = StrSql & " ," & itetipotope
    StrSql = StrSql & " ," & iteporctope
    StrSql = StrSql & " ," & CtrlNuloTXT(iteitemstope)
    StrSql = StrSql & " ," & iteprorr & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT item OFF"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procItems:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procItems"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procTipoProc(ByVal Linea As String)

Dim arrLinea
Dim tprocnro As String
Dim tprocdesc As String
Dim Empnro As String
Dim tliqnro As String
Dim final As String
Dim ajugcias As String
Dim tprocrecalculo As String
    
On Error GoTo E_procTipoProc
    
    'Formato tipo, tprocnro, tprocdesc, empnro, tliqnro, final, ajugcias, tprocrecalculo
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 7 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    tprocnro = arrLinea(1)
    tprocdesc = arrLinea(2)
    Empnro = arrLinea(3)
    tliqnro = arrLinea(4)
    final = arrLinea(5)
    ajugcias = arrLinea(6)
    tprocrecalculo = arrLinea(7)
    
    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT tipoproc ON"
    StrSql = StrSql & " INSERT INTO tipoproc("
    StrSql = StrSql & " tprocnro, tprocdesc, empnro, tliqnro, final, ajugcias, tprocrecalculo)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & tprocnro
    StrSql = StrSql & " ," & CtrlNuloTXT(tprocdesc)
    StrSql = StrSql & " ," & Empnro
    StrSql = StrSql & " ," & tliqnro
    StrSql = StrSql & " ," & final
    StrSql = StrSql & " ," & ajugcias
    StrSql = StrSql & " ," & tprocrecalculo & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT tipoproc OFF"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procTipoProc:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procTipoProc"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procProcesos(ByVal Linea As String)

Dim arrLinea
Dim pronro As String
Dim prodesc As String
Dim propend As String
Dim profeccorr As String
Dim profecplan As String
Dim pliqnro As String
Dim tprocnro As String
Dim prosist As String
Dim profecpago As String
Dim profecini As String
Dim profecfin As String
Dim Empnro As String
Dim proaprob As String
Dim proestdesc As String
    
On Error GoTo E_procProcesos
    
    'Formato tipo, pronro, prodesc, propend, profeccorr, profecplan, pliqnro, tprocnro,
    'prosist, profecpago, profecini, profecfin, empnro, proaprob, proestdesc
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 14 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    pronro = arrLinea(1)
    prodesc = arrLinea(2)
    propend = arrLinea(3)
    profeccorr = arrLinea(4)
    profecplan = arrLinea(5)
    pliqnro = arrLinea(6)
    tprocnro = arrLinea(7)
    prosist = arrLinea(8)
    profecpago = arrLinea(9)
    profecini = arrLinea(10)
    profecfin = arrLinea(11)
    Empnro = arrLinea(12)
    proaprob = arrLinea(13)
    proestdesc = arrLinea(14)
    
    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT proceso ON"
    StrSql = StrSql & " INSERT INTO proceso("
    StrSql = StrSql & " pronro, prodesc, propend, profeccorr, profecplan, pliqnro, tprocnro,"
    StrSql = StrSql & " prosist, profecpago, profecini, profecfin, empnro, proaprob, proestdesc)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & pronro
    StrSql = StrSql & " ," & CtrlNuloTXT(prodesc)
    StrSql = StrSql & " ," & propend
    StrSql = StrSql & " ," & cambiaFecha(profeccorr)
    StrSql = StrSql & " ," & cambiaFecha(profecplan)
    StrSql = StrSql & " ," & pliqnro
    StrSql = StrSql & " ," & tprocnro
    StrSql = StrSql & " ," & prosist
    StrSql = StrSql & " ," & cambiaFecha(profecpago)
    StrSql = StrSql & " ," & cambiaFecha(profecini)
    StrSql = StrSql & " ," & cambiaFecha(profecfin)
    StrSql = StrSql & " ," & Empnro
    StrSql = StrSql & " ," & proaprob
    StrSql = StrSql & " ," & CtrlNuloTXT(proestdesc) & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT proceso OFF"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procProcesos:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procProcesos"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procPeriodos(ByVal Linea As String)

Dim arrLinea
Dim pliqnro As String
Dim pliqdesc As String
Dim pliqdesde As String
Dim pliqhasta As String
Dim pliqmes As String
Dim pliqbackup As String
Dim pliqdepurado As String
Dim pliqbco As String
Dim pliqsuc As String
Dim pliqabierto As String
Dim pliqultimo As String
Dim pliqfecdep As String
Dim pliqanio As String
Dim pliqsist As String
Dim pliqdepant As String
Dim pliqtexto As String
Dim Empnro As String

On Error GoTo E_procPeriodos
    
    'Formato tipo, pliqnro, pliqdesc, pliqdesde, pliqhasta, pliqmes, pliqbackup, pliqdepurado, pliqbco, pliqsuc,
    'pliqabierto, pliqultimo, pliqfecdep, pliqanio, pliqsist,pliqdepant, pliqtexto, empnro

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 17 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    pliqnro = arrLinea(1)
    pliqdesc = arrLinea(2)
    pliqdesde = arrLinea(3)
    pliqhasta = arrLinea(4)
    pliqmes = arrLinea(5)
    pliqbackup = arrLinea(6)
    pliqdepurado = arrLinea(7)
    pliqbco = arrLinea(8)
    pliqsuc = arrLinea(9)
    pliqabierto = arrLinea(10)
    pliqultimo = arrLinea(11)
    pliqfecdep = arrLinea(12)
    pliqanio = arrLinea(13)
    pliqsist = arrLinea(14)
    pliqdepant = arrLinea(15)
    pliqtexto = arrLinea(16)
    Empnro = arrLinea(17)
    
    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT periodo ON"
    StrSql = StrSql & " INSERT INTO periodo("
    StrSql = StrSql & " pliqnro, pliqdesc, pliqdesde, pliqhasta, pliqmes, pliqbackup, pliqdepurado, pliqbco, pliqsuc,"
    StrSql = StrSql & " pliqabierto, pliqultimo, pliqfecdep, pliqanio, pliqsist,pliqdepant, pliqtexto, empnro)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & pliqnro
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqdesc)
    StrSql = StrSql & " ," & cambiaFecha(pliqdesde)
    StrSql = StrSql & " ," & cambiaFecha(pliqhasta)
    StrSql = StrSql & " ," & pliqmes
    StrSql = StrSql & " ," & pliqbackup
    StrSql = StrSql & " ," & pliqdepurado
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqbco)
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqsuc)
    StrSql = StrSql & " ," & pliqabierto
    StrSql = StrSql & " ," & pliqultimo
    StrSql = StrSql & " ," & cambiaFecha(pliqfecdep)
    StrSql = StrSql & " ," & pliqanio
    StrSql = StrSql & " ," & pliqsist
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqdepant)
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqtexto)
    StrSql = StrSql & " ," & Empnro & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT periodo OFF"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procPeriodos:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procPeriodos"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub BorrarTabla(ByVal tipo As Long)

On Error GoTo E_BorrarTabla

    Select Case tipo
        Case 1:
            StrSql = "truncate table tmp_empleado"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla tmp_empleado"
        Case 2:
            StrSql = "truncate table tipconcep"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla tipconcep"
        Case 3:
            StrSql = "truncate table concepto"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla concepto"
        Case 4:
            StrSql = "truncate table tipacum"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla tipacum"
        Case 5:
            StrSql = "truncate table acumulador"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla acumulador"
        Case 6:
            StrSql = "truncate table item"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla item"
        Case 7:
            StrSql = "truncate table tipoproc"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla tipoproc"
        Case 8:
            StrSql = "truncate table proceso"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla proceso"
        Case 9:
            StrSql = "truncate table periodo"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla periodo"
        Case 28:
            StrSql = "truncate table escala_ded"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla escala_ded"
        Case 29:
            StrSql = "DELETE FROM confrep WHERE repnro = 114"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra confrep 114"
        Case 30:
            StrSql = "truncate table escala"
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla periodo"
    End Select
    
Exit Sub

E_BorrarTabla:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BorrarTabla"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub BorrarBatch(ByVal Linea As String, ByRef listaBpronro As String)
'Borra la tabla batch proceso para los tipos recibo, F649 y libro ley
'en todo el mes de la fecha que viene en la linea
Dim arrLinea
Dim Fecha As String
Dim FecDesde As Date
Dim FecHasta As Date

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_BorrarBatch

    arrLinea = Split(Linea, Sep)
    
    If UBound(arrLinea) <> 24 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        HuboError = True
        Exit Sub
    End If
    
    Fecha = arrLinea(3)
    FecDesde = PrimerDiaMes(Month(Fecha), Year(Fecha))
    FecHasta = UltimoDiaMes(Month(Fecha), Year(Fecha))
    
    'Armo la lista de los procesos a borrar
    listaBpronro = "0"
    StrSql = "SELECT bpronro FROM batch_proceso"
    StrSql = StrSql & " WHERE btprcnro IN (26, 45, 50)"
    StrSql = StrSql & " AND bprcfecha >= " & ConvFecha(FecDesde)
    StrSql = StrSql & " AND bprcfecha <= " & ConvFecha(FecHasta)
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        listaBpronro = rs_Datos!bpronro
        rs_Datos.MoveNext
    Loop
    rs_Datos.Close
    
    'borro los procesos seleccionados
    StrSql = "DELETE batch_proceso"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBpronro & " )"
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    StrSql = "truncate table tmp_batch_proc"
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla batch_proceso y temporal"
    
Set rs_Datos = Nothing

Exit Sub

E_BorrarBatch:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BorrarBatch"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub BorrarRecibo(ByVal listaBpronro As String)
'Borra la tabla de rep_recibo y rep_recibo_det

On Error GoTo E_BorrarRecibo

    'borro los Recibos
    StrSql = "DELETE rep_recibo"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBpronro & " )"
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'borro los detalles de Recibos
    StrSql = "DELETE rep_recibo_det"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBpronro & " )"
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla rep_recibo y rep_recibo_det"
    
Exit Sub

E_BorrarRecibo:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BorrarRecibo"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub BorrarLibro(ByVal listaBpronro As String)
'Borra la tabla de rep_recibo y rep_reprecibodet

    
    'borro los rep_libroley
    StrSql = "DELETE rep_libroley"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBpronro & " )"
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'borro los detalles de rep_libroley_det
    StrSql = "DELETE rep_libroley_det"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBpronro & " )"
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'borro los detalles de rep_libroley_fam
    StrSql = "DELETE rep_libroley_fam"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBpronro & " )"
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla rep_libroley, rep_libroley_det y rep_libroley_fam"
    
Exit Sub

E_BorrarRecibo:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BorrarRecibo"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub BorrarF649(ByVal listaBpronro As String)
'Borra la tabla de rep_libroley y rep_libroley_det

On Error GoTo E_BorrarF649
    
    'borro los Recibos
    StrSql = "DELETE rep19"
    StrSql = StrSql & " WHERE bpronro IN ( " & listaBpronro & " )"
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 1) & "Borra Tabla rep19 (F649)"
    
Exit Sub

E_BorrarF649:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BorrarF649"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Sub

Public Function cambiaFecha(ByVal Fecha As String) As String

    If EsNulo(Fecha) Then
        cambiaFecha = "NULL"
    Else
        cambiaFecha = ConvFecha(Fecha)
    End If

End Function


Public Function mapearTenro(ByVal ternro As Long) As Long
Dim rs_ternro As New ADODB.Recordset
Dim aux As Long

On Error GoTo E_mapearTenro

    aux = 0
    'Con el ternro busco que legajo tiene el empleado en la base origen
    StrSql = "SELECT empleg"
    StrSql = StrSql & " FROM tmp_empleado"
    StrSql = StrSql & " WHERE ternro = " & ternro
    OpenRecordset StrSql, rs_ternro
    If Not rs_ternro.EOF Then
        aux = rs_ternro!empleg
    End If
    
    'Con el legajo origen busco el nuevo ternro en la base destino
    If aux <> 0 Then
        StrSql = "SELECT ternro"
        StrSql = StrSql & " FROM empleado"
        StrSql = StrSql & " WHERE empleg = " & aux
        OpenRecordset StrSql, rs_ternro
        If Not rs_ternro.EOF Then
            aux = rs_ternro!ternro
        End If
    End If
    
    mapearTenro = aux

If rs_ternro.State = adStateOpen Then rs_ternro.Close
Set rs_ternro = Nothing

Exit Function

E_mapearTenro:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: mapearTenro"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Function


Public Function mapearBpronro(ByVal bpronro As Long) As Long
Dim rs_Bpronro As New ADODB.Recordset
Dim aux As Long

On Error GoTo E_mapearBpronro

    aux = 0
    'busco el nuevo valor del bpronro en la base
    StrSql = "SELECT bpronro_dest"
    StrSql = StrSql & " FROM tmp_batch_proc"
    StrSql = StrSql & " WHERE bpronro_ori = " & bpronro
    OpenRecordset StrSql, rs_Bpronro
    If Not rs_Bpronro.EOF Then
        aux = rs_Bpronro!bpronro_dest
    End If
    
    mapearBpronro = aux

If rs_Bpronro.State = adStateOpen Then rs_Bpronro.Close
Set rs_Bpronro = Nothing

Exit Function

E_mapearBpronro:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: mapearBpronro"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
End Function


Public Sub procCabLiq(ByVal Linea As String)

Dim arrLinea
Dim cliqnro As String
Dim pronro As String
Dim Empleado As String
Dim ppagnro As String
Dim cliqtexto As String
Dim cliqdesde As String
Dim cliqhasta As String
Dim nrorecibo As String
Dim cliqnrocorr As String
Dim nroimp As String
Dim fechaimp As String
Dim entregado As String
Dim fechaentrega As String
Dim ternroAux As Long
Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procCabLiq

    'Formato tipo, cliqnro, pronro, Empleado, ppagnro, cliqtexto,
    'cliqdesde, cliqhasta, nrorecibo, cliqnrocorr, nroimp,
    'fechaimp , entregado, fechaentrega

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 13 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    cliqnro = arrLinea(1)
    pronro = arrLinea(2)
    Empleado = arrLinea(3)
    ppagnro = arrLinea(4)
    cliqtexto = arrLinea(5)
    cliqdesde = arrLinea(6)
    cliqhasta = arrLinea(7)
    nrorecibo = arrLinea(8)
    cliqnrocorr = arrLinea(9)
    nroimp = arrLinea(10)
    fechaimp = arrLinea(11)
    entregado = arrLinea(12)
    fechaentrega = arrLinea(13)
    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(Empleado)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro el ternro " & Empleado
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    
    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT cliqnro FROM cabliq WHERE cliqnro = " & cliqnro
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE cabliq WHERE cliqnro = " & cliqnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT cabliq ON"
    StrSql = StrSql & " INSERT INTO cabliq("
    StrSql = StrSql & " cliqnro, pronro, Empleado, ppagnro, cliqtexto,"
    StrSql = StrSql & " cliqdesde, cliqhasta, nrorecibo, cliqnrocorr, nroimp,"
    StrSql = StrSql & " fechaimp , entregado, fechaentrega)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & cliqnro
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & ppagnro
    StrSql = StrSql & " ," & CtrlNuloTXT(cliqtexto)
    StrSql = StrSql & " ," & cambiaFecha(cliqdesde)
    StrSql = StrSql & " ," & cambiaFecha(cliqhasta)
    StrSql = StrSql & " ," & nrorecibo
    StrSql = StrSql & " ," & cliqnrocorr
    StrSql = StrSql & " ," & nroimp
    StrSql = StrSql & " ," & cambiaFecha(fechaimp)
    StrSql = StrSql & " ," & entregado
    StrSql = StrSql & " ," & cambiaFecha(fechaentrega) & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT cabliq OFF"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procCabLiq:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procCabLiq"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procDetliq(ByVal Linea As String)

Dim arrLinea
Dim concnro As String
Dim dlimonto As String
Dim dlifec As String
Dim cliqnro As String
Dim dlicant As String
Dim dlimonto_base As String
Dim dliporcent As String
Dim dlitexto As String
Dim fornro As String
Dim tconnro As String
Dim dliretro As String
Dim ajustado As String
Dim dliqdesde As String
Dim dliqhasta As String

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procDetliq


    'Formato tipo, concnro, dlimonto, dlifec, cliqnro, dlicant, dlimonto_base, dliporcent, dlitexto, fornro,
    'tconnro, dliretro, ajustado, dliqdesde, dliqhasta

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 14 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    concnro = arrLinea(1)
    dlimonto = arrLinea(2)
    dlifec = arrLinea(3)
    cliqnro = arrLinea(4)
    dlicant = arrLinea(5)
    dlimonto_base = arrLinea(6)
    dliporcent = arrLinea(7)
    dlitexto = arrLinea(8)
    fornro = arrLinea(9)
    tconnro = arrLinea(10)
    dliretro = arrLinea(11)
    ajustado = arrLinea(12)
    dliqdesde = arrLinea(13)
    dliqhasta = arrLinea(14)
    
    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT cliqnro FROM detliq WHERE cliqnro = " & cliqnro
    StrSql = StrSql & " AND concnro = " & concnro
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE detliq WHERE cliqnro = " & cliqnro
        StrSql = StrSql & " AND concnro = " & concnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO detliq("
    StrSql = StrSql & " concnro, dlimonto, dlifec, cliqnro, dlicant, dlimonto_base, dliporcent, dlitexto, fornro,"
    StrSql = StrSql & " tconnro, dliretro, ajustado, dliqdesde, dliqhasta)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & concnro
    StrSql = StrSql & " ," & dlimonto
    StrSql = StrSql & " ," & cambiaFecha(dlifec)
    StrSql = StrSql & " ," & cliqnro
    StrSql = StrSql & " ," & dlicant
    StrSql = StrSql & " ," & dlimonto_base
    StrSql = StrSql & " ," & dliporcent
    StrSql = StrSql & " ," & CtrlNuloTXT(dlitexto)
    StrSql = StrSql & " ," & fornro
    StrSql = StrSql & " ," & tconnro
    StrSql = StrSql & " ," & dliretro
    StrSql = StrSql & " ," & ajustado
    StrSql = StrSql & " ," & dliqdesde
    StrSql = StrSql & " ," & dliqhasta & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procDetliq:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procDetliq"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procAcu_liq(ByVal Linea As String)

Dim arrLinea
Dim acuNro As String
Dim cliqnro As String
Dim almonto As String
Dim alcant As String
Dim alfecret As String
Dim almontoreal As String

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procAcu_liq


    'Formato tipo, acunro, cliqnro, almonto, alcant, alfecret, almontoreal

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 6 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    acuNro = arrLinea(1)
    cliqnro = arrLinea(2)
    almonto = arrLinea(3)
    alcant = arrLinea(4)
    alfecret = arrLinea(5)
    almontoreal = arrLinea(6)

    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT cliqnro FROM acu_liq WHERE cliqnro = " & cliqnro
    StrSql = StrSql & " AND acunro = " & acuNro
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE acu_liq WHERE cliqnro = " & cliqnro
        StrSql = StrSql & " AND acunro = " & acuNro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO acu_liq("
    StrSql = StrSql & " acunro, cliqnro, almonto, alcant, alfecret, almontoreal)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & acuNro
    StrSql = StrSql & " ," & cliqnro
    StrSql = StrSql & " ," & almonto
    StrSql = StrSql & " ," & alcant
    StrSql = StrSql & " ," & cambiaFecha(alfecret)
    StrSql = StrSql & " ," & almontoreal & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procAcu_liq:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procAcu_liq"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procAcu_mes(ByVal Linea As String)

Dim arrLinea
Dim ternro As String
Dim ternroAux As Long
Dim acuNro As String
Dim amanio As String
Dim ammonto As String
Dim amcant As String
Dim ammes As String
Dim ammontoreal As String

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procAcu_mes


    'Formato tipo, ternro, acunro, amanio, ammonto, amcant, ammes, ammontoreal

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 7 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    ternro = arrLinea(1)
    acuNro = arrLinea(2)
    amanio = arrLinea(3)
    ammonto = arrLinea(4)
    amcant = arrLinea(5)
    ammes = arrLinea(6)
    ammontoreal = arrLinea(7)
    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)
        
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If

    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT ternro FROM Acu_mes WHERE ternro = " & ternroAux
    StrSql = StrSql & " AND acunro = " & acuNro
    StrSql = StrSql & " AND amanio = " & amanio
    StrSql = StrSql & " AND ammes = " & ammes
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE Acu_mes WHERE ternro = " & ternroAux
        StrSql = StrSql & " AND acunro = " & acuNro
        StrSql = StrSql & " AND amanio = " & amanio
        StrSql = StrSql & " AND ammes = " & ammes
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO Acu_mes("
    StrSql = StrSql & " ternro, acunro, amanio, ammonto, amcant, ammes, ammontoreal)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & ternroAux
    StrSql = StrSql & " ," & acuNro
    StrSql = StrSql & " ," & amanio
    StrSql = StrSql & " ," & ammonto
    StrSql = StrSql & " ," & amcant
    StrSql = StrSql & " ," & ammes
    StrSql = StrSql & " ," & ammontoreal & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procAcu_mes:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procAcu_mes"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procDesliq(ByVal Linea As String)

Dim arrLinea
Dim itenro As String
Dim Empleado As String
Dim dlfecha As String
Dim pronro As String
Dim dlmonto As String
Dim dlprorratea As String
Dim ternroAux As Long

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procDesliq

    'Formato tipo, itenro, empleado, dlfecha, pronro, dlmonto, dlprorratea

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 6 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    itenro = arrLinea(1)
    Empleado = arrLinea(2)
    dlfecha = arrLinea(3)
    pronro = arrLinea(4)
    dlmonto = arrLinea(5)
    dlprorratea = arrLinea(6)
    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(Empleado)

    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro el ternro " & Empleado
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If


    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT empleado FROM desliq WHERE empleado = " & ternroAux
    StrSql = StrSql & " AND pronro = " & pronro
    StrSql = StrSql & " AND itenro = " & itenro
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE desliq WHERE empleado = " & ternroAux
        StrSql = StrSql & " AND pronro = " & pronro
        StrSql = StrSql & " AND itenro = " & itenro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO desliq("
    StrSql = StrSql & " itenro, empleado, dlfecha, pronro, dlmonto, dlprorratea)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & itenro
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & cambiaFecha(dlfecha)
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & dlmonto
    StrSql = StrSql & " ," & dlprorratea & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procDesliq:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procDesliq"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procFicharet(ByVal Linea As String)

Dim arrLinea
Dim Fecha As String
Dim importe As String
Dim pronro As String
Dim liqsistema As String
Dim Empleado As String
Dim ternroAux As Long

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procFicharet


    'Formato tipo, fecha, importe, pronro, liqsistema, empleado

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 5 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    Fecha = arrLinea(1)
    importe = arrLinea(2)
    pronro = arrLinea(3)
    liqsistema = arrLinea(4)
    Empleado = arrLinea(5)

    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(Empleado)

    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro el ternro " & Empleado
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If

    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT empleado FROM ficharet WHERE empleado = " & ternroAux
    StrSql = StrSql & " AND pronro = " & pronro
    StrSql = StrSql & " AND Fecha = " & cambiaFecha(Fecha)
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE ficharet WHERE empleado = " & ternroAux
        StrSql = StrSql & " AND pronro = " & pronro
        StrSql = StrSql & " AND Fecha = " & cambiaFecha(Fecha)
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO ficharet("
    StrSql = StrSql & " fecha, importe, pronro, liqsistema, empleado)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & cambiaFecha(Fecha)
    StrSql = StrSql & " ," & importe
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & liqsistema
    StrSql = StrSql & " ," & ternroAux & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procFicharet:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procFicharet"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procImpproarg(ByVal Linea As String)

Dim arrLinea
Dim acuNro As String
Dim cliqnro As String
Dim tconnro As String
Dim ipacant As String
Dim ipamonto As String

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procImpproarg


    'Formato tipo, acunro, cliqnro, tconnro, ipacant, ipamonto

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 5 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    acuNro = arrLinea(1)
    cliqnro = arrLinea(2)
    tconnro = arrLinea(3)
    ipacant = arrLinea(4)
    ipamonto = arrLinea(5)


    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT acunro FROM impproarg WHERE acunro = " & acuNro
    StrSql = StrSql & " AND cliqnro = " & cliqnro
    StrSql = StrSql & " AND tconnro = " & tconnro
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE impproarg WHERE acunro = " & acuNro
        StrSql = StrSql & " AND cliqnro = " & cliqnro
        StrSql = StrSql & " AND tconnro = " & tconnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO impproarg("
    StrSql = StrSql & " acunro, cliqnro, tconnro, ipacant, ipamonto)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & acuNro
    StrSql = StrSql & " ," & cliqnro
    StrSql = StrSql & " ," & tconnro
    StrSql = StrSql & " ," & ipacant
    StrSql = StrSql & " ," & ipamonto & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procImpproarg:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procImpproarg"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procImpmesarg(ByVal Linea As String)

Dim arrLinea
Dim ternro As String
Dim acuNro As String
Dim tconnro As String
Dim imaanio As String
Dim imames As String
Dim imacant As String
Dim imamonto As String
Dim ternroAux As Long

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procImpmesarg

    'Formato tipo, ternro, acunro, tconnro, imaanio, imames, imacant, imamonto

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 7 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    ternro = arrLinea(1)
    acuNro = arrLinea(2)
    tconnro = arrLinea(3)
    imaanio = arrLinea(4)
    imames = arrLinea(5)
    imacant = arrLinea(6)
    imamonto = arrLinea(7)

    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)

    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If

    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT acunro FROM impmesarg WHERE acunro = " & acuNro
    StrSql = StrSql & " AND tconnro = " & tconnro
    StrSql = StrSql & " AND imaanio = " & imaanio
    StrSql = StrSql & " AND imames = " & imames
    StrSql = StrSql & " AND ternro = " & ternroAux
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE impmesarg WHERE acunro = " & acuNro
        StrSql = StrSql & " AND tconnro = " & tconnro
        StrSql = StrSql & " AND imaanio = " & imaanio
        StrSql = StrSql & " AND imames = " & imames
        StrSql = StrSql & " AND ternro = " & ternroAux
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO impmesarg("
    StrSql = StrSql & " ternro, acunro, tconnro, imaanio, imames, imacant, imamonto)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & ternroAux
    StrSql = StrSql & " ," & acuNro
    StrSql = StrSql & " ," & tconnro
    StrSql = StrSql & " ," & imaanio
    StrSql = StrSql & " ," & imames
    StrSql = StrSql & " ," & imacant
    StrSql = StrSql & " ," & imamonto & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procImpmesarg:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procImpmesarg"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procTraza_gan(ByVal Linea As String)

Dim arrLinea
Dim pliqnro As String
Dim concnro As String
Dim empresa As String
Dim fecha_pago As String
Dim ternro As String
Dim msr As String
Dim nomsr As String
Dim nogan As String
Dim jubilacion As String
Dim osocial As String
Dim cuota_medico As String
Dim prima_seguro As String
Dim sepelio As String
Dim estimados As String
Dim otras As String
Dim donacion As String
Dim dedesp As String
Dim noimpo As String
Dim car_flia As String
Dim conyuge As String
Dim hijo As String
Dim otras_cargas As String
Dim retenciones As String
Dim promo As String
Dim saldo As String
Dim sindicato As String
Dim ret_mes As String
Dim mon_conyuge As String
Dim mon_hijo As String
Dim mon_otras As String
Dim viaticos As String
Dim amortizacion As String
Dim entidad1 As String
Dim entidad2 As String
Dim entidad3 As String
Dim entidad4 As String
Dim entidad5 As String
Dim entidad6 As String
Dim entidad7 As String
Dim entidad8 As String
Dim entidad9 As String
Dim entidad10 As String
Dim entidad11 As String
Dim entidad12 As String
Dim entidad13 As String
Dim entidad14 As String
Dim cuit_entidad1 As String
Dim cuit_entidad2 As String
Dim cuit_entidad3 As String
Dim cuit_entidad4 As String
Dim cuit_entidad5 As String
Dim cuit_entidad6 As String
Dim cuit_entidad7 As String
Dim cuit_entidad8 As String
Dim cuit_entidad9 As String
Dim cuit_entidad10 As String
Dim cuit_entidad11 As String
Dim cuit_entidad12 As String
Dim cuit_entidad13 As String
Dim cuit_entidad14 As String
Dim monto_entidad1 As String
Dim monto_entidad2 As String
Dim monto_entidad3 As String
Dim monto_entidad4 As String
Dim monto_entidad5 As String
Dim monto_entidad6 As String
Dim monto_entidad7 As String
Dim monto_entidad8 As String
Dim monto_entidad9 As String
Dim monto_entidad10 As String
Dim monto_entidad11 As String
Dim monto_entidad12 As String
Dim monto_entidad13 As String
Dim monto_entidad14 As String
Dim ganimpo As String
Dim ganneta As String
Dim total_entidad1 As String
Dim total_entidad2 As String
Dim total_entidad3 As String
Dim total_entidad4 As String
Dim total_entidad5 As String
Dim total_entidad6 As String
Dim total_entidad7 As String
Dim total_entidad8 As String
Dim total_entidad9 As String
Dim total_entidad10 As String
Dim total_entidad11 As String
Dim total_entidad12 As String
Dim total_entidad13 As String
Dim total_entidad14 As String
Dim pronro As String
Dim imp_deter As String
Dim eme_medicas As String
Dim seguro_optativo As String
Dim seguro_retiro As String
Dim tope_os_priv As String
Dim empleg As String
Dim deducciones As String
Dim art23 As String
Dim porcdeduc As String
Dim ternroAux As Long

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procTraza_gan

    'Formato tipo, pliqnro, concnro, empresa, fecha_pago, ternro, msr, nomsr, nogan, jubilacion,
    'osocial, cuota_medico, prima_seguro, sepelio, estimados, otras,donacion, dedesp,
    'noimpo, car_flia, conyuge, hijo, otras_cargas, retenciones, promo, saldo, sindicato,
    'ret_mes, mon_conyuge, mon_hijo, mon_otras, viaticos, amortizacion, entidad1, entidad2,
    'entidad3, entidad4, entidad5, entidad6, entidad7, entidad8, entidad9, entidad10, entidad11,
    'entidad12, entidad13, entidad14, cuit_entidad1, cuit_entidad2, cuit_entidad3, cuit_entidad4,
    'cuit_entidad5, cuit_entidad6, cuit_entidad7, cuit_entidad8, cuit_entidad9, cuit_entidad10,
    'cuit_entidad11, cuit_entidad12, cuit_entidad13, cuit_entidad14, monto_entidad1, monto_entidad2,
    'monto_entidad3, monto_entidad4, monto_entidad5, monto_entidad6, monto_entidad7, monto_entidad8,
    'monto_entidad9, monto_entidad10, monto_entidad11, monto_entidad12, monto_entidad13,
    'monto_entidad14, ganimpo, ganneta, total_entidad1, total_entidad2, total_entidad3,
    'total_entidad4, total_entidad5, total_entidad6, total_entidad7, total_entidad8, total_entidad9,
    'total_entidad10, total_entidad11, total_entidad12, total_entidad13, total_entidad14, pronro,
    'imp_deter, eme_medicas, seguro_optativo, seguro_retiro, tope_os_priv, empleg, deducciones, art23,
    'porcdeduc

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 100 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    pliqnro = arrLinea(1)
    concnro = arrLinea(2)
    empresa = arrLinea(3)
    fecha_pago = arrLinea(4)
    ternro = arrLinea(5)
    msr = arrLinea(6)
    nomsr = arrLinea(7)
    nogan = arrLinea(8)
    jubilacion = arrLinea(9)
    osocial = arrLinea(10)
    cuota_medico = arrLinea(11)
    prima_seguro = arrLinea(12)
    sepelio = arrLinea(13)
    estimados = arrLinea(14)
    otras = arrLinea(15)
    donacion = arrLinea(16)
    dedesp = arrLinea(17)
    noimpo = arrLinea(18)
    car_flia = arrLinea(19)
    conyuge = arrLinea(20)
    hijo = arrLinea(21)
    otras_cargas = arrLinea(22)
    retenciones = arrLinea(23)
    promo = arrLinea(24)
    saldo = arrLinea(25)
    sindicato = arrLinea(26)
    ret_mes = arrLinea(27)
    mon_conyuge = arrLinea(28)
    mon_hijo = arrLinea(29)
    mon_otras = arrLinea(30)
    viaticos = arrLinea(31)
    amortizacion = arrLinea(32)
    entidad1 = arrLinea(33)
    entidad2 = arrLinea(34)
    entidad3 = arrLinea(35)
    entidad4 = arrLinea(36)
    entidad5 = arrLinea(37)
    entidad6 = arrLinea(38)
    entidad7 = arrLinea(39)
    entidad8 = arrLinea(40)
    entidad9 = arrLinea(41)
    entidad10 = arrLinea(42)
    entidad11 = arrLinea(43)
    entidad12 = arrLinea(44)
    entidad13 = arrLinea(45)
    entidad14 = arrLinea(46)
    cuit_entidad1 = arrLinea(47)
    cuit_entidad2 = arrLinea(48)
    cuit_entidad3 = arrLinea(49)
    cuit_entidad4 = arrLinea(50)
    cuit_entidad5 = arrLinea(51)
    cuit_entidad6 = arrLinea(52)
    cuit_entidad7 = arrLinea(53)
    cuit_entidad8 = arrLinea(54)
    cuit_entidad9 = arrLinea(55)
    cuit_entidad10 = arrLinea(56)
    cuit_entidad11 = arrLinea(57)
    cuit_entidad12 = arrLinea(58)
    cuit_entidad13 = arrLinea(59)
    cuit_entidad14 = arrLinea(60)
    monto_entidad1 = arrLinea(61)
    monto_entidad2 = arrLinea(62)
    monto_entidad3 = arrLinea(63)
    monto_entidad4 = arrLinea(64)
    monto_entidad5 = arrLinea(65)
    monto_entidad6 = arrLinea(66)
    monto_entidad7 = arrLinea(67)
    monto_entidad8 = arrLinea(68)
    monto_entidad9 = arrLinea(69)
    monto_entidad10 = arrLinea(70)
    monto_entidad11 = arrLinea(71)
    monto_entidad12 = arrLinea(72)
    monto_entidad13 = arrLinea(73)
    monto_entidad14 = arrLinea(74)
    ganimpo = arrLinea(75)
    ganneta = arrLinea(76)
    total_entidad1 = arrLinea(77)
    total_entidad2 = arrLinea(78)
    total_entidad3 = arrLinea(79)
    total_entidad4 = arrLinea(80)
    total_entidad5 = arrLinea(81)
    total_entidad6 = arrLinea(82)
    total_entidad7 = arrLinea(83)
    total_entidad8 = arrLinea(84)
    total_entidad9 = arrLinea(85)
    total_entidad10 = arrLinea(86)
    total_entidad11 = arrLinea(87)
    total_entidad12 = arrLinea(88)
    total_entidad13 = arrLinea(89)
    total_entidad14 = arrLinea(90)
    pronro = arrLinea(91)
    imp_deter = arrLinea(92)
    eme_medicas = arrLinea(93)
    seguro_optativo = arrLinea(94)
    seguro_retiro = arrLinea(95)
    tope_os_priv = arrLinea(96)
    empleg = arrLinea(97)
    deducciones = arrLinea(98)
    art23 = arrLinea(99)
    porcdeduc = arrLinea(100)
    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If

    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT ternro FROM traza_gan WHERE ternro = " & ternroAux
    StrSql = StrSql & " AND pliqnro = " & pliqnro
    StrSql = StrSql & " AND pronro = " & pronro
    StrSql = StrSql & " AND concnro = " & concnro
    StrSql = StrSql & " AND empresa = " & empresa
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE traza_gan WHERE ternro = " & ternroAux
        StrSql = StrSql & " AND pliqnro = " & pliqnro
        StrSql = StrSql & " AND pronro = " & pronro
        StrSql = StrSql & " AND concnro = " & concnro
        StrSql = StrSql & " AND empresa = " & empresa
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO traza_gan("
    StrSql = StrSql & " pliqnro, concnro, empresa, fecha_pago, ternro, msr, nomsr, nogan, jubilacion,"
    StrSql = StrSql & " osocial, cuota_medico, prima_seguro, sepelio, estimados, otras,donacion, dedesp,"
    StrSql = StrSql & " noimpo, car_flia, conyuge, hijo, otras_cargas, retenciones, promo, saldo, sindicato,"
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
    StrSql = StrSql & " porcdeduc)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & pliqnro
    StrSql = StrSql & " ," & concnro
    StrSql = StrSql & " ," & empresa
    StrSql = StrSql & " ," & cambiaFecha(fecha_pago)
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & msr
    StrSql = StrSql & " ," & nomsr
    StrSql = StrSql & " ," & nogan
    StrSql = StrSql & " ," & jubilacion
    StrSql = StrSql & " ," & osocial
    StrSql = StrSql & " ," & cuota_medico
    StrSql = StrSql & " ," & prima_seguro
    StrSql = StrSql & " ," & sepelio
    StrSql = StrSql & " ," & estimados
    StrSql = StrSql & " ," & otras
    StrSql = StrSql & " ," & donacion
    StrSql = StrSql & " ," & dedesp
    StrSql = StrSql & " ," & noimpo
    StrSql = StrSql & " ," & car_flia
    StrSql = StrSql & " ," & conyuge
    StrSql = StrSql & " ," & hijo
    StrSql = StrSql & " ," & otras_cargas
    StrSql = StrSql & " ," & retenciones
    StrSql = StrSql & " ," & promo
    StrSql = StrSql & " ," & saldo
    StrSql = StrSql & " ," & sindicato
    StrSql = StrSql & " ," & ret_mes
    StrSql = StrSql & " ," & mon_conyuge
    StrSql = StrSql & " ," & mon_hijo
    StrSql = StrSql & " ," & mon_otras
    StrSql = StrSql & " ," & viaticos
    StrSql = StrSql & " ," & amortizacion
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad1)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad2)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad3)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad4)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad5)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad6)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad7)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad8)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad9)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad10)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad11)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad12)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad13)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad14)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad1)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad2)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad3)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad4)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad5)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad6)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad7)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad8)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad9)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad10)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad11)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad12)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad13)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad14)
    StrSql = StrSql & " ," & monto_entidad1
    StrSql = StrSql & " ," & monto_entidad2
    StrSql = StrSql & " ," & monto_entidad3
    StrSql = StrSql & " ," & monto_entidad4
    StrSql = StrSql & " ," & monto_entidad5
    StrSql = StrSql & " ," & monto_entidad6
    StrSql = StrSql & " ," & monto_entidad7
    StrSql = StrSql & " ," & monto_entidad8
    StrSql = StrSql & " ," & monto_entidad9
    StrSql = StrSql & " ," & monto_entidad10
    StrSql = StrSql & " ," & monto_entidad11
    StrSql = StrSql & " ," & monto_entidad12
    StrSql = StrSql & " ," & monto_entidad13
    StrSql = StrSql & " ," & monto_entidad14
    StrSql = StrSql & " ," & ganimpo
    StrSql = StrSql & " ," & ganneta
    StrSql = StrSql & " ," & total_entidad1
    StrSql = StrSql & " ," & total_entidad2
    StrSql = StrSql & " ," & total_entidad3
    StrSql = StrSql & " ," & total_entidad4
    StrSql = StrSql & " ," & total_entidad5
    StrSql = StrSql & " ," & total_entidad6
    StrSql = StrSql & " ," & total_entidad7
    StrSql = StrSql & " ," & total_entidad8
    StrSql = StrSql & " ," & total_entidad9
    StrSql = StrSql & " ," & total_entidad10
    StrSql = StrSql & " ," & total_entidad11
    StrSql = StrSql & " ," & total_entidad12
    StrSql = StrSql & " ," & total_entidad13
    StrSql = StrSql & " ," & total_entidad14
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & imp_deter
    StrSql = StrSql & " ," & eme_medicas
    StrSql = StrSql & " ," & seguro_optativo
    StrSql = StrSql & " ," & seguro_retiro
    StrSql = StrSql & " ," & tope_os_priv
    StrSql = StrSql & " ," & empleg
    StrSql = StrSql & " ," & deducciones
    StrSql = StrSql & " ," & art23
    StrSql = StrSql & " ," & porcdeduc & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procTraza_gan:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procTraza_gan"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procTraza_gan_item_top(ByVal Linea As String)

Dim arrLinea
Dim itenro As String
Dim ternro As String
Dim pronro As String
Dim empresa As String
Dim Monto As String
Dim ddjj As String
Dim old_liq As String
Dim liq As String
Dim prorr As String
Dim ternroAux As Long

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procTraza_gan_item_top

    'Formato tipo, itenro, ternro, pronro, empresa, monto, ddjj, old_liq, liq, prorr

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 9 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    itenro = arrLinea(1)
    ternro = arrLinea(2)
    pronro = arrLinea(3)
    empresa = arrLinea(4)
    Monto = arrLinea(5)
    ddjj = arrLinea(6)
    old_liq = arrLinea(7)
    liq = arrLinea(8)
    prorr = arrLinea(9)

    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT ternro FROM traza_gan_item_top WHERE ternro = " & ternroAux
    StrSql = StrSql & " AND itenro = " & itenro
    StrSql = StrSql & " AND pronro = " & pronro
    StrSql = StrSql & " AND empresa = " & empresa
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE traza_gan_item_top WHERE ternro = " & ternroAux
        StrSql = StrSql & " AND itenro = " & itenro
        StrSql = StrSql & " AND pronro = " & pronro
        StrSql = StrSql & " AND empresa = " & empresa
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO traza_gan_item_top("
    StrSql = StrSql & " itenro, ternro, pronro, empresa, monto, ddjj, old_liq, liq, prorr)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & itenro
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & empresa
    StrSql = StrSql & " ," & Monto
    StrSql = StrSql & " ," & ddjj
    StrSql = StrSql & " ," & old_liq
    StrSql = StrSql & " ," & liq
    StrSql = StrSql & " ," & prorr & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procTraza_gan_item_top:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procTraza_gan_item_top"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procDesmen(ByVal Linea As String)

Dim arrLinea
Dim itenro As String
Dim Empleado As String
Dim desmondec As String
Dim desmenprorra As String
Dim desano As String
Dim desfecdes As String
Dim desfechas As String
Dim descuit As String
Dim desrazsoc As String
Dim pronro As String
Dim ternroAux As Long

Dim rs_Datos As New ADODB.Recordset

On Error GoTo E_procDesmen

    'Formato tipo, itenro , Empleado, desmondec, desmenprorra, desano, desfecdes, desfechas, descuit, desrazsoc, pronro

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 10 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    itenro = arrLinea(1)
    Empleado = arrLinea(2)
    desmondec = arrLinea(3)
    desmenprorra = arrLinea(4)
    desano = arrLinea(5)
    desfecdes = arrLinea(6)
    desfechas = arrLinea(7)
    descuit = arrLinea(8)
    desrazsoc = arrLinea(9)
    pronro = arrLinea(10)

    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(Empleado)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro el ternro " & Empleado
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Si el registro ya existe en la base lo borro
    StrSql = "SELECT empleado FROM desmen WHERE empleado = " & ternroAux
    StrSql = StrSql & " AND itenro = " & itenro
    If EsNulo(desfecdes) Then
        StrSql = StrSql & " AND desfecdes IS NULL "
    Else
        StrSql = StrSql & " AND desfecdes = " & cambiaFecha(desfecdes)
    End If
    If EsNulo(desfechas) Then
        StrSql = StrSql & " AND desfechas IS NULL "
    Else
        StrSql = StrSql & " AND desfechas = " & cambiaFecha(desfechas)
    End If
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        rs_Datos.Close
        StrSql = "DELETE desmen WHERE empleado = " & ternroAux
        StrSql = StrSql & " AND itenro = " & itenro
        If EsNulo(desfecdes) Then
            StrSql = StrSql & " AND desfecdes IS NULL "
        Else
            StrSql = StrSql & " AND desfecdes = " & cambiaFecha(desfecdes)
        End If
        If EsNulo(desfechas) Then
            StrSql = StrSql & " AND desfechas IS NULL "
        Else
            StrSql = StrSql & " AND desfechas = " & cambiaFecha(desfechas)
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    'Inserto Datos
    StrSql = "INSERT INTO desmen("
    StrSql = StrSql & " itenro,empleado,desmondec,desmenprorra,desano,desfecdes,desfechas,descuit,desrazsoc,pronro)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & itenro
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & desmondec
    StrSql = StrSql & " ," & desmenprorra
    StrSql = StrSql & " ," & desano
    StrSql = StrSql & " ," & cambiaFecha(desfecdes)
    StrSql = StrSql & " ," & cambiaFecha(desfechas)
    StrSql = StrSql & " ," & CtrlNuloTXT(descuit)
    StrSql = StrSql & " ," & CtrlNuloTXT(desrazsoc)
    StrSql = StrSql & " ," & pronro & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

Exit Sub

E_procDesmen:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procDesmen"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Function PrimerDiaMes(ByVal mes As Integer, ByVal anio As Long) As Date

    PrimerDiaMes = CDate("01/" & CStr(mes) & "/" & CStr(anio))
    
End Function


Public Function UltimoDiaMes(ByVal mes As Integer, ByVal anio As Long) As Date
Dim aux As Date
    
    'Armo el primer dia del mes siguiente
    If mes = 12 Then
        aux = CDate("01/01/" & CStr(anio + 1))
    Else
        aux = CDate("01/" & CStr(mes + 1) & "/" & CStr(anio))
    End If
    
    'Le resto 1
    UltimoDiaMes = DateAdd("d", -1, aux)
    
End Function


Public Sub procBatch_proceso(ByVal Linea As String)

Dim arrLinea
Dim bpronro As String
Dim btprcnro As String
Dim bprcfecha As String
Dim iduser As String
Dim bprchora As String
Dim bprcempleados As String
Dim bprcfecdesde As String
Dim bprcfechasta As String
Dim bprcestado As String
Dim bprcparam  As String
Dim bprcprogreso As String
Dim bprcfecfin As String
Dim bprchorafin As String
Dim bprctiempo As String
Dim Empnro As String
Dim bprcurgente As String
Dim bprcterminar As String
Dim bprcConfirmado As String
Dim bprcfecInicioEj As String
Dim bprcFecFinEj As String
Dim bprcPid As String
Dim bprcHoraInicioEj As String
Dim bprcHoraFinEj As String
Dim bprctipomodelo As String
Dim bpronroNuevo As Long

On Error GoTo E_procBatch_proceso

    'Formato tipo, bpronro, btprcnro, bprcfecha, iduser, bprchora, bprcempleados, bprcfecdesde, bprcfechasta, bprcestado,
    'bprcparam , bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcurgente, bprcterminar,
    'bprcConfirmado, bprcfecInicioEj, bprcFecFinEj, bprcPid, bprcHoraInicioEj, bprcHoraFinEj, bprctipomodelo

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 24 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    bpronro = arrLinea(1)
    btprcnro = arrLinea(2)
    bprcfecha = arrLinea(3)
    iduser = arrLinea(4)
    bprchora = arrLinea(5)
    bprcempleados = arrLinea(6)
    bprcfecdesde = arrLinea(7)
    bprcfechasta = arrLinea(8)
    bprcestado = arrLinea(9)
    bprcparam = arrLinea(10)
    bprcprogreso = arrLinea(11)
    bprcfecfin = arrLinea(12)
    bprchorafin = arrLinea(13)
    bprctiempo = arrLinea(14)
    Empnro = arrLinea(15)
    bprcurgente = arrLinea(16)
    bprcterminar = arrLinea(17)
    bprcConfirmado = arrLinea(18)
    bprcfecInicioEj = arrLinea(19)
    bprcFecFinEj = arrLinea(20)
    bprcPid = arrLinea(21)
    bprcHoraInicioEj = arrLinea(22)
    bprcHoraFinEj = arrLinea(23)
    bprctipomodelo = arrLinea(24)
    
    MyBeginTrans
    
    'Inserto Datos
    StrSql = "INSERT INTO batch_proceso("
    StrSql = StrSql & " btprcnro, bprcfecha, iduser, bprchora, bprcempleados, bprcfecdesde, bprcfechasta, bprcestado,"
    StrSql = StrSql & " bprcparam , bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcurgente, bprcterminar,"
    StrSql = StrSql & " bprcConfirmado, bprcfecInicioEj, bprcFecFinEj, bprcPid, bprcHoraInicioEj, bprcHoraFinEj, bprctipomodelo)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & btprcnro
    StrSql = StrSql & " ," & cambiaFecha(bprcfecha)
    StrSql = StrSql & " ," & CtrlNuloTXT(iduser)
    StrSql = StrSql & " ," & CtrlNuloTXT(bprchora)
    StrSql = StrSql & " ," & CtrlNuloTXT(bprcempleados)
    StrSql = StrSql & " ," & cambiaFecha(bprcfecdesde)
    StrSql = StrSql & " ," & cambiaFecha(bprcfechasta)
    StrSql = StrSql & " ," & CtrlNuloTXT(bprcestado)
    StrSql = StrSql & " ," & CtrlNuloTXT(bprcparam)
    StrSql = StrSql & " ," & bprcprogreso
    StrSql = StrSql & " ," & cambiaFecha(bprcfecfin)
    StrSql = StrSql & " ," & CtrlNuloTXT(bprchorafin)
    StrSql = StrSql & " ," & CtrlNuloTXT(bprctiempo)
    StrSql = StrSql & " ," & Empnro
    StrSql = StrSql & " ," & bprcurgente
    StrSql = StrSql & " ," & bprcterminar
    StrSql = StrSql & " ," & bprcConfirmado
    StrSql = StrSql & " ," & cambiaFecha(bprcfecInicioEj)
    StrSql = StrSql & " ," & cambiaFecha(bprcFecFinEj)
    StrSql = StrSql & " ," & bprcPid
    StrSql = StrSql & " ," & CtrlNuloTXT(bprcHoraInicioEj)
    StrSql = StrSql & " ," & CtrlNuloTXT(bprcHoraFinEj)
    StrSql = StrSql & " ," & bprctipomodelo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Recupero el codigo del insertado
    bpronroNuevo = getLastIdentity(objConn, "batch_proceso")
    
    'Guardo el nuevo valor en el mapeo
    StrSql = "INSERT INTO tmp_batch_proc("
    StrSql = StrSql & " bpronro_ori ,bpronro_dest)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & bpronro
    StrSql = StrSql & " ," & bpronroNuevo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    MyCommitTrans

Exit Sub

E_procBatch_proceso:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procBatch_proceso"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procRep_recibo(ByVal Linea As String)

Dim arrLinea
Dim bpronro As String
Dim ternro As String
Dim pronro As String
Dim apellido As String
Dim nombre As String
Dim direccion As String
Dim Legajo As String
Dim pliqnro As String
Dim pliqmes As String
Dim pliqanio As String
Dim pliqdepant As String
Dim pliqfecdep As String
Dim pliqbco As String
Dim empfecalta As String
Dim sueldo As String
Dim categoria As String
Dim centrocosto As String
Dim localidad As String
Dim profecpago As String
Dim formapago As String
Dim empnombre As String
Dim empdire As String
Dim empcuit As String
Dim emplogo As String
Dim emplogoalto As String
Dim Cuil As String
Dim emplogoancho As String
Dim empfirma As String
Dim empfirmaalto As String
Dim empfirmaancho As String
Dim prodesc As String
Dim descripcion As String
Dim puesto As String
Dim tenro1 As String
Dim tenro2 As String
Dim tenro3 As String
Dim estrnro1 As String
Dim estrnro2 As String
Dim estrnro3 As String
Dim Orden As String
Dim Auxchar1 As String
Dim Auxchar2 As String
Dim Auxchar3 As String
Dim Auxchar4 As String
Dim Auxchar5 As String
Dim auxdeci1 As String
Dim auxdeci2 As String
Dim auxdeci3 As String
Dim auxdeci4 As String
Dim auxdeci5 As String
Dim ternroAux As Long
Dim bpronroAux As Long


On Error GoTo E_procRep_recibo

    'Formato tipo, bpronro, ternro, pronro, apellido, nombre, direccion, legajo, pliqnro, pliqmes,
    'pliqanio, pliqdepant, pliqfecdep, pliqbco, empfecalta, sueldo, categoria, centrocosto,
    'localidad, profecpago, formapago, empnombre, empdire, empcuit, emplogo, emplogoalto, cuil,
    'emplogoancho, empfirma, empfirmaalto, empfirmaancho, prodesc, descripcion, puesto, tenro1, tenro2,
    'tenro3, estrnro1, estrnro2, estrnro3, orden, auxchar1, auxchar2, auxchar3, auxchar4, auxchar5,
    'auxdeci1, auxdeci2, auxdeci3, auxdeci4, auxdeci5

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 50 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    bpronro = arrLinea(1)
    ternro = arrLinea(2)
    pronro = arrLinea(3)
    apellido = arrLinea(4)
    nombre = arrLinea(5)
    direccion = arrLinea(6)
    Legajo = arrLinea(7)
    pliqnro = arrLinea(8)
    pliqmes = arrLinea(9)
    pliqanio = arrLinea(10)
    pliqdepant = arrLinea(11)
    pliqfecdep = arrLinea(12)
    pliqbco = arrLinea(13)
    empfecalta = arrLinea(14)
    sueldo = arrLinea(15)
    categoria = arrLinea(16)
    centrocosto = arrLinea(17)
    localidad = arrLinea(18)
    profecpago = arrLinea(19)
    formapago = arrLinea(20)
    empnombre = arrLinea(21)
    empdire = arrLinea(22)
    empcuit = arrLinea(23)
    emplogo = arrLinea(24)
    emplogoalto = arrLinea(25)
    Cuil = arrLinea(26)
    emplogoancho = arrLinea(27)
    empfirma = arrLinea(28)
    empfirmaalto = arrLinea(29)
    empfirmaancho = arrLinea(30)
    prodesc = arrLinea(31)
    descripcion = arrLinea(32)
    puesto = arrLinea(33)
    tenro1 = arrLinea(34)
    tenro2 = arrLinea(35)
    tenro3 = arrLinea(36)
    estrnro1 = arrLinea(37)
    estrnro2 = arrLinea(38)
    estrnro3 = arrLinea(39)
    Orden = arrLinea(40)
    Auxchar1 = arrLinea(41)
    Auxchar2 = arrLinea(42)
    Auxchar3 = arrLinea(43)
    Auxchar4 = arrLinea(44)
    Auxchar5 = arrLinea(45)
    auxdeci1 = arrLinea(46)
    auxdeci2 = arrLinea(47)
    auxdeci3 = arrLinea(48)
    auxdeci4 = arrLinea(49)
    auxdeci5 = arrLinea(50)
    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Busco el bpronro nuevo
    bpronroAux = mapearBpronro(bpronro)
    
    If bpronroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el bpronro Origen " & bpronro & " en la tabla de mapeos."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO rep_recibo("
    StrSql = StrSql & " bpronro, ternro, pronro, apellido, nombre, direccion, legajo, pliqnro, pliqmes,"
    StrSql = StrSql & " pliqanio, pliqdepant, pliqfecdep, pliqbco, empfecalta, sueldo, categoria, centrocosto,"
    StrSql = StrSql & " localidad, profecpago, formapago, empnombre, empdire, empcuit, emplogo, emplogoalto, cuil,"
    StrSql = StrSql & " emplogoancho, empfirma, empfirmaalto, empfirmaancho, prodesc, descripcion, puesto, tenro1, tenro2,"
    StrSql = StrSql & " tenro3, estrnro1, estrnro2, estrnro3, orden, auxchar1, auxchar2, auxchar3, auxchar4, auxchar5,"
    StrSql = StrSql & " auxdeci1, auxdeci2, auxdeci3, auxdeci4, auxdeci5)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & bpronroAux
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & CtrlNuloTXT(apellido)
    StrSql = StrSql & " ," & CtrlNuloTXT(nombre)
    StrSql = StrSql & " ," & CtrlNuloTXT(direccion)
    StrSql = StrSql & " ," & Legajo
    StrSql = StrSql & " ," & pliqnro
    StrSql = StrSql & " ," & pliqmes
    StrSql = StrSql & " ," & pliqanio
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqdepant)
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqfecdep)
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqbco)
    StrSql = StrSql & " ," & CtrlNuloTXT(empfecalta)
    StrSql = StrSql & " ," & sueldo
    StrSql = StrSql & " ," & CtrlNuloTXT(categoria)
    StrSql = StrSql & " ," & CtrlNuloTXT(centrocosto)
    StrSql = StrSql & " ," & CtrlNuloTXT(localidad)
    StrSql = StrSql & " ," & CtrlNuloTXT(profecpago)
    StrSql = StrSql & " ," & CtrlNuloTXT(formapago)
    StrSql = StrSql & " ," & CtrlNuloTXT(empnombre)
    StrSql = StrSql & " ," & CtrlNuloTXT(empdire)
    StrSql = StrSql & " ," & CtrlNuloTXT(empcuit)
    StrSql = StrSql & " ," & CtrlNuloTXT(emplogo)
    StrSql = StrSql & " ," & emplogoalto
    StrSql = StrSql & " ," & CtrlNuloTXT(Cuil)
    StrSql = StrSql & " ," & emplogoancho
    StrSql = StrSql & " ," & CtrlNuloTXT(empfirma)
    StrSql = StrSql & " ," & empfirmaalto
    StrSql = StrSql & " ," & empfirmaancho
    StrSql = StrSql & " ," & CtrlNuloTXT(prodesc)
    StrSql = StrSql & " ," & CtrlNuloTXT(descripcion)
    StrSql = StrSql & " ," & CtrlNuloTXT(puesto)
    StrSql = StrSql & " ," & tenro1
    StrSql = StrSql & " ," & tenro2
    StrSql = StrSql & " ," & tenro3
    StrSql = StrSql & " ," & estrnro1
    StrSql = StrSql & " ," & estrnro2
    StrSql = StrSql & " ," & estrnro3
    StrSql = StrSql & " ," & Orden
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar1)
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar2)
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar3)
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar4)
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar5)
    StrSql = StrSql & " ," & auxdeci1
    StrSql = StrSql & " ," & auxdeci2
    StrSql = StrSql & " ," & auxdeci3
    StrSql = StrSql & " ," & auxdeci4
    StrSql = StrSql & " ," & auxdeci5 & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procRep_recibo:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procRep_recibo"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procRep_recibo_det(ByVal Linea As String)
Dim arrLinea
Dim bpronro As String
Dim ternro As String
Dim pronro As String
Dim cliqnro As String
Dim concabr As String
Dim Conccod As String
Dim concnro As String
Dim tconnro As String
Dim concimp As String
Dim dlicant As String
Dim dlimonto As String
Dim conctipo As String
Dim ternroAux As Long
Dim bpronroAux As Long


On Error GoTo E_procRep_recibo_det

    'Formato tipo, bpronro, ternro, pronro, cliqnro, concabr, conccod, concnro, tconnro, concimp, dlicant,
    'dlimonto, conctipo"

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 12 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    bpronro = arrLinea(1)
    ternro = arrLinea(2)
    pronro = arrLinea(3)
    cliqnro = arrLinea(4)
    concabr = arrLinea(5)
    Conccod = arrLinea(6)
    concnro = arrLinea(7)
    tconnro = arrLinea(8)
    concimp = arrLinea(9)
    dlicant = arrLinea(10)
    dlimonto = arrLinea(11)
    conctipo = arrLinea(12)
    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Busco el bpronro nuevo
    bpronroAux = mapearBpronro(bpronro)
    
    If bpronroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el bpronro Origen " & bpronro & " en la tabla de mapeos."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO rep_recibo_det("
    StrSql = StrSql & " bpronro, ternro, pronro, cliqnro, concabr, conccod, concnro, tconnro, concimp, dlicant,"
    StrSql = StrSql & " dlimonto, conctipo)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & bpronroAux
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & cliqnro
    StrSql = StrSql & " ," & CtrlNuloTXT(concabr)
    StrSql = StrSql & " ," & CtrlNuloTXT(Conccod)
    StrSql = StrSql & " ," & concnro
    StrSql = StrSql & " ," & tconnro
    StrSql = StrSql & " ," & concimp
    StrSql = StrSql & " ," & dlicant
    StrSql = StrSql & " ," & dlimonto
    StrSql = StrSql & " ," & CtrlNuloTXT(conctipo) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procRep_recibo_det:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procRep_recibo_det"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub



Public Sub procRep_libroley(ByVal Linea As String)
Dim arrLinea
Dim bpronro As String
Dim Legajo As String
Dim ternro As String
Dim pronro As String
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim empresa As String
Dim emprnro As String
Dim pliqnro As String
Dim fecpago As String
Dim fecalta As String
Dim fecbaja As String
Dim contrato As String
Dim categoria As String
Dim direccion As String
Dim puesto As String
Dim documento As String
Dim fecha_nac As String
Dim est_civil As String
Dim Cuil As String
Dim estado As String
Dim reg_prev As String
Dim lug_trab As String
Dim basico As String
Dim neto As String
Dim msr As String
Dim asi_flia As String
Dim dtos As String
Dim bruto As String
Dim prodesc As String
Dim descripcion As String
Dim pliqdesc As String
Dim pliqmes As String
Dim pliqanio As String
Dim profecpago As String
Dim pliqfecdep As String
Dim pliqbco As String
Dim ultima_pag_impr As String
Dim Auxchar1 As String
Dim Auxchar2 As String
Dim Auxchar3 As String
Dim Auxchar4 As String
Dim Auxchar5 As String
Dim auxdeci1 As String
Dim auxdeci2 As String
Dim auxdeci3 As String
Dim auxdeci4 As String
Dim auxdeci5 As String
Dim Orden As String
Dim tedabr1 As String
Dim tedabr2 As String
Dim tedabr3 As String
Dim estrdabr1 As String
Dim estrdabr2 As String
Dim estrdabr3 As String
Dim tipofam As String
Dim ternroAux As Long
Dim bpronroAux As Long


On Error GoTo E_procRep_libroley

    'Formato tipo, bpronro, legajo, ternro, pronro, apellido, apellido2, nombre, nombre2, empresa, emprnro,
    'pliqnro, fecpago, fecalta, fecbaja, contrato, categoria, direccion, puesto, documento, fecha_nac,
    'est_civil, cuil, estado, reg_prev, lug_trab, basico, neto, msr, asi_flia, dtos, bruto, prodesc,
    'descripcion, pliqdesc, pliqmes, pliqanio, profecpago, pliqfecdep, pliqbco, ultima_pag_impr, auxchar1,
    'auxchar2, auxchar3, auxchar4, auxchar5, auxdeci1, auxdeci2, auxdeci3, auxdeci4, auxdeci5, orden,
    'tedabr1 , tedabr2, tedabr3, estrdabr1, estrdabr2, estrdabr3, tipofam

    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 58 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    bpronro = arrLinea(1)
    Legajo = arrLinea(2)
    ternro = arrLinea(3)
    pronro = arrLinea(4)
    apellido = arrLinea(5)
    apellido2 = arrLinea(6)
    nombre = arrLinea(7)
    nombre2 = arrLinea(8)
    empresa = arrLinea(9)
    emprnro = arrLinea(10)
    pliqnro = arrLinea(11)
    fecpago = arrLinea(12)
    fecalta = arrLinea(13)
    fecbaja = arrLinea(14)
    contrato = arrLinea(15)
    categoria = arrLinea(16)
    direccion = arrLinea(17)
    puesto = arrLinea(18)
    documento = arrLinea(19)
    fecha_nac = arrLinea(20)
    est_civil = arrLinea(21)
    Cuil = arrLinea(22)
    estado = arrLinea(23)
    reg_prev = arrLinea(24)
    lug_trab = arrLinea(25)
    basico = arrLinea(26)
    neto = arrLinea(27)
    msr = arrLinea(28)
    asi_flia = arrLinea(29)
    dtos = arrLinea(30)
    bruto = arrLinea(31)
    prodesc = arrLinea(32)
    descripcion = arrLinea(33)
    pliqdesc = arrLinea(34)
    pliqmes = arrLinea(35)
    pliqanio = arrLinea(36)
    profecpago = arrLinea(37)
    pliqfecdep = arrLinea(38)
    pliqbco = arrLinea(39)
    ultima_pag_impr = arrLinea(40)
    Auxchar1 = arrLinea(41)
    Auxchar2 = arrLinea(42)
    Auxchar3 = arrLinea(43)
    Auxchar4 = arrLinea(44)
    Auxchar5 = arrLinea(45)
    auxdeci1 = arrLinea(46)
    auxdeci2 = arrLinea(47)
    auxdeci3 = arrLinea(48)
    auxdeci4 = arrLinea(49)
    auxdeci5 = arrLinea(50)
    Orden = arrLinea(51)
    tedabr1 = arrLinea(52)
    tedabr2 = arrLinea(53)
    tedabr3 = arrLinea(54)
    estrdabr1 = arrLinea(55)
    estrdabr2 = arrLinea(56)
    estrdabr3 = arrLinea(57)
    tipofam = arrLinea(58)
    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Busco el bpronro nuevo
    bpronroAux = mapearBpronro(bpronro)
    
    If bpronroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el bpronro Origen " & bpronro & " en la tabla de mapeos."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO rep_libroley("
    StrSql = StrSql & " bpronro, legajo, ternro, pronro, apellido, apellido2, nombre, nombre2, empresa, emprnro,"
    StrSql = StrSql & " pliqnro, fecpago, fecalta, fecbaja, contrato, categoria, direccion, puesto, documento, fecha_nac,"
    StrSql = StrSql & " est_civil, cuil, estado, reg_prev, lug_trab, basico, neto, msr, asi_flia, dtos, bruto, prodesc,"
    StrSql = StrSql & " descripcion, pliqdesc, pliqmes, pliqanio, profecpago, pliqfecdep, pliqbco, ultima_pag_impr, auxchar1,"
    StrSql = StrSql & " auxchar2, auxchar3, auxchar4, auxchar5, auxdeci1, auxdeci2, auxdeci3, auxdeci4, auxdeci5, orden,"
    StrSql = StrSql & " tedabr1 , tedabr2, tedabr3, estrdabr1, estrdabr2, estrdabr3, tipofam)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & bpronroAux
    StrSql = StrSql & " ," & Legajo
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & CtrlNuloTXT(apellido)
    StrSql = StrSql & " ," & CtrlNuloTXT(apellido2)
    StrSql = StrSql & " ," & CtrlNuloTXT(nombre)
    StrSql = StrSql & " ," & CtrlNuloTXT(nombre2)
    StrSql = StrSql & " ," & CtrlNuloTXT(empresa)
    StrSql = StrSql & " ," & emprnro
    StrSql = StrSql & " ," & pliqnro
    StrSql = StrSql & " ," & CtrlNuloTXT(fecpago)
    StrSql = StrSql & " ," & CtrlNuloTXT(fecalta)
    StrSql = StrSql & " ," & CtrlNuloTXT(fecbaja)
    StrSql = StrSql & " ," & CtrlNuloTXT(contrato)
    StrSql = StrSql & " ," & CtrlNuloTXT(categoria)
    StrSql = StrSql & " ," & CtrlNuloTXT(direccion)
    StrSql = StrSql & " ," & CtrlNuloTXT(puesto)
    StrSql = StrSql & " ," & CtrlNuloTXT(documento)
    StrSql = StrSql & " ," & CtrlNuloTXT(fecha_nac)
    StrSql = StrSql & " ," & CtrlNuloTXT(est_civil)
    StrSql = StrSql & " ," & CtrlNuloTXT(Cuil)
    StrSql = StrSql & " ," & estado
    StrSql = StrSql & " ," & CtrlNuloTXT(reg_prev)
    StrSql = StrSql & " ," & CtrlNuloTXT(lug_trab)
    StrSql = StrSql & " ," & basico
    StrSql = StrSql & " ," & neto
    StrSql = StrSql & " ," & msr
    StrSql = StrSql & " ," & asi_flia
    StrSql = StrSql & " ," & dtos
    StrSql = StrSql & " ," & bruto
    StrSql = StrSql & " ," & CtrlNuloTXT(prodesc)
    StrSql = StrSql & " ," & CtrlNuloTXT(descripcion)
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqdesc)
    StrSql = StrSql & " ," & pliqmes
    StrSql = StrSql & " ," & pliqanio
    StrSql = StrSql & " ," & cambiaFecha(profecpago)
    StrSql = StrSql & " ," & cambiaFecha(pliqfecdep)
    StrSql = StrSql & " ," & CtrlNuloTXT(pliqbco)
    StrSql = StrSql & " ," & ultima_pag_impr
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar1)
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar2)
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar3)
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar4)
    StrSql = StrSql & " ," & CtrlNuloTXT(Auxchar5)
    StrSql = StrSql & " ," & auxdeci1
    StrSql = StrSql & " ," & auxdeci2
    StrSql = StrSql & " ," & auxdeci3
    StrSql = StrSql & " ," & auxdeci4
    StrSql = StrSql & " ," & auxdeci5
    StrSql = StrSql & " ," & Orden
    StrSql = StrSql & " ," & CtrlNuloTXT(tedabr1)
    StrSql = StrSql & " ," & CtrlNuloTXT(tedabr2)
    StrSql = StrSql & " ," & CtrlNuloTXT(tedabr3)
    StrSql = StrSql & " ," & CtrlNuloTXT(estrdabr1)
    StrSql = StrSql & " ," & CtrlNuloTXT(estrdabr2)
    StrSql = StrSql & " ," & CtrlNuloTXT(estrdabr3)
    StrSql = StrSql & " ," & tipofam & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procRep_libroley:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procRep_libroley"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procRep_libroley_det(ByVal Linea As String)
Dim arrLinea
Dim bpronro As String
Dim ternro As String
Dim pronro As String
Dim concabr As String
Dim Conccod As String
Dim concnro As String
Dim concimp As String
Dim dlicant As String
Dim dlimonto As String
Dim conctipo As String
Dim ternroAux As Long
Dim bpronroAux As Long


On Error GoTo E_procRep_libroley_det

    'Formato tipo, bpronro, ternro, pronro, concabr, conccod, concnro, concimp, dlicant, dlimonto, conctipo
    
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 10 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    bpronro = arrLinea(1)
    ternro = arrLinea(2)
    pronro = arrLinea(3)
    concabr = arrLinea(4)
    Conccod = arrLinea(5)
    concnro = arrLinea(6)
    concimp = arrLinea(7)
    dlicant = arrLinea(8)
    dlimonto = arrLinea(9)
    conctipo = arrLinea(10)

    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Busco el bpronro nuevo
    bpronroAux = mapearBpronro(bpronro)
    
    If bpronroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el bpronro Origen " & bpronro & " en la tabla de mapeos."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO rep_libroley_det("
    StrSql = StrSql & " bpronro, ternro, pronro, concabr, conccod, concnro, concimp, dlicant, dlimonto, conctipo)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & bpronroAux
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & CtrlNuloTXT(concabr)
    StrSql = StrSql & " ," & CtrlNuloTXT(Conccod)
    StrSql = StrSql & " ," & concnro
    StrSql = StrSql & " ," & concimp
    StrSql = StrSql & " ," & dlicant
    StrSql = StrSql & " ," & dlimonto
    StrSql = StrSql & " ," & CtrlNuloTXT(conctipo) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procRep_libroley_det:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procRep_libroley_det"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procRep_libroley_fam(ByVal Linea As String)
Dim arrLinea
Dim bpronro As String
Dim ternro As String
Dim pronro As String
Dim nrodoc As String
Dim sigladoc As String
Dim ternrofam As String
Dim terape As String
Dim ternom As String
Dim terfecnac As String
Dim tersex As String
Dim famest As String
Dim famtrab As String
Dim faminc As String
Dim famsalario As String
Dim famfecvto As String
Dim famCargaDGI As String
Dim famDGIdesde As String
Dim famDGIhasta As String
Dim famemergencia As String
Dim paredesc As String
Dim ternroAux As Long
Dim bpronroAux As Long


On Error GoTo E_procRep_libroley_fam

    'Formato tipo, bpronro, ternro, pronro, nrodoc, sigladoc, ternrofam, terape, ternom, terfecnac, tersex,
    'famest, famtrab, faminc, famsalario, famfecvto, famCargaDGI, famDGIdesde, famDGIhasta,
    'famemergencia, paredesc
    
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 20 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    bpronro = arrLinea(1)
    ternro = arrLinea(2)
    pronro = arrLinea(3)
    nrodoc = arrLinea(4)
    sigladoc = arrLinea(5)
    ternrofam = arrLinea(6)
    terape = arrLinea(7)
    ternom = arrLinea(8)
    terfecnac = arrLinea(9)
    tersex = arrLinea(10)
    famest = arrLinea(11)
    famtrab = arrLinea(12)
    faminc = arrLinea(13)
    famsalario = arrLinea(14)
    famfecvto = arrLinea(15)
    famCargaDGI = arrLinea(16)
    famDGIdesde = arrLinea(17)
    famDGIhasta = arrLinea(18)
    famemergencia = arrLinea(19)
    paredesc = arrLinea(20)
    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Busco el bpronro nuevo
    bpronroAux = mapearBpronro(bpronro)
    
    If bpronroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el bpronro Origen " & bpronro & " en la tabla de mapeos."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO rep_libroley_fam("
    StrSql = StrSql & " bpronro, ternro, pronro, nrodoc, sigladoc, ternrofam, terape, ternom, terfecnac, tersex,"
    StrSql = StrSql & " famest, famtrab, faminc, famsalario, famfecvto, famCargaDGI, famDGIdesde, famDGIhasta,"
    StrSql = StrSql & " famemergencia, paredesc)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & bpronroAux
    StrSql = StrSql & " ," & ternroAux
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & CtrlNuloTXT(nrodoc)
    StrSql = StrSql & " ," & CtrlNuloTXT(sigladoc)
    StrSql = StrSql & " ," & ternrofam
    StrSql = StrSql & " ," & CtrlNuloTXT(terape)
    StrSql = StrSql & " ," & CtrlNuloTXT(ternom)
    StrSql = StrSql & " ," & cambiaFecha(terfecnac)
    StrSql = StrSql & " ," & tersex
    StrSql = StrSql & " ," & famest
    StrSql = StrSql & " ," & famtrab
    StrSql = StrSql & " ," & faminc
    StrSql = StrSql & " ," & famsalario
    StrSql = StrSql & " ," & cambiaFecha(famfecvto)
    StrSql = StrSql & " ," & famCargaDGI
    StrSql = StrSql & " ," & cambiaFecha(famDGIdesde)
    StrSql = StrSql & " ," & cambiaFecha(famDGIhasta)
    StrSql = StrSql & " ," & famemergencia
    StrSql = StrSql & " ," & CtrlNuloTXT(paredesc) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procRep_libroley_fam:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procRep_libroley_fam"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procF649(ByVal Linea As String)
Dim arrLinea
Dim repnro As String
Dim pliqnro As String
Dim pronro As String
Dim empresa As String
Dim bpronro As String
Dim Fecha As String
Dim Hora As String
Dim iduser As String
Dim empleg As String
Dim msr As String
Dim nomsr As String
Dim nogan As String
Dim jubilacion As String
Dim osocial As String
Dim cuota_medico As String
Dim prima_seguro As String
Dim sepelio As String
Dim estimados As String
Dim otras As String
Dim donacion As String
Dim dedesp As String
Dim noimpo As String
Dim car_flia As String
Dim conyuge As String
Dim hijo As String
Dim otras_cargas As String
Dim retenciones As String
Dim promo As String
Dim saldo As String
Dim Desde As String
Dim Hasta As String
Dim Cuil As String
Dim direccion As String
Dim sindicato As String
Dim ret_mes As String
Dim dir_calle As String
Dim dir_num As String
Dim dir_piso As String
Dim dir_dpto As String
Dim dir_localidad As String
Dim dir_pcia As String
Dim dir_cp As String
Dim cuit As String
Dim monto_letras As String
Dim emp_nombre As String
Dim emp_cuit As String
Dim ano As String
Dim mon_conyuge As String
Dim mon_hijo As String
Dim mon_otras As String
Dim viaticos As String
Dim amortizacion As String
Dim entidad1 As String
Dim entidad2 As String
Dim entidad3 As String
Dim entidad4 As String
Dim entidad5 As String
Dim entidad6 As String
Dim entidad7 As String
Dim entidad8 As String
Dim entidad9 As String
Dim entidad10 As String
Dim entidad11 As String
Dim entidad12 As String
Dim entidad13 As String
Dim entidad14 As String
Dim cuit_entidad1 As String
Dim cuit_entidad2 As String
Dim cuit_entidad3 As String
Dim cuit_entidad4 As String
Dim cuit_entidad5 As String
Dim cuit_entidad6 As String
Dim cuit_entidad7 As String
Dim cuit_entidad8 As String
Dim cuit_entidad9 As String
Dim cuit_entidad10 As String
Dim cuit_entidad11 As String
Dim cuit_entidad12 As String
Dim cuit_entidad13 As String
Dim cuit_entidad14 As String
Dim monto_entidad1 As String
Dim monto_entidad2 As String
Dim monto_entidad3 As String
Dim monto_entidad4 As String
Dim monto_entidad5 As String
Dim monto_entidad6 As String
Dim monto_entidad7 As String
Dim monto_entidad8 As String
Dim monto_entidad9 As String
Dim monto_entidad10 As String
Dim monto_entidad11 As String
Dim monto_entidad12 As String
Dim monto_entidad13 As String
Dim monto_entidad14 As String
Dim ganimpo As String
Dim ganneta As String
Dim total_entidad1 As String
Dim total_entidad2 As String
Dim total_entidad3 As String
Dim total_entidad4 As String
Dim total_entidad5 As String
Dim total_entidad6 As String
Dim total_entidad7 As String
Dim total_entidad8 As String
Dim total_entidad9 As String
Dim total_entidad10 As String
Dim total_entidad11 As String
Dim total_entidad12 As String
Dim total_entidad13 As String
Dim total_entidad14 As String
Dim imp_deter As String
Dim eme_medicas As String
Dim seguro_optativo As String
Dim seguro_retiro As String
Dim tope_os_priv As String
Dim prorratea As String
Dim suscribe As String
Dim caracter As String
Dim fecha_caracter As String
Dim fecha_devolucion As String
Dim dependencia_dgi As String
Dim anual_final As String
Dim ternro As String
Dim terape As String
Dim ternom As String
Dim ternroAux As Long
Dim bpronroAux As Long


On Error GoTo E_procF649

    'Formato tipo, repnro, pliqnro, pronro, empresa, bpronro, Fecha, Hora, iduser, empleg, msr, nomsr,
    'nogan, jubilacion, osocial, cuota_medico, prima_seguro, sepelio, estimados, otras, donacion,
    'dedesp, noimpo, car_flia, conyuge, hijo, otras_cargas, retenciones, promo, saldo, desde, hasta,
    'cuil, direccion, sindicato, ret_mes, dir_calle, dir_num, dir_piso, dir_dpto, dir_localidad,
    'dir_pcia, dir_cp, cuit, monto_letras, emp_nombre, emp_cuit, ano, mon_conyuge, mon_hijo, mon_otras,
    'viaticos, amortizacion, entidad1, entidad2, entidad3, entidad4, entidad5, entidad6, entidad7,
    'entidad8, entidad9, entidad10, entidad11, entidad12, entidad13, entidad14, cuit_entidad1,
    'cuit_entidad2, cuit_entidad3, cuit_entidad4, cuit_entidad5, cuit_entidad6, cuit_entidad7, cuit_entidad8,
    'cuit_entidad9, cuit_entidad10, cuit_entidad11, cuit_entidad12, cuit_entidad13, cuit_entidad14,
    'monto_entidad1, monto_entidad2, monto_entidad3, monto_entidad4, monto_entidad5, monto_entidad6,
    'monto_entidad7, monto_entidad8, monto_entidad9, monto_entidad10, monto_entidad11, monto_entidad12,
    'monto_entidad13, monto_entidad14, ganimpo, ganneta, total_entidad1, total_entidad2, total_entidad3,
    'total_entidad4, total_entidad5, total_entidad6, total_entidad7, total_entidad8, total_entidad9,
    'total_entidad10, total_entidad11, total_entidad12, total_entidad13, total_entidad14, imp_deter,
    'eme_medicas, seguro_optativo, seguro_retiro, tope_os_priv, prorratea, suscribe, caracter, fecha_caracter,
    'fecha_devolucion, dependencia_dgi, anual_final, ternro, terape, ternom
    
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 125 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    repnro = arrLinea(1)
    pliqnro = arrLinea(2)
    pronro = arrLinea(3)
    empresa = arrLinea(4)
    bpronro = arrLinea(5)
    Fecha = arrLinea(6)
    Hora = arrLinea(7)
    iduser = arrLinea(8)
    empleg = arrLinea(9)
    msr = arrLinea(10)
    nomsr = arrLinea(11)
    nogan = arrLinea(12)
    jubilacion = arrLinea(13)
    osocial = arrLinea(14)
    cuota_medico = arrLinea(15)
    prima_seguro = arrLinea(16)
    sepelio = arrLinea(17)
    estimados = arrLinea(18)
    otras = arrLinea(19)
    donacion = arrLinea(20)
    dedesp = arrLinea(21)
    noimpo = arrLinea(22)
    car_flia = arrLinea(23)
    conyuge = arrLinea(24)
    hijo = arrLinea(25)
    otras_cargas = arrLinea(26)
    retenciones = arrLinea(27)
    promo = arrLinea(28)
    saldo = arrLinea(29)
    Desde = arrLinea(30)
    Hasta = arrLinea(31)
    Cuil = arrLinea(32)
    direccion = arrLinea(33)
    sindicato = arrLinea(34)
    ret_mes = arrLinea(35)
    dir_calle = arrLinea(36)
    dir_num = arrLinea(37)
    dir_piso = arrLinea(38)
    dir_dpto = arrLinea(39)
    dir_localidad = arrLinea(40)
    dir_pcia = arrLinea(41)
    dir_cp = arrLinea(42)
    cuit = arrLinea(43)
    monto_letras = arrLinea(44)
    emp_nombre = arrLinea(45)
    emp_cuit = arrLinea(46)
    ano = arrLinea(47)
    mon_conyuge = arrLinea(48)
    mon_hijo = arrLinea(49)
    mon_otras = arrLinea(50)
    viaticos = arrLinea(51)
    amortizacion = arrLinea(52)
    entidad1 = arrLinea(53)
    entidad2 = arrLinea(54)
    entidad3 = arrLinea(55)
    entidad4 = arrLinea(56)
    entidad5 = arrLinea(57)
    entidad6 = arrLinea(58)
    entidad7 = arrLinea(59)
    entidad8 = arrLinea(60)
    entidad9 = arrLinea(61)
    entidad10 = arrLinea(62)
    entidad11 = arrLinea(63)
    entidad12 = arrLinea(64)
    entidad13 = arrLinea(65)
    entidad14 = arrLinea(66)
    cuit_entidad1 = arrLinea(67)
    cuit_entidad2 = arrLinea(68)
    cuit_entidad3 = arrLinea(69)
    cuit_entidad4 = arrLinea(70)
    cuit_entidad5 = arrLinea(71)
    cuit_entidad6 = arrLinea(72)
    cuit_entidad7 = arrLinea(73)
    cuit_entidad8 = arrLinea(74)
    cuit_entidad9 = arrLinea(75)
    cuit_entidad10 = arrLinea(76)
    cuit_entidad11 = arrLinea(77)
    cuit_entidad12 = arrLinea(78)
    cuit_entidad13 = arrLinea(79)
    cuit_entidad14 = arrLinea(80)
    monto_entidad1 = arrLinea(81)
    monto_entidad2 = arrLinea(82)
    monto_entidad3 = arrLinea(83)
    monto_entidad4 = arrLinea(84)
    monto_entidad5 = arrLinea(85)
    monto_entidad6 = arrLinea(86)
    monto_entidad7 = arrLinea(87)
    monto_entidad8 = arrLinea(88)
    monto_entidad9 = arrLinea(89)
    monto_entidad10 = arrLinea(90)
    monto_entidad11 = arrLinea(91)
    monto_entidad12 = arrLinea(92)
    monto_entidad13 = arrLinea(93)
    monto_entidad14 = arrLinea(94)
    ganimpo = arrLinea(95)
    ganneta = arrLinea(96)
    total_entidad1 = arrLinea(97)
    total_entidad2 = arrLinea(98)
    total_entidad3 = arrLinea(99)
    total_entidad4 = arrLinea(100)
    total_entidad5 = arrLinea(101)
    total_entidad6 = arrLinea(102)
    total_entidad7 = arrLinea(103)
    total_entidad8 = arrLinea(104)
    total_entidad9 = arrLinea(105)
    total_entidad10 = arrLinea(106)
    total_entidad11 = arrLinea(107)
    total_entidad12 = arrLinea(108)
    total_entidad13 = arrLinea(109)
    total_entidad14 = arrLinea(110)
    imp_deter = arrLinea(111)
    eme_medicas = arrLinea(112)
    seguro_optativo = arrLinea(113)
    seguro_retiro = arrLinea(114)
    tope_os_priv = arrLinea(115)
    prorratea = arrLinea(116)
    suscribe = arrLinea(117)
    caracter = arrLinea(118)
    fecha_caracter = arrLinea(119)
    fecha_devolucion = arrLinea(120)
    dependencia_dgi = arrLinea(121)
    anual_final = arrLinea(122)
    ternro = arrLinea(123)
    terape = arrLinea(124)
    ternom = arrLinea(125)
    
    'Busco el ternro del empleado en la base destino
    ternroAux = mapearTenro(ternro)
    
    If ternroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el mapeo del ternro  origen " & ternro & " en la base destino."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Busco el bpronro nuevo
    bpronroAux = mapearBpronro(bpronro)
    
    If bpronroAux = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el bpronro Origen " & bpronro & " en la tabla de mapeos."
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    
    'Inserto Datos
    StrSql = "INSERT INTO rep19("
    StrSql = StrSql & " pliqnro, pronro, empresa, bpronro, Fecha, Hora, iduser, empleg, msr, nomsr,"
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
    StrSql = StrSql & " monto_entidad13, monto_entidad14, ganimpo, ganneta, total_entidad1, total_entidad2, total_entidad3,"
    StrSql = StrSql & " total_entidad4, total_entidad5, total_entidad6, total_entidad7, total_entidad8, total_entidad9,"
    StrSql = StrSql & " total_entidad10, total_entidad11, total_entidad12, total_entidad13, total_entidad14, imp_deter,"
    StrSql = StrSql & " eme_medicas, seguro_optativo, seguro_retiro, tope_os_priv, prorratea, suscribe, caracter, fecha_caracter,"
    StrSql = StrSql & " fecha_devolucion, dependencia_dgi, anual_final, ternro, terape, ternom)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & pliqnro
    StrSql = StrSql & " ," & pronro
    StrSql = StrSql & " ," & empresa
    StrSql = StrSql & " ," & bpronroAux
    StrSql = StrSql & " ," & cambiaFecha(Fecha)
    StrSql = StrSql & " ," & CtrlNuloTXT(Hora)
    StrSql = StrSql & " ," & CtrlNuloTXT(iduser)
    StrSql = StrSql & " ," & empleg
    StrSql = StrSql & " ," & msr
    StrSql = StrSql & " ," & nomsr
    StrSql = StrSql & " ," & nogan
    StrSql = StrSql & " ," & jubilacion
    StrSql = StrSql & " ," & osocial
    StrSql = StrSql & " ," & cuota_medico
    StrSql = StrSql & " ," & prima_seguro
    StrSql = StrSql & " ," & sepelio
    StrSql = StrSql & " ," & estimados
    StrSql = StrSql & " ," & otras
    StrSql = StrSql & " ," & donacion
    StrSql = StrSql & " ," & dedesp
    StrSql = StrSql & " ," & noimpo
    StrSql = StrSql & " ," & car_flia
    StrSql = StrSql & " ," & conyuge
    StrSql = StrSql & " ," & hijo
    StrSql = StrSql & " ," & otras_cargas
    StrSql = StrSql & " ," & retenciones
    StrSql = StrSql & " ," & promo
    StrSql = StrSql & " ," & saldo
    StrSql = StrSql & " ," & cambiaFecha(Desde)
    StrSql = StrSql & " ," & cambiaFecha(Hasta)
    StrSql = StrSql & " ," & CtrlNuloTXT(Cuil)
    StrSql = StrSql & " ," & CtrlNuloTXT(direccion)
    StrSql = StrSql & " ," & sindicato
    StrSql = StrSql & " ," & ret_mes
    StrSql = StrSql & " ," & CtrlNuloTXT(dir_calle)
    StrSql = StrSql & " ," & CtrlNuloTXT(dir_num)
    StrSql = StrSql & " ," & CtrlNuloTXT(dir_piso)
    StrSql = StrSql & " ," & CtrlNuloTXT(dir_dpto)
    StrSql = StrSql & " ," & CtrlNuloTXT(dir_localidad)
    StrSql = StrSql & " ," & CtrlNuloTXT(dir_pcia)
    StrSql = StrSql & " ," & CtrlNuloTXT(dir_cp)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit)
    StrSql = StrSql & " ," & CtrlNuloTXT(monto_letras)
    StrSql = StrSql & " ," & CtrlNuloTXT(emp_nombre)
    StrSql = StrSql & " ," & CtrlNuloTXT(emp_cuit)
    StrSql = StrSql & " ," & CtrlNuloTXT(ano)
    StrSql = StrSql & " ," & mon_conyuge
    StrSql = StrSql & " ," & mon_hijo
    StrSql = StrSql & " ," & mon_otras
    StrSql = StrSql & " ," & viaticos
    StrSql = StrSql & " ," & amortizacion
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad1)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad2)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad3)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad4)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad5)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad6)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad7)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad8)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad9)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad10)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad11)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad12)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad13)
    StrSql = StrSql & " ," & CtrlNuloTXT(entidad14)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad1)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad2)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad3)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad4)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad5)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad6)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad7)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad8)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad9)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad10)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad11)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad12)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad13)
    StrSql = StrSql & " ," & CtrlNuloTXT(cuit_entidad14)
    StrSql = StrSql & " ," & monto_entidad1
    StrSql = StrSql & " ," & monto_entidad2
    StrSql = StrSql & " ," & monto_entidad3
    StrSql = StrSql & " ," & monto_entidad4
    StrSql = StrSql & " ," & monto_entidad5
    StrSql = StrSql & " ," & monto_entidad6
    StrSql = StrSql & " ," & monto_entidad7
    StrSql = StrSql & " ," & monto_entidad8
    StrSql = StrSql & " ," & monto_entidad9
    StrSql = StrSql & " ," & monto_entidad10
    StrSql = StrSql & " ," & monto_entidad11
    StrSql = StrSql & " ," & monto_entidad12
    StrSql = StrSql & " ," & monto_entidad13
    StrSql = StrSql & " ," & monto_entidad14
    StrSql = StrSql & " ," & ganimpo
    StrSql = StrSql & " ," & ganneta
    StrSql = StrSql & " ," & total_entidad1
    StrSql = StrSql & " ," & total_entidad2
    StrSql = StrSql & " ," & total_entidad3
    StrSql = StrSql & " ," & total_entidad4
    StrSql = StrSql & " ," & total_entidad5
    StrSql = StrSql & " ," & total_entidad6
    StrSql = StrSql & " ," & total_entidad7
    StrSql = StrSql & " ," & total_entidad8
    StrSql = StrSql & " ," & total_entidad9
    StrSql = StrSql & " ," & total_entidad10
    StrSql = StrSql & " ," & total_entidad11
    StrSql = StrSql & " ," & total_entidad12
    StrSql = StrSql & " ," & total_entidad13
    StrSql = StrSql & " ," & total_entidad14
    StrSql = StrSql & " ," & imp_deter
    StrSql = StrSql & " ," & eme_medicas
    StrSql = StrSql & " ," & seguro_optativo
    StrSql = StrSql & " ," & seguro_retiro
    StrSql = StrSql & " ," & tope_os_priv
    StrSql = StrSql & " ," & prorratea
    StrSql = StrSql & " ," & CtrlNuloTXT(suscribe)
    StrSql = StrSql & " ," & CtrlNuloTXT(caracter)
    StrSql = StrSql & " ," & cambiaFecha(fecha_caracter)
    StrSql = StrSql & " ," & cambiaFecha(fecha_devolucion)
    StrSql = StrSql & " ," & CtrlNuloTXT(dependencia_dgi)
    StrSql = StrSql & " ," & anual_final
    StrSql = StrSql & " ," & ternro
    StrSql = StrSql & " ," & CtrlNuloTXT(terape)
    StrSql = StrSql & " ," & CtrlNuloTXT(ternom) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procF649:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procF649"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Function nombre(ByVal tipo As Long) As String
Dim Salida As String
    
    Salida = ""
    Select Case tipo
        Case 1:
            Salida = "Empleados"
        Case 2:
            Salida = "TipoConc"
        Case 3:
            Salida = "Conceptos"
        Case 4:
            Salida = "Tipo Acumuladores"
        Case 5:
            Salida = "Acumuladores"
        Case 6:
            Salida = "Items"
        Case 7:
            Salida = "Tipo Procesos"
        Case 8:
            Salida = "Procesos"
        Case 9:
            Salida = "Periodos"
        Case 10:
            Salida = "CabLiq"
        Case 11:
            Salida = "Detliq"
        Case 12:
            Salida = "Acu_liq"
        Case 13:
            Salida = "Acu_mes"
        Case 14:
            Salida = "Desliq"
        Case 15:
            Salida = "Ficharet"
        Case 16:
            Salida = "Impproarg"
        Case 17:
            Salida = "Impmesarg"
        Case 18:
            Salida = "Traza_gan"
        Case 19:
            Salida = "Traza_gan_item_top"
        Case 20:
            Salida = "Desmen"
        Case 21:
            Salida = "Batch_proceso"
        Case 22:
            Salida = "Rep_recibo"
        Case 23:
            Salida = "Rep_recibo_det"
        Case 24:
            Salida = "Rep_libroley"
        Case 25:
            Salida = "Rep_libroley_det"
        Case 26:
            Salida = "Rep_libroley_fam"
        Case 27:
            Salida = "F649"
    End Select
    
    nombre = Salida

End Function


Public Sub procEscala_ded(ByVal Linea As String)

Dim arrLinea
Dim esd_topeinf As String
Dim esd_topesup As String
Dim esd_porcentaje As String
    
On Error GoTo E_procEscala_ded

    'Formato tipo, esd_topeinf, esd_topesup, esd_porcentaje
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 3 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    esd_topeinf = arrLinea(1)
    esd_topesup = arrLinea(2)
    esd_porcentaje = arrLinea(3)
    
    'Inserto Datos
    StrSql = "INSERT INTO escala_ded("
    StrSql = StrSql & " esd_topeinf, esd_topesup, esd_porcentaje)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & esd_topeinf
    StrSql = StrSql & " ," & esd_topesup
    StrSql = StrSql & " ," & esd_porcentaje & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procEscala_ded:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procEscala_ded"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procConfrep(ByVal Linea As String)

Dim arrLinea
Dim repnro As String
Dim confnrocol As String
Dim confetiq As String
Dim conftipo As String
Dim confval As String
Dim Empnro As String
Dim confaccion As String
Dim confval2 As String
    
On Error GoTo E_procConfrep

    'Formato tipo, repnro, confnrocol, confetiq, conftipo, confval, empnro, confaccion, confval2
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 8 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    repnro = arrLinea(1)
    confnrocol = arrLinea(2)
    confetiq = arrLinea(3)
    conftipo = arrLinea(4)
    confval = arrLinea(5)
    Empnro = arrLinea(6)
    confaccion = arrLinea(7)
    confval2 = arrLinea(8)
    
    'Inserto Datos
    StrSql = StrSql & " INSERT INTO confrep("
    StrSql = StrSql & " repnro, confnrocol, confetiq, conftipo, confval, empnro, confaccion, confval2)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & repnro
    StrSql = StrSql & " ," & confnrocol
    StrSql = StrSql & " ," & CtrlNuloTXT(confetiq)
    StrSql = StrSql & " ," & CtrlNuloTXT(conftipo)
    StrSql = StrSql & " ," & confval
    StrSql = StrSql & " ," & Empnro
    StrSql = StrSql & " ," & CtrlNuloTXT(confaccion)
    StrSql = StrSql & " ," & CtrlNuloTXT(confval2) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procConfrep:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procConfrep"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub


Public Sub procEscala(ByVal Linea As String)

Dim arrLinea
Dim escnro As String
Dim escinf As String
Dim escsup As String
Dim escporexe As String
Dim esccuota As String
Dim escano As String
Dim escmes As String
Dim esfecha As String
    
On Error GoTo E_procEscala

    'Formato tipo, escnro, escinf, escsup, escporexe, esccuota, escano, escmes, esfecha
    arrLinea = Split(Linea, Sep)
    
    'Control de parametro
    If UBound(arrLinea) <> 8 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La linea no coincide con el formato esperado."
        Flog.writeline Espacios(Tabulador * 1) & Linea
        CantRegErr = CantRegErr + 1
        HuboError = True
        Exit Sub
    End If
    
    'Recupero Datos
    escnro = arrLinea(1)
    escinf = arrLinea(2)
    escsup = arrLinea(3)
    escporexe = arrLinea(4)
    esccuota = arrLinea(5)
    escano = arrLinea(6)
    escmes = arrLinea(7)
    esfecha = arrLinea(8)
    
    'Inserto Datos
    StrSql = "SET IDENTITY_INSERT escala ON"
    StrSql = StrSql & " INSERT INTO escala("
    StrSql = StrSql & " escnro, escinf, escsup, escporexe, esccuota, escano, escmes, esfecha)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & " (" & escnro
    StrSql = StrSql & " ," & escinf
    StrSql = StrSql & " ," & escsup
    StrSql = StrSql & " ," & escporexe
    StrSql = StrSql & " ," & esccuota
    StrSql = StrSql & " ," & escano
    StrSql = StrSql & " ," & escmes
    StrSql = StrSql & " ," & cambiaFecha(esfecha) & ")"
    StrSql = StrSql & " SET IDENTITY_INSERT escala OFF"
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_procEscala:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: procEscala"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "Linea Procesada: " & Linea
    Flog.writeline "=================================================================="
    HuboError = True
End Sub



