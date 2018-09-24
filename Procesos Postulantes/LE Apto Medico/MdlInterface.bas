Attribute VB_Name = "MdlInterface"
Option Explicit

Global Const Version = "1.00" 'FGZ
Global Const FechaModificacion = "17/12/2009"
Global Const UltimaModificacion = " "

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

Global Const EncryptionKey = "56238"

Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegWarnings As Long
Global RegFecha As Date
Global NroProceso As Long

Global f
'Global HuboError As Boolean
Global Path
Global NArchivo
Global NroLinea As Long
Global LineaCarga As Long

Global separador As String
Global SeparadorDecimal As String
Global UsaEncabezado As Boolean

Global ErroresNov As Boolean

Global ErrCarga
Global LineaError
Global LineaOK

Global PisaNovedad As Boolean
Global Vigencia As Boolean
Global Vigencia_Desde As String
Global Vigencia_Hasta As String
Global Pisa As Boolean
Global TikPedNro As Long
Global nombrearchivo As String
Global acuNro As Long 'se usa en el modelo 216 de Citrusvil y se carga por confrep
Global nro_ModOrg  As Long

Global NroModelo As Long
Global DescripcionModelo As String
Global Primera_Vez As Boolean
Global Banco As Long
Global usuario As String
Global EncontroAlguno As Boolean

Global nrocolumna As Long
Global Tabs As Long

Global Pliqnro As Long
Global Obligatorio As Boolean
Global aux As String
Global Campoetiqueta As String
Global InsertarEnLote As Boolean


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de Interface.
' Autor      : Lisandro Moro
' Fecha      : 24/09/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim Nombre_Arch_Errores As String
Dim Nombre_Arch_Correctos As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
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

    'Obtiene los datos de como esta configurado el servidor actualmente
    Call ObtenerConfiguracionRegional
    
    
    Nombre_Arch = PathFLog & "AptoMedico_PostulantesLE" & "-" & NroProcesoBatch & ".log"
    Nombre_Arch_Errores = PathFLog & "Lineas_Errores" & "-" & NroProcesoBatch & ".log"
    Nombre_Arch_Correctos = PathFLog & "Lineas_Procesadas" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    Set FlogE = fs.CreateTextFile(Nombre_Arch_Errores, True)
    Set FlogP = fs.CreateTextFile(Nombre_Arch_Correctos, True)
    
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
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Numero, separador decimal    : " & NumeroSeparadorDecimal
    Flog.writeline "Numero, separador de miles   : " & NumeroSeparadorMiles
    Flog.writeline "Moneda, separador decimal    : " & MonedaSeparadorDecimal
    Flog.writeline "Moneda, separador de miles   : " & MonedaSeparadorMiles
    Flog.writeline "Formato de Fecha del Servidor: " & FormatoDeFechaCorto
    Flog.writeline "-----------------------------------------------------------------"
    
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 259 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    ErroresNov = False
    Primera_Vez = False
    tplaorden = 0
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        Flog.writeline Espacios(Tabulador * 0) & "Parametros del proceso = " & bprcparam
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call LevantarParamteros(bprcparam)
        LineaCarga = 0
        Call ComenzarTransferencia
    End If
    
    
Final:
    If Not HuboError Then
        If ErroresNov Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        End If
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Resumen
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "===================================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Lineas Leidas    : " & RegLeidos
    Flog.writeline Espacios(Tabulador * 0) & "Lineas Erroneas  : " & RegError
    Flog.writeline Espacios(Tabulador * 0) & "Warnings         : " & RegWarnings
    Flog.writeline Espacios(Tabulador * 0) & "Lineas Procesadas: " & RegLeidos - RegError
    Flog.writeline Espacios(Tabulador * 0) & "===================================================================="
    objConn.Close
    objconnProgreso.Close
    Flog.Close
    FlogE.Close
    FlogP.Close
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
    GoTo Final
End Sub


Private Sub LeeArchivo(ByVal nombrearchivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim carpeta
Dim Programado As Boolean

    If App.PrevInstance Then
        Flog.writeline Espacios(Tabulador * 0) & "Hay una instancia previa del proceso corriendo "
        Exit Sub
    End If
    'Espero hasta que se crea el archivo
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(nombrearchivo)
        If f.Size = 0 Then
            Flog.writeline Espacios(Tabulador * 0) & "No anda el getfile "
            Err.Number = 1
        End If
    Loop
    On Error GoTo 0
    Flog.writeline Espacios(Tabulador * 0) & "Archivo creado " & nombrearchivo
   
   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(nombrearchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProcesoBatch & "," & NroModelo & ",'" & Left(nombrearchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "inter_pin")
        Flog.writeline Espacios(Tabulador * 0) & "Ultimo inter_pin " & crpNro
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se pudo abrir el archivo " & nombrearchivo
    End If
                
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If rs_Modelo.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "No esta el modelo " & NroModelo
        Exit Sub
    End If
                    
    StrSql = "SELECT * FROM modelo_filas WHERE bpronro =" & NroProcesoBatch
    StrSql = StrSql & " ORDER BY fila "
    OpenRecordset StrSql, rs_Lineas
    If Not rs_Lineas.EOF Then
        rs_Lineas.MoveFirst
        Programado = False
    Else
        'Si no hay filas es porque se ejecutó programado
        Flog.writeline Espacios(Tabulador * 0) & "No hay filas seleccionadas"
        Programado = True
    End If
    
    'Determino la proporcion de progreso
    Progreso = 0
    If Not Programado Then
        CEmpleadosAProc = rs_Lineas.RecordCount
        If CEmpleadosAProc = 0 Then
            CEmpleadosAProc = 1
        End If
        IncPorc = (99 / CEmpleadosAProc)
    Else
        IncPorc = 1
    End If
    
    
    If Not Programado Then
        Do While Not f.AtEndOfStream And Not rs_Lineas.EOF
            strLinea = f.ReadLine
            NroLinea = NroLinea + 1
            If NroLinea = 1 And UsaEncabezado Then
                strLinea = f.ReadLine
            End If
            If Trim(strLinea) <> "" And NroLinea = rs_Lineas!fila Then
                
                Call Insertar_Linea_Segun_Modelo_Custom(strLinea)
                RegLeidos = RegLeidos + 1
                
                rs_Lineas.MoveNext
                
                'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
                'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
                'como colgado
                Progreso = Progreso + IncPorc
                Flog.writeline Espacios(Tabulador * 0) & "Progreso = " & CLng(Progreso) & " (Incremento = " & IncPorc & ")"
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 0) & "Progreso actualizado"
            End If
        Loop
    Else
        Do While Not f.AtEndOfStream
            strLinea = f.ReadLine
            NroLinea = NroLinea + 1
            If NroLinea = 1 And UsaEncabezado Then
                strLinea = f.ReadLine
            End If
            If Trim(strLinea) <> "" Then
                Call Insertar_Linea_Segun_Modelo_Custom(strLinea)
                
                'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
                'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
                'como colgado
                Progreso = Progreso + IncPorc
                Flog.writeline Espacios(Tabulador * 0) & "Progreso = " & CLng(Progreso) & " (Incremento = " & IncPorc & ")"
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 0) & "Progreso actualizado"
            End If
        Loop
    End If
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    f.Close
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Archivo procesado: " & nombrearchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Set f = fs.getfile(nombrearchivo)
    Archivo_Aux = Replace(Format(Now, "yyyy-mm-dd hh:mm:ss"), ":", "-") & " " & NArchivo
    
    On Error Resume Next
    f.Move Path & "\bk\" & Mid(Archivo_Aux, 1, Len(Archivo_Aux) - 3) & "bk"
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "La carpeta Destino no existe. Se creará."
        Set carpeta = fs.CreateFolder(Path & "\bk")
        f.Move Path & "\bk\" & Mid(Archivo_Aux, 1, Len(Archivo_Aux) - 3) & "bk"
    End If
    'desactivo el manejador de errores
    On Error GoTo 0
    
    Flog.writeline "archivo movido " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'Borrar el archivo
    'fs.Deletefile nombrearchivo, True
    
    
    
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    Exit Sub
    
CE:
    HuboError = True
    
    MyRollbackTrans
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Long
Dim pos2 As Long

Dim NombreArchivo1 As String
Dim NombreArchivo2 As String
Dim NombreArchivo3 As String


separador = "@"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Nro de Modelo
        pos1 = 1
        pos2 = InStr(pos1, parametros, separador) - 1
        NroModelo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        'Nombre del archivo a levantar
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, separador) - 1
        If pos2 > 0 Then
            nombrearchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        Else
            pos2 = Len(parametros)
            nombrearchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        End If
    End If
End If

End Sub


Public Sub ComenzarTransferencia()
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder

    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el registro de la tabla sistema nro 1"
        Exit Sub
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
        separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
        SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
        UsaEncabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
        DescripcionModelo = objRs!moddesc
        
        Flog.writeline Espacios(Tabulador * 1) & "Modelo " & NroModelo & " " & objRs!moddesc
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Directorio de importación :  " & Directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo
        Exit Sub
    End If
    
    'Algunos modelos no se comportan de la misma manera ==>
    Select Case NroModelo
    Case Else
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        Path = Directorio
        
        Dim fc, F1, s2
        Set Folder = fs.GetFolder(Directorio)
        Set CArchivos = Folder.Files
        
        HuboError = False
        EncontroAlguno = False
        For Each archivo In CArchivos
            EncontroAlguno = True
            If UCase(archivo.Name) = UCase(nombrearchivo) Or Trim(nombrearchivo) = "" Then
                NArchivo = archivo.Name
                Flog.writeline Espacios(Tabulador * 1) & "Procesando archivo " & archivo.Name
                Call LeeArchivo(Directorio & "\" & archivo.Name)
            End If
        Next
        If Not EncontroAlguno Then
            Flog.writeline Espacios(Tabulador * 1) & "No se encontró el archivo " & nombrearchivo
        End If
    End Select
End Sub

Public Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    
    ErroresNov = True
End Sub



Public Sub Escribir_Log(ByVal TipoLog As String, ByVal Lin As Long, ByVal Col As Long, ByVal msg As String, ByVal CantTab As Long, ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Escribe un mensage determinado en uno de 3 archivos de log
' Autor      : FGZ
' Fecha      : 18/04/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Select Case UCase(TipoLog)
    Case "FLOG" 'Archivo de Informacion de resumen
            Flog.writeline Espacios(Tabulador * CantTab) & msg
    Case "FLOGE" 'Archivo de Errores
            FlogE.writeline Espacios(Tabulador * CantTab) & "Linea " & Lin & " Columna " & Col & ": " & msg
            FlogE.writeline Espacios(Tabulador * CantTab) & strLinea
    Case "FLOGP" 'Archivo de lineas procesadas
            FlogP.writeline Espacios(Tabulador * CantTab) & "Linea " & Lin & " Columna " & Col & ": " & msg
    Case Else
        Flog.writeline Espacios(Tabulador * CantTab) & "Nombre de archivo de log incorrecto " & TipoLog
End Select

End Sub

Public Sub Insertar_Linea_Segun_Modelo_Custom(ByVal linea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador de acurdo al modelo
' Autor      : Lisandro Moro
' Fecha      : 24/09/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
' Interfaces Customisadas
Dim OK As Boolean

MyBeginTrans
    HuboError = False
    Select Case NroModelo
        Case 320: '
            Call LineaModelo_320(linea, OK)
        Case Else
            Texto = ": " & " No existe el modelo "
            Call Escribir_Log("floge", "", 0, Texto, Tabs, linea)
        End Select
    
MyCommitTrans
If Not HuboError Then
    'Flog.Writeline Espacios(Tabulador * 1) & "Transaccion Cometida"
Else
    Flog.writeline Espacios(Tabulador * 1) & "Transaccion Abortada"
End If

End Sub

Public Sub LineaModelo_320(ByVal strReg As String, ByRef OK As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Apto Medico de Postulantes de Legajo Electronico
' Autor      : FGZ
' Fecha      : 17/12/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
'   Formato:
'       Cuil
' ---------------------------------------------------------------------------------------------
Dim datos() As String
Dim Ternro As Long
Dim Ternro_Temp As Long
Dim a As Long
Dim tempString As String
Dim I As Long
Dim Actualizar_Postulante As Boolean
Dim Envia_Mail As Boolean
Dim TerMail As String
Dim NroLote As Long
Dim Fecha As Date
Dim MsgTxt As String

Dim rs_Ter As New ADODB.Recordset
Dim rs_Con As New ADODB.Recordset
Dim rs_Lote As New ADODB.Recordset
Dim rs_Lote_Post As New ADODB.Recordset

    Flog.writeline
    FlogE.writeline
    FlogP.writeline
    Flog.writeline "Procesando linea " & strReg
    
    'Si ocurre un error antes de insertar el tercero se aborta el postulante
    On Error GoTo Manejador_De_Error:
    
    Envia_Mail = False
    datos = Split("10" & separador & strReg, separador)
    For I = 0 To UBound(datos)
        datos(I) = Trim(datos(I))
    Next I
    
    'Cuil
    'datos(0) = 10
    datos(1) = StrToStr(CStr(datos(1)), 30) 'Numero de CUIL
    datos(1) = Replace(datos(1), ".", "") 'elimino puntos y comas
    datos(1) = Replace(datos(1), ",", "")
    
   
    '===========================================================
    'Bsucar el postulante en el Lote activo y poner su estado en apto
    'enviar el mail
    
    'Validaciones
    '       Que exista el postulante
    '       que el postulante esté en un lote activo
    
    
    '---------------------------------------------------------------------
    'Establecer la conexion a la BD temporal
    StrSql = " SELECT cnnro, cnstring FROM conexion WHERE cnnro = 2 "
    OpenRecordset StrSql, rs_Con
    If rs_Con.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "No se encuentra la conexion a la BD temporal."
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Abortado."
        Exit Sub
    End If
    
    On Error Resume Next
    'Abro la conexion a la BD Temporal
    OpenConnection rs_Con!cnstring, objconn2
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion. Debe Configurar bien la conexion a la BD temporal."
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Abortado."
        Exit Sub
    End If
    '---------------------------------------------------------------------
    
    On Error GoTo Manejador_De_Error:
        
    'Reviso si el postulante ya existe en la BD temporal
    StrSql = "SELECT tercero.ternro, tercero.teremail FROM ter_doc  "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = ter_doc.ternro "
    StrSql = StrSql & " WHERE ter_doc.tidnro = " & datos(0)
    StrSql = StrSql & " AND nrodoc = '" & datos(1) & "'"
    OpenRecordsetWithConn StrSql, rs_Ter, objconn2
    If rs_Ter.EOF Then
        'El postulante no existe en la BD temporal
        Flog.writeline Espacios(Tabulador * 0) & "No se encuentra un Postulante con ese CUIL " & datos(1)
        Exit Sub
    Else
        Ternro_Temp = rs_Ter!Ternro
        TerMail = rs_Ter!teremail
    End If
        
        
    'Busco el postulante en un lote activo
    StrSql = "SELECT pos_busqueda.busnro,busfin,busfecfin, pos_busqueda_comp.actnro, pos_busqueda_comp.estnro "
    StrSql = StrSql & " FROM pos_busqueda "
    StrSql = StrSql & " LEFT JOIN pos_busqueda_comp ON pos_busqueda_comp.busnro = pos_busqueda.busnro "
    StrSql = StrSql & " WHERE busfin = 0"
    StrSql = StrSql & " AND ternro = " & Ternro_Temp
    OpenRecordsetWithConn StrSql, rs_Lote, objconn2
    If Not rs_Lote.EOF Then
        NroLote = rs_Lote!busnro
        Fecha = rs_Lote!busfecfin
        
        'No se debe reenviar actualizar el estado y/o reenviar el mail si el estado es
        '               2 Completo
        '               6 Ausente
        '               7 Firmado
        
        'Estados
            ' 1 Inicial
            ' 2 Completado
            ' 3 No le interesa
            ' 4 Apto
            ' 5 No Apto
            ' 6 Ausente
            ' 7 Firmado
        
        Select Case rs_Lote!estnro
        Case 1:
            Actualizar_Postulante = True
            Envia_Mail = True
        Case 2, 3, 6, 7:
            Actualizar_Postulante = False
            Envia_Mail = False
        Case 4:
            Actualizar_Postulante = False
            Envia_Mail = True
        Case 5:
            Actualizar_Postulante = True
            Envia_Mail = True
        End Select
    Else
        Flog.writeline "El postulante no se encuentra en ningun lote activo."
        Actualizar_Postulante = False
    End If
        
        
    'Asocio el postulante al lote
    If Actualizar_Postulante Then
        'tabla pos_busqueda_comp
        StrSql = "UPDATE pos_busqueda_comp SET "
        StrSql = StrSql & " estnro = 4 "
        StrSql = StrSql & " WHERE ternro = " & Ternro_Temp
        StrSql = StrSql & " AND busnro = " & NroLote
        objconn2.Execute StrSql, , adExecuteNoRecords
        
        Envia_Mail = True
    End If


    If Envia_Mail Then
        Flog.writeline
        Flog.writeline "Envia mail al Postulante"
        Flog.writeline
        
        MsgTxt = TextoMail(Ternro_Temp, Fecha)
        
        Call EnviarMail(Ternro_Temp, TerMail, MsgTxt)
    End If
    
    
Fin:
Exit Sub

Manejador_De_Error:
    HuboError = True
    Flog.writeline "SQL ejecutada: " & StrSql
    Texto = ": " & Err.Description
    Call Escribir_Log("floge", NroLinea, nrocolumna, Texto, Tabs, strReg)
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strReg
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Public Sub ValidarPais(Pais As String, ByRef Nro_Pais As Long)
    Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Pais) Then
        StrSql = " SELECT * FROM pais WHERE UPPER(paisdesc) = '" & UCase(Pais) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO pais(paisdesc,paisdef) VALUES('"
            StrSql = StrSql & UCase(Pais) & "',0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Nro_Pais = getLastIdentity(objConn, "pais")
        Else
            Nro_Pais = rs_sub!Paisnro
        End If
    End If
End Sub

Public Sub ValidarProvincia(Provincia As String, ByRef Nro_Provincia As Long, Nro_Pais As Long)
    Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Provincia) Then
        StrSql = " SELECT * FROM provincia WHERE upper(provdesc) = '" & UCase(Provincia) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO provincia(provdesc,paisnro) VALUES('"
            StrSql = StrSql & UCase(Provincia) & "'," & Nro_Pais & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Nro_Provincia = getLastIdentity(objConn, "provincia")
        Else
            Nro_Provincia = rs_sub!provnro
        End If
    End If
End Sub

Public Sub ValidarLocalidad(Localidad As String, ByRef Nro_Localidad As Long, Nro_Pais As Long, Nro_Provincia As Long, Optional Nro_Zona As Long)
    Dim rs_sub As New ADODB.Recordset
    Dim Sql_Ins As String
    Dim SQL_Val As String
    
    If Not EsNulo(Localidad) Then
        StrSql = " SELECT * FROM localidad WHERE UPPER(locdesc) = '" & UCase(Localidad) & "'"
    '    If nro_pais <> 0 Then    '        StrSql = StrSql & " AND paisnro = " & nro_pais    '    End If    '    '    If nro_provincia <> 0 Then    '        StrSql = StrSql & " AND provnro = " & nro_provincia    '    End If
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            Sql_Ins = " INSERT INTO localidad(locdesc"
            SQL_Val = " VALUES('" & UCase(Localidad) & "'"
            If Nro_Pais <> 0 Then
                Sql_Ins = Sql_Ins & ",paisnro"
                SQL_Val = SQL_Val & "," & Nro_Pais
            End If
            If Nro_Zona <> 0 Then
                Sql_Ins = Sql_Ins & ",zonanro"
                SQL_Val = SQL_Val & "," & Nro_Zona
            End If
            If Nro_Provincia <> 0 Then
                Sql_Ins = Sql_Ins & ",provnro"
                SQL_Val = SQL_Val & "," & Nro_Provincia
            End If
            StrSql = Sql_Ins & ")" & SQL_Val & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Nro_Localidad = getLastIdentity(objConn, "localidad")
        Else
            Nro_Localidad = rs_sub!locnro
        End If
    End If
End Sub

Public Sub ValidarPartido(Partido As String, ByRef Nro_Partido As Long)
    Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Partido) Then
        StrSql = " SELECT * FROM partido WHERE UPPER(partnom) = '" & UCase(Partido) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO partido(partnom) VALUES('"
            StrSql = StrSql & UCase(Partido) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            Nro_Partido = getLastIdentity(objConn, "partido")
        Else
            Nro_Partido = rs_sub!partnro
        End If
    End If
End Sub

Public Sub ValidarZona(Zona As String, ByRef Nro_Zona As Long, Nro_Provincia As Long)
    Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Zona) Then
        StrSql = " SELECT * FROM zona WHERE UPPER(zonadesc) = '" & UCase(Zona) & "' AND provnro = " & Nro_Provincia
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO zona(zonadesc,provnro) VALUES('"
            StrSql = StrSql & UCase(Zona) & "'," & Nro_Provincia & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Nro_Zona = getLastIdentity(objConn, "zona")
        Else
            Nro_Zona = rs_sub!zonanro
        End If
    End If
End Sub

Public Function ValidaEstructura(tipoEstr As Long, ByRef estructura As String) As Long
    Dim Rs_Estr As New ADODB.Recordset
    
    StrSql = " SELECT estrnro FROM estructura WHERE UPPER(estructura.estrdabr) = '" & UCase(Mid(estructura, 1, 60)) & "'"
    StrSql = StrSql & " AND estructura.tenro = " & tipoEstr
    OpenRecordset StrSql, Rs_Estr
    If Not Rs_Estr.EOF Then
        'CodEst = Rs_Estr!estrNro
        ValidaEstructura = Rs_Estr!estrNro
    Else
        StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest)"
        StrSql = StrSql & " VALUES(" & tipoEstr & ",'" & UCase(Mid(estructura, 1, 60)) & "',1,-1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        'CodEst = getLastIdentity(objConn, "estructura")
        ValidaEstructura = getLastIdentity(objConn, "estructura")
    End If
End Function


Sub EnviarMail(ByVal Ternro As Long, ByVal Mail As String, ByVal Texto As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Envia mail a postulantes de Legajo Electronico
' Autor      : FGZ
' Fecha      : 17/12/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim bpronro As Long

    On Error GoTo CE
    
    MyBeginTrans
    
    StrSql = "insert into batch_proceso "
    StrSql = StrSql & "(btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
    StrSql = StrSql & "bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados) "
    StrSql = StrSql & "values (25," & ConvFecha(Date) & ",'" & usuario & "','" & FormatDateTime(Time, 4) & ":00'"
    StrSql = StrSql & ",null,null,'1','Preparando',null,null,null,null,0,null)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    bpronro = getLastIdentity(objConn, "batch_proceso")
    
    Call MailsAPostTLP(bpronro, Ternro, Mail, Texto)
    
    StrSql = "UPDATE batch_proceso SET bprcestado = 'Pendiente'"
    StrSql = StrSql & " WHERE bpronro = " & bpronro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    MyCommitTrans
    
    Exit Sub
CE:
    HuboError = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
    MyRollbackTrans
End Sub

Sub MailsAPostTLP(ByVal bpronro As Long, ByVal Ternro As Long, ByVal Mails As String, ByVal Texto As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Envia mail a postulantes de Legajo Electronico
' Autor      : FGZ
' Fecha      : 17/12/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim colcount As Integer
Dim Mensaje As String
Dim strSQLQuery As String
Dim param As String
Dim Opcion As Integer
Dim campo As Integer
Dim reemplazo As String
Dim AlertaFileName As String
Dim AlertaAttachFileName  As String
Dim fs2, AlertaFile
Dim Texto_Body As String
Dim dirsalidas As String

Dim objRs As New ADODB.Recordset

    On Error GoTo CE
            
            
    'Directorio Salidas
    StrSql = "select sis_dirsalidas from sistema"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        dirsalidas = objRs!sis_dirsalidas & "\attach"
    Else
        Flog.writeline "No se encuentra configurado sis_dirsalidas"
        Exit Sub
    End If
    If objRs.State = adStateOpen Then objRs.Close
            
    If Mails <> "" Then
        'Este es el mensaje en html completo. Es lo mismo que lo que se adjunta ------
        Texto_Body = Texto
        
        'Attach
        'Va a ser un archivo fijo que es un isntructivo de como completar los datos
        ' el mismo se encuentra en la carpeta attach/le/ y se llama instructivo.doc
        AlertaFileName = dirsalidas & "\msg_" & bpronro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now)
        AlertaAttachFileName = dirsalidas & "\le\instructivo.doc"
        
        'Tambien lo podria crear aca
            ''Creo el archivo a atachar en el mail
            'Set fs2 = CreateObject("Scripting.FileSystemObject")
            'Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
            'Flog.writeline "Fin Alerta: Positiva. Enviado a: " & Mails
            '
            'AlertaFile.writeline "<html><head>"
            'AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
            'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
            'AlertaFile.writeline "<title>Alerta - Busqueda &reg;</title></head><body>"
            'AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
            'AlertaFile.writeline "<table>" & Mensaje & "</table>"
            'AlertaFile.writeline "</body></html>"
            'AlertaFile.Close
        
        Call crearProcesosMensajeriaTLP(Mails, AlertaFileName, AlertaAttachFileName, Texto_Body)
    Else
        Flog.writeline "Alerta: No se encontraron mails definidos en el resultado de la consulta. No se han enviado mensajes."
    End If
    
    
    
    Exit Sub
CE:
    HuboError = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub



Sub crearProcesosMensajeriaTLP(ByVal mailBoxs As String, ByVal AlertaFileName As String, ByVal AlertaAttachFileName As String, ByVal Mensaje As String)
Dim objRs As New ADODB.Recordset
Dim fs2, MsgFile

    Set fs2 = CreateObject("Scripting.FileSystemObject")
    
    'Los nombres de los archivos para los mails de esta alerta empiezan con el bpronro de este proceso
    Set MsgFile = fs2.CreateTextFile(AlertaFileName & ".msg", True)
    
    MsgFile.writeline "[MailMessage]"
    MsgFile.writeline "FromName=Teleperformance. Dpto Recruiting."
    MsgFile.writeline "Subject=Alerta - Busqueda"
    
    MsgFile.writeline "Body1=" & Mensaje
    If Len(AlertaAttachFileName) > 0 Then
       MsgFile.writeline "Attachment=" & AlertaAttachFileName
    Else
       MsgFile.writeline "Attachment="
    End If
    
    MsgFile.writeline "Recipients=" & mailBoxs
    
    StrSql = "select cfgemailfrom,cfgemailhost,cfgemailport,cfgemailuser,cfgemailpassword from conf_email where cfgemailest = -1"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        MsgFile.writeline "FromAddress=" & objRs!cfgemailfrom
        MsgFile.writeline "Host=" & objRs!cfgemailhost
        MsgFile.writeline "Port=" & objRs!cfgemailport
        MsgFile.writeline "User=" & objRs!cfgemailuser
        MsgFile.writeline "Password=" & objRs!cfgemailpassword
    Else
        Flog.writeline "No existen datos configurados para el envio de emails, o no existe configuracion activa"
        Exit Sub
    End If
    If objRs.State = adStateOpen Then objRs.Close
End Sub


Private Function TextoMail(ByVal Ternro As Long, ByVal Fecha As Date) As String
Dim Texto_Body As String
Dim Link As String

Dim rs_Link As New ADODB.Recordset

    StrSql = " SELECT cnnro,cnstring FROM conexion "
    StrSql = StrSql & " WHERE cnnro = 3 "
    OpenRecordset StrSql, rs_Link
    If Not rs_Link.EOF Then
        Link = rs_Link!cnstring
    Else
        Flog.writeline "No se encuentra la conexion para asociar en el mail. Debe existir una conexion nro 3."
        Link = " LINK "
    End If


    Texto_Body = "¡ Felicitaciones Quedaste seleccionado para nuestra búsqueda ! " & Chr(13)
    Texto_Body = Texto_Body & "Para continuar con el proceso de incorporación te solicitamos ingresar a la siguiente dirección "
    Texto_Body = Texto_Body & Link & Encrypt(EncryptionKey, Ternro)
    Texto_Body = Texto_Body & Chr(13)
    Texto_Body = Texto_Body & " Es vital que completes esta información antes del " & Fecha
    Texto_Body = Texto_Body & " a las 24:00 hs, para poder contar con tus datos para el armado del contrato y la "
    Texto_Body = Texto_Body & "confirmación del día de la firma del mismo." & Chr(13)
    Texto_Body = Texto_Body & "En esta página encontraras información de nuestra Empresa y podrás completar tus datos personales." & Chr(13)
    Texto_Body = Texto_Body & "Te recomendamos que leas con atención ya que esta información que ingreses se utilizara para armar tu legajo."
    Texto_Body = Texto_Body & Chr(13)
    Texto_Body = Texto_Body & "Cualquier inconveniente que tengas, por favor comunicate al siguiente tel 5555-55555"
    Texto_Body = Texto_Body & Chr(13)
    Texto_Body = Texto_Body & Chr(13)
    Texto_Body = Texto_Body & Chr(13)
    Texto_Body = Texto_Body & "                                                          Departamente de Recruiting"

    
    If rs_Link.State = adStateOpen Then rs_Link.Close
    Set rs_Link = Nothing
    
    
    TextoMail = Texto_Body
End Function
