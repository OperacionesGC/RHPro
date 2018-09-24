Attribute VB_Name = "MdlInterface"
Option Explicit

Global Const Version = "1.00" 'FGZ
Global Const FechaModificacion = "17/12/2009"
Global Const UltimaModificacion = " "

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

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
Global Lote As String
Global NroLote As Long


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
    
    
    Nombre_Arch = PathFLog & "Importacion_PostulantesLE" & "-" & NroProcesoBatch & ".log"
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 258 AND bpronro =" & NroProcesoBatch
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
        
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, separador) - 1
            If pos2 > 0 Then
                Lote = Mid(parametros, pos1, pos2 - pos1 + 1)
            Else
                pos2 = Len(parametros)
                Lote = Mid(parametros, pos1, pos2 - pos1 + 1)
            End If
        Else
            pos2 = Len(parametros)
            nombrearchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
            
            Lote = "0"
        End If
        
        If EsNulo(Lote) Or Lote = "0" Then
            InsertarEnLote = False
            NroLote = 0
        Else
            InsertarEnLote = True
            NroLote = Lote
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
        Case 319: '
            Call LineaModelo_319(linea, OK)
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

Public Sub LineaModelo_319(ByVal strReg As String, ByRef OK As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Postulantes para Legajo Electronico
' Autor      : FGZ
' Fecha      : 14/12/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
    'Formato:
'        Apellidos
'        Nombres
'        Fecha de nacimiento (AAAA-MM-DD)
'        Numero de DNI
'        Direccion
'        Ciudad
'        Provincia
'        Codigo Postal
'        E -mail
'        Telefono
'        Cuil
'----------------------------
'1        Apellido 1
'2        Apellido 2
'3        Nombre 1
'4        Nombre 1
'5        Sexo
'6        Fecha de nacimiento (AAAA-MM-DD)
'7        Numero de DNI
'8        Direccion   Calle
'9        Direccion   Nro
'10       Direccion   Piso
'11       Direccion   Depto
'12       Direccion   Torre
'13       Direccion   Manzana
'14       Direccion   Sector
'15       Direccion   Ciudad
'16       Direccion   Provincia
'17       Direccion   Codigo Postal
'18       E-mail
'19       Telefono
'20       Cuil
' ---------------------------------------------------------------------------------------------
Dim datos() As String
Dim ternro As Long
Dim ternro_temp As Long
Dim NroDom As Long
Dim a As Long
Dim Xconst As Long
Dim ActPasos As Boolean
Dim tempString As String
Dim EstPosNro As Long
Dim I As Long
Dim Paisnro As Long
Dim Asociar_Postulante As Boolean
Dim Esta_En_BL As Boolean

Dim rs_Ter As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Con As New ADODB.Recordset
Dim rs_Lote As New ADODB.Recordset
Dim rs_Lote_Post As New ADODB.Recordset

    Flog.writeline
    FlogE.writeline
    FlogP.writeline
    Flog.writeline "Procesando linea " & strReg
    
    'Si ocurre un error antes de insertar el tercero se aborta el postulante
    On Error GoTo Manejador_De_Error:
    
    datos = Split("0" & separador & strReg, separador)
    For I = 0 To UBound(datos)
        datos(I) = Trim(datos(I))
    Next I
    
    'Apellidos
    datos(1) = StrToStr(datos(1), 25) 'terape
    datos(2) = StrToStr(datos(2), 25) 'terape2
    'Nombres
    datos(3) = StrToStr(datos(3), 25) 'ternom
    datos(4) = StrToStr(datos(4), 25)  'ternom2
    
    'Sexo
    If (UCase(datos(5)) = "M") Or (UCase(datos(5)) = "-1") Or (UCase(datos(5)) = "MASCULINO") Then
        datos(5) = -1
    Else
        datos(5) = 0
    End If
    
    'Fecha de nacimiento (AAAA-MM-DD)
    datos(6) = ConvFecha(CDate(datos(6))) 'terfecnac
    
    'Numero de DNI
    datos(0) = 1
    'datos(1) = TraerCodTipoDocumento(CStr(datos(1)))  'TipoDocumento
    datos(7) = StrToStr(CStr(datos(7)), 30) 'Numero de DNI
    datos(7) = Replace(datos(7), ".", "") 'elimino puntos y comas
    datos(7) = Replace(datos(7), ",", "")
    
    'Direccion
        datos(8) = StrToStr(datos(8), 30) 'Calle
        datos(9) = StrToStr(datos(9), 8) 'Numero
        datos(10) = StrToStr(datos(10), 8) 'Piso
        datos(11) = StrToStr(datos(11), 8) 'Depto
        datos(12) = StrToStr(datos(12), 8) 'Torre
        datos(13) = StrToStr(datos(13), 8) 'Manzana
        datos(14) = StrToStr(datos(14), 8) 'Sector
    'Ciudad
        datos(15) = TraerCodLocalidad(CStr(datos(15))) 'Localidad
    'Provincia
        datos(16) = TraerCodProvincia(CStr(datos(16)))   'Provincia
    'Codigo Postal
        datos(17) = StrToStr(datos(17), 12) 'CP
        
    'E -mail
        datos(18) = StrToStr(datos(18), 100) 'CP
        
    'Telefono
    datos(19) = validatelefono(StrToStr(datos(19), 20))
    
    'Cuil
    'TipoDoc = 10
    datos(20) = StrToStr(CStr(datos(20)), 30) 'Numero de DNI
    datos(20) = Replace(datos(20), ".", "") 'elimino puntos y comas
    datos(20) = Replace(datos(20), ",", "")
    
   
   
    'completo datos obligatorios no informadfos
    'datos(34) = TraerCodPartido(CStr(datos(34)))  'dirPartido
    Paisnro = TraerCodPais("Argentina")
    
   
   
    '===========================================================
    'Insertar en la BD local
    'replicar en la bd temporal
    'actualizar la bd local
    
    
    'Validaciones
    '       Black List
    '       las mismas que se hacen desde el alta manual de postulantes
    
    
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
    
        
    'Reviso si el postulante esta en la lista negra
    StrSql = "SELECT b_list.empleg,b_list.ternro,b_list.observacion, tercero.terape,tercero.ternom "
    StrSql = StrSql & " FROM b_list "
    StrSql = StrSql & " INNER JOIN tercero  ON tercero.ternro  = b_list.ternro "
    StrSql = StrSql & " INNER JOIN ter_doc dni  ON dni.ternro =  b_list.ternro "
    StrSql = StrSql & " WHERE dni.tidnro = " & datos(0)
    StrSql = StrSql & " AND dni.nrodoc = '" & datos(7) & "'"
    OpenRecordset StrSql, rs_Ter
    If rs_Ter.EOF Then
        Esta_En_BL = False
    Else
        Esta_En_BL = True
    End If
        
        
        
        
        
        
        
        
If Esta_En_BL Then
    Flog.writeline "El Postulante está en la lista negra. "
Else
    'Reviso si el postulante ya existe en la BD temporal
    StrSql = "SELECT * FROM ter_doc  "
    StrSql = StrSql & " WHERE ter_doc.tidnro = " & datos(0)
    StrSql = StrSql & " AND nrodoc = '" & datos(7) & "'"
    OpenRecordsetWithConn StrSql, rs_Ter, objconn2
    If rs_Ter.EOF Then
    
        'Busco el nro de tercero que va a tener en la BD temporal
         StrSql = "SELECT max(ternro) maximo FROM TERCERO "
         OpenRecordsetWithConn StrSql, rs_Ter, objconn2
         If Not rs_Ter.EOF Then
            ternro_temp = CLng(rs_Ter("maximo")) + 1
         End If
    
        StrSql = " INSERT INTO tercero (ternom, ternom2, terape, terape2, terfecnac, tersex, teremail) VALUES ("
        StrSql = StrSql & "'" & datos(3) & "'"      'ternom
        StrSql = StrSql & ",'" & datos(4) & "'"     'ternom2
        StrSql = StrSql & ",'" & datos(1) & "'"      'terape
        StrSql = StrSql & ",'" & datos(2) & "'"     'terape2
        StrSql = StrSql & "," & datos(6)            'terfecnac
        StrSql = StrSql & "," & CLng(datos(5))      'tersex
        StrSql = StrSql & ",'" & datos(18) & "'"    'teremail
        StrSql = StrSql & ")"
        objconn2.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Inserto en la tabla de tercero (BD temporal)"
        Flog.writeline StrSql
        
        '--Obtengo el ternro--
        'por alguna razon que no logro descubrir esto no funciona
        'ternro_temp = getLastIdentity(objconn2, "tercero")
        Flog.writeline "-----------------------------------------------"
        Flog.writeline "Codigo de Tercero = " & ternro_temp
        
        ' Inserto los datos como postulante
        StrSql = "INSERT INTO pos_postulante "
        StrSql = StrSql & "(ternro, telinterno, telcelular,tercerotemp) "
        StrSql = StrSql & " VALUES (" & ternro_temp & ",'" & datos(19) & "',''," & ternro_temp & ")"
        objconn2.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Complemento de postulante"
       
        '--Inserto el Registro correspondiente en ter_tip--
        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & ternro_temp & ",14)"
        objconn2.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Tipo de tercero"
        
        '-----------------------------------------------------------------------------
        'Cualquier error que ocurra de aquí en adelante sigue adelante
        ' tratando de insertar la mayor cantidad de datos posible
        On Error Resume Next
        
        'DNI
        StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
        StrSql = StrSql & " VALUES(" & ternro_temp & "," & datos(0) & ",'" & datos(7) & "')"
        objconn2.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "DNI"
        
        'CUIL
        StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
        StrSql = StrSql & " VALUES(" & ternro_temp & ",10,'" & datos(20) & "')"
        objconn2.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "CUIL"
    
        '--------------------------------------------------------------------------------
        '--Inserto el Domicilio--
        StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
        StrSql = StrSql & " VALUES(1," & ternro_temp & ",-1,2)"
        objconn2.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Domicilio"
    
        '--Obtengo el numero de domicilio en la tabla--
        NroDom = CLng(getLastIdentity(objconn2, "cabdom"))
    
        StrSql = " INSERT INTO detdom (domnro,calle,nro,sector,torre,piso,oficdepto,manzana,codigopostal,"
        StrSql = StrSql & "locnro,provnro,paisnro)" ', zonanro, partnro) "
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & NroDom
        StrSql = StrSql & ",'" & datos(8) & "'"
        StrSql = StrSql & ",'" & datos(9) & "'"
        StrSql = StrSql & ",'" & datos(14) & "'"
        StrSql = StrSql & ",'" & datos(12) & "'"
        StrSql = StrSql & ",'" & datos(10) & "'"
        StrSql = StrSql & ",'" & datos(11) & "'"
        StrSql = StrSql & ",'" & datos(13) & "'"
        StrSql = StrSql & ",'" & datos(17) & "'"
        StrSql = StrSql & "," & datos(15)
        StrSql = StrSql & "," & datos(16)
        StrSql = StrSql & "," & Paisnro
        'StrSql = StrSql & "," & datos(35)
        'StrSql = StrSql & "," & datos(34)
        StrSql = StrSql & ")"
        objconn2.Execute StrSql, , adExecuteNoRecords
        If Err Then
            Err.Clear
        End If
        Flog.writeline "Domicilio insertado."
    
        '--Telefonos-Personal--
        datos(19) = validatelefono(StrToStr(datos(19), 20))
        If datos(19) <> "" Then
            StrSql = " INSERT INTO telefono "
            StrSql = StrSql & " (domnro, telnro, telfax, teldefault, telcelular ) "
            StrSql = StrSql & " VALUES (" & NroDom & ", '" & datos(19) & "' ,0 , -1 ,0 ) "
            objconn2.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Err.Clear
            End If
            Flog.writeline "Telefono Personal Insertado."
        End If
        
'        '--Telefonos-Celular--
'        'FGZ 11/04/2005 - quieren que lo cargue como telefono comun
'        datos(19) = validatelefono(StrToStr(datos(19), 20))
'        If datos(19) <> "" Then
'            StrSql = " INSERT INTO telefono "
'            StrSql = StrSql & " (domnro, telnro, telfax, teldefault, telcelular ) "
'            StrSql = StrSql & " VALUES (" & NroDom & ", '" & datos(19) & "' , 0, -1, 0 ) "
'            objconn2.Execute StrSql, , adExecuteNoRecords
'            If Err Then
'                Err.Clear
'            End If
'            Flog.writeline "Telefono Celular Insertado."
'        End If
    Else
        'El postulante ya existe
        ternro_temp = rs_Ter!ternro
        'InsertarEnLote = False
    End If
        
        
    'Reviso si el postulante ya existe en la BD productiva
    StrSql = "SELECT * FROM ter_doc  "
    StrSql = StrSql & " WHERE ter_doc.tidnro = " & datos(0)
    StrSql = StrSql & " AND nrodoc = '" & datos(7) & "'"
    OpenRecordset StrSql, rs_Ter
    If rs_Ter.EOF Then
        StrSql = " INSERT INTO tercero (ternom, ternom2, terape, terape2, terfecnac, tersex, teremail) VALUES ("
        StrSql = StrSql & "'" & datos(3) & "'"      'ternom
        StrSql = StrSql & ",'" & datos(4) & "'"     'ternom2
        StrSql = StrSql & ",'" & datos(1) & "'"      'terape
        StrSql = StrSql & ",'" & datos(2) & "'"     'terape2
        StrSql = StrSql & "," & datos(6)            'terfecnac
        StrSql = StrSql & "," & CLng(datos(5))      'tersex
        StrSql = StrSql & ",'" & datos(18) & "'"    'teremail
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Tercero"
        Flog.writeline StrSql
    
        '--Obtengo el ternro--
        ternro = getLastIdentity(objConn, "tercero")
        Flog.writeline "-----------------------------------------------"
        Flog.writeline "Codigo de Tercero = " & ternro
        
        ' Inserto los datos como postulante
        StrSql = "INSERT INTO pos_postulante "
        StrSql = StrSql & "(ternro, telinterno, telcelular,tercerotemp) "
        StrSql = StrSql & " VALUES (" & ternro & ",'" & datos(19) & "',''," & ternro_temp & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Complemento de Postulantes"
        If Err Then
            Flog.writeline "Error"
            Flog.writeline StrSql
            Err.Clear
        End If
        
        
       
        '--Inserto el Registro correspondiente en ter_tip--
        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & ternro & ",14)"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Tipo de Tercero"
        If Err Then
            Flog.writeline "Error"
            Flog.writeline StrSql
            Err.Clear
        End If
        
        
        'DNI
        StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
        StrSql = StrSql & " VALUES(" & ternro & "," & datos(0) & ",'" & datos(7) & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "DNI"
        If Err Then
            Flog.writeline "Error"
            Flog.writeline StrSql
            Err.Clear
        End If
    
    
        'CUIL
        StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
        StrSql = StrSql & " VALUES(" & ternro & ",10,'" & datos(20) & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "CUIL"
        If Err Then
            Flog.writeline "Error"
            Flog.writeline StrSql
            Err.Clear
        End If
    
    
        '--------------------------------------------------------------------------------
        '--Inserto el Domicilio--
        StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
        StrSql = StrSql & " VALUES(1," & ternro & ",-1,2)"
        objConn.Execute StrSql, , adExecuteNoRecords
        If Err Then
            Flog.writeline "Error"
            Flog.writeline StrSql
            Err.Clear
        End If
    
        '--Obtengo el numero de domicilio en la tabla--
        NroDom = CLng(getLastIdentity(objConn, "cabdom"))
    
        StrSql = " INSERT INTO detdom (domnro,calle,nro,sector,torre,piso,oficdepto,manzana,codigopostal,"
        StrSql = StrSql & "locnro,provnro,paisnro)" ', zonanro, partnro) "
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & NroDom
        StrSql = StrSql & ",'" & datos(8) & "'"
        StrSql = StrSql & ",'" & datos(9) & "'"
        StrSql = StrSql & ",'" & datos(14) & "'"
        StrSql = StrSql & ",'" & datos(12) & "'"
        StrSql = StrSql & ",'" & datos(10) & "'"
        StrSql = StrSql & ",'" & datos(11) & "'"
        StrSql = StrSql & ",'" & datos(13) & "'"
        StrSql = StrSql & ",'" & datos(17) & "'"
        StrSql = StrSql & "," & datos(15)
        StrSql = StrSql & "," & datos(16)
        StrSql = StrSql & "," & Paisnro
        'StrSql = StrSql & "," & datos(35)
        'StrSql = StrSql & "," & datos(34)
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        If Err Then
            Err.Clear
        End If
        Flog.writeline "Domicilio insertado."
    
        '--Telefonos-Personal--
        datos(19) = validatelefono(StrToStr(datos(19), 20))
        If datos(19) <> "" Then
            StrSql = " INSERT INTO telefono "
            StrSql = StrSql & " (domnro, telnro, telfax, teldefault, telcelular ) "
            StrSql = StrSql & " VALUES (" & NroDom & ", '" & datos(19) & "' ,0 , -1 ,0 ) "
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Err.Clear
            End If
            Flog.writeline "Telefono Personal Insertado."
        End If
        
'        '--Telefonos-Celular--
'        'FGZ 11/04/2005 - quieren que lo cargue como telefono comun
'        datos(19) = validatelefono(StrToStr(datos(19), 20))
'        If datos(19) <> "" Then
'            StrSql = " INSERT INTO telefono "
'            StrSql = StrSql & " (domnro, telnro, telfax, teldefault, telcelular ) "
'            StrSql = StrSql & " VALUES (" & NroDom & ", '" & datos(19) & "' , 0, -1, 0 ) "
'            objConn.Execute StrSql, , adExecuteNoRecords
'            If Err Then
'                Err.Clear
'            End If
'            Flog.writeline "Telefono Celular Insertado."
'        End If
    Else
        'El postulante ya existe
        'InsertarEnLote = False
    End If
    
    
    '-------------------------------------------------------
    'Opcional. Asociar a un Lote
    '
    '   La idea es que si viene informado el lote ==> los postulantes importados se asocien a un lote en particular
    '       y sino simplemente se carguen los postulantes pero no se asignen a ningun lote
    
    '   Validaciones
    '       Se debe validar que el lote exista y que esté abierto
    '       se debe validar que no se supere la cantidad de postulantes asociados al lote
    
    
    If InsertarEnLote Then
        Flog.writeline "Asocio al Lote "
        Asociar_Postulante = True
        
        StrSql = "SELECT pos_busqueda.busnro,busfin "
        StrSql = StrSql & " FROM pos_busqueda "
        StrSql = StrSql & " LEFT JOIN pos_busqueda_comp ON pos_busqueda_comp.busnro = pos_busqueda.busnro "
        StrSql = StrSql & " WHERE busfin = 0"
        StrSql = StrSql & " AND ternro = " & ternro_temp
        OpenRecordsetWithConn StrSql, rs_Lote, objconn2
        If rs_Lote.EOF Then
            StrSql = "SELECT busnro,busestado, busfin, busfecfin, buscantmaxpost  "
            StrSql = StrSql & " FROM pos_busqueda "
            StrSql = StrSql & " LEFT JOIN pos_estadobusqueda ON pos_estadobusqueda.estbusnro = pos_busqueda.estbusnro "
            StrSql = StrSql & " WHERE busnro = " & NroLote
            OpenRecordsetWithConn StrSql, rs_Lote, objconn2
            If rs_Lote.EOF Then
                Flog.writeline "No se encuentra el Lote " & Lote
                Asociar_Postulante = False
            Else
                'si esta inactivo ==> no puedo agregar postulantes
                If CBool(rs_Lote!busfin) Then
                    Flog.writeline "El lote " & Lote & " se encuentra inactivo."
                    Asociar_Postulante = False
                Else
                    'Veo que no se supere la cantidad de postulantes de la búsqueda
                    StrSql = "SELECT DISTINCT Count(pos_busqueda_comp.ternro) suma "
                    StrSql = StrSql & "FROM pos_busqueda_comp "
                    StrSql = StrSql & "WHERE pos_busqueda_comp.busnro = " & NroLote
                    OpenRecordsetWithConn StrSql, rs_Lote_Post, objconn2
                    If Not rs_Lote_Post.EOF Then
                        If rs_Lote!buscantmaxpost <= rs_Lote_Post!suma Then
                            Flog.writeline "No se puede asociar el postulante al lote " & Lote & " se alcanzó la cantidad máxima."
                            Asociar_Postulante = False
                        End If
                    Else
                        'No hay postulantes asociados al lote. Lo inserto
                    End If
                End If
            End If
        Else
            Flog.writeline "El postulante ya se encuentra asociado a otro un lote activo. Lote nro " & rs_Lote!busnro
            Asociar_Postulante = False
        End If
        
        
        'Asocio el postulante al lote
        If Asociar_Postulante Then
            'tabla pos_busqueda_comp
            Flog.writeline "Agrego postulante "
            StrSql = "INSERT INTO pos_busqueda_comp (busnro,ternro,actnro,estnro) VALUES ("
            StrSql = StrSql & NroLote
            StrSql = StrSql & "," & ternro_temp
            StrSql = StrSql & ",0,1)"
            objconn2.Execute StrSql, , adExecuteNoRecords
        End If
    Else
        Flog.writeline "El postulante no se asociará a ningun Lote"
    End If
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

'Function TraerDomParticularLibre(terNro) As Integer
'    Dim tiposdom As String
'    Dim rs As New ADODB.Recordset
'    tiposdom = "(2,6,7,8,9,10,11)"
'
'    StrSql = " SELECT max(domnro) m "
'    StrSql = StrSql & " FROM cabdom "
'    StrSql = StrSql & " WHERE tidonro IN " & tiposdom
'    StrSql = StrSql & " AND ternro = " & terNro
'    'StrSql = StrSql & " ORDER BY tidonro DESC "
'    OpenRecordset StrSql, rs
'    If rs.EOF Then
'        rs.Close
'        TraerDomParticularLibre = 2
'    Else
'        If IsNull(rs("m")) Then 'no tiene ninguna cargado
'            TraerDomParticularLibre = 2
'        End If
'        If CInt(rs("m")) = 0 Then 'no tiene ninguna cargado
'            TraerDomParticularLibre = 2
'        End If
'        If CInt(rs("m")) = 11 Then 'estan completos
'            TraerDomParticularLibre = moverDomicilios(terNro)
'        Else
'            TraerDomParticularLibre = CInt(rs("m")) + 1
'        End If
'
'        rs.Close
'
'    End If
'
'
'End Function

Sub moverDomicilios(ternro)
    Dim rs_sub As New ADODB.Recordset
    Dim l_domnro As Long
    
    StrSql = " SELECT domnro FROM cabdom WHERE tidonro = 11 AND ternro = " & ternro
    OpenRecordset StrSql, rs_sub
    If Not rs_sub.EOF Then
        l_domnro = rs_sub("domnro")
        rs_sub.Close
        
        StrSql = " DELETE FROM telefono WHERE domnro = " & l_domnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " DELETE FROM detdom WHERE domnro = " & l_domnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
    
    StrSql = " DELETE FROM cabdom WHERE tidonro = 11 AND ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " UPDATE cabdom SET tidonro = 11, domdefault = 0 WHERE tidonro = 10 AND ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " UPDATE cabdom SET tidonro = 10, domdefault = 0 WHERE tidonro = 9 AND ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " UPDATE cabdom SET tidonro = 9, domdefault = 0 WHERE tidonro = 8 AND ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " UPDATE cabdom SET tidonro = 8, domdefault = 0 WHERE tidonro = 7 AND ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " UPDATE cabdom SET tidonro = 7, domdefault = 0 WHERE tidonro = 6 AND ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " UPDATE cabdom SET tidonro = 6, domdefault = 0 WHERE tidonro = 2 AND ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub AsignarEstructura(tipoEstr As Long, estrNro As Long, ternro As Long, Fecha As Date)
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Procedimiento que inserta la estructura. si existe una estructura del mismo tipo en el intervalo
    '               la estructura será actualizada.
    ' Autor      : Lisandro Moro
    ' Fecha      : 25/09/2009
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rs_his As New ADODB.Recordset
    Dim FBaja As Date
    FBaja = DateAdd("d", -1, CDate(Fecha))
    
    StrSql = "SELECT * FROM his_estructura "
    StrSql = StrSql & " WHERE tenro = " & tipoEstr
    StrSql = StrSql & " AND ternro = " & ternro
    StrSql = StrSql & " AND (htetdesde > " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " ORDER BY htetdesde "
    If rs_his.State = adStateOpen Then rs_his.Close
    OpenRecordset StrSql, rs_his
    If Not rs_his.EOF Then
        Texto = ": " & " Existe una estructura con fecha posterior a la ingeresada."
        Call Escribir_Log("floge", LineaCarga, 1, Texto, Tabs, "")
        Exit Sub
    End If
    rs_his.Close
    
    
    'FGZ - 24/11/2009
    'Le agregué este control porque si es el mismo registro hace macanas
    StrSql = " SELECT * FROM his_estructura "
    StrSql = StrSql & " WHERE tenro = " & tipoEstr
    StrSql = StrSql & " AND ternro = " & ternro
    StrSql = StrSql & " AND htetdesde = " & ConvFecha(Fecha)
    StrSql = StrSql & " AND htethasta is null "
    OpenRecordset StrSql, rs_his
    If rs_his.EOF Then
        'Entonces el registro que tiene no es el mismo que el que va a insertar
    
        StrSql = " SELECT * FROM his_estructura "
        StrSql = StrSql & " WHERE tenro = " & tipoEstr
        StrSql = StrSql & " AND ternro = " & ternro
        StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Fecha) & ")"
        StrSql = StrSql & " AND (htethasta >= " & ConvFecha(Fecha) & " OR htethasta is null)"
        StrSql = StrSql & " ORDER BY htetdesde "
        OpenRecordset StrSql, rs_his
        If Not rs_his.EOF Then
            StrSql = " UPDATE his_estructura SET htethasta = " & ConvFecha(FBaja)
            StrSql = StrSql & " WHERE tenro = " & tipoEstr
            StrSql = StrSql & " AND ternro = " & ternro
            StrSql = StrSql & " AND estrnro = " & rs_his!estrNro
            StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_his!htetdesde)
            objConn.Execute StrSql, , adExecuteNoRecords
            Texto = ": " & " Cierro la estructura anterior."
            Call Escribir_Log("floge", LineaCarga, 1, Texto, Tabs, "")
        End If
        
        StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
        StrSql = StrSql & ternro & "," & estrNro & "," & tipoEstr & "," & ConvFecha(Fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Texto = ": " & " Habro la nueva estructura (" & tipoEstr & ")"
        Call Escribir_Log("floge", LineaCarga, 1, Texto, Tabs, "")
    Else
        'No hago nada
    End If
    If rs_his.State = adStateOpen Then rs_his.Close
    Set rs_his = Nothing
End Sub

