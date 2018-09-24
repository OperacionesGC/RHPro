Attribute VB_Name = "MdlInterface"
Option Explicit

'Global Const Version = "1.00" 'Lisandro Moro
'Global Const FechaModificacion = "24/09/2009"
'Global Const UltimaModificacion = " "

'Global Const Version = "1.01" 'FGZ
'Global Const FechaModificacion = "24/11/2009"
'Global Const UltimaModificacion = " "   'Le agregué unos controles en la carga de estructuras

Global Const Version = "1.02" 'FGZ
Global Const FechaModificacion = "27/11/2009"
Global Const UltimaModificacion = " "   'Cuando era planificado andaba MAL

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
    
    
    Nombre_Arch = PathFLog & "Importacion_schering" & "-" & NroProcesoBatch & ".log"
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 255 AND bpronro =" & NroProcesoBatch
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
        'Flog.Writeline Espacios(Tabulador * 1) & "Levanta parametros"
        Call LevantarParamteros(bprcparam)
        'Flog.Writeline Espacios(Tabulador * 1) & "fin levanta parametros"
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
                'NroLinea = NroLinea + 1
                'rs_Lineas.MoveNext
            End If
            If Trim(strLinea) <> "" And NroLinea = rs_Lineas!fila Then
                
                'Flog.Writeline Espacios(Tabulador * 0) & "Linea " & NroLinea
                'Select Case rs_Modelo!modinterface
                '    Case 1:
                '        Call Insertar_Linea_Segun_Modelo_Estandar(strLinea)
                '        RegLeidos = RegLeidos + 1
                '    Case 2:
                        Call Insertar_Linea_Segun_Modelo_Custom(strLinea)
                        RegLeidos = RegLeidos + 1
                '    Case 3:
                '        Call Insertar_Linea_Segun_Modelo_MigraInicial(strLinea)
                '        RegLeidos = RegLeidos + 1
                '    Case Else
                '        Flog.writeline Espacios(Tabulador * 0) & "El Modelo " & NroModelo & " no tiene configurado el campo modinterface"
                'End Select
                
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
'    Case 222:
'        Call LineaModelo_222
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
        Case 283: '
            Call LineaModelo_283(linea, OK)
        Case Else
            Texto = ": " & " No existe el modelo "
            Call Escribir_Log("floge", "", 0, Texto, Tabs, linea)
        End Select
    
'If Not HuboError Then
'    MyCommitTrans
'Else
'    MyRollbackTrans
'End If
MyCommitTrans
If Not HuboError Then
    'Flog.Writeline Espacios(Tabulador * 1) & "Transaccion Cometida"
Else
    Flog.writeline Espacios(Tabulador * 1) & "Transaccion Abortada"
End If

End Sub

Public Sub LineaModelo_283(ByVal strReg As String, ByRef OK As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Empleados
' Autor      : Lisandro Moro
' Fecha      : 24/09/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Long
Dim pos2            As Long

Dim Legajo          As String   'LEGAJO                   -- empleado.empleg(6)
Dim Apellido        As String   'APELLIDO                 -- empleado.terape y tercero.terape(25)
Dim nombre          As String   'NOMBRE                   -- empleado.ternom y tercero.ternom(25)
Dim Calle           As String   'Calle                    -- detdom.calle(30)
Dim Nro             As String   'Número                   -- detdom.nro(8)
Dim Piso            As String   'Piso                     -- detdom.piso(8)
Dim Depto           As String   'Depto                    -- detdom.depto(8)
Dim Torre           As String   'Torre                    -- detdom.torre(8)
Dim Manzana         As String   'Manzana                  -- detdom.manzana(8)
Dim Cpostal         As String   'Cpostal                  -- detdom.codigopostal(12)
Dim Entre           As String   'Entre Calles             -- detdom.entrecalles(80)
Dim Barrio          As String   'Barrio                   -- detdom.barrio(30)
Dim Localidad       As String   'Localidad                -- detdom.locnro(30)localidad.locdesc(60)
Dim Partido         As String   'Partido                  -- detdom.partnro(30)partido.partnom(30)
Dim Zona            As String   'Zona                     -- detdom.zonanro(20)zona.zonadesc(60)
Dim Provincia       As String   'Provincia                -- detdom.provnro(20)provincia.provdesc(20)
Dim Pais            As String   'Pais                     -- detdom.paisnro(20)pais.paisdesc(20)
Dim Telefono        As String   'Telefono                 -- telefono.telnro(20)
Dim TelCelular      As String   'Telefono                 -- telefono.telnro(20)
Dim UnidadNegocio   As String   'Unidad de negocio        -- his_estructura.estrnro(60)
Dim Direccion       As String   'Direccion                -- his_estructura.estrnro(60)
Dim Gerencia        As String   'Gerencia                 -- his_estructura.estrnro(60)
Dim Departamento    As String   'Departamento             -- his_estructura.estrnro(60)
Dim Puesto          As String   'Puesto                   -- his_estructura.estrnro(60)
Dim Contrato        As String   'tipo de contrato         -- his_estructura.estrnro(60)
Dim OSocialElegida  As String   'Obra Social              -- his_estructura.estrnro(60)
Dim Banda           As String   'Banda                    -- his_estructura.estrnro(60)
Dim Sindicato       As String   'Sindicato                -- his_estructura.estrnro(60)
Dim Convenio        As String   'Convenio                 -- his_estructura.estrnro(60)
Dim Sucursal        As String   'Sucursal                 -- his_estructura.estrnro(60)
Dim Imputacion      As String   'Imputacion contable      -- his_estructura.estrnro(60)
Dim CCosto          As String   'C.Costo                  -- his_estructura.estrnro(60)
Dim LugarRecibo     As String   'LugarRecibo              -- his_estructura.estrnro(60)
Dim FechaCambio     As Date     'Fecha CAmbio estructura  -- his_estructura.fecalta(10)
Dim fechavalida  As Boolean


Dim l_nro_UnidadNegocio   As Long
Dim l_nro_Direccion       As Long
Dim l_nro_Gerencia        As Long
Dim l_nro_Departamento    As Long
Dim l_nro_Puesto          As Long
Dim l_nro_Contrato        As Long
Dim l_nro_OSocialElegida  As Long
Dim l_nro_Banda           As Long
Dim l_nro_Sindicato       As Long
Dim l_nro_Convenio        As Long
Dim l_nro_Sucursal        As Long
Dim l_nro_Imputacion      As Long
Dim l_nro_CCosto          As Long
Dim l_nro_LugarRecibo     As Long


Dim sql As String
Dim sqlc As String
Dim sqld As String
Dim tidonro As Long

Dim ternro As Long
Dim NroTercero          As Long
Dim Nro_Legajo          As Long
Dim Nro_Nrodom          As String

    Dim l_domicilio_ok As Boolean
    l_domicilio_ok = True
    Dim l_mover_domicilio As Boolean
    l_mover_domicilio = True
    Dim l_Calle As String
    Dim l_Nro As String
    Dim l_Piso As String
    Dim l_Depto As String
    Dim l_Torre As String
    Dim l_Manzana As String
    Dim l_Cpostal As String
    Dim l_Entre As String
    Dim l_Barrio As String
    Dim l_Localidad As String '
    Dim l_Partido As String
    Dim l_Zona As String
    Dim l_Provincia As String '
    Dim l_Pais As Integer '
    Dim l_nro_Pais As Long
    Dim l_nro_Provincia As Long
    Dim l_nro_Localidad As Long
    Dim l_nro_zona As Long
    Dim l_nro_partido As Long
    
    Dim l_auxchr1 As String
    Dim l_auxchr2 As String
    Dim l_email As String
    Dim l_kilometro As String
    Dim l_circunscripcion As String
    Dim l_cuerpo As String
    Dim l_lote As String
    Dim l_parcela As String
    Dim l_bloque As String
    Dim l_seccion As String
    Dim l_casa As String
    Dim l_cpa As String
    Dim domnro_ant As Long
    Dim nrodom As Long
    domnro_ant = 0

Dim rs As New ADODB.Recordset
Dim rs_emp As New ADODB.Recordset
Dim rs_Tel As New ADODB.Recordset

On Error GoTo SaltoLinea
    
'    Sigue = True 'Indica que si en el archivo viene mas de una vez un empleado, le crea las fases
'    ExisteLeg = False
'    'RegLeidos = RegLeidos + 1
'    LineaCarga = LineaCarga + 1
    
    Flog.writeline
    FlogE.writeline
    FlogP.writeline
    
    'Legajo
    nrocolumna = nrocolumna + 1
    Campoetiqueta = "Legajo"
    pos1 = 1
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Legajo = aux
    StrSql = "SELECT ternro FROM empleado WHERE empleado.empleg = " & Legajo
    OpenRecordset StrSql, rs_emp
    If rs_emp.EOF Then
        Texto = ": " & " - El Empleado No Existe. Legajo: " & Legajo
        nrocolumna = 1
        Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
        Exit Sub
    Else
        NroTercero = rs_emp!ternro
        ternro = NroTercero
    End If
    
    
    'Apellido
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Apellido = Left(aux, 25)
    
    'Nombre
    nrocolumna = nrocolumna + 1
    Campoetiqueta = "Nombre"
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    nombre = Left(aux, 25)
   
    'Calle
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Calle = Left(aux, 30)
    
    'Nro
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Nro = Left(aux, 8)
    
    'Piso
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Piso = Left(aux, 8)
    
    'Departamento
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Depto = Left(aux, 8)

    'Torre
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Torre = Left(aux, 8)
    
    'Manzana
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Manzana = Left(aux, 8)

    'Codigo Postal
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Cpostal = Left(aux, 12)

    'Entre calles
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Entre = Left(aux, 80)

    'Barrio
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Barrio = Left(aux, 30)

    'Localidad
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Localidad = Left(aux, 60)
    
    'Partido
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Partido = Left(aux, 30)
    
    'Zona
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Zona = Left(aux, 60)
    
    'Provincia
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Provincia = Left(aux, 20)
    
    'Pais
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Pais = Left(aux, 20)
   
    'Tel Particular
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Telefono = Left(aux, 20)
    
    'Tel Celular
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    TelCelular = Left(aux, 20)
   
   'Unidad de Negocio
    nrocolumna = nrocolumna + 1
    Campoetiqueta = "Direccion"
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    UnidadNegocio = Left(aux, 60)
   
    'Direccion
    nrocolumna = nrocolumna + 1
    Campoetiqueta = "Direccion"
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Direccion = Left(aux, 60)
    
    'Gerencia
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Gerencia = Left(aux, 60)
    
    'Departamento
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Departamento = Left(aux, 60)
    
    'Puesto
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Puesto = Left(aux, 60)
    
    'Contrato
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Contrato = Left(aux, 60)
    
    'OS Elegida
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    OSocialElegida = Left(aux, 60)
    
    'Banda
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Banda = Left(aux, 60)
    
    'Sindicato
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Sindicato = Left(aux, 60)
    
    'Convenio
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Convenio = Left(aux, 60)
    
    'Sucursal
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Sucursal = Left(aux, 60)
    
    'Imputacion
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Imputacion = Left(aux, 60)
    
    'Centro de Costo
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    CCosto = Left(aux, 60)
    
    'Lugar recivo
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    LugarRecibo = Left(aux, 60)
    
    'Fecha Cambio
    Dim F_Cambio As String
    
    nrocolumna = nrocolumna + 1
    pos1 = pos2 + 2
    pos2 = Len(strReg) 'InStr(pos1, strReg, separador) - 1
    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    F_Cambio = Left(aux, 10)
    If F_Cambio = "" Or EsNulo(F_Cambio) Or UCase(F_Cambio) = "NULL" Then
        F_Cambio = "Null"
        Texto = ": " & " - La Fecha de Alta es Obligatoria, no se puesden aingresar las estructuras."
        nrocolumna = 10
        Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
        'Ok = False
        RegError = RegError + 1
        'Exit Sub
        fechavalida = False
    Else
        FechaCambio = CDate(F_Cambio)
        fechavalida = True
    End If

    '------------------------------------------------------------------
    'Hasta Aqui los Datos Obligatorios del Empleado
    'Fin lectura de campos
    ' =====================================================================================================
    ' Inserto el Tercero
    'Actualizo el apellido del empleado
    If Apellido <> "" Then
        If UCase(Apellido) = "NULL" Then
'            StrSql = " UPDATE tercero SET terape = null WHERE ternro = " & NroTercero
'            objConn.Execute StrSql, , adExecuteNoRecords
'            StrSql = " UPDATE empleado SET terape = '' WHERE ternro = " & NroTercero
'            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = ": " & " No se puede borrarel Apellido o dejarlo vacio - "
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
        Else
            StrSql = " UPDATE tercero SET terape = '" & Apellido & "' WHERE ternro = " & NroTercero
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " UPDATE empleado SET terape = '" & Apellido & "' WHERE ternro = " & NroTercero
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = ": " & " Actualizo el Apellido - "
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
        End If
    Else
        'nada
        Texto = ": " & " No se puede dejar el Apellido en blanco- "
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
    End If
    
    'Actualizo el Nombre del empleado
    If nombre <> "" Then
        If UCase(nombre) = "NULL" Then
'            StrSql = " UPDATE tercero SET ternom = null WHERE ternro = " & NroTercero
'            objConn.Execute StrSql, , adExecuteNoRecords
'            StrSql = " UPDATE empleado SET ternom = '' WHERE ternro = " & NroTercero
'            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = ": " & "No se puede Borrar el Nombre o dejarlo vacio - "
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
        Else
            StrSql = " UPDATE tercero SET ternom = '" & nombre & "' WHERE ternro = " & NroTercero
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " UPDATE empleado SET ternom = '" & nombre & "' WHERE ternro = " & NroTercero
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = ": " & " Actualizo el Nombre - "
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
        End If
    Else
        'nada
        Texto = ": " & " No se puede dejar el Nombre en blanco- "
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
    End If
  

    ' Inserto el Domicilio
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    
    If Calle <> "" Or Nro <> "" Or Piso <> "" Or Depto <> "" Or Torre <> "" Or Manzana <> "" Or Cpostal <> "" Or Entre <> "" Or Barrio <> "" Or Localidad <> "" Or Partido <> "" Or Zona <> "" Or Provincia <> "" Or Pais <> "" Then
        'Asumo que hay una modificacion o alta

        StrSql = " SELECT cabdom.domnro, Calle , Nro, Piso, oficdepto, Torre, Manzana, codigopostal, entrecalles, Barrio, locnro, partnro, zonanro, provnro, paisnro "
        StrSql = StrSql & " , auxchr1, auxchr2, email, kilometro, circunscripcion, cuerpo, lote, parcela, bloque, seccion, casa, cpa "
        StrSql = StrSql & " FROM detdom "
        StrSql = StrSql & " INNER JOIN cabdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " WHERE ternro = " & NroTercero
        StrSql = StrSql & " AND tidonro = 2 "
        StrSql = StrSql & " ORDER BY domdefault "
        OpenRecordset StrSql, rs_emp
        If Not rs_emp.EOF Then
            'Busco el domicilio actual
            l_mover_domicilio = True
        
            nrodom = rs_emp("domnro")
            domnro_ant = nrodom
            
            If Calle <> "" Then
                If UCase(Calle) = "NULL" Then
                    l_Calle = "null"
                Else
                    l_Calle = "'" & Calle & "'"
                End If
            Else
                l_Calle = "'" & rs_emp("Calle") & "'"
            End If
            
            If Nro <> "" Then
                If UCase(Nro) = "NULL" Then
                    l_Nro = "null"
                Else
                    l_Nro = "'" & Nro & "'"
                End If
            Else
                l_Nro = "'" & rs_emp("Nro") & "'"
            End If
            
            If Piso <> "" Then
                If UCase(Piso) = "NULL" Then
                    l_Piso = "null"
                Else
                    l_Piso = "'" & Piso & "'"
                End If
            Else
                l_Piso = "'" & rs_emp("Piso") & "'"
            End If
            
            If Depto <> "" Then
                If UCase(Depto) = "NULL" Then
                    l_Depto = "null"
                Else
                    l_Depto = "'" & Depto & "'"
                End If
            Else
                l_Depto = "'" & rs_emp("oficdepto") & "'"
            End If
            
            If Torre <> "" Then
                If UCase(Torre) = "NULL" Then
                    l_Torre = "null"
                Else
                    l_Torre = "'" & Torre & "'"
                End If
            Else
                l_Torre = "'" & rs_emp("Torre") & "'"
            End If
            
            If Manzana <> "" Then
                If UCase(Manzana) = "NULL" Then
                    l_Manzana = "null"
                Else
                    l_Manzana = "'" & Manzana & "'"
                End If
            Else
                l_Manzana = "'" & rs_emp("Manzana") & "'"
            End If
            
            If Cpostal <> "" Then
                If UCase(Cpostal) = "NULL" Then
                    l_Cpostal = "null"
                Else
                    l_Cpostal = "'" & Cpostal & "'"
                End If
            Else
                l_Cpostal = "'" & rs_emp("codigopostal") & "'"
            End If
            
            If Entre <> "" Then
                If UCase(Entre) = "NULL" Then
                    l_Entre = "null"
                Else
                    l_Entre = "'" & Entre & "'"
                End If
            Else
                l_Entre = "'" & rs_emp("entrecalles") & "'"
            End If
            
            If Barrio <> "" Then
                If UCase(Barrio) = "NULL" Then
                    l_Barrio = "null"
                Else
                    l_Barrio = "'" & Barrio & "'"
                End If
            Else
                l_Barrio = "'" & rs_emp("Barrio") & "'"
            End If
            
            
            If Pais <> "" Then
                If UCase(Pais) = "NULL" Then
                    l_Pais = "null"
                    Texto = ": " & " - No se ingreso el Pais "
                    nrocolumna = 19
                    Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                    l_domicilio_ok = False
                    'Exit Sub
                Else
                    'l_Pais = Pais
                    Call ValidarPais(Pais, l_nro_Pais)
                End If
            Else
                'l_Pais = rs_emp("pais")
                Call ValidarPais(Pais, l_nro_Pais)
                Texto = ": " & " - No se ingreso el Pais "
                nrocolumna = 19
                Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                'Exit Sub
                l_domicilio_ok = False
            End If
            
            If Provincia <> "" Then
                If UCase(Provincia) = "NULL" Then
                    l_Provincia = "null"
                    Texto = ": " & " - No se ingreso La Provincia "
                    nrocolumna = 18
                    Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                    l_domicilio_ok = False
                    'Exit Sub
                Else
                    'l_Provincia = rs_emp("provnro")
                    Call ValidarProvincia(Provincia, l_nro_Provincia, l_nro_Pais)
                End If
            Else
                'l_Provincia = Provincia
                Texto = ": " & " - No se ingreso La Provincia "
                nrocolumna = 18
                Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                'Exit Sub
                l_domicilio_ok = False
            End If
            
            If Localidad <> "" Then
                If UCase(Localidad) = "NULL" Then
                    l_Localidad = "null"
                    Texto = ": " & " - No se ingreso La Localidad "
                    nrocolumna = 15
                    Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                    'Exit Sub
                    l_domicilio_ok = False
                Else
                    'l_Localidad = Localidad
                    Call ValidarLocalidad(Localidad, l_nro_Localidad, l_nro_Pais, l_nro_Provincia)
                End If
            Else
                'l_Localidad = rs_emp("locnro")
                Texto = ": " & " - No se ingreso La Localidad "
                nrocolumna = 15
                Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                'Exit Sub
                l_domicilio_ok = False
            End If
            
            
            If Partido <> "" Then
                If UCase(Partido) = "NULL" Then
                    l_Partido = "null"
                    l_nro_partido = 0
                Else
                    'l_Partido = partnro
                    Call ValidarPartido(Partido, l_nro_partido)
                End If
            Else
                l_nro_partido = IIf(EsNulo(rs_emp("partnro")), 0, rs_emp("partnro"))
            End If
            
            
            If Zona <> "" Then
                If UCase(Zona) = "NULL" Then
                    l_Zona = "null"
                    l_nro_zona = 0
                Else
                    'l_Zona = zonanro
                    Call ValidarZona(Zona, l_nro_zona, l_nro_Provincia)
                End If
            Else
                l_nro_zona = IIf(EsNulo(rs_emp("Zonanro")), 0, rs_emp("Zonanro"))
            End If

            l_auxchr1 = "'" & rs_emp("auxchr1") & "'"
            l_auxchr2 = "'" & rs_emp("auxchr2") & "'"
            l_email = "'" & rs_emp("email") & "'"
            l_kilometro = "'" & rs_emp("kilometro") & "'"
            l_circunscripcion = "'" & rs_emp("circunscripcion") & "'"
            l_cuerpo = "'" & rs_emp("cuerpo") & "'"
            l_lote = "'" & rs_emp("lote") & "'"
            l_parcela = "'" & rs_emp("parcela") & "'"
            l_bloque = "'" & rs_emp("bloque") & "'"
            l_seccion = "'" & rs_emp("seccion") & "'"
            l_casa = "'" & rs_emp("casa") & "'"
            l_cpa = "'" & rs_emp("cpa") & "'"
        
            
        Else
            'No posee domicilio actual
            l_mover_domicilio = False

            If UCase(Calle) = "NULL" Then
                l_Calle = "null"
            Else
                l_Calle = "'" & Calle & "'"
            End If
            
            If UCase(Nro) = "NULL" Then
                l_Nro = "null"
            Else
                l_Nro = "'" & Nro & "'"
            End If
        
        
            If UCase(Piso) = "NULL" Then
                l_Piso = "null"
            Else
                l_Piso = "'" & Piso & "'"
            End If
            
            If UCase(Depto) = "NULL" Then
                l_Depto = "null"
            Else
                l_Depto = "'" & Depto & "'"
            End If
            
            If UCase(Torre) = "NULL" Then
                l_Torre = "null"
            Else
                l_Torre = "'" & Torre & "'"
            End If
            
            If UCase(Manzana) = "NULL" Then
                l_Manzana = "null"
            Else
                l_Manzana = "'" & Manzana & "'"
            End If
            
            If UCase(Cpostal) = "NULL" Then
                l_Cpostal = "null"
            Else
                l_Cpostal = "'" & Cpostal & "'"
            End If
            
            If UCase(Entre) = "NULL" Then
                l_Entre = "null"
            Else
                l_Entre = "'" & Entre & "'"
            End If
            
            If UCase(Barrio) = "NULL" Then
                l_Barrio = "null"
            Else
                l_Barrio = "'" & Barrio & "'"
            End If
            
            If Pais <> "" Then
                If UCase(Pais) = "NULL" Then
                    l_Pais = "null"
                    Texto = ": " & " - No se ingreso el Pais "
                    nrocolumna = 19
                    Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                    'Exit Sub
                    l_domicilio_ok = False
                Else
                    'l_Pais = Pais
                    Call ValidarPais(Pais, l_nro_Pais)
                End If
            Else
                'l_Pais = rs_emp("pais")
                Call ValidarPais(Pais, l_nro_Pais)
                Texto = ": " & " - No se ingreso el Pais "
                nrocolumna = 19
                Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                'Exit Sub
                l_domicilio_ok = False
            End If
            
            If Provincia <> "" Then
                If UCase(Provincia) = "NULL" Then
                    l_Provincia = "null"
                    Texto = ": " & " - No se ingreso La Provincia "
                    nrocolumna = 18
                    Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                    'Exit Sub
                    l_domicilio_ok = False
                Else
                    'l_Provincia = rs_emp("provnro")
                    Call ValidarProvincia(Provincia, l_nro_Provincia, l_nro_Pais)
                End If
            Else
                'l_Provincia = Provincia
                Texto = ": " & " - No se ingreso La Provincia "
                nrocolumna = 18
                Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                'Exit Sub
                l_domicilio_ok = False
            End If
            
            If Localidad <> "" Then
                If UCase(Localidad) = "NULL" Then
                    l_Localidad = "null"
                    Texto = ": " & " - No se ingreso La Localidad "
                    nrocolumna = 15
                    Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                    'Exit Sub
                    l_domicilio_ok = False
                Else
                    'l_Localidad = Localidad
                    Call ValidarLocalidad(Localidad, l_nro_Localidad, l_nro_Pais, l_nro_Provincia)
                End If
            Else
                'l_Localidad = rs_emp("locnro")
                Texto = ": " & " - No se ingreso La Localidad "
                nrocolumna = 15
                Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
                'Exit Sub
                l_domicilio_ok = False
            End If
            
                
            If UCase(Partido) = "NULL" Then
                l_Partido = "null"
                l_nro_partido = 0
            Else
                'l_Partido = partnro
                Call ValidarPartido(Partido, l_nro_partido)
            End If
        
        
            If UCase(Zona) = "NULL" Then
                l_Zona = "null"
                l_nro_zona = 0
            Else
                'l_Zona = zonanro
                Call ValidarZona(Zona, l_nro_zona, l_nro_Provincia)
            End If

            l_auxchr1 = "null"
            l_auxchr2 = "null"
            l_email = "null"
            l_kilometro = "null"
            l_circunscripcion = "null"
            l_cuerpo = "null"
            l_lote = "null"
            l_parcela = "null"
            l_bloque = "null"
            l_seccion = "null"
            l_casa = "null"
            l_cpa = "null"

        End If
        

        If l_domicilio_ok Then
            If l_mover_domicilio Then
                moverDomicilios (ternro)
            End If
            
            StrSql = " INSERT INTO cabdom (tipnro, ternro, domdefault, tidonro) "
            StrSql = StrSql & " VALUES (1, " & ternro & ", -1, 2)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = ": " & "Inserte la cabecera del Domicilio - "
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
            
            nrodom = getLastIdentity(objConn, "cabdom")
            
            StrSql = " INSERT INTO detdom (domnro, Calle , Nro, Piso, oficdepto, Torre, Manzana, codigopostal"
            StrSql = StrSql & " , entrecalles, Barrio, locnro, partnro, zonanro, provnro, paisnro "
            StrSql = StrSql & " , auxchr1, auxchr2, email, kilometro, circunscripcion, cuerpo, lote"
            StrSql = StrSql & " , parcela, bloque, seccion, casa, cpa ) "
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "( " & nrodom & ","
            StrSql = StrSql & "" & l_Calle & ","
            StrSql = StrSql & "" & l_Nro & ","
            StrSql = StrSql & "" & l_Piso & ","
            StrSql = StrSql & "" & l_Depto & ","
            StrSql = StrSql & "" & l_Torre & ","
            StrSql = StrSql & "" & l_Manzana & ","
            StrSql = StrSql & "" & l_Cpostal & ","
            StrSql = StrSql & "" & l_Entre & ","
            StrSql = StrSql & "" & l_Barrio & ","
            StrSql = StrSql & "" & l_nro_Localidad & ","
            StrSql = StrSql & "" & IIf(l_nro_partido = 0, "Null", l_nro_partido) & ","
            StrSql = StrSql & "" & IIf(l_nro_zona = 0, "Null", l_nro_zona) & ","
            StrSql = StrSql & "" & l_nro_Provincia & ","
            StrSql = StrSql & "" & l_nro_Pais & ","
            StrSql = StrSql & "" & l_auxchr1 & ","
            StrSql = StrSql & "" & l_auxchr2 & ","
            StrSql = StrSql & "" & l_email & ","
            StrSql = StrSql & "" & l_kilometro & ","
            StrSql = StrSql & "" & l_circunscripcion & ","
            StrSql = StrSql & "" & l_cuerpo & ","
            StrSql = StrSql & "" & l_lote & ","
            StrSql = StrSql & "" & l_parcela & ","
            StrSql = StrSql & "" & l_bloque & ","
            StrSql = StrSql & "" & l_seccion & ","
            StrSql = StrSql & "" & l_casa & ","
            StrSql = StrSql & "" & l_cpa & ""
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = ": " & "Inserto el Domicilio - "
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                
            'Busco los telefonos anteriores y los copio al nuevo domicilio
            If domnro_ant <> 0 Then
                StrSql = "SELECT * from telefono WHERE domnro = " & domnro_ant
                OpenRecordset StrSql, rs_emp
                If Not rs_emp.EOF Then
                    Do While Not rs_emp.EOF
                        StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular, tipotel) "
                        StrSql = StrSql & " VALUES(" & nrodom & ",'" & rs_emp("telnro") & "'," & rs_emp("telfax") & "," & rs_emp("teldefault") & "," & rs_emp("telcelular") & ", " & rs_emp("tipotel") & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Texto = ": " & "Inserto el Telefono Principal asociado al domicilio anterior - "
                        Call Escribir_Log("flogp", LineaCarga, 20, Texto, Tabs + 1, strReg)
                        rs_emp.MoveNext
                    Loop
                Else
                End If
            End If
        Else
            Texto = ": " & " - No se modifica el domicilio "
            nrocolumna = 0
            Call Escribir_Log("flog", LineaCarga, nrocolumna, Texto, Tabs, strReg)
           ' Exit Sub
        End If
    End If
    

    If Telefono <> "" Or TelCelular <> "" Then
        StrSql = " SELECT detdom.domnro, Calle , Nro, Piso, oficdepto, Torre, Manzana, codigopostal, entrecalles, Barrio, locnro, partnro, zonanro, provnro, paisnro "
        StrSql = StrSql & " , auxchr1, auxchr2, email, kilometro, circunscripcion, cuerpo, lote, parcela, bloque, seccion, casa, cpa "
        StrSql = StrSql & " FROM detdom "
        StrSql = StrSql & " INNER JOIN cabdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " WHERE ternro = " & NroTercero
        StrSql = StrSql & " AND tidonro = 2 "
        StrSql = StrSql & " ORDER BY domdefault "
        OpenRecordset StrSql, rs_emp
        If Not rs_emp.EOF Then
            'Busco el domicilio actual
            nrodom = rs_emp("domnro")
            rs_emp.Close
            
            If Telefono <> "" Then
                StrSql = "SELECT telnro from telefono WHERE domnro = " & nrodom & " AND teldefault = -1 AND tipotel = 1"
                OpenRecordset StrSql, rs_emp
                If rs_emp.EOF Then
                    If Telefono <> "null" Then
                        StrSql = " INSERT INTO telefono(domnro,telnro,telfax,telcelular, tipotel) "
                        StrSql = StrSql & " VALUES(" & nrodom & ",'" & Telefono & "',0,0, 1)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Texto = ": " & "Inserte el Telefono Principal - "
                        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    End If
                Else
                    If Telefono <> "null" Then
                        StrSql = " UPDATE telefono SET telnro = '" & Telefono & "'"
                        StrSql = StrSql & " WHERE domnro = " & nrodom & " AND tipotel = 1 AND telnro = '" & rs_emp("telnro") & "'"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Texto = ": " & "Actualizo el Telefono Celular - "
                        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    Else
                        StrSql = " DELETE FROM telefono "
                        StrSql = StrSql & " WHERE domnro = " & nrodom & " AND tipotel = 1 AND telnro = '" & rs_emp("telnro") & "'"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Texto = ": " & "Borro el Telefono Principal - "
                        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    End If
                    
                End If
            End If
            If TelCelular <> "" Then
                StrSql = "SELECT * FROM telefono "
                StrSql = StrSql & " WHERE domnro =" & nrodom
                'StrSql = StrSql & " AND telnro ='" & TelCelular & "'"
                StrSql = StrSql & " AND tipotel = 2 "
                If rs_Tel.State = adStateOpen Then rs_Tel.Close
                OpenRecordset StrSql, rs_Tel
                If rs_Tel.EOF Then
                    If TelCelular <> "null" Then
                        StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular, tipotel) "
                        StrSql = StrSql & " VALUES(" & nrodom & ",'" & TelCelular & "',0,0,-1,2)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        Texto = ": " & "Inserte el Telefono Celular - "
                        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    End If
                Else
                    If TelCelular <> "null" Then
                        StrSql = " UPDATE telefono SET telnro = '" & TelCelular & "'"
                        StrSql = StrSql & " WHERE domnro = " & nrodom & " AND telcelular = -1 AND tipotel = 2 AND telnro = '" & rs_Tel("telnro") & "'"
                        Texto = ": " & "Actualizo el Telefono Celular - "
                        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    Else
                        StrSql = " DELETE FROM telefono "
                        StrSql = StrSql & " WHERE domnro = " & nrodom & " AND telcelular = -1 AND tipotel = 2 AND telnro = '" & rs_Tel("telnro") & "'"
                        Texto = ": " & "Borro el Telefono Celular - "
                        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        
        Else
            'Si no posee un domicilio,que es raro
            Texto = ": " & " - No se ingresaron los telefonos, no posee domicilio "
            nrocolumna = 20
            Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
            Exit Sub
        End If
    End If
    
  
  
 'Inserto las Estructuras
'  Call AsignarEstructura(41, nro_bancopago, NroTercero, F_Alta, F_Baja)
  
 
    If Not fechavalida Then
    Else
        
        If UnidadNegocio <> "" Then
            l_nro_UnidadNegocio = ValidaEstructura(44, UnidadNegocio)
            AsignarEstructura 44, l_nro_UnidadNegocio, NroTercero, FechaCambio
        End If
        
        If Direccion <> "" Then
            l_nro_Direccion = ValidaEstructura(35, Direccion)
            AsignarEstructura 35, l_nro_Direccion, NroTercero, FechaCambio
        End If
        
        If Gerencia <> "" Then
            l_nro_Gerencia = ValidaEstructura(6, Gerencia)
            AsignarEstructura 6, l_nro_Gerencia, NroTercero, FechaCambio
        End If
        
        If Departamento <> "" Then
            l_nro_Departamento = ValidaEstructura(9, Departamento)
            AsignarEstructura 9, l_nro_Departamento, NroTercero, FechaCambio
        End If
        
        If Puesto <> "" Then
            l_nro_Puesto = ValidaEstructura(4, Puesto)
            AsignarEstructura 4, l_nro_Puesto, NroTercero, FechaCambio
        End If
                
        If Contrato <> "" Then
            l_nro_Contrato = ValidaEstructura(18, Contrato)
            AsignarEstructura 18, l_nro_Contrato, NroTercero, FechaCambio
        End If
        
        If OSocialElegida <> "" Then
            l_nro_OSocialElegida = ValidaEstructura(17, OSocialElegida)
            AsignarEstructura 17, l_nro_OSocialElegida, NroTercero, FechaCambio
        End If
        
        If Banda <> "" Then
            l_nro_Banda = ValidaEstructura(45, Banda)
            AsignarEstructura 45, l_nro_Banda, NroTercero, FechaCambio
        End If
        
        If Sindicato <> "" Then
            l_nro_Sindicato = ValidaEstructura(16, Sindicato)
            AsignarEstructura 16, l_nro_Sindicato, NroTercero, FechaCambio
        End If
        
        If Convenio <> "" Then
            l_nro_Convenio = ValidaEstructura(19, Convenio)
            AsignarEstructura 19, l_nro_Convenio, NroTercero, FechaCambio
        End If
        
        If Sucursal <> "" Then
            l_nro_Sucursal = ValidaEstructura(1, Sucursal)
            AsignarEstructura 1, l_nro_Sucursal, NroTercero, FechaCambio
        End If
        
        If Imputacion <> "" Then
            l_nro_Imputacion = ValidaEstructura(46, Imputacion)
            AsignarEstructura 46, l_nro_Imputacion, NroTercero, FechaCambio
        End If
        
        If CCosto <> "" Then
            l_nro_CCosto = ValidaEstructura(5, CCosto)
            AsignarEstructura 5, l_nro_CCosto, NroTercero, FechaCambio
        End If
        
        If LugarRecibo <> "" Then
            l_nro_LugarRecibo = ValidaEstructura(47, LugarRecibo)
            AsignarEstructura 47, l_nro_LugarRecibo, NroTercero, FechaCambio
        End If
    End If
  
  
Texto = ": " & "Linea procesada correctamente "
Call Escribir_Log("flogp", LineaCarga, nrocolumna, Texto, Tabs + 1, strReg)
'LineaOK.Writeline Mid(strReg, 1, Len(strReg))
OK = True
         
FinLinea:
If rs.State = adStateOpen Then
    rs.Close
End If
Exit Sub

SaltoLinea:
    Texto = ": " & " - Error:" & Err.Description
    nrocolumna = 1
    Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
    MyRollbackTrans
    OK = False
    GoTo FinLinea
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
            Nro_Pais = rs_sub!paisnro
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

