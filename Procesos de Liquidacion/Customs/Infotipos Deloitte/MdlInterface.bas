Attribute VB_Name = "MdlInterface"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "06/02/2007"   'Version Inicial. Actualiza Automaticamente El Profit center segun CCosto

'Const Version = "1.02"
'Const FechaVersion = "01/08/2007"   'Nuevo Infotipo 9302 - Prestamos

Const Version = "1.03"
Const FechaVersion = "31/07/2009"   'Encriptacion de string connection

'-----------------------------------------------------------------------------------
Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegFecha As Date
Global NroProceso As Long

Global f
'Global HuboError As Boolean
Global Path
Global NArchivo
Global NroLinea As Long
Global LineaCarga As Long

Global Separador As String
Global UsaSeparadorDeCampos As Boolean
Global SeparadorDecimal As String
Global UsaEncabezado As Boolean

Global ErroresNov As Boolean

Global ErrCarga
Global LineaError
Global LineaOK

Global NroModelo As Long
Global DescripcionModelo As String
Global fExport
Global fNovedades
Global fCambios
Global Fecha_Desde As Date
Global Fecha_Hasta As Date
Global Primera_vez As Boolean
Global ArchivoAGenerar
Global ArchivoNovedades
Global ArchivoCambios

Global Fila_Infotipo_0000 As Long
Global Fila_Infotipo_0001 As Long
Global Fila_Infotipo_0002 As Long
Global Fila_Infotipo_0006 As Long
Global Fila_Infotipo_0008 As Long
Global Fila_Infotipo_0009 As Long
Global Fila_Infotipo_0014 As Long
Global Fila_Infotipo_0015 As Long
Global Fila_Infotipo_0021 As Long
Global Fila_Infotipo_0027 As Long
Global Fila_Infotipo_0041 As Long
Global Fila_Infotipo_0057 As Long
Global Fila_Infotipo_0185 As Long
Global Fila_Infotipo_0389 As Long
Global Fila_Infotipo_0390 As Long
Global Fila_Infotipo_0391 As Long
Global Fila_Infotipo_0392 As Long
Global Fila_Infotipo_0393 As Long
Global Fila_Infotipo_0394 As Long
Global Fila_Infotipo_2001 As Long
Global Fila_Infotipo_2010 As Long
Global Fila_Infotipo_9004 As Long
Global Fila_Infotipo_9302 As Long


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de Infotipos.
' Autor      : FGZ
' Fecha      : 23/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim Nombre_Arch2 As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
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

    'Abro la conexion
    'OpenConnection strconexion, objConn
    'OpenConnection strconexion, objconnProgreso
    
    
    Nombre_Arch = PathFLog & "Infotipos" & "-" & NroProcesoBatch & ".log"
    Nombre_Arch2 = PathFLog & "Infotipos_Errores" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    Set FlogE = fs.CreateTextFile(Nombre_Arch2, True)
        
    ' Obtengo el Process ID
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
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 64 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
'    ErroresNov = False
'    Primera_Vez = False
'    tplaorden = 0
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call LevantarParamteros(bprcparam)
    End If
    
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
    
    objConn.Close
    objconnProgreso.Close
    Flog.Close

End Sub




Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer

Dim NombreArchivo As String
Dim Importar As Boolean


Separador = "@"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        '1- Nro de Modelo
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        NroModelo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        '2- Nombre del archivo a levantar o generar
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
'        '3- Importa o Exporta (TRUE = Importa, FALSE = Exporta)
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Importar = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Importar = True
'
'        '4- Fecha desde
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Fecha_Desde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
'
'        '5- Fecha Hasta
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Fecha_Hasta = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
'
'        '6- Infotipo
'        pos1 = pos2 + 2
'        pos2 = Len(parametros)
'        InfotipoVal = CStr(Mid(parametros, pos1, pos2 - pos1 + 1))
        
    End If
End If
If Importar Then
    Call Importar_Infotipo(NroModelo, NombreArchivo)
Else
    Call Exportar_Infotipo(NroModelo, NombreArchivo)
End If
End Sub


Public Sub Importar_Infotipo(ByVal NroModelo As Long, ByVal NombreArchivo As String)
Dim Directorio As String
Dim CArchivos
Dim Archivo
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
        Separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, "NULL")
        UsaSeparadorDeCampos = True
        If Separador = "NULL" Then
            UsaSeparadorDeCampos = False
        Else
        End If
        SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
        UsaEncabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
        DescripcionModelo = objRs!moddesc
        
        Flog.writeline Espacios(Tabulador * 1) & "Directorio a buscar :  " & Directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo
        Exit Sub
    End If
    
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        Path = Directorio
        
        Dim fc, F1, s2
        Set Folder = fs.GetFolder(Directorio)
        Set CArchivos = Folder.Files
        
        HuboError = False
        'NArchivo = Archivo.Name
        Flog.writeline Espacios(Tabulador * 1) & "Procesando archivo " & NombreArchivo
        Flog.writeline
        Flog.writeline
        Flog.writeline
        Primera_vez = True
        ArchivoAGenerar = Directorio & "\" & Mid(NombreArchivo, 1, 27) & "_" & Format(Now, FormatoInternoHora) & ".xls"
        ArchivoNovedades = Directorio & "\Novadades_" & Mid(NombreArchivo, 1, 27) & "_" & Format(Now, FormatoInternoHora) & ".csv"
        ArchivoCambios = Directorio & "\Cambios_" & Mid(NombreArchivo, 1, 27) & "_" & Format(Now, FormatoInternoHora) & ".csv"
        Call LeeArchivo(Directorio & "\" & NombreArchivo)
        
End Sub


Public Sub Exportar_Infotipo(ByVal NroModelo As Long, ByVal NombreArchivo As String)
Dim Directorio As String
Dim Archivo
Dim Carpeta

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
        Separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, "")
        UsaSeparadorDeCampos = True
        If Separador = "" Then
            UsaSeparadorDeCampos = False
        End If
        SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
        UsaEncabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
        DescripcionModelo = objRs!moddesc
        
        Flog.writeline Espacios(Tabulador * 1) & "Directorio a buscar :  " & Directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo
        Exit Sub
    End If
    
    'Archivo de exportacion
    Archivo = Directorio & "\" & NombreArchivo
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    'Activo el manejador de errores
    On Error Resume Next
    Set fExport = fs.CreateTextFile(Archivo, True)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Set Carpeta = fs.CreateFolder(Directorio)
        Set fExport = fs.CreateTextFile(Archivo, True)
    End If
    'desactivo el manejador de errores
    On Error GoTo 0
        
    Call Generar_Archivo(NroModelo, Archivo)
    
End Sub


Public Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    ErroresNov = True
End Sub
