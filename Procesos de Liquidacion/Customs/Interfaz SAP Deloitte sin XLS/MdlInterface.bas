Attribute VB_Name = "MdlInterface"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "08/03/2006"
'Global Const UltimaModificacion = " " 'Version inicial

'Global Const Version = "1.02"
'Global Const FechaModificacion = "10/05/2006"
'Global Const UltimaModificacion = " " 'Agregados de algunos campos en el formato 1
'                                      'Consideraciones en formato 2 (IT0021 y IT0394)


'Global Const Version = "1.03"
'Global Const FechaModificacion = "29/05/2006"
'Global Const UltimaModificacion = " " 'Agregados de algunos campos en el formato 1

'Global Const Version = "1.04"
'Global Const FechaModificacion = "02/06/2006"
'Global Const UltimaModificacion = " " 'modificaciones varias

'Global Const Version = "1.05"
'Global Const FechaModificacion = "07/06/2006"
'Global Const UltimaModificacion = " " 'modificaciones varias
'                                        'CUIL
'                                        'En alta genero nro de legajo con uno mas al ultimo existente
'                                        'Tomo la fecha de inicio como la fecha de la medida

'Global Const Version = "1.06"
'Global Const FechaModificacion = "08/06/2006"
'Global Const UltimaModificacion = " " 'el puesto tambien depende de la funcion y es mapeo a mano por ahora queda pendiente = que categoria

'Global Const Version = "1.07"
'Global Const FechaModificacion = "23/06/2006"
'Global Const UltimaModificacion = " " 'Modificaiones sobre el formato2.
''                                           IT0021.
''                                           Lectura y descarte de IT que no se usan.

'Global Const Version = "1.08"
'Global Const FechaModificacion = "09/08/2006"
'Global Const UltimaModificacion = " " 'Modificaiones sobre el formato2.
''                                           IT0021.
''                                           IT2001 y IT2002.

'Global Const Version = "1.09"
'Global Const FechaModificacion = "10/08/2006"
'Global Const UltimaModificacion = " " 'Modificaiones sobre el formato2.
''                                           IT0007, IT0009 y IT0016.


'Global Const Version = "1.10"
'Global Const FechaModificacion = "18/09/2006"
'Global Const UltimaModificacion = " " 'Correcciones Varias:
''                                       'Novedades:
''                                           Caso 1: En rhpro hoy estan cargadas sin vigencia pero desde SAP me lo van
''                                                   a informar con vigencia ==> SI Existe en RHPro sin vigencia
''                                                                                   ==> modifico con vigencia Desde: fecha de alta reconocida y Hasta:fecha desde -1 de la nueva
''                                                                                       Inserto la nueva
''                                                                                   ELSE
''                                                                                       como siempre
''                                           Caso 2: generalmente me van a venir informadas con vigencia desde y sin hasta ==>
''                                                   debe cerrar la que tengo cargada y crear la nueva

Global Const Version = "1.11"
Global Const FechaModificacion = "27/09/2006"
Global Const UltimaModificacion = " " 'Modificaiones sobre el formato1.
'

'------------------------------------------------------------
'------------------------------------------------------------

Global Const Cantidad_Infotipos = 14

Public Type TE_Datos
    ID_campo As Long
    Campo As String
    Descripcion As String
    TipoDato As String
    Valor As String
    IT As String
End Type


Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegFecha As Date
Global NroProceso As Long

Global f
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

Global Fila_Infotipo As Long
Global Arr_Datos(1 To 14) As TE_Datos
Global Formato_IT As Integer
Global Primer_Linea As Boolean


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 08/03/2006
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
    
    Nombre_Arch = PathFLog & "Interfase_sap_DTT" & "-" & NroProcesoBatch & ".log"
    Nombre_Arch2 = PathFLog & "Interfase_sap_DTT_Errores" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    Set FlogE = fs.CreateTextFile(Nombre_Arch2, True)
        
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    
    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 126 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
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
    
Fin:
    Flog.writeline "Fin"
    If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
    Flog.Close
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
    GoTo Fin:
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
        
    'Call Generar_Archivo(NroModelo, Archivo)
    
End Sub


Public Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    ErroresNov = True
End Sub
