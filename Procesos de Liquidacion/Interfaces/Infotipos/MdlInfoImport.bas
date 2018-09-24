Attribute VB_Name = "MdlInfoImport"
Option Explicit

Public Sub LeeArchivo(ByVal NombreArchivo As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento
' Autor      : FGZ
' Fecha      : 23/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim NroModelo As Long

    If App.PrevInstance Then Exit Sub

    'Espero hasta que se crea el archivo
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then Err.Number = 1
    Loop
    On Error GoTo 0
   
   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProcesoBatch & "," & NroModelo & ",'" & Left(NombreArchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "inter_pin")
    End If
                
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If rs_Modelo.EOF Then
        Exit Sub
    End If
                
    StrSql = "SELECT * FROM modelo_filas WHERE bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_Lineas
    If Not rs_Lineas.EOF Then
        rs_Lineas.MoveFirst
    End If
    
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = rs_Lineas.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (100 / CEmpleadosAProc)
    
    Do While Not f.AtEndOfStream And Not rs_Lineas.EOF
        strLinea = f.ReadLine
        NroLinea = NroLinea + 1
        If NroLinea = 1 And UsaEncabezado Then
            strLinea = f.ReadLine
        End If
        If Trim(strLinea) <> "" And NroLinea = rs_Lineas!fila Then
            RegLeidos = RegLeidos + 1
            
            Select Case rs_Modelo!modinterface
                Case 4:
                    Call Insertar_Linea_Segun_Infotipo(strLinea)
                Case Else
                    Call Insertar_Linea_Segun_Infotipo(strLinea)
            End Select
            rs_Lineas.MoveNext
        End If
        
        'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
        'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
        'como colgado
        Progreso = Progreso + IncPorc
        'If Progreso > 100 Then Progreso = 100
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Loop
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    f.Close
    Flog.Writeline Espacios(Tabulador * 1) & "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'Borrar el archivo
    fs.Deletefile NombreArchivo, True
    
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    Exit Sub
    
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub




Public Sub Insertar_Linea_Segun_Infotipo(ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador segun infotipo
' Autor      : FGZ
' Fecha      : 23/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'0000    Informacion de alta
'0001    Asignacion Organizacional
'0002    Datos Personales
'0006    Direcciones
'0008    Emolumentos Basicos
'0009    Relacion Bancaria
'0016    Elementos de Contratos
'0021    Datos Familiares
'0032    Datos Internos de la Empresa
'0041    Datos de Fecha
'0057    Asociaciones
'0185    Identificacion Personal
'0390    Impuesto a las Ganancias - Deducciones
'0392    Seguridad Social
'2006    Derechos pendientes de Tiempos
'??????  Conversión de acumulados históricos de la liquidación
Dim rs_Infotipos As New ADODB.Recordset
Dim Infotipo As String

' tengo que leer el string y determinar el infotipo que corresponde

StrSql = "SELECT * FROM infotipos "
StrSql = StrSql & " WHERE activo = -1 "
StrSql = StrSql & " AND descripcion = '" & Infotipo & "'"
OpenRecordset StrSql, rs_Infotipos

'que pasa cuando es base ??????????????????

If Not rs_Infotipos.EOF Then
    Select Case rs_Infotipos!Descripcion
    Case "0000":
        Call Infotipo_0000(rs_Infotipos!infonro)
'    Case "0001":
'        Call Infotipo_0001(rs_Infotipos!infonro)
'    Case "0002":
'        Call Infotipo_0002(rs_Infotipos!infonro)
'    Case "0006":
'        Call Infotipo_0006(rs_Infotipos!infonro)
'    Case "0008":
'        Call Infotipo_0008(rs_Infotipos!infonro)
'    Case "0009":
'        Call Infotipo_0009(rs_Infotipos!infonro)
'    Case "0016":
'        Call Infotipo_0016(rs_Infotipos!infonro)
'    Case "0021":
'        Call Infotipo_0021(rs_Infotipos!infonro)
'    Case "0032":
'        Call Infotipo_0032(rs_Infotipos!infonro)
'    Case "0041":
'        Call Infotipo_0041(rs_Infotipos!infonro)
'    Case "0057":
'        Call Infotipo_0057(rs_Infotipos!infonro)
'    Case "0185":
'        Call Infotipo_0185(rs_Infotipos!infonro)
'    Case "0390":
'        Call Infotipo_0390(rs_Infotipos!infonro)
'    Case "0392":
'        Call Infotipo_0392(rs_Infotipos!infonro)
'    Case "2006":
'        Call Infotipo_2006(rs_Infotipos!infonro)
'    Case "????":
'        Call Infotipo_XXXX(rs_Infotipos!infonro)
    Case Else
    
    End Select
Else
    Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 0) & "Infotipo " & Infotipo & " desconocido "
    Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
End If

If rs_Infotipos.State = adStateOpen Then rs_Infotipos.Close
Set rs_Infotipos = Nothing
End Sub



Public Sub Infotipo_0000(ByVal Infotipo As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: infotipo 0000. Informacion de alta.
' Autor      : FGZ
' Fecha      : 23/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_InfoItems As New ADODB.Recordset

StrSql = "SELECT * FROM info_items "
StrSql = StrSql & " WHERE infonro =" & Infotipo
StrSql = StrSql & " AND activo = -1 "
StrSql = StrSql & " ORDER BY info_items.orden"
OpenRecordset StrSql, rs_InfoItems

Do While Not rs_InfoItems.EOF





    rs_InfoItems.MoveNext
Loop

If rs_InfoItems.State = adStateOpen Then rs_InfoItems.Close
Set rs_InfoItems = Nothing

End Sub
