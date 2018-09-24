Attribute VB_Name = "MdlInterface"
Option Explicit

Global NombreArchivo As String
Global NroLinea As Long
Global f
Global HuboError As Boolean
Global PisaNovedad As Boolean
Global Path
Global NArchivo
Global Separador As String
Global UsaEncabezado As Boolean

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: IMPORTACION DE Detalle de Cantidad de BULTOS  a  RH Pro
'              IDEA : importado un desglose de Acumulado Diario de GTI para un T.Hora espec¡fico.
'              Luego lee el archivo y crea los Desglose AD de GTI, siempre pisa (asumen un reg. x empleado x convinatoria)
'              Genera un log de error en el mismo TMP
'              configuraci¢n de CTTES para la IMPORTACION
' Autor      : FGZ
' Fecha      : 10/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset


' carga las configuraciones basicas, formato de fecha, string de conexion,
' tipo de BD y ubicacion del archivo de log
Call CargarConfiguracionesBasicas
    
'Abro la conexion
    OpenConnection strconexion, objConn
    
    Nombre_Arch = PathFLog & "det_bultos.err"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Call ComenzarTransferencia
    
    Flog.Close
End Sub


Private Sub Leer(ByVal NombreArchivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String

    If App.PrevInstance Then Exit Sub

    'Espero hasta que se crea el archivo de Novedades
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then Err.Number = 1
    Loop
    On Error GoTo 0
   
   'Abro el archivo de Novedades
    On Error GoTo CE
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    
    Do While Not f.AtEndOfStream
        strLinea = f.ReadLine
        NroLinea = NroLinea + 1
        If NroLinea = 1 And UsaEncabezado Then
            strLinea = f.ReadLine
            NroLinea = NroLinea + 1
        End If
        If Trim(strLinea) <> "" Then
            Call Insertar(strLinea)
        End If
    Loop
    
fin:
    Exit Sub
    
CE:
    GoTo fin
End Sub


Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer




If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Pisa o no las novedades
        pos1 = 1
        pos2 = Len(parametros)
        PisaNovedad = CBool(Mid(parametros, pos1, pos2))

'        Pos1 = 1
'        Pos2 = InStr(Pos1, parametros, ".") - 1
'        PisaNovedad = Mid(parametros, Pos1, Pos2)
        
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        Mantener_Liq = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
'        Pos1 = Pos2 + 2
'        Pos2 = Len(parametros)
'        HACE_TRAZA = CBool(Mid(parametros, Pos1, Pos2 - Pos1 + 1))
        
    End If
End If

End Sub



Public Sub Insertar(ByVal strLinea As String)
Dim pos1 As Integer
Dim pos2 As Integer
    
Dim id_producto_peras As Long '   AS INT INITIAL 1.
Dim id_producto_manzana As Long 'AS INT INITIAL 2.
Dim id_producto_carozo As Long 'AS INT INITIAL 3.  /* incluye: CIRUELA, DURAZNO Y PELONES */
Dim id_th_bultos As Long '        AS INT INITIAL 51  . /* THora de BULTOS */.

Dim cant_bultos_txt As String
Dim monto_bultos_txt As String
Dim cant_bultos   As Single
Dim monto_bultos As Single
Dim primera_parte As String
Dim empaque As Integer
Dim legajo As Long
Dim fecha_desde As String
Dim fecha_hasta As String
Dim producto_txt As String
Dim producto As Integer
Dim fecha_prod As Date

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_gti_achdiario As New ADODB.Recordset

id_producto_peras = 1
id_producto_manzana = 2
id_producto_carozo = 3
id_th_bultos = 51

'Levanto los datos que vienen en la linea
'strLinea1
empaque = CInt(Mid(strLinea, 1, 1))
legajo = CLng(Mid(strLinea, 2, 6))
fecha_desde = Mid(strLinea, 8, 10)
fecha_hasta = Mid(strLinea, 18, 10)
producto_txt = Trim(Mid(strLinea, 28, 10))

'strLinea2
'fecha_prod = CDate(Mid(strLinea, 1, 10))

'strLinea3
cant_bultos = CSng(Mid(strLinea, 37, 7))
If Len(strLinea) > 0 Then
    cant_bultos = cant_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
Else
    cant_bultos = cant_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
End If

'strLinea4
monto_bultos = CSng(Mid(strLinea, 44, 7))
If Len(strLinea) > 0 Then
    monto_bultos = monto_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
Else
    monto_bultos = monto_bultos + CInt(Mid(strLinea, Len(strLinea) - 1, 2) / 100)
End If


' ====================================================================
' control de errores
    StrSql = "SELECT * FROM empleado where empleg = " & legajo
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.writeline "Empleado Inexistente: " & legajo
        Exit Sub
    End If
    
   
    Select Case producto_txt
    Case "PERAS":
        producto = 1 'id-th-bultos = id-th-bultos-pera
    Case "MANZANAS":
        producto = 2 'id-th-bultos = id-th-bultos-manzana
    Case "DURAZNOS":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    Case "PELONES":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    Case "CIRUELAS":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    End Select

'Busco la Sucursal
    StrSql = " SELECT estrnro FROM his_estructura " & _
             " WHERE ternro = " & rs_Empleado!ternro & " AND " & _
             " tenro = 1 AND estrnro = " & empaque & " AND " & _
             " (htetdesde <= " & ConvFecha(fecha_hasta) & ") AND " & _
             " ((" & ConvFecha(fecha_hasta) & " <= htethasta) or (htethasta is null))"
    OpenRecordset StrSql, rs_Estructura

    If Not rs_Estructura.EOF Then
        StrSql = " SELECT * FROM sucursal " & _
                 " WHERE estrnro =" & rs_Estructura!estrnro
        OpenRecordset StrSql, rs_Sucursal
        
        If rs_Sucursal.EOF Then
            Flog.writeline "Empaque Inexistente: " & empaque
            Exit Sub
        Else
        
        End If
    Else
        Flog.writeline "Empaque Inexistente: " & empaque
        Exit Sub
    End If


'=============================================================
' Inserto
StrSql = "SELECT * FROM gti_achdiario WHERE " & _
         " ternro = " & rs_Empleado!ternro & _
         " AND thnro = " & id_th_bultos & _
         " AND achdfecha = " & ConvFecha(fecha_prod) & _
         " AND catnro = " & rs_Empleado!catnro & _
         " AND puenro = " & producto & _
         " AND sucursal = " & rs_Sucursal!ternro
OpenRecordset StrSql, rs_gti_achdiario

If rs_gti_achdiario.EOF Then
    StrSql = "INSERT INTO gti_achdiario (" & _
             "ternro,thnro,achdfecha,catnro,puenro,sucursal" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_th_bultos & _
             "," & ConvFecha(fecha_prod) & _
             "," & rs_Empleado!catnro & _
             "," & producto & _
             "," & rs_Sucursal!ternro & _
             " )"
    objConn.Execute StrSql, , adExecuteNoRecords
End If

End Sub



Public Sub Insertar_multiple(ByVal strLinea1 As String, ByVal strLinea2 As String, ByVal strLinea3 As String, ByVal strLinea4 As String)
Dim pos1 As Integer
Dim pos2 As Integer
    
Dim id_producto_peras As Long '   AS INT INITIAL 1.
Dim id_producto_manzana As Long 'AS INT INITIAL 2.
Dim id_producto_carozo As Long 'AS INT INITIAL 3.  /* incluye: CIRUELA, DURAZNO Y PELONES */
Dim id_th_bultos As Long '        AS INT INITIAL 51  . /* THora de BULTOS */.

Dim cant_bultos_txt As String
Dim monto_bultos_txt As String
Dim cant_bultos   As Single
Dim monto_bultos As Single
Dim primera_parte As String
Dim empaque As Integer
Dim legajo As Integer
Dim fecha_desde As String
Dim fecha_hasta As String
Dim producto_txt As String
Dim producto As Integer
Dim fecha_prod As Date

Dim rs_Empleado As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_gti_achdiario As New ADODB.Recordset

id_producto_peras = 1
id_producto_manzana = 2
id_producto_carozo = 3
id_th_bultos = 51

'Levanto los datos que vienen en la linea
'strLinea1
empaque = CInt(Mid(strLinea1, 1, 1))
legajo = Mid(strLinea1, 2, 6)
fecha_desde = Mid(strLinea1, 8, 10)
fecha_hasta = Mid(strLinea1, 18, 10)
producto_txt = Mid(strLinea1, 28, 10)

'strLinea2
fecha_prod = CDate(Mid(strLinea2, 1, 10))

'strLinea3
cant_bultos = CSng(Mid(strLinea3, 1, Len(strLinea3) - 2))
If Len(strLinea3) > 0 Then
    cant_bultos = cant_bultos + CInt(Mid(strLinea3, Len(strLinea3) - 1, 2) / 100)
Else
    cant_bultos = cant_bultos + CInt(Mid(strLinea3, Len(strLinea3) - 1, 2) / 100)
End If

'strLinea4
monto_bultos = CSng(Mid(strLinea4, 1, Len(strLinea4) - 2))
If Len(strLinea4) > 0 Then
    monto_bultos = monto_bultos + CInt(Mid(strLinea4, Len(strLinea4) - 1, 2) / 100)
Else
    monto_bultos = monto_bultos + CInt(Mid(strLinea4, Len(strLinea4) - 1, 2) / 100)
End If


' ====================================================================
' control de errores
    StrSql = "SELECT * FROM empleado where empleg = " & legajo
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.writeline "Empleado Inexistente: " & legajo
        Exit Sub
    End If
    
   
    Select Case producto_txt
    Case "PERAS":
        producto = 1 'id-th-bultos = id-th-bultos-pera
    Case "MANZANAS":
        producto = 2 'id-th-bultos = id-th-bultos-manzana
    Case "DURAZNOS":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    Case "PELONES":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    Case "CIRUELAS":
        producto = 3 'id-th-bultos = id-th-bultos-carozo
    End Select

'Busco la Sucursal
    StrSql = " SELECT estrnro FROM his_estructura " & _
             " WHERE ternro = " & rs_Empleado!ternro & " AND " & _
             " tenro = 1 AND estrnro = " & empaque & " AND " & _
             " (htetdesde <= " & ConvFecha(fecha_hasta) & ") AND " & _
             " ((" & ConvFecha(fecha_hasta) & " <= htethasta) or (htethasta is null))"
    OpenRecordset StrSql, rs_Estructura

    If Not rs_Estructura.EOF Then
        StrSql = " SELECT * FROM sucursal " & _
                 " WHERE estrnro =" & rs_Estructura!estrnro
        OpenRecordset StrSql, rs_Sucursal
        
        If rs_Sucursal.EOF Then
            Flog.writeline "Empaque Inexistente: " & empaque
            Exit Sub
        Else
        
        End If
    Else
        Flog.writeline "Empaque Inexistente: " & empaque
        Exit Sub
    End If


'=============================================================
' Inserto
StrSql = "SELECT * FROM gti_achdiario WHERE " & _
         " ternro = " & rs_Empleado!ternro & _
         " AND thnro = " & id_th_bultos & _
         " AND achdfecha = " & ConvFecha(fecha_prod) & _
         " AND catnro = " & rs_Empleado!catnro & _
         " AND puenro = " & producto & _
         " AND sucursal = " & rs_Sucursal!ternro
OpenRecordset StrSql, rs_gti_achdiario

If rs_gti_achdiario.EOF Then
    StrSql = "INSERT INTO gti_achdiario (" & _
             "ternro,thnro,achdfecha,catnro,puenro,sucursal" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_th_bultos & _
             "," & ConvFecha(fecha_prod) & _
             "," & rs_Empleado!catnro & _
             "," & producto & _
             "," & rs_Sucursal!ternro & _
             " )"
    objConn.Execute StrSql, , adExecuteNoRecords
End If

End Sub



Public Sub ComenzarTransferencia()
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder

    'Setear acá el path del archivo a levantar y separador de parametros
    Separador = ";"
    Directorio = "c:\logs\otros"
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Path = Directorio
    
    Dim fc, F1, s2
    Set Folder = fs.GetFolder(Directorio)
    Set CArchivos = Folder.Files
    
    For Each archivo In CArchivos
        If UCase(Right(archivo.Name, 4)) = ".TXT" Then
            NArchivo = archivo.Name
            Call Leer(Directorio & "\" & archivo.Name)
        End If
    Next
    
End Sub

