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
' Descripcion: IMPORTACION DE Total de Cantidad de BULTOS  a  RH Pro
'              IDEA : importado un desglose de Acumulado Diario de GTI para un T.Hora espec¡fico.
'              Luego lee el archivo y crea los Desglose AD de GTI, siempre pisa (asumen un reg. x empleado x convinatoria)
'              Genera un log de error en el mismo TMP
'              configuraci¢n de CTTES para la IMPORTACION
' Autor      : Alvaro Bayon
' Fecha      : 11/11/2004
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
    
'Configuración de conceptos para las novedades
Dim id_concepto_pera As Long
Dim id_concepto_manzana As Long
Dim id_concepto_carozo As Long
Dim concepto_pera As String
Dim concepto_manzana As String
Dim concepto_carozo As String

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

Dim fs1
Dim Flog1
Dim txtArchivoNov



'Conccod de productos
concepto_pera = "100"
concepto_manzana = "120"
concepto_carozo = "140"

'Obtengo el código del concepto pera
StrSql = "SELECT * FROM concepto WHERE " & _
" conccod = '" & concepto_pera & "'"
OpenRecordset StrSql, rs_gti_achdiario
If Not rs_gti_achdiario.EOF Then
    id_concepto_pera = rs_gti_achdiario!concnro
    
End If

'Obtengo el código del concepto manzana
StrSql = "SELECT * FROM concepto WHERE " & _
" conccod = '" & concepto_manzana & "'"
OpenRecordset StrSql, rs_gti_achdiario
If Not rs_gti_achdiario.EOF Then
    id_concepto_manzana = rs_gti_achdiario!concnro
End If

'Obtengo el código del concepto carozo
StrSql = "SELECT * FROM concepto WHERE " & _
" conccod = '" & concepto_carozo & "'"
OpenRecordset StrSql, rs_gti_achdiario
If Not rs_gti_achdiario.EOF Then
    id_concepto_carozo = rs_gti_achdiario!concnro
End If


'-----------------------------------------------------
'borrado de novedades

'Primero hago un backup de las novedades que voy a borrar
Set fs1 = CreateObject("Scripting.FileSystemObject")
txtArchivoNov = PathFLog & "novemp" & CStr(Format(Date, "yyyymmdd")) & Format(Time, "hhmm") & ".txt"
Set Flog1 = fs.CreateTextFile(txtArchivoNov, True)

StrSql = "SELECT *  FROM novemp WHERE " & _
" concnro = " & id_concepto_manzana & _
" OR concnro = " & id_concepto_pera & _
" OR concnro = " & id_concepto_carozo
OpenRecordset StrSql, rs_gti_achdiario
Do While Not rs_gti_achdiario.EOF
    Flog1.Write rs_gti_achdiario!concnro & "," & rs_gti_achdiario!tpanro & ","
    Flog1.Write rs_gti_achdiario!Empleado & "," & rs_gti_achdiario!nevalor & ","
    Flog1.Write rs_gti_achdiario!nevigencia & "," & rs_gti_achdiario!nedesde & ","
    Flog1.Write rs_gti_achdiario!nehasta & "," & rs_gti_achdiario!neretro & ","
    Flog1.Write rs_gti_achdiario!nepliqdesde & "," & rs_gti_achdiario!nepliqhasta & ","
    Flog1.Write rs_gti_achdiario!pronro & "," & rs_gti_achdiario!nenro & ","
    Flog1.Writeline
       
    rs_gti_achdiario.MoveNext
Loop

'Borro las novedades de peras, manzanas o carozo
StrSql = "DELETE FROM novemp WHERE " & _
" concnro = " & id_concepto_manzana & _
" OR concnro = " & id_concepto_pera & _
" OR concnro = " & id_concepto_carozo

objConn.Execute StrSql, , adExecuteNoRecords



'---------------------------------------------------------


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
        Flog.Writeline "Empleado Inexistente: " & legajo
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

'   Hasta aquí el código anterior
'=============================================================

If empaque = 1 Then
    Select Case producto
    Case 1:      'PERAS
        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_pera & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 163"
        OpenRecordset StrSql, rs_gti_achdiario
        If rs_gti_achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_pera & _
             "," & cant_bultos & _
             ", 163" & _
             " )"
        End If
    Case 2:          'MANZANAS
        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_manzana & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 163"
        OpenRecordset StrSql, rs_gti_achdiario
        If rs_gti_achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_manzana & _
             "," & cant_bultos & _
             ", 163" & _
             " )"
        End If
        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_manzana & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 51"
        OpenRecordset StrSql, rs_gti_achdiario
        If rs_gti_achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_manzana & _
             "," & cant_bultos & _
             ", 51" & _
             " )"
        End If
    
    Case 3:          'CAROZO
        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_carozo & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 163"
        OpenRecordset StrSql, rs_gti_achdiario
        If rs_gti_achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_carozo & _
             "," & cant_bultos & _
             ", 163" & _
             " )"
        End If

        'Busco la novedad. Si no existe la creo
        StrSql = "SELECT * FROM novemp WHERE " & _
         " concnro = " & id_concepto_carozo & _
         " AND empleado = " & rs_Empleado!ternro & _
         " AND tpanro = 51"
        OpenRecordset StrSql, rs_gti_achdiario
        If rs_gti_achdiario.EOF Then
            StrSql = "INSERT INTO novemp (" & _
             "empleado,concnro,nevalor,tpanro" & _
             ") VALUES (" & rs_Empleado!ternro & _
             "," & id_concepto_carozo & _
             "," & cant_bultos & _
             ", 51" & _
             " )"
        End If
    End Select
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

