Attribute VB_Name = "MdlWorkFiles"
Option Explicit

' Variables que contienen los nombres de las tablas temporales de acuerdo a la DB
Global TTempWFDia As String
Global TTempWFDiaLaboral As String
Global TTempWFTurno As String
Global TTempWFEmbudo As String
Global TTempWFJustif As String
Global TTempWFAuspa As String
Global TTempWFAd As String
Global TTempLstaEmple As String
Global TTempWF_tpa As String
Global TTempWF_impproarg As String
Global TTempWF_tt_sitrev As String


'---------------------------------------------------------------
' ESTRUCTURAS DE LOS WF
'---------------------------------------------------------------
' WF_tpa:
'        tipoparam integer
'        Nombre char(30)
'        valor double
'        Fecha Date
        
'wf_impproarg:
'       acunro integer
'       ipacant double
'       ipamonto double
'       tconnro integer
'       desborde double
'       tope_aporte integer (TRUE: APORTE, FALSE: CONTRIBUCIONES

'---------------------------------------------------------------


Public Sub CargarNombresTablasTemporales()
' Setea los nombres de las tablas temporales de acuerdo al tipo de DB

Select Case TipoBD
    Case 1: ' DB2
            TTempWF_tpa = "wf_tpa"
            TTempWF_impproarg = "wf_impproarg"
            TTempWF_tt_sitrev = "wf_tt_sitrev"
    Case 2: ' Informix
            TTempWF_tpa = "wf_tpa"
            TTempWF_impproarg = "wf_impproarg"
            TTempWF_tt_sitrev = "wf_tt_sitrev"
    Case 3: ' SQL Server
            TTempWF_tpa = "#wf_tpa"
            TTempWF_impproarg = "#wf_impproarg"
            TTempWF_tt_sitrev = "#wf_tt_sitrev"
    Case 4: ' Oracle 9
            TTempWF_tpa = "wf_tpa"
            TTempWF_impproarg = "wf_impproarg"
            TTempWF_tt_sitrev = "wf_tt_sitrev"
    End Select
End Sub


Public Sub CreateTempTable(NombreTabla As String)
Dim cadena As String

On Error GoTo CE

Select Case TipoBD
Case 1: ' DB2
        Select Case UCase(NombreTabla)
        Case "WF_TPA", "#WF_TPA":
            cadena = TTempWF_tpa & "(tipoparam integer, ftorden integer, Nombre char(30), valor double, fecha date)"
        Case "WF_IMPPROARG", "#WF_IMPPROARG":
            cadena = TTempWF_impproarg & "(acunro integer, ipacant double, ipamonto double " & _
                     ", tconnro integer, desborde double, tope_aporte integer)"
        Case "WF_TT_SITREV", "#WF_TT_SITREV":
            cadena = TTempWF_tt_sitrev & "(codigo integer, diaini integer)"
        End Select
        
        StrSql = "CREATE TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
        
Case 2: ' Informix
        Select Case UCase(NombreTabla)
        Case "WF_TPA", "#WF_TPA":
            cadena = TTempWF_tpa & "(tipoparam integer, ftorden integer, Nombre char(30), valor decimal (15,4), fecha date)"
        Case "WF_IMPPROARG", "#WF_IMPPROARG":
            cadena = TTempWF_impproarg & "(acunro integer, ipacant decimal (15,4), ipamonto decimal (15,4)" & _
                     ", tconnro integer, desborde decimal (15,4), tope_aporte integer)"
        Case "WF_TT_SITREV", "#WF_TT_SITREV":
            cadena = TTempWF_tt_sitrev & "(codigo integer, diaini integer)"
        End Select
        
        StrSql = "CREATE TEMP TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
        
Case 3: ' SQL Server
        Select Case UCase(NombreTabla)
        Case "WF_TPA", "#WF_TPA":
            cadena = TTempWF_tpa & "(tipoparam integer, ftorden integer, Nombre char(30), valor decimal (19,4), fecha datetime)"
        Case "WF_IMPPROARG", "#WF_IMPPROARG":
            cadena = TTempWF_impproarg & "(acunro integer, ipacant decimal (19,4), ipamonto decimal (19,4)" & _
                     ", tconnro integer, desborde decimal (19,4), tope_aporte integer)"
        Case "WF_TT_SITREV", "#WF_TT_SITREV":
            cadena = TTempWF_tt_sitrev & "(codigo integer, diaini integer)"
        End Select
        
        StrSql = "CREATE TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
Case 4: ' Oracle 9
        Select Case UCase(NombreTabla)
        Case "WF_TPA", "#WF_TPA":
            cadena = TTempWF_tpa & "(tipoparam double precision, ftorden double precision,Nombre char(30), valor FLOAT(19), fecha date)"
        Case "WF_IMPPROARG", "#WF_IMPPROARG":
            cadena = TTempWF_impproarg & "(acunro double precision, ipacant FLOAT(19), ipamonto FLOAT(19)" & _
                     ", tconnro double precision, desborde FLOAT(19), tope_aporte double precision)"
        Case "WF_TT_SITREV", "#WF_TT_SITREV":
            cadena = TTempWF_tt_sitrev & "(codigo integer, diaini integer)"
        End Select
        
        StrSql = "CREATE GLOBAL TEMPORARY TABLE " & cadena & " ON COMMIT PRESERVE ROWS"
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
End Select

CE:
    If TipoBD = 4 Then
        Call BorrarTempTable(NombreTabla)
    Else
        Call BorrarTempTable(NombreTabla)
        Call CreateTempTable(NombreTabla)
    End If

End Sub


Public Sub BorrarTempTable(NombreTabla As String)
Dim Tabla As String

Tabla = NombreTabla
        
        
If TipoBD = 4 Then
    StrSql = "TRUNCATE TABLE " & Tabla
    objConn.Execute StrSql, , adExecuteNoRecords
Else
    StrSql = "DROP TABLE " & Tabla
    objConn.Execute StrSql, , adExecuteNoRecords
End If

End Sub


Public Sub LimpiarTempTable(NombreTabla As String)
Dim Tabla As String

Tabla = NombreTabla

StrSql = "DELETE FROM " & Tabla
objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Limpiar_wf_impproarg(ByVal NombreTabla As String)
Dim rs_acumulador As New ADODB.Recordset
Dim i As Integer

'Call BorrarTempTable(TTempWF_impproarg)
Call CreateTempTable(TTempWF_impproarg)

'strsql = "SELECT * FROM acumulador WHERE acutopea = -1"
'OpenRecordset strsql, rs_acumulador

'Do While Not rs_acumulador.EOF
'
'    For i = 1 To 5 ' uno por cada tipo de concepto
'        Call insertar_wf_impproarg(rs_acumulador!acunro, CBool(rs_acumulador!acuimponible), CBool(rs_acumulador!acuimpcont), i)
'    Next i
'    rs_acumulador.MoveNext
'Loop

End Sub


Public Sub insertar_wf_impproarg(ByVal acunro As Long, ByVal imponible As Boolean, ByVal impcont As Boolean, ByVal tconnro As Long)
Dim Valor As String

If imponible Then
    Valor = "-1"
End If
If impcont Then
    Valor = "0"
End If

StrSql = "INSERT INTO " & TTempWF_impproarg & " (acunro , ipacant, ipamonto, tconnro, tope_aporte) VALUES (" & _
                     acunro & ",0,0," & tconnro & "," & Valor & ")"
objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub insertar_wf_tpa(ByVal tpanro As Long, ByVal Orden As Long, ByVal Nom As String, ByVal Valor As Single, ByVal Fecha As Date)
Dim nombre As String

nombre = "par" & Format(tpanro, "00000")
StrSql = "INSERT INTO " & TTempWF_tpa & " (tipoparam,ftorden,Nombre, valor, fecha) VALUES (" & _
        tpanro & "," & Orden & ",'" & nombre & "'," & Valor & "," & ConvFecha(Fecha) & ")"
objConn.Execute StrSql, , adExecuteNoRecords

End Sub
