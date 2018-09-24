Attribute VB_Name = "MdlWorkFiles"
Option Explicit

'--------------------------------------------------------------------------------------------------------
'ANSI data type     Oracle          MySql       PostGreSQL                  Most Portable
'--------------------------------------------------------------------------------------------------------
'integer            NUMBER(38)      integer(11)     integer                         integer
'smallint           NUMBER(38)      smallint(6)     smallint                        smallint
'tinyint                *           tinyint(4)          *                           numeric(4,0)
'numeric(p,s)       NUMBER(p,s)     decimal(p,s)    numeric(p,s)                    numeric(p,s)
'varchar(n)         VARCHAR2(n)     varchar(n)      character varying(n)            varchar(n)
'char(n)            CHAR(n)         varchar(n)      character(n)                    char(n)
'datetime           DATE            datetime        timestamp without time zone     have to autodetect
'float              FLOAT(126)      float           double precision                float
'real               FLOAT(63)       double          real                            real
'--------------------------------------------------------------------------------------------------------





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
Global TTempWF_BAE As String

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
            'TTempWF_BAE = "wf_bae"
    Case 2: ' Informix
            TTempWF_tpa = "wf_tpa"
            TTempWF_impproarg = "wf_impproarg"
            TTempWF_tt_sitrev = "wf_tt_sitrev"
            'TTempWF_BAE = "wf_bae"
    Case 3: ' SQL Server
            TTempWF_tpa = "#wf_tpa"
            TTempWF_impproarg = "#wf_impproarg"
            TTempWF_tt_sitrev = "#wf_tt_sitrev"
            'TTempWF_BAE = "#wf_bae"
    Case 4: ' Oracle 9
            TTempWF_tpa = "wf_tpa"
            TTempWF_impproarg = "wf_impproarg"
            TTempWF_tt_sitrev = "wf_tt_sitrev"
            'TTempWF_BAE = "wf_bae"
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
'        Case "WF_BAE", "#WF_BAE":
'            cadena = TTempWF_BAE & "(ternro integer, codigo_parte char(3), Cantidad double, Horas double, codigo_bae char(3), penal double, mes integer, anio intege)"
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
'        Case "WF_BAE", "#WF_BAE":
'            cadena = TTempWF_BAE & "(ternro integer, codigo_parte char(3), Cantidad decimal (15,4), Horas decimal (15,4), codigo_bae char(3), penal decimal (15,4), mes integer, anio intege)"
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
'        Case "WF_BAE", "#WF_BAE":
'            cadena = TTempWF_BAE & "(ternro integer, codigo_parte varchar(3), Cantidad decimal (19,4), Horas decimal (19,4), codigo_bae varchar(3), penal decimal (15,4), mes integer, anio intege)"
        End Select
        
        StrSql = "CREATE TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
Case 4: ' Oracle 9
        Select Case UCase(NombreTabla)
        Case "WF_TPA", "#WF_TPA":
            cadena = TTempWF_tpa & "(tipoparam double precision, ftorden double precision,Nombre char(30), valor FLOAT(63), fecha date)"
        Case "WF_IMPPROARG", "#WF_IMPPROARG":
            cadena = TTempWF_impproarg & "(acunro double precision, ipacant FLOAT(63), ipamonto FLOAT(63)" & _
                     ", tconnro double precision, desborde FLOAT(63), tope_aporte double precision)"
        Case "WF_TT_SITREV", "#WF_TT_SITREV":
            cadena = TTempWF_tt_sitrev & "(codigo integer, diaini integer)"
'        Case "WF_BAE", "#WF_BAE":
'            cadena = TTempWF_BAE & "(ternro double precision, codigo_parte char(3), Cantidad FLOAT(19), Horas FLOAT(19), codigo_bae char(3), penal FLOAT(19), mes integer, anio intege)"
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

Call CreateTempTable(TTempWF_impproarg)
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


Public Sub insertar_wf_tpa(ByVal tpanro As Long, ByVal Orden As Long, ByVal Nom As String, ByVal Valor As Double, ByVal Fecha As Date)
Dim nombre As String

nombre = "par" & Format(tpanro, "00000")
StrSql = "INSERT INTO " & TTempWF_tpa & " (tipoparam,ftorden,Nombre, valor, fecha) VALUES (" & _
        tpanro & "," & Orden & ",'" & nombre & "'," & Valor & "," & ConvFecha(Fecha) & ")"
objConn.Execute StrSql, , adExecuteNoRecords

End Sub
