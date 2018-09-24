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
Global TTempWF_EscalaUTM As String
Global TTempWF_Retroactivo As String



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
            TTempWF_EscalaUTM = "wf_escalautm"
            TTempWF_Retroactivo = "wf_retroactivo"
    Case 2: ' Informix
            TTempWF_tpa = "wf_tpa"
            TTempWF_impproarg = "wf_impproarg"
            TTempWF_tt_sitrev = "wf_tt_sitrev"
            TTempWF_EscalaUTM = "wf_escalautm"
            TTempWF_Retroactivo = "wf_retroactivo"
    Case 3: ' SQL Server
            TTempWF_tpa = "#wf_tpa"
            TTempWF_impproarg = "#wf_impproarg"
            TTempWF_tt_sitrev = "#wf_tt_sitrev"
            TTempWF_EscalaUTM = "#wf_escalautm"
            TTempWF_Retroactivo = "#wf_retroactivo"
    Case 4: ' Oracle 9
            TTempWF_tpa = "wf_tpa"
            TTempWF_impproarg = "wf_impproarg"
            TTempWF_tt_sitrev = "wf_tt_sitrev"
            TTempWF_EscalaUTM = "wf_escalautm"
            TTempWF_Retroactivo = "wf_retroactivo"
    End Select
End Sub


Public Sub CreateTempTable(NombreTabla As String)
Dim Cadena As String

On Error GoTo CE

Select Case TipoBD
Case 1: ' DB2
        Select Case UCase(NombreTabla)
        Case "WF_TPA", "#WF_TPA":
            Cadena = TTempWF_tpa & "(tipoparam integer, ftorden integer, Nombre char(30), valor double, fecha date)"
        Case "WF_IMPPROARG", "#WF_IMPPROARG":
            Cadena = TTempWF_impproarg & "(acunro integer, ipacant double, ipamonto double " & _
                     ", tconnro integer, desborde double, tope_aporte integer)"
        Case "WF_TT_SITREV", "#WF_TT_SITREV":
            Cadena = TTempWF_tt_sitrev & "(codigo integer, diaini integer)"
'        Case "WF_BAE", "#WF_BAE":
'            cadena = TTempWF_BAE & "(ternro integer, codigo_parte char(3), Cantidad double, Horas double, codigo_bae char(3), penal double, mes integer, anio intege)"
        End Select
        
        StrSql = "CREATE TABLE " & Cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
        
Case 2: ' Informix
        Select Case UCase(NombreTabla)
        Case "WF_TPA", "#WF_TPA":
            Cadena = TTempWF_tpa & "(tipoparam integer, ftorden integer, Nombre char(30), valor decimal (15,4), fecha date)"
        Case "WF_IMPPROARG", "#WF_IMPPROARG":
            Cadena = TTempWF_impproarg & "(acunro integer, ipacant decimal (15,4), ipamonto decimal (15,4)" & _
                     ", tconnro integer, desborde decimal (15,4), tope_aporte integer)"
        Case "WF_TT_SITREV", "#WF_TT_SITREV":
            Cadena = TTempWF_tt_sitrev & "(codigo integer, diaini integer)"
'        Case "WF_BAE", "#WF_BAE":
'            cadena = TTempWF_BAE & "(ternro integer, codigo_parte char(3), Cantidad decimal (15,4), Horas decimal (15,4), codigo_bae char(3), penal decimal (15,4), mes integer, anio intege)"
        End Select
        
        StrSql = "CREATE TEMP TABLE " & Cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
        
Case 3: ' SQL Server
        Select Case UCase(NombreTabla)
        Case "WF_TPA", "#WF_TPA":
            Cadena = TTempWF_tpa & "(tipoparam integer, ftorden integer, Nombre char(30), valor decimal (19,4), fecha datetime)"
        Case "WF_IMPPROARG", "#WF_IMPPROARG":
            Cadena = TTempWF_impproarg & "(acunro integer, ipacant decimal (19,4), ipamonto decimal (19,4)" & _
                     ", tconnro integer, desborde decimal (19,4), tope_aporte integer)"
        Case "WF_TT_SITREV", "#WF_TT_SITREV":
            Cadena = TTempWF_tt_sitrev & "(codigo integer, diaini integer)"
        Case "WF_EscalaUTM", "#WF_ESCALAUTM":
            Cadena = TTempWF_EscalaUTM & "(desde decimal (19,4), hasta decimal (19,4), factor decimal (19,4), rebaja decimal (19,4))"
        Case "WF_RETROACTIVO", "#WF_RETROACTIVO":
            Cadena = TTempWF_Retroactivo & "(monto decimal (19,4), anio integer, mes integer, concnro integer)"
        End Select
        
        StrSql = "CREATE TABLE " & Cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
Case 4: ' Oracle 9
        Select Case UCase(NombreTabla)
        Case "WF_TPA", "#WF_TPA":
            Cadena = TTempWF_tpa & "(tipoparam double precision, ftorden double precision,Nombre char(30), valor FLOAT(63), fecha date)"
        Case "WF_IMPPROARG", "#WF_IMPPROARG":
            Cadena = TTempWF_impproarg & "(acunro double precision, ipacant FLOAT(63), ipamonto FLOAT(63)" & _
                     ", tconnro double precision, desborde FLOAT(63), tope_aporte double precision)"
        Case "WF_TT_SITREV", "#WF_TT_SITREV":
            Cadena = TTempWF_tt_sitrev & "(codigo integer, diaini integer)"
        Case "WF_EscalaUTM", "#WF_ESCALAUTM":
            Cadena = TTempWF_EscalaUTM & "(desde FLOAT(19), hasta FLOAT(19), factor FLOAT(19), rebaja FLOAT(19))"
        Case "WF_RETROACTIVO", "#WF_RETROACTIVO":
            Cadena = TTempWF_Retroactivo & "(monto FLOAT(63), anio double precision, mes double precision, concnro double precision)"
        End Select
        
        StrSql = "CREATE GLOBAL TEMPORARY TABLE " & Cadena & " ON COMMIT PRESERVE ROWS"
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
Dim I As Integer

Call CreateTempTable(TTempWF_impproarg)
End Sub


Public Sub insertar_wf_impproarg(ByVal acuNro As Long, ByVal imponible As Boolean, ByVal impcont As Boolean, ByVal tconnro As Long)
Dim Valor As String

If imponible Then
    Valor = "-1"
End If
If impcont Then
    Valor = "0"
End If

StrSql = "INSERT INTO " & TTempWF_impproarg & " (acunro , ipacant, ipamonto, tconnro, tope_aporte) VALUES (" & _
                     acuNro & ",0,0," & tconnro & "," & Valor & ")"
objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub insertar_wf_tpa(ByVal tpanro As Long, ByVal Orden As Long, ByVal Nom As String, ByVal Valor As Double, ByVal Fecha As Date)
Dim nombre As String

nombre = "par" & Format(tpanro, "00000")
StrSql = "INSERT INTO " & TTempWF_tpa & " (tipoparam,ftorden,Nombre, valor, fecha) VALUES (" & _
        tpanro & "," & Orden & ",'" & nombre & "'," & Valor & "," & ConvFecha(Fecha) & ")"
objConn.Execute StrSql, , adExecuteNoRecords

End Sub
Public Sub insertar_wf_escalautm(ByVal Mes As Long, ByVal Anio As Long, ByVal utmPliq As Double)

' ---------------------------------------------------------------------------------------------
' Descripcion: Carga la tabla Escala con UTM (CHILE).
' Autor      : Maximiliano Breglia
' Fecha      : 03/12/2006
' Ultima Mod.: 20/11/2007 - Martin Ferraro
' Descripcion: Se paso por parametro la anio y mes del periodo y el utm para utilizarlo en el
'              recalculo del impuesto unico. Primero borro la tabla
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_escala_utm As New ADODB.Recordset
'Dim fecha_periodo As Date
Dim Max_escala As Integer
Dim v_desde As Double
Dim v_hasta As Double
Dim v_factor As Double
Dim v_rebaja As Double


    'fecha_periodo = buliq_periodo!pliqdesde
    
    StrSql = "DELETE FROM " & TTempWF_EscalaUTM
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = "SELECT * FROM escala "
    'StrSql = StrSql & " WHERE escfecha <= " & ConvFecha(fecha_periodo)
    StrSql = StrSql & " WHERE escano = " & Anio
    StrSql = StrSql & " AND escmes = " & Mes
    StrSql = StrSql & " Order BY escfecha desc"
    OpenRecordset StrSql, rs_escala_utm
            
    Do While Not rs_escala_utm.EOF
        'v_desde = rs_escala_utm!escinf * buliq_periodo!pliqutm
        v_desde = rs_escala_utm!escinf * utmPliq
        'v_hasta = rs_escala_utm!escsup * buliq_periodo!pliqutm
        v_hasta = rs_escala_utm!escsup * utmPliq
        v_factor = rs_escala_utm!escporexe
        'v_rebaja = rs_escala_utm!esccuota * buliq_periodo!pliqutm
        v_rebaja = rs_escala_utm!esccuota * utmPliq
    
        StrSql = "INSERT INTO " & TTempWF_EscalaUTM & " (desde , hasta, factor, rebaja) VALUES (" & _
                 v_desde & "," & v_hasta & "," & v_factor & "," & v_rebaja & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        rs_escala_utm.MoveNext
    Loop
    
    If rs_escala_utm.State = adStateOpen Then rs_escala_utm.Close
    Set rs_escala_utm = Nothing

End Sub

Public Sub insertar_wf_retroactivo(ByVal Monto As Double, ByVal Anio As Long, ByVal Mes As Long, ByVal ConcNro As Long)

StrSql = "INSERT INTO " & TTempWF_Retroactivo & " (monto ,anio, mes, concnro) VALUES (" & _
                     Monto & "," & Anio & "," & Mes & "," & ConcNro & ")"
objConn.Execute StrSql, , adExecuteNoRecords

End Sub


