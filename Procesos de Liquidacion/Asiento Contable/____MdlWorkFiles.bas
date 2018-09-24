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
Global TTempWFConcepto_Dist As String


Public Sub CargarNombresTablasTemporales()
' Setea los nombres de las tablas temporales de acuerdo al tipo de DB

Select Case TipoBD
    Case 1: ' DB2
            TTempWFConcepto_Dist = "wf_concepto_dist"
    Case 2: ' Informix
            TTempWFConcepto_Dist = "wf_concepto_dist"
    Case 3: ' SQL Server
            TTempWFConcepto_Dist = "#wf_concepto_dist"
    Case 4: ' Oracle 9
            TTempWFConcepto_Dist = "wf_concepto_dist"
    End Select
End Sub


Public Sub CreateTempTable(NombreTabla As String)
Dim cadena As String

On Error GoTo CE

Select Case TipoBD
Case 1: ' DB2
        Select Case UCase(NombreTabla)
        Case "WF_CONCEPTO_DIST", "#WF_CONCEPTO_DIST":
            cadena = TTempWFConcepto_Dist & "(ternro integer,concnro integer,pronro integer,masinro integer,tenro integer,estrnro integer,tenro2 integer,estrnro2 integer,tenro3 integer,estrnro3 integer,porcentaje double,monto double)"
        End Select
        
        StrSql = "CREATE TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
        
Case 2: ' Informix
        Select Case UCase(NombreTabla)
        Case "WF_CONCEPTO_DIST", "#WF_CONCEPTO_DIST":
            cadena = TTempWFConcepto_Dist & "(ternro integer,concnro integer,pronro integer,masinro integer,tenro integer,estrnro integer,tenro2 integer,estrnro2 integer,tenro3 integer,estrnro3 integer,porcentaje decimal (15,4),monto decimal (15,4))"
        End Select
        
        StrSql = "CREATE TEMP TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
        
Case 3: ' SQL Server
        Select Case UCase(NombreTabla)
        Case "WF_CONCEPTO_DIST", "#WF_CONCEPTO_DIST":
            cadena = TTempWFConcepto_Dist & "(ternro integer,concnro integer,pronro integer,masinro integer,tenro integer,estrnro integer,tenro2 integer,estrnro2 integer,tenro3 integer,estrnro3 integer,porcentaje decimal (19,4),monto decimal (19,4))"
        End Select
        
        StrSql = "CREATE TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
Case 4: ' Oracle 9
        Select Case UCase(NombreTabla)
        Case "WF_CONCEPTO_DIST", "#WF_CONCEPTO_DIST":
            cadena = TTempWFConcepto_Dist & "(ternro double precision,concnro double precision,pronro double precision,masinro double precision,tenro double precision,estrnro double precision,tenro2 double precision,estrnro2 double precision,tenro3 double precision,estrnro3 double precision,porcentaje FLOAT(63),monto FLOAT(63))"
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

