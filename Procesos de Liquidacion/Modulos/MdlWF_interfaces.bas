Attribute VB_Name = "MdlWF_interfaces"
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
Global TTempWF_MOV_HORARIOS As String



Public Sub CargarNombresTablasTemporales()
' Setea los nombres de las tablas temporales de acuerdo al tipo de DB
Select Case TipoBD
    Case 1: ' DB2
            TTempWF_MOV_HORARIOS = "wf_MOV_HORARIOS"
    Case 2: ' Informix
            TTempWF_MOV_HORARIOS = "wf_MOV_HORARIOS"
    Case 3: ' SQL Server
            TTempWF_MOV_HORARIOS = "#wf_MOV_HORARIOS"
    Case 4: ' Oracle 9
            TTempWF_MOV_HORARIOS = "WF_MOV_HORARIOS"
    End Select
End Sub


Public Sub CreateTempTable(NombreTabla As String)
Dim cadena As String

On Error GoTo CE

Select Case TipoBD
Case 1: ' DB2
        Select Case UCase(NombreTabla)
        Case "WF_MOV_HORARIOS", "#WF_MOV_HORARIOS":
            cadena = TTempWF_MOV_HORARIOS & "(idtarjeta char(50), fecdesde date, fechasta date)"
        End Select
        StrSql = "CREATE TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
Case 2: ' Informix
        Select Case UCase(NombreTabla)
        Case "WF_MOV_HORARIOS", "#WF_MOV_HORARIOS":
            cadena = TTempWF_MOV_HORARIOS & "(idtarjeta char(50), fecdesde date, fechasta date)"
        End Select
        StrSql = "CREATE TEMP TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
Case 3: ' SQL Server
        Select Case UCase(NombreTabla)
        Case "WF_MOV_HORARIOS", "#WF_MOV_HORARIOS":
            cadena = TTempWF_MOV_HORARIOS & "(idtarjeta varchar(50), fecdesde datetime, fechasta datetime)"
        End Select
        StrSql = "CREATE TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
Case 4: ' Oracle 9
        Select Case UCase(NombreTabla)
        Case "WF_MOV_HORARIOS", "#WF_MOV_HORARIOS":
            cadena = TTempWF_MOV_HORARIOS & "(idtarjeta char(50), fecdesde date, fechasta date)"
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


Public Sub Insertar_WF_MOV_HORARIOS(ByVal idtarjeta As String, ByVal FechaDesde As Date, ByVal FechaHasta As Date)
Dim nombre As String

StrSql = "INSERT INTO " & TTempWF_MOV_HORARIOS & " (idtarjeta,fecdesde, fechasta) VALUES ('" & _
        idtarjeta & "'," & ConvFecha(FechaDesde) & "," & ConvFecha(FechaHasta) & ")"
objConn.Execute StrSql, , adExecuteNoRecords

End Sub

