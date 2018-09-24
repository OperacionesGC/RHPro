VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Familiares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private listaFamiliares As String
Private Famparenro As String
Private Faminc As String
Private Famtrab As String
Private Famestudia As String
Private FamiDEmeM As Integer
Private cobraSalario As String
Private nivCodExt As String
Private rs_datos As New ADODB.Recordset
Public Property Get obtenerFamiliares() As String
   obtenerFamiliares = listaFamiliares
End Property
Public Property Get obtenerCondicion() As String
    If Faminc = -1 Then
        obtenerCondicion = "M"
    Else
        obtenerCondicion = "N"
    End If
End Property
Public Property Get obtenerTrabaja() As String
    If Famtrab = -1 Then
        obtenerTrabaja = "S"
    Else
        obtenerTrabaja = "N"
    End If
End Property
Public Property Get obtenerEstudia() As String
    If Famestudia = -1 Then
        obtenerEstudia = "S"
    Else
        obtenerEstudia = "N"
    End If
End Property
Public Property Get obtenerEmerMed() As Integer
    obtenerEmerMed = FamiDEmeM
End Property
Public Property Get obtenerRelacionFamiliar() As String
'    StrSql = " SELECT codexterno FROM mapeo_sap "
'    StrSql = StrSql & " WHERE tablaref='parentesco'"
'    StrSql = StrSql & " AND codInterno = " & Famparenro
'    StrSql = StrSql & " AND upper(left(infotipo,3))='PY_'"
'    OpenRecordset StrSql, rs_datos
'    If Not rs_datos.EOF Then
'        If Not EsNulo(rs_datos("codexterno")) Then
'            obtenerRelacionFamiliar = "O"
'        Else
'            obtenerRelacionFamiliar = rs_datos("codexterno")
'        End If
'    Else
'        obtenerRelacionFamiliar = "O"
'    End If
'    rs_datos.Close
    obtenerRelacionFamiliar = Famparenro

End Property
Public Property Get obtenerSalarioFamiliar() As String
    obtenerSalarioFamiliar = cobraSalario
End Property
Public Property Get obtenerNivel() As String
    obtenerNivel = nivCodExt
End Property

Public Sub buscarFamiliares(ByVal ternro As Integer)
Dim cant As Integer
listaFamiliares = "0"
cant = 0
    StrSql = " SELECT ternro FROM familiar "
    StrSql = StrSql & " WHERE empleado=" & ternro
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        Do While Not rs_datos.EOF
            listaFamiliares = listaFamiliares & "," & rs_datos("ternro")
            cant = cant + 1
        rs_datos.MoveNext
        Loop
        Flog.writeline "Cantidad de familiares del tercero:" & ternro & " " & cant
    Else
        Flog.writeline "El tercero: " & ternro & " no tiene familiares"
    End If
    rs_datos.Close
End Sub


Public Sub buscarRelacionFamiliar(ByVal ternro As Integer)
    StrSql = "SELECT * FROM familiar "
    StrSql = StrSql & " INNER JOIN parentesco P ON P.parenro = familiar.parenro "
    StrSql = StrSql & " LEFT JOIN estructura e ON e.estrnro = familiar.famemermed "
    StrSql = StrSql & " WHERE ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        Famparenro = IIf(EsNulo(rs_datos!parcodext), 0, rs_datos!parcodext)
        Faminc = IIf(EsNulo(rs_datos!Faminc), 0, rs_datos!Faminc)
        Famtrab = IIf(EsNulo(rs_datos!Famtrab), 0, rs_datos!Famtrab)
        Famestudia = IIf(EsNulo(rs_datos!Famestudia), 0, rs_datos!Famestudia)
        FamiDEmeM = IIf(EsNulo(rs_datos!estrcodext), 0, rs_datos!estrcodext)
    Else
        Famparenro = ""
        Faminc = 0
        Famtrab = 0
        Famestudia = 0
        FamiDEmeM = 0
        Flog.writeline "Ocurrio un error buscando la relacion familiar"
        Exit Sub
    End If
    rs_datos.Close
End Sub

Public Sub buscarSalarioFamiliar(ByVal ternro As Integer)
    StrSql = "SELECT famsalario FROM familiar"
    StrSql = StrSql & " WHERE ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If Not EsNulo(rs_datos!famsalario) Then
            If rs_datos!famsalario = -1 Then
                cobraSalario = "S"
            Else
                cobraSalario = "N"
            End If
        Else
            cobraSalario = "N"
        End If
    Else
        cobraSalario = "N"
    End If
    rs_datos.Close
End Sub

Public Sub buscarTitulo(ByVal ternro As Integer)
    StrSql = "SELECT nivcodext FROM estudio_Actual "
    StrSql = StrSql & " INNER JOIN nivest ON nivest.nivnro = estudio_actual.nivnro"
    StrSql = StrSql & " WHERE ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        nivCodExt = IIf(EsNulo(rs_datos!nivCodExt), "", rs_datos!nivCodExt)
    Else
        nivCodExt = "N"
    End If
    rs_datos.Close
End Sub
