VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Feriado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim nroConv As Long
Dim NombreConvenio As String
Dim nroSuc As Long
Dim NombreSucursal As String
Dim nroEstr As Long
Dim NombreEstr As String


Property Get Numero_Convenio() As Long
    Numero_Convenio = nroConv
End Property
Property Get Nombre_Convenio() As Long
    Nombre_Convenio = NombreConvenio
End Property
Property Get Numero_Sucursal() As Long
    Numero_Sucursal = nroSuc
End Property
Property Get Nombre_Sucursal() As Long
    Nombre_Sucursal = NombreSucursal
End Property

Property Set Conexion(ByRef cn As ADODB.Connection)
    Set objConn = cn
End Property
Property Set ConexionTraza(ByRef cn As ADODB.Connection)
    Set CnTraza = cn
End Property

Public Function Feriado(dia As Date, ternro As Long, depurar As Boolean) As Boolean
Dim objRsFeriado As New ADODB.Recordset

    Feriado = False

    'Determino el nro de la estructura
    StrSql = " SELECT * FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
    StrSql = StrSql & " WHERE Alcance_Testr.tanro = " & lngAlcanEstr & " And " & _
             " his_estructura.Ternro = " & ternro & " AND htetdesde <= " & ConvFecha(dia) & _
             " AND (htethasta >= " & ConvFecha(dia) & " Or htethasta Is Null )" & _
             " ORDER BY htetdesde DESC "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        nroEstr = objRs!estrnro
        NombreEstr = Trim(" " & objRs!estrdabr)
    End If
    
    StrSql = "SELECT * FROM Feriado WHERE ferifecha = " & ConvFecha(dia)
    OpenRecordset StrSql, objRsFeriado
    
    If Not objRsFeriado.EOF Then
        If objRsFeriado!tipferinro = 1 Then
            'Pais
            StrSql = "SELECT * FROM Pais WHERE paisdef = " & blnTRUE
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                If CDbl(objRsFeriado!fericodext) = objRs!paisnro Then Feriado = True
            End If
        Else
            StrSql = "SELECT * FROM Fer_estr WHERE estrnro = " & nroEstr & " AND ferinro = " & objRsFeriado!ferinro
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then Feriado = True
        End If
        
        If depurar Then GeneraTraza ternro, dia, "Es Feriado", Str(Feriado)
    End If
End Function

Private Sub GeneraTraza(ternro As Long, Fecha As Date, desc As String, Optional Valor As String = "?")

    StrSql = "INSERT INTO gti_traza(ternro,fecproc,descripcion,valor) VALUES (" & _
             ternro & "," & ConvFecha(Fecha) & ",'" & desc & "','" & Valor & "')"
    
    CnTraza.Execute StrSql, , adExecuteNoRecords
    
End Sub

