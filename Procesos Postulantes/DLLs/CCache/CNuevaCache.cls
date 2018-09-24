VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNuevaCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------------------------
' Descripcion:Modulo de Cache de Valores de Parametros
' Autor      :FGZ
' Fecha      :11/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

'definicion de la variable para la estructura Cache
Public Cache As New Collection


Public Sub Limpiar()
Dim X
' Borra la estructura del cache
'    For Each X In Cache
'        Cache.Remove 1
'    Next X
    
    For X = 1 To Cache.Count
       Cache.Remove 1
    Next X
End Sub


'Indica si el simbolo esta definido actualmente en la tabla.
Public Function EsSimboloDefinido(sName As String) As Boolean
    Dim nValue As Variant
    On Error Resume Next
    nValue = Cache(sName)
    If Err Then
        EsSimboloDefinido = False
    Else
        EsSimboloDefinido = True
    End If
End Function


'Borra el simbolo de la tabla. Si el simbolo no esta definfo, no hace nada.
Public Sub Borrar_Simbolo(sName As String)
    If EsSimboloDefinido(sName) Then
        Cache.Remove sName
    End If
End Sub


Public Sub Insertar_Simbolo(sName As String, nValue As Variant)
    If EsSimboloDefinido(sName) = False Then
        If IsNumeric(nValue) Then
            Cache.Add nValue, sName
        Else
            Cache.Add nValue, sName
        End If
    End If
End Sub


'Setea el valor del simbolo especificado.
'Da un error en tiempo de ejecucion si el simbolo ya existe o su valor no es numerico.
Public Property Let Valor(sName As String, nValue As Variant)
    If EsSimboloDefinido(sName) = True Then
        If IsNumeric(nValue) Then
            Borrar_Simbolo sName
            Insertar_Simbolo sName, nValue
        Else
            'Error
        End If
    Else
        'Error. "Simbolo no definido"
    End If
End Property

'Retorna el valor del simbolo especificado. Da un error en tiempo de ejecucion si el simbolo no existe.
Public Property Get Valor(sName As String) As Variant
    If EsSimboloDefinido(sName) = True Then
        Valor = Cache(sName)
    Else
        'Error. "Simbolo no definido"
    End If
End Property


'Retorna la cantidad de simbolos
Public Property Get Cantidad() As Integer
    Cantidad = Cache.Count
End Property


