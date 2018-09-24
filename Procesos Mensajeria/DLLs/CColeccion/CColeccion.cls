VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Coleccion para mantener los valores
Public m_cache As New Collection

Public Sub Add(sName As String, nValue As Variant)
    If IsSymbolDefined(sName) = False Then
        If IsNumeric(nValue) Then
            m_cache.Add nValue, sName
        Else
            m_cache.Add nValue, sName
        End If
    End If
End Sub

'Borra el simbolo de la tabla.
'Si el simbolo no esta definfo, no hace nada.
Public Sub Delete(sName As String)
    If IsSymbolDefined(sName) Then
        m_cache.Remove sName
    End If
End Sub

'Indica si el simbolo esta definido actualmente en la tabla.
Public Function IsSymbolDefined(sName As String) As Boolean
    Dim nValue As Variant
    On Error Resume Next
    nValue = m_cache(sName)
    If Err Then
        IsSymbolDefined = False
    Else
        IsSymbolDefined = True
    End If
End Function

'Setea el valor del simbolo especificado
'Da un error en tiempo de ejecucion si el simbolo ya existe o su valor no es numerico.
Public Property Let Value(sName As String, nValue As Variant)
    If IsSymbolDefined(sName) = True Then
        If IsNumeric(nValue) Then
            Delete sName
            Add sName, nValue
        Else
            'Error
        End If
    Else
        'Error. "Simbolo no definido"
    End If
End Property

'Retorna el valor del simbolo especificado.
'Da un error en tiempo de ejecucion si el simbolo no existe.
Public Property Get Value(sName As String) As Variant
    If IsSymbolDefined(sName) = True Then
        Value = m_cache(sName)
    Else
        'Error. "Simbolo no definido"
    End If
End Property

'Retorna la cantidad de simbolos
Public Property Get Count() As Integer
    Count = m_cache.Count
End Property


