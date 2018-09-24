VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSymbolTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CSymbolTable - Clase Tabla de Simbolos

' Esta tabla es usada por la clase CEval. Es una simple colección para hacer mas facil
' almacenar un alista de simbolos y sus valores asociados.

Option Explicit

'Coleccion para mantener los valores de los simbolos
Private m_Symbols As New Collection
'Dim m_Symbols As New Collection

'Agregar un simbolo a la Tabla de Simbolos
'Da un error en tiempo de ejecucion si el simbolo ya existe o su valor no es numerico.
Public Sub Add(sName As String, nValue As Variant)
    If IsSymbolDefined(sName) = False Then
        If IsNumeric(nValue) Then
            m_Symbols.Add nValue, sName
        Else
            m_Symbols.Add nValue, sName
            'Err.Raise vbObjectError + 1020, , "Símbolo con valor invalido"
        End If
    Else
        Err.Raise vbObjectError + 1021, , "El símbolo ya está definido"
    End If
End Sub

'Borra el simbolo de la tabla.
'Si el simbolo no esta definfo, no hace nada.
Public Sub Delete(sName As String)
    If IsSymbolDefined(sName) Then
        m_Symbols.Remove sName
    End If
End Sub

'Indica si el simbolo esta definido actualmente en la tabla.
Public Function IsSymbolDefined(sName As String) As Boolean
    Dim nValue As Variant
    On Error Resume Next
    nValue = m_Symbols(sName)
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
            Err.Raise vbObjectError + 1020, , "Símbolo con valor invalido"
        End If
    Else
        Err.Raise vbObjectError + 1022, , "Simbolo no definido"
    End If
End Property

'Retorna el valor del simbolo especificado.
'Da un error en tiempo de ejecucion si el simbolo no existe.
Public Property Get Value(sName As String) As Variant
    If IsSymbolDefined(sName) = True Then
        Value = m_Symbols(sName)
    Else
        'Err.Raise vbObjectError + 1022, , "Simbolo no definido"
    End If
End Property

'Retorna la cantidad de simbolos en la tabla.
Public Property Get Count() As Integer
    Count = m_Symbols.Count
End Property

