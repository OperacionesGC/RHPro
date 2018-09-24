VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CStack - Clase Stack

' Clase genérica para usar por cualquier algoritmo que use una Pila.
' Los items son almacenados como Variant,
' así que puede almacenar virtualmente cualquier tipo de Dato.

Option Explicit

Private m_StackArray() As Variant
Private m_nStackSize As Integer
Private m_nCurrPos As Integer

'tamaño de la Pila
Private Const INC_SIZE = 10

'Propiedad que retorna el nro de items de la Pila.
Public Property Get StackSize() As Integer
    StackSize = m_nCurrPos
End Property

'Inserta el valor al Tope de la Pila.
'La Pila es mas grande de lo que se necesita.
Public Sub Push(Value As Variant)
    m_nCurrPos = m_nCurrPos + 1
    If m_nCurrPos > m_nStackSize Then
        m_nStackSize = m_nStackSize + INC_SIZE
        ReDim Preserve m_StackArray(1 To m_nStackSize)
    End If
    m_StackArray(m_nCurrPos) = Value
End Sub

'Retorna el valor del elemento al Tope de la Pila removiendolo de la Pila.
'Si la Pila está vacia dá un Error.
Public Function Pop() As Variant
    If m_nCurrPos > 0 Then
        Pop = (m_StackArray(m_nCurrPos))
        m_nCurrPos = m_nCurrPos - 1
    Else
        ErrorEnExpresion = True
        'Err.Raise vbObjectError + 1011, , "Metodo Pop invocado sobre Pila vacia"
    End If
End Function

'Retorna el valor del elemento al Tope de la Pila sin remover de la Pila el elemento.
'Si la Pila está vacia dá un Error.
Public Function GetPopValue() As Variant
    If m_nCurrPos > 0 Then
        GetPopValue = m_StackArray(m_nCurrPos)
    Else
        ErrorEnExpresion = True
        'Err.Raise vbObjectError + 1012, , "Metodo GetPopValue invocado sobre Pila vacia"
    End If
End Function

'Inicialización de la Pila
Private Sub Class_Initialize()
    m_nStackSize = 0
    m_nCurrPos = 0
End Sub

