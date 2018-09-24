VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ccache"
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
Public Cache As New CColeccion

' Metodos de la estructura
Public Sub Cache_Borrar()
Dim X
' Borra la estructura del cache
    For Each X In objCache.Cache.m_cache
        objCache.Cache.Delete (X)
    Next X
End Sub
    
Public Function Cache_Existe(p_param As String) As Boolean
' Retorna TRUE si es que el parametro esta en el cache
Dim Esta As Boolean

    Esta = False
    On Error Resume Next
    If objCache.Cache.IsSymbolDefined(p_param) Then
        Esta = True
    Else
        Esta = False
    End If
    On Error GoTo 0
    
    Cache_Existe = Esta
End Function

Public Function Cache_Valor_Param(p_param As String) As Single
' retorna el valor del parametro en el cache
    
    Cache_Valor_Param = objCache.Cache.Value(p_param)

End Function

Public Sub Cache_Insertar(ByVal p_param As String, ByVal p_valor As Single)
' Inserta un elemento en el cache
Dim TablaCache As New CColeccion

    Set objCache.Cache = TablaCache
    objCache.Cache.Add p_param, p_valor
    
End Sub
