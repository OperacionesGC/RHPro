Attribute VB_Name = "MdlColeccion"
Option Explicit

'Coleccion para mantener listas de strings
'Public Coleccion As New Collection


'Test
    'Dim c As New Collection
    'Dim b As New Class1

    'c.Add "a string", "a"
    'c.Add b, "b"

    'Debug.Print "a", Exists(c, "a") ' True '
    'Debug.Print "b", Exists(c, "b") ' True '
    'Debug.Print "c", Exists(c, "c") ' False '
    'Debug.Print 1, Exists(c, 1) ' True '
    'Debug.Print 2, Exists(c, 2) ' True '
    'Debug.Print 3, Exists(c, 3) ' False '


Public Function Exists(Coleccion, index) As Boolean
On Error GoTo ExistsTryNonObject
    Dim o As Object

    Set o = Coleccion(index)
    Exists = True
    Exit Function

ExistsTryNonObject:
    Exists = ExistsNonObject(Coleccion, index)
End Function

Private Function ExistsNonObject(Coleccion, index) As Boolean
On Error GoTo ExistsNonObjectErrorHandler
    Dim v As Variant

    v = Coleccion(index)
    ExistsNonObject = True
    Exit Function

ExistsNonObjectErrorHandler:
    ExistsNonObject = False
End Function

