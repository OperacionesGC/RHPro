Imports System.Xml.Serialization
Imports System.IO

Public Class ManejoArchivo

#Region "DECLARACIONES PRIVADAS"
    Private _pathLog As String
    Const _Tabulador = 5
#End Region
    

#Region "Propiedades"
    Public Property PathLog() As String
        Get
            Return _pathLog
        End Get
        Set(ByVal value As String)
            '_pathLog = value
        End Set
    End Property
#End Region




    ''' <summary>
    ''' Guarda un Archivo Xml en en disco
    ''' </summary>
    ''' <param name="Objet">Objeto a guardar</param>
    ''' <param name="Path">Path del archivo</param>
    ''' <remarks></remarks>
    Public Sub GuardarXml(ByVal Objet As Object, ByVal Path As String)
        Dim tXmlSerializer As XmlSerializer = New XmlSerializer(Objet.GetType())
        Dim tStremWriter As StreamWriter = New StreamWriter(Path)

        Try
            tXmlSerializer.Serialize(tStremWriter, Objet)
            tStremWriter.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Carga un archivo xml desde el Disco Rigido solo si el archivo pertenece a la 
    ''' misma clase que el Objeto Estructura
    ''' </summary>
    ''' <param name="Estructura">tipo de Objeto que se desea cargar</param>
    ''' <param name="Path">Path donde se encuentra el archivo Xml</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CargarXml(ByVal Estructura As Object, ByVal Path As String) As Object
        Dim tSReader As StreamReader
        Dim tObjet As Object
        Dim tXmlSerializer As XmlSerializer

        Try
            tSReader = New StreamReader(Path)
            tXmlSerializer = New XmlSerializer(Estructura.GetType())
            tObjet = tXmlSerializer.Deserialize(tSReader)
            tSReader.Close()

            If Estructura.GetType.FullName = tObjet.GetType.FullName Then
                Return tObjet
            Else
                Return Nothing
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Private Function leerArchivo()
        Return ""
    End Function

    Public Function CrearArchivo(ByVal path As String, ByVal nombre As String) As Boolean
        Try
            Dim Archivo As System.IO.FileStream
            ' crea un archivo vacio prueba.txt   
            Archivo = System.IO.File.Create(path & nombre)
            _pathLog = path & nombre
            Archivo.Close()

        Catch oe As Exception
            MsgBox(oe.Message, MsgBoxStyle.Critical)
        End Try
    End Function


    Public Function EscribirLinea(ByVal texto As String, Optional ByVal Espacio As Integer = 0) As Boolean
        Dim oSW As New StreamWriter(_pathLog, True)

        oSW.WriteLine(Space(_Tabulador * Espacio) & texto)
        oSW.Flush()
        oSW.Close()
    End Function

End Class
