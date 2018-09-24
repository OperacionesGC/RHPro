Public Class ProcesamientoOnline
    Dim _Ternro As Long
    Dim _Fecha As String

    Public Property Ternro() As Long
        Get
            Return _Ternro
        End Get
        Set(ByVal value As Long)
            _Ternro = value
        End Set
    End Property

    Public Property Fecha() As String
        Get
            Return _Fecha
        End Get
        Set(ByVal value As String)
            _Fecha = value
        End Set
    End Property

End Class
