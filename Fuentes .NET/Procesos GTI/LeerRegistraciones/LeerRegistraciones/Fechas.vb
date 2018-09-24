Public Class Fechas
#Region "Constructor"
    Sub New()

    End Sub


#End Region

#Region "Metodos"

    ''Cambia la fecha al formato año mes dia hh mm toda junta, sin espacio. Rellena con 00 la hora y minuto
    Public Function cambiaFecha(ByVal Actual As String, Optional ByRef formato As String = "") As String
        Try
            Select Case formato
                Case "", "yyyy-mm-dd"
                    Return Actual.Substring(6, 4) & Actual.Substring(3, 2) & Actual.Substring(0, 2) & "0000"
                Case "dd-mm-yyyy"
                    Return Actual.Substring(0, 2) & Actual.Substring(3, 2) & Actual.Substring(6, 4)
                Case Else
                    Return "0000000000"
            End Select


        Catch ex As Exception
            Return "0000000000"
        End Try

    End Function


    Function ObtenerHoraDeFecha(ByVal Actual As String) As String
        Try
            Return Actual.Substring(11, 2) & Actual.Substring(14, 2)
        Catch ex As Exception
            Return ""
        End Try

    End Function

    Function convFechaQP(ByVal Actual As String, Optional ByRef formato As String = "") As String
        Select Case formato
            Case "", "dd-mm-yyyy"
                Return "'" & Actual.Substring(0, 2) & "/" & Actual.Substring(3, 2) & "/" & Actual.Substring(6, 4) & "'"
            Case "yyyy-mm-dd"
                Return "'" & Actual.Substring(6, 4) & "/" & Actual.Substring(3, 2) & "/" & Actual.Substring(0, 2) & "'"
            Case Else
                Return ""
        End Select
    End Function

    Function convFec(ByVal Actual As String, Optional ByRef formato As String = "") As String
        Select Case formato
            Case "", "dd-mm-yyyy"
                Return "'" & Actual.Substring(6, 2) & "/" & Actual.Substring(4, 2) & "/" & Actual.Substring(0, 4) & "'"
            Case "yyyy-mm-dd"
                Return "'" & Actual.Substring(0, 4) & "/" & Actual.Substring(4, 2) & "/" & Actual.Substring(6, 2) & "'"
            Case Else
                Return ""
        End Select
    End Function



#End Region


End Class
