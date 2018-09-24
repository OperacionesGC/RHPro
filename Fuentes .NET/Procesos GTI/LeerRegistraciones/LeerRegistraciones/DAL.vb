Imports System.Configuration
Imports System.Data
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Data.OleDb
Imports System
Imports System.Configuration.ConfigurationSettings



Namespace ConsultaBaseC
    Public Class DAL
        Public Shared Function [Error](ByVal Codigo As Integer, ByVal Idioma As String) As String
            Dim salida As String = BuscarError(Convert.ToString(Codigo) & Idioma)

            If salida Is Nothing Then
                Return "Codigo de error " & Codigo & " para el idioma " & Idioma & " no encontrado."
            Else
                Return salida
            End If
        End Function

        Public Shared Function constr(ByVal NroBase As String) As String
            'string cnnAux = Encriptar.Decrypt(DAL.EncrKy(), ConfigurationManager.ConnectionStrings[NroBase].ConnectionString);
            'Dim cnnAux As String = ConfigurationSettings.ConnectionStrings(NroBase).ConnectionString
            Dim cnnAux As String = NroBase
            'cnnAux = cnnAux & " User Id=" & Encriptar.Decrypt(DAL.EncrKy(), UsuESS()) & ";"
            'cnnAux = cnnAux & " Password=" & Encriptar.Decrypt(DAL.EncrKy(), PassESS()) & ";"

            If TipoBase(NroBase).ToUpper() = "ORA" Then
                Esquema(cnnAux, NroBase)
            End If

            Return cnnAux

        End Function

        Public Shared Sub Esquema(ByVal conexion As String, ByVal NroBase As String)
            Dim cn2 As New OleDbConnection()
            cn2.ConnectionString = conexion
            cn2.Open()
            Dim sqlSS As String = "ALTER SESSION SET NLS_SORT = BINARY"
            Dim cmd As New OleDbCommand(sqlSS, cn2)
            cmd.ExecuteNonQuery()
            sqlSS = "ALTER SESSION SET CURRENT_SCHEMA = " & BuscarEsquema(NroBase)
            cmd = New OleDbCommand(sqlSS, cn2)
            cmd.ExecuteNonQuery()
            If cn2.State = ConnectionState.Open Then
                cn2.Close()
            End If
        End Sub

        Public Shared Function constrUsu(ByVal User As String, ByVal Pass As String, ByVal segNT As String, ByVal NroBase As String) As String
            Dim cnnAux As String = ConfigurationManager.ConnectionStrings(NroBase).ConnectionString

            If segNT = "TrueValue" Then
                cnnAux = cnnAux & " Integrated Security=SSPI;"
            Else
                cnnAux = cnnAux & " User Id=" & User & ";"
                cnnAux = cnnAux & " Password=" & Pass & ";"
            End If

            Return cnnAux
        End Function

        Public Shared Function Bases() As DataTable
            'Creo la tabla de salida
            Dim tablaSalida As New DataTable("table")
            Dim Columna As New DataColumn()
            Columna.DataType = System.Type.[GetType]("System.String")
            Columna.ColumnName = "combo"
            Columna.AutoIncrement = False
            Columna.Unique = False
            tablaSalida.Columns.Add(Columna)

            Dim appSettings As NameValueCollection = ConfigurationManager.AppSettings
            Dim appSettingsEnum As IEnumerator = appSettings.Keys.GetEnumerator()

            Dim i As Integer = 0

            While appSettingsEnum.MoveNext()
                Dim key As String = appSettings.Keys(i)
                If isNumeric(key) Then
                    Dim fila As DataRow = tablaSalida.NewRow()
                    fila("combo") = appSettings(key).ToString()
                    tablaSalida.Rows.Add(fila)
                End If
                i += 1
            End While

            Return tablaSalida
        End Function

        Public Shared Function TipoBase(ByVal NroBase As String) As String
            Dim Salida As String = "SQL"

            Dim appSettings As NameValueCollection = ConfigurationManager.AppSettings
            Dim appSettingsEnum As IEnumerator = appSettings.Keys.GetEnumerator()

            Dim Ciclar As Boolean = True
            Dim Encontro As Boolean = False
            Dim i As Integer = 0
            Dim Fila As String
            Dim ArrFila As String()

            While Ciclar
                If isNumeric(appSettings.Keys(i)) Then
                    Fila = appSettings(appSettings.Keys(i)).ToString()
                    ArrFila = Fila.Split(New Char() {","c})
                    If ArrFila(1) = NroBase Then
                        Encontro = True
                        If ArrFila.Length >= 5 Then
                            Salida = ArrFila(4)
                        End If
                    End If
                Else
                    Ciclar = False
                End If

                i += 1

                If (i > appSettings.Count) OrElse (Encontro) Then
                    Ciclar = False
                End If
            End While

            Return Salida
        End Function

        Public Shared Function BuscarError(ByVal Clave As String) As String
            Return ConfigurationSettings.AppSettings(Clave)
        End Function

        Public Shared Function UsuESS() As String
            Return ConfigurationSettings.AppSettings("UsuESS")
        End Function

        Public Shared Function PassESS() As String
            Return ConfigurationSettings.AppSettings("PassESS")
        End Function

        Public Shared Function EncrKy() As String
            Return ConfigurationSettings.AppSettings("EncrKy")
        End Function

        Public Shared Function isNumeric(ByVal value As Object) As Boolean
            Try
                Dim d As Double = System.[Double].Parse(value.ToString(), System.Globalization.NumberStyles.Any)
                Return True
            Catch generatedExceptionName As System.FormatException
                Return False
            End Try
        End Function

        Public Shared Function DescEstr(ByVal Est As Integer) As String
            If ConfigurationSettings.AppSettings("TEDESC" & Est.ToString()) Is Nothing Then
                Return ""
            Else
                Return ConfigurationSettings.AppSettings("TEDESC" & Est.ToString())
            End If
        End Function

        Public Shared Function BuscarEsquema(ByVal NroBase As String) As String
            If ConfigurationSettings.AppSettings("SCHEMA" & NroBase) Is Nothing Then
                Return ""
            Else
                Return ConfigurationSettings.AppSettings("SCHEMA" & NroBase)
            End If
        End Function

        Public Shared Function NroEstr(ByVal Est As Integer) As String
            If ConfigurationSettings.AppSettings("TENRO" & Est.ToString()) Is Nothing Then
                Return "0"
            Else
                Return ConfigurationSettings.AppSettings("TENRO" & Est.ToString())
            End If
        End Function
    End Class
End Namespace
