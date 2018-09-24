Imports System.Data
Imports System.Configuration

Namespace ConsultaBaseC
    Public Class Encriptar
        Public Shared Function Encrypt(ByVal strEncryptionKey As String, ByVal strTextToEncrypt As String) As String
            Dim strTemp As String = ""

            For Outer As Integer = 0 To strEncryptionKey.Length - 1
                Dim Key As Integer = AscW(Convert.ToChar(strEncryptionKey.Substring(Outer, 1)))
                For Inner As Integer = 0 To strTextToEncrypt.Length - 1
                    strTemp = strTemp & Convert.ToString(ChrW((AscW(Convert.ToChar(strTextToEncrypt.Substring(Inner, 1)))) Xor Key))
                    Key = (Key + strEncryptionKey.Length) Mod 256
                Next
                strTextToEncrypt = strTemp
                strTemp = ""
            Next

            Return CadenaHex(strTextToEncrypt)
        End Function

        Public Shared Function Decrypt(ByVal strEncryptionKey As String, ByVal strTextToEncrypt As String) As String
            Dim strTemp As String = ""
            strTextToEncrypt = CadenaAscii(strTextToEncrypt)

            For Outer As Integer = 0 To strEncryptionKey.Length - 1
                Dim Key As Integer = AscW(Convert.ToChar(strEncryptionKey.Substring(Outer, 1)))
                For Inner As Integer = 0 To strTextToEncrypt.Length - 1
                    strTemp = strTemp & Convert.ToString(ChrW((AscW(Convert.ToChar(strTextToEncrypt.Substring(Inner, 1)))) Xor Key))
                    Key = (Key + strEncryptionKey.Length) Mod 256
                Next
                strTextToEncrypt = strTemp
                strTemp = ""
            Next

            Return strTextToEncrypt
        End Function

        Public Shared Function CadenaHex(ByVal strTextToEncrypt As String) As String
            Dim Buffer As String = ""

            For Outer As Integer = 0 To strTextToEncrypt.Length - 1
                strTextToEncrypt.Substring(Outer, 1)
                Dim Auxi As String = (AscW(Convert.ToChar(strTextToEncrypt.Substring(Outer, 1)))).ToString("X")
                If Auxi.Length < 2 Then
                    Auxi = "0" & Auxi
                End If
                Buffer = Buffer & Auxi
            Next

            Return Buffer
        End Function

        Public Shared Function CadenaAscii(ByVal strTextToEncrypt As String) As String
            Dim Buffer As String = ""

            Dim Outer As Integer = 0
            While Outer < strTextToEncrypt.Length
                Buffer = Buffer & Convert.ToString(ChrW(Integer.Parse(strTextToEncrypt.Substring(Outer, 2), System.Globalization.NumberStyles.HexNumber)))
                Outer = Outer + 2
            End While
            Return Buffer
        End Function
    End Class
End Namespace
