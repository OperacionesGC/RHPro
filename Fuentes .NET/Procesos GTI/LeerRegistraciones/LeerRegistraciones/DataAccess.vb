Imports System.IO
Imports System.Data.OleDb
Imports System.Data
Imports System.Windows.Forms

Public Class DataAccess

#Region "DECLARACIONES PRIVADAS"
    Private _PathProcesos As String
    Private _PathFLog As String
    Private _strformatoFservidor As String
    Private _strconexion As String
    Private _TipoBD As String
    Private _SCHEMA As String
    Private _EncriptStrconexion As Boolean
    Private _c_seed As String
    Private _PathSAP
#End Region


#Region "Constructor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal cargarConfi As Boolean)
        ' carga las configuraciones basicas para los procesos
        Dim strline As String
        Dim pos1 As Integer
        Dim pos2 As Integer
        Dim Encontro As Boolean
        Dim objReader As StreamReader

        If File.Exists(Application.StartupPath & "\rhproprocesos.ini") Then
            objReader = New StreamReader(Application.StartupPath & "\rhproprocesos.ini")
        Else
            If File.Exists(CurDir() & "\rhproappsrv.ini") Then
                objReader = New StreamReader(Application.StartupPath & "\RHProappSrv.ini")
            End If
        End If

        Try

            If Not Etiqueta Is Nothing Then
                Encontro = False
                Do While Not objReader.EndOfStream And Not Encontro
                    strline = objReader.ReadLine()
                    If InStr(1, UCase(strline), UCase(Etiqueta)) > 0 Then
                        Encontro = True
                    End If
                Loop
            End If

            ' Path del Proceso de SAP (lo usa el SAP)
            If Not objReader.EndOfStream Then
                strline = objReader.ReadLine()
                pos1 = InStr(1, strline, "[") + 1
                pos2 = InStr(1, strline, "]")
                _PathSAP = Mid(strline, pos1, pos2 - pos1)
                If Right(_PathSAP, 1) <> "\" Then _PathSAP = _PathSAP & "\"
            End If

            ' Path de los ejecutables de los procesos (lo usa el AppSrv)
            If Not objReader.EndOfStream Then
                strline = objReader.ReadLine()
                pos1 = InStr(1, strline, "[") + 1
                pos2 = InStr(1, strline, "]")
                _PathProcesos = Mid(strline, pos1, pos2 - pos1)
                If Right(_PathProcesos, 1) <> "\" Then _PathProcesos = _PathProcesos & "\"
            End If

            ' seteo del path del archivo de Log
            If Not objReader.EndOfStream Then
                strline = objReader.ReadLine()
                pos1 = InStr(1, strline, "[") + 1
                pos2 = InStr(1, strline, "]")
                _PathFLog = Mid(strline, pos1, pos2 - pos1)
                If Right(_PathFLog, 1) <> "\" Then _PathFLog = _PathFLog & "\"
            End If

            ' seteo del formato de Fecha del Servidor
            If Not objReader.EndOfStream Then
                strline = objReader.ReadLine()
                pos1 = InStr(1, strline, "[") + 1
                pos2 = InStr(1, strline, "]")
                _strformatoFservidor = Mid(strline, pos1, pos2 - pos1)
            End If

            ' seteo del string de conexion
            If Not objReader.EndOfStream Then
                strline = objReader.ReadLine()
                pos1 = InStr(1, strline, "[") + 1
                pos2 = InStr(1, strline, "]")
                _strconexion = Mid(strline, pos1, pos2 - pos1)
            End If

            ' seteo del tipo de Base de datos
            If Not objReader.EndOfStream Then
                strline = objReader.ReadLine()
                pos1 = InStr(1, strline, "[") + 1
                pos2 = InStr(1, strline, "]")
                _TipoBD = Mid(strline, pos1, pos2 - pos1)
            End If

            'Etiqueta (solo para el AppSrv)
            If Not objReader.EndOfStream Then
                strline = objReader.ReadLine()
            End If

            'FGZ - 23/06/2008 - Se agregó este parametro
            ' seteo del schema de Base de datos
            _SCHEMA = ""
            If Not objReader.EndOfStream Then
                strline = objReader.ReadLine()
                pos1 = InStr(1, strline, "[") + 1
                pos2 = InStr(1, strline, "]")
                _SCHEMA = Mid(strline, pos1, pos2 - pos1)
            End If

            objReader.Close()

        Catch ex As Exception
            Throw New Exception("Error al cargar el archivo de configuración")
        End Try
    End Sub

#End Region


#Region "Propiedades"
    Public Property PathFLog() As String
        Get
            Return _PathFLog
        End Get
        Set(ByVal value As String)
            _PathFLog = value
        End Set
    End Property

    Public Property PathProcesos() As String
        Get
            Return _PathProcesos
        End Get
        Set(ByVal value As String)
            _PathProcesos = value
        End Set
    End Property

    Public Property Conexion() As String
        Get
            Return _strconexion
        End Get
        Set(ByVal value As String)
            _strconexion = value
        End Set
    End Property

#End Region

#Region "Metodos"

    Public Function ConvFecha(ByVal dteFecha As Date) As String
        If UCase(_strformatoFservidor) = "DD/MM/YYYY" Then
            Return "'" & Format(C_Date(dteFecha), "dd/MM/yyyy") & "'"
        Else
            Return "'" & Format(C_Date(dteFecha), _strformatoFservidor) & "'"
        End If
    End Function

    Public Function C_Date(ByVal Fecha) As Date        
        Return Format(Fecha, "dd/MM/yyyy")
    End Function


    Public Function getLastIdentity(ByRef objConn As OleDbConnection, ByVal NombreTabla As String) As Object
        Dim dtTable As New DataTable
        Dim da As OleDbDataAdapter
        Dim cmd As New OleDbCommand()
        Dim StrSql As String = ""

        Select Case _TipoBD
            Case 1 'db2
                StrSql = "SELECT identity_val_local() as Codigo FROM sysibm.sysdummy1"
            Case 2 ' Informix
                StrSql = "select unique DBINFO('sqlca.sqlerrd1') as codigo from " & NombreTabla
            Case 3 ' sql server
                'StrSql = "SELECT @@IDENTITY as Codigo FROM " & NombreTabla & ""
                StrSql = "SELECT SCOPE_IDENTITY() as Codigo "
                'StrSql = "SELECT @@IDENTITY as Codigo"
            Case 4 ' Oracle 9
                StrSql = "select SEQ_" & UCase(NombreTabla) & ".CURRVAL as Codigo FROM DUAL"
        End Select
        
        cmd.CommandText = StrSql
        cmd.Connection = objConn        
        Dim Codigo As Long = cmd.ExecuteScalar


        If (Codigo > 0) Then
            Return Codigo
        Else
            Return -1
        End If

    End Function
#End Region
End Class
