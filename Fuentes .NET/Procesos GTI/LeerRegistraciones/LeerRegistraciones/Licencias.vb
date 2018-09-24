Imports System.Data
Imports System.Data.OleDb

Public Class Licencias
#Region "Constructor"
    Sub New()

    End Sub


#End Region

#Region "Metodos"
    Function TotalDiasCorrespondientes(ByVal ternro As String) As Double

        Dim StrSql As String
        Dim dsDatos As New DataSet
        Dim dtDatos As New DataTable
        Dim da As OleDbDataAdapter
        Dim conexionAux As New OleDbConnection
        Dim confiBasica As New DataAccess(True)
        Dim diasCorresp As Double

        'conexionAux = New OleDbConnection(confiBasica.Conexion)

        StrSql = "SELECT SUM(vdiascorcant) vdiascorcant FROM vacdiascor WHERE venc=0 AND ternro= " & ternro
        dtDatos = New DataTable
        'da = New OleDbDataAdapter(StrSql, conexionAux.ConnectionString)
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)

        diasCorresp = 0
        If (dtDatos.Rows.Count > 0) Then
            If Not IsDBNull(dtDatos.Rows(0).Item(0)) Then
                diasCorresp = CDbl(dtDatos.Rows(0).Item("vdiascorcant"))
            End If
            Return diasCorresp
        End If

        Try
            Return diasCorresp
        Catch ex As Exception
            Return "0"
        End Try

    End Function


    Function TotalDiasGozados(ByVal ternro As String) As Double
        Dim StrSql As String
        Dim dsDatos As New DataSet
        Dim dtDatos As New DataTable
        Dim da As OleDbDataAdapter
        Dim diasGozados As Double

        StrSql = " SELECT SUM(elcantdiashab) elcantdiashab FROM emp_lic " & _
                 " WHERE emp_lic.empleado = " & ternro & " AND tdnro = 2 AND licestnro = 2"
        dtDatos = New DataTable
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)

        diasGozados = 0
        If (dtDatos.Rows.Count > 0) Then
            If Not IsDBNull(dtDatos.Rows(0).Item(0)) Then
                diasGozados = CDbl(dtDatos.Rows(0).Item("elcantdiashab"))
            End If
            Return diasGozados
        End If

        Try
            Return diasGozados
        Catch ex As Exception
            Return "0"
        End Try
    End Function


    Function TotalDiasBeneficio(ByVal ternro As String) As Double
        Dim StrSql As String
        Dim dsDatos As New DataSet
        Dim dtDatos As New DataTable
        Dim da As OleDbDataAdapter
        Dim diasBeneficio As Double

        StrSql = " SELECT SUM(vdiascorcant) vdiascorcant FROM vacdiascor WHERE venc=3 AND ternro= " & ternro

        dtDatos = New DataTable
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)

        diasBeneficio = 0
        If (dtDatos.Rows.Count > 0) Then
            If Not IsDBNull(dtDatos.Rows(0).Item(0)) Then
                diasBeneficio = CDbl(dtDatos.Rows(0).Item("elcantdiashab"))
            End If
            Return diasBeneficio
        End If

        Try
            Return diasBeneficio
        Catch ex As Exception
            Return "0"
        End Try
    End Function

    Function TotalDiasVendidos(ByVal ternro As String) As Double
        Dim StrSql As String
        Dim dsDatos As New DataSet
        Dim dtDatos As New DataTable
        Dim da As OleDbDataAdapter
        Dim diasVendidos As Double

        StrSql = " SELECT SUM(cantvacvendidos) cantvacvendidos FROM vacvendidos " & _
                 " WHERE aprobado = -1 AND ternro = " & ternro

        dtDatos = New DataTable
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)

        diasVendidos = 0
        If (dtDatos.Rows.Count > 0) Then
            If Not IsDBNull(dtDatos.Rows(0).Item(0)) Then
                diasVendidos = CDbl(dtDatos.Rows(0).Item("elcantdiashab"))
            End If
            Return diasVendidos
        End If

        Try
            Return diasVendidos
        Catch ex As Exception
            Return "0"
        End Try
    End Function

#End Region



End Class
