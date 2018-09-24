Attribute VB_Name = "mdlconfrepadv"
Option Explicit

Private Type TConfRepAdv
    confnrocol As Long
    Etiqueta As String
    Tipo As String
    tipo2 As String
    tipo3 As String
    tipo4 As String
    tipo5 As String
    Confval As String
    confval2 As String
    confval3 As String
    confval4 As String
    confval5 As String
End Type

Const topeArreglo As Integer = 300
Global confRep(topeArreglo) As TConfRepAdv
Dim rs_Confrep As New ADODB.Recordset

Public Sub ArmoDatosConfrep(repnro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtiene configuración del confrep
' Autor      : Gonzalez Nicolás
' Fecha      : 18/06/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim contador As Long
Dim X

StrSql = "SELECT confetiq, confnrocol, conftipo, conftipo2, conftipo3, conftipo4, conftipo5"
StrSql = StrSql & ",confval, confval2, confval3, confval4, confval5 "
StrSql = StrSql & " FROM confrepAdv "
StrSql = StrSql & " WHERE repnro = " & repnro
StrSql = StrSql & " ORDER BY confrepAdv.confnrocol ASC"
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Texto = "No se encontro configuracion del reporte"
    HuboError = True
Else
    Do While Not rs_Confrep.EOF
        contador = CLng(rs_Confrep("confnrocol"))
        If contador > topeArreglo Then
           Flog.writeline " ERROR:La columna del confrep supera el limite establecido."
        Else
            confRep(contador).confnrocol = CLng(rs_Confrep("confnrocol"))
            'Guarda el resto de las columnas
            confRep(contador).Etiqueta = rs_Confrep("confetiq")
            'TIPOS
            If Not EsNulo(rs_Confrep("conftipo")) Then
                confRep(contador).Tipo = rs_Confrep("conftipo")
            Else
                confRep(contador).Tipo = ""
            End If
            
            If Not EsNulo(rs_Confrep("conftipo2")) Then
                confRep(contador).tipo2 = rs_Confrep("conftipo2")
            Else
                confRep(contador).tipo2 = ""
            End If
            
            If Not EsNulo(rs_Confrep("conftipo3")) Then
                confRep(contador).tipo3 = rs_Confrep("conftipo3")
            Else
                confRep(contador).tipo3 = ""
            End If

            If Not EsNulo(rs_Confrep("conftipo4")) Then
                confRep(contador).tipo4 = rs_Confrep("conftipo4")
            Else
                confRep(contador).tipo4 = ""
            End If

            If Not EsNulo(rs_Confrep("conftipo5")) Then
                confRep(contador).tipo5 = rs_Confrep("conftipo5")
            Else
                confRep(contador).tipo5 = ""
            End If

            '----------------------------------------------------
            'VALORES
            '----------------------------------------------------
            If Not EsNulo(rs_Confrep("confval")) Then
                confRep(contador).Confval = rs_Confrep("confval")
            Else
                confRep(contador).Confval = ""
            End If
            If Not EsNulo(rs_Confrep("confval2")) Then
                confRep(contador).confval2 = rs_Confrep("confval2")
            Else
                confRep(contador).confval2 = ""
            End If
            
            If Not EsNulo(rs_Confrep("confval3")) Then
                confRep(contador).confval3 = rs_Confrep("confval3")
            Else
                confRep(contador).confval3 = ""
            End If
            
            If Not EsNulo(rs_Confrep("confval4")) Then
                confRep(contador).confval4 = rs_Confrep("confval4")
            Else
                confRep(contador).confval4 = ""
            End If

            If Not EsNulo(rs_Confrep("confval5")) Then
                confRep(contador).confval5 = rs_Confrep("confval5")
            Else
                confRep(contador).confval5 = ""
            End If

            
        End If
        rs_Confrep.MoveNext
    Loop
End If
rs_Confrep.Close
End Sub

