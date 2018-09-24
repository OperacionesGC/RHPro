Attribute VB_Name = "varios"
Option Explicit


Public Function Minimo(ByVal X, ByVal Y)
    If X <= Y Then
        Minimo = X
    Else
        Minimo = Y
    End If
End Function


Public Sub LimpiarTraza(ByVal cabecera As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Limpia la Traza para un empleado/concepto.
' Autor      : FGZ
' Fecha      : 08/09/2003
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    
    StrSql = "DELETE FROM traza WHERE cliqnro = " & cabecera
    'StrSql = "EXEC Eliminar_traza " & cabecera
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub InsertarTraza(ByVal cliqnro As Long, ByVal concepto As Long, ByVal tpanro As Long, ByVal Desc As String, ByVal valor As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Graba un registro de traza para un empleado/concepto. {Traza.i}
' Autor      : Lic.Mauricio Heidt
' Fecha      : 27/10/1996
' Traduccion : FGZ
' Fecha      : 05/09/2003
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_traza As New ADODB.Recordset
Dim Aux_Texto As String

Aux_Texto = Left(Desc, 60)
'StrSql = "SELECT * FROM traza " & _
'         " WHERE cliqnro = " & cliqnro & _
'         " AND concnro = " & concepto & _
'         " AND tpanro = " & tpanro
'OpenRecordset StrSql, rs_traza
'
'If rs_traza.EOF Then
    StrSql = "INSERT INTO traza (cliqnro,concnro,tpanro,tradesc,travalor)" & _
             " VALUES (" & cliqnro & _
             "," & concepto & _
             "," & tpanro & _
             ",'" & Aux_Texto & _
             "'," & valor & _
             ")"
    objConn.Execute StrSql, , adExecuteNoRecords
'End If

End Sub


Public Function CantidadDeDias(ByVal PeriodoDesde As Date, ByVal PeriodoHasta As Date, ByVal Desde As Date, ByVal Hasta As Date) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias que caen dentro de un periodo (especificado como un
'              rango de fechas) .
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim FechaInicioAuxiliar As Date
Dim FechaFinAuxiliar As Date
    
    FechaInicioAuxiliar = IIf(Desde > PeriodoDesde, Desde, PeriodoDesde)
    If Not IsNull(Hasta) Then
        FechaFinAuxiliar = Hasta
        FechaFinAuxiliar = IIf(Hasta > PeriodoHasta, PeriodoHasta, Hasta)
    Else
        'FechaFinAuxiliar = PeriodoHasta
        FechaFinAuxiliar = IIf(Hasta > PeriodoHasta, PeriodoHasta, Hasta)
    End If
    
    CantidadDeDias = DateDiff("d", FechaInicioAuxiliar, FechaFinAuxiliar) + 1

End Function



Public Function EliminarCHInvalidos(ByVal cadena As String) As String
Dim ch As String
Dim i As Integer
Dim CadenaAux As String
    
    CadenaAux = ""
    
    i = 1
    ch = Mid$(cadena, i, 1)
    i = i + 1
    
    Do Until i > Len(cadena)
         
        Select Case Asc(ch)
        Case 8: ' backspace
            ch = Chr(32)
        Case 9: ' Tab
            ch = Chr(32)
        Case 10: 'Nueva Linea
            ch = Chr(32)
        Case 12: 'Form Feed
            ch = Chr(32)
        Case 13: 'Retorno de Carro
            ch = Chr(32)
        Case 32: 'Espacio en Blanco
            'ch = Chr(32)
        Case Else: ' lo dejo como esta
        
        End Select
    
        CadenaAux = CadenaAux & ch
        ch = Mid$(cadena, i, 1)
        i = i + 1
    Loop

EliminarCHInvalidos = CadenaAux

End Function

Public Function Biciesto(ByVal Año As Integer) As Boolean
Dim dias As Integer
Dim DiaInicio As Date
Dim DiaFin As Date

DiaInicio = CDate("01/02/" & Año)
DiaFin = CDate("01/03/" & Año) - 1

dias = DateDiff("d", DiaInicio, DiaFin) + 1
If dias = 28 Then
    Biciesto = False
Else
    Biciesto = True
End If
End Function


Public Function IsEmptyRecordset(ByVal rs As Recordset) As Boolean
    IsEmptyRecordset = ((rs.BOF = True) And (rs.EOF = True))
End Function



'Public Function Espacios(ByVal Cantidad As Integer) As String
'    Espacios = Space(Cantidad)
'End Function

Public Function EnLetras(Numero As String) As String
    Dim b, paso As Integer
    Dim Expresion, entero, deci, flag As String

    flag = "N"
    For paso = 1 To Len(Numero)
        If Mid(Numero, paso, 1) = "." Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(Numero, paso, 1) 'Extae la parte entera del numero
            Else
                deci = deci + Mid(Numero, paso, 1) 'Extrae la parte decimal del numero
            End If
        End If
    Next paso

    If Len(deci) = 1 Then
        deci = deci & "0"
    End If

    flag = "N"
    If CLng(Numero) >= -999999999 And CLng(Numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
        For paso = Len(entero) To 1 Step -1
            b = Len(entero) - (paso - 1)
            Select Case paso
            Case 3, 6, 9
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then
                            Expresion = Expresion & "cien "
                        Else
                            Expresion = Expresion & "ciento "
                        End If
                    Case "2"
                        Expresion = Expresion & "doscientos "
                    Case "3"
                        Expresion = Expresion & "trescientos "
                    Case "4"
                        Expresion = Expresion & "cuatrocientos "
                    Case "5"
                        Expresion = Expresion & "quinientos "
                    Case "6"
                        Expresion = Expresion & "seiscientos "
                    Case "7"
                        Expresion = Expresion & "setecientos "
                    Case "8"
                        Expresion = Expresion & "ochocientos "
                    Case "9"
                        Expresion = Expresion & "novecientos "
                End Select

            Case 2, 5, 8
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" Then
                            flag = "S"
                            Expresion = Expresion & "diez "
                        End If
                        If Mid(entero, b + 1, 1) = "1" Then
                            flag = "S"
                            Expresion = Expresion & "once "
                        End If
                        If Mid(entero, b + 1, 1) = "2" Then
                            flag = "S"
                            Expresion = Expresion & "doce "
                        End If
                        If Mid(entero, b + 1, 1) = "3" Then
                            flag = "S"
                            Expresion = Expresion & "trece "
                        End If
                        If Mid(entero, b + 1, 1) = "4" Then
                            flag = "S"
                            Expresion = Expresion & "catorce "
                        End If
                        If Mid(entero, b + 1, 1) = "5" Then
                            flag = "S"
                            Expresion = Expresion & "quince "
                        End If
                        If Mid(entero, b + 1, 1) > "5" Then
                            flag = "N"
                            Expresion = Expresion & "dieci"
                        End If

                    Case "2"
                        If Mid(entero, b + 1, 1) = "0" Then
                            Expresion = Expresion & "veinte "
                            flag = "S"
                        Else
                            Expresion = Expresion & "veinti"
                            flag = "N"
                        End If

                    Case "3"
                        If Mid(entero, b + 1, 1) = "0" Then
                            Expresion = Expresion & "treinta "
                            flag = "S"
                        Else
                            Expresion = Expresion & "treinta y "
                            flag = "N"
                        End If

                    Case "4"
                        If Mid(entero, b + 1, 1) = "0" Then
                            Expresion = Expresion & "cuarenta "
                            flag = "S"
                        Else
                            Expresion = Expresion & "cuarenta y "
                            flag = "N"
                        End If

                    Case "5"
                        If Mid(entero, b + 1, 1) = "0" Then
                            Expresion = Expresion & "cincuenta "
                            flag = "S"
                        Else
                            Expresion = Expresion & "cincuenta y "
                            flag = "N"
                        End If

                    Case "6"
                        If Mid(entero, b + 1, 1) = "0" Then
                            Expresion = Expresion & "sesenta "
                            flag = "S"
                        Else
                            Expresion = Expresion & "sesenta y "
                            flag = "N"
                        End If

                    Case "7"
                        If Mid(entero, b + 1, 1) = "0" Then
                            Expresion = Expresion & "setenta "
                            flag = "S"
                        Else
                            Expresion = Expresion & "setenta y "
                            flag = "N"
                        End If

                    Case "8"
                        If Mid(entero, b + 1, 1) = "0" Then
                            Expresion = Expresion & "ochenta "
                            flag = "S"
                        Else
                            Expresion = Expresion & "ochenta y "
                            flag = "N"
                        End If

                    Case "9"
                        If Mid(entero, b + 1, 1) = "0" Then
                            Expresion = Expresion & "noventa "
                            flag = "S"
                        Else
                            Expresion = Expresion & "noventa y "
                            flag = "N"
                        End If
                End Select

            Case 1, 4, 7
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                Expresion = Expresion & "uno "
                            Else
                                Expresion = Expresion & "un "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            Expresion = Expresion & "dos "
                        End If
                    Case "3"
                        If flag = "N" Then
                            Expresion = Expresion & "tres "
                        End If
                    Case "4"
                        If flag = "N" Then
                            Expresion = Expresion & "cuatro "
                        End If
                    Case "5"
                        If flag = "N" Then
                            Expresion = Expresion & "cinco "
                        End If
                    Case "6"
                        If flag = "N" Then
                            Expresion = Expresion & "seis "
                        End If
                    Case "7"
                        If flag = "N" Then
                            Expresion = Expresion & "siete "
                        End If
                    Case "8"
                        If flag = "N" Then
                            Expresion = Expresion & "ocho "
                        End If
                    Case "9"
                        If flag = "N" Then
                            Expresion = Expresion & "nueve "
                        End If
                End Select
            End Select
            If paso = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 6) Then
                    Expresion = Expresion & "mil "
                End If
            End If
            If paso = 7 Then
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    Expresion = Expresion & "millón "
                Else
                    Expresion = Expresion & "millones "
                End If
            End If
        Next paso

        If deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & Expresion & "con " & deci ' & "/100"
            Else
                EnLetras = Expresion & "con " & deci ' & "/100"
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & Expresion
            Else
                EnLetras = Expresion
            End If
        End If
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        EnLetras = ""
    End If
End Function


Public Sub BusMes(ByVal Mes As Integer, ByRef Des_Mes As String)
Select Case Mes
    Case 1:     Des_Mes = "Enero"
    Case 2:     Des_Mes = "Febrero"
    Case 3:     Des_Mes = "Marzo"
    Case 4:     Des_Mes = "Abril"
    Case 5:     Des_Mes = "Mayo"
    Case 6:     Des_Mes = "junio"
    Case 7:     Des_Mes = "Julio"
    Case 8:     Des_Mes = "Agosto"
    Case 9:     Des_Mes = "Septiembre"
    Case 10:    Des_Mes = "Octubre"
    Case 11:    Des_Mes = "Noviembre"
    Case 12:    Des_Mes = "Diciembre"
End Select
End Sub

