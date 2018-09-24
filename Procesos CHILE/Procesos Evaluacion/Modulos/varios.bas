Attribute VB_Name = "varios"
Option Explicit


Public Function Minimo(ByVal X, ByVal Y)
    If X <= Y Then
        Minimo = X
    Else
        Minimo = Y
    End If
End Function


Public Sub LimpiarTraza(ByVal Cabecera As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Limpia la Traza para un empleado/concepto.
' Autor      : FGZ
' Fecha      : 08/09/2003
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    
    StrSql = "DELETE FROM traza WHERE cliqnro = " & Cabecera
    'StrSql = "EXEC Eliminar_traza " & cabecera
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub InsertarTraza(ByVal cliqnro As Long, ByVal concepto As Long, ByVal Tpanro As Long, ByVal desc As String, ByVal Valor As Single)
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

ContadorProgreso = ContadorProgreso + 1
Aux_Texto = Left(desc, 60)
'StrSql = "SELECT * FROM traza " & _
'         " WHERE cliqnro = " & cliqnro & _
'         " AND concnro = " & concepto & _
'         " AND tpanro = " & tpanro
'OpenRecordset StrSql, rs_traza
'
'If rs_traza.EOF Then
    StrSql = "INSERT INTO traza (cliqnro,concnro,tpanro,tradesc,travalor,trafrecuencia)" & _
             " VALUES (" & cliqnro & _
             "," & concepto & _
             "," & Tpanro & _
             ",'" & Aux_Texto & _
             "'," & Valor & _
             ",'" & CStr(ContadorProgreso) & _
             "')"
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
    If Not EsNulo(Hasta) Then
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
Dim I As Integer
Dim CadenaAux As String
    
    CadenaAux = ""
    
    I = 1
    ch = Mid$(cadena, I, 1)
    I = I + 1
    
    Do Until I > Len(cadena)
         
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
        ch = Mid$(cadena, I, 1)
        I = I + 1
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



Public Function Espacios(ByVal Cantidad As Integer) As String
    Espacios = Space(Cantidad)
End Function

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

Public Sub AcotarStr(ByRef Str As String, ByVal Longitud As Integer, ByVal Completar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Restringe la cantidad de caracteres del string pasado como parametro
' Autor      : FGZ
' Fecha      : 09/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Str = Left(Str, Longitud)
    If Completar Then
        If Len(Str) < Longitud Then
            Str = Str & Space(Longitud - Len(Str))
        End If
    End If
End Sub

Public Function Format_Str(ByVal Str, ByVal Longitud As Integer, ByVal Completar As Boolean, ByVal Str_Completar As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Restringe la cantidad de caracteres del string pasado como parametro y lo completa
'              con el caracter pasado por parametro hasta la longitud (si completar es TRUE)
' Autor      : FGZ
' Fecha      : 12/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    If Not EsNulo(Str) Then
        Str = Left(Str, Longitud)
        If Completar Then
            If Len(Str) < Longitud Then
                Str = Str & String(Longitud - Len(Str), Str_Completar)
            End If
        End If
        Format_Str = Str
    Else
        If Completar Then
            Format_Str = String(Longitud, " ")
        Else
            Format_Str = ""
        End If
    End If
End Function

Public Function Format_StrNro(ByVal Str, ByVal Longitud As Integer, ByVal Completar As Boolean, ByVal Str_Completar As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Restringe la cantidad de caracteres del string pasado como parametro y lo completa
'              con el caracter pasado por parametro hasta la longitud (si completar es TRUE)
' Autor      : FGZ
' Fecha      : 12/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    If Not EsNulo(Str) Then
        Str = Left(Str, Longitud)
        If Completar Then
            If Len(Str) < Longitud Then
                Str = String(Longitud - Len(Str), Str_Completar) & Str
            End If
        End If
        Format_StrNro = Str
    Else
        Format_StrNro = ""
    End If
End Function

Public Function EsNulo(ByVal Objeto) As Boolean
    If IsNull(Objeto) Then
        EsNulo = True
    Else
        If UCase(Objeto) = "NULL" Or UCase(Objeto) = "" Then
            EsNulo = True
        Else
            EsNulo = False
        End If
    End If
End Function


Public Function FormatearParaSql(ByVal Str, ByVal Longitud As Integer, ByVal Izquierda As Boolean, ByVal Completar As Boolean) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Restringe la cantidad de caracteres del string pasado como parametro y completa para insertar en sql
' acotados desde la Izquierda o derecha, segun parametro y completa segun parametro
' Autor      : FGZ
' Fecha      : 28/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    If Not EsNulo(Str) Then
        If Completar Then
            If Len(Str) < Longitud Then
                If Izquierda Then
                    'completo con espacio a Derecha
                    Str = Str & Space(Longitud - Len(Str))
                Else
                    'completo con espacio a Izquierda
                    Str = Space(Longitud - Len(Str)) & Str
                End If
            Else
                If Izquierda Then
                    Str = Left(Str, Longitud)
                Else
                    Str = Right(Str, Longitud)
                End If
            End If
        Else
            If Izquierda Then
                Str = Left(Str, Longitud)
            Else
                Str = Right(Str, Longitud)
            End If
        End If
    Else
        Str = ""
    End If
    FormatearParaSql = "'" & Str & "'"
End Function


Public Function GetValor(ByVal Valor, ByVal Nulo As String)
    If EsNulo(Valor) Then
        GetValor = Nulo
    Else
        GetValor = Valor
    End If
End Function

Public Function GetFecha(ByVal Valor)
    If EsNulo(Valor) Then
        GetFecha = "NULL"
    Else
        GetFecha = ConvFecha(Valor)
    End If
End Function

Public Function GetString(ByVal campo)
  If Len(campo) <> 0 Then
     GetString = "'" & campo & "'"
  Else
     GetString = "NULL"
  End If
End Function 'getString(formName)

Public Function Format_Fecha(ByVal Str, ByVal Tipo As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Formatea una fecha de a cuerdo a un tipo/criterio
'
' Autor      : Scarpa D.
' Fecha      : 02/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Dim Salida As String
    Dim Fecha

    If Not EsNulo(Str) Then
       If Trim(Str) <> "" Then
            Select Case Tipo
               Case 1
                  Fecha = CDate(Str)
                  Salida = Year(Fecha) & Format_StrNro(Month(Fecha), 2, True, "0") & Format_StrNro(Day(Fecha), 2, True, "0")
               Case Else
                  Salida = Str
            End Select
            
            Format_Fecha = Salida
        Else
            Format_Fecha = ""
        End If
    Else
        Format_Fecha = ""
    End If
End Function

Public Function Cuil_Valido(ByVal Cuil As String, ByRef MensajeError As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Valida que el cuil sea correcto
'
' Autor      : FGZ
' Fecha      : 28/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valido As Boolean

Dim Totalsuma
Dim Digito
Dim Resto
Dim Numerototal
Dim Numero1
Dim Numero2
Dim Numero3
Dim N1
Dim N2
Dim N3
Dim N4
Dim N5
Dim N6
Dim N7
Dim N8
Dim N9
Dim N10
Dim Opcion

Valido = False
Opcion = ""


Numerototal = Cuil
Numero1 = Mid(Numerototal, 1, 2)
Numero2 = Mid(Numerototal, 4, 8)
Numero3 = Mid(Numerototal, 13, 1)

N1 = Mid(Numero1, 1, 1)
N2 = Mid(Numero1, 2, 1)

N3 = Mid(Numero2, 1, 1)
N4 = Mid(Numero2, 2, 1)
N5 = Mid(Numero2, 3, 1)
N6 = Mid(Numero2, 4, 1)
N7 = Mid(Numero2, 5, 1)
N8 = Mid(Numero2, 6, 1)
N9 = Mid(Numero2, 7, 1)
N10 = Mid(Numero2, 8, 1)

If Cuil = "" Then
    Opcion = ""
Else
    If Len(Numerototal) <> 13 Then
        Opcion = "El número de CUIL está mal ingresado, debe contener trece caracteres. "
    Else
        If Mid(Numerototal, 3, 1) <> "-" Then
            Opcion = "El tercer carácter debe ser un guión. "
        End If
        If Mid(Numerototal, 12, 1) <> "-" Then
            Opcion = "El decimosegundo carácter debe ser un guión. "
        End If
        If Not IsNumeric(Numero1) Then
            Opcion = "Los dos primeros dígitos deben ser numéricos. "
        End If
        If Not IsNumeric(Numero2) Then
            Opcion = "Los dígitos entre guiones deben ser numéricos. "
        End If
        If Not IsNumeric(Numero3) Then
            Opcion = "El último dígito debe ser numérico. "
        End If
    
        Totalsuma = N1 * 5 + N2 * 4 + N3 * 3 + N4 * 2 + N5 * 7 + N6 * 6 + N7 * 5 + N8 * 4 + N9 * 3 + N10 * 2
        Resto = Totalsuma Mod 11
        Select Case Resto
        Case 0:
            Digito = 0
        Case 1:
            Digito = 1
        Case Else
            Digito = 11 - Resto
        End Select
        
        If CInt(Numero3) <> CInt(Digito) Then
            Opcion = Opcion + ". El Digito verificador es incorrecto. "
        End If
    End If
End If
If (Opcion = "") Then
    Valido = True
Else
    Valido = False
End If
MensajeError = Opcion
Cuil_Valido = Valido
End Function



Public Function StrToFecha(ByVal Str As String, ByRef Ok As Boolean) 'As Date
Dim Fecha
Dim Dia As String
Dim Mes As String
Dim Anio As String

    Dia = Mid(Str, 7, 2)
    Mes = Mid(Str, 5, 2)
    Anio = Mid(Str, 1, 4)
    
    If Str = "99991231" Then
        Fecha = ""
        Ok = True
    Else
        If IsDate(Dia & "/" & Mes & "/" & Anio) Then
            Fecha = CDate(Dia & "/" & Mes & "/" & Anio)
            Ok = True
        Else
            Fecha = ""
            Ok = False
        End If
    End If
    StrToFecha = Fecha
End Function
