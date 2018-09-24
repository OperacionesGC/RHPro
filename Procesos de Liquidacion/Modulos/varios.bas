Attribute VB_Name = "varios"
Option Explicit


Public Function Minimo(ByVal X, ByVal Y)
    If X <= Y Then
        Minimo = X
    Else
        Minimo = Y
    End If
End Function

Public Function Maximo(ByVal X, ByVal Y)
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve el mayor de los dos valores dados como parametro
' Autor      : GdeCos
' Fecha      : 3/05/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
    If X >= Y Then
        Maximo = X
    Else
        Maximo = Y
    End If

End Function
'
'
'Public Sub LimpiarTraza(ByVal Cabecera As Long)
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Limpia la Traza para un empleado/concepto.
'' Autor      : FGZ
'' Fecha      : 08/09/2003
'' Ultima Mod :
'' Descripcion:
'' ---------------------------------------------------------------------------------------------
'
'    If ReusaTraza Then
'
'        StrSql = "UPDATE traza SET cliqnro = -1, travalor = 0"
'        StrSql = StrSql & " ,concnro = -1 ,tpanro = -1, tradesc = NULL"
'        StrSql = StrSql & " WHERE cliqnro = " & Cabecera
'
'    Else
'
'        StrSql = "DELETE FROM traza WHERE cliqnro = " & Cabecera
'
'    End If
'
'    objConn.Execute StrSql, , adExecuteNoRecords
'
'End Sub
'
'
'Public Sub InsertarTraza(ByVal cliqnro As Long, ByVal concepto As Long, ByVal tpanro As Long, ByVal desc As String, ByVal Valor As Double)
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Graba un registro de traza para un empleado/concepto. {Traza.i}
'' Autor      : Lic.Mauricio RHPro
'' Fecha      : 27/10/1996
'' Traduccion : FGZ
'' Fecha      : 05/09/2003
'' Ultima Mod :
'' Descripcion:
'' ---------------------------------------------------------------------------------------------
'Dim rs_Traza As New ADODB.Recordset
'Dim Aux_Texto As String
'
'On Error GoTo MLocal:
'
'    Aux_Texto = Left(desc, 60)
'
'    If ReusaTraza Then
'
'        'Busco una traza libre
'        StrSql = "SELECT tranro FROM traza WHERE cliqnro = -1 AND travalor = 0"
'        StrSql = StrSql & " AND concnro = -1 AND tpanro = -1 AND tradesc IS NULL"
'        OpenRecordset StrSql, rs_Traza
'
'        If Not rs_Traza.EOF Then
'            'Actualizo el registro encontrado
'            StrSql = "UPDATE traza"
'            StrSql = StrSql & " SET cliqnro = " & cliqnro
'            StrSql = StrSql & " ,concnro = " & concepto
'            StrSql = StrSql & " ,tpanro = " & tpanro
'            StrSql = StrSql & " ,tradesc = '" & Aux_Texto & "'"
'            StrSql = StrSql & " ,travalor = " & Valor
'            StrSql = StrSql & " WHERE tranro = " & rs_Traza!tranro
'            objConn.Execute StrSql, , adExecuteNoRecords
'        Else
'            'Creo un nuevo registro
'            StrSql = "INSERT INTO traza (cliqnro,concnro,tpanro,tradesc,travalor,trafrecuencia)"
'            StrSql = StrSql & " VALUES (" & cliqnro
'            StrSql = StrSql & " ," & concepto
'            StrSql = StrSql & " ," & tpanro
'            StrSql = StrSql & " ,'" & Aux_Texto
'            StrSql = StrSql & " '," & Valor
'            StrSql = StrSql & " ,'0000000')"
'            objConn.Execute StrSql, , adExecuteNoRecords
'        End If
'
'    Else
'
'        StrSql = "INSERT INTO traza (cliqnro,concnro,tpanro,tradesc,travalor,trafrecuencia)" & _
'                 " VALUES (" & cliqnro & _
'                 "," & concepto & _
'                 "," & tpanro & _
'                 ",'" & Aux_Texto & _
'                 "'," & Valor & _
'                 ",'" & Format(ContadorProgreso, "0000000") & _
'                 "')"
'        objConn.Execute StrSql, , adExecuteNoRecords
'        ContadorProgreso = ContadorProgreso + 1
'
'    End If
'
'    If rs_Traza.State = adStateOpen Then rs_Traza.Close
'    Set rs_Traza = Nothing
'
'Exit Sub
'MLocal:
''        Flog.Writeline
''        Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
''        Flog.Writeline Espacios(Tabulador * 0) & " Error insertando traza "
''        Flog.Writeline Espacios(Tabulador * 0) & " Error: " & Err.Description
''        Flog.Writeline
''        Flog.Writeline Espacios(Tabulador * 0) & "Ultimo SQL Ejecutado: " & StrSql
''        Flog.Writeline
''        Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
''        Flog.Writeline
'End Sub


Public Function CantidadDeDias(ByVal PeriodoDesde As Date, ByVal PeriodoHasta As Date, ByVal Desde As Date, ByVal Hasta As Date) As Long
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

    CantidadDeDias = IIf((DateDiff("d", FechaInicioAuxiliar, FechaFinAuxiliar) + 1 < 0), 0, DateDiff("d", FechaInicioAuxiliar, FechaFinAuxiliar) + 1)

End Function

Public Function CantidadDeDiasSykes(ByVal PeriodoDesde As Date, ByVal PeriodoHasta As Date, ByRef Desde As Date, ByRef Hasta As Date) As Long
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

    If Desde < PeriodoDesde And Hasta > PeriodoDesde Then
        Hasta = DateAdd("d", -1, PeriodoDesde)
        CantidadDeDiasSykes = 0
    Else
        CantidadDeDiasSykes = IIf((DateDiff("d", FechaInicioAuxiliar, FechaFinAuxiliar) + 1 < 0), 0, DateDiff("d", FechaInicioAuxiliar, FechaFinAuxiliar) + 1)
    End If
End Function

Public Function CantidadDeDiasSykesSV(ByVal PeriodoDesde As Date, ByVal PeriodoHasta As Date, ByRef Desde As Date, ByRef Hasta As Date) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Verifica si al menos una parte de las fechas son retroactivas (fecha desde). y que la fecha hasta este dentro del proceso de liquidacion.
'              rango de fechas) .
' Autor      : EAM 26/01/2015
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

'Comprueba que la fecha desde este fuera del proceso de liquidacion
If (DateDiff("d", PeriodoDesde, Desde) + 1 <= 0) Then
    'Comprueba que la fecha hasta este dentro del proceso de liquidaci�n
    If (DateDiff("d", PeriodoHasta, Hasta) <= 0) Then
        CantidadDeDiasSykesSV = 0
    Else
        CantidadDeDiasSykesSV = 1
    End If
Else
    CantidadDeDiasSykesSV = 1
End If

End Function

Public Function CantidadDeDiasHab(ByVal Ternro, ByVal PeriodoDesde As Date, ByVal PeriodoHasta As Date, ByVal Desde As Date, ByVal Hasta As Date) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias que caen dentro de un periodo (especificado como un
'              rango de fechas) .
' Autor      : FGZ
' Fecha      :
' Ultima Mod.: FGZ - 22/11/2012
' Descripcion: Estaba mal definido el objeto feriado (faltaba el new)
' ---------------------------------------------------------------------------------------------
Dim FechaInicioAuxiliar As Date
Dim FechaFinAuxiliar As Date
Dim Actual As Date
Dim aux As Long

Dim objFeriado As New Feriado
Dim empl As Long

    empl = Ternro
    
    FechaInicioAuxiliar = IIf(Desde > PeriodoDesde, Desde, PeriodoDesde)
    If Not EsNulo(Hasta) Then
        FechaFinAuxiliar = Hasta
        FechaFinAuxiliar = IIf(Hasta > PeriodoHasta, PeriodoHasta, Hasta)
    Else
        'FechaFinAuxiliar = PeriodoHasta
        FechaFinAuxiliar = IIf(Hasta > PeriodoHasta, PeriodoHasta, Hasta)
    End If
    
    aux = 0
    Actual = Desde
    Do While DateDiff("d", Actual, Hasta) >= 0
        
        If (Weekday(Actual) <> 1) And (Weekday(Actual) <> 7) Then
            'No es sabado ni domingo, verificar si es feriado
            If objFeriado.Feriado(Actual, empl, False) Then
              aux = aux + 1
            End If
        End If
        
        Actual = DateAdd("d", 1, Actual)
    
    Loop

    CantidadDeDiasHab = aux

End Function


Public Function CantidadDeHorasDias(ByVal PeriodoDesde As Date, ByVal PeriodoHasta As Date, ByVal Desde As Date, ByVal Hasta As Date, ByVal habilConf As String, ByVal noHabilConf As String, ByVal feriadoConf As String) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de horas que caen dentro de un periodo (especificado como un
'              rango de fechas) segun sea habil, no habil o feriado.
' Autor      : Martin Ferraro
' Fecha      : 17/12/2007
' Ultima Mod.: FGZ - 22/11/2012
' Descripcion: Estaba mal definido el objeto feriado (faltaba el new)
' ---------------------------------------------------------------------------------------------
Dim FechaInicioAuxiliar As Date
Dim FechaFinAuxiliar As Date
Dim CantidadDeDias As Long
Dim fechaFor As Date
Dim multiplicador As Long
Dim acumulado As Long
'Dim objFeriado As Feriado
Dim objFeriado As New Feriado

    'Me aseguro que el formato sea de 7 caracteres
    If Len(Trim(habilConf)) < 7 Then
        habilConf = habilConf & String(7 - Len(Trim(habilConf)), "0")
    End If
    If Len(Trim(noHabilConf)) < 7 Then
        noHabilConf = noHabilConf & String(7 - Len(Trim(noHabilConf)), "0")
    End If
    If Len(Trim(feriadoConf)) < 7 Then
        feriadoConf = feriadoConf & String(7 - Len(Trim(feriadoConf)), "0")
    End If
    
    'El formato es de lunes a domingo (LMMJVSD) paso el domingo al principio
    habilConf = Right(habilConf, 1) & Left(habilConf, 6)
    noHabilConf = Right(noHabilConf, 1) & Left(noHabilConf, 6)
    feriadoConf = Right(feriadoConf, 1) & Left(feriadoConf, 6)

    'Rangos de fechas  a tener en cuenta
    FechaInicioAuxiliar = IIf(Desde > PeriodoDesde, Desde, PeriodoDesde)
    If Not EsNulo(Hasta) Then
        FechaFinAuxiliar = Hasta
        FechaFinAuxiliar = IIf(Hasta > PeriodoHasta, PeriodoHasta, Hasta)
    Else
        FechaFinAuxiliar = IIf(Hasta > PeriodoHasta, PeriodoHasta, Hasta)
    End If
    
    acumulado = 0
    
    For fechaFor = FechaInicioAuxiliar To FechaFinAuxiliar
            
        multiplicador = 0
        
        If objFeriado.Feriado(fechaFor, buliq_empleado!Ternro, False) Then
            multiplicador = Mid(feriadoConf, Weekday(fechaFor), 1)
        Else
            'De lunes a viernes se considera habil y de sabado a domingo no habil
            Select Case Weekday(fechaFor)
                Case 2 To 6
                    multiplicador = Mid(habilConf, Weekday(fechaFor), 1)
                Case 1, 7
                    multiplicador = Mid(noHabilConf, Weekday(fechaFor), 1)
            End Select
        End If
        
        acumulado = acumulado + multiplicador
            
    Next fechaFor
    
    CantidadDeHorasDias = acumulado
    
End Function


Public Function Dias_Licencias_Ya_Marcados(ByVal FechaDeInicio As Date, ByVal FechaDeFin As Date, ByVal Desde As Date, ByVal Hasta As Date, ByVal Proceso As Long, ByVal Tercero As Long) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias de licencia que ya fueron marcados con el nro de proceso pasado como parametro
'              que caen dentro de un periodo (especificado como un rango de fechas) .
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Lic As New ADODB.Recordset
Dim Dias As Long

    Dias = 0

    StrSql = "SELECT elfechadesde, elfechahasta FROM emp_lic WHERE (empleado = " & Tercero & " )"
    StrSql = StrSql & " AND pronro = " & Proceso
    StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(FechaDeFin)
    StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(FechaDeInicio)
    OpenRecordset StrSql, rs_Lic
        
    Do While Not rs_Lic.EOF
        Dias = Dias + CantidadDeDias(FechaDeInicio, FechaDeFin, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
        rs_Lic.MoveNext
    Loop
    Dias_Licencias_Ya_Marcados = Dias
    
    'cierro
    If rs_Lic.State = adStateOpen Then rs_Lic.Close
    Set rs_Lic = Nothing
End Function


Public Function Dias_Licencias_Total(ByVal FechaDeInicio As Date, ByVal FechaDeFin As Date, ByVal Desde As Date, ByVal Hasta As Date, ByVal Tercero As Long) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad total de dias de licencia
'              que caen dentro de un periodo (especificado como un rango de fechas) .
' Autor      : FGZ
' Fecha      : 21/10/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Lic As New ADODB.Recordset
Dim Dias As Long

    Dias = 0

    StrSql = "SELECT elfechadesde, elfechahasta FROM emp_lic WHERE (empleado = " & Tercero & " )"
    StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(FechaDeFin)
    StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(FechaDeInicio)
    OpenRecordset StrSql, rs_Lic
        
    Do While Not rs_Lic.EOF
        Dias = Dias + CantidadDeDias(FechaDeInicio, FechaDeFin, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
        rs_Lic.MoveNext
    Loop
    Dias_Licencias_Total = Dias
    
    'cierro
    If rs_Lic.State = adStateOpen Then rs_Lic.Close
    Set rs_Lic = Nothing
End Function

Public Function Esta_de_Licencia(ByVal Fecha As Date, ByVal Tercero As Long) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna TRUE si el dia de la fecha esta de Licencia. Sino FALSE.
' Autor      : FGZ
' Fecha      : 31/10/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Lic As New ADODB.Recordset

    'StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & Tercero & " )"
    StrSql = "SELECT empleado FROM emp_lic WHERE (empleado = " & Tercero & " )"
    StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(Fecha)
    StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(Fecha)
    OpenRecordset StrSql, rs_Lic
    Esta_de_Licencia = Not rs_Lic.EOF
    
    'cierro
    If rs_Lic.State = adStateOpen Then rs_Lic.Close
    Set rs_Lic = Nothing
End Function

Public Function Esta_de_Licencia_Tipo(ByVal Fecha As Date, ByVal Tercero As Long, ByVal Tipos As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna TRUE si el dia de la fecha esta de Licencia de algunos de los tipos . Sino FALSE.
' Autor      : FGZ
' Fecha      : 15/05/2013
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Lic As New ADODB.Recordset

    StrSql = "SELECT empleado FROM emp_lic WHERE (empleado = " & Tercero & " )"
    StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(Fecha)
    StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND tdnro IN (" & Tipos & ")"
    OpenRecordset StrSql, rs_Lic
    Esta_de_Licencia_Tipo = Not rs_Lic.EOF
    
    'cierro
    If rs_Lic.State = adStateOpen Then rs_Lic.Close
    Set rs_Lic = Nothing
End Function



Public Function EliminarCHInvalidos(ByVal cadena As String) As String
Dim ch As String
Dim i As Long
Dim cadenaAux As String
    
    cadenaAux = ""
    
    i = 1
    ch = Mid$(cadena, i, 1)
    i = i + 1
    
    Do Until i > Len(cadena) + 1
         
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
        Case 39: 'comilla simple
            ch = Chr(180)
        Case Else: ' lo dejo como esta
        
        End Select
    
        cadenaAux = cadenaAux & ch
        ch = Mid$(cadena, i, 1)
        i = i + 1
    Loop

EliminarCHInvalidos = cadenaAux

End Function




Public Function Biciesto(ByVal A�o As Long) As Boolean
Dim Dias As Long
Dim DiaInicio As Date
Dim DiaFin As Date

DiaInicio = C_Date("01/02/" & A�o)
DiaFin = C_Date("01/03/" & A�o) - 1

Dias = DateDiff("d", DiaInicio, DiaFin) + 1
If Dias = 28 Then
    Biciesto = False
Else
    Biciesto = True
End If
End Function


Public Function IsEmptyRecordset(ByVal rs As Recordset) As Boolean
    IsEmptyRecordset = ((rs.BOF = True) And (rs.EOF = True))
End Function



Public Function Espacios(ByVal Cantidad As Long) As String
    Espacios = Space(Cantidad)
End Function

Public Function EnLetras(Numero As String) As String
    Dim b, paso As Long
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
                    Expresion = Expresion & "mill�n "
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


Public Sub BusMes(ByVal Mes As Long, ByRef Des_Mes As String)
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

'Public Sub PeriodoAnterior(ByVal Mes As Long, ByRef Des_Mes As Integer)
'Select Case Mes
'    Case 1:     Des_Mes = 12
'    Case 2:     Des_Mes = 1
'    Case 3:     Des_Mes = 2
'    Case 4:     Des_Mes = 3
'    Case 5:     Des_Mes = 4
'    Case 6:     Des_Mes = 5
'    Case 7:     Des_Mes = 6
'    Case 8:     Des_Mes = 7
'    Case 9:     Des_Mes = 8
'    Case 10:    Des_Mes = 9
'    Case 11:    Des_Mes = 10
'    Case 12:    Des_Mes = 11
'End Select
'End Sub
Public Sub PeriodosAnteriores(ByVal MesActual As Integer, ByVal AnioActual As Integer, ByRef Mes1 As Integer, ByRef Anio1 As Integer, ByRef Mes2 As Integer, ByRef Anio2 As Integer, ByRef Mes3 As Integer, ByRef Anio3 As Integer)

If MesActual = 1 Then
    Mes3 = 12
    Anio3 = AnioActual - 1
    Mes2 = 11
    Anio2 = AnioActual - 1
    Mes1 = 10
    Anio1 = AnioActual - 1
  Else
     If MesActual = 2 Then
        Mes3 = 1
        Anio3 = AnioActual
        Mes2 = 12
        Anio2 = AnioActual - 1
        Mes1 = 11
        Anio1 = AnioActual - 1
     Else
        If MesActual = 3 Then
            Mes3 = 2
            Anio3 = AnioActual
            Mes2 = 1
            Anio2 = AnioActual
            Mes1 = 12
            Anio1 = AnioActual - 1
        Else
            Mes3 = MesActual - 1
            Anio3 = AnioActual
            Mes2 = MesActual - 2
            Anio2 = AnioActual
            Mes1 = MesActual - 3
            Anio1 = AnioActual
        End If
    End If
  End If
End Sub


Public Sub AcotarStr(ByRef Str As String, ByVal longitud As Long, ByVal Completar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Restringe la cantidad de caracteres del string pasado como parametro
' Autor      : FGZ
' Fecha      : 09/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Str = Left(Str, longitud)
    If Completar Then
        If Len(Str) < longitud Then
            Str = Str & Space(longitud - Len(Str))
        End If
    End If
End Sub

Public Function Format_Str(ByVal Str, ByVal longitud As Long, ByVal Completar As Boolean, ByVal Str_Completar As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Restringe la cantidad de caracteres del string pasado como parametro y lo completa
'              con el caracter pasado por parametro hasta la longitud (si completar es TRUE)
' Autor      : FGZ
' Fecha      : 12/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    If Not EsNulo(Str) Then
        Str = Left(Str, longitud)
        If Completar Then
            If Len(Str) < longitud Then
                Str = Str & String(longitud - Len(Str), Str_Completar)
            End If
        End If
        Format_Str = Str
    Else
        If Completar Then
            Format_Str = String(longitud, " ")
        Else
            Format_Str = ""
        End If
    End If
End Function

Public Function Format_StrNro(ByVal Str, ByVal longitud As Long, ByVal Completar As Boolean, ByVal Str_Completar As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Restringe la cantidad de caracteres del string pasado como parametro y lo completa
'              con el caracter pasado por parametro hasta la longitud (si completar es TRUE)
' Autor      : FGZ
' Fecha      : 12/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    If Not EsNulo(Str) Then
        Str = Left(Str, longitud)
        If Completar Then
            If Len(Str) < longitud Then
                Str = String(longitud - Len(Str), Str_Completar) & Str
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

Public Function formatoNumero(Numero As Double, decimales As Integer) As Double
    'LM - devuelve un numero con la cantidad de decimales
    formatoNumero = Fix(Numero * (10 ^ decimales)) / (10 ^ decimales)

End Function

Public Function FormatearParaSql(ByVal Str, ByVal longitud As Long, ByVal Izquierda As Boolean, ByVal Completar As Boolean) As String
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
            If Len(Str) < longitud Then
                If Izquierda Then
                    'completo con espacio a Derecha
                    Str = Str & Space(longitud - Len(Str))
                Else
                    'completo con espacio a Izquierda
                    Str = Space(longitud - Len(Str)) & Str
                End If
            Else
                If Izquierda Then
                    Str = Left(Str, longitud)
                Else
                    Str = Right(Str, longitud)
                End If
            End If
        Else
            If Izquierda Then
                Str = Left(Str, longitud)
            Else
                Str = Right(Str, longitud)
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

Public Function GetString(ByVal Campo)
  If Len(Campo) <> 0 Then
     GetString = "'" & Campo & "'"
  Else
     GetString = "NULL"
  End If
End Function 'getString(formName)

Public Function Format_Fecha(ByVal Str, ByVal tipo As Long)
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
            Select Case tipo
               Case 1
                  Fecha = C_Date(Str)
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

Public Function Cuil_Valido_old(ByVal Cuil As String, ByRef MensajeError As String) As Boolean
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
        Opcion = "El n�mero de CUIL est� mal ingresado, debe contener trece caracteres. "
    Else
        If Mid(Numerototal, 3, 1) <> "-" Then
            Opcion = "El tercer car�cter debe ser un gui�n. "
        End If
        If Mid(Numerototal, 12, 1) <> "-" Then
            Opcion = "El decimosegundo car�cter debe ser un gui�n. "
        End If
        If Not IsNumeric(Numero1) Then
            Opcion = "Los dos primeros d�gitos deben ser num�ricos. "
        End If
        If Not IsNumeric(Numero2) Then
            Opcion = "Los d�gitos entre guiones deben ser num�ricos. "
        End If
        If Not IsNumeric(Numero3) Then
            Opcion = "El �ltimo d�gito debe ser num�rico. "
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
        
        If CLng(Numero3) <> CLng(Digito) Then
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
Cuil_Valido_old = Valido
End Function



Public Function StrToFecha(ByVal Str As String, ByRef OK As Boolean) 'As Date
Dim Fecha
Dim Dia As String
Dim Mes As String
Dim Anio As String

If Not EsNulo(Trim(Str)) Then
    Dia = Mid(Str, 7, 2)
    Mes = Mid(Str, 5, 2)
    Anio = Mid(Str, 1, 4)
    
    If Str = "99991231" Then
        Fecha = ""
        OK = True
    Else
        If IsDate(Dia & "/" & Mes & "/" & Anio) Then
            Fecha = C_Date(Dia & "/" & Mes & "/" & Anio)
            OK = True
        Else
            Fecha = ""
            OK = False
        End If
    End If
    StrToFecha = Fecha
Else
    Fecha = ""
    OK = True
End If


End Function


Public Function New_Generar_Cuil(ByVal tipo As String, ByVal NumDoc As String, ByVal Hombre As Boolean) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Genra el nro de cuil
' Autor      : FGZ
' Fecha      : 25/04/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Aux_nro As String
Dim Pre As String
Dim Valido As Boolean
Dim i As Long
Dim Mensaje As String

    If Len(NumDoc) < 8 Then
        NumDoc = String(8 - Len(NumDoc), "0") & NumDoc
    End If
    If Hombre Then
        If tipo = "LE" Then
            Pre = "20"
        Else
            Pre = "20"
        End If
    Else
        If tipo = "LC" Then
            Pre = "23"
        Else
            Pre = "27"
        End If
    End If
    
    Aux_nro = Pre & NumDoc
    i = 0
    Valido = False
    Do While i <= 9 And Not Valido
        If Cuil_Valido(Aux_nro & i, Mensaje) Then
            Valido = True
            Aux_nro = Aux_nro & i
        End If
        i = i + 1
    Loop
    If Valido Then
        New_Generar_Cuil = Aux_nro
    Else
        New_Generar_Cuil = 0
    End If
End Function

Public Function Generar_Cuil(ByVal NumDoc As String, ByVal Hombre As Boolean) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Genra el nro de cuil
' Autor      : FGZ
' Fecha      : 25/04/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Aux_nro As String
Dim Pre As String
Dim Valido As Boolean
Dim i As Long
Dim Mensaje As String

    If Len(NumDoc) < 8 Then
        NumDoc = String(8 - Len(NumDoc), "0") & NumDoc
    End If
    If Hombre Then
        Pre = "20"
    Else
        Pre = "27"
    End If
    
    Aux_nro = Pre & NumDoc
    i = 0
    Valido = False
    Do While i <= 9 And Not Valido
        If Cuil_Valido(Aux_nro & i, Mensaje) Then
            Valido = True
            Aux_nro = Aux_nro & i
        End If
        i = i + 1
    Loop
    If Valido Then
        Generar_Cuil = Aux_nro
    Else
        Generar_Cuil = 0
    End If
End Function

Public Function Generar_Rut_Uruguay(ByVal NumDoc As String, ByVal Hombre As Boolean) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Genra el nro de rut
' Autor      : JMH
' Fecha      : 07/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

        Generar_Rut_Uruguay = NumDoc
End Function

Public Function Generar_Rut_Chile(ByVal NumDoc As String, ByVal Hombre As Boolean) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Genra el nro de rut
' Autor      : JMH
' Fecha      : 07/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

        Generar_Rut_Chile = NumDoc
End Function


Public Function Cuil_Valido(ByVal strCUIL As String, ByRef MensajeError As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Valida el Nro de CUIL
' Autor      : FGZ
' Fecha      : 25/04/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'N�meros por los que hay que multiplicar cada d�gito del CUIL
Const FACTORES = "54327654321"
Dim lngSuma As Long
Dim i As Long
Dim Valido As Boolean
lngSuma = 0
    strCUIL = Replace(strCUIL, "-", "")
    Valido = False
    If Len(strCUIL) = 11 Then
        If IsNumeric(strCUIL) Then
            For i = 1 To Len(strCUIL) '11
                lngSuma = lngSuma + (CLng(Mid(strCUIL, i, 1)) * CLng(Mid(FACTORES, i, 1)))
            Next i
            Valido = (lngSuma Mod Len(strCUIL) = 0) '11 = 0)
        End If
    Else
        MensajeError = "El cuil debe tener 11 d�gitos"
    End If
    Cuil_Valido = Valido
End Function

Public Function Rut_Valido_Uruguay(ByVal strCUIL As String, ByRef MensajeError As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Valida el Nro de RUT
' Autor      : JMH
' Fecha      : 07/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valido As Boolean

    Valido = True

    Rut_Valido_Uruguay = Valido
    
End Function

Public Function Rut_Valido_Chile(ByVal strCUIL As String, ByRef MensajeError As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Valida el Nro de RUT
' Autor      : JMH
' Fecha      : 07/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valido As Boolean

    Valido = True

    Rut_Valido_Chile = Valido
    
End Function


Public Function EsUltimoRegistro(ByRef Reg As ADODB.Recordset) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve TRUE si es el ultimo registro del recordset
' Autor      : FGZ
' Fecha      : 17/06/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Reg.MoveNext
    If Reg.EOF Then
        EsUltimoRegistro = True
    Else
        EsUltimoRegistro = False
    End If
    Reg.MovePrevious
End Function

Public Function EsUltimoRegistroCuenta(ByRef Reg As ADODB.Recordset, ByVal Linea As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve TRUE si es el ultimo registro del recordset o el ultimo registro con la misma cuenta
' Autor      : FGZ
' Fecha      : 17/06/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Reg.MoveNext
    If Reg.EOF Then
        EsUltimoRegistroCuenta = True
    Else
        If Reg!Linea = Linea Then
            EsUltimoRegistroCuenta = False
        Else
            EsUltimoRegistroCuenta = True
        End If
    End If
    Reg.MovePrevious
End Function







Public Function HorasInterseccion(ByVal R1FD As Date, ByVal R1HD As String, ByVal R1FH As Date, ByVal R1HH As String, ByVal R2FD As Date, ByVal R2HD As String, ByVal R2FH As Date, ByVal R2HH As String) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de horas que caen en la interseccion de 2 rangos de pares fecha-hora
'       R1FD -----------------R1FH
'       R2FD -----------------R2FH
' Parametros entrada
'              R1FD --> Fecha desde del rango 1
'              R1FH --> Fecha hasta del rango 1
'              R2FD --> Fecha desde del rango 2
'              R2FH --> Fecha hasta del rango 2

'              R1HD --> Hora desde del rango 1
'              R1HH --> Hora hasta del rango 1
'              R2HD --> Hora desde del rango 2
'              R2HH --> Hora hasta del rango 2
' Autor      : FGZ
' Fecha      : 07/11/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim TotHor As Single
Dim Tdias As Integer
Dim Thoras As Integer
Dim Tmin As Integer

TotHor = 0

'Rango1     [---------------]
'Rango2           (----------------)
If Menor_Hora(R1FD, R1HD, R2FD, R2HD) And Menor_Hora(R2FD, R2HD, R1FH, R1HH) And Menor_Igual_Hora(R1FH, R1HH, R2FH, R2HH) Then
    RestaHs R2FD, R2HD, R1FH, R1HH, Tdias, Thoras, Tmin
    'HCdesde = R2HD
    'HChasta = R1HH
    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
End If

'Rango1         [---------------]
'Rango2     (----------------)
If Mayor_Igual_Hora(R1FD, R1HD, R2FD, R2HD) And Menor_Hora(R1FD, R1HD, R2FH, R2HH) And Mayor_Hora(R1FH, R1HH, R2FH, R2HH) Then
    RestaHs R1FD, R1HD, R2FH, R2HH, Tdias, Thoras, Tmin
    'HCdesde = R1HD
    'HChasta = R2HH
    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
End If

'rango1         [---------------]
'rango2     (-----------------------)
If Mayor_Igual_Hora(R1FD, R1HD, R2FD, R2HD) And Menor_Hora(R1FD, R1HD, R2FH, R2HH) And Menor_Igual_Hora(R1FH, R1HH, R2FH, R2HH) And Mayor_Hora(R1FH, R1HH, R2FD, R2HD) Then
    RestaHs R1FD, R1HD, R1FH, R1HH, Tdias, Thoras, Tmin
    'HCdesde = R1HD
    'HChasta = R1HH
    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
End If

'rango1     [---------------]
'rango2         (--------)
If Menor_Hora(R1FD, R1HD, R2FD, R2HD) And Mayor_Hora(R1FH, R1HH, R2FH, R2HH) Then
    RestaHs R2FD, R2HD, R2FH, R2HH, Tdias, Thoras, Tmin
    'HCdesde = R2HD
    'HChasta = R2HH
    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
End If
   
HorasInterseccion = TotHor
    
End Function


Public Function Mayor_Hora(ByVal F1 As Date, ByVal H1 As String, ByVal F2 As Date, ByVal H2 As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna TRUE si la hora 1 es > que la hora 2.
' Autor      : FGZ
' Fecha      : 26/10/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------

    Mayor_Hora = Not Menor_Igual_Hora(F1, H1, F2, H2)
End Function

Public Function Menor_Hora(ByVal F1 As Date, ByVal H1 As String, ByVal F2 As Date, ByVal H2 As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna TRUE si la hora 1 es > que la hora 2.
' Autor      : FGZ
' Fecha      : 26/10/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------

    Menor_Hora = Not Mayor_Igual_Hora(F1, H1, F2, H2)
End Function

Public Function Mayor_Igual_Hora(ByVal F1 As Date, ByVal H1 As String, ByVal F2 As Date, ByVal H2 As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna TRUE si la hora 1 es >= que la hora 2.
' Autor      : FGZ
' Fecha      : 26/10/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim R As Boolean

    If F1 > F2 Then
        R = True
    Else
        If F1 < F2 Then
            R = False
        Else
            If H1 >= H2 Then
                R = True
            Else
                R = False
            End If
        End If
    End If
    
    Mayor_Igual_Hora = R
End Function

Public Function Menor_Igual_Hora(ByVal F1 As Date, ByVal H1 As String, ByVal F2 As Date, ByVal H2 As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna TRUE si la hora 1 es <= que la hora 2.
' Autor      : FGZ
' Fecha      : 26/10/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim R As Boolean

    If F1 < F2 Then
        R = True
    Else
        If F1 > F2 Then
            R = False
        Else
            If H1 <= H2 Then
                R = True
            Else
                R = False
            End If
        End If
    End If
    
    Menor_Igual_Hora = R
End Function


Public Sub RestaHs(Fecha_Inicio As Date, hora_inicio As String, Fecha_Fin As Date, hora_fin As String, ByRef TotDias, ByRef tothoras As Integer, ByRef TotMin As Integer)
Dim total As Integer
Dim cantdias  As Integer
Dim canthoras As Integer
Dim Dia   As Integer '  cantidad de minutos en un dia
Dim Hora As Integer   ' cantidad de minutos en una hora

    Dia = 1440
    Hora = 60
    canthoras = 0
    If Not EsNulo(hora_fin) And Not EsNulo(hora_inicio) Then
        canthoras = (Int(Mid(hora_fin, 1, 2)) * Hora + _
                       Int(Mid(hora_fin, 3, 2))) - _
                      (Int(Mid(hora_inicio, 1, 2)) * Hora + _
                       Int(Mid(hora_inicio, 3, 2)))
    End If
    cantdias = DateDiff("d", Fecha_Inicio, Fecha_Fin) * Dia
    
    total = cantdias + canthoras
    TotDias = Int(total / Dia)
    tothoras = Int((total Mod Dia) / Hora)
    TotMin = (total Mod Hora)
End Sub

Public Function ValidarRuta(ByVal modarchdefault As String, ByVal carpeta As String, ByVal msgError As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Valida que exista una carpeta, sino la crea y devuelve la ruta completa.
' Autor      : Gonzalez Nicol�s
' Fecha      : 09/01/2012
' Ultima Mod : FGZ - 26/09/2012 - le agregu� control por si la variable carpeta ya viene con barra "\"
' ---------------------------------------------------------------------------------------------
    'Activo el manejador de errores
    'On Error Resume Next
    Dim Ruta, carpetanueva
    Dim fs ' Licho - CAS-17249 - HEIDT- INSTALACION DEL SEC EN TEST80
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    'FGZ - 26/09/2012 --------------------------------
    If Left$(carpeta, 1) = "\" Or Left$(carpeta, 2) = "\\" Then
        Ruta = modarchdefault & carpeta
    Else
        Ruta = modarchdefault & "\" & carpeta
    End If
    'Ruta = modarchdefault & "\" & carpeta
    'FGZ - 26/09/2012 --------------------------------
    
    If Not fs.FolderExists(Ruta) Then
        Flog.writeline "La carpeta " & carpeta & " no existe. Se crear�."
        'CREO LA CARPETA
        Set carpetanueva = fs.CreateFolder(Ruta)
        'Flog.writeline "error: " & Err.Number
        If Err.Number <> 0 Then
            'IMPRIME MSG. DE ERROR EN CASO QUE NO SE PUEDA CREAR LA CARPETA.
            Select Case msgError
                Case 1:
                    Flog.writeline "Error al crear la carpeta " & carpeta & " - Consulta al Administrador del Sistema."
            End Select
        End If
        ValidarRuta = Ruta
    Else
        ValidarRuta = Ruta
    End If
End Function

Public Sub moverArchivos(ByVal directorioOrigen As String, ByVal directorioDestino As String, ByVal borrar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Mueve los archivos de un directorio a otro.
' Autor      : Deluchi Ezequiel.
' Fecha      : 24/07/2012
' Ultima Mod :
' ---------------------------------------------------------------------------------------------
' Parametros
'----------------------------------------------------------------------------------------------
' directorioOrigen: directorio origen del archivo a mover
' directorioDestino: directorio destino donde se van a mover los archivos
' borrar: valor booleano para saber si borra (mover) o no (copiar) el archivo origen

Dim Ruta
Dim directorio
Dim fichero
Dim file
Dim fso As Object


'chequeo que exista la ruta Destino
Ruta = ValidarRuta(directorioDestino, "", 0)

Set fso = CreateObject("Scripting.FileSystemObject")
Set file = CreateObject("Scripting.FileSystemObject")

'Agrego la barra al final por si no la tiene
If Right(Ruta, 1) <> "\" Then
        Ruta = Ruta & "\"
End If

If Right(directorioOrigen, 1) <> "\" Then
        directorioOrigen = directorioOrigen & "\"
End If

Set directorio = fso.GetFolder(directorioOrigen)


For Each fichero In directorio.Files
    Set file = fso.GetFile(fichero)
    'Activo el manejador de errores
    On Error Resume Next

    fso.copyFile directorioOrigen + file.Name, Ruta + file.Name
    
    If borrar Then
        Flog.writeline "Moviendo el archivo: " & fichero.Name & " a la carpeta: " & Ruta
        fso.DeleteFile directorioOrigen + file.Name, True
    Else
        Flog.writeline "Copiando el archivo: " & fichero.Name & " a la carpeta: " & Ruta
    End If
    'desactivo el manejador de errores
    On Error GoTo 0

Next
Flog.writeline "Se movieron los archivos correctamente."
End Sub


Public Function Armar_Fecha(ByVal Dia, ByVal Mes, ByVal Anio) As Date
' ---------------------------------------------------------------------------------------------
' Descripcion: Arma una fecha controlando que no sea biciesto
' Autor      : FGZ
' Fecha      : 06/08/2014
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Fecha

    'Si al alta fu� un 29/02 ==> lo llevo al 28/09
    If Dia = 29 And Mes = 2 And Not Biciesto(Anio) Then
        Fecha = CDate(Dia - 1 & "/" & Mes & "/" & Anio)
    Else
        Fecha = CDate(Dia & "/" & Mes & "/" & Anio)
    End If
    Armar_Fecha = Fecha
End Function



Public Sub CalcularDiasVac(ByVal TipoVac As Long, ByVal FechaInicial As Date, ByRef FechaFinal As Date, ByVal Ternro As Long, ByRef CHabiles As Integer, ByRef cNoHabiles As Integer, ByRef cFeriados As Integer)
'Calcula la cantidad de dias habiles, no habiles y feriados entre 2 fechas teniendo en cuenta el tipo de vacacion

Dim i As Integer
Dim j As Integer
Dim objFeriado As New Feriado
Dim DHabiles(1 To 7) As Boolean
Dim EsFeriado As Boolean
Dim objRs As New ADODB.Recordset
Dim ExcluyeFeriados As Boolean
Dim Fecha As Date


    StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & TipoVac
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        DHabiles(1) = objRs!tpvhabiles__1
        DHabiles(2) = objRs!tpvhabiles__2
        DHabiles(3) = objRs!tpvhabiles__3
        DHabiles(4) = objRs!tpvhabiles__4
        DHabiles(5) = objRs!tpvhabiles__5
        DHabiles(6) = objRs!tpvhabiles__6
        DHabiles(7) = objRs!tpvhabiles__7
    
        ExcluyeFeriados = CBool(objRs!tpvferiado)
    Else
        Flog.writeline "No se encontro el tipo de Vacacion " & TipoVac
        Exit Sub
    End If

    
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = objConn
    
    i = 0
    j = 0
    CHabiles = 0
    cNoHabiles = 0
    cFeriados = 0
    
    Fecha = FechaInicial
    
    Do While Fecha <= FechaFinal
    
        EsFeriado = objFeriado.Feriado(Fecha, Ternro, False)
        
        If (EsFeriado) And Not ExcluyeFeriados Then
            cFeriados = cFeriados + 1
        Else
            If DHabiles(Weekday(Fecha)) Or (EsFeriado And ExcluyeFeriados) Then
                i = i + 1
                If DHabiles(Weekday(Fecha)) Then
                    CHabiles = CHabiles + 1
                End If
            Else
                cNoHabiles = cNoHabiles + 1
            End If
        End If
        Fecha = DateAdd("d", 1, Fecha)
    Loop
    
    Set objFeriado = Nothing
End Sub

Public Function EstaFirmado(ByVal FirmaActiva As Boolean, ByVal cysfircodext As Long, ByVal tipoFirma As Byte)
 Dim Firmado As Boolean
 Dim rs_firmas As New ADODB.Recordset
 Dim USA_DEBUG As String
    If FirmaActiva Then
        '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
        StrSql = "SELECT cysfirfin from cysfirmas WHERE cysfiryaaut = -1 AND cysfirfin = -1 " & _
                " AND cysfircodext = '" & cysfircodext & "' AND cystipnro = " & tipoFirma
        OpenRecordset StrSql, rs_firmas
        If rs_firmas.EOF Then
            Firmado = False
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
            End If
        Else
            Firmado = True
        End If
    Else
        Firmado = True
    End If
    
    EstaFirmado = Firmado
End Function




