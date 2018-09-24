Attribute VB_Name = "MdlExportacion2"
Public Sub AlmCuentaCC(ByVal CuentaCompleta As String, ByVal Monto As Double, ByVal Linea As Integer, ByVal descrip As String, ByVal dh As Integer)

'--------------------------------------------------------------------------------
'  Descripci¢n: Almacena la Cuenta (Sin el Centro de Costo) y el Monto
'               para generar el total de dicha cuenta - CAS-14674
'  Autor: Zamarbide Juan
'  Fecha: 09/01/2012
'-------------------------------------------------------------------------------
 Dim Cuenta As String
 Dim num As Integer
 Dim cc As Boolean
  
 cc = False
 'Cuenta = Mid(CuentaCompleta, 11, Len(CuentaCompleta))
 'If Cuenta <> "" Then
    cc = True
 'End If
 Cuenta = ""
 If cc Then
    ReDim Preserve Acuentas(lastpos) As TR_Cuenta
    Cuenta = Mid(CuentaCompleta, 1, 10)
    
    Flog.writeline "monto - cuenta 2 " & Monto & "- " & Cuenta
    If InThere(Cuenta, Acuentas, num) Then
       If Not dh Then
           Acuentas(num).Monto = Acuentas(num).Monto - truncar(Monto, 2)
       Else
           Acuentas(num).Monto = Acuentas(num).Monto + truncar(Monto, 2)
       End If
    Else
       Acuentas(lastpos - 1).Cuenta = Cuenta
       Acuentas(lastpos - 1).Monto = truncar(Monto, 2)
       lastpos = lastpos + 1
    End If
 End If
End Sub
Public Sub AlmCuentaCCvariable(ByVal CuentaCompleta As String, ByVal Monto As Double, ByVal Linea As Integer, ByVal descrip As String, ByVal dh As Integer, ByVal Desde As Integer, ByVal Hasta As Integer)

'--------------------------------------------------------------------------------
'  Descripci¢n: Almacena la Cuenta (Sin el Centro de Costo) y el Monto
'               para generar el total de dicha cuenta - CAS-14674
'  Autor: Sebastian Stremel
'  Fecha: 29/01/2013
'-------------------------------------------------------------------------------
 Dim Cuenta As String
 Dim num As Integer
 Dim cc As Boolean
  
 cc = False
 'Cuenta = Mid(CuentaCompleta, 11, Len(CuentaCompleta))
 'If Cuenta <> "" Then
    cc = True
 'End If
 Cuenta = ""
 If cc Then
    ReDim Preserve Acuentas2(lastpos) As TR_Cuenta
    Cuenta = Mid(CuentaCompleta, Desde, Hasta)
    
    Flog.writeline "monto - cuenta 1 " & Monto & "- " & Cuenta
    If InThere(Cuenta, Acuentas2, num) Then
       'If Not dh Then
       '    Acuentas2(num).Monto = Acuentas2(num).Monto - Monto
       'Else
           Acuentas2(num).Monto = Acuentas2(num).Monto + truncar(Monto, 2)
       'End If
    Else
       Acuentas2(lastpos - 1).Cuenta = Cuenta
       'Acuentas2(lastpos - 1).Monto = CDbl(Replace(truncar(Monto, 2), ",", ""))
       Acuentas2(lastpos - 1).Monto = truncar(Monto, 2)
       lastpos = lastpos + 1
    End If
 End If
End Sub

Public Function InThere(ByVal cta As String, ByRef Acuentas() As TR_Cuenta, ByRef num As Integer) As Boolean
Dim j As Integer
InThere = False
For j = 0 To UBound(Acuentas)
    If Acuentas(j).Cuenta = cta Then
        InThere = True
        num = j
        Exit For
    End If
Next
End Function
Public Sub ImporteTotal(ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve el importe total, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Parte_Entera As Long
Dim Parte_Decimal As Integer
Dim Numero

    Numero = Split(CStr(totalImporte), ".")
    Parte_Entera = Fix(totalImporte)
    Parte_Decimal = IIf(Round((totalImporte - Parte_Entera) * 100, 0) > 0, Round((totalImporte - Parte_Entera) * 100, 0), 0)

    Numero(0) = Parte_Entera
    'cadena = Format(Parte_Entera, String(Longitud - 3, "0")) & SeparadorDecimales & Format(Parte_Decimal, "00")
    If Completar Then
        cadena = Format(Parte_Entera, String(longitud - 3, "0")) & SeparadorDecimales
    Else
        cadena = CStr(Parte_Entera) & SeparadorDecimales
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    
    cadena = Replace(cadena, ",", ".")
    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Public Sub ImporteTotalDH(ByVal Completar As Boolean, ByVal longitud As Integer, ByVal SumaDebe As Boolean, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve el importe total, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo,
'               dependiendo si el total es del DEBE o del HABER.
'  Autor: FAF
'  Fecha: 11/10/2007
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
    
    If SumaDebe Then
        Numero = Split(CStr(totalImporteD), ".")
        Parte_Entera = Fix(totalImporteD)
        Parte_Decimal = CStr(Format(IIf(Round((totalImporteD - Parte_Entera) * 100, 0) <> 0, Round(Abs(totalImporteD - Parte_Entera) * 100, 0), 0), "##"))
    Else
        Numero = Split(CStr(totalImporteH), ".")
        Parte_Entera = Fix(totalImporteH)
        Parte_Decimal = CStr(Format(IIf(Round((totalImporteH - Parte_Entera) * 100, 0) <> 0, Round(Abs(totalImporteH - Parte_Entera) * 100, 0), 0), "##"))
    End If

    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    cadena = Numero(0) & SeparadorDecimales
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    
    cadena = Replace(cadena, ",", ".")
    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Public Sub totalRegistros(ByVal total As Long, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(total, String(longitud, "0"))
    
    Str_Salida = cadena
End Sub

Public Sub totalRegistrosCompletar(ByVal total As Long, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: .
'  Autor: Carmen Quintero
'  Fecha: 12/08/2015
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

     cadena = total

     If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
     End If
     
     Str_Salida = cadena
End Sub

Public Sub Importe_Format(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String, ByVal Signo As String, ByVal separador As String)
'--------------------------------------------------------------------------------
'Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               Si va al debe es + y - sino, con dos decimales seguidos con el
'               Separador
'Autor: FGZ
'Fecha: 25/04/2005
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera
Dim Parte_Decimal
Dim Numero

    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
'    Parte_Decimal = IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    
    Numero(0) = Parte_Entera
    If Completar Then
        cadena = Format(Numero(0), String(longitud - 3, "0"))
    Else
        cadena = Numero(0)
    End If
    If debe Then
        '21/05/2014
        'cadena = " " & cadena
    Else
        cadena = "-" & cadena
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    
    If debe Then
        '21/05/2014
        'Aux_Cadena = " " & Format(Numero(0), String(Longitud - 3, "0")) & Separador
        Aux_Cadena = Format(Numero(0), String(longitud - 3, "0")) & separador
    Else
        Aux_Cadena = "-" & Format(Numero(0), String(longitud - 3, "0")) & separador
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = totalImporte + Abs(CSng(Aux_Cadena))
    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub
Public Sub Importe_Format_decimal(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String, ByVal Signo As String, ByVal separador As String)
'--------------------------------------------------------------------------------
'Descripci¢n:   devuelve el importe de la linea en el siguiente formato:
'               Si va al debe es + y - sino.
'               coloca los decimales con el separador configurado en el modelo 234
'Autor: Sebastian Stremel
'Fecha: 06/11/2013
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera
Dim Parte_Decimal
Dim Numero

    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
'    Parte_Decimal = IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    
    Numero(0) = Parte_Entera
    If Completar Then
        cadena = Format(Numero(0), String(longitud - 3, "0"))
    Else
        cadena = Numero(0)
    End If
    If debe Then
        cadena = " " & cadena
    Else
        cadena = "-" & cadena
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        'cadena = cadena & Left(Numero(1) & "00", 2) 'linea seba 05/11/2013
        cadena = cadena & separador & Left(Numero(1) & "00", 2) 'linea seba 05/11/2013
    Else
        'cadena = cadena & "00" 'linea seba 05/11/2013 separador
        cadena = cadena & separador & "00" 'linea seba 05/11/2013 separador
    End If
    
    If debe Then
        Aux_Cadena = " " & Format(Numero(0), String(longitud - 3, "0")) & separador
    Else
        Aux_Cadena = "-" & Format(Numero(0), String(longitud - 3, "0")) & separador
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = totalImporte + Abs(CSng(Aux_Cadena))
    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Public Function EsUltimoRegistroItem(ByRef Reg As ADODB.Recordset) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve TRUE si es el ultimo registro del recordset del tipo de item
' Autor      : FGZ
' Fecha      : 17/06/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Hay As Boolean
Dim Aux_Pos As Long

    Hay = False
    Aux_Pos = Reg.AbsolutePosition
    If Not Reg!itemicfijo Then
        Reg.MoveNext
        Do While Not Reg.EOF And Not Hay
            If UCase(Reg!itemicprog) = "IMPORTE" Then
                Hay = True
            End If
            Reg.MoveNext
        Loop
        'Reposiciono
        Reg.MoveFirst
        Do While Not Reg.AbsolutePosition = Aux_Pos
            Reg.MoveNext
        Loop
        If Not Hay Then
            EsUltimoRegistroItem = True
        Else
            EsUltimoRegistroItem = False
        End If
    Else
        EsUltimoRegistroItem = False
    End If
End Function

Public Function EsUltimoRegistroItemABS(ByRef Reg As ADODB.Recordset) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve TRUE si es el ultimo registro del recordset del tipo de item
' Autor      : FGZ
' Fecha      : 17/06/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Hay As Boolean
Dim Aux_Pos As Long

    Hay = False
    Aux_Pos = Reg.AbsolutePosition
    If Not Reg!itemicfijo Then
        Reg.MoveNext
        Do While Not Reg.EOF And Not Hay
            If UCase(Reg!itemicprog) = "IMPORTEABS" Then
                Hay = True
            End If
            Reg.MoveNext
        Loop
        'Reposiciono
        Reg.MoveFirst
        Do While Not Reg.AbsolutePosition = Aux_Pos
            Reg.MoveNext
        Loop
        If Not Hay Then
            EsUltimoRegistroItemABS = True
        Else
            EsUltimoRegistroItemABS = False
        End If
    Else
        EsUltimoRegistroItemABS = False
    End If
End Function

Public Sub BusCotiza(ByVal Fecha As Date, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve la cotizacion del dolar a la fecha Fecha, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo. Si no la encuentra
'               devuelve 1. Completa con ceros a izq si Completar = "S"
'  Autor: Martin Ferraro
'  Fecha: 24/10/2006
'  Modificado: Raul Chinestra - La Parte decimal del nro estaba saliendo mal, por ejemplo 3.08 salia 3.80.
'-------------------------------------------------------------------------------

Dim Valor As Double
Dim i As Integer
Dim Parte_Entera As Long
Dim Parte_Decimal As Integer
Dim Numero
Dim cadena As String
    
    Valor = CalculaCotizacion(Fecha, 2)

    Numero = Split(CStr(Valor), ".")
    'Parte_Entera = Fix(Valor)
    'Parte_Decimal = IIf(Round((Valor - Parte_Entera) * 100, 0) > 0, Round((Valor - Parte_Entera) * 100, 0), 0)
    'Numero(0) = Parte_Entera
    Numero(0) = Fix(Valor)
    
    'cadena = Parte_Entera & SeparadorDecimales
    cadena = Numero(0) & SeparadorDecimales
    If UBound(Numero) > 0 Then
        'Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
   
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        Else
            If Len(cadena) > longitud Then
                cadena = Right(cadena, longitud)
            End If
        End If
    End If
    Str_Salida = cadena
    
End Sub


Public Function CalculaCotizacion(ByVal Fecha As Date, ByVal MonDest As Integer) As Double

Dim rs_cotizacion As New ADODB.Recordset
Dim Aux As Double

    'Busco el valor del dolar en base
    StrSql = " SELECT * FROM pais"
    StrSql = StrSql & " INNER JOIN moneda ON moneda.paisnro = pais.paisnro AND monorigen = -1"
    StrSql = StrSql & " INNER JOIN cotizamon ON cotizamon.monnro = moneda.monnro"
    StrSql = StrSql & " AND cotizamon.cotfecha <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND cotizamon.mondestnro = " & MonDest
    StrSql = StrSql & " Where pais.paisdef = -1"
    StrSql = StrSql & " ORDER BY cotfecha DESC"
    OpenRecordset StrSql, rs_cotizacion
    
    If Not rs_cotizacion.EOF Then
        Aux = IIf(EsNulo(rs_cotizacion!cotvalororigen), 1, rs_cotizacion!cotvalororigen)
    Else
        Aux = 1
    End If
    rs_cotizacion.Close

    Set rs_cotizacion = Nothing
    
    'Retorno el resultado
    CalculaCotizacion = Aux
    
End Function

Public Sub ImporteUSD(ByVal Fecha As Date, ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato en USD:
'               Con dos decimales seguidos con el separador definido en el modelo
'  Autor: Martin Ferraro
'  Fecha: 25/10/2006
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single
Dim MontoUSD As Double

MontoUSD = Monto / CalculaCotizacion(Fecha, 2)

Numero = Split(CStr(MontoUSD), ".")
Parte_Entera = Fix(MontoUSD)
Parte_Decimal = CStr(Format(IIf(Round(Abs((MontoUSD - Parte_Entera)) * 100, 0) <> 0, Round(Abs(MontoUSD - Parte_Entera) * 100, 0), 0), "##"))
If Len(Parte_Decimal) < 2 Then
    Parte_Decimal = "0" & Parte_Decimal
Else
    Parte_Decimal = Left(Parte_Decimal, 2)
End If

Numero(0) = Parte_Entera
If debe Then
    If Completar Then
        cadena = Format(Numero(0), String(longitud - 3, "0"))
    Else
        cadena = Numero(0)
    End If
Else
    If Completar Then
        cadena = Format(Numero(0), String(longitud - 3, "0"))
    Else
        cadena = Numero(0)
    End If
End If
If UBound(Numero) > 0 Then
    Numero(1) = Parte_Decimal
    cadena = cadena & SeparadorDecimales & Left(Numero(1) & "00", 2)
Else
    cadena = cadena & SeparadorDecimales & "00"
End If

'Para calcular el total
If debe Then
    Aux_Cadena = Numero(0) & "."
Else
    Aux_Cadena = Numero(0) & "."
End If
If UBound(Numero) > 0 Then
    Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
Else
    Aux_Cadena = Aux_Cadena & "00"
End If
Aux_Cadena = Replace(Aux_Cadena, ",", ".")

    

'cadena = Aux_Cadena
If Completar Then
    If Len(cadena) < longitud Then
        cadena = String(longitud - Len(cadena), "0") & cadena
    Else
        If Len(cadena) > longitud Then
            cadena = Right(cadena, longitud)
        End If
    End If
End If
Str_Salida = cadena
 
End Sub


Public Sub InterfaceFecha(ByVal Fecha As Date, ByVal longitud As Long, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve el string INTERFACE concatenado a Fecha en formato AAMM
'               trunca o completa con espacios hasta logitud
'  Autor: Martin Ferraro
'  Fecha: 25/10/2006
'-------------------------------------------------------------------------------
Dim fechaAux As String
    
    'Armo la fecha
    fechaAux = ""
    Call Fecha_Estandar(Fecha, "YYMM", False, longitud, fechaAux)
    
    'Concateno Interface
    fechaAux = "INTERFACE" & fechaAux
    
    'Completo o trunco segun longitud
    If Len(fechaAux) < longitud Then
        fechaAux = String(longitud - Len(fechaAux), " ") & fechaAux
    Else
        If Len(fechaAux) > longitud Then
            fechaAux = Right(fechaAux, longitud)
        End If
    End If
    
    Str_Salida = fechaAux

End Sub

Public Sub SecuenciaFecha(ByVal Codigo As Long, ByVal Fecha As Date, ByVal longitud As Long, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve el valor del Codigo + "-" concatenado a Fecha en formato AAMMDD
'               trunca o completa con espacios hasta logitud
'  Autor: Martin Ferraro
'  Fecha: 25/10/2006
'-------------------------------------------------------------------------------
Dim fechaAux As String
    
    'Armo la fecha
    fechaAux = ""
    Call Fecha_Estandar(Fecha, "YYMMDD", False, longitud, fechaAux)
        
    'Concateno las partes de la salida
    fechaAux = Trim(Str(Codigo)) & "-" & fechaAux

    'Completo o trunco segun longitud
    If Len(fechaAux) < longitud Then
        fechaAux = String(longitud - Len(fechaAux), "0") & fechaAux
    Else
        If Len(fechaAux) > longitud Then
            fechaAux = Right(fechaAux, longitud)
        End If
    End If
    
    Str_Salida = fechaAux

End Sub

Public Sub CC_Profit(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal Cuenta As String, ByVal pos As Long, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByVal periodoMes As Integer, ByVal periodoAnio As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
' Descripción:
' La resolucion del nro de profit depende de algunas condiciones
' 1era) Segun el centro de costo del empleado:  Si comienza con 11* ==> Profit = "0000110000"
'                                               Si comienza con 45* ==> Profit = "0000453000"
'       OJO --- Terner en cuenta que algunas cuentas no estan abiertas por centros de costo ==> poner el profit que traen
'
' 2da)  a_ Toda cuenta que comience con 3* ==> Profit = "0000990000"
'       b_ Toda cuenta que comience con 4* ==> Profit = "          " - 10 espacios
'       c_ el resto llevan el profit segun la primera condicion
'
' Autor: FGZ
' Fecha: 06/02/2007
'-------------------------------------------------------------------------------
Dim Aux As String
Dim Aux_CC As String
Dim Aux_Profit As String
Dim cadena As String

    Aux = Mid(Trim(Cuenta), 1, 1)
    Select Case Aux
    Case "3":
        Aux_Profit = "0000990000"
        Aux_CC = "          "
    Case "4":
        Aux_Profit = "          "
        Aux_CC = Mid(Trim(Cuenta), 10, 10)
    Case Else
        'Si el CC comienza con "11" ==> cadena = "0000110000"
        'Si el CC comienza con "45" ==> cadena = "0000453000"
        'Toma el Profit cargado en E2 en la cuenta que esta en el mismo lugar donde estaria el CCosto(E1)
        Aux_CC = "          "
        Aux_Profit = Mid(Trim(Cuenta), 10, 10)
    End Select
    cadena = Aux_CC & Aux_Profit
                    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = cadena & String(longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena

'Cierro y libero
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Public Sub ComparaCTA(ByVal Cuenta As String, ByVal Desde As Long, ByVal cant As Long, ByVal CompararCon As String, ByVal Fecha As Date, ByVal longitud As Long, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve Fecha en formato MMAA si la porcion de cuenta desde cant es igual
'               Sino devuelve vacio o espacios completados a long
'  Autor: Martin Ferraro
'  Fecha: 22/02/2007
'-------------------------------------------------------------------------------
Dim cadenaAux As String
Dim Salida As String
    
    Salida = ""
    
    If cant = -1 Then
        'Hasta el Final
        cadenaAux = Mid(Cuenta, Desde, Len(Cuenta))
    Else
        cadenaAux = Mid(Cuenta, Desde, cant)
    End If
    
    If cadenaAux = CompararCon Then
        If Not EsNulo(Fecha) Then
            Salida = Format(Fecha, "MMYY")
        End If
    End If
    
    If Len(Salida) < longitud Then
        Salida = Salida & String(longitud - Len(Salida), " ")
    End If
    
        
End Sub

Public Sub ComparaCTAZ(ByVal Cuenta As String, ByVal Desde As Long, ByVal cant As Long, ByVal CompararCon As String, ByVal CambiarPor As String, ByVal LongCoincideDesde As Long, ByVal LongCoincideHasta As Long, ByVal Desde2 As Long, ByVal Cant2 As Long, ByVal CompletarCon As String, ByVal longitud As Long, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: Devuelde CambiarPor si la porcion de cuenta Desde Cant es igual a CompararCon
'               Sino verifica que la Cuenta tenga una longitud de LongCoincide, entonces devuelve
'                   la porcion de Cuenta a partir de Desde2 y de Cant2 caracteres
'  Autor: Fernando Favre
'  Fecha: 09/10/2007
'-------------------------------------------------------------------------------
Dim Salida As String
    
    Salida = ""
    If CompletarCon = "" Then
        CompletarCon = " "
    End If
    If CambiarPor = "" Then
        CambiarPor = " "
    End If
    
    If Desde > Len(Cuenta) Or Mid(Cuenta, Desde, cant) = CompararCon Then
        Salida = CambiarPor
    Else
        If Len(Cuenta) <= LongCoincideDesde And Len(Cuenta) >= LongCoincideHasta Then
            Salida = Mid(Cuenta, Desde2, Cant2)
        Else
            Salida = CompletarCon
        End If
    End If
    
    If Len(Salida) < longitud Then
        Salida = Salida & String(longitud - Len(Salida), CompletarCon)
    End If
    
    Str_Salida = Salida
End Sub

Public Sub ComparaCuentaF(ByVal Cuenta As String, ByVal Desde As Long, ByVal cant As Long, ByVal CompararCon As String, ByVal CambiarPor As String, ByVal LongCoincideDesde As Long, ByVal LongCoincideHasta As Long, ByVal Desde2 As Long, ByVal Cant2 As Long, ByVal CompletarCon As String, ByVal longitud As Long, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripcion: Realiza la misma funcionalidad de SICTA con la diferencia que no respeta la longitud del campo si el valor devuelto por el ítem es menor.
'  Autor: Carmen Quintero
'  Fecha: 01/10/2013
'-------------------------------------------------------------------------------
Dim Salida As String
    
    Salida = ""
    If CompletarCon = "" Then
        CompletarCon = " "
    End If
    If CambiarPor = "" Then
        CambiarPor = " "
    End If
    
    If Desde > Len(Cuenta) Or Mid(Cuenta, Desde, cant) = CompararCon Then
        Salida = CambiarPor
    Else
        If Len(Cuenta) <= LongCoincideDesde And Len(Cuenta) >= LongCoincideHasta Then
            Salida = Mid(Cuenta, Desde2, Cant2)
        Else
            Salida = CompletarCon
        End If
    End If
    
    'If Len(Salida) < Longitud Then
    '    Salida = Salida & String(Longitud - Len(Salida), CompletarCon)
    'End If
    
    Str_Salida = Salida
End Sub

Public Sub ComparaCuentaFF(ByVal Cuenta As String, ByVal Desde As Long, ByVal cant As Long, ByVal CompararCon As String, ByVal CambiarPor As String, ByVal LongCoincideDesde As Long, ByVal LongCoincideHasta As Long, ByVal Desde2 As Long, ByVal Cant2 As Long, ByVal CompletarCon As String, ByVal longitud As Long, ByVal caracter As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripcion: Realiza la misma funcionalidad de ComparaCuentaF con la diferencia que si se encuentra el parametro "carater" se corta el resultado hasta la aparicion de este
'  Autor: Mauricio Zwenger
'  Fecha: 10/07/2015
'-------------------------------------------------------------------------------
Dim Salida As String
    
    Salida = ""
    If CompletarCon = "" Then
        CompletarCon = " "
    End If
    If CambiarPor = "" Then
        CambiarPor = " "
    End If
    
    If Desde > Len(Cuenta) Or Mid(Cuenta, Desde, cant) = CompararCon Then
        Salida = CambiarPor
    Else
    
        If Len(Cuenta) <= LongCoincideDesde And Len(Cuenta) >= LongCoincideHasta Then
            
            If InStr(Desde, Cuenta, caracter) > 0 Then
                Cuenta = Left(Cuenta, InStr(Desde, Cuenta, caracter) - 1)
            End If
            
            Salida = Mid(Cuenta, Desde2, Cant2)
        Else
            Salida = CompletarCon
        End If
    End If
    
    Str_Salida = Salida
End Sub



Public Sub ComparaCADS(ByVal Cuenta As String, ByRef ArrStr() As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve una cadena vacía si una subcadena esta dentro de una cadena de la cuenta contable
'               si no, devuelve una subcadena de la cuenta contable. El arreglo que viene por parámetros tiene 5 posiciones:
'               X = inicio cadena 1, Y = tamaño cadena 1, C = cadena a comparar, Q = inicio cadena 2, W = tamaño cadena 2
'  Autor: Zamarbide Juan
'  Fecha: 04/08/2011
'-------------------------------------------------------------------------------
Dim start, length, Subcad As String
Dim SubArr() As String
Dim cont As Integer
Dim estac As Boolean
    estac = False
    start = ArrStr(0)
    length = ArrStr(1)
    Subcad = ArrStr(2)
    cont = -1
    SubArr = Split(Subcad, " ")
    Subcad = Trim(Mid(Cuenta, CLng(start), CLng(length)))
    
    Do While Not UBound(SubArr) = cont
        If Subcad = SubArr(cont + 1) Then
            estac = True
        End If
        cont = cont + 1
    Loop
    start = ArrStr(3)
    length = ArrStr(4)
    If estac Then
        Str_Salida = ""
    Else
        Subcad = Trim(Mid(Cuenta, CLng(start), CLng(length)))
        Str_Salida = Subcad
    End If
End Sub

Public Sub CuentaReemplaza(ByVal Cuenta As String, ByVal Str1 As String, ByVal Str2 As String, ByVal DesdeMostrar As Long, ByVal CantMostrar As Long, ByVal longitud As Long, ByRef Str_Salida As String)

'--------------------------------------------------------------------------------
'  Descripci¢n: Devuelve la cuenta con el str1 cambiado por str2, desde DesdeMostrar
'               la cantidad de caracteres CantMostrar
'  Autor: Martin Ferraro
'  Fecha: 06/06/2007
'-------------------------------------------------------------------------------
Dim cadenaAux As String
Dim Salida As String
    
    'Reemplazo
    Salida = Replace(Cuenta, Str1, Str2)
    
    'Trunco el string
    If DesdeMostrar <= 0 Then DesdeMostrar = 1
    If CantMostrar < 0 Then CantMostrar = 0
    Salida = Mid(Salida, DesdeMostrar, CantMostrar)
        
    'Completo con espacios
    If Len(Salida) < longitud Then
        Salida = Salida & String(longitud - Len(Salida), " ")
    End If
        
        
    Str_Salida = Salida
End Sub

Public Sub ImporteTotalDebeHaber(ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Long, ByRef Str_Salida As String, ByVal nroliq As Long, ByVal ProcVol As Long, ByVal Asinro As String)

    Dim cadena As String
    Dim rs_TotalImportaDH As New ADODB.Recordset
    Dim dh
    
    totalImporteDebe = 0
    totalImporteHaber = 0
    dh = 0
    If debe = True Then
        dh = -1
    End If
    '________________________________________________________________________________
    'Calcula el total para D y H
    '--------------------------------------------------------------------------------
    StrSql = "SELECT sum(monto) monto FROM  proc_vol "
    StrSql = StrSql & " INNER JOIN linea_asi ON proc_vol.vol_cod = linea_asi.vol_cod "
    StrSql = StrSql & " WHERE proc_vol.pliqnro =" & nroliq
    If ProcVol <> 0 Then 'Si no son todos
        StrSql = StrSql & " AND linea_asi.vol_cod IN (" & ProcVol & ")"
    End If
    StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
    StrSql = StrSql & " AND linea_asi.cuenta <> '999999.999'"
    StrSql = StrSql & " AND dh = " & dh
    OpenRecordset StrSql, rs_TotalImportaDH
    If Not rs_TotalImportaDH.EOF Then
        If debe = True Then
            totalImporteDebe = rs_TotalImportaDH!Monto
        Else
            totalImporteHaber = rs_TotalImportaDH!Monto
        End If
    End If
    '--------------------------------------------------------------------------------
    cadena = ""
    If debe Then
        If Completar Then
            If Len(totalImporteDebe) < longitud Then
                cadena = String(longitud - Len(CStr(totalImporteDebe)), "0") & Replace(totalImporteDebe, ".", SeparadorDecimales)
            Else
                cadena = Left(CStr(Replace(totalImporteDebe, ".", SeparadorDecimales)), longitud)
            End If
        Else
            cadena = CStr(Replace(totalImporteDebe, ".", SeparadorDecimales))
        End If
    Else
        If Completar Then
            If Len(totalImporteHaber) < longitud Then
                cadena = String(longitud - Len(CStr(totalImporteHaber)), "0") & Replace(totalImporteHaber, ".", SeparadorDecimales)
            Else
                cadena = Left(CStr(Replace(totalImporteHaber, ".", SeparadorDecimales)), longitud)
            End If
        Else
            cadena = CStr(Replace(totalImporteHaber, ".", SeparadorDecimales))
        End If
    End If
    Str_Salida = cadena
End Sub

Public Sub ImporteTotalDH2(ByVal Completar As Boolean, ByVal longitud As Long, ByRef Str_Salida As String, ByVal nroliq As Long, ByVal ProcVol As Long, ByVal Asinro As String, ByVal usaSeparador As Boolean)

    Dim cadena As String
    Dim rs_TotalImportaDH As New ADODB.Recordset
    Dim totalImporteDebeHaber As String
    
    '________________________________________________________________________________
    'Calcula la suma del total del D mas la suma total del H
    '--------------------------------------------------------------------------------
    
    StrSql = "SELECT sum(monto) monto FROM  proc_vol "
    StrSql = StrSql & " INNER JOIN linea_asi ON proc_vol.vol_cod = linea_asi.vol_cod "
    StrSql = StrSql & " WHERE proc_vol.pliqnro =" & nroliq
    If ProcVol <> 0 Then 'Si no son todos
        StrSql = StrSql & " AND linea_asi.vol_cod IN (" & ProcVol & ")"
    End If
    StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
    StrSql = StrSql & " AND linea_asi.cuenta <> '999999.999'"
    StrSql = StrSql & " AND dh in (-1,0) "
    OpenRecordset StrSql, rs_TotalImportaDH
    
    totalImporteDebeHaber = "0.00"
    If Not rs_TotalImportaDH.EOF Then
        If Not IsNull(rs_TotalImportaDH!Monto) Then
            totalImporteDebeHaber = Format(rs_TotalImportaDH!Monto, ".00")
        End If
    End If
    '--------------------------------------------------------------------------------
    cadena = ""
    If Not usaSeparador Then
        totalImporteDebeHaber = Replace(totalImporteDebeHaber, ".", "")
    End If

        If Completar Then
            If Len(totalImporteDebeHaber) < longitud Then
                cadena = String(longitud - Len(CStr(totalImporteDebeHaber)), "0") & Replace(totalImporteDebeHaber, ".", SeparadorDecimales)
            Else
                cadena = Left(CStr(Replace(totalImporteDebeHaber, ".", SeparadorDecimales)), longitud)
            End If
        Else
            cadena = CStr(Replace(totalImporteDebeHaber, ".", SeparadorDecimales))
        End If

    Str_Salida = cadena
End Sub

Public Sub ImporteTotalHaberNegativo(ByVal Completar As Boolean, ByVal longitud As Long, ByRef Str_Salida As String, ByVal nroliq As Long, ByVal ProcVol As Long, ByVal Asinro As String, ByVal usaSeparador As Boolean)

    Dim cadena As String
    Dim rs_TotalImportaDH As New ADODB.Recordset
    Dim totalImporteDebeHaber As String
    
    '________________________________________________________________________________
    'Calcula la suma del total del Haber
    '--------------------------------------------------------------------------------
    
    StrSql = "SELECT sum(monto) monto FROM  proc_vol "
    StrSql = StrSql & " INNER JOIN linea_asi ON proc_vol.vol_cod = linea_asi.vol_cod "
    StrSql = StrSql & " WHERE proc_vol.pliqnro =" & nroliq
    If ProcVol <> 0 Then 'Si no son todos
        StrSql = StrSql & " AND linea_asi.vol_cod IN (" & ProcVol & ")"
    End If
    StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
    StrSql = StrSql & " AND linea_asi.cuenta <> '999999.999'"
    StrSql = StrSql & " AND dh in (0) "
    OpenRecordset StrSql, rs_TotalImportaDH
    
    totalImporteDebeHaber = "0.00"
    If Not rs_TotalImportaDH.EOF Then
        If Not IsNull(rs_TotalImportaDH!Monto) Then
            totalImporteDebeHaber = Format((rs_TotalImportaDH!Monto * -1), ".00")
        End If
    End If
    '--------------------------------------------------------------------------------
    cadena = ""
    If Not usaSeparador Then
        totalImporteDebeHaber = Replace(totalImporteDebeHaber, ".", "")
    End If
        
        If Completar Then
            If Len(totalImporteDebeHaber) < longitud Then
                cadena = String(longitud - Len(CStr(totalImporteDebeHaber)), "0") & Replace(totalImporteDebeHaber, ".", SeparadorDecimales)
            Else
                cadena = Left(CStr(Replace(totalImporteDebeHaber, ".", SeparadorDecimales)), longitud)
            End If
        Else
            cadena = CStr(Replace(totalImporteDebeHaber, ".", SeparadorDecimales))
        End If

    Str_Salida = cadena
End Sub

Public Sub primerCampo(ByRef texto1, ByRef texto2, ByRef cadena)

cadena = ""

If EsPrimeraLineaCuenta Then
    
    cadena = texto1
    
Else
    cadena = texto2
    
End If

End Sub


Public Sub Archivo_ASTO_SAP(ByVal Dir As String, ByVal Fecha As Date)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: Genera el archivo ASTOmmaa.txt para el volcado SAP de Halliburton
'  Autor: Fapitalle N.
'  Fecha: 18/08/2005
'-------------------------------------------------------------------------------
Dim fAstoSAP
Dim fs
Dim Archivo
Dim carpeta
Dim cadena As String

Archivo = Dir & "\ASTO" & Format(Fecha, "MMYY") & ".txt"
Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fAstoSAP = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Set carpeta = fs.CreateFolder(Dir)
    Set fAstoSAP = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores
On Error GoTo 0

cadena = "Constante" + ";" + "Blancos" + ";" + "Cuenta" + ";" + "Blancos" + ";" + "Entidad" + ";" + _
         "Blancos" + ";" + "Constante" + ";" + "Blanco" + ";" + "Moneda" + ";" + "Blanco" + ";" + _
         "Constante" + ";" + "Descripcion" + ";" + "Blanco" + ";" + "Fecha" + ";" + "Blanco" + ";" + _
         "Constante" + ";" + "Blanco" + ";" + "Importe" + ";" + "Debe/Haber"

fAstoSAP.writeline cadena
fAstoSAP.Close
    
End Sub

Public Sub Cuenta(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta.p
'  Descripci¢n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               999999999999.999999.99999999
'               ej: si la cuenta es: 11000003.521110.01
'                   debera salir:000011000003.521110.00000001
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    cadena = cadena & IIf(Len(cadena) = 10, ".1000", "")
    Str_Salida = cadena

End Sub


Public Sub Cuenta_1(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta.p
'  Descripci¢n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               999999999999.999999.99999999
'               ej: si la cuenta es: 11000003.521110.01
'                   debera salir:000011000003.521110.00000001
'  OBS  : TRAER LA UNIDAD DE NEGOCIO AL FRENTE DE LA CUENTA
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Subcad As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    cadena = cadena & IIf(Len(cadena) = 10, ".1000", "")
    
    'PARA TRAER LA UNIDAD DE NEGOCIO AL FRENTE DE LA CUENTA
    Subcad = cadena
    cadena = ""
    cadena = Mid(Subcad, 12, 4) & "." & Mid(Subcad, 1, 10)
    
    Str_Salida = cadena
End Sub


Public Sub Cuenta_2(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta2.p
'  Descripci¢n: devuelve los primeros numeros de la cuenta, hasta el primer punto
'               en un formato de 12 digitos.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    
    cadena = IIf(Mid(cadena, 12, 4) = "", "1000", Mid(cadena, 12, 4))
    Str_Salida = cadena

End Sub

Public Sub Cuenta_3(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta2.p
'  Descripci¢n: devuelve los primeros numeros de la cuenta, hasta el primer punto
'               en un formato de 6 digitos.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    cadena = IIf(Mid(cadena, 1, 6) = "", "000000", Mid(cadena, 1, 6))
    Str_Salida = cadena

End Sub

Public Sub Cuenta_4(ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/cuenta.p
'  Descripci¢n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               999999999999.999999.99999999
'               ej: si la cuenta es: 11000003.521110.01
'                   debera salir:000011000003.521110.00000001
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = ""
    i = 1
    Do While i <= Len(Cuenta)
        cadena = cadena + Mid(Cuenta, i, 1)
        i = i + 1
    Loop
    cadena = IIf(Mid(cadena, 8, 3) = "", "000", Mid(cadena, 8, 3))
    Str_Salida = cadena

End Sub


Public Sub Importe(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               Con dos decimales seguidos con el separador definido en el modelo
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    'Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0), "##"))
    Parte_Decimal = CStr(Format(IIf(Round(Abs((Monto - Parte_Entera)) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    Else
        Parte_Decimal = Left(Parte_Decimal, 2)
    End If
    
    Numero(0) = Parte_Entera
    If debe Then
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    Else
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & SeparadorDecimales & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & SeparadorDecimales & "00"
    End If
    
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = CDbl(Round(CDbl(totalImporte) + Abs(CDbl(cadena)), 2))
    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        If debe Then 'agregado por DNN el 08/01/2009
            Diferencia = Round(total + Abs(CDbl(Aux_Cadena)), 2)
        Else 'agregado por DNN el 08/01/2009
            Diferencia = Round(total - Abs(CDbl(Aux_Cadena)), 2) 'agregado por DNN el 08/01/2009
        End If 'agregado por DNN el 08/01/2009
        If Diferencia <> 0 Then
                totalImporte = Round(totalImporte - Abs(CDbl(cadena)), 2)
            'Monto = CSng(Aux_Cadena) + Diferencia
            Monto = -1 * total
            'CQ 24/02/2015
            Balancea = True
        Else
'            Total = Total + CSng(cadena)
'            Balancea = True
            If debe Then
                total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
            Else
                total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
            End If
            Balancea = True
        End If
    Else
        If debe Then
            total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
        Else
            total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
        End If
        Balancea = True
'
'        Balancea = True
'        Total = Total + CSng(cadena)
    End If
Loop

'cadena = Aux_Cadena
If Completar Then
    If Len(cadena) < longitud Then
        cadena = String(longitud - Len(cadena), "0") & cadena
    Else
        If Len(cadena) > longitud Then
            cadena = Right(cadena, longitud)
        End If
    End If
End If
Str_Salida = cadena
 
End Sub

Public Sub ImporteN(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: funciona igual que el item Importe, lo unico que no toma en cuenta que el asiento este balanceado
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

'Balancea = False
'Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    'Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0), "##"))
    Parte_Decimal = CStr(Format(IIf(Round(Abs((Monto - Parte_Entera)) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    Else
        Parte_Decimal = Left(Parte_Decimal, 2)
    End If
    
    Numero(0) = Parte_Entera
    If debe Then
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    Else
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & SeparadorDecimales & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & SeparadorDecimales & "00"
    End If
    
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = CDbl(Round(CDbl(totalImporte) + Abs(CDbl(cadena)), 2))
    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        If debe Then 'agregado por DNN el 08/01/2009
            Diferencia = Round(total + Abs(CDbl(Aux_Cadena)), 2)
        Else 'agregado por DNN el 08/01/2009
            Diferencia = Round(total - Abs(CDbl(Aux_Cadena)), 2) 'agregado por DNN el 08/01/2009
        End If 'agregado por DNN el 08/01/2009
        If Diferencia <> 0 Then
                totalImporte = Round(totalImporte - Abs(CDbl(cadena)), 2)
            'Monto = CSng(Aux_Cadena) + Diferencia
            Monto = -1 * total
        Else
'            Total = Total + CSng(cadena)
'            Balancea = True
            If debe Then
                total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
            Else
                total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
            End If
            'Balancea = True
        End If
    Else
        If debe Then
            total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
        Else
            total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
        End If
        'Balancea = True
'
'        Balancea = True
'        Total = Total + CSng(cadena)
    End If
'Loop

'cadena = Aux_Cadena
If Completar Then
    If Len(cadena) < longitud Then
        cadena = String(longitud - Len(cadena), "0") & cadena
    Else
        If Len(cadena) > longitud Then
            cadena = Right(cadena, longitud)
        End If
    End If
End If
Str_Salida = cadena
 
End Sub

Public Sub ImporteEsp(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               Con dos decimales seguidos con el separador definido en el modelo
'  Autor: DNN
'  Fecha: 23/01/2009
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single
Dim ii As Integer

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    'Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0), "##"))
    Parte_Decimal = CStr(Format(IIf(Round(Abs((Monto - Parte_Entera)) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    Else
        Parte_Decimal = Left(Parte_Decimal, 2)
    End If
    
    Numero(0) = Parte_Entera
    If debe Then
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    Else
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & SeparadorDecimales & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & SeparadorDecimales & "00"
    End If
    
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = CDbl(Round(CDbl(totalImporte) + Abs(CDbl(cadena)), 2))
    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        If debe Then 'agregado por DNN el 08/01/2009
            Diferencia = Round(total + CDbl(Aux_Cadena), 2)
        Else 'agregado por DNN el 08/01/2009
            Diferencia = Round(total - CDbl(Aux_Cadena), 2) 'agregado por DNN el 08/01/2009
        End If 'agregado por DNN el 08/01/2009
        If Diferencia <> 0 Then
            '======================================================
            'sebastian stremel - como no balancea el asiento y es incorrecto lo corto a la fuerza
            Flog.writeline "EL ASIENTO NO BALANCEA"
            Balancea = True
            '======================================================
            totalImporte = Round(totalImporte - Abs(CDbl(cadena)), 2)
            'Monto = CSng(Aux_Cadena) + Diferencia
            Monto = -1 * total
        Else
'            Total = Total + CSng(cadena)
'            Balancea = True
            If debe Then
                total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
            Else
                total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
            End If
            Balancea = True
        End If
    Else
        If debe Then
            total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
        Else
            total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
        End If
        Balancea = True
'
'        Balancea = True
'        Total = Total + CSng(cadena)
    End If
Loop

'cadena = Aux_Cadena
If Completar Then
    If Len(cadena) < longitud Then
        cadena = String(longitud - Len(cadena), "0") & cadena
    Else
        If Len(cadena) > longitud Then
            cadena = Right(cadena, longitud)
        End If
    End If
End If
Aux_Cadena = ""
'Tomo a la cadena y reemplazo los ceros de la izquierda por espacios
If Mid(cadena, 2, 1) <> "." And Mid(cadena, 2, 1) <> "," Then
    ii = 1
    Do While Mid(cadena, ii, 1) = "0"
        'Aux_Cadena = Mid(cadena, 1, ii)
        Aux_Cadena = Aux_Cadena & " "
        ii = ii + 1
    Loop
    If Not debe Then 'agrego el (-)
        Aux_Cadena = Mid(Aux_Cadena, 1, Len(Aux_Cadena) - 1) & "-"
    End If
End If
cadena = Aux_Cadena & Mid(cadena, ii, Len(cadena))
Str_Salida = cadena
 
End Sub

Public Sub ImporteEspS(ByVal Monto As Double, ByVal Completar As Boolean, ByVal longitud As Integer, ByVal debe As Boolean, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               Ceros a la izquierda, sin signo, con dos decimales seguidos con el separador definido en el modelo
'  Autor: JAZ
'  Fecha: 06/01/2011
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single
Dim ii As Integer

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Abs(Monto))
    'Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0), "##"))
    Parte_Decimal = CStr(Format(IIf(Round(Abs((Monto - Parte_Entera)) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    Else
        Parte_Decimal = Left(Parte_Decimal, 2)
    End If
    
    Numero(0) = Parte_Entera
    If debe Then
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    Else
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & "." & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "." & "00"
    End If
    
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = CDbl(Round(CDbl(totalImporte) + Abs(CDbl(cadena)), 2))
    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        If debe Then 'agregado por DNN el 08/01/2009
            Diferencia = Round(total + CDbl(Aux_Cadena), 2)
        Else 'agregado por DNN el 08/01/2009
            Diferencia = Round(total - CDbl(Aux_Cadena), 2) 'agregado por DNN el 08/01/2009
        End If 'agregado por DNN el 08/01/2009
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(cadena)), 2)
            'Monto = CSng(Aux_Cadena) + Diferencia
            Monto = -1 * total
        Else
'            Total = Total + CSng(cadena)
'            Balancea = True
            If debe Then
                total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
            Else
                total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
            End If
            Balancea = True
        End If
    Else
        If debe Then
            total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
        Else
            total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
        End If
        Balancea = True
'
'        Balancea = True
'        Total = Total + CSng(cadena)
    End If
Loop

'cadena = Aux_Cadena
If Completar Then
    If Len(cadena) < longitud Then
        cadena = String(longitud - Len(cadena), "0") & cadena
    Else
        If Len(cadena) > longitud Then
            cadena = Right(cadena, longitud)
        End If
    End If
End If
Aux_Cadena = ""
'Tomo a la cadena y reemplazo los ceros de la izquierda por espacios
If Mid(cadena, 2, 1) <> "." And Mid(cadena, 2, 1) <> "," Then
    ii = 1
    Do While Mid(cadena, ii, 1) = "0"
        'Aux_Cadena = Mid(cadena, 1, ii)
        Aux_Cadena = Aux_Cadena & " "
        ii = ii + 1
    Loop
End If
cadena = Aux_Cadena & Mid(cadena, ii, Len(cadena))
Str_Salida = cadena
 
End Sub

Public Sub ImporteDH(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               Con dos decimales seguidos con el separador definido en el modelo
'  Autor:
'  Fecha:
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero

'Si monto es 0 es porque se pidio un item distinto al que trae el registro de la base,
'por ejemplo item IMPORTED (importe debe) y viene el un monto del haber, entonces llega un 0 en monto.
If Monto <> 0 Then
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round(Abs((Monto - Parte_Entera)) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    Else
        Parte_Decimal = Left(Parte_Decimal, 2)
    End If
    
    Numero(0) = Parte_Entera
    If debe Then
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    Else
        If Completar Then
            cadena = Format(Numero(0), String(longitud - 3, "0"))
        Else
            cadena = Numero(0)
        End If
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        If Len(Numero(1)) > 1 Then
            cadena = cadena & SeparadorDecimales & Left(Numero(1), 2)
        Else
            cadena = cadena & SeparadorDecimales & "0" & Numero(1)
        End If
    'Else
        'cadena = cadena & SeparadorDecimales & "00"
    End If
        
    'LED 21/08/2012
    If debe Then
        totalImporteDebe = CDbl(Round(CDbl(totalImporteDebe) + Abs(CDbl(Replace(cadena, ",", "."))), 2))
    Else
        totalImporteHaber = CDbl(Round(CDbl(totalImporteHaber) + Abs(CDbl(Replace(cadena, ",", "."))), 2))
    End If
    'LED 21/08/2012 - Fin
    
    'cadena = Aux_Cadena
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        Else
            If Len(cadena) > longitud Then
                cadena = Right(cadena, longitud)
            End If
        End If
    End If
Else
    cadena = ""
End If
Str_Salida = cadena
 
End Sub
Public Sub ImporteCH(ByVal Monto As Double, ByVal debe As Boolean, ByVal tipo As String, ByVal longitud As Integer, ByRef Str_Salida As String, ByVal completa As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               Sin decimales y con Ceros o no a la Izquierda dependiendo si es D u H
'  Autor:  Zamarbide Juan
'  Fecha:  31/05/2012
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Caso As String
Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single
Dim comp As String

comp = completa

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round(Abs((Monto - Parte_Entera)) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    Else
        Parte_Decimal = Left(Parte_Decimal, 2)
    End If
    
    Numero(0) = Parte_Entera
    'Me fijo el tipo de caso
   
    If debe Then
        Caso = "D"
    Else
        Caso = "H"
    End If
    
    'Comparo para ver si es el caso, completo con ceros, si no vacío
    If tipo <> "" Then

        If tipo = Caso Then
            'seba
            If completa = "C" Then
                cadena = Format(Numero(0), String(longitud, "0"))
            Else
                cadena = Numero(0)
            End If
            'hasta aca
            'cadena = Format(Numero(0), String(Longitud, "0"))
        Else
            cadena = ""
        End If

    Else
        'seba
        If completa = "C" Then
            cadena = Format(Numero(0), String(longitud, "0"))
        Else
            cadena = Numero(0)
        End If
        'hasta aca
    End If
    'If UBound(Numero) > 0 Then
    '    Numero(1) = Parte_Decimal
    '    cadena = cadena & SeparadorDecimales & Left(Numero(1) & "00", 2)
    'Else
    '    cadena = cadena & SeparadorDecimales & "00"
    'End If
    
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0)
    Else
        Aux_Cadena = Numero(0)
    End If
    'If UBound(Numero) > 0 Then
    '    Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    'Else
    '    Aux_Cadena = Aux_Cadena & "00"
    'End If
    'Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = CDbl(Round(CDbl(totalImporte) + Abs(CDbl(Aux_Cadena)), 2))
    If debe Then
        totalImporteD = CDbl(Round(CDbl(totalImporteD) + Abs(CDbl(Aux_Cadena)), 2))
    Else
        totalImporteH = CDbl(Round(CDbl(totalImporteH) + Abs(CDbl(Aux_Cadena)), 2))
    End If
    'cadena = Aux_Cadena
    If EsUltimoItem And EsUltimoProceso Then
        If debe Then 'agregado por DNN el 08/01/2009
            Diferencia = Round(total + CDbl(Aux_Cadena), 2)
        Else 'agregado por DNN el 08/01/2009
            Diferencia = Round(total - CDbl(Aux_Cadena), 2) 'agregado por DNN el 08/01/2009
        End If 'agregado por DNN el 08/01/2009
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
            'Monto = CSng(Aux_Cadena) + Diferencia
            Monto = -1 * total
        Else
'            Total = Total + CSng(cadena)
'            Balancea = True
            If debe Then
                total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
            Else
                total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
            End If
            Balancea = True
        End If
    Else
        If debe Then
            total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
        Else
            total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
        End If
        Balancea = True
'
'        Balancea = True
'        Total = Total + CSng(cadena)
    End If
Loop
    
If tipo <> "" Then
    If Len(cadena) < longitud Then
        'seba
        If completa = "C" Then
            cadena = String(longitud - Len(cadena), "0") & cadena
                   
        End If
    Else
        
        If Len(cadena) > longitud Then
            cadena = Right(cadena, longitud)
        End If
    End If
End If
Str_Salida = cadena
 
End Sub
'Private Sub ImporteCtroCostos(ByVal Monto As Double, ByVal tipo As String, ByVal Longitud As Integer, ByRef Str_Salida As String, ByVal completa As String)
'Private Sub ImporteCtroCostos(ByRef ProcVol, ByVal Cuenta As String, ByVal ccosto As String, ByVal mount As Double, ByRef cadena, ByVal longitud As Integer)
Public Sub ImporteCtroCostos(ByRef ProcVol, ByVal Cuenta As String, ByVal ccosto As String, ByVal mount As Double, ByRef cadena, ByVal longitud As Integer, ByVal debe)


Dim num As Integer
Dim Resultado As Double
Dim Diferencia As Double
Dim Aux_Cadena As String
Dim Balancea As Boolean

num = 0
Resultado = 0
cadena = ""
If InThere(Cuenta, Acuentas, num) Then
    'Resultado = (mount / Acuentas(num).Monto) * 100
    'cadena = CStr(FormatNumber(Round(Resultado, 2), 2))
    'Resultado = Round((mount / Acuentas(num).Monto) * 100, 2)
    
    'Comentado 06/01/2014
    Resultado = truncar(mount, 2)
    'Resultado = Acuentas(num).Monto
    cadena = CStr(Resultado)
    'cadena = CStr(FormatNumber(Resultado, 2))
End If

'validacion seba para ver si es del debe o del haber
If ((debe = 0) Or (Len(Trim(ccosto)) <= 0)) Then ' si no tiene apertura o es del haber muestro 15 espacios
    cadena = ""
End If


If EsUltimoLineaCuenta1 Then
    'If Resultado + ResultadoAcumuladoPorCC > 100 Then
    If Resultado + ResultadoAcumuladoPorCC1 > Acuentas(num).Monto Then
        'Resultado = (100 - ResultadoAcumuladoPorCC)
        Resultado = (Acuentas(num).Monto - ResultadoAcumuladoPorCC1)
    Else
        'If Resultado + ResultadoAcumuladoPorCC < 100 Then
        If Resultado + ResultadoAcumuladoPorCC1 < Acuentas(num).Monto Then
            'Resultado = (100 - ResultadoAcumuladoPorCC)
            Resultado = (Acuentas(num).Monto - ResultadoAcumuladoPorCC1)
        End If
    End If
    
    'Agregado el 06/01/2014
    'Resultado = Resultado
    'fin
    
    cadena = CStr(Resultado)
    'validacion seba para ver si es del debe o del haber
    If ((debe = 0) Or (Len(Trim(ccosto)) <= 0)) Then ' si no tiene apertura o es del haber muestro 15 espacios
        cadena = ""
    End If
    'If cadena <> "" Then
        'If Len(cadena) < 6 Then
        '    cadena = Right("   " & cadena, 6)
        'End If
    cadena = Replace(cadena, ",", "")
    If Len(cadena) < longitud Then
        'cadena = cadena & String(Longitud - Len(cadena), " ")
        cadena = String(longitud - Len(cadena), " ") & cadena
    End If
    'End If
    ResultadoAcumuladoPorCC1 = 0
    EsUltimoLineaCuenta1 = False
Else
    cadena = Replace(cadena, ",", "")
    If Len(cadena) < longitud Then
        'cadena = cadena & String(Longitud - Len(cadena), " ")
        'Resultado = Replace(cadena, ",", "")
        cadena = String(longitud - Len(cadena), " ") & cadena
        'Resultado = cadena
    End If
    ResultadoAcumuladoPorCC1 = ResultadoAcumuladoPorCC1 + Resultado
    
End If

End Sub


Public Sub Hacer_Header(ByVal dh As Boolean, ByVal Cuenta As String, ByVal Asi_Cod As String, ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: Devuelve el encabezado por asi_cod para la exportacion de shering
'  Autor: Fapitalle N.
'  Fecha: 12/08/2005
'-------------------------------------------------------------------------------
Dim cadena As String
Dim cuenta_contable As String
Dim codigo1 As String
Dim Texto As String

    cadena = "HEADR" + Format(Fecha, "DDMMYYYY") + "SA" + Format(Fecha, "MM")
    
    If dh Then
        codigo1 = "40"
    Else
        codigo1 = "50"
    End If
    
    Select Case Asi_Cod
        Case 1:
            Texto = "HABERES Y RETENCIONES    "
            If Len(Cuenta) = 19 Then
                codigo1 = "39"
            End If
        Case 2:
            Texto = "APORTES PATRONALES       "
        Case 3:
            Texto = "PREVISIONES              "
        Case 4:
            Texto = "INTERES S/PRESTAMO " + Format(Fecha, "MMYYYY")
            If Len(Cuenta) = 19 Then
                codigo1 = "29"
            End If
        Case Else:  'no deberia darse
            Texto = "<<<..ASICOD.WRONG.....>>>"
    End Select
    
    Select Case Len(Cuenta)
        Case 19:
            cuenta_contable = "00" + Mid(Cuenta, 11, 9)
        Case 10:
            cuenta_contable = Cuenta
        Case 14:
            cuenta_contable = Mid(Cuenta, 1, 10)
        Case Else:  'no deberia darse
            cuenta_contable = "<<LENWRG>>"
    End Select
    
    cadena = cadena + Texto + codigo1 + cuenta_contable
    
    Str_Salida = cadena
End Sub

Public Function Hacer_Pie(ByRef Reg As ADODB.Recordset)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: Devuelve verdadero si se necesita el pie por asi_cod para la exportacion de shering
'  Autor: Fapitalle N.
'  Fecha: 16/08/2005
'-------------------------------------------------------------------------------
Dim Asi_Cod_Actual
Dim hacer As Boolean

    Asi_Cod_Actual = Reg!masinro
    Reg.MoveNext
    If Reg.EOF Then
        hacer = True
    Else
        If Asi_Cod_Actual <> Reg!masinro Then
            hacer = True
        Else
            hacer = False
        End If
    End If
    Reg.MovePrevious
    Hacer_Pie = hacer
End Function


Public Function truncar(ByVal value As Double, ByVal escala As Integer) As Double

    Dim lngPotencia As Long
    Dim cadena, cadena1, cadena2 As String
    Dim ca As Integer


    lngPotencia = 10 ^ escala
    cadena = CStr(value * lngPotencia)
    
    ca = InStr(1, cadena, ".")
    If ca <> 0 Then
        cadena = Mid(cadena, 1, ca - 1)
    End If
    
    cadena1 = Mid(cadena, 1, Len(cadena) - 2)
    cadena2 = Mid(cadena, Len(cadena) - 1, Len(cadena))
    truncar = CDbl(cadena1 & "." & cadena2)
    
    'Flog.writeline "cadena1 " & cadena1 & "-  cadena2  - " & cadena2
    'truncar = value / lngPotencia
    'truncar = value
End Function


Public Sub ImporteGrupo(ByRef ProcVol, ByVal Cuenta As String, ByVal mount As Double, ByRef cadena, ByVal longitud As Integer, ByVal debe)
Dim num As Integer
Dim Resultado As Double
Dim Diferencia As Double
Dim Aux_Cadena As String
Dim Balancea As Boolean
'Dim suma As Double
num = 0
Resultado = 0

cadena = ""
If InThere(Cuenta, Acuentas2, num) Then
    'Comentado 06/01/2014
    'Resultado = Round(Acuentas2(num).Monto, 2)
    Resultado = Acuentas2(num).Monto
    cadena = CStr(Resultado)
End If


If EsUltimoLineaCuenta1 Then
    'Comentado 06/01/2014
    'Resultado = Acuentas2(num).Monto
    'Resultado = Round(Acuentas2(num).Monto, 2)
    Resultado = Acuentas2(num).Monto
    cadena = CStr(Resultado)
    ResultadoAcumuladoGrupo = 0
    EsUltimoLineaCuenta1 = False
    cadena = Replace(cadena, ",", "")
Else
    ResultadoAcumuladoGrupo = Replace(Resultado, ",", "")
    cadena = CStr(Resultado)
    cadena = Replace(cadena, ",", "")
End If


If Len(cadena) < longitud Then
    cadena = String(longitud - Len(cadena), " ") & cadena
End If



End Sub

Public Sub Fecha1(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a¤o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del año (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Year(CDate(Fecha)) Mod 1900, "000") & Format(CDate(Fecha) - CDate("01/01/" & Year(CDate(Fecha))), "000")
        
    Str_Salida = cadena

End Sub

Public Sub Fecha2(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a¤o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del año (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Fecha, "ddmmyy")
        
    Str_Salida = cadena

End Sub

Public Sub Fecha3(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a¤o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del año (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Fecha, "ddmm")
        
    Str_Salida = cadena

End Sub

Public Sub Fecha4(ByVal Fecha As Date, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision en el siguiente formato:
'               MYYYY donde los primeros 2 corresponden al mes de 1 a 12
'               y los otros 4 digitos son para el año
'  Autor: Fapitalle N.
'  Fecha: 09/08/2005
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Fecha, "MYYYY")
        
    Str_Salida = cadena

End Sub


Public Sub Fecha_Estandar(ByVal Fecha As Date, ByVal Formato As String, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve la fecha de emision  en el siguiente formato:
'               999999 donde los primeros tres corresponden al a¤o (099 para
'               1999 y 100 para 2000) y los otros tres digitos son para
'               los dias del año (del 001 al 365).
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Fecha, Formato)
    'Cadena = Format(Fecha, "ddmmyy")
        
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = cadena & String(longitud - Len(cadena), " ")
        End If
    End If
        
    Str_Salida = cadena

End Sub

Public Sub nroLinea(ByVal Linea As Long, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/nrolinea.p
'  Descripci¢n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Linea, String(longitud, "0"))
        
'    If Completar Then
'        If Len(Cadena) < Longitud Then
'            Cadena = String(Longitud - Len(Cadena), "0") & Cadena
'        End If
'    End If
    Str_Salida = cadena
End Sub

Public Sub NroAsiento(ByVal asiento As Long, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(asiento, String(longitud, "0"))
        
    Str_Salida = cadena
End Sub

Public Sub NroAsientoZ(ByVal asiento As Long, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = asiento
    
    If Len(cadena) < longitud Then
        cadena = String(longitud - Len(cadena), " ") & cadena
    End If
    
        
    Str_Salida = cadena
End Sub


Public Sub Leyenda(ByVal Descripcion As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/leyasiento.p
'  Descripci¢n: devuelve el la descripcion.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String
    If Len(Descripcion) < cant Then
        cadena = Mid(Descripcion, pos, Len(Descripcion))
    Else
        cadena = Mid(Descripcion, pos, cant)
    End If

    If Completar Then
        If Len(cadena) < longitud Then
            cadena = cadena & String(longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena

End Sub

Public Sub LeyendaAux(ByVal Descripcion As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByVal nroLinea As Integer, ByVal lista As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/leyasiento.p
'  Descripci¢n: devuelve el la descripcion en las lineas que corresponda.
'  Autor: Sebastian Stremel
'  Fecha: 04/02/2016
'-------------------------------------------------------------------------------
Dim cadena As String
lista = "," & lista & ","
    If InStr(lista, "," & nroLinea & ",") Then
        If Len(Descripcion) < cant Then
            cadena = Mid(Descripcion, pos, Len(Descripcion))
        Else
            cadena = Mid(Descripcion, pos, cant)
        End If
    
        If Completar Then
            If Len(cadena) < longitud Then
                cadena = cadena & String(longitud - Len(cadena), " ")
            End If
        End If
        Str_Salida = cadena
    Else
        Str_Salida = ""
    End If
End Sub

Public Sub Leyenda1(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal pos As Long, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/leyasiento.p
'  Descripci¢n: devuelve el la descripcion.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Linea As New ADODB.Recordset
Dim Encontro As Boolean

    If Len(Descripcion) < cant Then
        cadena = Mid(Descripcion, pos, Len(Descripcion))
    Else
        cadena = Mid(Descripcion, pos, cant)
    End If

Encontro = False
StrSql = "SELECT * FROM mod_linea "
StrSql = StrSql & " WHERE mod_linea.masinro =" & Asi_Cod
StrSql = StrSql & " AND mod_linea.linaorden =" & Linea
OpenRecordset StrSql, rs_Mod_Linea
If Not rs_Mod_Linea.EOF Then
    '1er nivel de estructura
    If Not EsNulo(rs_Mod_Linea!lineanivternro1) Then
        If rs_Mod_Linea!lineanivternro1 = 32 Then
            Encontro = True
        End If
    End If
    '2do nivel de estructura
    If Not EsNulo(rs_Mod_Linea!lineanivternro2) Then
        If rs_Mod_Linea!lineanivternro2 = 32 Then
            Encontro = True
        End If
    End If
    '3er nivel de estructura
    If Not EsNulo(rs_Mod_Linea!lineanivternro3) Then
        If rs_Mod_Linea!lineanivternro3 = 32 Then
            Encontro = True
        End If
    End If
    
    If Encontro Then
        cadena = "Jornales " & cadena
    Else
        cadena = "Sueldos " & cadena
    End If
End If
    
If Completar Then
    If Len(cadena) < longitud Then
        cadena = cadena & String(longitud - Len(cadena), " ")
    End If
End If
Str_Salida = cadena

End Sub

Public Sub Leyenda2(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal pos As Long, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: devuelve la descripcion del modelo.
'  Autor: DOS
'  Fecha: 18/05/2005
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Asiento As New ADODB.Recordset

    StrSql = "SELECT * FROM mod_asiento "
    StrSql = StrSql & " WHERE masinro =" & Asi_Cod
    
    OpenRecordset StrSql, rs_Mod_Asiento
    
    cadena = ""
    
    If Not rs_Mod_Asiento.EOF Then
       cadena = rs_Mod_Asiento!masidesc
    End If
    
    rs_Mod_Asiento.Close
        
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = cadena & String(longitud - Len(cadena), " ")
        End If
    End If
    
    Str_Salida = cadena

End Sub

Public Sub Modelo_Nro(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal pos As Long, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: devuelve la descripcion del modelo.
'  Autor: DOS
'  Fecha: 18/05/2005
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Asiento As New ADODB.Recordset
  
    cadena = ""
    cadena = Asi_Cod
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    
    Str_Salida = cadena

End Sub


Public Sub Leyenda3(ByVal Asi_Cod As Long, ByVal Linea As Long, ByVal Descripcion As String, ByVal pos As Long, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByVal periodoMes As Integer, ByVal periodoAnio As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo:
'  Descripci¢n: devuelve la descripcion del modelo y el periodo.
'  Autor: DOS
'  Fecha: 18/05/2005
'-------------------------------------------------------------------------------
Dim cadena As String
Dim rs_Mod_Asiento As New ADODB.Recordset

    StrSql = "SELECT * FROM mod_asiento "
    StrSql = StrSql & " WHERE masinro =" & Asi_Cod
    
    OpenRecordset StrSql, rs_Mod_Asiento
    
    cadena = ""
    
    If Not rs_Mod_Asiento.EOF Then
       cadena = Left(rs_Mod_Asiento!masidesc, 7)
    End If
    
    rs_Mod_Asiento.Close
    
    If periodoMes < 10 Then
       cadena = cadena & " 0" & periodoMes
    Else
       cadena = cadena & " " & periodoMes
    End If
    
    cadena = cadena & " " & periodoAnio
        
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = cadena & String(longitud - Len(cadena), " ")
        End If
    End If
    
    Str_Salida = cadena

End Sub


Public Sub NroPeriodo(ByVal Periodo As Long, ByVal Inicial As Long, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: .
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String

    cadena = Format(Periodo + Inicial, String(longitud, "0"))
        
    Str_Salida = cadena
End Sub


Public Sub ImporteABS(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
 '--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea

   ' Flog.writeline "Monto " & Monto
    
    Numero = Split(CStr(Monto), ".")
    
    'Monto = truncar(Monto, 2)
    
    Parte_Entera = Fix(Monto)
    
    'Flog.writeline "Parte Entera " & Parte_Entera
    
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    
    'Flog.writeline "Parte Decimal " & Parte_Decimal
    
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    If Completar Then
        cadena = Format(Numero(0), String(longitud - 3, "0")) & SeparadorDecimales
    Else
        cadena = Numero(0) & SeparadorDecimales
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
        
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Numero(1)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
    total = Round(total + CDbl(cadena), 2)
    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        
        'Flog.writeline "TotalABS " & TotalABS
        
        ' Agregado 27/04/2015
        cadena = CStr(Format(Abs(TotalABS), "#.00"))
        'fin
        
        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
        
        'Flog.writeline "Diferencia " & Diferencia
        
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
            
            'Flog.writeline "Total Importe " & totalImporte
            
            total = Round(total - CDbl(cadena), 2)
            
            'Flog.writeline "total " & total
            If Diferencia < 0 Then
                'Monto = CSng(Aux_Cadena) - Diferencia
                Monto = TotalABS * -1
            Else
                'Monto = CSng(Aux_Cadena) + Diferencia
                Monto = -1 * TotalABS
            End If
            Balancea = True
        Else
            If debe Then
                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
            Else
                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
            End If
            Balancea = True
        End If
    Else
        Balancea = True
        If debe Then
            TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
        Else
            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
        End If
    End If
Loop

    'Flog.writeline "Cadena " & cadena
    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub


Public Sub ImporteABSSD(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
 '--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               SIN separador de decimales.
'  Autor: LED
'  Fecha: 04/11/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    'Agregado 01/07/2014
    Monto = Round(Monto, 2)
    'fin
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    If Completar Then
        cadena = String(longitud - 2 - Len(Trim(Numero(0))), " ") & Trim(Numero(0))
    Else
        cadena = Numero(0)
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", "")
    cadena = Replace(cadena, "-", "")
        
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Numero(1)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
    total = Round(total + CDbl(cadena), 2)
    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
            total = Round(total - CDbl(cadena), 2)
            If Diferencia < 0 Then
                Monto = TotalABS * -1
            Else
                Monto = -1 * TotalABS
            End If
            Balancea = True
        Else
            If debe Then
                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
            Else
                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
            End If
            Balancea = True
        End If
    Else
        Balancea = True
        If debe Then
            TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
        Else
            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
        End If
    End If
Loop

    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub
Public Sub ImporteABS_2(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el ".", el relleno es con espacios al final
'               La funcion es una ligera modificacion de ImporteABS
'  Autor: Fapitalle N.
'  Fecha: 10/08/2005
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    cadena = Numero(0) & "."
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
    
    If Completar Then
        cadena = cadena & String(longitud - Len(cadena), " ")
    End If
    
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Numero(1)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
    total = Round(total + CDbl(cadena), 2)
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
            total = Round(total - CDbl(cadena), 2)
            If Diferencia < 0 Then
                'Monto = CSng(Aux_Cadena) - Diferencia
                Monto = TotalABS * -1
            Else
                'Monto = CSng(Aux_Cadena) + Diferencia
                Monto = -1 * TotalABS
            End If
            Balancea = True
        Else
            If debe Then
                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
            Else
                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
            End If
            Balancea = True
        End If
    Else
        Balancea = True
        If debe Then
            TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
        Else
            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
        End If
    End If
Loop

    Str_Salida = cadena

End Sub


Public Sub ImporteABS_3(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el ",", el relleno es con ceros adelante
'               La funcion es una ligera modificacion de ImporteABS
'  Autor: Fapitalle N.
'  Fecha: 10/08/2005
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
'Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    cadena = Numero(0) & ","
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, "-", "")
    
    If Completar Then
        cadena = String(longitud - Len(cadena), "0") + cadena
    End If
    
    'Para calcular el total
'    If Debe Then
'        Aux_Cadena = Numero(0) & "."
'    Else
'        Aux_Cadena = Numero(0) & "."
'    End If
'    If UBound(Numero) > 0 Then
'        Aux_Cadena = Aux_Cadena & Numero(1)
'    Else
'        Aux_Cadena = Aux_Cadena & "00"
'    End If
'    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
'    totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
'    Total = Round(Total + CDbl(cadena), 2)
    'FGZ - 17/06/2005
'    If EsUltimoItem And EsUltimoProceso Then
'        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
'        If Diferencia <> 0 Then
'            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
'            Total = Round(Total - CDbl(cadena), 2)
'            If Diferencia < 0 Then
'                'Monto = CSng(Aux_Cadena) - Diferencia
'                Monto = TotalABS * -1
'            Else
'                'Monto = CSng(Aux_Cadena) + Diferencia
'                Monto = -1 * TotalABS
'            End If
'        Else
'            If Debe Then
'                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
'            Else
'                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
'            End If
''            Balancea = True
''        End If
'    Else
'        Balancea = True
'        If Debe Then
'            TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
'        Else
'            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
'        End If
'    End If
'Loop

    Str_Salida = cadena

End Sub


Public Sub ImporteABS_4(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByVal Cuenta As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo. Solamente realiza
'               la suma en el total y el balanceo la primera vez de la cuenta. Se puede utilizar
'               varias veces en la misma linea
'  Autor: FAF
'  Fecha: 11/10/2007
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    If Len(Parte_Decimal) < 2 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    If Completar Then
        cadena = Format(Numero(0), String(longitud - 3, "0")) & SeparadorDecimales
    Else
        cadena = Numero(0) & SeparadorDecimales
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
        
    If Cuenta = cuenta_ant Then
        Balancea = True
    Else
        'Para calcular el total
        If debe Then
            Aux_Cadena = Numero(0) & "."
        Else
            Aux_Cadena = Numero(0) & "."
        End If
        If UBound(Numero) > 0 Then
            Aux_Cadena = Aux_Cadena & Numero(1)
        Else
            Aux_Cadena = Aux_Cadena & "00"
        End If
        Aux_Cadena = Replace(Aux_Cadena, ",", ".")
        totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
        total = Round(total + CDbl(cadena), 2)
        
        'Para calcular el total en el debe y en el haber
        If debe Then
            totalImporteD = CDbl(Round(CDbl(totalImporteD) + Abs(CDbl(Aux_Cadena)), 2))
        Else
            totalImporteH = CDbl(Round(CDbl(totalImporteH) + Abs(CDbl(Aux_Cadena)), 2))
        End If
        
        'FGZ - 17/06/2005
        If EsUltimoItem And EsUltimoProceso Then
            Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 2)
            If Diferencia <> 0 Then
                totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 2)
                total = Round(total - CDbl(cadena), 2)
                If Diferencia < 0 Then
                    'Monto = CSng(Aux_Cadena) - Diferencia
                    Monto = TotalABS * -1
                Else
                    'Monto = CSng(Aux_Cadena) + Diferencia
                    Monto = -1 * TotalABS
                End If
                Balancea = True
            Else
                If debe Then
                    TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
                Else
                    TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
                End If
                Balancea = True
            End If
        Else
            Balancea = True
            If debe Then
                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 2)
            Else
                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 2)
            End If
        End If
    End If
Loop

    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena
    
    cuenta_ant = Cuenta
    
End Sub


Public Sub ImporteABS_old(ByVal Monto As Single, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con dos decimales seguidos y
'               el separador de decimales es el definido en el modelo.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera
Dim Parte_Decimal
Dim Numero

    Numero = Split(CStr(Monto), ".")
    'Parte_Entera = Fix(Monto)
    'Parte_Decimal = IIf(Round((Monto - Parte_Entera) * 100, 0) > 0, Round((Monto - Parte_Entera) * 100, 0), 0)

'    If Debe Then
'       totalImporte = totalImporte + Monto
'    Else
'       totalImporte = totalImporte - Monto
'    End If
    
    If Completar Then
        cadena = Format(Numero(0), String(longitud - 3, "0")) & SeparadorDecimales
    Else
        cadena = Numero(0) & SeparadorDecimales
    End If
    If UBound(Numero) > 0 Then
        cadena = cadena & Left(Numero(1) & "00", 2)
    Else
        cadena = cadena & "00"
    End If
    cadena = Replace(cadena, ",", ".")
    cadena = Replace(cadena, "-", "")
        
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Left(Numero(1) & "00", 2)
    Else
        Aux_Cadena = Aux_Cadena & "00"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = totalImporte + Abs(CSng(Aux_Cadena))
    'totalImporte = totalImporte + Abs(CSng(cadena))
    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Public Sub debehaber(ByVal debe As Boolean, ByVal debeCod As String, ByVal haberCod As String, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve debeCod o haberCod dependiendo si es debe o haber.
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If debe Then
        cadena = debeCod
    Else
        cadena = haberCod
    End If
    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    Str_Salida = cadena

End Sub

Public Sub CodModAsiento(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve el código del modelo del asiento de la siguiente manera:
'               desde la posicion Pos por una cantidad Cant y completa con ceros hasta Longitud
'               ej: si el codigo es: 124
'                   pos = 1, Cant = 2, Completar = True , Longitud = 5
'                   debera salir: 00012
'  Autor: Fernando Favre
'  Fecha: 23/03/2006
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If

    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
            'cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena

End Sub

Public Sub DOCinCTA(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Tidnro As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve la cuenta contable con la porcion que va desde POS hasta POS + Cant reemplazada
'               por el documento informado en el Tipo de Documento
'               ej: si la cuenta es: 34621785432684322 y el doc = 0077800
'                   pos = 3, Cant = 4, Tidnro = 20
'                   debera salir: 34007780085432684322
'  Autor: Fernando Favre
'  Fecha: 15/08/2006
'-------------------------------------------------------------------------------
Dim cadena As String
Dim NroDoc As String
Dim StrSql As String
Dim rs_consult As New ADODB.Recordset

    StrSql = "SELECT nrodoc FROM ter_doc WHERE tidnro = " & Tidnro
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        NroDoc = Trim(rs_consult!NroDoc)
    End If
    rs_consult.Close
    
    cadena = Replace(Cuenta, Mid(Cuenta, pos, cant), NroDoc)
    
    Str_Salida = cadena

End Sub
Public Sub NroCuenta_2(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida2 As String)
'--------------------------------------------------------------------------------
'  Descripcion: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               desde la posicion Pos por una cantidad Cant y completa con espacios hasta Longitud
'               ej: si la cuenta es: 11000003.521110.01
'                   pos = 1, Cant = 8, Completar = True , Longitud = 12
'                   debera salir: 000011000003
'  Autor: Sebastian Stremel
'  Fecha: 13/02/2012
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If

    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), " ") & cadena
            'cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida2 = cadena
End Sub

Public Sub NroCuenta_1(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               desde la posicion Pos por una cantidad Cant y completa con ceros hasta Longitud
'               ej: si la cuenta es: 11000003.521110.01
'                   pos = 1, Cant = 8, Completar = True , Longitud = 12
'                   debera salir: 000011000003
'  Autor: Fapitalle N.
'  Fecha: 09/08/2005
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If

    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
            'cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena
End Sub

Public Sub NroCuenta(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve la cuenta de la linea del asiento de la siguiente manera:

'               desde la posicion Pos por una cantidad Cant
'               OBS: Completa con ESPACIOS al FINAL,
'               ej: si la cuenta es: 11000003.521110.01
'                   pos = 1, Cant = 8, Completar = True , Longitud = 12
'                   debera salir: '11000003    '
'
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If

    If Completar Then
        If Len(cadena) < longitud Then
            'cadena = String(Longitud - Len(cadena), "0") & cadena
            cadena = cadena & String(longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena
End Sub

Public Sub NroCuentaVariable(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve la cuenta de la linea del asiento: comienza desde la posicion que se pasa por parametro y la longitud depende si es debe o haber
'  Autor: EAM
'  Fecha: 22/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If
    Str_Salida = cadena
End Sub

Public Sub finCuentaZoetis(ByVal Cuenta As String, ByVal Completar As Boolean, ByVal longitud As Integer, ByVal Fecha As String, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve la cuenta de la linea del asiento: comienza desde la posicion que se pasa por parametro y la longitud depende si es debe o haber
'  Autor: EAM
'  Fecha: 22/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String
Dim longitudCadena As Double
Dim posicion As Double
Dim salir As Boolean
            
    longitudCadena = Len(Cuenta)
    salir = False
    posicion = 0
    Do While longitudCadena > 0 And Not salir
        
        If Mid(Cuenta, longitudCadena, 1) = "-" Then
            salir = True
        Else
            posicion = posicion + 1
            longitudCadena = longitudCadena - 1
        End If
        
    Loop
    
    If longitudCadena > 0 Then
        'busco que tengo que devolver
        Select Case UCase(Left(Right(Cuenta, posicion), 1))
            
            Case "Z" 'caso que la linea contiene -ZLLLL
                cadena = Right(Cuenta, posicion)
            Case "P" 'caso que la linea contiene -PNOMINA
                cadena = Right(Cuenta, posicion)
            Case "0" To "9" 'caso que la linea contiene -NNN donde n es NNN es un codigo de tipo de estructura
                cadena = obtenerEstructura(CLng(Left(Right(Cuenta, posicion), 3)), Right(Cuenta, 4), Fecha)
        End Select
    Else
        cadena = ""
    End If
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = cadena & String(longitud - Len(cadena), " ")
        Else
            cadena = Left(cadena, longitud)
        End If
    Else
        cadena = Left(cadena, longitud)
    End If
    Str_Salida = cadena
End Sub

Function obtenerEstructura(ByVal Tenro As Long, ByVal Legajo As String, ByVal Fecha As String)
Dim rs_consult As New ADODB.Recordset
Dim Ternro As String
Dim cadena As String

    'Verifico que el empleado exista
    StrSql = " SELECT ternro FROM empleado WHERE empleg = " & CLng(Legajo)

    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        Ternro = rs_consult!Ternro
    Else
        cadena = ""
    End If
    
    Select Case Tenro = 40
        Case 40
            StrSql = " SELECT ter_doc.nrodoc FROM empleado " & _
                     " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = " & Tenro & _
                     " INNER JOIN seguro ON seguro.estrnro = his_estructura.estrnro " & _
                     " LEFT JOIN ter_doc ON ter_doc.ternro = seguro.ternro AND ter_doc.tidnro = 6 " & _
                     " WHERE Empleado.ternro = " & Ternro & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha) & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR his_estructura.htethasta is null)) "
            
        Case Else
            'si el empleado existe voy a recuperar el documento de tipo 6 de la estructura activa para el empleado
        StrSql = " SELECT ter_doc.nrodoc FROM empleado " & _
                 " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = " & Tenro & _
                 " INNER JOIN replica_estr ON replica_estr.estrnro = his_estructura.estrnro " & _
                 " LEFT JOIN ter_doc ON ter_doc.ternro = replica_estr.origen AND ter_doc.tidnro = 6 " & _
                 " WHERE Empleado.ternro = " & Ternro & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha) & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR his_estructura.htethasta is null)) "
    End Select
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        If Not IsNull(rs_consult!NroDoc) Then
            cadena = rs_consult!NroDoc
        Else
            cadena = ""
        End If
    Else
        cadena = ""
    End If
        
    obtenerEstructura = cadena
    
    If rs_consult.State = adStateOpen Then rs_consult.Close
    Set rs_consult = Nothing
    
End Function


Public Sub nrovolcod(ByVal vol_cod As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: Devuelve el
'
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(vol_cod) < cant Then
        cadena = Mid(vol_cod, pos, Len(vol_cod))
    Else
        cadena = Mid(vol_cod, pos, cant)
    End If

    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
            'cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena
End Sub

Public Sub NroCuenta_n(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve la cuenta de la linea del asiento de la siguiente manera:
'               desde la posicion Pos por una cantidad Cant
'               OBS: Completa con ESPACIOS al FINAL,
'               ej: si la cuenta es: 11000003.521110.01
'                   pos = 1, Cant = 8, Completar = True , Longitud = 12
'                   debera salir: '11000003    '
'
'  Autor: FGZ
'  Fecha: 26/10/2004
'-------------------------------------------------------------------------------
Dim cadena As String

    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If
    If IsNumeric(cadena) Then
        cadena = Format(cadena, String(longitud, "#"))
    End If
    If Completar Then
        If Len(cadena) < longitud Then
            'cadena = String(Longitud - Len(cadena), "0") & cadena
            cadena = cadena & String(longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena
End Sub

'Private Sub Porcentaje_CC(ByRef ProcVol, ByVal debe, ByVal completa, ByVal Cuenta As String, ByVal ccosto As String, ByVal mount As Double, ByRef cadena)
Public Sub Porcentaje_CC(ByRef ProcVol, ByVal Cuenta As String, ByVal ccosto As String, ByVal mount As Double, ByRef cadena)
'--------------------------------------------------------------------------------
'  Descripci¢n: Calcula el % de un CC con respecto al total de la cuenta ligada a dicho CC
'  Autor: Zamarbide Juan
'  Fecha: 09/01/2012
' Ultima Modificacion: FGZ - 16/08/2012 - se le agregó control de ultima linea para ver como estaba el procentaje y redondear para llegar al 100%
'-------------------------------------------------------------------------------
Dim num As Integer
Dim Resultado As Double
Dim Diferencia As Double
Dim Aux_Cadena As String
Dim Balancea As Boolean

num = 0
Resultado = 0
cadena = ""
If InThere(Cuenta, Acuentas, num) Then
    'Resultado = (mount / Acuentas(num).Monto) * 100
    'cadena = CStr(FormatNumber(Round(Resultado, 2), 2))
    Resultado = Round((mount / Acuentas(num).Monto) * 100, 2)
    cadena = CStr(FormatNumber(Resultado, 2))
End If
If cadena <> "" Then
    If Len(cadena) < 6 Then
        cadena = Right("   " & cadena, 6)
    End If
End If


If EsUltimoLineaCuenta Then
    If Resultado + ResultadoAcumuladoPorCC > 100 Then
        Resultado = (100 - ResultadoAcumuladoPorCC)
    Else
        If Resultado + ResultadoAcumuladoPorCC < 100 Then
            Resultado = (100 - ResultadoAcumuladoPorCC)
        End If
    End If
    cadena = CStr(FormatNumber(Resultado, 2))
    If cadena <> "" Then
    If Len(cadena) < 6 Then
        cadena = Right("   " & cadena, 6)
    End If
End If
    ResultadoAcumuladoPorCC = 0
    EsUltimoLineaCuenta = False
Else
    ResultadoAcumuladoPorCC = ResultadoAcumuladoPorCC + Resultado
End If

End Sub

'FB se agrego , ByVal TipoArchivo As Long
Sub generarArchivo(ByRef rs_Periodo As Recordset, ByRef rs_Mod_Asiento As Recordset, ByRef rs_desc As Recordset, ByRef rs_Sistema As Recordset, ByVal TipoArchivo As Long, ByRef Asinro As String, ByVal nroliq As Long, ByVal Empresa As Long, ByVal separarArchivo As Integer, ByVal rs_Procesos As Recordset)

Dim rs_Archivo As New ADODB.Recordset

Select Case TipoArchivo
    Case 1
        Archivo = directorio & "\asi_" & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & ".csv"
        nombreArchivoExp = "\asi_" & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & ".csv"
    Case 2
        tmpStr = "int_cont_AR"
        tmpStr = tmpStr & "_" & Format(CStr(Year(Date)), "0000")
        tmpStr = tmpStr & "_" & Format(CStr(Month(Date)), "00")
        tmpStr = tmpStr & "_" & Format(CStr(Day(Date)), "00")
        tmpStr = tmpStr & "_" & Format(CStr(Hour(Now)), "00")
        tmpStr = tmpStr & "" & Format(CStr(Minute(Now)), "00")
        tmpStr = tmpStr & "" & Format(CStr(Second(Now)), "00")
        tmpStr = tmpStr & "_01.txt"
        Archivo = directorio & "\" & tmpStr
    Case 3
        Archivo = directorio & "\SAP" & Format(rs_Periodo!pliqhasta, "MMYY") & ".txt"
        nombreArchivoExp = "\SAP" & Format(rs_Periodo!pliqhasta, "MMYY") & ".txt"
    Case 4
        Archivo = directorio & "\" & Format(Right(CStr(Asinro), 4), "0000") & Format(rs_Periodo!pliqhasta, "MMYY") & ".txt"
        nombreArchivoExp = "\" & Format(Right(CStr(Asinro), 4), "0000") & Format(rs_Periodo!pliqhasta, "MMYY") & ".txt"
    Case 5
        Archivo = directorio & "\CBSIS006" & Format(CStr(Day(Date)), "00") & Format(CStr(Month(Date)), "00") & Format(CStr(Year(Date)), "0000") & Format(CStr(Hour(Now)), "00") & Format(CStr(Minute(Now)), "00") & Format(CStr(Second(Now)), "00") & ".txt"
        nombreArchivoExp = "\CBSIS006" & Format(CStr(Day(Date)), "00") & Format(CStr(Month(Date)), "00") & Format(CStr(Year(Date)), "0000") & Format(CStr(Hour(Now)), "00") & Format(CStr(Minute(Now)), "00") & Format(CStr(Second(Now)), "00") & ".txt"
    Case 6
        Archivo = ""
        cadena = Day(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = cadena
        cadena = Month(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = Archivo & cadena & Year(Date)
        Archivo = directorio & "\" & rs_Procesos!masinro & "_" & rs_Procesos!vol_cod & "_" & Archivo & ".txt"
        nombreArchivoExp = "\" & rs_Procesos!masinro & "_" & rs_Procesos!vol_cod & "_" & Archivo & ".txt"
    Case 7
        StrSql = "SELECT * FROM mod_asiento "
        StrSql = StrSql & " WHERE masinro =" & rs_Procesos!masinro
        OpenRecordset StrSql, rs_Mod_Asiento
        Archivo = ""
        cadena = Day(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = cadena
        cadena = Month(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = Archivo & cadena & Year(Date)
        Archivo = directorio & "\" & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & "_" & Archivo & ".txt"
        nombreArchivoExp = "\" & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & "_" & Archivo & ".txt"
        rs_Mod_Asiento.Close
        
    Case 8
        StrSql = "SELECT * FROM mod_asiento "
        StrSql = StrSql & " WHERE masinro =" & rs_Procesos!masinro
        OpenRecordset StrSql, rs_Mod_Asiento
        Archivo = ""
        cadena = Day(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = cadena
        cadena = Month(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = Archivo & cadena & Year(Date)
        'Archivo = directorio & "\" & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_desc & "_" & Archivo & ".txt"
        'nombreArchivoExp = rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_desc & "_" & Archivo & ".txt"
        
        'Agregado por Carmen Quintero 17/05/2013
        If ((InStr(rs_Procesos!vol_desc, "/") > 0) Or (InStr(rs_Procesos!vol_desc, "\") > 0)) Then
            Descripcion = ValidarDesc(rs_Procesos!vol_desc)
        Else
            Descripcion = rs_Procesos!vol_desc
        End If
        Archivo = directorio & "\" & rs_Mod_Asiento!masidesc & "_" & Descripcion & "_" & Archivo & ".txt"
        nombreArchivoExp = rs_Mod_Asiento!masidesc & "_" & Descripcion & "_" & Archivo & ".txt"
        'Fin
        
        rs_Mod_Asiento.Close
    Case 9
        StrSql = "SELECT * FROM mod_asiento "
        StrSql = StrSql & " WHERE masinro =" & rs_Procesos!masinro
        OpenRecordset StrSql, rs_Mod_Asiento
        Archivo = ""
        cadena = Day(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = cadena
        cadena = Month(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = Archivo & cadena & Year(Date)
        'Archivo = directorio & "\" & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & "_" & rs_Procesos!vol_desc & "_" & Archivo & ".txt"
        'nombreArchivoExp = rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & "_" & rs_Procesos!vol_desc & "_" & Archivo & ".txt"
        
        'Agregado por Carmen Quintero 17/05/2013
        If ((InStr(rs_Procesos!vol_desc, "/") > 0) Or (InStr(rs_Procesos!vol_desc, "\") > 0)) Then
            Descripcion = ValidarDesc(rs_Procesos!vol_desc)
        Else
            Descripcion = rs_Procesos!vol_desc
        End If
        Archivo = directorio & "\" & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & "_" & Descripcion & "_" & Archivo & ".txt"
        nombreArchivoExp = rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & "_" & Descripcion & "_" & Archivo & ".txt"
        'Fin
        
        rs_Mod_Asiento.Close
    Case 10
        StrSql = "SELECT * FROM mod_asiento "
        StrSql = StrSql & " WHERE masinro =" & rs_Procesos!masinro
        OpenRecordset StrSql, rs_Mod_Asiento
        Archivo = ""
        cadena = Day(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = cadena
        cadena = Month(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = Archivo & cadena & Year(Date)
        'Archivo = directorio & "\" & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_desc & "_" & rs_Procesos!vol_cod & "_" & Archivo & ".txt"
        'nombreArchivoExp = rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_desc & "_" & rs_Procesos!vol_cod & "_" & Archivo & ".txt"
        
        'Agregado por Carmen Quintero 17/05/2013
        If ((InStr(rs_Procesos!vol_desc, "/") > 0) Or (InStr(rs_Procesos!vol_desc, "\") > 0)) Then
            Descripcion = ValidarDesc(rs_Procesos!vol_desc)
        Else
            Descripcion = rs_Procesos!vol_desc
        End If
        Archivo = directorio & "\" & rs_Mod_Asiento!masidesc & "_" & Descripcion & "_" & rs_Procesos!vol_cod & "_" & Archivo & ".txt"
        nombreArchivoExp = rs_Mod_Asiento!masidesc & "_" & Descripcion & "_" & rs_Procesos!vol_cod & "_" & Archivo & ".txt"
        'Fin
        
        rs_Mod_Asiento.Close
    Case 11
        StrSql = "SELECT * FROM mod_asiento "
        StrSql = StrSql & " WHERE masinro =" & rs_Procesos!masinro
        OpenRecordset StrSql, rs_Mod_Asiento
        Archivo = ""
        cadena = Day(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = cadena
        cadena = Month(Date)
        If Len(cadena) < 2 Then cadena = "0" & cadena
        Archivo = Archivo & cadena & Year(Date)
        'Archivo = directorio & "\" & rs_Mod_Asiento!masinro & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & rs_Procesos!vol_desc & "_" & Archivo & ".txt"
        'nombreArchivoExp = rs_Mod_Asiento!masinro & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & rs_Procesos!vol_desc & "_" & Archivo & ".txt"
        
        'Agregado por Carmen Quintero 17/05/2013
        If ((InStr(rs_Procesos!vol_desc, "/") > 0) Or (InStr(rs_Procesos!vol_desc, "\") > 0)) Then
            Descripcion = ValidarDesc(rs_Procesos!vol_desc)
        Else
            Descripcion = rs_Procesos!vol_desc
        End If
        Archivo = directorio & "\" & rs_Mod_Asiento!masinro & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & Descripcion & "_" & Archivo & ".txt"
        nombreArchivoExp = rs_Mod_Asiento!masinro & rs_Mod_Asiento!masidesc & "_" & rs_Procesos!vol_cod & Descripcion & "_" & Archivo & ".txt"
        'Fin
        
        rs_Mod_Asiento.Close
    Case 12
        StrSql = "SELECT * FROM mod_asiento "
        StrSql = StrSql & " WHERE masinro =" & rs_Procesos!masinro
        OpenRecordset StrSql, rs_Mod_Asiento
        'Archivo = directorio & "\asi_" & rs_Procesos!vol_cod & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "00") & ".csv"
        Archivo = directorio & "\" & rs_Procesos!vol_cod & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "00") & ".txt"
        'nombreArchivoExp = "\asi_" & rs_Procesos!vol_cod & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "00") & ".csv"
        nombreArchivoExp = "\" & rs_Procesos!vol_cod & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "00") & ".txt"
        rs_Mod_Asiento.Close
    Case 13
        Archivo = directorio & "\" & Format(rs_Procesos!vol_cod, "0000") & Format(rs_Periodo!pliqhasta, "MMYY") & ".txt"
        nombreArchivoExp = Format(rs_Procesos!vol_cod, "0000") & Format(rs_Periodo!pliqhasta, "MMYY") & ".txt"
    '__________________________________
    'sebastian case 14 para cardif
    '291RHPR[Entidad][Tipo][Período][Version]
    '__________________________________
    Case 14
        '[entidad]
        Archivo = "291RHPR"
        Flog.writeline "Archivo " & Archivo
        If (Empresa = 7) Then
            Archivo = Archivo & "CSV"
        Else
            If (Empresa = 8) Then
                Archivo = Archivo & "CSE"
            End If
        End If
        
        Flog.writeline "Archivo segun empresa " & Empresa & "-" & Archivo
        
        '[tipo]
        asientoNro = Split(Asinro, ",")
        StrSql = "SELECT masidesc FROM mod_asiento WHERE masinro = " & asientoNro(1)
        OpenRecordset StrSql, rs_desc
        If Not rs_desc.EOF Then
            descMod = rs_desc!masidesc
            descMod = Mid(descMod, 4, 5)
            Archivo = Archivo & descMod
        End If
        rs_desc.Close
        
        Flog.writeline "Archivo modelo " & Archivo
        
        '[periodo]
        StrSql = "SELECT pliqhasta FROM periodo WHERE pliqnro=" & nroliq
        OpenRecordset StrSql, rs_desc
        If Not rs_desc.EOF Then
            Dim fechaAux
            'fechaAux = Format(rs_desc!pliqhasta, "YYYY/MM")
            fechaAux = Format(rs_desc!pliqhasta, "YY/MM")
            Archivo = Archivo & Replace(fechaAux, "/", "")
            Archivo = Archivo & Format(Now(), "DD")
        End If
        
        Flog.writeline "Archivo liq " & Archivo
        '[version]
        ' Comentado 07/01/2014
        'l_nro = 1
        'ArchivoAux = Archivo
        'Archivo = ArchivoAux & l_nro & ".txt"
        'fin
        
        'Carmen Quintero 07/01/2014
        ArchivoAux = Archivo
        
        Flog.writeline "Archivo aux " & ArchivoAux
        
        StrSql = " SELECT versiongen, archivogen FROM liq_archivo "
        StrSql = StrSql & "WHERE archivogen = '" & ArchivoAux & "'"
        StrSql = StrSql & " AND fechagen = " & ConvFecha(Date) & ""
        OpenRecordset StrSql, rs_Archivo
        If rs_Archivo.EOF Then
            l_nro = 1
            Archivo = ArchivoAux & l_nro & ".txt"
            StrSql = "INSERT INTO liq_archivo (archivogen,versiongen,fechagen)" & _
                 " VALUES ('" & ArchivoAux & _
                 "'," & l_nro & _
                 "," & ConvFecha(Date) & _
                 ")"
            Flog.writeline "Inserto en la tabla liq_archivo " & StrSql
        Else
            l_nro = CInt(rs_Archivo("versiongen")) + 1
            Archivo = ArchivoAux & l_nro & ".txt"
            StrSql = "UPDATE liq_archivo SET versiongen = " & l_nro & ", fechagen = " & ConvFecha(Date) & ""
            StrSql = StrSql & " WHERE archivogen = '" & rs_Archivo("archivogen") & "'"
            Flog.writeline "Actualizo en la tabla liq_archivo " & StrSql
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        'fin
        
       
       'sebastian stremel
        'Si existe una direccion en la tabla sistema campo sis_expseguridad copio los archivos en la direccion
        StrSql = " SELECT sis_expseguridad FROM sistema "
        OpenRecordset StrSql, rs_Sistema
        
        If Not EsNulo(rs_Sistema!sis_expseguridad) Then
            Flog.writeline "Me fijo si existe el archivo en la carpeta: " & rs_Sistema!sis_expseguridad
            'moverArchivos directorio, rs_Sistema!sis_expseguridad, True
            Dire = rs_Sistema!sis_expseguridad
        Else
            Dire = directorio
            Flog.writeline "No existe carpeta configurada en la tabla sistema, los archivos no se moveran. "
        End If
       'hasta aca
       
       
        ' Comentado 07/01/2014
        'Do While existe_archivo(Archivo, Dire)
         '   l_nro = l_nro + 1
         '   Archivo = ArchivoAux & l_nro & ".txt"
            'Archivo = Archivo & l_nro & ".txt"
        'Loop
        'fin
        
        Archivo = directorio & "\" & Archivo 'Aca encontró el nombre que no existe y lo arma.
        
    Case 15:
        Archivo = directorio & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & ".txt"
        nombreArchivoExp = Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & ".txt"
    
    Case 16:
        Archivo = directorio & "\CBSCH000" & Format(CStr(Day(Date)), "00") & Format(CStr(Month(Date)), "00") & Format(CStr(Year(Date)), "0000") & Format(CStr(Hour(Now)), "00") & Format(CStr(Minute(Now)), "00") & Format(CStr(Second(Now)), "00") & ".txt"
        nombreArchivoExp = "\CBSCH000" & Format(CStr(Day(Date)), "00") & Format(CStr(Month(Date)), "00") & Format(CStr(Year(Date)), "0000") & Format(CStr(Hour(Now)), "00") & Format(CStr(Minute(Now)), "00") & Format(CStr(Second(Now)), "00") & ".txt"
        
    Case 17:
        Archivo = directorio & "\CCLGL.PRARGN.CCLJEES" & Format(Now, "YYYYMMDD-HHmmss") & ".txt"
        nombreArchivoExp = "\CCLGL.PRARGN.CCLJEES" & Format(Now, "YYYYMMDD-HHmmss") & ".txt"
           
    'hasta aca
    Case Else
        Archivo = directorio & "\asi_" & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & ".csv"
        nombreArchivoExp = "\asi_" & Format(CStr(rs_Periodo!pliqmes), "00") & Format(CStr(rs_Periodo!pliqanio), "0000") & ".csv"
        
End Select

Set fs = CreateObject("Scripting.FileSystemObject")

'Activo el manejador de errores

On Error Resume Next
' Set fExport = fs.CreateTextFile(Archivo, True)

 If TipoArchivo = 17 Then
        Set fExport = fs.CreateTextFile(Archivo, True)
        fExport.Charset = "UTF-8"
        If separarArchivo = -1 Then
    
            'mid(instr(nombreArchivoExp,".")
            If CLng(TipoArchivo) = CLng(15) Then
                Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & "\HEAD_" & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
                Set fAuxiliarDetalle = fs.CreateTextFile(directorio & "\ITEM_" & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
                Set fAuxiliarPie = fs.CreateTextFile(directorio & "\" & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)

            Else
                Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & "_cab" & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
                Set fAuxiliarDetalle = fs.CreateTextFile(directorio & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & "_det" & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
                Set fAuxiliarPie = fs.CreateTextFile(directorio & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & "_pie" & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
            End If

        Else
            Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & "\fencab.tmp", True)
            Set fAuxiliarDetalle = fs.CreateTextFile(directorio & "\fdet.tmp", True)
            Set fAuxiliarPie = fs.CreateTextFile(directorio & "\fpie.tmp", True)
        End If
        fAuxiliarEncabezado.Charset = "UTF-8"
        fAuxiliarDetalle.Charset = "UTF-8"
        fAuxiliarPie.Charset = "UTF-8"
Else
   Set fExport = fs.CreateTextFile(Archivo, True)
   If separarArchivo = -1 Then
    
        'mid(instr(nombreArchivoExp,".")
        If CLng(TipoArchivo) = CLng(15) Then
            Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & "\HEAD_" & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
            Set fAuxiliarDetalle = fs.CreateTextFile(directorio & "\ITEM_" & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
            Set fAuxiliarPie = fs.CreateTextFile(directorio & "\" & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
        Else
            Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & "_cab" & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
            Set fAuxiliarDetalle = fs.CreateTextFile(directorio & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & "_det" & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
            Set fAuxiliarPie = fs.CreateTextFile(directorio & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & "_pie" & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
        End If
    Else
        Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & "\fencab.tmp", True)
        Set fAuxiliarDetalle = fs.CreateTextFile(directorio & "\fdet.tmp", True)
        Set fAuxiliarPie = fs.CreateTextFile(directorio & "\fpie.tmp", True)
     End If
End If


'If separarArchivo = -1 Then

''mid(instr(nombreArchivoExp,".")
  '  If CLng(TipoArchivo) = CLng(15) Then
  '      Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & "\HEAD_" & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
  '      Set fAuxiliarDetalle = fs.CreateTextFile(directorio & "\ITEM_" & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
  '      Set fAuxiliarPie = fs.CreateTextFile(directorio & "\" & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
  '  Else
  '      Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & "_cab" & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
  '      Set fAuxiliarDetalle = fs.CreateTextFile(directorio & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & "_det" & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
  '      Set fAuxiliarPie = fs.CreateTextFile(directorio & Mid(nombreArchivoExp, 1, InStr(nombreArchivoExp, ".") - 1) & "_pie" & Mid(nombreArchivoExp, InStr(nombreArchivoExp, "."), Len(nombreArchivoExp)), True)
  '  End If
'Else
 '   Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & "\fencab.tmp", True)
  '  Set fAuxiliarDetalle = fs.CreateTextFile(directorio & "\fdet.tmp", True)
  '  Set fAuxiliarPie = fs.CreateTextFile(directorio & "\fpie.tmp", True)
'End If

If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(directorio)
    If TipoArchivo = 17 Then
        Set fExport = fs.CreateTextFile(Archivo, True)
        Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & "\fencab.tmp", True)
        Set fAuxiliarDetalle = fs.CreateTextFile(directorio & "\fdet.tmp", True)
        Set fAuxiliarPie = fs.CreateTextFile(directorio & "\fpie.tmp", True)
        fExport.Charset = "UTF-8"
        fAuxiliarEncabezado.Charset = "UTF-8"
        fAuxiliarDetalle.Charset = "UTF-8"
        fAuxiliarPie.Charset = "UTF-8"
    Else
        Set fExport = fs.CreateTextFile(Archivo, True)
        Set fAuxiliarEncabezado = fs.CreateTextFile(directorio & "\fencab.tmp", True)
        Set fAuxiliarDetalle = fs.CreateTextFile(directorio & "\fdet.tmp", True)
        Set fAuxiliarPie = fs.CreateTextFile(directorio & "\fpie.tmp", True)
    End If
End If

End Sub

Sub lineas(ByVal programa As String, ByVal Linea As Integer, ByVal Cantidad As Integer, ByRef Str_Salida As String)
Dim cantCarac As Integer
Dim Parametro As String
cantCarac = Len(programa)
If cantCarac = 8 Then
    'obtengo el ultimo caracter
    Parametro = Right(programa, 1)
    If UCase(Parametro) = "S" Then
        'completo con ceros
        Str_Salida = String(Cantidad - Len(Trim(Linea)), "0") & Linea
        
    Else
    'no completo
    Str_Salida = Linea
    End If
Else
    'no completo
    Str_Salida = Linea
End If

    'Str_Salida = linea
End Sub

Sub IMPORTECTA(ByVal Monto As Double, ByVal debe As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef cadena As String)
'--------------------------------------------------------------------------------
'  Archivo    : IMPORTECTA
'  Descripcion: devuelve el importe de la linea con 2 decimales y completa con ceros o no dependiendo
'               del Parametro
'  Autor      : Sebastian Stremel
'  Fecha      : 20/08/2014
'-------------------------------------------------------------------------------
Dim i As Integer

Dim Aux_Cadena As String
Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero



Numero = Split(CStr(Monto), ".")
Parte_Entera = Fix(Monto)
Parte_Decimal = CStr(Format(IIf(Round(Abs((Monto - Parte_Entera)) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))

If Len(Parte_Decimal) < 2 Then
    Parte_Decimal = "0" & Parte_Decimal
Else
    Parte_Decimal = Left(Parte_Decimal, 2)
End If

Numero(0) = Parte_Entera
cadena = Numero(0)
If UBound(Numero) > 0 Then
    Numero(1) = Parte_Decimal
    cadena = cadena & SeparadorDecimales & Left(Numero(1) & "00", 2)
Else
    cadena = cadena & SeparadorDecimales & "00"
End If

'Agregado 27/04/2015
Dim Diferencia As Single

Aux_Cadena = Numero(0) & "."

If UBound(Numero) > 0 Then
    Aux_Cadena = Aux_Cadena & Numero(1)
Else
    Aux_Cadena = Aux_Cadena & "00"
End If
Aux_Cadena = Replace(Aux_Cadena, ",", ".")
totalImporte = Round(totalImporte + Abs(CDbl(Aux_Cadena)), 2)
total = Round(total + CDbl(cadena), 2)

    If EsUltimoItem And EsUltimoProceso Then
        If debe Then 'agregado por DNN el 08/01/2009
            Diferencia = Round(total + Abs(CDbl(Aux_Cadena)), 2)
        Else 'agregado por DNN el 08/01/2009
            Diferencia = Round(total - Abs(CDbl(Aux_Cadena)), 2) 'agregado por DNN el 08/01/2009
        End If 'agregado por DNN el 08/01/2009
        If Diferencia <> 0 Then
                totalImporte = Round(totalImporte - Abs(CDbl(cadena)), 2)
            'Monto = CSng(Aux_Cadena) + Diferencia
            Monto = -1 * total
        Else
'            Total = Total + CSng(cadena)
'            Balancea = True
            If debe Then
                total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
            Else
                total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
            End If
            'Balancea = True
        End If
    Else
        If debe Then
            total = Round(total + CDbl(Abs(Aux_Cadena)), 2)
        Else
            total = Round(total - CDbl(Abs(Aux_Cadena)), 2)
        End If
        'Balancea = True
'
'        Balancea = True
'        Total = Total + CSng(cadena)
    End If 'fin

If Completar Then
    If Not debe Then
    'si es del haber lleva signo -
        If Len(cadena) < longitud - 1 Then
            cadena = "-" & String((longitud - 1) - Len(cadena), "0") & cadena
        Else
            If Len(cadena) > longitud Then
                cadena = Right(cadena, longitud)
            End If
        End If
    Else
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        Else
            If Len(cadena) > longitud Then
                cadena = Right(cadena, longitud)
            End If
        End If
    End If
Else
    If Not debe Then
        cadena = "-" & cadena
    Else
        cadena = cadena
    End If
End If


End Sub

Public Sub ImporteABSSR(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
 '--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con 4 decimales seguidos y SIN separador de decimales.
'  Autor: EAM
'  Fecha: 28/08/2014
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
    'Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    Parte_Decimal = CStr(IIf(Replace(Left(FormatNumber((Monto - Parte_Entera), 4), 6), "0.", "") <> 0, Replace(Left(FormatNumber((Monto - Parte_Entera), 4), 6), "0.", ""), 0))
    If Len(Parte_Decimal) < 4 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    If Completar Then
        cadena = String(longitud - 4 - Len(Trim(Numero(0))), " ") & Trim(Numero(0))
    Else
        cadena = Numero(0)
    End If
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & Left(Numero(1) & "00", 4)
    Else
        cadena = cadena & "0000"
    End If
    cadena = Replace(cadena, ",", "")
    cadena = Replace(cadena, "-", "")
        
    'Para calcular el total
    If debe Then
        Aux_Cadena = Numero(0) & "."
    Else
        Aux_Cadena = Numero(0) & "."
    End If
    If UBound(Numero) > 0 Then
        Aux_Cadena = Aux_Cadena & Numero(1)
    Else
        Aux_Cadena = Aux_Cadena & "0000"
    End If
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    totalImporte = totalImporte + Abs(CDbl(Aux_Cadena))
    total = total + CDbl(cadena)
    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        Diferencia = Round(TotalABS + CDbl(Aux_Cadena), 4)
        If Diferencia <> 0 Then
            totalImporte = Round(totalImporte - Abs(CDbl(Aux_Cadena)), 4)
            total = Round(total - CDbl(cadena), 4)
            If Diferencia < 0 Then
                Monto = TotalABS * -1
            Else
                Monto = -1 * TotalABS
            End If
            Balancea = True
        Else
            If debe Then
                TotalABS = Round(TotalABS + CDbl(Aux_Cadena), 4)
            Else
                TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 4)
            End If
            Balancea = True
        End If
    Else
        Balancea = True
        If debe Then
            TotalABS = TotalABS + CDbl(Aux_Cadena)
        Else
            TotalABS = Round(TotalABS - CDbl(Aux_Cadena), 4)
        End If
    End If
Loop
        
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    
    Str_Salida = cadena
    

End Sub
Public Sub ImporteABSCR3(ByVal Monto As Double, ByVal debe As Boolean, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String, ByRef dh As String)
 '--------------------------------------------------------------------------------
'  Archivo: conta/importe.p
'  Descripci¢n: devuelve el importe de la linea en el siguiente formato:
'               el monto esta expresado en valor absoluto, con 3 decimales seguidos y SIN separador de decimales.
'  Autor: EAM
'  Fecha: 28/08/2014
'-------------------------------------------------------------------------------
Dim i As Integer
Dim cadena As String
Dim Aux_Cadena As String

Dim Parte_Entera As String
Dim Parte_Decimal As String
Dim Numero
Dim Balancea As Boolean
Dim Diferencia As Single

Balancea = False
Do While Not Balancea
    
    Monto = Round(Monto, 4)
    
    Numero = Split(CStr(Monto), ".")
    Parte_Entera = Fix(Monto)
     
    'Parte_Decimal = CStr(Format(IIf(Round((Monto - Parte_Entera) * 100, 0) <> 0, Round(Abs(Monto - Parte_Entera) * 100, 0), 0), "##"))
    Parte_Decimal = CStr(IIf(Replace(Left(FormatNumber((Monto - Parte_Entera), 4), 6), "0.", "") <> 0, Replace(Left(FormatNumber((Monto - Parte_Entera), 4), 6), "0.", ""), 0))
    
    
    If Len(Parte_Decimal) < 4 Then
        Parte_Decimal = "0" & Parte_Decimal
    End If
    Numero(0) = Parte_Entera
    
    If Completar Then
        cadena = String(longitud - 4 - Len(Trim(Numero(0))), " ") & Trim(Numero(0))
    Else
        cadena = Numero(0)
    End If
        
    
    If UBound(Numero) > 0 Then
        Numero(1) = Parte_Decimal
        cadena = cadena & "." & Left(Numero(1) & "000", 4)
    Else
        cadena = cadena & "." & "0000"
    End If

    'Para calcular el total
    Aux_Cadena = cadena
    
    Aux_Cadena = Replace(Aux_Cadena, ",", ".")
    
    If debe Then
        totalImporteDD = totalImporteDD + Abs(FormatNumber(Aux_Cadena, 2))
    Else
        totalImporteHH = totalImporteHH + Abs(FormatNumber(Aux_Cadena, 2))
    End If
    
    totalImporte = totalImporte + CDbl(Aux_Cadena)
    total = total + CDbl(cadena)
    

    
    'FGZ - 17/06/2005
    If EsUltimoItem And EsUltimoProceso Then
        
        Flog.writeline Espacios(Tabulador * 0) & "totalImporteDD: " & totalImporteDD
        Flog.writeline Espacios(Tabulador * 0) & "totalImporteHH: " & totalImporteHH
        
        'Diferencia = FormatNumber(Abs(TotalABS) - Abs(Aux_Cadena), 4)
        'Diferencia = Abs(FormatNumber(totalImporte, 4)) - Abs(FormatNumber(Aux_Cadena, 2))
        Diferencia = Abs(FormatNumber(totalImporteDD, 2)) - Abs(FormatNumber(totalImporteHH, 2))
        Flog.writeline Espacios(Tabulador * 0) & "diferencia: " & Diferencia
        If Diferencia <> 0 Then
            'totalImporte = FormatNumber(totalImporte - Abs(CDbl(Aux_Cadena)), 4)
            'total = FormatNumber(total - CDbl(cadena), 4)
            If debe Then
                If Diferencia < 0 Then
                    cadena = CDbl(cadena) + Abs(CDbl(Diferencia))
                    dh = "D"
                Else
                    cadena = CDbl(cadena) - Abs(CDbl(Diferencia))
                    dh = "H"
                End If
            Else
                If Diferencia < 0 Then
                    cadena = CDbl(cadena) - Abs(CDbl(Diferencia))
                    dh = "D"
                Else
                    cadena = CDbl(cadena) + Abs(CDbl(Diferencia))
                    dh = "H"
                End If
            
            End If
            
            Balancea = True
        Else
            If debe Then
                TotalABS = FormatNumber(TotalABS + CDbl(Aux_Cadena), 2)
            Else
                TotalABS = FormatNumber(TotalABS - CDbl(Aux_Cadena), 2)
            End If
            Balancea = True
        End If
    Else
        Balancea = True
        If debe Then
            TotalABS = FormatNumber(TotalABS + CDbl(Aux_Cadena), 4)
        Else
            TotalABS = FormatNumber(TotalABS - CDbl(Aux_Cadena), 4)
        End If
    End If
Loop
        

    cadena = CStr(FormatNumber(cadena, 2))
    
    cadena = Replace(cadena, ",", "")
    cadena = Replace(cadena, ".", "")
    cadena = Replace(cadena, ",", "")
    cadena = Replace(cadena, "-", "")
    
    cadena = String(longitud - Len(Trim(CStr(cadena))), " ") & CStr(cadena)

    Str_Salida = cadena
    

End Sub

Public Sub Comprobante(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)
'--------------------------------------------------------------------------------
'  Descripci¢n: devuelve el Nro. de sucursal. Caso particular para BIA
'  Autor: Fernando Favre
'  Fecha: 28/03/2006
'-------------------------------------------------------------------------------
Dim cadena As String
    
    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If

    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
            'cadena = cadena & String(Longitud - Len(cadena), " ")
        End If
    End If
    Str_Salida = cadena

End Sub

Public Sub CuentaL(ByVal Linea As String, ByVal masinro As String, ByRef Cuenta As String)
Dim rs_linea As New ADODB.Recordset

StrSql = "SELECT linacuenta FROM mod_linea WHERE masinro = " & masinro & " AND linaorden = " & Linea
OpenRecordset StrSql, rs_linea
If Not rs_linea.EOF Then
    Cuenta = rs_linea!linacuenta
End If
End Sub

Public Sub Validacion(ByRef programa As String, ByVal longitud As Long, ByRef cadena As String, ByVal pliqnro As Long, ByVal volcado As Long, ByRef asiento As String)
    
    If Len(programa) > 18 Then
        If Mid(programa, 22, 1) = "S" Then
            If Mid(programa, 20, 1) = "S" Then
                Call ImporteTotalHaberNegativo(True, longitud, cadena, pliqnro, volcado, asiento, True)
            Else
                Call ImporteTotalHaberNegativo(True, longitud, cadena, pliqnro, volcado, asiento, False)
            End If
        Else
            If Mid(programa, 20, 1) = "S" Then
                Call ImporteTotalHaberNegativo(False, longitud, cadena, pliqnro, volcado, asiento, True)
            Else
                Call ImporteTotalHaberNegativo(False, longitud, cadena, pliqnro, volcado, asiento, False)
            End If
        End If
    Else
        cadena = " ERROR "
        Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & " o esta mal definido. Se debe indicar si el total es H (haber)."
    End If

End Sub
Public Sub subcta_legajo(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByVal cadenaABuscar As String, ByVal nuevaCadena As String, ByVal descLinea As String, ByRef Str_Salida As String)

Dim cadena As String
Dim rs_linea As New ADODB.Recordset

    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If
    
    If cadena = cadenaABuscar Then
        cadena = Replace(cadena, cadenaABuscar, nuevaCadena)
    Else
        If cadena <> "" Then
            'asumo que vino el legajo y los datos del empleado
            StrSql = " SELECT * From empleado "
            StrSql = StrSql & " WHERE empleg = " & cadena
            OpenRecordset StrSql, rs_linea
            If Not rs_linea.EOF Then
                cadena = ""
                If EsNulo(rs_linea!terape2) Then
                    cadena = rs_linea!terape
                Else
                    cadena = cadena & rs_linea!terape2 & " " & rs_linea!terape
                End If
                
                If EsNulo(rs_linea!ternom2) Then
                    cadena = cadena & "," & rs_linea!ternom
                Else
                    cadena = cadena & " " & rs_linea!ternom2 & " " & rs_linea!ternom
                End If
                cadena = descLinea & " " & cadena
            Else
                cadena = ""
            End If
            rs_linea.Close
            'hasta aca
        Else
            cadena = ""
        End If
    End If
    

        
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    
    Str_Salida = Replace(cadena, "#", "")
    
End Sub
Public Sub subcta_cc(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByVal cadenaABuscar As String, ByVal nuevaCadena As String, ByVal descLinea As String, ByRef Str_Salida As String)

Dim cadena As String
Dim rs_linea As New ADODB.Recordset

    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If

    If cadena = cadenaABuscar Then
        cadena = Replace(cadena, cadenaABuscar, nuevaCadena)
    Else
        'asumo que vino el centro de costo
        StrSql = " SELECT * From estructura   "
        StrSql = StrSql & " WHERE estrcodext = '" & cadena & "'"
        OpenRecordset StrSql, rs_linea
        If Not rs_linea.EOF Then
            If EsNulo(rs_linea!estrdabr) Then
                cadena = ""
            Else
                cadena = rs_linea!estrdabr
            End If
            cadena = descLinea & " " & cadena
        Else
            cadena = ""
        End If
        rs_linea.Close
        'hasta aca
    End If
    
    If cadena <> "" Then
        
    End If
    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    
    Str_Salida = Replace(cadena, "#", "")
End Sub
Public Sub nro_doc(ByVal Cuenta As String, ByVal pos As Integer, ByVal cant As Integer, ByVal Completar As Boolean, ByVal longitud As Integer, ByRef Str_Salida As String)

Dim cadena As String
Dim rs_linea As New ADODB.Recordset

    If Len(Cuenta) < cant Then
        cadena = Mid(Cuenta, pos, Len(Cuenta))
    Else
        cadena = Mid(Cuenta, pos, cant)
    End If


    'asumo que vino el centro de costo
    If cadena <> "" Then
        StrSql = " SELECT nrodoc From ter_doc  "
        StrSql = StrSql & " INNER JOIN empleado on empleado.ternro = ter_doc.ternro AND empleado.empleg=" & cadena
        StrSql = StrSql & " ORDER BY tidnro ASC "
        OpenRecordset StrSql, rs_linea
        If Not rs_linea.EOF Then
            cadena = rs_linea!NroDoc
        Else
            cadena = ""
        End If
        rs_linea.Close
    End If
    'hasta aca
    
    If Completar Then
        If Len(cadena) < longitud Then
            cadena = String(longitud - Len(cadena), "0") & cadena
        End If
    End If
    
    Str_Salida = cadena
End Sub

Public Sub leg_cc(ByVal Cuenta, ByVal Linea, ByVal masinro, ByVal posLeg, ByVal cantLeg, ByVal posCC, ByVal cantCC, ByVal completa, ByVal itemiclong, ByRef cadena)

Dim rs_linea As New ADODB.Recordset
Dim lineaAux As String

StrSql = "SELECT linacuenta FROM mod_linea "
StrSql = StrSql & " WHERE mod_linea.masinro =" & masinro
StrSql = StrSql & " AND mod_linea.linaorden =" & Linea
OpenRecordset StrSql, rs_linea
If Not rs_linea.EOF Then
    lineaAux = rs_linea!linacuenta
Else
    lineaAux = ""
End If
rs_linea.Close

If lineaAux <> "" Then
    If InStr(lineaAux, "L") > 0 Then 'tiene un legajo por lo tanto tiene prioridad
        'tengo que buscar el legajo
        cadena = Mid(Cuenta, posLeg, cantLeg)
        'con el legajo busco el documento
        If EsNulo(cadena) Then
            cadena = ""
        Else
            StrSql = " SELECT nrodoc From ter_doc  "
            StrSql = StrSql & " INNER JOIN empleado on empleado.ternro = ter_doc.ternro AND empleado.empleg=" & cadena
            StrSql = StrSql & " ORDER BY tidnro ASC "
            OpenRecordset StrSql, rs_linea
            If Not rs_linea.EOF Then
                cadena = rs_linea!NroDoc
            Else
                cadena = ""
            End If
            rs_linea.Close
        End If
    Else
        If InStr(lineaAux, "E1") > 0 Then
            'busco el c.costo
            cadena = Mid(Cuenta, posCC, cantCC)
        End If
    End If
Else
    cadena = ""
End If

End Sub

'End Sub
