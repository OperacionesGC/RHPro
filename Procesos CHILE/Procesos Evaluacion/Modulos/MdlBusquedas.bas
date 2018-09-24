Attribute VB_Name = "MdlBusquedas"
Option Explicit


Public Sub bus_AntEstructura(ByVal NroProg As Long, ByRef Dias As Integer, Meses As Integer, ByRef Anios As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Antiguedad en la Estructura a una Fecha
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim TipoEstr As Long      ' Tipo de Estructura
Dim TipoFecha As Integer    ' 1 - Primer dia del año
                            ' 2 - Ultimo dia del año
                            ' 3 - Inicio del proceso
                            ' 4 - Fin del proceso
                            ' 5 - Inicio del Periodo
                            ' 6 - Fin del periodo
                            ' 7 - Today
Dim Resultado As Integer    ' Tipo de resultado devuelto
                            ' 1 - En dias
                            ' 2 - En Meses
                            ' 3 - En Años
                            
Dim Param_cur As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim Aux_Fecha As Date
Dim FechaDesde As Date
Dim FechaHasta As Date

Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
   
   
    'Obtener los parametros de la Busqueda
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    OpenRecordset StrSql, Param_cur
    
    If Not Param_cur.EOF Then
        TipoEstr = Param_cur!Auxint1
        TipoFecha = Param_cur!Auxint2
    Else
        Exit Sub
    End If

    Select Case TipoFecha
    Case 1: 'Primer dia del año
        Aux_Fecha = CDate("01/01/" & Year(Date))
    Case 2: 'Ultimo dia del año
        Aux_Fecha = CDate("31/12/" & Year(Date))
    Case 3: 'Inicio del proceso
        Aux_Fecha = CDate("31/12/" & Year(Date))
    Case 4: 'Fin del proceso
        Aux_Fecha = CDate("31/12/" & Year(Date))
    Case 5: 'Inicio del periodo
        Aux_Fecha = CDate("31/12/" & Year(Date))
    Case 6: 'Fin del periodo
        Aux_Fecha = CDate("31/12/" & Year(Date))
    Case 7: 'Today
        Aux_Fecha = Date
    Case Else
        'tipo de fecha no valido
    End Select

    If Not EsNulo(Aux_Fecha) Then
        'Busco de estructura
        StrSql = " SELECT htetdesde,htethasta FROM his_estructura " & _
                 " WHERE ternro = " & Tercero & " AND " & _
                 " tenro =" & TipoEstr & " AND " & _
                 " (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND " & _
                 " ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
    
        If Not rs_Estructura.EOF Then
            FechaDesde = rs_Estructura!htetdesde
            FechaHasta = IIf(EsNulo(rs_Estructura!htethasta), Date, rs_Estructura!htethasta)
        End If
        
        Call Dif_Fechas(FechaDesde, Aux_Fecha, aux1, aux2, aux3)
        Dias = Dias + aux1
        Meses = Meses + aux2 + Int(Dias / 30)
        Anios = Anios + aux3 + Int(Meses / 12)
        Dias = Dias Mod 30
        Meses = Meses Mod 12
    End If
    
    
    
Fin:
'Cierro todo y libero
If Param_cur.State = adStateOpen Then Param_cur.Close
Set Param_cur = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing
End Sub



Public Sub bus_Grado()
' ---------------------------------------------------------------------------------------------
' Descripcion: Grado del Empleado
' Autor      : FGZ
' Fecha      : 27/06/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Empleado As New ADODB.Recordset
    
    'Busco los datos del empleado
    StrSql = "SELECT * FROM empleado WHERE ternro = " & Tercero
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el empleado."
        Exit Sub
    Else
        Valor = IIf(Not EsNulo(rs_Empleado!granro), rs_Empleado!granro, 0)
    End If
    
'Cierro todo y libero
If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
Set rs_Empleado = Nothing
End Sub


Public Sub CargarValoresdelaGrilla(ByVal rs As ADODB.Recordset, ByRef Arreglo)
' ---------------------------------------------------------------------------------------------
' Descripcion: Llena un arreglo con los valores de los registros de ValGrilla.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
rs.MoveFirst
i = 1
    Do While Not rs.EOF
        If Not EsNulo(rs!vgrvalor) Then
            Arreglo(i) = rs!vgrvalor
            i = i + 1
        End If
        
        rs.MoveNext
    Loop

End Sub




Public Sub bus_RDP(ByVal NroGrilla As Long, ByVal Cero_No_Encuentra As Boolean, ByVal Valor_Grilla, ByVal Operacion As Integer, ByVal Acumulativa As Boolean, ByVal tipoBus As Long, ByVal concnro As Long, ByVal Prog As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim ValAnt As Single
Dim Busq As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim Continuar As Boolean
Dim Parametros(5) As Integer
Dim grilla_val(10) As Single     ' para alojar los valores de:  valgrilla.val(i)

Dim TipoBase As Long
Dim TipoBaseVariable As Long

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim rs_Busqueda As New ADODB.Recordset

Dim NroBusqueda As Long
Dim TipoBusqueda As Long
Dim Encontro As Boolean
Dim Antdia As Integer
Dim Antmes As Integer
Dim Antanio As Integer

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    'El tipo Base de la antiguedad
    TipoBase = 4
    TipoBaseVariable = 15
    
    Continuar = True
    ant = 1
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Grilla de " & rs_cabgrilla!cgrdimension & "dimensiones "
    End If
    
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop

'Parametros Variables
' busco que parametro es el parametro del concepto
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "busco que parametro es el parametro del concepto "
    End If

    Continuar = True
    pvar = 1
    Do While (pvar <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case pvar
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBaseVariable = rs_tbase!tprogbase Then
                    Continuar = False
                    pvariable = True
                Else
                    pvar = pvar + 1
                End If
            End If
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBaseVariable = rs_tbase!tprogbase Then
                    Continuar = False
                    pvariable = True
                Else
                    pvar = pvar + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBaseVariable = rs_tbase!tprogbase Then
                    Continuar = False
                    pvariable = True
                Else
                    pvar = pvar + 1
                End If
            End If
        Case 4:
            If rs_cabgrilla!grparnro_4 = 15 Then 'si es el param. variable
                Continuar = False
                pvariable = True
            Else
                pvar = pvar + 1
            End If
        Case 5:
            If rs_cabgrilla!grparnro_5 = 15 Then 'si es el param. variable
                Continuar = False
                pvariable = True
            Else
                pvar = pvar + 1
            End If
        End Select
    Loop

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Resuelvo los indices de la grilla segun las busquedas por cada dimension"
    End If
   
    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad en la estructura
            Select Case j
            Case 1:
                Call bus_AntEstructura(rs_cabgrilla!grparnro_1, Antdia, Antmes, Antanio)
            Case 2:
                Call bus_AntEstructura(rs_cabgrilla!grparnro_2, Antdia, Antmes, Antanio)
            Case 3:
                Call bus_AntEstructura(rs_cabgrilla!grparnro_3, Antdia, Antmes, Antanio)
            Case 4:
                Call bus_AntEstructura(rs_cabgrilla!grparnro_4, Antdia, Antmes, Antanio)
            Case 5:
                Call bus_AntEstructura(rs_cabgrilla!grparnro_5, Antdia, Antmes, Antanio)
            End Select
            Parametros(j) = (Antanio * 12) + Antmes
        Case Else:
            Select Case j
            Case 1:
                'Call bus_Estructura(rs_cabgrilla!grparnro_1)
                Call bus_Grado
            Case 2:
                'Call bus_Estructura(rs_cabgrilla!grparnro_2)
                Call bus_Grado
            Case 3:
                'Call bus_Estructura(rs_cabgrilla!grparnro_3)
                Call bus_Grado
            Case 4:
                'Call bus_Estructura(rs_cabgrilla!grparnro_4)
                Call bus_Grado
            Case 5:
                'Call bus_Estructura(rs_cabgrilla!grparnro_5)
                Call bus_Grado
            End Select
            Parametros(j) = Valor
        End Select
    Next j

    If Not antig Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No busca antiguedad "
        End If
    
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant And j <> pvar Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            Else
                If pvariable Then
                    StrSql = StrSql & " AND vgrcoor_" & j & "<= " & Parametros(j)
                End If
            End If
        Next j
        If pvariable Then
            StrSql = StrSql & " ORDER BY vgrcoor_" & pvar & " DESC "
        End If
        OpenRecordset StrSql, rs_valgrilla
    
        If Not rs_valgrilla.EOF Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Cargo los Valores de la Grilla "
            End If
            Call CargarValoresdelaGrilla(rs_valgrilla, grilla_val)
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco el valor segun la operacion "
            End If
            Call BusValor(Operacion, Valor_Grilla, grilla_val, Valor)
        Else
            If Cero_No_Encuentra Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontró valor en grilla "
                    Flog.writeline Espacios(Tabulador * 4) & "Esta configurado que retorne cero si no lo encuentra "
                End If
            
                 Valor = 0
                 Bien = True
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontró valor en grilla "
                    Flog.writeline Espacios(Tabulador * 4) & "Retorna Falso "
                End If
                Bien = False
            End If
       End If
    Else 'Antig
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busca antiguedad "
        End If
    
        If Not Cero_No_Encuentra Then
            Bien = False
        Else
            Bien = True
            Valor = 0
        End If
    
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco la primera antiguedad de la escala menor a la del empleado "
            Flog.writeline Espacios(Tabulador * 4) & "de abajo hacia arriba "
        End If
    
        'Busco la primera antiguedad de la escala menor a la del empleado
        ' de abajo hacia arriba
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            End If
        Next j
            StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
        OpenRecordset StrSql, rs_valgrilla
        
        Encontro = False
        
        Do While Not rs_valgrilla.EOF And Not Encontro
            Select Case ant
            Case 1:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 2:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 3:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 4:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 5:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            End Select
            '--------------------------
            
            rs_valgrilla.MoveNext
        Loop
        If CBool(USA_DEBUG) Then
            If Encontro Then
                Flog.writeline Espacios(Tabulador * 4) & "Valor encontrado "
            Else
                Flog.writeline Espacios(Tabulador * 4) & "Valor no encontrado "
                If Cero_No_Encuentra Then
                    Flog.writeline Espacios(Tabulador * 4) & "Esta configurado que retorne cero si no lo encuentra "
                Else
                    Flog.writeline Espacios(Tabulador * 4) & "No Esta configurado que retorne cero si no lo encuentra. Retorna Falso "
                End If
            End If
        End If
    End If
    
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub


Public Sub BusValor(ByVal Op As Integer, ByVal valorGrilla, ByVal valgrilla, ByRef Valor As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula el valor con los eltos del arreglo con los valores de los registros de ValGrilla.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim cant As Integer
Dim Continuar As Boolean
Dim i As Integer

Select Case Op
Case 1:     'Sumatoria
    Valor = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            Valor = Valor + valgrilla(i)
        End If
    Next i

Case 2:     'Maximo
    Valor = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            If valgrilla(i) > Valor Then
                Valor = valgrilla(i)
            End If
        End If
    Next i
    
Case 3:     'Promedio
    Valor = 0
    cant = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            Valor = Valor + valgrilla(i)
            cant = cant + 1
        End If
    Next i

    If cant <> 0 Then
        Valor = Valor / cant
    End If

Case 4:     'Promedio sin cero
    Valor = 0
    cant = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            If valgrilla(i) <> 0 Then
                Valor = Valor + valgrilla(i)
                cant = cant + 1
            End If
        End If
    Next i

    If cant <> 0 Then
        Valor = Valor / cant
    End If

Case 5:     'Minimo
    Valor = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            If Valor = 0 Or valgrilla(i) < Valor Then
                Valor = valgrilla(i)
            End If
        End If
    Next i

Case 6:     'Primer valor no vacio desde el primero
    Valor = 0
    i = 1
    Continuar = True
    Do While i <= 10 And Continuar
        If valorGrilla(i) Then
            If valgrilla(i) <> 0 Then
                Valor = valgrilla(i)
                Continuar = False
            End If
        End If
        i = i + 1
    Loop

Case 7:     'Primer valor no vacio desde el ultimo
    Valor = 0
    i = 10
    Continuar = True
    Do While i >= 0 And Continuar
        If valorGrilla(i) Then
            If valgrilla(i) <> 0 Then
                Valor = valgrilla(i)
                Continuar = False
            End If
        End If
        i = i - 1
    Loop
End Select
End Sub

