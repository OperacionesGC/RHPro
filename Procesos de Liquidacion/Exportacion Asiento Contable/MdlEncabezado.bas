Attribute VB_Name = "MdlEncabezado"


Public Sub encabezado(ByVal rs_Items As Recordset, ByRef cadena As String, ByVal rs_Procesos As Recordset, ByVal rs_Periodo As Recordset)

'------------------------------------------------------------------------
' Genero el encabezado de la exportacion
'------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del encabezado del proceso de volcado "
Flog.writeline

'Cantidad_Warnings = 0
'Nro = Nro + 1 'Contador de Lineas

StrSql = "SELECT * FROM confitemicenc "
StrSql = StrSql & " INNER JOIN itemintcont ON confitemicenc.itemicnro = itemintcont.itemicnro "
If ModSalidaAsiento <> 0 Then
    StrSql = StrSql & "AND confitemicenc.moditenro = " & ModSalidaAsiento
End If
StrSql = StrSql & " ORDER BY confitemicenc.confitemicorden "
OpenRecordset StrSql, rs_Items
            
Enter = Chr(13) + Chr(10)
Fecha_Proc = Date
Aux_Linea = ""
primero = True
Do While Not rs_Items.EOF
    cadena = ""
    If rs_Items!itemicfijo Then
        If rs_Items!itemicvalorfijo = "" Then
            cadena = String(256, " ")
        Else
            cadena = rs_Items!itemicvalorfijo
        End If
    Else
        programa = UCase(rs_Items!itemicprog)
        Select Case programa
        Case "FECHA" To "FECHA YYYYYYYYYY"
            If Len(programa) >= 7 Then
                Formato = Mid(programa, 7, Len(programa) - 6)
            Else
                Formato = "DDMMYYYY"
            End If
                
            Select Case Formato
                Case "YYYDDD":
                    Call Fecha1(rs_Procesos!vol_fec_asiento, cadena)
                Case Else
                    Call Fecha_Estandar(rs_Procesos!vol_fec_asiento, Formato, True, rs_Items!itemiclong, cadena)
            End Select
        Case "HEADHALLISAP":
            Call Archivo_ASTO_SAP(directorio, rs_Periodo!pliqhasta)
            cadena = ";Company Code;6055;;Control Totals" + Enter
            cadena = cadena + ";Posting Date;" + Format(Fecha_Proc, "MM/DD/YY") + Enter
            cadena = cadena + ";Document Date;" + Format(Fecha_Proc, "MM/DD/YY") + Enter
            cadena = cadena + ";Reversal Entry Date" + Enter
            cadena = cadena + ";Document Type;SA" + Enter
            cadena = cadena + ";Currency;ARS" + Enter
            cadena = cadena + ";Reference Document;Sueldos" + Enter
            cadena = cadena + ";Document Header;Sueldos" + Enter
            cadena = cadena + ";Calculate Tax (Put X)" + Enter
        Case "LINEHALLISAP":
            cadena = "Line # ; SAP G/L Account ; Amount ; Tax Code ; Cost Center ; Internal Order ; Profit Center ; Personnel Number ; Intercompany ; Allocation ; Line Item Text ; Quantity ; UoM ; WBS Element ; Network ; Activity ; TP Profit Center ; Trading Partner ; Settlement Period ; Tax Jur code ; Asset Trans Type ; Tax Tran Type"
        Case "FECHA MYYYY":
            Call Fecha4(rs_Procesos!vol_fec_asiento, cadena)
        Case "ESPACIOS":
            cadena = String(rs_Items!itemiclong, " ")
        Case "FECHAACTUAL" To "FECHAACTUAL YYYYYYYY"
             If Len(programa) >= 13 Then
                 Formato = Mid(programa, 13, Len(programa) - 6)
             Else
                 Formato = "DDMMYYYY"
             End If
             Select Case Formato
                 Case "YYYDDD":
                      Call Fecha1(Date, cadena)
                 Case Else
                      Call Fecha_Estandar(Date, Formato, True, rs_Items!itemiclong, cadena)
             End Select
        Case "PROCESO":
            Cantidad = CLng(rs_Items!itemiclong)
            Call Leyenda(rs_Procesos!vol_desc, 1, CInt(Cantidad), True, rs_Items!itemiclong, cadena)
        Case "MODELO_NRO"
            Call Modelo_Nro(rs_Procesos!masinro, rs_Procesos!Linea, rs_Procesos!descLinea, 1, rs_Items!itemiclong, True, rs_Items!itemiclong, cadena)
        Case "VOLCOD" To "VOLCOD 99,99":
            If Len(programa) > 7 Then
                pos = CLng(InStr(1, programa, ","))
                posicion = Mid(programa, 8, pos - 8)
                Cantidad = Mid(programa, pos + 1, Len(programa) - pos)
                Call nrovolcod(rs_Procesos!vol_cod, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
            Else
                posicion = "1"
                Cantidad = rs_Items!itemiclong
                Call nrovolcod(rs_Procesos!vol_cod, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
            End If
        'LED - 04/07/2012
        Case "ENTER"
            cadena = Enter
            vinoEnter = True
        'LED - 04/07/2012 - Fin
        'FGZ - 01/08/2013 -------------------------
        Case "TAB 1" To "TAB 9":
            If Len(programa) > 4 Then
                Cantidad = Mid(programa, 5, 1)
            Else
                Cantidad = "1"
            End If
            cadena = String(CLng(Cantidad), Chr(9))
        'FGZ - 01/08/2013 -------------------------
        
        Case "IMPORTETOTAL":
            Call ImporteTotal(True, rs_Items!itemiclong, cadena)
        
        'LED 21/08/2012
        'Case "IMPORTETOTALDH A,A" To "IMPORTETOTALDH Z,Z":
        'FGZ - cambio de nombre porque el itema ya existia para el PIE
        Case "TOTALDEBEHABER A,A" To "TOTALDEBEHABER Z,Z":
            Fecha = Trim(Mid(programa, 16, 1))
            completa = (UCase(Mid(programa, 18, 1)) = "S")
            Select Case Fecha
                Case "D":
                    'Call ImporteTotalDebeHaber(True, Completa, rs_Items!itemiclong, cadena) 'NG
                    Call ImporteTotalDebeHaber(True, completa, rs_Items!itemiclong, cadena, nroliq, ProcVol, Asinro)
                Case "H":
                    'Call ImporteTotalDebeHaber(False, Completa, rs_Items!itemiclong, cadena) 'NG
                    Call ImporteTotalDebeHaber(False, completa, rs_Items!itemiclong, cadena, nroliq, ProcVol, Asinro)
            End Select
        'LED 21/08/2012 - Fin
        
        Case Else
            cadena = " ERROR "
            Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
        End Select
    End If
        
    'If Mid(cadena, 1, 2) <> "RR" Or primero Then 'Comentado versión 1.31
    If primero Then
         If Aux_Linea = "" Then
            Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
         Else
            If cadena <> Enter And vinoEnter Then
                Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
                vinoEnter = False
            Else
                Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
            End If
         
            'Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
            
         End If
    Else
        'Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong) 'Comentado versión 1.31
        If cadena <> Enter And vinoEnter Then
            Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
            vinoEnter = False
        Else
            Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
        End If
        
        'Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong) ' Agregado versión 1.31
    End If
    primero = False
    rs_Items.MoveNext
Loop
If Not primero Then
    Progreso = Progreso + 1
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
End If

'Escribo en el archivo de texto
If Trim(Aux_Linea) <> "" Then
   'fExport.writeline Aux_Linea '& Aux_Relleno
   fAuxiliarEncabezado.writeline Aux_Linea '& Aux_Relleno
   primero = True
End If

End Sub
