Attribute VB_Name = "mdlImportacionInterpay"
Public Sub import(ByVal strLinea As String, ByVal nroLinea As Long, ByVal separador As String)
Dim arrayLinea

Dim codMovimiento As String

    Flog.writeline Espacios(Tabulador * 0) & "Comienzo de la importacion para la linea: " & nroLinea & "."
    arrayLinea = Split(strLinea, separador)
    
    '-----------------------------------------------------------------------------------------------------------------------------
    codMovimiento = arrayLinea(0)
    Select Case UCase(codMovimiento)
        Case "ADDR":
            Call movADDR(arrayLinea)
        Case "ASGC":
            Call movASGC(arrayLinea)
        Case "COST":
            Call movCOST(arrayLinea)
        Case "HIRE":
            Call movHIRE(arrayLinea)
        Case "IDNO":
            Call movIDNO(arrayLinea)
        Case "PERS":
            Call movPERS(arrayLinea)
        Case "REVH":
            Call movREVH(arrayLinea)
        Case "REVT":
            Call movREVT(arrayLinea)
        Case "SLRY":
            Call movSLRY(arrayLinea)
        Case "TERM":
            Call movTERM(arrayLinea)
        Case "GHB":
            Call movGHB(arrayLinea)
        Case "LSA":
            Call movLSA(arrayLinea)
        Case "RNR":
            Call movRNR(arrayLinea)
    
    End Select

    Flog.writeline Espacios(Tabulador * 0) & "Fin importacion, Nro de linea: " & nroLinea & "."
    Flog.writeline Espacios(Tabulador * 0) & "-----------------------------------------------------------."
    
    
End Sub



Sub movADDR(ByVal arrayLinea)
'Changes - ADDR (Cambios de dirección)
Dim ternro As Long
Dim Pais As String
Dim Provincia As String
Dim zona As String
Dim partido As String
Dim Localidad As String
Dim Barrio As String
Dim entreCalles As String
Dim cp As String
Dim manzana As String
Dim torre As String
Dim depto As String
Dim piso As String
Dim Numero As String
Dim calle As String
Dim domnro As String
Dim paisnro As String
Dim locnro As String
Dim provnro As String


    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento ADDR"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    ternro = existeEmpleado(arrayLinea(4))
    If ternro <> 0 Then
        calle = Left(arrayLinea(23), 100)    'cortamos a 30 caracteres que es la long maxima del campo
        
        Numero = "-"    'Split(arrayLinea(23), ";")(1) 'por pedido del cliente informamos un "-" en este campo
        piso = ""   'Split(arrayLinea(23), ";")(2)
        depto = ""  'Split(arrayLinea(23), ";")(3)
        torre = ""  'Split(arrayLinea(23), ";")(4)

        manzana = ""    'Split(arrayLinea(24), ";")(0)
        entreCalles = ""    'Split(arrayLinea(24), ";")(1)
        
        Barrio = ""     'Split(arrayLinea(25), ";")(0)
        
        Provincia = arrayLinea(25)
        Localidad = arrayLinea(26)
                
        Pais = arrayLinea(29)
        
        If Trim(Pais) = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio de domicilio el pais debe ser informado."
            Exit Sub
        End If
        
        cp = arrayLinea(28)
        Call verificarLocalidad(Localidad, locnro)
        Call crearPais(Pais, Provincia, paisnro, provnro)
        Call actualizarDomicilio(ternro, paisnro, provnro, zona, partido, locnro, Barrio, entreCalles, cp, manzana, torre, depto, piso, Numero, calle, domnro)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio de domicilio el empleado debe existir."
        Exit Sub
    End If
    
    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento ADDR"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"

End Sub
Sub movREVH(ByVal arrayLinea)
Dim ternro As Long
Dim fechaAlta As String

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento REVH"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    ternro = existeEmpleado(arrayLinea(4))
    
    If ternro <> 0 Then
        fechaAlta = armarfecha(arrayLinea(1))
        
        Call actualizarFase(ternro, fechaAlta, "", "NULL")
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio Fecha de alta el empleado debe existir."
        Exit Sub
    End If
    
    Flog.writeline Espacios(Tabulador * 1) & "FIN Movimiento REVH"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
End Sub
Sub movREVT(ByVal arrayLinea)
Dim ternro As Long
Dim fechaBajaPrevista As String

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento REVT"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    ternro = existeEmpleado(arrayLinea(4))
    
    If ternro <> 0 Then
        fechaBajaPrevista = armarfecha(arrayLinea(2))
        
        Call bajaContrato(ternro, fechaBajaPrevista)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio Fecha de baja prevista el empleado debe existir."
        Exit Sub
    End If

    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento REVT"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
End Sub

Sub movIDNO(ByVal arrayLinea)
Dim rs_aux As New ADODB.Recordset

Dim ternro As Long

Dim tipoDni As String
Dim dni As String
Dim tipoCuil As String
Dim Cuil As String
Dim errorCuil As String

Dim AFJP As String
Dim afjpTenro As String
Dim AFJPEstrnro As Long
Dim AFJPTernro As Long

Dim convenio As String
Dim convenioTenro As String
Dim convenioEstrnro As Long
Dim convenioTernro As Long

Dim osElegida As String
Dim osElegidaTenro As String
Dim osElegidaEstrnro As Long
Dim osElegidaTernro As Long

Dim nacionalNro As Long

Dim sijp As String
Dim sijpTenro As String
Dim sijpEstrnro As Long
Dim sijpTernro As Long

Dim fechaDesde As String
Dim Inserto_estr As Boolean

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento IDNO"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    ternro = existeEmpleado(arrayLinea(4))
    If ternro <> 0 Then
    
        'busco documentos configurados
        StrSql = " SELECT confnrocol, confval FROM confrep WHERE repnro = 480 AND upper(conftipo) = 'DOC' AND confnrocol in (14,15)"
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            Do While Not rs_aux.EOF
                Select Case CLng(rs_aux!confnrocol)
                    Case 14: 'DNI columna 52
                        tipoDni = rs_aux!confval
                        Flog.writeline Espacios(Tabulador * 1) & "DNI configurado Tipo: " & tipoDni
                
                    Case 15: 'CUIL columna 51
                        tipoCuil = rs_aux!confval
                        Flog.writeline Espacios(Tabulador * 1) & "Cuil configurado Tipo: " & tipoCuil
                End Select
                
                rs_aux.MoveNext
            Loop
        Else
            Flog.writeline Espacios(Tabulador * 1) & "No hay documentos configurados"
        End If
        
        dni = arrayLinea(51)
        If Trim(dni) <> "" Then
            Call actualizarDocumento(ternro, tipoDni, dni)
        Else
            Flog.writeline Espacios(Tabulador * 1) & "DNI no informado."
        End If
        
        'busco la nacionalidad del empleado para validar el cuil
        StrSql = " SELECT nacionalnro FROM tercero WHERE ternro = " & ternro
        OpenRecordset StrSql, rs_aux
        If Not rs_aux Then
            nacionalNro = rs_aux!nacionalNro
        End If
        
        Cuil = arrayLinea(50)
        If Trim(Cuil) <> "" Then
            If Cuil_Valido605(Cuil, dni, errorCuil, "DNI", nacionalNro) Then
                Call actualizarDocumento(ternro, tipoCuil, Cuil)
                Flog.writeline Espacios(Tabulador * 1) & "Cuil actualizado."
            Else
                Flog.writeline Espacios(Tabulador * 1) & errorCuil
                Cuil = ""
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Cuil no informado."
        End If
        
        'seccion estructuras
        'recupero las estructuras configuradas en el reporte (480)
        StrSql = " SELECT confnrocol, confval FROM confrep WHERE repnro = 480 AND upper(conftipo) = 'TE' AND confnrocol in (10,11,12,13)"
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            Do While Not rs_aux.EOF
                Select Case CLng(rs_aux!confnrocol)
                    Case 10: 'AFJP columna 53
                        afjpTenro = rs_aux!confval
                        Flog.writeline Espacios(Tabulador * 1) & "AFJP configurado TE: " & unidadNegocioTenro

                    Case 11: 'Obra social columna 55
                        osElegidaTenro = rs_aux!confval
                        Flog.writeline Espacios(Tabulador * 1) & "Obra social configurado TE: " & cargoTenro

                    Case 12: 'SIJP columna 56
                        sijpTenro = rs_aux!confval
                        Flog.writeline Espacios(Tabulador * 1) & "SIJP configurado TE: " & formaPagoTenro
 
                    Case 13: 'Convenio columna 54
                        convenioTenro = rs_aux!confval
                        Flog.writeline Espacios(Tabulador * 1) & "Convenio configurado TE: " & bandaTenro
                End Select
                
                rs_aux.MoveNext
            Loop
        Else
            Flog.writeline Espacios(Tabulador * 1) & "No hay estructuras configuras"
        End If
        
        fechaDesde = armarfecha(arrayLinea(1))
        
        AFJP = arrayLinea(52)
        convenio = arrayLinea(53)
        osElegida = arrayLinea(54)
        sijp = arrayLinea(55)
            
        If AFJP <> "" Then
            Call ValidaEstructuraCodExt(afjpTenro, AFJP, AFJPEstrnro, Inserto_estr)
            If AFJPEstrnro <> 0 Then
                Call VerSiCrearTercero(afjpTenro, AFJP, AFJPTernro)
                Call AsignarEstructura(afjpTenro, AFJPEstrnro, ternro, fechaDesde, "Null")
                If Inserto_estr Then
                    Call VerSiCrearComplemento(afjpTenro, AFJPEstrnro, AFJP, AFJPTernro)
                End If
                Flog.writeline Espacios(Tabulador * 1) & "AFJP actualizado."
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "AFJP no informado."
        End If
        
        If convenio <> "" Then
            Call ValidaEstructuraCodExt(convenioTenro, convenio, convenioEstrnro, Inserto_estr)
            If convenioEstrnro <> 0 Then
                Call VerSiCrearTercero(convenioTenro, convenio, convenioTernro)
                Call AsignarEstructura(convenioTenro, convenioEstrnro, ternro, fechaDesde, "Null")
                If Inserto_estr Then
                    Call VerSiCrearComplemento(convenioTenro, convenioEstrnro, convenio, convenioTernro)
                End If
                Flog.writeline Espacios(Tabulador * 1) & "Convenio actualizado."
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Convenio no informado."
        End If
        
        If osElegida <> "" Then
            Call ValidaEstructuraCodExt(osElegidaTenro, osElegida, osElegidaEstrnro, Inserto_estr)
            If osElegidaEstrnro <> 0 Then
                Call VerSiCrearTercero(osElegidaTenro, osElegida, osElegidaTernro)
                Call AsignarEstructura(osElegidaTenro, osElegidaEstrnro, ternro, fechaDesde, "Null")
                If Inserto_estr Then
                    Call VerSiCrearComplemento(osElegidaTenro, osElegidaEstrnro, osElegida, osElegidaTernro)
                End If
                Flog.writeline Espacios(Tabulador * 1) & "Obra Social actualizada."
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Obra Social no informada."
        End If
        
        If sijp <> "" Then
            Call ValidaEstructuraCodExt(sijpTenro, sijp, sijpEstrnro, Inserto_estr)
            If sijpEstrnro <> 0 Then
                Call VerSiCrearTercero(sijpTenro, sijp, sijpTernro)
                Call AsignarEstructura(sijpTenro, sijpEstrnro, ternro, fechaDesde, "Null")
                If Inserto_estr Then
                    Call VerSiCrearComplemento(sijpTenro, sijpEstrnro, sijp, sijpTernro)
                End If
                Flog.writeline Espacios(Tabulador * 1) & "SIJP actualizado."
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "SIJP no informado."
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio de documentos o identificacion el empleado debe existir."
        Exit Sub
    End If
    
    Flog.writeline Espacios(Tabulador * 1) & "FIN Movimiento IDNO"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
End Sub

Sub movPERS(ByVal arrayLinea)
'Cambios de información personal
Dim rs_datos As New ADODB.Recordset
Dim rs_datosAux As New ADODB.Recordset
Dim ternro As Long
Dim fechaAlta As String
Dim Legajo As String
Dim apellido As String
Dim nombre As String
Dim sexo As Integer
Dim fechaNacimiento As String
Dim estCivil As String
Dim estCivilNro As Long
Dim paisNacimiento As String
Dim email As String
Dim provnro As String
Dim fechaAntiguedadRec As String
Dim fechaIngreso As String
    
    ternro = existeEmpleado(arrayLinea(4))
    Legajo = arrayLinea(4)
    If ternro <> 0 Then
        '------------------------------------------------------------------------------------------
        'apellido
        If arrayLinea(7) <> "" Then
            apellido = arrayLinea(7)
            Flog.writeline Espacios(Tabulador * 1) & "Apellido obtenido."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Apellido no informado."
        End If
        '------------------------------------------------------------------------------------------
        'Nombre
        If arrayLinea(8) <> "" Then
            nombre = Split(arrayLinea(8), " ")(0)
            Flog.writeline Espacios(Tabulador * 1) & "Primer Nombre obtenido."
            If UBound(Split(arrayLinea(8), " ")) > 0 Then
                nombre2 = Split(arrayLinea(8), " ")(1)
                Flog.writeline Espacios(Tabulador * 1) & "Segundo Nombre obtenido."
            Else
                nombre2 = ""
                Flog.writeline Espacios(Tabulador * 1) & "El empleado no tiene segundo Nombre."
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Nombre no informado."
            nombre = ""
            nombre2 = ""
        End If
        
        '------------------------------------------------------------------------------------------
        'Fecha de Alta
        If arrayLinea(1) <> "" Then
            fechaAlta = armarfecha(arrayLinea(1))
            Flog.writeline Espacios(Tabulador * 1) & "Fecha de alta obtenida."
        Else
            FecAlta = ""
            Flog.writeline Espacios(Tabulador * 1) & "Fecha de alta no informada."
        End If
        
        '------------------------------------------------------------------------------------------
        'Sexo
        If arrayLinea(12) <> "" Then
            If UCase(arrayLinea(12)) = "M" Then
                sexo = -1
            Else
                sexo = 0
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Sexo informado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Sexo no informado."
        End If

        '------------------------------------------------------------------------------------------
        'Fecha de Nacimiento
        If arrayLinea(13) <> "" Then
            fechaNacimiento = armarfecha(arrayLinea(13))
            Flog.writeline Espacios(Tabulador * 1) & "Fecha Nacimiento informada."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Fecha Nacimiento no informada."
            fechaNacimiento = ""
        End If


        '------------------------------------------------------------------------------------------
        'Estado Civil
        If arrayLinea(14) <> "" Then
            'StrSql = "SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(Mid(arrayLinea(14), 1, 30)) & "'"
            StrSql = "SELECT codexterno FROM mapeo_sap WHERE upper(infotipo) = 'IP_HIRE' AND UPPER(tablaref) = 'ESTCIVIL' AND UPPER(codinterno) = '" & UCase(Mid(arrayLinea(14), 1, 30)) & "'"
            OpenRecordset StrSql, rs_datos
            
            If Not rs_datos.EOF Then
                estCivil = Mid(rs_datos!codexterno, 1, 30)
                StrSql = "SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(estCivil) & "'"
                OpenRecordset StrSql, rs_datosAux
                If Not rs_datosAux.EOF Then
                    estCivilNro = rs_datosAux!estcivnro
                    Flog.writeline Espacios(Tabulador * 1) & "Estado Civil encontrado."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Estado Civil no encontrado, se creara."
                    estCivilNro = crearEstadoCivil(UCase(estCivil))
                End If
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Error: Estado Civil no configurado (mapeo)."
                HuboError = True
            End If
        End If

        '------------------------------------------------------------------------------------------
        'Pais Nacimiento
        If arrayLinea(15) <> "" Then
                Call crearPais(UCase(arrayLinea(15)), "", paisNacimiento, provnro)
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Pais de Nacimiento no informado."
        End If

        '------------------------------------------------------------------------------------------
        'Email
        If arrayLinea(34) <> "" Then
            email = arrayLinea(34)
            Flog.writeline Espacios(Tabulador * 1) & "Email encontrado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Email no informado."
            email = ""
        End If

        Call actualizarEmpleado(ternro, fechaAlta, "", paisNacimiento, Legajo, "", email, "", fechaNacimiento, sexo, estCivilNro, nombre, nombre2, apellido, "")
        
        fechaAntiguedadRec = armarfecha(arrayLinea(16))
        fechaIngreso = armarfecha(arrayLinea(17))

        Call actualizarFase(ternro, fechaAntiguedadRec, fechaIngreso, "NULL")
        
        Flog.writeline Espacios(Tabulador * 1) & "Datos generales del empleado Actualizados."
        
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para actualizar un empleado debe existir."
        Exit Sub
    End If


End Sub

Sub movSLRY(ByVal arrayLinea)
Dim rs_aux As New ADODB.Recordset
'Cambios de Salarios
Dim ternro As Long
Dim salario As String
Dim fechaDesde As String
Dim formaPago As String
Dim formaPagoTenro As String
Dim formaPagoTernro As Long
Dim formaPagoEstrnro As Long
Dim Conccod As String
Dim tpanro As Long
Dim Inserto_estr As Boolean

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento SLRY"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    ternro = existeEmpleado(arrayLinea(4))
    If ternro <> 0 Then
    
        'recupero las estructuras configuradas en el reporte (480)
        StrSql = " SELECT confnrocol, confval, confval2 FROM confrep WHERE repnro = 480 AND confnrocol in (8,17)"
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            Do While Not rs_aux.EOF
                Select Case CLng(rs_aux!confnrocol)
                    Case 8: 'Forma de Pago columna 24
                        formaPagoTenro = rs_aux!confval
                        Flog.writeline Espacios(Tabulador * 1) & "Forma de Pago configurado TE: " & centroCostoTenro
                    Case 17: 'Remuneracion a novedad columna 23
                        Conccod = rs_aux!confval2
                        tpanro = rs_aux!confval
                        Flog.writeline Espacios(Tabulador * 1) & "Novedad concepto: " & Conccod & " parametro: " & tpanro
                        
                End Select
                rs_aux.MoveNext
            Loop
        Else
            Flog.writeline Espacios(Tabulador * 1) & "No hay estructuras configuras"
        End If
        
        fechaDesde = armarfecha(arrayLinea(1))
        
        formaPago = arrayLinea(22)
        salario = arrayLinea(23)
        
        Call insertarNovedad(ternro, "", "", Conccod, tpanro, salario, 0)
        
        If formaPago <> "" Then
            Call ValidaEstructura(formaPagoTenro, formaPago, formaPagoEstrnro, Inserto_estr)
            Call VerSiCrearTercero(formaPagoTenro, formaPago, formaPagoTernro)
            Call AsignarEstructura(formaPagoTenro, formaPagoEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(formaPagoTenro, formaPagoEstrnro, formaPago, formaPagoTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Forma de Pago actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Forma de Pago no informado."
        End If
        
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio de salario el empleado debe existir."
        Exit Sub
    End If

    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento SLRY"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"

End Sub

Sub movCOST(ByVal arrayLinea)
Dim rs_aux As New ADODB.Recordset

Dim ternro As Long
Dim fechaDesde As String
Dim centroCosto As String
Dim centroCostoTenro As String
Dim centroCostoTernro As Long
Dim centroCostoEstrnro As Long
Dim Inserto_estr As Boolean

'Dim unidadNegocioTernro As Long
'Dim centroCostoTernro As Long
'Dim cargoTernro As Long
'Dim unidadNegocio As String
'Dim cargo As String     'puesto
'Dim unidadNegocioEstrnro As Long
'Dim cargoEstrnro As Long     'puesto

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento COST"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"

    ternro = existeEmpleado(arrayLinea(4))
    If ternro <> 0 Then
        fechaDesde = armarfecha(arrayLinea(1))
        
        'unidadNegocio = arrayLinea(18)
        'cargo = arrayLinea(20)
        
        'recupero las estructuras configuradas en el reporte (480)
        StrSql = " SELECT confnrocol, confval FROM confrep WHERE repnro = 480 AND upper(conftipo) = 'TE' AND confnrocol in (7)"
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            Do While Not rs_aux.EOF
                Select Case CLng(rs_aux!confnrocol)
                    Case 7: 'Centro de costo columna 22
                        centroCostoTenro = rs_aux!confval
                        Flog.writeline Espacios(Tabulador * 1) & "Centro de Costo configurado TE: " & centroCostoTenro
                End Select
                rs_aux.MoveNext
            Loop
        Else
            Flog.writeline Espacios(Tabulador * 1) & "No hay estructuras configuras"
        End If
            
        centroCosto = arrayLinea(21)
        
        If fechaDesde = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: No existe fecha desde para la estructura."
            Exit Sub
        End If
        
'        If unidadNegocio <> "" Then
'            Call ValidaEstructura(1, unidadNegocio, unidadNegocioEstrnro, Inserto_estr)
'            Call VerSiCrearTercero(1, unidadNegocio, unidadNegocioTernro)
'            Call AsignarEstructura(1, unidadNegocioEstrnro, ternro, fechaDesde, "Null")
'            If Inserto_estr Then
'                Call VerSiCrearComplemento(1, unidadNegocioEstrnro, unidadNegocio, unidadNegocioTernro)
'            End If
'            Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio actualizado."
'        Else
'            Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio no informado."
'        End If
                
'        If cargo <> "" Then
'            Call ValidaEstructura(4, cargo, cargoEstrnro, Inserto_estr)
'            Call VerSiCrearTercero(4, cargo, cargoTernro)
'            Call AsignarEstructura(4, cargoEstrnro, ternro, fechaDesde, "Null")
'            If Inserto_estr Then
'                Call VerSiCrearComplemento(4, cargoEstrnro, cargo, cargoTernro)
'            End If
'            Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) actualizado."
'        Else
'            Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) no informado."
'        End If
        
        If centroCosto <> "" Then
            Call ValidaEstructura(centroCostoTenro, centroCosto, centroCostoEstrnro, Inserto_estr)
            Call VerSiCrearTercero(centroCostoTenro, centroCosto, centroCostoTernro)
            Call AsignarEstructura(centroCostoTenro, centroCostoEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(centroCostoTenro, centroCostoEstrnro, centroCosto, centroCostoTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "centro de costo actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "centro de costo no informado."
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Estructuras de mov cost actualizados."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para cambios de estructura el empleado debe existir."
        Exit Sub
    End If

    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento COST"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"

End Sub

Sub movTERM(ByVal arrayLinea)
Dim ternro As Long
Dim fechaBajaPrevista As String
Dim codExtBaja As String

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento TERM"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    ternro = existeEmpleado(arrayLinea(4))
    
    If ternro <> 0 Then
        fechaBajaPrevista = armarfecha(arrayLinea(2))
        codExtBaja = arrayLinea(35)
        
        If fechaBajaPrevista <> "" Then
            Call cerrarFase(ternro, fechaBajaPrevista, codExtBaja)
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Fecha de baja prevista no informada."
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio Fecha de baja prevista el empleado debe existir."
        Exit Sub
    End If

    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento TERM"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
End Sub

Sub movGHB(ByVal arrayLinea)
Dim ternro As Long
Dim fechaInicio As String
Dim fechaFin As String
Dim Conccod As String
Dim Parametro As String
Dim Monto As String

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento GHB"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    ternro = existeEmpleado(arrayLinea(1))
    
    If ternro <> 0 Then
        fechaInicio = armarfecha(arrayLinea(4))
        fechaFin = armarfecha(arrayLinea(5))
        
        Conccod = Split(arrayLinea(6), ";")(0)
        Parametro = Split(arrayLinea(6), ";")(1)
        Monto = arrayLinea(7)
        
        Call insertarNovedad(ternro, fechaInicio, fechaFin, Conccod, Parametro, Monto, -1)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: el empleado debe existir para cargar la novedad."
        Exit Sub
    End If

    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento GHB"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
End Sub

Sub movLSA(ByVal arrayLinea)
Dim ternro As Long
Dim Conccod As String
Dim Parametro As String
Dim Monto As String

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento LSA"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    ternro = existeEmpleado(arrayLinea(4))
    
    If ternro <> 0 Then
        Conccod = Split(arrayLinea(6), ";")(0)
        Parametro = Split(arrayLinea(6), ";")(1)
        Monto = arrayLinea(7)
        
        Call insertarNovedad(ternro, "", "", Conccod, Parametro, Monto, 0)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: el empleado debe existir para cargar la novedad."
        Exit Sub
    End If

    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento LSA"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
End Sub
Sub movRNR(ByVal arrayLinea)
Dim ternro As Long
Dim Conccod As String
Dim Parametro As String
Dim Monto As String

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento RNR"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    ternro = existeEmpleado(arrayLinea(4))
    
    If ternro <> 0 Then
        Conccod = Split(arrayLinea(7), ";")(0)
        Parametro = Split(arrayLinea(7), ";")(1)
        Monto = arrayLinea(8)
        
        Call insertarNovedad(ternro, "", "", Conccod, Parametro, Monto, 0)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: el empleado debe existir para cargar la novedad."
        Exit Sub
    End If

    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento RNR"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
End Sub
Sub movHIRE(ByVal arrayLinea)
Dim rs_datos As New ADODB.Recordset
Dim rs_datosAux As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim ternro As Long
Dim Legajo As String
Dim reportaA As String
Dim apellido As String
Dim nombre As String
Dim sexo As String
Dim fechaNacimiento As String
Dim estCivil As String
Dim estCivilNro As String
Dim paisNacimiento As String
Dim fechaAlta As String
Dim fechaIngActual As String
Dim remuneracion As String
Dim calle As String
Dim nro As String
Dim piso As String
Dim depto As String
Dim torre As String
Dim manzana As String
Dim entreCalles As String
Dim Barrio As String
Dim Localidad As String
Dim Provincia As String
Dim provnro As String
Dim Pais As String

Dim tipoDni As String
Dim dni As String
Dim tipoCuil As String
Dim Cuil As String
Dim errorCuil As String

Dim nacionalidad As String
Dim telParticular As String
Dim telCelular As String
Dim email As String
Dim nroCuil As String
Dim nroDocumento As String
Dim fechaDesde As String
Dim zona As String
Dim partido As String
Dim cp As String
Dim Numero As String
Dim domnro As String
Dim paisnro As String
Dim locnro As String

Dim unidadNegocio As String
Dim unidadNegocioTenro As String
Dim unidadNegocioEstrnro As Long
Dim unidadNegocioTernro As Long


Dim cargo As String
Dim cargoTenro As String
Dim cargoEstrnro As Long
Dim cargoTernro As Long

Dim banda As String
Dim bandaTenro As String
Dim bandaEstrnro As Long
Dim bandaTernro As Long

Dim formaPago As String
Dim formaPagoTenro As String
Dim formaPagoEstrnro As Long
Dim formaPagoTernro As Long

Dim locationName As String
Dim locationNameTenro As String
Dim locationNameEstrnro As Long
Dim locationNameTernro As Long

Dim regimenHorario As String
Dim regimenHorarioTenro As String
Dim regimenHorarioEstrnro As Long
Dim regimenHorarioTernro As Long

Dim AFJP As String
Dim afjpTenro As String
Dim AFJPEstrnro As Long
Dim AFJPTernro As Long

Dim osElegida As String
Dim osElegidaTenro As String
Dim osElegidaEstrnro As Long
Dim osElegidaTernro As Long

Dim sijp As String
Dim sijpTenro As String
Dim sijpEstrnro As Long
Dim sijpTernro As Long

Dim industryFocusGroup As String
Dim industryFocusGroupTenro As String
Dim industryFocusGroupEstrnro As Long
Dim industryFocusGroupTernro As Long

Dim Conccod As String
Dim tpanro As String

Dim Inserto_estr As Boolean

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento HIRE"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
    
    Legajo = arrayLinea(4)
    ternro = existeEmpleado(Legajo)
    reportaA = existeEmpleado(arrayLinea(42))
    '------------------------------------------------------------------------------------------------------------
    'obtengo todos los datos del empleado
    '------------------------------------------------------------------------------------------
    'apellido
    If arrayLinea(7) <> "" Then
        apellido = arrayLinea(7)
        Flog.writeline Espacios(Tabulador * 1) & "Apellido obtenido."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Apellido no informado."
    End If
    '------------------------------------------------------------------------------------------
    'Nombre
    If arrayLinea(8) <> "" Then
        nombre = Split(arrayLinea(8), " ")(0)
        Flog.writeline Espacios(Tabulador * 1) & "Primer Nombre obtenido."
        If UBound(Split(arrayLinea(8), " ")) > 0 Then
            nombre2 = Split(arrayLinea(8), " ")(1)
            Flog.writeline Espacios(Tabulador * 1) & "Segundo Nombre obtenido."
        Else
            nombre2 = ""
            Flog.writeline Espacios(Tabulador * 1) & "El empleado no tiene segundo Nombre."
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Nombre no informado."
        nombre = ""
        nombre2 = ""
    End If
    
    '------------------------------------------------------------------------------------------
    'Sexo
    If arrayLinea(12) <> "" Then
        If UCase(arrayLinea(12)) = "M" Then
            sexo = -1
        Else
            sexo = 0
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Sexo informado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Sexo no informado."
    End If
    

    '------------------------------------------------------------------------------------------
    'Fecha de Nacimiento
    If arrayLinea(13) <> "" Then
        fechaNacimiento = armarfecha(arrayLinea(13))
        Flog.writeline Espacios(Tabulador * 1) & "Fecha Nacimiento informada."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Fecha Nacimiento no informada."
        fechaNacimiento = ""
    End If
    

    '------------------------------------------------------------------------------------------
    'Estado Civil
    If arrayLinea(14) <> "" Then
        'StrSql = "SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(Mid(arrayLinea(14), 1, 30)) & "'"
        StrSql = "SELECT codexterno FROM mapeo_sap WHERE upper(infotipo) = 'IP_HIRE' AND UPPER(tablaref) = 'ESTCIVIL' AND UPPER(codinterno) = '" & UCase(Mid(arrayLinea(14), 1, 30)) & "'"
        OpenRecordset StrSql, rs_datos
        
        If Not rs_datos.EOF Then
            estCivil = Mid(rs_datos!codexterno, 1, 30)
            StrSql = "SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(estCivil) & "'"
            OpenRecordset StrSql, rs_datosAux
            If Not rs_datos.EOF Then
                estCivilNro = rs_datosAux!estcivnro
                Flog.writeline Espacios(Tabulador * 1) & "Estado Civil encontrado."
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Estado Civil no encontrado, se creara."
                estCivilNro = crearEstadoCivil(UCase(estCivil))
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Error: Estado Civil no configurado (mapeo)."
            HuboError = True
        End If
    End If
    
    '------------------------------------------------------------------------------------------
    'Pais Nacimiento
    If arrayLinea(15) <> "" Then
        Call crearPais(UCase(Mid(arrayLinea(15), 1, 60)), "", paisNacimiento, provnro)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Pais de Nacimiento no informado."
    End If
    
    '------------------------------------------------------------------------------------------
    'Fecha de Alta
    If arrayLinea(16) <> "" Then
        fechaAlta = armarfecha(arrayLinea(16))
        Flog.writeline Espacios(Tabulador * 1) & "Fecha de alta obtenida."
    Else
        FecAlta = ""
        Flog.writeline Espacios(Tabulador * 1) & "Fecha de alta no informada."
    End If

    '------------------------------------------------------------------------------------------
    'Email
    If arrayLinea(34) <> "" Then
        email = arrayLinea(34)
        Flog.writeline Espacios(Tabulador * 1) & "Email encontrado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Email no informado."
        email = ""
    End If

    '------------------------------------------------------------------------------------------
    'Remuneracion
    If arrayLinea(22) <> "" Then
        remuneracion = Replace(FormatNumber(arrayLinea(22), 2), ",", "")
        Flog.writeline Espacios(Tabulador * 1) & "Remuneracion obtenida."
    Else
        remuneracion = "0.00"
        Flog.writeline Espacios(Tabulador * 1) & "Remuneracion no informada."
    End If

    '------------------------------------------------------------------------------------------
    'Nacionalidad
    If arrayLinea(29) <> "" Then
        nacionalidad = arrayLinea(29)
        StrSql = "SELECT nacionalnro FROM nacionalidad WHERE upper(nacionaldes) = '" & UCase(Mid(nacionalidad, 1, 30)) & "'"
        OpenRecordset StrSql, rs_datos
        If Not rs_datos.EOF Then
            nacionalidad = rs_datos!nacionalNro
            Flog.writeline Espacios(Tabulador * 1) & "Nacionalidad encontrada."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Nacionalidad no encontrada, se creara."
            nacionalidad = crearNacionalidad(UCase(Mid(nacionalidad, 1, 30)))
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: Nacionalidad no informada."
        Exit Sub
    End If
    
    If ternro = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "El empleado no existe se creara."
        
        Call crearEmpleado(ternro, fechaAlta, nacionalidad, paisNacimiento, Legajo, -1, email, remuneracion, fechaNacimiento, sexo, estCivilNro, nombre, nombre2, apellido, 0)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado existe se actualizaran los datos."
        Call actualizarEmpleado(ternro, fechaAlta, nacionalidad, paisNacimiento, Legajo, -1, email, remuneracion, fechaNacimiento, sexo, estCivilNro, nombre, nombre2, apellido, 0)
    End If
    
    If reportaA <> 0 Then
        Call actualizaReportaA(ternro, reportaA)
    End If
    '------------------------------------------------------------------------------------------
    'Seccion de domicilio
        calle = Left(arrayLinea(23), 30)    'cortamos a 30 caracteres que es la long maxima del campo
        
        Numero = "-"    'Split(arrayLinea(23), ";")(1) 'por pedido del cliente informamos un "-" en este campo
        Localidad = arrayLinea(25)
        Provincia = arrayLinea(26)
        Pais = arrayLinea(29)
        
        If Trim(Pais) = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio de domicilio el pais debe ser informado."
            Exit Sub
        End If
        
        cp = arrayLinea(28)
        Call verificarLocalidad(Localidad, locnro)
        Call crearPais(Pais, Provincia, paisnro, provnro)
        Call actualizarDomicilio(ternro, paisnro, provnro, zona, partido, locnro, Barrio, entreCalles, cp, manzana, torre, depto, piso, Numero, calle, domnro)
    
    '------------------------------------------------------------------------------------------
    'Seccion de estructuras
    fechaDesde = armarfecha(arrayLinea(16))
    fechaIngActual = armarfecha(arrayLinea(17))
    
    If fechaDesde = "" Then
        Flog.writeline Espacios(Tabulador * 1) & "Error: No existe fecha desde para la estructura."
        Exit Sub
    End If
    
    'actualizo documentos
    StrSql = " SELECT confnrocol, confval, confval2 FROM confrep WHERE repnro = 480 AND confnrocol in (14,15,16)"
    OpenRecordset StrSql, rs_aux
    If Not rs_aux.EOF Then
        Do While Not rs_aux.EOF
            Select Case CLng(rs_aux!confnrocol)
                Case 14: 'DNI columna 52
                    tipoDni = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "DNI configurado Tipo: " & tipoDni
                       
                Case 15: 'CUIL columna 51
                    tipoCuil = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Cuil configurado Tipo: " & tipoCuil
            
                Case 16: 'Remuneracion a novedad columna 23
                    Conccod = rs_aux!confval2
                    tpanro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Novedad concepto: " & Conccod & " parametro: " & tpanro
            
            End Select
            
            rs_aux.MoveNext
        Loop
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No hay documentos configurados"
    End If
    
    dni = arrayLinea(51)
    Cuil = arrayLinea(50)
    
    
    
    If Trim(dni) <> "" Then
        Call actualizarDocumento(ternro, tipoDni, dni)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "DNI no informado"
    End If
    
    
    
    If Trim(Cuil) <> "" Then
        If Cuil_Valido605(Cuil, dni, errorCuil, "DNI", nacionalidad) Then
            Call actualizarDocumento(ternro, tipoCuil, Cuil)
            Flog.writeline Espacios(Tabulador * 1) & "Cuil actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & errorCuil
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "CUIL no informado"
    End If
    
    'recupero las estructuras configuradas en el reporte (480)
    StrSql = " SELECT confnrocol, confval FROM confrep WHERE repnro = 480 AND upper(conftipo) = 'TE' AND confnrocol in (1,2,3,5,6,8,9,10,11,12)"
    OpenRecordset StrSql, rs_aux
    If Not rs_aux.EOF Then
        Do While Not rs_aux.EOF
            Select Case CLng(rs_aux!confnrocol)
                Case 1: 'Unidad de Negocio columna 19
                    unidadNegocioTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Unidad de Negocio configurado TE: " & unidadNegocioTenro
                
                Case 2: 'Banda columna 20
                    bandaTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Banda configurado TE: " & bandaTenro
            
                Case 3: 'Cargo columna 21
                    cargoTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Cargo configurado TE: " & cargoTenro
            

                Case 5: 'Regimen Horario columna 44
                    regimenHorarioTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Regimen Horario configurado TE: " & regimenHorarioTenro
            

                Case 6: '(Industry Focus Group columna 150
                    industryFocusGroupTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Industry Focus Group configurado TE: " & industryFocusGroupTenro
            

                Case 8: 'Forma de Pago columna 22
                    formaPagoTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Forma de Pago configurado TE: " & formaPagoTenro
            
                Case 9: 'Location Name columna 42
                    locationNameTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Location Name configurado TE: " & locationNameTenro
                        
                Case 10: 'AFJP columna 46
                    afjpTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "AFJP configurado TE: " & afjpTenro
            
                Case 11: 'Obra Social Elegida columna 47
                    osElegidaTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Obra Social Elegida configurado TE: " & osElegidaTenro
            
                Case 12: 'SIJP columna 48
                    sijpTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "SIJP configurado TE: " & sijpTenro
            End Select
            rs_aux.MoveNext
        Loop
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No hay estructuras configuras"
    End If
    
    Call insertarNovedad(ternro, "", "", Conccod, tpanro, remuneracion, 0)
    
    unidadNegocio = arrayLinea(18)
    banda = arrayLinea(19)
    cargo = arrayLinea(20)
    formaPago = arrayLinea(21)
    locationName = arrayLinea(41)
    regimenHorario = arrayLinea(43)
    AFJP = arrayLinea(45)
    osElegida = arrayLinea(46)
    sijp = arrayLinea(47)
    industryFocusGroup = arrayLinea(149)
    
    'analizamos la fase segun la fecha de alta
    Call actualizarFase(ternro, fechaAlta, fechaIngActual, "NULL")
    
    If unidadNegocio <> "" Then
        Call ValidaEstructura(unidadNegocioTenro, unidadNegocio, unidadNegocioEstrnro, Inserto_estr)
        Call VerSiCrearTercero(unidadNegocioTenro, unidadNegocio, unidadNegocioTernro)
        Call AsignarEstructura(unidadNegocioTenro, unidadNegocioEstrnro, ternro, fechaDesde, "Null")
        If Inserto_estr Then
            Call VerSiCrearComplemento(unidadNegocioTenro, unidadNegocioEstrnro, unidadNegocio, unidadNegocioTernro)
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio (TE configurable columna 8) actualizado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio (TE configurable columna 8) no informado."
    End If
    
    If banda <> "" Then
        Call ValidaEstructura(bandaTenro, banda, bandaEstrnro, Inserto_estr)
        Call VerSiCrearTercero(bandaTenro, banda, bandaTernro)
        Call AsignarEstructura(bandaTenro, bandaEstrnro, ternro, fechaDesde, "Null")
        If Inserto_estr Then
            Call VerSiCrearComplemento(bandaTenro, bandaEstrnro, banda, bandaTernro)
        End If
        Flog.writeline Espacios(Tabulador * 1) & "banda (TE configurable columna 9) actualizada."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "banda (TE configurable columna 9) no informada."
    End If
    
    If cargo <> "" Then
        Call ValidaEstructura(cargoTenro, cargo, cargoEstrnro, Inserto_estr)
        Call VerSiCrearTercero(cargoTenro, cargo, cargoTernro)
        Call AsignarEstructura(cargoTenro, cargoEstrnro, ternro, fechaDesde, "Null")
        If Inserto_estr Then
            Call VerSiCrearComplemento(cargoTenro, cargoEstrnro, cargo, cargoTernro)
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Cargo (TE configurable columna 10) actualizado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Cargo (TE configurable columna 10) no informado."
    End If
    
    If formaPago <> "" Then
        Call ValidaEstructura(formaPagoTenro, formaPago, formaPagoEstrnro, Inserto_estr)
        Call VerSiCrearTercero(formaPagoTenro, formaPago, formaPagoTernro)
        Call AsignarEstructura(formaPagoTenro, formaPagoEstrnro, ternro, fechaDesde, "Null")
        If Inserto_estr Then
            Call VerSiCrearComplemento(formaPagoTenro, formaPagoEstrnro, formaPago, formaPagoTernro)
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Forma de Pago (TE configurable columna 11) actualizado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Forma de Pago (TE configurable columna 11) no informado."
    End If
    
    If locationName <> "" Then
        Call ValidaEstructura(locationNameTenro, locationName, locationNameEstrnro, Inserto_estr)
        Call VerSiCrearTercero(locationNameTenro, locationName, locationNameTernro)
        Call AsignarEstructura(locationNameTenro, locationNameEstrnro, ternro, fechaDesde, "Null")
        If Inserto_estr Then
            Call VerSiCrearComplemento(locationNameTenro, locationNameEstrnro, locationName, locationNameTernro)
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Location Name (TE configurable columna 12) actualizado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Location Name (TE configurable columna 12) no informado."
    End If
    
    If regimenHorario <> "" Then
        Call ValidaEstructuraCodExt(regimenHorarioTenro, regimenHorario, regimenHorarioEstrnro, Inserto_estr)
        If regHorarioEstrnro <> 0 Then
            Call VerSiCrearTercero(regimenHorarioTenro, regimenHorario, regimenHorarioTernro)
            Call AsignarEstructura(regimenHorarioTenro, regimenHorarioEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(regimenHorarioTenro, regimenHorarioEstrnro, regimenHorario, regimenHorarioTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Regimen Horario (TE configurable columna 13) actualizado."
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Regimen Horario (TE configurable columna 13) no informado."
    End If
    
    If AFJP <> "" Then
        Call ValidaEstructuraCodExt(afjpTenro, AFJP, AFJPEstrnro, Inserto_estr)
        If AFJPEstrnro <> 0 Then
            Call VerSiCrearTercero(afjpTenro, AFJP, AFJPTernro)
            Call AsignarEstructura(afjpTenro, AFJPEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(afjpTenro, AFJPEstrnro, AFJP, AFJPTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "AFJP (TE configurable columna 14) actualizado."
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "AFJP (TE configurable columna 14) no informado."
    End If
    
    If osElegida <> "" Then
        Call ValidaEstructuraCodExt(osElegidaTenro, osElegida, osElegidaEstrnro, Inserto_estr)
        If osElegidaEstrnro <> 0 Then
            Call VerSiCrearTercero(osElegidaTenro, osElegida, osElegidaTernro)
            Call AsignarEstructura(osElegidaTenro, osElegidaEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(osElegidaTenro, osElegidaEstrnro, osElegida, osElegidaTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Obra social (TE configurable columna 15) actualizado."
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Obra Social (TE configurable columna 15) no informado."
    End If
    
    If sijp <> "" Then
        Call ValidaEstructuraCodExt(sijpTenro, sijp, sijpEstrnro, Inserto_estr)
        If sijpEstrnro <> 0 Then
            Call VerSiCrearTercero(sijpTenro, sijp, sijpTernro)
            Call AsignarEstructura(sijpTenro, sijpEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(sijpTenro, sijpEstrnro, sijp, sijpTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "SIJP (TE configurable columna 16) actualizado."
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "SIJP (TE configurable columna 16) no informado."
    End If
    
    If industryFocusGroup <> "" Then
        Call ValidaEstructura(industryFocusGroupTenro, industryFocusGroup, industryFocusGroupEstrnro, Inserto_estr)
        Call VerSiCrearTercero(industryFocusGroupTenro, industryFocusGroup, industryFocusGroupTernro)
        Call AsignarEstructura(industryFocusGroupTenro, industryFocusGroupEstrnro, ternro, fechaDesde, "Null")
        If Inserto_estr Then
            Call VerSiCrearComplemento(industryFocusGroupTenro, industryFocusGroupEstrnro, industryFocusGroup, industryFocusGroupTernro)
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Industry Focus Group (TE configurable columna 17) actualizado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Industry Focus Group (TE configurable columna 17) no informado."
    End If
    
    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento HIRE"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"
End Sub

Sub movASGC(ByVal arrayLinea)
Dim rs_aux As New ADODB.Recordset

Dim ternro As Long
Dim fechaDesde As String
Dim unidadNegocio As String
Dim unidadNegocioEstrnro As Long
Dim unidadNegocioTernro As Long
Dim unidadNegocioTenro As String

Dim banda As String
Dim bandaEstrnro As Long
Dim bandaTernro As Long
Dim bandaTenro As String

Dim regHorario As String
Dim regHorarioEstrnro As Long
Dim regHorarioTernro As Long
Dim regHorarioTenro As String

Dim cargo As String
Dim cargoEstrnro As Long
Dim cargoTernro As Long
Dim cargoTenro As String

Dim lugarPago As String
Dim lugarPagoEstrnro As Long
Dim lugarPagoTernro As Long
Dim lugarPagoTenro As String

Dim industryFocus As String
Dim industryFocusEstrnro As Long
Dim industryFocusTernro As Long
Dim industryFocusTenro As String

Dim Inserto_estr As Boolean
Dim reportaA As String

    Flog.writeline Espacios(Tabulador * 1) & "Comienzo Movimiento ASGC"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"

    'recupero las estructuras configuradas en el reporte (480)
    StrSql = " SELECT confnrocol, confval FROM confrep WHERE repnro = 480 AND upper(conftipo) = 'TE' AND confnrocol in (1,2,3,4,5,6)"
    OpenRecordset StrSql, rs_aux
    If Not rs_aux.EOF Then
        Do While Not rs_aux.EOF
            Select Case CLng(rs_aux!confnrocol)
                Case 1: 'Unidad de negocio columna 19
                    unidadNegocioTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio TE: " & unidadNegocioTenro
                Case 2: 'Banda columna 20
                    bandaTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Banda TE: " & bandaTenro
                Case 3: 'Cargo columna 21
                    cargoTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Cargo TE: " & cargoTenro
                Case 4: 'Lugar de Pago columna 42
                    lugarPagoTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Lugar de Pago TE: " & lugarPagoTenro
                Case 5: 'Regimen Horario columna 44
                    regHorarioTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Regimen Horario TE: " & regHorarioTenro
                Case 6: 'Industry Focus Group columna 46
                    industryFocusTenro = rs_aux!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Industry Focus Group TE: " & industryFocusTenro
            End Select
            rs_aux.MoveNext
        Loop
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No hay estructuras configuras"
    End If
    
    
    ternro = existeEmpleado(arrayLinea(4))
    Flog.writeline Espacios(Tabulador * 1) & "Controlo si existe reporta A."
    reportaA = existeEmpleado(arrayLinea(42))
    unidadNegocio = arrayLinea(18)
    banda = arrayLinea(19)
    cargo = arrayLinea(20)
    lugarPago = arrayLinea(41)
    regHorario = arrayLinea(43)
    industryFocus = arrayLinea(45)
    fechaDesde = armarfecha(arrayLinea(1))
    If ternro <> 0 Then
        
        If reportaA <> 0 Then
            Call actualizaReportaA(ternro, reportaA)
        End If
        
        If unidadNegocio <> "" Then
            Call ValidaEstructura(unidadNegocioTenro, unidadNegocio, unidadNegocioEstrnro, Inserto_estr)
            Call VerSiCrearTercero(unidadNegocioTenro, unidadNegocio, unidadNegocioTernro)
            Call AsignarEstructura(unidadNegocioTenro, unidadNegocioEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(unidadNegocioTenro, unidadNegocioEstrnro, unidadNegocio, unidadNegocioTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio no informado."
        End If
            
        If banda <> "" Then
            Call ValidaEstructura(bandaTenro, banda, bandaEstrnro, Inserto_estr)
            Call VerSiCrearTercero(bandaTenro, banda, bandaTernro)
            Call AsignarEstructura(bandaTenro, bandaEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(bandaTenro, bandaEstrnro, banda, bandaTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Banda actualizada."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "banda no informada."
        End If
            
        If cargo <> "" Then
            Call ValidaEstructura(cargoTenro, cargo, cargoEstrnro, Inserto_estr)
            Call VerSiCrearTercero(cargoTenro, cargo, cargoTernro)
            Call AsignarEstructura(cargoTenro, cargoEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(cargoTenro, cargoEstrnro, cargo, cargoTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) no informado."
        End If
    
        'lugar de pago
        If lugarPago <> "" Then
            Call ValidaEstructura(lugarPagoTenro, lugarPago, lugarPagoEstrnro, Inserto_estr)
            Call VerSiCrearTercero(lugarPagoTenro, lugarPago, lugarPagoTernro)
            Call AsignarEstructura(lugarPagoTenro, lugarPagoEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(lugarPagoTenro, lugarPagoEstrnro, lugarPago, lugarPagoTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Lugar de Pago actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Lugar de Pago no informado."
        End If
        
        If regHorario <> "" Then
            Call ValidaEstructuraCodExt(regHorarioTenro, regHorario, regHorarioEstrnro, Inserto_estr)
            If regHorarioEstrnro <> 0 Then
                Call VerSiCrearTercero(regHorarioTenro, regHorario, regHorarioTernro)
                Call AsignarEstructura(regHorarioTenro, regHorarioEstrnro, ternro, fechaDesde, "Null")
                If Inserto_estr Then
                    Call VerSiCrearComplemento(regHorarioTenro, regHorarioEstrnro, regHorario, regHorarioTernro)
                End If
                Flog.writeline Espacios(Tabulador * 1) & "Regimen Horario actualizado."
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Regimen Horario no informado."
        End If
        
        'Industry Focus
        If industryFocus <> "" Then
            Call ValidaEstructura(industryFocusTenro, industryFocus, industryFocusEstrnro, Inserto_estr)
            Call VerSiCrearTercero(industryFocusTenro, industryFocus, industryFocusTernro)
            Call AsignarEstructura(industryFocusTenro, industryFocusEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(industryFocusTenro, industryFocusEstrnro, industryFocus, industryFocusTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Industry Focus actualizado."
            
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Industry Focus no informado."
        End If
        
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para cambios en el movimiento ASGC el empleado debe existir."
        Exit Sub
    End If

    Flog.writeline Espacios(Tabulador * 1) & "Fin Movimiento ASGC"
    Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------------------------------------------------"

End Sub
