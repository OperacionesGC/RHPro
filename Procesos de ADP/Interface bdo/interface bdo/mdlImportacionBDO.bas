Attribute VB_Name = "mdlImportacionBDO"
Public Sub import(ByVal strLinea As String, ByVal nroLinea As Long, ByVal separador As String)
Dim arrayLinea

Dim codMovimiento As String

    Flog.writeline Espacios(Tabulador * 0) & "Comienzo de la importacion para la linea: " & nroLinea & "."
    arrayLinea = Split(strLinea, separador)
    
    '-----------------------------------------------------------------------------------------------------------------------------
    codMovimiento = arrayLinea(0)
    Select Case UCase(codMovimiento)
        Case "ADR":
            Call movADR(arrayLinea)
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
            'Call movREVH(arrayLinea)
        Case "REVT":
            Call movREVT(arrayLinea)
        Case "SLRY":
            Call movSLRY(arrayLinea)
        Case "TERM":
            Call movTERM(arrayLinea)
        'LSA - actualmente no se tienen estos bonos, pero se van a tener en un futuro
        'Case "LSA":
        '    Call movLSA(arrayLinea)
    
    End Select

    Flog.writeline Espacios(Tabulador * 0) & "Fin importacion, Nro de linea: " & nroLinea & "."
    Flog.writeline Espacios(Tabulador * 0) & "-----------------------------------------------------------."
    
    
End Sub



Sub movADR(ByVal arrayLinea)
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

    ternro = existeEmpleado(arrayLinea(4))
    If ternro <> 0 Then
        calle = Split(arrayLinea(23), ";")(0)
        Numero = Split(arrayLinea(23), ";")(1)
        piso = Split(arrayLinea(23), ";")(2)
        depto = Split(arrayLinea(23), ";")(3)
        torre = Split(arrayLinea(23), ";")(4)

        manzana = Split(arrayLinea(24), ";")(0)
        entreCalles = Split(arrayLinea(24), ";")(1)
        
        Barrio = Split(arrayLinea(25), ";")(0)
        Localidad = Split(arrayLinea(25), ";")(1)
        
        Provincia = Split(arrayLinea(26), ";")(0)
        Pais = Split(arrayLinea(26), ";")(1)
        
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
End Sub

Sub movREVT(ByVal arrayLinea)
Dim ternro As Long
Dim fechaBajaPrevista As String

    ternro = existeEmpleado(arrayLinea(4))
    
    If ternro <> 0 Then
        fechaBajaPrevista = armarfecha(arrayLinea(2))
        
        Call bajaContrato(ternro, fechaBajaPrevista)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio Fecha de baja prevista el empleado debe existir."
        Exit Sub
    End If
    
End Sub

Sub movIDNO(ByVal arrayLinea)
'Cambios de docuemntos/identificaciones
Dim ternro As Long
Dim nroDocumento As String
Dim nroCuil As String

    ternro = existeEmpleado(arrayLinea(4))
    If ternro <> 0 Then
    
        nroCuil = arrayLinea(50)
        If nroCuil <> "" Then
            Call actualizarDocumento(ternro, "10", nroCuil)
            Flog.writeline Espacios(Tabulador * 1) & "Cuil actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Cuil no informado."
        End If
        
        nroDocumento = arrayLinea(51)
        If nroDocumento <> "" Then
            Call actualizarDocumento(ternro, "1", nroDocumento)
            Flog.writeline Espacios(Tabulador * 1) & "DNI actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "DNI no informado."
        End If
        
        
        Flog.writeline Espacios(Tabulador * 1) & "documentos o identificacion Actualizados."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio de documentos o identificacion el empleado debe existir."
        Exit Sub
    End If
    
End Sub

Sub movPERS(ByVal arrayLinea)
'Cambios de información personal
Dim rs_datos As New ADODB.Recordset
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
        If arrayLinea(10) <> "" Then
            nombre = Split(arrayLinea(10), " ")(0)
            Flog.writeline Espacios(Tabulador * 1) & "Primer Nombre obtenido."
            If UBound(Split(arrayLinea(10), " ")) > 0 Then
                nombre2 = Split(arrayLinea(10), " ")(1)
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
            StrSql = "SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(Mid(arrayLinea(14), 1, 30)) & "'"
            OpenRecordset StrSql, rs_datos
            estCivil = Mid(arrayLinea(14), 1, 30)
            If Not rs_datos.EOF Then
                estCivilNro = rs_datos!estcivnro
                Flog.writeline Espacios(Tabulador * 1) & "Estado Civil encontrado."
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Estado Civil no encontrado, se creara."
                estCivilNro = crearEstadoCivil(UCase(Mid(arrayLinea(14), 1, 30)))
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Error: Estado Civil Obligatorio."
            HuboError = True
        End If

        '------------------------------------------------------------------------------------------
        'Pais Nacimiento
        If arrayLinea(15) <> "" Then
                Call crearPais(UCase(Mid(arrayLinea(15), 1, 60)), "", paisNacimiento, provnro)
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
        Flog.writeline Espacios(Tabulador * 1) & "Datos generales del empleado Actualizados."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para actualizar un empleado debe existir."
        Exit Sub
    End If


End Sub

Sub movSLRY(ByVal arrayLinea)
'Cambios de Salarios
Dim ternro As Long
Dim salario As String

    ternro = existeEmpleado(arrayLinea(4))
    If ternro <> 0 Then
        salario = arrayLinea(23)
        Call actualizarSalario(ternro, salario)
        Flog.writeline Espacios(Tabulador * 1) & "Salario del empleado actualizado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio de salario el empleado debe existir."
        Exit Sub
    End If

End Sub

Sub movCOST(ByVal arrayLinea)
'Cambios de centro de costo, unidad de negocio y puesto
Dim ternro As Long
Dim fechaDesde As String
Dim unidadNegocio As String
Dim centroCosto As String
Dim cargo As String     'puesto
Dim unidadNegocioEstrnro As Long
Dim centroCostoEstrnro As Long
Dim cargoEstrnro As Long     'puesto
Dim Inserto_estr As Boolean
Dim unidadNegocioTernro As Long
Dim centroCostoTernro As Long
Dim cargoTernro As Long

    ternro = existeEmpleado(arrayLinea(4))
    If ternro <> 0 Then
        fechaDesde = armarfecha(arrayLinea(1))
        unidadNegocio = arrayLinea(18)
        cargo = arrayLinea(20)
        centroCosto = arrayLinea(21)
        
        If fechaDesde = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: No existe fecha desde para la estructura."
            Exit Sub
        End If
        
        If unidadNegocio <> "" Then
            Call ValidaEstructura(1, unidadNegocio, unidadNegocioEstrnro, Inserto_estr)
            Call VerSiCrearTercero(1, unidadNegocio, unidadNegocioTernro)
            Call AsignarEstructura(1, unidadNegocioEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(1, unidadNegocioEstrnro, unidadNegocio, unidadNegocioTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio no informado."
        End If
                
        If cargo <> "" Then
            Call ValidaEstructura(4, cargo, cargoEstrnro, Inserto_estr)
            Call VerSiCrearTercero(4, cargo, cargoTernro)
            Call AsignarEstructura(4, cargoEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(4, cargoEstrnro, cargo, cargoTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) no informado."
        End If
        
        If centroCosto <> "" Then
            Call ValidaEstructura(5, centroCosto, centroCostoEstrnro, Inserto_estr)
            Call VerSiCrearTercero(5, centroCosto, centroCostoTernro)
            Call AsignarEstructura(5, centroCostoEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(5, centroCostoEstrnro, centroCosto, centroCostoTernro)
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

End Sub

Sub movTERM(ByVal arrayLinea)
Dim ternro As Long
Dim fechaBajaPrevista As String

    ternro = existeEmpleado(arrayLinea(4))
    
    If ternro <> 0 Then
        fechaBajaPrevista = armarfecha(arrayLinea(2))
        
        Call bajaContrato(ternro, fechaBajaPrevista)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para un cambio Fecha de baja prevista el empleado debe existir."
        Exit Sub
    End If

End Sub

Sub movHIRE(ByVal arrayLinea)
Dim rs_datos As New ADODB.Recordset
Dim ternro As Long
Dim Legajo As String
Dim apellido As String
Dim nombre As String
Dim sexo As String
Dim fechaNacimiento As String
Dim estCivil As String
Dim estCivilNro As String
Dim paisNacimiento As String
Dim fechaAlta As String
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
Dim Cpostal As String
Dim nacionalidad As String
Dim telParticular As String
Dim telCelular As String
Dim email As String
Dim nroCuil As String
Dim nroDocumento As String
Dim fechaDesde As String
Dim unidadNegocio As String
Dim unidadNegocioEstrnro As Long
Dim unidadNegocioTernro As Long
Dim cargo As String
Dim cargoEstrnro As Long
Dim cargoTernro As Long
Dim Inserto_estr As Boolean

    Legajo = arrayLinea(4)
    ternro = existeEmpleado(Legajo)
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
    If arrayLinea(10) <> "" Then
        nombre = Split(arrayLinea(10), " ")(0)
        Flog.writeline Espacios(Tabulador * 1) & "Primer Nombre obtenido."
        If UBound(Split(arrayLinea(10), " ")) > 0 Then
            nombre2 = Split(arrayLinea(10), " ")(1)
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
    If arrayLinea(16) <> "" Then
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
        StrSql = "SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(Mid(arrayLinea(14), 1, 30)) & "'"
        OpenRecordset StrSql, rs_datos
        estCivil = Mid(arrayLinea(14), 1, 30)
        If Not rs_datos.EOF Then
            estCivilNro = rs_datos!estcivnro
            Flog.writeline Espacios(Tabulador * 1) & "Estado Civil encontrado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Estado Civil no encontrado, se creara."
            estCivilNro = crearEstadoCivil(UCase(Mid(arrayLinea(14), 1, 30)))
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: Estado Civil Obligatorio."
        HuboError = True
    End If

    '------------------------------------------------------------------------------------------
    'Pais Nacimiento
    If arrayLinea(15) <> "" Then
        Call crearPais(UCase(Mid(arrayLinea(15), 1, 60)), "", paisNacimiento, provnro)
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
    
    '------------------------------------------------------------------------------------------
    'Pais Nacimiento
    If arrayLinea(15) <> "" Then
        Call crearPais(UCase(Mid(arrayLinea(15), 1, 60)), "", paisNacimiento, provnro)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Pais de Nacimiento no informado."
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
            nacionalidad = rs_datos!nacionalnro
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

    '------------------------------------------------------------------------------------------
    'Unidad de negocio
    fechaDesde = armarfecha(arrayLinea(16))
    
    If fechaDesde = "" Then
        Flog.writeline Espacios(Tabulador * 1) & "Error: No existe fecha desde para la estructura."
        Exit Sub
    End If
    
    unidadNegocio = arrayLinea(18)
    cargo = arrayLinea(20)
    
    'analizamos la fase segun la fecha de alta
    Call actualizarFase(ternro, fechaAlta, fechaBaja)
    
    If unidadNegocio <> "" Then
        Call ValidaEstructura(1, unidadNegocio, unidadNegocioEstrnro, Inserto_estr)
        Call VerSiCrearTercero(1, unidadNegocio, unidadNegocioTernro)
        Call AsignarEstructura(1, unidadNegocioEstrnro, ternro, fechaDesde, "Null")
        If Inserto_estr Then
            Call VerSiCrearComplemento(1, unidadNegocioEstrnro, unidadNegocio, unidadNegocioTernro)
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio actualizado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio no informado."
    End If
    
    
    If cargo <> "" Then
        Call ValidaEstructura(4, cargo, cargoEstrnro, Inserto_estr)
        Call VerSiCrearTercero(4, cargo, cargoTernro)
        Call AsignarEstructura(4, cargoEstrnro, ternro, fechaDesde, "Null")
        If Inserto_estr Then
            Call VerSiCrearComplemento(4, cargoEstrnro, cargo, cargoTernro)
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) actualizado."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) no informado."
    End If
    
End Sub

Sub movASGC(ByVal arrayLinea)
Dim ternro As Long
Dim fechaDesde As String
Dim unidadNegocio As String
Dim unidadNegocioEstrnro As Long
Dim unidadNegocioTernro As Long

Dim regHorario As String
Dim regHorarioEstrnro As Long
Dim regHorarioTernro As Long
Dim cargo As String
Dim cargoEstrnro As Long
Dim cargoTernro As Long
Dim Inserto_estr As Boolean
Dim reportaA As String

    ternro = existeEmpleado(arrayLinea(4))
    Flog.writeline Espacios(Tabulador * 1) & "Controlo si existe reporta A."
    reportaA = existeEmpleado(arrayLinea(42))
    cargo = arrayLinea(20)
    unidadNegocio = arrayLinea(18)
    regHorario = arrayLinea(43)
    
    If ternro <> 0 Then
        fechaDesde = armarfecha(arrayLinea(1))
        
        If unidadNegocio <> "" Then
            Call ValidaEstructura(1, unidadNegocio, unidadNegocioEstrnro, Inserto_estr)
            Call VerSiCrearTercero(1, unidadNegocio, unidadNegocioTernro)
            Call AsignarEstructura(1, unidadNegocioEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(1, unidadNegocioEstrnro, unidadNegocio, unidadNegocioTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Unidad de negocio no informado."
        End If
            
        If cargo <> "" Then
            Call ValidaEstructura(4, cargo, cargoEstrnro, Inserto_estr)
            Call VerSiCrearTercero(4, cargo, cargoTernro)
            Call AsignarEstructura(4, cargoEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(4, cargoEstrnro, cargo, cargoTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Cargo (puesto) no informado."
        End If
    
        If regHorario <> "" Then
            Call ValidaEstructura(21, regHorario, regHorarioEstrnro, Inserto_estr)
            Call VerSiCrearTercero(21, regHorario, regHorarioTernro)
            Call AsignarEstructura(21, regHorarioEstrnro, ternro, fechaDesde, "Null")
            If Inserto_estr Then
                Call VerSiCrearComplemento(21, regHorarioEstrnro, regHorario, regHorarioTernro)
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Regimen Horario actualizado."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Regimen Horario no informado."
        End If
        
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error: para cambios en el movimiento ASGC el empleado debe existir."
        Exit Sub
    End If

End Sub
