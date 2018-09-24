Attribute VB_Name = "mdlExportacionInterpay"
Option Explicit

Function expEmpleado(ByVal ternro As Long, ByVal separador As String)
 Dim rsDatosEmp  As New ADODB.Recordset
 Dim rsAux  As New ADODB.Recordset
 Dim rsConfrep  As New ADODB.Recordset
 Dim strLinea As String
 
 Dim unidadNegocioTenro As String
 Dim bandaTenro As String
 Dim puestoTenro As String
 Dim afjpTenro As String
 Dim convenioTenro As String
 Dim obraSocialTenro As String
 Dim sijpTenro As String
 Dim ifgTenro As String
 Dim Conccod As String
 Dim tpanro As String
 Dim cuilTidnro As String
 Dim dniTidnro As String
 Dim estrnroFueraConv As String
 Dim listaConccod As String

    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Generando la exportacion para el ternro: " & ternro & "."

    'busco los tipos de estructura configurados
    StrSql = " SELECT confnrocol, confval, confval2 FROM confrep WHERE repnro = 480 AND confnrocol in (1,2,3,6,10,11,12,13,14,15,18,19,20)"
    OpenRecordset StrSql, rsConfrep
    If Not rsConfrep.EOF Then
        
        listaConccod = "'0'"
        Conccod = "0"
        tpanro = "0"
        Do While Not rsConfrep.EOF
            Select Case CLng(rsConfrep!confnrocol)
                Case 1: 'Unidad de Negocio columna 14
                    unidadNegocioTenro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "unidad de Negocio configurado Tipo: " & unidadNegocioTenro
                
                Case 2: 'Banda columna 15
                    bandaTenro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Banda configurado Tipo: " & bandaTenro
                
                Case 3: 'Cargo (puesto) columna  16
                    puestoTenro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Cargo (Puesto) configurado Tipo: " & puestoTenro
                
                Case 6: 'Industry Focus Group columna 38
                    ifgTenro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Industry Focus Group configurado Tipo: " & ifgTenro
                
                Case 10: 'AFJP columna 33
                    afjpTenro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "AFJP configurado Tipo: " & afjpTenro
            
                Case 11: 'Obra social columna 35
                    obraSocialTenro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Obra social configurado Tipo: " & obraSocialTenro
                    
                 Case 12: 'SIJP columna 34
                    sijpTenro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "SIJP configurado Tipo: " & sijpTenro
                    
                Case 13: 'Convenio columna 37
                    convenioTenro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Convenio configurado Tipo: " & convenioTenro
            
                Case 14: 'DNI columna 32
                    dniTidnro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "DNI configurado Tipo: " & dniTidnro
            
                Case 15: 'CUIL columna 31
                    cuilTidnro = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "CUIL configurado Tipo: " & cuilTidnro
            
                Case 18: 'Codigo de concepto y parametro de novedad columna 18
                    Conccod = Conccod & "," & rsConfrep!confval2
                    tpanro = tpanro & "," & rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Columna de concepto y parametro para novedad encontrada."
            
                Case 19: 'Codigo de estructura de Fuera de convenio
                    estrnroFueraConv = rsConfrep!confval
                    Flog.writeline Espacios(Tabulador * 1) & "Columna de estructura fuera de convenio encontrada."
            
                Case 20: 'Codigo de estructura de Fuera de convenio
                    listaConccod = listaConccod & ",'" & rsConfrep!confval2 & "'"
            
            End Select
            
            rsConfrep.MoveNext
        Loop
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No hay documentos configurados"
    End If

    On Error GoTo CE
        
    StrSql = "SELECT empleg,tercero.ternom,tercero.ternom2,tercero.terape,tercero.terape2,tercero.terfecnac,pais.paiscodext " & _
            ",tercero.terfecing,estcivil.estcivdesabr,tercero.tersex,empleado.empfecalta,empleado.empemail, dni.nrodoc dni, cuil.nrodoc cuil" & _
            ",empleado.empfbajaprev,empleado.empest,empleado.empremu,empleado.empreporta " & _
            " FROM empleado " & _
            " INNER JOIN tercero ON tercero.ternro = empleado.ternro  " & _
            " LEFT JOIN pais ON pais.paisnro = tercero.paisnro " & _
            " INNER JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro " & _
            " LEFT JOIN ter_doc dni ON dni.ternro = empleado.ternro AND dni.tidnro = " & dniTidnro & _
            " LEFT JOIN ter_doc cuil ON cuil.ternro = empleado.ternro AND cuil.tidnro = " & cuilTidnro & _
            " WHERE empleado.ternro = " & ternro
    OpenRecordset StrSql, rsDatosEmp
    If Not rsDatosEmp.EOF Then
        'legajo y empresa
        strLinea = rsDatosEmp!empleg                                'pos 1 legajo
        strLinea = strLinea & separador                             'pos 2 vacio
        strLinea = strLinea & separador                             'pos 3 vacio
        'primer y segundo nombre
        strLinea = strLinea & separador & rsDatosEmp!ternom         'pos 4 nombre
        If Not IsNull(rsDatosEmp!ternom2) Then
            strLinea = strLinea & " " & rsDatosEmp!ternom2
        End If
        
        strLinea = strLinea & separador                             'pos 5 vacio
        strLinea = strLinea & separador                             'pos 6 vacio
        strLinea = strLinea & separador                             'pos 7 vacio
        
        'sexo
        If (CLng(rsDatosEmp!tersex) = -1) Then
            strLinea = strLinea & separador & "M"                   'pos 8 sexo
        Else
            strLinea = strLinea & separador & "F"
        End If
        'fecha de nacimiento
        strLinea = strLinea & separador & formatoFecha(rsDatosEmp!terfecnac, "ddmmyyyy")     'pos 9 fecha de nacimiento
        'estado civil
        strLinea = strLinea & separador & mapeoSap("IP_HIRE", rsDatosEmp!estcivdesabr)  'pos 10 estado civil
        'pais de nacimiento
        strLinea = strLinea & separador & rsDatosEmp!paiscodext       'pos 11 pais de nacimiento
        
        'fecha de fase mas antigua
        strLinea = strLinea & separador & formatoFecha(calcularFase(ternro, separador, -1), "ddmmyyyy") 'pos 12 fecha fecha de alta actual

        'fecha de ingreso
        strLinea = strLinea & separador & formatoFecha(calcularFase(ternro, separador, 0), "ddmmyyyy")     'pos 13 de ingreso reconocida
        
        'Estructura Unidad de Negocio
        strLinea = strLinea & separador & obtenerEstructura(ternro, unidadNegocioTenro, Date, "estrdabr")   'pos 14 sub business
        strLinea = strLinea & separador & obtenerEstructura(ternro, bandaTenro, Date, "estrdabr")   'pos 15 banda
        strLinea = strLinea & separador & obtenerEstructura(ternro, puestoTenro, Date, "estrdabr")    'pos 16 cargo (puesto)
        
        strLinea = strLinea & separador                                         'pos 17 vacio
        
        
        If poseeEstructura(ternro, 19, Date, estrnroFueraConv) Then
            'busco la novedad - el empleado no posee la es estructura configura Fuera de convenio
            strLinea = strLinea & separador & Format(obtenerNovedad(ternro, Conccod, tpanro, Date, 0), ".00")   'pos 18 remuneracion
        Else
            'busco la liquidacion en el mes anterior
            strLinea = strLinea & separador & Format(obtenerSueldo(ternro, listaConccod, Date), ".00")   'pos 18 remuneracion
        End If
        
        strLinea = strLinea & separador & armarDireccion(ternro, separador)     'pos 19-28 direccion y telefono

        strLinea = strLinea & separador & rsDatosEmp!empemail                   'pos 29 email
    
        strLinea = strLinea & separador                                         'pos 30 vacio
        'Cuil
        strLinea = strLinea & separador & IIf(EsNulo(rsDatosEmp!Cuil), "", rsDatosEmp!Cuil) 'pos 31 cuil
        
        'DNI
        strLinea = strLinea & separador & IIf(EsNulo(rsDatosEmp!dni), "", rsDatosEmp!dni) 'pos 32 dni
        
        strLinea = strLinea & separador & obtenerEstructura(ternro, afjpTenro, Date, "estrcodext")    'pos 33 AFJP
        strLinea = strLinea & separador & obtenerEstructura(ternro, convenioTenro, Date, "estrcodext")    'pos 34 Convenio
        strLinea = strLinea & separador & obtenerEstructura(ternro, obraSocialTenro, Date, "estrcodext")    'pos 35 Obra Social
        strLinea = strLinea & separador                                                                 'pos 36 vacio
        strLinea = strLinea & separador & obtenerEstructura(ternro, sijpTenro, Date, "estrcodext")    'pos 37 SIJP
        strLinea = strLinea & separador & obtenerEstructura(ternro, ifgTenro, Date, "estrdabr")    'pos 38 Industry Focus Group

    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado no tiene cargado algunno de los siguientes datos: "
        Flog.writeline Espacios(Tabulador * 2) & "Empresa, nacionalidad, pais de nacimiento o estado civil."
    End If
        
    Flog.writeline Espacios(Tabulador * 1) & "Fin de generacion de exportacion para el ternro: " & ternro & "."
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------------------"
    'buscamos el reporta A
'    StrSql = "SELECT empleg FROM empleado WHERE empleado.ternro= " & IIf(EsNulo(rsDatosEmp!empreporta), 0, rsDatosEmp!empreporta)
'    OpenRecordset StrSql, rsAux
'    If rsAux.EOF Then
'        Flog.writeline Espacios(Tabulador * 1) & "El empleado no posee reporta A."
'        strLinea = strLinea & separador & "N/A"
'    Else
'        strLinea = strLinea & separador & rsAux!empleg
'    End If
    
    
    GoTo datosOk
CE:
    strLinea = ""
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 389. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Function
datosOk:
    
    expEmpleado = strLinea
    Set rsDatosEmp = Nothing
End Function

