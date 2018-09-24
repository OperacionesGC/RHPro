Attribute VB_Name = "MdlEmpleados"
Option Explicit

Dim Ternro As String
Dim l_sql As String
Dim NroDom As Integer
Dim empleg As String
Dim terape As String
Dim terape2 As String
Dim ternom As String
Dim ternom2 As String
Dim terfecnac As String
Dim tersex As String
Dim teremail As String
Dim nacionalnro As String
Dim nacionalidad As String
Dim estcivnro As String
Dim empfecalta As String
Dim empremu As String
Dim calle As String
Dim callenro As String
Dim piso As String
Dim depto As String
Dim locnro As String
Dim provnro As String
Dim codigopostal As String
Dim telnro As String
Dim paisnro As String
Dim nrodoc As String
Dim nrocuil As String
Dim empresaEstrnro As String
Dim tipmotnro As String
Dim caunro As String
Dim Estrnro As String
Dim Mensaje As String
Dim Empnro As String
'Dim estrnro As String

Dim l_ACODE As String
Dim l_EMPLEADO As String
Dim l_FECINGRESO As String
Dim l_NOMBRE As String
Dim l_calle As String
Dim l_NUMERO As String
Dim l_piso As String
Dim l_depto As String
Dim l_LOCALIDAD As String
Dim l_PROVINICIA As String
Dim l_CODPOSTAL As String
Dim l_NRORNIC As String
Dim l_NACIONALID As String
Dim l_SEXO As String
Dim l_FECNACTO As String
Dim l_ESTCIVIL As String
Dim l_NROCUIL As String
Dim l_APODERADO As String
Dim l_CONFIDENC As String
Dim l_CONDSINDIC As String
Dim l_EMAILP As String
Dim l_TELEFONO As String
Dim l_CODOSOCIAL As String
Dim l_REBCONTPAT As String
Dim l_CONDICION As String
Dim l_CODJUBILAC As String
Dim l_ACTIVIDAD As String
Dim l_RCODE As String
Dim l_ADATE As String
Dim l_SDOJORNAL As String
Dim l_IMPUTACION As String


Public Sub Generar_Empleado(node As IXMLDOMNode)
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Procedimiento
    ' Autor      : Lisandro Moro
    ' Fecha      : 13/07/2011
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    
    Dim cadena As Variant
    Dim subCadena As Variant
    Dim error As Boolean
    error = False
    Dim b As Long
    Dim child As IXMLDOMNodeList
    
    On Error GoTo ErrorEmpleado
    
    l_ACODE = ""
    l_EMPLEADO = ""
    l_NOMBRE = ""
    l_NACIONALID = ""
    l_SEXO = ""
    l_ESTCIVIL = ""
    l_FECNACTO = ""
    l_EMAILP = ""
    l_FECINGRESO = ""
    l_calle = ""
    l_NUMERO = ""
    l_piso = ""
    l_depto = ""
    l_LOCALIDAD = ""
    l_PROVINICIA = ""
    l_CODPOSTAL = ""
    l_TELEFONO = ""
    l_NROCUIL = ""
    l_NRORNIC = ""
    l_APODERADO = ""
    l_CONFIDENC = ""
    l_CONDSINDIC = ""
    l_CODOSOCIAL = ""
    l_REBCONTPAT = ""
    l_CONDICION = ""
    l_CODJUBILAC = ""
    l_ACTIVIDAD = ""
    l_RCODE = ""
    l_ADATE = ""
    l_SDOJORNAL = ""
    l_IMPUTACION = ""
    
    Flog.writeline "===================================================================="
    For b = 0 To node.childNodes.length - 1
        Flog.writeline node.childNodes(b).nodeName & " :" & node.childNodes.Item(b).Text
        Select Case CStr(node.childNodes(b).nodeName)
            Case "f_ACODE"
                l_ACODE = node.childNodes.Item(b).Text
            Case "f_EMPLEADO"
                l_EMPLEADO = node.childNodes.Item(b).Text
            Case "f_NOMBRE"
                l_NOMBRE = node.childNodes.Item(b).Text
            Case "f_NACIONALID"
                l_NACIONALID = node.childNodes.Item(b).Text
            Case "f_SEXO"
                l_SEXO = node.childNodes.Item(b).Text
            Case "f_ESTCIVIL"
                l_ESTCIVIL = node.childNodes.Item(b).Text
            Case "f_FECNACTO"
                l_FECNACTO = node.childNodes.Item(b).Text
            Case "f_EMAILP"
                l_EMAILP = node.childNodes.Item(b).Text
            Case "f_FECINGRESO"
                l_FECINGRESO = node.childNodes.Item(b).Text
            Case "f_CALLE"
                l_calle = node.childNodes.Item(b).Text
            Case "f_NUMERO"
                l_NUMERO = node.childNodes.Item(b).Text
            Case "f_PISO"
                l_piso = node.childNodes.Item(b).Text
            Case "f_DEPTO"
                l_depto = node.childNodes.Item(b).Text
            Case "f_LOCALIDAD"
                l_LOCALIDAD = node.childNodes.Item(b).Text
            Case "f_PROVINICIA"
                l_PROVINICIA = node.childNodes.Item(b).Text
            Case "f_CODPOSTAL"
                l_CODPOSTAL = node.childNodes.Item(b).Text
            Case "f_TELEFONO"
                l_TELEFONO = node.childNodes.Item(b).Text
            Case "f_NROCUIL"
                l_NROCUIL = node.childNodes.Item(b).Text
            Case "f_NRORNIC"
                l_NRORNIC = node.childNodes.Item(b).Text
            Case "f_APODERADO"
                l_APODERADO = node.childNodes.Item(b).Text
            Case "f_CONFIDENC"
                l_CONFIDENC = node.childNodes.Item(b).Text
            Case "f_CONDSINDIC"
                l_CONDSINDIC = node.childNodes.Item(b).Text
            Case "f_CODOSOCIAL"
                l_CODOSOCIAL = node.childNodes.Item(b).Text
            Case "f_REBCONTPAT"
                l_REBCONTPAT = node.childNodes.Item(b).Text
            Case "f_CONDICION"
                l_CONDICION = node.childNodes.Item(b).Text
            Case "f_CODJUBILAC"
                l_CODJUBILAC = node.childNodes.Item(b).Text
            Case "f_ACTIVIDAD"
                l_ACTIVIDAD = node.childNodes.Item(b).Text
            Case "f_RCODE"
                l_RCODE = node.childNodes.Item(b).Text
            Case "f_ADATE"
                l_ADATE = node.childNodes.Item(b).Text
            Case "f_SDOJORNAL"
                l_SDOJORNAL = node.childNodes.Item(b).Text
            Case "f_IMPUTACION"
                l_IMPUTACION = node.childNodes.Item(b).Text
            Case Else
                Flog.writeline Espacios(Tabulador * 0) & "ERROR: EXISTE un tag no declarado: " & node.childNodes(b).nodeName
        End Select
    Next b
    
    'LineaCarga = NroLinea
    
'    l_ACODE = node.selectSingleNode("//f_ACODE").Text
'
'    l_EMPLEADO = node.selectSingleNode("//f_EMPLEADO").Text
'    l_NOMBRE = node.selectSingleNode("//f_NOMBRE").Text
'    l_NACIONALID = node.selectSingleNode("//f_NACIONALID").Text
'    l_SEXO = node.selectSingleNode("//f_SEXO").Text
'    l_ESTCIVIL = node.selectSingleNode("//f_ESTCIVIL").Text
'    l_FECNACTO = node.selectSingleNode("//f_FECNACTO").Text
'    l_EMAILP = node.selectSingleNode("//f_EMAILP").Text
'    l_FECINGRESO = node.selectSingleNode("//f_FECINGRESO").Text
'
'    l_calle = node.selectSingleNode("//f_CALLE").Text
'    l_NUMERO = node.selectSingleNode("//f_NUMERO").Text
'    l_piso = node.selectSingleNode("//f_PISO").Text
'    l_depto = node.selectSingleNode("//f_DEPTO").Text
'    l_LOCALIDAD = node.selectSingleNode("//f_LOCALIDAD").Text
'    l_PROVINICIA = node.selectSingleNode("//f_PROVINICIA").Text
'    l_CODPOSTAL = node.selectSingleNode("//f_CODPOSTAL").Text
'    l_TELEFONO = node.selectSingleNode("//f_TELEFONO").Text
'
'    l_NROCUIL = node.selectSingleNode("//f_NROCUIL").Text
'
'    l_NRORNIC = node.selectSingleNode("//f_NRORNIC").Text
'
'    l_APODERADO = node.selectSingleNode("//f_APODERADO").Text
'    l_CONFIDENC = node.selectSingleNode("//f_CONFIDENC").Text
'    l_CONDSINDIC = node.selectSingleNode("//f_CONDSINDIC").Text
'    l_CODOSOCIAL = node.selectSingleNode("//f_CODOSOCIAL").Text
'    l_REBCONTPAT = node.selectSingleNode("//f_REBCONTPAT").Text
'    l_CONDICION = node.selectSingleNode("//f_CONDICION").Text
'    l_CODJUBILAC = node.selectSingleNode("//f_CODJUBILAC").Text
'    l_ACTIVIDAD = node.selectSingleNode("//f_ACTIVIDAD").Text
'    l_RCODE = node.selectSingleNode("//f_RCODE").Text
'    l_ADATE = node.selectSingleNode("//f_ADATE").Text
'    l_SDOJORNAL = node.selectSingleNode("//f_SDOJORNAL").Text

    l_ADATE = a_fecha(l_ADATE)
    'tipmotnro = getMapeoSap("TIPOMOTIVO", l_RCODE)
    tipmotnro = "7"
    
    If l_RCODE <> "" Then
        caunro = getMapeoSap("CAUSA", l_RCODE)
        If caunro = "" Then
            Flog.writeline Espacios(Tabulador * 0) & "ATENCION: No se encontro la causa, se coloca en NULL."
            caunro = "NULL"
        End If
    End If
    
    'Exit Sub
    If l_EMPLEADO = "" Then
        empleg = "0"
    Else
        empleg = l_EMPLEADO
    End If
    
    ternom = ""
    ternom2 = ""
    terape = ""
    terape2 = ""
    If l_NOMBRE <> "" Then
        cadena = Split(Trim(l_NOMBRE), ",")
        If UBound(cadena) >= 0 Then
            'Apellidos
            subCadena = Split(Trim(cadena(0)), " ")
            If UBound(subCadena) >= 0 Then
                terape = subCadena(0)
            End If
            If UBound(subCadena) >= 1 Then
                terape2 = subCadena(1)
            End If
            'nombres
            If UBound(cadena) >= 1 Then
                subCadena = Split(Trim(cadena(1)), " ")
                If UBound(subCadena) >= 0 Then
                    ternom = subCadena(0)
                End If
                If UBound(subCadena) >= 1 Then
                    ternom2 = subCadena(1)
                End If
            End If
        End If
    End If
    
    terfecnac = a_fecha(l_FECNACTO)
    
    If l_SEXO <> "" Then
        If l_SEXO = "M" Then
            tersex = "-1"
        Else
            tersex = "0"
        End If
    Else
        tersex = ""
    End If
    
    teremail = l_EMAILP
    
    If l_NACIONALID <> "" Then
        nacionalnro = getMapeoSap("NACIONALIDAD", l_NACIONALID)
        If nacionalnro = "" Then
            nacionalnro = "NULL"
        Else
            If Not IsNumeric(nacionalnro) Then
                nacionalnro = "NULL"
            End If
        End If
    End If
    
    If l_ESTCIVIL <> "" Then
        estcivnro = getMapeoSap("ESTCIVIL", l_ESTCIVIL)
    End If
    
    Legajo = l_EMPLEADO
    
    empfecalta = a_fecha(l_FECINGRESO)
    
    empremu = l_SDOJORNAL
    
    calle = l_calle
    callenro = l_NUMERO
    piso = l_piso
    depto = l_depto
    
    paisnro = ""
    If paisnro = "" Then
        paisnro = "3"
    End If
    
    If l_PROVINICIA <> "" Then
        provnro = getMapeoSap("PROVINCIA", l_PROVINICIA)
        If provnro = "" Then provnro = "1" 'no informada
        Flog.writeline Espacios(Tabulador * 0) & "No se configuro la PROVINCIA, se establece en: 1."
    End If
    
    'locnro = getMapeoSap("LOCALIDAD", l_LOCALIDAD)
    locnro = getLocalidadNro(l_LOCALIDAD, provnro, paisnro) '10/01/2012
    
    codigopostal = l_CODPOSTAL
    
    'telnro = Replace(l_TELEFONO, "-", "")
    telnro = l_TELEFONO
    
    'paisnro = getMapeoSap("PAIS", "ARGENTINA")
    
        
    If Len(l_NROCUIL) = 11 Then
        Flog.writeline Espacios(Tabulador * 0) & "Convirtiendo CUIL: " & l_NROCUIL
        l_NROCUIL = Left(l_NROCUIL, 2) & "-" & Mid(l_NROCUIL, 3, 8) & "-" & Right(l_NROCUIL, 1)
        Flog.writeline Espacios(Tabulador * 0) & "CUIL Generado: " & l_NROCUIL
    End If
    If Not Cuil_Valido(l_NROCUIL, Mensaje) Then
        Flog.writeline Espacios(Tabulador * 0) & Mensaje
        Flog.writeline Espacios(Tabulador * 0) & "No se puede obtener el DNI en base al CUIL."
        nrocuil = ""
        nrodoc = ""
    Else
        nrocuil = l_NROCUIL
        nrodoc = Mid(l_NROCUIL, 4, 8)
    End If
        
    
    'Dim rs_sub As New ADODB.Recordset
    Dim a As Integer
    
'    Select Case l_ACODE
'        Case "U1" 'Hire - Nuevo
'            Debug.Print ""
'            'insertarEmpleado
'        Case "U2" 'Rehire - renuevo
'            'ModificarEmpleado
'        Case "U6" 'Return from Leave
'
'        Case "U8" 'Separation
'            'BajaEmpleado
'        Case Else
'
'    End Select
    
    
    Err.Clear
    
    On Error GoTo ErrorEmpleado
    
    MyBeginTrans
    
        ' Busco el tercero
    Ternro = buscar_tercero(empleg)

    If Ternro = "0" And l_ACODE <> "U1" Then
        Flog.writeline Espacios(Tabulador * 0) & "ERROR al identificar el Tercero - Legajo: " & empleg
        Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------"
        HuboError = True
        HuboErrorLocal = True
        Exit Sub
    End If
    
    'Si encontre el ternro es una modificacion sino es un alta
    'If Ternro <> "0" Then
        Ternro = insertar_tercero(Ternro, ternom, ternom2, terape, terape2, terfecnac, tersex, teremail, nacionalnro, estcivnro)
    'End If
    
    ' -- ter_tip tipo tercero --
    insertar_tertip Ternro, 1
    
    ' -- Complemento --
    insertar_complemento Ternro, Legajo, ternom, ternom2, terape, terape2, empfecalta, teremail, empremu
    
    ' -- Fases --
    insertar_fases Ternro, empfecalta, l_ACODE, caunro, l_ADATE
    
    '--Inserto el CUIL--
    If nrodoc <> "" Then
        insertar_documento Ternro, nrocuil, 10
    End If
    '--Inserto el Documento--
    If nrodoc <> "" Then
        insertar_documento Ternro, nrodoc, 1
    End If
    
    '--Inserto el Domicilio--
    If calle <> "" Then
        NroDom = insertar_domicilio(Ternro, calle, callenro, piso, depto, locnro, provnro, codigopostal, paisnro)
    End If
    
    '--Telefonos--
    If l_TELEFONO <> "" Then
        'telnro = Replace(l_TELEFONO, "-", "")
        telnro = l_TELEFONO
        insertar_telefono NroDom, telnro
    End If
    
    '-- Estructura EMPRESA --
    'F_NRORNIC   :  Empresa "CAR1" para General Mills, "CAR2" para La Salteña.
    If l_NRORNIC <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- NRONIC -----"
        insertar_Estructura Ternro, 10, getMapeoSap("EMPRESA", l_NRORNIC), l_ADATE, tipmotnro
    End If
    
    'F_APODERADO : Indica la planta donde desarrola actividad. San Fernando o Burzaco. Tomar estructura sucursal.
    If l_APODERADO <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- APODERADO -----"
        insertar_Estructura Ternro, 1, getMapeoSap("SUCURSAL", l_APODERADO), l_ADATE, tipmotnro
    End If
    
    
    'F_CONFIDENC :  0 para Nómina General, 1 para Confidenciales. Tomar cód. externo de la estructura Grupo de Seguridad
    If l_CONFIDENC <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- CONFIDENC -----"
        'Obtengo la empresa del empleado (estrnro)
        Estrnro = getEstrNro(Ternro, 10)
        Flog.writeline Espacios(Tabulador * 1) & "Empresa estructura obtenida - Estrnro:" & Estrnro
        
        'Obtengo la empresa del empleado (empnro)
        Empnro = getEmpnro(Estrnro)
        Flog.writeline Espacios(Tabulador * 1) & "Empresa obtenida - Empnro:" & Empnro
        
        ' Obtengo el grupo de seguridad
        Estrnro = getEstrnro_EstrCodExt_empnro(7, l_CONFIDENC, Empnro)
        Flog.writeline Espacios(Tabulador * 1) & "Grupo de Seguridad obtenido - estrnro:" & Estrnro

        insertar_Estructura Ternro, 7, Estrnro, l_ADATE, tipmotnro
        'insertar_Estructura Ternro, 7, getMapeoSap("Grupo de Seguridad", l_CONFIDENC), l_ADATE, tipmotnro
    End If
    
    'F_CONDSINDIC: Código de Sindicato. Tomar cod. Externo de la estructura sindicato.
    If l_CONDSINDIC <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- CONDSINDIC -----"
        
        'Obtengo la empresa del empleado (estrnro)
        Estrnro = getEstrNro(Ternro, 10)
        'Obtengo la empresa del empleado (empnro)
        Empnro = getEmpnro(Estrnro)
        ' Obtengo el Sindicato
        Estrnro = getEstrnro_EstrCodExt_empnro(16, l_CONDSINDIC, Empnro)
        
        insertar_Estructura Ternro, 16, Estrnro, l_ADATE, tipmotnro
        'insertar_Estructura Ternro, 16, getMapeoSap("SINDICATO", l_CONDSINDIC), l_ADATE, tipmotnro
    End If
    
    'F_CODOSOCIAL: Código de Obra Social. Tomar cod. Externo de la estructura OS elegida.
    If l_CODOSOCIAL <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- CODOSOCIAL -----"
        insertar_Estructura Ternro, 23, getMapeoSap("PLAN DE OS ELEGIDO", l_CODOSOCIAL), l_ADATE, tipmotnro
    End If
    
    'F_REBCONTPAT: Código de Contratación para SICOSS (En las altas viene 014 y al mes siguiente del fin de Contrato 008) Tomar cod. Externo de la estructura contrato
    If l_REBCONTPAT <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- REBCONTPAT -----"
        insertar_Estructura Ternro, 18, getMapeoSap("CONTRATO", l_REBCONTPAT), l_ADATE, tipmotnro
    End If
    
    'F_CONDICION : Siempre viene 1 (Mensual para Waldbott)
    If l_CONDICION <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- CONDICION -----"
        insertar_Estructura Ternro, 22, getMapeoSap("FORMA DE LIQUIDACION", l_CONDICION), l_ADATE, tipmotnro
    End If
    
    'F_CODJUBILAC: Viene siempre 00 (Reparto para Waldbott)
    If l_CODJUBILAC <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- CODJUBILAC -----"
        insertar_Estructura Ternro, 15, getMapeoSap("CAJA DE JUBILACION", l_CODJUBILAC), l_ADATE, tipmotnro
    End If
    
    'F_ACTIVIDAD : Viene siempre 49 (Para Waldbott es el códgio de Actividad No Clasificada).
    If l_ACTIVIDAD <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- ACTIVIDAD -----"
        insertar_Estructura Ternro, 29, getMapeoSap("ACTIVIDAD", l_ACTIVIDAD), l_ADATE, tipmotnro
    End If
    
    'l_IMPUTACION: Centro de costos, biene el codigo Externo de la estructura
    If l_IMPUTACION <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- IMPUTACION -----"
        If IsNumeric(l_IMPUTACION) Then
            l_IMPUTACION = CStr(CLng(l_IMPUTACION))
        Else
            l_IMPUTACION = l_IMPUTACION
        End If
        Estrnro = getEstrnro_EstrCodExt(5, l_IMPUTACION)
        If Estrnro <> "" Then
            insertar_Estructura Ternro, 5, Estrnro, l_ADATE, tipmotnro
        End If
    End If

    'F_SDOJORNAL : Sueldo Basico del personal fuera de convenio (este campo no se actualiza para los de convenio aunque SAP envíe el dato) Cpto 01000, tener en cuenta estrctura convenio.
    If l_SDOJORNAL <> "" Then
        Flog.writeline Espacios(Tabulador * 1) & "----- SDOJORNAL -----"
        'getMapeoSap("CONVENIO", "00") 'cableado - documentar
        If esFueraDeConvenio(Ternro, getMapeoSap("CONVENIO", "00"), l_ADATE, 19) Then
            Insertar_Novedad Ternro, "01000", "1", l_SDOJORNAL, l_ADATE, "", "Inrefase SAP-RHPRO"
        End If
    End If
    
    'rs_sub.Close
    
    If rs.State = adStateOpen Then rs.Close
    'If rs_sql.State = adStateOpen Then rs_sql.Close
    
    Err.Clear
    
    MyCommitTrans
    
    Exit Sub

ErrorEmpleado:
    Flog.writeline Espacios(Tabulador * 0) & " Error: ErrorEmpleado."
    Flog.writeline Espacios(Tabulador * 0) & "error al insertar el tercero " & ternom & "," & terape
    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & error
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    HuboError = True
    HuboErrorLocal = True
    MyRollbackTrans
    Exit Sub

    
'MyBeginTrans

MyCommitTrans



End Sub


'Function insertarEmpleado()
'
'    Dim rs_sub As New ADODB.Recordset
'    Dim a As Integer
'    'Dim ActPasos As Boolean
'    'Dim estact
'    'Dim carrcomp
'    'Dim Provincia As Integer
'
'    Err.Clear
'
'    On Error GoTo ErrorEmpleado
'
'    '--Inserto el Tercero--
'    Ternro = insertar_tercero("", ternom, ternom2, terape, terape2, terfecnac, tersex, teremail, nacionalnro, estcivnro)
'
'    If Ternro <> 0 Then
'
'        On Error GoTo 0
'        On Error Resume Next
'        'si da error  no puedo seguir
'
'        'ter_tip tipo tercero
'        insertar_tertip Ternro, 1
'
'        '--Complemento--
'        insertar_complemento Ternro, Legajo, ternom, ternom2, terape, terape2, empfecalta, teremail, empremu
'
'        'Fases
'        insertar_fases Ternro, empfecalta, l_ACODE
'
'        '--Inserto el Documento--
'        insertar_documento Ternro, nrodoc, 10
'
'
'        '--Inserto el Domicilio--
'        NroDom = insertar_domicilio(Ternro, calle, callenro, piso, depto, locnro, provnro, codigopostal, paisnro)
'
'        '--Telefonos--
'        telnro = Replace(l_TELEFONO, "-", "")
'        insertar_telefono NroDom, telnro
'
'        '-- Estructura EMPRESA --
'        'insertar_Estructura 10
'
'
'    End If
'
'    rs_sub.Close
'
'    If rs.State = adStateOpen Then rs.Close
'    'If rs_sql.State = adStateOpen Then rs_sql.Close
'
'    Err.Clear
'    'IniciarVariablesBumeran
'
'    Exit Function
'
'ErrorEmpleado:
'    Flog.writeline Espacios(Tabulador * 0) & "error al insergar el tercero " & ternom & "," & terape
'    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
'    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
'    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
'    Flog.writeline Espacios(Tabulador * 0) & error
'    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
'    If rs.State = adStateOpen Then rs.Close
'    'IniciarVariablesBumeran
'    Exit Function
'
'End Function

Function validatelefono(cadena As String) As String
    Dim a As Integer
    Dim car As String
    Dim cadenacompleta As String
    For a = 1 To Len(cadena)
        car = Asc(Mid(cadena, a, 1))
        If Not (car > 47 And car < 58) Or (car > 39 And car < 43) Or (car = 45) Or (car = 32) Or (car = 35) Then
            cadenacompleta = CStr(cadenacompleta) & CStr(Chr(car))
        Else
            cadenacompleta = cadenacompleta & CStr(Chr(car))
        End If
    Next a
    validatelefono = cadenacompleta
End Function


Function getNacionalidadNro(cod As String) As Long
    'Faltan los codigos de las nacionalidades
    Dim coddes As String
    If cod = "AR" Then
        coddes = "ARGENTINA"
    Else
    
    End If
    getNacionalidadNro = TraerCodNacionalidad(coddes)
    
    'Dim rs_sub As New ADODB.Recordset
    'StrSql = "INSERT INTO idinivel (idinivdesabr) "
    'StrSql = StrSql & " VALUES('" & idinivdesabr & "')"
    'objConn.Execute StrSql, , adExecuteNoRecords
    'StrSql = " SELECT MAX(idinivnro) AS Maxidinivnro FROM idinivel "
    'OpenRecordset StrSql, rs_sub
    'getNacionalidadNro = CInt(rs_sub!Maxidinivnro)
End Function

Function getLocalidadNro(locdesc As String, provnro As String, paisnro As String) As Long
    ' Lisandro Moro - 10/01/2012 - Se creo
    ' Lisandro Moro - 30/01/2012 - se agregaron la provincia y el pais
    Dim coddes As String
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Localidad: " & locdesc & " - provnro:" & provnro & " - paisnro: " & paisnro
    Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(locdesc) Then
        StrSql = " SELECT locnro FROM localidad WHERE locdesc = '" & locdesc & "'"
        StrSql = StrSql & " AND provnro = " & provnro
        StrSql = StrSql & " AND paisnro = " & paisnro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "No se encontro la Localidad: " & locdesc
            getLocalidadNro = 1
        Else
            getLocalidadNro = rs_sub!locnro
            Flog.writeline Espacios(Tabulador * 1) & "Se encontro la Localidad: " & locdesc & "(" & getLocalidadNro & ")"
        End If
    Else
        getLocalidadNro = 1 'NO INFORMADA
    End If
    
    If rs_sub.State = adStateOpen Then rs_sub.Close
    Set rs_sub = Nothing
    
    
End Function
Function getProvinciaNro(cod As String)
    Dim coddes As String
    If cod = "BA" Then
        coddes = "BUENOS AIRES"
    Else
        
    End If
    getProvinciaNro = TraerCodProvincia(coddes)

End Function

Function getEstadoCivilNro(cod As String)
    Dim coddes As String
    If cod = "S" Then
        coddes = "SOLTERO"
    End If
    getEstadoCivilNro = TraerCodEstadoCivil(coddes)
End Function

Function getEstrnro_EstrCodExt_empnro(Tenro, CodExt, Empnro)
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Busca la estructura segun el cod extendido y la empresa.
    ' Autor      : Lisandro Moro
    ' Fecha      : 15/11/2011
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(CodExt) Then
        StrSql = " SELECT estrnro FROM estructura WHERE tenro = " & Tenro & " AND estrcodext = '" & CodExt & "'"
        StrSql = StrSql & " AND empnro = " & Empnro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            getEstrnro_EstrCodExt_empnro = ""
            Flog.writeline Espacios(Tabulador * 1) & "ERROR - No se encontro la estructura por codigo externo : '" & CodExt & "' - Tenro: " & Tenro & " - empnro: " & Empnro
            HuboErrorLocal = True
        Else
            getEstrnro_EstrCodExt_empnro = rs_sub!Estrnro
            Flog.writeline Espacios(Tabulador * 1) & "Se encontro la estructura (" & getEstrnro_EstrCodExt_empnro & ") por codigo externo : '" & CodExt & "' - Tenro: " & Tenro & " - empnro: " & Empnro
        End If
        If rs_sub.State = adStateOpen Then rs_sub.Close
        Set rs_sub = Nothing
    Else
        getEstrnro_EstrCodExt_empnro = ""
        Flog.writeline Espacios(Tabulador * 1) & "ERROR - No se encontro la estructura por codigo externo : '" & CodExt & "' - Tenro: " & Tenro & " - empnro: " & Empnro
    End If

End Function

Function getEstrnro_EstrCodExt(Tenro, CodExt)
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Busca la estructura segun el cod extendido.
    ' Autor      : Lisandro Moro
    ' Fecha      : 15/11/2011
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(CodExt) Then
        StrSql = " SELECT estrnro FROM estructura WHERE tenro = " & Tenro & " AND estrcodext = '" & CodExt & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            getEstrnro_EstrCodExt = ""
            Flog.writeline Espacios(Tabulador * 1) & "ERROR - No se encontro la estructura por codigo externo : '" & CodExt & "' - Tenro: " & Tenro
            HuboErrorLocal = True
        Else
            getEstrnro_EstrCodExt = rs_sub!Estrnro
        End If
        If rs_sub.State = adStateOpen Then rs_sub.Close
        Set rs_sub = Nothing
    Else
        getEstrnro_EstrCodExt = ""
    End If
    
End Function
Function getEstrNro(Ternro, Tenro)
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Busca la estructura del empleado.
    ' Autor      : Lisandro Moro
    ' Fecha      : 10/01/2012
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim rs_sub As New ADODB.Recordset
    Dim l_fecha As Date
    'l_fecha = Date
    l_fecha = l_ADATE

    StrSql = " SELECT estrnro FROM his_estructura WHERE ternro = " & Ternro
    StrSql = StrSql & " AND tenro = " & Tenro
    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(l_fecha) & " AND (htethasta >= " & ConvFecha(l_fecha) & " OR htethasta is null)) "
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        'Texto = "El empleado posee la estructura activa a la fecha - ternro: " & Ternro & " - Fecha : " & l_fecha
        Flog.writeline Espacios(Tabulador * 1) & "ERROR - El empleado NO posee la estructura activa a la fecha - ternro: " & Ternro & " - Fecha : " & l_fecha
        HuboErrorLocal = True
        getEstrNro = ""
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado posee la estructura activa a la fecha - ternro: " & Ternro & " - Fecha : " & l_fecha
        getEstrNro = rs_sub!Estrnro
    End If
    rs_sub.Close
End Function

Function getEmpnro(Estrnro)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca la empresa de la estructura.
' Autor      : Lisandro Moro
' Fecha      : 10/01/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
If Not EsNulo(Estrnro) Then
    StrSql = " SELECT empnro FROM empresa WHERE estrnro = " & Estrnro
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        getEmpnro = ""
        Flog.writeline Espacios(Tabulador * 1) & "ERROR - No se encontro la empresa: estrnro = " & Estrnro
        HuboErrorLocal = True
    Else
        getEmpnro = rs_sub!Empnro
    End If
    If rs_sub.State = adStateOpen Then rs_sub.Close
    Set rs_sub = Nothing
Else
    getEmpnro = ""
End If

End Function

Function getMapeoSap(Tabla, codigo)
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Busca los codigos mapeados.
    ' Autor      : Lisandro Moro
    ' Fecha      : 14/07/2011
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(codigo) Then
        StrSql = " SELECT codinterno FROM mapeo_sap WHERE tablaref = '" & Tabla & "' AND codexterno = '" & codigo & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            getMapeoSap = ""
            Flog.writeline Espacios(Tabulador * 1) & "ERROR - No se encontro el mapeo para la tabla: " & Tabla & " - codigo: " & codigo
            HuboErrorLocal = True
        Else
            getMapeoSap = rs_sub!codinterno
            Flog.writeline Espacios(Tabulador * 1) & "Se encontro el mapeo: " & getMapeoSap & " para la tabla: " & Tabla & " - codigo: " & codigo
        End If
        If rs_sub.State = adStateOpen Then rs_sub.Close
        Set rs_sub = Nothing
    Else
        getMapeoSap = ""
        Flog.writeline Espacios(Tabulador * 1) & "ERROR - No se encontro el mapeo para la tabla: " & Tabla & " - codigo: " & codigo
        HuboErrorLocal = True
    End If

    
End Function

Function insertar_tercero(l_ternro, l_ternom, l_ternom2, l_terape, l_terape2, l_terfecnac, l_tersex, l_teremail, l_nacionalnro, l_estcivnro)
    Dim rs_sub As New ADODB.Recordset
    'tersex
    If Ternro = "0" Or Ternro = "" Then
    
        If l_tersex = "" Then l_tersex = "0" 'el unico requisito
        
        If l_terape <> "" Or l_ternom <> "" Then
        
            StrSql = " INSERT INTO tercero (ternom,ternom2, terape, terape2 ,terfecnac,tersex,teremail, nacionalnro, estcivnro) VALUES ("
            StrSql = StrSql & "'" & l_ternom & "'"
            If l_ternom2 = "" Then
                StrSql = StrSql & ", null"
            Else
                StrSql = StrSql & ",'" & l_ternom2 & "'"
            End If
            StrSql = StrSql & ",'" & l_terape & "'"
            If l_terape2 = "" Then
                StrSql = StrSql & ", null"
            Else
                StrSql = StrSql & ",'" & l_terape2 & "'"
            End If
            StrSql = StrSql & "," & ConvFecha(l_terfecnac)
            StrSql = StrSql & "," & CInt(l_tersex)
            StrSql = StrSql & ",'" & l_teremail & "'"
            StrSql = StrSql & "," & l_nacionalnro
            If l_estcivnro = "" Then
                StrSql = StrSql & ",1" 'Sin datos
            Else
                StrSql = StrSql & "," & l_estcivnro
            End If
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto en la tabla de tercero"
                
            '--Obtengo el ternro--
            insertar_tercero = getLastIdentity(objConn, "tercero")
            Flog.writeline Espacios(Tabulador * 1) & "-----------------------------------------------"
            Flog.writeline Espacios(Tabulador * 1) & "Codigo de Tercero = " & Ternro
        Else
            insertar_tercero = "0"
            Flog.writeline Espacios(Tabulador * 1) & "No se puede Crear el TERCERO, el Nombre O el Apellido estan en Blanco."
            HuboErrorLocal = True
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el TERCERO - " & l_ternro
        If l_ternom <> "" Then
            StrSql = " UPDATE tercero SET ternom = '" & l_ternom & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el TERCERO Nombre:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el TERCERO(ternom:" & l_ternom & ")"
            End If
        End If
        If l_ternom2 <> "" Then
            StrSql = " UPDATE tercero SET ternom2 = '" & l_ternom2 & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el TERCERO Nombre2:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el TERCERO(ternom2:" & l_ternom2 & ")"
            End If
        End If
        If l_terape <> "" Then
            StrSql = " UPDATE tercero SET terape = '" & l_terape & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el TERCERO Apellido:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el TERCERO(terape:" & l_terape & ")"
            End If
        End If
        If l_terape2 <> "" Then
            StrSql = " UPDATE tercero SET terape2 = '" & l_terape2 & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el TERCERO Apellido2:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el TERCERO(terape2:" & l_terape2 & ")"
            End If
        End If
        'If l_empfecalta <> "" Then
        '    StrSql = " UPDATE tercero SET empfecalta = " & ConvFecha(l_empfecalta) & " WHERE ternro = " & l_ternro
        '    objConn.Execute StrSql, , adExecuteNoRecords
        '    If Err Then
        '        Flog.writeline Espacios(Tabulador * 1) &  "Error al ACTUALIZAR el TERCERO fecha alta:" & Err.Description
        '        Flog.writeline Espacios(Tabulador * 1) &  StrSql
        '        Err.Clear
        '    Else
        '        Flog.writeline Espacios(Tabulador * 1) &  "ACTUALIZO el TERCERO(fecha alta:" & l_empfecalta & ")"
        '    End If
        'End If
        If l_terfecnac <> "" Then
            StrSql = " UPDATE tercero SET terfecnac = " & ConvFecha(l_terfecnac) & " WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el TERCERO fecha terfecnac:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el TERCERO(fecha terfecnac:" & terfecnac & ")"
            End If
        End If
        If l_teremail <> "" Then
            StrSql = " UPDATE tercero SET teremail = '" & l_teremail & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el TERCERO email:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el TERCERO(email:" & l_teremail & ")"
            End If
        End If
        'l_nacionalnro , l_estcivnro
        If l_nacionalnro <> "" And l_nacionalnro <> "0" And l_nacionalnro <> "Null" Then
            StrSql = " UPDATE tercero SET nacionalnro = '" & l_nacionalnro & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el TERCERO nacionalnro:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el TERCERO(nacionalnro:" & l_nacionalnro & ")"
            End If
        End If
        If l_estcivnro <> "" Then
            StrSql = " UPDATE tercero SET estcivnro = '" & l_estcivnro & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el TERCERO estcivnro:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el TERCERO(estcivnro:" & l_estcivnro & ")"
            End If
        End If
        insertar_tercero = l_ternro
    End If
    
    If rs_sub.State = adStateOpen Then rs_sub.Close
    Set rs_sub = Nothing
    
End Function

Sub insertar_tertip(l_ternro, l_tipnro)
    Dim rs_sub As New ADODB.Recordset
    StrSql = " SELECT ternro, tipnro FROM ter_tip WHERE tipnro = 1 AND ternro = " & l_ternro
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        '--Inserto el Registro correspondiente en ter_tip--
        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & l_ternro & ",1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 0) & "Inserto el tipo de tercero 1 en ter_tip para el tercero " & l_ternro
    End If
    rs_sub.Close
    Set rs_sub = Nothing
End Sub

Sub insertar_complemento(l_ternro, l_legajo, l_ternom, l_ternom2, l_terape, l_terape2, l_empfecalta, l_teremail, l_empremu)
    Dim rs_sub As New ADODB.Recordset
    
    Flog.writeline Espacios(Tabulador * 0) & "Inserto Complemento Empleado."
        
    'empleg    'empest    'ternro    'terape    'ternom    'empnro    'expreo    'empcostohora
    
    StrSql = " SELECT ternro FROM empleado WHERE ternro = " & l_ternro
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        
        If l_legajo = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "ERROR al generar el complemento, falta el Legajo."
            HuboErrorLocal = True
            Exit Sub
        End If
        If l_ternom = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "ERROR al generar el complemento, falta el Nombre."
            HuboErrorLocal = True
            Exit Sub
        End If
        If l_terape = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "ERROR al generar el complemento, falta el Apellido."
            HuboErrorLocal = True
            Exit Sub
        End If
        If l_ternom = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "ERROR al generar el complemento, falta el Nombre."
            HuboErrorLocal = True
            Exit Sub
        End If
    
        StrSql = " INSERT INTO empleado "
        StrSql = StrSql & " (empleg, ternro, ternom, ternom2, terape, terape2, empfecalta, empfaltagr, empest, empemail)" ', empremu) "
        StrSql = StrSql & " VALUES (" & l_legajo & ", " & l_ternro & ",'" & l_ternom & "', '" & l_ternom2 & "' ,'" & l_terape & "', '" & l_terape2 & "' ," & ConvFecha(l_empfecalta) & "," & ConvFecha(l_empfecalta) & ", -1,'" & l_teremail & "')" '," & l_empremu & " ) "
        objConn.Execute StrSql, , adExecuteNoRecords
        If Err Then
            Flog.writeline Espacios(Tabulador * 1) & "Error al insertar el Complemento EMPLEADO:" & Err.Description
            Flog.writeline Espacios(Tabulador * 1) & l_legajo & " - ternro:" & l_ternro & " - ternom:" & l_ternom & " - ternom2:" & l_ternom2 & " - empfecalta:" & l_empfecalta & " - empfecalta:" & l_empfecalta & " - teremail:" & l_teremail & " - empremu:" & l_empremu
            HuboErrorLocal = True
            Flog.writeline Espacios(Tabulador * 1) & StrSql
            Err.Clear
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Inserto el Complemento EMPLEADO:"
            Flog.writeline Espacios(Tabulador * 1) & l_legajo & " - ternro:" & l_ternro & " - ternom:" & l_ternom & " - ternom2:" & l_ternom2 & " - empfecalta:" & l_empfecalta & " - empfecalta:" & l_empfecalta & " - teremail:" & l_teremail & " - empremu:" & l_empremu
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & " Actualizando Datos del Complemento del EMPLEADO - ternro: " & l_ternro
        Flog.writeline Espacios(Tabulador * 1) & "Parametros : legajo:" & l_legajo & " - ternro:" & l_ternro & " - ternom:" & l_ternom & " - ternom2:" & l_ternom2 & " - empfecalta:" & l_empfecalta & " - empfecalta:" & l_empfecalta & " - teremail:" & l_teremail & " - empremu:" & l_empremu
        If l_ternom <> "" Then
            StrSql = " UPDATE empleado SET ternom = '" & l_ternom & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el Complemento EMPLEADO Nombre:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Complemento EMPLEADO(ternom:" & l_ternom & ")"
            End If
        End If
        If l_ternom2 <> "" Then
            StrSql = " UPDATE empleado SET ternom2 = '" & l_ternom2 & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el Complemento EMPLEADO Nombre2:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Complemento EMPLEADO(ternom2:" & l_ternom2 & ")"
            End If
        End If
        If l_terape <> "" Then
            StrSql = " UPDATE empleado SET terape = '" & l_terape & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el Complemento EMPLEADO Apellido:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Complemento EMPLEADO(terape:" & l_terape & ")"
            End If
        End If
        If l_terape2 <> "" Then
            StrSql = " UPDATE empleado SET terape2 = '" & l_terape2 & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el Complemento EMPLEADO Apellido2:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Complemento EMPLEADO(terape2:" & l_terape2 & ")"
            End If
        End If
        If l_empfecalta <> "" Then
            StrSql = " UPDATE empleado SET empfecalta = " & ConvFecha(l_empfecalta) & " WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el Complemento EMPLEADO fecha alta:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Complemento EMPLEADO(fecha alta:" & l_empfecalta & ")"
            End If
        End If
        If l_empfecalta <> "" Then
            StrSql = " UPDATE empleado SET empfaltagr = " & ConvFecha(l_empfecalta) & " WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el Complemento EMPLEADO fecha alta2:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Complemento EMPLEADO(fecha alta2:" & l_empfecalta & ")"
            End If
        End If
        If l_teremail <> "" Then
            StrSql = " UPDATE empleado SET empemail = '" & l_teremail & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el Complemento EMPLEADO email:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Complemento EMPLEADO(email:" & l_teremail & ")"
            End If
        End If
        If l_empremu <> "" Then
            StrSql = " UPDATE empleado SET empremu = '" & l_empremu & "' WHERE ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el Complemento EMPLEADO empremu:" & Err.Description
                Flog.writeline Espacios(Tabulador * 1) & StrSql
                HuboErrorLocal = True
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Complemento EMPLEADO(empremu:" & l_empremu & ")"
            End If
        End If

    End If
    rs_sub.Close
    Set rs_sub = Nothing
End Sub

Sub insertar_fases(l_ternro, l_empfecalta, l_ACODE, caunro, l_bajfec)
    Dim rs_sub As New ADODB.Recordset
    Dim l_htethasta As String
    
    Flog.writeline Espacios(Tabulador * 0) & " FASES "
    
    If l_ACODE <> "U8" Then
        If l_empfecalta = "" Then
            'Flog.writeline Espacios(Tabulador * 1) & "No se actualizan fases. No hay fecha de alta."
            Exit Sub
        End If
        Flog.writeline Espacios(Tabulador * 1) & " -- FASES --"
        
        l_htethasta = DateAdd("d", -1, l_empfecalta)
        
        'Si tiene una fase con la misma fecha desde la dejo abierta
        StrSql = " SELECT * FROM fases WHERE empleado = " & l_ternro
        StrSql = StrSql & " AND altfec = " & ConvFecha(l_empfecalta)
        OpenRecordset StrSql, rs_sub
        If Not rs_sub.EOF Then
            If EsNulo(rs_sub("bajfec")) Then
                Texto = " El empleado posee la fase activa - ternro: " & l_ternro & " - Fecha desde: " & l_empfecalta
                Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
                Exit Sub
            Else
                StrSql = " UPDATE fases SET bajfec = null, estado = -1 "
                StrSql = StrSql & " WHERE empleado = " & l_ternro
                StrSql = StrSql & " AND altfec = " & ConvFecha(l_empfecalta)
                objConn.Execute StrSql, , adExecuteNoRecords
                Texto = " Actualizo la fase actual - ternro: " & l_ternro & " - Fecha desde: " & l_empfecalta
                Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
                Exit Sub
            End If
        End If
        
        StrSql = " SELECT * FROM fases WHERE empleado = " & l_ternro
        StrSql = StrSql & " AND ( altfec <= " & ConvFecha(l_empfecalta) & " AND (bajfec >= " & ConvFecha(l_empfecalta) & " OR bajfec is null)) "
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
            StrSql = StrSql & " VALUES( " & l_ternro & "," & ConvFecha(l_empfecalta) & ",null,null,-1,-1,-1,-1,-1,-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = "Inserto la Fase - ternro: " & l_ternro & " - Fecha desde: " & l_empfecalta
            Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
        Else
            
            ' BORRO futuras fases
            StrSql = " SELECT * FROM fases WHERE empleado = " & l_ternro
            StrSql = StrSql & " AND  altfec > " & ConvFecha(l_empfecalta)
            OpenRecordset StrSql, rs_sub
            If rs_sub.EOF Then
            Else
                Do While Not rs_sub.EOF
                    StrSql = " DELETE FROM fases WHERE fases.fasnro = " & rs_sub!fasnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Texto = "Borro la Fase - ternro: " & l_ternro & " - Fecha desde: " & rs_sub!altfec & " - hasta " & rs_sub!bajfec
                    Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
                    
                    rs_sub.MoveNext
                Loop
            End If
            
            'cierro la fase anterior y creo la siguiente
            'Cierro las fases
            StrSql = " UPDATE fases SET estado=0, bajfec = " & ConvFecha(l_htethasta)
            StrSql = StrSql & " WHERE empleado = " & l_ternro
            StrSql = StrSql & " AND (altfec <= " & ConvFecha(l_empfecalta)
            StrSql = StrSql & " AND (bajfec >= " & ConvFecha(l_empfecalta) & " OR bajfec is null)) "
            'StrSql = StrSql & " AND altfec = " & rs_sub("altfec")
            'StrSql = StrSql & " AND (bajfec is null OR bajfec >= " & l_empfecalta & " )"
            'StrSql = StrSql & " AND (bajfec is null) "
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = "CIERRO la Fase - ternro: " & l_ternro & " - Fecha hasta: " & Date
            Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
            
            
            
            StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
            StrSql = StrSql & " VALUES( " & l_ternro & "," & ConvFecha(l_empfecalta) & ",null,"
            StrSql = StrSql & "null,-1,-1,-1,-1,-1,-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = "Inserto la Fase - ternro: " & l_ternro & " - Fecha desde: " & l_empfecalta
            Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
            
            Texto = ": " & "Fase Actualizada - "
            Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
            
        End If
        
        rs_sub.Close
        Set rs_sub = Nothing
    Else
        If l_bajfec = "" Then
           l_bajfec = Date
        End If
        
        'Cierro las fases
        StrSql = " UPDATE fases SET estado = 0, bajfec = " & ConvFecha(l_bajfec)
        StrSql = StrSql & " , caunro= " & caunro
        StrSql = StrSql & " WHERE empleado = " & l_ternro
        StrSql = StrSql & " AND bajfec is null "
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Texto = "CIERRO la Fase - ternro: " & l_ternro & " - Fecha hasta: " & l_bajfec
        Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
        
        Texto = ": " & "Fase Actualizada - "
        Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
        
        StrSql = " UPDATE empleado SET empest = 0 WHERE ternro = " & l_ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        If Err Then
            Flog.writeline Espacios(Tabulador * 1) & "Error al ACTUALIZAR el Complemento EMPLEADO ESTADO:" & Err.Description
            Flog.writeline Espacios(Tabulador * 1) & StrSql
            Err.Clear
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Complemento EMPLEADO(ESTADO: INACTIVO)"
        End If
        
    End If
    

End Sub

Sub insertar_documento(l_ternro, l_nrodoc, l_tidnro)
    Dim rs_sub As New ADODB.Recordset
    Flog.writeline Espacios(Tabulador * 1) & "  -- DOCUMENTOS -- "
    nrodoc = Replace(nrodoc, ".", "") 'elimino puntos y comas
    nrodoc = Replace(nrodoc, ",", "")
    
    If nrodoc <> "" Then
        StrSql = " SELECT * FROM ter_doc WHERE ternro = " & l_ternro & " AND tidnro = " & l_tidnro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
            StrSql = StrSql & " VALUES(" & l_ternro & ", " & l_tidnro & ",'" & l_nrodoc & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al insertar el CUIL - ternro:" & l_ternro & " - nrodoc:" & l_nrodoc & " - tidnro:" & l_tidnro
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Inserto el Documento CUIL - ternro:" & l_ternro & " - nrodoc:" & l_nrodoc & " - tidnro:" & l_tidnro
            End If
        Else
            StrSql = " UPDATE ter_doc SET nrodoc = '" & l_nrodoc & "'"
            StrSql = StrSql & " WHERE ternro = " & l_ternro & " AND tidnro = " & l_tidnro
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZO el Documento CUIL - ternro:" & l_ternro & " - nrodoc:" & l_nrodoc & " - tidnro:" & l_tidnro
        End If
        rs_sub.Close
        Set rs_sub = Nothing
    Else
        Flog.writeline Espacios(Tabulador * 1) & "ERROR: Documento Vacio."
    End If
    
End Sub
        
Function insertar_domicilio(l_ternro, l_calle, l_callenro, l_piso, l_depto, l_locnro, l_provnro, l_codigopostal, l_paisnro)
    Dim rs_sub As New ADODB.Recordset
    Dim l_NroDom
    Flog.writeline Espacios(Tabulador * 1) & " -- DOMICILIOS  --"
    
    StrSql = " SELECT * FROM cabdom WHERE ternro = " & l_ternro & " AND tipnro = 1 AND domdefault = -1 AND tidonro = 2 AND modnro = 1"
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro la cabecera del domicilio, se creara uno nuevo."
        
        StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro, modnro) "
        StrSql = StrSql & " VALUES(1," & l_ternro & ",-1,2,1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        If Err Then
            Flog.writeline Espacios(Tabulador * 1) & "Error al insertar la cabecera del Domicilio."
            Err.Clear
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Inserto la cabecera del Domicilio."
        End If
        
        '--Obtengo el numero de domicilio en la tabla--
        l_NroDom = getLastIdentity(objConn, "cabdom")
        Flog.writeline Espacios(Tabulador * 1) & "Obtengo la cabecera del domicilio ( " & l_NroDom & ")."
        
    Else
        l_NroDom = rs_sub("domnro")
        Flog.writeline Espacios(Tabulador * 1) & "Obtengo la cabecera del domicilio ( " & l_NroDom & ")."
    End If


    '--Si mo tiene algun dato le agregamos unos ficticios--
    
    If l_locnro = "" Then l_locnro = "1" 'no informada
    If l_provnro = "" Then l_provnro = "1" 'no informada
    If l_paisnro = "" Then l_paisnro = "1" 'no informada
    
    Err.Clear
    
    If l_calle = "" Then
        Flog.writeline Espacios(Tabulador * 1) & "No se genera/actualiza el domicilio, la calle no puede ser vacio."
    Else
    
        StrSql = "SELECT * FROM detdom WHERE domnro = " & l_NroDom
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
    
            StrSql = " INSERT INTO detdom (domnro,calle,nro,piso,oficdepto,codigopostal,"
            StrSql = StrSql & "locnro,provnro,paisnro) "
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & l_NroDom
            StrSql = StrSql & ",'" & CStr(l_calle) & "'"
            StrSql = StrSql & ",'" & CStr(l_callenro) & "'"
            StrSql = StrSql & ",'" & CStr(l_piso) & "'"
            StrSql = StrSql & ",'" & CStr(l_depto) & "'"
            StrSql = StrSql & ",'" & CStr(l_codigopostal) & "'"
            StrSql = StrSql & "," & l_locnro
            StrSql = StrSql & "," & l_provnro
            StrSql = StrSql & "," & l_paisnro
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al insertar el detalle Domicilio."
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Inserto el detalle Domicilio."
            End If
        Else
            StrSql = " UPDATE detdom SET "
            StrSql = StrSql & "calle = '" & CStr(l_calle) & "'"
            StrSql = StrSql & ",nro = '" & CStr(l_callenro) & "'"
            StrSql = StrSql & ",piso = '" & CStr(l_piso) & "'"
            StrSql = StrSql & ",oficdepto = '" & CStr(l_depto) & "'"
            StrSql = StrSql & ",codigopostal = '" & CStr(l_codigopostal) & "'"
            StrSql = StrSql & ",locnro = " & CInt(l_locnro)
            StrSql = StrSql & ",provnro = " & l_provnro
            StrSql = StrSql & ",paisnro = " & l_paisnro
            StrSql = StrSql & " WHERE domnro = " & l_NroDom
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline Espacios(Tabulador * 1) & "Error al Actualizar el detalle Domicilio."
                Err.Clear
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Actualizo el detalle Domicilio."
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Detalle Domiciliol: " & l_calle & " - " & l_callenro & " - " & l_piso & " - " & l_depto & " - " & l_locnro & " - " & l_provnro & " - " & l_codigopostal & " - " & l_paisnro
        End If
    
    End If
    
    insertar_domicilio = l_NroDom
    
    rs_sub.Close
    Set rs_sub = Nothing

End Function

Sub insertar_telefono(l_NroDom, l_telnro)
    Dim rs_sub As New ADODB.Recordset
    Flog.writeline Espacios(Tabulador * 1) & " -- TELEFONOS -- "
    If Trim(l_telnro) <> "" And l_NroDom <> 0 Then
        StrSql = " SELECT * from telefono where domnro = " & l_NroDom & " AND telnro = '" & l_telnro & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
             StrSql = " INSERT INTO telefono "
             StrSql = StrSql & " (domnro, telnro, telfax, teldefault, telcelular, tipotel ) "
             StrSql = StrSql & " VALUES (" & l_NroDom & ",'" & Left(l_telnro, 20) & "',0,-1,0,1 ) "
             objConn.Execute StrSql, , adExecuteNoRecords
        Else
             'StrSql = " UPDATE telefono SET telnro = '" & Left(l_telnro, 20) & "'"
             'StrSql = StrSql & " where domnro = " & l_NroDom & " AND telnro = '" & l_telnro & "'"
             'objConn.Execute StrSql, , adExecuteNoRecords
        End If
        If Err Then
            Flog.writeline Espacios(Tabulador * 1) & "Error al insertar el Telefono "
            HuboErrorLocal = True
            Err.Clear
        Else
            Flog.writeline Espacios(Tabulador * 1) & " Inserto el telefono "
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Telefono:" & l_telnro & " - domnro: " & l_NroDom
    End If
    rs_sub.Close
    Set rs_sub = Nothing
End Sub

Function buscar_tercero(l_legajo) As String
    Dim rs_sub As New ADODB.Recordset
    
    StrSql = " SELECT ternro FROM empleado WHERE empleg = " & l_legajo
    OpenRecordset StrSql, rs_sub
    If Not rs_sub.EOF Then
        buscar_tercero = CStr(rs_sub("ternro"))
        Flog.writeline Espacios(Tabulador * 1) & "Se encontro el tercero: " & rs_sub("ternro") & "  - Legajo:" & l_legajo
    Else
        buscar_tercero = "0"
        Flog.writeline Espacios(Tabulador * 1) & "Error al encontrar el tercero - Legajo:" & l_legajo
    End If
    
    rs_sub.Close
    Set rs_sub = Nothing
End Function

Function a_fecha(Fecha As String) As String
    If Fecha <> "" Then
        a_fecha = Mid(Fecha, 1, 2)
        a_fecha = a_fecha & "/" & Mid(Fecha, 3, 2)
        a_fecha = a_fecha & "/" & Mid(Fecha, 5, 4)
        If IsDate(a_fecha) Then
            a_fecha = a_fecha
        Else
            a_fecha = ""
        End If
    End If
End Function

Sub insertar_Estructura(l_ternro, l_tenro, l_estrnro, l_fecha, l_tipmotnro)
    Dim rs_sub As New ADODB.Recordset
    Dim l_htethasta As String
    Flog.writeline Espacios(Tabulador * 0) & " -- ESTRUCTURAS -- "
    Flog.writeline Espacios(Tabulador * 1) & "GENERAR ESTRUCTURA: tenro: " & l_tenro & " - estrnro: " & l_estrnro & " - fecha:" & l_fecha & " - Tipo Motivo:" & l_tipmotnro
    
    If l_estrnro = "" Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR estructura EN BLANCO. No se crea la estructura."
        Exit Sub
    End If
    
    If l_fecha = "" Then
        l_fecha = gethtetdesde(l_ternro, l_tenro)
    End If
    '_______________________________
    'VALIDO QUE EXISTA LA ESTRUCTURA
    StrSql = "SELECT estrnro FROM estructura WHERE estrnro = " & l_estrnro
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        Texto = "ERROR estructura Número: " & l_estrnro & " No existe. Verificar Mapeo"
        HuboErrorLocal = True
        Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
        rs_sub.Close
        Exit Sub
    End If
    rs_sub.Close
    '------------------------------
    
    l_htethasta = DateAdd("d", -1, l_fecha)
    
    If l_tipmotnro = "" Then l_tipmotnro = "0"
    
    StrSql = " SELECT * FROM his_estructura WHERE ternro = " & l_ternro
    StrSql = StrSql & " AND tenro = " & l_tenro
    StrSql = StrSql & " AND estrnro = " & l_estrnro
    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(l_fecha) & " AND (htethasta >= " & ConvFecha(l_fecha) & " OR htethasta is null)) "
    'StrSql = StrSql & " AND (( htetdesde <= " & ConvFecha(l_fecha) & " AND (htethasta >= " & ConvFecha(l_fecha) & " OR htethasta is null)) "
    'StrSql = StrSql & " OR   ( htetdesde >= " & ConvFecha(l_fecha) & " )) "
    OpenRecordset StrSql, rs_sub
    If Not rs_sub.EOF Then
        Texto = "El empleado posee la estructura activa a la fecha - ternro: " & l_ternro & " - estrnro: " & l_estrnro & " - Fecha : " & l_fecha
        Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
        rs_sub.Close
        Exit Sub
    Else

    End If
    rs_sub.Close
    
    StrSql = " SELECT * FROM his_estructura WHERE ternro = " & l_ternro
    StrSql = StrSql & " AND tenro = " & l_tenro
    StrSql = StrSql & " AND htetdesde = " & ConvFecha(l_fecha)
    OpenRecordset StrSql, rs_sub
    If Not rs_sub.EOF Then
    
    
        StrSql = " UPDATE his_estructura SET estrnro = " & l_estrnro
        StrSql = StrSql & ",tipmotnro = " & l_tipmotnro
        StrSql = StrSql & " WHERE tenro = " & rs_sub!Tenro
        StrSql = StrSql & " AND estrnro = " & rs_sub!Estrnro
        StrSql = StrSql & " AND ternro = " & l_ternro
        StrSql = StrSql & " AND htetdesde = " & ConvFecha(l_fecha)
        objConn.Execute StrSql, , adExecuteNoRecords
        Texto = "ACTUALIZO la estructura - estrnro: " & rs_sub!Estrnro & " - a la estructura: " & l_estrnro & " - ternro: " & rs_sub!Ternro & " - hasta " & l_htethasta
        Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
        Exit Sub
    End If
    
    'Borro si hay alguna con fecha desde mayor
    StrSql = " SELECT * FROM his_estructura WHERE ternro = " & l_ternro
    StrSql = StrSql & " AND tenro = " & l_tenro
    StrSql = StrSql & " AND htetdesde > " & ConvFecha(l_fecha)
    OpenRecordset StrSql, rs_sub
    If Not rs_sub.EOF Then
        Do While Not rs_sub.EOF
            StrSql = " DELETE FROM his_estructura "
            StrSql = StrSql & " WHERE tenro = " & rs_sub!Tenro
            StrSql = StrSql & " AND estrnro = " & rs_sub!Estrnro
            StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_sub!htetdesde)
            StrSql = StrSql & " AND ternro = " & l_ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            Texto = "BORRO la estructura - ternro: " & l_ternro & " - ternro: " & rs_sub!Ternro & " - hasta " & l_htethasta
            Texto = Texto & " - La estructura posee una fecha mayor a la fecha de la estructura que estoy creando."
            Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
            rs_sub.MoveNext
        Loop
    End If
    
    
    StrSql = " SELECT * FROM his_estructura WHERE ternro = " & l_ternro
    StrSql = StrSql & " AND tenro = " & l_tenro
    'StrSql = StrSql & " AND estrnro <> " & l_estrnro
    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(l_fecha) & " AND (htethasta >= " & ConvFecha(l_fecha) & " OR htethasta is null)) "
    'StrSql = StrSql & " AND (( htetdesde <= " & ConvFecha(l_fecha) & " AND (htethasta >= " & ConvFecha(l_fecha) & " OR htethasta is null)) "
    'StrSql = StrSql & " OR   ( htetdesde >= " & ConvFecha(l_fecha) & " )) "
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        
        StrSql = " INSERT INTO his_estructura(tenro, ternro, estrnro, htetdesde, htethasta, tipmotnro)"
        StrSql = StrSql & " VALUES( " & l_tenro & "," & l_ternro & "," & l_estrnro & "," & ConvFecha(l_fecha) & ", null, " & l_tipmotnro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Texto = "Inserto el historico estructura - ternro: " & l_ternro & " - estrnro: " & l_estrnro & " - Fecha desde: " & l_fecha
        Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
    
    Else
        If rs_sub("estrnro") <> l_estrnro Then
            
            'Cierro la anterior
            StrSql = " UPDATE his_estructura SET htethasta = " & ConvFecha(l_htethasta)
            StrSql = StrSql & ",tipmotnro = " & l_tipmotnro
            StrSql = StrSql & " WHERE tenro = " & rs_sub!Tenro
            StrSql = StrSql & " AND estrnro = " & rs_sub!Estrnro
            StrSql = StrSql & " AND ternro = " & l_ternro
            StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_sub!htetdesde)
            objConn.Execute StrSql, , adExecuteNoRecords
            Texto = "CIERRO la estructura - estrnro: " & rs_sub!Estrnro & " - ternro: " & rs_sub!Ternro & " - hasta " & l_htethasta
            Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
            
            'Creo la nueva
            StrSql = " INSERT INTO his_estructura(tenro, ternro, estrnro, htetdesde, htethasta, tipmotnro)"
            StrSql = StrSql & " VALUES( " & l_tenro & "," & l_ternro & "," & l_estrnro & "," & ConvFecha(l_fecha) & ",null, " & l_tipmotnro & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Texto = "Inserto el historico estructura - ternro: " & l_ternro & " - estrnro: " & l_estrnro & " - Fecha desde: " & l_fecha
            Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
        Else
            'Es la misma, no hago anda
        End If
    End If
    rs_sub.Close
        
    Set rs_sub = Nothing
End Sub

Function gethtetdesde(l_ternro, l_tenro) As String
    Dim rs_sub As New ADODB.Recordset
    
    StrSql = " SELECT htetdesde, tenro FROM his_estructura WHERE ternro = " & l_ternro '& " AND tenro = " & l_tenro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Date)
    StrSql = StrSql & " ORDER BY htetdesde desc "
    OpenRecordset StrSql, rs_sub
    If Not rs_sub.EOF Then
        gethtetdesde = rs_sub("htetdesde")
        'Flog.writeline Espacios(Tabulador * 1) &  "Se encontro el historico de estructura para el tipo de estructura: " & l_tenro & " la fecha : " & rs_sub("htetdesde") & "  - ternro:" & l_ternro
        Flog.writeline Espacios(Tabulador * 1) & "Se encontro el historico de estructura para la fecha : " & rs_sub("htetdesde") & "  - ternro:" & l_ternro
    Else
        rs_sub.Close
        StrSql = " SELECT htetdesde, tenro FROM his_estructura WHERE ternro = " & l_ternro '& " AND htetdesde <= " & ConvFecha(Date)
        StrSql = " ORDER BY htetdesde desc "
        OpenRecordset StrSql, rs_sub
        If Not rs_sub.EOF Then
            gethtetdesde = rs_sub("htetdesde")
            Flog.writeline Espacios(Tabulador * 1) & "Se encontro el historico de estructura para el tipo de estructura: " & rs_sub("tenro") & " la fecha : " & rs_sub("htetdesde") & "  - ternro:" & l_ternro
        Else
            gethtetdesde = Date
            Flog.writeline Espacios(Tabulador * 1) & "No se encontro historico de estructura: " & l_tenro & " - ternro:" & l_ternro
            Flog.writeline Espacios(Tabulador * 1) & "    Asumo la actual: " & Date
        End If
    End If
    
    rs_sub.Close
    Set rs_sub = Nothing

End Function

Public Sub Insertar_Novedad(l_ternro As String, l_conccod As String, l_tpanro As Long, l_Monto As String, l_Fecha_Desde As String, l_Fecha_Hasta As String, Texto As String)
    Flog.writeline Espacios(Tabulador * 1) & " -- NOVEDADES --"
    Dim ConcNro As String
    'Dim tpanro As Long
    Dim Valor As Single
    Dim FechaDesde As Date
    Dim FechaHasta As Date
    'FechaHasta = Date
    
    Dim rs_sub As New ADODB.Recordset

    StrSql = "SELECT * FROM concepto "
    StrSql = StrSql & " WHERE conccod = '" & l_conccod & "'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        ConcNro = rs!ConcNro
    Else
        ConcNro = "0"
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro el CONCEPTO: conccod:" & l_conccod & " - No se puede continuar "
        Exit Sub
    End If

    If l_Fecha_Desde = "" Then
        l_Fecha_Desde = gethtetdesde(l_ternro, "0") 'busco la fecha para cualquier estructura.
    End If
    FechaDesde = CDate(l_Fecha_Desde)
    If Not EsNulo(l_Fecha_Hasta) Then
        FechaHasta = CDate(l_Fecha_Hasta)
    End If

    'Inserta la novedad para el monto
    If ConcNro <> "0" And l_tpanro <> "0" Then
        Valor = Monto
        
        StrSql = "SELECT * FROM novemp WHERE "
        StrSql = StrSql & " concnro = " & ConcNro
        StrSql = StrSql & " AND tpanro = " & l_tpanro
        StrSql = StrSql & " AND empleado = " & l_ternro
        StrSql = StrSql & " AND (nevigencia = 0 "
        StrSql = StrSql & " OR (nevigencia = -1 "
        If Not EsNulo(FechaHasta) Then
            StrSql = StrSql & " AND (nedesde <= " & ConvFecha(FechaHasta)
            StrSql = StrSql & " AND nehasta >= " & ConvFecha(FechaDesde) & ")"
            StrSql = StrSql & " OR  (nedesde <= " & ConvFecha(FechaHasta)
            StrSql = StrSql & " AND nehasta is null )))"
        Else
            StrSql = StrSql & " AND nehasta is null OR nehasta >= " & ConvFecha(FechaDesde) & "))"
        End If
        If rs_sub.State = adStateOpen Then rs_sub.Close
        OpenRecordset StrSql, rs_sub
    
        If Not rs_sub.EOF Then
            'A lo sumo va a actualizar una sola
            Do While Not rs_sub.EOF
                If Not CBool(rs_sub!nevigencia) Then
                    'Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
                    Flog.writeline Espacios(Tabulador * 0) & "Linea " & NroLinea & ": No se puede insertar la novedad porque ya existe una sin vigencia"
                Else
                    If rs_sub!nedesde = FechaDesde Then
                        If EsNulo(rs_sub!neHasta) Then
                        
                        Else
                            'ya la tengo ==> actualizo el monto y la fecha hasta
                            If Not EsNulo(rs_sub!neHasta) Then
                                If Not EsNulo(FechaHasta) Then
                                    If FechaHasta < rs_sub!neHasta Then
                                        StrSql = "UPDATE novemp SET "
                                        StrSql = StrSql & " nehasta = " & ConvFecha(FechaHasta)
                                        StrSql = StrSql & " ,nevalor = " & Valor
                                        StrSql = StrSql & " WHERE nenro = " & rs_sub!nenro
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If rs_sub!nedesde < FechaDesde Then
                            'es la nueva vigencia, actualizo la anterior e Inserto la nueva
                            StrSql = "UPDATE novemp SET "
                            StrSql = StrSql & " nehasta = " & ConvFecha(FechaDesde - 1)
                            StrSql = StrSql & " WHERE nenro = " & rs_sub!nenro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        
                            'es la nueva vigencia, Inserto
                            StrSql = "INSERT INTO novemp ("
                            StrSql = StrSql & "empleado,concnro,tpanro,nevalor,nevigencia,nedesde"
                            If Not EsNulo(FechaHasta) Then
                                StrSql = StrSql & ",nehasta"
                            End If
                            StrSql = StrSql & " ,netexto"
                            StrSql = StrSql & ") VALUES (" & l_ternro
                            StrSql = StrSql & "," & ConcNro
                            StrSql = StrSql & "," & l_tpanro
                            StrSql = StrSql & "," & Valor
                            StrSql = StrSql & ",-1"
                            StrSql = StrSql & "," & ConvFecha(FechaDesde)
                            If Not EsNulo(FechaHasta) Then
                                StrSql = StrSql & "," & ConvFecha(FechaHasta)
                            End If
                            StrSql = StrSql & ",'" & Texto & "'"
                            StrSql = StrSql & " )"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
                rs_sub.MoveNext
            Loop
        Else
                StrSql = "INSERT INTO novemp ("
                StrSql = StrSql & "empleado,concnro,tpanro,nevalor,nevigencia,nedesde"
                If Not EsNulo(FechaHasta) Then
                    StrSql = StrSql & ",nehasta"
                End If
                StrSql = StrSql & " ,netexto"
                
                StrSql = StrSql & ") VALUES (" & l_ternro
                StrSql = StrSql & "," & ConcNro
                StrSql = StrSql & "," & l_tpanro
                StrSql = StrSql & "," & Valor
                StrSql = StrSql & ",-1"
                StrSql = StrSql & "," & ConvFecha(FechaDesde)
                If Not EsNulo(FechaHasta) Then
                    StrSql = StrSql & "," & ConvFecha(FechaHasta)
                End If
                StrSql = StrSql & ",'" & Texto & "'"
                StrSql = StrSql & " )"
                objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'Call InsertarLogNovedad("", Concnro, tpanro, EsMonto, Valor, Fecha_Desde, Fecha_Hasta, Texto)
        Flog.writeline Espacios(Tabulador * 0) & "Linea " & NroLinea & ": Inserto la novedad - ternro: " & l_ternro & " - concnro:" & ConcNro & " - Monto:" & Valor & " - F desde:" & FechaDesde
    End If
    
    'cierro y libero
    If rs_sub.State = adStateOpen Then rs_sub.Close
    Set rs_sub = Nothing
End Sub

Function esFueraDeConvenio(l_ternro, l_convnro, l_fecha, l_tenro)
    Dim rs_sub As New ADODB.Recordset
    Dim l_htethasta As String
    Flog.writeline Espacios(Tabulador * 1) & " -- CONVENIO -- "
    Flog.writeline Espacios(Tabulador * 0) & "Buscando Convenio: tenro: 19 - tenro: " & l_tenro & " - fecha:" & l_fecha
    
    If l_fecha = "" Then
        l_fecha = gethtetdesde(l_ternro, l_tenro)
    End If
    
    l_htethasta = DateAdd("d", -1, l_fecha)
        
    StrSql = " SELECT * FROM his_estructura WHERE ternro = " & l_ternro
    StrSql = StrSql & " AND tenro = " & l_tenro
    'StrSql = StrSql & " AND estrnro = " & l_convnro
    StrSql = StrSql & " AND ( htetdesde <= " & ConvFecha(l_fecha) & " AND (htethasta >= " & ConvFecha(l_fecha) & " OR htethasta is null)) "
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        'no tiene la estructura convenio
        Texto = "El empleado no posee la estructura CONVENIO - ternro: " & l_ternro & " - tenro: " & l_tenro & " - Fecha desde: " & l_fecha
        Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
        esFueraDeConvenio = False
    Else
        'Esta en un convenio
        If CStr(rs_sub("estrnro")) = CStr(l_convnro) Then
            'esta fuera de convenio
            Texto = "El empleado esta FUERA de CONVENIO - ternro: " & l_ternro & " - tenro: " & l_tenro & " - estrnro: " & CStr(rs_sub("estrnro")) & " - Fecha desde: " & l_fecha
            Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
            esFueraDeConvenio = True
        Else
            Texto = "El empleado esta EN CONVENIO - ternro: " & l_ternro & " - tenro: " & l_tenro & " - estrnro: " & CStr(rs_sub("estrnro")) & " - Fecha desde: " & l_fecha
            Call Escribir_Log("flog", NroLinea, 1, Texto, 1, "")
            'esta en algun convio
            esFueraDeConvenio = False
        End If
    End If
    rs_sub.Close
    Set rs_sub = Nothing
End Function
