Attribute VB_Name = "mdlImportacion"
Public Sub import_modelo1000(ByVal strLinea As String, ByVal destino As Long, ByRef str_error As String)
'Datos del empleado
Dim arrayLinea
Dim indice As Long
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Datos As New ADODB.Recordset
Dim ternro As Long
Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim Nombre2 As String
Dim fecnac As String
Dim pais As String
Dim Nacionalidad As String
Dim FecIng As String
Dim EstCiv As String
Dim Sexo As String
Dim fecAlta As String
Dim Estudia As String
Dim NivEstudio As String
Dim Email As String
Dim FecVtoCont As String
Dim estado As String
Dim Remuneracion As String
Dim ModOrg As String
Dim ReportaA As String
Dim huboError As Boolean
Dim separador As String

    separador = SeparadorModelo(1000)
    arrayLinea = Split(strLinea, separador)
    Flog.writeline Espacios(Tabulador * 0) & "Comienzo de modelo 1000 - Empleados."
    Flog.writeline Espacios(Tabulador * 0) & "Busco el empleado: " & arrayLinea(1)
    'Veo si el empleado existe
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & arrayLinea(1)
    Legajo = arrayLinea(1)
    OpenRecordset StrSql, rs_Empleado
    If Not rs_Empleado.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " existe en el sistema."
        ternro = rs_Empleado!ternro
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " No existe en el sistema, se creara."
        ternro = 0
    End If
    huboError = False
    str_error = str_error & "<TABLE width='50%' border='1' bordercolor='#333333' cellpadding='0' cellspacing='0'>" & vbCrLf
    str_error = str_error & "<tr><TH style='background-color:#d13528;color:#FFFFFF;' colspan='2' align='center'><b>Modelo 1000, para el legajo: " & arrayLinea(1) & "</TH></tr>" & vbCrLf

    'Recorro los datos de la linea
    For indice = 2 To UBound(arrayLinea)
        
        Select Case indice
             '------------------------------------------------------------------------------------------
             Case 2: 'Apellidos
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If InStr(arrayLinea(indice), "@") = 0 Then
                        apellido = arrayLinea(indice)
                    Else
                        apellido = Split(arrayLinea(indice), "@")(0)
                        apellido2 = Split(arrayLinea(indice), "@")(1)
                    End If
                    Flog.writeline Espacios(Tabulador * 1) & "Apellidos obtenidos."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Apellidos Obligatorios."
                    str_error = str_error & "<tr><td>Apellidos invalidos</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 3: 'Nombres
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If InStr(arrayLinea(indice), "@") = 0 Then
                        nombre = arrayLinea(indice)
                    Else
                        nombre = Split(arrayLinea(indice), "@")(0)
                        Nombre2 = Split(arrayLinea(indice), "@")(1)
                    End If
                    Flog.writeline Espacios(Tabulador * 1) & "Nombres obtenidos."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Nombres Obligatorios."
                    str_error = str_error & "<tr><td>Nombres invalidos</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 4: 'Fecha de nacimiento
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    fecnac = arrayLinea(indice)
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Fecha Nacimiento Obligatoria."
                    str_error = str_error & "<tr><td>Fecha de nacimiento invalidos</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 5: 'Pais
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = "SELECT paisnro FROM pais WHERE upper(paisdesc) = '" & UCase(Mid(arrayLinea(indice), 1, 60)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        pais = rs_Datos!paisnro
                        Flog.writeline Espacios(Tabulador * 1) & "Pais encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Pais no encontrado, se creara."
                        StrSql = " INSERT INTO pais (paisdesc,paisdef) " & _
                                 " VALUES ('" & UCase(Mid(arrayLinea(indice), 1, 60)) & "',0)"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        pais = getLastIdentity(objConn, "pais")
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Pais Obligatorio."
                    str_error = str_error & "<tr><td>Pais invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 6: 'Nacionalidad
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = "SELECT nacionalnro FROM nacionalidad WHERE upper(nacionaldes) = '" & UCase(Mid(arrayLinea(indice), 1, 30)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        Nacionalidad = rs_Datos!nacionalnro
                        Flog.writeline Espacios(Tabulador * 1) & "Nacionalidad encontrada."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Nacionalidad no encontrada, se creara."
                        StrSql = " INSERT INTO nacionalidad (nacionaldes,nacionaldefault) " & _
                                 " VALUES ('" & UCase(Mid(arrayLinea(indice), 1, 30)) & "',0)"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        Nacionalidad = getLastIdentity(objConn, "nacionalidad")
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Nacionalidad Obligatoria."
                    str_error = str_error & "<tr><td>Nacionalidad invalida</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 7: 'Fecha de Ingreso
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    FecIng = arrayLinea(indice)
                Else
                    FecIng = ""
                    'str_error = str_error & "<tr><td>Fecha de Ingreso invalido</td></tr>" & vbCrLf
                    'huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 8: 'Estado Civil
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = "SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(Mid(arrayLinea(indice), 1, 30)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        EstCiv = rs_Datos!estcivnro
                        Flog.writeline Espacios(Tabulador * 1) & "Estado Civil encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Estado Civil no encontrado, se creara."
                        StrSql = " INSERT INTO estcivil (estcivdesabr) " & _
                                 " VALUES ('" & UCase(Mid(arrayLinea(indice), 1, 30)) & "')"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        EstCiv = getLastIdentity(objConn, "estcivil")
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Estado Civil Obligatorio."
                    str_error = str_error & "<tr><td>Estado civil invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 9: 'Sexo
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If UCase(arrayLinea(indice)) = "M" Then
                        Sexo = -1
                    Else
                        Sexo = 0
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Sexo Obligatorio."
                    str_error = str_error & "<tr><td>Sexo invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            
            '------------------------------------------------------------------------------------------
             Case 10: 'Fecha de Alta tabla empleado
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    fecAlta = arrayLinea(indice)
                Else
                    fecAlta = ""
                    'str_error = str_error & "<tr><td>Fecha de Alta del empleado invalida</td></tr>" & vbCrLf
                    'huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 11: 'Estudia
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If UCase(arrayLinea(indice)) = "SI" Then
                        Estudia = -1
                    Else
                        Estudia = 0
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Estudia Obligatorio."
                    str_error = str_error & "<tr><td>Campo Estudia invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 12: 'Nivel de estudio
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = "SELECT nivnro FROM nivest WHERE upper(nivdesc) = '" & UCase(Mid(arrayLinea(indice), 1, 40)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        NivEstudio = rs_Datos!nivnro
                        Flog.writeline Espacios(Tabulador * 1) & "Nivel de Estudio encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Nivel de Estudio no encontrado, se creara."
                        StrSql = " INSERT INTO nivest (nivdesc,nivsist,nivestfli) " & _
                                 " VALUES ('" & UCase(Mid(arrayLinea(indice), 1, 40)) & "',0,0)"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        NivEstudio = getLastIdentity(objConn, "nivest")
                    End If

                Else
                    If Estudia = -1 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Error: Nivel de Estudio obligatorio."
                        str_error = str_error & "<tr><td>Campo Estudia invalido</td></tr>" & vbCrLf
                        huboError = True
                    Else
                        NivEstudio = "null"
                    End If
                End If
            
            '------------------------------------------------------------------------------------------
            Case 13: 'Email
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    Email = arrayLinea(indice)
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Email Obligatorio."
                    str_error = str_error & "<tr><td>Email invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
            Case 14: 'Fecha vencimiento contrato
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    FecVtoCont = arrayLinea(indice)
                Else
                    FecVtoCont = ""
                End If
            '------------------------------------------------------------------------------------------
            Case 15: 'Estado
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If UCase(arrayLinea(indice)) = "ACTIVO" Then
                        estado = -1
                    Else
                        estado = 0
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Estado del empleado Obligatorio."
                    str_error = str_error & "<tr><td>Estado del empleado invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 16: 'Remuneracion
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    Remuneracion = arrayLinea(indice)
                Else
                    Remuneracion = 0
                End If
            '------------------------------------------------------------------------------------------
             Case 17: 'Modelo de Organizacion
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = "SELECT tplatenro FROM adptemplate WHERE upper(tplatedesabr) = '" & UCase(Mid(arrayLinea(indice), 1, 30)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        ModOrg = rs_Datos!tplatenro
                        Flog.writeline Espacios(Tabulador * 1) & "Modelo de Organizacion encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Modelo de Organizacion no encontrado, se creara."
                        StrSql = " INSERT INTO adptemplate (tplatedesabr,tplatedefault) " & _
                                 " VALUES ('" & UCase(Mid(arrayLinea(indice), 1, 30)) & "',0)"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        ModOrg = getLastIdentity(objConn, "adptemplate")
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Modelo de Organizacion Obligatorio."
                    str_error = str_error & "<tr><td>Modelo de Organizacion invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 18: 'Reporta A
                ReportaA = 0
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If UCase(arrayLinea(indice)) <> "N/A" Then
                        StrSql = "SELECT ternro FROM empleado WHERE empleg = " & arrayLinea(indice)
                        OpenRecordset StrSql, rs_Datos
                        If Not rs_Datos.EOF Then
                            ReportaA = rs_Datos!ternro
                            Flog.writeline Espacios(Tabulador * 1) & "Reporta A encontrado."
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Reporta A no existe."
                            str_error = str_error & "<tr><td>Reporta A invalido</td></tr>" & vbCrLf
                            huboError = True
                        End If
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Reporta A Obligatorio."
                    str_error = str_error & "<tr><td>Reporta A invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
        End Select
    Next
        
    'Si el empleado no existe lo tengo que crear
    If Not huboError Then
        str_error = str_error & "<tr><td colspan='2' align='center'><b>No hay Errores</b></td></tr>" & vbCrLf
        If ternro = 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "El empleado no existe."
            StrSql = " INSERT INTO tercero(ternom,ternom2,terape,terape2,terfecnac,tersex,estcivnro,terfecing,nacionalnro,paisnro)"
            StrSql = StrSql & " VALUES('" & nombre & "','" & Nombre2 & "','" & apellido & "','" & apellido2 & "'," & ConvFecha(fecnac) & "," & Sexo & "," & EstCiv & ","
            If UCase(FecIng) <> "N/A" And UCase(FecIng) <> "" Then
                StrSql = StrSql & ConvFecha(FecIng) & ","
            Else
                StrSql = StrSql & "null,"
            End If
            StrSql = StrSql & Nacionalidad & ","
            StrSql = StrSql & pais & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Flog.writeline Espacios(Tabulador * 1) & "Inserto en la tabla tercero."
            ternro = getLastIdentity(objConn, "tercero")
            Flog.writeline Espacios(Tabulador * 1) & "Nuevo numero de tercero: " & ternro & "."
            
            'inserto el ter_tip correspondiente a empleado (1)
            StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & ternro & ",1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Inserto el empleado
            StrSql = " INSERT INTO empleado(empleg,empfecalta,empest,empfbajaprev,"
            StrSql = StrSql & "ternro,nivnro,empestudia,terape,terape2,ternom,ternom2,empemail, "
            StrSql = StrSql & "tplatenro,empremu"
            If ReportaA <> 0 Then
                StrSql = StrSql & ",empreporta"
            End If
           
            StrSql = StrSql & ") VALUES("
            StrSql = StrSql & Legajo & ","
            
            If UCase(fecAlta) <> "N/A" And UCase(fecAlta) <> "" Then
                StrSql = StrSql & ConvFecha(fecAlta) & ","
            Else
                StrSql = StrSql & " null,"
            End If
            
            StrSql = StrSql & estado & ","
            
            If UCase(FecVtoCont) <> "N/A" And UCase(FecVtoCont) <> "" Then
                StrSql = StrSql & ConvFecha(FecVtoCont) & ","
            Else
                StrSql = StrSql & " null,"
            End If
            
            StrSql = StrSql & ternro & "," & NivEstudio & "," & Estudia & ",'" & apellido & "','" & apellido2 & "','"
            StrSql = StrSql & nombre & "','" & Nombre2 & "','" & Email & "'," & ModOrg & "," & Remuneracion
            If ReportaA <> 0 Then
                StrSql = StrSql & "," & ReportaA
            End If
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Inserto en la tabla empleado."
            Else
                'empleado existe, actualizo tercero
                Flog.writeline Espacios(Tabulador * 1) & "El empleado existe."
                StrSql = " UPDATE tercero SET " & _
                        " ternom = '" & nombre & "'," & _
                        " ternom2 = '" & Nombre2 & "'," & _
                        " terape = '" & apellido & "'," & _
                        " terape2 = '" & apellido2 & "'," & _
                        " terfecnac = " & ConvFecha(fecnac) & "," & _
                        " tersex = " & Sexo & "," & _
                        " estcivnro = " & EstCiv
                If UCase(FecIng) <> "N/A" And UCase(FecIng) <> "" Then
                    StrSql = StrSql & " ,terfecing = " & ConvFecha(FecIng)
                Else
                    StrSql = StrSql & " ,terfecing = null "
                End If
                StrSql = StrSql & ", nacionalnro = " & Nacionalidad & "," & _
                        " paisnro = " & pais & _
                        " WHERE ternro = " & ternro
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Actualizada la tabla tercero."
                
                'empleado existe, actualizo
                StrSql = " UPDATE empleado SET " & _
                         " empleg = " & Legajo & ","

                If UCase(fecAlta) <> "N/A" And UCase(fecAlta) <> "" Then
                    StrSql = StrSql & " empfecalta = " & ConvFecha(fecAlta) & ","
                Else
                    StrSql = StrSql & " empfecalta = null, "
                End If
                
                StrSql = StrSql & "empest = " & estado
                
                If UCase(FecVtoCont) <> "N/A" Then
                    StrSql = StrSql & " ,empfbajaprev = " & ConvFecha(FecVtoCont) & ","
                Else
                    StrSql = StrSql & " ,empfbajaprev = null ,"
                End If
                StrSql = StrSql & " nivnro = " & NivEstudio & "," & _
                         " empestudia = " & Estudia & "," & _
                         " terape = '" & apellido & "'," & _
                         " terape2 = '" & apellido2 & "'," & _
                         " ternom = '" & nombre & "'," & _
                         " ternom2 = '" & Nombre2 & "'," & _
                         " empemail = '" & Email & "'," & _
                         " tplatenro = " & ModOrg & "," & _
                         " empremu = " & Remuneracion
                If ReportaA <> 0 Then
                    StrSql = StrSql & ",empreporta = " & ReportaA
                End If
                StrSql = StrSql & " WHERE ternro = " & ternro
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Actualizada la tabla empleado."
            End If
    End If
    str_error = str_error & "</table>" & vbCrLf
    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "

    Flog.writeline Espacios(Tabulador * 0) & "Fin de modelo 1000 - Empleados."
    
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    If rs_Datos.State = adStateOpen Then rs_Datos.Close
    
End Sub

Public Sub import_modelo1001(ByVal strLinea As String, ByVal destino As Long, ByRef str_error As String)
'Domicilios
Dim arrayLinea
Dim Legajo As Long
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Datos As New ADODB.Recordset
Dim rs_tel As New ADODB.Recordset
Dim pais As String
Dim tipoDomicilio As String
Dim calle As String
Dim Numero As String
Dim piso As String
Dim depto As String
Dim torre As String
Dim manzana As String
Dim cp As String
Dim entreCalles As String
Dim barrio As String
Dim localidad As String
Dim partido As String
Dim zona As String
Dim provincia As String
Dim TelPart As String
Dim TelLab As String
Dim TelCel As String
Dim domnro As Long
Dim modeloDomicilio As String
Dim separador As String
Dim huboError As Boolean


    separador = SeparadorModelo(1001)
    arrayLinea = Split(strLinea, separador)

    Flog.writeline Espacios(Tabulador * 0) & "Comienzo del modelo 1001 - Domicilios."
    Flog.writeline Espacios(Tabulador * 1) & "Busco el empleado: " & arrayLinea(1)
    'Chequeo si el empleado existe
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & arrayLinea(1)
    Legajo = arrayLinea(1)
    OpenRecordset StrSql, rs_Empleado
    If Not rs_Empleado.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " existe en el sistema."
        ternro = rs_Empleado!ternro
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " No existe en el sistema."
        Exit Sub
    End If
        
    huboError = False
    str_error = str_error & "<TABLE width='50%' border='1' bordercolor='#333333' cellpadding='0' cellspacing='0'>" & vbCrLf
    str_error = str_error & "<tr><TH style='background-color:#d13528;color:#FFFFFF;' colspan='2' align='center'><b>Modelo 1001, para el legajo: " & arrayLinea(1) & "</TH></tr>" & vbCrLf

    TelPart = 0
    TelLab = 0
    TelCel = 0
    'Recorro los datos de la linea
    For indice = 2 To UBound(arrayLinea)
        Select Case indice
        '------------------------------------------------------------------------------------------
        Case 2: 'Modelo de Domicilio
            If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                modeloDomicilio = Trim(arrayLinea(indice))
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Error: Modelo de domicilio Obligatorio."
                str_error = str_error & "<tr><td>Modelo de domicilio invalido</td></tr>" & vbCrLf
                huboError = True
            End If
            
        '------------------------------------------------------------------------------------------
        Case 3: 'Tipo de Domicilio
            If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                StrSql = " SELECT tidonro FROM tipodomi WHERE upper(tidodes) = '" & UCase(Mid(arrayLinea(indice), 1, 30)) & "'"
                OpenRecordset StrSql, rs_Datos
                If Not rs_Datos.EOF Then
                    tipoDomicilio = rs_Datos!tidonro
                    Flog.writeline Espacios(Tabulador * 1) & "Tipo de Domicilio encontrado."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Tipo de Domicilio no existe, se creara."
                    StrSql = " INSERT INTO tipodomi (tidodes) VALUES " & _
                             " ('" & Left(arrayLinea(indice), 30) & "') "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    tipoDomicilio = getLastIdentity(objConn, "tipodomi")
                End If
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Error: Tipo Domicilio Obligatorio."
                str_error = str_error & "<tr><td>Tipo de domicilio invalido</td></tr>" & vbCrLf
                huboError = True
            End If
        '------------------------------------------------------------------------------------------
        Case 4: 'Calle
            If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                calle = Left(arrayLinea(indice), 30)
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Error: Calle Obligatoria."
                str_error = str_error & "<tr><td>Calle invalida</td></tr>" & vbCrLf
                huboError = True
            End If
            
        '------------------------------------------------------------------------------------------
        Case 5: 'Numero
            If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                Numero = Left(arrayLinea(indice), 8)
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Error: Nro de calle Obligatorio."
                str_error = str_error & "<tr><td>Numero de Calle invalido</td></tr>" & vbCrLf
                huboError = True
            End If
            
        '------------------------------------------------------------------------------------------
        Case 6: 'piso
            If UCase(arrayLinea(indice)) <> "N/A" Then
                piso = Left(arrayLinea(indice), 8)
            End If
        '------------------------------------------------------------------------------------------
        Case 7: 'depto
            If UCase(arrayLinea(indice)) <> "N/A" Then
                depto = Left(arrayLinea(indice), 8)
            End If
        '------------------------------------------------------------------------------------------
        Case 8: 'torre
            If UCase(arrayLinea(indice)) <> "N/A" Then
                torre = Left(arrayLinea(indice), 8)
            End If
        '------------------------------------------------------------------------------------------
        Case 9: 'manzana
            If UCase(arrayLinea(indice)) <> "N/A" Then
                manzana = Left(arrayLinea(indice), 8)
            End If
        '------------------------------------------------------------------------------------------
        Case 10: 'codigo postal
            cp = Left(arrayLinea(indice), 12)
        
        Case 11: 'Entre Calles
            If UCase(arrayLinea(indice)) <> "N/A" Then
                entreCalles = Left(arrayLinea(indice), 80)
            End If
        '------------------------------------------------------------------------------------------
        Case 12: 'Barrio
            If UCase(arrayLinea(indice)) <> "N/A" Then
                barrio = Left(arrayLinea(indice), 30)
            End If
        '------------------------------------------------------------------------------------------
        Case 13: 'Localidad
            
            If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                StrSql = " SELECT locnro FROM localidad WHERE upper(locdesc) = '" & UCase(Mid(arrayLinea(indice), 1, 60)) & "'"
                OpenRecordset StrSql, rs_Datos
                If Not rs_Datos.EOF Then
                    localidad = rs_Datos!locnro
                    Flog.writeline Espacios(Tabulador * 1) & "Localidad encontrada."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Localidad no existe, se creara."
                    StrSql = " INSERT INTO localidad (locdesc) VALUES " & _
                             " ('" & Left(arrayLinea(indice), 60) & "') "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    localidad = getLastIdentity(objConn, "localidad")
                End If
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Error: Localidad Obligatoria."
                str_error = str_error & "<tr><td>Localidad no valida</td></tr>" & vbCrLf
                huboError = True
            End If

        '------------------------------------------------------------------------------------------
        Case 14: 'Distrito o partido
            If UCase(arrayLinea(indice)) <> "N/A" Then
                StrSql = " SELECT partnro FROM partido WHERE upper(partnom) = '" & UCase(Mid(arrayLinea(indice), 1, 30)) & "'"
                OpenRecordset StrSql, rs_Datos
                If Not rs_Datos.EOF Then
                    partido = rs_Datos!partnro
                    Flog.writeline Espacios(Tabulador * 1) & "Partido o Distrito encontrado."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Partido o Distrito no existe, se creara."
                    StrSql = " INSERT INTO partido (partnom) VALUES " & _
                             " ('" & Left(arrayLinea(indice), 30) & "') "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    partido = getLastIdentity(objConn, "partido")
                End If
            Else
                partido = "null"
            End If
        '------------------------------------------------------------------------------------------
        Case 15: 'zona
            If UCase(arrayLinea(indice)) <> "N/A" Then
                StrSql = " SELECT zonanro FROM zona WHERE upper(zonadesc) = '" & UCase(Mid(arrayLinea(indice), 1, 60)) & "'"
                OpenRecordset StrSql, rs_Datos
                If Not rs_Datos.EOF Then
                    zona = rs_Datos!zonanro
                    Flog.writeline Espacios(Tabulador * 1) & "Partido o Distrito encontrado."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Partido o Distrito no existe, se creara."
                    'Necesito la provincia para insertar la zona
                    zona = 0
                    zonadesc = Left(arrayLinea(indice), 60)
                End If
            Else
                zona = "null"
            End If
        '------------------------------------------------------------------------------------------
        Case 16: 'Provincia
            If UCase(arrayLinea(indice)) <> "N/A" Then
                StrSql = " SELECT provnro FROM provincia WHERE upper(provdesc) = '" & UCase(Mid(arrayLinea(indice), 1, 60)) & "'"
                OpenRecordset StrSql, rs_Datos
                If Not rs_Datos.EOF Then
                    provincia = rs_Datos!provnro
                    Flog.writeline Espacios(Tabulador * 1) & "Provincia encontrada."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Provincia no existe, se creara."
                    'pais fijo en 3 (argentina) por ahora .
                    StrSql = " INSERT INTO provincia (provdesc,paisnro) VALUES " & _
                             " ('" & Left(arrayLinea(indice), 60) & "',3) "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    provincia = getLastIdentity(objConn, "provincia")
                End If
                
                If CStr(zona) = "0" Then 'si zona es 0 es que no la encontro, case 14
                    'ahora que tengo la provincia inserto la zona
                    StrSql = " INSERT INTO zona (zonadesc,provnro) VALUES " & _
                             " ('" & Left(zonadesc, 60) & "'," & provincia & ") "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    zona = getLastIdentity(objConn, "zona")
                End If
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Error: Provincia Obligatoria."
                str_error = str_error & "<tr><td>Provincia no valida</td></tr>" & vbCrLf
                huboError = True
            End If
        '------------------------------------------------------------------------------------------
        Case 17: 'Pais - Fijo por ahora - ARGENTINA
            pais = 3
        '------------------------------------------------------------------------------------------
        Case 18: 'Telefono particular
            If UCase(arrayLinea(indice)) <> "N/A" Then
                TelPart = Mid(arrayLinea(indice), 1, 20)
            End If
        '------------------------------------------------------------------------------------------
        Case 19: 'Telefono laboral
            If UCase(arrayLinea(indice)) <> "N/A" Then
                TelLab = Mid(arrayLinea(indice), 1, 20)
            End If
        '------------------------------------------------------------------------------------------
        Case 20: 'Telefono celular
            If UCase(arrayLinea(indice)) <> "N/A" Then
                TelCel = Mid(arrayLinea(indice), 1, 20)
            End If
        '------------------------------------------------------------------------------------------
        End Select
    Next
    
    If Not huboError Then
       str_error = str_error & "<tr><td colspan='2' align='center'><b>No hay Errores</b></td></tr>" & vbCrLf
       StrSql = " SELECT domnro FROM cabdom WHERE tidonro = " & tipoDomicilio & " AND ternro = " & ternro
       OpenRecordset StrSql, rs_Datos
       
       If rs_Datos.EOF Then
           Flog.writeline Espacios(Tabulador * 1) & "No existe domicilio para el empleado, se creara."
           'No existe el domicilio para el empleado
           StrSql = " INSERT INTO cabdom (tipnro,ternro,domdefault,tidonro,modnro) VALUES " & _
                    " (1," & ternro & ",0," & tipoDomicilio & ",1) "
           objConn.Execute StrSql, , adExecuteNoRecords
           
           Flog.writeline Espacios(Tabulador * 1) & "Cabecera de domicilio creada."
           domnro = getLastIdentity(objConn, "cabdom")
           
           StrSql = " INSERT INTO detdom (domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal, " & _
                    " entrecalles,barrio,locnro,partnro,zonanro,provnro,paisnro) VALUES " & _
                    " (" & domnro & ",'" & calle & "','" & Numero & "','" & piso & "','" & depto & "'," & _
                    " '" & torre & "','" & manzana & "','" & cp & "','" & entreCalles & "','" & barrio & "'," & _
                       localidad & "," & partido & "," & zona & "," & provincia & "," & pais & ")"
           objConn.Execute StrSql, , adExecuteNoRecords
           Flog.writeline Espacios(Tabulador * 1) & "Detalle de domicilio creado."
           
           Select Case CStr(destino)
               '------------------------------------------------------------------------------------------
               Case "2": 'Version R2
                   If CStr(TelPart) <> "0" Then
                       StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular) " & _
                                " VALUES (" & domnro & ",'" & TelPart & "',0,-1,0) "
                       objConn.Execute StrSql, , adExecuteNoRecords
                       Flog.writeline Espacios(Tabulador * 1) & "Telefono particular R2, creado."
                   End If
                   
                   If CStr(TelLab) <> "0" Then
                       StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular) " & _
                                " VALUES (" & domnro & ",'" & TelLab & "',0,0,0) "
                       objConn.Execute StrSql, , adExecuteNoRecords
                       Flog.writeline Espacios(Tabulador * 1) & "Telefono laboral R2, creado."
                   
                   End If
                   
                   If CStr(TelCel) <> "0" Then
                       StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular) " & _
                                " VALUES (" & domnro & ",'" & TelCel & "',0,0,-1) "
                       objConn.Execute StrSql, , adExecuteNoRecords
                       Flog.writeline Espacios(Tabulador * 1) & "Telefono celular R2, creado."
                   End If
               '------------------------------------------------------------------------------------------
               Case "3": 'Version R3
                   If CStr(TelPart) <> "0" Then
                       StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular,tipotel) " & _
                                " VALUES (" & domnro & ",'" & TelPart & "',0,-1,0,1) "
                       objConn.Execute StrSql, , adExecuteNoRecords
                       Flog.writeline Espacios(Tabulador * 1) & "Telefono particular R3, creado."
                   End If
                   
                   If CStr(TelLab) <> "0" Then
                       StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular,tipotel) " & _
                                " VALUES (" & domnro & ",'" & TelLab & "',0,0,0,2) "
                       objConn.Execute StrSql, , adExecuteNoRecords
                       Flog.writeline Espacios(Tabulador * 1) & "Telefono laboral R3, creado."
                   
                   End If
                   
                   If CStr(TelCel) <> "0" Then
                       StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular,tipotel) " & _
                                " VALUES (" & domnro & ",'" & TelCel & "',0,0,-1,3) "
                       objConn.Execute StrSql, , adExecuteNoRecords
                       Flog.writeline Espacios(Tabulador * 1) & "Telefono celular R3, creado."
                   End If
               '------------------------------------------------------------------------------------------
           End Select
       Else
           'ya tiene domicilio
           StrSql = " UPDATE detdom SET " & _
                    " calle = '" & calle & "'," & _
                    " nro = '" & Numero & "'," & _
                    " piso = '" & piso & "'," & _
                    " oficdepto = '" & depto & "'," & _
                    " torre = '" & torre & "'," & _
                    " manzana = '" & manzana & "'," & _
                    " codigopostal = '" & cp & "'," & _
                    " entrecalles = '" & entreCalles & "'," & _
                    " barrio = '" & barrio & "'," & _
                    " locnro = " & localidad & "," & _
                    " partnro = " & partido & "," & _
                    " zonanro = " & zona & "," & _
                    " provnro = " & provincia & "," & _
                    " paisnro = " & pais & _
                    " WHERE domnro = " & rs_Datos!domnro
           objConn.Execute StrSql, , adExecuteNoRecords
           Flog.writeline Espacios(Tabulador * 1) & "Domicilio Actualizado."
           
           Select Case CStr(destino)
               Case "2": 'Version R2
                    If CStr(TelPart) <> "0" Then
                        'chequeo si existe telefono particular
                        StrSql = " SELECT domnro FROM  telefono WHERE  teldefault = -1 AND domnro = " & rs_Datos!domnro
                        OpenRecordset StrSql, rs_tel
                        If Not rs_tel.EOF Then
                            'existe el telefono, lo actualizo
                            StrSql = " UPDATE telefono SET " & _
                                     " telnro = ' " & TelPart & "'," & _
                                     " telfax = 0, telcelular = 0 " & _
                                     " WHERE domnro = " & rs_Datos!domnro & " AND teldefault = -1 "
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline Espacios(Tabulador * 1) & "Telefono particular R2, actualizado."
                        Else
                            'No existe el telefono, lo inserto
                            StrSql = " INSERT INTO telefono (domnro,telnro,telfax,tedefault,telcelular) VALUES " & _
                                     " (" & rs_Datos!domnro & ",'" & TelPart & "',0,-1,0) "
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline Espacios(Tabulador * 1) & "Telefono particular R2, insertado."
                        End If
                   End If
                   
                   If CStr(TelLab) <> "0" Then
                       'chequeo si existe telefono particular
                       StrSql = " SELECT domnro FROM  telefono WHERE  telfax = 0 AND teldefault = 0 AND telcelular = 0 AND domnro = " & rs_Datos!domnro
                       OpenRecordset StrSql, rs_tel
                        If Not rs_tel.EOF Then
                            'existe el telefono, lo actualizo
                            StrSql = " UPDATE telefono SET " & _
                                     " telnro = ' " & TelLab & "'," & _
                                     " telfax = 0, teldefault = 0 ,telcelular = 0 " & _
                                     " WHERE domnro = " & rs_Datos!domnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline Espacios(Tabulador * 1) & "Telefono laboral R2, actualizado."
                       Else
                            'No existe el telefono, lo inserto
                            StrSql = " INSERT INTO telefono (domnro,telnro,telfax,tedefault,telcelular) VALUES " & _
                                     " (" & rs_Datos!domnro & ",'" & TelLab & "',0,0,0) "
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline Espacios(Tabulador * 1) & "Telefono laboral R2, insertado."
                       End If
                   End If
                   
                   If CStr(TelCel) <> "0" Then
                       'chequeo si existe telefono particular
                       StrSql = " SELECT domnro FROM  telefono WHERE telcelular = -1 AND domnro = " & rs_Datos!domnro
                       OpenRecordset StrSql, rs_tel
                        If Not rs_tel.EOF Then
                            'existe el telefono, lo actualizo
                            StrSql = " UPDATE telefono SET " & _
                                     " telnro = ' " & TelCel & "'," & _
                                     " telfax = 0, teldefault = 0 ,telcelular = -1 " & _
                                     " WHERE domnro = " & rs_Datos!domnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline Espacios(Tabulador * 1) & "Telefono celular R2, actualizado."
                        Else
                            'No existe el telefono, lo inserto
                            StrSql = " INSERT INTO telefono (domnro,telnro,telfax,tedefault,telcelular) VALUES " & _
                                     " (" & rs_Datos!domnro & ",'" & TelCel & "',0,0,-1) "
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline Espacios(Tabulador * 1) & "Telefono celular R2, insertado."
                        End If
                   End If
    
               Case "3": 'Version R3
                   If CStr(TelPart) <> "0" Then
                       'Chequeo si existe telefono particular
                       StrSql = " SELECT domnro FROM Telefono WHERE domnro = " & rs_Datos!domnro & " AND tipotel = 1 "
                       OpenRecordset StrSql, rs_tel
                       If rs_tel.EOF Then
                           StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular,tipotel) VALUES " & _
                                    " (" & rs_Datos!domnro & ",'" & TelPart & "',0,-1,0,1 )"
                           objConn.Execute StrSql, , adExecuteNoRecords
                           Flog.writeline Espacios(Tabulador * 1) & "Telefono particular R3, insertado."
                       Else
                           StrSql = " UPDATE telefono SET " & _
                                    " telnro = '" & TelPart & "'," & _
                                    " telfax = 0, teldefault = -1 ,telcelular = 0, tipotel = 1 " & _
                                    " WHERE domnro = " & rs_Datos!domnro
                           objConn.Execute StrSql, , adExecuteNoRecords
                           Flog.writeline Espacios(Tabulador * 1) & "Telefono particular R3, actualizado."
                       End If
                   End If
                   If CStr(TelLab) <> "0" Then
                       StrSql = " SELECT domnro FROM Telefono WHERE domnro = " & rs_Datos!domnro & " AND tipotel = 3 "
                       OpenRecordset StrSql, rs_tel
                       If rs_tel.EOF Then
                           StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular,tipotel) VALUES " & _
                                    " (" & rs_Datos!domnro & ",'" & TelPart & "',0,0,0,3 )"
                           objConn.Execute StrSql, , adExecuteNoRecords
                           Flog.writeline Espacios(Tabulador * 1) & "Telefono laboral R3, insertado."
                       Else
                           StrSql = " UPDATE telefono SET " & _
                                    " telnro = '" & TelLab & "'," & _
                                    " telfax = 0, teldefault = 0 ,telcelular = 0, tipotel = 3 " & _
                                    " WHERE domnro = " & rs_Datos!domnro
                           objConn.Execute StrSql, , adExecuteNoRecords
                           Flog.writeline Espacios(Tabulador * 1) & "Telefono laboral R3, actualizado."
                       End If
                   End If
                   
                   If CStr(TelCel) <> "0" Then
                       StrSql = " SELECT domnro FROM Telefono WHERE domnro = " & rs_Datos!domnro & " AND tipotel = 2 "
                       OpenRecordset StrSql, rs_tel
                       If rs_tel.EOF Then
                           StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular,tipotel) VALUES " & _
                                    " (" & rs_Datos!domnro & ",'" & TelPart & "',0,0,-1,3 )"
                           objConn.Execute StrSql, , adExecuteNoRecords
                           Flog.writeline Espacios(Tabulador * 1) & "Telefono celular R3, insertado."
                       Else
                           StrSql = " UPDATE telefono SET " & _
                                    " telnro = '" & TelCel & "'," & _
                                    " telfax = 0, teldefault = 0 ,telcelular = -1, tipotel = 2 " & _
                                    " WHERE domnro = " & rs_Datos!domnro
                           objConn.Execute StrSql, , adExecuteNoRecords
                           Flog.writeline Espacios(Tabulador * 1) & "Telefono celular R3, actualizado."
                       End If
                   End If
           End Select
           
       End If
    End If
    str_error = str_error & "</table>" & vbCrLf
    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "

    Flog.writeline Espacios(Tabulador * 0) & "Fin de modelo 1001 - Domicilios."
    
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    If rs_Datos.State = adStateOpen Then rs_Datos.Close
End Sub
Public Sub import_modelo1002(ByVal strLinea As String, ByVal destino As Long, ByRef str_error As String)
'Documentos
Dim arrayLinea
Dim Legajo As Long
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Datos As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim tipoDocumento As String
Dim nroDocumento As String
Dim Nro_Institucion As String
Dim separador As String
Dim huboError As Boolean

    separador = SeparadorModelo(1002)
    arrayLinea = Split(strLinea, separador)

    Flog.writeline Espacios(Tabulador * 0) & "Comienzo del modelo 1002 - Documentos."
    Flog.writeline Espacios(Tabulador * 1) & "Busco el empleado: " & arrayLinea(1)
    'Chequeo si el empleado existe
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & arrayLinea(1)
    Legajo = arrayLinea(1)
    OpenRecordset StrSql, rs_Empleado
    If Not rs_Empleado.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " existe en el sistema."
        ternro = rs_Empleado!ternro
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " No existe en el sistema."
        Exit Sub
    End If
        
    huboError = False
    str_error = str_error & "<TABLE width='50%' border='1' bordercolor='#333333' cellpadding='0' cellspacing='0'>" & vbCrLf
    str_error = str_error & "<tr><TH style='background-color:#d13528;color:#FFFFFF;' colspan='2' align='center'><b>Modelo 1002, para el legajo: " & arrayLinea(1) & "</TH></tr>" & vbCrLf
    
    'Recorro los datos de la linea
    For indice = 2 To UBound(arrayLinea)
        Select Case indice
            '------------------------------------------------------------------------------------------
            Case 2: 'Tipo de Documento
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = " SELECT tidnro FROM tipodocu WHERE tidsigla = '" & UCase(Mid(arrayLinea(indice), 1, 8)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        tipoDocumento = rs_Datos!tidnro
                        Flog.writeline Espacios(Tabulador * 1) & "Tipo de Documento encontrado."
                    Else
                        
                        'busco la primera institucion, si no existe la creo
                        StrSql = " SELECT * FROM institucion "
                        If rs.State = adStateOpen Then rs.Close
                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            Nro_Institucion = rs!instnro
                        Else
                            'creo una
                            StrSql = " INSERT INTO institucion (instdes,instabre) VALUES ('NACIONAL','NAC')"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Nro_Institucion = getLastIdentity(objConn, "institucion")
                        End If
                        
                        StrSql = " INSERT INTO tipodocu (tidnom,tidsigla,tidsist,instnro,tidunico,tidvalunico) VALUES " & _
                                 " ('" & Left(arrayLinea(indice), 8) & "','" & Left(arrayLinea(indice), 8) & "',0," & Nro_Institucion & ",0,0)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        tipoDocumento = getLastIdentity(objConn, "tipodocu")
                        Flog.writeline Espacios(Tabulador * 1) & "Tipo de Documento no encontrado, se creara."
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Tipo de Documento Obligatorio."
                    str_error = str_error & "<tr><td>Tipo de Documento invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
            Case 3: 'Numero de Documento
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    nroDocumento = Mid(arrayLinea(indice), 1, 30)
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Nro de Documento Obligatorio."
                    str_error = str_error & "<tr><td>Numero de Documento invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
        End Select
    Next
        
    If Not huboError Then
        str_error = str_error & "<tr><td colspan='2' align='center'><b>No hay Errores</b></td></tr>" & vbCrLf
        'chequeo si ya existe el documento
        StrSql = " SELECT tidnro,ternro FROM ter_doc WHERE ternro = " & ternro & " AND tidnro = " & tipoDocumento
        OpenRecordset StrSql, rs_Datos
        If rs_Datos.EOF Then
            StrSql = " INSERT INTO ter_doc (tidnro,ternro,nrodoc) VALUES " & _
                     " (" & tipoDocumento & "," & ternro & ",'" & nroDocumento & "') "
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Documento insertado."
        Else
            StrSql = " UPDATE ter_doc SET " & _
                     " nrodoc = '" & nroDocumento & "'" & _
                     " WHERE tidnro = " & rs_Datos!tidnro & " AND ternro = " & ternro
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Documento insertado."
        End If
    End If
    str_error = str_error & "</table>" & vbCrLf
    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "
    Flog.writeline Espacios(Tabulador * 0) & "Fin del modelo 1002 - Documentos."
    
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    If rs_Datos.State = adStateOpen Then rs_Datos.Close
    If rs.State = adStateOpen Then rs.Close
    
End Sub
Public Sub import_modelo1003(ByVal strLinea As String, ByVal destino As Long, ByRef str_error As String)
'Fases
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Datos As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim NroCausa As String
Dim fecAlta As String
Dim fecBaja As String
Dim estado As String
Dim sueldo As String
Dim vacaciones As String
Dim indenmizacion As String
Dim real As String
Dim altaRec As String
Dim separador As String
Dim huboError As Boolean

    separador = SeparadorModelo(1003)
    arrayLinea = Split(strLinea, separador)

    Flog.writeline Espacios(Tabulador * 0) & "Comienzo del modelo 1003 - Fases."
    Flog.writeline Espacios(Tabulador * 1) & "Busco el empleado: " & arrayLinea(1)
    'Chequeo si el empleado existe
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & arrayLinea(1)
    Legajo = arrayLinea(1)
    OpenRecordset StrSql, rs_Empleado
    If Not rs_Empleado.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " existe en el sistema."
        ternro = rs_Empleado!ternro
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " No existe en el sistema."
        Exit Sub
    End If
        
    huboError = False
    str_error = str_error & "<TABLE width='50%' border='1' bordercolor='#333333' cellpadding='0' cellspacing='0'>" & vbCrLf
    str_error = str_error & "<tr><TH style='background-color:#d13528;color:#FFFFFF;' colspan='2' align='center'><b>Modelo 1003, para el legajo: " & arrayLinea(1) & "</TH></tr>" & vbCrLf
    
    'Recorro los datos de la linea
    For indice = 2 To UBound(arrayLinea)
        Select Case indice
            '------------------------------------------------------------------------------------------
            Case 2: 'Causa Baja
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    StrSql = " SELECT caunro FROM causa WHERE upper(caudes) = '" & UCase(Mid(arrayLinea(indice), 1, 80)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        NroCausa = rs_Datos!caunro
                        Flog.writeline Espacios(Tabulador * 1) & "Tipo de causa encontrada."
                    Else
                        StrSql = " INSERT INTO causa (caudes,causist,caudesvin,empnro) VALUES " & _
                                 " ('" & Left(arrayLinea(indice), 80) & "',0,-1,0)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        NroCausa = getLastIdentity(objConn, "causa")
                        Flog.writeline Espacios(Tabulador * 1) & "Tipo de causa no encontrada, se creara."
                    End If
                Else
                    NroCausa = 0
                    Flog.writeline Espacios(Tabulador * 1) & "Causa de baja sin informar."
                End If
            '------------------------------------------------------------------------------------------
            Case 3: 'fecalta
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    fecAlta = arrayLinea(indice)
                    Flog.writeline Espacios(Tabulador * 1) & "Fecha de Alta encontrada."
                Else
                    str_error = str_error & "<tr><td>Fecha de alta de fase invalida</td></tr>" & vbCrLf
                    Flog.writeline Espacios(Tabulador * 1) & "Fecha de Alta sin informar (obligatoria)."
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
            Case 4: 'fecBaja
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    fecBaja = arrayLinea(indice)
                    Flog.writeline Espacios(Tabulador * 1) & "Fecha de baja informada."
                Else
                    fecBaja = "null"
                    Flog.writeline Espacios(Tabulador * 1) & "Fecha de baja sin informar."
                End If
            '------------------------------------------------------------------------------------------
            Case 5: 'Estado
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    Flog.writeline Espacios(Tabulador * 1) & "Estado de la fase informada."
                    If UCase(Trim(arrayLinea(indice))) = "ACTIVA" Then
                        estado = "-1"
                    Else
                        estado = "0"
                    End If
                Else
                    str_error = str_error & "<tr><td>Estado de fase invalido</td></tr>" & vbCrLf
                    Flog.writeline Espacios(Tabulador * 1) & "Estado de la fase sin informar (Obligatorio)."
                    huboError = True
                End If

            '------------------------------------------------------------------------------------------
            Case 6: 'Sueldo
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    Flog.writeline Espacios(Tabulador * 1) & "Sueldo informado."
                    If UCase(Trim(arrayLinea(indice))) = "SI" Then
                        sueldo = "-1"
                    Else
                        sueldo = "0"
                    End If
                Else
                    str_error = str_error & "<tr><td>Sueldo invalido</td></tr>" & vbCrLf
                    Flog.writeline Espacios(Tabulador * 1) & "Sueldo sin informar (Obligatorio)."
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
            Case 7: 'Vacaciones
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    Flog.writeline Espacios(Tabulador * 1) & "Vacaciones informadas."
                    If UCase(Trim(arrayLinea(indice))) = "SI" Then
                        vacaciones = "-1"
                    Else
                        vacaciones = "0"
                    End If
                Else
                    str_error = str_error & "<tr><td>Vacaciones fase invalido</td></tr>" & vbCrLf
                    huboError = True
                    Flog.writeline Espacios(Tabulador * 1) & "Vacaciones sin informar (Obligatorio)."
                End If
            '------------------------------------------------------------------------------------------
            Case 8: 'Indenmizacion
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    Flog.writeline Espacios(Tabulador * 1) & "Indenmizacion informada."
                    If UCase(Trim(arrayLinea(indice))) = "SI" Then
                        indenmizacion = "-1"
                    Else
                        indenmizacion = "0"
                    End If
                Else
                    str_error = str_error & "<tr><td>Indenmizacion fase invalido</td></tr>" & vbCrLf
                    Flog.writeline Espacios(Tabulador * 1) & "Indenmizacion sin informar (Obligatorio)."
                    huboError = True
                End If
                
            '------------------------------------------------------------------------------------------
            Case 9: 'Real
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    Flog.writeline Espacios(Tabulador * 1) & "Fase Real informada."
                    If UCase(Trim(arrayLinea(indice))) = "SI" Then
                        real = "-1"
                    Else
                        real = "0"
                    End If
                Else
                    str_error = str_error & "<tr><td>Fase Real invalida</td></tr>" & vbCrLf
                    Flog.writeline Espacios(Tabulador * 1) & "Fase Real sin informar (Obligatorio)."
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
            Case 10: 'AltaRec
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    Flog.writeline Espacios(Tabulador * 1) & "Alta Reconocida informada."
                    If UCase(Trim(arrayLinea(indice))) = "SI" Then
                        altaRec = "-1"
                    Else
                        altaRec = "0"
                    End If
                Else
                    str_error = str_error & "<tr><td>Alta Reconocida invalida</td></tr>" & vbCrLf
                    Flog.writeline Espacios(Tabulador * 1) & "Alta Reconocida sin informar (Obligatorio)."
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
        End Select
    Next
    

    If Not huboError Then
            str_error = str_error & "<tr><td colspan='2' align='center'><b>No hay Errores</b></td></tr>" & vbCrLf
        '-----------------------------------------------------------------------
        'Controles de fases ----------------------------------------------------
        'Si no existe fase ==> simplemente crea la fase
        StrSql = "SELECT * FROM fases WHERE empleado = " & ternro
        StrSql = StrSql & " ORDER BY altfec DESC"
        OpenRecordset StrSql, rs
        If rs.EOF Then
            If NroCausa <> 0 Then
              StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
              StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
              StrSql = StrSql & " VALUES(" & ternro & "," & NroCausa & "," & cambiaFecha(fecAlta) & "," & cambiaFecha(fecBaja) & ","
              StrSql = StrSql & estado & ",0," & sueldo & "," & vacaciones & "," & indenmizacion & ","
              StrSql = StrSql & real & "," & altaRec & ")"
              
              objConn.Execute StrSql, , adExecuteNoRecords
              
            Else
              StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
              StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
              StrSql = StrSql & " VALUES(" & ternro & ",Null," & cambiaFecha(fecAlta) & "," & cambiaFecha(fecBaja) & ","
              StrSql = StrSql & estado & ",0," & sueldo & "," & vacaciones & "," & indenmizacion & ","
              StrSql = StrSql & real & "," & altaRec & ")"
              objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
        Else    'Ya tiene fases que arranca en esa fecha ==> Actualizo
            StrSql = "SELECT * FROM fases "
            StrSql = StrSql & " WHERE empleado = " & ternro
            StrSql = StrSql & " AND altfec = " & cambiaFecha(fecAlta)
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                'Si la fecha hasta del registro encontrado es mayor o nulo ==> actualizo
                '   sino error
                If EsNulo(rs!bajfec) Then
                    'Actualizo
                    StrSql = " UPDATE fases SET "
                    StrSql = StrSql & " bajfec = " & cambiaFecha(fecBaja)
                    StrSql = StrSql & ",estado = " & estado
                    StrSql = StrSql & ",sueldo = " & sueldo
                    StrSql = StrSql & ",vacaciones = " & vacaciones
                    StrSql = StrSql & ",indemnizacion = " & indenmizacion
                    StrSql = StrSql & ",real = " & real
                    StrSql = StrSql & ",fasrecofec = " & altaRec
                    If NroCausa <> 0 Then
                        StrSql = StrSql & ",caunro = " & NroCausa
                    End If
                    StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                Else
                    If UCase(fecBaja) = "NULL" Then
                        'no debe existir ningun otro registro
                        'sino
                        'Error
                        StrSql = "SELECT fasnro FROM fases "
                        StrSql = StrSql & " WHERE empleado = " & ternro
                        StrSql = StrSql & " AND altfec > " & cambiaFecha(fecAlta)
                        StrSql = StrSql & " AND fasnro <> " & rs!fasnro
                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            Flog.writeline Espacios(Tabulador * 1) & "La fase se superpone con otra ya existente."
                            Exit Sub
                        Else
                            'Actualizo
                            StrSql = " UPDATE fases SET "
                            StrSql = StrSql & " bajfec = " & cambiaFecha(fecBaja)
                            StrSql = StrSql & ",estado = " & estado
                            StrSql = StrSql & ",sueldo = " & sueldo
                            StrSql = StrSql & ",vacaciones = " & vacaciones
                            StrSql = StrSql & ",indemnizacion = " & indenmizacion
                            StrSql = StrSql & ",real = " & real
                            StrSql = StrSql & ",fasrecofec = " & altaRec
                            If NroCausa <> 0 Then
                                StrSql = StrSql & ",caunro = " & NroCausa
                            End If
                            StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                        End If
                    Else
                        If rs!bajfec >= FBaja Then
                            'Actualizo
                            StrSql = " UPDATE fases SET "
                            StrSql = StrSql & " bajfec = " & cambiaFecha(fecBaja)
                            StrSql = StrSql & ",estado = " & estado
                            StrSql = StrSql & ",sueldo = " & sueldo
                            StrSql = StrSql & ",vacaciones = " & vacaciones
                            StrSql = StrSql & ",indemnizacion = " & indenmizacion
                            StrSql = StrSql & ",real = " & real
                            StrSql = StrSql & ",fasrecofec = " & altaRec
                            If NroCausa <> 0 Then
                                StrSql = StrSql & ",caunro = " & NroCausa
                            End If
                            StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "La fase se superpone con otra ya existente."
                            Exit Sub
        
                        End If
                    End If
                End If
            Else
                If UCase(fecBaja) = "NULL" Then
                    StrSql = "SELECT fasnro FROM fases "
                    StrSql = StrSql & " WHERE empleado = " & ternro
                    StrSql = StrSql & " AND bajfec IS NULL "
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        Flog.writeline Espacios(Tabulador * 1) & "La fase se superpone con otra ya existente."
                        Exit Sub
        
                    Else
                        'fecha desde nueva tiene que ser mayor que todas las fases existentes
                        StrSql = "SELECT fasnro FROM fases "
                        StrSql = StrSql & " WHERE empleado = " & ternro
                        StrSql = StrSql & " AND bajfec >= " & cambiaFecha(fecAlta)
                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            'Error. No se puede actualizar
                            Flog.writeline Espacios(Tabulador * 1) & "La fase se superpone con otra ya existente."
                            Exit Sub
        
                        Else
                            'Inserto
                            If NroCausa <> 0 Then
                              StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
                              StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                              StrSql = StrSql & " VALUES(" & ternro & "," & NroCausa & "," & cambiaFecha(fecAlta) & "," & cambiaFecha(fecBaja) & ","
                              StrSql = StrSql & estado & ",0," & sueldo & "," & vacaciones & "," & indenmizacion & ","
                              StrSql = StrSql & real & "," & altaRec & ")"
                              objConn.Execute StrSql, , adExecuteNoRecords
        
                            Else
                              StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
                              StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                              StrSql = StrSql & " VALUES(" & ternro & ",Null," & cambiaFecha(fecAlta) & "," & cambiaFecha(fecBaja) & ","
                              StrSql = StrSql & estado & ",0," & sueldo & "," & vacaciones & "," & indenmizacion & ","
                              StrSql = StrSql & real & "," & altaRec & ")"
                              objConn.Execute StrSql, , adExecuteNoRecords
                            End If
                            
                        End If
                    End If
                Else
                    'SI existe fases que cruce la nueva ==>
                    StrSql = "SELECT fasnro FROM fases "
                    StrSql = StrSql & " WHERE empleado = " & ternro
                    StrSql = StrSql & " AND ("
                    StrSql = StrSql & " (altfec <= " & cambiaFecha(fecAlta) & " AND bajfec >=" & cambiaFecha(fecBaja) & ")"
                    StrSql = StrSql & " OR "
                    StrSql = StrSql & " (altfec <= " & cambiaFecha(fecAlta) & " AND bajfec >=" & cambiaFecha(fecAlta) & ")"
                    StrSql = StrSql & " OR "
                    StrSql = StrSql & " (altfec >= " & cambiaFecha(fecAlta) & " AND bajfec <=" & cambiaFecha(fecBaja) & ")"
                    StrSql = StrSql & " OR "
                    StrSql = StrSql & " (altfec >= " & cambiaFecha(fecAlta) & " AND altfec <= " & cambiaFecha(fecBaja) & " AND bajfec >=" & cambiaFecha(fecBaja) & ")"
                    StrSql = StrSql & " OR "
                    StrSql = StrSql & " (altfec >= " & cambiaFecha(fecAlta) & " AND altfec <= " & cambiaFecha(fecBaja) & " AND bajfec IS NULL)"
                    StrSql = StrSql & " OR "
                    StrSql = StrSql & " (altfec <= " & cambiaFecha(fecAlta) & " AND bajfec IS NULL)"
                    StrSql = StrSql & " )"
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        Flog.writeline Espacios(Tabulador * 1) & "La fase se superpone con otra ya existente."
                        Exit Sub
                    Else
                        'Inserto
                        If NroCausa <> 0 Then
                          StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
                          StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                          StrSql = StrSql & " VALUES(" & ternro & "," & NroCausa & "," & cambiaFecha(fecAlta) & "," & cambiaFecha(fecBaja) & ","
                          StrSql = StrSql & estado & ",0," & sueldo & "," & vacaciones & "," & indenmizacion & ","
                          StrSql = StrSql & real & "," & altaRec & ")"
                          objConn.Execute StrSql, , adExecuteNoRecords
                          
                        Else
                          StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
                          StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                          StrSql = StrSql & " VALUES(" & ternro & ",Null," & cambiaFecha(fecAlta) & "," & cambiaFecha(fecBaja) & ","
                          StrSql = StrSql & estado & ",0," & sueldo & "," & vacaciones & "," & indenmizacion & ","
                          StrSql = StrSql & real & "," & altaRec & ")"
                          objConn.Execute StrSql, , adExecuteNoRecords
                        End If
        
                    End If
                End If
            End If
        End If
    End If 'fin if hubo_Error
    str_error = str_error & "</table>" & vbCrLf
    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "
    
    Flog.writeline Espacios(Tabulador * 0) & "Fin del modelo 1003 - Fases."
    
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    If rs_Datos.State = adStateOpen Then rs_Datos.Close
    If rs.State = adStateOpen Then rs.Close

End Sub
Public Sub import_modelo1004(ByVal strLinea As String, ByVal destino As Long, ByRef str_error As String)
'Datos de familiares
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Datos As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim ternro As Long
Dim NroTercero As Long
Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim Nombre2 As String
Dim fecnac As String
Dim pais As String
Dim paisNac As String
Dim Nacionalidad As String
Dim EstCiv As String
Dim Sexo As String
Dim fecAlta As String
Dim Estudia As String
Dim parentesco As String
Dim nroDocumento As String
Dim tipoDocumento As String
Dim calle As String
Dim Numero As String
Dim piso As String
Dim depto As String
Dim torre As String
Dim manzana As String
Dim cp As String
Dim entreCalles As String
Dim barrio As String
Dim localidad As String
Dim partido As String
Dim zona As String
Dim provincia As String
Dim telefono As String
Dim Nro_Institucion As String
Dim nro_osocial As String
Dim nro_planos As String
Dim nro_aviso As String
Dim nro_salario As String
Dim nro_gan As String
Dim fechaInicio As String
Dim IngresoDom As Boolean
Dim separador As String
Dim huboError As Boolean

    separador = SeparadorModelo(1004)
    arrayLinea = Split(strLinea, separador)

    Flog.writeline Espacios(Tabulador * 0) & "Comienzo del modelo 1004 - Familiar."
    Flog.writeline Espacios(Tabulador * 1) & "Busco el empleado: " & arrayLinea(1)
    'Chequeo si el empleado existe
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & arrayLinea(1)
    Legajo = arrayLinea(1)
    OpenRecordset StrSql, rs_Empleado
    If Not rs_Empleado.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " existe en el sistema."
        ternro = rs_Empleado!ternro
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " No existe en el sistema."
        Exit Sub
    End If
        
    huboError = False
    str_error = str_error & "<TABLE width='50%' border='1' bordercolor='#333333' cellpadding='0' cellspacing='0'>" & vbCrLf
    str_error = str_error & "<tr><TH style='background-color:#d13528;color:#FFFFFF;' colspan='2' align='center'><b>Modelo 1004, para el legajo: " & arrayLinea(1) & "</TH></tr>" & vbCrLf
    
    'Recorro los datos de la linea
    For indice = 2 To UBound(arrayLinea)
        Select Case indice
            Case 2: 'Apellidos
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If InStr(arrayLinea(indice), "@") = 0 Then
                        apellido = arrayLinea(indice)
                    Else
                        apellido = Split(arrayLinea(indice), "@")(0)
                        apellido2 = Split(arrayLinea(indice), "@")(1)
                    End If
                    Flog.writeline Espacios(Tabulador * 1) & "Apellidos obtenidos."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Apellido Obligatorio."
                    str_error = str_error & "<tr><td>Apellido invalido</td></tr>" & vbCrLf
                    huboError = True
                End If

            '------------------------------------------------------------------------------------------
             Case 3: 'Nombres
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If InStr(arrayLinea(indice), "@") = 0 Then
                        nombre = arrayLinea(indice)
                    Else
                        nombre = Split(arrayLinea(indice), "@")(0)
                        Nombre2 = Split(arrayLinea(indice), "@")(1)
                    End If
                    Flog.writeline Espacios(Tabulador * 1) & "Apellidos obtenidos."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Nombre Obligatorio."
                    str_error = str_error & "<tr><td>Nombre invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 4: 'Fecha de nacimiento
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    fecnac = arrayLinea(indice)
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Fecha de nacimiento Obligatorio."
                    str_error = str_error & "<tr><td>Fecha de nacimiento invalida</td></tr>" & vbCrLf
                    huboError = True
                End If
                
            '------------------------------------------------------------------------------------------
             Case 5: 'Pais del tercero
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = "SELECT paisnro FROM pais WHERE upper(paisdesc) = '" & UCase(Mid(arrayLinea(indice), 1, 60)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        paisNac = rs_Datos!paisnro
                        Flog.writeline Espacios(Tabulador * 1) & "Pais encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Pais no encontrado, se creara."
                        StrSql = " INSERT INTO pais (paisdesc,paisdef) " & _
                                 " VALUES ('" & UCase(Mid(arrayLinea(indice), 1, 60)) & "',0)"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        paisNac = getLastIdentity(objConn, "pais")
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Pais de nacimiento Obligatorio."
                    str_error = str_error & "<tr><td>Pais de nacimiento invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 6: 'Nacionalidad
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = "SELECT nacionalnro FROM nacionalidad WHERE upper(nacionaldes) = '" & UCase(Mid(arrayLinea(indice), 1, 30)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        Nacionalidad = rs_Datos!nacionalnro
                        Flog.writeline Espacios(Tabulador * 1) & "Nacionalidad encontrada."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Nacionalidad no encontrada, se creara."
                        StrSql = " INSERT INTO nacionalidad (nacionaldes,nacionaldefault) " & _
                                 " VALUES ('" & UCase(Mid(arrayLinea(indice), 1, 30)) & "',0)"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        Nacionalidad = getLastIdentity(objConn, "nacionalidad")
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Nacionalidad Obligatoria."
                    str_error = str_error & "<tr><td>Nacionalidad invalida</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 7: 'Estado Civil
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = "SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(Mid(arrayLinea(indice), 1, 30)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        EstCiv = rs_Datos!estcivnro
                        Flog.writeline Espacios(Tabulador * 1) & "Estado Civil encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Estado Civil no encontrado, se creara."
                        StrSql = " INSERT INTO estcivil (estcivdesabr) " & _
                                 " VALUES ('" & UCase(Mid(arrayLinea(indice), 1, 30)) & "')"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        EstCiv = getLastIdentity(objConn, "estcivil")
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Estado civil Obligatorio."
                    str_error = str_error & "<tr><td>Estado civil invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 8: 'Sexo
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If UCase(arrayLinea(indice)) = "M" Then
                        Sexo = -1
                    Else
                        Sexo = 0
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Sexo Obligatorio."
                    str_error = str_error & "<tr><td>Sexo invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
             Case 9: 'Parentesco
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = " SELECT parenro FROM parentesco WHERE upper(paredesc) = '" & UCase(Mid(arrayLinea(indice), 1, 40)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        parentesco = rs_Datos!parenro
                        Flog.writeline Espacios(Tabulador * 1) & "Parentesco encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Parentesco no encontrado, se creara."
                        StrSql = " INSERT INTO parentesco (paredesc,paresist,parepadre,parereqpadre) " & _
                                 " VALUES ('" & Left(arrayLinea(indice), 40) & "',0,0,0)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        parentesco = getLastIdentity(objConn, "parentesco")
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Parentesco Obligatorio."
                    str_error = str_error & "<tr><td>Parentesco invalido</td></tr>" & vbCrLf
                    huboError = True
                End If

                
            '------------------------------------------------------------------------------------------
             Case 10: 'Estudia
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    If UCase(arrayLinea(indice)) = "SI" Then
                        Estudia = -1
                    Else
                        Estudia = 0
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Estudia Obligatorio."
                    str_error = str_error & "<tr><td>Campo Estudia invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
                
            '------------------------------------------------------------------------------------------
             Case 11: 'Nivel de estudio
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = "SELECT nivnro FROM nivest WHERE upper(nivdesc) = '" & UCase(Mid(arrayLinea(indice), 1, 40)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        NivEstudio = rs_Datos!nivnro
                        Flog.writeline Espacios(Tabulador * 1) & "Nivel de Estudio encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Nivel de Estudio no encontrado, se creara."
                        StrSql = " INSERT INTO nivest (nivdesc,nivsist,nivestfli) " & _
                                 " VALUES ('" & UCase(Mid(arrayLinea(indice), 1, 40)) & "',0,0)"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        NivEstudio = getLastIdentity(objConn, "nivest")
                    End If
                Else
                    If Estudia = -1 Then
                        Flog.writeline Espacios(Tabulador * 1) & "Error: El empleado Estudia, Nivel de Estudio Obligatorio."
                        str_error = str_error & "<tr><td>Si Estudia debe tener un nivel de estudio</td></tr>" & vbCrLf
                        huboError = True
                    End If
                End If
            
            '------------------------------------------------------------------------------------------
            Case 12: 'Tipo de Documento
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    StrSql = " SELECT tidnro FROM tipodocu WHERE tidsigla = '" & UCase(Mid(arrayLinea(indice), 1, 8)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        tipoDocumento = rs_Datos!tidnro
                        Flog.writeline Espacios(Tabulador * 1) & "Tipo de Documento encontrado."
                    Else
                        'busco la primera institucion, si no existe la creo
                        StrSql = " SELECT * FROM institucion "
                        If rs.State = adStateOpen Then rs.Close
                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            Nro_Institucion = rs!instnro
                        Else
                            'creo una nueva
                            StrSql = " INSERT INTO institucion (instdes,instabre) VALUES ('NACIONAL','NAC')"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Nro_Institucion = getLastIdentity(objConn, "institucion")
                            Flog.writeline Espacios(Tabulador * 1) & "Institucion de expedicion, se creara."
                        End If
                        
                    
                        StrSql = " INSERT INTO tipodocu (tidnom,tidsigla,tidsist,instnro,tidunico,tidvalunico) VALUES " & _
                                 " ('" & Left(arrayLinea(indice), 8) & "','" & Left(arrayLinea(indice), 8) & "',0," & Nro_Institucion & ",0,0)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        tipoDocumento = getLastIdentity(objConn, "tipodocu")
                        Flog.writeline Espacios(Tabulador * 1) & "Tipo de Documento no encontrado, se creara."
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Tipo de Documento Obligatorio."
                    str_error = str_error & "<tr><td>Tipo de Documento invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
            Case 13: 'Numero de Documento
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    nroDocumento = Mid(arrayLinea(indice), 1, 30)
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Nro de Documento Obligatorio."
                    str_error = str_error & "<tr><td>Numero de Documento invalido</td></tr>" & vbCrLf
                    huboError = True
                End If
            '------------------------------------------------------------------------------------------
            Case 14: 'Calle
                IngresoDom = True
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    calle = Left(arrayLinea(indice), 30)
                Else
                    IngresoDom = False
                End If
            '------------------------------------------------------------------------------------------
            Case 15: 'Numero
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    Numero = Left(arrayLinea(indice), 8)
                End If
            '------------------------------------------------------------------------------------------
            Case 16: 'piso
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    piso = Left(arrayLinea(indice), 8)
                End If
            '------------------------------------------------------------------------------------------
            Case 17: 'depto
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    depto = Left(arrayLinea(indice), 8)
                End If
            '------------------------------------------------------------------------------------------
            Case 18: 'torre
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    torre = Left(arrayLinea(indice), 8)
                End If
            '------------------------------------------------------------------------------------------
            Case 19: 'manzana
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    manzana = Left(arrayLinea(indice), 8)
                End If
            '------------------------------------------------------------------------------------------
            Case 20: 'codigo postal
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    cp = Left(arrayLinea(indice), 12)
                End If
            
            '------------------------------------------------------------------------------------------
            Case 21: 'Entre Calles
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    entreCalles = Left(arrayLinea(indice), 80)
                End If
            '------------------------------------------------------------------------------------------
            Case 22: 'Barrio
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    barrio = Left(arrayLinea(indice), 30)
                End If
            '------------------------------------------------------------------------------------------
            Case 23: 'Localidad
                StrSql = " SELECT locnro FROM localidad WHERE upper(locdesc) = '" & UCase(Mid(arrayLinea(indice), 1, 60)) & "'"
                OpenRecordset StrSql, rs_Datos
                If Not rs_Datos.EOF Then
                    localidad = rs_Datos!locnro
                    Flog.writeline Espacios(Tabulador * 1) & "Localidad encontrada."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Localidad no existe, se creara."
                    StrSql = " INSERT INTO localidad (locdesc) VALUES " & _
                             " ('" & Left(arrayLinea(indice), 60) & "') "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    localidad = getLastIdentity(objConn, "localidad")
                End If
            '------------------------------------------------------------------------------------------
            Case 24: 'Distrito o partido
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    StrSql = " SELECT partnro FROM partido WHERE upper(partnom) = '" & UCase(Mid(arrayLinea(indice), 1, 30)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        partido = rs_Datos!partnro
                        Flog.writeline Espacios(Tabulador * 1) & "Partido o Distrito encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Partido o Distrito no existe, se creara."
                        StrSql = " INSERT INTO partido (partnom) VALUES " & _
                                 " ('" & Left(arrayLinea(indice), 30) & "') "
                        objConn.Execute StrSql, , adExecuteNoRecords
                        partido = getLastIdentity(objConn, "partido")
                    End If
                Else
                    partido = "null"
                End If
            '------------------------------------------------------------------------------------------
            Case 25: 'zona
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    StrSql = " SELECT zonanro FROM zona WHERE upper(zonadesc) = '" & UCase(Mid(arrayLinea(indice), 1, 60)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        zona = rs_Datos!zonanro
                        Flog.writeline Espacios(Tabulador * 1) & "Partido o Distrito encontrado."
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Partido o Distrito no existe, se creara."
                        'Necesito la provincia para insertar la zona
                        zona = 0
                        zonadesc = Left(arrayLinea(indice), 60)
                    End If
                Else
                    zona = "null"
                End If
            '------------------------------------------------------------------------------------------
            Case 26: 'Provincia
                StrSql = " SELECT provnro FROM provincia WHERE upper(provdesc) = '" & UCase(Mid(arrayLinea(indice), 1, 60)) & "'"
                OpenRecordset StrSql, rs_Datos
                If Not rs_Datos.EOF Then
                    provincia = rs_Datos!provnro
                    Flog.writeline Espacios(Tabulador * 1) & "Provincia encontrada."
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Provincia no existe, se creara."
                    'pais fijo en 3 (argentina) por ahora .
                    StrSql = " INSERT INTO provincia (provdesc,paisnro) VALUES " & _
                             " ('" & Left(arrayLinea(indice), 60) & "',3) "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    provincia = getLastIdentity(objConn, "provincia")
                End If
                
                If CStr(zona) = "0" Then 'si zona es 0 es que no la encontro, case 14
                    'ahora que tengo la provincia inserto la zona
                    StrSql = " INSERT INTO zona (zonadesc,provnro) VALUES " & _
                             " ('" & Left(zonadesc, 60) & "'," & provincia & ") "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    zona = getLastIdentity(objConn, "zona")
                End If
            '------------------------------------------------------------------------------------------
            Case 27: 'Pais - Fijo por ahora - ARGENTINA
                pais = 3
            '------------------------------------------------------------------------------------------
            Case 28: 'Telefono particular
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    telefono = arrayLinea(indice)
                End If
            '------------------------------------------------------------------------------------------
            Case 29: 'Obra Social
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    StrSql = " SELECT ternro FROM osocial WHERE UPPER(osdesc) = '" & UCase(Mid(arrayLinea(indice), 1, 100)) & "'"
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        nro_osocial = rs!ternro
                    Else
                        nro_osocial = "0"
                    End If
        
                Else
                    nro_osocial = "0"
                End If
            '------------------------------------------------------------------------------------------
            Case 30: 'Plan Obra Social
                If UCase(arrayLinea(indice)) <> "N/A" Then
                    If CStr(nro_osocial) <> "0" Then
                        StrSql = " SELECT plnro FROM planos WHERE UPPER(plnom) = '" & UCase(Mid(arrayLinea(indice), 1, 60)) & "'"
                        StrSql = StrSql & " AND osocial = " & nro_osocial

                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            nro_planos = rs!plnro
                        Else
                            nro_planos = "0"
                        End If
                    Else
                        nro_planos = "0"
                    End If
                    
                Else
                    nro_planos = "0"
                End If
            '------------------------------------------------------------------------------------------
            Case 31: 'Avisar ante emergencia
            If UCase(arrayLinea(indice)) = "SI" Then
                nro_aviso = -1
            Else
                nro_aviso = 0
            End If
            '------------------------------------------------------------------------------------------
            Case 32: 'Paga Salario fliar
                If UCase(arrayLinea(indice)) = "SI" Then
                    nro_salario = -1
                Else
                    nro_salario = 0
                End If
            '------------------------------------------------------------------------------------------
            Case 33: 'ganancias
                If UCase(arrayLinea(indice)) = "SI" Then
                    nro_gan = -1
                Else
                    nro_gan = 0
                End If
            '------------------------------------------------------------------------------------------
            Case 34: 'Feha de inicio vinculo
                If UCase(arrayLinea(indice)) = "N/A" Then
                        fechaInicio = "Null"
                Else
                    fechaInicio = arrayLinea(indice)
                End If
                If fechaInicio = "Null" Then
                    ' Busco la fecha de alta reconocida
                     StrSql = "SELECT altfec FROM fases "
                     StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = fases.empleado "
                     StrSql = StrSql & " WHERE empleado.ternro = " & ternro
                     StrSql = StrSql & " AND fasrecofec = -1 "
                     OpenRecordset StrSql, rs
                     If Not rs.EOF Then
                        'Calculo la fecha de Inicio
                        If IsDate(fecnac) Then
                            If CDate(rs!altfec) > CDate(fecnac) Then
                                    fechaInicio = rs!altfec
                            Else
                                    fechaInicio = fecnac
                            End If
                        End If
                    End If
                End If
            '------------------------------------------------------------------------------------------
            Case 35: 'Fecha de vencimiento
                If UCase(arrayLinea(indice)) <> "N/A" Then
                        fechaVto = arrayLinea(indice)
                Else
                    fechaVto = "Null"
                End If
                
        End Select
    Next
    ' Veo que la fecha de vencimiento no sea menor que la de inicio
    If IsDate(fechaVto) Then
       If CDate(fechaInicio) > CDate(fechaVto) Then
           Flog.writeline Espacios(Tabulador * 1) & "La Fecha de inicio es mayor a la de fecha de vencimiento."
           Exit Sub
        End If
    End If
    
    If Not huboError Then
        str_error = str_error & "<tr><td colspan='2' align='center'><b>No hay Errores</b></td></tr>" & vbCrLf
    
        fechaInicio = ConvFecha(fechaInicio)
        fecnac = ConvFecha(fecnac)
    
        '-------------------------------------------------------------------------------------------------
        StrSql = "SELECT * FROM tercero "
        StrSql = StrSql & " INNER JOIN ter_tip ON tercero.ternro = ter_tip.ternro AND ter_tip.tipnro = 3 "
        StrSql = StrSql & " INNER JOIN familiar ON familiar.ternro = tercero.ternro AND familiar.empleado = " & ternro
        StrSql = StrSql & " WHERE ternom = '" & nombre & "' AND ternom2 = '" & Nombre2 & "'"
        StrSql = StrSql & " AND terape = '" & apellido & "' AND terape2 = '" & apellido2 & "'"
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
            'Inserto el tercero asociado al familiar
            
            StrSql = " INSERT INTO tercero(ternom,ternom2,terape,terape2,terfecnac,tersex,nacionalnro,paisnro,estcivnro)"
            StrSql = StrSql & " VALUES('" & nombre & "','" & Nombre2 & "','" & apellido & "','" & apellido2 & "'," & fecnac & "," & Sexo & ","
            If Nacionalidad <> 0 Then
              StrSql = StrSql & Nacionalidad & ","
            Else
              StrSql = StrSql & "Null,"
            End If
            If paisNac <> 0 Then
              StrSql = StrSql & paisNac & ","
            Else
              StrSql = StrSql & "Null,"
            End If
            StrSql = StrSql & EstCiv & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        
            NroTercero = getLastIdentity(objConn, "tercero")
            
            'Inserto el Registro correspondiente en ter_tip
            StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & NroTercero & ",3)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Inserto el Familiar
            StrSql = " INSERT INTO familiar(empleado,ternro,parenro,famest,famestudia,famcernac,famsalario,famemergencia,famcargadgi,osocial,plnro,famternro,famfec,famfecvto)"
            StrSql = StrSql & " values(" & ternro & "," & NroTercero & "," & parentesco & ",-1," & Estudia & ",0," & nro_salario & "," & nro_aviso & "," & nro_gan & "," & nro_osocial & "," & nro_planos & ",0," & fechaInicio & "," & fechaVto & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Inserto los estudios de familiar
            If Estudia = -1 Then
                StrSql = " INSERT INTO estudio_actual (ternro, nivnro) VALUES (" & NroTercero & "," & NivEstudio & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Familiar no existe, se creara."
        
        Else
            'Actualizo los datos
            StrSql = "UPDATE tercero SET "
            StrSql = StrSql & " terfecnac = " & fecnac
            StrSql = StrSql & " ,tersex = " & Sexo
            If Nacionalidad <> 0 Then
                StrSql = StrSql & " ,nacionalnro = " & Nacionalidad
            End If
            If paisNac <> 0 Then
                StrSql = StrSql & " ,paisnro = " & paisNac
            End If
            StrSql = StrSql & " WHERE ternro = " & rs!ternro
            objConn.Execute StrSql, , adExecuteNoRecords
        
            NroTercero = rs!ternro
        
            StrSql = "UPDATE familiar SET "
            StrSql = StrSql & " parenro = " & parentesco
            StrSql = StrSql & " ,famestudia = " & Estudia
            StrSql = StrSql & " ,famsalario = " & nro_salario
            StrSql = StrSql & " ,famemergencia = " & nro_aviso
            StrSql = StrSql & " ,famcargadgi = " & nro_gan
            StrSql = StrSql & " ,osocial = " & nro_osocial
            StrSql = StrSql & " ,plnro = " & nro_planos
            StrSql = StrSql & " ,famternro = 0"
            StrSql = StrSql & " WHERE empleado = " & ternro
            StrSql = StrSql & " AND ternro = " & NroTercero
            objConn.Execute StrSql, , adExecuteNoRecords
        
            If Estudia = -1 Then
                StrSql = " SELECT ternro FROM estudio_actual WHERE ternro = " & NroTercero
                
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    StrSql = " INSERT INTO estudio_actual (ternro, nivnro) VALUES (" & NroTercero & "," & NivEstudio & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = " UPDATE estudio_actual SET nivnro = " & NivEstudio
                    StrSql = StrSql & "WHERE ternro = " & NroTercero
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            
        
            End If
            Flog.writeline Espacios(Tabulador * 1) & "Familiar actualizado."
        End If
        
        'Inserto los Documentos
        'Valido que no exista el documento
        StrSql = "SELECT * FROM ter_doc WHERE nrodoc = '" & nroDocumento & "'"
        StrSql = StrSql & " AND tidnro = " & tipoDocumento
        OpenRecordset StrSql, rs_Datos
        If Not rs_Datos.EOF Then
                Flog.writeline Espacios(Tabulador * 1) & "Documento existente."
        Else
            If nroDocumento <> "" And nroDocumento <> "N/A" And TipDoc <> "N/A" Then
                StrSql = "SELECT * FROM ter_doc WHERE ternro = " & NroTercero
                StrSql = StrSql & " AND tidnro = " & tipoDocumento
        
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                    StrSql = StrSql & " VALUES(" & NroTercero & "," & tipoDocumento & ",'" & nroDocumento & "')"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Documento creado."
                Else
                    StrSql = " UPDATE ter_doc SET "
                    StrSql = StrSql & " nrodoc = '" & nroDocumento & "'"
                    StrSql = StrSql & " WHERE ternro = " & NroTercero
                    StrSql = StrSql & " AND tidnro = " & tipoDocumento
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Documento actualizado."
                End If
            End If
        End If
        
        
        
        'Inserto el Domicilio
         If Not IngresoDom = False Then
            StrSql = "SELECT * FROM cabdom  "
            StrSql = StrSql & " WHERE tipnro = 1"
            StrSql = StrSql & " AND ternro = " & NroTercero
            StrSql = StrSql & " AND domdefault = -1"
            StrSql = StrSql & " AND tidonro = 2"
            
            OpenRecordset StrSql, rs
            If rs.EOF Then
                StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
                StrSql = StrSql & " VALUES(1," & NroTercero & ",-1,2)"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                NroDom = getLastIdentity(objConn, "cabdom")
                
                StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,"
                StrSql = StrSql & "locnro,provnro,paisnro,barrio,partnro,zonanro) "
                StrSql = StrSql & " VALUES (" & NroDom & ",'" & calle & "','" & Nro_Nrodom & "','" & piso & "','"
                StrSql = StrSql & depto & "','" & torre & "','" & manzana & "','" & cp & "'," & localidad & ","
                StrSql = StrSql & provincia & "," & pais & ",'" & barrio & "'," & partido & "," & zona & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Flog.writeline Espacios(Tabulador * 1) & "Domicilio creado."
        
                If telefono <> "" Then
                    Select Case CStr(destino)
                        Case "2":
                            StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                            StrSql = StrSql & " VALUES(" & NroDom & ",'" & telefono & "',0,-1,0)"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline Espacios(Tabulador * 1) & "Telefono familiar creado (R2)."
                        Case "3":
                            StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
                            StrSql = StrSql & " VALUES(" & NroDom & ",'" & telefono & "',0,-1,0,1)"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline Espacios(Tabulador * 1) & "Telefono creado."
                    End Select
                End If
            Else
                StrSql = " UPDATE detdom SET "
                StrSql = StrSql & " calle =" & "'" & calle & "'"
                StrSql = StrSql & ",nro =" & "'" & Numero & "'"
                StrSql = StrSql & ",piso =" & "'" & piso & "'"
                StrSql = StrSql & ",oficdepto =" & "'" & depto & "'"
                StrSql = StrSql & ",torre =" & "'" & torre & "'"
                StrSql = StrSql & ",manzana =" & "'" & manzana & "'"
                StrSql = StrSql & ",codigopostal =" & "'" & cp & "'"
                StrSql = StrSql & ",entrecalles =" & "'" & entreCalles & "'"
                StrSql = StrSql & ",locnro =" & localidad
                StrSql = StrSql & ",provnro =" & provincia
                StrSql = StrSql & ",paisnro =" & pais
                StrSql = StrSql & ", partnro = " & partido
                StrSql = StrSql & ", zonanro =" & zona
                StrSql = StrSql & " WHERE domnro = " & rs!domnro
                objConn.Execute StrSql, , adExecuteNoRecords
            
                Flog.writeline Espacios(Tabulador * 1) & "Domicilio Actualizado."
            
                If telefono <> "" Then
                    StrSql = "SELECT * FROM telefono "
                    StrSql = StrSql & " WHERE domnro =" & rs!domnro
                    StrSql = StrSql & " AND telnro ='" & telefono & "'"
        
                    OpenRecordset StrSql, rs_Datos
                    If rs_Datos.EOF Then
                        Select Case CStr(destino)
                            Case "2":
                                StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                                StrSql = StrSql & " VALUES(" & rs!domnro & ",'" & telefono & "',0,-1,0)"
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline Espacios(Tabulador * 1) & "Telefono Familiar."
                            
                            Case "3":
                                StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
                                StrSql = StrSql & " VALUES(" & rs!domnro & ",'" & telefono & "',0,-1,0,1)"
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline Espacios(Tabulador * 1) & "Telefono Familiar."
                        End Select
                    End If
                End If
            End If
        End If

    End If ' fin if huboError
    str_error = str_error & "</table>" & vbCrLf
    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "
    Flog.writeline Espacios(Tabulador * 0) & "Fin del modelo 1004 - Familiar."
    
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    If rs.State = adStateOpen Then rs.Close
End Sub

Public Sub import_modelo1005(ByVal strLinea As String, ByVal destino As Long, ByRef str_error As String)
'Datos del empleado
'Documentos
Dim arrayLinea
Dim Legajo As Long
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Datos As New ADODB.Recordset
Dim rs_Est As New ADODB.Recordset
Dim tipoEstructura As String
Dim estructura As String
Dim fechaDesde As String
Dim fechaHasta As String
Dim separador As String
Dim huboError As Boolean
Dim estrnro As Long
Dim Inserto_estr As Boolean
Dim ter_est As Long
Dim ternro As Long

    separador = SeparadorModelo(1005)
    arrayLinea = Split(strLinea, separador)

    Flog.writeline Espacios(Tabulador * 0) & "Comienzo del modelo 1005 - Estructuras."
    Flog.writeline Espacios(Tabulador * 1) & "Busco el empleado: " & arrayLinea(1)
    'Chequeo si el empleado existe
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & arrayLinea(1)
    Legajo = arrayLinea(1)
    OpenRecordset StrSql, rs_Empleado
    If Not rs_Empleado.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " existe en el sistema."
        ternro = rs_Empleado!ternro
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & arrayLinea(1) & " No existe en el sistema."
        Exit Sub
    End If
        
    huboError = False
    str_error = str_error & "<TABLE width='50%' border='1' bordercolor='#333333' cellpadding='0' cellspacing='0'>" & vbCrLf
    str_error = str_error & "<tr><TH style='background-color:#d13528;color:#FFFFFF;' colspan='2' align='center'><b>Modelo 1005, para el legajo: " & arrayLinea(1) & "</TH></tr>" & vbCrLf
    
    'Recorro los datos de la linea
    For indice = 2 To UBound(arrayLinea)
        Select Case indice
            '------------------------------------------------------------------------------------------
            Case 2: 'Tipo de Estructura
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    
                    StrSql = "SELECT tenro FROM tipoestructura WHERE UPPER(tedabr) = '" & UCase(Mid(arrayLinea(indice), 1, 25)) & "'"
                    OpenRecordset StrSql, rs_Datos
                    If Not rs_Datos.EOF Then
                        tipoEstructura = rs_Datos!Tenro
                    Else
                        StrSql = "INSERT INTO tipoestructura(tedabr,tesist,tedepbaja,cenro) VALUES('" & Mid(arrayLinea(indice), 1, 25) & "',0,0,1)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        tipoEstructura = getLastIdentity(objConn, "tipoestructura")
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Tipo Estructura Obligatorio."
                    str_error = str_error & "<tr><td>Tipo de Estrutura invalida</td></tr>" & vbCrLf
                    huboError = True
                End If
            
            '------------------------------------------------------------------------------------------
            Case 3: 'Estructura
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    estructura = arrayLinea(indice)
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Estructura Obligatoria."
                    str_error = str_error & "<tr><td>Estrutura invalida</td></tr>" & vbCrLf
                    huboError = True
                End If
                
            '------------------------------------------------------------------------------------------
            Case 4: 'Fecha Desde
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    fechaDesde = arrayLinea(indice)
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Error: Fecha Desde Obligatoria."
                    str_error = str_error & "<tr><td>Fecha Desde invalida</td></tr>" & vbCrLf
                    huboError = True
                End If
                
            '------------------------------------------------------------------------------------------
            Case 5: 'Fecha Hasta
                If UCase(arrayLinea(indice)) <> "N/A" And UCase(arrayLinea(indice)) <> "" Then
                    fechaHasta = arrayLinea(indice)
                Else
                    fechaHasta = ""
                End If
                
        End Select
    Next
    
    If UCase(fechaHasta) <> "NULL" And UCase(fechaDesde) <> "NULL" And UCase(fechaHasta) <> "" And UCase(fechaDesde) <> "" Then
        If CDate(fechaDesde) > CDate(fechaHasta) Then
            str_error = str_error & "<tr><td>Error: Fecha Hasta menor a Fecha Desde de la estructura</td></tr>" & vbCrLf
            huboError = True
        End If
    End If

    If Not huboError Then
        str_error = str_error & "<tr><td colspan='2' align='center'><b>No hay Errores</b></td></tr>" & vbCrLf
        Call ValidaEstructura(tipoEstructura, estructura, estrnro, Inserto_estr)
        
        'Validación para llamar o no a crearTercero, dependiendo del tipo de estructura
        Call VerSiCrearTercero(tipoEstructura, estructura, ter_est)
        
        If Inserto_estr Then
            'Veo si tengo que crear el complemento, dependiendo del tipo de estructura
            Select Case tipoEstructura
                Case 23
                    'Plan de Obra Social Elegida
                    'Tengo que buscar la Obra Social Elegida
                    StrSql = "SELECT os.ternro FROM his_estructura his "
                    StrSql = StrSql & " INNER JOIN estructura est ON est.tenro = his.tenro AND est.estrnro = his.estrnro "
                    StrSql = StrSql & " INNER JOIN osocial os ON os.osdesc = est.estrdabr "
                    StrSql = StrSql & " WHERE his.tenro = 17 and his.ternro = " & ternro
                    StrSql = StrSql & " ORDER BY htetdesde DESC, htethasta ASC "
                    OpenRecordset StrSql, rs_Est
                    If Not rs_Est.EOF Then
                        ter_est = rs_Est!ternro
                        Call VerSiCrearComplemento(tipoEstructura, estrnro, estructura, ter_est)
                    End If
                Case 25
                    'Plan de Obrea Social Por Ley
                    'Tengo que buscar la Obra Social por Ley
                    StrSql = "SELECT os.ternro FROM his_estructura his "
                    StrSql = StrSql & " INNER JOIN estructura est ON est.tenro = his.tenro AND est.estrnro = his.estrnro "
                    StrSql = StrSql & " INNER JOIN osocial os ON os.osdesc = est.estrdabr "
                    StrSql = StrSql & " WHERE his.tenro = 24 and his.ternro = " & ternro
                    StrSql = StrSql & " ORDER BY htetdesde DESC, htethasta ASC "
                    OpenRecordset StrSql, rs_Est
                    If Not rs_Est.EOF Then
                        ter_est = rs_Est!ternro
                        Call VerSiCrearComplemento(tipoEstructura, estrnro, estructura, ter_est)
                    End If
                Case Else
                    Call VerSiCrearComplemento(tipoEstructura, estrnro, estructura, ter_est)
            End Select

        End If
        Call AsignarEstructura(tipoEstructura, estrnro, ternro, fechaDesde, fechaHasta)
    End If ' fin if huboError
    str_error = str_error & "</table>" & vbCrLf
    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "
    Flog.writeline Espacios(Tabulador * 0) & "Fin del modelo 1005 - Estructuras."
    
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    If rs.State = adStateOpen Then rs.Close
    
End Sub

Public Sub ValidaEstructura(ByRef TipoEstr As String, ByRef Valor As String, ByRef CodEst As Long, ByRef Inserto_estr As Boolean)
Dim Rs_Estr As New ADODB.Recordset
Dim d_estructura As String
Dim CodExt As String
Dim l_pos1 As Long
Dim l_pos2 As Long

    StrSql = " SELECT estrnro FROM estructura WHERE UPPER(estructura.estrdabr) = '" & UCase(Mid(Valor, 1, 60)) & "'"
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    
    OpenRecordset StrSql, Rs_Estr
    If Not Rs_Estr.EOF Then
                
            CodEst = Rs_Estr!estrnro
            Inserto_estr = False
            
    Else
            
        StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
        StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & UCase(Mid(Valor, 1, 60)) & "',1,-1,'')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        CodEst = getLastIdentity(objConn, "estructura")
        
        Inserto_estr = True
    End If

End Sub

Public Sub VerSiCrearTercero(Tenro As String, Valor As String, ByRef CodTer)


  Select Case Tenro

    Case 1
        'Sucursal
        CreaTercero 10, Valor, CodTer
    Case 10
        'Empresa
        CreaTercero 10, Valor, CodTer
    Case 15
        'Caja de Jubilacion
        CreaTercero 6, Valor, CodTer
    Case 16
        'Sindicato
        CreaTercero 5, Valor, CodTer
    Case 17
        'OS Elegida
        CreaTercero 4, Valor, CodTer
    Case 24
        'Obra social por Ley
        CreaTercero 4, Valor, CodTer
    Case 28
        'Agencia
        CreaTercero 7, Valor, CodTer
    Case 40
        'ART
        CreaTercero 8, Valor, CodTer
    Case 41
        'Banco de Pago
        CreaTercero 13, Valor, CodTer
    Case Else
        'Cuando no se crea el tercero
        CodTer = 0

  End Select
 
End Sub

Public Sub CreaTercero(TipoTer As Long, Valor As String, ByRef CodTer)

Dim rs As New ADODB.Recordset
Dim rs_Ter As New ADODB.Recordset

Dim d_estructura As String
Dim l_pos1 As Long
Dim l_pos2 As Long

    
  d_estructura = Valor
    
  StrSql = " SELECT * FROM tercero "
  StrSql = StrSql & " INNER JOIN ter_tip ON tercero.ternro = ter_tip.ternro AND ter_tip.tipnro =" & TipoTer
  StrSql = StrSql & " WHERE terrazsoc = '" & Valor & "'"
  If rs_Ter.State = adStateOpen Then rs_Ter.Close
  OpenRecordset StrSql, rs_Ter
  If rs_Ter.EOF Then
    
      StrSql = " INSERT INTO tercero(terrazsoc,tersex)"
      StrSql = StrSql & " VALUES('" & Mid(d_estructura, 1, 60) & "',-1)"
      objConn.Execute StrSql, , adExecuteNoRecords
    
      CodTer = getLastIdentity(objConn, "tercero")
    
      StrSql = " INSERT INTO ter_tip(ternro,tipnro) "
      StrSql = StrSql & " VALUES(" & CodTer & "," & TipoTer & ")"
      objConn.Execute StrSql, , adExecuteNoRecords
    Else
        CodTer = rs_Ter!ternro
    End If

    If rs_Ter.State = adStateOpen Then rs_Ter.Close
    Set rs_Ter = Nothing
End Sub

Public Sub VerSiCrearComplemento(Tenro As String, codEstr As Long, Valor As String, CodTer As Long)

  Select Case Tenro

    Case 1
        'Sucursal
        Complementos1 CodTer, codEstr
    Case 4
        'Puesto
        Complementos4 codEstr, Valor
    Case 10
        'Empresa
        Complementos10 CodTer, codEstr, Valor
    Case 15
        'Caja de Jubilacion
        Complementos15 CodTer, codEstr
    Case 16
        'Sindicato
        Complementos16 CodTer, codEstr
    Case 17
        'OS Elegida
        Complementos17 CodTer, codEstr, Valor
    Case 18
        'Contrato
        Complementos18 CodTer, codEstr, Valor
    Case 19
        'Convenio
        Complementos19 codEstr
    Case 22
        'Forma de Liquidacion
        Complementos22 CodTer, codEstr, Valor
    Case 23
        'Plan de Obra social Elegida
        Complementos23 CodTer, codEstr, Valor
    Case 24
        'Obra social por Ley
        Complementos17 CodTer, codEstr, Valor
    Case 25
        'Plan de Obra social por Ley
        Complementos23 CodTer, codEstr, Valor
    Case 28
        'Agencia
        Complementos28 CodTer, codEstr, Valor
    Case 40
        'ART
        Complementos40 CodTer, codEstr, Valor
    Case 41
        'Banco de Pago
        Complementos41 CodTer, codEstr, Valor


  End Select
 
End Sub

Public Sub AsignarEstructura(TipoEstr As String, CodEst As Long, CodTer As Long, FAlta As String, FBaja As String)
    ' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta la estructura. si existe una estructura del mismo tipo en el intervalo
'               la estructura será actualizada.
' ---------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset
Dim rs_his As New ADODB.Recordset

Dim F_Cierre_Temp As Date


    F_Cierre_Temp = DateAdd("d", -1, CDate(FAlta))

    If CodEst <> 0 Then
        If nro_ModOrg <> 0 Then
            StrSql = "SELECT * FROM adptte_estr WHERE tplatenro = " & nro_ModOrg & " AND tenro = " & TipoEstr
            OpenRecordset StrSql, rs
            If rs.EOF Then
                tplaorden = tplaorden + 1
                StrSql = "INSERT INTO adptte_estr(tplatenro,tenro,tplaestroblig,tplaestrorden) VALUES (" & nro_ModOrg & "," & TipoEstr & ",0," & tplaorden & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    
        StrSql = "SELECT * FROM his_estructura "
        StrSql = StrSql & " WHERE tenro = " & TipoEstr
        StrSql = StrSql & " AND ternro = " & CodTer
        StrSql = StrSql & " AND (htetdesde <= " & cambiaFecha(FAlta) & ") AND"
        StrSql = StrSql & " ((" & cambiaFecha(FAlta) & " <= htethasta) or (htethasta is null))"
        StrSql = StrSql & " ORDER BY htetdesde "
        If rs_his.State = adStateOpen Then rs_his.Close
        OpenRecordset StrSql, rs_his
        If Not rs_his.EOF Then
            If Pisa Then
                If rs_his!estrnro = CodEst Then

                    StrSql = " UPDATE his_estructura SET htetdesde = " & cambiaFecha(FAlta)
                    StrSql = StrSql & ",htethasta = " & cambiaFecha(FBaja)
                    StrSql = StrSql & " WHERE tenro = " & TipoEstr
                    StrSql = StrSql & " AND ternro = " & CodTer
                    StrSql = StrSql & " AND estrnro = " & rs_his!estrnro
                    StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_his!htetdesde)
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Actualiza -- Actualizo Tipo de Estructura " & TipoEstr & "p/ el tercero: " & CodTer
    
                Else

                    StrSql = " UPDATE his_estructura SET "
                    StrSql = StrSql & " estrnro = " & CodEst
                    StrSql = StrSql & ",htetdesde = " & cambiaFecha(FAlta)
                    StrSql = StrSql & " WHERE tenro = " & TipoEstr
                    StrSql = StrSql & " AND ternro = " & CodTer
                    StrSql = StrSql & " AND estrnro = " & rs_his!estrnro
                    StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_his!htetdesde)
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'Else
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Actualiza -- Actualizo Tipo de Estructura " & TipoEstr & "p/ el tercero: " & CodTer
    
                End If
            Else ' no Pisa
                'FGZ - 23/07/2010 -  se agregó este control
                If rs_his!estrnro = CodEst Then
                    If Not UCase(FBaja) = "NULL" Then
                        StrSql = " UPDATE his_estructura SET "
                        StrSql = StrSql & "htethasta = " & cambiaFecha(FBaja)
                        StrSql = StrSql & " WHERE tenro = " & TipoEstr
                        StrSql = StrSql & " AND ternro = " & CodTer
                        StrSql = StrSql & " AND estrnro = " & rs_his!estrnro
                        StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_his!htetdesde)
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        Flog.writeline Espacios(Tabulador * 1) & "Actualiza -- Actualizo Tipo de Estructura " & TipoEstr & "p/ el tercero: " & CodTer
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "Sin Accion p/ Tipo de Estructura " & TipoEstr & "p/ el tercero: " & CodTer
                        'Nada

                    End If
                Else
                    
                    If EsNulo(rs_his!htethasta) Then
                        
                        If (FAlta) = ConvFecha(rs_his!htetdesde) Then
                            ' si la fecha es = Reeemplazar la estructura anterior
                            
                            StrSql = " UPDATE his_estructura SET "
                            StrSql = StrSql & " estrnro = " & CodEst & ", "
                            StrSql = StrSql & " htethasta = " & cambiaFecha(FBaja)
                            StrSql = StrSql & " WHERE tenro = " & TipoEstr
                            StrSql = StrSql & " AND ternro = " & CodTer
                            StrSql = StrSql & " AND estrnro = " & rs_his!estrnro
                            StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_his!htetdesde)
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Flog.writeline Espacios(Tabulador * 1) & "Actualiza -- Actualizo Tipo de Estructura " & TipoEstr & "p/ el tercero: " & CodTer
                        
                        Else
                        
                            StrSql = " UPDATE his_estructura SET "
                            StrSql = StrSql & "htethasta = " & ConvFecha(F_Cierre_Temp)
                            StrSql = StrSql & " WHERE tenro = " & TipoEstr
                            StrSql = StrSql & " AND ternro = " & CodTer
                            StrSql = StrSql & " AND estrnro = " & rs_his!estrnro
                            StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_his!htetdesde)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        
                    
                            'FGZ - 23/07/2010 - se cambió esta parte
                            If UCase(FBaja) = "NULL" Then
                                StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
                                StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & cambiaFecha(FAlta) & ")"
                                objConn.Execute StrSql, , adExecuteNoRecords
                            Else
                                StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                                StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & cambiaFecha(FAlta) & "," & cambiaFecha(FBaja) & ")"
                                objConn.Execute StrSql, , adExecuteNoRecords
                            End If
                        
                            Flog.writeline Espacios(Tabulador * 1) & "Inserto  Tipo de Estructura " & TipoEstr & "p/ el tercero: " & CodTer
                        
                        End If
                        
                    Else

                        Flog.writeline Espacios(Tabulador * 1) & "Ya existe una estructura de tipo  " & TipoEstr
                    End If
                End If
            End If
        Else
            ' ver si la Fecha de Alta es menor que la fecha desde del Tipo de Estructura.
            StrSql = "SELECT * FROM his_estructura "
            StrSql = StrSql & " WHERE tenro=" & TipoEstr
            StrSql = StrSql & "   AND ternro=" & CodTer
            StrSql = StrSql & "   AND htetdesde > " & cambiaFecha(FAlta)
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            
            If Not rs.EOF Then
                Flog.writeline Espacios(Tabulador * 1) & "Ya existe una estructura de tipo " & TipoEstr & " con Fecha de inicio mayor a la Fecha de Alta."
            Else
            
                'FGZ - 23/07/2010 - se cambió esta parte
                If UCase(FBaja) = "NULL" Then
                    StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
                    StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & cambiaFecha(FAlta) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                    StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & cambiaFecha(FAlta) & "," & cambiaFecha(FBaja) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                                    
                Flog.writeline Espacios(Tabulador * 1) & "Inserto el Tipo de Estructura " & TipoEstr & " para el tercero: " & CodTer
            End If
            rs.Close
            
            
        End If
        
    
    Else ' de CodEst <> 0
        'Flog.writeline Espacios(Tabulador * 1) & " Al cargar el tipo de Estructura " & TipoEstr & " - CodEst = 0 "
    End If
    
    If rs_his.State = adStateOpen Then rs_his.Close
    Set rs_his = Nothing

End Sub

Public Sub Complementos1(CodTer As Long, codEstr As Long)

    StrSql = " INSERT INTO sucursal(estrnro,ternro,sucest) VALUES(" & codEstr & "," & CodTer & ",-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos3(CodTer As Long, codEstr As Long)

    StrSql = " INSERT INTO categoria(estrnro,convnro) VALUES(" & codEstr & "," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos4(codEstr As Long, Valor As String)

    StrSql = " INSERT INTO puesto(estrnro,puedesc,puenroreemp) VALUES(" & codEstr & ",'" & Valor & "',0)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos10(CodTer As Long, codEstr As Long, Valor As String)

    StrSql = " INSERT INTO empresa(estrnro,ternro,empnom) VALUES(" & codEstr & "," & CodTer & ",'" & Valor & "')"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos15(CodTer As Long, codEstr As Long)

    ' Hay que crear un Tipo de Caja de Jubilacion "Migracion"

    StrSql = " INSERT INTO cajjub(estrnro,ternro,cajest,ticnro) VALUES(" & codEstr & "," & CodTer & ",-1,1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos16(CodTer As Long, codEstr As Long)

    StrSql = " INSERT INTO gremio(estrnro,ternro) VALUES(" & codEstr & "," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos17(CodTer As Long, codEstr As Long, Valor As String)
' Ultima Modificacion:  FGZ
' Fecha:                17/12/2004
'---------------------------------------------------------
Dim rs_17 As New ADODB.Recordset

    StrSql = "SELECT * FROM osocial  where osdesc = '" & Valor & "'"
    If rs_17.State = adStateOpen Then rs_17.Close
    OpenRecordset StrSql, rs_17
    
    If rs_17.EOF Then
        StrSql = " INSERT INTO osocial(ternro,osdesc) VALUES(" & CodTer & ",'" & Valor & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    StrSql = "SELECT * FROM replica_estr  where origen = " & CodTer
    StrSql = StrSql & " AND estrnro = " & codEstr
    If rs_17.State = adStateOpen Then rs_17.Close
    OpenRecordset StrSql, rs_17
    If rs_17.EOF Then
        StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & CodTer & "," & codEstr & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    If rs_17.State = adStateOpen Then rs_17.Close
    Set rs_17 = Nothing
End Sub

Public Sub Complementos18(CodTer As Long, codEstr As Long, Valor As String)
Dim rs_tipocont As New ADODB.Recordset
Dim rs_TC As New ADODB.Recordset
Dim CodTC As Long


    
    StrSql = "SELECT * FROM tipocont  where tcdabr = '" & Valor & "'"
    OpenRecordset StrSql, rs_tipocont
    
    If rs_tipocont.EOF Then
        '22-11-06 -Diego Rosso - se agregaron los campos tcdesc(se pone = a tcdabr) y leynro
        StrSql = " INSERT INTO tipocont(tcdabr,estrnro,tcind,tcdesc,leynro) VALUES('" & Valor & "'," & codEstr & ",-1,'" & Valor & "',1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        CodTC = getLastIdentity(objConn, "tipocont")
        
        'StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & CodTC & "," & CodEstr & ")"
        'objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub

Public Sub Complementos19(codEstr As Long)

    StrSql = " INSERT INTO convenios(estrnro) VALUES(" & codEstr & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos22(CodTer As Long, codEstr As Long, Valor As String)

    StrSql = " INSERT INTO formaliq(estrnro,folisistema) VALUES(" & codEstr & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos23(CodTer As Long, codEstr As Long, Valor As String)

Dim rs_pos As New ADODB.Recordset
Dim CodPlan As Long

    ' Hay que ver la relacion entra la Osocial y el Plan

    StrSql = " INSERT INTO planos(plnom,osocial) VALUES('" & Valor & "'," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    CodPlan = getLastIdentity(objConn, "planos")
    
    StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & CodPlan & "," & codEstr & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    

End Sub

Public Sub Complementos28(CodTer As Long, codEstr As Long, Valor As String)

    StrSql = " INSERT INTO agencia(estrnro,ternro,agedes,ageest) VALUES(" & codEstr & "," & CodTer & ",'" & Valor & "'" & ",-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub
Public Sub Complementos40(codEstr As Long, CodTer As Long, Valor As String)

    StrSql = " INSERT INTO seguro(ternro,estrnro,segdesc,segest) VALUES(" & codEstr & "," & CodTer & ",'" & Valor & "',-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos41(codEstr As Long, CodTer As Long, Valor As String)

    StrSql = " INSERT INTO banco(ternro,estrnro,bansucdesc,banest) VALUES(" & codEstr & "," & CodTer & ",'" & Valor & "',-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub
