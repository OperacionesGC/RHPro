Attribute VB_Name = "mdlValidar"
Function existeEmpleado(ByVal legajo As String)
'Dado un legajo si el empleado existe devuelve el nro de tercero sino devuelve 0
Dim rs_Empleado As New ADODB.Recordset
Dim ternro As Long
    'Controlo si el empleado existe o si hay que crearlo
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & legajo
    OpenRecordset StrSql, rs_Empleado
    If Not rs_Empleado.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "El empleado leg: " & legajo & " existe en el sistema."
        ternro = rs_Empleado!ternro
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado leg: " & legajo & " No existe en el sistema."
        ternro = 0
    End If
    existeEmpleado = ternro
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
End Function

Function crearNacionalidad(ByVal nacionalidadDesc As String)
    StrSql = " INSERT INTO nacionalidad (nacionaldes,nacionaldefault) VALUES ('" & nacionalidadDesc & "',0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    crearNacionalidad = getLastIdentity(objConn, "nacionalidad")
End Function

Function crearEstadoCivil(ByVal estadoCivilDesc As String)
    StrSql = " INSERT INTO estcivil (estcivdesabr) VALUES ('" & estadoCivilDesc & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    crearEstadoCivil = getLastIdentity(objConn, "estcivil")
End Function

Sub verificarLocalidad(ByVal Localidad As String, ByRef locnro As String)
Dim rs_datos As New ADODB.Recordset

    StrSql = " SELECT locnro FROM localidad WHERE upper(locdesc) = '" & UCase(Mid(Localidad, 1, 60)) & "'"
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        locnro = rs_datos!locnro
        Flog.writeline Espacios(Tabulador * 1) & "Localidad encontrada."
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Localidad no existe, se creara."
        StrSql = " INSERT INTO localidad (locdesc) VALUES ('" & Localidad & "') "
        objConn.Execute StrSql, , adExecuteNoRecords
        locnro = getLastIdentity(objConn, "localidad")
    End If
    
End Sub

Function crearPartido(ByVal partidoDesc As String)
    StrSql = " INSERT INTO partido (partnom) VALUES ('" & partidoDesc & "') "
    objConn.Execute StrSql, , adExecuteNoRecords
    crearPartido = getLastIdentity(objConn, "partido")
End Function

Sub crearPais(ByVal Pais As String, ByVal Provincia As String, ByRef paisnro As String, ByRef provnro As String)
Dim rs_datos As New ADODB.Recordset

    If Len(Pais) > 0 Then
        'chequeo si el pais existe
        StrSql = "SELECT paisnro FROM pais WHERE upper(paisdesc) = '" & UCase(Mid(Pais, 1, 60)) & "'"
        OpenRecordset StrSql, rs_datos
        If rs_datos.EOF Then
            'si pais es 0 tengo que crearlo
            StrSql = " INSERT INTO pais (paisdesc) VALUES ('" & Pais & "') "
            objConn.Execute StrSql, , adExecuteNoRecords
            paisnro = getLastIdentity(objConn, "partido")
        Else
            paisnro = rs_datos!paisnro
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Pais no informado."
    End If
    
    If Len(Provincia) > 0 Then
        StrSql = " SELECT provnro FROM provincia WHERE upper(provdesc) = '" & UCase(Left(Provincia, 60)) & "'"
        OpenRecordset StrSql, rs_datos
        If rs_datos.EOF Then
            StrSql = " INSERT INTO provincia (provdesc,paisnro) VALUES ('" & Provincia & "'," & paisnro & ") "
            objConn.Execute StrSql, , adExecuteNoRecords
            provnro = getLastIdentity(objConn, "provincia")
        Else
            provnro = rs_datos!provnro
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Provincia no informada."
    End If
        
    'If CStr(zona) = "0" Then
    '    StrSql = " INSERT INTO zona (zonadesc,provnro) VALUES ('" & zonadesc & "'," & provincia & ") "
    '    objConn.Execute StrSql, , adExecuteNoRecords
    '    zona = getLastIdentity(objConn, "zona")
    'End If
    If rs_datos.State = adStateOpen Then rs_datos.Close
    Set rs_datos = Nothing
End Sub

Sub crearEmpleado(ByRef ternro As Long, ByRef FecIng As String, ByRef nacionalidad As String, ByRef Pais As String, ByVal legajo As String, ByVal estado As String, ByVal email As String, ByVal remuneracion As String, ByVal fecnac As String, ByVal sexo As String, ByVal EstCiv As String, ByVal nombre As String, ByVal nombre2 As String, ByVal apellido As String, ByVal reportaA As String)
     StrSql = " INSERT INTO tercero(ternom,ternom2,terape,terfecnac,tersex,estcivnro,terfecing,nacionalnro,paisnro)"
     StrSql = StrSql & " VALUES('" & nombre & "','" & nombre2 & "','" & apellido & "'," & ConvFecha(fecnac) & "," & sexo & "," & EstCiv & ","
     If UCase(FecIng) <> "N/A" And UCase(FecIng) <> "" Then
         StrSql = StrSql & ConvFecha(FecIng) & ","
     Else
         StrSql = StrSql & "null,"
     End If
     StrSql = StrSql & nacionalidad & ","
     StrSql = StrSql & Pais & ")"
     objConn.Execute StrSql, , adExecuteNoRecords
     
     Flog.writeline Espacios(Tabulador * 1) & "Inserto en la tabla tercero."
     ternro = getLastIdentity(objConn, "tercero")
     Flog.writeline Espacios(Tabulador * 1) & "Nuevo numero de tercero: " & ternro & "."
     
     'inserto el ter_tip correspondiente a empleado (1)
     StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & ternro & ",1)"
     objConn.Execute StrSql, , adExecuteNoRecords
     
     'Inserto el empleado
     StrSql = " INSERT INTO empleado(empleg,empfecalta,empest,"
     StrSql = StrSql & "ternro,terape,ternom,empemail, empremu"
     If CLng(reportaA) <> 0 Then
         StrSql = StrSql & ",empreporta"
     End If
    
     StrSql = StrSql & ") VALUES("
     StrSql = StrSql & legajo & ","
     
     If UCase(FecIng) <> "" Then
         StrSql = StrSql & ConvFecha(FecIng) & ","
     Else
         StrSql = StrSql & " null,"
     End If
     
     StrSql = StrSql & estado & ","
     
     StrSql = StrSql & ternro & ",'" & apellido & "','" & nombre & "'"
     StrSql = StrSql & ",'" & email & "'," & remuneracion
     If CLng(reportaA) <> 0 Then
         StrSql = StrSql & "," & reportaA
     End If
     StrSql = StrSql & ")"
     objConn.Execute StrSql, , adExecuteNoRecords
     Flog.writeline Espacios(Tabulador * 1) & "Inserto en la tabla empleado."
     
End Sub

Sub actualizarEmpleado(ByRef ternro As Long, ByRef FecIng As String, ByRef nacionalidad As String, ByRef Pais As String, ByVal legajo As String, ByVal estado As String, ByVal email As String, ByVal remuneracion As String, ByVal fecnac As String, ByVal sexo As String, ByVal EstCiv As String, ByVal nombre As String, ByVal nombre2 As String, ByVal apellido As String, ByVal reportaA As String)
    
    Flog.writeline Espacios(Tabulador * 1) & "Comienza el analisis del empleado."

    StrSql = " UPDATE tercero SET " & _
            " ternom = '" & nombre & "'," & _
            " ternom2 = '" & nombre2 & "'," & _
            " terape = '" & apellido & "'," & _
            " terfecnac = " & ConvFecha(fecnac) & "," & _
            " tersex = " & sexo & "," & _
            " estcivnro = " & EstCiv
    
    If UCase(FecIng) <> "" Then
        StrSql = StrSql & " ,terfecing = " & ConvFecha(FecIng)
    Else
        StrSql = StrSql & " ,terfecing = null "
    End If
    
    If UCase(nacionalidad) <> "" Then
        StrSql = StrSql & ", nacionalnro = " & nacionalidad
    End If
    
    
    StrSql = StrSql & " ,paisnro = " & Pais & " WHERE ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Actualizada la tabla tercero."
    
    'empleado existe, actualizo
    StrSql = " UPDATE empleado SET " & _
             " empleg = " & legajo
    
    If UCase(FecAlta) <> "" Then
        StrSql = StrSql & " ,empfecalta = " & ConvFecha(FecIng)
    Else
        StrSql = StrSql & " ,empfecalta = null "
    End If
    
    If UCase(FecAlta) <> "" Then
        StrSql = StrSql & " ,empest = " & estado
    End If
    
    StrSql = StrSql & " ,terape = '" & apellido & "'" & _
             " ,ternom = '" & nombre & "'" & _
             " ,empemail = '" & email & "'"
    
    If CStr(remuneracion) <> "" Then
        StrSql = StrSql & " ,empremu = " & remuneracion
    End If
    
    If CStr(reportaA) <> "" Then
        StrSql = StrSql & ",empreporta = " & reportaA
    End If
    
    StrSql = StrSql & " WHERE ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Actualizada la tabla empleado."
    
    'Actualizo la fase segun la fecha de ingreso informada
'    StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
'    StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
'    StrSql = StrSql & " VALUES(" & ternro & ",Null," & cambiaFecha(fecAlta) & "," & cambiaFecha(fecBaja) & ","
'    StrSql = StrSql & "-1,0,-1,-1,-1,-1,-1)"
    Flog.writeline Espacios(Tabulador * 1) & "Fin actualizacion del empleado."

End Sub

Sub actualizarFase(ByVal ternro As String, ByVal FecAlta As String, ByVal fechaBaja As String)
Dim NroCausa As Integer
NroCausa = 0 'en el archivo de entrada no se informa nro causa
    Flog.writeline Espacios(Tabulador * 1) & "Comienza el analisis de fases."
    'Si no existe fase ==> simplemente crea la fase
    StrSql = "SELECT * FROM fases WHERE empleado = " & ternro
    StrSql = StrSql & " ORDER BY altfec DESC"
    OpenRecordset StrSql, rs
    If rs.EOF Then
        If NroCausa <> 0 Then
            StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
            StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
            StrSql = StrSql & " VALUES(" & ternro & "," & NroCausa & "," & cambiaFecha(FecAlta) & "," & cambiaFecha(fecBaja) & ","
            StrSql = StrSql & "-1,0,-1,-1,-1,-1,-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
            StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
            StrSql = StrSql & " VALUES(" & ternro & ",Null," & cambiaFecha(FecAlta) & "," & cambiaFecha(fecBaja) & ","
            StrSql = StrSql & "-1,0,-1,-1,-1,-1,-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
    Else    'Ya tiene fases que arranca en esa fecha ==> Actualizo
        StrSql = "SELECT * FROM fases "
        StrSql = StrSql & " WHERE empleado = " & ternro
        StrSql = StrSql & " AND altfec = " & cambiaFecha(FecAlta)
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            'Si la fecha hasta del registro encontrado es mayor o nulo ==> actualizo
            '   sino error
            If EsNulo(rs!bajfec) Then
                'Actualizo
                StrSql = " UPDATE fases SET "
                StrSql = StrSql & " bajfec = " & cambiaFecha(fecBaja)
                If NroCausa <> 0 Then
                    StrSql = StrSql & ",caunro = " & NroCausa
                End If
                StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
                objConn.Execute StrSql, , adExecuteNoRecords
                
            Else
                If UCase(fecBaja) = "NULL" Then
                    'no debe existir ningun otro registro sino Error
                    StrSql = "SELECT fasnro FROM fases "
                    StrSql = StrSql & " WHERE empleado = " & ternro
                    StrSql = StrSql & " AND altfec > " & cambiaFecha(FecAlta)
                    StrSql = StrSql & " AND fasnro <> " & rs!fasnro
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        Flog.writeline Espacios(Tabulador * 1) & "La fase se superpone con otra ya existente. Ternro: " & ternro & " Fecha Alta: " & FecAlta
                        Exit Sub
                    Else
                        'Actualizo
                        StrSql = " UPDATE fases SET "
                        StrSql = StrSql & " bajfec = " & cambiaFecha(fecBaja)
                        If NroCausa <> 0 Then
                            StrSql = StrSql & ",caunro = " & NroCausa
                        End If
                        StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                    End If
                Else
                    If CDate(rs!bajfec) >= CDate(FBaja) Then
                        'Actualizo
                        StrSql = " UPDATE fases SET "
                        StrSql = StrSql & " bajfec = " & cambiaFecha(fecBaja)
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
                    StrSql = StrSql & " AND bajfec >= " & cambiaFecha(FecAlta)
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
                          StrSql = StrSql & " VALUES(" & ternro & "," & NroCausa & "," & cambiaFecha(FecAlta) & "," & cambiaFecha(fecBaja) & ","
                          StrSql = StrSql & "-1,0,-1,-1,-1,-1,0)"
                          objConn.Execute StrSql, , adExecuteNoRecords
    
                        Else
                          StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
                          StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                          StrSql = StrSql & " VALUES(" & ternro & ",Null," & cambiaFecha(FecAlta) & "," & cambiaFecha(fecBaja) & ","
                          StrSql = StrSql & "-1,0,-1,-1,-1,-1,0)"
                          objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                        
                    End If
                End If
            Else
                'SI existe fases que cruce la nueva ==>
                StrSql = "SELECT fasnro FROM fases "
                StrSql = StrSql & " WHERE empleado = " & ternro
                StrSql = StrSql & " AND ("
                StrSql = StrSql & " (altfec <= " & cambiaFecha(FecAlta) & " AND bajfec >=" & cambiaFecha(fecBaja) & ")"
                StrSql = StrSql & " OR "
                StrSql = StrSql & " (altfec <= " & cambiaFecha(FecAlta) & " AND bajfec >=" & cambiaFecha(FecAlta) & ")"
                StrSql = StrSql & " OR "
                StrSql = StrSql & " (altfec >= " & cambiaFecha(FecAlta) & " AND bajfec <=" & cambiaFecha(fecBaja) & ")"
                StrSql = StrSql & " OR "
                StrSql = StrSql & " (altfec >= " & cambiaFecha(FecAlta) & " AND altfec <= " & cambiaFecha(fecBaja) & " AND bajfec >=" & cambiaFecha(fecBaja) & ")"
                StrSql = StrSql & " OR "
                StrSql = StrSql & " (altfec >= " & cambiaFecha(FecAlta) & " AND altfec <= " & cambiaFecha(fecBaja) & " AND bajfec IS NULL)"
                StrSql = StrSql & " OR "
                StrSql = StrSql & " (altfec <= " & cambiaFecha(FecAlta) & " AND bajfec IS NULL)"
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
                        StrSql = StrSql & " VALUES(" & ternro & "," & NroCausa & "," & cambiaFecha(FecAlta) & "," & cambiaFecha(fecBaja) & ","
                        StrSql = StrSql & "-1,0,-1,-1,-1,-1,0)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
                        StrSql = StrSql & " sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                        StrSql = StrSql & " VALUES(" & ternro & ",Null," & cambiaFecha(FecAlta) & "," & cambiaFecha(fecBaja) & ","
                        StrSql = StrSql & "-1,0,-1,-1,-1,-1,0)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
    
                End If
            End If
        End If
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Fin del analisis de Fases."

End Sub

Sub actualizarDocumento(ByVal ternro As String, ByVal tipoDocumento As String, ByVal nroDocumento As String)
Dim rs_datos As New ADODB.Recordset
    Flog.writeline Espacios(Tabulador * 1) & "Comienza el analisis de Documentos."
    StrSql = " SELECT tidnro,ternro FROM ter_doc WHERE ternro = " & ternro & " AND tidnro = " & tipoDocumento
    OpenRecordset StrSql, rs_datos
    If rs_datos.EOF Then
        StrSql = " INSERT INTO ter_doc (tidnro,ternro,nrodoc) VALUES " & _
                 " (" & tipoDocumento & "," & ternro & ",'" & nroDocumento & "') "
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Documento insertado."
    Else
        StrSql = " UPDATE ter_doc SET " & _
                 " nrodoc = '" & nroDocumento & "'" & _
                 " WHERE tidnro = " & rs_datos!tidnro & " AND ternro = " & ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Documento actualizado."
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Fin del analisis de Documentos."
    If rs_datos.State = adStateOpen Then rs_datos.Close
End Sub

Sub actualizarDomicilio(ByVal ternro As String, ByVal Pais As String, ByVal Provincia As String, ByVal zona As String, ByVal partido As String, ByVal Localidad As String, ByVal Barrio As String, ByVal entreCalles As String, ByVal cp As String, ByVal manzana As String, ByVal torre As String, ByVal depto As String, ByVal piso As String, ByVal Numero As String, ByVal calle As String, ByRef domnro As String)
Dim rs_datos As New ADODB.Recordset
    
    Flog.writeline Espacios(Tabulador * 1) & "Comienzo del analisis de Domicilio."
    StrSql = " SELECT * FROM cabdom  WHERE tipnro = 1 AND ternro = " & ternro & " AND domdefault = -1 AND tidonro = 2"
    OpenRecordset StrSql, rs_datos
   
    If rs_datos.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe domicilio para el empleado, se creara."
        'No existe el domicilio para el empleado
        StrSql = " INSERT INTO cabdom (tipnro,ternro,domdefault,tidonro,modnro) VALUES (1," & ternro & ",-1,2,1) "
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline Espacios(Tabulador * 1) & "Cabecera de domicilio creada."
        domnro = getLastIdentity(objConn, "cabdom")
        
        StrSql = " INSERT INTO detdom (domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal, " & _
                 " entrecalles,barrio"
        If Localidad <> "" Then
            StrSql = StrSql & ",locnro"
        End If
        If partido <> "" Then
            StrSql = StrSql & ",partnro"
        End If
        If zona <> "" Then
            StrSql = StrSql & ",zonanro"
        End If
        StrSql = StrSql & " ,provnro,paisnro) VALUES"
        
        StrSql = StrSql & " (" & domnro & ",'" & calle & "','" & Numero & "','" & piso & "','" & depto & "',"
        StrSql = StrSql & " '" & torre & "','" & manzana & "','" & cp & "','" & entreCalles & "','" & Barrio & "'"
        StrSql = StrSql & "," & Localidad
        
        If partido <> "" Then
            StrSql = StrSql & "," & partido
        End If
        If zona <> "" Then
            StrSql = StrSql & "," & zona
        End If
        StrSql = StrSql & "," & Provincia & "," & Pais & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Detalle de domicilio creado."
    
    Else
        domnro = rs_datos!domnro
        'ya tiene domicilio
        StrSql = " UPDATE detdom SET "
        If calle <> "" Then
            StrSql = StrSql & " calle = '" & calle & "'"
        End If
        
        StrSql = StrSql & " ,nro = '" & Numero & "'"
        StrSql = StrSql & " ,piso = '" & piso & "'"
        StrSql = StrSql & " ,oficdepto = '" & depto & "'"
        StrSql = StrSql & " ,torre = '" & torre & "'"
        StrSql = StrSql & " ,manzana = '" & manzana & "'"

        If cp <> "" Then
            StrSql = StrSql & " ,codigopostal = '" & cp & "'"
        End If
        
        StrSql = StrSql & " ,entrecalles = '" & entreCalles & "'"
        StrSql = StrSql & " ,barrio = '" & Barrio & "'"
        StrSql = StrSql & " ,locnro = " & Localidad
        
        If partido <> "" Then
            StrSql = StrSql & " ,partnro = " & partido
        End If
        If zona <> "" Then
            StrSql = StrSql & " ,zonanro = " & zona
        End If
        If Provincia <> "" Then
            StrSql = StrSql & " ,provnro = " & Provincia
        End If
        
        StrSql = StrSql & " ,paisnro = " & Pais
        StrSql = StrSql & " WHERE domnro = " & domnro

        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Domicilio Actualizado."
        
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Fin analisis de Domicilio."
    
    If rs_datos.State = adStateOpen Then rs_datos.Close
End Sub

Sub actualizarTelefono(ByVal nroTelefono As String, ByVal tipoTelefono As String, ByVal nrodom As String)
Dim rs_tel As New ADODB.Recordset

    Flog.writeline Espacios(Tabulador * 1) & "Comienza analisis de Telefono."
    If CStr(nroTelefono) <> "0" Then
        'Chequeo si existe telefono particular
        StrSql = " SELECT domnro FROM Telefono WHERE domnro = " & nrodom & " AND tipotel =  " & tipoTelefono
        OpenRecordset StrSql, rs_tel
        If rs_tel.EOF Then
            StrSql = " INSERT INTO telefono (domnro,telnro,telfax,teldefault,telcelular,tipotel) VALUES " & _
                     " (" & nrodom & ",'" & nroTelefono & "',0,-1,0," & tipoTelefono & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Telefono particular, creado."
        Else
            StrSql = " UPDATE telefono SET " & _
                     " telnro = '" & nroTelefono & "'," & _
                     " telfax = 0, teldefault = -1 ,telcelular = 0 " & _
                     " WHERE domnro = " & nrodom & " AND tipotel = " & tipoTelefono
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Telefono particular, actualizado."
        End If
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Fin analisis de Telefono."
    If rs_tel.State = adStateOpen Then rs_tel.Close
End Sub

Sub actualizarSalario(ByVal ternro As String, ByVal salario As Double)
    StrSql = " UPDATE empleado SET empremu = " & salario & " WHERE ternro = " & ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 1) & "Salario Actualizado."
End Sub



Function calcularFase(ByVal ternro As String, ByVal separador As String)
Dim rs_fases As New ADODB.Recordset
Dim strLinea As String
    
    StrSql = " SELECT altfec FROM fases WHERE empleado = " & ternro & " ORDER BY altfec ASC "
    OpenRecordset StrSql, rs_fases
    If Not rs_fases.EOF Then
        strLinea = rs_fases!altfec
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado no posee fases."
        strLinea = ""
    End If
    
    If rs_fases.State = adStateOpen Then rs_fases.Close
    calcularFase = strLinea
End Function

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

Public Sub VerSiCrearTercero(tenro As String, Valor As String, ByRef CodTer)


  Select Case tenro

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

Public Sub VerSiCrearComplemento(tenro As String, codEstr As Long, Valor As String, CodTer As Long)

  Select Case tenro

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
'               la estructura ser� actualizada.
' ---------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset
Dim rs_his As New ADODB.Recordset

Dim F_Cierre_Temp As Date


    F_Cierre_Temp = DateAdd("d", -1, CDate(FAlta))
    pisa = -1
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
            If pisa Then
                If rs_his!estrnro = CodEst Then

                    StrSql = " UPDATE his_estructura SET htetdesde = " & cambiaFecha(FAlta)
                    StrSql = StrSql & ",htethasta = " & cambiaFecha(FBaja)
                    StrSql = StrSql & " WHERE tenro = " & TipoEstr
                    StrSql = StrSql & " AND ternro = " & CodTer
                    StrSql = StrSql & " AND estrnro = " & rs_his!estrnro
                    StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_his!htetdesde)
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Actualiza -- Actualizo Tipo de Estructura " & TipoEstr & " p/ el tercero: " & CodTer
    
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
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Actualiza -- Actualizo Tipo de Estructura " & TipoEstr & " p/ el tercero: " & CodTer
    
                End If
            Else ' no Pisa
                'FGZ - 23/07/2010 -  se agreg� este control
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
                        
                        If ConvFecha(FAlta) = ConvFecha(rs_his!htetdesde) Then
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
                        
                    
                            'FGZ - 23/07/2010 - se cambi� esta parte
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
            
                'FGZ - 23/07/2010 - se cambi� esta parte
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


Public Function obtenerEstructura(ByVal ternro As Long, ByVal tenro As Long, ByVal fecha As String)
Dim rs_estruct As New ADODB.Recordset
Dim Salida As String
    
    StrSql = " SELECT estrdabr FROM his_estructura " & _
             " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro AND estructura.tenro = " & tenro & _
             " WHERE ternro = " & ternro & " AND (htetdesde <= " & ConvFecha(Date) & " AND (htethasta >= " & ConvFecha(Date) & " or htethasta is null))"
    OpenRecordset StrSql, rs_estruct
    If Not rs_estruct.EOF Then
        Salida = rs_estruct("estrdabr")
    Else
        Salida = ""
    End If
    
    obtenerEstructura = Salida
    
    If rs_estruct.State = adStateOpen Then rs_estruct.Close
    Set rs_estruct = Nothing

End Function


Function armarDireccion(ByVal ternro As Long, ByVal separador As String)
'llega con un separador al principio
Dim rs_dire As New ADODB.Recordset
Dim Salida As String

    StrSql = " SELECT calle,nro,piso,oficdepto,torre, manzana, entrecalles, barrio, locdesc, provdesc " & _
             " ,codigopostal, paisdesc, tel_part.telnro telpart, tel_cel.telnro telcel FROM cabdom " & _
             " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
             " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
             " INNER JOIN provincia ON provincia.provnro = detdom.provnro " & _
             " INNER JOIN pais ON pais.paisnro = detdom.paisnro " & _
             " LEFT JOIN telefono tel_part ON tel_part.domnro = cabdom.domnro AND tel_part.tipotel = 1 " & _
             " LEFT JOIN telefono tel_cel ON tel_cel.domnro = cabdom.domnro AND tel_part.tipotel = 2 " & _
             " WHERE ternro = " & ternro & " AND domdefault = -1 "
    
    OpenRecordset StrSql, rs_dire
    If Not rs_dire.EOF Then
        'comienza la primer linea de direccion
        Salida = rs_dire("calle") & ";" & rs_dire("nro") & ";" & rs_dire("piso")
        Salida = Salida & ";" & rs_dire("oficdepto") & ";" & rs_dire("torre")
        'comienza la segunda linea de direccion
        Salida = Salida & separador
        Salida = Salida & rs_dire("manzana") & ";" & rs_dire("entrecalles")
        'comienza la tercera linea de direccion
        Salida = Salida & separador
        Salida = Salida & rs_dire("barrio") & ";" & rs_dire("locdesc")
        'provincia / pais
        Salida = Salida & separador
        Salida = Salida & rs_dire("provdesc") & ";" & rs_dire("paisdesc")
        'codigo postal
        Salida = Salida & separador & rs_dire("codigopostal")
        'pais de domicilio
        Salida = Salida & separador & rs_dire("paisdesc")
        'telefono particular
        Salida = Salida & separador & IIf(EsNulo(rs_dire("telpart")), "", rs_dire("telpart"))
        'telefono celular
        Salida = Salida & separador & IIf(EsNulo(rs_dire("telcel")), "", rs_dire("telcel"))
        
    Else
        Salida = ";;;;"
        'comienza la segunda linea de direccion
        Salida = Salida & separador
        Salida = Salida & ";"
        'comienza la tercera linea de direccion
        Salida = Salida & separador
        Salida = Salida & ";"
        'provincia / pais
        Salida = Salida & separador
        Salida = Salida & ";"
        'codigo postal
        Salida = Salida & separador
        'pais de domicilio
        Salida = Salida & separador
        'telefono particular
        Salida = Salida & separador
        'telefono celular
        Salida = Salida & separador
    End If

    armarDireccion = Salida
    
    If rs_dire.State = adStateOpen Then rs_dire.Close
    Set rs_dire = Nothing

End Function

Sub bajaContrato(ByVal ternro As Long, ByVal fechaBajaPrevista As String)
Dim rs_contrato As New ADODB.Recordset
    
    'busco el contrato activo del empleado
    StrSql = " SELECT estrnro FROM his_estructura " & _
             " WHERE tenro = 18 AND htethasta is null AND ternro = " & ternro
    OpenRecordset StrSql, rs_contrato
    
    If Not rs_contrato.EOF Then
        'encontramos un contrato activo, cerramos la estructura
        StrSql = " UPDATE his_estructura SET htethasta = " & cambiaFecha(fechaBajaPrevista) & _
                 " WHERE ternro = " & ternro & " AND tenro = 18 AND htethasta is null AND estrnro = " & rs_contrato!estrnro
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "Fecha de vencimiento de contrato actualizada."
    Else
        'si no busco el contrato mas nuevo que tuvo el empleado
        StrSql = " SELECT htethasta, estrnro FROM his_estructura " & _
                 " WHERE tenro = 18  AND ternro = " & ternro & " ORDER BY htethasta DESC "
        OpenRecordset StrSql, rs_contrato
        If Not rs_contrato.EOF Then
            'encontramos un contrato activo, cerramos la estructura
            StrSql = " UPDATE his_estructura SET htethasta = " & cambiaFecha(fechaBajaPrevista) & _
                     " WHERE ternro = " & ternro & " AND tenro = 18 AND estrnro = " & rs_contrato!estrnro & " AND htethasta = " & rs_contrato!htethasta
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "Fecha de vencimiento de contrato actualizada."
        Else
            Flog.writeline Espacios(Tabulador * 1) & "El empleado no tiene contrato, no es posible acutalizar la fecha de vencimiento de contrato."
            Exit Sub
        End If
    End If
    
    If rs_contrato.State = adStateOpen Then rs_contrato.Close
    Set rs_contrato = Nothing

End Sub

Function armarfecha(ByVal fecha As String)

    If fecha <> "" Then
        armarfecha = Left(fecha, 2) & "/" & Mid(fecha, 3, 2) & "/" & Right(fecha, 4)
    Else
        armarfecha = ""
    End If
End Function
