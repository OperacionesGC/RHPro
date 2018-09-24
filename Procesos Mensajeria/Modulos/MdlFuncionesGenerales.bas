Attribute VB_Name = "MdlFuncionesGenerales"
Public Function StrToStr(cadena As String, Longitud As Integer)
    StrToStr = CStr(Left(Trim(CStr(cadena)), Longitud))
End Function
Public Function StrToInt(cadena As String) As Integer
    'On Error GoTo cero:
    StrToInt = CInt(cadena)
'cero:
'    StrToInt = 0
End Function
Public Function StrToFec(cadena As String)
    On Error GoTo cero:
    StrToFec = ConvFecha(cadena)
cero:
    StrToFec = "'NULL'"
End Function
Public Function StrToBool(cadena As String)

End Function
Public Function StrToDbl(cadena As String) As Double
    'On Error GoTo cero:
    StrToDbl = CDbl(cadena)
'cero:
    'StrToDbl = 0
End Function
Public Sub InsertarPaso(terceros As Integer, paso As Integer)
    If Not EsNulo(terceros) Then
        StrSql = "INSERT INTO paso_ext (pasnro, extnro,extestado, extfecha, extusuario) "
        StrSql = StrSql & "  VALUES( " & paso & " , " & terceros & ",-1," & ConvFecha(Date) & " , '" & Left(usuario, 20) & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub
Function TieneIdioma(l_ternro As Integer, l_idioma As Integer) As Boolean
    Dim rs_sub As New ADODB.Recordset
    StrSql = " SELECT empleado, idinro FROM emp_idi WHERE empleado = " & l_ternro & " and idinro = " & l_idioma
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        TieneIdioma = False
    Else
        TieneIdioma = True
    End If
End Function

Public Function TraerCodEstadoCivil(EstCivdesabr As String) As Integer
    If Not EsNulo(EstCivdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        Select Case UCase(EstCivdesabr)
            'Los case son datos cero, sino creo uno nuevo
            Case "Sin Datos", "NO ESPECIFICADO", ""
                TraerCodEstadoCivil = 1
            Case "CASADO", "CASADO/A"
                TraerCodEstadoCivil = 2
            Case "CONVIVENCIA"
                TraerCodEstadoCivil = 3
            Case "DIVORCIADO", "DIVORCIADO/A"
                TraerCodEstadoCivil = 4
            Case "SEPARADO", "SEPARADO/A"
                TraerCodEstadoCivil = 5
            Case "SEPARADO DE HECHO"
                TraerCodEstadoCivil = 6
            Case "SEPARADO LEGAL"
                TraerCodEstadoCivil = 7
            Case "SOLTERO", "SOLTERO/A"
                TraerCodEstadoCivil = 8
            Case "VIUDO", "VIUDO/A"
                TraerCodEstadoCivil = 9
            Case Else
                StrSql = " SELECT estcivnro FROM estcivil WHERE estcivdesabr = '" & EstCivdesabr & "'"
                OpenRecordset StrSql, rs_sub
                If rs_sub.EOF Then
                    StrSql = "INSERT INTO estcivil (estcivdesabr) VALUES('"
                    StrSql = StrSql & Left(UCase(EstCivdesabr), 30) & "')"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    StrSql = " SELECT MAX(estcivnro) AS Maxestcivnro FROM estcivil "
                    OpenRecordset StrSql, rs_sub
                        
                    TraerCodEstadoCivil = rs_sub!Maxestcivnro
                Else
                    TraerCodEstadoCivil = rs_sub!estcivnro
                End If
        End Select
    Else
        TraerCodEstadoCivil = 1 'Sin datos
    End If
End Function
Public Function TraerCodTipoDocumento(Sigla As String)
    If Not EsNulo(Sigla) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT tidnro FROM tipodocu WHERE tidsigla = '" & Left(Sigla, 8) & "' OR tidnom = '" & Left(Sigla, 30) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO tipodocu (tidsigla, tidnom, tidpers, tidsist, instnro,tidunico) VALUES('"
            StrSql = StrSql & Left(Sigla, 8) & "','" & Left(Sigla, 30) & "',0,0,7,0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(tidnro) AS Maxtidnro FROM tipodocu "
            OpenRecordset StrSql, rs_sub
                
            TraerCodTipoDocumento = rs_sub!Maxtidnro
        Else
            TraerCodTipoDocumento = rs_sub!tidnro
        End If
    Else
        'TraerCodTipoDocumento = TraerCodTipoDocumento("dni")
    End If
End Function
Public Function TraerCodLocalidad(Localidad As String)
    If Not EsNulo(Localidad) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT locnro FROM localidad WHERE locdesc = '" & Localidad & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO localidad (locdesc) VALUES('"
            StrSql = StrSql & UCase(Left(Localidad, 30)) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(locnro) AS Maxlocnro FROM localidad "
            OpenRecordset StrSql, rs_sub
                
            TraerCodLocalidad = rs_sub!Maxlocnro
        Else
            TraerCodLocalidad = rs_sub!locnro
        End If
    Else
        TraerCodLocalidad = 1 'NO INFORMADA
    End If
End Function
Public Function TraerCodProvincia(Provincia As String)
    If Not EsNulo(Provincia) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT provnro FROM Provincia WHERE provdesc = '" & Provincia & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Provincia (provdesc) VALUES('"
            StrSql = StrSql & UCase(Left(Provincia, 20)) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(provnro) AS Maxprovnro FROM Provincia "
            OpenRecordset StrSql, rs_sub
                
            TraerCodProvincia = rs_sub!Maxprovnro
        Else
            TraerCodProvincia = rs_sub!provnro
        End If
    Else
        TraerCodProvincia = 1 'no informada
    End If
End Function

Public Function TraerCodPartido(Partido As String) As Integer
    If Not EsNulo(Partido) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT partnro FROM Partido WHERE partnom = '" & Partido & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Partido (partnom) VALUES('"
            StrSql = StrSql & UCase(Left(Partido, 30)) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(partnro) AS Maxpartnro FROM Partido "
            OpenRecordset StrSql, rs_sub
                
            TraerCodPartido = rs_sub!Maxpartnro
        Else
            TraerCodPartido = rs_sub!partnro
        End If
    Else
        TraerCodPartido = 1 'Sin datos
    End If
End Function


Public Function TraerCodZona(Zona As String, provnro As Integer)
    If Not EsNulo(Zona) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT zonanro FROM Zona WHERE zonadesc = '" & Zona & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO zona (zonadesc, provnro) VALUES('"
            StrSql = StrSql & Left(Zona, 20) & "'," & provnro & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(zonanro) AS Maxzonanro FROM zona "
            OpenRecordset StrSql, rs_sub
                
            TraerCodZona = rs_sub!Maxzonanro
        Else
            TraerCodZona = rs_sub!zonanro
        End If
    End If
End Function
Public Function TraerCodPais(Paisdesc As String)
    If Not EsNulo(Paisdesc) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT paisnro FROM Pais WHERE paisdesc = '" & Paisdesc & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Pais (paisdesc) VALUES('"
            StrSql = StrSql & Left(UCase(Paisdesc), 20) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(paisnro) AS Maxpaisnro FROM pais "
            OpenRecordset StrSql, rs_sub
                
            TraerCodPais = rs_sub!Maxpaisnro
        Else
            TraerCodPais = rs_sub!paisnro
        End If
    Else
        TraerCodPais = 1
    End If
End Function
Public Function TraerCodNacionalidad(Nacionaldes As String)
    If Not EsNulo(Nacionaldes) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT Nacionalnro FROM Nacionalidad WHERE Nacionaldes = '" & Nacionaldes & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Nacionalidad (Nacionaldes) VALUES('"
            StrSql = StrSql & Left(Nacionaldes, 20) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(Nacionalnro) AS MaxNacionalnro FROM Nacionalidad "
            OpenRecordset StrSql, rs_sub
                
            TraerCodNacionalidad = rs_sub!MaxNacionalnro
        Else
            TraerCodNacionalidad = rs_sub!nacionalnro
        End If
    End If
End Function

Public Function TraerCodNivelEstudio(nivdesc As String)
    If Not EsNulo(nivdesc) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT nivnro FROM nivest WHERE nivdesc = '" & nivdesc & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            
            StrSql = " INSERT INTO nivest (nivdesc,nivsist,nivobligatorio,nivestfli) VALUES ("
            StrSql = StrSql & "'" & Left(UCase(nivdesc), 40) & "'" & ",0,0,0 )"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(nivnro) AS Maxnivnro FROM nivest "
            OpenRecordset StrSql, rs_sub
                
            TraerCodNivelEstudio = CInt(rs_sub!Maxnivnro)
        Else
            TraerCodNivelEstudio = CInt(rs_sub!nivnro)
        End If
    End If
End Function
Public Function TraerCodCarrera(Carredudesabr As String)
    If Not EsNulo(Carredudesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT carredunro FROM cap_carr_edu WHERE carredudesabr = '" & Carredudesabr & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO cap_carr_edu (Carredudesabr) "
            StrSql = StrSql & " VALUES('" & Left(UCase(Carredudesabr), 60) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(carredunro) AS Maxcarredunro FROM cap_carr_edu "
            OpenRecordset StrSql, rs_sub
                
            TraerCodCarrera = CInt(rs_sub!Maxcarredunro)
        Else
            TraerCodCarrera = CInt(rs_sub!carredunro)
        End If
    Else
        TraerCodCarrera = "NULL"
    End If
End Function
Public Function TraerCodCausa(caudes As String)
    If Not EsNulo(caudes) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT caunro FROM causa WHERE caudes = '" & caudes & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO causa (caudes) "
            StrSql = StrSql & " VALUES('" & Left(UCase(caudes), 60) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(caunro) AS Maxcaunro FROM causa "
            OpenRecordset StrSql, rs_sub
                
            TraerCodCausa = CInt(rs_sub!Maxcaunro)
        Else
            TraerCodCausa = CInt(rs_sub!caunro)
        End If
    End If
End Function
Public Function TraerCodTitulo(Titdesabr As String, nivnro As Integer)
    If Not EsNulo(Titdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT titnro FROM titulo WHERE titdesabr = '" & Left(UCase(Trim(Titdesabr)), 40) & "'"
        'StrSql = StrSql & " AND nivnro = " & nivnro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO titulo (titdesabr, nivnro ) "
            StrSql = StrSql & " VALUES('" & Left(UCase(Trim(Titdesabr)), 40) & "'," & nivnro & ")"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(titnro) AS Maxtitnro FROM titulo "
            OpenRecordset StrSql, rs_sub
                
            TraerCodTitulo = CInt(rs_sub!Maxtitnro)
        Else
            TraerCodTitulo = CInt(rs_sub!titnro)
        End If
    Else
        TraerCodTitulo = "Null"
    End If
End Function
Public Function TraerCodTituloSolo(Titdesabr As String)
    If Not EsNulo(Titdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT titnro FROM titulo WHERE titdesabr = '" & Titdesabr & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO titulo (titdesabr) "
            StrSql = StrSql & " VALUES('" & Left(UCase(Titdesabr), 40) & "')"
'            StrSql = "INSERT INTO nivest (titdesabr) "
'            StrSql = StrSql & " VALUES('" & Titdesabr & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(titnro) AS Maxtitnro FROM titulo "
            OpenRecordset StrSql, rs_sub
                
            TraerCodTituloSolo = CInt(rs_sub!Maxtitnro)
        Else
            TraerCodTituloSolo = CInt(rs_sub!titnro)
        End If
    End If
End Function
Public Function TraerCodInstitucion(Instdes As String)
    If Not EsNulo(Instdes) Then
        Dim rs_sub As New ADODB.Recordset
        Dim Arreglo
            Dim cadena As String
            Dim a As Integer
            Arreglo = Split(Instdes)
            If UBound(Arreglo) <= 0 Then
                cadena = Left(Trim(Arreglo(a)), 3)
            Else
                For a = 0 To UBound(Arreglo)
                    cadena = cadena & Left(Trim(Arreglo(a)), 1)
                Next a
            End If
        StrSql = " SELECT instnro FROM institucion WHERE instdes = '" & UCase(Instdes) & "'"
        StrSql = StrSql & " OR instabre = '" & UCase(Left(Instdes, 30)) & "'"
        StrSql = StrSql & " OR instabre = '" & UCase(Left(cadena, 30)) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
       
            StrSql = " INSERT INTO institucion (instdes,instabre, instedu) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(Instdes), 200) & "','" & Left(UCase(cadena), 30) & "',-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            

            StrSql = " SELECT MAX(instnro) AS Maxinstnro FROM institucion "
            OpenRecordset StrSql, rs_sub

            TraerCodInstitucion = CInt(rs_sub!Maxinstnro)
        Else
            TraerCodInstitucion = CInt(rs_sub!instnro)
        End If
    End If
End Function
Public Function TraerCodInstitucionAbreviada(Instabre As String)
    If Not EsNulo(Instabre) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT instnro FROM institucion WHERE Instabre = '" & UCase(Instabre) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO institucion (instdes,instabre, instedu) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(Instabre), 200) & "','" & Left(UCase(Instabre), 30) & "',-1)"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(instnro) AS Maxinstnro FROM institucion "
            OpenRecordset StrSql, rs_sub

            TraerCodInstitucionAbreviada = CInt(rs_sub!Maxinstnro)
        Else
            TraerCodInstitucionAbreviada = CInt(rs_sub!instnro)
        End If
    Else
        TraerCodInstitucionAbreviada = 7 'NO informada
    End If
End Function

Public Function TraerCodCargo(Cardesabr As String)
    If Not EsNulo(Cardesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT carnro FROM cargo WHERE cardesabr = '" & Cardesabr & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO cargo (cardesabr ) "
            StrSql = StrSql & " VALUES('" & Left(Cardesabr, 50) & "')"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(carnro) AS Maxcarnro FROM cargo "
            OpenRecordset StrSql, rs_sub

            TraerCodCargo = CInt(rs_sub!Maxcarnro)
        Else
            TraerCodCargo = CInt(rs_sub!carnro)
        End If
    Else
        StrSql = "INSERT INTO cargo (cardesabr ) "
        StrSql = StrSql & " VALUES('" & Left(Cardesabr, 50) & "')"

        objConn.Execute StrSql, , adExecuteNoRecords

        StrSql = " SELECT MAX(carnro) AS Maxcarnro FROM cargo "
        OpenRecordset StrSql, rs_sub

        TraerCodCargo = CInt(rs_sub!Maxcarnro)
    End If
End Function
Public Function TraerCodTipoCurso(tipcurdesabr As String) As Integer
    If Not EsNulo(tipcurdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT tipcurnro FROM cap_tipocurso WHERE tipcurdesabr = '" & UCase(tipcurdesabr) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_tipocurso (tipcurdesabr) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(tipcurdesabr), 50) & "')"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(tipcurnro) AS Maxtipcurnro FROM cap_tipocurso "
            OpenRecordset StrSql, rs_sub

            TraerCodTipoCurso = CInt(rs_sub!Maxtipcurnro)
        Else
            TraerCodTipoCurso = CInt(rs_sub!tipcurnro)
        End If
    End If
End Function

Public Function TraerCodCurso(curdesabr As String, tipcurnro As Integer) As Integer
    If Not EsNulo(curdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT curnro FROM cap_curso WHERE curdesabr = '" & UCase(curdesabr) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_curso (curdesabr,curcodext, tipcurnro) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(curdesabr), 50) & "','" & Left(UCase(curdesext), 25) & "', " & tipcurnro & " )"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(curnro) AS Maxcurnro FROM cap_curso "
            OpenRecordset StrSql, rs_sub

            TraerCodCurso = CInt(rs_sub!Maxcurnro)
        Else
            TraerCodCurso = CInt(rs_sub!curnro)
        End If
    End If
End Function



'Public Function TraerCodEltoana(eltanadesabr As String, espnro As Integer)
'    If Not EsNulo(eltanadesabr) Then
'        Dim rs_sub As New ADODB.Recordset
'        StrSql = " SELECT eltananro FROM eltoana WHERE eltanadesabr = '" & Trim(eltanadesabr) & "' and espnro = " & CInt(espnro)
'        OpenRecordset StrSql, rs_sub
'        If rs_sub.EOF Then
'            StrSql = "INSERT INTO eltoana (eltanadesabr, espnro ) "
'            StrSql = StrSql & " VALUES('" & Left(Trim(eltanadesabr), 40) & "'," & espnro & ")"
'
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            StrSql = " SELECT MAX(eltananro) AS Maxeltananro FROM eltoana "
'            OpenRecordset StrSql, rs_sub
'
'            TraerCodEltoana = CInt(rs_sub!Maxeltananro)
'        Else
'            TraerCodEltoana = CInt(rs_sub!eltananro)
'        End If
'    Else
'        StrSql = "INSERT INTO eltoana (eltanadesabr, espnro ) "
'        StrSql = StrSql & " VALUES('" & Left(Trim(eltanadesabr), 40) & "', " & espnro & ")"
'
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        StrSql = " SELECT MAX(eltananro) AS Maxeltananro FROM eltoana "
'        OpenRecordset StrSql, rs_sub
'
'        TraerCodEltoana = CInt(rs_sub!Maxeltananro)
'    End If
'End Function

Public Function TraerCodEltoana(eltanadesabr As String, espnro As Integer)
'Public Function TraerCodEltoana(eltanadesabr As String) As Integer
    If Not EsNulo(eltanadesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT eltananro FROM eltoana WHERE eltanadesabr = '" & Trim(eltanadesabr) & "' and espnro = " & CInt(espnro)
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO eltoana (eltanadesabr, espnro ) "
            StrSql = StrSql & " VALUES('" & Left(Trim(eltanadesabr), 40) & "'," & espnro & ")"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(eltananro) AS Maxeltananro FROM eltoana "
            OpenRecordset StrSql, rs_sub

            TraerCodEltoana = CInt(rs_sub!Maxeltananro)
        Else
            TraerCodEltoana = CInt(rs_sub!eltananro)
        End If
    Else
        StrSql = "INSERT INTO eltoana (eltanadesabr, espnro ) "
        StrSql = StrSql & " VALUES('" & Left(Trim(eltanadesabr), 40) & "', " & espnro & ")"

        objConn.Execute StrSql, , adExecuteNoRecords

        StrSql = " SELECT MAX(eltananro) AS Maxeltananro FROM eltoana "
        OpenRecordset StrSql, rs_sub

        TraerCodEltoana = CInt(rs_sub!Maxeltananro)
    End If
End Function

Public Function TraerEspecializacion(espdesabr As String)
    If Not EsNulo(espdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT espnro FROM especializacion WHERE espdesabr = '" & espdesabr & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO especializacion (espdesabr) "
            StrSql = StrSql & " VALUES('" & Left(espdesabr, 40) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(espnro) AS Maxespnro FROM especializacion "
            OpenRecordset StrSql, rs_sub
                
            TraerEspecializacion = CInt(rs_sub!Maxespnro)
        Else
            TraerEspecializacion = CInt(rs_sub!espnro)
        End If
    End If
End Function


Public Function TraerCodNivelEspecializacion(espnivdesabr As String)
    If Not EsNulo(espnivdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT espnivnro FROM espnivel WHERE espnivdesabr = '" & espnivdesabr & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO espnivel (espnivdesabr) "
            StrSql = StrSql & " VALUES('" & Left(espnivdesabr, 40) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(espnivnro) AS Maxespnivnro FROM espnivel "
            OpenRecordset StrSql, rs_sub
                
            TraerCodNivelEspecializacion = CInt(rs_sub!Maxespnivnro)
        Else
            TraerCodNivelEspecializacion = CInt(rs_sub!espnivnro)
        End If
    End If
End Function
Public Function TraerCodEspecializacion(espdesabr As String)
    If Not EsNulo(espdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT espnro FROM especializacion WHERE espdesabr = '" & espdesabr & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO especializacion (espdesabr) "
            StrSql = StrSql & " VALUES('" & Left(espdesabr, 40) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(espnro) AS Maxespnro FROM especializacion "
            OpenRecordset StrSql, rs_sub
                
            TraerCodEspecializacion = CInt(rs_sub!Maxespnro)
        Else
            TraerCodEspecializacion = CInt(rs_sub!espnro)
        End If
    End If
End Function
Public Function TraerCodProcedencia(Prodesabr As String)
    If Not EsNulo(Trim(Left(Prodesabr, 30))) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT pronro FROM pos_procedencia WHERE prodesabr = '" & Trim(Left(Prodesabr, 30)) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO pos_procedencia (prodesabr) "
            StrSql = StrSql & " VALUES('" & Trim(Left(Prodesabr, 30)) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(pronro) AS Maxpronro FROM pos_procedencia "
            OpenRecordset StrSql, rs_sub
                
            TraerCodProcedencia = CInt(rs_sub!Maxpronro)
        Else
            TraerCodProcedencia = CInt(rs_sub!pronro)
        End If
    End If
End Function

Public Function TraerCodListaEmpresa(lempdes As String)
    lempdes = Left(lempdes, 60)
    If Not EsNulo(lempdes) Then
        Dim Rs_Estr As New ADODB.Recordset
        StrSql = " SELECT lempnro FROM listaemp WHERE lempdes = '" & lempdes & "'"
        OpenRecordset StrSql, Rs_Estr
        If Not Rs_Estr.EOF Then
            TraerCodListaEmpresa = CInt(Rs_Estr!lempnro)
        Else
            StrSql = " INSERT INTO listaemp(lempdes)"
            StrSql = StrSql & " VALUES('" & Left(UCase(lempdes), 60) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(lempnro) AS MaxEmpnro FROM listaemp "
            OpenRecordset StrSql, Rs_Estr
            
            TraerCodListaEmpresa = CInt(Rs_Estr!MaxEmpnro)
        End If
    Else
        TraerCodListaEmpresa = 0
    End If
End Function
Public Function TraerCodIdioma(ididesc As String)
    If Not EsNulo(ididesc) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT idinro FROM Idioma WHERE ididesc = '" & ididesc & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO idioma (ididesc) "
            StrSql = StrSql & " VALUES('" & Left(ididesc, 30) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(idinro) AS Maxidinro FROM idioma "
            OpenRecordset StrSql, rs_sub
                
            TraerCodIdioma = CInt(rs_sub!Maxidinro)
        Else
            TraerCodIdioma = CInt(rs_sub!idinro)
        End If
    End If
End Function
Public Function TraerCodIdiNivel(idnivdesabr As String)
    If Not EsNulo(idnivdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT idnivnro FROM idinivel WHERE idnivdesabr = '" & idnivdesabr & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO idinivel (idnivdesabr) "
            StrSql = StrSql & " VALUES('" & Left(idnivdesabr, 30) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(idnivnro) AS Maxidnivnro FROM idinivel "
            OpenRecordset StrSql, rs_sub
                
            TraerCodIdiNivel = CInt(rs_sub!Maxidnivnro)
        Else
            TraerCodIdiNivel = CInt(rs_sub!idnivnro)
        End If
    End If
End Function

Function validatelefono(cadena As String)
    Dim a As Integer
    Dim car As String
    Dim cadenacompleta As String
    For a = 1 To Len(cadena)
        car = Asc(Mid(cadena, a, 1))
        If (car > 47 And car < 58) Or (car > 39 And car < 43) Or (car = 45) Or (car = 32) Or (car = 35) Then
            cadenacompleta = CStr(cadenacompleta) & CStr(Chr(car))
        End If
    Next a
    validatelefono = cadenacompleta
End Function

