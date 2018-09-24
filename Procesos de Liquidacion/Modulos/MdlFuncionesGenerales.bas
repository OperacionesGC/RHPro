Attribute VB_Name = "MdlFuncionesGenerales"
Public Function StrToStr(cadena As String, Longitud As Long)
    StrToStr = CStr(Left(Trim(CStr(cadena)), Longitud))
End Function
Public Function StrToInt(cadena As String) As Long
    'On Error GoTo cero:
    StrToInt = CLng(cadena)
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
Public Sub InsertarPaso(terceros As Long, paso As Long)
    If Not EsNulo(terceros) Then
        StrSql = "INSERT INTO paso_ext (pasnro, extnro,extestado, extfecha, extusuario) "
        StrSql = StrSql & "  VALUES( " & paso & " , " & terceros & ",-1," & ConvFecha(Date) & " , '" & Left(usuario, 20) & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub
Function TieneIdioma(l_ternro As Long, l_idioma As Long) As Boolean
    Dim rs_sub As New ADODB.Recordset
    StrSql = " SELECT empleado, idinro FROM emp_idi WHERE empleado = " & l_ternro & " and idinro = " & l_idioma
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        TieneIdioma = False
    Else
        TieneIdioma = True
    End If
End Function

Public Function TraerCodTipInformacion(DescTipInf As String) As Long

    Dim rs_sub As New ADODB.Recordset
    Dim Aux_Nro_TipInf As Long
 
    Aux_Nro_TipInf = 0
    
    If DescTipInf <> "" Then
        StrSql = " SELECT tipinfnro FROM pos_tipinfgral WHERE upper(tipinfdesabr) = '" & Left(UCase(DescTipInf), 50) & "'"
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
            StrSql = "INSERT INTO pos_tipinfgral (tipinfdesabr,tipinfdesext) VALUES('"
            StrSql = StrSql & Left(UCase(DescTipInf), 50) & "',' ')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(tipinfnro) AS Maxtipinfnro FROM pos_tipinfgral "
            OpenRecordset StrSql, rs_sub
                
            Aux_Nro_TipInf = rs_sub!Maxtipinfnro
        Else
            Aux_Nro_TipInf = rs_sub!tipinfnro
        End If
    Else
        Aux_Nro_TipInf = 0
    End If
    TraerCodTipInformacion = Aux_Nro_TipInf
    
    If rs_sub.State = adStateOpen Then rs_sub.Close
    Set rs_sub = Nothing
    
End Function
Public Function TraerCodItemInformacion(DescItemInf As String, NroTipInf As Long) As Long

    Dim rs_sub As New ADODB.Recordset
    Dim Aux_Nro_Item As Long
 
    Aux_Nro_Item = 0
    EliminarCHInvalidos (DescItemInf)
    If DescItemInf <> "" Then
        StrSql = " SELECT infnro FROM pos_infgral WHERE upper(infdesabr) = '" & Left(UCase(DescItemInf), 50) & "'"
        StrSql = StrSql & " AND tipinfnro = " & NroTipInf
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
            StrSql = "INSERT INTO pos_infgral (infdesabr,infdesext,tipinfnro) VALUES('"
            StrSql = StrSql & Left(UCase(DescItemInf), 50) & "',' '," & NroTipInf & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            
            StrSql = " SELECT MAX(infnro) AS Maxitemnro FROM pos_infgral "
            OpenRecordset StrSql, rs_sub
                
            Aux_Nro_Item = rs_sub!Maxitemnro
        Else
            Aux_Nro_Item = rs_sub!infnro
        End If
    Else
        Aux_Nro_Item = 0
    End If
    TraerCodItemInformacion = Aux_Nro_Item
    
    If rs_sub.State = adStateOpen Then rs_sub.Close
    Set rs_sub = Nothing
    
End Function
Public Function AsociarParAPostulante(NroItemInfo As Long, Ternro As Long) As Long

    Dim rs_sub As New ADODB.Recordset
    Dim Aux_asocio As Long
 
    Aux_asocio = 0
   
    StrSql = " SELECT * FROM pos_infter WHERE infnro = " & NroItemInfo
    StrSql = StrSql & " AND ternro = " & Ternro
    OpenRecordset StrSql, rs_sub
    
    If rs_sub.EOF Then
        StrSql = "INSERT INTO pos_infter (infnro,ternro) "
        StrSql = StrSql & "VALUES(" & NroItemInfo & "," & Ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Aux_asocio = -1
    End If
        
    AsociarParAPostulante = Aux_asocio
    
    If rs_sub.State = adStateOpen Then rs_sub.Close
    Set rs_sub = Nothing
    
End Function
Public Function TraerCodEstadoCivil(EstCivdesabr As String) As Long
Dim rs_sub As New ADODB.Recordset
Dim Aux_Nro_Estcivil As Long

    Aux_Nro_Estcivil = 0
    If Not EsNulo(EstCivdesabr) Then
        StrSql = " SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & Left(UCase(EstCivdesabr), 30) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO estcivil (estcivdesabr) VALUES('"
            StrSql = StrSql & Left(UCase(EstCivdesabr), 30) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(estcivnro) AS Maxestcivnro FROM estcivil "
            OpenRecordset StrSql, rs_sub
                
            Aux_Nro_Estcivil = rs_sub!Maxestcivnro
        Else
            Aux_Nro_Estcivil = rs_sub!estcivnro
        End If
    Else
        Aux_Nro_Estcivil = 0
    End If
    TraerCodEstadoCivil = Aux_Nro_Estcivil
    
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function


'Public Function TraerCodEstadoCivil_old(EstCivdesabr As String) as long
'Dim rs_sub As New ADODB.Recordset
'
'    If Not EsNulo(EstCivdesabr) Then
'        Select Case UCase(EstCivdesabr)
'            'Los case son datos cero, sino creo uno nuevo
'            Case "Sin Datos", "NO ESPECIFICADO", "", "N/A"
'                TraerCodEstadoCivil = 1
'            Case "CASADO", "CASADO/A"
'                TraerCodEstadoCivil = 2
'            Case "CONVIVENCIA"
'                TraerCodEstadoCivil = 3
'            Case "DIVORCIADO", "DIVORCIADO/A"
'                TraerCodEstadoCivil = 4
'            Case "SEPARADO", "SEPARADO/A"
'                TraerCodEstadoCivil = 5
'            Case "SEPARADO DE HECHO"
'                TraerCodEstadoCivil = 6
'            Case "SEPARADO LEGAL"
'                TraerCodEstadoCivil = 7
'            Case "SOLTERO", "SOLTERO/A"
'                TraerCodEstadoCivil = 8
'            Case "VIUDO", "VIUDO/A"
'                TraerCodEstadoCivil = 9
'            Case Else
'                StrSql = " SELECT estcivnro FROM estcivil WHERE estcivdesabr = '" & EstCivdesabr & "'"
'                OpenRecordset StrSql, rs_sub
'                If rs_sub.EOF Then
'                    StrSql = "INSERT INTO estcivil (estcivdesabr) VALUES('"
'                    StrSql = StrSql & Left(UCase(EstCivdesabr), 30) & "')"
'                    objConn.Execute StrSql, , adExecuteNoRecords
'
'                    StrSql = " SELECT MAX(estcivnro) AS Maxestcivnro FROM estcivil "
'                    OpenRecordset StrSql, rs_sub
'
'                    TraerCodEstadoCivil = rs_sub!Maxestcivnro
'                Else
'                    TraerCodEstadoCivil = rs_sub!estcivnro
'                End If
'        End Select
'    Else
'        TraerCodEstadoCivil = 1 'Sin datos
'    End If
'
'
'If rs_sub.State = adStateOpen Then rs_sub.Close
'Set rs_sub = Nothing
'End Function


Public Function TraerCodTipoDocumento(Sigla As String)
    If Not EsNulo(Sigla) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT tidnro FROM tipodocu WHERE upper(tidsigla) = '" & UCase(Left(Sigla, 8)) & "' OR upper(tidnom) = '" & UCase(Left(Sigla, 30)) & "'"
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
        StrSql = " SELECT locnro FROM localidad WHERE upper(locdesc) = '" & Left(UCase(Localidad), 30) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO localidad (locdesc) VALUES('"
            StrSql = StrSql & UCase(Left(Localidad, 30)) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(locnro) AS Maxlocnro FROM localidad "
            OpenRecordset StrSql, rs_sub
                
            TraerCodLocalidad = rs_sub!Maxlocnro
        Else
            TraerCodLocalidad = rs_sub!Locnro
        End If
    Else
        TraerCodLocalidad = 1 'NO INFORMADA
    End If
End Function
Public Function TraerCodLocalidadPorNro(LocalidadNro As Integer)
    If Not EsNulo(LocalidadNro) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT locnro FROM localidad WHERE locnro = " & LocalidadNro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodLocalidadPorNro = -1
        Else
            TraerCodLocalidadPorNro = rs_sub!Locnro
        End If
    Else
        TraerCodLocalidadPorNro = 1 'NO INFORMADA
    End If
End Function
Public Function TraerCodProvincia(Provincia As String)
    If Not EsNulo(Provincia) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT provnro FROM Provincia WHERE upper(provdesc) = '" & Left(UCase(Provincia), 20) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Provincia (provdesc) VALUES('"
            StrSql = StrSql & UCase(Left(Provincia, 20)) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(provnro) AS Maxprovnro FROM Provincia "
            OpenRecordset StrSql, rs_sub
                
            TraerCodProvincia = rs_sub!Maxprovnro
        Else
            TraerCodProvincia = rs_sub!Provnro
        End If
    Else
        TraerCodProvincia = 1 'no informada
    End If
End Function
Public Function TraerCodProvinciaPorNro(ProvinciaNro As Integer)
    If Not EsNulo(ProvinciaNro) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT provnro FROM Provincia WHERE provnro = " & ProvinciaNro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodProvinciaPorNro = -1
        Else
            TraerCodProvinciaPorNro = rs_sub!Provnro
        End If
    Else
        TraerCodProvinciaPorNro = 1 'no informada
    End If
End Function

Public Function TraerCodPartido(Partido As String) As Long
    If Not EsNulo(Partido) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT partnro FROM Partido WHERE upper(partnom) = '" & Left(UCase(Partido), 30) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Partido (partnom) VALUES('"
            StrSql = StrSql & UCase(Left(Partido, 30)) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(partnro) AS Maxpartnro FROM Partido "
            OpenRecordset StrSql, rs_sub
                
            TraerCodPartido = rs_sub!Maxpartnro
        Else
            TraerCodPartido = rs_sub!Partnro
        End If
    Else
        TraerCodPartido = 1 'Sin datos
    End If
End Function
Public Function TraerCodPartidoPorNro(PartidoNro As Integer) As Long
    If Not EsNulo(PartidoNro) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT partnro FROM Partido WHERE partnro = " & PartidoNro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodPartidoPorNro = -1
        Else
            TraerCodPartidoPorNro = rs_sub!Partnro
        End If
    Else
        TraerCodPartidoPorNro = 1 'Sin datos
    End If
End Function

Public Function TraerCodZona(Zona As String, Provnro As Long)
    If Not EsNulo(Zona) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT zonanro FROM Zona WHERE upper(zonadesc) = '" & Left(UCase(Zona), 20) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO zona (zonadesc, provnro) VALUES('"
            StrSql = StrSql & Left(Zona, 20) & "'," & Provnro & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(zonanro) AS Maxzonanro FROM zona "
            OpenRecordset StrSql, rs_sub
                
            TraerCodZona = rs_sub!Maxzonanro
        Else
            TraerCodZona = rs_sub!Zonanro
        End If
    End If
End Function
Public Function TraerCodPais(Paisdesc As String)
    If Not EsNulo(Paisdesc) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT paisnro FROM Pais WHERE upper(paisdesc) = '" & Left(UCase(Paisdesc), 20) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Pais (paisdesc) VALUES('"
            StrSql = StrSql & Left(UCase(Paisdesc), 20) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(paisnro) AS Maxpaisnro FROM pais "
            OpenRecordset StrSql, rs_sub
                
            TraerCodPais = rs_sub!Maxpaisnro
        Else
            TraerCodPais = rs_sub!Paisnro
        End If
    Else
        TraerCodPais = 1
    End If
End Function
Public Function TraerCodPaisPorNro(PaisNumero As Integer)
    If Not EsNulo(PaisNumero) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT paisnro FROM Pais WHERE paisnro = " & PaisNumero
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodPaisPorNro = -1
        Else
            TraerCodPaisPorNro = rs_sub!Paisnro
        End If
    Else
        TraerCodPaisPorNro = 1
    End If
End Function
Public Function TraerCodNacionalidad(Nacionaldes As String)
    If Not EsNulo(Nacionaldes) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT Nacionalnro FROM Nacionalidad WHERE upper(Nacionaldes) = '" & Left(UCase(Nacionaldes), 20) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO Nacionalidad (Nacionaldes) VALUES('"
            StrSql = StrSql & Left(UCase(Nacionaldes), 20) & "')"
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
        StrSql = " SELECT nivnro FROM nivest WHERE upper(nivdesc) = '" & Left(UCase(nivdesc), 40) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            
            StrSql = " INSERT INTO nivest (nivdesc,nivsist,nivobligatorio,nivestfli) VALUES ("
            StrSql = StrSql & "'" & Left(UCase(nivdesc), 40) & "'" & ",0,0,0 )"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(nivnro) AS Maxnivnro FROM nivest "
            OpenRecordset StrSql, rs_sub
                
            TraerCodNivelEstudio = CLng(rs_sub!Maxnivnro)
        Else
            TraerCodNivelEstudio = CLng(rs_sub!nivnro)
        End If
    End If
End Function
Public Function TraerCodCarrera(Carredudesabr As String)
    If Not EsNulo(Carredudesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT carredunro FROM cap_carr_edu WHERE carredudesabr = '" & Carredudesabr & "'"
        'StrSql = " SELECT carredunro FROM cap_carr_edu WHERE carredudesext = '" & Carredudesabr & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO cap_carr_edu (Carredudesabr) "
            StrSql = StrSql & " VALUES('" & Left(UCase(Carredudesabr), 90) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(carredunro) AS Maxcarredunro FROM cap_carr_edu "
            OpenRecordset StrSql, rs_sub
                
            TraerCodCarrera = CLng(rs_sub!Maxcarredunro)
        Else
            TraerCodCarrera = CLng(rs_sub!CarreduNro)
        End If
    Else
        TraerCodCarrera = "NULL"
    End If
End Function
Public Function TraerCodCarreraPorNro(CarreduNro As Integer)
    If Not EsNulo(CarreduNro) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT carredunro FROM cap_carr_edu WHERE carredunro = " & CarreduNro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodCarreraPorNro = -1
        Else
            TraerCodCarreraPorNro = CLng(rs_sub!CarreduNro)
        End If
    Else
        TraerCodCarreraPorNro = "NULL"
    End If
End Function
Public Function TraerCodCausa(caudes As String)
    If Not EsNulo(caudes) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT caunro FROM causa WHERE upper(caudes) = '" & UCase(Left(caudes, 60)) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO causa (caudes) "
            StrSql = StrSql & " VALUES('" & Left(UCase(caudes), 60) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(caunro) AS Maxcaunro FROM causa "
            OpenRecordset StrSql, rs_sub
                
            TraerCodCausa = CLng(rs_sub!Maxcaunro)
        Else
            TraerCodCausa = CLng(rs_sub!caunro)
        End If
    End If
End Function
Public Function TraerCodTitulo(Titdesabr As String, nivnro As Long)
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
                
            TraerCodTitulo = CLng(rs_sub!Maxtitnro)
        Else
            TraerCodTitulo = CLng(rs_sub!TitNro)
        End If
    Else
        TraerCodTitulo = "Null"
    End If
End Function
Public Function TraerCodTituloPorNro(TitNro As Integer, nivnro As Long)
    If Not EsNulo(TitNro) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT titnro FROM titulo WHERE titnro = " & TitNro
        'StrSql = StrSql & " AND nivnro = " & nivnro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodTituloPorNro = -1
        Else
            TraerCodTituloPorNro = CLng(rs_sub!TitNro)
        End If
    Else
        TraerCodTituloPorNro = "Null"
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
                
            TraerCodTituloSolo = CLng(rs_sub!Maxtitnro)
        Else
            TraerCodTituloSolo = CLng(rs_sub!TitNro)
        End If
    End If
End Function
Public Function TraerCodInstitucion(Instdes As String)
    If Not EsNulo(Instdes) Then
        Dim rs_sub As New ADODB.Recordset
        Dim Arreglo
            Dim cadena As String
            Dim a As Long
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

            TraerCodInstitucion = CLng(rs_sub!Maxinstnro)
        Else
            TraerCodInstitucion = CLng(rs_sub!InstNro)
        End If
    End If
End Function
Public Function TraerCodInstitucionAbreviada(Instabre As String)
    If Not EsNulo(Instabre) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT instnro FROM institucion WHERE UPPER(Instabre) = '" & UCase(Instabre) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO institucion (instdes,instabre, instedu) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(Instabre), 200) & "','" & Left(UCase(Instabre), 30) & "',-1)"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(instnro) AS Maxinstnro FROM institucion "
            OpenRecordset StrSql, rs_sub

            TraerCodInstitucionAbreviada = CLng(rs_sub!Maxinstnro)
        Else
            TraerCodInstitucionAbreviada = CLng(rs_sub!InstNro)
        End If
    Else
        TraerCodInstitucionAbreviada = 7 'NO informada
    End If
End Function
Public Function TraerCodInstitucionDescripcion(InstDesc As String)
    If Not EsNulo(InstDesc) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT instnro FROM institucion WHERE UPPER(Instdes) = '" & UCase(InstDesc) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO institucion (instdes,instabre, instedu) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(InstDesc), 200) & "','" & Left(UCase(InstDesc), 90) & "',-1)"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(instnro) AS Maxinstnro FROM institucion "
            OpenRecordset StrSql, rs_sub

            TraerCodInstitucionDescripcion = CLng(rs_sub!Maxinstnro)
        Else
            TraerCodInstitucionDescripcion = CLng(rs_sub!InstNro)
        End If
    Else
        TraerCodInstitucionDescripcion = 7 'NO informada
    End If
End Function
Public Function TraerCodInstitucionPorNro(InstNro As Integer)
    If Not EsNulo(InstNro) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT instnro FROM institucion WHERE instnro = " & InstNro
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodInstitucionPorNro = -1
        Else
            TraerCodInstitucionPorNro = CLng(rs_sub!InstNro)
        End If
    Else
        TraerCodInstitucionPorNro = 7 'NO informada
    End If
End Function

Public Function TraerCodCargo(Cardesabr As String)
    If Not EsNulo(Cardesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT carnro FROM cargo WHERE upper(cardesabr) = '" & UCase(Left(Cardesabr, 50)) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO cargo (cardesabr ) "
            StrSql = StrSql & " VALUES('" & UCase(Left(Cardesabr, 50)) & "')"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(carnro) AS Maxcarnro FROM cargo "
            OpenRecordset StrSql, rs_sub

            TraerCodCargo = CLng(rs_sub!Maxcarnro)
        Else
            TraerCodCargo = CLng(rs_sub!carnro)
        End If
    Else
        StrSql = "INSERT INTO cargo (cardesabr ) "
        StrSql = StrSql & " VALUES('" & Left(Cardesabr, 50) & "')"

        objConn.Execute StrSql, , adExecuteNoRecords

        StrSql = " SELECT MAX(carnro) AS Maxcarnro FROM cargo "
        OpenRecordset StrSql, rs_sub

        TraerCodCargo = CLng(rs_sub!Maxcarnro)
    End If
End Function
Public Function TraerCodTipoCurso(tipcurdesabr As String) As Long
    If Not EsNulo(tipcurdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT tipcurnro FROM cap_tipocurso WHERE upper(tipcurdesabr) = '" & Left(UCase(tipcurdesabr), 50) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_tipocurso (tipcurdesabr) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(tipcurdesabr), 50) & "')"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(tipcurnro) AS Maxtipcurnro FROM cap_tipocurso "
            OpenRecordset StrSql, rs_sub

            TraerCodTipoCurso = CLng(rs_sub!Maxtipcurnro)
        Else
            TraerCodTipoCurso = CLng(rs_sub!tipcurnro)
        End If
    End If
End Function

Public Function TraerCodCurso(curdesabr As String, tipcurnro As Long) As Long
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

            TraerCodCurso = CLng(rs_sub!Maxcurnro)
        Else
            TraerCodCurso = CLng(rs_sub!curnro)
        End If
    End If
End Function

'Public Function TraerCodEltoana(eltanadesabr As String, espnro as long)
'    If Not EsNulo(eltanadesabr) Then
'        Dim rs_sub As New ADODB.Recordset
'        StrSql = " SELECT eltananro FROM eltoana WHERE eltanadesabr = '" & Trim(eltanadesabr) & "' and espnro = " & clng(espnro)
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
'            TraerCodEltoana = clng(rs_sub!Maxeltananro)
'        Else
'            TraerCodEltoana = clng(rs_sub!eltananro)
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
'        TraerCodEltoana = clng(rs_sub!Maxeltananro)
'    End If
'End Function

Public Function TraerCodEltoana(eltanadesabr As String, espnro As Long)
'Public Function TraerCodEltoana(eltanadesabr As String) as long
    Flog.writeline Espacios(Tabulador * 2) & "TraerCodEltoana(" & eltanadesabr & ", " & espnro & ")"
    If Not EsNulo(eltanadesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = "SELECT eltananro FROM eltoana WHERE upper(eltanadesabr) = '" & UCase(Left(Trim(eltanadesabr), 40)) & "' and espnro = " & CLng(espnro)
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO eltoana (eltanadesabr, espnro ) "
            StrSql = StrSql & " VALUES('" & UCase(Left(Trim(eltanadesabr), 40)) & "'," & CLng(espnro) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(eltananro) AS Maxeltananro FROM eltoana "
            OpenRecordset StrSql, rs_sub

            TraerCodEltoana = CLng(rs_sub!Maxeltananro)
        Else
            TraerCodEltoana = CLng(rs_sub!eltananro)
        End If
    Else
        Flog.writeline Espacios(Tabulador * 2) & "eltanadesabr NULA. Busco eltoana"
        StrSql = "INSERT INTO eltoana (eltanadesabr, espnro ) "
        StrSql = StrSql & " VALUES('" & Left(Trim(eltanadesabr), 40) & "', " & espnro & ")"

        objConn.Execute StrSql, , adExecuteNoRecords

        StrSql = " SELECT MAX(eltananro) AS Maxeltananro FROM eltoana "
        OpenRecordset StrSql, rs_sub

        TraerCodEltoana = CLng(rs_sub!Maxeltananro)
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
                
            TraerEspecializacion = CLng(rs_sub!Maxespnro)
        Else
            TraerEspecializacion = CLng(rs_sub!espnro)
        End If
    End If
End Function


Public Function TraerCodNivelEspecializacion(espnivdesabr As String)
    If Not EsNulo(espnivdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT espnivnro FROM espnivel WHERE upper(espnivdesabr) = '" & Left(UCase(espnivdesabr), 40) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO espnivel (espnivdesabr) "
            StrSql = StrSql & " VALUES('" & UCase(Left(espnivdesabr, 40)) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(espnivnro) AS Maxespnivnro FROM espnivel "
            OpenRecordset StrSql, rs_sub
                
            TraerCodNivelEspecializacion = CLng(rs_sub!Maxespnivnro)
        Else
            TraerCodNivelEspecializacion = CLng(rs_sub!espnivnro)
        End If
    End If
End Function
Public Function TraerCodEspecializacion(espdesabr As String)
    If Not EsNulo(espdesabr) Then
        Flog.writeline Espacios(Tabulador * 2) & "Especializacion: " & espdesabr
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT espnro FROM especializacion WHERE espdesabr = '" & espdesabr & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO especializacion (espdesabr) "
            StrSql = StrSql & " VALUES('" & Left(espdesabr, 40) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(espnro) AS Maxespnro FROM especializacion "
            OpenRecordset StrSql, rs_sub
                
            TraerCodEspecializacion = CLng(rs_sub!Maxespnro)
        Else
            TraerCodEspecializacion = CLng(rs_sub!espnro)
        End If
    Else
        Flog.writeline Espacios(Tabulador * 2) & "Especializacion NULA"
    End If
End Function
Public Function TraerCodProcedencia(Prodesabr As String)
    If Not EsNulo(Trim(Left(Prodesabr, 30))) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT pronro FROM pos_procedencia WHERE upper(prodesabr) = '" & UCase(Trim(Left(Prodesabr, 30))) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO pos_procedencia (prodesabr) "
            StrSql = StrSql & " VALUES('" & UCase(Trim(Left(Prodesabr, 30))) & "')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(pronro) AS Maxpronro FROM pos_procedencia "
            OpenRecordset StrSql, rs_sub
                
            TraerCodProcedencia = CLng(rs_sub!Maxpronro)
        Else
            TraerCodProcedencia = CLng(rs_sub!pronro)
        End If
    Else
        StrSql = " SELECT pronro FROM pos_procedencia order by pronro "
        OpenRecordset StrSql, rs_sub
        
        TraerCodProcedencia = CLng(rs_sub!pronro)

    End If
    Flog.writeline Espacios(Tabulador * 0) & "Procedencia: " & TraerCodProcedencia
End Function

Public Function TraerCodListaEmpresa(lempdes As String)
    lempdes = Left(lempdes, 60)
    If Not EsNulo(lempdes) Then
        Dim Rs_Estr As New ADODB.Recordset
        StrSql = " SELECT lempnro FROM listaemp WHERE upper(lempdes) = '" & UCase(Left(lempdes, 60)) & "'"
        OpenRecordset StrSql, Rs_Estr
        If Not Rs_Estr.EOF Then
            TraerCodListaEmpresa = CLng(Rs_Estr!lempnro)
        Else
            StrSql = " INSERT INTO listaemp(lempdes)"
            StrSql = StrSql & " VALUES('" & Left(UCase(lempdes), 60) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(lempnro) AS MaxEmpnro FROM listaemp "
            OpenRecordset StrSql, Rs_Estr
            
            TraerCodListaEmpresa = CLng(Rs_Estr!MaxEmpnro)
        End If
    Else
        TraerCodListaEmpresa = 0
    End If
End Function
Public Function TraerCodIdioma(ididesc As String)
    If Not EsNulo(ididesc) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT idinro FROM Idioma WHERE upper(ididesc) = '" & UCase(Left(ididesc, 30)) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO idioma (ididesc) "
            StrSql = StrSql & " VALUES('" & UCase(Left(ididesc, 30)) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(idinro) AS Maxidinro FROM idioma "
            OpenRecordset StrSql, rs_sub
                
            TraerCodIdioma = CLng(rs_sub!Maxidinro)
        Else
            TraerCodIdioma = CLng(rs_sub!idinro)
        End If
    End If
End Function
Public Function TraerCodIdiNivel(idnivdesabr As String)
    If Not EsNulo(idnivdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT idnivnro FROM idinivel WHERE upper(idnivdesabr) = '" & UCase(Left(idnivdesabr, 30)) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = "INSERT INTO idinivel (idnivdesabr,idnivvalor) "
            StrSql = StrSql & " VALUES('" & UCase(Left(idnivdesabr, 30)) & "',0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = " SELECT MAX(idnivnro) AS Maxidnivnro FROM idinivel "
            OpenRecordset StrSql, rs_sub
                
            TraerCodIdiNivel = CLng(rs_sub!Maxidnivnro)
        Else
            TraerCodIdiNivel = CLng(rs_sub!idnivnro)
        End If
    Else
        TraerCodIdiNivel = 0
    End If
End Function

Function validatelefono(cadena As String)
    Dim a As Long
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

' ---------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------

Public Function TraerCodTipoDocumento_2(Sigla As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Sigla) Then
        StrSql = " SELECT tidnro FROM tipodocu WHERE tidnro = " & Sigla
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodTipoDocumento_2 = 0
        Else
            TraerCodTipoDocumento_2 = rs_sub!tidnro
        End If
    Else
        TraerCodTipoDocumento_2 = 0
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerTelefonosMask(ByVal Modnro As Long, ByVal tipotel As Long, ByVal ModTelcarga As String, ByRef Telnro As String, ByRef Telcarga As String, ByRef TelNroMask As String, ByRef TelCargaMask As String, ByVal Paisnro, ByVal Provnro, ByVal Zonanro, ByVal Partnro, ByVal Locnro, ByVal Vianro)
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve telefonos enmascarados
' Autor      : Gonzalez Nicolás
' Fecha      : 09/12/2015
' Ultima Mod.: 10/12/2015 - Gonzalez Nicolás - Se agrega control de telnro enmascarado
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Dim Telpos1 As Long
    Dim DatCarga As String
    Dim ArrDatCargaTel
    Dim UsaTabDef
    Dim MaxLenght
    Dim Tabla
    Dim Campo
    Dim Mask
    Dim MaskOblig
    Dim EjMask
    Dim Expr
    Dim ArrParam(6, 2)
    Dim a
    Dim ArrModTelcarga
    'Dim Telcarga
    'Dim TelNroMask
    'Dim TelCargaMask
    
    Dim TelCargaMaskAux
    Dim TelnroAux
    
    Dim retvalAux
    Dim tablaAux
    Dim rsTel As New ADODB.Recordset
    
    Dim posicion, Mask2
    Dim posicionNew
    Dim esNumero
    
    Dim modtelmask
    modtelmask = False
    'ENMASCARA telnro = telnromask
    StrSql = "SELECT modtelmask FROM modelo_etiq_tel WHERE modnro = " & Modnro & " AND modtelmask = -1"
    OpenRecordset StrSql, rsTel
    If Not rsTel.EOF Then
        modtelmask = True
    End If
    
    ArrModTelcarga = Split(ModTelcarga, "#!#")
    
    'Flog.writeline "ModTelcarga;" & ModTelcarga
    'Flog.writeline "TELNRO;" & telnro
    'Flog.writeline "Telcarga;" & Telcarga
    'Flog.writeline "TelNroMask;" & TelNroMask
    'Flog.writeline "TelCargaMask; " & TelCargaMask
    
    'Byval paisnro,Byval provnro,Byval zonanro,Byval partnro,Byval locnro,Byval vianro
    
    'SETEO PARAMETROS FIJOS
    ArrParam(0, 0) = ""
    ArrParam(0, 1) = 0
    'POSICIÓN DEBE CORRESPONDER AL ID DE LA TABLA modelo_etiq_oritel | campo =>| modcodorinro
    ArrParam(1, 0) = "paisnro"
    ArrParam(1, 1) = Paisnro
    
    ArrParam(2, 0) = "provnro"
    ArrParam(2, 1) = Provnro
    
    ArrParam(3, 0) = "zonanro"
    ArrParam(3, 1) = Zonanro
    
    ArrParam(4, 0) = "partnro"
    ArrParam(4, 1) = Partnro
    
    ArrParam(5, 0) = "locnro"
    ArrParam(5, 1) = Locnro
    
    ArrParam(6, 0) = "vianro"
    ArrParam(6, 1) = Vianro

    'Codigo Pais[-1,3,1,,+**,0,0+,,-1]
    'LDN [-1,3,2,,(***),0,E(),,-1]
    'Número [0,10,0,,**********,0,,,0]
    TelnroAux = ""
    For a = 0 To UBound(ArrModTelcarga)
        Telpos1 = InStr(1, ArrModTelcarga(a), "[") - 1
        '--------------------------------------------------
        'Extraigo label
        'l_Txt = Left(ArrModTelcarga(a), Telpos1)
        '--------------------------------------------------
        'Armo nuevo string limpio para desglosar parametros
        Telpos1 = Telpos1 + 2
        DatCarga = Mid(ArrModTelcarga(a), Telpos1, Len(ArrModTelcarga(a)))
        DatCarga = Left(DatCarga, Len(DatCarga) - 1)
        
        ArrDatCargaTel = Split(DatCarga, ",")
        '--------------------------------------------------
        'Usa tabla default => -1 | Tabla especifica => 0 (debe existir en modelo_etiq_oritel)
        UsaTabDef = ArrDatCargaTel(0)
        '--------------------------------------------------
        'Maxlenght del campo
        MaxLenght = ArrDatCargaTel(1)
        '--------------------------------------------------
        'Tabla de Búsqueda
        Tabla = CLng(ArrDatCargaTel(2))
        '--------------------------------------------------

        '--------------------------------------------------
        'Campo de Búsqueda / Opcional
        Campo = ""
        If ArrDatCargaTel(3) <> "" Then
            Campo = ArrDatCargaTel(3)
        End If
        
        'Máscara
        Mask = ""
        If ArrDatCargaTel(4) <> "" Then
            Mask = ArrDatCargaTel(4)
            'Actualizo al len de la máscara
            MaxLenght = Len(Mask)
        End If

        'Máscara Obligatoria [-1] [0]
        MaskOblig = "0" 'default no obligatorio
        If ArrDatCargaTel(5) <> "" Then
            MaskOblig = ArrDatCargaTel(5)
        End If
        
        'Ej de Máscara
        EjMask = ""
        If ArrDatCargaTel(6) <> "" Then
            EjMask = ArrDatCargaTel(6)
        End If

        'Debe ser una Expresión regular.
        Expr = ""
        If ArrDatCargaTel(7) <> "" Then
            Expr = ArrDatCargaTel(7)
        End If

        'Habilita Verificador de código
'        l_Verif = False
'       If ArrDatCargaTel(8) <> "" Then
'          If CStr(ArrDatCargaTel(8)) = "-1" Then
'                l_Verif = True
'            End If
'        End If
                
        'BUSCO EL VALOR S/ CORRESPONDA

        'OBTENGO NOMBRE DE CAMPO Y TABLA --
        retvalAux = ""
        If UBound(ArrParam) >= Tabla Then 'Controlo que el código que venga no sea superior al tope de arreglo
            If CLng(UsaTabDef) = -1 And Tabla <> 0 Then
                'Obtiene valor de tabla Genérica para -> (PAIS,PROVINCIA,ZONA,PARTIDO,LOCALIDAD,VIA
                StrSql = "SELECT modelo_etiq_codtel.modcodorinro,modcodoriid,modcodtelcod FROM modelo_etiq_codtel"
                StrSql = StrSql & " INNER JOIN modelo_etiq_oritel on modelo_etiq_oritel.modcodorinro = modelo_etiq_codtel.modcodorinro"
                StrSql = StrSql & " WHERE modelo_etiq_codtel.modcodorinro = " & Tabla
                StrSql = StrSql & " AND modcodoriid = " & ArrParam(Tabla, 1)
               ' Flog.writeline StrSql
                OpenRecordset StrSql, rsTel
                If Not rsTel.EOF Then
                    retvalAux = rsTel!modcodtelcod
                End If
                rsTel.Close
            Else
                If Tabla <> 0 Then
                    StrSql = "SELECT modcodoritabla,modcodoricampo FROM modelo_etiq_oritel WHERE modcodorinro = " & Tabla
                    OpenRecordset StrSql, rsTel
                    If Not rsTel.EOF Then
                        tablaAux = rsTel!modcodoritabla
                        If Campo = "" Then 'SI NO HAY UN CAMPO ESPECIFICO DEFINIDO ASIGNO EL DEFAULT DE TABLA
                            Campo = rsTel!modcodoricampo
                        End If
                    End If
                    rsTel.Close
        
                    'Obtiene campo de la tabla seleccionada (PAIS,PROVINCIA,ZONA,PARTIDO,LOCALIDAD,VIA)
                    StrSql = "SELECT " & Campo & " FROM " & tablaAux
                    StrSql = StrSql & " WHERE " & ArrParam(Tabla, 0) & " =" & ArrParam(Tabla, 1)
                    OpenRecordset StrSql, rsTel
                    If Not rsTel.EOF Then
                        retvalAux = rsTel(0)
                    End If
                End If
            End If
        End If
        
       
       
        '----------------------------------------------------------------
        'Recupero posición y simbolo a mostrar
        If (EjMask <> "") Then
            '1er caracter
            posicion = Left(EjMask, 1)
            '2do caracter
            Mask = Mid(EjMask, 2, Len(EjMask))
            If (Len(EjMask) > 2) Then
                '2do caracter
                Mask = Mid(EjMask, 2, 1)
                If (Len(EjMask) = 3) Then
                    '3er caracter
                    Mask2 = Mid(EjMask, 3, Len(EjMask))
                End If
            End If
        End If
        
       
        posicionNew = 0
        '----------------------------------------------------------------
        If (EjMask <> "") Then
            If (retvalAux <> "") Then
                If (posicion = "L") Then 'El parametro se muestra al final del string
                     posicionNew = -1
                ElseIf (posicion = "E") Then 'Concatena al principio y final
                     posicionNew = -2
                Else 'El parametro se muestra en la posición indicada
                     posicionNew = CLng(posicion)
                End If
            End If
        End If
        
        
        'ES UN NÚMERO SIN CODIGOS | FORMATEO S/CARACTERES CONFIGURADOS
        esNumero = False
        TelCargaMaskAux = ""
        If CLng(UsaTabDef) = 0 And Tabla = 0 Then
            retvalAux = Telnro
            'APLICO LA MÁSCARA SI ES QUE TIENE
            TelCargaMaskAux = Telnro
            esNumero = True
        End If
        
        
        'TelnroAux = ""
        
        ' ARMO TELEFONO PARA GRABAR CON MÁSCARA
        parte1 = ""
        parte2 = ""
        parte3 = ""
        If esNumero = False Then
            If (posicionNew > 0) Then
                If (Len(TextAux) >= posicion) Then
                    If (Mid(TextAux, posicion - 1, posicion) <> Mask) Then
                        parte1 = Mid(TextAux, 0, posicion - 1)
                        parte2 = Mid(TextAux, posicion - 1, Len(TextAux))
                        TextAux2 = CStr(parte1) & CStr(Mask) & CStr(parte2)
                    Else
                        TextAux2 = TextAux
                    End If
                Else
                    TextAux2 = TextAux
                End If
            Else
                If Mask <> "" Then
                    If posicion = "L" Then ' CONCATENO MASCARA AL FINAL POSICIÓN
                        'TextAux2 = TextAux2.toString() + sep + TextAux.toString() + Mask.toString();
                    ElseIf posicion = "E" Then ' CONCATENO AL PRINCIPIO Y FINAL
                        'TextAux2 = TextAux2.toString() + Mask.toString() + TextAux.toString() + Mask2.toString();
                        'EXCEPCION PARA PERU (CELULAR) FIJO 1
                        If Modnro = 10 And tipotel = 2 Then
                            retvalAux = "1"
                        End If
                        TelCargaMaskAux = CStr(Mask) & CStr(retvalAux) & CStr(Mask2)
                        'TelnroAux = ""
                    Else ' CONCATENO MASCARA EN 1er POSICIÓN
                        TelCargaMaskAux = CStr(Mask) & CStr(retvalAux)
                    End If
                
                Else 'NO UTILIZA MASCARA%>
                    'TextAux2 = TextAux2.toString() + TextAux.toString();
                    TelCargaMaskAux = CStr(TelCargaMaskAux) & CStr(retvalAux)
                End If
            End If
        End If

    
        If a = 0 Then
             'POSICIÓN
            Telcarga = IIf(retvalAux <> "", a & "," & retvalAux, "")
            'MASCARA APLICADA
            TelNroMask = IIf(retvalAux <> "", CStr(TelCargaMaskAux), ArrDatCargaTel(4))
            'MASCARA Y POSICIÓN APLICADA
            TelCargaMask = a & "," & IIf(retvalAux <> "", CStr(TelCargaMaskAux), ArrDatCargaTel(4))
        Else
            'POSICIÓN
            If Telcarga = "" Then
                Telcarga = Telcarga & IIf(retvalAux <> "", a & "," & retvalAux, "")
            Else
                Telcarga = Telcarga & IIf(retvalAux <> "", "#!#" & a & "," & retvalAux, "")
            End If
            'MASCARA APLICADA
            TelNroMask = CStr(TelNroMask) & " " & IIf(retvalAux <> "", CStr(TelCargaMaskAux), ArrDatCargaTel(4))
            'MASCARA Y POSICIÓN APLICADA
            TelCargaMask = CStr(TelCargaMask) & "#!#" & a & "," & IIf(retvalAux <> "", CStr(TelCargaMaskAux), ArrDatCargaTel(4))
        End If
        TelnroAux = CStr(TelnroAux) & CStr(retvalAux)
    Next
    Telnro = TelnroAux
    
    If modtelmask = True Then 'MUESTRA telnro = telnromask
        Telnro = TelNroMask
    End If
    'FIN DESGLOSE TELEFONOS
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    '----------------------------------------------------------------------------------------|
End Function
Public Function TraerCodNacionalidad_2(Nacionaldes As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Nacionaldes) Then
        StrSql = " SELECT Nacionalnro FROM Nacionalidad WHERE Nacionalnro = " & Nacionaldes
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodNacionalidad_2 = 0
        Else
            TraerCodNacionalidad_2 = rs_sub!nacionalnro
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodEstadoCivil_2(EstCivdesabr As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
Dim Aux_Nro_Estcivil As Long

    Aux_Nro_Estcivil = 0
    If Not EsNulo(EstCivdesabr) Then
        StrSql = " SELECT estcivnro FROM estcivil WHERE estcivnro = " & EstCivdesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            Aux_Nro_Estcivil = 0
        Else
            Aux_Nro_Estcivil = rs_sub!estcivnro
        End If
    Else
        Aux_Nro_Estcivil = 0
    End If
    TraerCodEstadoCivil_2 = Aux_Nro_Estcivil
    
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function


Public Function TraerCodPais_2(Paisdesc As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Paisdesc) Then
        StrSql = " SELECT paisnro FROM Pais WHERE paisnro = " & Paisdesc
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodPais_2 = 1
        Else
            TraerCodPais_2 = rs_sub!Paisnro
        End If
    Else
        TraerCodPais_2 = 1
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function


Public Function TraerCodLocalidad_2(Localidad As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Localidad) Then
        StrSql = " SELECT locnro FROM localidad WHERE locnro = " & Localidad
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodLocalidad_2 = 1
        Else
            TraerCodLocalidad_2 = rs_sub!Locnro
        End If
    Else
        TraerCodLocalidad_2 = 1 'NO INFORMADA
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodProvincia_2(Provincia As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Provincia) Then
        StrSql = " SELECT provnro FROM Provincia WHERE provnro = " & Provincia
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodProvincia_2 = 1 'no informada
        Else
            TraerCodProvincia_2 = rs_sub!Provnro
        End If
    Else
        TraerCodProvincia_2 = 1 'no informada
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function


Public Function TraerCodPartido_2(Partido As String) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Partido) Then
        StrSql = " SELECT partnro FROM Partido WHERE partnro = " & Partido
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodPartido_2 = 1
        Else
            TraerCodPartido_2 = rs_sub!Partnro
        End If
    Else
        TraerCodPartido_2 = 1 'Sin datos
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function


Public Function TraerCodZona_2(Zona As String, Provnro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Zona) Then
        StrSql = " SELECT zonanro FROM Zona WHERE zonanro = " & Zona
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodZona_2 = 0
        Else
            TraerCodZona_2 = rs_sub!Zonanro
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodProcedencia_2(Prodesabr As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Trim(Left(Prodesabr, 30))) Then
        StrSql = " SELECT pronro FROM pos_procedencia WHERE pronro = " & Prodesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodProcedencia_2 = 0
        Else
            TraerCodProcedencia_2 = CLng(rs_sub!pronro)
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function






Public Function TraerCodNivelEstudio_2(nivdesc As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    
    If Not EsNulo(nivdesc) Then
        StrSql = " SELECT nivnro FROM nivest WHERE nivnro = " & nivdesc
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodNivelEstudio_2 = 0
        Else
            TraerCodNivelEstudio_2 = CLng(rs_sub!nivnro)
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodCarrera_2(Carredudesabr As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset

    If Not EsNulo(Carredudesabr) Then
        StrSql = " SELECT carredunro FROM cap_carr_edu WHERE carredunro = " & Carredudesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodCarrera_2 = "NULL"
        Else
            TraerCodCarrera_2 = CLng(rs_sub!CarreduNro)
        End If
    Else
        TraerCodCarrera_2 = "NULL"
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodCausa_2(caudes As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    
    If Not EsNulo(caudes) Then
        StrSql = " SELECT caunro FROM causa WHERE caunro = " & caudes
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodCausa_2 = 0
        Else
            TraerCodCausa_2 = CLng(rs_sub!caunro)
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodTitulo_2(Titdesabr As String, nivnro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Titdesabr) Then
        StrSql = " SELECT titnro FROM titulo WHERE titnro = " & Titdesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodTitulo_2 = "Null"
        Else
            TraerCodTitulo_2 = CLng(rs_sub!TitNro)
        End If
    Else
        TraerCodTitulo_2 = "Null"
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodTituloSolo_2(Titdesabr As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Titdesabr) Then
        StrSql = " SELECT titnro FROM titulo WHERE titnro = " & Titdesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodTituloSolo_2 = 0
        Else
            TraerCodTituloSolo_2 = CLng(rs_sub!TitNro)
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodInstitucion_2(Instdes As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
Dim Arreglo
Dim cadena As String
Dim a As Long
    
    TraerCodInstitucion_2 = 0
    
    If Not EsNulo(Instdes) Then
        StrSql = " SELECT instnro FROM institucion WHERE instnro = " & Instdes
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodInstitucion_2 = 0
        Else
            TraerCodInstitucion_2 = CLng(rs_sub!InstNro)
        End If
    End If
    
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodInstitucionAbreviada_2(Instabre As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(Instabre) Then
        StrSql = " SELECT instnro FROM institucion WHERE instnro = " & Instabre
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodInstitucionAbreviada_2 = 7
        Else
            TraerCodInstitucionAbreviada_2 = CLng(rs_sub!InstNro)
        End If
    Else
        TraerCodInstitucionAbreviada_2 = 7 'NO informada
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodCargo_2(Cardesabr As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset

    If Not EsNulo(Cardesabr) Then
        StrSql = " SELECT carnro FROM cargo WHERE carnro = " & Cardesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodCargo_2 = 0
        Else
            TraerCodCargo_2 = CLng(rs_sub!carnro)
        End If
    Else
        TraerCodCargo_2 = 0
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodTipoCurso_2(tipcurdesabr As String) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset

    TraerCodTipoCurso_2 = 0
    
    If Not EsNulo(tipcurdesabr) Then
        StrSql = " SELECT tipcurnro FROM cap_tipocurso WHERE tipcurnro = " & tipcurdesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodTipoCurso_2 = 0
        Else
            TraerCodTipoCurso_2 = CLng(rs_sub!tipcurnro)
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodCurso_2(curdesabr As String, tipcurnro As Long) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    If Not EsNulo(curdesabr) Then
        StrSql = " SELECT curnro FROM cap_curso WHERE curnro = " & curdesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodCurso_2 = 0
        Else
            TraerCodCurso_2 = CLng(rs_sub!curnro)
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function


Public Function TraerCodListaEmpresa_2(lempdes As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Rs_Estr As New ADODB.Recordset

    lempdes = Left(lempdes, 60)
    If Not EsNulo(lempdes) Then
        StrSql = " SELECT lempnro FROM listaemp WHERE lempnro = " & lempdes
        OpenRecordset StrSql, Rs_Estr
        If Not Rs_Estr.EOF Then
            TraerCodListaEmpresa_2 = CLng(Rs_Estr!lempnro)
        Else
            TraerCodListaEmpresa_2 = 0
        End If
    Else
        TraerCodListaEmpresa_2 = 0
    End If
If Rs_Estr.State = adStateOpen Then Rs_Estr.Close
Set Rs_Estr = Nothing
End Function

Public Function TraerCodIdioma_2(ididesc As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset

    If Not EsNulo(ididesc) Then
        StrSql = " SELECT idinro FROM Idioma WHERE idinro = " & ididesc
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodIdioma_2 = 0
        Else
            TraerCodIdioma_2 = CLng(rs_sub!idinro)
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodIdiNivel_2(idnivdesabr As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    
    If Not EsNulo(idnivdesabr) Then
        StrSql = " SELECT idnivnro FROM idinivel WHERE idnivnro = " & idnivdesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodIdiNivel_2 = 0
        Else
            TraerCodIdiNivel_2 = CLng(rs_sub!idnivnro)
        End If
    Else
        TraerCodIdiNivel_2 = 0
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodEspecializacion_2(espdesabr As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    
    If Not EsNulo(espdesabr) Then
        Flog.writeline Espacios(Tabulador * 2) & "Especializacion: " & espdesabr
        StrSql = " SELECT espnro FROM especializacion WHERE espnro = " & espdesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodEspecializacion_2 = 0
        Else
            TraerCodEspecializacion_2 = CLng(rs_sub!espnro)
        End If
    Else
        Flog.writeline Espacios(Tabulador * 2) & "Especializacion NULA"
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodNivelEspecializacion_2(espnivdesabr As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : FGZ
' Fecha      : 02/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
    
    If Not EsNulo(espnivdesabr) Then
        StrSql = " SELECT espnivnro FROM espnivel WHERE espnivnro = " & espnivdesabr
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodNivelEspecializacion_2 = 0
        Else
            TraerCodNivelEspecializacion_2 = CLng(rs_sub!espnivnro)
        End If
    End If
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function TraerCodTipoNota_2(tiponota As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Customizacion de la Funcion.
' Autor      : RCH
' Fecha      : 30/08/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_sub As New ADODB.Recordset
Dim Arreglo
Dim cadena As String
Dim a As Long
    
    TraerCodTipoNota_2 = 0
    
    If Not EsNulo(tiponota) Then
        StrSql = " SELECT tnonro FROM tiponota WHERE tnonro = " & tiponota
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodTipoNota_2 = 0
        Else
            TraerCodTipoNota_2 = CLng(rs_sub!tnonro)
        End If
    End If
    
If rs_sub.State = adStateOpen Then rs_sub.Close
Set rs_sub = Nothing
End Function

Public Function cantidadCampos(Linea As String, separador As String) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve la cantidad de campos que tiene una linea de la exportación
' Autor      : Gustavo Ring
' Fecha      : 30/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim cant As Long
Dim I As Long
Dim campos As Integer
Dim anterior As String

    campos = 0
    cant = Len(Linea)
    anterior = ""
    I = 1
    While I <= cant
        If Mid(Linea, I, 1) = separador Then
            campos = campos + 1
        End If
        If anterior = ";" And Mid(Linea, I, 1) = ";" Then
                I = cant
        End If
        I = I + 1
    Wend
    cantidadCampos = campos
End Function

Public Function TraerCodEventoCrear(evecodext As String, evedesabr As String, curnro As Integer, centrocap As Integer) As Integer

' ---------------------------------------------------------------------------------------------
' Descripcion: devuelve el cod del evento, si no existe lo crea
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.: 25/07/2007 Gustavo Ring - se inicializan campos de costos y cantidad de alumnos
' Descripcion:
' ---------------------------------------------------------------------------------------------

  If Not EsNulo(evecodext) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT evenro FROM cap_evento WHERE evecodext = '" & evecodext & "'"
        StrSql = StrSql & " AND evedesabr='" & evedesabr & "'"
        
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_evento (evedesabr , evecodext,estevenro,eveubi,"
            StrSql = StrSql & " eveforeva,curnro,centrocap,eveniveva,eveorigen,evereqasi,eveabierto, "
            StrSql = StrSql & " evecostoind,evecanplaalu,evecanrealalu,evecostogral) "
            StrSql = StrSql & " VALUES ('" & Left(UCase(evedesabr), 50)
            StrSql = StrSql & "','" & evecodext & "',1,-1,1," & curnro & "," & centrocap
            StrSql = StrSql & ",0,1,0,-1,0,0,0,0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(evenro) maxevenro FROM cap_evento "
            OpenRecordset StrSql, rs_sub
            TraerCodEventoCrear = rs_sub!maxevenro
        Else
            TraerCodEventoCrear = rs_sub!evenro
        End If
    End If
    
End Function

Public Sub pasaracurso(evenro As Integer)

' ---------------------------------------------------------------------------------------------
' Descripcion: Pasa un evento a estado en curso
' Autor      : Gustavo Ring
' Fecha      : 25/07/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
  
  Dim rs_sub1 As New ADODB.Recordset
   
   StrSql = " UPDATE cap_evento SET estevenro= 5 WHERE evenro = " & evenro
   OpenRecordset StrSql, rs_sub1

End Sub
Public Function TraerCodEvento(evecodext As String, evedesabr As String) As Integer

' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve 0 si el evento no existe, sino devuelve evenro
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim eventonro As Integer

    eventonro = 0
    
    If Not EsNulo(evecodext) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT evenro FROM cap_evento WHERE evecodext = '" & evecodext & "'"
        StrSql = StrSql & " AND evedesabr='" & evedesabr & "'"
        OpenRecordset StrSql, rs_sub
        If Not (rs_sub.EOF) Then
            eventonro = rs_sub!evenro
        End If
    End If
    
    TraerCodEvento = eventonro
    
End Function

Public Sub ActualizarCantParticipantes(evenro As Integer)

' ---------------------------------------------------------------------------------------------
' Descripcion: Actualiza cantidad de participantes del evento
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim cantPart As Integer

    cantPart = 0
    
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT evecanplaalu FROM cap_evento WHERE evenro = " & evenro
        OpenRecordset StrSql, rs_sub
        cantPart = rs_sub("evecanplaalu") + 1
        StrSql = "UPDATE cap_evento SET evecanplaalu=" & cantPart
        StrSql = StrSql & ",evecanrealalu = " & cantPart
        StrSql = StrSql & " WHERE evenro= '" & evenro & "'"
        objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub ActualizarFechas(evenro As Integer)

' ---------------------------------------------------------------------------------------------
' Descripcion: Actualiza Fecha de Inicio y Fin del evento
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim fechafin As Date
Dim fechaini As Date

        Dim rs_sub As New ADODB.Recordset
        
        StrSql = " SELECT * FROM cap_eventomodulo "
        StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro=cap_calendario.evmonro "
        StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & evenro
        StrSql = StrSql & " ORDER BY calfecha ASC"
        
        OpenRecordset StrSql, rs_sub
        fechainicio = rs_sub("calfecha")
        StrSql = "UPDATE cap_evento SET evefecini=" & ConvFecha(fechainicio)
        StrSql = StrSql & " WHERE evenro= " & evenro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        rs_sub.Close
        
        StrSql = " SELECT * FROM cap_eventomodulo "
        StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro=cap_calendario.evmonro "
        StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & evenro
        StrSql = StrSql & " ORDER BY calfecha DESC"
        
        OpenRecordset StrSql, rs_sub
        fechafin = rs_sub("calfecha")
        StrSql = "UPDATE cap_evento SET evefecfin =" & ConvFecha(fechafin)
        StrSql = StrSql & " WHERE evenro= " & evenro
        objConn.Execute StrSql, , adExecuteNoRecords
        
End Sub

Public Function TraerCodEvento2(evecodext As String) As Integer

' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve 0 si el evento no existe, sino devuelve evenro
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim eventonro As Integer

    eventonro = 0
    
    If Not EsNulo(evecodext) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT evenro FROM cap_evento WHERE evecodext = '" & evecodext & "'"
        OpenRecordset StrSql, rs_sub
        If Not (rs_sub.EOF) Then
            eventonro = rs_sub!evenro
        End If
    End If
    
    TraerCodEvento2 = eventonro
    
End Function

Public Function TraerCodTerno(empleg As String) As Integer

' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve 0 si el tercero no existe, sino devuelve ternro
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Ternro As Long

    Ternro = 0
    
    If Not EsNulo(empleg) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT ternro FROM empleado WHERE empleg = " & empleg
        OpenRecordset StrSql, rs_sub
        If Not (rs_sub.EOF) Then
            Ternro = rs_sub!Ternro
        End If
    End If
    
    TraerCodTerno = Ternro
    
End Function

Public Function TraerCodResponsable(razonsocial As String) As Integer
    
' --------------------------------------------------------------------------------------------------
' Descripcion: Devuelve el código del responsable de la formación si no existe crea un nuevo Tercero
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' --------------------------------------------------------------------------------------------------
    
    If Not EsNulo(razonsocial) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT tercero.ternro FROM tercero "
        StrSql = StrSql & " INNER JOIN ter_tip on tipnro=15 and ter_tip.ternro=tercero.ternro "
        StrSql = StrSql & " WHERE tercero.terrazsoc = '" & UCase(razonsocial) & "'"
        
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
            StrSql = " INSERT INTO tercero (terrazsoc , tersex) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(razonsocial), 60) & "', -1 )"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = " SELECT MAX(ternro) terceronro FROM tercero "
            OpenRecordset StrSql, rs_sub
            TraerCodResponsable = rs_sub!terceronro
        Else
            TraerCodResponsable = rs_sub!Ternro
        End If
    End If
End Function

Public Sub AsignarTerceroComoCentroCap(Ternro As Integer)
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Asigna al tercero la propiedad que es centro de formación
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.: 25/07/2007 - Gustavo Ring - Se graba en cap_centrocap al formador
' Descripcion:
' ---------------------------------------------------------------------------------------------
         
    If Not EsNulo(Ternro) Then
        Dim rs_sub As New ADODB.Recordset
        
        StrSql = " SELECT * FROM ter_tip "
        StrSql = StrSql & " WHERE tipnro = 15 and ternro=" & Ternro
        
        OpenRecordset StrSql, rs_sub
                
        If rs_sub.EOF Then
            StrSql = " INSERT INTO ter_tip (tipnro , ternro) "
            StrSql = StrSql & "  VALUES (15 ," & Ternro & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        rs_sub.Close
        
        StrSql = " SELECT * FROM cap_centrocap "
        StrSql = StrSql & " WHERE tipcennro = 1 and ternro=" & Ternro
        
        OpenRecordset StrSql, rs_sub
                
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_centrocap (tipcennro , ternro) "
            StrSql = StrSql & "  VALUES (1 ," & Ternro & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
    End If
End Sub



Public Function TraerCodCursoSinCrear(curcodext As String) As Integer

' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve 0 si el curso no existe, sino devuelve evenro
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim curnro As Integer

    curnro = 0
    
    If Not EsNulo(curcodext) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT curnro FROM cap_curso WHERE curcodext = '" & UCase(curcodext) & "'"
        OpenRecordset StrSql, rs_sub
        If Not (rs_sub.EOF) Then
            curnro = rs_sub!curnro
        End If
    End If
    
    TraerCodCursoSinCrear = curnro
    
End Function


Public Function TraerCodCursoDeloitte(curdesabr As String, tipcurnro As Long, curcodext As String) As Integer
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve curnro, si no existe lo crea
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    If Not EsNulo(curdesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT curnro FROM cap_curso WHERE curdesabr = '" & UCase(curdesabr) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_curso (curdesabr,curcodext, tipcurnro) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(curdesabr), 50) & "','" & Left(UCase(curcodext), 25) & "', " & tipcurnro & " )"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(curnro) AS Maxcurnro FROM cap_curso "
            OpenRecordset StrSql, rs_sub

            TraerCodCursoDeloitte = rs_sub!Maxcurnro
        Else
            TraerCodCursoDeloitte = rs_sub!curnro
        End If
    End If
End Function

Public Function TraerCodModulo(moddesabr As String) As Integer
    
' ---------------------------------------------------------------------------------------------------
' Descripcion: Devuelve modnro, si no existe lo crea y se establece la relacion con curso si no esta
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------------
    
    Dim Modnro As Integer

    If Not EsNulo(moddesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT modnro FROM cap_modulo WHERE moddesabr = '" & UCase(moddesabr) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_modulo (moddesabr,tipmodnro, modmodal) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(moddesabr), 50)
            StrSql = StrSql & "',1,-1 )"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(modnro) AS Maxmodnro FROM cap_modulo "
            OpenRecordset StrSql, rs_sub

            TraerCodModulo = rs_sub!Maxmodnro
        Else
            TraerCodModulo = rs_sub!Modnro
        End If
    End If
    
    rs_sub.Close
    
            
End Function

Public Function TraerCodModulo2(moddesabr As String, tipmodnro As Long) As Long
    
' ---------------------------------------------------------------------------------------------------
' Descripcion: Devuelve modnro, si no existe lo crea
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.: Se agrego el tipo de modulo, si lo encuentra por nombre devuelve el valos, sino lo crea con el tipo de modulo
' Descripcion:
' ---------------------------------------------------------------------------------------------------
    
Dim Modnro As Integer

    If Not EsNulo(moddesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT modnro FROM cap_modulo WHERE moddesabr = '" & UCase(moddesabr) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_modulo (moddesabr,tipmodnro, modmodal) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(moddesabr), 50)
            StrSql = StrSql & "'," & tipmodnro & " ,-1 )"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(modnro) AS Maxmodnro FROM cap_modulo "
            OpenRecordset StrSql, rs_sub

            TraerCodModulo2 = rs_sub!Maxmodnro
            Modnro = rs_sub!Maxmodnro
        Else
            TraerCodModulo2 = rs_sub!Modnro
            Modnro = rs_sub!Modnro
        End If
    End If
    
    rs_sub.Close
    
            
End Function

Public Function TraerCodModulo0(moddesabr As String) As Integer
    
' ---------------------------------------------------------------------------------------------------
' Descripcion: Devuelve modnro, si no existe devuelve 0
' Autor      : lisandro Moro
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------------
    
Dim Modnro As Integer

    If Not EsNulo(moddesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT modnro FROM cap_modulo WHERE moddesabr = '" & UCase(moddesabr) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            TraerCodModulo0 = 0
        Else
            TraerCodModulo0 = rs_sub!Modnro
        End If
    End If
    
    rs_sub.Close
    
            
End Function

Public Function TraerCodTipoModulo(tipmoddesabr As String) As Long
    
' ---------------------------------------------------------------------------------------------------
' Descripcion: Devuelve tipmodnro, si no existe lo crea y se establece
' Autor      : Lisandro Moro
' Fecha      : 09/02/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------------
    
    Dim tipmodnro As Integer

    If Not EsNulo(tipmoddesabr) Then
        Dim rs_sub As New ADODB.Recordset
        StrSql = " SELECT tipmodnro FROM cap_tipomodulo WHERE tipmoddesabr = '" & UCase(tipmoddesabr) & "'"
        OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_tipomodulo (tipmoddesabr) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(tipmoddesabr), 50) & "')"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(tipmodnro) AS Maxtipmodnro FROM cap_tipomodulo "
            OpenRecordset StrSql, rs_sub

            TraerCodTipoModulo = rs_sub!Maxtipmodnro
        Else
            TraerCodTipoModulo = rs_sub!tipmodnro
        End If
    End If
    
    rs_sub.Close
    
            
End Function


Public Function actualizar_candidato(Ternro As Integer, evenro As Integer) As Boolean
' ---------------------------------------------------------------------------------------------------
' Descripcion: Actualiza la tabla de candidatos
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------------

    Dim rs_sub As New ADODB.Recordset
    
    StrSql = " SELECT * FROM cap_candidato WHERE evenro = " & evenro
    StrSql = StrSql & " AND ternro=" & Ternro
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_candidato (evenro,ternro,conf,recdip,confpart,invitado,invext) "
            StrSql = StrSql & "  VALUES (" & evenro & "," & Ternro & ",-1,0,-1,-1,0"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Call ActualizarCantParticipantes(evenro)
            actualizar_candidato = True
    Else
            actualizar_candidato = False
    End If
            
End Function

Public Sub actualizar_calendario_participante(Ternro As Integer, evenro As Integer)
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Relaciona un participante con todos los calendarios del evento
' Autor      : Gustavo Ring
' Fecha      : 25/07/2007
' Ultima Mod.: 09/02/2010 - Lisandro Moro - Se agrego la validacion si el ya existe la relacion ternro, calnro
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_cal As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
        
    StrSql = " SELECT calnro FROM cap_eventomodulo "
    StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro "
    StrSql = StrSql & " WHERE evenro = " & evenro
    OpenRecordset StrSql, rs_cal
    If Not rs_cal.EOF Then
        While Not rs_cal.EOF
            StrSql = " SELECT * FROM cap_partcal WHERE ternro = " & Ternro & " AND calnro = " & rs_cal("calnro")
            OpenRecordset StrSql, rs
            If rs.EOF Then
                StrSql = " INSERT INTO cap_partcal (ternro,calnro) "
                StrSql = StrSql & "  VALUES (" & Ternro & "," & rs_cal("calnro") & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                Texto = " El participante ya posee el calendario asosciado(ternro:" & Ternro & " - calnro " & rs_cal("calnro") & " ) "
                Call Escribir_Log("flog", NroLinea, NroColumna, Texto, Tabs, strLinea)
            End If
            rs_cal.MoveNext
            rs.Close
        Wend
    Else
        Texto = " No existen calendarios para el evento:" & evenro
        Call Escribir_Log("flog", NroLinea, NroColumna, Texto, Tabs, strLinea)
        Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
    End If
    rs_cal.Close
    Set rs_cal = Nothing
    
End Sub


Public Function TraerCodCalendario(evmonro As Integer, calfecha As String, calhordes As String, calhorhas As String, lugnro As Integer) As Integer
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve calnro, si no existe lo crea
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    Dim Dia As String
    
    StrSql = " SELECT calnro FROM cap_calendario WHERE calfecha = '" & calfecha
    StrSql = StrSql & "' AND calhordes='" & calhordes & "'"
    StrSql = StrSql & " AND calhorhas='" & calhorhas & "'"
    StrSql = StrSql & " AND evmonro=" & evmonro
    StrSql = StrSql & " AND lugnro = " & lugnro
        
    OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
        
            Select Case Weekday(Replace(ConvFecha(CDate(calfecha)), "'", ""))
                Case 1
                    Dia = "Domingo"
                Case 2
                    Dia = "Lunes"
                Case 3
                    Dia = "Martes"
                Case 4
                    Dia = "Miercoles"
                Case 5
                    Dia = "Jueves"
                Case 6
                    Dia = "Viernes"
                Case 7
                    Dia = "Sabado"
            End Select
            
            StrSql = " INSERT INTO cap_calendario (evmonro,calfecha,caldia,calhordes,calhorhas,lugnro) "
            StrSql = StrSql & "  VALUES (" & evmonro
            StrSql = StrSql & "," & ConvFecha(CDate(calfecha)) & ",'" & Dia & "','" & calhordes
            StrSql = StrSql & "','" & calhorhas & "'," & lugnro & ")"
                        
            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(calnro) AS Maxcalnro FROM cap_calendario "
            OpenRecordset StrSql, rs_sub

            TraerCodCalendario = rs_sub!Maxcalnro
        Else
            TraerCodCalendario = rs_sub!calnro
        End If
    
End Function
    
Public Sub actualizar_cap_asistencia(calnro As Integer, Ternro As Integer, asipre As Integer)
    
' ---------------------------------------------------------------------------------------------
' Descripcion: actualiza la tabla de asistencias
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    Dim asievehorini As String
    Dim asievehorfin As String
    
    StrSql = " SELECT calnro,calhordes,calhorhas FROM cap_calendario "
    StrSql = StrSql & " WHERE calnro=" & calnro
        
    OpenRecordset StrSql, rs_sub
    
    If Not rs_sub.EOF Then
            
            If Not IsNull(rs_sub("calhordes")) Then
                  asievehorini = rs_sub("calhordes")
            Else
                  asievehorini = ""
            End If
            
            If Not IsNull(rs_sub("calhorhas")) Then
                  asievehorfin = rs_sub("calhorhas")
            Else
                  asievehorfin = ""
            End If
            
    End If
    
    rs_sub.Close
    
    StrSql = " SELECT calnro,ternro FROM cap_asistencia WHERE ternro = " & Ternro
    StrSql = StrSql & " AND calnro=" & calnro
            
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        StrSql = " INSERT INTO cap_asistencia (calnro,ternro,asievehorini,asievehorfin,asipre) "
        StrSql = StrSql & "  VALUES (" & calnro & "," & Ternro
        StrSql = StrSql & ",'" & asievehorini & "','" & asievehorfin
        StrSql = StrSql & "'," & asipre & ")"
                            
        objConn.Execute StrSql, , adExecuteNoRecords
    
    Else
        StrSql = " UPDATE cap_asistencia SET "
        StrSql = StrSql & " asipre = " & asipre
        StrSql = StrSql & " WHERE ternro = " & Ternro
        StrSql = StrSql & " AND calnro = " & calnro
  
        objConn.Execute StrSql, , adExecuteNoRecords
    
    End If

End Sub
    

Public Function TraerCodCalendario2(calfecha As String, calhordes As String, evenro As Integer) As Integer
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve calnro, sino existe devuelve 0
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    Dim calnro As Integer
    
    StrSql = " SELECT calnro FROM cap_calendario "
    StrSql = StrSql & " INNER JOIN  cap_eventomodulo ON evenro = " & evenro
    StrSql = StrSql & " WHERE calfecha = '" & calfecha & "' AND calhordes='" & calhordes & "'"
    
    calnro = 0
    OpenRecordset StrSql, rs_sub
        If Not (rs_sub.EOF) Then
              calnro = rs_sub!calnro
        End If
    TraerCodCalendario2 = calnro
    
End Function

Public Function TraerCodLugar(lugdesabr As String) As Integer
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve lugnro, si no existe lo crea
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion: ver
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    
    StrSql = " SELECT lugnro FROM cap_lugar WHERE lugdesabr = '" & lugdesabr & "'"
            
    OpenRecordset StrSql, rs_sub
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_lugar (lugdesabr) "
            StrSql = StrSql & "  VALUES ('" & Left(UCase(lugdesabr), 50) & "')"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(lugnro) AS Maxlugnro FROM cap_lugar "
            OpenRecordset StrSql, rs_sub

            TraerCodLugar = rs_sub!Maxlugnro
        Else
            TraerCodLugar = rs_sub!lugnro
        End If
    
End Function
    
Public Function TraerCodEventoModulo(evenro As Integer, Modnro As Integer) As Integer
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve evmonro, si no existe lo crea
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    
    StrSql = " SELECT evmonro FROM cap_eventoModulo WHERE evenro = " & evenro
    StrSql = StrSql & " AND modnro = " & Modnro
            
    OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
            StrSql = " INSERT INTO cap_eventomodulo (evenro,modnro) "
            StrSql = StrSql & "  VALUES (" & evenro & "," & Modnro & ")"

            objConn.Execute StrSql, , adExecuteNoRecords

            StrSql = " SELECT MAX(evmonro) maxevmonro FROM cap_eventomodulo "
            OpenRecordset StrSql, rs_sub

            TraerCodEventoModulo = rs_sub!maxevmonro
        Else
            TraerCodEventoModulo = rs_sub!evmonro
        End If
    
End Function


Public Function controlHora(ByVal hora As String) As Boolean
    
' ---------------------------------------------------------------------------------------------
' Descripcion: controla la hora
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
Dim cant As Integer
Dim I As Integer
Dim controla As Boolean

controla = True

hora = Replace(hora, ":", "")

cant = Len(hora)
If cant <> 4 Then
        controla = False
        
End If

If controla Then
        For I = 1 To 4
          
          If (Asc(Mid(hora, I, 1)) < Asc(0)) Or (Asc(Mid(hora, I, 1)) > Asc(9)) Then
             controla = False
          End If
        Next
End If

If controla Then
        If CInt(Mid(hora, 1, 2)) < "0" Then controla = False
        If CInt(Mid(hora, 1, 2)) > "24" Then controla = False
        If CInt(Mid(hora, 3, 4)) < "0" Then controla = False
        If CInt(Mid(hora, 3, 4)) > "60" Then controla = False
End If

controlHora = controla

End Function

Public Function controlNumero(Numero As String) As Boolean
    
' ---------------------------------------------------------------------------------------------
' Descripcion: controla números con tamaño menor a 5 digitos
' Autor      : Gustavo Ring
' Fecha      : 29/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
Dim cant As Integer
Dim I As Integer
Dim controla As Boolean

controla = True

cant = Len(Numero)
If cant > 5 Then
        controla = False
End If

If cant > 0 Then
        For I = 1 To cant
          
          If (Asc(Mid(Numero, I, 1)) < Asc(0)) Or (Asc(Mid(Numero, I, 1)) > Asc(9)) Then
             controla = False
          End If
        Next
End If

controlNumero = controla

End Function

Public Function esNum(Numero As String) As Boolean
    
' ---------------------------------------------------------------------------------------------
' Descripcion: controla cadenas de números
' Autor      : Gustavo Ring
' Fecha      : 30/08/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
Dim cant As Integer
Dim I As Integer
Dim controla As Boolean

controla = True
cant = Len(Numero)

If cant > 0 Then
        For I = 1 To cant
          
          If (Asc(Mid(Numero, I, 1)) < Asc(0)) Or (Asc(Mid(Numero, I, 1)) > Asc(9)) Then
             controla = False
          End If
        Next
Else
        controla = False
End If

esNum = controla

End Function


Public Function TraerGrado(gradesabr As String) As Integer
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve el código del grado de la banda salarial
' Autor      : Gustavo Ring
' Fecha      : 06/08/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    Dim granro As Integer
    
    StrSql = " SELECT granro FROM  grado "
    StrSql = StrSql & " WHERE gradesabr='" & gradesabr & "'"
    
    granro = 0
    OpenRecordset StrSql, rs_sub
    
    If Not (rs_sub.EOF) Then
            granro = rs_sub!granro
    End If
    TraerGrado = granro
    
End Function

Public Function TraerOrigenBanda(obdesabr As String) As Integer
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve el código del origen de la banda salarial
' Autor      : Gustavo Ring
' Fecha      : 06/08/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    Dim obnro As Integer
    
    StrSql = " SELECT obnro FROM  origenbanda"
    StrSql = StrSql & " WHERE obdesabr='" & obdesabr & "'"
    
    obnro = 0
    OpenRecordset StrSql, rs_sub
    
    If Not (rs_sub.EOF) Then
            obnro = rs_sub!obnro
    End If
    TraerOrigenBanda = obnro
    
End Function

Public Function Superposicionfechas(fecdes As String, fechas As String, granro) As Boolean
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve true si hay superposicion de fechas en el mismo grado
' Autor      : Gustavo Ring
' Fecha      : 06/08/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    Dim superpos As Boolean
        
    StrSql = " SELECT * FROM banda_salarial  "
    StrSql = StrSql & " WHERE granro=" & granro
    StrSql = StrSql & " AND ((bsfecdesde <= " & ConvFecha(CDate(fecdes))
    StrSql = StrSql & " AND bsfechasta >= " & ConvFecha(CDate(fecdes)) & ") OR "
    StrSql = StrSql & " (bsfecdesde <= " & ConvFecha(CDate(fechas))
    StrSql = StrSql & " AND bsfechasta >= " & ConvFecha(CDate(fechas)) & ")) "
        
    superpos = False
    
    OpenRecordset StrSql, rs_sub
    
    If Not (rs_sub.EOF) Then
            superpos = True
    End If
    
    Superposicionfechas = superpos

End Function

Public Function importeValido(importe As String) As Boolean
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Devuelve true si el importe es valido
' Autor      : Gustavo Ring
' Fecha      : 06/08/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
 Dim I As Integer
 Dim datos() As String
 Dim Cantidad As Integer
 Dim esCorrecto As Boolean
 
 esCorrecto = True
 
 datos = Split(importe, NumeroSeparadorDecimal)
 Cantidad = UBound(datos)
 
 For I = 0 To Cantidad
        datos(I) = datos(I)
        If Not (esNum(datos(I))) Then
            esCorrecto = False
        End If
 Next I
       
 importeValido = esCorrecto
 
End Function

Public Function insertarorigen(txt_origen) As Integer
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta el origen y devuelve el obnro
' Autor      : Gustavo Ring
' Fecha      : 05/09/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    
    StrSql = "INSERT INTO origenbanda "
    StrSql = StrSql & "(obdesabr, obdesext) "
    StrSql = StrSql & "VALUES ('" & txt_origen
    StrSql = StrSql & "', '" & txt_origen
    StrSql = StrSql & "')"

    objConn.Execute StrSql, , adExecuteNoRecords
 
    StrSql = " SELECT MAX(obnro) max FROM origenbanda "
    OpenRecordset StrSql, rs_sub

    insertarorigen = rs_sub!Max
        
End Function

Public Function insertargrado(txt_grado) As Integer
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta el grado y devuelve el granro
' Autor      : Gustavo Ring
' Fecha      : 05/09/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
    
    
  
    StrSql = "INSERT INTO grado "
    StrSql = StrSql & "(gradesabr, "
    StrSql = StrSql & " gradesext ) "
    StrSql = StrSql & " values ('"
    StrSql = StrSql & txt_grado & "','"
    StrSql = StrSql & txt_grado & "')"

    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = " SELECT MAX(granro) max FROM grado "
    OpenRecordset StrSql, rs_sub

    insertargrado = rs_sub!Max
        
End Function

Public Sub insertarBanda(bsdesc As String, granro As Integer, bsfecdesde As String, bsfechasta As String, obnro As String, bsinterna As Integer, bszonaa As Double, bszonab As Double, bszonac As Double, bszonaab As Double, bszonabc As Double)
    
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta la banda salarial
' Autor      : Gustavo Ring
' Fecha      : 05/09/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    Dim rs_sub As New ADODB.Recordset
        
    StrSql = "INSERT INTO banda_salarial "
    StrSql = StrSql & "(bsdesc, granro, bsfecdesde, bsfechasta, obnro, bsinterna"
    StrSql = StrSql & ", bszonaa, bszonab, bszonac, bszonaab, bszonabc)"
    StrSql = StrSql & "VALUES ('" & bsdesc & "'," & granro & "," & ConvFecha(bsfecdesde)
    StrSql = StrSql & "," & ConvFecha(bsfechasta) & "," & obnro & "," & bsinterna
    StrSql = StrSql & "," & bszonaa & "," & bszonab & "," & bszonac
    StrSql = StrSql & "," & bszonaab & "," & bszonabc
    StrSql = StrSql & ")"


    objConn.Execute StrSql, , adExecuteNoRecords
  
End Sub

Public Function politicaActiva(polnro As Long) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Chequeo si la politica esta activa
' Autor      : Lisandor Moro
' Fecha      : 13/07/2016
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Dim rs_sub As New ADODB.Recordset

    StrSql = "SELECT cabpolestado FROM gti_cabpolitica "
    StrSql = StrSql & " WHERE cabpolnivel = " & polnro & " AND cabpolestado = -1"
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        politicaActiva = False
    Else
        politicaActiva = True
    End If
    If rs_sub.State = adStateOpen Then rs_sub.Close
    Set rs_sub = Nothing
    
End Function

