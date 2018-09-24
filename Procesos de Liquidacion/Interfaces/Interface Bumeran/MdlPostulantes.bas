Attribute VB_Name = "MdlPostulantes"
Option Explicit

Dim ternro As Long
Dim l_sql As String
Dim NroDom As Integer
Dim idcalificador() As Integer

    Dim terape As String          'a_apellido
    Dim calle As String           'a_calle
    'a_cambiares  (Ver q es?)
    Dim locnro As Integer         'a_ciudad
    Dim codigopostal As String    'a_cp
    Dim oficdepto As String       'a_dpto
    Dim teremail As String        'a_email
    Dim terfecnac As Date         'a_fnacimiento
    'a_idusuario  (desaparece)
    Dim ternom As String          'a_nombre
    Dim nrodoc As String          'a_nrodoc
    Dim nro As String             'a_numero
    Dim paisnro As Integer        'a_pai_idpais
    Dim nacionalnro As Integer       'a_pai_idpais_naciopais
    Dim piso As String            'a_piso
    Dim provnro As String        'a_pro_idprovincia_vivepro
    Dim tersex As Boolean         'a_sexo
    Dim tidnro As Integer         'a_tdd_idtipodedocumento
'- <computacion> (especializaciones eltoama y nivel)
    'idcalificador(ver q desaparece)
    Dim espnro() As Integer        'idconocimiento
    Dim espnivnro() As Integer      'idnivel
'- <curriculum>
    Dim posfecpres() As String        'FechaAlta
    'frecuencia(de cobro)(desaparece)
    Dim posrempre() As Double       'Minimo(sueldo)
    'objetivos (ver)(desaparece)
    'pue_idpuesto(ver)(desaparece)
    'puesto(ver)(desaparece)
    Dim posref() As String          'referencias
    'tdt_idtipodetrabajo (Ver de agregar)
'- <curriculum_area>
    'are_idarea (Ver)(desaparece)
'- <curriculum_industria>
    'ind_idindustria (Ver)(desaparece)
'- <estudio>
    'are_idareaestudio(area q desaparece)
    Dim capfechasta() As String       'ffin
    Dim capfecdesde() As String       'finicio
    Dim instnro() As Integer        'inins_idinstitucion
    Dim institucion() As Integer    'Institucion(Nueva, agregada a mano por el postulante)
    'pai_idpais (Desaparece, no tenemos la relacion con el pais)
    Dim capprom() As String         'promedio
    Dim caprango() As String        '(60)rng_idrango
    Dim nivnro() As Long         'tde_idtipodeestudio
    Dim titulo() As String          'titulo(Nueva, agregada a mano por el postulante)
    Dim titnro() As Integer
'- <experiencialaboral>
    'are_idarea(area q desaparece)
    Dim empatareas() As String      'descripcion
    Dim Empnro() As Integer         'empresa
    Dim empadesde() As String         'ffin
    Dim empahasta() As String         'finicio
    'ind_idindustria (desaparece)
    'pai_idpais (desaparece)
    'pue_idpuesto
    Dim carnro() As Integer         'puesto
'- <idiomas>
    'Dim idcalificador
    Dim idinro() As Integer         'idconocimiento
    Dim idnivel() As String         'idnivel(empidlee,empidhabla, empidescr)
'- <telefono>
    Dim Categoria() As Integer      'Categoria(fax, default o celular)
    Dim telnro() As String          'prefix + Numero
    'prefix (desaparece)
    Dim telfax() As Integer
    Dim teldefault() As Integer
    Dim telcelular() As Integer

Public Sub Insertar_Postulante_Segun_Modelo_Estandar(Rec_postulantes)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador de acurdo al modelo
' Autor      : Lisandro Moro
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'MyBeginTrans
    Select Case NroModelo
    Case 232: 'Interface Postulantes Bumerang
        'Call Bumeran(Rec_postulantes)
        'Inicializo los datos de bumeran
        'Call CargarDatosBumeran
        Call LeerXmlBumeran(Rec_postulantes, 0)
    End Select
'MyCommitTrans
End Sub

Sub LeerXmlBumeran(rs, hijo)
    Dim Columna As String
    Dim rsChild         'ADODB.Recordset
    Dim Col             'ADODB.Field
    Dim rsChils As ADODB.Recordset
    'Set rsChild = Server.CreateObject("ADODB.Recordset")
    Set rsChild = New ADODB.Recordset
    
    On Error Resume Next
   
   ' While rs.EOF <> True
        'ternro = TraerNuevoCodigoPostulante
        For Each Col In rs.Fields
            If Col.Name <> "$Text" Then   ' $Text to be ignored
                If Col.Type <> adChapter Then
                  ' Output the non-chaptered column
                    'MsgBox String((hijo), " ") & Col.Name & ": " & Col.Value & vbCrLf
                    Call Bumeran(Col.Name, Col.Value, CInt(hijo))
                Else
                    'Text2.Text = Text2.Text & vbCrLf
                    ' Retrieve the Child recordset
                    Columna = CStr(Col.Name)
                    Set rsChild = Col.Value
                    If Not rsChild.EOF Then rsChild.MoveFirst
                     If Err Then
                         'Print ("Error: " & Error)
                         'MsgBox ("End")
                     End If
                     Select Case Columna
                         Case "computacion"
                             Call ArmarEspecializaciones(rsChild)
                         Case "curriculum"
                             ArmarComplemento (rsChild)
                         Case "curriculum_area"
                             'rsChild.MoveLast
                         Case "curriculum_industria"
                             'rsChild.MoveLast
                         Case "estudio"
                             ArmarEstudiosFormales (rsChild)
                         Case "experiencialaboral"
                             ArmarEmpleosAnteriores (rsChild)
                         Case "idiomas"
                             ArmarIdiomas (rsChild)
                         Case "telefono"
                             ArmarTelefonos (rsChild)
                     End Select
                    ' LeerXmlBumeran rsChild, hijo + 1
                     
                     rsChild.Close
                     Set rsChild = Nothing
                End If
            Else
                'MsgBox "$Text", , Col.Name & "-" & Col.Value
            End If
        Next
   '     rs.MoveNext
        InsertarPostulanteBumeran
  '  Wend
    'rsChild.Close
    'rsChild = Nothing
End Sub
Function InsertarPostulanteBumeran()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que se encarga de insertar un postulante.
' Autor      : JMH
' Fecha      : 19/04/2006
' Ultima Mod.: FGZ - 11/05/2007
' Descripcion: Le agregué el pasinro a tercero
' ---------------------------------------------------------------------------------------------
    Dim rs_sub As New ADODB.Recordset
    Dim a As Integer
    Dim ActPasos As Boolean
    Dim estact
    Dim carrcomp
    Dim Provincia As Integer
    
    l_sql = "  "
    l_sql = l_sql & ""
    
    Err.Clear
    On Error GoTo ErrorTercero
    
    'Busco si ya existe el Postulante
    nrodoc = Replace(nrodoc, ".", "")
    StrSql = " SELECT * FROM ter_doc "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = ter_doc.ternro "
    StrSql = StrSql & " INNER JOIN ter_tip ON ter_tip.ternro = tercero.ternro AND tipnro = 14 "
    StrSql = StrSql & " WHERE nrodoc = '" & nrodoc & "'"
    OpenRecordset StrSql, rs_sub
    If Not rs_sub.EOF Then
       ternro = rs_sub!ternro
       ModificarPostulanteBumeran
    Else
    
        '--Inserto el Tercero--
        'FGZ - 11/05/2007 - le agregué el paisnro
        StrSql = " INSERT INTO tercero (ternom,terape,terfecnac,tersex,teremail, nacionalnro,paisnro) VALUES ("
        StrSql = StrSql & "'" & ternom & "'"
        StrSql = StrSql & ",'" & terape & "'"
        StrSql = StrSql & "," & ConvFecha(terfecnac)
        StrSql = StrSql & "," & CInt(tersex)
        StrSql = StrSql & ",'" & teremail & "'"
        StrSql = StrSql & "," & nacionalnro
        StrSql = StrSql & "," & paisnro
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline "Inserto en la tabla de tercero"
        
        '--Obtengo el ternro--
        ternro = getLastIdentity(objConn, "tercero")
        Flog.Writeline "-----------------------------------------------"
        Flog.Writeline "Codigo de Tercero = " & ternro
        
        
        If ternro <> 0 Then
        
            On Error GoTo 0
            On Error Resume Next
            'si da error  no puedo seguir
            
            '--Inserto el Registro correspondiente en ter_tip--
            StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & ternro & ",14)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline "Inserto el tipo de tercero 14 en ter_tip"
        
            '--Inserto el Documento--
            If tidnro <> 0 Then
                If tidnro > 4 Then tidnro = 1 'Cable
                nrodoc = Replace(nrodoc, ".", "") 'elimino puntos y comas
                nrodoc = Replace(nrodoc, ",", "")
                StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                StrSql = StrSql & " VALUES(" & ternro & "," & tidnro & ",'" & nrodoc & "')"
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.Writeline "Error al insertar el documento"
                    Err.Clear
                Else
                    Flog.Writeline "Inserto el Documento"
                End If
            End If
        
            '--Inserto el Domicilio--
            StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
            StrSql = StrSql & " VALUES(1," & ternro & ",-1,2)"
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.Writeline "Error al insertar el Domicilio"
                Err.Clear
            Else
                Flog.Writeline "Inserto el Domicilio"
            End If
            
            '--Obtengo el numero de domicilio en la tabla--
            NroDom = getLastIdentity(objConn, "cabdom")
    
            '--Si mo tiene algun dato le agregamos unos ficticios--
            'If Trim(calle) = "" Then calle = Null
            'If Trim(nro) = "" Then nro = Null
            'If Trim(piso) = "" Then piso = Null
            'If Trim(oficdepto) = "" Then oficdepto = Null
            'If Trim(codigopostal) = "" Then codigopostal = Null
            If locnro = 0 Then locnro = 1 'no informada
            If provnro = CStr(0) Then provnro = "1" 'no informada
            If provnro = "" Then provnro = "1" 'no informada
            If paisnro = 0 Then paisnro = 1 'no informada
            Provincia = CInt(provnro)
            Err.Clear
            StrSql = " INSERT INTO detdom (domnro,calle,nro,piso,oficdepto,codigopostal,"
            StrSql = StrSql & "locnro,provnro,paisnro) "
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & NroDom
            StrSql = StrSql & ",'" & CStr(calle) & "'"
            StrSql = StrSql & ",'" & CStr(nro) & "'"
            StrSql = StrSql & ",'" & CStr(piso) & "'"
            StrSql = StrSql & ",'" & CStr(oficdepto) & "'"
            StrSql = StrSql & ",'" & CStr(codigopostal) & "'"
            StrSql = StrSql & "," & CInt(locnro)
            StrSql = StrSql & "," & Provincia
            StrSql = StrSql & "," & CInt(paisnro)
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.Writeline "Error al insertar el Domicilio2"
                Err.Clear
            Else
                Flog.Writeline "Inserto el Domicilio"
            End If
    
        
            '--Telefonos--
            For a = 0 To UBound(telnro) - 1
                If Trim(telnro(a)) <> "" Then
                    StrSql = " SELECT * from telefono where domnro = " & NroDom & " AND telnro = '" & telnro(a) & "'"
                    OpenRecordset StrSql, rs_sub
                    If rs_sub.EOF Then
                         StrSql = " INSERT INTO telefono "
                         StrSql = StrSql & " (domnro, telnro, telfax, teldefault, telcelular ) "
                         StrSql = StrSql & " VALUES (" & NroDom & ", '" & Left(telnro(a), 20) & "' ," & telfax(a) & "," & teldefault(a) & "," & telcelular(a) & " ) "
                         objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    If Err Then
                        Flog.Writeline "Error al insertar el Telefono "
                        Err.Clear
                    Else
                        Flog.Writeline " Inserto el telefono "
                    End If
                End If
            Next a
        
            '--Complemento--
            For a = 0 To UBound(posrempre) - 1 'entra solo una vez
                StrSql = " INSERT INTO pos_postulante "
                'FGZ - 16/04/2007 - Le agregué el estado, campo estposnro con default en 4
                StrSql = StrSql & " (posrempre, ternro, posfecpres, posref, procnro,estposnro) "
                StrSql = StrSql & " VALUES (" & posrempre(a) & ", " & ternro & " ," & ConvFecha(posfecpres(a)) & ",'" & posref(a) & "'," & TraerCodProcedencia("Bumeran") & ",4 ) "
                'StrSql = StrSql & " Go "
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.Writeline "Error al insertar el Complemento " & Err.Description
                    Flog.Writeline StrSql
                    Err.Clear
                Else
                    Flog.Writeline "Inserte el Complemento "
                End If
                a = UBound(posrempre) - 1 'entra solo una vez
            Next a
        
            '--Empleos Anteriores--57
            ActPasos = False
            For a = 0 To UBound(Empnro) - 1
                StrSql = " INSERT INTO empant "
                StrSql = StrSql & " ( empleado, empatareas, lempnro, empadesde, emmpahasta, carnro, empaini, empafin ) "
                StrSql = StrSql & " VALUES (" & ternro & ", '" & empatareas(a) & "' ," & Empnro(a) & "," & empadesde(a) & "," & empahasta(a) & "," & carnro(a) & "," & empadesde(a) & "," & empahasta(a) & " ) "
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.Writeline "Error al insertar el empleo anterior "
                    Err.Clear
                Else
                    Flog.Writeline "Inserte Empleo anterior "
                    ActPasos = True
                End If
            Next a
            If ActPasos Then
                Call InsertarPaso(ternro, 57)
            End If
            ActPasos = False
            
            '--Inserto los estudios formales--49
            For a = 0 To UBound(nivnro) - 1
                If (CInt(nivnro(a)) <> 0) Then
                    If UCase(capfechasta(a)) = "NULL" Then
                        estact = -1
                        carrcomp = 0
                    Else
                        estact = 0
                        carrcomp = -1
                    End If
                    StrSql = " SELECT * from cap_estformal where nivnro = " & nivnro(a) & " and ternro = " & ternro & " and instnro = " & instnro(a) & " and titnro = " & titnro(a)
                    OpenRecordset StrSql, rs_sub
                    If rs_sub.EOF Then
                        StrSql = " INSERT INTO cap_estformal "
                        StrSql = StrSql & " ( nivnro, ternro, capfecdes, capfechas, instnro, capprom, caprango, titnro, capcomp, capestact ) "
                        StrSql = StrSql & " VALUES (" & nivnro(a) & ", " & ternro & " ," & capfecdesde(a) & "," & capfechasta(a) & "," & instnro(a) & ",'" & capprom(a) & " ','" & caprango(a) & "'," & titnro(a) & ", " & carrcomp & ", " & estact & " ) "
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    If Err Then
                        Flog.Writeline "Error al insetar el estudio Formal" & nivnro(a)
                        Err.Clear
                    Else
                        Flog.Writeline "Inserte el estudio Formal " & nivnro(a)
                        ActPasos = True
                    End If
                End If
            Next a
            If ActPasos Then
                Call InsertarPaso(ternro, 49)
            End If
            ActPasos = False
        
            '--Idiomas--53
            For a = 0 To UBound(idinro) - 1
                If Not TieneIdioma(ternro, idinro(a)) Then
                    StrSql = " INSERT INTO emp_idi "
                    StrSql = StrSql & " (idinro, empleado, empidlee, empidhabla, empidescr) "
                    If idcalificador(a) = "16" Then
                        StrSql = StrSql & " VALUES (" & idinro(a) & ", " & ternro & " , NULL , NULL, " & idnivel(a) & " ) "
                    Else
                        StrSql = StrSql & " VALUES (" & idinro(a) & ", " & ternro & " , NULL , " & idnivel(a) & ", NULL ) "
                    End If
                Else
                    StrSql = " UPDATE emp_idi SET "
                    StrSql = StrSql & " idinro = " & idinro(a) & ", empleado = " & ternro
                    If idcalificador(a) = "16" Then
                        StrSql = StrSql & ", empidescr = " & idnivel(a)
                    Else
                        StrSql = StrSql & ", empidhabla = " & idnivel(a)
                    End If
                    StrSql = StrSql & " where empleado = " & ternro & " and idinro = " & idinro(a)
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.Writeline "Error al inserte el Idioma " & idinro(a)
                    Err.Clear
                Else
                    Flog.Writeline "Inserte el Idioma " & idinro(a)
                    ActPasos = True
                End If
            Next a
            If ActPasos Then
                Call InsertarPaso(ternro, 53)
            End If
            ActPasos = False
        
            
            '--Especialidades--51
            For a = 0 To UBound(espnro) - 1
                StrSql = " INSERT INTO especemp "
                StrSql = StrSql & " (eltananro, ternro, espnivnro, espmeses, espfecha) "
                StrSql = StrSql & " VALUES (" & espnro(a) & ", " & ternro & " ," & espnivnro(a) & ", NULL, NULL ) "
                'StrSql = StrSql & " GO "
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.Writeline "Error al inserte Especialidades " & espnro(a)
                    Err.Clear
                Else
                    Flog.Writeline "Inserte la especialidad " & espnro(a)
                    ActPasos = True
                End If
            Next a
            If ActPasos Then
                Call InsertarPaso(ternro, 51)
            End If
            ActPasos = False
            
            
            '    'Postgrados
            '    If tiene_postgrado Then
            '      StrSql = " INSERT INTO cap_estformal(ternro,nivnro,titnro,instnro,capcomp,capactual)"
            '      StrSql = StrSql & " VALUES ("
            '      StrSql = StrSql & NroTercero
            '      StrSql = StrSql & "," & Ne_Nro
            '      StrSql = StrSql & "," & PosTitulo_Nro
            '      StrSql = StrSql & "," & Institucion_Nro
            '      StrSql = StrSql & "," & CInt(Ne_Completo)
            '      StrSql = StrSql & "," & IIf(tiene_postgrado, 0, -1)
            '      StrSql = StrSql & ")"
            '      objConn.Execute StrSql, , adExecuteNoRecords
            '      Flog.Writeline "Inserte el Nivel de estudio - Postgrado"
            '    End If
            '
            '    'Empleos anteriores
            '    If Not IsNull(Cargo_Descripcion) Then
            '
            '    End If
            '
        End If
    
    End If
    rs_sub.Close
    
    If rs.State = adStateOpen Then rs.Close
    'If rs_sql.State = adStateOpen Then rs_sql.Close
    
    Err.Clear
    IniciarVariablesBumeran
    
    Exit Function

ErrorTercero:
    Flog.Writeline "error al insergar el tercero " & ternom & "," & terape
    Flog.Writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline "Error: " & Err.Number
    Flog.Writeline "Decripcion: " & Err.Description
    Flog.Writeline Error
    Flog.Writeline "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    IniciarVariablesBumeran
    Exit Function
  
End Function

Function ModificarPostulanteBumeran()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que se encarga de modificar un postulante.
' Autor      : JMH
' Fecha      : 19/04/2006
' Ultima Mod.: FGZ - 11/05/2007
' Descripcion: Le agregué el pasinro a tercero
' ---------------------------------------------------------------------------------------------

    Dim rs_sub As New ADODB.Recordset
    Dim rs_Aux As New ADODB.Recordset
    Dim a As Integer
    Dim ActPasos As Boolean
    Dim estact
    Dim carrcomp
    Dim Provincia As Integer
    
    l_sql = "  "
    l_sql = l_sql & ""
    
    Err.Clear
    On Error GoTo ErrorTercero
    
    '--Modifico el Tercero--
    StrSql = " UPDATE tercero SET "
    StrSql = StrSql & " ternom = '" & ternom & "'"
    StrSql = StrSql & ", terape = '" & terape & "'"
    StrSql = StrSql & ", terfecnac = " & ConvFecha(terfecnac)
    StrSql = StrSql & ", tersex = " & CInt(tersex)
    StrSql = StrSql & ", teremail = '" & teremail & "'"
    StrSql = StrSql & ", nacionalnro = " & nacionalnro
    'FGZ - 11/05/2007 - le agregué el paisnro
    StrSql = StrSql & ", paisnro = " & paisnro
    StrSql = StrSql & " where ternro = " & ternro
                
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Modifico en la tabla de tercero al tercero: & " & ternro
    
    If ternro <> 0 Then
    
        On Error GoTo 0
        On Error Resume Next
        'si da error  no puedo seguir
        
        '--Modifico el Documento--
        If tidnro <> 0 Then
            If tidnro > 4 Then tidnro = 1 'Cable
            'nrodoc = Replace(nrodoc, ".", "") 'elimino puntos y comas
            'nrodoc = Replace(nrodoc, ",", "")
            
            StrSql = " UPDATE ter_doc SET "
            StrSql = StrSql & " nrodoc = '" & nrodoc & "'"
            StrSql = StrSql & ", tidnro = " & CInt(tidnro)
            StrSql = StrSql & " WHERE ternro = " & ternro
                      
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.Writeline "Error al modificar el documento"
                Err.Clear
            Else
                Flog.Writeline "Modificoo el Documento"
            End If
        End If
    
        '--Modifico el Domicilio--
        Dim domnro As Long
        
        StrSql = " SELECT domnro "
        StrSql = StrSql & " FROM cabdom "
        StrSql = StrSql & " WHERE ternro = " & ternro & " AND tipnro = 1 "
        StrSql = StrSql & " AND domdefault = -1 AND tidonro = 2 "
        OpenRecordset StrSql, rs_sub
        If Not rs_sub.EOF Then
           domnro = rs_sub!domnro
           
           If locnro = 0 Then locnro = 1 'no informada
           If provnro = CStr(0) Then provnro = "1" 'no informada
           If provnro = "" Then provnro = "1" 'no informada
           If paisnro = 0 Then paisnro = 1 'no informada
           Provincia = CInt(provnro)
           Err.Clear
           
           StrSql = " UPDATE detdom SET "
           StrSql = StrSql & " calle = '" & CStr(calle) & "'"
           StrSql = StrSql & ", nro = '" & CStr(nro) & "'"
           StrSql = StrSql & ", piso = '" & CStr(piso) & "'"
           StrSql = StrSql & ", oficdepto = '" & CStr(oficdepto) & "'"
           StrSql = StrSql & ", codigopostal = '" & CStr(codigopostal) & "'"
           StrSql = StrSql & ", locnro = " & CInt(locnro)
           StrSql = StrSql & ", provnro = " & Provincia
           StrSql = StrSql & ", paisnro = " & CInt(paisnro)
           StrSql = StrSql & " WHERE domnro = " & domnro
        
           objConn.Execute StrSql, , adExecuteNoRecords
           If Err Then
              Flog.Writeline "Error al insertar el Domicilio"
              Err.Clear
           Else
              Flog.Writeline "Inserto el Domicilio"
           End If
        Else
            Flog.Writeline "Error al buscar la cabecera del domicilio "
            Err.Clear
        End If
        rs_sub.Close
        
        '--Telefonos--
        Dim HayTelefonos As Boolean
        
        HayTelefonos = False
        For a = 0 To UBound(telnro) - 1
        
            If a = 0 Then
               Flog.Writeline " Busco si se cargaron Telefonos para el Tercero: " & ternro
               If VienenTelefonos() = True Then
                  HayTelefonos = True
                  StrSql = " SELECT * from telefono where domnro = " & domnro
                  OpenRecordset StrSql, rs_sub
                  
                  'Borro los telefonos que tiene asociado ese Postulante
                  Do While Not rs_sub.EOF
                     StrSql = " DELETE FROM telefono "
                     StrSql = StrSql & " WHERE  domnro = " & domnro & " AND telnro = '" & rs_sub!telnro & "'"
                     objConn.Execute StrSql, , adExecuteNoRecords
                     If Err Then
                        Flog.Writeline "Error al Borrar el Teléfono "
                        Err.Clear
                     End If
                     rs_sub.MoveNext
                  Loop
                  rs_sub.Close
                  Flog.Writeline " Se borraron todos los Teléfonos "
               End If
            End If
            
            'Si la variable esta en TRUE entonces quiere decir
            'que en el XML se cargaron el Teléfonos para el Postulante
            If HayTelefonos = True Then
                StrSql = " INSERT INTO telefono "
                StrSql = StrSql & " (domnro, telnro, telfax, teldefault, telcelular ) "
                StrSql = StrSql & " VALUES (" & domnro & ", '" & Left(telnro(a), 20) & "' ," & telfax(a) & "," & teldefault(a) & "," & telcelular(a) & " ) "
                objConn.Execute StrSql, , adExecuteNoRecords
                
                If Err Then
                   Flog.Writeline "Error al insertar el Teléfono "
                   Err.Clear
                Else
                   Flog.Writeline " Inserto el Teléfono "
                End If
            End If
        Next a
    
        '--Complemento--
        For a = 0 To UBound(posrempre) - 1 'entra solo una vez
        
            StrSql = " UPDATE pos_postulante SET "
            StrSql = StrSql & " posrempre = " & posrempre(a)
            StrSql = StrSql & ", posfecpres = " & ConvFecha(posfecpres(a))
            StrSql = StrSql & ", posref = '" & posref(a) & "'"
            StrSql = StrSql & ", procnro = " & TraerCodProcedencia("Bumeran")
            'FGZ - 16/04/2007 - Le agregué el estado, campo estposnro con default en 4
            StrSql = StrSql & ", estposnro = 4"
            StrSql = StrSql & " WHERE ternro = " & ternro
           
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.Writeline "Error al insertar el Complemento " & Err.Description
                Flog.Writeline StrSql
                Err.Clear
            Else
                Flog.Writeline "Inserte el Complemento "
            End If
            a = UBound(posrempre) - 1 'entra solo una vez
            
        Next a
    
        '--Empleos Anteriores--57
        Dim HayEmpleos As Boolean
        
        HayEmpleos = False
        ActPasos = False
        For a = 0 To UBound(Empnro) - 1
        
            If a = 0 Then
               Flog.Writeline " Busco si se cargaron Empleos Anteriores para el Tercero: " & ternro
               If VienenEmpleosAnteriores() = True Then
                  HayEmpleos = True
                  StrSql = " SELECT empantnro FROM empant WHERE empleado = " & ternro
                  OpenRecordset StrSql, rs_sub
                  
                  'Borro los Empleos Anteriores que tiene asociado ese Postulante
                 Do While Not rs_sub.EOF
                     StrSql = " DELETE FROM empant "
                     StrSql = StrSql & " WHERE  empantnro = " & rs_sub!empantnro
                     objConn.Execute StrSql, , adExecuteNoRecords
                     If Err Then
                        Flog.Writeline "Error al Borrar el Empleo Anterior "
                        Err.Clear
                     End If
                     rs_sub.MoveNext
                  Loop
                  rs_sub.Close
                  Flog.Writeline " Se borraron todos los Empleos Anteriores "
               End If
            End If
            
            'Si la variable esta en TRUE entonces quiere decir
            'que en el XML se cargaron los Empleos Anteriores
            'para el Postulante y estos deben insertarse
            If HayEmpleos = True Then
                StrSql = " INSERT INTO empant "
                StrSql = StrSql & " ( empleado, empatareas, lempnro, empadesde, emmpahasta, carnro, empaini, empafin ) "
                StrSql = StrSql & " VALUES (" & ternro & ", '" & empatareas(a) & "' ," & Empnro(a) & "," & empadesde(a) & "," & empahasta(a) & "," & carnro(a) & "," & empadesde(a) & "," & empahasta(a) & " ) "
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.Writeline "Error al insertar el empleo anterior "
                    Err.Clear
                Else
                    Flog.Writeline "Inserte Empleo anterior "
                    ActPasos = True
                End If
            End If
        Next a
        If ActPasos Then
            Flog.Writeline "Actualizo el paso para los Empleos Anteriores. "
            Call EliminarPaso(ternro, 57)
            Call InsertarPaso(ternro, 57)
        End If
        ActPasos = False
        
        '--Inserto los estudios formales--49
        Dim HayEstudios As Boolean
        
        HayEstudios = False
        For a = 0 To UBound(nivnro) - 1
        
            If a = 0 Then
               Flog.Writeline " Busco si se cargaron Estudios Formales para el Tercero: " & ternro
               If VienenEstudiosFormales() = True Then
                  HayEstudios = True
                  StrSql = " SELECT instnro, titnro FROM cap_estformal WHERE nivnro = " & nivnro(a)
                  StrSql = StrSql & " AND ternro = " & ternro
                  StrSql = StrSql & " AND instnro = " & instnro(a)
                  StrSql = StrSql & " AND titnro = " & titnro(a)
                  OpenRecordset StrSql, rs_sub
                  
                  'Borro los Empleos Anteriores que tiene asociado ese Postulante
                  Do While Not rs_sub.EOF
                     StrSql = " DELETE FROM cap_estformal "
                     StrSql = StrSql & " WHERE  ternro = " & ternro
                     StrSql = StrSql & " AND  instnro = " & rs_sub!instnro
                     StrSql = StrSql & " AND  titnro = " & rs_sub!titnro
                     objConn.Execute StrSql, , adExecuteNoRecords
                     
                     If Err Then
                        Flog.Writeline "Error al Borrar el Estudio Formal "
                        Err.Clear
                     End If
                     rs_sub.MoveNext
                  Loop
                  rs_sub.Close
                  Flog.Writeline " Se borraron todos los Estudios Formales "
               End If
            End If
            
            If HayEstudios = True Then
                If (CInt(nivnro(a)) <> 0) Then
                    If UCase(capfechasta(a)) = "NULL" Then
                        estact = -1
                        carrcomp = 0
                    Else
                        estact = 0
                        carrcomp = -1
                    End If
                    
                    StrSql = " INSERT INTO cap_estformal "
                    StrSql = StrSql & " ( nivnro, ternro, capfecdes, capfechas, instnro, capprom, caprango, titnro, capcomp, capestact ) "
                    StrSql = StrSql & " VALUES (" & nivnro(a) & ", " & ternro & " ," & capfecdesde(a) & "," & capfechasta(a) & "," & instnro(a) & ",'" & capprom(a) & " ','" & caprango(a) & "'," & titnro(a) & ", " & carrcomp & ", " & estact & " ) "
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    If Err Then
                        Flog.Writeline "Error al insetar el estudio Formal" & nivnro(a)
                        Err.Clear
                    Else
                        Flog.Writeline "Inserte el estudio Formal " & nivnro(a)
                        ActPasos = True
                    End If
                End If
            End If
        Next a
        
        If ActPasos Then
            Flog.Writeline "Actualizo el paso para los Estudios Formales. "
            Call EliminarPaso(ternro, 49)
            Call InsertarPaso(ternro, 49)
        End If
        ActPasos = False
    
        '--Idiomas--53
        For a = 0 To UBound(idinro) - 1
            If Not TieneIdioma(ternro, idinro(a)) Then
                StrSql = " INSERT INTO emp_idi "
                StrSql = StrSql & " (idinro, empleado, empidlee, empidhabla, empidescr) "
                If idcalificador(a) = "16" Then
                    StrSql = StrSql & " VALUES (" & idinro(a) & ", " & ternro & " , NULL , NULL, " & idnivel(a) & " ) "
                Else
                    StrSql = StrSql & " VALUES (" & idinro(a) & ", " & ternro & " , NULL , " & idnivel(a) & ", NULL ) "
                End If
            Else
                StrSql = " UPDATE emp_idi SET "
                StrSql = StrSql & " idinro = " & idinro(a) & ", empleado = " & ternro
                If idcalificador(a) = "16" Then
                    StrSql = StrSql & ", empidescr = " & idnivel(a)
                Else
                    StrSql = StrSql & ", empidhabla = " & idnivel(a)
                End If
                StrSql = StrSql & " where empleado = " & ternro & " and idinro = " & idinro(a)
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.Writeline "Error al inserte el Idioma " & idinro(a)
                Err.Clear
            Else
                Flog.Writeline "Inserte el Idioma " & idinro(a)
                ActPasos = True
            End If
        Next a
        If ActPasos Then
            Flog.Writeline "Actualizo el paso para los Idiomas. "
            Call EliminarPaso(ternro, 53)
            Call InsertarPaso(ternro, 53)
        End If
        ActPasos = False
    
        
        '--Especialidades--51
        Dim HayEspecialidades As Boolean
        
        HayEspecialidades = False
        For a = 0 To UBound(espnro) - 1
        
            If a = 0 Then
               Flog.Writeline " Busco si se cargaron Especializaciones para el Tercero: " & ternro
               If VienenEspecializaciones() Then
                  HayEspecialidades = True
                  StrSql = " SELECT * FROM especemp WHERE ternro = " & ternro
                  OpenRecordset StrSql, rs_sub
                  
                  'Borro los Empleos Anteriores que tiene asociado ese Postulante
                  Do While Not rs_sub.EOF
                     StrSql = " DELETE FROM especemp "
                     StrSql = StrSql & " WHERE  ternro = " & ternro
                     StrSql = StrSql & " AND  eltananro = " & rs_sub!eltananro
                     StrSql = StrSql & " AND  espnivnro = " & rs_sub!espnivnro
                     objConn.Execute StrSql, , adExecuteNoRecords
                     
                     If Err Then
                        Flog.Writeline "Error al Borrar la Especialidad "
                        Err.Clear
                     End If
                     rs_sub.MoveNext
                  Loop
                  rs_sub.Close
                  Flog.Writeline " Se borraron todas las Especialidades "
               End If
            End If
            
            If HayEspecialidades = True Then
                StrSql = " INSERT INTO especemp "
                StrSql = StrSql & " (eltananro, ternro, espnivnro, espmeses, espfecha) "
                StrSql = StrSql & " VALUES (" & espnro(a) & ", " & ternro & " ," & espnivnro(a) & ", NULL, NULL ) "
                objConn.Execute StrSql, , adExecuteNoRecords
                
                If Err Then
                    Flog.Writeline "Error al inserte Especialidades " & espnro(a)
                    Err.Clear
                Else
                    Flog.Writeline "Inserte la especialidad " & espnro(a)
                    ActPasos = True
                End If
            End If
        Next a
        If ActPasos Then
            Flog.Writeline "Actualizo el paso para las Especialidades. "
            Call EliminarPaso(ternro, 51)
            Call InsertarPaso(ternro, 51)
        End If
        ActPasos = False
        
        
        '    'Postgrados
        '    If tiene_postgrado Then
        '      StrSql = " INSERT INTO cap_estformal(ternro,nivnro,titnro,instnro,capcomp,capactual)"
        '      StrSql = StrSql & " VALUES ("
        '      StrSql = StrSql & NroTercero
        '      StrSql = StrSql & "," & Ne_Nro
        '      StrSql = StrSql & "," & PosTitulo_Nro
        '      StrSql = StrSql & "," & Institucion_Nro
        '      StrSql = StrSql & "," & CInt(Ne_Completo)
        '      StrSql = StrSql & "," & IIf(tiene_postgrado, 0, -1)
        '      StrSql = StrSql & ")"
        '      objConn.Execute StrSql, , adExecuteNoRecords
        '      Flog.Writeline "Inserte el Nivel de estudio - Postgrado"
        '    End If
        '
        '    'Empleos anteriores
        '    If Not IsNull(Cargo_Descripcion) Then
        '
        '    End If
        '
    End If
    
    
    If rs.State = adStateOpen Then rs.Close
    'If rs_sql.State = adStateOpen Then rs_sql.Close
    
    'Err.Clear
    'IniciarVariablesBumeran
    
    Exit Function

ErrorTercero:
    Flog.Writeline "error al insergar el tercero " & ternom & "," & terape
    Flog.Writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline "Error: " & Err.Number
    Flog.Writeline "Decripcion: " & Err.Description
    Flog.Writeline Error
    Flog.Writeline "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    IniciarVariablesBumeran
    Exit Function
  
End Function

Function ArmarEspecializaciones(ByVal rs As ADODB.Recordset)
    Dim Col
    Dim Valores
    Dim Aux_Codigo As Long
    Dim a
        a = 0
    
    ReDim espnro(rs.RecordCount)
    ReDim espnivnro(rs.RecordCount)
    
    
    Aux_Codigo = CInt(TraerEspecializacion("Computación"))
    
    While rs.EOF <> True
        For Each Col In rs.Fields
            Valores = Col.Value
            If Col.Name <> "$Text" Then
                Select Case Col.Name
                    Case "idcalificador"
                    Case "idconocimiento"
                        'espnro(a) = CInt(TraerCodEltoana(CStr(l_Conocimientos(Valores)), "Computacion"))
                        espnro(a) = CInt(TraerCodEltoana(CStr(l_Conocimientos(Valores)), Aux_Codigo))
                        'espnro(a) = CInt(TraerCodEltoana(CStr(l_Conocimientos(Valores))))
                    Case "idnivel"
                        espnivnro(a) = CInt(TraerCodNivelEspecializacion(CStr(l_Niveles(Valores))))
                End Select
            End If
        Next
        rs.MoveNext
        a = a + 1
    Wend

                    
End Function

Function VienenEspecializaciones()
    Dim I As Integer
    Dim Salir As Boolean
    
    Salir = False
    I = 0
    Do While Salir = False And I <= UBound(espnro)
          If espnro(I) <> 0 Then
             Salir = True
          End If
          I = I + 1
    Loop
    
    VienenEspecializaciones = Salir
    
End Function


Function ArmarComplemento(ByVal rs As Recordset)
    Dim Col
    Dim Valores
    Dim a
        a = 0
    'ReDim Sql_Compl(rs.RecordCount)
    ReDim posfecpres(rs.RecordCount)
    ReDim posrempre(rs.RecordCount)
    ReDim posref(rs.RecordCount)
    While rs.EOF <> True
        For Each Col In rs.Fields
            If Col.Name <> "$Text" Then
                Valores = Col.Value
                Select Case Col.Name
                    Case "falta"
                        posfecpres(a) = CDate(Valores)
                    Case "frecuencia"
                    Case "minimo"
                        If Valores = "" Or IsNull(Valores) Then Valores = 0
                        posrempre(a) = CDbl(Valores)
                    Case "objetivos"    'desaparece
                    Case "pue_idpuesto" 'desaparece
                    Case "referencias"
                        posref(a) = Left(CStr(Valores), 250)
                    Case "tdt_idtipodetrabajo"
                End Select
            End If
        Next
        rs.MoveNext
        a = a + 1
    Wend
End Function
Function ArmarEstudiosFormales(ByVal rs As Recordset)
    Dim Col
    Dim Valores
    Dim a
        a = 0
    'ReDim Sql_EstFormal(rs.RecordCount)
    ReDim capfechasta(rs.RecordCount)
    ReDim capfecdesde(rs.RecordCount)
    ReDim instnro(rs.RecordCount)
    ReDim capprom(rs.RecordCount)
    ReDim caprango(rs.RecordCount)
    ReDim nivnro(rs.RecordCount)
    ReDim titnro(rs.RecordCount)
    While rs.EOF <> True
        For Each Col In rs.Fields
            If Col.Name <> "$Text" Then
                Valores = Col.Value
                Select Case Col.Name
                    Case "are_idareaestudio" '(desaparece)
                    Case "ffin"
                        If Valores = "" Then
                            capfechasta(a) = "null"
                        Else
                            capfechasta(a) = ConvFecha(Valores)
                        End If
                    Case "finicio"
                        If Valores = "" Then
                            capfecdesde(a) = "null"
                        Else
                            capfecdesde(a) = ConvFecha(Valores)
                        End If
                    Case "ins_idinstitucion"
                        If (Valores = 0 Or Valores = "") Then
                            instnro(a) = 0
                        Else
                            instnro(a) = CInt(TraerCodInstitucion(CStr(l_Instituciones(Valores, 0))))
                        End If
                    Case "institucion"
                        If instnro(a) = 0 Or CStr(instnro(a)) = "" Then ' si no hay una definida arriba, creo una
                            If Valores <> "" Then
                                instnro(a) = CInt(TraerCodInstitucion(CStr(Valores)))
                            Else
                                instnro(a) = ""
                            End If
                            'CInt(TraerCodInstitucion(CStr(l_Instituciones(Valores, 0))))
                        End If
                    Case "pai_idpais" '(desaparece)
                    Case "promedio"
                        capprom(a) = Left(CStr(Valores), 30)
                    Case "rng_idrango" 'cap_estformal
                        If Valores <> "" Then
                            caprango(a) = Left(CStr(l_Rango(CInt(Valores))), 60)
                        Else
                            caprango(a) = ""
                        End If
                    Case "tde_idtipodeestudio"
                        nivnro(a) = CInt(TraerCodNivelEstudio(CStr(l_Tipos_de_estudios(CInt(Valores)))))
                    Case "titulo"   'Descripto por el postulante.....
                        If Valores = "" Then
                            titnro(a) = 0
                        Else
                            titnro(a) = CInt(TraerCodTitulo(CStr(Valores), nivnro(a)))
                        End If
                End Select
            End If
        Next
        rs.MoveNext
        a = a + 1
    Wend
End Function

Function VienenEstudiosFormales()
    Dim I As Integer
    Dim Salir As Boolean
    
    Salir = False
    I = 0
    Do While Salir = False And I <= UBound(nivnro)
          If nivnro(I) <> 0 Then
             Salir = True
          End If
          I = I + 1
    Loop
    
    VienenEstudiosFormales = Salir
    
End Function


Function ArmarEmpleosAnteriores(ByVal rs As Recordset)
    Dim Col
    Dim Valores
    Dim a
        a = 0
    ReDim empatareas(rs.RecordCount)
    ReDim Empnro(rs.RecordCount)
    ReDim empahasta(rs.RecordCount)
    ReDim empadesde(rs.RecordCount)
    ReDim carnro(rs.RecordCount)
    'ReDim Sql_Empant(rs.RecordCount)
    While rs.EOF <> True
        For Each Col In rs.Fields
            If Col.Name <> "$Text" Then
                Valores = Col.Value
                Select Case Col.Name
                    Case "are_idarea" '(desaparece)
                    Case "descripcion"
                        empatareas(a) = Left(CStr(Valores), 200)
                        empatareas(a) = Replace(empatareas(a), vbCrLf, ". ")
                        empatareas(a) = Replace(empatareas(a), vbCr, ". ")
                    Case "empresa"
                        Empnro(a) = CInt(TraerCodListaEmpresa(CStr(Valores)))
                    Case "ffin"
                        If Valores = "" Then
                            empahasta(a) = "NULL"
                        Else
                            empahasta(a) = ConvFecha(Valores)
                        End If
                    Case "finicio"
                        If Valores = "" Then
                            empadesde(a) = "NULL"
                        Else
                            empadesde(a) = ConvFecha(Valores)
                        End If
                    Case "ind_idindustria" '(desaparece)
                    Case "pai_idpais"      '(desaparece)
                    Case "pue_idpuesto"
                    Case "puesto"
                        If Valores = "" Or IsNull(Valores) Then
                            carnro(a) = CInt(TraerCodCargo(CStr("Ninguno")))
                        Else
                            carnro(a) = CInt(TraerCodCargo(CStr(Valores)))
                        End If
                End Select
            End If
        Next
        rs.MoveNext
        a = a + 1
    Wend
    
End Function
Function VienenEmpleosAnteriores()
    Dim I As Integer
    Dim Salir As Boolean
    
    Salir = False
    I = 0
    Do While Salir = False And I <= UBound(Empnro)
          If Empnro(I) <> 0 Then
             Salir = True
          End If
          I = I + 1
    Loop
    
    VienenEmpleosAnteriores = Salir
    
End Function

Function ArmarIdiomas(ByVal rs As Recordset)
    Dim Col
    Dim Valores
    Dim a
        a = 0
    ReDim idinro(rs.RecordCount)
    ReDim idnivel(rs.RecordCount)
    Dim Calificador()
    ReDim Calificador(rs.RecordCount)
    ReDim idcalificador(rs.RecordCount)
    'Dim Arreglo(rs.RecordCount, 2)
    'ReDim Sql_Idioma(rs.RecordCount)
    While rs.EOF <> True
        For Each Col In rs.Fields
            If Col.Name <> "$Text" Then
                Valores = Col.Value
                Select Case Col.Name
                    Case "idcalificador"
                        ' 16 - escrito
                        ' 17 - oral
                        'Arreglo(rs.AbsolutePosition, 0) = Valores
                        idcalificador(a) = CInt(Valores)
                        Calificador(a) = Valores
                    Case "idconocimiento"
                        If Valores = 0 Or IsNull(Valores) Then Valores = 0
                        idinro(a) = CInt(TraerCodIdioma(CStr(l_Conocimientos(Valores))))
                        'Arreglo(rs.AbsolutePosition, 1) = idinro
                    Case "idnivel"
                        If Valores = "" Or IsNull(Valores) Then Valores = 0
                        idnivel(a) = CInt(TraerCodIdiNivel(CStr(l_Niveles(Valores))))
                        'Arreglo(rs.AbsolutePosition, 2) = idnivel
                End Select
            End If
        Next
        rs.MoveNext
        a = a + 1
    Wend
End Function

Function VienenIdiomas()
    Dim I As Integer
    Dim Salir As Boolean
    
    Salir = False
    I = 0
    Do While Salir = False And I <= UBound(idinro)
          If idinro(I) <> 0 Then
             Salir = True
          End If
          I = I + 1
    Loop
    
    VienenIdiomas = Salir
    
End Function


Function ArmarTelefonos(ByVal rs As Recordset)
    Dim Col
    Dim Valores
    Dim Categoria As Integer
    Dim a
        a = 0
    'ReDim Sql_Tel(rs.RecordCount)
    ReDim telfax(rs.RecordCount)
    ReDim teldefault(rs.RecordCount)
    ReDim telcelular(rs.RecordCount)
    ReDim telnro(rs.RecordCount)
    While rs.EOF <> True
        For Each Col In rs.Fields
            If Col.Name <> "$Text" Then
                Valores = Col.Value
                Select Case Col.Name
                    Case "categoria"
                        If Valores = 0 Or IsNull(Valores) Then Valores = 1
                        Categoria = CInt(Valores)
                        Select Case Categoria
                            Case 1  'telefono
                                telfax(a) = 0
                                teldefault(a) = -1
                                telcelular(a) = 0
                            Case 2  'alternativo
                                telfax(a) = 0
                                teldefault(a) = 0
                                telcelular(a) = 0
                            Case 3  'Fax
                                telfax(a) = -1
                                teldefault(a) = 0
                                telcelular(a) = 0
                            Case Else 'alternativo
                                telfax(a) = 0
                                teldefault(a) = 0
                                telcelular(a) = 0
                        End Select
                    Case "numero"
                        telnro(a) = telnro(a) & CStr(Valores)
                    Case "prefix"
                        If Trim(Valores) <> "" Or Not IsNull(Valores) Then
                            telnro(a) = CStr(Valores) & "-" & CStr(telnro(a))
                        End If
                End Select
            End If
        Next
        telnro(a) = validatelefono(telnro(a))
        rs.MoveNext
        a = a + 1
    Wend
End Function

Function VienenTelefonos()
    Dim I As Integer
    Dim Salir As Boolean
    
    Salir = False
    I = 0
    Do While Salir = False And I <= UBound(telnro)
          If telnro(I) <> "" Then
             Salir = True
          End If
          I = I + 1
    Loop
    
    VienenTelefonos = Salir
    
End Function


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
Function TieneIdioma(l_ternro As Long, l_idioma As Integer) As Boolean
    Dim rs_sub As New ADODB.Recordset
    StrSql = " SELECT empleado, idinro FROM emp_idi WHERE empleado = " & l_ternro & " and idinro = " & l_idioma
    OpenRecordset StrSql, rs_sub
    If rs_sub.EOF Then
        TieneIdioma = False
    Else
        TieneIdioma = True
    End If
End Function
Public Sub IniciarVariablesBumeran()
Dim I As Long

    ternro = 0
    terape = ""
    calle = ""
    'a_cambiares  (Ver q es?)
    locnro = 0
    codigopostal = ""
    oficdepto = ""
    teremail = ""
    'FGZ - 09/05/2007 - No se puede asignar el estado NULL a una variable
    'terfecnac = Null
    terfecnac = "00:00:00"
    'FGZ - 09/05/2007 - No se puede asignar el estado NULL a una variable
    
    'a_idusuario  (desaparece)
    ternom = ""
    nrodoc = ""
    nro = ""
    paisnro = 0
    nacionalnro = 0
    piso = ""
    provnro = 0
    'FGZ - 09/05/2007 - No se puede asignar el estado NULL a una variable
    'tersex = Null
    tersex = False
    'FGZ - 09/05/2007 - No se puede asignar el estado NULL a una variable
    tidnro = 0
    '- <computacion> (especializaciones eltoama y nivel)
    'idcalificador(ver q desaparece)
    ReDim espnro(0)
    ReDim espnivnro(0)
    '- <curriculum>
    ReDim posfecpres(0) ' = Null
    'frecuencia(de cobro)
    ReDim posrempre(0) ' = 0
    'objetivos (ver)
    'pue_idpuesto(ver)
    'puesto(ver)
    ReDim posref(0) ' = ""
    'tdt_idtipodetrabajo (Ver de agregar)
    '- <curriculum_area>
    'are_idarea (Ver)
    '- <curriculum_industria>
    'ind_idindustria (Ver)
    '- <estudio>
    'are_idareaestudio(area q desaparece)
    
    'FGZ - 09/05/2007 - No se puede asignar el estado NULL a una variable
    'capfechasta = Null
    'capfecdesde = Null
    ReDim capfechasta(0)
    ReDim capfecdesde(0)
    'FGZ - 09/05/2007 - No se puede asignar el estado NULL a una variable
    
    ReDim instnro(0) ' = 0
    ReDim institucion(0) ' = 0
    'pai_idpais (Desaparece, no tenemos la relacion con el pais)
    ReDim capprom(0) ' = ""
    ReDim caprango(0) ' = "" = ""
    ReDim nivnro(0) ' = "" = 0
    ReDim titulo(0) ' = "" = ""
    '- <experiencialaboral>
    'are_idarea(area q desaparece)
    ReDim empatareas(0) ' = ""
    ReDim Empnro(0) ' = 0
    
    'FGZ - 09/05/2007 - No se puede asignar el estado NULL a una variable
    'empadesde = Null
    'empahasta = Null
    ReDim empadesde(0)
    ReDim empahasta(0)
    'FGZ - 09/05/2007 - No se puede asignar el estado NULL a una variable
    
    'ind_idindustria (desaparece)
    'pai_idpais (desaparece)
    'pue_idpuesto
    ReDim carnro(0) ' = 0
    '- <idiomas>
    'idcalificador            'Desaparece
    ReDim idinro(0) ' = 0
    ReDim idnivel(0)  '= 0
    '- <telefono>
    ReDim Categoria(0)  '= 0
    ReDim telnro(0) ' = ""
    'prefix (desaparece)
    
End Sub
Public Sub Bumeran(titulo As String, Valor As String, hijo As Integer) ', Subtitulo As String)
' Descripcion: Interface de Postulantes de Bumeran
' Autor      : Lisandro Moro
' Fecha      : 26/08/2004
' Ultima Mod.:
    Select Case titulo
        Case "a_apellido"
            terape = Left(CStr(Valor), 25)
        Case "a_calle"
            calle = CStr(Valor)
        Case "a_cambiares"
        Case "a_ciudad"
            locnro = CInt(TraerCodLocalidad(Valor))
        Case "a_cp"
            codigopostal = CStr(Valor)
        Case "a_dpto"
            oficdepto = CStr(Valor)
        Case "a_email"
            teremail = CStr(Valor)
        Case "a_fnacimiento"
            terfecnac = CDate(Valor)
        Case "a_idusuario"
        Case "a_nombre"
            ternom = Left(CStr(Valor), 25)
        Case "a_nrodoc"
            nrodoc = CStr(Valor)
            If nrodoc = "" Then nrodoc = "0"
        Case "a_numero"
            nro = CStr(Valor)
        Case "a_pai_idpais"
            If Valor = "" Then
                paisnro = "NULL"
            Else
                paisnro = CInt(TraerCodPais(CStr(l_Paises(CInt(Valor)))))
            End If
        Case "a_pai_idpais_naciopais"
            If Valor = "" Then
               nacionalnro = "NULL"
            Else
               nacionalnro = CInt(TraerCodNacionalidad(CStr(l_Paises(CInt(Valor)))))
            End If
        Case "a_piso"
            piso = CStr(Valor)
        Case "a_pro_idprovincia_vivepro"
            If Valor = "" Then
                provnro = "NULL"
            Else
                provnro = CInt(TraerCodProvincia(CStr(l_Provincias(CInt(Valor), 0))))
            End If
        Case "a_sexo"
            tersex = CBool(Valor)
        Case "a_tdd_idtipodedocumento"
            If Valor = "" Or Valor = "0" Or IsNull(Valor) Then
                tidnro = 1 ' dni
            Else
                tidnro = CInt(TraerCodTipoDocumento(Replace(CStr(l_Tipos_de_documentos(CInt(Valor), 0)), ".", "")))
            End If
    'Computacion
'        Case "idcalificador"
'        Case "idconocimiento"
'            espnro = CInt(TraerCodEltoana(CStr(l_Conocimientos(valor)), "Computacion"))
'        Case "idnivel"
'            espnivnro = CInt(TraerCodNivelEspecializacion(CStr(l_Niveles(valor))))
    'Curriculum
'        Case "falta"
'            posfecpres = CDate(valor)
'        Case "frecuencia"
'        Case "minimo"
'            If valor = "" Or IsNull(valor) Then valor = 0
'            posrempre = CDbl(valor)
'        Case "objetivos"    'desaparece
'        Case "pue_idpuesto" 'desaparece
'        Case "referencias"
'            posref = CStr(valor)
'        Case "tdt_idtipodetrabajo"
    'curriculum_area
        Case "are_idarea" 'desaparece
    'curriculum_industria
        Case "ind_idindustria" 'desaparece
    'estudio
'        Case "are_idareaestudio" '(desaparece)
'        Case "ffin"
'            capfechasta = valor
'        Case "finicio"
'            capfecdesde = valor
'        Case "ins_idinstitucion"
'            If (CInt(valor) <> 0 Or valor <> "") Then
'                instnro = CInt(TraerCodInstitucion(CStr(l_Instituciones(valor, 0))))
'            Else
'                instnro = 0
'            End If
'        Case "institucion"
'            If instnro = 0 Then ' si no hay una definida arriba, creo una
'                institucion = CInt(TraerCodInstitucion(CStr(l_Instituciones(valor, 0))))
'            End If
'        Case "pai_idpais" '(desaparece)
'        Case "promedio"
'            capprom = CStr(valor)
'        Case "rng_idrango" 'cap_estformal
'            If valor <> "" Then
'                caprango = CStr(l_Rango(CInt(valor)))
'            Else
'                caprango = ""
'            End If
'        Case "tde_idtipodeestudio"
'            nivnro = CInt(TraerCodNivelEstudio(CStr(l_Tipos_de_estudios(CInt(valor)))))
'        Case "titulo"   'Descripto por el postulante.....
'            titulo = CStr(TraerCodTituloSolo(CStr(valor)))
    'experiencia laboral
'        Case "are_idarea" '(desaparece)
'        Case "descripcion"
'            empatareas = CStr(valor)
'        Case "empresa"
'            Empnro = CInt(TraerCodListaEmpresa(CStr(valor)))
'        Case "ffin"
'            empahasta = valor
'        Case "finicio"
'            empadesde = valor
'        Case "ind_idindustria" '(desaparece)
'        Case "pai_idpais"      '(desaparece)
'        Case "pue_idpuesto"
'        Case "puesto"
'            If valor = "" Or IsNull(valor) Then valor = 0
'            carnro = CInt(valor)
    'idiomas
'        Case "idcalificador"
'        Case "idconocimiento"
'            If valor = 0 Or IsNull(valor) Then valor = 0
'            idinro = CInt(TraerCodIdioma(CStr(l_Conocimientos(valor))))
'        Case "idnivel"
'            If valor = 0 Or IsNull(valor) Then valor = 0
'            idnivel = CInt(TraerCodIdiNivel(CInt(l_Niveles(valor))))
    'telefono
'        Case "categoria"
'            If valor = 0 Or IsNull(valor) Then valor = 0
'            Categoria = CInt(valor)
'        Case "numero"
'            telnro = telnro & CStr(valor)
'        Case "prefix"
'            telnro = CStr(valor) & telnro
    End Select
End Sub
Sub Postulantes()

'Dim rs As New ADODB.Recordset
'Dim rs_sql As New ADODB.Recordset
'
'  If rs.State = adStateOpen Then rs.Close
'  If rs_sql.State = adStateOpen Then rs_sql.Close
'
'  Set rs = Nothing
'  Set rs_sql = Nothing
End Sub
Function TraerNuevoCodigoPostulante()
'    Dim rs_sub As New ADODB.Recordset
'    StrSql = "INSERT INTO idinivel (idinivdesabr) "
'    StrSql = StrSql & " VALUES('" & idinivdesabr & "')"
'
'    objConn.Execute StrSql, , adExecuteNoRecords
'
'    StrSql = " SELECT MAX(idinivnro) AS Maxidinivnro FROM idinivel "
'    OpenRecordset StrSql, rs_sub
'
'    TraerCodIdiNivel = CInt(rs_sub!Maxidinivnro)
End Function
Public Sub EliminarPaso(terceros As Long, paso As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que se encarga de eliminar el paso para un dado postulante.
' Autor      : JMH
' Fecha      : 19/04/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    If Not EsNulo(terceros) Then
        StrSql = "DELETE FROM paso_ext WHERE pasnro =" & paso & " And extnro = " & terceros
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub
