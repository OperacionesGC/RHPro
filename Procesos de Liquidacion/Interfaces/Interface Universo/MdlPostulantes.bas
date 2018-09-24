Attribute VB_Name = "MdlPostulantes"
Option Explicit

Dim ternro As Long
Dim l_sql As String
Dim NroDom As Long
Dim idcalificador() As String
    'Dim id As Long
    Global ID As Long
Global ultimo_id As Long
    Dim terape As String          'a_apellido
    Dim calle As String           'a_calle
    'a_cambiares  (Ver q es?)
    Dim locnro As Long         'a_ciudad
    Dim codigopostal As String    'a_cp
    Dim oficdepto As String       'a_dpto
    Dim teremail As String        'a_email
    Dim terfecnac As Date         'a_fnacimiento
    'a_idusuario  (desaparece)
    Dim ternom As String          'a_nombre
    Dim nrodoc As String          'a_nrodoc
    Dim nro As String             'a_numero
    Dim paisnro As Long        'a_pai_idpais
    Dim nacionalnro As Long       'a_pai_idpais_naciopais
    Dim piso As String            'a_piso
    Dim provnro As String        'a_pro_idprovincia_vivepro
    Dim tersex As Boolean         'a_sexo
    Dim tidnro As Long         'a_tdd_idtipodedocumento
'- <computacion> (especializaciones eltoama y nivel)
    'idcalificador(ver q desaparece)
    Dim espnro() As Long        'idconocimiento
    Dim espnivnro() As Long      'idnivel
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
    Dim instnro() As Long        'inins_idinstitucion
    Dim institucion() As Long    'Institucion(Nueva, agregada a mano por el postulante)
    'pai_idpais (Desaparece, no tenemos la relacion con el pais)
    Dim capprom() As String         'promedio
    Dim caprango() As String        '(60)rng_idrango
    Dim nivnro() As Long         'tde_idtipodeestudio
    Dim titulo() As String          'titulo(Nueva, agregada a mano por el postulante)
    Dim titnro() As Long
'- <experiencialaboral>
    'are_idarea(area q desaparece)
    Dim empatareas() As String      'descripcion
    Dim Empnro() As Long         'empresa
    Dim empadesde() As String         'ffin
    Dim empahasta() As String         'finicio
    'ind_idindustria (desaparece)
    'pai_idpais (desaparece)
    'pue_idpuesto
    Dim carnro() As Long         'puesto
'- <idiomas>
    'Dim idcalificador
    Dim idinro() As Long         'idconocimiento
    Dim idnivel() As String         'idnivel(empidlee,empidhabla, empidescr)
'- <telefono>
    Dim Categoria() As Long      'Categoria(fax, default o celular)
    Dim telnro() As String          'prefix + Numero
    'prefix (desaparece)
    Dim telfax() As Long
    Dim teldefault() As Long
    Dim telcelular() As Long

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
    Case 278: 'Interface Postulantes Universo
        Call LeerXmlUniverso(Rec_postulantes, 0)
    End Select
'MyCommitTrans
End Sub

Public Sub Actualizar_Postulante_Segun_Modelo_Estandar(Rec_postulantes)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador de acurdo al modelo
' Autor      : FGZ
' Fecha      : 07/10/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Select Case NroModelo
    Case 278: 'Interface Postulantes Universo
        Call LeerXmlUniverso2(Rec_postulantes, 0)
    End Select
End Sub

Sub LeerXmlUniverso(rs, hijo)
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
                    Call Universo(Col.Name, Col.Value, CLng(hijo))
                Else
                    'Text2.Text = Text2.Text & vbCrLf
                    ' Retrieve the Child recordset
                    Columna = CStr(Col.Name)
                    Set rsChild = Col.Value
                    If Not rsChild.EOF Then rsChild.MoveFirst
                     If Err Then
                         Flog.writeline "Error: " & Error
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
                    ' LeerXmlUniverso rsChild, hijo + 1
                     
                     rsChild.Close
                     Set rsChild = Nothing
                End If
            Else
                'MsgBox "$Text", , Col.Name & "-" & Col.Value
            End If
        Next
   '     rs.MoveNext
        InsertarPostulanteUniverso
  '  Wend
    'rsChild.Close
    'rsChild = Nothing
End Sub

Sub LeerXmlUniverso2(rs, hijo)
    Dim Columna As String
    Dim rsChild         'ADODB.Recordset
    Dim Col             'ADODB.Field
    Dim rsChils As ADODB.Recordset
    Set rsChild = New ADODB.Recordset
    
    On Error Resume Next
   
        For Each Col In rs.Fields
            If Col.Name <> "$Text" Then   ' $Text to be ignored
                If Col.Type <> adChapter Then
                  ' Output the non-chaptered column
                    'MsgBox String((hijo), " ") & Col.Name & ": " & Col.Value & vbCrLf
                    Call Universo(Col.Name, Col.Value, CLng(hijo))
                End If
            End If
        Next
        ActualizarPostulanteUniverso
  End Sub


Function InsertarPostulanteUniverso()
    Dim rs_sub As New ADODB.Recordset
    Dim a As Long
    Dim ActPasos As Boolean
    Dim estact
    Dim carrcomp
    Dim Provincia As Long
    
    l_sql = "  "
    l_sql = l_sql & ""
    
    Err.Clear
    On Error GoTo ErrorTercero
    
    'Busco si ya existe el Postulante
    Flog.writeline
    Flog.writeline "Busco si el documento informado ya existe para algun postulante"
    
    nrodoc = Replace(nrodoc, ".", "")
    StrSql = " SELECT * FROM ter_doc "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = ter_doc.ternro "
    StrSql = StrSql & " INNER JOIN ter_tip ON ter_tip.ternro = tercero.ternro AND tipnro = 14 "
    StrSql = StrSql & " WHERE nrodoc = '" & nrodoc & "'"
    OpenRecordset StrSql, rs_sub
    If Not rs_sub.EOF Then
        Flog.writeline
        Flog.writeline "Hay un postulante con ese documento " & nrodoc
        ternro = rs_sub!ternro
        ModificarPostulanteUniverso
    Else
        Flog.writeline
        Flog.writeline "Postulante Nuevo "
        Flog.writeline
        
        '--Inserto el Tercero--
        Flog.writeline
        Flog.writeline "Tercero"
        
        StrSql = " INSERT INTO tercero (ternom,terape,terfecnac,tersex,teremail, nacionalnro) VALUES ("
        StrSql = StrSql & "'" & ternom & "'"
        StrSql = StrSql & ",'" & terape & "'"
        StrSql = StrSql & "," & ConvFecha(terfecnac)
        StrSql = StrSql & "," & CInt(tersex)
        StrSql = StrSql & ",'" & teremail & "'"
        StrSql = StrSql & "," & nacionalnro
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Inserto en la tabla de tercero"
        
        '--Obtengo el ternro--
        ternro = getLastIdentity(objConn, "tercero")
        Flog.writeline "-----------------------------------------------"
        Flog.writeline "Codigo de Tercero = " & ternro
        
        If ternro <> 0 Then
        
            On Error GoTo 0
            On Error Resume Next
            'si da error  no puedo seguir
            
            '--Inserto el Registro correspondiente en ter_tip--
            StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & ternro & ",14)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Inserto el tipo de tercero 14 en ter_tip"
        
            '--Inserto el Documento--
            If tidnro <> 0 Then
                If tidnro > 4 Then tidnro = 1 'Cable
                nrodoc = Replace(nrodoc, ".", "") 'elimino puntos y comas
                nrodoc = Replace(nrodoc, ",", "")
                StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                StrSql = StrSql & " VALUES(" & ternro & "," & tidnro & ",'" & nrodoc & "')"
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.writeline "Error al insertar el documento"
                    Err.Clear
                Else
                    Flog.writeline "Inserto el Documento"
                End If
            End If
        
            '--Inserto el Domicilio--
            Flog.writeline
            Flog.writeline "Domicilio"
            
            StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
            StrSql = StrSql & " VALUES(1," & ternro & ",-1,2)"
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline "Error al insertar el Domicilio"
                Err.Clear
            Else
                Flog.writeline "Inserto el Domicilio"
            End If
            
            '--Obtengo el numero de domicilio en la tabla--
            NroDom = getLastIdentity(objConn, "cabdom")
            Flog.writeline "    Cabecera de domicilio " & NroDom
        
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
            Provincia = CLng(provnro)
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
            StrSql = StrSql & "," & CLng(locnro)
            StrSql = StrSql & "," & Provincia
            StrSql = StrSql & "," & CLng(paisnro)
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline "Error al insertar el Domicilio"
                Err.Clear
            Else
                Flog.writeline "Domicilio Insertado"
            End If
    
        
            '--Telefonos--
            Flog.writeline
            Flog.writeline "Telefonos"
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
                        Flog.writeline "Error al insertar el Telefono " & telnro(a)
                        Err.Clear
                    Else
                        Flog.writeline " Inserto el telefono "
                    End If
                End If
            Next a
        
            '--Complemento--
            Flog.writeline
            Flog.writeline "Complemento"
            
            For a = 0 To UBound(posrempre) - 1 'entra solo una vez
                StrSql = " INSERT INTO pos_postulante "
                'FGZ - 16/04/2007 - Le agregué el estado, campo estposnro con default en 4
                StrSql = StrSql & " (posrempre, ternro, posfecpres, posref, procnro, estposnro, arepronro) "
                StrSql = StrSql & " VALUES (" & posrempre(a) & ", " & ternro & " ," & ConvFecha(posfecpres(a)) & ",'" & posref(a) & "'," & TraerCodProcedencia("Universo") & ",4, " & ID & " ) "
                'StrSql = StrSql & " Go "
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.writeline "Error al insertar el Complemento " & Err.Description
                    Flog.writeline StrSql
                    Err.Clear
                Else
                    Flog.writeline "Inserte el Complemento "
                End If
                a = UBound(posrempre) - 1 'entra solo una vez
            Next a
        
            '--Empleos Anteriores--58 ex 57
            Flog.writeline
            Flog.writeline "Empleaos Anteriores"
            
            ActPasos = False
            For a = 0 To UBound(Empnro) - 1
                'StrSql = " INSERT INTO empant "
                'StrSql = StrSql & " ( empleado, empatareas, lempnro, empadesde, emmpahasta, carnro, empaini, empafin ) "
                'StrSql = StrSql & " VALUES (" & ternro & ", '" & empatareas(a) & "' ," & Empnro(a) & "," & empadesde(a) & "," & empahasta(a) & "," & carnro(a) & "," & empadesde(a) & "," & empahasta(a) & " ) "
                
                'FGZ - 02/09/2009 - le cambié esto porque a veces no cargan la empresa y cuando eso sucede este insert rompe
                StrSql = " INSERT INTO empant "
                StrSql = StrSql & " ( empleado, empatareas "
                If Empnro(a) > 0 Then
                    StrSql = StrSql & ",lempnro"
                End If
                StrSql = StrSql & ", empadesde, emmpahasta, carnro, empaini, empafin ) "
                StrSql = StrSql & " VALUES (" & ternro & ", '" & empatareas(a) & "'"
                If Empnro(a) > 0 Then
                    StrSql = StrSql & "," & Empnro(a)
                End If
                StrSql = StrSql & "," & empadesde(a) & "," & empahasta(a) & "," & carnro(a) & "," & empadesde(a) & "," & empahasta(a) & " ) "
                
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.writeline "Error al insertar el empleo anterior "
                    Err.Clear
                Else
                    Flog.writeline "Inserte Empleo anterior "
                    ActPasos = True
                End If
            Next a
            If ActPasos Then
                Call InsertarPaso(ternro, 58)
            End If
            ActPasos = False
            
            '--Inserto los estudios formales--50 ex 49
            Flog.writeline
            Flog.writeline "Estudios Formales"
            
            For a = 0 To UBound(nivnro) - 1
                If (CLng(nivnro(a)) <> 0) Then
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
                        Flog.writeline "Error al insetar el estudio Formal" & nivnro(a)
                        Err.Clear
                    Else
                        Flog.writeline "Inserte el estudio Formal " & nivnro(a)
                        ActPasos = True
                    End If
                End If
            Next a
            If ActPasos Then
                Call InsertarPaso(ternro, 50)
            End If
            ActPasos = False
        
            '--Idiomas--54 ex 53
            Flog.writeline
            Flog.writeline "Idiomas"
            
            For a = 0 To UBound(idinro) - 1
                If Not TieneIdioma(ternro, idinro(a)) Then
                    StrSql = " INSERT INTO emp_idi "
                    StrSql = StrSql & " (idinro, empleado, empidlee, empidhabla, empidescr) "
                    StrSql = StrSql & " VALUES (" & idinro(a) & ", " & ternro & " , " & idnivel(a) & ", " & idnivel(a) & ", " & idnivel(a) & " ) "
                Else
                    StrSql = " UPDATE emp_idi SET "
                    StrSql = StrSql & " idinro = " & idinro(a) & ", empleado = " & ternro
                    StrSql = StrSql & ", empidescr = " & idnivel(a)
                    StrSql = StrSql & ", empidhabla = " & idnivel(a)
                    StrSql = StrSql & ", empidlee = " & idnivel(a)
                    StrSql = StrSql & " where empleado = " & ternro & " and idinro = " & idinro(a)
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.writeline "Error al inserte el Idioma " & idinro(a)
                    Err.Clear
                Else
                    Flog.writeline "Inserte el Idioma " & idinro(a)
                    ActPasos = True
                End If
            Next a
            If ActPasos Then
                Call InsertarPaso(ternro, 54)
            End If
            ActPasos = False
        
            
            '--Especialidades--52 ex 51
            Flog.writeline
            Flog.writeline "Especialidades"
            
            For a = 0 To UBound(espnro) - 1
                StrSql = " INSERT INTO especemp "
                StrSql = StrSql & " (eltananro, ternro, espnivnro, espmeses, espfecha) "
                StrSql = StrSql & " VALUES (" & espnro(a) & ", " & ternro & " ," & espnivnro(a) & ", NULL, NULL ) "
                'StrSql = StrSql & " GO "
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err Then
                    Flog.writeline "Error al inserte Especialidades " & espnro(a)
                    Err.Clear
                Else
                    Flog.writeline "Inserte la especialidad " & espnro(a)
                    ActPasos = True
                End If
            Next a
            If ActPasos Then
                Call InsertarPaso(ternro, 52)
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
        Else
            Flog.writeline
            Flog.writeline "No se pudo insertar el Tercero. Getidentity retornó 0."
        End If
    
    End If
    If rs_sub.State = adStateOpen Then rs_sub.Close
    
    If rs.State = adStateOpen Then rs.Close
    'If rs_sql.State = adStateOpen Then rs_sql.Close
    
    Err.Clear
    IniciarVariablesUniverso
    
    Exit Function

ErrorTercero:
    Flog.writeline "error al insergar el tercero " & ternom & "," & terape
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline Error
    Flog.writeline "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    IniciarVariablesUniverso
    Exit Function
  
End Function

Function ActualizarPostulanteUniverso()
    Dim rs_sub As New ADODB.Recordset
    Dim a As Long
    Dim ActPasos As Boolean
    Dim estact
    Dim carrcomp
    Dim Provincia As Long
    
    l_sql = "  "
    l_sql = l_sql & ""
    
    Err.Clear
    On Error GoTo ErrorTercero
    
    'Busco si ya existe el Postulante
    Flog.writeline
    Flog.writeline "Busco si el postulante ya existe "
    
    nrodoc = Replace(nrodoc, ".", "")
    StrSql = " SELECT * FROM ter_doc "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = ter_doc.ternro "
    StrSql = StrSql & " INNER JOIN ter_tip ON ter_tip.ternro = tercero.ternro AND tipnro = 14 "
    StrSql = StrSql & " WHERE nrodoc = '" & nrodoc & "'"
    OpenRecordset StrSql, rs_sub
    If Not rs_sub.EOF Then
       ternro = rs_sub!ternro
       ModificarPostulanteUniverso2
    Else
        Flog.writeline "El postulante no se encuentra. Problema con la numeracion del mismo "
    End If
    rs_sub.Close
    
    If rs.State = adStateOpen Then rs.Close
    'If rs_sql.State = adStateOpen Then rs_sql.Close
    
    Err.Clear
    IniciarVariablesUniverso
    
    Exit Function

ErrorTercero:
    Flog.writeline "error al insergar el tercero " & ternom & "," & terape
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline Error
    Flog.writeline "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    IniciarVariablesUniverso
    Exit Function
End Function

Function ModificarPostulanteUniverso()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que se encarga de modificar un postulante.
' Autor      : JMH
' Fecha      : 19/04/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    Dim rs_sub As New ADODB.Recordset
    Dim rs_Aux As New ADODB.Recordset
    Dim a As Long
    Dim ActPasos As Boolean
    Dim estact
    Dim carrcomp
    Dim Provincia As Long
    
    l_sql = "  "
    l_sql = l_sql & ""
    
    Err.Clear
    On Error GoTo ErrorTercero
    
    Flog.writeline
    Flog.writeline "El postulante ya existe, Actulizando... "
    
    '--Modifico el Tercero--
    Flog.writeline
    Flog.writeline "Tercero"
    
    StrSql = " UPDATE tercero SET "
    StrSql = StrSql & " ternom = '" & ternom & "'"
    StrSql = StrSql & ", terape = '" & terape & "'"
    StrSql = StrSql & ", terfecnac = " & ConvFecha(terfecnac)
    StrSql = StrSql & ", tersex = " & CInt(tersex)
    StrSql = StrSql & ", teremail = '" & teremail & "'"
    StrSql = StrSql & ", nacionalnro = " & nacionalnro
    StrSql = StrSql & " where ternro = " & ternro
                
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Modifico en la tabla de tercero al tercero: & " & ternro
    
    If ternro <> 0 Then
    
        On Error GoTo 0
        On Error Resume Next
        'si da error  no puedo seguir
        
        '--Modifico el Documento--
        Flog.writeline
        Flog.writeline "Documento"
        
        If tidnro <> 0 Then
            If tidnro > 4 Then tidnro = 1 'Cable
            'nrodoc = Replace(nrodoc, ".", "") 'elimino puntos y comas
            'nrodoc = Replace(nrodoc, ",", "")
            
            StrSql = " UPDATE ter_doc SET "
            StrSql = StrSql & " nrodoc = '" & nrodoc & "'"
            StrSql = StrSql & ", tidnro = " & CLng(tidnro)
            StrSql = StrSql & " WHERE ternro = " & ternro
                      
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline "Error al modificar el documento"
                Err.Clear
            Else
                Flog.writeline "Modificoo el Documento"
            End If
        End If
    
        '--Modifico el Domicilio--
        Flog.writeline
        Flog.writeline "Domicilio"
        
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
           Provincia = CLng(provnro)
           Err.Clear
           
           StrSql = " UPDATE detdom SET "
           StrSql = StrSql & " calle = '" & CStr(calle) & "'"
           StrSql = StrSql & ", nro = '" & CStr(nro) & "'"
           StrSql = StrSql & ", piso = '" & CStr(piso) & "'"
           StrSql = StrSql & ", oficdepto = '" & CStr(oficdepto) & "'"
           StrSql = StrSql & ", codigopostal = '" & CStr(codigopostal) & "'"
           StrSql = StrSql & ", locnro = " & CLng(locnro)
           StrSql = StrSql & ", provnro = " & Provincia
           StrSql = StrSql & ", paisnro = " & CLng(paisnro)
           StrSql = StrSql & " WHERE domnro = " & domnro
        
           objConn.Execute StrSql, , adExecuteNoRecords
           If Err Then
              Flog.writeline "Error al insertar el Domicilio"
              Err.Clear
           Else
              Flog.writeline "Inserto el Domicilio"
           End If
        Else
            Flog.writeline "Error al buscar la cabecera del domicilio "
            Err.Clear
        End If
        rs_sub.Close
        
        '--Telefonos--
        Flog.writeline
        Flog.writeline "Telefono"
        
        Dim HayTelefonos As Boolean
        
        HayTelefonos = False
        For a = 0 To UBound(telnro) - 1
        
            If a = 0 Then
               Flog.writeline " Busco si se cargaron Telefonos para el Tercero: " & ternro
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
                        Flog.writeline "Error al Borrar el Teléfono "
                        Err.Clear
                     End If
                     rs_sub.MoveNext
                  Loop
                  rs_sub.Close
                  Flog.writeline " Se borraron todos los Teléfonos "
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
                   Flog.writeline "Error al insertar el Teléfono "
                   Err.Clear
                Else
                   Flog.writeline " Inserto el Teléfono "
                End If
            End If
        Next a
    
        '--Complemento--
        Flog.writeline
        Flog.writeline "Complemento"
        
        For a = 0 To UBound(posrempre) - 1 'entra solo una vez
        
            StrSql = " UPDATE pos_postulante SET "
            StrSql = StrSql & " posrempre = " & posrempre(a)
            StrSql = StrSql & ", posfecpres = " & ConvFecha(posfecpres(a))
            StrSql = StrSql & ", posref = '" & posref(a) & "'"
            StrSql = StrSql & ", procnro = " & TraerCodProcedencia("Universo")
            'FGZ - 16/04/2007 - Le agregué el estado, campo estposnro con default en 4
            StrSql = StrSql & ", estposnro = 4"
            StrSql = StrSql & ", arepronro = " & ID
            StrSql = StrSql & " WHERE ternro = " & ternro
           
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline "Error al insertar el Complemento " & Err.Description
                Flog.writeline StrSql
                Err.Clear
            Else
                Flog.writeline "Inserte el Complemento con id " & ID
            End If
            a = UBound(posrempre) - 1 'entra solo una vez
            
        Next a
    
        '--Empleos Anteriores--58 ex 57
        Flog.writeline
        Flog.writeline "Empleos Anteriores"
        
        Dim HayEmpleos As Boolean
        
        HayEmpleos = False
        ActPasos = False
        For a = 0 To UBound(Empnro) - 1
        
            If a = 0 Then
               Flog.writeline " Busco si se cargaron Empleos Anteriores para el Tercero: " & ternro
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
                        Flog.writeline "Error al Borrar el Empleo Anterior "
                        Err.Clear
                     End If
                     rs_sub.MoveNext
                  Loop
                  rs_sub.Close
                  Flog.writeline " Se borraron todos los Empleos Anteriores "
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
                    Flog.writeline "Error al insertar el empleo anterior "
                    Err.Clear
                Else
                    Flog.writeline "Inserte Empleo anterior "
                    ActPasos = True
                End If
            End If
        Next a
        If ActPasos Then
            Flog.writeline "Actualizo el paso para los Empleos Anteriores. "
            Call EliminarPaso(ternro, 58)
            Call InsertarPaso(ternro, 58)
        End If
        ActPasos = False
        
        '--Inserto los estudios formales--50 ex 49
        Flog.writeline
        Flog.writeline "Estudios Formales"
        
        Dim HayEstudios As Boolean
        
        HayEstudios = False
        For a = 0 To UBound(nivnro) - 1
        
            If a = 0 Then
               Flog.writeline " Busco si se cargaron Estudios Formales para el Tercero: " & ternro
               If VienenEstudiosFormales() = True Then
                  HayEstudios = True
                  StrSql = " SELECT * FROM cap_estformal WHERE ternro = " & ternro
                  OpenRecordset StrSql, rs_sub
                  
                  'Borro los Empleos Anteriores que tiene asociado ese Postulante
                  Do While Not rs_sub.EOF
                     StrSql = " DELETE FROM cap_estformal "
                     StrSql = StrSql & " WHERE ternro = " & ternro
                     StrSql = StrSql & " AND instnro = " & rs_sub!instnro
                     StrSql = StrSql & " AND titnro = " & rs_sub!titnro
                     StrSql = StrSql & " AND nivnro = " & rs_sub!nivnro
                     objConn.Execute StrSql, , adExecuteNoRecords
                     
                     If Err Then
                        Flog.writeline "Error al Borrar el Estudio Formal "
                        Err.Clear
                     End If
                     rs_sub.MoveNext
                  Loop
                  rs_sub.Close
                  Flog.writeline " Se borraron todos los Estudios Formales "
               End If
            End If
            
            If HayEstudios = True Then
                If (CLng(nivnro(a)) <> 0) Then
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
                        Flog.writeline "Error al insetar el estudio Formal" & nivnro(a)
                        Err.Clear
                    Else
                        Flog.writeline "Inserte el estudio Formal " & nivnro(a)
                        ActPasos = True
                    End If
                End If
            End If
        Next a
        
        If ActPasos Then
            Flog.writeline "Actualizo el paso para los Estudios Formales. "
            Call EliminarPaso(ternro, 50)
            Call InsertarPaso(ternro, 50)
        End If
        ActPasos = False
    
        '--Idiomas--54 ex 53
        Flog.writeline
        Flog.writeline "Idiomas"
        
        For a = 0 To UBound(idinro) - 1
            If Not TieneIdioma(ternro, idinro(a)) Then
                StrSql = " INSERT INTO emp_idi "
                StrSql = StrSql & " (idinro, empleado, empidlee, empidhabla, empidescr) "
                StrSql = StrSql & " VALUES (" & idinro(a) & ", " & ternro & " , " & idnivel(a) & ", " & idnivel(a) & ", " & idnivel(a) & " ) "
            Else
                StrSql = " UPDATE emp_idi SET "
                StrSql = StrSql & " idinro = " & idinro(a) & ", empleado = " & ternro
                StrSql = StrSql & ", empidescr = " & idnivel(a)
                StrSql = StrSql & ", empidhabla = " & idnivel(a)
                StrSql = StrSql & ", empidlee = " & idnivel(a)
                StrSql = StrSql & " where empleado = " & ternro & " and idinro = " & idinro(a)
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            If Err Then
                Flog.writeline "Error al inserte el Idioma " & idinro(a)
                Err.Clear
            Else
                Flog.writeline "Inserte el Idioma " & idinro(a)
                ActPasos = True
            End If
        Next a
        If ActPasos Then
            Flog.writeline "Actualizo el paso para los Idiomas. "
            Call EliminarPaso(ternro, 54)
            Call InsertarPaso(ternro, 54)
        End If
        ActPasos = False
    
        
        '--Especialidades--52 ex 51
        Flog.writeline
        Flog.writeline "Especialidades"
        
        Dim HayEspecialidades As Boolean
        
        HayEspecialidades = False
        For a = 0 To UBound(espnro) - 1
        
            If a = 0 Then
               Flog.writeline " Busco si se cargaron Especializaciones para el Tercero: " & ternro
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
                        Flog.writeline "Error al Borrar la Especialidad "
                        Err.Clear
                     End If
                     rs_sub.MoveNext
                  Loop
                  rs_sub.Close
                  Flog.writeline " Se borraron todas las Especialidades "
               End If
            End If
            
            If HayEspecialidades = True Then
                StrSql = " INSERT INTO especemp "
                StrSql = StrSql & " (eltananro, ternro, espnivnro, espmeses, espfecha) "
                StrSql = StrSql & " VALUES (" & espnro(a) & ", " & ternro & " ," & espnivnro(a) & ", NULL, NULL ) "
                objConn.Execute StrSql, , adExecuteNoRecords
                
                If Err Then
                    Flog.writeline "Error al inserte Especialidades " & espnro(a)
                    Err.Clear
                Else
                    Flog.writeline "Inserte la especialidad " & espnro(a)
                    ActPasos = True
                End If
            End If
        Next a
        If ActPasos Then
            Flog.writeline "Actualizo el paso para las Especialidades. "
            Call EliminarPaso(ternro, 52)
            Call InsertarPaso(ternro, 52)
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
    IniciarVariablesUniverso
    
    Exit Function

ErrorTercero:
    Flog.writeline "error al insergar el tercero " & ternom & "," & terape
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline Error
    Flog.writeline "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    IniciarVariablesUniverso
    Exit Function
  
End Function


Function ModificarPostulanteUniverso2()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que se encarga de modificar solo el id de un postulante ya existente.
' Autor      : FGZ
' Fecha      : 07/10/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    Dim rs_sub As New ADODB.Recordset
    Dim rs_Aux As New ADODB.Recordset
    Dim a As Long
    Dim ActPasos As Boolean
    Dim estact
    Dim carrcomp
    Dim Provincia As Long
    
    l_sql = "  "
    l_sql = l_sql & ""
    
    Err.Clear
    On Error GoTo ErrorTercero
    
    Flog.writeline
    Flog.writeline "El postulante ya existe, Actulizando... "
    If ternro <> 0 Then
        StrSql = " UPDATE pos_postulante SET "
        StrSql = StrSql & " arepronro = " & ID
        StrSql = StrSql & " WHERE ternro = " & ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        If Err Then
            Flog.writeline "Error al insertar el Complemento " & Err.Description
            Flog.writeline StrSql
            Err.Clear
        Else
            Flog.writeline "Actualizo el Complemento con id " & ID
        End If
    End If
    If rs.State = adStateOpen Then rs.Close
    
    IniciarVariablesUniverso
    
    Exit Function

ErrorTercero:
    Flog.writeline "error al insergar el tercero " & ternom & "," & terape
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline Error
    Flog.writeline "Linea " & RegLeidos & " del archivo procesado"
    If rs.State = adStateOpen Then rs.Close
    IniciarVariablesUniverso
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
    
    
    'Aux_Codigo = CInt(TraerEspecializacion("Computación"))
    
    While rs.EOF <> True
        For Each Col In rs.Fields
            Valores = Col.Value
            If Col.Name <> "$Text" Then
                Select Case Col.Name
                    Case "idcalificador"
                        Aux_Codigo = CLng(TraerEspecializacion(CStr(Valores)))
                    Case "idconocimiento"
                        'espnro(a) = CInt(TraerCodEltoana(CStr(l_Conocimientos(Valores)), "Computacion"))
                        espnro(a) = CLng(TraerCodEltoana(CStr(Valores), Aux_Codigo))
                        'espnro(a) = CInt(TraerCodEltoana(CStr(l_Conocimientos(Valores))))
                    Case "idnivel"
                        espnivnro(a) = CLng(TraerCodNivelEspecializacion(CStr(Valores)))
                End Select
            End If
        Next
        rs.MoveNext
        a = a + 1
    Wend

                    
End Function

Function VienenEspecializaciones()
    Dim i As Long
    Dim Salir As Boolean
    
    Salir = False
    i = 0
    Do While Salir = False And i <= UBound(espnro)
          If espnro(i) <> 0 Then
             Salir = True
          End If
          i = i + 1
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
                            instnro(a) = CLng(TraerCodInstitucion(CStr(Valores)))
                        End If
                    Case "institucion"
                        If instnro(a) = 0 Or CStr(instnro(a)) = "" Then ' si no hay una definida arriba, creo una
                            If Valores <> "" Then
                                instnro(a) = CLng(TraerCodInstitucion(CStr(Valores)))
                            Else
                                instnro(a) = 0
                            End If
                            'CInt(TraerCodInstitucion(CStr(l_Instituciones(Valores, 0))))
                        End If
                    Case "pai_idpais" '(desaparece)
                    Case "promedio"
                        capprom(a) = Left(CStr(Valores), 30)
                    Case "rng_idrango" 'cap_estformal
                        If Valores <> "" Then
                            caprango(a) = Left(CStr(Valores), 60)
                        Else
                            caprango(a) = ""
                        End If
                    Case "tde_idtipodeestudio"
                        nivnro(a) = CLng(TraerCodNivelEstudio(CStr(Valores)))
                    Case "titulo"   'Descripto por el postulante.....
                        If Valores = "" Then
                            titnro(a) = 0
                        Else
                            titnro(a) = CLng(TraerCodTitulo(CStr(Valores), nivnro(a)))
                        End If
                End Select
            End If
        Next
        rs.MoveNext
        a = a + 1
    Wend
End Function

Function VienenEstudiosFormales()
    Dim i As Long
    Dim Salir As Boolean
    
    Salir = False
    i = 0
    Do While Salir = False And i <= UBound(nivnro)
          If nivnro(i) <> 0 Then
             Salir = True
          End If
          i = i + 1
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
                        Empnro(a) = CLng(TraerCodListaEmpresa(CStr(Valores)))
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
                            carnro(a) = CLng(TraerCodCargo(CStr("Ninguno")))
                        Else
                            carnro(a) = CLng(TraerCodCargo(CStr(Valores)))
                        End If
                End Select
            End If
        Next
        rs.MoveNext
        a = a + 1
    Wend
    
End Function
Function VienenEmpleosAnteriores()
    Dim i As Long
    Dim Salir As Boolean
    
    Salir = False
    i = 0
    Do While Salir = False And i <= UBound(Empnro)
          If Empnro(i) <> 0 Then
             Salir = True
          End If
          i = i + 1
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
                        idcalificador(a) = CStr(Valores)
                        Calificador(a) = Valores
                    Case "idconocimiento"
                        If Valores = 0 Or IsNull(Valores) Then Valores = "Ninguno"
                        idinro(a) = CLng(TraerCodIdioma(CStr(Valores)))
                        'Arreglo(rs.AbsolutePosition, 1) = idinro
                    Case "idnivel"
                        Select Case Valores
                            Case "B"
                                Valores = "Básico"
                            Case "I"
                                Valores = "Intermedio"
                            Case "A"
                                Valores = "Avanzado"
                            Case "N"
                                Valores = "Nativo"
                            Case Else
                                Valores = "Ninguno"
                        End Select
                        idnivel(a) = CLng(TraerCodIdiNivel(CStr(Valores)))
                        'Arreglo(rs.AbsolutePosition, 2) = idnivel
                End Select
            End If
        Next
        rs.MoveNext
        a = a + 1
    Wend
End Function

Function VienenIdiomas()
    Dim i As Long
    Dim Salir As Boolean
    
    Salir = False
    i = 0
    Do While Salir = False And i <= UBound(idinro)
          If idinro(i) <> 0 Then
             Salir = True
          End If
          i = i + 1
    Loop
    
    VienenIdiomas = Salir
    
End Function


Function ArmarTelefonos(ByVal rs As Recordset)
    Dim Col
    Dim Valores
    Dim Categoria As String
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
                        If Valores = "" Or IsNull(Valores) Then Valores = "P"
                        Categoria = CStr(Valores)
                        Select Case Categoria
                            Case "P"  'telefono
                                telfax(a) = 0
                                teldefault(a) = -1
                                telcelular(a) = 0
                            Case "C"  'Celular
                                telfax(a) = 0
                                teldefault(a) = 0
                                telcelular(a) = -1
                            Case "M"  'Mensajes
                                telfax(a) = 0
                                teldefault(a) = 0
                                telcelular(a) = 0
                            Case Else 'alternativo
                                telfax(a) = -1
                                teldefault(a) = 0
                                telcelular(a) = 0
                        End Select
                    Case "numero"
                        telnro(a) = telnro(a) & CStr(Valores)
                    Case "prefix"
                        If Not IsNull(Valores) Then
                            If Trim(Valores) <> "" Then
                                telnro(a) = CStr(Valores) & "-" & CStr(telnro(a))
                            End If
                        End If
                End Select
            End If
        Next
        telnro(a) = validatelefono(telnro(a))
        ' Voy eliminando la menor cantidad de caracteres posibles
        If Len(telnro(a)) > 20 Then telnro(a) = Replace(telnro(a), " ", "")
        If Len(telnro(a)) > 20 Then telnro(a) = Replace(telnro(a), "(", "")
        If Len(telnro(a)) > 20 Then telnro(a) = Replace(telnro(a), ")", "")
        If Len(telnro(a)) > 20 Then telnro(a) = Replace(telnro(a), "-", "")
        If Len(telnro(a)) > 20 Then telnro(a) = Right(telnro(a), 20)
        rs.MoveNext
        a = a + 1
    Wend
End Function

Function VienenTelefonos()
    Dim i As Long
    Dim Salir As Boolean
    
    Salir = False
    i = 0
    Do While Salir = False And i <= UBound(telnro)
          If telnro(i) <> "" Then
             Salir = True
          End If
          i = i + 1
    Loop
    
    VienenTelefonos = Salir
    
End Function


Function validatelefono(cadena As String) As String
    Dim a As Long
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
Public Sub IniciarVariablesUniverso()
    ternro = 0
    terape = ""
    calle = ""
    'a_cambiares  (Ver q es?)
    locnro = 0
    codigopostal = ""
    oficdepto = ""
    teremail = ""
    terfecnac = Date
    'a_idusuario  (desaparece)
    ternom = ""
    nrodoc = ""
    nro = ""
    paisnro = 0
    nacionalnro = 0
    piso = ""
    provnro = 0
    'tersex = Null
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
    ReDim capfechasta(0)
    ReDim capfecdesde(0)
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
    ReDim empadesde(0)
    ReDim empahasta(0)
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
Public Sub Universo(titulo As String, Valor As String, hijo As Long) ', Subtitulo As String)
' Descripcion: Interface de Postulantes de Universo
' Autor      : Lisandro Moro
' Fecha      : 26/08/2004
' Ultima Mod.:
    Select Case titulo
        'FGZ - 07/10/2008
        'Case "id"
        Case "id", "usuario id"
            ID = CLng(Valor)
            Flog.writeline "Leyendo id: " & ID
        Case "a_apellido"
            terape = Left(CStr(Valor), 25)
        Case "a_calle"
            calle = CStr(Valor)
        Case "a_cambiares"
        Case "a_ciudad"
            locnro = CLng(TraerCodLocalidad(Valor))
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
        Case "a_idpais" 'NUEVO
            If Valor = "" Then
                paisnro = "NULL"
            Else
                paisnro = CLng(TraerCodPais(CStr((Valor))))
            End If
        Case "a_pai_idpais"
            If Valor = "" Then
                paisnro = "NULL"
            Else
                paisnro = CLng(TraerCodPais(CStr(Valor)))
            End If
        Case "a_pai_idpais_naciopais"
            If Valor = "" Then
               nacionalnro = "NULL"
            Else
               nacionalnro = CLng(TraerCodNacionalidad(CStr(Valor)))
            End If
        Case "a_piso"
            piso = CStr(Valor)
        Case "a_pro_idprovincia_vivepro"
            If Valor = "" Then
                provnro = "NULL"
            Else
                provnro = CLng(TraerCodProvincia(CStr(Valor)))
            End If
        Case "a_sexo"
            If Valor = "M" Then
                tersex = True
            Else
                tersex = False
            End If
        Case "a_tdd_idtipodedocumento"
            If Valor = "" Or Valor = "0" Or IsNull(Valor) Then
                tidnro = 1 ' dni
            Else
                tidnro = CLng(TraerCodTipoDocumento(Replace(CStr(Valor), ".", "")))
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
