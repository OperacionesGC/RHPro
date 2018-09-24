Attribute VB_Name = "mdlExportacion"
Option Explicit



'Exportación del modelo 1000
Function expModelo1000(ByVal ternro As Long, ByVal separador As String)
 Dim rsDatosEmp  As New ADODB.Recordset
 Dim rsAux  As New ADODB.Recordset
 Dim strLinea As String

    On Error GoTo CE

    StrSql = "SELECT empleg,tercero.ternom,tercero.ternom2,tercero.terape,tercero.terape2,tercero.terfecnac,pais.paisdesc,nacionalidad.nacionaldes " & _
            ",tercero.terfecing,estcivil.estcivdesabr,tercero.tersex,empleado.empfecalta,empleado.empestudia,nivest.nivdesc,empleado.empemail" & _
            ",empleado.empfbajaprev,empleado.empest,empleado.empremu,adptemplate.tplatedesabr,empleado.empreporta" & _
            " FROM empleado " & _
            " INNER JOIN tercero ON tercero.ternro = empleado.ternro  " & _
            " INNER JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro " & _
            " INNER JOIN pais ON pais.paisnro = tercero.paisnro " & _
            " INNER JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro " & _
            " LEFT JOIN nivest ON nivest.nivnro = empleado.nivnro " & _
            " INNER JOIN adptemplate ON adptemplate.tplatenro = empleado.tplatenro " & _
            " WHERE empleado.ternro = " & ternro
    OpenRecordset StrSql, rsDatosEmp

    strLinea = "1000" & separador & rsDatosEmp!empleg
    strLinea = strLinea & separador & rsDatosEmp!terape & "@" & rsDatosEmp!terape2
    strLinea = strLinea & separador & rsDatosEmp!ternom & "@" & rsDatosEmp!ternom2
    strLinea = strLinea & separador & rsDatosEmp!terfecnac
    strLinea = strLinea & separador & rsDatosEmp!paisdesc
    strLinea = strLinea & separador & rsDatosEmp!nacionaldes
    
    If EsNulo(rsDatosEmp!terfecing) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!terfecing
    End If
    
    strLinea = strLinea & separador & rsDatosEmp!estcivdesabr
    
    If (CLng(rsDatosEmp!tersex) = -1) Then
        strLinea = strLinea & separador & "M"
    Else
        strLinea = strLinea & separador & "F"
    End If
        

    If EsNulo(rsDatosEmp!empfecalta) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!empfecalta
    End If
    
    
    If (CLng(rsDatosEmp!empestudia) = 0) Then
        strLinea = strLinea & separador & "No"
    Else
        strLinea = strLinea & separador & "Si"
    End If
    
    
    If EsNulo(rsDatosEmp!nivdesc) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!nivdesc
    End If
    
    strLinea = strLinea & separador & rsDatosEmp!empemail
    
    If EsNulo(rsDatosEmp!empfbajaprev) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!empfbajaprev
    End If
    
    If rsDatosEmp!empest = 0 Then
        strLinea = strLinea & separador & "Inactivo"
    Else
        strLinea = strLinea & separador & "Activo"
    End If
    
    strLinea = strLinea & separador & rsDatosEmp!empremu
    strLinea = strLinea & separador & rsDatosEmp!tplatedesabr
    
    StrSql = "SELECT empleg FROM empleado WHERE empleado.ternro= " & IIf(EsNulo(rsDatosEmp!empreporta), 0, rsDatosEmp!empreporta)
    OpenRecordset StrSql, rsAux
    If rsAux.EOF Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsAux!empleg
    End If
    
    
    GoTo datosOk
CE:
    strLinea = ""
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 1000. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Function
datosOk:
    Call sincronizar_Det(ternro, 1000)
    expModelo1000 = strLinea
    Set rsDatosEmp = Nothing
    
End Function



'Exportación del modelo 1001 - Domicilio
Function expModelo1001(ByVal ternro As Long, ByVal separador As String)
 Dim rsDatosEmp  As New ADODB.Recordset
 Dim rsAux  As New ADODB.Recordset
 Dim strLinea As String

    On Error GoTo CE

    StrSql = "SELECT empleg,cabdom.domnro,tipodomi.tidodes,detdom.calle,detdom.nro,detdom.piso,detdom.oficdepto,detdom.torre,detdom.manzana " & _
            " ,detdom.codigopostal,detdom.EntreCalles,detdom.Barrio,localidad.locdesc,partido.partnom partido,zona.zonadesc,provincia.provdesc " & _
            " ,pais.paisdesc" & _
            " FROM empleado " & _
            " INNER JOIN cabdom ON cabdom.ternro = empleado.ternro " & _
            " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
            " INNER JOIN tipodomi ON tipodomi.tidonro = cabdom.tidonro " & _
            " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
            " LEFT JOIN partido ON partido.partnro = detdom.partnro  " & _
            " LEFT JOIN zona ON zona.zonanro = localidad.zonanro " & _
            " INNER JOIN provincia ON provincia.provnro = localidad.provnro " & _
            " INNER JOIN pais ON pais.paisnro = provincia.paisnro " & _
            " WHERE empleado.ternro = " & ternro
    OpenRecordset StrSql, rsDatosEmp
    
    'Modelo (Fijo) | Legajo | Pais (Fijo ARGENTINA)
    strLinea = "1001" & separador & rsDatosEmp!empleg & separador & "ARGENTINA "
        
    'Descripcion del tipo de domicilio(tipodomi)
    strLinea = strLinea & separador & rsDatosEmp!tidodes
    
    'Nombre de la calle (detdom.calle)
    strLinea = strLinea & separador & rsDatosEmp!calle
    
    'Número de la calle (detdom.nro)
    strLinea = strLinea & separador & rsDatosEmp!Nro
    
    'Número de la calle (detdom.piso)
    If EsNulo(rsDatosEmp!piso) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!piso
    End If
    
    'Departamento u oficina (detdom.oficdepto)
    If EsNulo(rsDatosEmp!oficdepto) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!oficdepto
    End If
    
    'Torre (detdom.torre)
    If EsNulo(rsDatosEmp!torre) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!torre
    End If
    
    'Manzana (detdom.manzana)
    If EsNulo(rsDatosEmp!manzana) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!manzana
    End If
    
    'Código Postal (detdom.codigopostal)
    strLinea = strLinea & separador & rsDatosEmp!codigopostal

    'Entre calle (detdom.EntreCalles)
    If EsNulo(rsDatosEmp!entreCalles) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!entreCalles
    End If
    
    'Entre calle (detdom.Barrio)
    If EsNulo(rsDatosEmp!barrio) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!barrio
    End If
    
    'Localidad (localidad.locdesc)
    strLinea = strLinea & separador & rsDatosEmp!locdesc
        
    'Distrito (localidad.partido)
    If EsNulo(rsDatosEmp!partido) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!partido
    End If
        
    'Zona (zona.zonadesco)
    If EsNulo(rsDatosEmp!zonadesc) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!zonadesco
    End If
    
    'Provincia (provdesc.provdesc)
    strLinea = strLinea & separador & rsDatosEmp!provdesc
    
    'País(pais.paisdesc)
    strLinea = strLinea & separador & rsDatosEmp!paisdesc
    
    StrSql = "SELECT telnro From telefono " & _
            " INNER JOIN tipotel ON tipotel.titelnro = telefono.tipotel " & _
            " INNER JOIN cabdom ON cabdom.domnro = telefono.domnro " & _
            " WHERE cabdom.domnro= " & rsDatosEmp!domnro & " AND tipotel.titelnro= 1"
    OpenRecordset StrSql, rsAux
    
    'Telefono personal (telefono.telnro)
    
    If rsAux.EOF Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsAux!telnro
    End If
    
    
    StrSql = "SELECT telnro From telefono " & _
            " INNER JOIN tipotel ON tipotel.titelnro = telefono.tipotel " & _
            " INNER JOIN cabdom ON cabdom.domnro = telefono.domnro " & _
            " WHERE cabdom.domnro= " & rsDatosEmp!domnro & " AND tipotel.titelnro= 3"
    OpenRecordset StrSql, rsAux
    
    'Telefono Laboral - Fax (telefono.telnro)
    If rsAux.EOF Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsAux!telnro
    End If
    
    
    StrSql = "SELECT telnro From telefono " & _
            " INNER JOIN tipotel ON tipotel.titelnro = telefono.tipotel " & _
            " INNER JOIN cabdom ON cabdom.domnro = telefono.domnro " & _
            " WHERE cabdom.domnro= " & rsDatosEmp!domnro & " AND tipotel.titelnro= 2"
    OpenRecordset StrSql, rsAux
    
    'Telefono Celular (telefono.telnro)
    If rsAux.EOF Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsAux!telnro
    End If
    
    GoTo datosOk
CE:
    strLinea = ""
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 1001. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Function
    
datosOk:
    Call sincronizar_Det(ternro, 1001)
    expModelo1001 = strLinea
    Set rsDatosEmp = Nothing
    
End Function



'Exportación del modelo 1002 - Documentos
Function expModelo1002(ByVal ternro As Long, ByVal separador As String)
 Dim rsDatosEmp  As New ADODB.Recordset
 Dim strLinea As String

'1   Modelo  Nro de Modelo   Fijo "1002"
'2   Legajo  Nro Legajo  Empleg
'3   TipoDoc Tipo de Documento   Sigla
'4   NroDoc  Nro de documento

    On Error GoTo CE

    StrSql = "SELECT empleg,tipodocu.tidsigla,ter_doc.nrodoc FROM empleado  " & _
            " INNER JOIN ter_doc on ter_doc.ternro = empleado.ternro " & _
            " INNER JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro " & _
            " WHERE empleado.ternro = " & ternro
    OpenRecordset StrSql, rsDatosEmp
    
    
    strLinea = ""
    Do While Not rsDatosEmp.EOF
        'Modelo (Fijo) | Legajo | descripcion del tipo de documento (tipodocu.tidnom) | nro de documento (ter_doc.nrodoc)
        strLinea = strLinea & "1002" & separador & rsDatosEmp!empleg & separador & rsDatosEmp!tidsigla & separador & rsDatosEmp!nrodoc
        rsDatosEmp.MoveNext
        If Not rsDatosEmp.EOF Then
            strLinea = strLinea & Chr(10) + Chr(13)
        End If
    Loop
    
    
    GoTo datosOk
CE:
    strLinea = ""
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 1002. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Function
    
datosOk:
    Call sincronizar_Det(ternro, 1002)
    expModelo1002 = strLinea
    Set rsDatosEmp = Nothing
    
End Function



'Exportación del modelo 1003 - FASES
Function expModelo1003(ByVal ternro As Long, ByVal separador As String)
 Dim rsDatosEmp  As New ADODB.Recordset
 Dim strLinea As String

'1   Modelo  Nro de Modelo   Fijo "1003"
'2   Legajo  Nro Legajo  Empleg
'3   CausaBaja   Causa  de Baja  Puede ser nula
'4   FecAlta Fecha de alta
'5   FecBaja Fecha de baja   Puede ser nula
'6   Estado  Activa o inactiva   Activa / Inactiva
'7   Sueldo  Si es fase para sueldos o no    SI/NO
'8   Vac Si es fase para vacaciones o no SI/NO
'9   Indem   Si es fase para indemnizaciones o no    SI/NO
'10  al    Si es fase real o no    SI/NO
'    AltaReRec Si es Fecha de Alta reconocida o no SI/NO


    On Error GoTo CE

    StrSql = "SELECT empleg,causa.caudes,fases.altfec, fases.bajfec,fases.estado,fases.sueldo,fases.vacaciones " & _
            " ,fases.indemnizacion,fases.real,fases.fasrecofec " & _
            " FROM empleado  " & _
            " INNER JOIN fases on fases.empleado = empleado.ternro " & _
            " LEFT JOIN causa ON causa.caunro = fases.caunro " & _
            " WHERE empleado.ternro = " & ternro & " ORDER BY altfec DESC "
    OpenRecordset StrSql, rsDatosEmp
    
    'Modelo (Fijo) | Legajo
    strLinea = "1003" & separador & rsDatosEmp!empleg
    
    'Causa de baja
    If EsNulo(rsDatosEmp!caudes) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!caudes
    End If
    
    'Fecha de alta
    strLinea = strLinea & separador & rsDatosEmp!altfec
    
    'Fecha de baja
    If EsNulo(rsDatosEmp!bajfec) Then
        strLinea = strLinea & separador & "N/A"
    Else
        strLinea = strLinea & separador & rsDatosEmp!bajfec
    End If

    'Estado (activo | inactivo)
    If (rsDatosEmp!estado = 0) Then
        strLinea = strLinea & separador & "Inactiva"
    Else
        strLinea = strLinea & separador & "Activa"
    End If
    
    'Sueldo
    If (rsDatosEmp!sueldo = 0) Then
        strLinea = strLinea & separador & "NO"
    Else
        strLinea = strLinea & separador & "SI"
    End If
    
    'Vacaciones
    If (rsDatosEmp!vacaciones = 0) Then
        strLinea = strLinea & separador & "NO"
    Else
        strLinea = strLinea & separador & "SI"
    End If
    
    'Indemnizacion
    If (rsDatosEmp!indemnizacion = 0) Then
        strLinea = strLinea & separador & "NO"
    Else
        strLinea = strLinea & separador & "SI"
    End If

    'Indemnizacion
    If (rsDatosEmp!real = 0) Then
        strLinea = strLinea & separador & "NO"
    Else
        strLinea = strLinea & separador & "SI"
    End If

    'Fecha de alta reconocida
    If (rsDatosEmp!fasrecofec = 0) Then
        strLinea = strLinea & separador & "NO"
    Else
        strLinea = strLinea & separador & "SI"
    End If


    GoTo datosOk
CE:
    strLinea = ""
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 1003. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Function
datosOk:
    Call sincronizar_Det(ternro, 1003)
    expModelo1003 = strLinea
    Set rsDatosEmp = Nothing
    
End Function



'Exportación del modelo 1004 - Datos de los Familiares
Function expModelo1004(ByVal ternro As Long, ByVal separador As String)
 Dim rsDatosEmp  As New ADODB.Recordset
 Dim rsAux  As New ADODB.Recordset
 Dim rsDocumentos As New ADODB.Recordset
 Dim strLinea As String
 Dim empleg As Long

    On Error GoTo CE
' StrSql = "SELECT empleg,tercero.ternom,tercero.ternom2,tercero.terape,tercero.terape2,tercero.terfecnac,pais.paisdesc,nacionalidad.nacionaldes " & _
'            ",estcivil.estcivdesabr, ,tercero.tersex, tercero.terfecingempleado.empfecalta,empleado.empestudia,nivest.nivdesc,empleado.empemail" & _
'            ",empleado.empfbajaprev,empleado.empest,empleado.empremu,adptemplate.tplatedesabr,empleado.empreporta"
    StrSql = " SELECT  empleg FROM empleado  WHERE ternro= " & ternro
    OpenRecordset StrSql, rsDatosEmp
    
    If Not rsDatosEmp.EOF Then
        empleg = rsDatosEmp!empleg
    Else
        GoTo CE
    End If
    
    'Datos del Familiar
    StrSql = "SELECT tercero.ternro,tercero.ternom,tercero.ternom2,tercero.terape,tercero.terape2,tercero.terfecnac,pais.paisdesc,nacionalidad.nacionaldes " & _
            ",estcivil.estcivdesabr,tercero.tersex,parentesco.paredesc,familiar.famestudia,osocial.osdesc,planos.plnom,familiar.famemergencia " & _
            ",familiar.famsalario,familiar.famcargadgi, familiar.famfec, familiar.famfecvto" & _
            " FROM familiar " & _
            " INNER JOIN tercero ON tercero.ternro = familiar.ternro  " & _
            " INNER JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro " & _
            " INNER JOIN pais ON pais.paisnro = tercero.paisnro " & _
            " INNER JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro " & _
            " INNER JOIN parentesco ON parentesco.parenro = familiar.parenro " & _
            " LEFT JOIN planos ON planos.osocial = familiar.osocial " & _
            " LEFT JOIN osocial ON osocial.ternro = familiar.osocial " & _
            " WHERE familiar.empleado = " & ternro
    OpenRecordset StrSql, rsDatosEmp
    
    strLinea = ""
    Do While Not rsDatosEmp.EOF
        
        'Obtiene el documento del familiar
        StrSql = "SELECT tipodocu.tidsigla,ter_doc.nrodoc FROM tercero " & _
                " INNER JOIN ter_doc on ter_doc.ternro = tercero.ternro " & _
                " INNER JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro " & _
                " WHERE Tercero.ternro = " & rsDatosEmp!ternro
        OpenRecordset StrSql, rsDocumentos
        
        Do While Not rsDocumentos.EOF

            strLinea = strLinea & "1004" & separador & empleg
            strLinea = strLinea & separador & rsDatosEmp!terape & "@" & rsDatosEmp!terape2
            strLinea = strLinea & separador & rsDatosEmp!ternom & "@" & rsDatosEmp!ternom2
            strLinea = strLinea & separador & rsDatosEmp!terfecnac
            strLinea = strLinea & separador & rsDatosEmp!paisdesc
            strLinea = strLinea & separador & rsDatosEmp!nacionaldes
            strLinea = strLinea & separador & rsDatosEmp!estcivdesabr
            
            If (CLng(rsDatosEmp!tersex) = -1) Then
                strLinea = strLinea & separador & "M"
            Else
                strLinea = strLinea & separador & "F"
            End If
            
            'Descripcion del parentezco
            strLinea = strLinea & separador & rsDatosEmp!paredesc
            
            
            'Si el familiar estudia- Sigue la misma lógica que "familiar_01.asp" Que no significa que este bien, sino que es consistente.
            StrSql = "SELECT nivdesc FROM estudio_actual " & _
                    " LEFT JOIN nivest ON nivest.nivnro = estudio_actual.nivnro " & _
                    " WHERE ternro = " & rsDatosEmp!ternro
            OpenRecordset StrSql, rsAux
        
            If rsAux.EOF Then
                strLinea = strLinea & separador & "NO" & separador & "N/A"
            Else
                strLinea = strLinea & separador & "SI" & separador & rsAux!nivdesc
            End If
            
            strLinea = strLinea & separador & rsDocumentos!tidsigla & separador & rsDocumentos!nrodoc
            
            'Busca los datos del domicilio
            StrSql = "SELECT cabdom.domnro, detdom.calle,detdom.nro,detdom.piso,detdom.oficdepto,detdom.torre,detdom.manzana " & _
                    " ,detdom.codigopostal,detdom.EntreCalles,detdom.Barrio,localidad.locdesc,partido.partnom partido,zona.zonadesc,provincia.provdesc " & _
                    " ,pais.paisdesc " & _
                    " FROM tercero " & _
                    " INNER JOIN cabdom ON cabdom.ternro = tercero.ternro " & _
                    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
                    " INNER JOIN tipodomi ON tipodomi.tidonro = cabdom.tidonro " & _
                    " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
                    " LEFT JOIN partido ON partido.partnro = detdom.partnro  " & _
                    " LEFT JOIN zona ON zona.zonanro = localidad.zonanro " & _
                    " INNER JOIN provincia ON provincia.provnro = localidad.provnro " & _
                    " INNER JOIN pais ON pais.paisnro = provincia.paisnro " & _
                    " WHERE tercero.ternro = " & rsDatosEmp!ternro
            OpenRecordset StrSql, rsAux
    
            If Not rsAux.EOF Then
                'Nombre de la calle (detdom.calle)
                strLinea = strLinea & separador & rsAux!calle
        
                'Número de la calle (detdom.nro)
                strLinea = strLinea & separador & rsAux!Nro
        
                'Número de la calle (detdom.piso)
                If EsNulo(rsAux!piso) Then
                    strLinea = strLinea & separador & "N/A"
                Else
                    strLinea = strLinea & separador & rsAux!piso
                End If
        
                'Departamento u oficina (detdom.oficdepto)
                If EsNulo(rsAux!oficdepto) Then
                    strLinea = strLinea & separador & "N/A"
                Else
                    strLinea = strLinea & separador & rsAux!oficdepto
                End If
        
                'Torre (detdom.torre)
                If EsNulo(rsAux!torre) Then
                    strLinea = strLinea & separador & "N/A"
                Else
                    strLinea = strLinea & separador & rsAux!torre
                End If
        
                'Manzana (detdom.manzana)
                If EsNulo(rsAux!manzana) Then
                    strLinea = strLinea & separador & "N/A"
                Else
                    strLinea = strLinea & separador & rsAux!manzana
                End If
        
                'Código Postal (detdom.codigopostal)
                strLinea = strLinea & separador & rsAux!codigopostal
    
                'Entre calle (detdom.EntreCalles)
                If EsNulo(rsAux!entreCalles) Then
                    strLinea = strLinea & separador & "N/A"
                Else
                    strLinea = strLinea & separador & rsAux!entreCalles
                End If
        
                'Entre calle (detdom.Barrio)
                If EsNulo(rsAux!barrio) Then
                    strLinea = strLinea & separador & "N/A"
                Else
                    strLinea = strLinea & separador & rsAux!barrio
                End If
        
                'Localidad (localidad.locdesc)
                strLinea = strLinea & separador & rsAux!locdesc
            
                'Distrito (localidad.partido)
                If EsNulo(rsAux!partido) Then
                    strLinea = strLinea & separador & "N/A"
                Else
                    strLinea = strLinea & separador & rsAux!partido
                End If
            
                'Zona (zona.zonadesco)
                If EsNulo(rsAux!zonadesc) Then
                    strLinea = strLinea & separador & "N/A"
                Else
                    strLinea = strLinea & separador & rsAux!zonadesco
                End If
        
                'Provincia (provdesc.provdesc)
                strLinea = strLinea & separador & rsAux!provdesc
        
                'País(pais.paisdesc)
                strLinea = strLinea & separador & rsAux!paisdesc
            
                StrSql = "SELECT telnro From telefono " & _
                        " INNER JOIN tipotel ON tipotel.titelnro = telefono.tipotel " & _
                        " INNER JOIN cabdom ON cabdom.domnro = telefono.domnro " & _
                        " WHERE cabdom.domnro= " & rsAux!domnro & " AND tipotel.titelnro= 1"
                OpenRecordset StrSql, rsAux
                
                'Telefono personal (telefono.telnro)
                If rsAux.EOF Then
                    strLinea = strLinea & separador & "N/A"
                Else
                    strLinea = strLinea & separador & rsAux!telnro
                End If
            
            Else
                strLinea = strLinea & separador & "N/A" & separador & "N/A" & separador & "N/A" & separador & "N/A" & separador & "N/A"
                strLinea = strLinea & separador & "N/A" & separador & "N/A" & separador & "N/A" & separador & "N/A" & separador & "N/A"
                strLinea = strLinea & separador & "N/A" & separador & "N/A" & separador & "N/A" & separador & "N/A" & separador & "N/A"
            End If
            
            'Obra social - Fax (familiar.OSocial)
            If EsNulo(rsDatosEmp!osdesc) Then
                strLinea = strLinea & separador & "N/A"
            Else
                strLinea = strLinea & separador & rsDatosEmp!osdesc
            End If
            
            'Plan de obra social (planos.plnom)
            If EsNulo(rsDatosEmp!plnom) Then
                strLinea = strLinea & separador & "N/A"
            Else
                strLinea = strLinea & separador & rsDatosEmp!plnom
            End If
            
            
            'Aviso de emergencia(familiar.famemergencia)
            If EsNulo(rsDatosEmp!famemergencia) Or (rsDatosEmp!famemergencia = 0) Then
                strLinea = strLinea & separador & "NO"
            Else
                strLinea = strLinea & separador & "SI"
            End If
            
            'salario familiar (familiar.famsalario)
            If EsNulo(rsDatosEmp!famsalario) Or (rsDatosEmp!famsalario = 0) Then
                strLinea = strLinea & separador & "NO"
            Else
                strLinea = strLinea & separador & "SI"
            End If
            
            'Ganancia (familiar.famcargadgi)
            If EsNulo(rsDatosEmp!famcargadgi) Or (rsDatosEmp!famcargadgi = 0) Then
                strLinea = strLinea & separador & "NO"
            Else
                strLinea = strLinea & separador & "SI"
            End If
            
            'Fecha de Inicio del Vinculo
            If EsNulo(rsDatosEmp!famfec) Then
                strLinea = strLinea & separador & "N/A"
            Else
                strLinea = strLinea & separador & rsDatosEmp!famfec
            End If
            
            'Fecha de Inicio de Vto del Vinculo
            If EsNulo(rsDatosEmp!famfecvto) Then
                strLinea = strLinea & separador & "N/A"
            Else
                strLinea = strLinea & separador & rsDatosEmp!famfecvto
            End If
                                               
            rsDocumentos.MoveNext
            If Not rsDocumentos.EOF Then
                strLinea = strLinea & Chr(10) + Chr(13)
            End If
        Loop
        rsDatosEmp.MoveNext
        If Not rsDatosEmp.EOF Then
            strLinea = strLinea & Chr(10) + Chr(13)
        End If
    Loop
    
    GoTo datosOk
CE:
    strLinea = ""
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 1004. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Function
datosOk:
    Call sincronizar_Det(ternro, 1004)
    expModelo1004 = strLinea
    Set rsDatosEmp = Nothing
    
End Function


'Exportación del modelo 1005 - Estructuras
Function expModelo1005(ByVal ternro As Long, ByVal separador As String)
 Dim rsDatosEmp  As New ADODB.Recordset
 Dim strLinea As String

'Modelo  Nro de Modelo
'Legajo  Nro Legajo
'TipoEstr    Descripción del tipo de estructura
'Estructura  Descripción de la estructura
'FecDesde    Fecha desde
'FecHasta    Fecha Hasta


    On Error GoTo CE

    StrSql = "SELECT empleg, tipoestructura.tenro, tedabr, estrdabr, teorden, his_estructura.htetdesde, his_estructura.htethasta " & _
            " FROM his_estructura " & _
            " INNER JOIN tipoestructura ON his_estructura.tenro = tipoestructura.tenro " & _
            " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
            " INNER JOIN empleado ON his_estructura.ternro = empleado.ternro " & _
            " WHERE his_estructura.ternro =" & ternro & " AND his_estructura.htetdesde <= " & ConvFecha(Date) & _
            " AND ((his_estructura.htethasta IS NULL) OR (his_estructura.htethasta >=" & ConvFecha(Date) & " ))"
    OpenRecordset StrSql, rsDatosEmp
    
    
    strLinea = ""
    Do While Not rsDatosEmp.EOF
        'Modelo (Fijo) | Legajo | descripcion del tipo de documento (tipodocu.tidnom) | nro de documento (ter_doc.nrodoc)
        strLinea = strLinea & "1005" & separador & rsDatosEmp!empleg & separador & rsDatosEmp!tedabr & separador & rsDatosEmp!estrdabr & separador & rsDatosEmp!htetdesde & separador & rsDatosEmp!htethasta
        rsDatosEmp.MoveNext
        If Not rsDatosEmp.EOF Then
            strLinea = strLinea & Chr(10) + Chr(13)
        End If
    Loop
    
    
    GoTo datosOk
CE:
    strLinea = ""
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 1005. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Function
datosOk:
    Call sincronizar_Det(ternro, 1005)
    expModelo1005 = strLinea
    Set rsDatosEmp = Nothing
    
End Function

'Sincroniza los datos para mostrar los recibos de sueldo desde ess
Sub expModelo405(ByVal progreso As Double, ByVal bpronro As Long)
 
 Dim rs_Datos As New ADODB.Recordset
 Dim rs_datosSinc As New ADODB.Recordset
 Dim rs_datosSinc2 As New ADODB.Recordset
 Dim rs_datosSinc3 As New ADODB.Recordset
 Dim rs_datosSinc4 As New ADODB.Recordset
 Dim strConexionExt As String
 Dim porc As Double
 Dim cjpb As String
 Dim irpf As String
 Dim cliqnro As Long
 Dim confcjpb As Boolean    'variable para saber si busco un CO (true) o un AC (false)
 Dim confirpf As Boolean    'variable para saber si busco un CO (true) o un AC (false)

    On Error GoTo CE
    
    Flog.writeline Espacios(Tabulador * 0) & "Ingreso al modelo 405"
    
    
    '*******************************************************************************************************************
    'Buscamos los terceros del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de terceros tipo empleado"
    progreso = progreso + porc
    StrSql = " SELECT tercero.* FROM tercero " & _
             " INNER JOIN empleado ON empleado.ternro = tercero.ternro AND empleado.empest = -1 "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT ternro FROM tercero WHERE ternro = " & rs_Datos!ternro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("tercero", rs_Datos, "ternro = " & rs_Datos!ternro)
        Else
            Call insertarDatos("tercero", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de terceros tipo empleado"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los empleados del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de empleados activos"
    progreso = progreso + porc
    StrSql = " SELECT empleado.* FROM empleado " & _
             " INNER JOIN tercero ON tercero.ternro = empleado.ternro " & _
             " WHERE empleado.empest = -1 "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT ternro FROM empleado WHERE ternro = " & rs_Datos!ternro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("empleado", rs_Datos, "ternro = " & rs_Datos!ternro)
        Else
            Call insertarDatos("empleado", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de empleados activos"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    StrSql = "SELECT empsinc_det.bpronro FROM empsinc " & _
            " INNER JOIN empsinc_det ON empsinc_det.ternro = empsinc.esternro AND empsinc_det.ternro = 0 " & _
            " WHERE empsinc.esternro = 0 "
    rs_Datos.CursorType = 3
    OpenRecordset StrSql, rs_Datos
    
    If rs_Datos.RecordCount = 0 Then
        porc = 100
    End If
    porc = CLng(100 / IIf(rs_Datos.RecordCount = 0, 1, rs_Datos.RecordCount))
    progreso = 100 - progreso
    
    If Not rs_Datos.EOF Then
        
        'Busco la configuracion de los conceptos o acumuladores que muestran el cjpb y irpf
        StrSql = "SELECT conftipo, confval2, confval, confnrocol FROM confrep WHERE repnro = 60 AND confnrocol in (404,405) "
        OpenRecordset StrSql, rs_datosSinc
        cjpb = "0"
        irpf = "0"
        Do While Not rs_datosSinc.EOF
            Select Case rs_datosSinc!confnrocol
                Case 404
                    If UCase(rs_datosSinc!conftipo) = "CO" Then
                        confcjpb = True
                        cjpb = rs_datosSinc!confval2
                    Else
                        confcjpb = False
                        cjpb = rs_datosSinc!confval
                    End If
                    
                Case 405
                    If UCase(rs_datosSinc!conftipo) = "CO" Then
                        confirpf = True
                        irpf = rs_datosSinc!confval2
                    Else
                        confirpf = False
                        irpf = rs_datosSinc!confval
                    End If
            
            End Select
            rs_datosSinc.MoveNext
        Loop
        
        Flog.writeline Espacios(Tabulador * 0) & "CJPB configurado: " & cjpb
        Flog.writeline Espacios(Tabulador * 0) & "IRPF configurado: " & irpf
        
        Do While Not rs_Datos.EOF
            progreso = progreso + porc
            'Cabecera del recibo de sueldo
            StrSql = " SELECT * FROM rep_recibo WHERE bpronro = " & rs_Datos!bpronro
            OpenRecordset StrSql, rs_datosSinc
            Do While Not rs_datosSinc.EOF
                StrSql = " SELECT * FROM rep_recibo " & _
                         " WHERE pronro = " & rs_datosSinc!pronro & " AND bpronro = " & rs_datosSinc!bpronro & " AND ternro = " & rs_datosSinc!ternro
                OpenRecordsetExt StrSql, rs_datosSinc3, ExtConn
                If Not rs_datosSinc3.EOF Then
                    Call actualizarRegistro("rep_recibo", rs_datosSinc, "pronro = " & rs_datosSinc3!pronro & " AND bpronro = " & rs_datosSinc3!bpronro & " AND ternro = " & rs_datosSinc3!ternro)
                Else
                    Call insertarDatos("rep_recibo", rs_datosSinc)
                End If
                
                rs_datosSinc.MoveNext
            Loop
            
            'Detalle del recibo de sueldo
            StrSql = " SELECT * FROM rep_recibo_det WHERE bpronro = " & rs_Datos!bpronro
            OpenRecordset StrSql, rs_datosSinc2
            Do While Not rs_datosSinc2.EOF
                StrSql = " SELECT rep_recibo_det.* FROM rep_recibo_det " & _
                         " WHERE cliqnro = " & rs_datosSinc2!cliqnro & " AND conccod = '" & rs_datosSinc2!Conccod & "'"
                OpenRecordsetExt StrSql, rs_datosSinc3, ExtConn
                If Not rs_datosSinc3.EOF Then
                    Call actualizarRegistro("rep_recibo_det", rs_datosSinc2, "cliqnro = " & rs_datosSinc3!cliqnro & " AND concnro = " & rs_datosSinc3!concnro)
                Else
                    Call insertarDatos("rep_recibo_det", rs_datosSinc2)
                End If
                'Busco los detliq del empleado
                cliqnro = rs_datosSinc2!cliqnro
                
                'actualizo el cjpb
                If confcjpb Then
                    'Busco el valor del sistema completo
                    StrSql = " SELECT detliq.* FROM detliq " & _
                             " INNER JOIN concepto ON concepto.concnro =  detliq.concnro AND concepto.conccod=  '" & cjpb & "'" & _
                             " WHERE detliq.cliqnro = " & rs_datosSinc2!cliqnro
                    OpenRecordset StrSql, rs_datosSinc4
                    If Not rs_datosSinc4.EOF Then
                        'es un concepto busco si existe en el sistema reducido y lo actualizo o inserto
                        StrSql = " SELECT detliq.* FROM detliq " & _
                                 " INNER JOIN concepto ON concepto.concnro =  detliq.concnro AND concepto.conccod=  '" & cjpb & "'" & _
                                 " WHERE detliq.cliqnro = " & rs_datosSinc2!rs_datosSinc4
                        OpenRecordsetExt StrSql, rs_datosSinc3, ExtConn
                        If Not rs_datosSinc3.EOF Then
                            Call actualizarRegistro("detliq", rs_datosSinc4, "cliqnro = " & rs_datosSinc4!cliqnro & " AND concnro = " & rs_datosSinc4!concnro)
                        Else
                            Call insertarDatos("detliq", rs_datosSinc4)
                        End If
                    End If
                Else
                    'Busco el valor del concepto del sistema completo
                    StrSql = " SELECT acu_liq.* FROM acu_liq WHERE acunro = " & cjpb & " AND cliqnro = " & rs_datosSinc2!cliqnro
                    OpenRecordset StrSql, rs_datosSinc4
                    If Not rs_datosSinc4.EOF Then
                        'es un acumulador busco si existe en el sistema reducido y lo actualizo o inserto
                        StrSql = " SELECT acu_liq.* FROM acu_liq WHERE acunro = " & cjpb & " AND cliqnro = " & rs_datosSinc4!cliqnro
                        OpenRecordsetExt StrSql, rs_datosSinc3, ExtConn
                        If Not rs_datosSinc3.EOF Then
                            Call actualizarRegistro("acu_liq", rs_datosSinc4, "cliqnro = " & rs_datosSinc4!cliqnro & " AND acunro = " & rs_datosSinc4!acunro)
                        Else
                            Call insertarDatos("acu_liq", rs_datosSinc4)
                        End If
                    End If
                End If
                
                'actualizo el irpf
                If confirpf Then
                    'Busco el valor del concepto del sistema completo
                    StrSql = " SELECT detliq.* FROM detliq " & _
                             " INNER JOIN concepto ON concepto.concnro =  detliq.concnro AND concepto.conccod=  '" & irpf & "'" & _
                             " WHERE detliq.cliqnro = " & rs_datosSinc2!cliqnro
                    OpenRecordset StrSql, rs_datosSinc4
                        If Not rs_datosSinc4.EOF Then
                        'es un concepto busco si existe en el sistema reducido y lo actualizo o inserto
                        StrSql = " SELECT detliq.* FROM detliq " & _
                                 " INNER JOIN concepto ON concepto.concnro =  detliq.concnro AND concepto.conccod=  '" & irpf & "'" & _
                                 " WHERE detliq.cliqnro = " & rs_datosSinc4!cliqnro
                        OpenRecordsetExt StrSql, rs_datosSinc3, ExtConn
                        If Not rs_datosSinc3.EOF Then
                            Call actualizarRegistro("detliq", rs_datosSinc4, "cliqnro = " & rs_datosSinc4!cliqnro & " AND concnro = " & rs_datosSinc4!concnro)
                        Else
                            Call insertarDatos("detliq", rs_datosSinc4)
                        End If
                    End If
                Else
                    'Busco el valor del concepto del sistema completo
                    StrSql = " SELECT acu_liq.* FROM acu_liq WHERE acunro = " & irpf & " AND cliqnro = " & rs_datosSinc2!cliqnro
                    OpenRecordset StrSql, rs_datosSinc4
                    If Not rs_datosSinc4.EOF Then
                        'es un acumulador busco si existe en el sistema reducido y lo actualizo o inserto
                        StrSql = " SELECT acu_liq.* FROM acu_liq WHERE acunro = " & irpf & " AND cliqnro = " & rs_datosSinc2!cliqnro
                        OpenRecordsetExt StrSql, rs_datosSinc4, ExtConn
                        If Not rs_datosSinc3.EOF Then
                            Call actualizarRegistro("acu_liq", rs_datosSinc4, "cliqnro = " & rs_datosSinc4!cliqnro & " AND acunro = " & rs_datosSinc4!acunro)
                        Else
                            Call insertarDatos("acu_liq", rs_datosSinc4)
                        End If
                    End If
                End If
                
                rs_datosSinc2.MoveNext
            Loop
            
            Flog.writeline Espacios(Tabulador * 0) & "Comienza la busqueda de recibos aprobados del sistema completo"
            StrSql = " SELECT bpronro, aprobado FROM gesthistoliq WHERE aprobado = -1 AND bpronro = " & rs_Datos!bpronro
            OpenRecordset StrSql, rs_datosSinc
            Do While Not rs_datosSinc.EOF
                Call recibosAprobados(rs_datosSinc)
                rs_datosSinc.MoveNext
            Loop
            Flog.writeline Espacios(Tabulador * 0) & "Finalizo la busqueda de recibos aprobados del sistema completo"
            
            
            Flog.writeline Espacios(Tabulador * 0) & "Comienza la busqueda de recibos aprobados por los usuario del sistema"
            StrSql = " SELECT * FROM rep_recibo_digital WHERE estado = 0 "
            OpenRecordset StrSql, rs_datosSinc
            Do While Not rs_datosSinc.EOF
                Call recibosAprobadosUsuario(rs_datosSinc)
                rs_datosSinc.MoveNext
            Loop
            Flog.writeline Espacios(Tabulador * 0) & "Finalizo la busqueda de recibos aprobados por los usuario del sistema"
            
            'Borro el registro de sincronizacion
            StrSql = " DELETE FROM empsinc_det WHERE ternro = 0 AND bpronro = " & rs_Datos!bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Muevo al proximo registro desincronizado
            rs_Datos.MoveNext
            
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Loop

    
        'Busco los datos del lado de la base reducida y actualizo el sistema completo
        Flog.writeline Espacios(Tabulador * 0) & "Comienza la busqueda de recibos aprobados por el usuario del sistema reducido"
        StrSql = " SELECT cliqnro, estado, fechaestado, Obs FROM rep_recibo_digital WHERE estado = 0 "
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        Do While Not rs_datosSinc.EOF
            StrSql = " UPDATE rep_recibo_digital SET " & _
                     " estado = " & rs_datosSinc!estado & ", fechaestado = '" & rs_datosSinc!fechaestado & "', obs = '" & rs_datosSinc!Obs & "'" & _
                     " WHERE cliqnro = " & rs_datosSinc!cliqnro
            objConn.Execute StrSql, , adExecuteNoRecords

            rs_datosSinc.MoveNext
        Loop
        Flog.writeline Espacios(Tabulador * 0) & "Finalizo la busqueda de recibos aprobados por el usuario del sistema reducido"

    Else
        Flog.writeline Espacios(Tabulador * 0) & "No hay datos para sincronizar"
    End If
    
    
    GoTo datosOk
CE:
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 405. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Sub
datosOk:
   ' Call sincronizar_Det(ternro, 1005)
    Set rs_Datos = Nothing
    Set rs_datosSinc = Nothing
    Set rs_datosSinc2 = Nothing
    Set rs_datosSinc3 = Nothing
    Set rs_datosSinc4 = Nothing

End Sub

'sincroniza los datos para mostrar en el tablero de GTI desde ESS
Sub expModelo406(ByVal fechaDesde As String, ByVal fechaHasta As String, ByVal progreso As Double, ByVal bpronro As Long)
Dim rs_Datos As New ADODB.Recordset
Dim rs_datosSinc As New ADODB.Recordset
Dim strConexionExt As String
Dim porc As Double


 
    On Error GoTo CE
    Flog.writeline Espacios(Tabulador * 0) & "Ingreso al modelo 406"
    
    'si se disparo automaticamente tomo el mes actual
    If Trim(fechaDesde) = "" Then
        fechaDesde = "01/" & Month(Date) & "/" & Year(Date)
        fechaHasta = DateAdd("d", -1, DateAdd("m", 1, "01/" & Month(Date) & "/" & Year(Date)))
    End If
    
    Flog.writeline Espacios(Tabulador * 0) & "Fechas de busqueda de datos, desde: " & fechaDesde & ", hasta:" & fechaHasta
    
    'cantidad de secciones de actualizacion
    porc = CLng(100 / 23)
    
    
    
    '*******************************************************************************************************************
    'Buscamos los paises del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de paises"
    progreso = (100 - progreso) + porc
    StrSql = " SELECT * FROM pais "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT paisnro FROM pais WHERE paisnro = " & rs_Datos!paisnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
                Call actualizarRegistro("pais", rs_Datos, "paisnro = " & rs_Datos!paisnro)
        Else
            Call insertarDatos("pais", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de paises"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    
    '*******************************************************************************************************************
    'Buscamos los lugar de nacimiento del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de lugar de nacimiento"
    progreso = progreso + porc
    StrSql = " SELECT * FROM lugar_nac "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT lugarnro FROM lugar_nac WHERE lugarnro = " & rs_Datos!lugarnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("lugar_nac", rs_Datos, "lugarnro = " & rs_Datos!lugarnro)
        Else
            Call insertarDatos("lugar_nac", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de lugar de nacimiento"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    '*******************************************************************************************************************
    'Buscamos los terceros del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de terceros tipo empleado"
    progreso = progreso + porc
    StrSql = " SELECT tercero.* FROM tercero " & _
             " INNER JOIN empleado ON empleado.ternro = tercero.ternro AND empleado.empest = -1 "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT ternro FROM tercero WHERE ternro = " & rs_Datos!ternro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("tercero", rs_Datos, "ternro = " & rs_Datos!ternro)
        Else
            Call insertarDatos("tercero", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de terceros tipo empleado"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los empleados del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de empleados activos"
    progreso = progreso + porc
    StrSql = " SELECT empleado.* FROM empleado " & _
             " INNER JOIN tercero ON tercero.ternro = empleado.ternro " & _
             " WHERE empleado.empest = -1 "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT ternro FROM empleado WHERE ternro = " & rs_Datos!ternro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("empleado", rs_Datos, "ternro = " & rs_Datos!ternro)
        Else
            Call insertarDatos("empleado", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de empleados activos"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los tipos de hora
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de tipos de horas"
    progreso = progreso + porc
    StrSql = " SELECT * FROM tiphora "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT thnro FROM tiphora WHERE thnro = " & rs_Datos!thnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("tiphora", rs_Datos, "thnro = " & rs_Datos!thnro)
        Else
            Call insertarDatos("tiphora", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de tipos de horas"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los tipos de licencias
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de tipos de licencias"
    progreso = progreso + porc
    StrSql = " SELECT * FROM tipdia "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT tdnro FROM tipdia WHERE tdnro = " & rs_Datos!tdnroo
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("tipdia", rs_Datos, "tdnro = " & rs_Datos!tdnro)
        Else
            Call insertarDatos("tipdia", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de tipos de licencias"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los tipos de tarjetas del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de tipos de tarjetas"
    progreso = progreso + porc
    StrSql = " SELECT * FROM gti_tiptar "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT tptrnro FROM gti_tiptar WHERE tptrnro = " & rs_Datos!tptrnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_tiptar", rs_Datos, "tptrnro = " & rs_Datos!tptrnro)
        Else
            Call insertarDatos("gti_tiptar", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de tipos de tarjetas"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los relojes del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de relojes"
    progreso = progreso + porc
    StrSql = "SELECT * FROM gti_reloj"
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT relnro FROM gti_reloj WHERE relnro = " & rs_Datos!relnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_reloj", rs_Datos, "relnro = " & rs_Datos!relnro)
        Else
            Call insertarDatos("gti_reloj", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de relojes"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los motivo de novedad horaria del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de motivo de novedad horaria"
    progreso = progreso + porc
    StrSql = "SELECT * FROM gti_motivo"
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT motnro FROM gti_motivo WHERE motnro = " & rs_Datos!motnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_motivo", rs_Datos, "motnro = " & rs_Datos!motnro)
        Else
            Call insertarDatos("gti_motivo", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de motivo de novedad horaria"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los tipos de justificacion
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de tipos de justificacion"
    progreso = progreso + porc
    StrSql = "SELECT * FROM gti_tipojust"
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT tjusnro FROM gti_tipojust WHERE tjusnro = " & rs_Datos!tjusnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_tipojust", rs_Datos, "tjusnro = " & rs_Datos!tjusnro)
        Else
            Call insertarDatos("gti_tipojust", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de tipos de justificacion"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos las justificaciones
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de justificacion"
    progreso = progreso + porc
    StrSql = " SELECT * FROM gti_justificacion " & _
             " WHERE (jusdesde <= " & cambiaFecha(fechaDesde) & " AND (jushasta >= " & cambiaFecha(fechaHasta) & " or jushasta >= " & cambiaFecha(fechaDesde) & ")) " & _
             " OR (jusdesde >= " & cambiaFecha(fechaDesde) & " AND (jusdesde <= " & cambiaFecha(fechaHasta) & "))"

    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT jusnro FROM gti_justificacion WHERE jusnro = " & rs_Datos!jusnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_justificacion", rs_Datos, "jusnro = " & rs_Datos!jusnro)
        Else
            Call insertarDatos("gti_justificacion", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de justificacion"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    '*******************************************************************************************************************
    'Buscamos las subturnos
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de subturnos"
    progreso = progreso + porc
    StrSql = "SELECT * FROM gti_subturno "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT subturnro FROM gti_subturno WHERE subturnro = " & rs_Datos!subturnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_subturno", rs_Datos, "subturnro = " & rs_Datos!subturnro)
        Else
            Call insertarDatos("gti_subturno", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de subturnos"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos las dias de los dias del subturnos
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de dias del subturnos"
    progreso = progreso + porc
    StrSql = "SELECT * FROM gti_dias "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT dianro FROM gti_dias WHERE dianro = " & rs_Datos!dianro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_dias", rs_Datos, "dianro = " & rs_Datos!dianro)
        Else
            Call insertarDatos("gti_dias", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de dias del subturnos"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los tipos de partes
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de tipos de partes"
    progreso = progreso + porc
    StrSql = "SELECT * FROM gti_tipoparte "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT gtpnro FROM gti_tipoparte WHERE gtpnro = " & rs_Datos!gtpnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_tipoparte", rs_Datos, "gtpnro = " & rs_Datos!gtpnro)
        Else
            Call insertarDatos("gti_tipoparte", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de tipos de partes"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos las cabeceras de partes
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de cabeceras de partes"
    progreso = progreso + porc
    StrSql = " SELECT * FROM gti_cabparte WHERE " & _
             " (gcpdesde <= " & cambiaFecha(fechaDesde) & " AND (gcphasta >= " & cambiaFecha(fechaHasta) & " OR gcphasta >= " & cambiaFecha(fechaDesde) & ")) " & _
             " OR (gcpdesde >= " & cambiaFecha(fechaDesde) & " AND (gcpdesde <= " & cambiaFecha(fechaHasta) & ")) "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT gcpnro FROM gti_cabparte WHERE gcpnro = " & rs_Datos!gcpnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_cabparte", rs_Datos, "gcpnro = " & rs_Datos!gcpnro)
        Else
            Call insertarDatos("gti_cabparte", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de cabeceras de partes"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos los tipos de novedad
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de tipos de novedad"
    progreso = progreso + porc
    StrSql = "SELECT * FROM gti_tiponovedad "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT gtnovnro FROM gti_tiponovedad WHERE gtnovnro = " & rs_Datos!gtnovnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_tiponovedad", rs_Datos, "gtnovnro = " & rs_Datos!gtnovnro)
        Else
            Call insertarDatos("gti_tiponovedad", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de tipos de novedad"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos las anormalidades
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de anormalidades"
    progreso = progreso + porc
    StrSql = "SELECT * FROM gti_anormalidad "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT normnro FROM gti_anormalidad WHERE normnro = " & rs_Datos!normnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_anormalidad", rs_Datos, "normnro = " & rs_Datos!normnro)
        Else
            Call insertarDatos("gti_anormalidad", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de anormalidades"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    '*******************************************************************************************************************
    'Buscamos el horario cumplido
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de horario cumplido"
    progreso = progreso + porc
    StrSql = " SELECT * FROM gti_horcumplido WHERE " & _
             " (hordesde <= " & cambiaFecha(fechaDesde) & " AND (horhasta >= " & cambiaFecha(fechaHasta) & " OR horhasta >= " & cambiaFecha(fechaDesde) & ")) " & _
             " OR (hordesde >= " & cambiaFecha(fechaDesde) & " AND (hordesde <= " & cambiaFecha(fechaHasta) & "))"
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT hornro FROM gti_horcumplido WHERE hornro = " & rs_Datos!hornro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_horcumplido", rs_Datos, "hornro = " & rs_Datos!hornro)
        Else
            Call insertarDatos("gti_horcumplido", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de horario cumplido"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos el acumulado diario
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de acumulado diario"
    progreso = progreso + porc
    StrSql = " SELECT * FROM gti_acumdiario " & _
             " WHERE adfecha >= " & cambiaFecha(fechaDesde) & " AND adfecha <= " & cambiaFecha(fechaHasta) & _
             " ORDER BY ternro "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT ternro FROM gti_acumdiario WHERE ternro = " & rs_Datos!ternro & _
                 " AND thnro = " & rs_Datos!thnro & " AND adfecha = " & cambiaFecha(rs_Datos!adfecha)
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_acumdiario", rs_Datos, "ternro = " & rs_Datos!ternro & " AND thnro = " & rs_Datos!thnro & " AND adfecha = " & cambiaFecha(rs_Datos!adfecha))
        Else
            Call insertarDatos("gti_acumdiario", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de acumulado diario"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    '*******************************************************************************************************************
    'Buscamos el parte de asignacion horaria
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de parte de asignacion horaria"
    progreso = progreso + porc
    StrSql = " SELECT * FROM gti_detturtemp WHERE " & _
             " (gttempdesde <= " & cambiaFecha(fechaDesde) & " AND (gttemphasta >= " & cambiaFecha(fechaHasta) & " or gttemphasta >= " & cambiaFecha(fechaDesde) & ")) " & _
             " OR (gttempdesde >= " & cambiaFecha(fechaDesde) & " AND (gttempdesde <= " & cambiaFecha(fechaHasta) & ")) " & _
             " ORDER BY ternro "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = " SELECT ternro FROM gti_detturtemp WHERE gcpnro = " & rs_Datos!gcpnro & " AND ternro = " & rs_Datos!ternro & _
                 " AND ttemphdesde1 = '" & rs_Datos!ttemphdesde1 & "' AND ttemphdesde2 = '" & rs_Datos!ttemphdesde2 & "'" & _
                 " AND ttemphdesde3 = '" & rs_Datos!ttemphdesde3 & "' AND ttemphhasta1 = '" & rs_Datos!ttemphhasta1 & "'" & _
                 " AND ttemphhasta2 = '" & rs_Datos!ttemphhasta2 & "' AND ttemphhasta3 = '" & rs_Datos!ttemphhasta3 & "'" & _
                 " AND gttempdesde = " & cambiaFecha(rs_Datos!gttempdesde) & " AND  gttemphasta = " & cambiaFecha(rs_Datos!gttemphasta)
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_detturtemp", rs_Datos, "gcpnro = " & rs_Datos!gcpnro & " AND ternro = " & rs_Datos!ternro & " AND ttemphdesde1 = '" & rs_Datos!ttemphdesde1 & "' AND ttemphdesde2 = '" & rs_Datos!ttemphdesde2 & "' AND ttemphdesde3 = '" & rs_Datos!ttemphdesde3 & "' AND ttemphhasta1 = '" & rs_Datos!ttemphhasta1 & "' AND ttemphhasta2 = '" & rs_Datos!ttemphhasta2 & "' AND ttemphhasta3 = '" & rs_Datos!ttemphhasta3 & "' AND gttempdesde = " & cambiaFecha(rs_Datos!gttempdesde) & " AND  gttemphasta = " & cambiaFecha(rs_Datos!gttemphasta))
        Else
            Call insertarDatos("gti_detturtemp", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de parte de asignacion horaria"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    '*******************************************************************************************************************
    'Buscamos el procesamiento de empleados
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de procesamiento de empleados"
    progreso = progreso + porc
    StrSql = " SELECT * FROM gti_proc_emp WHERE " & _
             " fecha >= " & cambiaFecha(fechaDesde) & " AND fecha <= " & cambiaFecha(fechaHasta) & _
             " ORDER BY ternro "

    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = " SELECT ternro FROM gti_proc_emp WHERE ternro = " & rs_Datos!ternro & " AND fecha = " & cambiaFecha(rs_Datos!Fecha)
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_proc_emp", rs_Datos, "ternro = " & rs_Datos!ternro & " AND fecha = " & cambiaFecha(rs_Datos!Fecha))
        Else
            Call insertarDatos("gti_proc_emp", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de procesamiento de empleados"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos las licencias de los empleados
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de licencias de los empleados"
    progreso = progreso + porc
    StrSql = " SELECT * FROM emp_lic WHERE " & _
             " (elfechadesde <= " & cambiaFecha(fechaDesde) & " AND (elfechahasta >= " & cambiaFecha(fechaHasta) & " or elfechahasta >= " & cambiaFecha(fechaDesde) & ")) " & _
             " OR (elfechadesde >= " & cambiaFecha(fechaDesde) & " AND (elfechadesde <= " & cambiaFecha(fechaHasta) & ")) " & _
             " ORDER BY empleado "

    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = " SELECT emp_licnro FROM emp_lic WHERE emp_licnro = " & rs_Datos!emp_licnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("emp_lic", rs_Datos, "emp_licnro = " & rs_Datos!emp_licnro)
        Else
            Call insertarDatos("emp_lic", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de licencias de los empleados"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos las novedades de los empleados
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de novedades de los empleados"
    progreso = progreso + porc
    StrSql = " SELECT * FROM gti_novedad WHERE " & _
             " (gnovdesde <= " & cambiaFecha(fechaDesde) & " AND (gnovhasta >= " & cambiaFecha(fechaHasta) & " or gnovhasta >= " & cambiaFecha(fechaDesde) & ")) " & _
             " OR (gnovdesde >= " & cambiaFecha(fechaDesde) & " AND (gnovdesde <= " & cambiaFecha(fechaHasta) & ")) "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = " SELECT gnovnro FROM gti_novedad WHERE gnovnro = " & rs_Datos!gnovnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("gti_novedad", rs_Datos, "gnovnro = " & rs_Datos!gnovnro)
        Else
            Call insertarDatos("gti_novedad", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de novedades de los empleados"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    GoTo datosOk
CE:
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 406. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Sub
datosOk:
   ' Call sincronizar_Det(ternro, 1005)
    Set rs_Datos = Nothing
    Set rs_datosSinc = Nothing
End Sub

Sub expModelo407(ByVal progreso As Double, ByVal bpronro As Long)
Dim rs_Datos As New ADODB.Recordset
Dim rs_datosSinc As New ADODB.Recordset
Dim strConexionExt As String
Dim porc As Double
Dim bpronroReducido As String

 
    On Error GoTo CE
    Flog.writeline Espacios(Tabulador * 0) & "Ingreso al modelo 407"
    
    'cantidad de secciones de actualizacion
    porc = CLng(100 / 9)
        
    '*******************************************************************************************************************
    'Busco las estructuras puesto del sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de Estructura puesto"
    progreso = (100 - progreso) + porc
    StrSql = " SELECT * FROM estructura WHERE tenro = 4 "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT estrnro FROM estructura WHERE estrnro = " & rs_Datos!estrnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("estructura", rs_Datos, "estrnro = " & rs_Datos!estrnro)
        Else
            Call insertarDatos("estructura", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de puestos"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    '*******************************************************************************************************************
    'Buscamos los bacth procesos
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de batch procesos"
    progreso = progreso + porc
    'Busco el ultimo bpronro que existe en el sistema reducido
    StrSql = " SELECT top 1 bpronro FROM batch_proceso WHERE btprcnro = 279 ORDER BY bpronro DESC "
    OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
    bpronroReducido = 0
    If Not rs_datosSinc.EOF Then
        bpronroReducido = rs_datosSinc!bpronro
    End If
    
    '*******************************************************************************************************************
    'Buscamos los batch_procesos mayores al ultimo del sistema reducido
    StrSql = " SELECT * FROM batch_proceso WHERE btprcnro = 279 AND bpronro > " & bpronroReducido
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'No es necesario actulizar nada porque no existen
        Call insertarDatos("batch_proceso", rs_Datos)
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de batch procesos"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    '*******************************************************************************************************************
    'Buscamos los datos de la tabla periodo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de periodos"
    progreso = progreso + porc
    StrSql = " SELECT * FROM periodo "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT pliqnro FROM periodo WHERE pliqnro = " & rs_Datos!pliqnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("periodo", rs_Datos, "pliqnro = " & rs_Datos!pliqnro)
        Else
            Call insertarDatos("periodo", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de periodos"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    '*******************************************************************************************************************
    'Buscamos el ultimo bpronro de la tabla rep_det_irpf_cab en el sistema reducido
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de rep_det_irpf_cab"
    StrSql = " SELECT top 1 bpronro FROM rep_det_irpf_cab ORDER BY bpronro DESC "
    OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
    bpronroReducido = 0
    If Not rs_datosSinc.EOF Then
        bpronroReducido = rs_datosSinc!bpronro
    End If
        
    'Buscamos los procesos mas nuevos de la tabla rep_det_irpf_cab en el sistema completo
    progreso = progreso + porc
    StrSql = " SELECT * FROM rep_det_irpf_cab WHERE bpronro > " & bpronroReducido
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'No es necesario actulizar nada porque no existen
        Call insertarDatos("rep_det_irpf_cab", rs_Datos)
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de rep_det_irpf_cab "
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    
    '*******************************************************************************************************************
    'Buscamos el ultimo bpronro de la tabla rep_det_IRPF_det en el sistema reducido
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de rep_det_IRPF_det"
    StrSql = " SELECT top 1 bpronro FROM rep_det_IRPF_det ORDER BY bpronro DESC "
    OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
    bpronroReducido = 0
    If Not rs_datosSinc.EOF Then
        bpronroReducido = rs_datosSinc!bpronro
    End If
        
    'Buscamos los procesos mas nuevos de la tabla rep_det_IRPF_det en el sistema completo
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de rep_det_IRPF_det"
    progreso = progreso + porc
    StrSql = " SELECT * FROM rep_det_IRPF_det WHERE bpronro > " & bpronroReducido
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'No es necesario actulizar nada porque no existen
        Call insertarDatos("rep_det_IRPF_det", rs_Datos)
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de rep_det_IRPF_det "
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
           
    '*******************************************************************************************************************
    'Buscamos las escalas configuradas uru_irpfcab
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de uru_irpfcab"
    progreso = progreso + porc
    StrSql = " SELECT * FROM uru_irpfcab "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT cabnro FROM uru_irpfcab WHERE cabnro = " & rs_Datos!cabnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("uru_irpfcab", rs_Datos, "cabnro = " & rs_Datos!cabnro)
        Else
            Call insertarDatos("uru_irpfcab", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de uru_irpfcab"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos las escalas configuradas uru_irpfdet
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de uru_irpfdet"
    progreso = progreso + porc
    StrSql = " SELECT * FROM uru_irpfdet "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT detnro FROM uru_irpfdet WHERE detnro = " & rs_Datos!detnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("uru_irpfdet", rs_Datos, "detnro = " & rs_Datos!detnro)
        Else
            Call insertarDatos("uru_irpfdet", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de uru_irpfdet"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos las escalas configuradas uru_irpfdedcab
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de uru_irpfdedcab"
    progreso = progreso + porc
    StrSql = " SELECT * FROM uru_irpfdedcab "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT cabnro FROM uru_irpfdedcab WHERE cabnro = " & rs_Datos!cabnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("uru_irpfdedcab", rs_Datos, "cabnro = " & rs_Datos!cabnro)
        Else
            Call insertarDatos("uru_irpfdedcab", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de uru_irpfdedcab"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    '*******************************************************************************************************************
    'Buscamos las escalas configuradas uru_irpfdeddet
    Flog.writeline Espacios(Tabulador * 0) & "Actualizacion de uru_irpfdeddet"
    progreso = progreso + porc
    StrSql = " SELECT * FROM uru_irpfdeddet "
    OpenRecordset StrSql, rs_Datos
    Do While Not rs_Datos.EOF
        'controlo que exista o no en el sistema reducido
        StrSql = "SELECT detnro FROM uru_irpfdeddet WHERE detnro = " & rs_Datos!detnro
        OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
        If Not rs_datosSinc.EOF Then
            Call actualizarRegistro("uru_irpfdeddet", rs_Datos, "detnro = " & rs_Datos!detnro)
        Else
            Call insertarDatos("uru_irpfdeddet", rs_Datos)
        End If
        rs_Datos.MoveNext
    Loop
    Flog.writeline Espacios(Tabulador * 0) & "Fin Actualizacion de uru_irpfdeddet"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    GoTo datosOk
CE:
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Flog.writeline Espacios(Tabulador * 0) & "Error al tratar de recuperar los datos del modelo 407. "
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "__________________________________________________________"
    Exit Sub
datosOk:
    Set rs_Datos = Nothing
    Set rs_datosSinc = Nothing
End Sub

Sub insertarDatos(ByVal tabla As String, ByVal rs_Datos As ADODB.Recordset)
 Dim indice As Integer
 Dim NuevoValor As String
    
    'armo el query para insertar los datos
    StrSql = " INSERT INTO " & tabla & " ("
    For indice = 0 To rs_Datos.Fields.Count - 1
         StrSql = StrSql & rs_Datos(indice).Name & ","
    Next
    'quito la ultima ","
    StrSql = Left(StrSql, Len(StrSql) - 1)
    StrSql = StrSql & ") VALUES ("
    
    For indice = 0 To rs_Datos.Fields.Count - 1
                                 
        Select Case VarType(rs_Datos.Fields(rs_Datos(indice).Name))
            Case 8: 'tipo cadena
                NuevoValor = "'" & rs_Datos(indice) & "'"
            Case 7: 'tipo fecha
                NuevoValor = cambiaFecha(rs_Datos(indice))
            Case Else: 'cualquier otro tipo
                NuevoValor = IIf(IsNull(rs_Datos(indice)), "NULL", rs_Datos(indice))
        End Select
    
         
         StrSql = StrSql & NuevoValor & ","
    Next
    
    'quito la ultima ","
    StrSql = Left(StrSql, Len(StrSql) - 1)
    StrSql = StrSql & ")"
    ExtConn.Execute StrSql, , adExecuteNoRecords
    
End Sub
Sub actualizarRegistro(ByVal tabla As String, ByVal rs_Datos As ADODB.Recordset, ByVal clausulaWhere As String)
 Dim indice As Integer
 Dim NuevoValor As String
 
 
    StrSql = " UPDATE " & tabla & " SET "
    For indice = 0 To rs_Datos.Fields.Count - 1
        StrSql = StrSql & rs_Datos(indice).Name & " = "
         
        Select Case VarType(rs_Datos.Fields(rs_Datos(indice).Name))
            Case 8: 'tipo cadena
                NuevoValor = "'" & rs_Datos(indice) & "'"
            Case 7: 'tipo fecha
                NuevoValor = cambiaFecha(rs_Datos(indice))
            Case Else: 'cualquier otro tipo
                NuevoValor = IIf(IsNull(rs_Datos(indice)), "NULL", rs_Datos(indice))
        End Select
        StrSql = StrSql & NuevoValor & ", "
    Next
    
    'quito la ultima ", "
    StrSql = Left(StrSql, Len(StrSql) - 2)
        
    StrSql = StrSql & " WHERE " & clausulaWhere

    ExtConn.Execute StrSql, , adExecuteNoRecords
End Sub
Sub recibosAprobados(ByVal rs_Datos As ADODB.Recordset)
Dim rs_datosSinc As New ADODB.Recordset
Dim rs_datosSinc2 As New ADODB.Recordset
   
    'Busco los datos del lado de la base reducida
    StrSql = " SELECT aprobado FROM gesthistoliq WHERE bpronro = " & rs_Datos!bpronro
    OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
    If rs_datosSinc.EOF Then
        'Si no hay datos tengo que insertar todo lo que encuentre
        StrSql = " SELECT * FROM gesthistoliq WHERE bpronro = " & rs_Datos!bpronro
        OpenRecordset StrSql, rs_datosSinc2
        Do While Not rs_datosSinc2.EOF
            Call insertarDatos("gesthistoliq", rs_datosSinc2)
            rs_datosSinc2.MoveNext
        Loop
        
    Else
        'Existen los registros tengo que actualizar el estado
        StrSql = " UPDATE gesthistoliq SET aprobado = " & rs_Datos!aprobado & " WHERE bpronro = " & rs_Datos!bpronro
        ExtConn.Execute StrSql, , adExecuteNoRecords
    End If
    
Set rs_datosSinc = Nothing
Set rs_datosSinc2 = Nothing

End Sub

Sub recibosAprobadosUsuario(ByVal rs_Datos As ADODB.Recordset)
Dim rs_datosSinc As New ADODB.Recordset
    'Busco los datos del lado de la base reducida
    StrSql = " SELECT estado, fechaestado, obs FROM rep_recibo_digital WHERE cliqnro = " & rs_Datos!cliqnro
    OpenRecordsetExt StrSql, rs_datosSinc, ExtConn
    If rs_datosSinc.EOF Then
        'No existen datos se insertan del lado del sistema
        Call insertarDatos("rep_recibo_digital", rs_Datos)
    Else
        'Si existe lo actualizo
        StrSql = " UPDATE rep_recibo_digital SET " & _
                 " estado = " & rs_Datos!estado & ", fechaestado = '" & rs_Datos!fechaestado & "', obs = '" & rs_Datos!Obs & "'" & _
                 " WHERE cliqnro = " & rs_Datos!cliqnro
        ExtConn.Execute StrSql, , adExecuteNoRecords
    End If
    
Set rs_datosSinc = Nothing

End Sub
