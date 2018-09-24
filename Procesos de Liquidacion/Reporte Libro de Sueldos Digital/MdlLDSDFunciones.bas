Attribute VB_Name = "MdlLDSDFunciones"

Function getTerDoc(ByVal Ternro As Long, tidnro As Long, Tipodato As Integer, CodChar As Integer) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna Datos de los documentos
' Autor      : Gonzalez Nicolás
' Fecha      : 17/06/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    'Retorna SIGLA o Número de Documento S/parametro
    StrSql = "SELECT tipodocu.tidnro,tipodocu.tidnom,tipodocu.tidsigla,institucion.instdes,institucion.instabre,ter_doc.nrodoc  "
    StrSql = StrSql & " FROM  tipodocu "
    StrSql = StrSql & " LEFT JOIN institucion ON institucion.instnro = tipodocu.instnro "
    StrSql = StrSql & " LEFT JOIN ter_doc ON ter_doc.tidnro = tipodocu.tidnro"
    StrSql = StrSql & " WHERE tipodocu.tidnro =" & tidnro
    StrSql = StrSql & " AND ter_doc.ternro = " & Ternro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Select Case Tipodato
            Case 0: 'tidsigla
                getTerDoc = nullToString(rs("tidsigla"))
            Case 1: 'nrodoc
                getTerDoc = nullToString(rs("nrodoc"))
                If CodChar = 1 Then
                    getTerDoc = Replace(getTerDoc, "-", "")
                End If
            Case 2: 'instdes
                getTerDoc = nullToString(rs("instdes"))
            Case 3: 'instabre
                getTerDoc = nullToString(rs("instabre"))
        End Select
        
    Else
        getTerDoc = ""
    
    End If
    rs.Close
End Function
Public Function getZona(tipo, Fecha_Fin_Periodo)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna Datos de LA ZONA
' Autor      : Gonzalez Nicolás
' Fecha      : 24/07/2015
' Ultima Mod.: 16/10/2015 - NG - Para 1y2 se usa ternro de sucursal y empresa
' Descripcion:
' ----------------------------------------------------------------

    Dim rs_sucursal As New ADODB.Recordset
    Dim rs_zona As New ADODB.Recordset
    Dim Continuar
    Dim Aux_Zona
    'Dim Aux_Localidad
    Continuar = True
    Aux_Zona = "00"
    'De acuerdo a la opcion busco la zona
    If Continuar = True Then
        Select Case CInt(tipo)
        Case 1: 'Sucursal
            'Cargo el tipo de estructura segun sea sucursal
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
            StrSql = StrSql & "  tenro = 1 AND "
            StrSql = StrSql & " (htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
            StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                StrSql = " SELECT ternro,Estrnro FROM sucursal "
                StrSql = StrSql & " WHERE estrnro =" & rs_Estructura!Estrnro
                OpenRecordset StrSql, rs_sucursal
                If Not rs_sucursal.EOF Then
                    StrSql = " SELECT zona.zonanro,zona.zonacod,locdesc,provdesc FROM detdom "
                    StrSql = StrSql & " INNER JOIN zona ON zona.zonanro = detdom.zonanro "
                    StrSql = StrSql & " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro "
                    StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
                    StrSql = StrSql & " LEFT JOIN provincia ON provincia.provnro = detdom.provnro "
                    StrSql = StrSql & " WHERE cabdom.ternro = " & rs_sucursal!Ternro
                    OpenRecordset StrSql, rs_zona
                    If Not rs_zona.EOF Then
                        Aux_Zona = Left(CStr(IIf(Not EsNulo(rs_zona!zonacod), rs_zona!zonacod, "00")), 2)
                    End If
                End If  ' If Not rs_Sucursal.EOF Then
            End If ' If Not rs_Estructura.EOF Then
        Case 2: 'Empresa
            'Cargo el tipo de estructura segun sea sucursal o Empresa
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
            StrSql = StrSql & " tenro = 10 AND "
            StrSql = StrSql & " (htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
            StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                StrSql = " SELECT ternro FROM empresa "
                StrSql = StrSql & " WHERE estrnro =" & rs_Estructura!Estrnro
                OpenRecordset StrSql, rs_sucursal
                If Not rs_sucursal.EOF Then
                    StrSql = " SELECT zona.zonanro,zona.zonacod,locdesc,provdesc FROM detdom "
                    StrSql = StrSql & " INNER JOIN zona ON zona.zonanro = detdom.zonanro "
                    StrSql = StrSql & " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro "
                    StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
                    StrSql = StrSql & " LEFT JOIN provincia ON provincia.provnro = detdom.provnro "
                    StrSql = StrSql & " WHERE cabdom.ternro = " & rs_sucursal!Ternro
                    OpenRecordset StrSql, rs_zona
                    If Not rs_zona.EOF Then
                        Aux_Zona = Left(CStr(IIf(Not EsNulo(rs_zona!zonacod), rs_zona!zonacod, "00")), 2)
                    End If
                End If  ' If Not rs_Sucursal.EOF Then
            End If ' If Not rs_Estructura.EOF Then
        Case 3: 'Empleado
            StrSql = " SELECT zona.zonanro,zona.zonacod,locdesc,provdesc FROM detdom "
            StrSql = StrSql & " INNER JOIN zona ON zona.zonanro = detdom.zonanro "
            StrSql = StrSql & " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro "
            StrSql = StrSql & "INNER JOIN localidad ON localidad.locnro = detdom.locnro "
            StrSql = StrSql & " LEFT JOIN provincia ON provincia.provnro = detdom.provnro "
            StrSql = StrSql & " WHERE cabdom.ternro = " & Ternro
            OpenRecordset StrSql, rs_zona
            If Not rs_zona.EOF Then
                Aux_Zona = Left(CStr(IIf(Not IsNull(rs_zona!zonacod), rs_zona!zonacod, "00")), 2)
            End If
        Case 4: 'Laboral
            StrSql = " SELECT zona.zonanro,zona.zonacod,locdesc,provdesc FROM detdom "
            StrSql = StrSql & " INNER JOIN zona ON zona.zonanro = detdom.zonanro "
            StrSql = StrSql & " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro "
            StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
            StrSql = StrSql & " LEFT JOIN provincia ON provincia.provnro = detdom.provnro "
            StrSql = StrSql & " WHERE cabdom.ternro = " & Ternro
            OpenRecordset StrSql, rs_zona
            If Not rs_zona.EOF Then
                Aux_Zona = Left(CStr(IIf(Not EsNulo(rs_zona!zonacod), rs_zona!zonacod, "00")), 2)
            End If
        End Select
        getZona = Format(Aux_Zona, "00")
    End If
End Function
Public Function Buscar_SituacionRevistaConfig(confval, confval2) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: SI ENCUENTRA AL MENOS UNA CONFIGURADA RETORNA CONFVAL2 DEL CONFREP
' Autor      : Gonzalez Nicolás
' Fecha      : 22/07/2015
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
    Dim rs_aux As New ADODB.Recordset
    Dim codigoExt As String
    codigoExt = "0"
    If confval2 <> "" Then 'Solo se controla si esta configurado x confrep
        StrSql = "SELECT estrnro FROM his_estructura "
        StrSql = StrSql & " WHERE ternro= " & Ternro & " AND his_estructura.estrnro IN(" & confval & ")"
        StrSql = StrSql & " AND (htethasta< " & ConvFecha(DatosPeriodo.pliqhasta) & " OR htethasta IS NULL) ORDER BY htethasta,htetdesde DESC"
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            codigoExt = confval2
        End If
    End If
    Buscar_SituacionRevistaConfig = codigoExt
End Function
Public Function getContratoActual(ByVal Tcodnro As String, ByVal Fecha As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna Código del contrato actual
' Autor      : Gonzalez Nicolás
' Fecha      : 23/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    StrSql = " SELECT estrnro FROM his_estructura "
    StrSql = StrSql & " WHERE ternro = " & Ternro
    StrSql = StrSql & " AND tenro = 18  "
    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Fecha) & ")  "
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= htethasta) or (htethasta is null))"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs!Estrnro
        StrSql = StrSql & " AND tcodnro = " & Tcodnro
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            getContratoActual = Left(CStr(rs_Estr_cod!nrocod), 3)
        Else
            'Flog.Writeline "No se encontró el codigo interno para el Tipo de Contrato."
            getContratoActual = "-1"
        End If
    Else
        'Flog.Writeline "No se encontró el Tipo de Contrato."
        getContratoActual = "0"
    End If
End Function
Public Function getCodSiniestrado(ByVal Tcodnro As String, ByVal Tenro As Integer, ByVal Fecha As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna Código de Siniestrado
' Autor      : Gonzalez Nicolás
' Fecha      : 16/10/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    StrSql = " SELECT estrnro FROM his_estructura "
    StrSql = StrSql & " WHERE ternro = " & Ternro
    StrSql = StrSql & " AND tenro = " & Tenro
    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Fecha) & ")  "
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= htethasta) or (htethasta is null))"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs!Estrnro
        StrSql = StrSql & " AND tcodnro = " & Tcodnro
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            getCodSiniestrado = Left(CStr(rs_Estr_cod!nrocod), 2)
        Else
            getCodSiniestrado = "-1"
        End If
    Else
        getCodSiniestrado = "00"
    End If
End Function
Public Function getCodOS(ByVal Tcodnro As String, ByVal Fecha As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna Código de Obra Social
' Autor      : Gonzalez Nicolás
' Fecha      : 24/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
StrSql = " SELECT his_estructura.estrnro FROM his_estructura "
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro
StrSql = StrSql & " AND his_estructura.tenro = 17  "
StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha) & ")  "
StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
    StrSql = StrSql & " AND tcodnro = " & Tcodnro
    OpenRecordset StrSql, rs_Estr_cod
    If Not rs_Estr_cod.EOF Then
        getCodOS = Left(rs_Estr_cod!nrocod, 6)
    Else
        getCodOS = "-1"
    End If
Else
    getCodOS = "0"
End If

End Function
Public Function getCondicionSIJP(ByVal Tcodnro As String, ByVal Fecha As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna Codigo de SIJP
' Autor      : Gonzalez Nicolás
' Fecha      : 23/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    'Flog.Writeline "Buscar la condicion"
    StrSql = " SELECT his_estructura.estrnro FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro
    StrSql = StrSql & "  AND his_estructura.tenro = 31 AND "
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs!Estrnro
        StrSql = StrSql & " AND tcodnro = " & Tcodnro
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            getCondicionSIJP = Left(CStr(rs_Estr_cod!nrocod), 2)
        Else
            'Flog.Writeline "No se encontró el codigo interno para la Condicion de SIJP"
            getCondicionSIJP = "-1"
        End If
    Else
        'Flog.Writeline "No se encontro la Condicion de SIJP"
        getCondicionSIJP = "0"
    End If
End Function
Public Function getActividad(ByVal Tcodnro As String, ByVal Fecha As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna Codigo de Actividad
' Autor      : Gonzalez Nicolás
' Fecha      : 23/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    StrSql = " SELECT his_estructura.estrnro FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro
    StrSql = StrSql & " AND his_estructura.tenro = 29 "
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha) & ") "
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & Tcodnro
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            getActividad = Left(rs_Estr_cod!nrocod, 2)
        Else
            getActividad = "-1"
        End If
    Else
        getActividad = "0"
    End If
End Function
Public Function getEstructura(ByVal Tcodnro As String, ByVal Tenro As Long, ByVal Fecha As Date, ByVal Salida As String, ByVal Escape As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna Codigo externo de una estructura
' Autor      : Gonzalez Nicolás
' Fecha      : 23/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    StrSql = " SELECT his_estructura.estrnro,estructura.estrcodext FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.ternro = " & Ternro
    StrSql = StrSql & " AND his_estructura.tenro = " & Tenro
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha) & ") "
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        If Tcodnro <> "0" Then 'BUSCO POR TIPO DE CODIGO
            StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
            StrSql = StrSql & " AND tcodnro = " & Tcodnro
            OpenRecordset StrSql, rs_Estr_cod
            If Not rs_Estr_cod.EOF Then
                getEstructura = IIf(EsNulo(rs_Estr_cod!nrocod), Escape, rs_Estr_cod!nrocod)
            Else
                getEstructura = Escape
            End If
        Else
            Select Case UCase(Salida)
                Case "ESTRCODEXT":
                    getEstructura = IIf(EsNulo(rs_Estructura!estrcodext), Escape, rs_Estructura!estrcodext)
            End Select
        End If
    Else
        getEstructura = Escape
    End If
End Function
Function nullToString(Texto As Variant) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: CONTROL DE NULOS
' Autor      : Gonzalez Nicolás
' Fecha      : 17/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    If EsNulo(Texto) Then
        nullToString = ""
    Else
        nullToString = Texto
    End If

End Function

Function getEmpresaTipoEmp()
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna codigo DGI asociado a una empresa
' Autor      : Gonzalez Nicolás
' Fecha      : 18/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
StrSql = "SELECT tipempcoddgi FROM empresa "
StrSql = StrSql & " INNER JOIN tipempdor ON empresa.tipempnro = tipempdor.tipempnro"
StrSql = StrSql & " WHERE empresa.Estrnro = " & ListParam.Empresa
OpenRecordset StrSql, rs
If Not rs.EOF Then
    getEmpresaTipoEmp = nullToString(rs("tipempcoddgi"))
Else
    getEmpresaTipoEmp = ""
End If
rs.Close
    
End Function

Public Function Calcular_Edad(ByVal Fecha As Date) As Integer
'...........................................................................
' Archivo       : edad.i                              fecha ini. : 20/01/92
' Autor         : FGZ
' Fecha         :
' Ultima Mod    : 25/07/2005 - Se calcula a la fecha fin del periodo.
'...........................................................................
Dim años  As Integer
Dim ALaFecha As Date

    ALaFecha = C_Date(Date)

    años = Year(ALaFecha) - Year(Fecha)
    If Month(ALaFecha) < Month(Fecha) Then
       años = años - 1
    Else
        If Month(ALaFecha) = Month(Fecha) Then
            If Day(ALaFecha) < Day(Fecha) Then
                años = años - 1
            End If
        End If
    End If
    Calcular_Edad = años
End Function

Function getValoresLiq(ByVal busEtiq As String, ByVal Ternro As Long, ByVal codigoConAcu As String, ByVal mesLiq As Integer, ByVal anioLiq As Long, ByVal cliqnro As String, ByVal Cantidad As Long)
'Busca un Concepto, Acumulador Mesual o Acumulador Liquidacion segun la etiqueta configurada en el confRep.
'(Concepto -> COC | COM) -- (Acumulador Mensual -> ACM | ACC) -- (Acumulador de Liquidacion -> ALC | ALM)
 Dim rsValorLiq As New ADODB.Recordset
    
    Select Case busEtiq
        Case "COC", "CO" '-------------------------------------------------------------------------------------------------------------
            'Busco todos lod detliq entre los meses
            StrSql = "SELECT sum(detliq.dlicant) dlicant, sum(detliq.dlimonto) dlimonto  FROM periodo " & _
                    " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
                    " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro AND cabliq.empleado = " & Ternro & _
                    " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
                    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concnro IN('" & codigoConAcu & "')" & _
                    " WHERE periodo.pliqmes= " & mesLiq & "  AND periodo.pliqanio= " & anioLiq & _
                    " AND cabliq.cliqnro IN(" & cliqnro & ")"
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "COC" Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!dlicant), 0, rsValorLiq!dlicant)
                Else
                    getValoresLiq = IIf(EsNulo(rsValorLiq!dlimonto), 0, rsValorLiq!dlimonto)
                End If
            Else
                getValoresLiq = 0
            End If
        
        Case "ACC", "ACM" '-------------------------------------------------------------------------------------------------------------
            StrSql = "SELECT SUM(ammonto) ammonto, SUM(amcant) amcant FROM acu_mes WHERE acu_mes.ammes= " & mesLiq & "  AND acu_mes.amanio= " & anioLiq & " AND ternro = " & Ternro & _
                    " AND acunro IN  (" & codigoConAcu & ")"
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "ACC" Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!amcant), 0, rsValorLiq!amcant)
                Else
                    getValoresLiq = IIf(EsNulo(rsValorLiq!ammonto), 0, rsValorLiq!ammonto)
                End If
            Else
                getValoresLiq = 0
            End If
        
        Case "ACL", "AC" '-------------------------------------------------------------------------------------------------------------
            StrSql = "SELECT sum(acu_liq.alcant) alcant, sum(acu_liq.almonto) almonto FROM periodo " & _
                    " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
                    " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
                    " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
                    " WHERE cabliq.empleado = " & Ternro & _
                    " AND periodo.pliqmes = " & mesLiq & _
                    " AND periodo.pliqanio = " & anioLiq & _
                    " AND acu_liq.acunro IN (" & codigoConAcu & ")" & _
                    " AND cabliq.cliqnro IN(" & cliqnro & ")"
            ' " AND periodo.pliqmes >= " & (mesLiq - 12)
            OpenRecordset StrSql, rsValorLiq
            If Not rsValorLiq.EOF Then
                If busEtiq = "ACL" Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!alcant), 0, rsValorLiq!alcant)
                Else
                    getValoresLiq = IIf(EsNulo(rsValorLiq!almonto), 0, rsValorLiq!almonto)
                End If
            Else
                getValoresLiq = 0
            End If
       Case "IM": 'Imponibles
            StrSql = "SELECT * FROM impproarg WHERE cliqnro IN (" & cliqnro & ")"
            StrSql = StrSql & " AND acunro =" & codigoConAcu
            StrSql = StrSql & " AND cliqnro IN(" & cliqnro & ")"
            OpenRecordset StrSql, rsValorLiq
            If Not rsValorLiq.EOF Then
                If Cantidad = 6 Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!ipacant), 0, rsValorLiq!ipacant)
                ElseIf Cantidad = 7 Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!ipamonto), 0, rsValorLiq!ipamonto)
                Else
                    getValoresLiq = 0
                End If
            Else
                getValoresLiq = 0
            End If
    
        Case "I1": 'Imponible Sueldo
            StrSql = "SELECT ipacant,ipamonto FROM impproarg WHERE cliqnro IN (" & cliqnro & ")"
            StrSql = StrSql & " AND acunro =" & codigoConAcu
            StrSql = StrSql & " AND tconnro = 1"
            StrSql = StrSql & " AND cliqnro IN(" & cliqnro & ")"
            OpenRecordset StrSql, rsValorLiq
            If Not rsValorLiq.EOF Then
                 If Cantidad = 6 Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!ipacant), 0, rsValorLiq!ipacant)
                ElseIf Cantidad = 7 Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!ipamonto), 0, rsValorLiq!ipamonto)
                Else
                    getValoresLiq = 0
                End If
            Else
                getValoresLiq = 0
            End If
    
        Case "I2": 'Imponible LAR
            StrSql = "SELECT ipacant,ipamonto FROM impproarg WHERE cliqnro IN (" & cliqnro & ")"
            StrSql = StrSql & " AND acunro =" & codigoConAcu
            StrSql = StrSql & " AND tconnro = 2"
            StrSql = StrSql & " AND cliqnro IN(" & cliqnro & ")"
            OpenRecordset StrSql, rsValorLiq
            If Not rsValorLiq.EOF Then
                If Cantidad = 6 Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!ipacant), 0, rsValorLiq!ipacant)
                ElseIf Cantidad = 7 Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!ipamonto), 0, rsValorLiq!ipamonto)
                Else
                    getValoresLiq = 0
                End If
            Else
                getValoresLiq = 0
            End If
    
        Case "I3": 'Imponible SAC
            StrSql = "SELECT ipacant,ipamonto FROM impproarg WHERE cliqnro IN (" & cliqnro & ")"
            StrSql = StrSql & " AND acunro =" & codigoConAcu
            StrSql = StrSql & " AND tconnro = 3"
            StrSql = StrSql & " AND cliqnro IN(" & cliqnro & ")"
            OpenRecordset StrSql, rsValorLiq
            If Not rsValorLiq.EOF Then
                If Cantidad = 6 Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!ipacant), 0, rsValorLiq!ipacant)
                ElseIf Cantidad = 7 Then
                    getValoresLiq = IIf(EsNulo(rsValorLiq!ipamonto), 0, rsValorLiq!ipamonto)
                Else
                    getValoresLiq = 0
                End If
            Else
                getValoresLiq = 0
            End If

    End Select
    
End Function
Public Function Format_Data(ByVal Fecha As Date, ByVal Formato As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato de fecha para impresion
' Autor      : Gonzalez Nicolás
' Fecha      : 16/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    If IsDate(Fecha) Then
        Select Case Formato
            Case "AAAAMMDD"
                Format_Data = Year(Fecha)
                Format_Data = Format_Data & Format_StrLR(Month(Fecha), 2, "L", True, "0")
                Format_Data = Format_Data & Format_StrLR(Day(Fecha), 2, "L", True, "0")
            Case "AAAAMM"
                Format_Data = Year(Fecha)
                Format_Data = Format_Data & Format_StrLR(Month(Fecha), 2, "L", True, "0")
        End Select
    Else
      Format_Data = ""
    End If
End Function
Public Function getUnidadConc(tconnro)
    Dim aux
    Dim Arraux
    Dim Resultado
    Dim Encontre
    Encontre = False
    aux = confRep(24).confval
    Resultado = ""
    If aux <> "" Then
        If InStr(aux, ",") > 0 Then
           Arraux = Split(aux, ",")
           For a = 0 To UBound(Arraux)
                If CLng(Arraux(a)) = CLng(tconnro) Then
                    Resultado = "$"
                    Encontre = True
                    Exit For
                End If
           Next
        End If
    End If
        
    
    If Encontre = False Then
        aux = confRep(24).confval2
        If aux <> "" Then
            If InStr(aux, ",") > 0 Then
               Arraux = Split(aux, ",")
               For a = 0 To UBound(Arraux)
                    If CLng(Arraux(a)) = CLng(tconnro) Then
                        Resultado = "%"
                        Encontre = True
                        Exit For
                    End If
               Next
            End If
        End If
    End If
            
    If Encontre = False Then
        aux = confRep(24).confval3
        If aux <> "" Then
            If InStr(aux, ",") > 0 Then
               Arraux = Split(aux, ",")
               For a = 0 To UBound(Arraux)
                    If CLng(Arraux(a)) = CLng(tconnro) Then
                        Resultado = "A"
                        Encontre = True
                        Exit For
                    End If
               Next
            End If
        End If
    End If
    
    If Encontre = False Then
        aux = confRep(24).confval4
        If aux <> "" Then
            If InStr(aux, ",") > 0 Then
               Arraux = Split(aux, ",")
               For a = 0 To UBound(Arraux)
                    If CLng(Arraux(a)) = CLng(tconnro) Then
                        Resultado = "Q"
                        Encontre = True
                        Exit For
                    End If
               Next
            End If
        End If
    End If
            
    If Encontre = False Then
        aux = confRep(25).confval
        If aux <> "" Then
            If InStr(aux, ",") > 0 Then
               Arraux = Split(aux, ",")
               For a = 0 To UBound(Arraux)
                    If CLng(Arraux(a)) = CLng(tconnro) Then
                        Resultado = "M"
                        Encontre = True
                        Exit For
                    End If
               Next
            End If
        End If
    End If
    
    If Encontre = False Then
        aux = confRep(25).confval2
        If aux <> "" Then
            If InStr(aux, ",") > 0 Then
               Arraux = Split(aux, ",")
               For a = 0 To UBound(Arraux)
                    If CLng(Arraux(a)) = CLng(tconnro) Then
                        Resultado = "D"
                        Encontre = True
                        Exit For
                    End If
               Next
            End If
        End If
    End If
            
    If Encontre = False Then
        aux = confRep(25).confval3
        If aux <> "" Then
            If InStr(aux, ",") > 0 Then
               Arraux = Split(aux, ",")
               For a = 0 To UBound(Arraux)
                    If CLng(Arraux(a)) = CLng(tconnro) Then
                        Resultado = "H"
                        Encontre = True
                        Exit For
                    End If
               Next
            End If
        End If
    End If
    getUnidadConc = Resultado
        
End Function
Function getFamiliar(ByVal Ternro As Long, ByVal parenro As Long, ByVal Tipodato As Long, ByVal Escape As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna datos de un familiar
' Autor      : Gonzalez Nicolás
' Fecha      : 18/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Dim NomComp As String
    Dim contador As Long
    StrSql = "SELECT tercero.terape,tercero.terape2,tercero.ternom,tercero.ternom2 "
    StrSql = StrSql & " ,tercero.nacionalnro "
    StrSql = StrSql & " FROM familiar"
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = familiar.ternro"
    StrSql = StrSql & " WHERE familiar.Empleado =  " & Ternro
    StrSql = StrSql & " AND familiar.parenro = " & parenro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Select Case Tipodato
            Case 0:
                getFamiliar = "1"
            Case 1:
                contador = 0
                Do While Not rs.EOF
                    contador = contador + 1
                    rs.MoveNext
                Loop
                getFamiliar = CStr(contador)
            'Case 0: 'NOMBRE COMPLETO
            '    getFamiliar = GetNombreFormat(rs!Terape, IIf(IsNull(rs!Terape2), "", rs!Terape2), rs!Ternom, IIf(IsNull(rs!Ternom2), "", rs!Ternom2), 60)
           'Case 1: 'NACIONALIDAD FAMILIAR
                'getFamiliar = getPaisNac(IIf(IsNull(rs!nacionalnro), 0, rs!nacionalnro), 1)
        End Select
    Else
        getFamiliar = Escape
    End If
End Function
Public Function Format_StrLR(ByVal Str, ByVal Longitud As Long, ByVal Posicion As String, ByVal Completar As Boolean, ByVal Str_Completar As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Restringe la cantidad de caracteres del string pasado como parametro y lo completa
'              con el caracter pasado por parametro a la izq/der hasta la longitud (si completar es TRUE)
' Autor      : Gonzalez Nicolás
' Fecha      : 17/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
If Not EsNulo(Str) Then
    Str = Left(Str, Longitud)
    If Completar Then
        If Len(Str) < Longitud Then
            If UCase(Posicion) = "R" Then
                Str = Str & String(Longitud - Len(Str), Str_Completar)
            Else
                Str = String(Longitud - Len(Str), Str_Completar) & Str
            End If
        End If
    End If
    'Corta el string según Tope
    Format_StrLR = UCase(Str)
Else
    If Completar Then
        Format_StrLR = String(Longitud, " ")
    Else
        Format_StrLR = ""
    End If
End If

End Function
Function getEmpleadoEstr(Ternro, Tenro, Tipodato, Fdesde, FHasta, Escape)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna datos de una determinada estructura
' Autor      : Gonzalez Nicolás
' Fecha      : 20/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    StrSql = "SELECT estructura.estrnro,estructura.estrdabr,estructura.estrdext,estructura.estrcodext  FROM his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON  estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " WHERE his_estructura.Ternro = " & Ternro
    StrSql = StrSql & " AND his_estructura.Tenro = " & Tenro
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(FHasta) & " AND (his_estructura.htethasta >=" & ConvFecha(Fdesde) & " OR his_estructura.htethasta IS NULL))"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Select Case Tipodato
            Case "ESTRNRO"
                getEmpleadoEstr = IIf(IsNull(rs!Estrnro), "0", rs!Estrnro)
            Case "ESTRDABR"
                getEmpleadoEstr = IIf(EsNulo(rs!estrdabr), Escape, rs!estrdabr)
            Case "ESTRDEXT"
                getEmpleadoEstr = IIf(EsNulo(rs!estrdext), Escape, rs!estrdext)
            Case "ESTRCODEXT"
                getEmpleadoEstr = IIf(EsNulo(rs!estrcodext), Escape, rs!estrcodext)
            Case "HTETDESDE":
                getEmpleadoEstr = IIf(EsNulo(rs!htetdesde), Escape, rs!htetdesde)
        End Select
    Else
        'NO HAY DATOS
        getEmpleadoEstr = Escape
    End If
    rs.Close
End Function

Function getDatosPedidoPag(Ternro, Pliqnro, cliqnro, Tipodato, Escape)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna datos del Pedido de Pago
' Autor      : Gonzalez Nicolás
' Fecha      : 20/07/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    StrSql = "SELECT pago.pagomonto, pago.ctabcbu, pago.ctabnro, pago.pagosec, estructura.estrdabr, proceso.prodesc ,pedidopago.ppagfecped"
    StrSql = StrSql & ",formapago.fpagsigla,formapago.fpagdesext"
    StrSql = StrSql & " FROM pedidopago"
    StrSql = StrSql & " INNER JOIN pago ON pedidopago.ppagnro = pago.ppagnro"
    StrSql = StrSql & " INNER JOIN cabliq ON pago.pagorigen = cabliq.cliqnro"
    StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro"
    StrSql = StrSql & " LEFT JOIN banco ON pedidopago.bannro = banco.ternro"
    StrSql = StrSql & " LEFT JOIN estructura ON estructura.estrnro = banco.estrnro"
    StrSql = StrSql & " LEFT JOIN formapago ON formapago.fpagnro = pago.fpagnro"
    StrSql = StrSql & " WHERE pedidopago.tppanro IN(1,2,3) "
    StrSql = StrSql & " AND pago.pagorigen = " & cliqnro
    StrSql = StrSql & " AND pedidopago.pliqnro = " & Pliqnro
    StrSql = StrSql & " AND pago.ternro = " & Ternro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Select Case Tipodato
            Case "CTACBU"
                getDatosPedidoPag = IIf(IsNull(rs!ctabcbu), Escape, rs!ctabcbu)
            Case "SIGLA"
                getDatosPedidoPag = IIf(IsNull(rs!fpagsigla), Escape, rs!fpagsigla)
            Case "DESEXT"
                getDatosPedidoPag = IIf(IsNull(rs!fpagdesext), Escape, rs!fpagdesext)
            Case "FECPED"
                getDatosPedidoPag = IIf(IsNull(rs!ppagfecped), Escape, rs!ppagfecped)
        End Select
    Else
        'NO HAY DATOS
        getDatosPedidoPag = Escape
    End If
    rs.Close
End Function
Public Function getCodMiSimpl(ByVal Ternro As Long, ByVal Tcodnro As String, ByVal Tenro As Long, ByVal Fdesde As Date, ByVal FHasta As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna Codigo de MI Simplifación
' Autor      : Gonzalez Nicolás
' Fecha      : 19/10/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    StrSql = "SELECT estructura.estrnro,estructura.estrdabr,estructura.estrdext,estructura.estrcodext  FROM his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON  estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " WHERE his_estructura.Ternro = " & Ternro
    StrSql = StrSql & " AND his_estructura.Tenro = " & Tenro
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(FHasta) & " AND (his_estructura.htethasta >=" & ConvFecha(Fdesde)
    StrSql = StrSql & " OR his_estructura.htethasta IS NULL))"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs!Estrnro
        StrSql = StrSql & " AND tcodnro = " & Tcodnro
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            getCodMiSimpl = CStr(rs_Estr_cod!nrocod)
        Else
            'No se encontró el codigo interno"
            getCodMiSimpl = "-1"
        End If
    Else
        'No se encontro estructura
        getCodMiSimpl = "0"
    End If
End Function
