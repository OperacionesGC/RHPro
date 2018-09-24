Attribute VB_Name = "repRecibos3"

Sub generarDatosRecibo500(Pronro, Ternro, acunroSueldo, tituloReporte, orden)
'Descripción : BOLETA DE PAGO (QUINCENAL) SYKES - SV
'Codigo de forma de liq quincenal
'Const quincenal1 = 553
'Const quincenal2 = 555

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim Nombre
'Dim Direccion
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
'Dim Cuil
Dim empFecAlta
'Dim EmpFecAltaRec
'Dim empFecAltarec2
Dim Sueldo
'Dim Categoria

Dim localidad
Dim proFecPago
Dim pliqhasta
'Dim FormaPago
'Dim Puesto

Dim EmpEstrnro
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String
Dim quincena As String

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset


On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

'-----------------------------------------------------------------------------------
'BUSCO CONFIGURACIÓN DE REPORTE PARA EL MODELO QUE SE PROCESA.
'-----------------------------------------------------------------------------------
' || Esta función se utiliza a partir de los modelos 500 ||
' --->  Devuelve Recordset
'-----------------------------------------------------------------------------------
'StrSql = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & tipoRecibo * 100
'StrSql = StrSql & " AND (confnrocol < " & (tipoRecibo * 100) + 100
'StrSql = StrSql & " ORDER BY confnrocol ASC "

'Dim rs_confrep  As New ADODB.Recordset
Dim EmpresaTenro As Long
Dim UnAdminTenro As Long
Dim CCostoTenro As Long
Dim PlazaTenro As Long
Dim IVMTenro As Long
Dim ListaQuin As String
Dim CentroCosto As String
Dim UnidadAdmin As String
Dim Plaza As String
Dim IVM As String
Dim AcuSueldoMen As Long

'Obtengo configuración de reporte
Dim rs_GetConfrepRs As New ADODB.Recordset
StrSql = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & tipoRecibo * 100
StrSql = StrSql & " AND confnrocol < " & (tipoRecibo * 100) + 100 & ")"
StrSql = StrSql & " ORDER BY confnrocol ASC "
'GetConfrepRs tipoRecibo
OpenRecordset StrSql, rs_GetConfrepRs

If Not rs_GetConfrepRs.EOF Then
    Do While Not rs_GetConfrepRs.EOF
        Select Case rs_GetConfrepRs!confnrocol
            Case 50000: 'Empresa
                        EmpresaTenro = rs_GetConfrepRs!confval
            Case 50001: 'Unidad Administrativa
                        UnAdminTenro = rs_GetConfrepRs!confval
            Case 50002: 'Centro de Costo
                        CCostoTenro = rs_GetConfrepRs!confval
            Case 50003: 'Plaza
                        PlazaTenro = rs_GetConfrepRs!confval
            Case 50004: 'IVM
                        IVMTenro = rs_GetConfrepRs!confval
            Case 50005: 'Salario Mensual (AC)
                        AcuSueldoMen = rs_GetConfrepRs!confval
'            Case 50006: 'Quincena (lista de modelos)
'                    If EsNulo(rs_GetConfrepRs!confval2) Then
'                        ListaQuin = rs_GetConfrepRs!confval2
'                    Else
'                        ListaQuin = "0"
'                        Flog.writeline "Error de Configuración. Verifique columna: " & rs_GetConfrepRs!confnrocol
'                    End If
        End Select
        rs_GetConfrepRs.MoveNext
    Loop
Else
    Flog.writeline "No se encontró configuración para el Modelo 500"
    Exit Sub
End If
rs_GetConfrepRs.Close
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'------------------------------------------------------------------
'Obtengo el nro de cabecera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   cliqnro = rsConsult!cliqnro
   proFecPago = rsConsult!proFecPago
   profecini = rsConsult!profecini
   proDesc = rsConsult!proDesc
Else
   Flog.writeline "Error al obtener los datos del proceso"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   Apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
   empFecAlta = rsConsult!empfaltagr
   If IsNull(rsConsult!empremu) Then
      'Sueldo = 0
   Else
      'Sueldo = rsConsult!empremu
   End If
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & Pronro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If



'------------------------------------------------------------------
'Busco los datos para Unidad Administrativa
'------------------------------------------------------------------
If tenro1 <> 0 Then
    If estrnro1 <> 0 Then
        EmpEstrnro1 = estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro1
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!Estrnro
        End If
    End If
End If



'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

If tenro2 <> 0 Then
    If estrnro2 <> 0 Then
        EmpEstrnro2 = estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro2
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------

If tenro3 <> 0 Then
    If estrnro3 <> 0 Then
        EmpEstrnro3 = estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro3
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco el valor del cuil
'------------------------------------------------------------------
'StrSql = " SELECT cuil.nrodoc "
'StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
'StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'   Cuil = rsConsult!Nrodoc
'Else
''   Flog.writeline "Error al obtener los datos del cuil"
''   GoTo MError
'End If

'------------------------------------------------------------------
'Busco el valor de la direccion y localidad del empleado
'------------------------------------------------------------------
'StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,detdom.piso,detdom.oficdepto,detdom.barrio"
'StrSql = StrSql & " FROM  cabdom "
'StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
'StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
'StrSql = StrSql & " WHERE  cabdom.domdefault = -1 AND cabdom.ternro = " & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'   Direccion = IIf(sinDatos(rsConsult!calle), "", rsConsult!calle)
'   Direccion = Direccion & " " & IIf(sinDatos(rsConsult!Nro), "", rsConsult!Nro)
'   Direccion = Direccion & " " & IIf(sinDatos(rsConsult!Piso), "", rsConsult!Piso & "P")
'   Direccion = Direccion & " " & IIf(sinDatos(rsConsult!oficdepto), "", """" & rsConsult!oficdepto & """")
' ' Direccion = Direccion & " " & IIf(sinDatos(rsConsult!barrio), "", rsConsult!barrio)
'   Direccion = Direccion & "," & IIf(sinDatos(rsConsult!locdesc), "", rsConsult!locdesc)
'
'Else
'   Direccion = ""
''   Flog.writeline "Error al obtener los datos de la localidad"
''   GoTo MError
'End If



'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo
'Sueldo = 0
'If Sueldo = 0 Then
StrSql = " SELECT almonto"
StrSql = StrSql & " From acu_liq"
StrSql = StrSql & " Where acunro = " & AcuSueldoMen
StrSql = StrSql & " AND cliqnro = " & cliqnro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Sueldo = rsConsult!almonto
Else
   Flog.writeline "Error al obtener los datos del sueldo"
   Sueldo = 0
End If
'End If

'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------

'StrSql = " SELECT estrdabr "
'StrSql = StrSql & " From his_estructura"
'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
'StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 3 And his_estructura.ternro = " & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'   Categoria = rsConsult!estrdabr
'Else
''   Flog.writeline "Error al obtener los datos de la categoria"
''   GoTo MError
'End If

'------------------------------------------------------------------
'Busco el valor del puesto
'------------------------------------------------------------------

'StrSql = " SELECT estrdabr "
'StrSql = StrSql & " From his_estructura"
'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
'StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'Puesto = ""
'
'If Not rsConsult.EOF Then
'   Puesto = rsConsult!estrdabr
'Else
''   Flog.writeline "Error al obtener los datos del puesto"
''   GoTo MError
'End If


'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & CCostoTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrdabr
Else
    CentroCosto = ""
    Flog.writeline "No se encontró Centro de Costo"
End If



'------------------------------------------------------------------
'Busco el valor del Unidad Administrativa
'------------------------------------------------------------------
StrSql = " SELECT estrdabr,estrcodext "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & UnAdminTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   UnidadAdmin = rsConsult!estrdabr & "@" & rsConsult!estrcodext
Else
    UnidadAdmin = "@"
    Flog.writeline "No se encontró Unidad Administrativa"
End If


'------------------------------------------------------------------
'Busco el valor de Plaza
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & PlazaTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Plaza = rsConsult!estrdabr
Else
    Plaza = ""
    Flog.writeline "No se encontró Plaza"
End If


'------------------------------------------------------------------
'Busco el valor de IVM
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & IVMTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    IVM = rsConsult!estrdabr
Else
    IVM = ""
    Flog.writeline "No se encontró IVM"
End If





'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
'StrSql = " SELECT ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc"
'StrSql = StrSql & " From pago"
'StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro AND pago.pagorigen=" & cliqnro
'StrSql = StrSql & " LEFT JOIN banco     ON banco.ternro = pago.banternro"
'StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'  If rsConsult!fpagbanc = "-1" Then
'    FormaPago = rsConsult!fpagdescabr & " " & rsConsult!terrazsoc & " " & rsConsult!ctabnro
'  Else
'    FormaPago = rsConsult!fpagdescabr
'  End If
'Else
''   Flog.writeline "Error al obtener los datos de la forma de pago"
''   GoTo MError
'End If
'
'rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de la forma de liquidacion del empleado para ver si es quincenal
'------------------------------------------------------------------
quincena = " "

StrSql = "SELECT his_estructura.estrnro " & _
    " From his_estructura" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = 22"
OpenRecordset StrSql, rs_estructura

If Not rs_estructura.EOF Then
    
    If (rs_estructura!Estrnro = quincenal1) Or (rs_estructura!Estrnro = quincenal2) Then
        'Miro que quincena es
        StrSql = " SELECT proceso.pronro, proceso.tprocnro, tprocdesc"
        StrSql = StrSql & " From Proceso"
        StrSql = StrSql & " INNER JOIN tipoproc ON tipoproc.tprocnro = proceso.tprocnro"
        StrSql = StrSql & " Where Proceso.pronro = " & Pronro
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            If Not EsNulo(rsConsult!tprocdesc) Then
                quincena = Left(rsConsult!tprocdesc, 1)
            End If
        End If
        rsConsult.Close
    End If

End If

rs_estructura.Close

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = " & EmpresaTenro
OpenRecordset StrSql, rs_estructura

EmpEstrnro = 0

If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    Exit Sub
Else
    EmpNombre = UCase(rs_estructura!empnom)
    EmpEstrnro = rs_estructura!Estrnro
End If

'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    EmpDire = "   "
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro & ", " & rs_Domicilio!locdesc
End If

'Consulta para obtener el cuit de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa"
    'Exit Sub
    EmpCuit = "  "
Else
    EmpCuit = rs_cuit!NroDoc
End If

'Consulta para buscar el logo de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_logo
If rs_logo.EOF Then
    Flog.writeline "No se encontró el Logo de la Empresa"
    'Exit Sub
    EmpLogo = ""
    EmpLogoAlto = 0
    EmpLogoAncho = 0
Else
    EmpLogo = rs_logo!tipimdire & rs_logo!terimnombre
    EmpLogoAlto = rs_logo!tipimaltodef
    EmpLogoAncho = rs_logo!tipimanchodef
End If

'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    'Exit Sub
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If

'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & EmpEstrnro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqdepant = rsConsult!periodoant
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!Banco
Else
   pliqdepant = ""
   pliqfecdep = ""
   pliqbco = ""
   Flog.writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco la Fecha de Alta del empleado. Corresponde a la fecha desde de la fase que
'contiene la marca de fecha de Alta Reconocida.
'------------------------------------------------------------------
'StrSql = " SELECT altfec "
'StrSql = StrSql & " FROM fases "
'StrSql = StrSql & " WHERE empleado= " & Ternro & " AND fasrecofec = -1 "
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'   EmpFecAltaRec = rsConsult!altfec
'Else
'   Flog.writeline "Error al obtener la Fecha de Alta del Empleado"
'   EmpFecAltaRec = ""
'   'GoTo MError
'End If

'------------------------------------------------------------------
'Busco la Fecha de Baja del empleado. Corresponde a la fecha desde de la fase que
'contiene la marca de fecha de Baja Reconocida 2.
'------------------------------------------------------------------
'StrSql = " SELECT bajfec "
'StrSql = StrSql & " FROM fases "
'StrSql = StrSql & " WHERE empleado= " & Ternro & " AND estado = 0 "
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'   empFecAltarec2 = rsConsult!bajfec
'Else
'   Flog.writeline "Error al obtener la Fecha de Alta del Empleado"
'   empFecAltarec2 = ""
'   'GoTo MError
'End If

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,direccion,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,pliqdepant,"
StrSql = StrSql & " pliqfecdep,pliqbco,cuil,empfecalta,"
StrSql = StrSql & " sueldo,categoria,centrocosto,localidad,"
StrSql = StrSql & " profecpago,empnombre,empdire,empcuit,emplogo,emplogoalto,emplogoancho,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,formapago,prodesc,descripcion,puesto, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden"
StrSql = StrSql & " ,auxchar1,auxchar2,auxchar3,auxchar4"
StrSql = StrSql & ",modeloRecibo"
StrSql = StrSql & ")"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"

'StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & ",''" ' Direccion

StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"

'StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",''" 'Cuil

StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & Sueldo

'StrSql = StrSql & ",'" & Mid(Categoria, 1, 20) & "'"
StrSql = StrSql & ",''" ' Categoria

StrSql = StrSql & ",'" & Mid(CentroCosto, 1, 25) & "'"
StrSql = StrSql & ",'" & Mid(localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & Mid(EmpDire, 1, 200) & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho

'StrSql = StrSql & ",'" & FormaPago & "'"
StrSql = StrSql & ",''" ' FormaPago

StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"

'StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & ",''" ' Puesto

StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden

'StrSql = StrSql & ",'" & EmpFecAltaRec & "'"
'StrSql = StrSql & ",'" & empFecAltarec2 & "'"
StrSql = StrSql & ",'" & Mid(UnidadAdmin, 1, 100) & "'" ' Auxchar1
StrSql = StrSql & ",'" & Mid(Plaza, 1, 100) & "'" ' Auxchar2
StrSql = StrSql & ",'" & Mid(IVM, 1, 100) & "'" 'Auxchar3
StrSql = StrSql & ",'" & Mid(quincena, 1, 100) & "'" 'Auxchar4
StrSql = StrSql & "," & tipoRecibo

StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

Flog.writeline "SQL INSERT: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado
'------------------------------------------------------------------

StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
    
OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
  
    StrSql = " INSERT INTO rep_recibo_det "
    StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
    StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
    StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & "," & Pronro
    StrSql = StrSql & "," & rsConsult!cliqnro
    StrSql = StrSql & ",'" & rsConsult!concabr & "'"
    StrSql = StrSql & ",'" & rsConsult!ConcCod & "'"
    StrSql = StrSql & "," & rsConsult!ConcNro
    StrSql = StrSql & "," & rsConsult!tconnro
    StrSql = StrSql & "," & rsConsult!concimp
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
    StrSql = StrSql & ",'" & arrTipoConc(rsConsult!tconnro) & "')"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rsConsult.MoveNext

Loop

rsConsult.Close

Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub
Sub generarDatosRecibo501(Pronro, Ternro, acunroSueldo, tituloReporte, orden)
'Descripción : COMPROBANTE DE PAGO SYKES - SV
'NG  03/07/2015 -  Se agregó pliqanio a la query de búsqueda
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim Nombre
Dim apellidoReporta
Dim nombreReporta
'Dim Direccion
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
'Dim Cuil
Dim empFecAlta
'Dim EmpFecAltaRec
'Dim empFecAltarec2
Dim Sueldo
'Dim Categoria

Dim localidad
Dim proFecPago
Dim pliqhasta
'Dim FormaPago
'Dim Puesto

Dim EmpEstrnro
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String
Dim quincena As String

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset
'Dim OrdenAux As Long

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

'-----------------------------------------------------------------------------------
'BUSCO CONFIGURACIÓN DE REPORTE PARA EL MODELO QUE SE PROCESA.
'-----------------------------------------------------------------------------------
' || Esta función se utiliza a partir de los modelos 500 ||
' --->  Devuelve Recordset
'-----------------------------------------------------------------------------------
'StrSql = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & tipoRecibo * 100
'StrSql = StrSql & " AND (confnrocol < " & (tipoRecibo * 100) + 100
'StrSql = StrSql & " ORDER BY confnrocol ASC "

'Dim rs_confrep  As New ADODB.Recordset
Dim EmpresaTenro As Long
Dim UnAdminTenro As Long
'Dim CCostoTenro As Long
Dim PlazaTenro As Long
'Dim IVMTenro As Long
Dim ListaQuin As String
'Dim CentroCosto As String
Dim UnidadAdmin As String
Dim Plaza As String
'Dim IVM As String
Dim AcuSueldoMen As Long

Dim AcuIngPriQuin As Long
Dim AcuDescLeyPriQuin As Long
Dim AcuOtrDescPriQuin As Long
Dim AcuDevPriQuin As Long
Dim AcuIngSegQuin As Long
Dim AcuDescLeySegQuin As Long
Dim AcuOtrDescSegQuin As Long
Dim AcuDevSegQuin As Long

Dim IngresoPriQuin As Long
Dim IngresoDescLeyPriQuin As Long
Dim OtrosDescPriQuin As Long
Dim DevPriQuin As Long
Dim IngresoSegQuin As Long
Dim IngresoDescLeySegQuin As Long
Dim OtrosDescSegQuin As Long
Dim DevSegQuin As Long

Dim ModeloQuin1 As String
Dim ModeloQuin2 As String

Dim tprocnro As Long

'Obtengo configuración de reporte
Dim rs_GetConfrepRs As New ADODB.Recordset
StrSql = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & tipoRecibo * 100
StrSql = StrSql & " AND confnrocol < " & (tipoRecibo * 100) + 100 & ")"
StrSql = StrSql & " ORDER BY confnrocol ASC "
'GetConfrepRs tipoRecibo
OpenRecordset StrSql, rs_GetConfrepRs

If Not rs_GetConfrepRs.EOF Then
    Do While Not rs_GetConfrepRs.EOF
        Select Case rs_GetConfrepRs!confnrocol
            Case 50100: 'Empresa
                        EmpresaTenro = rs_GetConfrepRs!confval
            Case 50101: 'Unidad Administrativa
                        UnAdminTenro = rs_GetConfrepRs!confval
            Case 50102: 'Plaza
                        PlazaTenro = rs_GetConfrepRs!confval
            Case 50103: 'Salario Mensual (AC)
                        AcuSueldoMen = rs_GetConfrepRs!confval
            Case 50104: 'Ingresos 1ra Quincena
                        AcuIngPriQuin = rs_GetConfrepRs!confval
            Case 50105: 'Descuentos de ley (1ra Q)
                        AcuDescLeyPriQuin = rs_GetConfrepRs!confval
            Case 50106: 'Otros Descuentos (1ra Q)
                        AcuOtrDescPriQuin = rs_GetConfrepRs!confval
            Case 50107: 'Devengado (1ra Q)
                        AcuDevPriQuin = rs_GetConfrepRs!confval
            Case 50108: 'Ingresos 2da Quincena
                        AcuIngSegQuin = rs_GetConfrepRs!confval
            Case 50109: 'Descuentos de ley (2da Q)
                        AcuDescLeySegQuin = rs_GetConfrepRs!confval
            Case 50110: 'Otros Descuentos (2da Q)
                        AcuOtrDescSegQuin = rs_GetConfrepRs!confval
            Case 50111: 'Devengado (2da Q)
                        AcuDevSegQuin = rs_GetConfrepRs!confval
            Case 50112: 'Modelo 1ra Quincena
                        ModeloQuin1 = rs_GetConfrepRs!confval2
            Case 50113: 'Modelo 2da Quincena
                        ModeloQuin2 = rs_GetConfrepRs!confval2
        End Select
        rs_GetConfrepRs.MoveNext
    Loop
Else
    Flog.writeline "No se encontró configuración para el Modelo " & tipoRecibo
    Exit Sub
End If

rs_GetConfrepRs.Close
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'------------------------------------------------------------------
'Obtengo el nro de cabecera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini,  tipoproc.tprocnro FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN tipoproc ON tipoproc.tprocnro = proceso.tprocnro "
StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   cliqnro = rsConsult!cliqnro
   proFecPago = rsConsult!proFecPago
   profecini = rsConsult!profecini
   proDesc = rsConsult!proDesc
   tprocnro = rsConsult!tprocnro
Else
   Flog.writeline "Error al obtener los datos del proceso"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = "SELECT empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2 "
StrSql = StrSql & ",empleado.empfaltagr,empleado.empremu ,empleado.empreporta "
StrSql = StrSql & ",empReporta.terape terapeReporta,empReporta.terape2 terape2Reporta "
StrSql = StrSql & ",empReporta.ternom ternomReporta,empReporta.ternom2 ternom2Reporta "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " LEFT JOIN empleado empReporta ON empReporta.ternro = empleado.empreporta "
StrSql = StrSql & " WHERE empleado.ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   Apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
   empFecAlta = rsConsult!empfaltagr
   If EsNulo(rsConsult!empreporta) Then
        'Flog.writeline "El empleado Legajo " & rsConsult!empleg
        Flog.writeline "No se encontró Reporta A (Jefe Inmediato)"
        nombreReporta = ""
        apellidoReporta = ""
   Else
        If Ternro <> rsConsult!empreporta Then
            nombreReporta = rsConsult!ternomReporta & " " & rsConsult!ternom2Reporta
            apellidoReporta = rsConsult!terapeReporta & " " & rsConsult!terape2Reporta
        Else
            Flog.writeline "No se encontró Reporta A (Jefe Inmediato)"
            nombreReporta = ""
            apellidoReporta = ""
        End If
   End If
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & Pronro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If



'------------------------------------------------------------------
'Busco los datos para Unidad Administrativa
'------------------------------------------------------------------
If tenro1 <> 0 Then
    If estrnro1 <> 0 Then
        EmpEstrnro1 = estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro1
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!Estrnro
        End If
    End If
End If



'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------
If tenro2 <> 0 Then
    If estrnro2 <> 0 Then
        EmpEstrnro2 = estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro2
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------
If tenro3 <> 0 Then
    If estrnro3 <> 0 Then
        EmpEstrnro3 = estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro3
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'StrSql = " SELECT almonto"
'StrSql = StrSql & " From acu_liq"
'StrSql = StrSql & " Where acunro = " & AcuSueldoMen
'StrSql = StrSql & " AND cliqnro = " & cliqnro

StrSql = " SELECT acu_mes.ammonto"
StrSql = StrSql & " From acu_liq"
StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.acunro =  acu_liq.acunro AND acu_mes.TERNRO = " & Ternro
StrSql = StrSql & " Where acu_mes.acunro = " & AcuSueldoMen
StrSql = StrSql & " AND cliqnro = " & cliqnro
StrSql = StrSql & "  AND acu_mes.ammes = " & pliqmes

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   'Sueldo = rsConsult!almonto
   Sueldo = rsConsult!ammonto
Else
   Flog.writeline "No se encontraron valores para el Salario Mensual" 'Error al obtener los datos del sueldo"
   Sueldo = 0
End If

'=========================================
'   AC DE LA 1ra QUINCENA
'=========================================

''------------------------------------------------------------------------------------------
''INGRESOS | 1ra Quincena
'StrSql = " SELECT almonto"
'StrSql = StrSql & " From acu_liq"
'StrSql = StrSql & " Where acunro = " & AcuIngPriQuin
'StrSql = StrSql & " AND cliqnro = " & cliqnro
'OpenRecordset StrSql, rsConsult
'If Not rsConsult.EOF Then
'   IngresoPriQuin = rsConsult!almonto
'Else
'   Flog.writeline "No se encontraron valores para el Ingreso de la 1ra Quincena "
'   IngresoPriQuin = 0
'End If
'
'
''------------------------------------------------------------------------------------------
''DESCUENTOS DE LEY | 1ra Quincena
'StrSql = " SELECT almonto"
'StrSql = StrSql & " From acu_liq"
'StrSql = StrSql & " Where acunro = " & AcuDescLeyPriQuin
'StrSql = StrSql & " AND cliqnro = " & cliqnro
'OpenRecordset StrSql, rsConsult
'If Not rsConsult.EOF Then
'   DescLeyPriQuin = rsConsult!almonto
'Else
'   Flog.writeline "No se encontraron valores para Descuentos de Ley para la 1ra Quincena "
'   DescLeyPriQuin = 0
'End If
'
'
''------------------------------------------------------------------------------------------
''OTROS DESCUENTOS | 1ra Quincena
'StrSql = " SELECT almonto"
'StrSql = StrSql & " From acu_liq"
'StrSql = StrSql & " Where acunro = " & AcuOtrDescPriQuin
'StrSql = StrSql & " AND cliqnro = " & cliqnro
'OpenRecordset StrSql, rsConsult
'If Not rsConsult.EOF Then
'   OtrosDescPriQuin = rsConsult!almonto
'Else
'   Flog.writeline "No se encontraron valores para Otros Descuentos para la 1ra Quincena "
'   OtrosDescPriQuin = 0
'End If
'
'
''DEVENGADO | 1ra Quincena
'StrSql = " SELECT almonto"
'StrSql = StrSql & " From acu_liq"
'StrSql = StrSql & " Where acunro = " & AcuDevPriQuin
'StrSql = StrSql & " AND cliqnro = " & cliqnro
'OpenRecordset StrSql, rsConsult
'If Not rsConsult.EOF Then
'   DevPriQuin = rsConsult!almonto
'Else
'   Flog.writeline "No se encontraron valores para Devengado para la 1ra Quincena "
'   DevPriQuin = 0
'End If
'
''=========================================
''   AC DE LA 2da QUINCENA
''=========================================
''------------------------------------------------------------------------------------------
''INGRESOS | 2da Quincena
'StrSql = " SELECT almonto"
'StrSql = StrSql & " From acu_liq"
'StrSql = StrSql & " Where acunro = " & AcuIngSegQuin
'StrSql = StrSql & " AND cliqnro = " & cliqnro
'OpenRecordset StrSql, rsConsult
'If Not rsConsult.EOF Then
'   IngresoSegQuin = rsConsult!almonto
'Else
'   Flog.writeline "No se encontraron valores para el Ingreso de la 2da Quincena "
'   IngresoSegQuin = 0
'End If
'
'
''------------------------------------------------------------------------------------------
''DESCUENTOS DE LEY | 2da Quincena
'StrSql = " SELECT almonto"
'StrSql = StrSql & " From acu_liq"
'StrSql = StrSql & " Where acunro = " & AcuDescLeyPriQuin
'StrSql = StrSql & " AND cliqnro = " & cliqnro
'OpenRecordset StrSql, rsConsult
'If Not rsConsult.EOF Then
'   DescLeySegQuin = rsConsult!almonto
'Else
'   Flog.writeline "No se encontraron valores para Descuentos de Ley para la 2da Quincena "
'   DescLeySegiQuin = 0
'End If
'
'
''------------------------------------------------------------------------------------------
''OTROS DESCUENTOS | 2da Quincena
'StrSql = " SELECT almonto"
'StrSql = StrSql & " From acu_liq"
'StrSql = StrSql & " Where acunro = " & AcuOtrDescPriQuin
'StrSql = StrSql & " AND cliqnro = " & cliqnro
'OpenRecordset StrSql, rsConsult
'If Not rsConsult.EOF Then
'   OtrosDescSegQuin = rsConsult!almonto
'Else
'   Flog.writeline "No se encontraron valores para Otros Descuentos para la 2da Quincena "
'   OtrosDescSegQuin = 0
'End If
'
'
''DEVENGADO | 2da Quincena
'StrSql = " SELECT almonto"
'StrSql = StrSql & " From acu_liq"
'StrSql = StrSql & " Where acunro = " & AcuDevSegQuin
'StrSql = StrSql & " AND cliqnro = " & cliqnro
'OpenRecordset StrSql, rsConsult
'If Not rsConsult.EOF Then
'   DevSegQuin = rsConsult!almonto
'Else
'   Flog.writeline "No se encontraron valores para Devengado para la 2da Quincena "
'   DevSegQuin = 0
'End If





'------------------------------------------------------------------
'Busco el valor del Unidad Administrativa
'------------------------------------------------------------------
StrSql = " SELECT estrdabr,estrcodext "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & UnAdminTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   'UnidadAdmin = rsConsult!estrdabr & "@" & rsConsult!estrcodext
   UnidadAdmin = rsConsult!estrdabr
Else
    UnidadAdmin = "@"
    Flog.writeline "No se encontró Unidad Administrativa"
End If


'------------------------------------------------------------------
'Busco el valor de Plaza
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & PlazaTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Plaza = rsConsult!estrdabr
Else
    Plaza = ""
    Flog.writeline "No se encontró Plaza"
End If

'------------------------------------------------------------------
'Busco los datos de la forma de liquidacion del empleado para ver si es quincenal
'------------------------------------------------------------------

StrSql = " SELECT * FROM tipoproc"
StrSql = StrSql & " INNER JOIN proceso ON proceso.tprocnro = tipoproc.tprocnro"
StrSql = StrSql & " WHERE tipoproc.tprocnro IN(3) "
StrSql = StrSql & " AND proceso.pronro =" & Pronro
OpenRecordset StrSql, rs_estructura
If Not rs_estructura.EOF Then
End If

'quincena = " "
'
'StrSql = "SELECT his_estructura.estrnro " & _
'    " From his_estructura" & _
'    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
'    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
'    " AND his_estructura.ternro = " & Ternro & _
'    " AND his_estructura.tenro  = 22"
'OpenRecordset StrSql, rs_estructura
'
'If Not rs_estructura.EOF Then
'
'    If (rs_estructura!Estrnro = quincenal1) Or (rs_estructura!Estrnro = quincenal2) Then
'        'Miro que quincena es
'        StrSql = " SELECT proceso.pronro, proceso.tprocnro, tprocdesc"
'        StrSql = StrSql & " From Proceso"
'        StrSql = StrSql & " INNER JOIN tipoproc ON tipoproc.tprocnro = proceso.tprocnro"
'        StrSql = StrSql & " Where Proceso.pronro = " & Pronro
'        OpenRecordset StrSql, rsConsult
'        If Not rsConsult.EOF Then
'            If Not EsNulo(rsConsult!tprocdesc) Then
'                quincena = Left(rsConsult!tprocdesc, 1)
'            End If
'        End If
'        rsConsult.Close
'    End If
'
'End If
'
'rs_estructura.Close

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = " & EmpresaTenro
OpenRecordset StrSql, rs_estructura

EmpEstrnro = 0

If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    Exit Sub
Else
    EmpNombre = UCase(rs_estructura!empnom)
    EmpEstrnro = rs_estructura!Estrnro
End If

'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    'Exit Sub
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If


'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------
StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,direccion,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,pliqdepant,"
StrSql = StrSql & " pliqfecdep,pliqbco,cuil,empfecalta,"
StrSql = StrSql & " sueldo,categoria,centrocosto,localidad,"
StrSql = StrSql & " profecpago,empnombre,empdire,empcuit,emplogo,emplogoalto,emplogoancho,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,formapago,prodesc,descripcion,puesto, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden"
StrSql = StrSql & " ,auxchar1,auxchar2,auxchar3,auxchar4"
'StrSql = StrSql & " ,auxdeci1,auxdeci2,auxdeci3,auxdeci4"
StrSql = StrSql & ",modeloRecibo"
StrSql = StrSql & ")"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & ",''" ' Direccion
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",''" 'Cuil
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & Sueldo
StrSql = StrSql & ",''" ' Categoria
StrSql = StrSql & ",''" 'CentroCosto
StrSql = StrSql & ",''" 'localidad
StrSql = StrSql & ",''" 'proFecPago
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",''" ' EmpDire
StrSql = StrSql & ",''" 'EmpCuit
StrSql = StrSql & ",''" ' EmpLogo
StrSql = StrSql & ",0" ' EmpLogoAlto
StrSql = StrSql & ",0" '  EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",''" ' FormaPago
StrSql = StrSql & ",'" & Mid(proDesc, 1, 60) & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & ",''" ' Puesto
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Mid(UnidadAdmin, 1, 100) & "'" ' Auxchar1
StrSql = StrSql & ",'" & Mid(Plaza, 1, 100) & "'" ' Auxchar2
StrSql = StrSql & ",'" & nombreReporta & " " & apellidoReporta & "'" 'Auxchar3 (Jefe Inmediato)


StrSqlAux = StrSql & ",'1'" 'Auxchar4
'StrSqlAux = StrSqlAux & "," & numberForSQL(IngresoPriQuin) 'Auxdeci1 (INGRESOS)
'StrSqlAux = StrSqlAux & "," & numberForSQL(DescLeyPriQuin) 'Auxdeci2 (DESCUENTO DE LEY)
'StrSqlAux = StrSqlAux & "," & numberForSQL(OtrosDescPriQuin) 'Auxdeci3 (OTROS DESCUENTOS)
'StrSqlAux = StrSqlAux & "," & numberForSQL(DevPriQuin) 'Auxdeci4 (DEVENGADO)
StrSqlAux = StrSqlAux & "," & tipoRecibo
StrSqlAux = StrSqlAux & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------
Flog.writeline "SQL INSERT: " & StrSqlAux
objConn.Execute StrSqlAux, , adExecuteNoRecords


'
'StrSqlAux = StrSql & ",'2'" 'Auxchar4
'StrSqlAux = StrSqlAux & "," & numberForSQL(IngresoSegQuin) 'Auxdeci1 (INGRESOS)
'StrSqlAux = StrSqlAux & "," & numberForSQL(DescLeySegQuin) 'Auxdeci2 (DESCUENTO DE LEY)
'StrSqlAux = StrSqlAux & "," & numberForSQL(OtrosDescSegQuin) 'Auxdeci3 (OTROS DESCUENTOS)
'StrSqlAux = StrSqlAux & "," & numberForSQL(DevSegQuin) 'Auxdeci4 (DEVENGADO)
'StrSqlAux = StrSqlAux & "," & tipoRecibo
'StrSqlAux = StrSqlAux & ")"
'Flog.writeline "SQL INSERT: " & StrSqlAux
'objConn.Execute StrSqlAux, , adExecuteNoRecords



'------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado
'------------------------------------------------------------------
'StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
'StrSql = StrSql & " FROM cabliq "
'StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
'StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
'StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
'StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
'StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
Dim ArrAcuPriQuin(7) As Long
ArrAcuPriQuin(0) = AcuIngPriQuin
ArrAcuPriQuin(1) = AcuDescLeyPriQuin
ArrAcuPriQuin(2) = AcuOtrDescPriQuin
ArrAcuPriQuin(3) = AcuDevPriQuin
ArrAcuPriQuin(4) = AcuIngSegQuin
ArrAcuPriQuin(5) = AcuDescLeySegQuin
ArrAcuPriQuin(6) = AcuOtrDescSegQuin
ArrAcuPriQuin(7) = AcuDevSegQuin

If (tprocnro = 21 Or tprocnro = 23) Then
    'For a = 0 To UBound(ArrAcuPriQuin)
    For a = 0 To 3
        'StrSql = " SELECT acu_mes.acunro,acu_mes.amcant,acu_mes.ammonto,cabliq.cliqnro FROM cabliq"
        StrSql = " SELECT acu_mes.acunro,acu_mes.amcant,acu_liq.almonto ammonto,cabliq.cliqnro FROM cabliq"
        StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
        StrSql = StrSql & " AND cabliq.empleado = " & Ternro
        StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.acunro =  acu_liq.acunro AND acu_mes.TERNRO = " & Ternro
        StrSql = StrSql & " WHERE acu_mes.acunro = " & ArrAcuPriQuin(a)
        'Agregado 20/08/2014
        StrSql = StrSql & "  AND acu_mes.ammes = " & pliqmes
        'AGREGADO 03/07/2015
        StrSql = StrSql & " AND acu_mes.amanio =" & pliqanio
        'fin
        'StrSql = StrSql & " WHERE acu_mes.acunro IN ("
        '    For a = 0 To UBound(ArrAcuPriQuin)
        '        StrSql = StrSql & ArrAcuPriQuin(a) & ","
        '    Next
        'StrSql = StrSql & ")"
        
        Flog.writeline "SQL consulta por AC configurado: " & StrSql
        
        OpenRecordset StrSql, rsConsult
        '
        If Not rsConsult.EOF Then
            Do Until rsConsult.EOF
                StrSql = " INSERT INTO rep_recibo_det "
                StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
                StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
                StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & "," & Ternro
                StrSql = StrSql & "," & Pronro
                StrSql = StrSql & "," & rsConsult!cliqnro
                
                'StrSql = StrSql & ",'" & rsConsult!concabr & "'"
                'StrSql = StrSql & ",'" & rsConsult!ConcCod & "'"
                'StrSql = StrSql & "," & rsConsult!ConcNro
                'StrSql = StrSql & "," & rsConsult!tconnro
                'StrSql = StrSql & "," & rsConsult!concimp
                'StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
                'StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
                'StrSql = StrSql & ",'" & arrTipoConc(rsConsult!tconnro) & "')"
                
                StrSql = StrSql & ",'" & rsConsult!acuNro & "'"
                StrSql = StrSql & ",'" & rsConsult!acuNro & "'"
                StrSql = StrSql & "," & rsConsult!acuNro
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",-1"
                StrSql = StrSql & "," & numberForSQL(rsConsult!amcant)
                StrSql = StrSql & "," & numberForSQL(rsConsult!ammonto)
                StrSql = StrSql & ",'')"
                objConn.Execute StrSql, , adExecuteNoRecords
                rsConsult.MoveNext
            Loop
        Else
            'Si no encuentra el AC lo inserta en 0
            StrSql = " INSERT INTO rep_recibo_det "
            StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
            StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
            StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & "(" & NroProceso
            StrSql = StrSql & "," & Ternro
            StrSql = StrSql & "," & Pronro
            StrSql = StrSql & ",0"
    
            StrSql = StrSql & ",'" & ArrAcuPriQuin(a) & "'"
            StrSql = StrSql & ",'" & ArrAcuPriQuin(a) & "'"
            StrSql = StrSql & "," & ArrAcuPriQuin(a)
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",-1"
            StrSql = StrSql & "," & numberForSQL(0)
            StrSql = StrSql & "," & numberForSQL(0)
            StrSql = StrSql & ",'')"
            objConn.Execute StrSql, , adExecuteNoRecords
        
        
        End If
        rsConsult.Close
    Next
Else
    If (tprocnro = 22 Or tprocnro = 24) Then
        For a = 4 To 7
            'StrSql = " SELECT acu_mes.acunro,acu_mes.amcant,acu_mes.ammonto,cabliq.cliqnro FROM cabliq"
            StrSql = " SELECT acu_mes.acunro,acu_mes.amcant,acu_liq.almonto ammonto,cabliq.cliqnro FROM cabliq"
            StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
            StrSql = StrSql & " AND cabliq.empleado = " & Ternro
            StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro"
            StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro"
            StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.acunro =  acu_liq.acunro AND acu_mes.TERNRO = " & Ternro
            StrSql = StrSql & " WHERE acu_mes.acunro = " & ArrAcuPriQuin(a)
            'Agregado 20/08/2014
            StrSql = StrSql & "  AND acu_mes.ammes = " & pliqmes
            'AGREGADO 03/07/2015
            StrSql = StrSql & " AND acu_mes.amanio =" & pliqanio
            'fin
            'StrSql = StrSql & " WHERE acu_mes.acunro IN ("
            '    For a = 0 To UBound(ArrAcuPriQuin)
            '        StrSql = StrSql & ArrAcuPriQuin(a) & ","
            '    Next
            'StrSql = StrSql & ")"
            
            Flog.writeline "SQL consulta por AC configurado: " & StrSql
            
            OpenRecordset StrSql, rsConsult
            '
            If Not rsConsult.EOF Then
                Do Until rsConsult.EOF
                    StrSql = " INSERT INTO rep_recibo_det "
                    StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
                    StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
                    StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
                    StrSql = StrSql & " VALUES"
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & "," & Ternro
                    StrSql = StrSql & "," & Pronro
                    StrSql = StrSql & "," & rsConsult!cliqnro
                    
                    'StrSql = StrSql & ",'" & rsConsult!concabr & "'"
                    'StrSql = StrSql & ",'" & rsConsult!ConcCod & "'"
                    'StrSql = StrSql & "," & rsConsult!ConcNro
                    'StrSql = StrSql & "," & rsConsult!tconnro
                    'StrSql = StrSql & "," & rsConsult!concimp
                    'StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
                    'StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
                    'StrSql = StrSql & ",'" & arrTipoConc(rsConsult!tconnro) & "')"
                    
                    StrSql = StrSql & ",'" & rsConsult!acuNro & "'"
                    StrSql = StrSql & ",'" & rsConsult!acuNro & "'"
                    StrSql = StrSql & "," & rsConsult!acuNro
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",-1"
                    StrSql = StrSql & "," & numberForSQL(rsConsult!amcant)
                    StrSql = StrSql & "," & numberForSQL(rsConsult!ammonto)
                    StrSql = StrSql & ",'')"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    rsConsult.MoveNext
                Loop
            Else
                'Si no encuentra el AC lo inserta en 0
                StrSql = " INSERT INTO rep_recibo_det "
                StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
                StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
                StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
                StrSql = StrSql & " VALUES"
                StrSql = StrSql & "(" & NroProceso
                StrSql = StrSql & "," & Ternro
                StrSql = StrSql & "," & Pronro
                StrSql = StrSql & ",0"
        
                StrSql = StrSql & ",'" & ArrAcuPriQuin(a) & "'"
                StrSql = StrSql & ",'" & ArrAcuPriQuin(a) & "'"
                StrSql = StrSql & "," & ArrAcuPriQuin(a)
                StrSql = StrSql & ",0"
                StrSql = StrSql & ",-1"
                StrSql = StrSql & "," & numberForSQL(0)
                StrSql = StrSql & "," & numberForSQL(0)
                StrSql = StrSql & ",'')"
                objConn.Execute StrSql, , adExecuteNoRecords
            
            
            End If
            rsConsult.Close
        Next
    End If
End If
Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub
Sub generarDatosRecibo502(Pronro, Ternro, acunroSueldo, tituloReporte, orden)
'Descripción : BOLETA DE PAGO (QUINCENAL) SYKES - SV
'Codigo de forma de liq quincenal
'Const quincenal1 = 553
'Const quincenal2 = 555

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim Nombre
'Dim Direccion
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
'Dim Cuil
Dim empFecAlta
'Dim EmpFecAltaRec
'Dim empFecAltarec2
Dim Sueldo
'Dim Categoria

Dim localidad
Dim proFecPago
Dim pliqhasta
'Dim FormaPago
'Dim Puesto

Dim EmpEstrnro
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String
Dim quincena As String

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset


On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

'-----------------------------------------------------------------------------------
'BUSCO CONFIGURACIÓN DE REPORTE PARA EL MODELO QUE SE PROCESA.
'-----------------------------------------------------------------------------------
' || Esta función se utiliza a partir de los modelos 500 ||
' --->  Devuelve Recordset
'-----------------------------------------------------------------------------------
'StrSql = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & tipoRecibo * 100
'StrSql = StrSql & " AND (confnrocol < " & (tipoRecibo * 100) + 100
'StrSql = StrSql & " ORDER BY confnrocol ASC "

'Dim rs_confrep  As New ADODB.Recordset
Dim EmpresaTenro As Long
Dim UnAdminTenro As Long
Dim CCostoTenro As Long
Dim PlazaTenro As Long
Dim IVMTenro As Long
Dim ListaQuin As String
Dim CentroCosto As String
Dim UnidadAdmin As String
Dim Plaza As String
Dim IVM As String
Dim AcuSueldoMen As Long

'Obtengo configuración de reporte
Dim rs_GetConfrepRs As New ADODB.Recordset
StrSql = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & tipoRecibo * 100
StrSql = StrSql & " AND confnrocol < " & (tipoRecibo * 100) + 100 & ")"
StrSql = StrSql & " ORDER BY confnrocol ASC "
'GetConfrepRs tipoRecibo
OpenRecordset StrSql, rs_GetConfrepRs

If Not rs_GetConfrepRs.EOF Then
    Do While Not rs_GetConfrepRs.EOF
        Select Case rs_GetConfrepRs!confnrocol
            Case 50200: 'Empresa
                        EmpresaTenro = rs_GetConfrepRs!confval
            Case 50201: 'Unidad Administrativa
                        UnAdminTenro = rs_GetConfrepRs!confval
            Case 50202: 'Centro de Costo
                        CCostoTenro = rs_GetConfrepRs!confval
            Case 50203: 'Plaza
                        PlazaTenro = rs_GetConfrepRs!confval
            Case 50204: 'IVM
                        IVMTenro = rs_GetConfrepRs!confval
            Case 50205: 'Salario Mensual (AC)
                        AcuSueldoMen = rs_GetConfrepRs!confval
'            Case 50206: 'Quincena (lista de modelos)
'                    If EsNulo(rs_GetConfrepRs!confval2) Then
'                        ListaQuin = rs_GetConfrepRs!confval2
'                    Else
'                        ListaQuin = "0"
'                        Flog.writeline "Error de Configuración. Verifique columna: " & rs_GetConfrepRs!confnrocol
'                    End If
        End Select
        rs_GetConfrepRs.MoveNext
    Loop
Else
    Flog.writeline "No se encontró configuración para el Modelo 500"
    Exit Sub
End If
rs_GetConfrepRs.Close
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'------------------------------------------------------------------
'Obtengo el nro de cabecera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   cliqnro = rsConsult!cliqnro
   proFecPago = rsConsult!proFecPago
   profecini = rsConsult!profecini
   proDesc = rsConsult!proDesc
Else
   Flog.writeline "Error al obtener los datos del proceso"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   Apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
   empFecAlta = rsConsult!empfaltagr
   If IsNull(rsConsult!empremu) Then
      'Sueldo = 0
   Else
      'Sueldo = rsConsult!empremu
   End If
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & Pronro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If



'------------------------------------------------------------------
'Busco los datos para Unidad Administrativa
'------------------------------------------------------------------
If tenro1 <> 0 Then
    If estrnro1 <> 0 Then
        EmpEstrnro1 = estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro1
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!Estrnro
        End If
    End If
End If



'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

If tenro2 <> 0 Then
    If estrnro2 <> 0 Then
        EmpEstrnro2 = estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro2
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------

If tenro3 <> 0 Then
    If estrnro3 <> 0 Then
        EmpEstrnro3 = estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro3
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!Estrnro
        End If
    End If
End If


'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo
'Sueldo = 0
'If Sueldo = 0 Then
StrSql = " SELECT almonto"
StrSql = StrSql & " From acu_liq"
StrSql = StrSql & " Where acunro = " & AcuSueldoMen
StrSql = StrSql & " AND cliqnro = " & cliqnro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Sueldo = rsConsult!almonto
Else
   Flog.writeline "Error al obtener los datos del sueldo"
   Sueldo = 0
End If
'End If

'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------

'StrSql = " SELECT estrdabr "
'StrSql = StrSql & " From his_estructura"
'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
'StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 3 And his_estructura.ternro = " & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'   Categoria = rsConsult!estrdabr
'Else
''   Flog.writeline "Error al obtener los datos de la categoria"
''   GoTo MError
'End If

'------------------------------------------------------------------
'Busco el valor del puesto
'------------------------------------------------------------------

'StrSql = " SELECT estrdabr "
'StrSql = StrSql & " From his_estructura"
'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
'StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'Puesto = ""
'
'If Not rsConsult.EOF Then
'   Puesto = rsConsult!estrdabr
'Else
''   Flog.writeline "Error al obtener los datos del puesto"
''   GoTo MError
'End If


'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & CCostoTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrdabr
Else
    CentroCosto = ""
    Flog.writeline "No se encontró Centro de Costo"
End If



'------------------------------------------------------------------
'Busco el valor del Unidad Administrativa
'------------------------------------------------------------------
StrSql = " SELECT estrdabr,estrcodext "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & UnAdminTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   UnidadAdmin = rsConsult!estrdabr & "@" & rsConsult!estrcodext
Else
    UnidadAdmin = "@"
    Flog.writeline "No se encontró Unidad Administrativa"
End If


'------------------------------------------------------------------
'Busco el valor de Plaza
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & PlazaTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Plaza = rsConsult!estrdabr
Else
    Plaza = ""
    Flog.writeline "No se encontró Plaza"
End If


'------------------------------------------------------------------
'Busco el valor de IVM
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= " & IVMTenro
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    IVM = rsConsult!estrdabr
Else
    IVM = ""
    Flog.writeline "No se encontró IVM"
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
'StrSql = " SELECT ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc"
'StrSql = StrSql & " From pago"
'StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro AND pago.pagorigen=" & cliqnro
'StrSql = StrSql & " LEFT JOIN banco     ON banco.ternro = pago.banternro"
'StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'  If rsConsult!fpagbanc = "-1" Then
'    FormaPago = rsConsult!fpagdescabr & " " & rsConsult!terrazsoc & " " & rsConsult!ctabnro
'  Else
'    FormaPago = rsConsult!fpagdescabr
'  End If
'Else
''   Flog.writeline "Error al obtener los datos de la forma de pago"
''   GoTo MError
'End If
'
'rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de la forma de liquidacion del empleado para ver si es quincenal
'------------------------------------------------------------------
'quincena = " "
'
'StrSql = "SELECT his_estructura.estrnro " & _
'    " From his_estructura" & _
'    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
'    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
'    " AND his_estructura.ternro = " & Ternro & _
'    " AND his_estructura.tenro  = 22"
'OpenRecordset StrSql, rs_estructura
'
'If Not rs_estructura.EOF Then
'
'    If (rs_estructura!Estrnro = quincenal1) Or (rs_estructura!Estrnro = quincenal2) Then
'        'Miro que quincena es
'        StrSql = " SELECT proceso.pronro, proceso.tprocnro, tprocdesc"
'        StrSql = StrSql & " From Proceso"
'        StrSql = StrSql & " INNER JOIN tipoproc ON tipoproc.tprocnro = proceso.tprocnro"
'        StrSql = StrSql & " Where Proceso.pronro = " & Pronro
'        OpenRecordset StrSql, rsConsult
'        If Not rsConsult.EOF Then
'            If Not EsNulo(rsConsult!tprocdesc) Then
'                quincena = Left(rsConsult!tprocdesc, 1)
'            End If
'        End If
'        rsConsult.Close
'    End If
'
'End If
'
'rs_estructura.Close

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = " & EmpresaTenro
OpenRecordset StrSql, rs_estructura

EmpEstrnro = 0

If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    Exit Sub
Else
    EmpNombre = UCase(rs_estructura!empnom)
    EmpEstrnro = rs_estructura!Estrnro
End If

'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    EmpDire = "   "
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro & ", " & rs_Domicilio!locdesc
End If

'Consulta para obtener el cuit de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa"
    'Exit Sub
    EmpCuit = "  "
Else
    EmpCuit = rs_cuit!NroDoc
End If

'Consulta para buscar el logo de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_logo
If rs_logo.EOF Then
    Flog.writeline "No se encontró el Logo de la Empresa"
    'Exit Sub
    EmpLogo = ""
    EmpLogoAlto = 0
    EmpLogoAncho = 0
Else
    EmpLogo = rs_logo!tipimdire & rs_logo!terimnombre
    EmpLogoAlto = rs_logo!tipimaltodef
    EmpLogoAncho = rs_logo!tipimanchodef
End If

'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    'Exit Sub
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If

'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & EmpEstrnro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqdepant = rsConsult!periodoant
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!Banco
Else
   pliqdepant = ""
   pliqfecdep = ""
   pliqbco = ""
   Flog.writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco la Fecha de Alta del empleado. Corresponde a la fecha desde de la fase que
'contiene la marca de fecha de Alta Reconocida.
'------------------------------------------------------------------
StrSql = " SELECT altfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE empleado= " & Ternro & " AND fasrecofec = -1 "

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   empFecAltaRec = rsConsult!altfec
Else
   Flog.writeline "No se encontró fecha de Alta Reconocida."
   empFecAltaRec = ""
   'GoTo MError
End If

'------------------------------------------------------------------
'Busco la Fecha de Baja del empleado. Corresponde a la fecha desde de la fase que
'------------------------------------------------------------------
StrSql = " SELECT bajfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE empleado= " & Ternro & " AND estado = 0 "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   empFecAltarec2 = rsConsult!bajfec
Else
   Flog.writeline "No se encontró fecha de baja."
   empFecAltarec2 = ""
   'GoTo MError
End If

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,direccion,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,pliqdepant,"
StrSql = StrSql & " pliqfecdep,pliqbco,cuil,empfecalta,"
StrSql = StrSql & " sueldo,categoria,centrocosto,localidad,"
StrSql = StrSql & " profecpago,empnombre,empdire,empcuit,emplogo,emplogoalto,emplogoancho,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,formapago,prodesc,descripcion,puesto, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden"
StrSql = StrSql & " ,auxchar1,auxchar2,auxchar3,auxchar4,auxchar5"
StrSql = StrSql & ",modeloRecibo"
StrSql = StrSql & ")"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"

'StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & ",''" ' Direccion

StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"

'StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",''" 'Cuil

StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & Sueldo

'StrSql = StrSql & ",'" & Mid(Categoria, 1, 20) & "'"
StrSql = StrSql & ",''" ' Categoria

StrSql = StrSql & ",'" & Mid(CentroCosto, 1, 25) & "'"
StrSql = StrSql & ",'" & Mid(localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & Mid(EmpDire, 1, 200) & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho

'StrSql = StrSql & ",'" & FormaPago & "'"
StrSql = StrSql & ",''" ' FormaPago

StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"

'StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & ",''" ' Puesto

StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden

StrSql = StrSql & ",'" & Mid(UnidadAdmin, 1, 100) & "'" ' Auxchar1
StrSql = StrSql & ",'" & Mid(Plaza, 1, 100) & "'" ' Auxchar2
StrSql = StrSql & ",'" & Mid(IVM, 1, 100) & "'" 'Auxchar3

'StrSql = StrSql & ",'" & Mid(quincena, 1, 100) & "'" 'Auxchar4

StrSql = StrSql & ",'" & empFecAltaRec & "'" 'Auxchar4
StrSql = StrSql & ",'" & empFecAltarec2 & "'" 'Auxchar5



StrSql = StrSql & "," & tipoRecibo

StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

Flog.writeline "SQL INSERT: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado
'------------------------------------------------------------------

StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
    
OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
  
    StrSql = " INSERT INTO rep_recibo_det "
    StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
    StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
    StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & "," & Pronro
    StrSql = StrSql & "," & rsConsult!cliqnro
    StrSql = StrSql & ",'" & rsConsult!concabr & "'"
    StrSql = StrSql & ",'" & rsConsult!ConcCod & "'"
    StrSql = StrSql & "," & rsConsult!ConcNro
    StrSql = StrSql & "," & rsConsult!tconnro
    StrSql = StrSql & "," & rsConsult!concimp
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
    StrSql = StrSql & ",'" & arrTipoConc(rsConsult!tconnro) & "')"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rsConsult.MoveNext

Loop

rsConsult.Close

Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub
'--------------------------------------------------------------------
' Se encarga de generar el recibo Mensual de Bolivia Modelo 503
'--------------------------------------------------------------------
Sub generarDatosRecibo503(Pronro, Ternro, acunroSueldo, tituloReporte, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim Nombre
Dim Direccion
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim empFecAlta
Dim Sueldo
Dim SalarioNeto
Dim CentroCosto
Dim Departamento
Dim localidad
Dim proFecPago
Dim pliqhasta
Dim Banco
Dim CodBanco
Dim nroCuenta
Dim NIT
Dim Padron
Dim Seguro
Dim Nro_Padron
Dim Nro_Seguro
Dim Nro_NIT
Dim Sucursal

Dim UFV
Dim TipoUFV
Dim Valor_TipoUFV
Dim SaldoIVA
Dim Valor_SaldoIVA

Dim HHMM
Dim PMAM
Dim HoraProceso

Dim tel

Dim EmpEstrnro
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String
Dim ColumnasConfrep As String
Dim ModeloLiq As String
 
Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3
Dim ObraSocial
Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset
Dim rs_Telefono As New ADODB.Recordset

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

Flog.writeline "Modelo 503"

'Obtengo configuración de reporte
Dim rs_GetConfrepRs As New ADODB.Recordset
StrSql = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & tipoRecibo * 100
StrSql = StrSql & " AND confnrocol < " & (tipoRecibo * 100) + 100 & ")"
StrSql = StrSql & " ORDER BY confnrocol ASC "
'GetConfrepRs tipoRecibo
OpenRecordset StrSql, rs_GetConfrepRs
Flog.writeline "confrep" & StrSql
If Not rs_GetConfrepRs.EOF Then
    Do While Not rs_GetConfrepRs.EOF
        Select Case rs_GetConfrepRs!confnrocol
             Case 50300: 'Tipo Cambio UFV
                        If rs_GetConfrepRs!conftipo = "CO" Then
                            TipoUFV = True
                            UFV = rs_GetConfrepRs!confval
                        Else
                            TipoUFV = False
                            UFV = rs_GetConfrepRs!confval2
                        End If
            Case 50301: 'Saldo IVA sgte Mes
                        If rs_GetConfrepRs!conftipo = "CO" Then
                           TipoSaldoIVA = True
                           SaldoIVA = rs_GetConfrepRs!confval
                        Else
                            TipoSaldoIVA = False
                            SaldoIVA = rs_GetConfrepRs!confval2
                        End If
            Case 50302: 'NIT
                           NIT = rs_GetConfrepRs!confval
            Case 50303: 'Padron
                           Padron = rs_GetConfrepRs!confval
            Case 50304: 'Numero Seguro
                           Seguro = rs_GetConfrepRs!confval
        End Select
         rs_GetConfrepRs.MoveNext
    Loop
Else
    Flog.writeline "No se encontró configuración para el Modelo 503"
    'Exit Sub
End If
rs_GetConfrepRs.Close


'------------------------------------------------------------------
'Obtengo el nro de cabezera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   cliqnro = rsConsult!cliqnro
   proFecPago = rsConsult!proFecPago
   profecini = rsConsult!profecini
   proDesc = rsConsult!proDesc
Else
   Flog.writeline "Error al obtener los datos del proceso"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del Modelo de Liquidación
'------------------------------------------------------------------
ModeloLiq = " "

StrSql = " SELECT tipoproc.tprocdesc FROM proceso"
StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
StrSql = StrSql & " AND proceso.pronro = " & Pronro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    ModeloLiq = rsConsult!tprocdesc
End If

'------------------------------------------------------------------
'Obtengo Tipo Cambio UFV
'------------------------------------------------------------------
If TipoUFV Then
     StrSql = " SELECT detliq.dlimonto valor"
     StrSql = StrSql & " FROM detliq "
     StrSql = StrSql & " INNER JOIN concepto ON concepto.conccod = " & UFV
     StrSql = StrSql & " AND concepto.concnro = detliq.concnro "
     StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
Else
     StrSql = " SELECT almonto valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & UFV
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
     Valor_TipoUFV = Replace(rsConsult!Valor, ",", ".")
Else
    Valor_TipoUFV = 0
End If
Flog.writeline "Tipo Cambio UFV " & StrSql

'------------------------------------------------------------------
'Saldo IVA sgte Mes
'------------------------------------------------------------------
If TipoSaldoIVA Then
     StrSql = " SELECT detliq.dlimonto valor "
     StrSql = StrSql & " FROM detliq "
     StrSql = StrSql & " INNER JOIN concepto ON concepto.conccod = " & SaldoIVA
     StrSql = StrSql & " AND concepto.concnro = detliq.concnro "
     StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
Else
     StrSql = " SELECT almonto valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & SaldoIVA
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Valor_SaldoIVA = Replace(rsConsult!Valor, ",", ".")
Else
    Valor_SaldoIVA = 0
End If
Flog.writeline "Saldo IVA sgte Mes " & StrSql

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   Apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
   'empFecAlta = rsConsult!empfaltagr
   If IsNull(rsConsult!empremu) Then
      Sueldo = 0
   Else
      Sueldo = Replace(rsConsult!empremu, ",", ".")
   End If
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
Flog.writeline "Buscando los datos de la cuenta del empleado."

StrSql = " SELECT * FROM ctabancaria LEFT JOIN banco ON banco.ternro = ctabancaria.banco WHERE ctabestado=-1 AND ctabancaria.ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
        Banco = rsConsult!Bandesc
        CodBanco = rsConsult!Banco
        nroCuenta = rsConsult!ctabnro
        Flog.writeline "Datos de la cuenta bancaria obtenidos"
  Else
        Banco = ""
        CodBanco = ""
        nroCuenta = ""
        Flog.writeline "El empleado no tiene cuentas bancarias activas"
End If
 
'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & Pronro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

If tenro1 <> 0 Then
    If estrnro1 <> 0 Then
        EmpEstrnro1 = estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro1
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!Estrnro
        End If
    End If
End If
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

If tenro2 <> 0 Then
    If estrnro2 <> 0 Then
        EmpEstrnro2 = estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro2
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!Estrnro
        End If
    End If
End If
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------
If tenro3 <> 0 Then
    If estrnro3 <> 0 Then
        EmpEstrnro3 = estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro3
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!Estrnro
        End If
    End If
End If
 
'------------------------------------------------------------------
'Busco el valor de la sucursal
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 1 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

Sucursal = ""

If Not rsConsult.EOF Then
   Sucursal = rsConsult!estrdabr
Else
   Flog.writeline "Error al obtener los datos de la sucursal"
'   GoTo MError
End If
 
'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
        
    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
       Sueldo = Replace(rsConsult!almonto, ",", ".")
    Else
       Flog.writeline "Error al obtener los datos del sueldo. Se queda con el valor de la remuneracion cargada en ADP."
       Sueldo = 0
    End If

'------------------------------------------------------------------
'Busco el valor del Departamento
'------------------------------------------------------------------
StrSql = " SELECT estructura.estrnro,estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= 9"
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   Departamento = rsConsult!estrdabr
Else
    Departamento = ""
    Flog.writeline "No se encontró el Departamento"
End If

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------
Flog.writeline "Busco los datos de la empresa"
StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = 10"
OpenRecordset StrSql, rs_estructura

EmpEstrnro = 0

If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    Exit Sub
Else
    EmpNombre = rs_estructura!empnom
    EmpEstrnro = rs_estructura!Estrnro
End If

'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc, detdom.piso, detdom.oficdepto, detdom.domnro From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    EmpDire = "   "
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro
    If Not EsNulo(rs_Domicilio!Piso) Then
        EmpDire = EmpDire & " P. " & rs_Domicilio!Piso
    End If
Flog.writeline " Domicilio de la Empresa " & StrSql
' Telefono principal
    StrSql = "SELECT telnro FROM telefono " & _
        "WHERE domnro = " & rs_Domicilio!domnro & " AND teldefault = -1 "
    OpenRecordset StrSql, rs_Telefono
    If Not rs_Telefono.EOF Then
        tel = rs_Telefono!telnro
    Else
        tel = " "
    End If
    Flog.writeline "Telefono de la Empresa " & StrSql
    If Not EsNulo(rs_Domicilio!oficdepto) Then
        EmpDire = EmpDire & " Dpto. " & rs_Domicilio!oficdepto
    End If
    EmpDire = EmpDire & " - " & rs_Domicilio!locdesc
End If

'Consulta para obtener el cuit de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa"
    'Exit Sub
    EmpCuit = "  "
Else
    EmpCuit = rs_cuit!NroDoc
End If

'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    'Exit Sub
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If

'----------------------------------------------------------------
'Padron
'----------------------------------------------------------------
StrSql = "SELECT nrodoc FROM ter_doc "
StrSql = StrSql & "INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro "
StrSql = StrSql & "WHERE ternro = " & rs_estructura!Ternro
StrSql = StrSql & " AND tipodocu.tidnro = " & Padron
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    Nro_Padron = rsConsult!NroDoc
Else
    Nro_Padron = " "
End If

Flog.writeline "Query para el NIT " & StrSql
Flog.writeline "El Nro de Padron es " & Nro_Padron


'----------------------------------------------------------------
'Seguro
'----------------------------------------------------------------
StrSql = "SELECT nrodoc FROM ter_doc "
StrSql = StrSql & "INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro "
StrSql = StrSql & "WHERE ternro = " & rs_estructura!Ternro
StrSql = StrSql & " AND tipodocu.tidnro = " & Seguro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    Nro_Seguro = rsConsult!NroDoc
Else
    Nro_Seguro = " "
End If

Flog.writeline "Query para el Seguro " & StrSql
Flog.writeline "El Nro de Seguro es " & Nro_Seguro


'----------------------------------------------------------------
'NIT
'----------------------------------------------------------------
StrSql = "SELECT nrodoc FROM ter_doc "
StrSql = StrSql & "INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro "
StrSql = StrSql & "WHERE ternro = " & rs_estructura!Ternro
StrSql = StrSql & " AND tipodocu.tidnro = " & NIT
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    Nro_NIT = rsConsult!NroDoc
Else
    Nro_NIT = " "
End If

Flog.writeline "Query para el NIT " & StrSql
Flog.writeline "El Nro de NIT es " & Nro_NIT

'------------------------------------------------------------------
'Busco la Fecha de Alta del empleado. Corresponde a la fecha desde de la fase que
'contiene la marca de fecha de Alta Reconocida.
'------------------------------------------------------------------
StrSql = " SELECT altfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE fasrecofec = -1 AND empleado= " & Ternro '& " ORDER BY altfec "
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    empFecAlta = rsConsult!altfec
Else
    Flog.writeline Espacios(Tabulador * 1) & "Error al obtener la Fecha de Alta del Empleado. Se obtiene de la fase marcada como Fecha Alta Reconocida."
    empFecAlta = "&nbsp;"
End If

'------------------------------------------------------------------
'Busco la Hora en que se procesa
'------------------------------------------------------------------
Hora = TimeValue(Now)

HHMM = Mid(Hora, 1, 5)
PMAM = Mid(Hora, 10, 4)
HoraProceso = HHMM & "" & PMAM

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------
Flog.writeline "Armo la SQL para guardar los datos"

StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,"
StrSql = StrSql & " empfecalta,"
StrSql = StrSql & " sueldo, centrocosto,"
StrSql = StrSql & " profecpago,empnombre,empdire,empcuit,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,prodesc,descripcion, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden,auxchar1,auxchar2,auxchar3,modeloRecibo,auxdeci1, "
StrSql = StrSql & " auxdeci2, auxchar4, auxchar5, auxchar6, auxchar7,auxchar8,auxchar9)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & Sueldo
StrSql = StrSql & ",'" & Departamento & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & Sucursal & "'"      'No se usa
StrSql = StrSql & ",'" & EmpCuit & "'"      'No se usa
StrSql = StrSql & ",'" & EmpFirma & "'"     'No se usa
StrSql = StrSql & "," & EmpFirmaAlto        'No se usa
StrSql = StrSql & "," & EmpFirmaAncho       'No se usa
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"   'No se usa Descripcion
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & CodBanco & "'"
StrSql = StrSql & ",'" & Banco & "'"
StrSql = StrSql & ",'" & nroCuenta & "'"
StrSql = StrSql & "," & tipoRecibo
StrSql = StrSql & "," & Valor_TipoUFV
StrSql = StrSql & "," & Valor_SaldoIVA
StrSql = StrSql & ",'" & HoraProceso & "'"
StrSql = StrSql & ",'" & tel & "'"
StrSql = StrSql & ",'" & Nro_Padron & "'"
StrSql = StrSql & ",'" & Nro_Seguro & "'"
StrSql = StrSql & ",'" & Nro_NIT & "'"
StrSql = StrSql & ",'" & ModeloLiq & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

Flog.writeline "SQL INSERT: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado
'------------------------------------------------------------------

StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
    
OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
  
    StrSql = " INSERT INTO rep_recibo_det "
    StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
    StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
    StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & "," & Pronro
    StrSql = StrSql & "," & rsConsult!cliqnro
    StrSql = StrSql & ",'" & rsConsult!concabr & "'"
    StrSql = StrSql & ",'" & rsConsult!ConcCod & "'"
    StrSql = StrSql & "," & rsConsult!ConcNro
    StrSql = StrSql & "," & rsConsult!tconnro
    StrSql = StrSql & "," & rsConsult!concimp
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
    StrSql = StrSql & ",'" & arrTipoConc(rsConsult!tconnro) & "')"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rsConsult.MoveNext

Loop

rsConsult.Close

Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub

'--------------------------------------------------------------------
' Se encarga de generar el recibo Quinquenios de Bolivia Modelo 504
'--------------------------------------------------------------------
Sub generarDatosRecibo504(Pronro, Ternro, acunroSueldo, tituloReporte, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim Nombre
Dim Direccion
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim empFecAlta

Dim MesActual
Dim AnioActual

Dim MesNro1 As Integer
Dim AnioNro1 As Integer
Dim MesNro2 As Integer
Dim AnioNro2 As Integer
Dim MesNro3 As Integer
Dim AnioNro3 As Integer

Dim Sueldo
Dim SalarioNeto
Dim CentroCosto
Dim localidad
Dim proFecPago
Dim pliqhasta
Dim Banco
Dim CodBanco
Dim nroCuenta
Dim NIT
Dim sexo
Dim FechaNac
Dim Puesto
Dim FechaBaj
Dim CalcPromedioMes1
Dim CalcPromedioMes2
Dim CalcPromedioMes3

Dim CalcPromMes1
Dim CalcPromMes2
Dim CalcPromMes3

Dim esCalcPromedioMes1
Dim esCalcPromedioMes2
Dim esCalcPromedioMes3

Dim anios
Dim Meses
Dim Dias

Dim Fecha_aux As Date
Dim anios_aux As Integer
Dim meses_aux As Integer
Dim dias_aux As Integer
Dim diasHab_aux As Integer

Dim EmpEstrnro
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String
Dim ColumnasConfrep As String
 
Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3
Dim ObraSocial
Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

Flog.writeline "Modelo 504"

'------------------------------------------------------------------
'Obtengo el nro de cabezera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini, pliqmes, pliqanio FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   cliqnro = rsConsult!cliqnro
   proFecPago = rsConsult!proFecPago
   profecini = rsConsult!profecini
   proDesc = rsConsult!proDesc
   MesActual = rsConsult!pliqmes
   AnioActual = rsConsult!pliqanio
Else
   Flog.writeline "Error al obtener los datos del proceso"
   GoTo MError
End If

Call PeriodosAnteriores(MesActual, AnioActual, MesNro1, AnioNro1, MesNro2, AnioNro2, MesNro3, AnioNro3)

'Obtengo configuración de reporte
Dim rs_GetConfrepRs As New ADODB.Recordset
StrSql = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & tipoRecibo * 100
StrSql = StrSql & " AND confnrocol < " & (tipoRecibo * 100) + 100 & ")"
StrSql = StrSql & " ORDER BY confnrocol ASC "
'GetConfrepRs tipoRecibo
OpenRecordset StrSql, rs_GetConfrepRs
Flog.writeline "confrep" & StrSql
If Not rs_GetConfrepRs.EOF Then
    Do While Not rs_GetConfrepRs.EOF
        Select Case rs_GetConfrepRs!confnrocol
            Case 50400: 'Calculo Promedio
                        If rs_GetConfrepRs!conftipo = "CO" Then
                            esCalcPromedioMes1 = True
                            CalcPromedioMes1 = rs_GetConfrepRs!confval
                        Else
                            esCalcPromedioMes1 = False
                            CalcPromedioMes1 = rs_GetConfrepRs!confval2
                        End If
            Case 50401: 'Calculo Promedio
                        If rs_GetConfrepRs!conftipo = "CO" Then
                            esCalcPromedioMes2 = True
                            CalcPromedioMes2 = rs_GetConfrepRs!confval
                        Else
                            esCalcPromedioMes2 = False
                            CalcPromedioMes2 = rs_GetConfrepRs!confval2
                        End If
            Case 50402: 'Calculo Promedio
                        If rs_GetConfrepRs!conftipo = "CO" Then
                            esCalcPromedioMes3 = True
                            CalcPromedioMes3 = rs_GetConfrepRs!confval
                        Else
                            esCalcPromedioMes3 = False
                            CalcPromedioMes3 = rs_GetConfrepRs!confval2
                        End If
        End Select
         rs_GetConfrepRs.MoveNext
    Loop
Else
    Flog.writeline "No se encontró configuración para el Modelo 504"
    'Exit Sub
End If
rs_GetConfrepRs.Close


'------------------------------------------------------------------
'Obtengo Calculo Promedio Mes 1
'------------------------------------------------------------------
If esCalcPromedioMes1 Then
'     StrSql = " SELECT detliq.dlimonto valor "
'     StrSql = StrSql & " FROM detliq "
'     StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And detliq.concnro = " & CalcPromedioMes1
    StrSql = "SELECT detliq.dlimonto valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro=3 AND pliqmes = " & MesNro1
    StrSql = StrSql & " AND pliqanio = " & AnioNro1 & " And detliq.ConcNro = " & CalcPromedioMes1
Else
'     StrSql = " SELECT almonto valor"
'     StrSql = StrSql & " From acu_liq"
'     StrSql = StrSql & " Where acunro = " & CalcPromedioMes1
'     StrSql = StrSql & " AND cliqnro = " & cliqnro
    StrSql = "SELECT sum(almonto) valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro = 3 AND pliqmes = " & MesNro1
    StrSql = StrSql & " AND pliqanio = " & AnioNro1
    StrSql = StrSql & " AND acu_liq.acuNro = " & CalcPromedioMes1
    StrSql = StrSql & " AND empleado = " & Ternro
    StrSql = StrSql & " GROUP BY empleado"
End If
OpenRecordset StrSql, rsConsult
Flog.writeline "CalcPromedioMes1" & StrSql
If Not rsConsult.EOF Then
    CalcPromMes1 = Replace(rsConsult!Valor, ",", ".")
Else
    CalcPromMes1 = 0
End If

'------------------------------------------------------------------
'Obtengo Calculo Promedio Mes 2
'------------------------------------------------------------------
If esCalcPromedioMes2 Then
'     StrSql = " SELECT detliq.dlimonto valor "
'     StrSql = StrSql & " FROM detliq "
'     StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And detliq.concnro = " & CalcPromedioMes2
    StrSql = "SELECT detliq.dlimonto valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro=3 AND pliqmes = " & MesNro2
    StrSql = StrSql & " AND pliqanio = " & AnioNro2 & " And detliq.ConcNro = " & CalcPromedioMes2
Else
'     StrSql = " SELECT almonto valor"
'     StrSql = StrSql & " From acu_liq"
'     StrSql = StrSql & " Where acunro = " & CalcPromedioMes2
'     StrSql = StrSql & " AND cliqnro = " & cliqnro
    StrSql = "SELECT sum(almonto) valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro = 3 AND pliqmes = " & MesNro2
    StrSql = StrSql & " AND pliqanio = " & AnioNro2
    StrSql = StrSql & " AND acu_liq.acuNro = " & CalcPromedioMes2
    StrSql = StrSql & " AND empleado = " & Ternro
    StrSql = StrSql & " GROUP BY empleado"


End If
OpenRecordset StrSql, rsConsult
Flog.writeline "CalcPromedioMes2" & StrSql
If Not rsConsult.EOF Then
    CalcPromMes2 = Replace(rsConsult!Valor, ",", ".")
Else
    CalcPromMes2 = 0
End If

'------------------------------------------------------------------
'Obtengo Calculo Promedio Mes 3
'------------------------------------------------------------------
If esCalcPromedioMes3 Then
'     StrSql = " SELECT detliq.dlimonto valor "
'     StrSql = StrSql & " FROM detliq "
'     StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And detliq.concnro = " & CalcPromedioMes3
    StrSql = "SELECT detliq.dlimonto valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro=3 AND pliqmes = " & MesNro3
    StrSql = StrSql & " AND pliqanio = " & AnioNro3 & " And detliq.ConcNro = " & CalcPromedioMes3
Else
'     StrSql = " SELECT almonto valor"
'     StrSql = StrSql & " From acu_liq"
'     StrSql = StrSql & " Where acunro = " & CalcPromedioMes3
'     StrSql = StrSql & " AND cliqnro = " & cliqnro
    StrSql = "SELECT sum(almonto) valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro = 3 AND pliqmes = " & MesNro3
    StrSql = StrSql & " AND pliqanio = " & AnioNro3
    StrSql = StrSql & " AND acu_liq.acuNro = " & CalcPromedioMes3
    StrSql = StrSql & " AND empleado = " & Ternro
    StrSql = StrSql & " GROUP BY empleado"
End If
OpenRecordset StrSql, rsConsult
Flog.writeline "CalcPromedioMes3" & StrSql
If Not rsConsult.EOF Then
    CalcPromMes3 = Replace(rsConsult!Valor, ",", ".")
Else
    CalcPromMes3 = 0
End If

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   Apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
   If IsNull(rsConsult!empremu) Then
      Sueldo = 0
   Else
      Sueldo = Replace(rsConsult!empremu, ",", ".")
   End If
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
Flog.writeline "Buscando los datos de la cuenta del empleado."

StrSql = " SELECT * FROM ctabancaria LEFT JOIN banco ON banco.ternro = ctabancaria.banco WHERE ctabestado=-1 AND ctabancaria.ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
        Banco = rsConsult!Bandesc
        CodBanco = rsConsult!Banco
        nroCuenta = rsConsult!ctabnro
        Flog.writeline "Datos de la cuenta bancaria obtenidos"
  Else
        Banco = ""
        CodBanco = ""
        nroCuenta = ""
        Flog.writeline "El empleado no tiene cuentas bancarias activas"
End If
 
'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & Pronro
       
OpenRecordset StrSql, rsConsult
Dim Des_Mes As String
If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   If pliqmes = 1 Then
        pliqmes = 10
        Call BusMes(10, Des_Mes)
        DescMes1 = Des_Mes
        Call BusMes(11, Des_Mes)
        DescMes2 = Des_Mes
       Call BusMes(12, Des_Mes)
        DescMes3 = Des_Mes
   Else
        If pliqmes = 2 Then
            Call BusMes(11, Des_Mes)
            DescMes1 = Des_Mes
            Call BusMes(12, Des_Mes)
            DescMes2 = Des_Mes
            Call BusMes(1, Des_Mes)
            DescMes3 = Des_Mes
        Else
            If pliqmes = 3 Then
                Call BusMes(12, Des_Mes)
                DescMes1 = Des_Mes
                Call BusMes(2, Des_Mes)
                DescMes2 = Des_Mes
                Call BusMes(1, Des_Mes)
                DescMes3 = Des_Mes
            Else
                Call BusMes((pliqmes - 3), Des_Mes)
                DescMes1 = Des_Mes
                Call BusMes((pliqmes - 2), Des_Mes)
                DescMes2 = Des_Mes
                Call BusMes((pliqmes - 1), Des_Mes)
                DescMes3 = Des_Mes
            End If
        End If
   End If
   pliqanio = rsConsult!pliqanio
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

If tenro1 <> 0 Then
    If estrnro1 <> 0 Then
        EmpEstrnro1 = estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro1
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!Estrnro
        End If
    End If
End If
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

If tenro2 <> 0 Then
    If estrnro2 <> 0 Then
        EmpEstrnro2 = estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro2
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!Estrnro
        End If
    End If
End If
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------
If tenro3 <> 0 Then
    If estrnro3 <> 0 Then
        EmpEstrnro3 = estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro3
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!Estrnro
        End If
    End If
End If
 
'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
        
    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
       Sueldo = Replace(rsConsult!almonto, ",", ".")
    Else
       Flog.writeline "Error al obtener los datos del sueldo. Se queda con el valor de la remuneracion cargada en ADP."
       Sueldo = 0
    End If

'------------------------------------------------------------------
'Busco el valor de los días trabajados
'------------------------------------------------------------------
Flog.writeline "Busco el NIT"

    StrSql = " SELECT nrodoc FROM tipodocu"
    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.tidnro=tipodocu.tidnro"
    StrSql = StrSql & " WHERE tipodocu.tidsigla='NIT' " & "AND ter_doc.ternro = " & Ternro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        NIT = rsConsult!NroDoc
    Else
        NIT = 0
    End If

'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= 5"
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrdabr
Else
    CentroCosto = ""
    Flog.writeline "No se encontró Centro de Costo"
End If

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------
Flog.writeline "Busco los datos de la empresa"
StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = 10"
OpenRecordset StrSql, rs_estructura

EmpEstrnro = 0

If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    Exit Sub
Else
    EmpNombre = rs_estructura!empnom
    EmpEstrnro = rs_estructura!Estrnro
End If

'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc, detdom.piso, detdom.oficdepto From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    EmpDire = "   "
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro
    If Not EsNulo(rs_Domicilio!Piso) Then
        EmpDire = EmpDire & " P. " & rs_Domicilio!Piso
    End If
    If Not EsNulo(rs_Domicilio!oficdepto) Then
        EmpDire = EmpDire & " Dpto. " & rs_Domicilio!oficdepto
    End If
    EmpDire = EmpDire & " - " & rs_Domicilio!locdesc
End If

'Consulta para obtener el cuit de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa"
    'Exit Sub
    EmpCuit = "  "
Else
    EmpCuit = rs_cuit!NroDoc
End If

'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    'Exit Sub
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If

'------------------------------------------------------------------
'Busco el valor del puesto
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & Ternro

OpenRecordset StrSql, rsConsult

Puesto = ""

If Not rsConsult.EOF Then
   Puesto = rsConsult!estrdabr
Else
   Flog.writeline "Error al obtener los datos del puesto"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco la Fecha de Alta del empleado. Corresponde a la fecha desde de la fase que
'contiene la marca de fecha de Alta Reconocida.
'------------------------------------------------------------------
StrSql = " SELECT altfec "
StrSql = StrSql & " FROM fases "
'StrSql = StrSql & " WHERE empleado= " & Ternro '& " ORDER BY altfec ASC "
StrSql = StrSql & " WHERE empleado= " & Ternro & " AND fasrecofec = -1 ORDER BY altfec ASC "
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    empFecAlta = rsConsult!altfec
Else
    Flog.writeline Espacios(Tabulador * 1) & "Error al obtener la Fecha de Alta del Empleado. Se obtiene de la fase marcada como Fecha Alta Reconocida."
    empFecAlta = "&nbsp;"
End If

'------------------------------------------------------------------------------------
'Busco Fecha de Baja
'------------------------------------------------------------------------------------

StrSql = " SELECT bajfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE empleado= " & Ternro
StrSql = StrSql & " ORDER BY bajfec DESC"
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    FechaBaj = rsConsult!bajfec
Else
    FechaBaj = ""
End If

'------------------------------------------------------------------------------------
'Consulta para obtener la Fecha de Nacimiento y el Sexo del Empleado
'------------------------------------------------------------------------------------

StrSql = "SELECT tersex, terfecnac FROM tercero "
StrSql = StrSql & "WHERE ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    sexo = rsConsult!tersex
    FechaNac = rsConsult!terfecnac
    If EsNulo(sexo) Then
        sexo = ""
    Else
        If sexo = -1 Then
            sexo = "M"
        Else
            sexo = "F"
        End If
    End If
Else
    sexo = ""
    FechaNac = ""
End If

'------------------------------------------------------------------
'Calculo la antiguedad
'------------------------------------------------------------------
Antiguedad = ""
Fecha_aux = CDate("01" & "/" & pliqmes & "/" & pliqanio)
Fecha_aux = DateAdd("m", 1, Fecha_aux)
Fecha_aux = DateAdd("d", -1, Fecha_aux)

If Fecha_aux < empFecAlta Then
    Antiguedad = "0 año/s 0 mes/ses"
    anios = 0
    Meses = 0
    Dias = 0
Else
    Call bus_Antiguedad(Ternro, "REAL", Fecha_aux, dias_aux, meses_aux, anios_aux, diasHab_aux)
    'Antiguedad = anios_aux & " año/s " & meses_aux & " mes/es"
    anios = anios_aux
    Meses = meses_aux
    Dias = diasHab_aux
End If

'------------------------------------------------------------------
'Obtengo la fecha de pago del proceso
'------------------------------------------------------------------
'StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini FROM cabliq "
'StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
'StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'   proFecPago = rsConsult!proFecPago
'Else
'   Flog.writeline "Error al obtener los datos del proceso"
'   GoTo MError
'End If

StrSql = "SELECT altfec,fases.* FROM fases"
StrSql = StrSql & " INNER JOIN causa ON fases.caunro = causa.caunro"
StrSql = StrSql & " WHERE fases.Empleado = " & Ternro
StrSql = StrSql & " AND causa.caunro = 13 AND bajfec <= " & ConvFecha(pliqhasta) & " "
StrSql = StrSql & " ORDER BY fases.bajfec desc "
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    proFecPago = rsConsult!bajfec
Else
    proFecPago = " "
    Flog.writeline "Error al obtener Fecha Ultimo Anticipo Indemnizacion " & StrSql
End If

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------
Flog.writeline "Armo la SQL para guardar los datos"

StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,"
StrSql = StrSql & " empfecalta,categoria,"
StrSql = StrSql & " centrocosto,puesto,"
StrSql = StrSql & " profecpago,empnombre,empdire,empcuit,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,prodesc,descripcion, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden,auxchar1,auxchar2,auxchar3,auxchar4,auxchar5,modeloRecibo, "
StrSql = StrSql & " auxdeci1,auxdeci2,auxdeci3,auxdeci4,auxdeci5,auxdeci6,auxchar6,auxchar7,auxchar8)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & ",'" & FechaBaj & "'"
StrSql = StrSql & ",'" & CentroCosto & "'"
StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"      'No se usa
StrSql = StrSql & ",'" & FechaBaj & "'"
StrSql = StrSql & ",'" & EmpFirma & "'"     'No se usa
StrSql = StrSql & "," & EmpFirmaAlto        'No se usa
StrSql = StrSql & "," & EmpFirmaAncho       'No se usa
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"   'No se usa Descripcion
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & CodBanco & "'"
StrSql = StrSql & ",'" & Banco & "'"
StrSql = StrSql & ",'" & nroCuenta & "'"
StrSql = StrSql & ",'" & sexo & "'"
StrSql = StrSql & ",'" & FechaNac & "'"
StrSql = StrSql & "," & tipoRecibo
StrSql = StrSql & "," & anios
StrSql = StrSql & "," & Meses
StrSql = StrSql & "," & Dias
StrSql = StrSql & "," & CalcPromMes1
StrSql = StrSql & "," & CalcPromMes2
StrSql = StrSql & "," & CalcPromMes3
StrSql = StrSql & ",'" & DescMes1 & "'"
StrSql = StrSql & ",'" & DescMes2 & "'"
StrSql = StrSql & ",'" & DescMes3 & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

Flog.writeline "SQL INSERT: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado
'------------------------------------------------------------------

StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
   
OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
  
    StrSql = " INSERT INTO rep_recibo_det "
    StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
    StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
    StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & "," & Pronro
    StrSql = StrSql & "," & rsConsult!cliqnro
    StrSql = StrSql & ",'" & rsConsult!concabr & "'"
    StrSql = StrSql & ",'" & rsConsult!ConcCod & "'"
    StrSql = StrSql & "," & rsConsult!ConcNro
    StrSql = StrSql & "," & rsConsult!tconnro
    StrSql = StrSql & "," & rsConsult!concimp
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
    StrSql = StrSql & ",'" & arrTipoConc(rsConsult!tconnro) & "')"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rsConsult.MoveNext

Loop

rsConsult.Close

Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub

'--------------------------------------------------------------------
' Se encarga de generar el recibo Quinquenios de Bolivia Modelo 505
'--------------------------------------------------------------------
Sub generarDatosRecibo505(Pronro, Ternro, acunroSueldo, tituloReporte, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim Nombre
Dim Direccion
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim empFecAlta
Dim Sueldo
Dim SalarioNeto
Dim CentroCosto
Dim localidad
Dim proFecPago
Dim pliqhasta
Dim Banco
Dim CodBanco
Dim nroCuenta
Dim NIT
Dim sexo
Dim FechaNac
Dim Puesto
Dim FechaBaj
Dim CalcPromedioMes1
Dim CalcPromedioMes2
Dim CalcPromedioMes3
Dim MesNro As Integer
Dim MesActual
Dim AnioActual

Dim MesNro1 As Integer
Dim AnioNro1 As Integer
Dim MesNro2 As Integer
Dim AnioNro2 As Integer
Dim MesNro3 As Integer
Dim AnioNro3 As Integer

Dim CalcPromMes1
Dim CalcPromMes2
Dim CalcPromMes3

Dim esCalcPromedioMes1
Dim esCalcPromedioMes2
Dim esCalcPromedioMes3

Dim anios
Dim Meses
Dim Dias

Dim Fecha_aux As Date
Dim anios_aux As Integer
Dim meses_aux As Integer
Dim dias_aux As Integer
Dim diasHab_aux As Integer

Dim EmpEstrnro
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String
Dim ColumnasConfrep As String
 
Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3
Dim ObraSocial
Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset
Dim FechaUltimoAnticipo As String

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

Flog.writeline "Modelo 505"

'------------------------------------------------------------------
'Obtengo el nro de cabecera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini,pliqmes,pliqanio FROM cabliq"
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
       
OpenRecordset StrSql, rsConsult
Flog.writeline "Nro de Cabecera de Liquidacion" & StrSql

If Not rsConsult.EOF Then
   cliqnro = rsConsult!cliqnro
   proFecPago = rsConsult!proFecPago
   profecini = rsConsult!profecini
   proDesc = rsConsult!proDesc
   MesActual = rsConsult!pliqmes
   AnioActual = rsConsult!pliqanio
Else
   Flog.writeline "Error al obtener los datos del proceso"
   GoTo MError
End If
'   MesNro1 = MesActual
'   AnioNro1 = AnioActual
'   If MesActual = 1 Then
'        Call PeriodoAnterior(MesActual, MesNro)
'        MesNro2 = MesNro
'        AnioNro2 = AnioActual - 1
'        Call PeriodoAnterior(MesNro2, MesNro)
'        MesNro3 = MesNro
'        AnioNro3 = AnioActual - 1
'   Else
'        If pliqmes = 2 Then
'            Call PeriodoAnterior(MesActual, MesNro)
'            MesNro2 = MesNro
'            AnioNro2 = AnioActual
'            Call PeriodoAnterior(MesNro2, MesNro)
'            MesNro3 = MesNro
'            AnioNro3 = AnioActual - 1
'        Else
'           Call PeriodoAnterior(MesActual, MesNro)
'            MesNro2 = MesNro
'            AnioNro2 = AnioActual
'            Call PeriodoAnterior(MesNro2, MesNro)
'            MesNro3 = MesNro
'            AnioNro3 = AnioActual
'        End If
'   End If

Call PeriodosAnteriores(MesActual, AnioActual, MesNro1, AnioNro1, MesNro2, AnioNro2, MesNro3, AnioNro3)

'Obtengo configuración de reporte
Dim rs_GetConfrepRs As New ADODB.Recordset
StrSql = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & tipoRecibo * 100
StrSql = StrSql & " AND confnrocol < " & (tipoRecibo * 100) + 100 & ")"
StrSql = StrSql & " ORDER BY confnrocol ASC "
'GetConfrepRs tipoRecibo
OpenRecordset StrSql, rs_GetConfrepRs
Flog.writeline "confrep" & StrSql
If Not rs_GetConfrepRs.EOF Then
    Do While Not rs_GetConfrepRs.EOF
        Select Case rs_GetConfrepRs!confnrocol
            Case 50500: 'Calculo Promedio
                        If rs_GetConfrepRs!conftipo = "CO" Then
                            esCalcPromedioMes1 = True
                            CalcPromedioMes1 = rs_GetConfrepRs!confval
                        Else
                            esCalcPromedioMes1 = False
                            CalcPromedioMes1 = rs_GetConfrepRs!confval2
                        End If
            Case 50501: 'Calculo Promedio
                        If rs_GetConfrepRs!conftipo = "CO" Then
                            esCalcPromedioMes2 = True
                            CalcPromedioMes2 = rs_GetConfrepRs!confval
                        Else
                            esCalcPromedioMes2 = False
                            CalcPromedioMes2 = rs_GetConfrepRs!confval2
                        End If
            Case 50502: 'Calculo Promedio
                        If rs_GetConfrepRs!conftipo = "CO" Then
                            esCalcPromedioMes3 = True
                            CalcPromedioMes3 = rs_GetConfrepRs!confval
                        Else
                            esCalcPromedioMes3 = False
                            CalcPromedioMes3 = rs_GetConfrepRs!confval2
                        End If
           Case 50513: 'Fecha Ultimo Anticipo
                        If rs_GetConfrepRs!conftipo = "FUN" Then
                            FechaUltimoAnticipo = rs_GetConfrepRs!confval2
                        End If
        End Select
         rs_GetConfrepRs.MoveNext
    Loop
Else
    Flog.writeline "No se encontró configuración para el Modelo 505"
    'Exit Sub
End If
rs_GetConfrepRs.Close


'------------------------------------------------------------------
'Obtengo Calculo Promedio Mes 1
'------------------------------------------------------------------
If esCalcPromedioMes1 Then
'     StrSql = " SELECT detliq.dlimonto valor "
'     StrSql = StrSql & " FROM detliq "
'     StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And detliq.concnro = " & CalcPromedioMes1
    StrSql = "SELECT detliq.dlimonto valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro=3 AND pliqmes = " & MesNro1
    StrSql = StrSql & " AND pliqanio = " & AnioNro1 & " And detliq.ConcNro = " & CalcPromedioMes1
Else
'     StrSql = " SELECT almonto valor"
'     StrSql = StrSql & " From acu_liq"
'     StrSql = StrSql & " Where acunro = " & CalcPromedioMes1
'     StrSql = StrSql & " AND cliqnro = " & cliqnro
    StrSql = "SELECT sum(almonto) valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro = 3 AND pliqmes = " & MesNro1
    StrSql = StrSql & " AND pliqanio = " & AnioNro1
    StrSql = StrSql & " AND acu_liq.acuNro = " & CalcPromedioMes1
    StrSql = StrSql & " AND empleado = " & Ternro
    StrSql = StrSql & " GROUP BY empleado"
End If
OpenRecordset StrSql, rsConsult
Flog.writeline "CalcPromedioMes1" & StrSql
If Not rsConsult.EOF Then
    CalcPromMes1 = Replace(rsConsult!Valor, ",", ".")
Else
    CalcPromMes1 = 0
End If

'------------------------------------------------------------------
'Obtengo Calculo Promedio Mes 2
'------------------------------------------------------------------
If esCalcPromedioMes2 Then
'     StrSql = " SELECT detliq.dlimonto valor "
'     StrSql = StrSql & " FROM detliq "
'     StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And detliq.concnro = " & CalcPromedioMes2
    StrSql = "SELECT detliq.dlimonto valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro=3 AND pliqmes = " & MesNro2
    StrSql = StrSql & " AND pliqanio = " & AnioNro2 & " And detliq.ConcNro = " & CalcPromedioMes2
Else
'     StrSql = " SELECT almonto valor"
'     StrSql = StrSql & " From acu_liq"
'     StrSql = StrSql & " Where acunro = " & CalcPromedioMes2
'     StrSql = StrSql & " AND cliqnro = " & cliqnro
    StrSql = "SELECT sum(almonto) valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro = 3 AND pliqmes = " & MesNro2
    StrSql = StrSql & " AND pliqanio = " & AnioNro2
    StrSql = StrSql & " AND acu_liq.acuNro = " & CalcPromedioMes2
    StrSql = StrSql & " AND empleado = " & Ternro
    StrSql = StrSql & " GROUP BY empleado"


End If
OpenRecordset StrSql, rsConsult
Flog.writeline "CalcPromedioMes2" & StrSql
If Not rsConsult.EOF Then
    CalcPromMes2 = Replace(rsConsult!Valor, ",", ".")
Else
    CalcPromMes2 = 0
End If

'------------------------------------------------------------------
'Obtengo Calculo Promedio Mes 3
'------------------------------------------------------------------
If esCalcPromedioMes3 Then
'     StrSql = " SELECT detliq.dlimonto valor "
'     StrSql = StrSql & " FROM detliq "
'     StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And detliq.concnro = " & CalcPromedioMes3
    StrSql = "SELECT detliq.dlimonto valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro=3 AND pliqmes = " & MesNro3
    StrSql = StrSql & " AND pliqanio = " & AnioNro3 & " And detliq.ConcNro = " & CalcPromedioMes3
Else
'     StrSql = " SELECT almonto valor"
'     StrSql = StrSql & " From acu_liq"
'     StrSql = StrSql & " Where acunro = " & CalcPromedioMes3
'     StrSql = StrSql & " AND cliqnro = " & cliqnro
    StrSql = "SELECT sum(almonto) valor FROM proceso "
    StrSql = StrSql & "INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & "INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & "INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & "WHERE tprocnro = 3 AND pliqmes = " & MesNro3
    StrSql = StrSql & " AND pliqanio = " & AnioNro3
    StrSql = StrSql & " AND acu_liq.acuNro = " & CalcPromedioMes3
    StrSql = StrSql & " AND empleado = " & Ternro
    StrSql = StrSql & " GROUP BY empleado"
End If
OpenRecordset StrSql, rsConsult
Flog.writeline "CalcPromedioMes3" & StrSql
If Not rsConsult.EOF Then
    CalcPromMes3 = Replace(rsConsult!Valor, ",", ".")
Else
    CalcPromMes3 = 0
End If

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   Apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
   If IsNull(rsConsult!empremu) Then
      Sueldo = 0
   Else
      Sueldo = Replace(rsConsult!empremu, ",", ".")
   End If
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
Flog.writeline "Buscando los datos de la cuenta del empleado."

StrSql = " SELECT * FROM ctabancaria LEFT JOIN banco ON banco.ternro = ctabancaria.banco WHERE ctabestado=-1 AND ctabancaria.ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
        Banco = rsConsult!Bandesc
        CodBanco = rsConsult!Banco
        nroCuenta = rsConsult!ctabnro
        Flog.writeline "Datos de la cuenta bancaria obtenidos"
  Else
        Banco = ""
        CodBanco = ""
        nroCuenta = ""
        Flog.writeline "El empleado no tiene cuentas bancarias activas"
End If
 
'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & Pronro
       
OpenRecordset StrSql, rsConsult
Dim Des_Mes As String
If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   If pliqmes = 1 Then
        pliqmes = 10
        Call BusMes(10, Des_Mes)
        DescMes1 = Des_Mes
        Call BusMes(11, Des_Mes)
        DescMes2 = Des_Mes
       Call BusMes(12, Des_Mes)
        DescMes3 = Des_Mes
   Else
        If pliqmes = 2 Then
            Call BusMes(11, Des_Mes)
            DescMes1 = Des_Mes
            Call BusMes(12, Des_Mes)
            DescMes2 = Des_Mes
            Call BusMes(1, Des_Mes)
            DescMes3 = Des_Mes
        Else
            If pliqmes = 3 Then
                Call BusMes(12, Des_Mes)
                DescMes1 = Des_Mes
                Call BusMes(2, Des_Mes)
                DescMes2 = Des_Mes
                Call BusMes(1, Des_Mes)
                DescMes3 = Des_Mes
            Else
                Call BusMes((pliqmes - 3), Des_Mes)
                DescMes1 = Des_Mes
                Call BusMes((pliqmes - 2), Des_Mes)
                DescMes2 = Des_Mes
                Call BusMes((pliqmes - 1), Des_Mes)
                DescMes3 = Des_Mes
            End If
        End If
   End If
   pliqanio = rsConsult!pliqanio
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

If tenro1 <> 0 Then
    If estrnro1 <> 0 Then
        EmpEstrnro1 = estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro1
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!Estrnro
        End If
    End If
End If
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

If tenro2 <> 0 Then
    If estrnro2 <> 0 Then
        EmpEstrnro2 = estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro2
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!Estrnro
        End If
    End If
End If
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------
If tenro3 <> 0 Then
    If estrnro3 <> 0 Then
        EmpEstrnro3 = estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro3
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!Estrnro
        End If
    End If
End If
 
'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
        
    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
       Sueldo = Replace(rsConsult!almonto, ",", ".")
    Else
       Flog.writeline "Error al obtener los datos del sueldo. Se queda con el valor de la remuneracion cargada en ADP."
       Sueldo = 0
    End If

'------------------------------------------------------------------
'Busco el valor de los días trabajados
'------------------------------------------------------------------
Flog.writeline "Busco el NIT"

    StrSql = " SELECT nrodoc FROM tipodocu"
    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.tidnro=tipodocu.tidnro"
    StrSql = StrSql & " WHERE tipodocu.tidsigla='NIT' " & "AND ter_doc.ternro = " & Ternro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        NIT = rsConsult!NroDoc
    Else
        NIT = 0
    End If

'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ")"
StrSql = StrSql & " AND his_estructura.tenro= 5"
StrSql = StrSql & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrdabr
Else
    CentroCosto = ""
    Flog.writeline "No se encontró Centro de Costo"
End If

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------
Flog.writeline "Busco los datos de la empresa"
StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = 10"
OpenRecordset StrSql, rs_estructura

EmpEstrnro = 0

If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    Exit Sub
Else
    EmpNombre = rs_estructura!empnom
    EmpEstrnro = rs_estructura!Estrnro
End If

'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc, detdom.piso, detdom.oficdepto From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    EmpDire = "   "
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro
    If Not EsNulo(rs_Domicilio!Piso) Then
        EmpDire = EmpDire & " P. " & rs_Domicilio!Piso
    End If
    If Not EsNulo(rs_Domicilio!oficdepto) Then
        EmpDire = EmpDire & " Dpto. " & rs_Domicilio!oficdepto
    End If
    EmpDire = EmpDire & " - " & rs_Domicilio!locdesc
End If

'Consulta para obtener el cuit de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa"
    'Exit Sub
    EmpCuit = "  "
Else
    EmpCuit = rs_cuit!NroDoc
End If

'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    'Exit Sub
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If

'------------------------------------------------------------------
'Busco el valor del puesto
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & Ternro

OpenRecordset StrSql, rsConsult

Puesto = ""

If Not rsConsult.EOF Then
   Puesto = rsConsult!estrdabr
Else
   Flog.writeline "Error al obtener los datos del puesto"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco la Fecha de Alta del empleado. Corresponde a la fecha desde de la fase que
'contiene la marca de fecha de Alta Reconocida.
'------------------------------------------------------------------
StrSql = " SELECT altfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE empleado= " & Ternro & " AND fasrecofec = -1 ORDER BY altfec ASC "
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    empFecAlta = rsConsult!altfec
Else
    Flog.writeline Espacios(Tabulador * 1) & "Error al obtener la Fecha de Alta del Empleado. Se obtiene de la fase marcada como Fecha Alta Reconocida."
    empFecAlta = "&nbsp;"
End If

'------------------------------------------------------------------------------------
'Busco Fecha de Baja
'------------------------------------------------------------------------------------

StrSql = " SELECT bajfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE empleado= " & Ternro
StrSql = StrSql & " ORDER BY bajfec DESC"
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    FechaBaj = rsConsult!bajfec
Else
    FechaBaj = ""
End If

'------------------------------------------------------------------------------------
'Consulta para obtener la Fecha de Nacimiento y el Sexo del Empleado
'------------------------------------------------------------------------------------

StrSql = "SELECT tersex, terfecnac FROM tercero "
StrSql = StrSql & "WHERE ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    sexo = rsConsult!tersex
    FechaNac = rsConsult!terfecnac
    If EsNulo(sexo) Then
        sexo = ""
    Else
        If sexo = -1 Then
            sexo = "M"
        Else
            sexo = "F"
        End If
    End If
Else
    sexo = ""
    FechaNac = ""
End If

'------------------------------------------------------------------
'Calculo la antiguedad
'------------------------------------------------------------------
Antiguedad = ""
Fecha_aux = CDate("01" & "/" & pliqmes & "/" & pliqanio)
Fecha_aux = DateAdd("m", 1, Fecha_aux)
Fecha_aux = DateAdd("d", -1, Fecha_aux)

If Fecha_aux < empFecAlta Then
    Antiguedad = "0 año/s 0 mes/ses"
    anios = 0
    Meses = 0
    Dias = 0
Else
    Call bus_Antiguedad(Ternro, "REAL", Fecha_aux, dias_aux, meses_aux, anios_aux, diasHab_aux)
    'Antiguedad = anios_aux & " año/s " & meses_aux & " mes/es"
    anios = anios_aux
    Meses = meses_aux
    Dias = diasHab_aux
End If

'------------------------------------------------------------------
'Obtengo la fecha de pago del proceso
'------------------------------------------------------------------
'StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini FROM cabliq "
'StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
'StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'   proFecPago = rsConsult!proFecPago
'Else
'   Flog.writeline "Error al obtener los datos del proceso"
'   GoTo MError
'End If

StrSql = "SELECT altfec,fases.* FROM fases"
StrSql = StrSql & " INNER JOIN causa ON fases.caunro = causa.caunro"
StrSql = StrSql & " WHERE fases.Empleado = " & Ternro
StrSql = StrSql & " AND causa.caunro = 13 AND bajfec <= " & ConvFecha(pliqhasta) & " "
StrSql = StrSql & " ORDER BY fases.bajfec desc "
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    proFecPago = rsConsult!bajfec
Else
    proFecPago = " "
    Flog.writeline "Error al obtener Fecha Ultimo Anticipo Indemnizacion " & StrSql
End If


'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------
Flog.writeline "Armo la SQL para guardar los datos"

StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,"
StrSql = StrSql & " empfecalta,categoria,"
StrSql = StrSql & " centrocosto,puesto,"
StrSql = StrSql & " profecpago,empnombre,empdire,empcuit,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,prodesc,descripcion, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden,auxchar1,auxchar2,auxchar3,auxchar4,auxchar5,modeloRecibo, "
StrSql = StrSql & " auxdeci1,auxdeci2,auxdeci3,auxdeci4,auxdeci5,auxdeci6,auxchar6,auxchar7,auxchar8,auxchar9)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & ",'" & FechaBaj & "'"
StrSql = StrSql & ",'" & CentroCosto & "'"
StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"      'No se usa
StrSql = StrSql & ",'" & FechaBaj & "'"
StrSql = StrSql & ",'" & EmpFirma & "'"     'No se usa
StrSql = StrSql & "," & EmpFirmaAlto        'No se usa
StrSql = StrSql & "," & EmpFirmaAncho       'No se usa
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"   'No se usa Descripcion
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & CodBanco & "'"
StrSql = StrSql & ",'" & Banco & "'"
StrSql = StrSql & ",'" & nroCuenta & "'"
StrSql = StrSql & ",'" & sexo & "'"
StrSql = StrSql & ",'" & FechaNac & "'"
StrSql = StrSql & "," & tipoRecibo
StrSql = StrSql & "," & anios
StrSql = StrSql & "," & Meses
StrSql = StrSql & "," & Dias
StrSql = StrSql & "," & CalcPromMes1
StrSql = StrSql & "," & CalcPromMes2
StrSql = StrSql & "," & CalcPromMes3
StrSql = StrSql & ",'" & DescMes1 & "'"
StrSql = StrSql & ",'" & DescMes2 & "'"
StrSql = StrSql & ",'" & DescMes3 & "'"
StrSql = StrSql & ",'" & FechaUltimoAnticipo & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

Flog.writeline "SQL INSERT: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado
'------------------------------------------------------------------

StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
   
OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
  
    StrSql = " INSERT INTO rep_recibo_det "
    StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
    StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
    StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & "," & Pronro
    StrSql = StrSql & "," & rsConsult!cliqnro
    StrSql = StrSql & ",'" & rsConsult!concabr & "'"
    StrSql = StrSql & ",'" & rsConsult!ConcCod & "'"
    StrSql = StrSql & "," & rsConsult!ConcNro
    StrSql = StrSql & "," & rsConsult!tconnro
    StrSql = StrSql & "," & rsConsult!concimp
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
    StrSql = StrSql & ",'" & arrTipoConc(rsConsult!tconnro) & "')"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rsConsult.MoveNext

Loop

rsConsult.Close

Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub

Public Sub GetConfrepRs(ByVal Modelo As Long)
'Public Function GetConfrepRs(ByVal modelo As Long) As ADODB.Recordset
'-----------------------------------------------------------------------------------
' - || Esta función se utiliza a partir de los modelos 500 || -
'-----------------------------------------------------------------------------------
' Descripcion: Devuelve Recordset con el total de filas configuradas para el modelo en confrep
'              Ejemplo: para el modelo de recibo 500 devuelve columnas desde confnrocol 50000 hasta 50999
' Autor      : Gonzalez Nicolás
' Fecha      : 19/02/2014
' Ultima Mod :
' Descripcion:
'-----------------------------------------------------------------------------------

    'Dim rs_GetConfrepRs As New ADODB.Recordset
    Dim StrSqlConfrep As String
    StrSqlConfrep = " SELECT * FROM confrep WHERE repnro = 60 AND (confnrocol >= " & Modelo * 100
    StrSqlConfrep = StrSqlConfrep & " AND confnrocol < " & (tipoRecibo * 100) + 100 & ")"
    StrSqlConfrep = StrSqlConfrep & " ORDER BY confnrocol ASC "
    'OpenRecordset StrSqlConfrep, rs_GetConfrepRs
    
End Sub


'--------------------------------------------------------------------


