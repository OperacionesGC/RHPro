Attribute VB_Name = "repRecibos4"

'--------------------------------------------------------------------
'Recibo Dalkia Modelo 182
'--------------------------------------------------------------------
Sub generarDatosRecibo182(Pronro, Ternro, acunroSueldo, tituloReporte, orden, zonaDomicilio)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim cliqnro
Dim profecini As String
Dim Cont As Long

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim Nombre
Dim Direccion
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim Cuil
Dim empFecAlta
Dim Sueldo
Dim Categoria
Dim CentroCosto
Dim Localidad
Dim proFecPago
Dim pliqhasta
Dim pliqdesde
Dim FormaPago
Dim Puesto
Dim ValorHora
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
Dim oSocial

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3
Dim EmpTernro

Dim DiasVacTomados
Dim DivPersonal
Dim Departamento
Dim LugarDeTrabajo
Dim Banco
Dim CuentaBancaria
Dim PagoMonto
Dim FormaDePagoNro

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

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
      Sueldo = 0
   Else
      Sueldo = rsConsult!empremu
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
   pliqdesde = rsConsult!pliqdesde
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
'Busco el valor del cuil
'------------------------------------------------------------------
StrSql = " SELECT cuil.nrodoc "
StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Cuil = rsConsult!NroDoc
Else
'   Flog.writeline "Error al obtener los datos del cuil"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor de la direccion y localidad
'------------------------------------------------------------------

Direccion = ""
Localidad = ""

Select Case zonaDomicilio
   'Direccion de la sucursal
   Case 1

        StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc"
        StrSql = StrSql & " From his_estructura"
        StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(proFecPago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(proFecPago) & ") AND tenro=1 AND his_estructura.ternro=" & Ternro
        StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
        StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
    
        '---LOG---
        Flog.writeline "Buscando datos de la direccion de la sucursal"
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           Direccion = rsConsult!calle & " " & rsConsult!nro & ", " & rsConsult!locdesc
           Localidad = rsConsult!locdesc
        End If
        
        rsConsult.Close
    
  'Direccion de la empresa
  Case 2

        StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
            " From his_estructura" & _
            " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
            " WHERE his_estructura.htetdesde <=" & ConvFecha(proFecPago) & " AND " & _
            " (his_estructura.htethasta >= " & ConvFecha(proFecPago) & " OR his_estructura.htethasta IS NULL)" & _
            " AND his_estructura.ternro = " & Ternro & _
            " AND his_estructura.tenro  = 10"
        
        OpenRecordset StrSql, rsConsult
        
        EmpTernro = 0
        
        If Not rsConsult.EOF Then
            EmpTernro = rsConsult!Ternro
        End If
        
        rsConsult.Close
        
        'Consulta para obtener la direccion de la empresa
        StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc From cabdom " & _
            " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & EmpTernro & _
            " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
        
        '---LOG---
        Flog.writeline "Buscando datos de la direccion de la empresa"
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           Direccion = rsConsult!calle & " " & rsConsult!nro & ", " & rsConsult!locdesc
           Localidad = rsConsult!locdesc
        End If
        
        rsConsult.Close
    
  'Direccion del empleado
  Case 3

        StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,detdom.piso,detdom.oficdepto,detdom.barrio"
        StrSql = StrSql & " FROM  cabdom "
        StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
        StrSql = StrSql & " WHERE  cabdom.domdefault = -1 AND cabdom.ternro = " & Ternro
       
        '---LOG---
        Flog.writeline "Buscando datos de la direccion del empleado"
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           Direccion = rsConsult!calle & " " & rsConsult!nro & ", " & rsConsult!locdesc
           Localidad = rsConsult!locdesc
        End If
        
        rsConsult.Close
    
End Select

'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

If Sueldo = 0 Then
    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
    
    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
       Sueldo = rsConsult!almonto
    Else
       Flog.writeline "Error al obtener los datos del sueldo"
       Sueldo = 0
       'GoTo MError
    End If
End If

'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------

StrSql = " SELECT estrdabr, estrcodext, estructura.estrnro "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 3 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

Categoria = ""
If Not rsConsult.EOF Then
   Categoria = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la categoria"
'   GoTo MError
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
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If


'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor de la Division de personal (Gerencia)
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=6 AND his_estructura.ternro=" & Ternro
       
OpenRecordset StrSql, rsConsult

DivPersonal = ""
If Not rsConsult.EOF Then
   DivPersonal = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la div de personal"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del Departamento (Sector)
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=2 AND his_estructura.ternro=" & Ternro
       
OpenRecordset StrSql, rsConsult

Departamento = ""
If Not rsConsult.EOF Then
   Departamento = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del departamento"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del Lugar de Trabajo(Sucursal)
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=1 AND his_estructura.ternro=" & Ternro
       
OpenRecordset StrSql, rsConsult

LugarDeTrabajo = ""
If Not rsConsult.EOF Then
   LugarDeTrabajo = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del Lugar de Trabajo"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
StrSql = " SELECT pago.fpagnro, ctabnro,ctabcbu,bandesc,pago.pagomonto,fpagdescabr "
StrSql = StrSql & " From pago"
StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro AND pago.pagorigen=" & cliqnro
StrSql = StrSql & " LEFT JOIN banco     ON banco.ternro = pago.banternro"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   FormaPago = rsConsult!fpagdescabr
   If Not EsNulo(rsConsult!ctabnro) Then
       CuentaBancaria = rsConsult!ctabnro
       If IsNull(rsConsult!Bandesc) Then
          Banco = ""
       Else
          Banco = rsConsult!Bandesc
       End If
   Else
       CuentaBancaria = rsConsult!ctabcbu
       Banco = "CBU"
   End If
   PagoMonto = rsConsult!PagoMonto
   FormaDePagoNro = rsConsult!fpagnro
Else
   Flog.writeline "Fin de archivo en Formas de Pago"
   FormaPago = ""
   CuentaBancaria = ""
   PagoMonto = 0
   FormaDePagoNro = 1000
   Banco = ""
'   Flog.writeline "Error al obtener los datos de la forma de pago"
'   GoTo MError
End If

Flog.writeline "Contenido de CuentaBancaria = " & CuentaBancaria

rsConsult.Close

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

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
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc,codigopostal,barrio From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    EmpDire = "   "
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro & " - " & rs_Domicilio!codigopostal & " " & rs_Domicilio!barrio & " - " & rs_Domicilio!locdesc
    Localidad = rs_Domicilio!locdesc
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

'------------------------------------------------------------------------------------------------------
' Busco las vacaciones pendientes
'------------------------------------------------------------------------------------------------------

StrSql = " SELECT DISTINCT * "
StrSql = StrSql & " FROM novemp "
StrSql = StrSql & " WHERE novemp.empleado = " & Ternro & " AND novemp.tpanro = " & paramVacPend
StrSql = StrSql & " AND concnro = " & concnroVacPend
StrSql = StrSql & " AND nedesde <= " & ConvFecha(pliqhasta)
StrSql = StrSql & " ORDER BY nedesde DESC "

OpenRecordset StrSql, rsConsult

DiasVacTomados = 0
If Not rsConsult.EOF Then
    DiasVacTomados = rsConsult!nevalor
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la obra social elegida
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 17 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

oSocial = ""

If Not rsConsult.EOF Then
   oSocial = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor de la hora según la escala
'------------------------------------------------------------------

StrSql = "SELECT vgrvalor from ("
StrSql = StrSql & "  SELECT * from("
StrSql = StrSql & "  SELECT tipoestructura.tenro as tenroA, tedabr as tedabrA, estrdabr as estrdabrA, teorden as teordenA, "
StrSql = StrSql & "  estructura.estrnro as estrnroA,his_estructura.ternro as ternroA "
StrSql = StrSql & " FROM his_estructura INNER JOIN tipoestructura "
StrSql = StrSql & " ON his_estructura.tenro = tipoestructura.tenro INNER JOIN estructura"
StrSql = StrSql & " ON estructura.estrnro = his_estructura.estrnro INNER JOIN claseestructura"
StrSql = StrSql & " ON tipoestructura.cenro = claseestructura.cenro LEFT JOIN adptte_estr "
StrSql = StrSql & " ON tipoestructura.tenro = adptte_estr.tenro AND tplatenro = 30"
StrSql = StrSql & " WHERE his_estructura.tenro =3 ) a"
StrSql = StrSql & " inner join ("
StrSql = StrSql & " SELECT tipoestructura.tenro as tenroB, tedabr as tedabrB, estrdabr as estrdabrB, teorden as teordenB,"
StrSql = StrSql & " estructura.estrnro as estrnroB,his_estructura.ternro as ternroB"
StrSql = StrSql & " FROM his_estructura INNER JOIN tipoestructura"
StrSql = StrSql & " ON his_estructura.tenro = tipoestructura.tenro INNER JOIN estructura"
StrSql = StrSql & " ON estructura.estrnro = his_estructura.estrnro INNER JOIN claseestructura"
StrSql = StrSql & " ON tipoestructura.cenro = claseestructura.cenro LEFT JOIN adptte_estr"
StrSql = StrSql & " ON tipoestructura.tenro = adptte_estr.tenro AND tplatenro = 30"
StrSql = StrSql & " WHERE his_estructura.tenro =19 and estructura.estrnro=627  ) b"
StrSql = StrSql & "  on a.ternroA = b.ternroB ) c "
StrSql = StrSql & " join (select * from valgrilla where cgrnro = 22 and vgrcoor_1= 627) d"
StrSql = StrSql & " on d.vgrcoor_1 =c.estrnrob and d.vgrcoor_2 =c.estrnroa"
StrSql = StrSql & " where c.ternroA = " & Ternro

       
OpenRecordset StrSql, rsConsult

ValorHora = 0
If Not rsConsult.EOF Then
   ValorHora = rsConsult!vgrvalor
Else
   ValorHora = 0
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
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden, auxdeci1, auxchar1, auxchar2, auxchar3,"
StrSql = StrSql & " auxchar4,auxchar5,auxdeci2,auxdeci3,auxdeci4, auxchar6)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & numberForSQL(Sueldo)
StrSql = StrSql & ",'" & Mid(Categoria, 1, 20) & "'"
StrSql = StrSql & ",'" & Mid(CentroCosto, 1, 25) & "'"
StrSql = StrSql & ",'" & Mid(Localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & FormaPago & "'"
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden
StrSql = StrSql & "," & numberForSQL(DiasVacTomados)
StrSql = StrSql & ",'" & DivPersonal & "'"
StrSql = StrSql & ",'" & Departamento & "'"
StrSql = StrSql & ",'" & LugarDeTrabajo & "'"
StrSql = StrSql & ",'" & CuentaBancaria & "'"
StrSql = StrSql & ",'" & Banco & "'"
StrSql = StrSql & "," & numberForSQL(PagoMonto)
StrSql = StrSql & "," & numberForSQL(arrFormaPago(FormaDePagoNro))
StrSql = StrSql & "," & numberForSQL(ValorHora)
StrSql = StrSql & ",'" & Mid(oSocial, 1, 100) & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'---------------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado, de la columna 1
'---------------------------------------------------------------------------

For Cont = 1 To cantAcumGrupo1
    
    StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
    StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
    StrSql = StrSql & " INNER JOIN con_acum ON concepto.concnro = con_acum.concnro AND con_acum.acunro = " & acumGrupo1(Cont)
    StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
        
    OpenRecordset StrSql, rsConsult
    
    Do Until rsConsult.EOF
      
        StrSql = " SELECT ternro FROM rep_recibo_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND ternro = " & Ternro
        StrSql = StrSql & " AND pronro = " & Pronro
        StrSql = StrSql & " AND concnro = " & rsConsult!ConcNro
        
        OpenRecordset StrSql, rsConsult2
        
        If rsConsult2.EOF Then
            
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
              StrSql = StrSql & "," & acumGrupo1(Cont)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
              StrSql = StrSql & ",'1')"
              
              objConn.Execute StrSql, , adExecuteNoRecords
              
        End If
        
        rsConsult2.Close
        
        rsConsult.MoveNext
    
    Loop
    
    rsConsult.Close

Next

'---------------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado, de la columna 2
'---------------------------------------------------------------------------

For Cont = 1 To cantAcumGrupo2
    
    StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
    StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
    StrSql = StrSql & " INNER JOIN con_acum ON concepto.concnro = con_acum.concnro AND con_acum.acunro = " & acumGrupo2(Cont)
    StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
        
    OpenRecordset StrSql, rsConsult
    
    Do Until rsConsult.EOF
      
        StrSql = " SELECT ternro FROM rep_recibo_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND ternro = " & Ternro
        StrSql = StrSql & " AND pronro = " & Pronro
        StrSql = StrSql & " AND concnro = " & rsConsult!ConcNro
        
        OpenRecordset StrSql, rsConsult2
        
        If rsConsult2.EOF Then
            
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
              StrSql = StrSql & "," & acumGrupo2(Cont)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
              StrSql = StrSql & ",'2')"
              
              objConn.Execute StrSql, , adExecuteNoRecords
              
        End If
        
        rsConsult2.Close
        
        rsConsult.MoveNext
    
    Loop
    
    rsConsult.Close

Next


'---------------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado, de la columna 3
'---------------------------------------------------------------------------

For Cont = 1 To cantAcumGrupo3
    
    StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
    StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
    StrSql = StrSql & " INNER JOIN con_acum ON concepto.concnro = con_acum.concnro AND con_acum.acunro = " & acumGrupo3(Cont)
    StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
        
    OpenRecordset StrSql, rsConsult
    
    Do Until rsConsult.EOF
      
        StrSql = " SELECT ternro FROM rep_recibo_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND ternro = " & Ternro
        StrSql = StrSql & " AND pronro = " & Pronro
        StrSql = StrSql & " AND concnro = " & rsConsult!ConcNro
        
        OpenRecordset StrSql, rsConsult2
        
        If rsConsult2.EOF Then
            
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
              StrSql = StrSql & "," & acumGrupo3(Cont)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
              StrSql = StrSql & ",'3')"
              
              objConn.Execute StrSql, , adExecuteNoRecords
              
        End If
        
        rsConsult2.Close
        
        rsConsult.MoveNext
    
    Loop
    
    rsConsult.Close

Next

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
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub generarDatosRecibo183(Pronro, Ternro, acunroSueldo, tituloReporte, orden)

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
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim Cuil
Dim empFecAlta
Dim Sueldo
Dim Categoria
Dim CentroCosto
Dim Localidad
Dim proFecPago
Dim pliqhasta
Dim FormaPago
Dim Puesto
Dim Banco
Dim nroCuenta

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
      Sueldo = 0
   Else
      Sueldo = rsConsult!empremu
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
'Busco el valor del cuil
'------------------------------------------------------------------
StrSql = " SELECT cuil.nrodoc "
StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Cuil = rsConsult!NroDoc
Else
'   Flog.writeline "Error al obtener los datos del cuil"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor de la direccion y localidad
'------------------------------------------------------------------

StrSql = " SELECT detdom.calle,detdom.nro,detdom.piso,detdom.oficdepto,localidad.locdesc"
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(proFecPago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(proFecPago) & ") AND tenro=1 AND his_estructura.ternro=" & Ternro
StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
       
OpenRecordset StrSql, rsConsult
Direccion = ""
If Not rsConsult.EOF Then
   Direccion = rsConsult!calle & " " & rsConsult!nro
    
    '02/10/2006 - Martin Ferraro - Se agrego piso y dpto a la dir del la empresa
    If Not EsNulo(rsConsult!Piso) Then
        Direccion = Direccion & " P. " & rsConsult!Piso
    End If
    If Not EsNulo(rsConsult!oficdepto) Then
        Direccion = Direccion & " Dpto. " & rsConsult!oficdepto
    End If
    
    Direccion = Direccion & ", " & rsConsult!locdesc
   
   Localidad = Direccion
Else
'   Flog.writeline "Error al obtener los datos de la localidad"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

If Sueldo = 0 Then
    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
    
    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
       Sueldo = rsConsult!almonto
    Else
       Flog.writeline "Error al obtener los datos del sueldo"
       Sueldo = 0
       'GoTo MError
    End If
End If

'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 3 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Categoria = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la categoria"
'   GoTo MError
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
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If


'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
StrSql = " SELECT ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc"
StrSql = StrSql & " From pago"
StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro AND pago.pagorigen=" & cliqnro
StrSql = StrSql & " LEFT JOIN banco     ON banco.ternro = pago.banternro"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
  If rsConsult!fpagbanc = "-1" Then
    FormaPago = rsConsult!fpagdescabr & " " & rsConsult!terrazsoc & " " & rsConsult!ctabnro
  Else
    FormaPago = rsConsult!fpagdescabr
  End If
Else
'   Flog.writeline "Error al obtener los datos de la forma de pago"
'   GoTo MError
End If

rsConsult.Close

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

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
    '02/10/2006 - Martin Ferraro - Se agrego piso y dpto a la dir del la empresa
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

'--------------------------------------------------------------------
'Obtengo la configuracion del tipo de documento resolucion 23/05/2015
'--------------------------------------------------------------------
StrSql = " SELECT confval FROM confrep WHERE repnro = 60 AND confnrocol = 358 "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    tidnroRST = rsConsult!confval
    Flog.writeline "Codigo de tipo de documento Resolucion encontrado."
Else
    tidnroRST = 0
    Flog.writeline "Codigo de tipo de documento Resolucion No encontrado."
End If

'Consulta para buscar el numero de resolucion de la empresa
StrSql = "SELECT nrodoc FROM ter_doc " & _
         " Where ternro =" & rs_estructura!Ternro & " AND ter_doc.tidnro = " & tidnroRST
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No se encontró nro de resolucion de la empresa"
    'Exit Sub
    codResolucion = 0
Else
    codResolucion = rsConsult!NroDoc
End If
'--------------------------------------------------------------------
'Obtengo la configuracion del tipo de documento resolucion 23/05/2015
'--------------------------------------------------------------------



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

'------------------------------------------------------------------
'Busco el valor de la obra social
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 17 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

ObraSocial = ""

If Not rsConsult.EOF Then
   ObraSocial = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
Flog.writeline "Buscando los datos de la cuenta del empleado."

StrSql = " SELECT * FROM ctabancaria LEFT JOIN banco ON banco.ternro = ctabancaria.banco WHERE ctabestado=-1 AND ctabancaria.ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
        Banco = rsConsult!Bandesc
        nroCuenta = rsConsult!ctabnro
        Flog.writeline "Datos de la cuenta bancaria obtenidos"
  Else
        Banco = ""
        nroCuenta = ""
        Flog.writeline "El empleado no tiene cuentas bancarias activas"
End If


rsConsult.Close

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
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden,auxchar1,auxchar2,auxdeci1,auxchar3)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & Sueldo
StrSql = StrSql & ",'" & Mid(Categoria, 1, 20) & "'"
StrSql = StrSql & ",'" & Mid(CentroCosto, 1, 25) & "'"
StrSql = StrSql & ",'" & Mid(Localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & FormaPago & "'"
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & ObraSocial & "'"
StrSql = StrSql & ",'" & Banco & "'"
StrSql = StrSql & "," & 0
StrSql = StrSql & ",'" & codResolucion & "'"
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
' Se encarga de generar los datos para La Caja
'--------------------------------------------------------------------
Sub generarDatosRecibo184(Pronro, Ternro, acunroSueldo, tituloReporte, orden)

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
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim Cuil
Dim empFecAlta
Dim Sueldo
Dim Localidad
Dim proFecPago
Dim pliqhasta
Dim FormaPago
Dim Banco
Dim Cuenta
Dim Sucursal
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
Dim empFecBaja

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim DOC
Dim tipo
Dim TipoDoc

Dim Destino_Cod
Dim Destino_Des
Dim l_Destino_Cod
Dim l_Destino_Des
Dim l_Destino

Dim l_Etiq1
Dim l_Etiq2
Dim l_Etiq3

Dim l_EmpCodExt

Dim Calificacion
Dim l_calificacion
Dim remu_estr
Dim l_remu_estr
Dim Remu_AC
Dim l_Remu_AC
Dim BancoCodExt
Dim FecAltReco
Dim DescFecAltReco As String
Dim TenroAltReco
Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

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
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu,empfecbaja "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   Apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
   empFecAlta = rsConsult!empfaltagr
   empFecBaja = rsConsult!empFecBaja
   If IsNull(rsConsult!empremu) Then
      Sueldo = 0
   Else
      Sueldo = rsConsult!empremu
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
'Busco la Fecha de Baja
'------------------------------------------------------------------
StrSql = " SELECT bajfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE fases.empleado = " & Ternro & " AND fases.real = -1 "
StrSql = StrSql & " ORDER BY altfec DESC"
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    If Not IsNull(rsConsult!bajfec) Then
        empFecBaja = rsConsult!bajfec
    Else
        empFecBaja = "  "
    End If
Else
    empFecBaja = "  "
End If

rsConsult.Close


'------------------------------------------------------------------
'Busco la Fecha de Alta Reconocida con estado Inactiva
'------------------------------------------------------------------
FecAltReco = ""
DescFecAltReco = ""
StrSql = " SELECT altfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE fases.empleado = " & Ternro & " AND fases.fasrecofec = -1 "
StrSql = StrSql & " AND estado =0 "
StrSql = StrSql & " ORDER BY altfec DESC"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If Not EsNulo(rsConsult!altfec) Then
        FecAltReco = rsConsult!altfec
        DescFecAltReco = "FECHA DE ANTIGÜEDAD RECONOCIDA"
    End If
End If
rsConsult.Close

'------------------------------------------------------------------
'Si el empleado tiene la estructura configurada cambio descripción
'------------------------------------------------------------------
If FecAltReco <> "" Then
    'Recupero TE del confrep
    '------------------------------------------------------------
    TenroAltReco = 0
    StrSql = " SELECT confval FROM confrep WHERE repnro = 60 AND conftipo = 'TE' "
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        TenroAltReco = rsConsult!confval
    End If
    If TenroAltReco <> 0 Then
        StrSql = " SELECT estrdabr "
        StrSql = StrSql & " From his_estructura"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = " & TenroAltReco
        StrSql = StrSql & " AND his_estructura.ternro = " & Ternro
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            DescFecAltReco = "FECHA DE INGRESO RECONOCIDA"
        End If
    End If
End If


'------------------------------------------------------------------
'Busco el valor del cuil
'------------------------------------------------------------------
StrSql = " SELECT cuil.nrodoc "
StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Cuil = rsConsult!NroDoc
Else
'   Flog.writeline "Error al obtener los datos del cuil"
'   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la direccion y localidad
'------------------------------------------------------------------

StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc"
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(proFecPago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(proFecPago) & ") AND tenro=1 AND his_estructura.ternro=" & Ternro
StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
OpenRecordset StrSql, rsConsult

Localidad = ""
If Not rsConsult.EOF Then
   'Direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
   Direccion = rsConsult!locdesc
   Localidad = Direccion
Else
'   Flog.writeline "Error al obtener los datos de la localidad"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

If Sueldo = 0 Then
    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
    
    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
       Sueldo = rsConsult!almonto
    Else
       Flog.writeline "Error al obtener los datos del sueldo"
       Sueldo = 0
       'GoTo MError
    End If
End If

'------------------------------------------------------------------
'Busco el valor de la Sucursal
'------------------------------------------------------------------
Flog.writeline "Busco el valor de la Sucursal del empleado"
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 1 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Sucursal = rsConsult!estrdabr
Else
   Flog.writeline "Error al obtener los datos de la sucursal del empleado"
   Sucursal = ""
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
StrSql = " SELECT ctabnro,ctabcbu,fpagdescabr,tercero.terrazsoc,fpagbanc,estrcodext"
StrSql = StrSql & " From pago"
StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro AND pago.pagorigen=" & cliqnro
StrSql = StrSql & " INNER JOIN banco ON banco.ternro = pago.banternro"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro= banco.estrnro"
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
  If rsConsult!fpagbanc = "-1" Then
    FormaPago = rsConsult!fpagdescabr
    Banco = rsConsult!terrazsoc
    BancoCodExt = rsConsult!estrcodext
    If BancoCodExt = "017" Then
        Cuenta = rsConsult!ctabnro
    Else
        Cuenta = rsConsult!ctabcbu
    End If
'    If Not EsNulo(rsConsult!ctabcbu) And (rsConsult!ctabcbu <> "0") Then
'        Cuenta = rsConsult!ctabcbu
'    Else
'       Cuenta = rsConsult!ctabnro
'    End If
  Else
    FormaPago = rsConsult!fpagdescabr
    Banco = rsConsult!terrazsoc
    If BancoCodExt = "017" Then
        Cuenta = rsConsult!ctabnro
    Else
        Cuenta = rsConsult!ctabcbu
    End If
'    If Not EsNulo(rsConsult!ctabcbu) Then
'        Cuenta = rsConsult!ctabcbu
'    Else
'        Cuenta = rsConsult!ctabnro
'    End If
  End If
Else
'   Flog.writeline "Error al obtener los datos de la forma de pago"
'   GoTo MError
End If

rsConsult.Close

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom  " & _
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
Flog.writeline "Cuit de la Empresa " & StrSql
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
Flog.writeline "Cargas Sociales " & StrSql
rsConsult.Close

StrSql = " SELECT confval FROM confrep WHERE repnro = 60 AND conftipo = 'TDO' "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    TipoDoc = rsConsult!confval
    Flog.writeline "Codigo de tipo de documento Resolucion encontrado."
Else
    TipoDoc = 1
    Flog.writeline "Codigo de tipo de documento Resolucion No encontrado."
End If

'------------------------------------------------------------------
'Busco Documento
'------------------------------------------------------------------
StrSql = " SELECT nrodoc, tidsigla "
StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc ON (tercero.ternro=ter_doc.ternro and ter_doc.tidnro = " & TipoDoc & ")"
StrSql = StrSql & " INNER JOIN tipodocu ON tipodocu.tidnro=ter_doc.tidnro "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    DOC = rsConsult!NroDoc
    tipo = rsConsult!tidsigla
Else
    Flog.writeline "No se encontraro el Documento."
End If

rsConsult.Close

'DESTINO
StrSql = " SELECT confval,confval2 FROM confrep WHERE repnro = 60 AND conftipo = 'DES' "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Destino_Cod = rsConsult!confval
    Destino_Des = rsConsult!confval2
Else
    Destino_Cod = 0
    Destino_Des = 0
    Flog.writeline "Codigo de Estructura No Encontrada"
End If

StrSql = " SELECT estrcodext "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
StrSql = StrSql & " And estructura.tenro = " & Destino_Cod
StrSql = StrSql & " And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    l_Destino_Cod = Left(rsConsult!estrcodext, 4)
Else
    l_Destino_Cod = 0
    Flog.writeline "Estructura No Encontrada"
End If

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
StrSql = StrSql & " And estructura.tenro = " & Destino_Des
StrSql = StrSql & " And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    l_Destino_Des = Left(rsConsult!estrdabr, 4)
Else
    l_Destino_Des = ""
    Flog.writeline "Estructura No Encontrada"
End If
l_Destino = l_Destino_Cod & l_Destino_Des

StrSql = "SELECT estrcodext FROM estructura WHERE estrnro= " & EmpEstrnro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    l_EmpCodExt = rsConsult!estrcodext
Else
    l_EmpCodExt = 0
End If

l_Etiq1 = ""
l_Etiq2 = ""
l_Etiq3 = ""
l_Etiq4 = 0

StrSql = "SELECT * FROM confrep WHERE repnro = 60 and conftipo='CON' and confnrocol = 385"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    l_estr = rsConsult!confval
    l_tipo_estr = rsConsult!confval2
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " From his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " "
    StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
    StrSql = StrSql & " AND his_estructura.estrnro = " & l_estr & "  And his_estructura.ternro = " & Ternro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then           'Encontro Estructura Confidencial
            StrSql = " SELECT estrdabr "
            StrSql = StrSql & " From his_estructura"
            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " "
            StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
            StrSql = StrSql & " AND his_estructura.tenro = " & l_tipo_estr & "  And his_estructura.ternro = " & Ternro
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                l_Etiq1 = "CATEGORIA"
                l_Etiq2 = rsConsult!estrdabr
            End If
            StrSql = "SELECT * FROM confrep WHERE repnro = 60 and confnrocol=389"
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                 l_tipo_estr = rsConsult!confval2
                 StrSql = " SELECT estrcodext "
                 StrSql = StrSql & " From his_estructura"
                 StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
                 StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " "
                 StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
                 StrSql = StrSql & " AND his_estructura.tenro = " & l_tipo_estr & "  And his_estructura.ternro = " & Ternro
                 OpenRecordset StrSql, rsConsult
                 If Not rsConsult.EOF Then
                    l_Etiq3 = "FUNCION"
                    l_Etiq4 = rsConsult!estrcodext
                End If
            End If
    End If
Else
        Flog.writeline "No Definio la Estructura Confidencial"
End If
Flog.writeline "Busco Estructura Nomina General"
StrSql = "SELECT * FROM confrep WHERE repnro=60 and conftipo='GRA'"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    l_estr = rsConsult!confval
    l_tipo_estr = rsConsult!confval2
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " From his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " "
    StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
    StrSql = StrSql & " AND his_estructura.estrnro = " & l_estr & "  And his_estructura.ternro = " & Ternro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then           'Encontro Estructura Confidencial
        StrSql = " SELECT estrdabr "
        StrSql = StrSql & " From his_estructura"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " "
        StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
        StrSql = StrSql & " AND his_estructura.tenro = " & l_tipo_estr & "  And his_estructura.ternro = " & Ternro
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            l_Etiq1 = "FUNCION"
            l_Etiq2 = rsConsult!estrdabr
        End If
        StrSql = "SELECT * FROM confrep WHERE repnro = 60 and confnrocol=390"
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
             l_tipo_estr = rsConsult!confval2
             StrSql = " SELECT estrcodext "
             StrSql = StrSql & " From his_estructura"
             StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
             StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " "
             StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
             StrSql = StrSql & " AND his_estructura.tenro = " & l_tipo_estr & "  And his_estructura.ternro = " & Ternro
             OpenRecordset StrSql, rsConsult
             If Not rsConsult.EOF Then
                l_Etiq3 = "CATEGORIA"
                l_Etiq4 = rsConsult!estrcodext
            End If
        End If
    End If
Else
    Flog.writeline "No Definio la Estructura Nomina General"
End If

'Calificacion Personal
StrSql = "SELECT confval FROM confrep WHERE repnro=60 AND conftipo='TE1'"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Calificacion = rsConsult!confval
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " From his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
    StrSql = StrSql & " And estructura.tenro = " & Calificacion
    StrSql = StrSql & " And his_estructura.ternro = " & Ternro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        l_calificacion = rsConsult!estrdabr
    Else
        l_calificacion = ""
        Flog.writeline "Estructura No Encontrada"
    End If
Else
    Flog.writeline "El confrep no tiene la Estructura Remuneracion Configurada "
    l_calificacion = ""
End If

'Remunercion Estructura
StrSql = "SELECT confval FROM confrep WHERE repnro=60 AND conftipo='TE2'"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    remu_estr = rsConsult!confval
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " From his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") "
    StrSql = StrSql & " And estructura.tenro = " & remu_estr
    StrSql = StrSql & " And his_estructura.ternro = " & Ternro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        l_remu_estr = rsConsult!estrdabr
    Else
        l_remu_estr = ""
        Flog.writeline "Estructura No Encontrada"
    End If
Else
    Flog.writeline "El confrep no tiene la Estructura Remuneracion Configurada "
    l_remu_estr = ""
End If

'Acumulador
StrSql = "SELECT confval FROM confrep WHERE repnro=60 AND confnrocol=387"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Remu_AC = rsConsult!confval
    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & Remu_AC
    StrSql = StrSql & " AND cliqnro = " & cliqnro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        l_Remu_AC = rsConsult!almonto
    Else
        l_Remu_AC = 0
        Flog.writeline "Acumulador No Encontrado"
    End If
Else
    Flog.writeline "El confrep no tiene el Acumulador Configurado en la Columna 387 "
    l_Remu_AC = 0
End If

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,direccion,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,pliqdepant,"
StrSql = StrSql & " pliqfecdep,pliqbco,cuil,empfecalta,"
StrSql = StrSql & " sueldo,localidad,"
StrSql = StrSql & " profecpago,formapago,empnombre,empdire,empcuit,emplogo,emplogoalto,emplogoancho,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,prodesc,descripcion,puesto, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden, auxchar1, auxchar3,"
StrSql = StrSql & " auxchar6,auxchar7,auxchar9,auxchar10,auxchar11,auxchar2,auxchar4,auxchar5,auxchar12,auxchar8,auxdeci1,auxchar13,auxchar14)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & Sueldo
StrSql = StrSql & ",'" & Mid(Localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & FormaPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & ",'" & l_remu_estr & "'"  'En el Puesto va la Estructura Remuneracion configurada en confrep como TIPO TE2
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Cuenta & "'"
StrSql = StrSql & ",'" & Banco & "'"
StrSql = StrSql & ",'" & DOC & "'"
StrSql = StrSql & ",'" & tipo & "'"
StrSql = StrSql & ",'" & l_Destino_Des & "'"
StrSql = StrSql & ",'" & l_Destino_Cod & "'"
StrSql = StrSql & ",'" & l_EmpCodExt & "'"
StrSql = StrSql & ",'" & l_Etiq1 & "'"
StrSql = StrSql & ",'" & l_Etiq2 & "'"
StrSql = StrSql & ",'" & l_Etiq3 & "'"
StrSql = StrSql & ",'" & l_Etiq4 & "'"
StrSql = StrSql & ",'" & l_calificacion & "'"
StrSql = StrSql & "," & l_Remu_AC
StrSql = StrSql & ",'" & DescFecAltReco & "'"  'DESCRIPCION | auxchar13
StrSql = StrSql & ",'" & FecAltReco & "'"  'fecha de alta reconocida | auxchar14
StrSql = StrSql & ")"


'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

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
' Se encarga de generar los datos para el Recibo Uruguay
'--------------------------------------------------------------------
Sub generarDatosRecibo185(Pronro, Ternro, acunroSueldo, tituloReporte, orden)

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
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim Cuil
Dim empFecAlta
Dim Sueldo As Double
Dim Categoria
Dim CentroCosto
Dim Localidad
Dim proFecPago
Dim pliqhasta
Dim FormaPago
Dim Puesto
'Dim MontoTicketBruto As Double
'Dim MontoTicketNeto As Double
Dim ModeloLiq
Dim ImpIRPF
'Dim ImpDescuentoIRPF

Dim EmpEstrnro
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim CodRUC As Integer
Dim CodBPS As Integer
Dim EmpMTSS As String
Dim EmpBPS As String
Dim EmpRUC As String
Dim EmpBSE As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String
Dim CtaSucDesc

Dim codCargo As Long
Dim codSector As Long

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim esConcTotal As Boolean
Dim concTotal As String
Dim valorTotal As Double

Dim Banco As String
Dim nroCuenta As String

Dim NroTransaccion

Dim Tipo_Grupo
Dim Tipo_Subgrupo
Dim Tipo_CJPB
Dim Tipo_BPS
Dim Tipo_MTSS
Dim Tipo_BSE
Dim Tipo_PlanillaTrabajo

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

codCargo = 0
codSector = 0

'------------------------------------------------------------------------------'
'  Buscar la configuracion del confrep, para los Tipo de Documentos RUC y BPS  '
'------------------------------------------------------------------------------'
Flog.writeline "Obtengo los datos del confrep para los Tipo de Documentos RUC y BPS."
CodRUC = 0

'RUC
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 391 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No esta configurado el ConfRep para el Tipo de Documento RUC (col 391)."
Else
    If rsConsult!conftipo = "DOC" Then
        CodRUC = rsConsult!confval
    Else
        Flog.writeline " El Tipo de Columna 391 del confrep no es:Tipo Documento (DOC). "
    End If
End If
rsConsult.Close

'Grupo
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 397 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
   Flog.writeline "No esta configurado el ConfRep para el Grupo (col 397)."
   Tipo_Grupo = 0
Else
   Tipo_Grupo = rsConsult!confval2
End If

'SubGrupo
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 398 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
   Flog.writeline "No esta configurado el ConfRep para el SubGrupo (col 398)."
   Tipo_Subgrupo = 0
Else
   Tipo_Subgrupo = rsConsult!confval2
End If

'CJPB
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 399 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
   Flog.writeline "No esta configurado el ConfRep para el CJPB (col 399)."
   Tipo_CJPB = 0
Else
   Tipo_CJPB = rsConsult!confval2
End If

'BPS
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 400 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
   Flog.writeline "No esta configurado el ConfRep para el BPS (col 400)."
   Tipo_BPS = 0
Else
   Tipo_BPS = rsConsult!confval2
End If

'MTSS
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 401 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
   Flog.writeline "No esta configurado el ConfRep para el MTSS (col 401)."
   Tipo_MTSS = 0
Else
   Tipo_MTSS = rsConsult!confval2
End If

'BSE
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 402 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
   Flog.writeline "No esta configurado el ConfRep para el BSE (col 402)."
   Tipo_BSE = 0
Else
   Tipo_BSE = rsConsult!confval2
End If

'Planilla de Trabajo
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 403 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
   Flog.writeline "No esta configurado el ConfRep para la Planilla de Trabajo (col 403)."
   Tipo_PlanillaTrabajo = 0
Else
   Tipo_PlanillaTrabajo = rsConsult!confval2
End If

StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol IN (2,395,396) "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No esta configurado el ConfRep para el Tipo de Estructura SECTOR Y CARGO"
Else
    Do While Not rsConsult.EOF
        Select Case rsConsult!confnrocol
            Case 2: 'Neto
                If UCase(rsConsult!conftipo) = "CO" Then
                    esConcTotal = True
                    concTotal = rsConsult!confval2
                Else
                    If UCase(rsConsult!conftipo) = "AC" Then
                        esConcTotal = False
                        concTotal = rsConsult!confval
                    Else
                        Flog.writeline "La columna 2 debe ser de tipo CO o AC "
                        concTotal = 0
                    End If
                End If
            Case 395: 'Cargo
                If UCase(rsConsult!conftipo) = "TE" Then
                    codCargo = rsConsult!confval
                Else
                    codCargo = 0
                End If
            Case 396: 'Sector
                If UCase(rsConsult!conftipo) = "TE" Then
                    codSector = rsConsult!confval
                Else
                    codSector = 0
                End If
        End Select
    rsConsult.MoveNext
    Loop
End If
rsConsult.Close
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
Else
    Flog.writeline "Error al obtener los datos del empleado"
    GoTo MError
End If

'------------------------------------------------------------------
'Busco la Fecha de Alta del empleado. Corresponde a la fecha desde de la fase que
'esta activa.
'------------------------------------------------------------------
StrSql = " SELECT altfec FROM fases "
StrSql = StrSql & " WHERE empleado= " & Ternro & " AND estado = -1 ORDER BY altfec DESC"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   empFecAlta = rsConsult!altfec
Else
   Flog.writeline "Error al obtener la Fecha de Alta del Empleado"
   empFecAlta = ""
   'GoTo MError
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
'Busco el valor de la cedula de identidad
'------------------------------------------------------------------
StrSql = " SELECT cuil.nrodoc "
StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro <= 5) "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
StrSql = StrSql & " ORDER BY cuil.tidnro ASC "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Cuil = rsConsult!NroDoc
Else
    Cuil = " "
    Flog.writeline "Error al obtener los datos del cuil"
End If

'------------------------------------------------------------------
'Busco el valor de la direccion y localidad
'------------------------------------------------------------------
StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc"
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(proFecPago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(proFecPago) & ") AND tenro=10 AND his_estructura.ternro=" & Ternro
StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = empresa.ternro"
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
OpenRecordset StrSql, rsConsult
Direccion = ""
Localidad = ""
If Not rsConsult.EOF Then
    'Direccion = rsConsult!calle & " " & rsConsult!nro & ", " & rsConsult!locdesc
    Direccion = rsConsult!calle & " " & rsConsult!nro
    'localidad = Direccion
End If

'------------------------------------------------------------------
' Busco VHora definido en la columna 44 del confrep
'Se cambio -- se usa col 44 para el Sueldo.
'------------------------------------------------------------------
Sueldo = 0

If esSueldoRecibo Then
    StrSql = " SELECT detliq.dlimonto valor "
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And detliq.concnro = " & SueldoRecibo
Else
    StrSql = " SELECT almonto valor"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & SueldoRecibo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Sueldo = rsConsult!Valor
Else
    Flog.writeline " No se encontraron datos de de SUELDO/JORNAL. (col44)"
    Sueldo = 0
    'GoTo MError
End If
rsConsult.Close

'BUSCO EL CAMPO PESOS NETO

If esConcTotal Then
    StrSql = " SELECT detliq.dlimonto valor "
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro =  detliq.concnro AND concepto.conccod=  '" & concTotal & "'"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
Else
    StrSql = " SELECT almonto valor"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & concTotal
    StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    valorTotal = rsConsult!Valor    'Neto
Else
    Flog.writeline " No se encontraron datos del campo pesos neto (col 2)"
    valorTotal = 0                  'Neto
    'GoTo MError
End If
rsConsult.Close
'HASTA ACA

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
'Busco el valor del cargo
'------------------------------------------------------------------
Categoria = " "

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro =" & codCargo & " And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Categoria = rsConsult!estrdabr
End If

'------------------------------------------------------------------
'Busco el valor del puesto
'------------------------------------------------------------------
Puesto = ""

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Puesto = rsConsult!estrdabr
End If

'------------------------------------------------------------------
'Busco el valor del centro de costo, en realidad se busca el Sector del empleado
'------------------------------------------------------------------
CentroCosto = " "

StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=" & codSector & " AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    CentroCosto = rsConsult!estrdabr
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
StrSql = " SELECT pago.ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc, fpagbanc, ctabsucdesc "
StrSql = StrSql & " From pago"
StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro  AND pago.pagorigen=" & cliqnro
StrSql = StrSql & " LEFT JOIN banco ON banco.ternro = pago.banternro"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
StrSql = StrSql & " INNER JOIN ctabancaria ON ctabancaria.banco=pago.banternro "
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    If rsConsult!fpagbanc = "-1" Then
        FormaPago = "Banco"
        nroCuenta = rsConsult!ctabnro
        Banco = rsConsult!terrazsoc
        CtaSucDesc = rsConsult!ctabsucdesc
    Else
        FormaPago = "Efectivo"
        nroCuenta = ""
        Banco = ""
        CtaSucDesc = ""
    End If
    Flog.writeline "Obtengo Forma Pago - Nro de Cuenta - Banco " & StrSql
Else
    FormaPago = "" 'banco - efectivo
    Banco = ""
    nroCuenta = ""
    Flog.writeline "Error al obtener los datos de la forma de pago "
End If
rsConsult.Close

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------
EmpEstrnro = 0
EmpNombre = " "

StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = 10"
OpenRecordset StrSql, rs_estructura
If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    GoTo MError
Else
    EmpNombre = rs_estructura!empnom
    EmpEstrnro = rs_estructura!Estrnro
End If


' -------------------------------------------------------------------------
'Consulta para obtener el cuit de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa"
    EmpCuit = "  "
Else
    EmpCuit = rs_cuit!NroDoc
End If
rs_cuit.Close

' -------------------------------------------------------------------------
'Consulta para obtener el BSE de la empresa
'Flog.writeline "Buscando estructura BSE tipo " & CodBSE
'StrSql = "SELECT cuit.nrodoc FROM tercero "
'StrSql = StrSql & " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro AND cuit.tidnro = " & CodBSE & ")"
'StrSql = StrSql & " WHERE tercero.ternro =" & rs_estructura!Ternro
'OpenRecordset StrSql, rs_cuit
'If rs_cuit.EOF Then
'    Flog.writeline "No se encontró el BSE de la empresa"
'    EmpBSE = " "
'Else
'    Flog.writeline "BSE = " & rs_cuit!NroDoc
'    EmpBSE = rs_cuit!NroDoc
'End If
'rs_cuit.Close

' -------------------------------------------------------------------------
'Consulta para buscar el logo de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_logo
If rs_logo.EOF Then
    Flog.writeline "No se encontró el Logo de la Empresa"
    EmpLogo = ""
    EmpLogoAlto = 0
    EmpLogoAncho = 0
Else
    EmpLogo = rs_logo!tipimdire & rs_logo!terimnombre
    EmpLogoAlto = rs_logo!tipimaltodef
    EmpLogoAncho = rs_logo!tipimanchodef
End If

' -------------------------------------------------------------------------
'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If

' -------------------------------------------------------------------------
'Consulta para obtener el MTSS de la empresa
'StrSql = " SELECT *"
'StrSql = StrSql & " From ter_doc"
'StrSql = StrSql & " Where ter_doc.ternro = " & rs_estructura!Ternro & " And ter_doc.tidnro = " & CodMTSS
'OpenRecordset StrSql, rs_cuit
'If rs_cuit.EOF Then
'    Flog.writeline "No se encontró el MTSS"
'    EmpMTSS = "  "
'Else
'    Flog.writeline "MTSS = " & rs_cuit!NroDoc
'    EmpMTSS = IIf(EsNulo(rs_cuit!NroDoc), " ", rs_cuit!NroDoc)
'End If
'rs_cuit.Close
' -------------------------------------------------------------------------
'Consulta para obtener el BPS de la empresa
'StrSql = "SELECT cuit.nrodoc FROM tercero "
'StrSql = StrSql & " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro AND cuit.tidnro = " & CodBPS & ")"
'StrSql = StrSql & " WHERE tercero.ternro =" & rs_estructura!Ternro
'OpenRecordset StrSql, rs_cuit
'If rs_cuit.EOF Then
'    Flog.writeline "No se encontró el BPS de la Empresa"
'    EmpBPS = "  "
'Else
'    EmpBPS = rs_cuit!NroDoc
'End If
'rs_cuit.Close
' -------------------------------------------------------------------------
'Consulta para obtener el RUC de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero "
StrSql = StrSql & " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro AND cuit.tidnro = " & CodRUC & ")"
StrSql = StrSql & " WHERE tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el RUC de la Empresa"
    EmpRUC = "  "
Else
    EmpRUC = rs_cuit!NroDoc
End If
rs_cuit.Close



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
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco Nro de Transaccion
'------------------------------------------------------------------
StrSql = " SELECT ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc, fpagbanc,NroTransaccion "
StrSql = StrSql & " From pago"
StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro  AND pago.pagorigen=" & cliqnro
StrSql = StrSql & " LEFT JOIN banco ON banco.ternro = pago.banternro"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    If EsNulo(rsConsult!NroTransaccion) Then
        NroTransaccion = 0
    Else
        NroTransaccion = rsConsult!NroTransaccion
    End If
Else
    NroTransaccion = 0
End If
Flog.writeline "Nro de Transaccion " & StrSql
'-----------------------------------------------------------------------------------------


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
StrSql = StrSql & " empfirmaalto,empfirmaancho,prodesc,descripcion,puesto, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden,"
StrSql = StrSql & " auxchar1, auxchar2, auxchar3, auxchar4, auxchar11, auxchar12, auxchar13, auxchar14, "
StrSql = StrSql & " auxchar5, auxdeci6, auxchar, auxchar6, auxdeci7,auxchar7"
StrSql = StrSql & ")"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & numberForSQL(Sueldo)
StrSql = StrSql & ",'" & Mid(Categoria, 1, 100) & "'"
StrSql = StrSql & ",'" & Mid(CentroCosto, 1, 100) & "'"
StrSql = StrSql & ",'" & Mid(Localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden

StrSql = StrSql & ",'" & Tipo_MTSS & "'"        'auxchar1
StrSql = StrSql & ",'" & Tipo_BPS & "'"         'auxchar2
StrSql = StrSql & ",'" & EmpRUC & "'"           'axuchar3
StrSql = StrSql & ",'" & Tipo_BSE & "'"         'auxchar4
StrSql = StrSql & ",'" & Tipo_PlanillaTrabajo & "'"    'auxchar11
StrSql = StrSql & ",'" & Tipo_Grupo & "'"       'auxchar12
StrSql = StrSql & ",'" & Tipo_Subgrupo & "'"    'auxchar13
StrSql = StrSql & ",'" & Tipo_CJPB & "'"        'auxchar14
StrSql = StrSql & ",'" & ModeloLiq & "'"        'auxchar5
StrSql = StrSql & "," & numberForSQL(valorTotal)    'auxdeci6
StrSql = StrSql & ",'" & nroCuenta & "'"        'auxchar
StrSql = StrSql & ",'" & Banco & "'"            'auxchar6
StrSql = StrSql & "," & NroTransaccion          'auxdeci7
StrSql = StrSql & ",'" & CtaSucDesc & "'"       'auxchar7
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------
objConn.Execute StrSql, , adExecuteNoRecords


'Flog.Writeline "======================================================================================================================================="
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

'Flog.writeline " Se sale del procedimiento de generarDatosRecibo50. "

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
' Se encarga de generar los datos para agco
'--------------------------------------------------------------------
Sub generarDatosRecibo186(Pronro, Ternro, acunroSueldo, tituloReporte, orden)

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
Dim pliqdesc
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim Cuil
Dim NroDoc
Dim empFecAlta
Dim Sueldo
Dim Categoria
Dim Sector
Dim Localidad
Dim proFecPago
Dim pliqhasta
Dim FormaPago
Dim Puesto
Dim ObraSocial
Dim CabliqNumero
Dim Fecha_aux
Dim anios_aux As Integer
Dim meses_aux As Integer
Dim dias_aux As Integer
Dim diasHab_aux As Integer
Dim Antiguedad
Dim CentroCosto
Dim BancoPago

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
    
Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim ConvenioNro As Integer
Dim SucursalNro As Integer
Dim CategoriaNro As Integer
Dim FormaLiqNro As Integer
Dim valorAntig
Dim AntigGrilla
Dim salir
Dim Contrato

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

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
   Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos del proceso."
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
      Sueldo = 0
   Else
      Sueldo = rsConsult!empremu
   End If
Else
   Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos del empleado."
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro AND proceso.pronro= " & Pronro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   pliqhasta = rsConsult!pliqhasta
   pliqdesc = rsConsult!pliqdesc
Else
   Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos del Período actual."
   GoTo MError
End If

'------------------------------------------------------------------
'Busco la Fecha de Alta del empleado. Corresponde a la fecha desde de la fase que
'contiene la marca de fecha de Alta Reconocida.
'------------------------------------------------------------------
StrSql = " SELECT altfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE empleado= " & Ternro & " ORDER BY altfec "
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    empFecAlta = rsConsult!altfec
Else
'    rsConsult.Close
'    StrSql = " SELECT altfec "
'    StrSql = StrSql & " FROM fases "
'    StrSql = StrSql & " WHERE empleado= " & ternro & " ORDER BY altfec DESC "
'    OpenRecordset StrSql, rsConsult
'    If Not rsConsult.EOF Then
'        empFecAlta = rsConsult!altfec
'    Else
        Flog.writeline Espacios(Tabulador * 1) & "Error al obtener la Fecha de Alta del Empleado"
        empFecAlta = "&nbsp;"
'    End If
End If

'------------------------------------------------------------------
'Calculo la antiguedad
'------------------------------------------------------------------
Fecha_aux = CDate("01" & "/" & pliqmes & "/" & pliqanio)

'FB - 17/06/2014 - Se corrige el cálculo de la antiguedad
Fecha_aux = DateAdd("m", 1, Fecha_aux)
Fecha_aux = DateAdd("d", -1, Fecha_aux)


If Fecha_aux < empFecAlta Then
    Antiguedad = "Años: 0 Meses: 0"
Else
    Call bus_Antiguedad(Ternro, "REAL", Fecha_aux, dias_aux, meses_aux, anios_aux, diasHab_aux)
    
'    anios_aux = DateDiff("yyyy", empFecAlta, Fecha_aux)
'    If Month(empFecAlta) > Month(Fecha_aux) Or (Month(empFecAlta) = Month(Fecha_aux) And Day(empFecAlta) > Day(Fecha_aux)) Then
'        anios_aux = anios_aux - 1
'    End If
'    meses_aux = DateDiff("m", empFecAlta, Fecha_aux) - (anios_aux * 12)
'    If Day(empFecAlta) > Day(Fecha_aux) Or (Month(empFecAlta) = Month(Fecha_aux) And Day(empFecAlta) > Day(Fecha_aux)) Then
'        meses_aux = meses_aux - 1
'    End If
    Antiguedad = "Años: " & anios_aux & " Meses: " & meses_aux
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
'Busco el valor del cuil
'------------------------------------------------------------------
StrSql = " SELECT tipodocu.tidsigla, cuil.nrodoc "
StrSql = StrSql & " FROM tercero INNER JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
StrSql = StrSql & " INNER JOIN tipodocu ON tipodocu.tidnro = cuil.tidnro "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Cuil = rsConsult!tidsigla & "-" & rsConsult!NroDoc
Else
   Flog.writeline Espacios(Tabulador * 1) & "No se encontraron los datos del Cuil."
   Cuil = "&nbsp;"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del nro documento
'------------------------------------------------------------------
StrSql = " SELECT tipodocu.tidsigla, ter_doc.nrodoc "
StrSql = StrSql & " FROM tercero INNER JOIN ter_doc ON (tercero.ternro=ter_doc.ternro and ter_doc.tidnro<=5) "
StrSql = StrSql & " INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   NroDoc = rsConsult!tidsigla & "-" & rsConsult!NroDoc
Else
   Flog.writeline Espacios(Tabulador * 1) & "No se encontraron los datos del Nro documento."
   NroDoc = "&nbsp;"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------
StrSql = " SELECT his_estructura.estrnro,estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 3 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Categoria = rsConsult!estrdabr
   CategoriaNro = rsConsult!Estrnro
Else
   Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos de la Categoría."
   Categoria = "&nbsp;"
   CategoriaNro = 0
'   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del sector
'------------------------------------------------------------------
Flog.writeline "Busco el valor del sector"
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 2 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

Sector = ""

If Not rsConsult.EOF Then
   Sector = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del contrato
'------------------------------------------------------------------
StrSql = " SELECT his_estructura.estrnro,estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 18 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Contrato = rsConsult!estrdabr
Else
   Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos del Contrato."
   Contrato = "&nbsp;"
'   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco Centro de Costo
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 5 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrdabr
Else
   Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos del Centro Costo."
   CentroCosto = "&nbsp;"
'   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco Puesto
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Puesto = rsConsult!estrdabr
Else
   Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos del Puesto."
   Puesto = "&nbsp;"
'   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la obra social
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 17 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   ObraSocial = rsConsult!estrdabr
Else
   Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos de la Obra Social."
   ObraSocial = "&nbsp;"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el domicilio de la sucursal
'------------------------------------------------------------------
StrSql = " SELECT his_estructura.estrnro,detdom.calle,detdom.nro,localidad.locdesc "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 1 And his_estructura.ternro = " & Ternro
StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro = estructura.estrnro "
StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro AND cabdom.tidonro = 5 "
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    Localidad = rsConsult!calle & " " & rsConsult!nro & " - " & rsConsult!locdesc
    SucursalNro = rsConsult!Estrnro
Else
    Localidad = "&nbsp;"
    SucursalNro = 0
    Flog.writeline Espacios(Tabulador * 1) & "Error al obtener el Domicilio de la sucursal del empleado."
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor de Básico
'------------------------------------------------------------------
Sueldo = 0
If esSueldoConc Then
    StrSql = " SELECT detliq.dlimonto valor "
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And detliq.concnro = " & acunroSueldo
Else
    StrSql = " SELECT almonto valor"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Sueldo = rsConsult!Valor
Else
   Flog.writeline "Error al obtener los datos del Básico. Se debe configurar en la columna 1 del confrep."
   Sueldo = 0
'   GoTo MError
End If

    
'------------------------------------------------------------------
'Busco el valor de la Antigüedad
'------------------------------------------------------------------
valorAntig = 0
If SueldoRecibo <> 0 Then
    'Columna 42
    If esSueldoRecibo Then
        StrSql = " SELECT novemp.nevalor valor "
        StrSql = StrSql & " FROM novemp "
        StrSql = StrSql & " WHERE novemp.empleado = " & Ternro & " And novemp.concnro = " & SueldoRecibo & " AND novemp.tpanro = 35"
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           valorAntig = rsConsult!Valor
           Flog.writeline Espacios(Tabulador * 1) & "---> El valor de la Antigüedad se obtuvo de la novedad para el concepto (concnro=" & SueldoRecibo & ") y parámetro 35."
        Else
           'MENSUALES
           If esGratificacionConc Then
                StrSql = " SELECT novemp.nevalor valor "
                StrSql = StrSql & " FROM novemp "
                StrSql = StrSql & " WHERE novemp.empleado = " & Ternro & " And novemp.concnro = " & Gratificacion & " AND novemp.tpanro = 35"
                OpenRecordset StrSql, rsConsult
                
                If Not rsConsult.EOF Then
                   valorAntig = rsConsult!Valor
                   Flog.writeline Espacios(Tabulador * 1) & "---> El valor de la Antigüedad se obtuvo de la novedad para el concepto (concnro=" & Gratificacion & ") y parámetro 35."
                End If
           Else
                Flog.writeline "Error al obtener el valor de la Antigüedad. La columna 43 del confrep debe ser un concepto."
           End If
           'FIN MENSUALES
        End If
    
    Else
        Flog.writeline "Error al obtener el valor de la Antigüedad. La columna 44 del confrep debe ser un concepto, de tipo 'CO'."
        valorAntig = 0
    End If
    
Else
    Flog.writeline "Error al obtener el valor de la Antigüedad. No esta configurado el concepto en la columna 44 del confrep."
    valorAntig = 0
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
StrSql = " SELECT ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc"
StrSql = StrSql & " From pago"
StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro AND pago.pagorigen=" & cliqnro
StrSql = StrSql & " LEFT JOIN banco     ON banco.ternro = pago.banternro"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
  If rsConsult!fpagbanc = "-1" Then
    FormaPago = rsConsult!ctabnro
    BancoPago = rsConsult!terrazsoc
  Else
    FormaPago = rsConsult!fpagdescabr
    BancoPago = "&nbsp;"
  End If
Else
   Flog.writeline "Error al obtener los datos de la forma de pago."
   FormaPago = "&nbsp;"
   BancoPago = "&nbsp;"
'   GoTo MError
End If

rsConsult.Close

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------
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
    Flog.writeline "No se encontró la Empresa."
    EmpNombre = "&nbsp;"
    GoTo MError
Else
    EmpNombre = rs_estructura!empnom
    EmpEstrnro = rs_estructura!Estrnro
End If

'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,detdom.codigopostal,localidad.locdesc From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el Domicilio de la Empresa."
    'Exit Sub
    EmpDire = "&nbsp;"
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro & "(" & rs_Domicilio!codigopostal & ") " & rs_Domicilio!locdesc
End If

'Consulta para obtener el cuit de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa."
    'Exit Sub
    EmpCuit = "&nbsp;"
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
    Flog.writeline "No se encontró el Logo de la Empresa."
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
    Flog.writeline "No se encontró el Firma de la Empresa."
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
   pliqdepant = "&nbsp;"
   pliqfecdep = "&nbsp;"
   pliqbco = "&nbsp;"
   Flog.writeline "No se encontraron los datos de las Cargas Sociales."
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------
StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,direccion,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,pliqdepant,"
StrSql = StrSql & " pliqfecdep,pliqbco,cuil,empfecalta,"
StrSql = StrSql & " sueldo,categoria,centrocosto,localidad,profecpago,empnombre,"
StrSql = StrSql & " empdire,empcuit,emplogo,emplogoalto,emplogoancho,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,formapago,prodesc,descripcion,puesto, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden,"
StrSql = StrSql & " auxchar1, auxchar2, auxchar3, auxchar4, auxchar5,auxchar6, auxdeci1)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & Sueldo
StrSql = StrSql & ",'" & Mid(Categoria, 1, 25) & "'"
StrSql = StrSql & ",'" & Mid(CentroCosto, 1, 25) & "'"
StrSql = StrSql & ",'" & Mid(Localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & Mid(FormaPago, 1, 200) & "'"
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & ",'" & Mid(Puesto, 1, 60) & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Mid(ObraSocial, 1, 100) & "'"
StrSql = StrSql & ",'" & Antiguedad & "'"
StrSql = StrSql & ",'" & NroDoc & "'"
StrSql = StrSql & ",'" & Mid(BancoPago, 1, 100) & "'"
StrSql = StrSql & ",'" & Mid(Contrato, 1, 100) & "'"
StrSql = StrSql & ",'" & Sector
StrSql = StrSql & "'," & valorAntig
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

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

Sub generarDatosRecibo187(Pronro, Ternro, acunroSueldo, tituloReporte, orden)
'--------------------------------------------------------------------
' Boleta de Pago - Perú - Standard
'--------------------------------------------------------------------
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
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim NroDoc

'Fechas de fases
Dim empFecAlta
Dim empFecBaja
'---------------

Dim Sueldo
Dim SalarioNeto
Dim Categoria
Dim CentroCosto
Dim Localidad
Dim proFecPago
Dim pliqhasta
Dim pliqdesde
Dim FormaPago
Dim Puesto
Dim Banco
Dim nroCuenta

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
Dim UbicacionCod As Integer

'Auxchar
Dim Estructuracargo As String
Dim EstructuraAFP As String
Dim Campolibre As Double
Dim DescCampoLibre As String



'Dim Ubicacion As String
'-----
Dim EtiquetaFormaPago As String
Dim CodFormaPago As Integer
 
'Auxdeci
Dim HrsTrabajadas
Dim DiasVacaciones
Dim Faltas
Dim Diasdescmedico
Dim Diassubsidio
Dim Diaslicsg
Dim Diassusp
Dim Tardanza

' conceptos o acumuladores
Dim DiasVacacionesCO
Dim FaltasCO
Dim DiasdescmedicoCO
Dim DiassubsidioCO
Dim DiaslicsgCO
Dim DiassuspCO
Dim TardanzaCO

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3
'Dim ObraSocial
Dim CarnetAFP

Dim EstructuraUbicacion
Dim EstructuraSede
Dim EstructuraCCosto
Dim EstructuraArea

Dim periodo_vac_desde As String
Dim periodo_vac_hasta As String

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset

Dim valorSueldo As Double
Dim codSueldo As String
Dim codAfp As Long
Dim TipoDocRegPat
Dim RegPat
Dim Doc113
Dim TipoDoc113
Dim canthoras
Dim esconcepto
Dim AporteEmp
Dim AporteEmpDesc
Dim NroConcepto
Dim EmpTernro

Dim usaPagoDto As Boolean
Dim usadiasseparados As Boolean
Dim licVac As Integer
Dim CantDiasVac
Dim codConcVac

codSueldo = "0"

Dim moneda As String
Dim TipoCambio
Dim COtipocambio
On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

Flog.writeline "Modelo 179"

'------------------------------------------------------------------
'LEVANTO DEL CONFREP EL VALOR DE CARNET AFP
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 243"
StrSql = StrSql & " AND conftipo = 'DOC' "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    codAfp = rsConsult!confval
Else
    codAfp = 0
    Flog.writeline " No se configuro la columna 243 como tipo de DOC, es necesaria para obtener el carnet afp "
End If
rsConsult.Close

'------------------------------------------------------------------
'LEVANTO DEL CONFREP LOS VALORES DEL CONCEPTO SUELDO
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 242"
StrSql = StrSql & " AND conftipo = 'CO' "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Do While Not rsConsult.EOF
        codSueldo = codSueldo & ", " & rsConsult!confval2
    rsConsult.MoveNext
    Loop
Else
    Flog.writeline " No se configuro la columna 242 como tipo de concepto, es necesaria para obtener el suedo "
End If
rsConsult.Close

'------------------------------------------------------------------
'LEVANTO DEL CONFREP LOS VALORES PARA SABER SI USA concepto para vacaciones
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 390"
StrSql = StrSql & " AND upper(conftipo) = 'VAC' "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
      codConcVac = IIf(EsNulo(rsConsult!confval2), 0, rsConsult!confval2)
        Flog.writeline "Concepto vacaciones.: " & codConcVac
Else
      codConcVac = 0
        Flog.writeline "Concepto vacaciones.: " & codConcVac
End If
rsConsult.Close



'------------------------------------------------------------------
'LEVANTO DEL CONFREP LOS VALORES PARA SABER SI USA PAGO/DTO
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 390"
StrSql = StrSql & " AND upper(conftipo) = 'LIC' "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!confval = 1 Then
        usaPagoDto = True
        Flog.writeline "Usa pago/dto."
    Else
        usaPagoDto = False
        licVac = IIf(EsNulo(rsConsult!confval2), 0, rsConsult!confval2)
        Flog.writeline "No usa pago/dto, usa licencia de vacaciones. Tipo configurado: " & licVac
        
    End If
Else
    Flog.writeline " No se configuro la columna 242 como tipo de concepto, es necesaria para obtener el suedo "
End If
rsConsult.Close

'------------------------------------------------------------------
'Obtengo el nro de cabecera de liquidación y la fecha de pago del proceso
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
rsConsult.Close



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
   pliqdesde = rsConsult!pliqdesde
   'proDesc = proDesc & " - " & rsConsult!pliqdesde & " - " & pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & IIf(EsNulo(rsConsult!ternom2), "", rsConsult!ternom2)
   Apellido = rsConsult!terape & " " & IIf(EsNulo(rsConsult!terape2), "", rsConsult!terape2)
   Legajo = rsConsult!empleg
   'empFecAlta = rsConsult!empfaltagr
   If IsNull(rsConsult!empremu) Then
      Sueldo = 0
   Else
      Sueldo = rsConsult!empremu
   End If
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco la fecha de Alta y Baja de la última fase Activa
'------------------------------------------------------------------
StrSql = "SELECT empleado,altfec,bajfec FROM fases "
StrSql = StrSql & " WHERE empleado= " & Ternro
StrSql = StrSql & " ORDER BY altfec DESC"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   empFecAlta = rsConsult!altfec
   empFecBaja = IIf(EsNulo(rsConsult!bajfec), "", rsConsult!bajfec)
   
   '********************************
   'FALTA AGREGAR AL INSERT (BAJFEC)
   '********************************
Else
   Flog.writeline "No existe fase para el empleado ternro:" & Ternro
   GoTo MError
End If
rsConsult.Close


'------------------------------------------------------------------
'Busco la estructura "CARGO" desde el confrep | COLUMNA FIJA 117 (CARGO)
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = (SELECT confval from confrep WHERE repnro = 60 and  confnrocol=117)  And his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   'Departamento = rsConsult!Estrnro & "@@" & rsConsult!estrdabr
   Estructuracargo = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
   Flog.writeline "No se pudo obtener los datos de cargo. Revisar configuración de Columna 117 del confrep"
   'GoTo MError
End If
rsConsult.Close
 
'------------------------------------------------------------------
'Busco la estructura "Tipo de trabajador" desde el confrep | COLUMNA FIJA 376 (Tipo de trabajador)
'------------------------------------------------------------------
EstructuraTipotrabajador = ""
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = (SELECT confval from confrep WHERE repnro = 60 and  confnrocol=376)  And his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   'Departamento = rsConsult!Estrnro & "@@" & rsConsult!estrdabr
   EstructuraTipotrabajador = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
   Flog.writeline "No se pudo obtener los datos de la estructura Tipo de trabajador. Revisar configuración de Columna 376 del confrep"
   'GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco la estructura "Tipo de trabajador" desde el confrep | COLUMNA FIJA 377 (Ocupación)
'------------------------------------------------------------------
EstructuraOcupacion = ""
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = (SELECT confval from confrep WHERE repnro = 60 and  confnrocol=377)  And his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   'Departamento = rsConsult!Estrnro & "@@" & rsConsult!estrdabr
   EstructuraOcupacion = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
   Flog.writeline "No se pudo obtener los datos de la estructura Ocupación. Revisar configuración de Columna 377 del confrep"
   'GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco la estructura "AFP" desde el confrep | COLUMNA FIJA 118 (AFP)
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = (SELECT confval from confrep WHERE repnro = 60 and  confnrocol=118)  And his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   'Departamento = rsConsult!Estrnro & "@@" & rsConsult!estrdabr
   EstructuraAFP = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
   Flog.writeline "No se pudo obtener los datos de AFP. Revisar configuración de Columna 118 del confrep"
   'GoTo MError
End If
rsConsult.Close





'------------------------------------------------------------------
'Busco la estructura libre desde el confrep | COLUMNA FIJA 797 (Campo libre)
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = (SELECT confval from confrep WHERE repnro = 60 and  confnrocol=797)  And his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   'Departamento = rsConsult!Estrnro & "@@" & rsConsult!estrdabr
   DescCampoLibre = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
   Flog.writeline "No se pudo obtener los datos del campo libre. Revisar configuración de Columna 797 del confrep"
   'GoTo MError
End If
rsConsult.Close



'------------------------------------------------------------------
'BUSCO EL VALOR DEL CAMPO SUELDO DEL RECIBO
'------------------------------------------------------------------
If codSueldo <> "0" Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod IN (" & codSueldo & ")"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
       valorSueldo = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
    Else
       Flog.writeline "Error al obtener el valor del campo sueldo."
       valorSueldo = 0
    End If
    rsConsult.Close
Else
    valorSueldo = 0
End If
 

'------------------------------------------------------------------
'Busco el valor de los días trabajados | COLUMNA 49
'------------------------------------------------------------------
If ConcDiasTrabajados <> "" And Not IsNull(ConcDiasTrabajados) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & ConcDiasTrabajados & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
       DiasTrabajados = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
    Else
       Flog.writeline "Error al obtener los dias trabajados."
       DiasTrabajados = 0
    End If
    rsConsult.Close
Else
    DiasTrabajados = 0
End If
 
'------------------------------------------------------------------
'Busco el valor de la cantidad de horas | COLUMNA 90
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 90 "
Flog.writeline "Horas Trabajadas" & StrSql
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!conftipo = "CO" Then
        canthoras = rsConsult!confval
        esconcepto = True
    Else
        canthoras = rsConsult!confval2
        esconcepto = False
    End If
End If
rsConsult.Close

If esconcepto Then
    StrSql = "SELECT detliq.dlimonto valor"
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
    StrSql = StrSql & " AND concepto.conccod = " & canthoras
Else
     StrSql = " SELECT almonto valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & canthoras
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
Flog.writeline "Horas Trabajadas" & StrSql
If Not rsConsult.EOF Then
    HrsTrabajadas = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
Else
    HrsTrabajadas = 0
 End If
 
rsConsult.Close


DiasVacacionesCO = 0
'------------------------------------------------------------------
'Busco el valor de la cantidad de Dias de vacaciones | COLUMNA 790
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 790 "
Flog.writeline "Dias de vacaciones" & StrSql
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!conftipo = "CO" Then
        DiasVacacionesCO = rsConsult!confval2
        esconcepto = True
    Else
        DiasVacacionesCO = rsConsult!confval
        esconcepto = False
    End If
End If
rsConsult.Close

If esconcepto Then
    StrSql = "SELECT detliq.dlicant valor"
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
    StrSql = StrSql & " AND concepto.conccod = " & DiasVacacionesCO
Else
     StrSql = " SELECT alcant valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & DiasVacacionesCO
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If

OpenRecordset StrSql, rsConsult
Flog.writeline "Dias de vacaciones : " & StrSql
If Not rsConsult.EOF Then
    DiasVacaciones = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
Else
    DiasVacaciones = 0
 End If
 
rsConsult.Close


FaltasCO = 0

'------------------------------------------------------------------
'Busco el valor de la cantidad de faltas | COLUMNA 791
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 791 "
Flog.writeline "FAltas" & StrSql
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!conftipo = "CO" Then
        FaltasCO = rsConsult!confval2
        esconcepto = True
    Else
        FaltasCO = rsConsult!confval
        esconcepto = False
    End If
End If
rsConsult.Close

If esconcepto Then
    StrSql = "SELECT detliq.dlicant valor"
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
    StrSql = StrSql & " AND concepto.conccod = " & FaltasCO
Else
     StrSql = " SELECT alcant valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & FaltasCO
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If

OpenRecordset StrSql, rsConsult
Flog.writeline "faltas : " & StrSql
If Not rsConsult.EOF Then
    Faltas = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
Else
    Faltas = 0
 End If
 
rsConsult.Close


DiasdescmedicoCO = 0

'------------------------------------------------------------------
'Busco el valor de la cantidad de Dias desc med | COLUMNA 792
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 792 "
Flog.writeline "Dias desc med" & StrSql
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!conftipo = "CO" Then
        DiasdescmedicoCO = rsConsult!confval2
        esconcepto = True
    Else
        DiasdescmedicoCO = rsConsult!confval
        esconcepto = False
    End If
End If
rsConsult.Close

If esconcepto Then
    StrSql = "SELECT detliq.dlicant valor"
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
    StrSql = StrSql & " AND concepto.conccod = " & DiasdescmedicoCO
Else
     StrSql = " SELECT alcant valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & DiasdescmedicoCO
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
Flog.writeline "Dias des med : " & StrSql
If Not rsConsult.EOF Then
    Diasdescmedico = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
Else
    Diasdescmedico = 0
 End If
 
rsConsult.Close


DiassubsidioCO = 0

'------------------------------------------------------------------
'Busco el valor de la cantidad de Dias subsidio | COLUMNA 793
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 793 "
Flog.writeline "Dias de subsidio" & StrSql
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!conftipo = "CO" Then
        DiassubsidioCO = rsConsult!confval2
        esconcepto = True
    Else
        DiassubsidioCO = rsConsult!confval
        esconcepto = False
    End If
End If
rsConsult.Close

If esconcepto Then
    StrSql = "SELECT detliq.dlicant valor"
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
    StrSql = StrSql & " AND concepto.conccod = " & DiassubsidioCO
Else
     StrSql = " SELECT alcant valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & DiassubsidioCO
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
Flog.writeline "Dias de subsidio : " & StrSql
If Not rsConsult.EOF Then
    Diassubsidio = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
Else
    Diassubsidio = 0
 End If
 
rsConsult.Close


DiaslicsgCO = 0

'------------------------------------------------------------------
'Busco el valor de la cantidad de Dias Lic S/G | COLUMNA 794
'-----------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 794 "
Flog.writeline "Dias de lic s/g" & StrSql
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!conftipo = "CO" Then
        DiaslicsgCO = rsConsult!confval2
        esconcepto = True
    Else
        DiaslicsgCO = rsConsult!confval
        esconcepto = False
    End If
End If
rsConsult.Close

If esconcepto Then
    StrSql = "SELECT detliq.dlicant valor"
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
    StrSql = StrSql & " AND concepto.conccod = " & DiaslicsgCO
Else
     StrSql = " SELECT alcant valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & DiaslicsgCO
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
Flog.writeline "Dias Lic S/G : " & StrSql
If Not rsConsult.EOF Then
    Diaslicsg = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
Else
    Diaslicsg = 0
 End If
 
rsConsult.Close

DiassuspCO = 0

'------------------------------------------------------------------
'Busco el valor de la cantidad de Dias de suspension| COLUMNA 795
'-----------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 795 "
Flog.writeline "Dias de suspension" & StrSql
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!conftipo = "CO" Then
        DiassuspCO = rsConsult!confval2
        esconcepto = True
    Else
        DiassuspCO = rsConsult!confval
        esconcepto = False
    End If
End If
rsConsult.Close

If esconcepto Then
    StrSql = "SELECT detliq.dlicant valor"
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
    StrSql = StrSql & " AND concepto.conccod = " & DiassuspCO
Else
     StrSql = " SELECT alcant valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & DiassuspCO
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
Flog.writeline "Dias de suspension: " & StrSql
If Not rsConsult.EOF Then
    Diassusp = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
Else
    Diassusp = 0
 End If
 
rsConsult.Close

'------------------------------------------------------------------------------------------
TardanzaCO = 0
'------------------------------------------------------------------
'Busco el valor de la cantidad de Tardanza | COLUMNA 796
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND confnrocol = 796 "
Flog.writeline "Tardanza:" & StrSql
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!conftipo = "CO" Then
        TardanzaCO = rsConsult!confval2
        esconcepto = True
    Else
        TardanzaCO = rsConsult!confval
        esconcepto = False
    End If
End If
rsConsult.Close

If esconcepto Then
    StrSql = "SELECT detliq.dlicant valor"
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
    StrSql = StrSql & " AND concepto.conccod = " & TardanzaCO
Else
     StrSql = " SELECT alcant valor"
     StrSql = StrSql & " From acu_liq"
     StrSql = StrSql & " Where acunro = " & TardanzaCO
     StrSql = StrSql & " AND cliqnro = " & cliqnro
End If

OpenRecordset StrSql, rsConsult
Flog.writeline "Dias Tardanza  : " & StrSql
If Not rsConsult.EOF Then
    Tardanza = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
Else
    Tardanza = 0
 End If
 
rsConsult.Close


'------------------------------------------------------------------
'Busco el valor del tipo de cambio -- Columna tipo TCA
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND conftipo = 'TCA' "
Flog.writeline "Concepto Tipo de cambio : " & StrSql
OpenRecordset StrSql, rsConsult
COtipocambio = 0
If Not rsConsult.EOF Then
    If Not IsNull(rsConsult!confval2) And rsConsult!confval2 <> "" Then
        COtipocambio = rsConsult!confval2
    Else
        COtipocambio = 0
    End If
End If
rsConsult.Close

    StrSql = "SELECT detliq.dlimonto valor"
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
    StrSql = StrSql & " AND concepto.conccod = " & COtipocambio
    OpenRecordset StrSql, rsConsult
Flog.writeline "Consulta Tipo de cambio :" & StrSql
If Not rsConsult.EOF Then
    TipoCambio = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
Else
    TipoCambio = 0
 End If
rsConsult.Close

FormaPago = ""

'---------------------sebastian stremel 05/04/2013-----------------

'------------------------------------------------------------------
'Busco la estructura "UBICACION" desde el confrep | COLUMNA FIJA 300 (UBICACION)
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = (SELECT confval from confrep WHERE repnro = 60 and  confnrocol=300)  And his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    EstructuraUbicacion = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
    EstructuraUbicacion = ""
    Flog.writeline "No se pudo obtener los datos de ubicacion. Revisar configuración de Columna 300 del confrep"
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco la estructura "SEDE" desde el confrep | COLUMNA FIJA 301 (SEDE)
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = (SELECT confval from confrep WHERE repnro = 60 and  confnrocol=301)  And his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    EstructuraSede = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
    EstructuraSede = ""
    Flog.writeline "No se pudo obtener los datos de sede. Revisar configuración de Columna 301 del confrep"
   'GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco la estructura "CENTRO DE COSTO" desde el confrep | COLUMNA FIJA 302 (C.COSTO)
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = (SELECT confval from confrep WHERE repnro = 60 and  confnrocol=302)  And his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    EstructuraCCosto = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
    EstructuraCCosto = ""
    Flog.writeline "No se pudo obtener los datos de centro de costo. Revisar configuración de Columna 302 del confrep"
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco la estructura "AREA" desde el confrep | COLUMNA FIJA 303 (AREA)
'------------------------------------------------------------------
StrSql = " SELECT estrdabr,estrcodext "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = (SELECT confval from confrep WHERE repnro = 60 and  confnrocol=303)  And his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    EstructuraArea = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
    EstructuraArea = ""
    Flog.writeline "No se pudo obtener los datos del area. Revisar configuración de Columna 303 del confrep"
End If
rsConsult.Close

'-----------------------HASTA ACA---------------------------------

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
           EmpEstrnro1 = IIf(EsNulo(rsConsult!Estrnro), "", rsConsult!Estrnro)
        End If
        rsConsult.Close
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
           EmpEstrnro2 = IIf(EsNulo(rsConsult!Estrnro), "", rsConsult!Estrnro)
        End If
        rsConsult.Close
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
           EmpEstrnro3 = IIf(EsNulo(rsConsult!Estrnro), "", rsConsult!Estrnro)
        End If
        rsConsult.Close
    End If
End If

'------------------------------------------------------------------
'Busco Tipo de documento 1 ó 3
'------------------------------------------------------------------
StrSql = "SELECT doc.nrodoc FROM tercero "
StrSql = StrSql & " INNER JOIN ter_doc doc ON tercero.ternro = doc.ternro "
StrSql = StrSql & " INNER JOIN tipodocu_pais td ON td.tidnro = doc.tidnro  and td.tidcod <= 5 "
StrSql = StrSql & " WHERE tercero.ternro =" & Ternro
StrSql = StrSql & " ORDER BY tidcod ASC "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    NroDoc = IIf(EsNulo(rsConsult!NroDoc), "", rsConsult!NroDoc)
Else
    Flog.writeline "Error al obtener los datos del Documento para el Legajo: " & Legajo
   'GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco Tipo de documento tipo 7 | CARNET AFP
'------------------------------------------------------------------
StrSql = " SELECT ter_doc.ternro,ter_doc.tidnro, ter_doc.nrodoc FROM tercero"
StrSql = StrSql & " LEFT JOIN ter_doc ON tercero.ternro=ter_doc.ternro"
StrSql = StrSql & " WHERE Tercero.Ternro = " & Ternro
StrSql = StrSql & " AND ter_doc.tidnro=" & codAfp
StrSql = StrSql & " ORDER BY tidnro ASC"
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    CarnetAFP = IIf(EsNulo(rsConsult!NroDoc), "", rsConsult!NroDoc)
Else
    CarnetAFP = ""
    Flog.writeline "Error al obtener los datos del Documento para el Legajo: " & Legajo
   'GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la direccion y localidad de la sucursal
'------------------------------------------------------------------
StrSql = " SELECT detdom.calle,detdom.nro,detdom.piso,detdom.oficdepto,localidad.locdesc"
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(proFecPago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(proFecPago) & ") AND tenro=1 AND his_estructura.ternro=" & Ternro
StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
       
OpenRecordset StrSql, rsConsult
Direccion = ""
If Not rsConsult.EOF Then
    Direccion = IIf(EsNulo(rsConsult!calle), "", rsConsult!calle) & " - " & IIf(EsNulo(rsConsult!nro), "", rsConsult!nro)
    If Not EsNulo(rsConsult!Piso) Then
        Direccion = Direccion & " P. " & IIf(EsNulo(rsConsult!Piso), "", rsConsult!Piso)
    End If
    If Not EsNulo(rsConsult!oficdepto) Then
        Direccion = Direccion & " Dpto. " & IIf(EsNulo(rsConsult!oficdepto), "", rsConsult!oficdepto)
    End If
    
    Direccion = Direccion & ", " & IIf(EsNulo(rsConsult!locdesc), "", rsConsult!locdesc)
   
    Localidad = Direccion
Else
'   Flog.writeline "Error al obtener los datos de la localidad"
'   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

' ahora si tiene valor ==> lo recalcula sino deja el empremu
'If Sueldo = 0 Then
    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
       Sueldo = IIf(EsNulo(rsConsult!almonto), 0, rsConsult!almonto)
    Else
       Flog.writeline "Error al obtener los datos del sueldo. Se queda con el valor de la remuneracion cargada en ADP."
       'Sueldo = 0
       'GoTo MError
    End If
    rsConsult.Close
'End If
'------------------------------------------------------------------
'Busco el valor del acumulador Salario Neto SALE DE LA COLUMNA 2 DEL CONFREP
'------------------------------------------------------------------
'Busco la configuracion del confrep
Flog.writeline "Obtengo el acumulador neto desde el confrep"
       
    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = (select confval from confrep WHERE repnro = 60 and  confnrocol = 2)"
    StrSql = StrSql & " AND cliqnro = " & cliqnro
    OpenRecordset StrSql, rsConsult

    If Not rsConsult.EOF Then
       SalarioNeto = IIf(EsNulo(rsConsult!almonto), 0, rsConsult!almonto)
    Else
       Flog.writeline "Error al obtener los datos del acumulador Salario Neto"
       SalarioNeto = 0
       'GoTo MError
    End If
    rsConsult.Close
'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------
Categoria = ""

'------------------------------------------------------------------
'Busco el valor del puesto | COLUMNA FIJA 117 (CARGO)
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult

Puesto = ""
If Not rsConsult.EOF Then
   Puesto = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del centro de costo
''------------------------------------------------------------------
CentroCosto = ""
  
'------------------------------------------------------------------
'Busco el banco y la cuenta del empleado
'------------------------------------------------------------------
Banco = ""
nroCuenta = ""

StrSql = " SELECT banco.bandesc, ctabancaria.ctabnro, ctabancaria.ctabcbu "
StrSql = StrSql & " From ctabancaria "
StrSql = StrSql & " INNER JOIN banco ON ctabancaria.banco = banco.ternro "
StrSql = StrSql & " WHERE ctabancaria.ternro= " & Ternro
StrSql = StrSql & " AND ctabancaria.ctabestado = -1"

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    If Not EsNulo(rsConsult!ctabnro) Then
       Banco = rsConsult!Bandesc
       nroCuenta = rsConsult!ctabnro
    End If
Else
    Flog.writeline "Error al obtener los datos de Cuenta + Banco"
End If

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

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
    EmpNombre = IIf(EsNulo(rs_estructura!empnom), "", rs_estructura!empnom)
    EmpEstrnro = IIf(EsNulo(rs_estructura!Estrnro), "", rs_estructura!Estrnro)
    EmpTernro = IIf(EsNulo(rs_estructura!Ternro), 0, rs_estructura!Ternro)
End If
'rs_estructura.Close

StrSql = " SELECT * FROM ter_imag "
StrSql = StrSql & " INNER JOIN tipoimag ON tipoimag.tipimnro=ter_imag.tipimnro "
StrSql = StrSql & " WHERE ternro =" & rs_estructura!Ternro
StrSql = StrSql & " AND ter_imag.tipimnro=1"
OpenRecordset StrSql, rs_Domicilio
If Not rs_Domicilio.EOF Then
    EmpLogo = IIf(EsNulo(rs_Domicilio!tipimdire), "", rs_Domicilio!tipimdire) & IIf(EsNulo(rs_Domicilio!terimnombre), "", rs_Domicilio!terimnombre)
    EmpLogoAncho = IIf(EsNulo(rs_Domicilio!tipimanchodef), 0, rs_Domicilio!tipimanchodef)
    EmpLogoAlto = IIf(EsNulo(rs_Domicilio!tipimaltodef), 0, rs_Domicilio!tipimaltodef)
    Flog.writeline "Se encontro el logo de la empresa: " & EmpLogo
Else
    EmpLogo = ""
    EmpLogoAncho = 0
    EmpLogoAlto = 0
    Flog.writeline "No se encontro el logo de la empresa"
End If
rs_Domicilio.Close

'Consulta para obtener la direccion de la empresa
'Direccion + Numero + Piso + Mza + Lote + Nombre de zona + Distrito
' se cambia el formato de la direccion de la empresa
'Via + Direccion + Numero + Piso + Mza + Lote + Block + Nombre de zona + Distrito + Departamento + Provincia
StrSql = "SELECT via.viadesc, detdom.calle, detdom.nro , detdom.piso,  detdom.manzana, detdom.lote " & _
    " ,detdom.bloque,   detdom.auxchr4, localidad.locdesc, provincia.provdesc,partido.partnom  From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.domdefault = -1 AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN partido ON detdom.partnro = partido.partnro " & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro " & _
    " LEFT JOIN via ON detdom.vianro = via.vianro " & _
    " LEFT JOIN provincia ON detdom.provnro = provincia.provnro "
 Flog.writeline "domicilio empresa: " & StrSql
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    EmpDire = "   "
Else
    If Not EsNulo(rs_Domicilio!viadesc) Then
        EmpDire = rs_Domicilio!viadesc
    Else
        EmpDire = ""
    End If

    EmpDire = EmpDire & "@" & IIf(EsNulo(rs_Domicilio!calle), "", rs_Domicilio!calle) & "@" & IIf(EsNulo(rs_Domicilio!nro), "", rs_Domicilio!nro)
    If Not EsNulo(rs_Domicilio!Piso) Then
        EmpDire = EmpDire & "@" & rs_Domicilio!Piso
    Else
        EmpDire = EmpDire & "@" & ""
    End If
    
    If Not EsNulo(rs_Domicilio!manzana) Then
        EmpDire = EmpDire & "@" & rs_Domicilio!manzana
    Else
        EmpDire = EmpDire & "@" & ""
    End If
    
    If Not EsNulo(rs_Domicilio!lote) Then
        EmpDire = EmpDire & "@" & rs_Domicilio!lote
    Else
        EmpDire = EmpDire & "@" & ""
    End If
    
     If Not EsNulo(rs_Domicilio!bloque) Then
        EmpDire = EmpDire & "@" & rs_Domicilio!bloque
    Else
        EmpDire = EmpDire & "@" & ""
    End If
    
    If Not EsNulo(rs_Domicilio!auxchr4) Then
        EmpDire = EmpDire & "@" & rs_Domicilio!auxchr4
    Else
        EmpDire = EmpDire & "@" & ""
    End If
    
    If Not EsNulo(rs_Domicilio!locdesc) Then
       EmpDire = EmpDire & "@" & rs_Domicilio!locdesc
    Else
        EmpDire = EmpDire & "@" & ""
    End If
    
    If Not EsNulo(rs_Domicilio!provdesc) Then
       EmpDire = EmpDire & "@" & rs_Domicilio!provdesc
    Else
        EmpDire = EmpDire & "@" & ""
    End If
    
    If Not EsNulo(rs_Domicilio!partnom) Then
       EmpDire = EmpDire & "@" & rs_Domicilio!partnom
    Else
        EmpDire = EmpDire & "@" & ""
    End If
    

End If
rs_Domicilio.Close

'Consulta para obtener el RUC de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero "
StrSql = StrSql & " INNER JOIN ter_doc cuit ON tercero.ternro = cuit.ternro "
StrSql = StrSql & " INNER JOIN tipodocu_pais td ON td.tidnro = cuit.tidnro  and td.tidcod = 6 "
StrSql = StrSql & " WHERE tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el tipo de documento RUC de la Empresa"
    EmpCuit = ""
Else
    EmpCuit = IIf(EsNulo(rs_cuit!NroDoc), "", rs_cuit!NroDoc)
End If
rs_cuit.Close


'Firma de Recibo
StrSql = " SELECT estrcodext "
StrSql = StrSql & " From estructura"
StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(profecfin) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(profecfin) & ") AND his_estructura.tenro=52 AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   FirmaRecibo = rsConsult!estrcodext
Else
   FirmaRecibo = 2
End If


'Consulta para buscar la firma de la empresa
'StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
'    " From ter_imag " & _
'    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
'    " AND ter_imag.ternro =" & rs_estructura!Ternro
'OpenRecordset StrSql, rs_firma
'If rs_firma.EOF Then
'    Flog.writeline "No se encontró el Firma de la Empresa."
'    'Exit Sub
'    EmpFirma = ""
'    EmpFirmaAlto = 0
'    EmpFirmaAncho = 0
'Else
'    EmpFirma = IIf(EsNulo(rs_firma!tipimdire), "", rs_firma!tipimdire) & IIf(EsNulo(rs_firma!terimnombre), "", rs_firma!terimnombre)
'    EmpFirmaAlto = IIf(EsNulo(rs_firma!tipimaltodef), "", rs_firma!tipimaltodef)
'    EmpFirmaAncho = IIf(EsNulo(rs_firma!tipimanchodef), "", rs_firma!tipimanchodef)
'End If
'Flog.writeline "Cierro Firma"
'rs_estructura.Close
'rs_firma.Close



'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = " & FirmaRecibo & " AND tipoimag.tipimnro = ter_imag.tipimnro" & _
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
   pliqdepant = ""
   pliqfecdep = ""
   pliqbco = ""

'________________________________________________________________
'BUSCO EL PERIODO VACACIONAL
If usaPagoDto Then
    StrSql = " SELECT min(elfechadesde) desde, max(elfechahasta) hasta  FROM emp_lic "
    StrSql = StrSql & " INNER JOIN vacpagdesc ON vacpagdesc.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " WHERE Empleado =" & Ternro & " And tdnro = 2 And pliqnro = " & pliqnro
    StrSql = StrSql & " AND vacpagdesc.pronro =" & Pronro
    StrSql = StrSql & " AND pago_dto in(1,3)" '1 y 3 son los codigos de pago
    OpenRecordset StrSql, rs_cuit
    If Not rs_cuit.EOF Then
        Flog.writeline "Se encontro el periodo vacacional"
        periodo_vac_desde = IIf(EsNulo(rs_cuit!Desde), "", rs_cuit!Desde)
        periodo_vac_hasta = IIf(EsNulo(rs_cuit!Hasta), "", rs_cuit!Hasta)
    Else
        Flog.writeline "No se encontro el periodo vacacional, query: " & StrSql
        periodo_vac_desde = ""
        periodo_vac_hasta = ""
    End If
    Flog.writeline "Cierro Periodo Vacacional"
    rs_cuit.Close
Else
  StrSql = " SELECT min(elfechadesde) desde, max(elfechahasta) hasta ,SUM(elcantdias) dias"
    StrSql = StrSql & " FROM emp_lic "
    StrSql = StrSql & " WHERE Empleado = " & Ternro
    StrSql = StrSql & " AND tdnro = " & licVac
    StrSql = StrSql & " AND pronro = " & Pronro
    'StrSql = StrSql & " WHERE ("
    'StrSql = StrSql & " (elfechadesde <= " & ConvFecha(pliqdesde) & " AND (elfechahasta is null or elfechahasta >= " & ConvFecha(pliqhasta)
    'StrSql = StrSql & " or elfechahasta >= " & ConvFecha(pliqdesde) & " )) OR"
    'StrSql = StrSql & " (elfechadesde >= " & ConvFecha(pliqdesde) & " AND (elfechadesde <= " & ConvFecha(pliqhasta) & "))"
    'StrSql = StrSql & " )"
    OpenRecordset StrSql, rs_cuit
    If Not rs_cuit.EOF Then
        Flog.writeline "Se encontro el periodo vacacional"
        periodo_vac_desde = IIf(EsNulo(rs_cuit!Desde), "", rs_cuit!Desde)
        periodo_vac_hasta = IIf(EsNulo(rs_cuit!Hasta), "", rs_cuit!Hasta)
        
    Else
        Flog.writeline "No se encontro el periodo vacacional, query: " & StrSql
        periodo_vac_desde = ""
        periodo_vac_hasta = ""
    End If
    rs_cuit.Close
End If


'------------------------------------------------------------------
'BUSCO EL VALOR DEL Concepto de vacaciones
'------------------------------------------------------------------
If codConcVac <> "0" Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod IN (" & codConcVac & ")"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        
       CantDiasVac = IIf(EsNulo(rsConsult!Valor), 0, rsConsult!Valor)
       periodo_vac_desde = pliqdesde
       periodo_vac_hasta = DateAdd("d", CantDiasVac, periodo_vac_desde)
       periodo_vac_hasta = DateAdd("d", -1, periodo_vac_hasta)
    
       Flog.writeline "periodo vacacional desde: " & periodo_vac_desde
       Flog.writeline "periodo vacacional hasta: " & periodo_vac_hasta
    End If
    rsConsult.Close
End If
 

'------------------------------------------------------------------
'LEVANTO DEL CONFREP EL TIPO DOCUMENTO REG. PAT.
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND conftipo = 'REG' "
OpenRecordset StrSql, rsConsult
Flog.writeline " REG: PAT. " & StrSql
If Not rsConsult.EOF Then
    TipoDocRegPat = rsConsult!confval
Else
    TipoDocRegPat = 0
    Flog.writeline " No se configuro el tipo de DOC Reg Pat"
End If
rsConsult.Close

'Busco el Tipo de Documento REG. PAT
StrSql = " SELECT tidnom, nrodoc FROM ter_doc "
StrSql = StrSql & " INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro"
StrSql = StrSql & " WHERE ternro=" & EmpTernro
StrSql = StrSql & " AND ter_doc.tidnro=" & TipoDocRegPat
OpenRecordset StrSql, rsConsult
Flog.writeline " REG: PAT. " & StrSql
If Not rsConsult.EOF Then
    RegPat = IIf(EsNulo(rsConsult!NroDoc), "", rsConsult!NroDoc)
Else
    RegPat = 0
End If

'------------------------------------------------------------------
'LEVANTO DEL CONFREP EL TIPO DOCUMENTO 113
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND conftipo = 'SUP' "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    TipoDoc113 = rsConsult!confval
Else
    TipoDoc113 = 0
    Flog.writeline " No se configuro el tipo de DOC 113"
End If
rsConsult.Close

'Busco el Tipo de Documento 113
StrSql = " SELECT tidnom, nrodoc FROM ter_doc "
StrSql = StrSql & " INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro"
StrSql = StrSql & " WHERE ternro=" & EmpTernro
StrSql = StrSql & " AND ter_doc.tidnro=" & TipoDoc113
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Doc113 = IIf(EsNulo(rsConsult!NroDoc), "", rsConsult!NroDoc)
Else
    Doc113 = 0
End If

'------------------------------------------------------------------
'Busco el valor de los Aportes del Empleador
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=60 "
StrSql = StrSql & " AND conftipo= 'EMP' "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
        NroConcepto = rsConsult!confval
Else
    NroConcepto = 0
End If
rsConsult.Close

StrSql = "SELECT detliq.dlimonto valor, concabr"
StrSql = StrSql & " FROM detliq "
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro
StrSql = StrSql & " AND concepto.conccod = " & NroConcepto
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    AporteEmp = Replace(rsConsult!Valor, ",", ".")
    AporteEmpDesc = rsConsult!concabr
Else
    AporteEmp = 0
    AporteEmpDesc = ""
End If
rsConsult.Close

'________________________________________________________________
'_____________________Busco la forma de pago _________________

StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And (his_estructura.estrnro IN (SELECT confval from confrep WHERE repnro = 60 and  conftipo='FFP')  And his_estructura.Ternro = " & Ternro & ")"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then

   FormaPago = "-1"
   proDesc = proDesc & " - " & pliqdesde & " - " & pliqhasta
Else
   FormaPago = "0"
   Flog.writeline "No se pudo obtener los datos de la forma de pago"
 
End If
 Flog.writeline "forma de pago:" & StrSql
rsConsult.Close

'_____________________Busco la estructura moneda_________________
StrSql = " SELECT estrdabr "
StrSql = StrSql & " FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " INNER JOIN confrep c ON c.confval = his_estructura.tenro and c.repnro=60 and c.confnrocol=397 and c.conftipo='TE' "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = c.confval And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
 Flog.writeline "Consulta moneda:" & StrSql
If Not rsConsult.EOF Then
    moneda = IIf(EsNulo(rsConsult!estrdabr), "", rsConsult!estrdabr)
Else
    Flog.writeline "No se encontro la estructura moneda."
    moneda = ""
End If
rsConsult.Close
'________________________________________________________________

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " ("
StrSql = StrSql & " bpronro,ternro,pronro"
StrSql = StrSql & " ,apellido,Nombre,direccion,Legajo"
StrSql = StrSql & " ,pliqnro,pliqmes,pliqanio,pliqdepant"
StrSql = StrSql & " ,pliqfecdep,pliqbco,cuil,empfecalta"
StrSql = StrSql & " ,sueldo,categoria,centrocosto,localidad"
StrSql = StrSql & " ,profecpago,empnombre,empdire,empcuit,emplogo,emplogoalto,emplogoancho,empfirma"
StrSql = StrSql & " ,empfirmaalto,empfirmaancho,formapago,prodesc,descripcion "
StrSql = StrSql & " ,tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden"
StrSql = StrSql & " ,auxchar1,auxchar2,auxchar3,auxchar4,auxchar5"
StrSql = StrSql & " ,auxdeci1,auxdeci2,auxdeci3,auxchar6"
StrSql = StrSql & " ,auxchar7,auxchar8,auxchar9,auxchar10,auxchar11,auxchar12,auxchar13, auxdeci4,auxdeci7,puesto,auxchar,auxdeci5"
StrSql = StrSql & " ,auxdeci10,auxdeci11,auxdeci12,auxdeci13,auxdeci14,auxdeci15,auxdeci16,auxchar14) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & NroDoc & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & Replace(Sueldo, ",", ".")
StrSql = StrSql & ",'" & Doc113 & "'" 'Categoria
StrSql = StrSql & ",'" & Mid(EstructuraCCosto, 1, 100) & "'"
StrSql = StrSql & ",'" & RegPat & "'" ' Localidad
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & FormaPago & "'"
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden

'Auxchar
StrSql = StrSql & ",'" & CarnetAFP & "'"

StrSql = StrSql & ",'" & periodo_vac_desde & "'"
StrSql = StrSql & ",'" & periodo_vac_hasta & "'"

StrSql = StrSql & ",'" & EstructuraAFP & "'"   'AFP
StrSql = StrSql & ",'" & Estructuracargo & "'" ' CARGO

'Auxdeci
StrSql = StrSql & "," & Replace(SalarioNeto, ",", ".")
StrSql = StrSql & "," & DiasTrabajados
StrSql = StrSql & "," & HrsTrabajadas
StrSql = StrSql & ",'" & IIf(EsNulo(empFecBaja), "", empFecBaja) & "'" 'sebastian stremel

'auxchar
StrSql = StrSql & ",'" & EstructuraUbicacion & "'"
StrSql = StrSql & ",'" & EstructuraSede & "'"
StrSql = StrSql & ",'" & EstructuraArea & "'"
'auxchar10 y auxchar11, 12 y 13
StrSql = StrSql & ",'" & Banco & "'"
StrSql = StrSql & ",'" & nroCuenta & "'"
StrSql = StrSql & ",'" & EstructuraTipotrabajador & "'"
StrSql = StrSql & ",'" & EstructuraOcupacion & "'"
'Auxdeci
StrSql = StrSql & "," & IIf(EsNulo(valorSueldo), 0, valorSueldo)
StrSql = StrSql & "," & AporteEmp
StrSql = StrSql & ",'" & AporteEmpDesc & "'"
StrSql = StrSql & ",'" & moneda & "'"
StrSql = StrSql & "," & TipoCambio
StrSql = StrSql & "," & DiasVacaciones
StrSql = StrSql & "," & Faltas
StrSql = StrSql & "," & Diasdescmedico
StrSql = StrSql & "," & Diassubsidio
StrSql = StrSql & "," & Diaslicsg
StrSql = StrSql & "," & Diassusp
StrSql = StrSql & "," & Tardanza
StrSql = StrSql & ",'" & DescCampoLibre
StrSql = StrSql & "')"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

'Flog.writeline "SQL INSERT: " & StrSql

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

Flog.writeline "----------------------------------------------------------------------------------------"
Flog.writeline ""
Flog.writeline ""

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
' Se encarga de generar los datos para el recibo de Deloitte Personal
'--------------------------------------------------------------------
Sub generarDatosRecibo188(Pronro, Ternro, acunroSueldo, tituloReporte, orden)

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
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim Cuil
Dim empFecAlta
Dim Sueldo
Dim Categoria
Dim CentroCosto
Dim Localidad
Dim proFecPago
Dim pliqhasta
Dim FormaPago
Dim Puesto

Dim empFecBaja
Dim oSocial
Dim Gerencia
Dim reportaa
Dim empreporta
Dim grupoSeguridad
Dim Sucursal

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

Dim RRHHFirma As String
Dim RRHHFirmaAlto As Integer
Dim RRHHFirmaAncho As Integer
    
Dim proDesc As String
Dim tidnroRST As String
Dim codResolucion As String

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim Fecha_aux As Date
Dim anios_aux As Integer
Dim meses_aux As Integer
Dim dias_aux As Integer
Dim diasHab_aux As Integer
Dim Antiguedad

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

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
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu, empfecbaja, empreporta "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   Apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!empleg
   empFecAlta = rsConsult!empfaltagr
   empFecBaja = rsConsult!empFecBaja
   empreporta = rsConsult!empreporta
   
Else
   Flog.writeline "Error al obtener los datos del empleado"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco la fecha de alta y baja
'------------------------------------------------------------------

StrSql = " SELECT * FROM fases WHERE empleado = " & Ternro & _
         " ORDER BY altfec DESC "

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   If Not IsNull(rsConsult!bajfec) Then
      empFecBaja = rsConsult!bajfec
   Else
      empFecBaja = ""
   End If
Else
   empFecBaja = ""
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del reporta a
'------------------------------------------------------------------
If IsNull(empreporta) Then
    reportaa = ""
Else
    StrSql = " SELECT empleg,terape,terape2,ternom,ternom2 "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " WHERE ternro= " & empreporta
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       reportaa = rsConsult!terape & " " & rsConsult!terape2
       reportaa = reportaa & ", " & rsConsult!ternom & " " & rsConsult!ternom2
    Else
       reportaa = ""
    End If

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
'Calculo la antiguedad
'------------------------------------------------------------------
Antiguedad = ""
Fecha_aux = CDate("01" & "/" & pliqmes & "/" & pliqanio)
Fecha_aux = DateAdd("m", 1, Fecha_aux)
Fecha_aux = DateAdd("d", -1, Fecha_aux)

If Fecha_aux < empFecAlta Then
    Antiguedad = "0 año/s 0 mes/ses"
Else
    Call bus_Antiguedad(Ternro, "REAL", Fecha_aux, dias_aux, meses_aux, anios_aux, diasHab_aux)
    Antiguedad = anios_aux & " año/s " & meses_aux & " mes/es"
End If


'------------------------------------------------------------------
'Busco el valor del cuil
'------------------------------------------------------------------
StrSql = " SELECT cuil.nrodoc "
StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Cuil = rsConsult!NroDoc
Else
'   Flog.writeline "Error al obtener los datos del cuil"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor de la direccion y localidad
'------------------------------------------------------------------

StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc"
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(proFecPago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(proFecPago) & ") AND tenro=1 AND his_estructura.ternro=" & Ternro
StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
       
OpenRecordset StrSql, rsConsult
Direccion = ""
If Not rsConsult.EOF Then
   Direccion = rsConsult!calle & " " & rsConsult!nro & ", " & rsConsult!locdesc
   Localidad = Direccion
Else
'   Flog.writeline "Error al obtener los datos de la localidad"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'buscar el valor del sueldo


StrSql = " SELECT almonto"
StrSql = StrSql & " From acu_liq"
StrSql = StrSql & " Where acunro = " & acunroSueldo
StrSql = StrSql & " AND cliqnro = " & cliqnro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Sueldo = rsConsult!almonto
Else
   Flog.writeline "Error al obtener los datos del sueldo"
   Sueldo = 0
   'GoTo MError
End If


'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 3 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Categoria = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la categoria"
'   GoTo MError
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
'Busco el valor del grupo de seguridad
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 7 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

grupoSeguridad = ""

If Not rsConsult.EOF Then
   grupoSeguridad = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del grupo de seguridad"
'   GoTo MError
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
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor de la obra social elegida
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 17 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

oSocial = ""

If Not rsConsult.EOF Then
   oSocial = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor de la gerencia
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 6 And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsConsult

Gerencia = ""

If Not rsConsult.EOF Then
   Gerencia = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la gerencia"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & Ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
StrSql = " SELECT * FROM ctabancaria WHERE ctabestado=-1 AND ternro = " & Ternro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then

  If rsConsult.RecordCount = 1 Then
     
     'La forma de pago la saca de ctabancaria, por que sino cuando es cheque no la encuentra
     StrSql = "SELECT ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc"
     StrSql = StrSql & " From ctabancaria"
     StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = ctabancaria.fpagnro"
     StrSql = StrSql & " LEFT JOIN banco ON banco.ternro =  ctabancaria.banco"
     StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
     StrSql = StrSql & " WHERE ctabestado=-1 AND ctabancaria.ternro=" & Ternro
  
  Else
     
     StrSql = " SELECT ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc"
     StrSql = StrSql & " From pago"
     StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro AND pago.pagorigen=" & cliqnro
     StrSql = StrSql & " LEFT JOIN banco     ON banco.ternro = pago.banternro"
     StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
  
  End If
  
  rsConsult.Close
       
  OpenRecordset StrSql, rsConsult

  If Not rsConsult.EOF Then
     If CBool(rsConsult!fpagbanc) Then
        FormaPago = rsConsult!fpagdescabr & "&nbsp;&nbsp;&nbsp;&nbsp;" & rsConsult!terrazsoc & "&nbsp;&nbsp;&nbsp;&nbsp;Cuenta:" & rsConsult!ctabnro
     Else
        FormaPago = rsConsult!fpagdescabr
     End If
  Else
     'Me fijo si la ctabancaria no tiene asociado un banco, o sea es en efectivo
     StrSql = "SELECT fpagdescabr "
     StrSql = StrSql & " From ctabancaria"
     StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = ctabancaria.fpagnro"
     StrSql = StrSql & " WHERE ctabestado=-1 AND ctabancaria.ternro=" & Ternro
     
     rsConsult.Close
     OpenRecordset StrSql, rsConsult
     
     If Not rsConsult.EOF Then
        FormaPago = rsConsult!fpagdescabr
     Else
        Flog.writeline "Error, no se encuentra la forma de pago del empleado"
     End If
     
  End If

Else
   Flog.writeline "Error, el empleado no tiene cuentas bancarias"
End If

rsConsult.Close

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    "(his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    "AND his_estructura.ternro = " & Ternro & _
    "AND his_estructura.tenro  = 10"
OpenRecordset StrSql, rs_estructura

EmpEstrnro = 0

If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    Exit Sub
Else
    EmpEstrnro = rs_estructura!Estrnro
    EmpNombre = rs_estructura!empnom
End If

'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc,piso,codigopostal From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    EmpDire = "   "
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro & " Piso " & rs_Domicilio!Piso & "<br>" & rs_Domicilio!codigopostal & " - " & rs_Domicilio!locdesc
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

'------------------------------------------------------------------
'Obtengo la configuracion del tipo de documento resolucion
'------------------------------------------------------------------
StrSql = " SELECT confval FROM confrep WHERE repnro = 60 AND confnrocol = 358 "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    tidnroRST = rsConsult!confval
    Flog.writeline "Codigo de tipo de documento Resolucion encontrado."
Else
    tidnroRST = 0
    Flog.writeline "Codigo de tipo de documento Resolucion No encontrado."
End If

'Consulta para buscar el numero de resolucion de la empresa
StrSql = "SELECT nrodoc FROM ter_doc " & _
         " Where ternro =" & rs_estructura!Ternro & " AND ter_doc.tidnro = " & tidnroRST
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No se encontró nro de resolucion de la empresa"
    'Exit Sub
    codResolucion = 0
Else
    codResolucion = rsConsult!NroDoc
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

'Consulta para buscar la firma del responsable de RRHH
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 15 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
    
    OpenRecordset StrSql, rs_firma
    
If rs_firma.EOF Then
    Flog.writeline "No se encontró la firma del responsable de RRHH"
    'Exit Sub
    RRHHFirma = ""
    RRHHFirmaAlto = 0
    RRHHFirmaAncho = 0
Else
    RRHHFirma = rs_firma!tipimdire & rs_firma!terimnombre
    RRHHFirmaAlto = rs_firma!tipimaltodef
    RRHHFirmaAncho = rs_firma!tipimanchodef
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
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden, auxchar1,auxchar2,auxchar3,auxchar4,auxchar5,auxchar6, auxdeci1, auxdeci2, auxchar7 )"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & Nombre & "'"
StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & Sueldo
StrSql = StrSql & ",'" & Mid(Categoria, 1, 20) & "'"
StrSql = StrSql & ",'" & Mid(CentroCosto, 1, 25) & "'"
StrSql = StrSql & ",'" & Mid(Localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & FormaPago & "'"
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Mid(empFecBaja, 1, 100) & "'"
StrSql = StrSql & ",'" & Mid(oSocial, 1, 100) & "'"
StrSql = StrSql & ",'" & Mid(Gerencia, 1, 100) & "'"
StrSql = StrSql & ",'" & Mid(Antiguedad, 1, 100) & "'"
StrSql = StrSql & ",'" & Mid(Sucursal, 1, 100) & "'"
StrSql = StrSql & ",'" & RRHHFirma & "'"
StrSql = StrSql & "," & RRHHFirmaAlto
StrSql = StrSql & "," & RRHHFirmaAncho
StrSql = StrSql & ",'" & codResolucion & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

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
' Se encarga de generar los datos para Recibo de PERSONAL
'--------------------------------------------------------------------
Sub generarDatosRecibo189(Pronro, Ternro, acunroSueldo, tituloReporte, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim Nombre
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim empFecAlta
Dim Sueldo
Dim proFecPago
Dim pliqhasta

Dim EmpEstrnro
Dim EmpNombre As String
Dim CodRUC As Integer
Dim CodIPS As Integer
Dim CodMJT As Integer
Dim CodEmpDoc As Integer
Dim DocEmp As String
Dim EmpMJT As String
Dim EmpIPS As String
Dim EmpRUC As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim Banco As String
Dim Cuenta As String
Dim Gerencia

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset

Dim empFecBaja
Dim causaBaja

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

'-----------------------------------------------------------------------------------------'
'  Buscar la configuracion del confrep, para los Tipo de Documentos RUC - IPS - MJT - DOC '
'-----------------------------------------------------------------------------------------'
Flog.writeline "Obtengo los datos del confrep para los Tipo de Documentos RUC - IPS - MJT - DOC ."
CodRUC = 0
CodIPS = 0
CodMJT = 0

'RUC
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 110 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No esta configurado el ConfRep para el Tipo de Documento RUC (col 110)."
Else
    CodRUC = rsConsult!confval
End If
rsConsult.Close

'IPS
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 398 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
   Flog.writeline "No esta configurado el ConfRep para el Tipo de Documento IPS(col 398)."
Else
   CodIPS = rsConsult!confval
End If
rsConsult.Close

'MJT
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 399 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No esta configurado el ConfRep para el Tipo de Documento MJT (col 399)."
Else
    CodMJT = rsConsult!confval
End If
rsConsult.Close

StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 60 AND confnrocol = 400 "
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "No esta configurado el ConfRep para el Tipo de Documento del Empleado (col 400)."
Else
    CodEmpDoc = rsConsult!confval
End If
rsConsult.Close


'------------------------------------------------------------------
'Obtengo el nro de cabezera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini, proceso.profecplan FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    cliqnro = rsConsult!cliqnro
    proFecPago = rsConsult!proFecplan
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
'Busco la Fecha de Alta del empleado. Corresponde a la fecha desde de la fase +
'antigua con real en true.
'------------------------------------------------------------------
StrSql = " SELECT altfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE real = -1 AND fases.empleado=" & Ternro & " order by altfec ASC"
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   empFecAlta = rsConsult!altfec
   Flog.writeline "    Fecha de Alta del Empleado Fase + antigua OK"
Else
   Flog.writeline "Error al obtener la Fecha de Alta del Empleado Fase + antigua"
   empFecAlta = ""
   'GoTo MError
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
'Busco el banco y la cuenta del empleado
'------------------------------------------------------------------
StrSql = " SELECT banco.bandesc, ctabancaria.ctabnro, ctabancaria.ctabcbu "
StrSql = StrSql & " From ctabancaria "
StrSql = StrSql & " INNER JOIN banco ON ctabancaria.banco = banco.ternro "
StrSql = StrSql & " WHERE ctabancaria.ternro= " & Ternro
StrSql = StrSql & " AND ctabancaria.ctabestado = -1"

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   If Not EsNulo(rsConsult!ctabcbu) Then
       Banco = rsConsult!Bandesc
       Cuenta = rsConsult!ctabcbu
   Else
       If Not EsNulo(rsConsult!ctabnro) Then
            Banco = rsConsult!Bandesc
            Cuenta = rsConsult!ctabnro
       Else
            Banco = ""
            Cuenta = ""
       End If
   End If
Flog.writeline "    los datos de Cuenta + Banco (OK)"
Else
    Banco = ""
    Cuenta = ""
    Flog.writeline "Error al obtener los datos de Cuenta + Banco"
End If

'------------------------------------------------------------------
'Busco el valor del GERENCIA
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 6 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult

Gerencia = ""
If Not rsConsult.EOF Then
   Gerencia = rsConsult!estrdabr
Else
   Flog.writeline "Error al obtener los datos de Gerencia"
   GoTo MError
End If


'------------------------------------------------------------------
' Busco VHora definido en la columna 44 del confrep
'Se cambio -- se usa col 44 para el Sueldo.
'------------------------------------------------------------------
Sueldo = 0

If esSueldoRecibo Then
    StrSql = " SELECT detliq.dlimonto valor "
    StrSql = StrSql & " FROM detliq "
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And detliq.concnro = " & SueldoRecibo
Else
    StrSql = " SELECT almonto valor"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & SueldoRecibo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
End If
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Sueldo = rsConsult!Valor
Else
    Flog.writeline " No se encontraron datos de de SUELDO/JORNAL. (col44)"
    Sueldo = 0
    'GoTo MError
End If

' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------
EmpEstrnro = 0
EmpNombre = " "

StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = 10"
OpenRecordset StrSql, rs_estructura
If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    GoTo MError
Else
    EmpNombre = rs_estructura!empnom
    EmpEstrnro = rs_estructura!Estrnro
End If

' -------------------------------------------------------------------------
'Consulta para buscar el logo de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_logo
If rs_logo.EOF Then
    Flog.writeline "No se encontró el Logo de la Empresa"
    EmpLogo = ""
    EmpLogoAlto = 0
    EmpLogoAncho = 0
Else
    EmpLogo = rs_logo!tipimdire & rs_logo!terimnombre
    EmpLogoAlto = rs_logo!tipimaltodef
    EmpLogoAncho = rs_logo!tipimanchodef
End If

' -------------------------------------------------------------------------
'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If

' -------------------------------------------------------------------------
'Consulta para obtener el MJT de la empresa
Flog.writeline "Buscando el MTJ tipo de doc " & CodMJT
StrSql = " SELECT *"
StrSql = StrSql & " From ter_doc"
StrSql = StrSql & " Where ter_doc.ternro = " & rs_estructura!Ternro & " And ter_doc.tidnro = " & CodMJT
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el MJT"
    EmpMJT = "  "
Else
    Flog.writeline "MJT = " & rs_cuit!NroDoc
    EmpMJT = IIf(EsNulo(rs_cuit!NroDoc), " ", rs_cuit!NroDoc)
End If

' -------------------------------------------------------------------------
'Consulta para obtener el IPS de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero "
StrSql = StrSql & " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro AND cuit.tidnro = " & CodIPS & ")"
StrSql = StrSql & " WHERE tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el IPS de la Empresa"
    EmpIPS = "  "
Else
    EmpIPS = rs_cuit!NroDoc
End If

' -------------------------------------------------------------------------
'Consulta para obtener el RUC de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero "
StrSql = StrSql & " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro AND cuit.tidnro = " & CodRUC & ")"
StrSql = StrSql & " WHERE tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el RUC de la Empresa"
    EmpRUC = "  "
Else
    EmpRUC = rs_cuit!NroDoc
End If

' -------------------------------------------------------------------------
'Consulta para obtener el Documento del Empleado
Flog.writeline "Buscando el CodEmpDoc tipo de doc " & CodEmpDoc
StrSql = " SELECT *"
StrSql = StrSql & " From ter_doc"
StrSql = StrSql & " Where ter_doc.ternro = " & Ternro & " And ter_doc.tidnro = " & CodEmpDoc
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CodEmpDoc"
    DocEmp = "  "
Else
    Flog.writeline "DocEmp = " & rs_cuit!NroDoc
    DocEmp = IIf(EsNulo(rs_cuit!NroDoc), " ", rs_cuit!NroDoc)
End If

' -------------------------------------------------------------------------

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
End If
rsConsult.Close
'-----------------------------------------------------------------------------------------


'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------
StrSql = " INSERT INTO rep_recibo "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,pliqdepant,"
StrSql = StrSql & " pliqfecdep,pliqbco,empfecalta,"
StrSql = StrSql & " sueldo,"
StrSql = StrSql & " profecpago,empnombre,emplogo,emplogoalto,emplogoancho,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,prodesc,descripcion, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden,"
StrSql = StrSql & " auxchar1, auxchar2, auxchar3,"
StrSql = StrSql & " auxchar4, auxchar5, auxchar6,auxchar7,auxchar8"
StrSql = StrSql & ")"
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
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & "," & numberForSQL(Sueldo)
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden

StrSql = StrSql & ",'" & EmpMJT & "'"
StrSql = StrSql & ",'" & EmpIPS & "'"
StrSql = StrSql & ",'" & EmpRUC & "'"

StrSql = StrSql & ",'" & Banco & "'"
StrSql = StrSql & ",'" & Cuenta & "'"
StrSql = StrSql & ",'" & Gerencia & "'"
StrSql = StrSql & ",'" & DocEmp & "'"
StrSql = StrSql & ",'" & fecEstr & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------
objConn.Execute StrSql, , adExecuteNoRecords


'Flog.Writeline "======================================================================================================================================="
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

'Flog.writeline " Se sale del procedimiento de generarDatosRecibo50. "

Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub

