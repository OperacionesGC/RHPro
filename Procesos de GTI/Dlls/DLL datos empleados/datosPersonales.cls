VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "datosPersonales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private empLegajo As Long
Private empNombre As String
Private empNombre2 As String
Private empApellido As String
Private empApellido2 As String
Private empSexo As String
Private empOficina As String
Private empReporta As String
Private empFNacimiento As String
Private empFechaBaja As String
Private empFechaAlta As String
Private empEstadoCivil As String
Private empPais As String
Private empPaisDesc As String
Private empCausaBaja As String
Private empFechaReingreso As String
Private empCuil As String
Private empLugarNac As String
Private empTitulo As String
Private empnrodoc As String
Private empSalarioFamiliar As String
Private empFechaEstadoCivil As String
Private empCtaBancaria As String
Private empFormaPago As String
Private empEmail As String
Private empNacionalidad As String
Private empInstitucion  As String
Private empTipoDoc As String
Private empEstado As String
Private empMoneda As String
Private empLicencia As String
Private empNivel As String
Private empTelInterno As String
Private empDocPais As Integer
Dim rs_datos As New ADODB.Recordset
Public Property Get obtenerLegajo() As Long
    obtenerLegajo = empLegajo
End Property
Public Property Get obtenerNombreApellido(ByVal campo As String) As String
    Select Case campo
        Case "nombre":
            obtenerNombreApellido = empNombre
        Case "apellido":
            obtenerNombreApellido = empApellido
        Case "nombre2":
            obtenerNombreApellido = empNombre2
        Case "apellido2":
            obtenerNombreApellido = empApellido2
    End Select
End Property
Public Property Get obtenerSexo() As String
    obtenerSexo = empSexo
End Property
Public Property Get obtenerEstructura() As String
    obtenerEstructura = empOficina
End Property
Public Property Get obtenerReporta() As String
    obtenerReporta = empReporta
End Property
Public Property Get obtenerFNacimiento() As String
    obtenerFNacimiento = empFNacimiento
End Property
Public Property Get obtenerFechaBaja() As String
    obtenerFechaBaja = empFechaBaja
End Property
Public Property Get obtenerFechaAlta() As String
    obtenerFechaAlta = empFechaAlta
End Property
Public Property Get obtenerEmail() As String
    obtenerEmail = empEmail
End Property
Public Property Get obtenerEstadoCivil() As String
    obtenerEstadoCivil = empEstadoCivil
End Property
Public Property Get obtenerPaisDelEmpleado() As String
    obtenerPaisDelEmpleado = empPais
End Property
Public Property Get obtenerDescPaisDelEmpleado() As String
    obtenerDescPaisDelEmpleado = empPaisDesc
End Property
Public Property Get obtenerCausaBajaEmpleado() As String
    obtenerCausaBajaEmpleado = empCausaBaja
End Property
Public Property Get obtenerFechaReingreso() As String
    obtenerFechaReingreso = empFechaReingreso
End Property
Public Property Get obtenerCUIL() As String
    obtenerCUIL = empCuil
End Property
Public Property Get obtenerLugarNac() As String
    obtenerLugarNac = empLugarNac
End Property
Public Property Get obtenerTitulo() As String
    obtenerTitulo = empTitulo
End Property
Public Property Get obtenerNroDoc() As String
    obtenerNroDoc = empnrodoc
End Property
Public Property Get obtenerSalarioFamiliar() As String
    obtenerSalarioFamiliar = empSalarioFamiliar
End Property
Public Property Get obtenerCtaBancaria() As String
    obtenerCtaBancaria = empCtaBancaria
End Property
Public Property Get obtenerTipoCta() As String
    obtenerTipoCta = empFormaPago
End Property
Public Property Get obtenerInstitucion() As String
    obtenerInstitucion = empInstitucion
End Property
Public Property Get obtenerTelInterno() As String
    obtenerTelInterno = empTelInterno
End Property
Public Property Get obtenerNivel() As String
    obtenerNivel = empNivel
End Property
Public Property Get obtenerTipoDoc() As String
    obtenerTipoDoc = empTipoDoc
End Property
Public Property Get obtenerNacionalidad() As String
    obtenerNacionalidad = empNacionalidad
End Property
Public Property Get obtenerEstado() As String
    obtenerEstado = empEstado
'    If empEstado = -1 Then
'        obtenerEstado = "A"
'    Else
'        StrSql = " SELECT codexterno FROM mapeo_sap "
'        StrSql = StrSql & " WHERE tablaref='caunro'"
'        StrSql = StrSql & " AND codInterno = " & empEstado
'        StrSql = StrSql & " AND upper(left(infotipo,3))='PY_'"
'        Flog.writeline "query estado: " & StrSql
'        OpenRecordset StrSql, rs_datos
'        If Not rs_datos.EOF Then
'            If Not EsNulo(rs_datos("codexterno")) Then
'                obtenerEstado = rs_datos("codexterno")
'            Else
'                obtenerEstado = 0
'            End If
'        Else
'            obtenerEstado = 0
'        End If
'        rs_datos.Close
'    End If
    
End Property
Public Property Get obtenerMoneda() As String
   obtenerMoneda = empMoneda
End Property
Public Property Get obtenerFechaEstadoCivil() As String
    obtenerFechaEstadoCivil = empFechaEstadoCivil
End Property
Public Property Get obtenerLicencia() As String
    obtenerLicencia = empLicencia
End Property
Public Property Get obtenerDocPais() As String
    obtenerDocPais = empDocPais
End Property
Public Sub buscarDatosPersonales(ByVal ternro As Integer)
On Error GoTo CE
    StrSql = " SELECT empleado.empleg, tercero.terape, tercero.terape2, tercero.ternom, tercero.ternom2, tercero.tersex, empleado.empreporta, R.empleg reporta"
    StrSql = StrSql & ", terfecnac, terestciv, paisdesc, ter_doc.nrodoc cuil, lugar_nac.lugardesc, tercero.paisnro, tercero.terfecestciv, nacionalcodext, teremail "
    StrSql = StrSql & ", estcivcodext, empleado.empinterno, tercero.docpais"
    StrSql = StrSql & " FROM tercero "
    StrSql = StrSql & " LEFT JOIN empleado ON empleado.ternro = tercero.ternro"
    StrSql = StrSql & " LEFT JOIN empleado R ON r.ternro = empleado.empreporta "
    StrSql = StrSql & " LEFT JOIN pais ON pais.paisnro = tercero.docpais "
    StrSql = StrSql & " LEFT JOIN ter_doc ON ter_doc.ternro = tercero.ternro AND ter_doc.tidnro = 10 "
    StrSql = StrSql & " LEFT JOIN lugar_nac ON lugar_nac.lugarnro = tercero.lugarnro "
    StrSql = StrSql & " LEFT JOIN nacionalidad n ON n.nacionalnro = tercero.nacionalnro"
    StrSql = StrSql & " LEFT JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro "
    StrSql = StrSql & " WHERE tercero.ternro=" & ternro
    Flog.writeline "Query datos personales: " & StrSql
    OpenRecordset StrSql, rs_datos
    
    If Not rs_datos.EOF Then
        'legajo
        If Not EsNulo(rs_datos("empleg")) Then
            empLegajo = rs_datos("empleg")
        Else
            empLegajo = 0
        End If
        
        'Segundo apellido
        If Not EsNulo(rs_datos("terape2")) Then
            empApellido2 = rs_datos("terape2")
        End If
        
        'apellido
        If Not EsNulo(rs_datos("terape")) Then
            empApellido = rs_datos("terape")
        End If
        
        'segundo Nombre
        If Not EsNulo(rs_datos("ternom2")) Then
            empNombre2 = rs_datos("ternom2")
        End If
        
        'Nombre
        If Not EsNulo(rs_datos("ternom")) Then
            empNombre = rs_datos("ternom")
        End If
        
        'sexo
        If rs_datos("tersex") = -1 Then
            empSexo = "M"
        Else
            empSexo = "F"
        End If
        
        If Not EsNulo(rs_datos("reporta")) Then
            empReporta = rs_datos("reporta")
        Else
            empReporta = ""
        End If
        
        If Not EsNulo(rs_datos("terfecnac")) Then
            empFNacimiento = rs_datos("terfecnac")
        Else
            empFNacimiento = ""
        End If
        
        empEstadoCivil = IIf(EsNulo(rs_datos("estcivcodext")), "", rs_datos("estcivcodext"))
        
        'fecha del estado civil
        empFechaEstadoCivil = IIf(EsNulo(rs_datos("terfecestciv")), "", rs_datos("terfecestciv"))
        
        If Not EsNulo(rs_datos("paisnro")) Then
            empPais = rs_datos("paisnro")
        Else
            empPais = ""
        End If
        
        If Not EsNulo(rs_datos("paisdesc")) Then
            empPaisDesc = rs_datos("paisdesc")
        Else
            empPaisDesc = ""
        End If
        
        'CUIL DEL EMPLEADO
        If Not EsNulo(rs_datos("cuil")) Then
            empCuil = rs_datos("cuil")
        Else
            empCuil = ""
        End If
        
        'lugar de nacimiento
        If Not EsNulo(rs_datos("lugardesc")) Then
            empLugarNac = rs_datos("lugardesc")
        Else
            empLugarNac = ""
        End If
        
        'Email
        If Not EsNulo(rs_datos("teremail")) Then
            empEmail = rs_datos("teremail")
        Else
            empEmail = ""
        End If
        
        'Nacionalidad
        If Not EsNulo(rs_datos("nacionalcodext")) Then
            empNacionalidad = rs_datos("nacionalcodext")
        Else
            empNacionalidad = 0
        End If
        
        'telefono interno
        If Not EsNulo(rs_datos("empinterno")) Then
            empTelInterno = rs_datos("empinterno")
        Else
            empTelInterno = 0
        End If
        
        'docpais
        If Not EsNulo(rs_datos("docpais")) Then
            empDocPais = rs_datos("docpais")
        Else
            empDocPais = 0
        End If
        
    Else
        empLegajo = 0
        empNombre = ""
        empEstadoCivil = ""
        empCuil = ""
        empLugarNac = ""
        empTelInterno = 0
        empDocPais = 0
    End If
    rs_datos.Close
    Exit Sub
CE:
    Flog.writeline "Se produjo un error en la clase datosPersonales"
    Flog.writeline "descripcion del error: " & Err.Description
    Exit Sub
End Sub
Public Sub buscarEstructuras(ByVal ternro As Integer, ByVal tenro As Integer, ByVal Fecha As String, ByVal campo As String, Optional ByVal mapeo As Boolean)
On Error GoTo errorEst
        
    StrSql = " SELECT e." & campo & " valor FROM his_estructura"
    StrSql = StrSql & " INNER JOIN estructura E ON e.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ((htetdesde <=" & ConvFecha(Fecha) & ") And ((htethasta >=" & ConvFecha(Fecha) & ") Or (htethasta is Null)))"
    StrSql = StrSql & " AND his_estructura.tenro=" & tenro
    StrSql = StrSql & " AND his_estructura.ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        empOficina = rs_datos("valor")
    Else
        empOficina = 0
    End If
    rs_datos.Close
    
'    'VER MAPEO DE ESTRUCTURA
'    'busco en la tabla de mapeo
'    StrSql = " SELECT codexterno FROM mapeo_sap "
'    StrSql = StrSql & " WHERE tablaref='estructura'"
'    StrSql = StrSql & " AND codInterno = '" & tenro & "@" & empOficina & "'"
'    StrSql = StrSql & " AND upper(left(infotipo,3))='PY_'"
'    OpenRecordset StrSql, rs_datos
'    If Not rs_datos.EOF Then
'        If Not EsNulo(rs_datos("codexterno")) Then
'            empOficina = rs_datos("codexterno")
'        Else
'            empOficina = 0
'        End If
'    Else
'        empOficina = 0
'    End If
'    rs_datos.Close
Exit Sub
errorEst:
    Flog.writeline "Error en estructuras"
    Flog.writeline "ERROR: " & Err.Description
    Exit Sub
End Sub

Public Sub buscarFechaBaja(ByVal ternro As Integer, ByVal Fecha As String)
    StrSql = " SELECT bajfec, fases.caunro, causa.caucod FROM FASES "
    StrSql = StrSql & " LEFT JOIN causa ON  causa.caunro = fases.caunro "
    StrSql = StrSql & " WHERE (altfec<=" & ConvFecha(Fecha) & ") AND (bajfec >= " & ConvFecha(Fecha) & " OR bajfec is null )"
    StrSql = StrSql & " AND empleado=" & ternro
    StrSql = StrSql & " ORDER BY altfec DESC "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If Not EsNulo(rs_datos("bajfec")) Then
            empFechaBaja = rs_datos("bajfec")
            empCausaBaja = rs_datos("caucod")
        Else
            empFechaBaja = ""
            empCausaBaja = ""
        End If
    Else
        empFechaBaja = ""
        empCausaBaja = ""
    End If
    rs_datos.Close
    

'    'busco en la tabla de mapeo
'    StrSql = " SELECT codexterno FROM mapeo_sap "
'    StrSql = StrSql & " WHERE tablaref='causa'"
'    StrSql = StrSql & " AND codInterno = " & empCausaBaja
'    StrSql = StrSql & " AND upper(left(infotipo,3))='PY_'"
'    OpenRecordset StrSql, rs_datos
'    If Not rs_datos.EOF Then
'        If Not EsNulo(rs_datos("codexterno")) Then
'            empCausaBaja = rs_datos("codexterno")
'        Else
'            empCausaBaja = ""
'        End If
'    Else
'        empCausaBaja = ""
'    End If
'    rs_datos.Close
    
End Sub
Public Sub buscarFechaAlta(ByVal ternro As Integer, ByVal Fecha As String)
    StrSql = " SELECT altfec FROM FASES "
    StrSql = StrSql & " WHERE fasrecofec=-1"
    StrSql = StrSql & " AND empleado=" & ternro
    StrSql = StrSql & " ORDER BY altfec DESC "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If Not EsNulo(rs_datos("altfec")) Then
            empFechaAlta = rs_datos("altfec")
        Else
            empFechaAlta = ""
        End If
    Else
        empFechaAlta = ""
    End If
    rs_datos.Close
End Sub

Public Sub buscarFechaReingreso(ByVal ternro As Integer, ByVal Fecha As String)
    StrSql = " SELECT altfec FROM FASES "
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " empleado=" & ternro
    StrSql = StrSql & " ORDER BY altfec DESC "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If Not EsNulo(rs_datos("altfec")) Then
            empFechaReingreso = rs_datos("altfec")
        Else
            empFechaReingreso = ""
        End If
    Else
        empFechaReingreso = ""
    End If
    rs_datos.Close
End Sub

Public Sub buscarTitulo(ByVal ternro As Integer)
Dim titnro As Integer

    StrSql = " SELECT titulo.titnro, institucion.instdes, cap_estformal.nivnro, nivest.nivcodext FROM cap_estformal "
    StrSql = StrSql & " INNER JOIN titulo ON titulo.titnro = cap_estformal.titnro "
    StrSql = StrSql & " INNER JOIN institucion ON institucion.instnro = titulo.instnro "
    StrSql = StrSql & " INNER JOIN nivest ON nivest.nivnro = cap_estformal.nivnro "
    StrSql = StrSql & " WHERE ternro =" & ternro & " AND capcomp = -1 "
    StrSql = StrSql & " ORDER BY cap_estformal.nivnro DESC , capfechas DESC "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If Not EsNulo(rs_datos("titnro")) Then
            empTitulo = rs_datos("titnro")
        Else
            empTitulo = 0
        End If
        
        If Not EsNulo(rs_datos("instdes")) Then
            empInstitucion = rs_datos("instdes")
        Else
            empInstitucion = ""
        End If
        
        If Not EsNulo(rs_datos("nivcodext")) Then
            empNivel = rs_datos("nivcodext")
        Else
            empNivel = "N"
        End If
        
    Else
        empTitulo = 0
        empInstitucion = ""
        empNivel = "N"
    End If
    rs_datos.Close
    
'    'busco en la tabla de mapeo
'    StrSql = " SELECT codexterno FROM mapeo_sap "
'    StrSql = StrSql & " WHERE upper(tablaref)='TITULO'"
'    StrSql = StrSql & " AND codInterno = " & titnro
'    StrSql = StrSql & " AND upper(left(infotipo,3))='PY_'"
'    OpenRecordset StrSql, rs_datos
'    If Not rs_datos.EOF Then
'        If Not EsNulo(rs_datos("codexterno")) Then
'            empTitulo = rs_datos("codexterno")
'        Else
'            empTitulo = 0
'        End If
'    Else
'        empTitulo = 0
'    End If
'    rs_datos.Close
End Sub

Public Sub buscarNroDoc(ByVal ternro As Integer, ByVal tipoDoc As Integer)
    StrSql = " SELECT nroDoc, t.tidcod, tidsigla  FROM ter_doc "
    If tipoDoc = 0 Then
        StrSql = StrSql & " INNER JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro "
        StrSql = StrSql & " INNER JOIN tipodocu_pais t ON t.tidnro = ter_doc.tidnro AND t.tidcod <= 5"
    Else
        StrSql = StrSql & " INNER JOIN tipodocu on tipodocu.tidnro = ter_doc.tidnro "
        StrSql = StrSql & " INNER JOIN tipodocu_pais t ON t.tidnro = ter_doc.tidnro AND t.tidcod = " & tipoDoc
    End If
    StrSql = StrSql & " WHERE ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If Not EsNulo(rs_datos("nrodoc")) Then
            empnrodoc = rs_datos("nrodoc")
            empTipoDoc = rs_datos("tidsigla")
        Else
            empnrodoc = ""
            empTipoDoc = 0
        End If
    Else
        empnrodoc = ""
        empTipoDoc = 0
    End If
    rs_datos.Close
End Sub

Public Sub buscarSalarioFamiliar(ByVal ternro As Integer)
    StrSql = " SELECT * FROM familiar "
    StrSql = StrSql & " WHERE empleado=" & ternro
    StrSql = StrSql & " AND famsalario=-1 "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        'si no es fin de archivo ==> cobra salario familiar
        empSalarioFamiliar = "S"
    Else
        empSalarioFamiliar = "N"
    End If
    rs_datos.Close
End Sub

Public Sub buscarCtaBancaria(ByVal ternro As Integer)
    StrSql = "SELECT ctabancaria.ctabnro, ctabancaria.fpagnro FROM ctabancaria "
    StrSql = StrSql & " INNER JOIN banco ON banco.ternro = ctabancaria.banco "
    StrSql = StrSql & " WHERE ctabancaria.ternro =" & ternro & " AND ctabestado = -1"
    'StrSql = StrSql & " AND banco.estrnro =" & banco
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If Not EsNulo(rs_datos("ctabnro")) Then
            empCtaBancaria = rs_datos("ctabnro")
            empFormaPago = rs_datos("fpagnro")
        Else
            empCtaBancaria = ""
            empFormaPago = 0
        End If
    Else
        empCtaBancaria = ""
        empFormaPago = 0
    End If
    rs_datos.Close
    
    'busco la moneda con la forma de pago
    StrSql = "SELECT moneda.monnro , moneda.mondesabr, moneda.codext "
    StrSql = StrSql & " From formapago "
    StrSql = StrSql & " INNER JOIN moneda ON moneda.monnro = formapago.monnro and formapago.fpagnro=" & empFormaPago
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        empMoneda = IIf(EsNulo(rs_datos("codext")), 0, rs_datos("codext"))
    Else
        empMoneda = 1
    End If
    rs_datos.Close
    
'    'busco en la tabla de mapeo
'    StrSql = " SELECT codexterno FROM mapeo_sap "
'    StrSql = StrSql & " WHERE tablaref='formapago'"
'    StrSql = StrSql & " AND codInterno = " & empFormaPago
'    StrSql = StrSql & " AND upper(left(infotipo,3))='PY_'"
'    OpenRecordset StrSql, rs_datos
'    If Not rs_datos.EOF Then
'        If Not EsNulo(rs_datos("codexterno")) Then
'            empFormaPago = rs_datos("codexterno")
'        Else
'            empFormaPago = 0
'        End If
'    Else
'        empFormaPago = 0
'    End If
'    rs_datos.Close
End Sub

Public Sub buscarEstado(ByVal ternro As Integer)
    StrSql = " SELECT altfec, bajfec, tipnro FROM FASES "
    StrSql = StrSql & " INNER JOIN ter_tip ON ter_tip.ternro = fases.empleado "
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " fases.empleado=" & ternro
    StrSql = StrSql & " AND ((altfec<= " & ConvFecha(Date)
    StrSql = StrSql & " )AND (bajfec>= " & ConvFecha(Date) & " OR bajfec is null))"
    StrSql = StrSql & " ORDER BY bajfec DESC "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        If rs_datos!tipnro = "2" Then
            empEstado = "P"
        Else
            If rs_datos!tipnro = "1" Then
                empEstado = "A" 'esta activo y es empleado
            Else
                empEstado = ""
            End If
        End If
    Else
        'a la fecha esta de baja
        StrSql = " SELECT altfec, bajfec,caunro FROM FASES "
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " empleado=" & ternro
        StrSql = StrSql & " ORDER BY bajfec DESC "
        OpenRecordset StrSql, rs_datos
        If Not rs_datos.EOF Then
            'empEstado = IIf(EsNulo(rs_datos("caunro")), 0, rs_datos("caunro"))
            empEstado = "E" 'esta de baja el empleado
        Else
            empEstado = ""
        End If
    End If
    rs_datos.Close
End Sub
Public Sub buscarLicencia(ByVal ternro As Integer, ByVal tipoLic)
    StrSql = "SELECT * FROM emp_lic "
    StrSql = StrSql & " WHERE empleado=" & ternro
    StrSql = StrSql & " AND ((elfechadesde<= " & ConvFecha(Date)
    StrSql = StrSql & " )AND (elfechahasta>= " & ConvFecha(Date) & " OR elfechahasta is null))"
    StrSql = StrSql & " AND tdnro=" & tipoLic
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        empLicencia = rs_datos("emp_licnro")
    Else
        empLicencia = ""
    End If
    rs_datos.Close
End Sub



