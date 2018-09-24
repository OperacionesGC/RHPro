VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscarTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Dim P_Asignacion As Boolean
'Dim Nro_fpgo As Long
'Dim Fecha_Inicio As Date
'Dim p_turcomp As Boolean
Dim PFecha As Date

Private Type TTurno
    tiene_turno As Boolean
    Numero As Long
    Tipo As Integer
    Nombre As String
End Type

Private Type TJustif
    Tiene_Justif As Boolean
    justif_turno As Boolean
    Numero As Long
End Type

Private Type Templ
    Legajo As Long
    Grupo As Long
    NombreGrupo As String
End Type
Dim Empl As Templ
Dim Tur As TTurno
Dim Justif As TJustif

Dim NombreFPago As String

Property Set Conexion(ByRef cn As ADODB.Connection)
    Set objConn = cn
End Property
Property Set ConexionTraza(ByRef cn As ADODB.Connection)
    Set CnTraza = cn
End Property

Property Let Grupo_Nombre(N As String)
    Empl.NombreGrupo = N
End Property
Property Get Grupo_Nombre() As String
    Grupo_Nombre = Empl.NombreGrupo
End Property
Property Let Tiene_Justif(TieneJustif As Boolean)
    Justif.Tiene_Justif = TieneJustif
End Property
Property Let justif_turno(jus_turno As Boolean)
    Justif.justif_turno = jus_turno
End Property
Property Get justif_turno() As Boolean
    justif_turno = Justif.justif_turno
End Property
Property Get Tiene_Justif() As Boolean
    Tiene_Justif = Justif.Tiene_Justif
End Property
Property Let Justif_Numero(nro As Long)
    Justif.Numero = nro
End Property
Property Get Justif_Numero() As Long
    Justif_Numero = Justif.Numero
End Property
Property Let Turno_Nombre(N As String)
    Tur.Nombre = N
End Property
Property Get Turno_Nombre() As String
    Turno_Nombre = Tur.Nombre
End Property
Property Let FPago_Nombre(N As String)
    NombreFPago = N
End Property
Property Get FPago_Nombre() As String
    FPago_Nombre = NombreFPago
End Property

Property Let FechaInicio(Fecha As Date)
    Fecha_Inicio = Fecha
End Property
Property Get FechaInicio() As Date
    FechaInicio = Fecha_Inicio
End Property

Property Let Numero_FPago(nro As Long)
    Nro_fpgo = nro
End Property
Property Get Numero_FPago() As Long
    Numero_FPago = Nro_fpgo
End Property

Property Let Turno_Numero(nro As Long)
    Tur.Numero = nro
End Property
Property Get Turno_Numero() As Long
    Turno_Numero = Tur.Numero
End Property

Property Let tiene_turno(tiene As Boolean)
    Tur.tiene_turno = tiene
End Property
Property Get tiene_turno() As Boolean
    tiene_turno = Tur.tiene_turno
End Property

Property Let Turno_Tipo(Tipo As Integer)
    Tur.Tipo = Tipo
End Property
Property Get Turno_Tipo() As Integer
    Turno_Tipo = Tur.Tipo
End Property

Property Let Empleado_Legajo(nro As Long)
    Empl.Legajo = nro
End Property
Property Get Empleado_Legajo() As Long
    Empleado_Legajo = Empl.Legajo
End Property
Property Let Empleado_Grupo(nro As Long)
    Empl.Grupo = nro
End Property
Property Get Empleado_Grupo() As Long
    Empleado_Grupo = Empl.Grupo
End Property

Property Let Compensa_Turno(compensa As Boolean)
    p_turcomp = compensa
End Property
Property Get Compensa_Turno() As Boolean
    Compensa_Turno = p_turcomp
End Property

Property Let Tiene_PAsignacion(tiene As Boolean)
    P_Asignacion = tiene
End Property

Property Get Tiene_PAsignacion() As Boolean
    Tiene_PAsignacion = P_Asignacion
End Property




'Private Sub GeneraTraza(Ternro As Long, Fecha As Date, Desc As String, Optional valor As String = "?")
'
'    StrSql = "INSERT INTO gti_traza(ternro,fecproc,descripcion,valor) VALUES (" & _
'             Ternro & "," & ConvFecha(Fecha) & ",'" & Desc & "','" & valor & "')"
'
'    CnTraza.Execute StrSql, , adExecuteNoRecords
'
'End Sub


Public Sub Buscar_Turno(Fecha As Date, lngTernro As Long, depurar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el turno del empleado en la fecha.
' Autor      :
' Fecha      :
' Ultima Mod.: FGZ - 07/07/2008
' Descripcion: Se agregaron los chequeosa de circuito de firmas
' ---------------------------------------------------------------------------------------------
Dim objRsTurnoFPago As New ADODB.Recordset
Dim objRsReldTur As New ADODB.Recordset
Dim objRsTurnoFPGrupo As New ADODB.Recordset
Dim objLicAut As New ADODB.Recordset
'Dim Parte_FT As Boolean
Dim rs_FT As New ADODB.Recordset
Dim rs_Firma As New ADODB.Recordset
Dim Firmado As Boolean
Dim Nov_Firmada As Boolean
Dim Encontro_Jus As Boolean


Dim rs As New ADODB.Recordset
Dim rs_Tur As New ADODB.Recordset

    PFecha = Fecha
    Justif.Tiene_Justif = False
    Justif.justif_turno = False
    Tur.tiene_turno = False
    P_Asignacion = False
    Tur.Nombre = ""
    Tur.Numero = 0
    Tur.tiene_turno = False
    Tur.Tipo = 0
    Justif.justif_turno = False
    Justif.Numero = 0
    Justif.Tiene_Justif = False
    
    StrSql = ""
    Select Case TipoBD
    Case 4:
        StrSql = "SELECT * FROM ("
    End Select
    StrSql = StrSql & "(SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN emp_lic ON gti_justificacion.juscodext = emp_lic.emp_licnro "
    StrSql = StrSql & " WHERE (ternro = " & lngTernro & ") "
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND emp_lic.licestnro = 2"
    StrSql = StrSql & " AND jussigla = 'LIC'"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
    StrSql = StrSql & " )UNION ("
    StrSql = StrSql & " SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN gti_novedad ON gti_justificacion.juscodext = gti_novedad.gnovnro "
    StrSql = StrSql & " WHERE (Ternro = " & lngTernro & ")"
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND jussigla = 'NOV'"
    StrSql = StrSql & " AND jussigla <> 'ALM'"
    StrSql = StrSql & " AND (gti_novedad.ft = 0 OR (gti_novedad.ft = -1 AND gti_novedad.ftap = -1))"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " )UNION ("
    StrSql = StrSql & " SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " WHERE (Ternro = " & lngTernro & ")"
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND jussigla <> 'LIC'"
    StrSql = StrSql & " AND jussigla <> 'NOV'"
    'FGZ - 14/8/2008 - le agregué esta linea por las justificaciones automaticas de la Politica 400
    StrSql = StrSql & " AND jussigla <> 'ALM'"
    'FGZ - 14/8/2008 - le agregué esta linea por las justificaciones automaticas de la Politica 400
    StrSql = StrSql & ")"
    Select Case TipoBD
    Case 4:
        StrSql = StrSql & ")"
    End Select
    StrSql = StrSql & " ORDER BY juseltipo "
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        'Esto no quiere decir que no hay justificaciones sino que no hay o no hay aprobadas
    Else
        'FGZ - 10/06/2011 -------------------------------------------------------
        ' Se agregó control de firmas para Novedades horarias (que no sean automaticas - pol 400)
        Encontro_Jus = False
        Do While Not objRs.EOF And Not Encontro_Jus
            If objRs!jussigla = "NOV" Then
                'Verificar si esta en el NIVEL FINAL DE FIRMA ACTIVO para Novedades Horarias de GTI
                If Firma_Novedades Then
                    StrSql = "SELECT * FROM cysfirmas "
                    StrSql = StrSql & " WHERE cysfirfin = -1"
                    StrSql = StrSql & " AND cysfircodext = '" & objRs!juscodext & "' "
                    StrSql = StrSql & " AND cystipnro = 7"
                    OpenRecordset StrSql, rs
                    If rs.EOF Then
                        Nov_Firmada = False
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * 6) & "Hay una novedad horaria NO aprobado. Se descarta."
                        End If
                    Else
                        Nov_Firmada = True
                        Encontro_Jus = True
                    End If
                Else
                    Nov_Firmada = True
                    Encontro_Jus = True
                End If
            Else
                Nov_Firmada = True
                Encontro_Jus = True
            End If
                
            'FGZ - 10/06/2011 -------------------------------------------------------
            If Nov_Firmada Then
                Justif.Tiene_Justif = True
                Justif.Numero = objRs!jusnro
                If Not IsNull(objRs!turnro) Then
                    StrSql = "SELECT * FROM gti_turno WHERE turnro = " & objRs!turnro
                    OpenRecordset StrSql, rs_Tur
                    If Not rs_Tur.EOF Then
                        Tur.tiene_turno = True
                        Tur.Numero = rs_Tur!turnro
                        Tur.Nombre = Trim(rs_Tur!turdesabr)
                        Tur.Tipo = rs_Tur!TipoTurno
                        Justif.justif_turno = True
                        'FGZ - 10/06/2011 -------------------------------------------------------
                        'le saque el exit y lo manejo abajo con la condicion de si tiene justificacion
                        'Exit Sub
                    End If
                End If
            End If
            objRs.MoveNext
        Loop
    
    End If
    
If Not Justif.justif_turno Then
    'Si no tiene justificaci¢n, busco los partes de Asignaci¢n de horas
    StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & lngTernro & ") AND " & _
             "(gttempdesde <= " & ConvFecha(Fecha) & ") and (" & ConvFecha(Fecha) & " <= gttemphasta)"
    OpenRecordset StrSql, objRs
    'If Not objRs.EOF Then
    Firmado = False
    Do While Not objRs.EOF And Not Firmado
        'FGZ - 31/05/2010  --------------------------------------------------------------------------
        'Verifico que no haya sido generado fuera de termino y en ese caso reviso que esté aprobado
        StrSql = "SELECT input_ft.idnro,input_ft.origen, gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
        StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
        StrSql = StrSql & " WHERE idtipoinput = 6 "
        StrSql = StrSql & " AND origen = " & objRs!GCPNRO
        OpenRecordset StrSql, rs_FT
        If Not rs_FT.EOF Then
            'El parte fué cargado fuera de termimo
            If rs_FT!ftap = -1 Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria fuera de termino aprobado."
                End If
                Firmado = True
                Call InsertarFT(rs_FT!idnro, 6, rs_FT!Origen)
                'Parte_FT = True
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria fuera de termino NO aprobado. Se descarta."
                End If
                Firmado = False
                'Parte_FT = False
            End If
        Else
            'Verificar si esta en el NIVEL FINAL DE FIRMA ACTIVO para partes de asignacion Horaria
            StrSql = "select * from cystipo where cystipnro = 17"
            OpenRecordset StrSql, rs_Firma
            If Not rs_Firma.EOF Then
                If rs_Firma!cystipact = -1 Then
                    StrSql = "SELECT * FROM cysfirmas "
                    StrSql = StrSql & " WHERE cysfirfin = -1"
                    StrSql = StrSql & " AND cysfircodext = '" & objRs!GCPNRO & "' "
                    StrSql = StrSql & " AND cystipnro = 17"
                    OpenRecordset StrSql, rs
                    If rs.EOF Then
                        Firmado = False
                    Else
                        Firmado = True
                    End If
                Else
                    Firmado = True
                End If
            Else
                Firmado = True
            End If
        End If
        If Firmado Then
            P_Asignacion = True
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria sin fin de firma. Se descartó"
            End If
        End If
    'End If
        objRs.MoveNext
    Loop
    
 
    'Si no tiene justificación busca los partes de Cambio de Turno
    StrSql = "SELECT gti_turno.turdesabr,gti_turforpago.turnro,gti_turforpago.fpgonro,gti_reldtur.grtddesde, "
    StrSql = StrSql & "gti_reldtur.grtoffset, gti_turno.turcompensa, gti_turno.tipoturno,"
    StrSql = StrSql & " gti_formapago.fpgodesabr, gti_reldtur.gcpnro "
    StrSql = StrSql & " FROM  gti_reldtur "
    StrSql = StrSql & " INNER JOIN gti_turforpago ON "
    StrSql = StrSql & " gti_reldtur.turnro = gti_turforpago.turfpagnro "
    StrSql = StrSql & " INNER JOIN gti_turno ON gti_turno.turnro=gti_turforpago.turnro "
    StrSql = StrSql & " INNER JOIN gti_formapago ON gti_turforpago.fpgonro = gti_formapago.fpgonro "
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " (ternro = " & lngTernro & " ) AND "
    StrSql = StrSql & " (grtddesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= grtdhasta) "
    StrSql = StrSql & " OR (grtdhasta is null) ) "
    OpenRecordset StrSql, objRs
    Do While Not objRs.EOF
        'FGZ - 31/05/2010  --------------------------------------------------------------------------
        'Verifico que no haya sido generado fuera de termino y en ese caso reviso que esté aprobado
        StrSql = "SELECT input_ft.idnro,input_ft.origen, gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
        StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
        StrSql = StrSql & " WHERE idtipoinput = 7 "
        StrSql = StrSql & " AND origen = " & objRs!GCPNRO
        OpenRecordset StrSql, rs_FT
        If Not rs_FT.EOF Then
            'El parte fué cargado fuera de termimo
            If rs_FT!ftap = -1 Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de cambio de turno cargado fuera de termino aprobado."
                End If
                Firmado = True
                Call InsertarFT(rs_FT!idnro, 7, rs_FT!Origen)
                'Parte_FT = True
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de cambio de turno fuera de termino NO aprobado. Se descarta."
                End If
                Firmado = False
                'Parte_FT = False
            End If
        Else
            'Chequeo si tiene circuito de firma activo para los partes de asignacion horaria
            'Verificar si esta ACTIVO para partes de cambio de turno
            StrSql = "select * from cystipo where cystipnro = 4"
            OpenRecordset StrSql, rs_Firma
            If Not rs_Firma.EOF Then
                If rs_Firma!cystipact = -1 Then
                    StrSql = "SELECT * FROM cysfirmas "
                    StrSql = StrSql & " WHERE cysfirfin = -1"
                    StrSql = StrSql & " AND cysfircodext = '" & objRs!GCPNRO & "' "
                    StrSql = StrSql & " AND cystipnro = 4"
                    OpenRecordset StrSql, rs
                    If rs.EOF Then
                        Firmado = False
                    Else
                        Firmado = True
                    End If
                Else
                    Firmado = True
                End If
            Else
                Firmado = True
            End If
        End If
        If Firmado Then
            Tur.tiene_turno = True
            Tur.Numero = objRs!turnro
            Tur.Nombre = Trim(objRs!turdesabr)
            NombreFPago = Trim(objRs!fpgodesabr)
            Nro_fpgo = objRs!fpgonro
            Fecha_Inicio = DateAdd("d", objRs!grtddesde, -(0 & objRs!grtoffset))
            p_turcomp = objRs!turcompensa     'Fecha de inicio del turno
            Tur.Tipo = objRs!TipoTurno
            
            StrSql = " SELECT * FROM his_estructura "
            StrSql = StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
            StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
            StrSql = StrSql & " WHERE (tanro = " & lngAlcanGrupo & ") AND (ternro = " & Empleado.Ternro & ") AND "
            StrSql = StrSql & " (htetdesde <= " & ConvFecha(Fecha) & ") AND "
            StrSql = StrSql & " ((" & ConvFecha(Fecha) & " <= htethasta) or (htethasta is null))"
            StrSql = StrSql & " ORDER BY alcance_testr.alteorden DESC, his_estructura.htetdesde Desc "
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                Empl.NombreGrupo = Trim(objRs!estrdabr)
                Empl.Grupo = objRs!estrnro
            End If
            Exit Sub
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de cambio de turno sin fin de firma. Se descartó"
            End If
        End If
        
        'Siguiente
        objRs.MoveNext
    Loop

    'Buscar si la fecha tiene un Turno Asociado en forma Directa en el Histórico
    StrSql = " SELECT estructura.estrdabr,his_estructura.htetdesde,gti_turfpgogru.*,gti_formapago.fpgodesabr,gti_formapago.fpgonro,gti_turno.turnro,gti_turno.TipoTurno,gti_turno.turcompensa,gti_turno.turdesabr,Alcance_Testr.alteorden " & _
             " From his_estructura " & _
             " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro " & _
             " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " & _
             " INNER JOIN gti_turfpgogru ON gti_turfpgogru.estrnro = estructura.estrnro " & _
             " INNER JOIN gti_turforpago ON gti_turforpago.turfpagnro = gti_turfpgogru.turfpagnro " & _
             " INNER JOIN gti_formapago ON gti_formapago.fpgonro = gti_turforpago.fpgonro " & _
             " INNER JOIN gti_turno ON gti_turno.turnro = gti_turforpago.turnro " & _
             " Where (Alcance_Testr.tanro = " & lngAlcanGrupo & ") AND " & _
             " (his_estructura.ternro = " & lngTernro & ") AND " & _
             " (htetdesde <= " & ConvFecha(Fecha) & ")  AND " & _
             "((htethasta >= " & ConvFecha(Fecha) & ")" & _
             " OR (htethasta is null )) AND (fechavalidez <= " & ConvFecha(Fecha) & " ) " & _
             " ORDER BY Alcance_Testr.alteorden DESC,his_estructura.htetdesde DESC,gti_turfpgogru.FechaValidez Desc "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'Existe un turno asociado para la fecha
        Tur.tiene_turno = True
        Tur.Numero = objRs!turnro
        Tur.Nombre = Trim(objRs!turdesabr)
        NombreFPago = Trim(objRs!fpgodesabr)
        Empl.NombreGrupo = Trim(objRs!estrdabr)
        Empl.Grupo = objRs!estrnro
        Tur.Tipo = objRs!TipoTurno
        p_turcomp = objRs!turcompensa
        Fecha_Inicio = DateAdd("d", objRs!FechaValidez, -(0 & objRs!offset))
        Nro_fpgo = objRs!fpgonro
    Else
        'Buscar el Turno Actual del empleado */
        
    End If
End If

'cierro y libero
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
If rs_Firma.State = adStateOpen Then rs_Firma.Close
Set rs_Firma = Nothing
If rs_FT.State = adStateOpen Then rs_FT.Close
Set rs_FT = Nothing
If rs_Tur.State = adStateOpen Then rs_Tur.Close
Set rs_Tur = Nothing
    
End Sub


Public Sub Buscar_Turno_OLD(Fecha As Date, lngTernro As Long, depurar As Boolean)
Dim objRsTurnoFPago As New ADODB.Recordset
Dim objRsReldTur As New ADODB.Recordset
Dim objRsTurnoFPGrupo As New ADODB.Recordset
Dim objLicAut As New ADODB.Recordset

'    'FGZ - 01/06/2007 ----- Mejoras -----
'    Cantidad_Turnos = Cantidad_Turnos + 1
'    'FGZ - 01/06/2007 ----- Mejoras -----
    
    PFecha = Fecha
    Justif.Tiene_Justif = False
    Justif.justif_turno = False
    Tur.tiene_turno = False
    P_Asignacion = False
    
    'FGZ - 27/06/2007 - le agregué estas inicializaciones -------------------
    Tur.Nombre = ""
    Tur.Numero = 0
    Tur.tiene_turno = False
    Tur.Tipo = 0
    
    Justif.justif_turno = False
    Justif.Numero = 0
    Justif.Tiene_Justif = False
    'FGZ - 27/06/2007 - le agregué estas inicializaciones -------------------
    
    
'     StrSql = "SELECT * FROM gti_justificacion WHERE (ternro = " & lngTernro & ") AND " & _
'              "(juseltipo <> 3) AND (jusdesde <= " & ConvFecha(Fecha) & ") AND " & _
'              "(" & ConvFecha(Fecha) & " <= jushasta)"
    
'     StrSql = "SELECT * FROM gti_justificacion WHERE (ternro = " & lngTernro & ") AND " & _
'              " (jusdesde <= " & ConvFecha(Fecha) & ") AND " & _
'              "(" & ConvFecha(Fecha) & " <= jushasta)"

    'FGZ - 15/08/2006 - debe ordenar las justificaciones por el campo juseltipo
    
    
    'FGZ - 13/05/2008  - levanto solamente las licencias aprobadas
'    StrSql = "SELECT * FROM gti_justificacion WHERE (ternro = " & lngTernro & ") AND "
'    StrSql = StrSql & " (jusdesde <= " & ConvFecha(Fecha) & ") AND "
'    StrSql = StrSql & "(" & ConvFecha(Fecha) & " <= jushasta) "
'    StrSql = StrSql & " order by juseltipo "
    
    StrSql = ""
    Select Case TipoBD
    Case 4:
        StrSql = "SELECT * FROM ("
    End Select
    StrSql = StrSql & "(SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN emp_lic ON gti_justificacion.juscodext = emp_lic.emp_licnro "
    StrSql = StrSql & " WHERE (ternro = " & lngTernro & ") "
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND emp_lic.licestnro = 2"
    StrSql = StrSql & " AND jussigla = 'LIC'"
    StrSql = StrSql & " )UNION ("
    StrSql = StrSql & " SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " WHERE (Ternro = " & lngTernro & ")"
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND jussigla <> 'LIC'"
    StrSql = StrSql & ")"
    Select Case TipoBD
    Case 4:
        StrSql = StrSql & ")"
    End Select
    StrSql = StrSql & " ORDER BY juseltipo "
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        'Esto no quiere decir que no hay justificaciones sino que no hay o no hay aprobadas
    Else
        Justif.Tiene_Justif = True
        Justif.Numero = objRs!jusnro
        If Not IsNull(objRs!turnro) Then
            StrSql = "SELECT * FROM gti_turno WHERE turnro = " & objRs!turnro
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                Tur.tiene_turno = True
                Tur.Numero = objRs!turnro
                Tur.Nombre = Trim(objRs!turdesabr)
                Tur.Tipo = objRs!TipoTurno
                Justif.justif_turno = True
                Exit Sub
            End If
        End If
    End If
    
    ' Si no tiene justificaci¢n, busco los partes de Asignaci¢n de horas
    StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & lngTernro & ") AND " & _
             "(gttempdesde <= " & ConvFecha(Fecha) & ") and (" & ConvFecha(Fecha) & " <= gttemphasta)"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then P_Asignacion = True
    
    ' Si no tiene justificación busca los partes de Cambio de Turno
'    StrSql = " SELECT6 gti_turno.turdesabr,gti_turforpago.turnro,gti_turforpago.fpgonro,gti_reldtur.grtddesde, gti_reldtur.grtoffset,gti_turno.turcompensa,gti_turno.tipoturno,gti_formapago.fpgodesabr " & _
'             " FROM  gti_reldtur " & _
'             " INNER JOIN gti_turforpago ON gti_reldtur.turnro = gti_turforpago.turfpagnro " & _
'             " INNER JOIN gti_turno ON gti_turno.turnro=gti_turforpago.turnro " & _
'             " INNER JOIN gti_formapago ON gti_turforpago.fpgonro = gti_formapago.fpgonro " & _
'             " WHERE (ternro = " & lngTernro & " ) AND " & _
'             " (grtddesde <= " & ConvFecha(Fecha) & ") AND " & _
'             " ((" & ConvFecha(Fecha) & " >= grtdhasta) OR (grtdhasta is null) ) "

    StrSql = "SELECT gti_turno.turdesabr,gti_turforpago.turnro,gti_turforpago.fpgonro,gti_reldtur.grtddesde, "
    StrSql = StrSql & "gti_reldtur.grtoffset, gti_turno.turcompensa, gti_turno.tipoturno,"
    StrSql = StrSql & " gti_formapago.fpgodesabr "
    StrSql = StrSql & " FROM  gti_reldtur "
    StrSql = StrSql & " INNER JOIN gti_turforpago ON "
    StrSql = StrSql & " gti_reldtur.turnro = gti_turforpago.turfpagnro "
    StrSql = StrSql & " INNER JOIN gti_turno ON gti_turno.turnro=gti_turforpago.turnro "
    StrSql = StrSql & " INNER JOIN gti_formapago ON gti_turforpago.fpgonro = gti_formapago.fpgonro "
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " (ternro = " & lngTernro & " ) AND "
    StrSql = StrSql & " (grtddesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= grtdhasta) "
    StrSql = StrSql & " OR (grtdhasta is null) ) "

    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Tur.tiene_turno = True
        Tur.Numero = objRs!turnro
        Tur.Nombre = Trim(objRs!turdesabr)
        NombreFPago = Trim(objRs!fpgodesabr)
        Nro_fpgo = objRs!fpgonro
        Fecha_Inicio = DateAdd("d", objRs!grtddesde, -(0 & objRs!grtoffset))
        p_turcomp = objRs!turcompensa     'Fecha de inicio del turno
        Tur.Tipo = objRs!TipoTurno
        
        StrSql = " SELECT * FROM his_estructura "
        StrSql = StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
        StrSql = StrSql & " WHERE (tanro = " & lngAlcanGrupo & ") AND (ternro = " & Empleado.Ternro & ") AND "
        StrSql = StrSql & " (htetdesde <= " & ConvFecha(Fecha) & ") AND "
        StrSql = StrSql & " ((" & ConvFecha(Fecha) & " <= htethasta) or (htethasta is null))"
        StrSql = StrSql & " ORDER BY alcance_testr.alteorden DESC, his_estructura.htetdesde Desc "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Empl.NombreGrupo = Trim(objRs!estrdabr)
            Empl.Grupo = objRs!estrnro
        End If
        Exit Sub
    End If

    ' Buscar si la fecha tiene un Turno Asociado en forma Directa en el Histórico
    StrSql = " SELECT estructura.estrdabr,his_estructura.htetdesde,gti_turfpgogru.*,gti_formapago.fpgodesabr,gti_formapago.fpgonro,gti_turno.turnro,gti_turno.TipoTurno,gti_turno.turcompensa,gti_turno.turdesabr,Alcance_Testr.alteorden " & _
             " From his_estructura " & _
             " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro " & _
             " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " & _
             " INNER JOIN gti_turfpgogru ON gti_turfpgogru.estrnro = estructura.estrnro " & _
             " INNER JOIN gti_turforpago ON gti_turforpago.turfpagnro = gti_turfpgogru.turfpagnro " & _
             " INNER JOIN gti_formapago ON gti_formapago.fpgonro = gti_turforpago.fpgonro " & _
             " INNER JOIN gti_turno ON gti_turno.turnro = gti_turforpago.turnro " & _
             " Where (Alcance_Testr.tanro = " & lngAlcanGrupo & ") AND " & _
             " (his_estructura.ternro = " & lngTernro & ") AND " & _
             " (htetdesde <= " & ConvFecha(Fecha) & ")  AND " & _
             "((htethasta >= " & ConvFecha(Fecha) & ")" & _
             " OR (htethasta is null )) AND (fechavalidez <= " & ConvFecha(Fecha) & " ) " & _
             " ORDER BY Alcance_Testr.alteorden DESC,his_estructura.htetdesde DESC,gti_turfpgogru.FechaValidez Desc "
    
    If objRs.State = adStateOpen Then objRs.Close
    'objrs.MaxRecords = 1
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
        'Existe un turno asociado para la fecha
        Tur.tiene_turno = True
        Tur.Numero = objRs!turnro
        Tur.Nombre = Trim(objRs!turdesabr)
        NombreFPago = Trim(objRs!fpgodesabr)
        Empl.NombreGrupo = Trim(objRs!estrdabr)
        Empl.Grupo = objRs!estrnro
        Tur.Tipo = objRs!TipoTurno
        p_turcomp = objRs!turcompensa
        Fecha_Inicio = DateAdd("d", objRs!FechaValidez, -(0 & objRs!offset))
        Nro_fpgo = objRs!fpgonro
    Else
        'Buscar el Turno Actual del empleado */
        
'        StrSql = "SELECT estructura.estrdabr,gti_turfpgogru.*,gti_formapago.fpgodesabr," & _
'                 " gti_formapago.fpgonro,gti_turno.turnro,gti_turno.TipoTurno,gti_turno.turcompensa , gti_turno.turdesabr,alcance_testr.alteorden " & _
'                 " From estruc_actual " & _
'                 " INNER JOIN Alcance_Testr ON estruc_actual.tenro = Alcance_Testr.tenro " & _
'                 " INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro " & _
'                 " INNER JOIN gti_turfpgogru ON gti_turfpgogru.estrnro = estructura.estrnro " & _
'                 " INNER JOIN gti_turforpago ON gti_turforpago.turfpagnro = gti_turfpgogru.turfpagnro " & _
'                 " INNER JOIN gti_formapago ON gti_formapago.fpgonro = gti_turforpago.fpgonro " & _
'                 " INNER JOIN gti_turno ON gti_turno.turnro = gti_turforpago.turnro " & _
'                 " Where (Alcance_Testr.tanro = " & lngAlcanGrupo & " ) AND " & _
'                 " (estruc_actual.ternro = " & lngTernro & ") " & _
'                 " ORDER BY alcance_testr.alteorden DESC, gti_turfpgogru.FechaValidez Desc "
'
'        If objrs.State = adStateOpen Then objrs.Close
'        objrs.MaxRecords = 1
'        OpenRecordset StrSql, objrs
'
'        If Not objrs.EOF Then
'            Tur.Tiene_Turno = True
'            Tur.Numero = objrs!turnro
'            Tur.Nombre = Trim(objrs!turdesabr)
'            NombreFPago = Trim(objrs!fpgodesabr)
'            Empl.NombreGrupo = Trim(objrs!estrdabr)
'            Empl.Grupo = objrs!estrnro
'            Tur.Tipo = objrs!TipoTurno
'            P_turcomp = objrs!turcompensa
'            Fecha_inicio = DateAdd("d", objrs!FechaValidez, -objrs!offset)
'            Nro_fpgo = objrs!fpgonro
'        End If
    End If
'
    'Empleado.Grupo = nro_grupo
    If depurar Then
        'GeneraTraza lngTernro, PFecha, "Tiene turno asociado", Str(Tur.tiene_turno)
        'GeneraTraza lngTernro, PFecha, "Fecha de validez", Format(Fecha_inicio, "DD/MM/YYYY")
        'GeneraTraza lngTernro, PFecha, "Código del turno", Str(Tur.Numero)
        'If Tur.Tipo = 1 Then
        '    GeneraTraza lngTernro, PFecha, "Tipo turno", "Normal"
        'ElseIf Tur.Tipo = 2 Then
        '    GeneraTraza lngTernro, PFecha, "Tipo turno", "Libre"
        'End If
        'GeneraTraza lngTernro, PFecha, "Código del grupo", Str(Empl.Grupo)
        'GeneraTraza lngTernro, PFecha, "Forma de pago", Str(Nro_fpgo)
        'GeneraTraza lngTernro, PFecha, "Tiene parte de asignación", Str(P_asignacion)
    End If
    
End Sub



Private Sub Class_Terminate()
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub
