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
Dim StrSql As String

Dim P_asignacion As Boolean
Dim Nro_fpgo As Long
Dim Fecha_inicio As Date
Dim p_turcomp As Boolean
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
    Fecha_inicio = Fecha
End Property
Property Get FechaInicio() As Date
    FechaInicio = Fecha_inicio
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
    P_asignacion = tiene
End Property

Property Get Tiene_PAsignacion() As Boolean
    Tiene_PAsignacion = P_asignacion
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
Dim objRsTurnoFPago As New ADODB.Recordset
Dim objRsReldTur As New ADODB.Recordset
Dim objRsTurnoFPGrupo As New ADODB.Recordset

    PFecha = Fecha
    Justif.Tiene_Justif = False
    Justif.justif_turno = False
    Tur.tiene_turno = False
    P_asignacion = False
    
     StrSql = "SELECT * FROM gti_justificacion WHERE (ternro = " & lngTernro & ") AND " & _
              "(juseltipo <> 3) AND (jusdesde <= " & ConvFecha(Fecha) & ") AND " & _
              "(" & ConvFecha(Fecha) & " <= jushasta)"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'Existe Una Justificacion Particular
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
    If Not objRs.EOF Then P_asignacion = True
    
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
        Fecha_inicio = DateAdd("d", objRs!grtddesde, -(0 & objRs!grtoffset))
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
        Fecha_inicio = DateAdd("d", objRs!FechaValidez, -(0 & objRs!offset))
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
