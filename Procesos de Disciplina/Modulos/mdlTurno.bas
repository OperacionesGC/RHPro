Attribute VB_Name = "mdlTurno"
Option Explicit

Public Type TipoTurno
    compensa As Boolean
    Nro_Grupo As Long
    Nro_Justif As Long
    Tiene_Justif As Boolean
    justif_turno As Boolean
    Fecha_Inicio As Date
    Nro_FPago As Long
    NombreFPago As String
    Nro_Turno As Long
    tiene_turno As Boolean
    Tipo_Turno As Integer
    P_Asignacion As Boolean
End Type

Public Type TTurno
    tiene_turno As Boolean
    Numero As Long
    Tipo As Integer
    Nombre As String
End Type

Public Type TJustif
    Tiene_Justif As Boolean
    justif_turno As Boolean
    Numero As Long
End Type

Public Type TEmp
    Legajo As Long
    Grupo As Long
    NombreGrupo As String
End Type

Public Type TDia
    Dia_Libre As Boolean
    Nro_Dia As Long
    Nro_Subturno As Long
    NombreSubTurno As String
    Orden_Dia As Long
    Trabaja As Boolean
    Genera As Integer
End Type

'FGZ - 15/11/2006
Public Type THT
    FE1 As Date
    E1 As String
    FS1 As Date
    S1 As String
    FE2 As Date
    E2 As String
    FS2 As Date
    S2 As String
    FE3 As Date
    E3 As String
    FS3 As Date
    S3 As String
End Type

'Feriado
Public Type TFeriado
    Feriado As Boolean
    Por_Estructura As Boolean
    NroEstr As Long
    NombreEstr As String
End Type

'Dia
Dim blnTrabaja  As Boolean
Dim Ordendia As Integer
'Dim Nro_Dia As Long
Dim blnDia_libre As Boolean
Dim Nro_Subturno As Long
Dim NombreSTurno As String
Dim PFecha As Date

'Feriado
Dim nroConv As Long
Dim NombreConvenio As String
Dim nroSuc As Long
Dim NombreSucursal As String
Dim NroEstr As Long
Dim NombreEstr As String


Public Sub Buscar_Turno_Nuevo(ByVal Fecha As Date, ByVal Tercero As Long, ByVal depurar As Boolean, ByRef T As TipoTurno, ByRef Turno As TTurno, ByRef Justif As TJustif, ByRef Empleado As Templeado)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el turno del empleado en la fecha.
' Autor      : FGZ
' Fecha      : 03/11/2004
' Ultima Mod.: 03/11/2004
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
Dim rs_Firma As New ADODB.Recordset
Dim Firmado As Boolean
Dim rs_FT As New ADODB.Recordset


    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 5) & "Inicio Buscar_Turno_Nuevo()"
        Flog.writeline
    End If
    
    T.Tiene_Justif = False
    T.justif_turno = False
    T.tiene_turno = False
    T.P_Asignacion = False
    
    Justif.justif_turno = False
    Justif.Tiene_Justif = False
    
    Turno.tiene_turno = False
    
    'FGZ - 13/05/2008  - levanto solamente las licencias aprobadas
'     StrSql = "SELECT * FROM gti_justificacion WHERE (ternro = " & Tercero & ") AND " & _
'              "(juseltipo <> 3) AND (jusdesde <= " & ConvFecha(Fecha) & ") AND " & _
'              "(" & ConvFecha(Fecha) & " <= jushasta)"
    
    StrSql = ""
    Select Case TipoBD
    Case 4:
        StrSql = "SELECT * FROM ("
    End Select
    StrSql = StrSql & "(SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN emp_lic ON gti_justificacion.juscodext = emp_lic.emp_licnro "
    StrSql = StrSql & " WHERE (ternro = " & Tercero & ") "
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND emp_lic.licestnro = 2 "
    StrSql = StrSql & " AND jussigla = 'LIC' AND juseltipo <> 3 "
'    'FGZ - 19/05/2010 ------------ Control FT -------------
'    StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
'    'FGZ - 19/05/2010 ------------ Control FT -------------
'    StrSql = StrSql & " )UNION ("
'    StrSql = StrSql & " SELECT gti_justificacion.* FROM gti_justificacion "
'    StrSql = StrSql & " WHERE (Ternro = " & Tercero & ")"
'    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
'    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
'    StrSql = StrSql & " AND jussigla <> 'LIC' AND juseltipo <> 3"
'    'FGZ - 14/8/2008 - le agregué esta linea por las justificaciones automaticas de la Politica 400
'    StrSql = StrSql & " AND jussigla <> 'ALM'"
'    'FGZ - 14/8/2008 - le agregué esta linea por las justificaciones automaticas de la Politica 400
'    StrSql = StrSql & ")"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
    StrSql = StrSql & " )UNION ("
    StrSql = StrSql & " SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN gti_novedad ON gti_justificacion.juscodext = gti_novedad.gnovnro "
    StrSql = StrSql & " WHERE (Ternro = " & Tercero & ")"
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND jussigla = 'NOV'"
    StrSql = StrSql & " AND jussigla <> 'ALM'"
    StrSql = StrSql & " AND (gti_novedad.ft = 0 OR (gti_novedad.ft = -1 AND gti_novedad.ftap = -1))"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " )UNION ("
    StrSql = StrSql & " SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " WHERE (Ternro = " & Tercero & ")"
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
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
    
        'verifico que sea un input_FT
        
    
    
        'Existe Una Justificacion Particular
        Justif.Tiene_Justif = True
        Justif.Numero = rs!jusnro
        If Not IsNull(rs!turnro) Then
            StrSql = "SELECT * FROM gti_turno WHERE turnro = " & rs!turnro
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                Turno.tiene_turno = True
                Turno.Numero = rs!turnro
                Turno.Nombre = Trim(rs!turdesabr)
                Turno.Tipo = rs!TipoTurno
                
                Justif.justif_turno = True
                
                T.Nro_Turno = Turno.Numero
                T.justif_turno = True
                T.Tiene_Justif = True
                T.tiene_turno = True
                Exit Sub
            End If
        End If
    End If
    
    'Si no tiene justificaci¢n, busco los partes de Asignaci¢n de horas
    StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & Tercero & ") AND " & _
             " (gttempdesde <= " & ConvFecha(Fecha) & ") and (" & ConvFecha(Fecha) & " <= gttemphasta)"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    'If Not rs.EOF Then
    Firmado = False
    Do While Not rs.EOF And Not Firmado
        'FGZ - 31/05/2010  --------------------------------------------------------------------------
        'Verifico que no haya sido generado fuera de termino y en ese caso reviso que esté aprobado
        StrSql = "SELECT input_ft.idnro,input_ft.origen, gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
        StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
        StrSql = StrSql & " WHERE idtipoinput = 6 "
        StrSql = StrSql & " AND origen = " & rs!gcpnro
        OpenRecordset StrSql, rs_FT
        If Not rs_FT.EOF Then
            'El parte fué cargado fuera de termimo
            If rs_FT!ftap = -1 Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria fuera de termino aprobado."
                End If
                Firmado = True
                Call InsertarFT(rs_FT!idnro, 6, rs_FT!Origen)
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria fuera de termino NO aprobado. Se descarta."
                End If
                Firmado = False
            End If
        Else
            'Verificar si esta en el NIVEL FINAL DE FIRMA ACTIVO para partes de cambio de turno
            StrSql = "select * from cystipo where cystipnro = 17"
            OpenRecordset StrSql, rs_Firma
            If Not rs_Firma.EOF Then
                If rs_Firma!cystipact = -1 Then
                    StrSql = "SELECT * FROM cysfirmas "
                    StrSql = StrSql & " WHERE cysfirfin = -1"
                    StrSql = StrSql & " AND cysfircodext = '" & rs!gcpnro & "' "
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
            'FGZ - 03/03/2010 - Comenté esta linea porque afectaba a una variable global
            '                   que mantiene si hay partes de asignacion horaria en el dia que se esta procesando
            '                   Cuando esta funcion se usa para buscar el turno de dias posteriores puede causar problemas
            'P_Asignacion = True
            T.P_Asignacion = True
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria sin fin de firma. Se descartó"
            End If
        End If
    'End If
    
        rs.MoveNext
    Loop
    
    
'------------------ FGZ - Circuito de firmas -------------------------
'    'Si no tiene justificación busca los partes de Cambio de Turno
'    StrSql = "SELECT gti_turno.turdesabr,gti_turforpago.turnro,gti_turforpago.fpgonro,gti_reldtur.grtddesde, "
'    StrSql = StrSql & "gti_reldtur.grtoffset, gti_turno.turcompensa, gti_turno.tipoturno,"
'    StrSql = StrSql & " gti_formapago.fpgodesabr "
'    StrSql = StrSql & " FROM  gti_reldtur "
'    StrSql = StrSql & " INNER JOIN gti_turforpago ON gti_reldtur.turnro = gti_turforpago.turfpagnro "
'    StrSql = StrSql & " INNER JOIN gti_turno ON gti_turno.turnro=gti_turforpago.turnro "
'    StrSql = StrSql & " INNER JOIN gti_formapago ON gti_turforpago.fpgonro = gti_formapago.fpgonro "
'    StrSql = StrSql & " WHERE "
'    StrSql = StrSql & " (ternro = " & Tercero & " ) AND "
'    StrSql = StrSql & " (grtddesde <= " & ConvFecha(Fecha) & ")"
'    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= grtdhasta) "
'    StrSql = StrSql & " OR (grtdhasta is null) ) "
'    If rs.State = adStateOpen Then rs.Close
'    OpenRecordset StrSql, rs
'    If Not rs.EOF Then
'        Turno.tiene_turno = True
'        Turno.Numero = rs!turnro
'        Turno.Nombre = Trim(rs!turdesabr)
'        Turno.Tipo = rs!TipoTurno
'
'        T.tiene_turno = True
'        T.Nro_Turno = rs!turnro
'        T.Nro_FPago = rs!fpgonro
'        T.Fecha_Inicio = DateAdd("d", rs!grtddesde, -(0 & rs!grtoffset))
'        T.compensa = rs!turcompensa     'Fecha de inicio del turno
'        T.NombreFPago = Trim(rs!fpgodesabr)
'
'        StrSql = " SELECT * FROM his_estructura "
'        StrSql = StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
'        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
'        StrSql = StrSql & " WHERE (tanro = " & lngAlcanGrupo & ") AND (ternro = " & Empleado.Ternro & ") AND "
'        StrSql = StrSql & " (htetdesde <= " & ConvFecha(Fecha) & ") AND "
'        StrSql = StrSql & " ((" & ConvFecha(Fecha) & " <= htethasta) or (htethasta is null))"
'        StrSql = StrSql & " ORDER BY alcance_testr.alteorden DESC, his_estructura.htetdesde Desc "
'        If rs.State = adStateOpen Then rs.Close
'        OpenRecordset StrSql, rs
'        If Not rs.EOF Then
'            Empleado.NombreGrupo = Trim(rs!estrdabr)
'            Empleado.Grupo = rs!estrnro
'        End If
'        Exit Sub
'    End If

    'Si no tiene justificación busca los partes de Cambio de Turno
    StrSql = "SELECT gti_turno.turdesabr,gti_turforpago.turnro,gti_turforpago.fpgonro,gti_reldtur.grtddesde, "
    StrSql = StrSql & "gti_reldtur.grtoffset, gti_turno.turcompensa, gti_turno.tipoturno,"
    StrSql = StrSql & " gti_formapago.fpgodesabr, gti_reldtur.gcpnro "
    StrSql = StrSql & " FROM  gti_reldtur "
    StrSql = StrSql & " INNER JOIN gti_turforpago ON gti_reldtur.turnro = gti_turforpago.turfpagnro "
    StrSql = StrSql & " INNER JOIN gti_turno ON gti_turno.turnro=gti_turforpago.turnro "
    StrSql = StrSql & " INNER JOIN gti_formapago ON gti_turforpago.fpgonro = gti_formapago.fpgonro "
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " (ternro = " & Tercero & " ) AND "
    StrSql = StrSql & " (grtddesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= grtdhasta) "
    StrSql = StrSql & " OR (grtdhasta is null) ) "
    OpenRecordset StrSql, rs
    Do While Not rs.EOF
        'FGZ - 31/05/2010  --------------------------------------------------------------------------
        'Verifico que no haya sido generado fuera de termino y en ese caso reviso que esté aprobado
        StrSql = "SELECT input_ft.idnro,input_ft.origen, gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
        StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
        StrSql = StrSql & " WHERE idtipoinput = 7 "
        StrSql = StrSql & " AND origen = " & rs!gcpnro
        OpenRecordset StrSql, rs_FT
        If Not rs_FT.EOF Then
            'El parte fué cargado fuera de termimo
            If rs_FT!ftap = -1 Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de cambio de turno cargado fuera de termino aprobado."
                End If
                Firmado = True
                Call InsertarFT(rs_FT!idnro, 7, rs_FT!Origen)
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de cambio de turno fuera de termino NO aprobado. Se descarta."
                End If
                Firmado = False
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
                    StrSql = StrSql & " AND cysfircodext = '" & rs!gcpnro & "' "
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
            If depurar Then
                Flog.writeline Espacios(Tabulador * 6) & "Firmado"
            End If

            Turno.tiene_turno = True
            Turno.Numero = rs!turnro
            Turno.Nombre = Trim(rs!turdesabr)
            Turno.Tipo = rs!TipoTurno
            
            T.tiene_turno = True
            T.Nro_Turno = rs!turnro
            T.Nro_FPago = rs!fpgonro
            T.Fecha_Inicio = DateAdd("d", rs!grtddesde, -(0 & rs!grtoffset))
            T.compensa = rs!turcompensa     'Fecha de inicio del turno
            T.NombreFPago = Trim(rs!fpgodesabr)
            
            If depurar Then
                Flog.writeline Espacios(Tabulador * 6) & "Busco la estructura"
            End If
            
            'FGZ - 24/09/2008 - Se cambió el query
            'StrSql = " SELECT * FROM his_estructura "
            StrSql = " SELECT his_estructura.estrnro, estructura.estrdabr FROM his_estructura "
            StrSql = StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
            StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
            StrSql = StrSql & " WHERE (tanro = " & lngAlcanGrupo & ") AND (ternro = " & Empleado.Ternro & ") AND "
            StrSql = StrSql & " (htetdesde <= " & ConvFecha(Fecha) & ") AND "
            StrSql = StrSql & " ((" & ConvFecha(Fecha) & " <= htethasta) or (htethasta is null))"
            StrSql = StrSql & " ORDER BY alcance_testr.alteorden DESC, his_estructura.htetdesde Desc "
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                Empleado.NombreGrupo = Trim(rs!estrdabr)
                Empleado.Grupo = rs!estrnro
            End If
            Exit Sub
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de cambio de turno sin fin de firma. Se descartó"
            End If
        End If
        
        'Siguiente
        If Not rs.EOF Then
            rs.MoveNext
        End If
    Loop
    If depurar Then
        Flog.writeline Espacios(Tabulador * 6) & "Fin Circuito"
    End If

'------------------ FGZ - Circuito de firmas -------------------------


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
             " (his_estructura.ternro = " & Tercero & ") AND " & _
             " (htetdesde <= " & ConvFecha(Fecha) & ")  AND " & _
             "((htethasta >= " & ConvFecha(Fecha) & ")" & _
             " OR (htethasta is null )) AND (fechavalidez <= " & ConvFecha(Fecha) & " ) " & _
             " ORDER BY Alcance_Testr.alteorden DESC,his_estructura.htetdesde DESC,gti_turfpgogru.FechaValidez Desc "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        'Existe un turno asociado para la fecha
        
        Turno.tiene_turno = True
        Turno.Numero = rs!turnro
        Turno.Nombre = Trim(rs!turdesabr)
        Turno.Tipo = rs!TipoTurno
        
        T.tiene_turno = True
        T.Nro_Turno = rs!turnro
        T.NombreFPago = Trim(rs!fpgodesabr)
        T.compensa = rs!turcompensa
        T.Fecha_Inicio = DateAdd("d", rs!FechaValidez, -(0 & rs!offset))
        T.Nro_FPago = rs!fpgonro
        
        Empleado.NombreGrupo = Trim(rs!estrdabr)
        Empleado.Grupo = rs!estrnro
    Else
        'Buscar el Turno Actual del empleado */
    End If
    
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 5) & "Fin Buscar_Turno_Nuevo()"
        Flog.writeline
    End If
    
    'cierro y libero
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_FT.State = adStateOpen Then rs_FT.Close
    Set rs_FT = Nothing
End Sub


Public Sub Buscar_Dia_Nuevo(ByVal Fecha As Date, ByVal Fecha_Inicio As Date, ByVal Nro_Turno As Long, ByVal Ternro As Long, ByVal P_Asignacion As Boolean, ByVal depurar As Boolean, ByRef Dia As TDia)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el dia del turno del empleado en la fecha.
' Autor      : FGZ
' Fecha      : 03/11/2004
' Ultima Mod.: 03/11/2004
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim num_dia As Integer
Dim dif_dias As Integer
Dim Firmado As Boolean

Dim rs As New ADODB.Recordset
Dim rs_Firma As New ADODB.Recordset
Dim rs_F As New ADODB.Recordset
Dim rs_FT As New ADODB.Recordset


'    If depurar Then
'        Flog.writeline
'        Flog.writeline Espacios(Tabulador * 5) & "Inicio Buscar_Dia_Nuevo()"
'        Flog.writeline
'    End If
    
    Dia.Trabaja = False
    Dia.Orden_Dia = 0
    Dia.Nro_Dia = 0
    Dia.Dia_Libre = False
    Dia.Nro_Subturno = 0
    Dia.NombreSubTurno = ""
    
    PFecha = Fecha

    StrSql = "SELECT * FROM gti_turno WHERE turnro = " & Nro_Turno
    OpenRecordset StrSql, rs
    
    Dia.Trabaja = False
    'Ordendia = -1 '?
    'Nro_Dia = -1 '?
    dif_dias = DateDiff("d", Fecha_Inicio, Fecha) + 1
    num_dia = dif_dias Mod rs!turtamanio
    If (num_dia = 0) Then num_dia = rs!turtamanio  'es el primer dia del turno
    
   
    ' Buscar el dia Correspondiente
    StrSql = "SELECT gti_dias.dianro,gti_dias.subturnro,gti_dias.diaorden,gti_dias.Dialibre,gti_subturno.subturdesabr, gti_subturno.subtgen FROM gti_subturno INNER JOIN gti_dias ON (gti_subturno.subturnro = gti_dias.subturnro) WHERE " & _
             " (turnro = " & Nro_Turno & ") AND (gti_dias.diaorden <= " & num_dia & ") ORDER BY diaorden DESC "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Dia.Trabaja = True
        Dia.Orden_Dia = rs!diaorden
        Dia.Nro_Dia = rs!dianro
        Dia.Nro_Subturno = rs!subturnro
        Dia.NombreSubTurno = rs!subturdesabr
        Dia.Dia_Libre = rs!Dialibre
        'FGZ - 31/07/2009 - le agregue ese campo
        Dia.Genera = IIf(EsNulo(rs!subtgen), 0, rs!subtgen) 'rs!subtgen
        'FGZ - 31/07/2009 - le agregue ese campo
    End If



    'FGZ - 02/06/2010 ---------------------------------------------------------------------------------
    'If P_Asignacion Then
    '    StrSql = "SELECT ttemplibre FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND " & _
    '             " (gttempdesde <= " & ConvFecha(Fecha) & ") AND " & _
    '             " (" & ConvFecha(Fecha) & " <= gttemphasta)"
    '    OpenRecordset StrSql, rs
    '    If Not rs.EOF Then Dia.Dia_Libre = rs!ttemplibre
    'End If
    
    If P_Asignacion Then
        StrSql = "SELECT gcpnro, ttemplibre FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND " & _
                 " (gttempdesde <= " & ConvFecha(Fecha) & ") AND " & _
                 " (" & ConvFecha(Fecha) & " <= gttemphasta)"
        OpenRecordset StrSql, rs
        Firmado = False
        Do While Not rs.EOF And Not Firmado
            'FGZ - 31/05/2010  --------------------------------------------------------------------------
            'Verifico que no haya sido generado fuera de termino y en ese caso reviso que esté aprobado
            StrSql = "SELECT input_ft.idnro,input_ft.origen, gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
            StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
            StrSql = StrSql & " WHERE idtipoinput = 6 "
            StrSql = StrSql & " AND origen = " & rs!gcpnro
            OpenRecordset StrSql, rs_FT
            If Not rs_FT.EOF Then
                'El parte fué cargado fuera de termimo
                If rs_FT!ftap = -1 Then
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria fuera de termino aprobado."
                    End If
                    Firmado = True
                    Call InsertarFT(rs_FT!idnro, 6, rs_FT!Origen)
                Else
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria fuera de termino NO aprobado. Se descarta."
                    End If
                    Firmado = False
                End If
            Else
                'Verificar si esta en el NIVEL FINAL DE FIRMA ACTIVO para partes de asignacion Horaria
                StrSql = "select * from cystipo where cystipnro = 17"
                OpenRecordset StrSql, rs_Firma
                If Not rs_Firma.EOF Then
                    If rs_Firma!cystipact = -1 Then
                        StrSql = "SELECT * FROM cysfirmas "
                        StrSql = StrSql & " WHERE cysfirfin = -1"
                        StrSql = StrSql & " AND cysfircodext = '" & rs!gcpnro & "' "
                        StrSql = StrSql & " AND cystipnro = 17"
                        OpenRecordset StrSql, rs_F
                        If rs_F.EOF Then
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
                Dia.Dia_Libre = rs!ttemplibre
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria sin fin de firma. Se descartó"
                End If
            End If
            rs.MoveNext
        Loop
    End If

    If depurar Then
        GeneraTraza Ternro, PFecha, "Código del subturno", Str(Dia.Nro_Subturno)
        GeneraTraza Ternro, PFecha, "Trabaja", Str(Dia.Trabaja)
        GeneraTraza Ternro, PFecha, "Orden del día", Str(Dia.Orden_Dia)
        GeneraTraza Ternro, PFecha, "Día libre (Franco)", Str(Dia.Dia_Libre)
    End If
   
'Libero
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_Firma.State = adStateOpen Then rs_Firma.Close
    Set rs_Firma = Nothing
    If rs_F.State = adStateOpen Then rs_F.Close
    Set rs_F = Nothing
    If rs_FT.State = adStateOpen Then rs_FT.Close
    Set rs_FT = Nothing
End Sub



Public Function EsFeriado_Nuevo(ByVal Dia As Date, ByVal Ternro As Long, ByVal depurar As Boolean) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Determina si ese dia es feiado para el empleado en el turno.
' Autor      : FGZ
' Fecha      : 03/11/2004
' Ultima Mod.: 03/11/2004
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Feriado As New ADODB.Recordset
Dim rs As New ADODB.Recordset

'    If depurar Then
'        Flog.writeline
'        Flog.writeline Espacios(Tabulador * 5) & "Inicio EsFeriado_Nuevo()"
'        Flog.writeline
'    End If
    EsFeriado_Nuevo = False

    'Determino el nro de la estructura
    StrSql = " SELECT estructura.estrnro,estructura.estrdabr FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
    StrSql = StrSql & " WHERE Alcance_Testr.tanro = " & lngAlcanEstr & " And " & _
             " his_estructura.Ternro = " & Ternro & " AND htetdesde <= " & ConvFecha(Dia) & _
             " AND (htethasta >= " & ConvFecha(Dia) & " Or htethasta Is Null )" & _
             " ORDER BY htetdesde DESC "
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        NroEstr = rs!estrnro
        NombreEstr = Trim(" " & rs!estrdabr)
    End If
    
    StrSql = "SELECT ferinro, tipferinro,fericodext FROM Feriado WHERE ferifecha = " & ConvFecha(Dia)
    OpenRecordset StrSql, rs_Feriado
    
    If Not rs_Feriado.EOF Then
        If rs_Feriado!tipferinro = 1 Then
            'Pais
            StrSql = "SELECT paisnro FROM Pais WHERE paisdef = " & blnTRUE
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                If Val(rs_Feriado!fericodext) = rs!paisnro Then EsFeriado_Nuevo = True
            End If
        Else
            StrSql = "SELECT * FROM Fer_estr WHERE estrnro = " & NroEstr & " AND ferinro = " & rs_Feriado!ferinro
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then EsFeriado_Nuevo = True
        End If
        'If depurar Then GeneraTraza Ternro, Dia, "Es Feriado", Str(Feriado)
    End If
    
'    If depurar Then
'        Flog.writeline
'        Flog.writeline Espacios(Tabulador * 5) & "Fin EsFeriado_Nuevo()"
'        Flog.writeline
'    End If
'cierro y libero
If rs.State = adStateOpen Then rs.Close
If rs_Feriado.State = adStateOpen Then rs_Feriado.Close

Set rs = Nothing
Set rs_Feriado = Nothing
End Function



Public Sub buscar_horas_turno(ByRef tdias_oblig As Single, ByRef max_horas As Single, ByRef horas_min As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Determina la cantidad de horas del dia para el empleado en el turno.
' Autor      :
' Fecha      :
' Ultima Mod.: FGZ - 01/06/2007
' Descripcion: select * por los campos necesarios
' ---------------------------------------------------------------------------------------------
Dim a As String
Dim objRs As New ADODB.Recordset

    If P_Asignacion Then
        StrSql = "SELECT diacanthoras,diamaxhoras,diaminhoras FROM gti_detturtemp WHERE (ternro =" & Empleado.Ternro & ") AND " & _
                 "(gttempdesde <= " & ConvFecha(p_fecha) & ") AND " & _
                 "(" & ConvFecha(p_fecha) & " <= gttemphasta)"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
              tdias_oblig = objRs!diacanthoras
              max_horas = objRs!diamaxhoras
              horas_min = objRs!diaminhoras
        End If
    Else
        StrSql = "SELECT diacanthoras,diamaxhoras,diaminhoras FROM gti_dias WHERE dianro = " & Nro_Dia
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
               tdias_oblig = objRs!diacanthoras
               max_horas = objRs!diamaxhoras
               horas_min = objRs!diaminhoras
         End If
    End If
    GeneraTraza Empleado.Ternro, p_fecha, "Horas del Turno"
    GeneraTraza Empleado.Ternro, p_fecha, "Cantidad de Horas Mínimas del Turno", Str(horas_min)
    GeneraTraza Empleado.Ternro, p_fecha, "Cantidad de Horas Máximas del Turno", Str(max_horas)
    GeneraTraza Empleado.Ternro, p_fecha, "Cantidad de Horas Obligatorias del Turno", Str(tdias_oblig)

If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing
End Sub


Public Sub Cambiar_Horas(ByVal nro_desg As Long, ByRef hora_desde_desg As String, ByRef hora_hasta_desg As String, ByRef fecha_desde_desg As Date, fecha_hasta_desg As Date)

Dim hora_entrada1 As String
Dim hora_entrada2 As String
Dim hora_entrada3 As String
Dim hora_salida1 As String
Dim hora_salida2 As String
Dim hora_salida3 As String

Dim horres As String
Dim fecres As Date
Dim ok As Boolean
Dim objrsDesgl As New ADODB.Recordset
Dim objRs As New ADODB.Recordset

'def buffer buf_desghora for gti_desghora
'def buffer b_turno for gti_turno.
'def buffer b_subturno for gti_subturno.
'def buffer b_dias for gti_dias.

   If P_Asignacion Then
        StrSql = "SELECT ttemphdesde1,ttemphdesde2,ttemphdesde3,ttemphhasta1,ttemphhasta2,ttemphhasta3 FROM gti_detturtemp WHERE (ternro =" & Empleado.Ternro & " ) and (" & _
                 "gttempdesde <= " & ConvFecha(p_fecha) & ") and (" & _
                 ConvFecha(p_fecha) & " <= gttemphasta)"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
             hora_entrada1 = "" & objRs!ttemphdesde1
             hora_entrada2 = "" & objRs!ttemphdesde2
             hora_entrada3 = "" & objRs!ttemphdesde3
             hora_salida1 = "" & objRs!ttemphhasta1
             hora_salida2 = "" & objRs!ttemphhasta2
             hora_salida3 = "" & objRs!ttemphhasta3
         End If
    Else
        StrSql = "SELECT diahoradesde1,diahoradesde2,diahoradesde3,diahorahasta1,diahorahasta2,diahorahasta3 FROM gti_dias where dianro =" & Nro_Dia
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            hora_entrada1 = objRs!diahoradesde1
            hora_entrada2 = objRs!diahoradesde2
            hora_entrada3 = objRs!diahoradesde3
            hora_salida1 = objRs!diahorahasta1
            hora_salida2 = objRs!diahorahasta2
            hora_salida3 = objRs!diahorahasta3
        End If
    End If


'    ***** Cambio segun lo definido en el desglose de horas ******
    StrSql = "SELECT * FROM  gti_desghora WHERE desghnro = " & nro_desg
    OpenRecordset StrSql, objrsDesgl
    If Not objrsDesgl.EOF Then
        Select Case objrsDesgl!desghorfijo
            Case 1 ' Horario Fijo
                hora_desde_desg = objrsDesgl!desghoradesde
                hora_hasta_desg = objrsDesgl!desghorahasta
                
                fecha_desde_desg = p_fecha
                fecha_hasta_desg = p_fecha
                
            Case 2 ' Horario Variable

                Select Case objrsDesgl!desde_entrada
                    Case 1
                       If objrsDesgl!desdehorentant <> "" Then
                            objFechasHoras.RestaXHoras p_fecha, hora_entrada1, objrsDesgl!desdehorentant, fecres, horres
                            ok = objFechasHoras.ValidarHora(horres)
                            If Not ok Then Exit Sub
                            hora_desde_desg = horres
                            fecha_desde_desg = fecres
                            If objrsDesgl!hastahorentant <> "" Then
                                 objFechasHoras.RestaXHoras p_fecha, hora_entrada1, objrsDesgl!hastahorentant, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_hasta_desg = horres
                                 fecha_hasta_desg = fecres
                             End If
                             If objrsDesgl!hastahorentpos <> "" Then
                                 objFechasHoras.SumoHoras p_fecha, hora_entrada1, objrsDesgl!hastahorentpos, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_hasta_desg = horres
                                 fecha_hasta_desg = fecres
                             End If
                             If objrsDesgl!hastahorsalant <> "" Then
                                 objFechasHoras.RestaXHoras p_fecha, hora_salida1, objrsDesgl!hastahorsalant, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_hasta_desg = horres
                                 fecha_hasta_desg = fecres
                             End If
                             If objrsDesgl!hastahorsalpos <> "" Then
                                 objFechasHoras.SumoHoras p_fecha, hora_salida1, objrsDesgl!hastahorsalpos, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_hasta_desg = horres
                                 fecha_hasta_desg = fecres
                             End If
                         End If
                       
                    If objrsDesgl!desdehorentpos <> "" Then
                        objFechasHoras.SumoHoras p_fecha, hora_entrada1, objrsDesgl!desdehorentpos, fecres, horres
                        ok = objFechasHoras.ValidarHora(horres)
                        If Not ok Then Exit Sub
                        hora_desde_desg = horres
                        fecha_desde_desg = fecres
                        If objrsDesgl!hastahorentpos <> "" Then
                             objFechasHoras.SumoHoras p_fecha, hora_entrada1, objrsDesgl!hastahorentpos, fecres, horres
                             ok = objFechasHoras.ValidarHora(horres)
                             If Not ok Then Exit Sub
                             hora_hasta_desg = horres
                             fecha_hasta_desg = fecres
                        End If
                        If objrsDesgl!hastahorsalant <> "" Then
                             objFechasHoras.RestaXHoras p_fecha, hora_salida1, objrsDesgl!hastahorsalpos, fecres, horres
                             ok = objFechasHoras.ValidarHora(horres)
                             If Not ok Then Exit Sub
                             hora_hasta_desg = horres
                             fecha_hasta_desg = fecres
                        End If
                        If objrsDesgl!hastahorsalpos <> "" Then
                             objFechasHoras.SumoHoras p_fecha, hora_salida1, objrsDesgl!hastahorsalpos, fecres, horres
                             ok = objFechasHoras.ValidarHora(horres)
                             If Not ok Then Exit Sub
                             hora_hasta_desg = horres
                             fecha_hasta_desg = fecres
                         End If
                    End If
                    If objrsDesgl!desdehorsalant <> "" Then
                         objFechasHoras.RestaXHoras p_fecha, hora_salida1, objrsDesgl!desdehorsalant, fecres, horres
                         ok = objFechasHoras.ValidarHora(horres)
                         If Not ok Then Exit Sub
                         hora_desde_desg = horres
                         fecha_desde_desg = fecres
                         If objrsDesgl!hastahorsalant <> "" Then
                             objFechasHoras.RestaXHoras p_fecha, hora_salida1, objrsDesgl!hastahorsalant, fecres, horres
                             ok = objFechasHoras.ValidarHora(horres)
                             If Not ok Then Exit Sub
                             hora_hasta_desg = horres
                             fecha_hasta_desg = fecres
                         End If
                         If objrsDesgl!hastahorsalpos <> "" Then
                            objFechasHoras.SumoHoras p_fecha, hora_salida1, objrsDesgl!hastahorsalpos, fecres, horres
                            ok = objFechasHoras.ValidarHora(horres)
                            If Not ok Then Exit Sub
                            hora_hasta_desg = horres
                            fecha_hasta_desg = fecres
                         End If
                    End If
                    If objrsDesgl!desdehorsalpos <> "" Then
                       objFechasHoras.SumoHoras p_fecha, hora_salida1, objrsDesgl!desdehorsalpos, fecres, horres
                       ok = objFechasHoras.ValidarHora(horres)
                       If Not ok Then Exit Sub
                       hora_desde_desg = horres
                       fecha_desde_desg = fecres
                       If objrsDesgl!hastahorsalpos <> "" Then
                          objFechasHoras.SumoHoras p_fecha, hora_salida1, objrsDesgl!hastahorsalpos, fecres, horres
                          ok = objFechasHoras.ValidarHora(horres)
                          If Not ok Then Exit Sub
                          hora_hasta_desg = horres
                          fecha_hasta_desg = fecres
                       End If
                    End If
                Case 2 'Segunda Hora de entrada/salida
                  If objrsDesgl!desdehorentant <> "" Then
                     objFechasHoras.RestaXHoras p_fecha, hora_entrada2, objrsDesgl!desdehorentant, fecres, horres
                     ok = objFechasHoras.ValidarHora(horres)
                     If Not ok Then Exit Sub
                     hora_desde_desg = horres
                     fecha_desde_desg = fecres
                     If objrsDesgl!hastahorentant <> "" Then
                        objFechasHoras.RestaXHoras p_fecha, hora_entrada2, objrsDesgl!hastahorentant, fecres, horres
                        ok = objFechasHoras.ValidarHora(horres)
                        If Not ok Then Exit Sub
                        hora_hasta_desg = horres
                        fecha_hasta_desg = fecres
                     End If
                     If objrsDesgl!hastahorentpos <> "" Then
                        objFechasHoras.SumoHoras p_fecha, hora_entrada2, objrsDesgl!hastahorentpos, fecres, horres
                        ok = objFechasHoras.ValidarHora(horres)
                        If Not ok Then Exit Sub
                        hora_hasta_desg = horres
                        fecha_hasta_desg = fecres
                     End If
                     If objrsDesgl!hastahorsalant <> "" Then
                        objFechasHoras.RestaXHoras p_fecha, hora_salida2, objrsDesgl!hastahorsalant, fecres, horres
                        ok = objFechasHoras.ValidarHora(horres)
                        If Not ok Then Exit Sub
                        hora_hasta_desg = horres
                        fecha_hasta_desg = fecres
                     End If
                     If objrsDesgl!hastahorsalpos <> "" Then
                        objFechasHoras.SumoHoras p_fecha, hora_salida2, objrsDesgl!hastahorsalpos, fecres, horres
                        ok = objFechasHoras.ValidarHora(horres)
                        If Not ok Then Exit Sub
                        hora_hasta_desg = horres
                        fecha_hasta_desg = fecres
                     End If
                  End If
                  If objrsDesgl!desdehorentpos <> "" Then
                     objFechasHoras.SumoHoras p_fecha, hora_entrada2, objrsDesgl!desdehorentpos, fecres, horres
                     ok = objFechasHoras.ValidarHora(horres)
                     If Not ok Then Exit Sub
                     hora_desde_desg = horres
                     fecha_desde_desg = fecres
                     If objrsDesgl!hastahorentpos <> "" Then
                        objFechasHoras.SumoHoras p_fecha, hora_entrada2, objrsDesgl!hastahorentpos, fecres, horres
                        ok = objFechasHoras.ValidarHora(horres)
                        If Not ok Then Exit Sub
                        hora_hasta_desg = horres
                        fecha_hasta_desg = fecres
                     End If
                     If objrsDesgl!hastahorsalant <> "" Then
                        objFechasHoras.RestaXHoras p_fecha, hora_salida2, objrsDesgl!hastahorsalpos, fecres, horres
                        ok = objFechasHoras.ValidarHora(horres)
                        If Not ok Then Exit Sub
                        hora_hasta_desg = horres
                        fecha_hasta_desg = fecres
                     End If
                     If objrsDesgl!hastahorsalpos <> "" Then
                        objFechasHoras.SumoHoras p_fecha, hora_salida2, objrsDesgl!hastahorsalpos, fecres, horres
                        ok = objFechasHoras.ValidarHora(horres)
                        If Not ok Then Exit Sub
                        hora_hasta_desg = horres
                        fecha_hasta_desg = fecres
                     End If
                 End If
                 If objrsDesgl!desdehorsalant <> "" Then
                     objFechasHoras.RestaXHoras p_fecha, hora_salida2, objrsDesgl!desdehorsalant, fecres, horres
                     ok = objFechasHoras.ValidarHora(horres)
                     If Not ok Then Exit Sub
                     hora_desde_desg = horres
                     fecha_desde_desg = fecres
                     If objrsDesgl!hastahorsalant <> "" Then
                         objFechasHoras.RestaXHoras p_fecha, hora_salida2, objrsDesgl!hastahorsalant, fecres, horres
                         ok = objFechasHoras.ValidarHora(horres)
                         If Not ok Then Exit Sub
                         hora_hasta_desg = horres
                         fecha_hasta_desg = fecres
                     End If
                     If objrsDesgl!hastahorsalpos <> "" Then
                         objFechasHoras.SumoHoras p_fecha, hora_salida2, objrsDesgl!hastahorsalpos, fecres, horres
                         ok = objFechasHoras.ValidarHora(horres)
                         If Not ok Then Exit Sub
                         hora_hasta_desg = horres
                         fecha_hasta_desg = fecres
                     End If
                 End If
                 If objrsDesgl!desdehorsalpos <> "" Then
                     objFechasHoras.SumoHoras p_fecha, hora_salida2, objrsDesgl!desdehorsalpos, fecres, horres
                     ok = objFechasHoras.ValidarHora(horres)
                     If Not ok Then Exit Sub
                     hora_desde_desg = horres
                     fecha_desde_desg = fecres
                     If objrsDesgl!hastahorsalpos <> "" Then
                         objFechasHoras.SumoHoras p_fecha, hora_salida2, objrsDesgl!hastahorsalpos, fecres, horres
                         ok = objFechasHoras.ValidarHora(horres)
                         If Not ok Then Exit Sub
                         hora_hasta_desg = horres
                         fecha_hasta_desg = fecres
                     End If
                 End If
             Case 3   ' Tercera Hora de entrada/salida
                 If objrsDesgl!desdehorentant <> "" Then
                       objFechasHoras.RestaXHoras p_fecha, hora_entrada3, objrsDesgl!desdehorentant, fecres, horres
                       ok = objFechasHoras.ValidarHora(horres)
                       If Not ok Then Exit Sub
                       hora_desde_desg = horres
                       fecha_desde_desg = fecres
                       If objrsDesgl!hastahorentant <> "" Then
                           objFechasHoras.RestaXHoras p_fecha, hora_entrada3, objrsDesgl!hastahorentant, fecres, horres
                           ok = objFechasHoras.ValidarHora(horres)
                           If Not ok Then Exit Sub
                           hora_hasta_desg = horres
                           fecha_hasta_desg = fecres
                        End If
                        If objrsDesgl!hastahorentpos <> "" Then
                           objFechasHoras.SumoHoras p_fecha, hora_entrada3, objrsDesgl!hastahorentpos, fecres, horres
                           ok = objFechasHoras.ValidarHora(horres)
                           If Not ok Then Exit Sub
                           hora_hasta_desg = horres
                           fecha_hasta_desg = fecres
                        End If
                        If objrsDesgl!hastahorsalant <> "" Then
                           objFechasHoras.RestaXHoras p_fecha, hora_salida3, objrsDesgl!hastahorsalant, fecres, horres
                           ok = objFechasHoras.ValidarHora(horres)
                           If Not ok Then Exit Sub
                           hora_hasta_desg = horres
                           fecha_hasta_desg = fecres
                        End If
                        If objrsDesgl!hastahorsalpos <> "" Then
                           objFechasHoras.SumoHoras p_fecha, hora_salida3, objrsDesgl!hastahorsalpos, fecres, horres
                           ok = objFechasHoras.ValidarHora(horres)
                           If Not ok Then Exit Sub
                           hora_hasta_desg = horres
                           fecha_hasta_desg = fecres
                        End If
                 End If
                 If objrsDesgl!desdehorentpos <> "" Then
                       objFechasHoras.SumoHoras p_fecha, hora_entrada3, objrsDesgl!desdehorentpos, fecres, horres
                       ok = objFechasHoras.ValidarHora(horres)
                       If Not ok Then Exit Sub
                       hora_desde_desg = horres
                       fecha_desde_desg = fecres
                       If objrsDesgl!hastahorentpos <> "" Then
                           objFechasHoras.SumoHoras p_fecha, hora_entrada3, objrsDesgl!hastahorentpos, fecres, horres
                           ok = objFechasHoras.ValidarHora(horres)
                           If Not ok Then Exit Sub
                           hora_hasta_desg = horres
                           fecha_hasta_desg = fecres
                       End If
                       If objrsDesgl!hastahorsalant <> "" Then
                           objFechasHoras.RestaXHoras p_fecha, hora_salida3, objrsDesgl!hastahorsalpos, fecres, horres
                           ok = objFechasHoras.ValidarHora(horres)
                           If Not ok Then Exit Sub
                           hora_hasta_desg = horres
                           fecha_hasta_desg = fecres
                       End If
                       If objrsDesgl!hastahorsalpos <> "" Then
                           objFechasHoras.SumoHoras p_fecha, hora_salida3, objrsDesgl!hastahorsalpos, fecres, horres
                           ok = objFechasHoras.ValidarHora(horres)
                           If Not ok Then Exit Sub
                           hora_hasta_desg = horres
                           fecha_hasta_desg = fecres
                       End If
                 End If
                 If objrsDesgl!desdehorsalant <> "" Then
                       objFechasHoras.RestaXHoras p_fecha, hora_salida3, objrsDesgl!desdehorsalant, fecres, horres
                       ok = objFechasHoras.ValidarHora(horres)
                       If Not ok Then Exit Sub
                       hora_desde_desg = horres
                       fecha_desde_desg = fecres
                       If objrsDesgl!hastahorsalant <> "" Then
                           objFechasHoras.RestaXHoras p_fecha, hora_salida3, objrsDesgl!hastahorsalant, fecres, horres
                           ok = objFechasHoras.ValidarHora(horres)
                           If Not ok Then Exit Sub
                           hora_hasta_desg = horres
                           fecha_hasta_desg = fecres
                       End If
                       If objrsDesgl!hastahorsalpos <> "" Then
                           objFechasHoras.SumoHoras p_fecha, hora_salida3, objrsDesgl!hastahorsalpos, fecres, horres
                           ok = objFechasHoras.ValidarHora(horres)
                           If Not ok Then Exit Sub
                           hora_hasta_desg = horres
                           fecha_hasta_desg = fecres
                       End If
                 End If
                 If objrsDesgl!desdehorsalpos <> "" Then
                       objFechasHoras.SumoHoras p_fecha, hora_salida3, objrsDesgl!desdehorsalpos, fecres, horres
                       ok = objFechasHoras.ValidarHora(horres)
                       If Not ok Then Exit Sub
                       hora_desde_desg = horres
                       fecha_desde_desg = fecres
                       If objrsDesgl!hastahorsalpos <> "" Then
                          objFechasHoras.SumoHoras p_fecha, hora_salida3, objrsDesgl!hastahorsalpos, fecres, horres
                          ok = objFechasHoras.ValidarHora(horres)
                          If Not ok Then Exit Sub
                          hora_hasta_desg = horres
                          fecha_hasta_desg = fecres
                       End If
                 End If
           End Select
           Case 3   'Fijo/Variable
               If objrsDesgl!desghoradesde <> "" Then
                    hora_desde_desg = objrsDesgl!desghoradesde
                    Select Case objrsDesgl!desde_entrada
                       Case 1  ' Primera hora de entrada salida */
                           If objrsDesgl!hastahorentant <> "" Then
                               objFechasHoras.RestaXHoras p_fecha, hora_entrada1, objrsDesgl!hastahorentant, fecres, horres
                               ok = objFechasHoras.ValidarHora(horres)
                               If Not ok Then Exit Sub
                               hora_hasta_desg = horres
                               fecha_hasta_desg = fecres
                           End If
                           If objrsDesgl!hastahorentpos <> "" Then
                               objFechasHoras.SumoHoras p_fecha, hora_entrada1, objrsDesgl!hastahorentpos, fecres, horres
                               ok = objFechasHoras.ValidarHora(horres)
                               If Not ok Then Exit Sub
                               hora_hasta_desg = horres
                               fecha_hasta_desg = fecres
                           End If
                           If objrsDesgl!hastahorsalant <> "" Then
                               objFechasHoras.RestaXHoras p_fecha, hora_salida1, objrsDesgl!hastahorsalant, fecres, horres
                               ok = objFechasHoras.ValidarHora(horres)
                               If Not ok Then Exit Sub
                               hora_hasta_desg = horres
                               fecha_hasta_desg = fecres
                           End If
                           If objrsDesgl!hastahorsalpos <> "" Then
                               objFechasHoras.SumoHoras p_fecha, hora_salida1, objrsDesgl!hastahorsalpos, fecres, horres
                               ok = objFechasHoras.ValidarHora(horres)
                               If Not ok Then Exit Sub
                               hora_hasta_desg = horres
                               fecha_hasta_desg = fecres
                           End If
                    Case 2  ' Segunda Hora de entrada/salida
                         If objrsDesgl!hastahorentant <> "" Then
                             objFechasHoras.RestaXHoras p_fecha, hora_entrada2, objrsDesgl!hastahorentant, fecres, horres
                             ok = objFechasHoras.ValidarHora(horres)
                             If Not ok Then Exit Sub
                             hora_hasta_desg = horres
                             fecha_hasta_desg = fecres
                          End If
                          If objrsDesgl!hastahorentpos <> "" Then
                              objFechasHoras.SumoHoras p_fecha, hora_entrada2, objrsDesgl!hastahorentpos, fecres, horres
                              ok = objFechasHoras.ValidarHora(horres)
                              If Not ok Then Exit Sub
                              hora_hasta_desg = horres
                              fecha_hasta_desg = fecres
                          End If
                          If objrsDesgl!hastahorsalant <> "" Then
                              objFechasHoras.RestaXHoras p_fecha, hora_salida2, objrsDesgl!hastahorsalant, fecres, horres
                              ok = objFechasHoras.ValidarHora(horres)
                              If Not ok Then Exit Sub
                              hora_hasta_desg = horres
                              fecha_hasta_desg = fecres
                           End If
                           If objrsDesgl!hastahorsalpos <> "" Then
                              objFechasHoras.SumoHoras p_fecha, hora_salida2, objrsDesgl!hastahorsalpos, fecres, horres
                              ok = objFechasHoras.ValidarHora(horres)
                              If Not ok Then Exit Sub
                              hora_hasta_desg = horres
                              fecha_hasta_desg = fecres
                            End If
                    Case 3   ' Tercera Hora de entrada/salida
                         If objrsDesgl!hastahorentant <> "" Then
                             objFechasHoras.RestaXHoras p_fecha, hora_entrada3, objrsDesgl!hastahorentant, fecres, horres
                             ok = objFechasHoras.ValidarHora(horres)
                             If Not ok Then Exit Sub
                             hora_hasta_desg = horres
                             fecha_hasta_desg = fecres
                         End If
                         If objrsDesgl!hastahorentpos <> "" Then
                              objFechasHoras.SumoHoras p_fecha, hora_entrada3, objrsDesgl!hastahorentpos, fecres, horres
                              ok = objFechasHoras.ValidarHora(horres)
                              If Not ok Then Exit Sub
                              hora_hasta_desg = horres
                              fecha_hasta_desg = fecres
                         End If
                         If objrsDesgl!hastahorsalant <> "" Then
                              objFechasHoras.RestaXHoras p_fecha, hora_salida3, objrsDesgl!hastahorsalant, fecres, horres
                              ok = objFechasHoras.ValidarHora(horres)
                              If Not ok Then Exit Sub
                              hora_hasta_desg = horres
                              fecha_hasta_desg = fecres
                         End If
                         If objrsDesgl!hastahorsalpos <> "" Then
                              objFechasHoras.SumoHoras p_fecha, hora_salida3, objrsDesgl!hastahorsalpos, fecres, horres
                              ok = objFechasHoras.ValidarHora(horres)
                              If Not ok Then Exit Sub
                              hora_hasta_desg = horres
                              fecha_hasta_desg = fecres
                         End If
                     End Select
            Else
            
                If objrsDesgl!desghorahasta <> "" Then
                    hora_hasta_desg = objrsDesgl!desghorahasta
                    Select Case objrsDesgl!desde_entrada
                       Case 1  ' Primera hora de entrada salida
                           If objrsDesgl!desdehorentant <> "" Then
                                objFechasHoras.RestaXHoras p_fecha, hora_entrada1, objrsDesgl!desdehorentant, fecres, horres
                                ok = objFechasHoras.ValidarHora(horres)
                                If Not ok Then Exit Sub
                                hora_desde_desg = horres
                                fecha_desde_desg = fecres
                           End If
                           If objrsDesgl!desdehorentpos <> "" Then
                                 objFechasHoras.SumoHoras p_fecha, hora_entrada1, objrsDesgl!desdehorentpos, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_desde_desg = horres
                                 fecha_desde_desg = fecres
                           End If
                           If objrsDesgl!desdehorsalant <> "" Then
                                 objFechasHoras.RestaXHoras p_fecha, hora_salida1, objrsDesgl!desdehorsalant, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_desde_desg = horres
                                 fecha_desde_desg = fecres
                           End If
                           If objrsDesgl!desdehorsalpos <> "" Then
                                 objFechasHoras.SumoHoras p_fecha, hora_salida1, objrsDesgl!desdehorsalpos, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_desde_desg = horres
                                 fecha_desde_desg = fecres
                           End If
                       Case 2  ' Primera hora de entrada salida */
                           If objrsDesgl!desdehorentant <> "" Then
                                 objFechasHoras.RestaXHoras p_fecha, hora_entrada2, objrsDesgl!desdehorentant, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_desde_desg = horres
                                 fecha_desde_desg = fecres
                            End If
                            If objrsDesgl!desdehorentpos <> "" Then
                                objFechasHoras.SumoHoras p_fecha, hora_entrada2, objrsDesgl!desdehorentpos, fecres, horres
                                ok = objFechasHoras.ValidarHora(horres)
                                If Not ok Then Exit Sub
                                hora_desde_desg = horres
                                fecha_desde_desg = fecres
                            End If
                            If objrsDesgl!desdehorsalant <> "" Then
                                objFechasHoras.RestaXHoras p_fecha, hora_salida2, objrsDesgl!desdehorsalant, fecres, horres
                                ok = objFechasHoras.ValidarHora(horres)
                                If Not ok Then Exit Sub
                                hora_desde_desg = horres
                                fecha_desde_desg = fecres
                            End If
                            If objrsDesgl!desdehorsalpos <> "" Then
                                 objFechasHoras.SumoHoras p_fecha, hora_salida2, objrsDesgl!desdehorsalpos, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_desde_desg = horres
                                 fecha_desde_desg = fecres
                            End If
                        Case 3  ' Primera hora de entrada salida
                            If objrsDesgl!desdehorentant <> "" Then
                                 objFechasHoras.RestaXHoras p_fecha, hora_entrada3, objrsDesgl!desdehorentant, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_desde_desg = horres
                                 fecha_desde_desg = fecres
                            End If
                            If objrsDesgl!desdehorentpos <> "" Then
                                 objFechasHoras.SumoHoras p_fecha, hora_entrada3, objrsDesgl!desdehorentpos, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_desde_desg = horres
                                 fecha_desde_desg = fecres
                            End If
                            If objrsDesgl!desdehorsalant <> "" Then
                                 objFechasHoras.RestaXHoras p_fecha, hora_salida3, objrsDesgl!desdehorsalant, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_desde_desg = horres
                                 fecha_desde_desg = fecres
                            End If
                            If objrsDesgl!desdehorsalpos <> "" Then
                                 objFechasHoras.SumoHoras p_fecha, hora_salida3, objrsDesgl!desdehorsalpos, fecres, horres
                                 ok = objFechasHoras.ValidarHora(horres)
                                 If Not ok Then Exit Sub
                                 hora_desde_desg = horres
                                 fecha_desde_desg = fecres
                            End If
                      End Select
                  End If
            End If
        End Select

    End If


End Sub



Public Sub Buscar_HTeorico_Turno_Movil(ByVal NroTer As Long, ByVal Fecha As Date, ByVal hora_desde As String, ByVal fecha_desde As Date, ByVal hora_hasta As String, ByVal fecha_hasta As Date, ByRef Encontro As Boolean)
' --------------------------------------------------------------
' Descripcion: Setea el horario teorico del legajo.
' Autor: FGZ - 24/10/2005
' Ultima modificacion:
' --------------------------------------------------------------
Dim Opcion As Integer
Dim F_Desde As Date
Dim F_Hasta As Date
Dim H_Desde As String
Dim H_Hasta As String
Dim Aux_Hora As String
Dim Aux_Hora_R As String
Dim Aux_Fecha As Date
Dim Horas_Dia As Single

Dim entrada As Boolean

Dim rs As New ADODB.Recordset
Dim rs_Dia As New ADODB.Recordset


    'Call buscar_horas_turno(horas_oblig, max_horas, horas_min)

    'Busco la cantidad de horas para ese dia en ese turno
    StrSql = "SELECT diacanthoras FROM gti_dias WHERE subturnro = " & Nro_Subturno
    StrSql = StrSql & " ORDER BY diaorden"
    OpenRecordset StrSql, rs_Dia
    If Not rs_Dia.EOF Then
        Horas_Dia = rs_Dia!diacanthoras
    End If

    'Busco la primer registracion dentro de la ventana
    StrSql = "SELECT regfecha, reghora, regentsal FROM gti_registracion WHERE ternro= " & NroTer
    StrSql = StrSql & " AND regfecha = " & ConvFecha(Fecha)
    StrSql = StrSql & " AND ( regllamada = 0 OR regllamada is null )"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " AND (gti_registracion.ft = 0 OR (gti_registracion.ft = -1 AND gti_registracion.ftap = -1))"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " ORDER BY ternro ASC, regfecha ASC, reghora ASC"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Aux_Fecha = rs!regfecha
        Aux_Hora = rs!reghora
        
        If Not EsNulo(rs!regentsal) Then
            If UCase(rs!regentsal) = "E" Then
                entrada = True
            Else
                entrada = False
            End If
        Else
            entrada = False
        End If
        
        'Redondeo/Fraccionamiento
        If TipoRedondeo <> 0 Then
            Aux_Hora_R = objFechasHoras.FraccionaHs(Aux_Hora, TipoRedondeo)
            If Not objFechasHoras.ValidarHora(Aux_Hora) Then
                If depurar Then
                    Flog.writeline "Hora no valida. " & Aux_Hora
                End If
                Exit Sub
            End If
        
            'Call objFechasHoras.Redondeo_Horas_Tipo(Aux_Hora, TipoRedondeo, Aux_Hora_R)
        Else
            If depurar Then
                Flog.writeline "No se redondea la registracion."
            End If
        End If
    
        If entrada Then
            F_Desde = Aux_Fecha
            H_Desde = Aux_Hora_R
        Else
            F_Hasta = Aux_Fecha
            H_Hasta = Aux_Hora_R
        End If
                
        If entrada Then
            'Sumo a la hora la cantidad de hs del turno
            Call objFechasHoras.SumoHoras(F_Desde, H_Desde, Horas_Dia, F_Hasta, H_Hasta)
            If Not objFechasHoras.ValidarHora(H_Hasta) Then
                If depurar Then
                    Flog.writeline "Hora no valida. " & H_Hasta
                End If
                Exit Sub
            End If
        Else
            'Resto a la hora la cantidad de hs del turno
            Call objFechasHoras.RestaXHoras(F_Hasta, H_Hasta, Horas_Dia, F_Desde, H_Desde)
            If Not objFechasHoras.ValidarHora(H_Desde) Then
                If depurar Then
                    Flog.writeline "Hora no valida. " & H_Desde
                End If
                Exit Sub
            End If
        End If
        hora_desde = H_Desde
        hora_hasta = H_Hasta
        fecha_desde = F_Desde
        fecha_hasta = F_Hasta
    Else
        Encontro = False
    End If


    'libero
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing

    If rs_Dia.State = adStateOpen Then rs_Dia.Close
    Set rs_Dia = Nothing
End Sub


Public Sub Calcular_HT(ByVal Fecha As Date, ByVal Nro_Dia As Long, ByRef Dia As THT)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de calculo de Horario cumplido.
' Autor      :
' Fecha      :
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
    
    StrSql = "SELECT diahoradesde1,diahoradesde2,diahoradesde3,diahorahasta1,diahorahasta2,diahorahasta3 FROM gti_dias WHERE dianro = " & Nro_Dia
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
    
        If (objRs!diahoradesde1 <> "0000" Or objRs!diahorahasta1 <> "0000") Then
            Dia.E1 = objRs!diahoradesde1
            Dia.S1 = objRs!diahorahasta1
            
            Dia.FE1 = Fecha
            If Dia.S1 < Dia.E1 Then Fecha = DateAdd("d", 1, Fecha)
            Dia.FS1 = Fecha
        Else
            Dia.E1 = ""
            Dia.S1 = ""
        End If
        If (objRs!diahoradesde2 <> "0000" Or objRs!diahorahasta2 <> "0000") Then
            Dia.E2 = objRs!diahoradesde2
            Dia.S2 = objRs!diahorahasta2
            
            If (Dia.E2 < Dia.S1) Then Fecha = DateAdd("d", 1, Fecha)
            Dia.FE2 = Fecha
            If Dia.S2 < Dia.E2 Then Fecha = DateAdd("d", 1, Fecha)
            Dia.FS2 = Fecha
        Else
            Dia.E2 = ""
            Dia.S2 = ""
        End If
        If (objRs!diahoradesde3 <> "0000" Or objRs!diahorahasta3 <> "0000") Then
            Dia.E3 = objRs!diahoradesde3
            Dia.S3 = objRs!diahorahasta3
            
            If Dia.E3 < Dia.S2 Then Fecha = DateAdd("d", 1, Fecha)
            Dia.FE3 = Fecha
            If Dia.S3 < Dia.S3 Then Fecha = DateAdd("d", 1, Fecha)
            Dia.FS3 = Fecha
        Else
            Dia.E3 = ""
            Dia.S3 = ""
        End If
    End If
End Sub


Public Sub Horario_Teorico()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de calculo de Horario cumplido.
' Autor      :
' Fecha      :
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Fecha As Date
    
    Fecha = p_fecha
    
    StrSql = "SELECT diahoradesde1,diahoradesde2, diahoradesde3,diahorahasta1,diahorahasta2,diahorahasta3 FROM gti_dias WHERE dianro = " & Nro_Dia
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
    
        If (objRs!diahoradesde1 <> "0000" Or objRs!diahorahasta1 <> "0000") Then
            E1 = objRs!diahoradesde1
            S1 = objRs!diahorahasta1
            
            FE1 = Fecha
            If S1 < E1 Then Fecha = DateAdd("d", 1, Fecha)
            FS1 = Fecha
        Else
            E1 = ""
            S1 = ""
        End If
        If (objRs!diahoradesde2 <> "0000" Or objRs!diahorahasta2 <> "0000") Then
            E2 = objRs!diahoradesde2
            S2 = objRs!diahorahasta2
            
            If (E2 < S1) Then Fecha = DateAdd("d", 1, Fecha)
            FE2 = Fecha
            If S2 < E2 Then Fecha = DateAdd("d", 1, Fecha)
            FS2 = Fecha
        Else
            E2 = ""
            S2 = ""
        End If
        If (objRs!diahoradesde3 <> "0000" Or objRs!diahorahasta3 <> "0000") Then
            E3 = objRs!diahoradesde3
            S3 = objRs!diahorahasta3
            
            If E3 < S2 Then Fecha = DateAdd("d", 1, Fecha)
            FE3 = Fecha
            If S3 < S3 Then Fecha = DateAdd("d", 1, Fecha)
            FS3 = Fecha
        Else
            E3 = ""
            S3 = ""
        End If
    End If
End Sub



Public Sub BuscarHoras_1erDiaST(ByVal Ternro As Long, ByVal Fecha As Date, ByRef Horas As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Determina la cantidad de horas del 1er dia del subturno en el turno del empleado.
' Autor      : FGZ
' Fecha      : 31/03/2010
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Firmado As Boolean
Dim Turno As Long
Dim Encontro_turno As Boolean
Dim rs_FT As New ADODB.Recordset


Dim rs As New ADODB.Recordset
Dim rs_Firma As New ADODB.Recordset


    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 5) & "Inicio BuscarHoras_1erDiaST()"
        Flog.writeline
    End If
    
    Encontro_turno = False
    
    'Busca los partes de Cambio de Turno
    StrSql = "SELECT gti_turno.turdesabr,gti_turforpago.turnro,gti_turforpago.fpgonro,gti_reldtur.grtddesde, "
    StrSql = StrSql & "gti_reldtur.grtoffset, gti_turno.turcompensa, gti_turno.tipoturno,"
    StrSql = StrSql & " gti_formapago.fpgodesabr, gti_reldtur.gcpnro "
    StrSql = StrSql & " FROM  gti_reldtur "
    StrSql = StrSql & " INNER JOIN gti_turforpago ON gti_reldtur.turnro = gti_turforpago.turfpagnro "
    StrSql = StrSql & " INNER JOIN gti_turno ON gti_turno.turnro=gti_turforpago.turnro "
    StrSql = StrSql & " INNER JOIN gti_formapago ON gti_turforpago.fpgonro = gti_formapago.fpgonro "
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " (ternro = " & Ternro & " ) AND "
    StrSql = StrSql & " (grtddesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(Fecha) & " <= grtdhasta) "
    StrSql = StrSql & " OR (grtdhasta is null) ) "
    OpenRecordset StrSql, rs
    Do While Not rs.EOF And Not Encontro_turno
        'FGZ - 31/05/2010  --------------------------------------------------------------------------
        'Verifico que no haya sido generado fuera de termino y en ese caso reviso que esté aprobado
        StrSql = "SELECT gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
        StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
        StrSql = StrSql & " WHERE idtipoinput = 7 "
        StrSql = StrSql & " AND origen = " & objRs!gcpnro
        OpenRecordset StrSql, rs_FT
        If Not rs_FT.EOF Then
            'El parte fué cargado fuera de termimo
            If rs_FT!ftap = -1 Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de cambio de turno cargado fuera de termino aprobado."
                End If
                Firmado = True
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de cambio de turno fuera de termino NO aprobado. Se descarta."
                End If
                Firmado = False
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
                    StrSql = StrSql & " AND cysfircodext = '" & rs!gcpnro & "' "
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
            Turno = rs!turnro
            Encontro_turno = True
        End If
        
        'Siguiente
        If Not rs.EOF Then
            rs.MoveNext
        End If
    Loop
    
    If Not Encontro_turno Then
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
                 " (his_estructura.ternro = " & Ternro & ") AND " & _
                 " (htetdesde <= " & ConvFecha(Fecha) & ")  AND " & _
                 "((htethasta >= " & ConvFecha(Fecha) & ")" & _
                 " OR (htethasta is null )) AND (fechavalidez <= " & ConvFecha(Fecha) & " ) " & _
                 " ORDER BY Alcance_Testr.alteorden DESC,his_estructura.htetdesde DESC,gti_turfpgogru.FechaValidez Desc "
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Turno = rs!turnro
            Encontro_turno = True
        End If
    End If
    
    'si tiene turno ==> busco el subturno
    If Encontro_turno Then
        'Buscar el dia Correspondiente
        StrSql = "SELECT diacanthoras "
        StrSql = StrSql & " FROM gti_subturno "
        StrSql = StrSql & " INNER JOIN gti_dias ON (gti_subturno.subturnro = gti_dias.subturnro) "
        StrSql = StrSql & " WHERE (turnro = " & Turno & ") AND (gti_dias.diaorden >= 1)"
        StrSql = StrSql & " ORDER BY diaorden "
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Horas = rs!diacanthoras
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * 6) & "No se encontró subturno para el turno " & Turno
            End If
        End If
    End If
    
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 6) & "Horas del 1er dia del subturno = " & Horas
        Flog.writeline Espacios(Tabulador * 5) & "Fin BuscarHoras_1erDiaST()"
        Flog.writeline
    End If
    
'cierro y libero
    If rs.State = adStateOpen Then rs.Close
    If rs_Firma.State = adStateOpen Then rs_Firma.Close
    Set rs = Nothing
    Set rs_Firma = Nothing
End Sub

