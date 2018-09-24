VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscarDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim blnTrabaja  As Boolean
Dim Ordendia As Integer
'FGZ - 15/05/2007 ---- le saque estas definiciones dado que las var estan definidas global
'Dim Nro_Dia As Long
'Dim Nro_Subturno As Long
'FGZ - 15/05/2007 ---- le saque estas definiciones dado que las var estan definidas global
Dim blnDia_libre As Boolean
Dim NombreSTurno As String
Dim PFecha As Date

'Private Type Templeado
'    Legajo As Long
'    Grupo As Long
'End Type

Dim Empl As Templeado

Property Let Dia_Libre(libre As Boolean)
    blnDia_libre = libre
End Property
Property Get Dia_Libre() As Boolean
    Dia_Libre = blnDia_libre
End Property

Property Let Trabaja(trab As Boolean)
    blnTrabaja = trab
End Property
Property Get Trabaja() As Boolean
    Trabaja = blnTrabaja
End Property

Property Let Orden_Dia(Orden As Integer)
    Ordendia = Orden
End Property
Property Get Orden_Dia() As Integer
    Orden_Dia = Ordendia
End Property

Property Let Numero_Dia(nro As Long)
    Nro_Dia = nro
End Property
Property Get Numero_Dia() As Long
    Numero_Dia = Nro_Dia
End Property
Property Let SubTurno_Numero(nro As Long)
    Nro_Subturno = nro
End Property
Property Get SubTurno_Numero() As Long
    SubTurno_Numero = Nro_Subturno
End Property
Property Let SubTurno_Nombre(N As String)
     NombreSTurno = N
End Property
Property Get SubTurno_Nombre() As String
    SubTurno_Nombre = NombreSTurno
End Property


Property Let Empleado_Grupo(nro As Long)
    Empl.Grupo = nro
End Property
Property Get Empleado_Grupo() As Long
    Empleado_Grupo = Empl.Grupo
End Property

Property Let Empleado_Legajo(nro As Long)
    Empl.Legajo = nro
End Property
Property Get Empleado_Legajo() As Long
    Empleado_Legajo = Empl.Legajo
End Property

Property Set Conexion(ByRef cn As ADODB.Connection)
    Set objConn = cn
End Property
Property Set ConexionTraza(ByRef cn As ADODB.Connection)
    Set CnTraza = cn
End Property

Public Sub Buscar_Dia(Fecha As Date, Fecha_Inicio As Date, Nro_Turno As Long, Ternro As Long, P_Asignacion As Boolean, depurar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el dia del empleado en la fecha.
' Autor      :
' Fecha      :
' Ultima Mod.: FGZ - 07/07/2008
' Descripcion: Se agregaron los chequeosa de circuito de firmas
' ---------------------------------------------------------------------------------------------
Dim num_dia As Integer
Dim dif_dias As Integer
    
Dim rs_FT As New ADODB.Recordset
Dim rs_Firma As New ADODB.Recordset
Dim Firmado As Boolean
Dim rs As New ADODB.Recordset
    
    
'    'FGZ - 01/06/2007 ----- Mejoras -----
'    Cantidad_Dias = Cantidad_Dias + 1
'    'FGZ - 01/06/2007 ----- Mejoras -----
    
    blnTrabaja = False
    Ordendia = 0
    Nro_Dia = 0
    blnDia_libre = False
    Nro_Subturno = 0
    NombreSTurno = ""

    PFecha = Fecha

    StrSql = "SELECT * FROM gti_turno WHERE turnro = " & Nro_Turno
    OpenRecordset StrSql, objRs
    
    blnTrabaja = False
    Ordendia = -1 '?
    Nro_Dia = -1 '?
    dif_dias = DateDiff("d", Fecha_Inicio, Fecha) + 1
    num_dia = dif_dias Mod objRs!turtamanio
    If (num_dia = 0) Then num_dia = objRs!turtamanio  'es el primer dia del turno */
    
   
    ' Buscar el d¡a Correspondiente
    'StrSql = "SELECT * FROM gti_subturno WHERE turnro = " & Nro_Turno & " ORDER BY subturorden"
    'OpenRecordset StrSql, objrsSubTurno
    
    StrSql = "SELECT gti_dias.*,gti_subturno.subturdesabr, gti_subturno.subtgen FROM gti_subturno INNER JOIN gti_dias ON (gti_subturno.subturnro = gti_dias.subturnro) WHERE " & _
             " (turnro = " & Nro_Turno & ") AND (gti_dias.diaorden <= " & num_dia & ") ORDER BY diaorden DESC "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        blnTrabaja = True
        Ordendia = objRs!diaorden
        Nro_Dia = objRs!dianro
        Nro_Subturno = objRs!subturnro
        NombreSTurno = objRs!subturdesabr
        blnDia_libre = objRs!Dialibre
        'FGZ - 31/07/2009 - le agregue ese campo
        Subturno_Genera = IIf(EsNulo(objRs!subtgen), 0, objRs!subtgen)
        'FGZ - 31/07/2009 - le agregue ese campo
    End If


    If P_Asignacion Then
        StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND " & _
                 " (gttempdesde <= " & ConvFecha(Fecha) & ") AND " & _
                 " (" & ConvFecha(Fecha) & " <= gttemphasta)"
        OpenRecordset StrSql, objRs
        Firmado = False
        Do While Not objRs.EOF And Not Firmado
            'FGZ - 31/05/2010  --------------------------------------------------------------------------
            'Verifico que no haya sido generado fuera de termino y en ese caso reviso que esté aprobado
            StrSql = "SELECT input_ft.idnro,input_ft.origen, gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
            StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
            StrSql = StrSql & " WHERE idtipoinput = 6 "
            StrSql = StrSql & " AND origen = " & objRs!gcpnro
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
                        StrSql = StrSql & " AND cysfircodext = '" & objRs!gcpnro & "' "
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
                blnDia_libre = objRs!ttemplibre
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria sin fin de firma. Se descartó"
                End If
            End If
            objRs.MoveNext
        Loop
    End If
        
'        StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND " & _
'                 " (gttempdesde <= " & ConvFecha(Fecha) & ") AND " & _
'                 " (" & ConvFecha(Fecha) & " <= gttemphasta)"
'        OpenRecordset StrSql, objRs
'        If Not objRs.EOF Then
'            'Verificar si esta en el NIVEL FINAL DE FIRMA ACTIVO para partes de cambio de turno
'            StrSql = "select * from cystipo where cystipnro = 17"
'            OpenRecordset StrSql, rs_Firma
'            If Not rs_Firma.EOF Then
'                If rs_Firma!cystipact = -1 Then
'                    StrSql = "SELECT * FROM cysfirmas "
'                    StrSql = StrSql & " WHERE cysfirfin = -1"
'                    StrSql = StrSql & " AND cysfircodext = '" & objRs!gcpnro & "' "
'                    StrSql = StrSql & " AND cystipnro = 17"
'                    OpenRecordset StrSql, rs
'                    If rs.EOF Then
'                        Firmado = False
'                    Else
'                        Firmado = True
'                    End If
'                Else
'                    Firmado = True
'                End If
'            Else
'                Firmado = True
'            End If
'
'            If Firmado Then
'                blnDia_libre = objRs!ttemplibre
'            Else
'                If depurar Then
'                    Flog.writeline Espacios(Tabulador * 6) & "Hay un parte de asignacion horaria sin fin de firma. Se descartó"
'                End If
'            End If
'        End If
    

    If depurar Then
        GeneraTraza Ternro, PFecha, "Código del subturno", Str(Nro_Subturno)
        GeneraTraza Ternro, PFecha, "Trabaja", Str(blnTrabaja)
        GeneraTraza Ternro, PFecha, "Orden del día", Str(Ordendia)
        GeneraTraza Ternro, PFecha, "Día libre (Franco)", Str(blnDia_libre)
    End If
    
    'cierro y libero
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_Firma.State = adStateOpen Then rs_Firma.Close
    Set rs_Firma = Nothing
    If rs_FT.State = adStateOpen Then rs_FT.Close
    Set rs_FT = Nothing
    
End Sub



Private Sub GeneraTraza(Ternro As Long, Fecha As Date, Desc As String, Optional valor As String = "?")

'    StrSql = "INSERT INTO gti_traza(ternro,fecproc,descripcion,valor) VALUES (" & _
'             Ternro & "," & ConvFecha(Fecha) & ",'" & Desc & "','" & valor & "')"
'    CnTraza.Execute StrSql, , adExecuteNoRecords
    
End Sub


