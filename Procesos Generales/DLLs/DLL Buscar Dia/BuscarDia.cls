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
Dim Nro_dia As Long
Dim blnDia_libre As Boolean
Dim Nro_subturno As Long
Dim NombreSTurno As String
Dim PFecha As Date

Private Type Templeado
    Legajo As Long
    Grupo As Long
End Type

Dim Empl As Templeado

Property Let Dia_libre(libre As Boolean)
    blnDia_libre = libre
End Property
Property Get Dia_libre() As Boolean
    Dia_libre = blnDia_libre
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
    Nro_dia = nro
End Property
Property Get Numero_Dia() As Long
    Numero_Dia = Nro_dia
End Property
Property Let SubTurno_Numero(nro As Long)
    Nro_subturno = nro
End Property
Property Get SubTurno_Numero() As Long
    SubTurno_Numero = Nro_subturno
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

Public Sub Buscar_Dia(Fecha As Date, Fecha_inicio As Date, nro_turno As Long, Ternro As Long, P_asignacion As Boolean, depurar As Boolean)
Dim num_dia As Integer
Dim dif_dias As Integer

    blnTrabaja = False
    Ordendia = 0
    Nro_dia = 0
    blnDia_libre = False
    Nro_subturno = 0
    NombreSTurno = ""

    PFecha = Fecha

    StrSql = "SELECT * FROM gti_turno WHERE turnro = " & nro_turno
    OpenRecordset StrSql, objrs
    
    blnTrabaja = False
    Ordendia = -1 '?
    Nro_dia = -1 '?
    dif_dias = DateDiff("d", Fecha_inicio, Fecha) + 1
    num_dia = dif_dias Mod objrs!turtamanio
    If (num_dia = 0) Then num_dia = objrs!turtamanio  'es el primer dia del turno */
    
   
    ' Buscar el d¡a Correspondiente
    'StrSql = "SELECT * FROM gti_subturno WHERE turnro = " & Nro_Turno & " ORDER BY subturorden"
    'OpenRecordset StrSql, objrsSubTurno
    
    StrSql = "SELECT gti_dias.*,gti_subturno.subturdesabr FROM gti_subturno INNER JOIN gti_dias ON (gti_subturno.subturnro = gti_dias.subturnro) WHERE " & _
             " (turnro = " & nro_turno & ") AND (gti_dias.diaorden <= " & num_dia & ") ORDER BY diaorden DESC "
    OpenRecordset StrSql, objrs
    If Not objrs.EOF Then
        blnTrabaja = True
        Ordendia = objrs!diaorden
        Nro_dia = objrs!dianro
        Nro_subturno = objrs!subturnro
        NombreSTurno = objrs!subturdesabr
        blnDia_libre = objrs!Dialibre
    End If


    If P_asignacion Then
        StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND " & _
                 " (gttempdesde <= " & ConvFecha(Fecha) & ") AND " & _
                 " (" & ConvFecha(Fecha) & " <= gttemphasta)"
        OpenRecordset StrSql, objrs
        If Not objrs.EOF Then blnDia_libre = objrs!ttemplibre
    End If



    If depurar Then
        GeneraTraza Ternro, PFecha, "Código del subturno", Str(Nro_subturno)
        GeneraTraza Ternro, PFecha, "Trabaja", Str(blnTrabaja)
        GeneraTraza Ternro, PFecha, "Orden del día", Str(Ordendia)
        GeneraTraza Ternro, PFecha, "Día libre (Franco)", Str(blnDia_libre)
    End If
End Sub



Private Sub GeneraTraza(Ternro As Long, Fecha As Date, Desc As String, Optional valor As String = "?")

    StrSql = "INSERT INTO gti_traza(ternro,fecproc,descripcion,valor) VALUES (" & _
             Ternro & "," & ConvFecha(Fecha) & ",'" & Desc & "','" & valor & "')"
    
    CnTraza.Execute StrSql, , adExecuteNoRecords
    
End Sub



