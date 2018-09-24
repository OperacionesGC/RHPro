Attribute VB_Name = "mdlFases"
Option Explicit

  
Public Sub Main()

Dim fechaDesde As Date

Dim strcmdLine As String

Dim rs_his As New ADODB.Recordset
Dim rs_fases As New ADODB.Recordset
Dim fechaHasta As Date
Dim FechaBaja As String
Dim EstructuraActual As Integer
Dim EmpleadoActual As Integer

' carga las configuraciones basicas, formato de fecha, string de conexion,
' tipo de BD y ubicacion del archivo de log
Call CargarConfiguracionesBasicas

       
OpenConnection strconexion, objConn

'   TE: Contrato actual
'   contrato de tiempo determinado
'   ordenados por tercero, esctructura y fecha descendente
StrSql = "SELECT * "
StrSql = StrSql & " FROM  his_estructura"
StrSql = StrSql & " INNER JOIN replica_estr on his_estructura.estrnro=replica_estr.estrnro "
StrSql = StrSql & " INNER JOIN tipocont on tipocont.tcnro=replica_estr.origen AND tcind=-1"
StrSql = StrSql & " WHERE tenro = 18"
StrSql = StrSql & " ORDER BY ternro,htetdesde DESC"

OpenRecordset StrSql, rs_his

EmpleadoActual = 0
EstructuraActual = 0
' Recorro el histórico de estructuras para los contratos de tiempo determinado
Do While Not rs_his.EOF

    'Para cada empleado proceso sólo el último contrato
    'Imito el comportamiento de un next de progress
    If (EmpleadoActual <> rs_his!Ternro) Then
        
        EmpleadoActual = rs_his!Ternro

        'Si la his_estructura tiene fecha hasta
        If Len(rs_his("htethasta")) > 0 Then
            'Si la fecha hasta es mayor a la fecha de hoy limpio
            If DateDiff("d", rs_his("htethasta"), Date) < 0 Then
                fechaHasta = rs_his("htethasta")
            
                StrSql = "UPDATE his_estructura SET htethasta=null"
                StrSql = StrSql & " WHERE estrnro =" & rs_his("estrnro")
                StrSql = StrSql & " AND ternro=" & rs_his("ternro")
                StrSql = StrSql & " AND htetdesde=" & rs_his("htetdesde")
                StrSql = StrSql & " AND htethasta=" & rs_his("htethasta")
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
        
        StrSql = "SELECT *  "
        StrSql = StrSql & " FROM fases"
        StrSql = StrSql & " WHERE empleado = " & rs_his("ternro")
        StrSql = StrSql & " ORDER BY altfec DESC"
        OpenRecordset StrSql, rs_fases
        
        'Me interesa solo la última fase
        If Not rs_fases.EOF Then
            If Len(rs_fases("bajfec")) > 0 Then
                'Si la fecha de baja es mayor que la fecha de hoy limpio
                If DateDiff("d", rs_fases("bajfec"), Date) < 0 Then
                    FechaBaja = rs_fases("bajfec")
                    
                    StrSql = "UPDATE fases SET bajfec=null, caunro=null "
                    StrSql = StrSql & " WHERE fasnro = " & rs_fases("fasnro")
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
        
        If (IsNull(rs_fases("bajfec"))) And (IsNull(rs_his("htethasta"))) Then
            'sdgsdg
        Else
            If (Not rs_fases.EOF) And (Not IsNull(rs_fases("bajfec"))) Then
                StrSql = "UPDATE empleado SET empfbajaprev=" & ConvFecha(rs_fases("bajfec"))
                StrSql = StrSql & " WHERE ternro = " & rs_his("ternro")
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                If (Not rs_his.EOF) And (Not IsNull(rs_his("htethasta"))) Then
                    StrSql = "UPDATE empleado SET empfbajaprev=" & ConvFecha(rs_his("htethasta"))
                    StrSql = StrSql & " WHERE ternro = " & rs_his("ternro")
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
    End If
    
    rs_his.MoveNext

Loop

rs_fases.Close
Set rs_fases = Nothing
rs_his.Close
Set rs_his = Nothing
        
       
        

Exit Sub


End Sub
    
    

