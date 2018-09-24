VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActualizarJustificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Archivo:         gtijus07.p
' fecha ini.     : 09/01/97
' Nombre progr.  : justificación de una licencia o curso
' Tipo programa  : proceiento
' autor          : Juan Angel Ringhetti
' Sistema:         Personal
' Base de datos  : PER
' Descripcion    : generación de justificaciones para módulo de licencias
':                 y capacitación(curso)
' Historia       : 12/4/97   -   Mauricio  : le agregue el parametro sigla
'                : 06/08/97  -   Maximiliano Breglia : parametro abm y corresp
'                                                      modif segun el parametro.
'                  18/06/99  -   Muzzolón Leandro: Agregué la hora desde y hasta
'                  22/01/02  -   Ezequiel Loustau: Se agregaron unos parametros de entrada por modif.
'                                                  en la tabla de emp_lic. Inicialmente este cambio
'                                                  solamente involucra a las licencias.

Property Set Conexion(ByRef cn As ADODB.Connection)
    Set objConn = cn
End Property



Public Sub Actualizar_justificacion(nroter As Integer, nrotur As Long, tipojus As Long, codext As Long, _
                                    desde As Date, horadesde As String, hasta As Date, horahasta As String, _
                                    compl As Boolean, sigla As String, tipoel As Integer, orden As String, _
                                    maxhoras As String, abm As Integer)

Dim rs As New ADODB.Recordset
Dim borrar As Boolean

    borrar = False
    
    StrSql = "SELECT * FROM gti_justificacion WHERE ternro = " & nroter & _
             " and jussigla = '" & sigla & "' and juscodext = " & codext
    
    rs.Open StrSql, objConn
    
    If abm = 2 Then
    
         If Not rs.EOF Then
         
                StrSql = "DELETE FROM gti_justificacion WHERE ternro = " & nroter & _
                         " and jussigla = '" & sigla & "' and juscodext = " & codext
                
                objConn.Execute StrSql, , adExecuteNoRecords
                
                borrar = True
                
                Exit Sub
         
         Else
            
                borrar = True
                
                Exit Sub
         
         End If
     
    Else
    
        If rs.EOF Then
        
           StrSql = "INSERT INTO gti_justificacion ( jusanterior,juscodext,jusdesde,jusdiacompleto,jushasta,jussigla,jussistema,ternro,tjusnro,turnro,jushoradesde,jushorahasta,juseltipo,juselorden,juselmaxhoras )" & _
                    " VALUES( " & CInt(True) & "," & codext & ",'" & ConvFecha(desde) & "'," & CInt(compl) & ",'" & ConvFecha(hasta) & "','" & sigla & "'," & CInt(True) & "," & nroter & "," & tipojus & "," & nrotur & ",'" & horadesde & "','" & horahasta & "'," & tipoel & "," & orden & "," & maxhoras & ")"
        
           objConn.Execute StrSql, , adExecuteNoRecords
           
        Else
        
            StrSql = "UPDATE gti_justificacion SET jusanterior = " & CInt(True) & " juscodext = " & codext & " jusdesde = '" & ConvFecha(desde) & "' jusdiacompleto = " & CInt(compl) & " jushasta = '" & ConvFecha(hasta) & "' jussigla = '" & sigla & "' jussistema = " & CInt(True) & _
                     " ternro = " & nroter & " tjusnro = " & tipojus & " turnro = " & nrotur & " jushoradesde = & '" & horadesde & "' jushorahasta = '" & horahasta & "' juseltipo = " & tipoel & " juselorden = " & orden & " juselmaxhoras = " & maxhoras
    
        End If
        
    End If


End Sub
