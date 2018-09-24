Attribute VB_Name = "mdlANexus"
Option Explicit

'Para Sql server
'Global Const strConexionNexus = "DSN=Nexushr-RHPro;database=nexus;uid=sa;pwd="

'Para Informix
Global Const strConexionNexus = "DSN=Nexushr-RHPro"

Global objConnNexus As New ADODB.Connection

'Variables de ADO
Dim objRs As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim rsAnrcab As New ADODB.Recordset
Dim rsHistliq As New ADODB.Recordset
Dim rsFactor As New ADODB.Recordset
Dim rsHistCon As New ADODB.Recordset
Dim rsEstructura As New ADODB.Recordset
Dim rsRango As New ADODB.Recordset
Dim rsConc As New ADODB.Recordset


Public Sub Main()
Dim fechaDesde As Date
Dim fechaHasta As Date
Dim Fecha As Date
Dim objRs As New ADODB.Recordset
Dim objrsEmpleado As New ADODB.Recordset
Dim NroProceso As Long
Dim strCmdLine  As String
Dim nro_analisis As Long

On Error GoTo ce
       
    
    OpenConnection strconexion, objConn
               
    Call ProcesoMigrador
       
        
    If objConn.State = adStateOpen Then objConn.Close

    Exit Sub
ce:
'    MyRollbackTrans
    If objConn.State = adStateOpen Then objConn.Close
End Sub

Public Sub OpenRecordsetNexus(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    objRs.Open strSQLQuery, objConnNexus, adOpenDynamic, lockType, adCmdText
End Sub


Private Sub ProcesoMigrador()

'Variables locales
Dim Orden As Long


Dim MiConcepto As String

Dim rsNexus As New ADODB.Recordset
Dim rs As New ADODB.Recordset

'Código -------------------------------------------------------------------

'Abro la conexion para Nexus
OpenConnection strConexionNexus, objConnNexus

'Obtengo el último concepto de RHPRO
StrSql = " SELECT concnro FROM concepto ORDER BY concnro DESC"
OpenRecordset StrSql, rs

Orden = rs!concnro + 1

'Obtengo los conceptos de Nexus
StrSql = " SELECT estr_liq,cod_cpto,nombre,descripcion FROM concepto_defin"
OpenRecordsetNexus StrSql, rsNexus


Do While Not rsNexus.EOF

    MiConcepto = Trim(rsNexus!estr_liq) + Trim(rsNexus!cod_cpto)
    
    
    'Voy insertando en la tabla de conceptos de rhpro
    StrSql = "INSERT INTO concepto(conccod" & _
        ",concabr,concext,tconnro,concretro,concvalid" & _
        ",concniv,concimp,concpuente,concorden) VALUES ('" & _
        MiConcepto & "','" & rsNexus!Nombre & "','" & _
        rsNexus!descripcion & "',17,0,-1,0,-1,0," & 100000 & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

    
    'Tengo que cargar en concorden el serial que acabo de insertar
    Orden = getLastIdentity(objConn, "concepto")
    
    StrSql = "UPDATE concepto SET" & _
        " concorden = " & Orden & _
        " WHERE concnro = " & Orden
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rsNexus.MoveNext
    Debug.Print Orden
Loop
    
    
End Sub

