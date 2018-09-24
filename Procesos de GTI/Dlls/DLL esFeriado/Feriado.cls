VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Feriado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim nroConv As Long
Dim NombreConvenio As String
Dim nroSuc As Long
Dim NombreSucursal As String
Dim NroEstr As Long
Dim NombreEstr As String

Property Get Numero_Convenio() As Long
    Numero_Convenio = nroConv
End Property
Property Get Nombre_Convenio() As Long
    Nombre_Convenio = NombreConvenio
End Property
Property Get Numero_Sucursal() As Long
    Numero_Sucursal = nroSuc
End Property
Property Get Nombre_Sucursal() As Long
    Nombre_Sucursal = NombreSucursal
End Property

Property Set Conexion(ByRef cn As ADODB.Connection)
    Set objConn = cn
End Property
Property Set ConexionTraza(ByRef cn As ADODB.Connection)
    Set CnTraza = cn
End Property

Public Function Feriado(Dia As Date, Ternro As Long, depurar As Boolean) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      :
' Fecha      :
' Ultima Mod.:
' Ultima Mod.: FGZ - 09/03/2007 - Estaba calculando mal la variable Feriado_Por_Estructura cuando el feriado es nacional
' Ultima Mod : FGZ - 01/06/2007 - Mejoras ---- * por los campos necesariso
' Ultima Mod : Diego Rosso - 17-09-2007 - Se reestructuro la funcion. Se agrego Inner join con Fer_estr.
' ---------------------------------------------------------------------------------------------
Dim objRsFeriado As New ADODB.Recordset

'    'FGZ - 01/06/2007 ----- Mejoras -----
'    Cantidad_Feriados = Cantidad_Feriados + 1
'    'FGZ - 01/06/2007 ----- Mejoras -----
    
    Feriado = False
    On Error GoTo MError
    
    Feriado_Por_Estructura = False
    Feriado_Laborable = False
    
    'If depurar Then
    '    Flog.writeline ""
    '    Flog.writeline "Es Feriado "
    'End If
    
    'Se fija si el dia esta declarado como Feriado
    StrSql = "SELECT feriado.tipferinro, feriado.fericodext, feriado.ferinro,ferilaborable FROM Feriado "
    StrSql = StrSql & " WHERE ferifecha = " & ConvFecha(Dia)
    OpenRecordset StrSql, objRsFeriado
    If Not objRsFeriado.EOF Then
        If objRsFeriado!tipferinro = 1 Then
            'Pais
            StrSql = "SELECT paisnro FROM Pais WHERE paisdef = " & blnTRUE
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                'FGZ - 09/03/2007 ------------
                'Cambié esta linea
                'If Val(objRsFeriado!fericodext) = objRs!paisnro Then Feriado = True
                'X
                If Val(objRsFeriado!fericodext) = objRs!paisnro Then
                    Feriado = True
                    Feriado_Por_Estructura = False
                End If
                'FGZ - 09/03/2007 ------------
            End If
        Else
            'FGZ - 01/06/2007 - Mejoras ----
            'StrSql = "SELECT * FROM Fer_estr WHERE estrnro = " & nroEstr & " AND ferinro = " & objRsFeriado!ferinro
           ' StrSql = "SELECT estrnro FROM Fer_estr WHERE estrnro = " & NroEstr & " AND ferinro = " & objRsFeriado!ferinro
           ' OpenRecordset StrSql, objRs
             
             'Determino el nro de la estructura
             'FGZ - 27/01/2009 ------- cambié el query
            'StrSql = " SELECT estructura.estrnro, estructura.estrdabr FROM his_estructura "
            'StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
            'StrSql = StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
            ''17-09-2007 Diego Rosso
            'StrSql = StrSql & " INNER JOIN Fer_estr on Fer_estr.estrnro=estructura.estrnro "
            'StrSql = StrSql & " WHERE Alcance_Testr.tanro = " & lngAlcanEstr & " And " & _
            '                  " his_estructura.Ternro = " & Ternro & " AND htetdesde <= " & ConvFecha(Dia) & _
            '                  " AND (htethasta >= " & ConvFecha(Dia) & " Or htethasta Is Null )" & _
            '                  " ORDER BY htetdesde DESC "
            'por esto
            StrSql = " SELECT estructura.estrnro, estructura.estrdabr FROM Fer_estr "
            StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro = Fer_estr.estrnro"
            StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
            StrSql = StrSql & " WHERE ferinro = " & objRsFeriado!ferinro
            StrSql = StrSql & " AND (his_estructura.Ternro = " & Ternro & " AND htetdesde <= " & ConvFecha(Dia)
            StrSql = StrSql & " AND (htethasta >= " & ConvFecha(Dia) & " Or htethasta Is Null ))"
            StrSql = StrSql & " ORDER BY htetdesde DESC "
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                NroEstr = objRs!estrnro
                NombreEstr = Trim(" " & objRs!estrdabr)
                Feriado_Por_Estructura = True
                Feriado = True
            Else
                Feriado = False
            End If
        End If
        
        'FGZ - 10/08/2010 --------
        If Feriado Then
            If Not EsNulo(objRsFeriado!ferilaborable) Then
                If CBool(objRsFeriado!ferilaborable) Then
                    Feriado_Laborable = True
                Else
                    Feriado_Laborable = False
                End If
            Else
                Feriado_Laborable = False
            End If
        Else
            Feriado_Laborable = False
        End If
        
        If depurar Then GeneraTraza Ternro, Dia, "Es Feriado", Str(Feriado)
    End If
    
                    
    If depurar Then
        If Feriado Then
            If Feriado_Por_Estructura Then
                Flog.writeline Espacios(Tabulador * (Nivel_Tab_Log + 1)) & "Es Feriado por estructura" & IIf(Feriado_Laborable, " Laborable ", " NO Laborable ")
            Else
                Flog.writeline Espacios(Tabulador * (Nivel_Tab_Log + 1)) & "Es Feriado Nacional" & IIf(Feriado_Laborable, " Laborable ", " NO Laborable ")
            End If
        End If
    End If
    Exit Function

MError:
    Flog.writeline
    Flog.writeline "Error procesando Empleado:" & Ternro & " " & Dia
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl: " & StrSql
End Function

Private Sub GeneraTraza(Ternro As Long, Fecha As Date, Desc As String, Optional valor As String = "?")

'    StrSql = "INSERT INTO gti_traza(ternro,fecproc,descripcion,valor) VALUES (" & _
'             Ternro & "," & ConvFecha(Fecha) & ",'" & Desc & "','" & valor & "')"
'    CnTraza.Execute StrSql, , adExecuteNoRecords
    
End Sub

