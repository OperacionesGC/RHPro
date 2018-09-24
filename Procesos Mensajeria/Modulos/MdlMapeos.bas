Attribute VB_Name = "MdlMapeos"
Option Explicit

'----------------------------------------------------------------
'Busca cual es el mapeo de un codigo Heidt a un codigo SAP
Public Function CalcularMapeo(ByVal Parametro, ByVal Tabla, ByVal Default)

    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim correcto As Boolean
    Dim Salida
    
    If IsNull(Parametro) Then
       correcto = False
    Else
       correcto = Parametro <> ""
    End If
           
    Salida = Default

    If correcto Then
        StrSql = " SELECT * FROM infotipos_mapeo " & _
                 " WHERE tablaref = '" & Tabla & "' " & _
                 "   AND codinterno = '" & Parametro & "' "
        
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Salida = CStr(IIf(Not IsNull(rs_Consult!codexterno), rs_Consult!codexterno, Default))
        Else
            Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la tabla " & Tabla & " con el codigo interno " & Parametro
        End If
        
        rs_Consult.Close
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "Parametro incorrecto al calcular el mapero de la tabla " & Tabla
    End If
    
    CalcularMapeo = Salida

End Function


'----------------------------------------------------------------
'Busca cual es el mapeo de un codigo SAP a un codigo Heidt
Public Function CalcularMapeoInv(ByVal Parametro, ByVal Tabla, ByVal Default)

    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim correcto As Boolean
    Dim Salida
    
    If IsNull(Parametro) Then
       correcto = False
    Else
       correcto = Parametro <> ""
    End If
           
    Salida = Default

    If correcto Then
        StrSql = " SELECT * FROM infotipos_mapeo " & _
                 " WHERE tablaref = '" & Tabla & "' " & _
                 "   AND codexterno = '" & Parametro & "' "
        
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Salida = CStr(IIf(Not IsNull(rs_Consult!codinterno), rs_Consult!codinterno, Default))
        Else
            Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la tabla " & Tabla & " con el codigo externo " & Parametro
        End If
        
        rs_Consult.Close
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "Parametro incorrecto al calcular el mapero de la tabla " & Tabla
    End If
    
    CalcularMapeoInv = Salida

End Function


Public Function CalcularMapeoSubtipo(ByVal Inf As String, ByVal Parametro, ByVal Tabla, ByVal Default)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula el mapeo inverso del subtipo para el infotipo pasado por parametro
' Autor      : FGZ
' Fecha      : 11/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset
Dim correcto As Boolean
Dim Salida
    
    If IsNull(Parametro) Then
       correcto = False
    Else
       correcto = Parametro <> ""
    End If
           
    Salida = Default

    If correcto Then
        StrSql = " SELECT * FROM infotipos_mapeo " & _
                 " WHERE tablaref = '" & Tabla & "' " & _
                 " AND codexterno = '" & Parametro & "' " & _
                 " AND infotipo = '" & Inf & "' "
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Salida = CStr(IIf(Not IsNull(rs_Consult!codinterno), rs_Consult!codinterno, Default))
        Else
            Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la tabla " & Tabla & " con el codigo externo " & Parametro
        End If
        
        rs_Consult.Close
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "Parametro incorrecto al calcular el mapero de la tabla " & Tabla
    End If
    
    CalcularMapeoSubtipo = Salida
End Function

Public Sub Mapear(ByVal Tabla As String, ByVal CodExt As String, CodInt As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: 'Crea el Mapeo entre la tablas de SAP y las de RHPRO
' Autor      : FGZ
' Fecha      : 09/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
    
        StrSql = " SELECT * FROM infotipos_mapeo "
        StrSql = StrSql & " WHERE tablaref = '" & UCase(Tabla) & "' "
        StrSql = StrSql & " AND codexterno = '" & UCase(CodExt) & "' "
        OpenRecordset StrSql, rs
        
        If Not rs.EOF Then
            If UCase(rs!codinterno) <> UCase(CodInt) Then
                StrSql = "UPDATE infotipos_mapeo SET "
                StrSql = StrSql & " codinterno = '" & Format_Str(UCase(CodInt), 10, False, "") & "'"
                StrSql = StrSql & " WHERE tablaref = '" & UCase(Tabla) & "' "
                StrSql = StrSql & " AND codexterno = '" & UCase(CodExt) & "' "
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.Writeline Espacios(Tabulador * 3) & "Mapeo modificado: tabla: " & Tabla & " codigo SAP: " & CodExt & " codigo RHPro: " & CodInt
            End If
        Else
            StrSql = "INSERT INTO infotipos_mapeo (tablaref,codexterno,codinterno) "
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & "'" & Format_Str(UCase(Tabla), 10, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(UCase(CodExt), 10, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(UCase(CodInt), 10, False, "") & "'"
            'StrSql = StrSql & ",'" & Infotipo & "'"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline Espacios(Tabulador * 3) & "Mapeo insertado: tabla: " & Tabla & " codigo SAP: " & CodExt & " codigo RHPro: " & CodInt
        End If
    
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Sub




Public Sub CalcularMapeoNomina(ByVal Nomina As String, ByVal EsMonto As Boolean, ByRef concnro As Long, ByRef Tpanro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca cual es el mapeo de una Nomina a un concepto parametro de RHPro.
' Autor      : FGZ
' Fecha      : 07/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset
Dim correcto As Boolean

    If EsNulo(Nomina) Then
       correcto = False
    Else
       correcto = True
    End If

    If correcto Then
        StrSql = " SELECT * FROM infotipos_mapeo_cc_Nominas "
        StrSql = StrSql & " WHERE cc_nomina = '" & Nomina & "' "
        StrSql = StrSql & " AND esMonto = " & CInt(EsMonto)
        OpenRecordset StrSql, rs_Consult

        If Not rs_Consult.EOF Then
            concnro = rs_Consult!concnro
            Tpanro = rs_Consult!Tpanro
        Else
            Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la nomina " & Nomina & " para " & IIf(EsMonto, "el monto ", "la cantidad")
        End If

        rs_Consult.Close
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "Nomina incorrecta para calcular el mapeo " & Nomina
    End If
    
'cierro y libero
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing
End Sub


Public Sub CalcularMapeoNominaDDJJ(ByVal Nomina As String, ByRef Itenro As Long, ByRef Desmen As Boolean, ByRef Acumula As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca cual es el mapeo de una Nomina a un item de DDJJ de RHPro.
' Autor      : FGZ
' Fecha      : 28/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Acumula = False
Desmen = True
Select Case Nomina
Case "MI10":
    Itenro = 5
    Acumula = True
Case "MI15":
    Itenro = 5
    Acumula = True
Case "MI20":
    Itenro = 6
Case "MI25":
    Itenro = 14
Case "MI30":
    Itenro = 1
Case "MI33":
    Itenro = 2
Case "MI35":
    Itenro = 0
    Flog.Writeline Espacios(Tabulador * 3) & "Nomina sin mapeo asociado " & Nomina
Case "MI40":
    Itenro = 0
    Desmen = False
Case "MI50":
    Itenro = 13
Case "MI55":
    Itenro = 0
    Flog.Writeline Espacios(Tabulador * 3) & "Nomina sin mapeo asociado " & Nomina
Case "MI60":
    Itenro = 9
Case "MI65":
    Itenro = 8
Case "MI70":
    Itenro = 18
Case "MI75":
    Itenro = 20
Case "MI76":
    Itenro = 21
Case "MI77":
    Itenro = 0
    Flog.Writeline Espacios(Tabulador * 3) & "Nomina sin mapeo asociado " & Nomina
Case "MI85":
    Itenro = 0
    Flog.Writeline Espacios(Tabulador * 3) & "Nomina sin mapeo asociado " & Nomina
Case "MI90":
    Itenro = 15
Case "MI92":
    Itenro = 10
Case "MI93":
    Itenro = 11
Case "MI94":
    Itenro = 12
Case Else
    Flog.Writeline Espacios(Tabulador * 3) & "Nomina incorrecta para calcular el mapeo " & Nomina
End Select
    
End Sub

