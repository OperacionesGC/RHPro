Attribute VB_Name = "MdlMapeos"
Option Explicit

'----------------------------------------------------------------
'Busca cual es el mapeo de un codigo RHPro a un codigo SAP
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
        Flog.Writeline Espacios(Tabulador * 3) & "Parametro incorrecto al calcular el mapeo de la tabla " & Tabla
    End If
    
    CalcularMapeo = Salida

End Function


'----------------------------------------------------------------
'Busca cual es el mapeo de un codigo SAP a un codigo RHPro
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
        StrSql = " SELECT * FROM mapeo_sap " & _
                 " WHERE tablaref = '" & Tabla & "' " & _
                 "   AND upper(codexterno) = '" & UCase(Parametro) & "' "
        
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Salida = CStr(IIf(Not IsNull(rs_Consult!codinterno), rs_Consult!codinterno, Default))
        Else
            Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la tabla " & Tabla & " con el codigo externo " & Parametro
        End If
        
        rs_Consult.Close
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "Parametro incorrecto al calcular el mapeo de la tabla " & Tabla
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
        Flog.Writeline Espacios(Tabulador * 3) & "Parametro incorrecto al calcular el mapeo de la tabla " & Tabla
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
    
        StrSql = " SELECT * FROM mapeo_sap "
        StrSql = StrSql & " WHERE tablaref = '" & UCase(Tabla) & "' "
        StrSql = StrSql & " AND codexterno = '" & UCase(CodExt) & "' "
        OpenRecordset StrSql, rs
        
        If Not rs.EOF Then
            If UCase(rs!codinterno) <> UCase(CodInt) Then
                StrSql = "UPDATE mapeo_sap SET "
                StrSql = StrSql & " codinterno = '" & Format_Str(UCase(CodInt), 60, False, "") & "'"
                StrSql = StrSql & " WHERE tablaref = '" & UCase(Tabla) & "' "
                StrSql = StrSql & " AND codexterno = '" & UCase(CodExt) & "' "
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.Writeline Espacios(Tabulador * 3) & "Mapeo modificado: tabla: " & Tabla & " codigo SAP: " & CodExt & " codigo RHPro: " & CodInt
            End If
        Else
            StrSql = "INSERT INTO mapeo_sap (tablaref,codexterno,codinterno) "
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & "'" & Format_Str(UCase(Tabla), 30, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(UCase(CodExt), 60, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(UCase(CodInt), 60, False, "") & "'"
            'StrSql = StrSql & ",'" & Infotipo & "'"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline Espacios(Tabulador * 3) & "Mapeo insertado: tabla: " & Tabla & " codigo SAP: " & CodExt & " codigo RHPro: " & CodInt
        End If
    
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Sub



Public Sub CalcularMapeoNominaAutomatico(ByVal Nomina As String, ByVal Fecha_Fin As Date, ByRef concnro As Long, ByRef tpanro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca cual es el mapeo de una Nomina a un concepto parametro de RHPro.
' Autor      : FGZ
' Fecha      : 03/03/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Concepto As New ADODB.Recordset
Dim rs_CftSegun As New ADODB.Recordset
Dim rs_His_Estructura As New ADODB.Recordset
Dim rs_con_for_tpa As New ADODB.Recordset

Dim Grupo As Long
Dim Alcance As Boolean

        StrSql = " SELECT * FROM concepto "
        StrSql = StrSql & " WHERE upper(concepto.conctexto) = '" & UCase(Nomina) & "' "
        OpenRecordset StrSql, rs_Concepto

        If Not rs_Concepto.EOF Then
            concnro = rs_Concepto!concnro
            
            StrSql = "SELECT * FROM cft_segun "
            StrSql = StrSql & " WHERE concnro = " & concnro
            'StrSql = StrSql & " AND tpanro = " & rs_For_Tpa!TPANRO
            StrSql = StrSql & " AND fornro = " & rs_Concepto!fornro & " AND (("
            StrSql = StrSql & " nivel = 0 AND origen = " & Empleado.Tercero & ") OR "
            StrSql = StrSql & " (nivel = 1) OR "
            StrSql = StrSql & " (nivel = 2)) "
            StrSql = StrSql & " ORDER BY nivel"
            OpenRecordset StrSql, rs_CftSegun
            
            Do While Not rs_CftSegun.EOF And Not Alcance
                tpanro = rs_CftSegun!tpanro
                If rs_CftSegun!Nivel = 1 Then
                    StrSql = " SELECT tenro, estrnro FROM his_estructura "
                    StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero & " AND "
                    StrSql = StrSql & " tenro =" & rs_CftSegun!Entidad & " AND "
                    StrSql = StrSql & " (htetdesde <= " & ConvFecha(Fecha_Fin) & ") AND "
                    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin) & " <= htethasta) or (htethasta is null))"
                    If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
                    OpenRecordset StrSql, rs_His_Estructura

                    If Not rs_His_Estructura.EOF Then
                        If rs_CftSegun!Nivel = 1 And rs_CftSegun!Origen = rs_His_Estructura!Estrnro Then
                            Alcance = True
                            Grupo = rs_CftSegun!Origen
                            'Selecc = rs_CftSegun!Selecc
                        Else
                            Alcance = False
                            Grupo = 0
                        End If
                    Else
                        Alcance = False
                    End If
                Else
                    'el alcance es global
                    Alcance = True
                    Grupo = 0
                End If 'If rs_CftSegun.nivel = 1 Then
                
                If Not Alcance Then
                    rs_CftSegun.MoveNext
                End If
            Loop
            
            If Not Alcance Then
                Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la nomina " & Nomina
            Else
                If EsNulo(rs_CftSegun!Selecc) Then
                    StrSql = "SELECT * FROM con_for_tpa WHERE concnro = " & concnro & _
                             " AND fornro =" & rs_Concepto!fornro & _
                             " AND tpanro =" & tpanro & _
                             " AND nivel =" & rs_CftSegun!Nivel
                Else
                    StrSql = "SELECT * FROM con_for_tpa WHERE concnro = " & concnro & _
                             " AND fornro =" & rs_Concepto!fornro & _
                             " AND tpanro =" & tpanro & _
                             " AND nivel =" & rs_CftSegun!Nivel & _
                             " AND selecc ='" & rs_CftSegun!Selecc & "'"
                End If
                OpenRecordset StrSql, rs_con_for_tpa
            
                If rs_con_for_tpa.EOF Then
                    Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la nomina " & Nomina
                End If
            End If
        Else
            Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la nomina " & Nomina
        End If
        
        
        'Flog.writeline Espacios(Tabulador * 3) & "Nomina incorrecta para calcular el mapeo " & Nomina
    
'cierro y libero
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
Set rs_Concepto = Nothing
If rs_CftSegun.State = adStateOpen Then rs_CftSegun.Close
Set rs_CftSegun = Nothing
If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
Set rs_His_Estructura = Nothing
If rs_con_for_tpa.State = adStateOpen Then rs_con_for_tpa.Close
Set rs_con_for_tpa = Nothing
End Sub

Public Sub MapearNominasAutomaticamente()
' ---------------------------------------------------------------------------------------------
' Descripcion: Genera todos los mapeo de nominas automaticamente de acuerdo a la descripcion extendida
'              del concepto para todos los conceptos de RHPro.
' Autor      : FGZ
' Fecha      : 01/04/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Concepto As New ADODB.Recordset
Dim rs_CftSegun As New ADODB.Recordset
Dim rs_con_for_tpa As New ADODB.Recordset
Dim concnro As Long
Dim tpanro As Long
Dim Conccod As String

    On Error GoTo MELocal
    
    StrSql = " SELECT * FROM concepto "
    StrSql = StrSql & " INNER JOIN for_tpa ON concepto.fornro = for_tpa.fornro "
    OpenRecordset StrSql, rs_Concepto

    Do While Not rs_Concepto.EOF
        If Not EsNulo(Trim(rs_Concepto!Conctexto)) Then
            concnro = rs_Concepto!concnro
            tpanro = rs_Concepto!tpanro
            Conccod = rs_Concepto!Conccod
            
            'revisar que el parametro se resuelva por novedad
            StrSql = "SELECT * FROM cft_segun " & _
                    " WHERE concnro = " & concnro & _
                    " AND tpanro = " & tpanro & _
                    " AND fornro = " & rs_Concepto!fornro & " AND nivel <> 1 " & _
                    " ORDER BY nivel"
            OpenRecordset StrSql, rs_CftSegun
            Do While Not rs_CftSegun.EOF
                If EsNulo(rs_CftSegun!Selecc) Then
                    StrSql = "SELECT * FROM con_for_tpa WHERE concnro = " & concnro & _
                             " AND fornro =" & rs_Concepto!fornro & _
                             " AND tpanro =" & tpanro & _
                             " AND nivel =" & rs_CftSegun!Nivel
                Else
                    StrSql = "SELECT * FROM con_for_tpa WHERE concnro = " & concnro & _
                             " AND fornro =" & rs_Concepto!fornro & _
                             " AND tpanro =" & tpanro & _
                             " AND nivel =" & rs_CftSegun!Nivel & _
                             " AND selecc ='" & rs_CftSegun!Selecc & "'"
                End If
                OpenRecordset StrSql, rs_con_for_tpa
                
                Do While Not rs_con_for_tpa.EOF
                    If Not CBool(rs_con_for_tpa!cftauto) Then
                        'Insertar el mapeo de la nimona
                        Call Insertar_Mapeo_Nomina(concnro, tpanro, UCase(Trim(rs_Concepto!Conctexto)), True)
                    End If
                    
                    rs_con_for_tpa.MoveNext
                Loop
                
                rs_CftSegun.MoveNext
            Loop
        Else
            Flog.Writeline Espacios(Tabulador * 1) & "El concepto " & rs_Concepto!Conccod & " no tiene nomina asociada "
        End If
        
        'Siguiente concepto
        rs_Concepto.MoveNext
    Loop

'cierro y libero
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
Set rs_Concepto = Nothing
If rs_CftSegun.State = adStateOpen Then rs_CftSegun.Close
Set rs_CftSegun = Nothing
If rs_con_for_tpa.State = adStateOpen Then rs_con_for_tpa.Close
Set rs_con_for_tpa = Nothing

Exit Sub

MELocal:
    'Resume Next
    Flog.Writeline Espacios(Tabulador * 1) & "Error en mapeos de nominas " & Err.Description
End Sub


'Public Sub MapearNominasAutomaticamente_old()
' ---------------------------------------------------------------------------------------------
' Descripcion: Genera todos los mapeo de nominas automaticamente de acuerdo a la descripcion extendida
'              del concepto para todos los conceptos de RHPro.
'Autor:        FGZ
' Fecha      : 01/04/2005
' Ultima Mod.:
'Descripcion:
' ---------------------------------------------------------------------------------------------
'Dim rs_Concepto As New ADODB.Recordset
'Dim rs_CftSegun As New ADODB.Recordset
'Dim rs_His_Estructura As New ADODB.Recordset
'Dim rs_con_for_tpa As New ADODB.Recordset
'
'Dim Grupo As Long
'Dim Alcance As Boolean
'
'    StrSql = " SELECT * FROM concepto "
'    OpenRecordset StrSql, rs_Concepto
'
'    Do While Not rs_Concepto.EOF
'        If Not EsNulo(rs_Concepto!Conctexto) Then
'            Concnro = rs_Concepto!Concnro
'
'            StrSql = "SELECT * FROM cft_segun "
'            StrSql = StrSql & " WHERE concnro = " & Concnro
'            StrSql = StrSql & " AND tpanro = " & rs_For_Tpa!Tpanro
'            StrSql = StrSql & " AND fornro = " & rs_Concepto!fornro & " AND (("
'            StrSql = StrSql & " nivel = 0 AND origen = " & Empleado.Tercero & ") OR "
'            StrSql = StrSql & " (nivel = 1) OR "
'            StrSql = StrSql & " (nivel = 2)) "
'            StrSql = StrSql & " ORDER BY nivel"
'            OpenRecordset StrSql, rs_CftSegun
'
'            Do While Not rs_CftSegun.EOF And Not Alcance
'                Tpanro = rs_CftSegun!Tpanro
'                If rs_CftSegun!Nivel <> 1 Then
'                    Alcance = True
'                    Grupo = 0
'                End If
'            Loop
'
'            If Not Alcance Then
'                Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la nomina " & Nomina
'            Else
'                If EsNulo(rs_CftSegun!Selecc) Then
'                    StrSql = "SELECT * FROM con_for_tpa WHERE concnro = " & Concnro & _
'                             " AND fornro =" & rs_Concepto!fornro & _
'                             " AND tpanro =" & Tpanro & _
'                             " AND nivel =" & rs_CftSegun!Nivel
'                Else
'                    StrSql = "SELECT * FROM con_for_tpa WHERE concnro = " & Concnro & _
'                             " AND fornro =" & rs_Concepto!fornro & _
'                             " AND tpanro =" & Tpanro & _
'                             " AND nivel =" & rs_CftSegun!Nivel & _
'                             " AND selecc ='" & rs_CftSegun!Selecc & "'"
'                End If
'                OpenRecordset StrSql, rs_con_for_tpa
'
'                If rs_con_for_tpa.EOF Then
'                    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la nomina " & Nomina
'                Else
'                    Insertar el mapeo de la nimona
'                    Call Insertar_Mapeo_Nomina(Concnro, Tpanro, UCase(Trim(rs_Concepto!Conctexto)))
'                End If
'            End If
'        End If
'
'        Siguiente concepto
'        rs_Concepto.MoveNext
'    Loop
'
'
'        Flog.writeline Espacios(Tabulador * 3) & "Nomina incorrecta para calcular el mapeo " & Nomina
'
'cierro y libero
'If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
'Set rs_Concepto = Nothing
'If rs_CftSegun.State = adStateOpen Then rs_CftSegun.Close
'Set rs_CftSegun = Nothing
'If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
'Set rs_His_Estructura = Nothing
'If rs_con_for_tpa.State = adStateOpen Then rs_con_for_tpa.Close
'Set rs_con_for_tpa = Nothing
'End Sub




Public Sub CalcularMapeoNomina(ByVal Nomina As String, ByVal EsMonto As Boolean, ByRef concnro As Long, ByRef tpanro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca cual es el mapeo de una Nomina a un concepto parametro de RHPro.
' Autor      : FGZ
' Fecha      : 07/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset
Dim correcto As Boolean


    concnro = 0
    tpanro = 0
    
    If EsNulo(Nomina) Then
       correcto = False
    Else
       correcto = True
    End If

    If correcto Then
        StrSql = " SELECT * FROM mapeo_sap_Nomina "
        StrSql = StrSql & " WHERE cc_nomina = '" & Nomina & "' "
        StrSql = StrSql & " AND esMonto = " & CInt(EsMonto)
        OpenRecordset StrSql, rs_Consult

        If Not rs_Consult.EOF Then
            concnro = rs_Consult!concnro
            tpanro = rs_Consult!tpanro
        Else
            StrSql = " SELECT * FROM mapeo_sap_Nomina "
            StrSql = StrSql & " WHERE cc_nomina = '" & Nomina & "' "
            StrSql = StrSql & " AND esMonto = " & CInt(Not EsMonto)
            OpenRecordset StrSql, rs_Consult
            If rs_Consult.EOF Then
                Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo para la nomina " & Nomina & " para " & IIf(EsMonto, "el monto ", "la cantidad")
            End If
        End If

        rs_Consult.Close
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "Nomina incorrecta para calcular el mapeo " & Nomina
    End If
    
'cierro y libero
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing
End Sub


Public Sub CalcularMapeoNominaDDJJ(ByVal Nomina As String, ByRef Itenro As Long, ByRef Desmen As Boolean, ByRef Acumula As Boolean, ByRef Cantidad As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca cual es el mapeo de una Nomina a un item de DDJJ de RHPro.
' Autor      : FGZ
' Fecha      : 28/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Cantidad = False
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
    Itenro = 22
Case "MI85":
    Itenro = 7
Case "MI90":
    Itenro = 15
Case "MI92":
    Itenro = 10
    Acumula = True
    Cantidad = True
Case "MI93":
    Itenro = 11
    Acumula = True
    Cantidad = True
Case "MI94":
    Itenro = 12
    Acumula = True
    Cantidad = True
Case Else
    Flog.Writeline Espacios(Tabulador * 3) & "Nomina incorrecta para calcular el mapeo " & Nomina
End Select
    
End Sub


Public Sub Insertar_Mapeo_Nomina(ByVal concnro As Long, ByVal tpanro As Long, ByVal Nomina As String, ByVal EsMonto As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta el mapeo de una Nomina a un concepto parametro de RHPro.
' Autor      : FGZ
' Fecha      : 02/04/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset

        StrSql = " SELECT * FROM mapeo_sap_Nomina "
        StrSql = StrSql & " WHERE cc_nomina = '" & Nomina & "' "
        'StrSql = StrSql & " AND esMonto = " & CInt(EsMonto)
        StrSql = StrSql & " AND concnro = " & concnro
        StrSql = StrSql & " AND tpanro = " & tpanro
        OpenRecordset StrSql, rs_Consult
    
        If rs_Consult.EOF Then
            'Insertar mapeo
            StrSql = "INSERT INTO mapeo_sap_Nomina (cc_nomina,concnro,tpanro,esmonto)"
            StrSql = StrSql & " VALUES ( "
            StrSql = StrSql & "'" & Nomina & "',"
            StrSql = StrSql & concnro & ","
            StrSql = StrSql & tpanro & ","
            StrSql = StrSql & CInt(EsMonto)
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

'cierro y libero
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing
End Sub
