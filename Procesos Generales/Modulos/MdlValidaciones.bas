Attribute VB_Name = "MdlValidaciones"
Option Explicit

Public Sub ValidarLocalidad(Localidad As String, ByRef Nro_Localidad As Long, Nro_Pais As Long, Nro_Provincia As Long)
Dim rs_sub As New ADODB.Recordset
Dim Sql_Ins As String
Dim SQL_Val As String

If Not EsNulo(Localidad) Then
    StrSql = " SELECT * FROM localidad WHERE UPPER(locdesc) = '" & Format_Str(UCase(Localidad), 30, False, "") & "'"
'    If nro_pais <> 0 Then
'        StrSql = StrSql & " AND paisnro = " & nro_pais
'    End If
'
'    If nro_provincia <> 0 Then
'        StrSql = StrSql & " AND provnro = " & nro_provincia
'    End If
    OpenRecordset StrSql, rs_sub
    
    If rs_sub.EOF Then
    
        Sql_Ins = " INSERT INTO localidad(locdesc"
        SQL_Val = " VALUES('" & Format_Str(UCase(Localidad), 30, False, "") & "'"
    
        If Nro_Pais <> 0 Then
            Sql_Ins = Sql_Ins & ",paisnro"
            SQL_Val = SQL_Val & "," & Nro_Pais
        End If
        If Nro_Provincia <> 0 Then
            Sql_Ins = Sql_Ins & ",provnro"
            SQL_Val = SQL_Val & "," & Nro_Provincia
        End If
        StrSql = Sql_Ins & ")" & SQL_Val & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Nro_Localidad = getLastIdentity(objConn, "localidad")
    Else
    
        Nro_Localidad = rs_sub!locnro
    
        If Nro_Provincia = 0 Then
            Nro_Provincia = IIf(Not EsNulo(rs_sub!provnro), rs_sub!provnro, 0)
        End If
        If Nro_Pais = 0 Then
            Nro_Pais = IIf(Not EsNulo(rs_sub!PaisNro), rs_sub!PaisNro, 0)
        End If
    End If
End If
End Sub

Public Sub ValidarPartido(Partido As String, ByRef Nro_Partido As Long)

Dim rs_sub As New ADODB.Recordset

If Not EsNulo(Partido) Then
    StrSql = " SELECT * FROM partido WHERE UPPER(partnom) = '" & Format_Str(UCase(Partido), 30, False, "") & "'"
    OpenRecordset StrSql, rs_sub
    
    If rs_sub.EOF Then
    
        StrSql = "INSERT INTO partido(partnom) VALUES('"
        StrSql = StrSql & Format_Str(UCase(Partido), 30, False, "") & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        Nro_Partido = getLastIdentity(objConn, "partido")
        
'        StrSql = " SELECT MAX(partnro) AS MaxPart FROM partido "
'        'StrSql = " SELECT @@IDENTITY AS MaxPart "
'        OpenRecordset StrSql, rs_sub
'
'        nro_partido = rs_sub!MaxPart
    
    Else
        
        Nro_Partido = rs_sub!partnro
    
    End If
End If
End Sub

Public Sub ValidarZona(Zona As String, ByRef nro_zona As Integer, Nro_Provincia As Integer)

Dim rs_sub As New ADODB.Recordset

    If Not EsNulo(Zona) Then
        StrSql = " SELECT * FROM zona WHERE UPPER(zonadesc) = '" & Format_Str(UCase(Zona), 20, False, "") & "' AND provnro = " & Nro_Provincia
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
        
            StrSql = "INSERT INTO zona(zonadesc,provnro) VALUES('"
            StrSql = StrSql & Format_Str(UCase(Zona), 20, False, "") & "'," & Nro_Provincia & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
'            StrSql = " SELECT MAX(zona) AS MaxZona FROM zona "
'            'StrSql = " SELECT @@IDENTITY AS MaxZona "
'            OpenRecordset StrSql, rs_sub
'
'            nro_zona = rs_sub!MaxZona
            nro_zona = getLastIdentity(objConn, "zona")
        Else
            
            nro_zona = rs_sub!zonanro
        
        End If
    End If

End Sub

Public Sub ValidarProvincia(Provincia As String, ByRef Nro_Provincia As Long, Nro_Pais As Long)

Dim rs_sub As New ADODB.Recordset

If Not EsNulo(Provincia) Then
    'StrSql = " SELECT * FROM provincia WHERE provdesc = '" & Provincia & "' AND paisnro = " & nro_pais
    StrSql = " SELECT * FROM provincia WHERE upper(provdesc) = '" & Format_Str(UCase(Provincia), 20, False, "") & "'"
    OpenRecordset StrSql, rs_sub
    
    If rs_sub.EOF Then
    
        StrSql = "INSERT INTO provincia(provdesc,paisnro) VALUES('"
        StrSql = StrSql & Format_Str(UCase(Provincia), 20, False, "") & "'," & Nro_Pais & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Nro_Provincia = getLastIdentity(objConn, "provincia")
        
        
'        StrSql = " SELECT MAX(provnro) AS MaxProv FROM provincia "
'        'StrSql = " SELECT @@IDENTITY AS MaxProv "
'        OpenRecordset StrSql, rs_sub
'
'        nro_provincia = rs_sub!MaxProv
    
    Else
        
        Nro_Provincia = rs_sub!provnro
    
    End If
End If
End Sub

Public Sub ValidarPais(Pais As String, ByRef Nro_Pais As Long)

Dim rs_sub As New ADODB.Recordset

    If Not EsNulo(Pais) Then
        StrSql = " SELECT * FROM pais WHERE UPPER(paisdesc) = '" & Format_Str(UCase(Pais), 20, False, "") & "'"
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
            StrSql = "INSERT INTO pais(paisdesc,paisdef) VALUES('"
            StrSql = StrSql & Format_Str(UCase(Pais), 20, False, "") & "',0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Nro_Pais = getLastIdentity(objConn, "pais")
        Else
            Nro_Pais = rs_sub!PaisNro
        End If
    End If


End Sub

Public Sub ValidarEstadoCivil(ByVal EstadoCivil As String, ByRef Nro_EstadoCivil As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Validacion.
' Autor      : FGZ
' Fecha      : 09/02/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset

    StrSql = " SELECT estcivnro FROM estcivil WHERE UPPER(estcivdesabr) = '" & Format_Str(UCase(EstadoCivil), 30, False, "") & "'"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Nro_EstadoCivil = rs!EstCivNro
    Else
        StrSql = " INSERT INTO estcivil(estcivdesabr) VALUES ('" & Format_Str(UCase(EstadoCivil), 30, False, "") & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Nro_EstadoCivil = getLastIdentity(objConn, "estcivil")
    End If
        
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Sub

Public Sub ValidarMoneda(ByVal Moneda As String, ByVal Codigo As String, ByVal Pais As Long, ByRef Nro_Moneda As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Validacion.
' Autor      : FGZ
' Fecha      : 09/02/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset

    StrSql = " SELECT monnro FROM moneda WHERE UPPER(mondesabr) = '" & Format_Str(UCase(Moneda), 30, False, "") & "'"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Nro_Moneda = rs!monnro
    Else
    
        StrSql = " INSERT INTO moneda(mondesabr,paisnro,monorigen,monlocal,moninternac) VALUES ("
        StrSql = StrSql & "'" & Format_Str(UCase(Moneda), 30, False, "") & "'"
        If UCase(Codigo) = "ARS" Then
            StrSql = StrSql & ",3,-1,-1,0"
        Else
            StrSql = StrSql & ",1,0,0,-1"
        End If
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Nro_Moneda = getLastIdentity(objConn, "moneda")
    End If
        
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Sub


Public Sub CalcularLegajo(NroEmp As Integer, ByRef Legajo As String)

Dim rs_leg As New ADODB.Recordset
Dim rs_emp As New ADODB.Recordset

Dim NroLegajo As Integer
Dim Continuar As Boolean

    
        StrSql = "SELECT MAX(empleg) AS ProxLegajo FROM empleado"
        StrSql = StrSql & " WHERE ternro in (SELECT ternro FROM his_estructura"
        StrSql = StrSql & " WHERE tenro = 10 AND estrnro = " & NroEmp & " AND htethasta IS NULL)"
        OpenRecordset StrSql, rs_leg
        
        NroLegajo = rs_leg!ProxLegajo + 1
        
        Continuar = True
                
        Do While Continuar
        
            StrSql = "SELECT empleg FROM empleado WHERE empleg = " & NroLegajo
            OpenRecordset StrSql, rs_emp
        
            If rs_emp.EOF Then
                Continuar = False
            Else
                NroLegajo = NroLegajo + 1
            End If
        
        Loop
        
        Legajo = Str(NroLegajo)
        
End Sub

Public Sub ValidaEstructura(TipoEstr As Long, ByRef Valor As String, ByRef CodEst As Long, ByRef Inserto_estr As Boolean)
Dim Rs_Estr As New ADODB.Recordset

Dim d_estructura As String
Dim CodExt As String
Dim l_pos1 As Integer
Dim l_pos2 As Integer


    If InStr(1, Valor, "$") > 0 Then
        l_pos1 = InStr(1, Valor, "$")
        l_pos2 = Len(Valor)
    
        d_estructura = Mid(Valor, l_pos1 + 2, l_pos2)
        If l_pos1 <> 0 Then
            CodExt = Trim(Mid(Valor, 1, l_pos1 - 1))
        Else
            CodExt = ""
        End If
    Else
        d_estructura = Valor
        CodExt = ""
    End If
    
    Valor = d_estructura
    
    StrSql = " SELECT * FROM estructura WHERE UPPER(estructura.estrdabr) = '" & Format_Str(UCase(d_estructura), 60, False, "") & "'"
    If Not EsNulo(CodExt) Then
        StrSql = StrSql & " AND estrcodext ='" & Format_Str(UCase(CodExt), 20, False, "") & "'"
    End If
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    OpenRecordset StrSql, Rs_Estr
        
    If Not Rs_Estr.EOF Then
            CodEst = Rs_Estr!Estrnro
            Inserto_estr = False
            
            If Not EsNulo(CodExt) Then
                If EsNulo(Rs_Estr!estrcodext) Then
                    ' ACTUALIZO el codigo externro
                    StrSql = " UPDATE estructura SET estrcodext ='" & Format_Str(UCase(CodExt), 20, False, "") & "'"
                    StrSql = StrSql & " WHERE estrnro = " & Rs_Estr!Estrnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    'FGZ - 30/08/2005
                    'ya existe, no hago nada
                    
'                    If Left(UCase(Rs_Estr!estrcodext), 20) <> Left(UCase(CodExt), 20) Then
'                        ' ACTUALIZO el codigo externro
'                        StrSql = " UPDATE estructura SET estrcodext ='" & Format_Str(UCase(CodExt), 20, False, "") & "'"
'                        StrSql = StrSql & " WHERE estrnro = " & Rs_Estr!Estrnro
'                        objConn.Execute StrSql, , adExecuteNoRecords
'                    End If
                End If
            End If
    Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & Format_Str(UCase(d_estructura), 60, False, "") & "',1,-1,'" & Format_Str(UCase(CodExt), 20, False, "") & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            CodEst = getLastIdentity(objConn, "estructura")
            
            Inserto_estr = True
    End If

If Rs_Estr.State = adStateOpen Then Rs_Estr.Close
Set Rs_Estr = Nothing
End Sub
Public Sub CreaComplemento(TipoEstr As Long, CodTer As Long, CodEstr As Long, Valor As String)


  Select Case TipoEstr

    Case 1
        Complementos1 CodTer, CodEstr
    Case 3
        Complementos3 CodTer, CodEstr
    Case 4
        Complementos4 CodEstr, Valor
    Case 10
        Complementos10 CodTer, CodEstr, Valor
    Case 15
        Complementos15 CodTer, CodEstr
    Case 16
        Complementos16 CodTer, CodEstr
    Case 17
        Complementos17 CodTer, CodEstr, Valor
    Case 18
        Complementos18 CodTer, CodEstr, Valor
    Case 19
        Complementos19 CodEstr
    Case 22
        Complementos22 CodTer, CodEstr, Valor
    Case 23
        Complementos23 CodTer, CodEstr, Valor
    Case 24
        Complementos17 CodTer, CodEstr, Valor
    Case 40
        Complementos40 CodTer, CodEstr, Valor
    Case 41
        Complementos41 CodTer, CodEstr, Valor

  End Select
 
End Sub
Public Sub Complementos1(CodTer As Long, CodEstr As Long)

    StrSql = " INSERT INTO sucursal(estrnro,ternro,sucest) VALUES(" & CodEstr & "," & CodTer & ",-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos3(CodTer As Long, CodEstr As Long)

    StrSql = " INSERT INTO categoria(estrnro,convnro) VALUES(" & CodEstr & "," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos4(CodEstr As Long, Valor As String)

    StrSql = " INSERT INTO puesto(estrnro,puedesc,puenroreemp) VALUES(" & CodEstr & ",'" & Format_Str(UCase(Valor), 60, False, "") & "',0)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos10(CodTer As Long, CodEstr As Long, Valor As String)

    StrSql = " INSERT INTO empresa(estrnro,ternro,empnom) VALUES(" & CodEstr & "," & CodTer & ",'" & Format_Str(UCase(Valor), 100, False, "") & "')"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos15(CodTer As Long, CodEstr As Long)

    ' Hay que crear un Tipo de Caja de Jubilacion "Migracion"

    StrSql = " INSERT INTO cajjub(estrnro,ternro,cajest,ticnro) VALUES(" & CodEstr & "," & CodTer & ",-1,1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos16(CodTer As Long, CodEstr As Long)

    StrSql = " INSERT INTO gremio(estrnro,ternro) VALUES(" & CodEstr & "," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos17(CodTer As Long, CodEstr As Long, Valor As String)
' Ultima Modificacion:  FGZ
' Fecha:                17/12/2004
'---------------------------------------------------------
Dim rs_17 As New ADODB.Recordset

    StrSql = "SELECT * FROM osocial  where osdesc = '" & Format_Str(UCase(Valor), 100, False, "") & "'"
    If rs_17.State = adStateOpen Then rs_17.Close
    OpenRecordset StrSql, rs_17
    
    If rs_17.EOF Then
        StrSql = " INSERT INTO osocial(ternro,osdesc) VALUES(" & CodTer & ",'" & Format_Str(UCase(Valor), 100, False, "") & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    StrSql = "SELECT * FROM replica_estr  where origen = " & CodTer
    StrSql = StrSql & " AND estrnro = " & CodEstr
    If rs_17.State = adStateOpen Then rs_17.Close
    OpenRecordset StrSql, rs_17
    If rs_17.EOF Then
        StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & CodTer & "," & CodEstr & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    If rs_17.State = adStateOpen Then rs_17.Close
    Set rs_17 = Nothing
End Sub

Public Sub Complementos18(CodTer As Long, CodEstr As Long, Valor As String)
'FGZ - 16/03/2006
'se agregaron los campos "leynro" y "tcreduccion" sino el asp no lo mostraba

Dim rs_tipocont As New ADODB.Recordset
Dim rs_TC As New ADODB.Recordset
Dim CodTC As Long


    
    StrSql = "SELECT * FROM tipocont  where tcdabr = '" & Format_Str(UCase(Valor), 60, False, "") & "'"
    OpenRecordset StrSql, rs_tipocont
    
    If rs_tipocont.EOF Then
        StrSql = " INSERT INTO tipocont(tcdabr,estrnro,tcind,leynro,tcreduccion) VALUES('" & Format_Str(UCase(Valor), 60, False, "") & "'," & CodEstr & ",-1,1,0)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
'        StrSql = " SELECT MAX(tcnro) AS CodTC FROM tipocont "
'        'StrSql = " SELECT @@IDENTITY AS CodTC "
'        OpenRecordset StrSql, rs_TC
        
        CodTC = getLastIdentity(objConn, "tipocont")
        
        StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & CodTC & "," & CodEstr & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

    End If
End Sub

Public Sub Complementos19(CodEstr As Long)

    StrSql = " INSERT INTO convenios(estrnro) VALUES(" & CodEstr & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos22(CodTer As Long, CodEstr As Long, Valor As String)

    StrSql = " INSERT INTO formaliq(estrnro,folisistema) VALUES(" & CodEstr & ",-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos23(CodTer As Long, CodEstr As Long, Valor As String)

Dim rs_pos As New ADODB.Recordset
Dim CodPlan As Long

    ' Hay que ver la relacion entra la Osocial y el Plan

    StrSql = " INSERT INTO planos(plnom,osocial) VALUES('" & Format_Str(UCase(Valor), 60, False, "") & "'," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    CodPlan = getLastIdentity(objConn, "planos")
    
    StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & CodPlan & "," & CodEstr & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    

End Sub

Public Sub Complementos40(CodEstr As Long, CodTer As Long, Valor As String)

    StrSql = " INSERT INTO seguro(ternro,estrnro,segdesc,segest) VALUES(" & CodEstr & "," & CodTer & ",'" & Format_Str(UCase(Valor), 200, False, "") & "',-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos41(CodEstr As Long, CodTer As Long, Valor As String)
Dim rs As New ADODB.Recordset

    StrSql = "SELECT * FROM banco WHERE bansucdesc = '" & Format_Str(UCase(Valor), 40, False, "") & "'"
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
        StrSql = " INSERT INTO banco(ternro,estrnro,bansucdesc,banest) VALUES(" & CodEstr & "," & CodTer & ",'" & UCase(Format_Str(Valor, 40, False, "")) & "',-1)"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

End Sub

Public Sub CreaTercero(TipoTer As Long, Valor As String, ByRef CodTer)

Dim rs As New ADODB.Recordset
Dim rs_Ter As New ADODB.Recordset

Dim d_estructura As String
Dim l_pos1 As Integer
Dim l_pos2 As Integer

    
  d_estructura = Valor
    
  StrSql = " SELECT * FROM tercero "
  StrSql = StrSql & " INNER JOIN ter_tip ON tercero.ternro = ter_tip.ternro AND ter_tip.tipnro =" & TipoTer
  StrSql = StrSql & " WHERE terrazsoc = '" & Valor & "'"
  If rs_Ter.State = adStateOpen Then rs_Ter.Close
  OpenRecordset StrSql, rs_Ter
  If rs_Ter.EOF Then
    
      StrSql = " INSERT INTO tercero(terrazsoc,tersex)"
      StrSql = StrSql & " VALUES('" & Mid(d_estructura, 1, 60) & "',-1)"
      objConn.Execute StrSql, , adExecuteNoRecords
    
      CodTer = getLastIdentity(objConn, "tercero")
    
      StrSql = " INSERT INTO ter_tip(ternro,tipnro) "
      StrSql = StrSql & " VALUES(" & CodTer & "," & TipoTer & ")"
      objConn.Execute StrSql, , adExecuteNoRecords
    Else
        CodTer = rs_Ter!ternro
    End If

    If rs_Ter.State = adStateOpen Then rs_Ter.Close
    Set rs_Ter = Nothing
End Sub

Public Sub ValidaEstructuraCodExt(TipoEstr As Integer, ByRef Valor As String, ByRef CodEst As Integer, ByRef Inserto_estr As Boolean)

Dim Rs_Estr As New ADODB.Recordset

Dim d_estructura As String
Dim CodExt As String
Dim l_pos1 As Byte
Dim l_pos2 As Byte


    d_estructura = Valor
    StrSql = " SELECT estrnro FROM estructura WHERE estructura.estrcodext = '" & Format_Str(UCase(Valor), 20, False, "") & "'"
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    OpenRecordset StrSql, Rs_Estr
        
    If Not Rs_Estr.EOF Then
            CodEst = Rs_Estr!Estrnro
            Inserto_estr = False
    Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & Format_Str(UCase(d_estructura), 60, False, "") & "',1,-1,'" & Format_Str(UCase(d_estructura), 20, False, "") & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            CodEst = getLastIdentity(objConn, "estructura")

            Inserto_estr = True
    End If

End Sub

Public Sub ValidaCategoria(TipoEstr As Integer, ByRef Valor As String, nroConv As Integer, ByRef CodEst As Integer, ByRef Inserto_estr As Boolean)
Dim pos1 As Integer
Dim pos2 As Integer

Dim Rs_Estr As New ADODB.Recordset
Dim Rs_Conv As New ADODB.Recordset
Dim Rs_NroC As New ADODB.Recordset


Dim d_estructura As String
Dim l_pos1 As Byte
Dim l_pos2 As Byte
Dim CodExt As String

Dim nro_convenio As Integer

    If InStr(1, Valor, "$") > 0 Then
        l_pos1 = InStr(1, Valor, "$")
        l_pos2 = Len(Valor)
    
        d_estructura = Mid(Valor, l_pos1 + 2, l_pos2)
        If l_pos1 <> 0 Then
            CodExt = Mid(Valor, 1, l_pos1 - 1)
        Else
            CodExt = ""
        End If
    Else
        d_estructura = Valor
        CodExt = ""
    End If
    
    Valor = d_estructura
    
    If nroConv <> 0 Then
    
        StrSql = "SELECT * FROM convenios WHERE estrnro = " & nroConv
        OpenRecordset StrSql, Rs_NroC
    
        If Not Rs_NroC.EOF Then
        
            nro_convenio = Rs_NroC!convnro
            
        End If
        
    
    End If
    
    
            
    StrSql = " SELECT estrnro FROM estructura WHERE UPPER(estructura.estrdabr) = '" & Format_Str(UCase(d_estructura), 60, False, "") & "'"
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    OpenRecordset StrSql, Rs_Estr
        
    If Not Rs_Estr.EOF Then
                
          StrSql = " SELECT convnro FROM categoria WHERE categoria.estrnro = " & Rs_Estr!Estrnro
          OpenRecordset StrSql, Rs_Conv
                
          If (Not Rs_Conv.EOF) And (nro_convenio = Rs_Conv!convnro) Then
            
            CodEst = Rs_Estr!Estrnro
            Inserto_estr = False
                
          Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & Format_Str(UCase(d_estructura), 60, False, "") & "',1,-1,'" & Format_Str(UCase(CodExt), 20, False, "") & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            CodEst = getLastIdentity(objConn, "estructura")
            
            Inserto_estr = True
            
            If nro_convenio <> 0 Then
                StrSql = " INSERT INTO categoria(estrnro,convnro) VALUES(" & CodEst & "," & nro_convenio & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
          End If
                
            
    Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & UCase(Mid(d_estructura, 1, 60)) & "',1,-1,'" & UCase(Mid(CodExt, 1, 20)) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            CodEst = getLastIdentity(objConn, "estructura")
            
            Inserto_estr = True
            
            If nro_convenio <> 0 Then
                StrSql = " INSERT INTO categoria(estrnro,convnro) VALUES(" & CodEst & "," & nro_convenio & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
    End If


End Sub


Public Sub ValidaCategoriaCodExt(TipoEstr As Integer, ByRef Valor As String, nroConv As Integer, ByRef CodEst As Integer, ByRef Inserto_estr As Boolean)
Dim pos1 As Integer
Dim pos2 As Integer

Dim Rs_Estr As New ADODB.Recordset
Dim Rs_Conv As New ADODB.Recordset

Dim d_estructura As String
Dim l_pos1 As Byte
Dim l_pos2 As Byte
Dim CodExt As String

    If InStr(1, Valor, "$") > 0 Then
        l_pos1 = InStr(1, Valor, "$")
        l_pos2 = Len(Valor)
    
        d_estructura = Mid(Valor, l_pos1 + 2, l_pos2)
        If l_pos1 <> 0 Then
            CodExt = Mid(Valor, 1, l_pos1 - 1)
        Else
            CodExt = ""
        End If
    Else
        d_estructura = Valor
        CodExt = ""
    End If
    
    Valor = d_estructura
    
    
    StrSql = " SELECT estrnro FROM estructura WHERE UPPER(estructura.estrcodext) = '" & Format_Str(UCase(d_estructura), 20, False, "") & "'"
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    OpenRecordset StrSql, Rs_Estr
        
    If Not Rs_Estr.EOF Then
                
          StrSql = " SELECT convnro FROM categoria WHERE categoria.estrnro = " & Rs_Estr!Estrnro
          OpenRecordset StrSql, Rs_Conv
                
          If (Not Rs_Conv.EOF) And (nroConv = Rs_Conv!convnro) Then
            
            CodEst = Rs_Estr!Estrnro
            Inserto_estr = False
                
          Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & Format_Str(UCase(d_estructura), 60, False, "") & "',1,-1," & Format_Str(UCase(CodExt), 20, False, "") & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            CodEst = getLastIdentity(objConn, "estructura")
            
            Inserto_estr = True
            
            StrSql = " INSERT INTO categoria(estrnro,convnro) VALUES(" & CodEst & "," & nroConv & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
          End If
                
            
    Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & UCase(Mid(d_estructura, 1, 60)) & "',1,-1," & UCase(Mid(CodExt, 1, 20)) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            CodEst = getLastIdentity(objConn, "estructura")
            
            Inserto_estr = True
            
            StrSql = " INSERT INTO categoria(estrnro,convnro) VALUES(" & CodEst & "," & nroConv & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub

Public Sub ValidarFormaPago(ByVal FormaPago As String, ByVal Codigo As String, ByVal Moneda As Long, ByRef Nro_FormaPago As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Validacion.
' Autor      : FGZ
' Fecha      : 09/02/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset

    StrSql = " SELECT * FROM formapago WHERE UPPER(fpagdescabr) = '" & Format_Str(UCase(FormaPago), 30, False, "") & "'"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Nro_FormaPago = rs!fpagnro
    Else
    
        StrSql = " INSERT INTO formapago(fpagdescabr,fpagbanc,acunro,monnro) VALUES ("
        StrSql = StrSql & "'" & Format_Str(UCase(FormaPago), 30, False, "") & "'"
        
        Select Case UCase(Codigo)
        Case "A", "I", "T", "U":
            StrSql = StrSql & ",-1"
        Case Else
            StrSql = StrSql & ",0"
        End Select
        StrSql = StrSql & ",6"
        StrSql = StrSql & "," & Moneda
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Nro_FormaPago = getLastIdentity(objConn, "formapago")
    End If
        
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Sub

