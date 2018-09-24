Attribute VB_Name = "MdlInterfacesMigraInicial"
'Global ErrCarga
'Global LineaError
'Global LineaOK


Public Sub Insertar_Linea_Segun_Modelo_MigraInicial(ByVal Linea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador de acurdo al modelo
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim NombreArchivo1 As String
Dim NombreArchivo2 As String
Dim NombreArchivo3 As String
Dim Ok As Boolean

MyBeginTrans

    Ok = True

    Select Case NroModelo
    
' Interfaces de Migracion 0
        
    ' Segun Formato en Plantilla de Formatos Migracion Inicial
    
    Case 600: 'Familiares
        Call LineaModelo_600(Linea)
        
    Case 601: 'Familiares - Goyaike
        Call LineaModelo_601(Linea)
        
    ' Segun Formato en Plantilla de Formatos Migracion Inicial
        
    Case 605: 'Empleados
        Ok = True
        Call LineaModelo_605(Linea, Ok)
        
    Case 610: 'DesmenFamiliar
        Call LineaModelo_610
        
    ' Segun Formato en Plantilla de Formatos Migracion Inicial
    
    Case 615: 'DDJJ
        Call LineaModelo_615(Linea)
        
    ' Segun Formato en Plantilla de Formatos Migracion Inicial
    
    Case 620: 'Desglose de Ganancias - Se hizo para Accor
        Call LineaModelo_620(Linea)
        
    Case 625: 'Liquidaciones - Se hizo para Accor
        Call LineaModelo_625(Linea)
        
    ' Segun Formato en Plantilla de Formatos Migracion Inicial

    Case 630: 'Historico de Estructuras
        Call LineaModelo_630(Linea)
        
    ' Segun Formato en Plantilla de Formatos Migracion Inicial

    Case 635: 'Titulos
        Call LineaModelo_635(Linea)
        
    ' Segun Formato en Plantilla de Formatos Migracion Inicial
        
    Case 640: 'Fases
        Call LineaModelo_640(Linea)
        
    Case 645: 'Acumuladores Mensuales
        Call LineaModelo_645(Linea)
        
    Case 650: 'Empleados CODELCO
        Ok = True
        Call LineaModelo_650(Linea, Ok)
        If Ok Then
            MyCommitTrans
            Call LineaModelo_653(Linea, Ok)
        End If
    Case 651: 'Password CODELCO
        Ok = True
        Call LineaModelo_651(Linea, Ok)
    Case 652: 'Historicos de Estructuras CODELCO
        Ok = True
        Call LineaModelo_652(Linea, Ok)
    Case 653: 'Empreporta CODELCO
        Ok = True
        Call LineaModelo_653(Linea, Ok)
    End Select
    
    If Ok Then
        MyCommitTrans
    Else
        MyRollbackTrans
    End If
    
End Sub


Public Sub ValidarLocalidad(Localidad As String, ByRef nro_localidad As Long, nro_pais As Long, nro_provincia As Long)
Dim rs_sub As New ADODB.Recordset
Dim Sql_Ins As String
Dim SQL_Val As String

If Not EsNulo(Localidad) Then
    StrSql = " SELECT * FROM localidad WHERE UPPER(locdesc) = '" & UCase(Localidad) & "'"
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
        SQL_Val = " VALUES('" & UCase(Localidad) & "'"
    
        If nro_pais <> 0 Then
        
            Sql_Ins = Sql_Ins & ",paisnro"
            SQL_Val = SQL_Val & "," & nro_pais
        
        End If
    
        If nro_provincia <> 0 Then
            Sql_Ins = Sql_Ins & ",provnro"
            SQL_Val = SQL_Val & "," & nro_provincia
        End If
        
        StrSql = Sql_Ins & ")" & SQL_Val & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        nro_localidad = getLastIdentity(objConn, "localidad")
        
    Else
    
        nro_localidad = rs_sub!locnro
    
    End If
End If
End Sub

Public Sub ValidarPartido(Partido As String, ByRef nro_partido As Long)

Dim rs_sub As New ADODB.Recordset

If Not EsNulo(Partido) Then
    StrSql = " SELECT * FROM partido WHERE UPPER(partnom) = '" & UCase(Partido) & "'"
    OpenRecordset StrSql, rs_sub
    
    If rs_sub.EOF Then
    
        StrSql = "INSERT INTO partido(partnom) VALUES('"
        StrSql = StrSql & UCase(Partido) & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        nro_partido = getLastIdentity(objConn, "partido")
        
'        StrSql = " SELECT MAX(partnro) AS MaxPart FROM partido "
'        'StrSql = " SELECT @@IDENTITY AS MaxPart "
'        OpenRecordset StrSql, rs_sub
'
'        nro_partido = rs_sub!MaxPart
    
    Else
        
        nro_partido = rs_sub!partnro
    
    End If
End If
End Sub

Public Sub ValidarZona(Zona As String, ByRef nro_zona As Long, nro_provincia As Long)

Dim rs_sub As New ADODB.Recordset

    If Not EsNulo(Zona) Then
        StrSql = " SELECT * FROM zona WHERE UPPER(zonadesc) = '" & UCase(Zona) & "' AND provnro = " & nro_provincia
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
        
            StrSql = "INSERT INTO zona(zonadesc,provnro) VALUES('"
            StrSql = StrSql & UCase(Zona) & "'," & nro_provincia & ")"
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

Public Sub ValidarProvincia(Provincia As String, ByRef nro_provincia As Long, nro_pais As Long)

Dim rs_sub As New ADODB.Recordset

If Not EsNulo(Provincia) Then
    'StrSql = " SELECT * FROM provincia WHERE provdesc = '" & Provincia & "' AND paisnro = " & nro_pais
    StrSql = " SELECT * FROM provincia WHERE upper(provdesc) = '" & UCase(Provincia) & "'"
    OpenRecordset StrSql, rs_sub
    
    If rs_sub.EOF Then
    
        StrSql = "INSERT INTO provincia(provdesc,paisnro) VALUES('"
        StrSql = StrSql & UCase(Provincia) & "'," & nro_pais & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        nro_provincia = getLastIdentity(objConn, "provincia")
        
        
'        StrSql = " SELECT MAX(provnro) AS MaxProv FROM provincia "
'        'StrSql = " SELECT @@IDENTITY AS MaxProv "
'        OpenRecordset StrSql, rs_sub
'
'        nro_provincia = rs_sub!MaxProv
    
    Else
        
        nro_provincia = rs_sub!provnro
    
    End If
End If
End Sub

Public Sub ValidarPais(Pais As String, ByRef nro_pais As Long)

Dim rs_sub As New ADODB.Recordset

    If Not EsNulo(Pais) Then
        StrSql = " SELECT * FROM pais WHERE UPPER(paisdesc) = '" & UCase(Pais) & "'"
        OpenRecordset StrSql, rs_sub
        
        If rs_sub.EOF Then
        
            StrSql = "INSERT INTO pais(paisdesc,paisdef) VALUES('"
            StrSql = StrSql & UCase(Pais) & "',0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nro_pais = getLastIdentity(objConn, "pais")
        Else
            nro_pais = rs_sub!paisnro
        End If
    End If


End Sub

Public Sub CalcularLegajo(NroEmp As Long, ByRef Legajo As String)

Dim rs_leg As New ADODB.Recordset
Dim rs_emp As New ADODB.Recordset

Dim NroLegajo As Long
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

Public Function Nombre_EstructuraValido(ByVal Tenro As Long, ByRef EstrDesc As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que valida y normaliza ciertos nombre de estructura de Ciertos tipos.
' Autor      : FGZ
' Fecha      : 23/01/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Encontro As Boolean
Dim Tiene As Boolean

Dim rs As New ADODB.Recordset
    
    Encontro = False
    StrSql = " SELECT * FROM confrep WHERE repnro = 120 "
    StrSql = StrSql & " AND conftipo = 'VAL' "
    StrSql = StrSql & " AND confval = " & Tenro
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Tiene = False
    Else
        Tiene = True
    End If
    Do While Not rs.EOF And Not Encontro
        If UCase(EstrDesc) = UCase(rs!confetiq) Then
            EstrDesc = rs!confetiq
            Encontro = True
        End If
    
        rs.MoveNext
    Loop

    If Tiene Then
        Nombre_EstructuraValido = Encontro
    Else
        Nombre_EstructuraValido = True
    End If
    
'Cierro y libero
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
End Function


Public Sub ValidaEstructura(TipoEstr As Long, ByRef Valor As String, ByRef CodEst As Long, ByRef Inserto_estr As Boolean)

Dim Rs_Estr As New ADODB.Recordset

Dim d_estructura As String
Dim CodExt As String
Dim l_pos1 As Long
Dim l_pos2 As Long


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
    
    StrSql = " SELECT estrnro FROM estructura WHERE UPPER(estructura.estrdabr) = '" & UCase(Mid(d_estructura, 1, 60)) & "'"
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    OpenRecordset StrSql, Rs_Estr
        
    If Not Rs_Estr.EOF Then
                
            CodEst = Rs_Estr!Estrnro
            Inserto_estr = False
            
    Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & UCase(Mid(d_estructura, 1, 60)) & "',1,-1,'" & Mid(CodExt, 1, 20) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            CodEst = getLastIdentity(objConn, "estructura")
            
            Inserto_estr = True
    End If


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

    StrSql = " INSERT INTO puesto(estrnro,puedesc,puenroreemp) VALUES(" & CodEstr & ",'" & Valor & "',0)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos10(CodTer As Long, CodEstr As Long, Valor As String)

    StrSql = " INSERT INTO empresa(estrnro,ternro,empnom) VALUES(" & CodEstr & "," & CodTer & ",'" & Valor & "')"
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

    StrSql = "SELECT * FROM osocial  where osdesc = '" & Valor & "'"
    If rs_17.State = adStateOpen Then rs_17.Close
    OpenRecordset StrSql, rs_17
    
    If rs_17.EOF Then
        StrSql = " INSERT INTO osocial(ternro,osdesc) VALUES(" & CodTer & ",'" & Valor & "')"
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
Dim rs_tipocont As New ADODB.Recordset
Dim rs_TC As New ADODB.Recordset
Dim CodTC As Long


    
    StrSql = "SELECT * FROM tipocont  where tcdabr = '" & Valor & "'"
    OpenRecordset StrSql, rs_tipocont
    
    If rs_tipocont.EOF Then
        StrSql = " INSERT INTO tipocont(tcdabr,estrnro,tcind) VALUES('" & Valor & "'," & CodEstr & ",-1)"
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

    StrSql = " INSERT INTO planos(plnom,osocial) VALUES('" & Valor & "'," & CodTer & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    CodPlan = getLastIdentity(objConn, "planos")
    
    StrSql = " INSERT INTO replica_estr(origen,estrnro) VALUES (" & CodPlan & "," & CodEstr & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    

End Sub

Public Sub Complementos40(CodEstr As Long, CodTer As Long, Valor As String)

    StrSql = " INSERT INTO seguro(ternro,estrnro,segdesc,segest) VALUES(" & CodEstr & "," & CodTer & ",'" & Valor & "',-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Complementos41(CodEstr As Long, CodTer As Long, Valor As String)

    StrSql = " INSERT INTO banco(ternro,estrnro,bansucdesc,banest) VALUES(" & CodEstr & "," & CodTer & ",'" & Valor & "',-1)"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub CreaTercero(TipoTer As Long, Valor As String, ByRef CodTer)

Dim rs As New ADODB.Recordset
Dim rs_Ter As New ADODB.Recordset

Dim d_estructura As String
Dim l_pos1 As Long
Dim l_pos2 As Long

    
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

Public Sub ValidaEstructuraCodExt(TipoEstr As Long, ByRef Valor As String, ByRef CodEst As Long, ByRef Inserto_estr As Boolean)

Dim Rs_Estr As New ADODB.Recordset

Dim d_estructura As String
Dim CodExt As String
Dim l_pos1 As Byte
Dim l_pos2 As Byte


    d_estructura = Valor
    StrSql = " SELECT estrnro FROM estructura WHERE upper(estructura.estrcodext) = '" & UCase(Mid(Valor, 1, 20)) & "'"
    StrSql = StrSql & " AND estructura.tenro = " & TipoEstr
    OpenRecordset StrSql, Rs_Estr
        
    If Not Rs_Estr.EOF Then
            CodEst = Rs_Estr!Estrnro
            Inserto_estr = False
    Else
            StrSql = " INSERT INTO estructura(tenro,estrdabr,empnro,estrest,estrcodext)"
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & UCase(Mid(d_estructura, 1, 60)) & "',1,-1,'" & UCase(Mid(d_estructura, 1, 20)) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            CodEst = getLastIdentity(objConn, "estructura")

            Inserto_estr = True
    End If

End Sub

Public Sub ValidaCategoria(TipoEstr As Long, ByRef Valor As String, nroConv As Long, ByRef CodEst As Long, ByRef Inserto_estr As Boolean)
Dim pos1 As Long
Dim pos2 As Long

Dim Rs_Estr As New ADODB.Recordset
Dim Rs_Conv As New ADODB.Recordset
Dim Rs_NroC As New ADODB.Recordset


Dim d_estructura As String
Dim l_pos1 As Byte
Dim l_pos2 As Byte
Dim CodExt As String

Dim nro_convenio As Long

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
    
    
            
    StrSql = " SELECT estrnro FROM estructura WHERE UPPER(estructura.estrdabr) = '" & UCase(Mid(d_estructura, 1, 60)) & "'"
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
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & UCase(Mid(d_estructura, 1, 60)) & "',1,-1,'" & UCase(Mid(CodExt, 1, 20)) & "')"
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


Public Sub ValidaCategoriaCodExt(TipoEstr As Long, ByRef Valor As String, nroConv As Long, ByRef CodEst As Long, ByRef Inserto_estr As Boolean)
Dim pos1 As Long
Dim pos2 As Long

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
    
    
    StrSql = " SELECT estrnro FROM estructura WHERE UPPER(estructura.estrcodext) = '" & UCase(Mid(d_estructura, 1, 20)) & "'"
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
            StrSql = StrSql & " VALUES(" & TipoEstr & ",'" & UCase(Mid(d_estructura, 1, 60)) & "',1,-1," & UCase(Mid(CodExt, 1, 20)) & ")"
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


Public Sub AsignarEstructura(TipoEstr As Long, CodEst As Long, CodTer As Long, FAlta As String, FBaja As String)
Dim rs As New ADODB.Recordset
Dim rs_his As New ADODB.Recordset

Dim FCierre As Date

    If CodEst <> 0 Then
    
        If nro_ModOrg <> 0 Then
        
            StrSql = "SELECT * FROM adptte_estr WHERE tplatenro = " & nro_ModOrg & " AND tenro = " & TipoEstr
            OpenRecordset StrSql, rs
            
            If rs.EOF Then
            
                tplaorden = tplaorden + 1
                StrSql = "INSERT INTO adptte_estr(tplatenro,tenro,tplaestroblig,tplaestrorden) VALUES (" & nro_ModOrg & "," & TipoEstr & ",0," & tplaorden & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            
            End If
    
        End If
    
        StrSql = "SELECT * FROM his_estructura "
        StrSql = StrSql & " WHERE ternro = " & CodTer
        StrSql = StrSql & " AND tenro = " & TipoEstr
        StrSql = StrSql & " AND htethasta IS NULL "
        StrSql = StrSql & " ORDER BY htetdesde "
        OpenRecordset StrSql, rs_his
        
        If Not rs_his.EOF Then
        
            If rs_his!Estrnro <> CodEst Then
        
                FCierre = CDate(Mid(FAlta, 2, 2) & "/" & Mid(FAlta, 5, 2) & "/" & Mid(FAlta, 8, 4))
        
                FCierre = FCierre - 1
                
                If ConvFecha(FCierre) < rs_his!htetdesde Then
                    
                    ErrCarga.Writeline "Las Fechas se Superponen"
                    Exit Sub
        
                End If
        
                StrSql = " UPDATE his_estructura SET htethasta = " & ConvFecha(FCierre)
                StrSql = StrSql & " WHERE tenro = " & TipoEstr & " AND ternro = " & CodTer & " AND estrnro = " & rs_his!Estrnro
                StrSql = StrSql & " AND htethasta IS NULL "
                objConn.Execute StrSql, , adExecuteNoRecords
                        
                If Not FBaja = "Null" Then
                    StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                    StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & FAlta & "," & FBaja & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
                    StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & FAlta & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
        
        
            Else
        
                If Pisa Then
                    If Not FBaja = "Null" Then
                    
                        StrSql = " UPDATE his_estructura SET htetdesde = " & FAlta
                        StrSql = StrSql & ",htethasta = " & FBaja
                        StrSql = StrSql & " WHERE tenro = " & TipoEstr & " AND ternro = " & CodTer & " AND estrnro = " & CodEst
                        StrSql = StrSql & " AND htethasta IS NULL "
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                    Else
                        
                        StrSql = " UPDATE his_estructura SET htetdesde = " & FAlta
                        StrSql = StrSql & " WHERE tenro = " & TipoEstr & " AND ternro = " & CodTer & " AND estrnro = " & CodEst
                        StrSql = StrSql & " AND htethasta IS NULL "
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                    End If
                End If
        
            End If
            
        Else
        
            If Not FBaja = "Null" Then
            
                StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & FAlta & "," & FBaja & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
            Else
                StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
                StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & FAlta & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
            End If
            
        End If

    End If

    If rs_his.State = adStateOpen Then rs_his.Close
    
    Set rs_his = Nothing
    
End Sub


Public Sub AsignarEstructura_NEW(TipoEstr As Long, CodEst As Long, CodTer As Long, FAlta As String, FBaja As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta la estructura. si existe una estructura del mismo tipo en el intervalo
'               la estructura será actualizada.
' Autor      : FGZ
' Fecha      : 22/04/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
Dim rs_his As New ADODB.Recordset

    If CodEst <> 0 Then
        If nro_ModOrg <> 0 Then
            StrSql = "SELECT * FROM adptte_estr WHERE tplatenro = " & nro_ModOrg & " AND tenro = " & TipoEstr
            OpenRecordset StrSql, rs
            If rs.EOF Then
                tplaorden = tplaorden + 1
                StrSql = "INSERT INTO adptte_estr(tplatenro,tenro,tplaestroblig,tplaestrorden) VALUES (" & nro_ModOrg & "," & TipoEstr & ",0," & tplaorden & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    
        StrSql = "SELECT * FROM his_estructura "
        'StrSql = StrSql & " WHERE estrnro =" & CodEst
        StrSql = StrSql & " WHERE tenro = " & TipoEstr
        StrSql = StrSql & " AND ternro = " & CodTer
        StrSql = StrSql & " AND (htetdesde <= " & FAlta & ") AND"
        StrSql = StrSql & " ((" & FAlta & " <= htethasta) or (htethasta is null))"
        StrSql = StrSql & " ORDER BY htetdesde "
        If rs_his.State = adStateOpen Then rs_his.Close
        OpenRecordset StrSql, rs_his
        If Not rs_his.EOF Then
            If Pisa Then
                If rs_his!Estrnro = CodEst Then
                    'If Not FBaja = "Null" Then
                        StrSql = " UPDATE his_estructura SET htetdesde = " & FAlta
                        StrSql = StrSql & ",htethasta = " & FBaja
                        StrSql = StrSql & " WHERE tenro = " & TipoEstr
                        StrSql = StrSql & " AND ternro = " & CodTer
                        StrSql = StrSql & " AND estrnro = " & rs_his!Estrnro
                        StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_his!htetdesde)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    'Else
                    '    StrSql = " UPDATE his_estructura SET htetdesde = " & FAlta
                    '    objConn.Execute StrSql, , adExecuteNoRecords
                    'End If
                Else
                    'If Not FBaja = "Null" Then
                        StrSql = " UPDATE his_estructura SET "
                        StrSql = StrSql & " estrnro = " & CodEst
                        StrSql = StrSql & ",htetdesde = " & FAlta
                        StrSql = StrSql & ",htethasta = " & FBaja
                        StrSql = StrSql & " WHERE tenro = " & TipoEstr
                        StrSql = StrSql & " AND ternro = " & CodTer
                        StrSql = StrSql & " AND estrnro = " & rs_his!Estrnro
                        StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_his!htetdesde)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    'Else
                    '    StrSql = " UPDATE his_estructura SET htetdesde = " & FAlta
                    '    StrSql = StrSql & ",estrnro = " & CodEst
                    '    objConn.Execute StrSql, , adExecuteNoRecords
                    'End If
                End If
            Else
                Texto = ": " & "Ya existe una estructura de tipo " & TipoEstr
                Call Escribir_Log("floge", LineaCarga, 1, Texto, Tabs, "")
            End If
        Else
            If Not FBaja = "Null" Then
                StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & FAlta & "," & FBaja & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
                StrSql = StrSql & CodTer & "," & CodEst & "," & TipoEstr & "," & FAlta & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    End If
    
    If rs_his.State = adStateOpen Then rs_his.Close
    Set rs_his = Nothing
End Sub




Public Sub LineaModelo_600(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Familiares
' Autor      : MAB
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As Long    ' Legajo del Empleado
Dim Apellido        As String  ' Apellido del Familiar
Dim nombre          As String  ' Nombre del Familiar
Dim Fnac            As String  ' Fecha de Nacimiento del Familiar
Dim NAC             As String  ' Nacionalidad del Familiar
Dim PaisNac         As String  ' Pais de Nacimiento
Dim EstCiv          As String  ' Estado Civil
Dim Sexo            As String  ' Sexo del Familiar
Dim GPare           As String  ' Grado de Parentesco
Dim Disc            As String  ' Discapacitado
Dim Estudia         As String  ' Estudia
Dim NivEst          As String  ' Nivel de Estudio
Dim TipDoc          As String  ' Tipo de Documento del Familiar
Dim NroDoc          As String  ' Nº de Documento del Familiar
Dim Calle           As String   'Calle                    -- detdom.calle
Dim Nro             As String   'Número                   -- detdom.nro
Dim Piso            As String   'Piso                     -- detdom.piso
Dim Depto           As String   'Depto                    -- detdom.depto
Dim Torre           As String   'Torre                    -- detdom.torre
Dim Manzana         As String   'Manzana                  -- detdom.manzana
Dim Cpostal         As String   'Cpostal                  -- detdom.codigopostal
Dim Entre           As String   'Entre Calles             -- detdom.entrecalles
Dim Barrio          As String   'Barrio                   -- detdom.barrio
Dim Localidad       As String   'Localidad                -- detdom.locnro
Dim Partido         As String   'Partido                  -- detdom.partnro
Dim Zona            As String   'Zona                     -- detdom.zonanro
Dim Provincia       As String   'Provincia                -- detdom.provnro
Dim Pais            As String   'Pais                     -- detdom.paisnro
Dim Telefono        As String   'Telefono                 -- telefono.telnro
Dim ObraSocial      As String   'Obra Social
Dim PlanOSocial     As String   'Plan Obra Social
Dim AvisoEmer       As String   'Aviso ante Emergencia
Dim PagaSalario     As String   'Paga Salario Familiar
Dim Ganancias       As String   'Se lo toma para ganancias

Dim Cuil            As String  ' CUIL del Familiar
Dim ESC             As String  ' Escolaridad
Dim GRADO           As String  ' Grado al que concurre
Dim NroTDoc         As String

Dim pos1            As Long
Dim pos2            As Long

Dim NroTercero      As Long
Dim NroEmpleado     As Long
Dim CodTerFam       As String
Dim nro_nrodom      As Long
Dim nro_nacionalidad As Long
Dim nro_paisnac      As Long
Dim nro_estciv      As Long
Dim nro_Sexo        As Long
Dim nro_estudia     As Long
Dim nro_osocial     As Long
Dim nro_planos      As Long
Dim nro_aviso       As Long
Dim nro_salario     As Long
Dim nro_gan         As Long
Dim nro_disc        As Long
Dim nro_paren        As Long
Dim nro_barrio          As Long
Dim nro_localidad       As Long
Dim nro_partido         As Long
Dim nro_zona            As Long
Dim nro_provincia       As Long
Dim nro_pais            As Long
Dim OSocial             As String
Dim ter_osocial         As Long
Dim Inserto_estr        As Boolean

Dim IngresoDom          As Boolean

Dim StrSql          As String
Dim rs              As New ADODB.Recordset
Dim rs_sql              As New ADODB.Recordset


    RegLeidos = RegLeidos + 1
    LineaCarga = LineaCarga + 1
    
    On Error GoTo SaltoLinea
    
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
    
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    If IsNumeric(Trim(Mid(strReg, pos1, pos2 - pos1))) Then
        Legajo = Trim(Mid(strReg, pos1, pos2 - pos1))
    Else
        Flog.Writeline Espacios(Tabulador * 1) & "El legajo no es numerico "
        'FlogE.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El legajo no es numerico"
        ErrCarga.Writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El legajo no es numerico"
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Apellido = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    nombre = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Fnac = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If Fnac = "N/A" Or Fnac = " " Then
        Fnac = "''"
    Else
       Fnac = ConvFecha(Fnac)
    End If
    
'    pos1 = pos2 + 1
'    pos2 = InStr(pos1 + 1, strReg, Separador)
'    PaisNac = Trim(Mid(strReg, pos1, pos2 - pos1))
'    StrSql = " SELECT paisnro FROM pais WHERE paisdesc = '" & PaisNac & "'"
'    OpenRecordset StrSql, rs
'
'    If Not rs.EOF Then
'        nro_paisnac = rs!paisnro
'    Else
'        nro_paisnac = 0
'    End If
'    rs.Close
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NAC = UCase(Trim(Mid(strReg, pos1, pos2 - pos1)))
    StrSql = " SELECT nacionalnro FROM nacionalidad WHERE upper(nacionaldes) = '" & NAC & "'"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        nro_nacionalidad = rs!nacionalnro
    Else
        nro_nacionalidad = 0
    End If
    rs.Close
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    EstCiv = UCase(Trim(Mid(strReg, pos1, pos2 - pos1)))
    StrSql = " SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(EstCiv) & "'"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        nro_estciv = rs!estcivnro
    Else
        nro_estciv = 0
    End If
    rs.Close
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Sexo = UCase(Trim(Mid(strReg, pos1, pos2 - pos1)))
    If Sexo = "M" Or Sexo = "MASCULINO" Then
        nro_Sexo = -1
    Else
        nro_Sexo = 0
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    GPare = UCase(Trim(Mid(strReg, pos1, pos2 - pos1)))
    StrSql = " SELECT parenro FROM parentesco WHERE upper(paredesc) = '" & UCase(GPare) & "'"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        nro_paren = rs!parenro
    Else
        nro_paren = 0
    End If
    rs.Close
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Disc = UCase(Trim(Mid(strReg, pos1, pos2 - pos1)))
    If Disc = "N/A" Or Disc = "NO" Then
        nro_disc = 0
    Else
        nro_disc = -1
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Estudia = UCase(Trim(Mid(strReg, pos1, pos2 - pos1)))
    If Estudia = "N/A" Or Estudia = "NO" Then
        nro_estudia = 0
    Else
        nro_estudia = -1
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NivEst = Trim(Mid(strReg, pos1, pos2 - pos1))
' Por ahora no hago nada con el nivel de estudio porque en Accor no lo pasaron

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    TipDoc = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    
    If TipDoc <> "N/A" Then
        StrSql = " SELECT tidnro FROM tipodocu WHERE UPPER(tidsigla) = '" & UCase(TipDoc) & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_tdocumento = rs_sql!tidnro
        Else
            StrSql = " INSERT INTO tipodocu(tidnom,tidsigla,tidsist,instnro,tidunico) VALUES ('" & UCase(Tipodoc) & "','" & UCase(Tipodoc) & "',0,0,0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_tdocumento = getLastIdentity(objConn, "tipodocu")
            
        End If
    Else
        nro_tdocumento = 0
    End If
    
'    If nro_tdocumento = 0 Then
'        LineaError.Writeline Mid(strReg, 1, Len(strReg))
'        ErrCarga.Writeline "Linea: " & LineaCarga & " - El Tipo de Documento no Existe."
'        Ok = False
'        Exit Sub
'    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NroDoc = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If NroDoc = "N/A" Then
        NroDoc = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Calle = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    IngresoDom = True
    
    If Calle = "N/A" Then
        Calle = ""
        IngresoDom = False
        
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Nro = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If (Nro <> "N/A") Then
        nro_nrodom = Nro
    Else
        nro_nrodom = 0
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Piso = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If Piso = "N/A" Then
        Piso = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Depto = Trim(Mid(strReg, pos1, pos2 - pos1))

    If Depto = "N/A" Then
        Depto = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Torre = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If Torre = "N/A" Then
        Torre = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Manzana = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If Manzana = "N/A" Then
        Manzana = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Cpostal = Trim(Mid(strReg, pos1, pos2 - pos1))

    If Cpostal = "N/A" Then
        Cpostal = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Entre = Trim(Mid(strReg, pos1, pos2 - pos1))

    If Entre = "N/A" Then
        Entre = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Barrio = Trim(Mid(strReg, pos1, pos2 - pos1))

    If Barrio = "N/A" Then
        Barrio = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Localidad = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Partido = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Zona = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Provincia = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Pais = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If Pais <> "N/A" Then
        Call ValidarPais(Pais, nro_pais)
    Else
        nro_pais = 0
    End If
    
    If (Provincia <> "N/A") And (Pais <> "N/A") Then
        Call ValidarProvincia(Provincia, nro_provincia, nro_pais)
    Else
        nro_provincia = 0
    End If
    
    If (Localidad <> "N/A") And (Provincia <> "N/A") And (Pais <> "N/A") Then
        Call ValidarLocalidad(Localidad, nro_localidad, nro_pais, nro_provincia)
    Else
        nro_localidad = 0
    End If
    
    If Partido <> "N/A" Then
        Call ValidarPartido(Partido, nro_partido)
    Else
        nro_partido = 0
    End If
    
    If Zona <> "N/A" Then
        Call ValidarZona(Zona, nro_zona, nro_provincia)
    Else
        nro_zona = 0
    End If
    
    If (IngresoDom = True) And (nro_localidad = 0) Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - Debe Ingresar la Localidad."
        Ok = False
        Exit Sub
    End If
    
    If (IngresoDom = True) And (nro_provincia = 0) Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - Debe Ingresar la Provincia."
        Ok = False
        Exit Sub
    End If
    
    If (IngresoDom = True) And (nro_pais = 0) Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - Debe Ingresar la Pais."
        Ok = False
        Exit Sub
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Telefono = Mid(strReg, pos1, pos2 - pos1)
    
    If Telefono = "N/A" Then
        Telefono = ""
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    ObraSocial = Trim(Mid(strReg, pos1, pos2 - pos1))
    If ObraSocial = "N/A" Or ObraSocial = "" Then
        nro_osocial = 0
    Else
        StrSql = " SELECT ternro FROM osocial WHERE UPPER(osdesc) = '" & UCase(ObraSocial) & "'"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            nro_osocial = rs!ternro
        Else
            nro_osocial = 0
        End If
        rs.Close
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    PlanOSocial = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    If PlanOSocial = "N/A" Or PlanOSocial = "" Then
        nro_planos = 0
    Else
        If nro_osocial <> 0 Then
            StrSql = " SELECT plnro FROM planos WHERE UPPER(plnom) = '" & UCase(PlanOSocial) & "'"
            StrSql = StrSql & " AND osocial = " & nro_osocial
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                nro_planos = rs!plnro
            Else
                nro_planos = 0
            End If
            rs.Close
        Else
            nro_planos = 0
        End If
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    AvisoEmer = Trim(Mid(strReg, pos1, pos2 - pos1))
    If AvisoEmer = "N/A" Or AvisoEmer = "NO" Then
        nro_aviso = 0
    Else
        nro_aviso = -1
    End If

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    PagaSalario = Trim(Mid(strReg, pos1, pos2 - pos1))
    If PagaSalario = "N/A" Or PagaSalario = "NO" Then
        nro_salario = 0
    Else
        nro_salario = -1
    End If

    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Ganancias = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Ganancias = "N/A" Or Ganancias = "NO" Then
        nro_gan = 0
    Else
        nro_gan = -1
    End If

' Busco el empleado asociado

  StrSql = "SELECT ternro FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, rs
  NroEmpleado = rs!ternro

  If rs.State = adStateOpen Then
    rs.Close
  End If

' Inserto el tercero asociado al familiar

  StrSql = " INSERT INTO tercero(ternom,terape,terfecnac,tersex,nacionalnro,paisnro,estcivnro)"
  StrSql = StrSql & " VALUES('" & nombre & "','" & Apellido & "'," & Fnac & "," & nro_Sexo & ","
  If nro_nacionalidad <> 0 Then
    StrSql = StrSql & nro_nacionalidad & ","
  Else
    StrSql = StrSql & "Null,"
  End If
  If nro_paisnac <> 0 Then
    StrSql = StrSql & nro_paisnac & ","
  Else
    StrSql = StrSql & "Null,"
  End If
  StrSql = StrSql & nro_estciv & ")"
  objConn.Execute StrSql, , adExecuteNoRecords

  NroTercero = getLastIdentity(objConn, "tercero")
  
  Flog.Writeline "Codigo de Tercero-Familiar = " & NroTercero

' Inserto el Familiar

  StrSql = " INSERT INTO familiar(empleado,ternro,parenro,famest,famestudia,famcernac,faminc,famsalario,famemergencia,famcargadgi,osocial,plnro,famternro)"
  StrSql = StrSql & " values(" & NroEmpleado & "," & NroTercero & "," & nro_paren & ",-1," & nro_estudia & ",0," & nro_disc & "," & nro_salario & "," & nro_aviso & "," & nro_gan & "," & nro_osocial & "," & nro_planos & ",0)"
  objConn.Execute StrSql, , adExecuteNoRecords

  Flog.Writeline "Inserte el Familiar - " & Legajo & " - " & Apellido & " - " & nombre

' Inserto el Registro correspondiente en ter_tip

  StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & NroTercero & ",3)"
  objConn.Execute StrSql, , adExecuteNoRecords

' Inserto los Documentos
  If NroDoc <> "" And NroDoc <> "N/A" And TipDoc <> "N/A" Then
    StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
    StrSql = StrSql & " VALUES(" & NroTercero & "," & nro_tdocumento & ",'" & NroDoc & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el DU - "
  End If
  
  If rs.State = adStateOpen Then rs.Close
  
' Inserto el Domicilio
  
  If Not IngresoDom = False Then
      StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
      StrSql = StrSql & " VALUES(1," & NroTercero & ",-1,2)"
      objConn.Execute StrSql, , adExecuteNoRecords
      
      CodDom = getLastIdentity(objConn, "cabdom")
      
      StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,"
      StrSql = StrSql & "locnro,provnro,paisnro,barrio,partnro,zonanro) "
      StrSql = StrSql & " VALUES (" & CodDom & ",'" & Calle & "','" & Nro & "','" & Piso & "','"
      StrSql = StrSql & Depto & "','" & Torre & "','" & Manzana & "','" & Cpostal & "'," & nro_localidad & ","
      StrSql = StrSql & nro_provincia & "," & nro_pais & ",'" & Barrio & "'," & nro_partido & "," & nro_zona & ")"
      objConn.Execute StrSql, , adExecuteNoRecords
      
      Flog.Writeline "Inserte el Domicilio - "
      
      If Telefono <> "" Then
        StrSql = " INSERT INTO telefono(domnro,telnro,teldefault) "
        StrSql = StrSql & " VALUES(" & CodDom & ",'" & Telefono & "',-1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline "Inserte el Telefono - "
      End If
      
  End If
  
  LineaOK.Writeline Mid(strReg, 1, Len(strReg))
  Ok = True
         
  If rs.State = adStateOpen Then
      rs.Close
  End If

  Exit Sub

SaltoLinea:

    LineaError.Writeline Mid(strReg, 1, Len(strReg))
    ErrCarga.Writeline "Linea: " & LineaCarga & " - " & Err.Description
    MyRollbackTrans
    Ok = False
  
End Sub

Public Sub LineaModelo_601(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Familiares Customizada - Goyaike
' Autor      : MAB
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As Long ' Legajo del Empleado
Dim Apellido        As String  ' Apellido del Familiar
Dim nombre          As String  ' Nombre del Familiar
Dim NroOSL          As String
Dim NroOSE          As String
Dim OSE             As String
Dim PlanOSE         As String
Dim PlanOdon        As String
Dim Beca            As String
Dim FPC             As String
Dim Seguro          As String

Dim pos1            As Long
Dim pos2            As Long

Dim NroTercero      As Long
Dim NroEmpleado     As Long
Dim NroFamiliar     As Long
Dim CodTerFam       As String
Dim nro_seg             As Long
Dim Inserto_estr        As Boolean

Dim StrSql          As String
Dim rs              As New ADODB.Recordset

'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
    
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Apellido = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    nombre = Trim(Mid(strReg, pos1, pos2 - pos1))

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NroOSL = Trim(Mid(strReg, pos1, pos2 - pos1))
    If NroOSL = "N/A" Then
        NroOSL = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    OSE = Trim(Mid(strReg, pos1, pos2 - pos1))
    If OSE = "N/A" Then
        OSE = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NroOSE = Trim(Mid(strReg, pos1, pos2 - pos1))
    If NroOSE = "N/A" Then
        NroOSE = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    PlanOSE = Trim(Mid(strReg, pos1, pos2 - pos1))
    If PlanOSE = "N/A" Then
        PlanOSE = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    PlanOdon = Trim(Mid(strReg, pos1, pos2 - pos1))
    If PlanOdon = "N/A" Then
        PlanOdon = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Beca = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Beca = "N/A" Then
        Beca = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    FPC = Trim(Mid(strReg, pos1, pos2 - pos1))
    If FPC = "N/A" Then
        FPC = ""
    End If
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Seguro = Trim(Mid(strReg, pos1, pos2 - pos1))
    If Seguro = "N/A" Or Seguro = "NO" Then
        nro_seg = 0
    Else
        nro_seg = -1
    End If

' Busco el empleado asociado

  StrSql = "SELECT ternro FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, rs
  NroEmpleado = rs!ternro

  If rs.State = adStateOpen Then
    rs.Close
  End If
  
' Busco al familiar por el nombre y apellido

  StrSql = "SELECT familiar.ternro FROM familiar "
  StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = familiar.ternro "
  StrSql = StrSql & " WHERE familiar.empleado = " & NroEmpleado
  StrSql = StrSql & " AND tercero.terape = '" & Apellido & "'"
  StrSql = StrSql & " AND tercero.ternom = '" & nombre & "'"
  OpenRecordset StrSql, rs
  
  NroFamiliar = 0
  
  If Not rs.EOF Then
    NroFamiliar = rs!ternro
    ' Inserto las Notas
    If NroOSL <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",26,'" & NroOSL & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto NroOSL"
    End If
    If OSE <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",27,'" & OSE & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto OSE"
    End If
    If NroOSE <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",6,'" & NroOSE & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto NroOSE"
    End If
    If PlanOSE <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",28,'" & PlanOSE & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto PlanOSE"
    End If
    If PlanOdon <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",29,'" & PlanOdon & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto PlanOdon"
    End If
    If Beca <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",30,'" & Beca & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto Beca"
    End If
    If FPC <> "" Then
      StrSql = " INSERT INTO notas_ter(ternro,tnonro,notatxt)"
      StrSql = StrSql & " VALUES(" & NroFamiliar & ",31,'" & FPC & "')"
      objConn.Execute StrSql, , adExecuteNoRecords
      Flog.Writeline Apellido & ", " & nombre & " - Inserto FPC"
    End If
  End If
  
  If rs.State = adStateOpen Then
    rs.Close
  End If
  
  If NroFamiliar <> 0 Then
    ' Asigno Benef. Seguro de Vida
    StrSql = "UPDATE familiar SET fambensegvida = " & nro_seg
    StrSql = StrSql & " WHERE familiar.ternro = " & NroFamiliar
    objConn.Execute StrSql, , adExecuteNoRecords
  End If
  
  If rs.State = adStateOpen Then
    rs.Close
  End If

End Sub




Public Sub LineaModelo_605(ByVal strReg As String, ByRef Ok As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Empleados
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Long
Dim pos2            As Long

Dim Legajo          As String   'LEGAJO                   -- empleado.empleg
Dim Apellido        As String   'APELLIDO                 -- empleado.terape y tercero.terape
Dim nombre          As String   'NOMBRE                   -- empleado.ternom y tercero.ternom
Dim Fnac            As String   'FNAC                     -- tercero.terfecna
Dim Nacionalidad    As String   'Nacionalidad             -- tercero.nacionalnro
Dim PaisNac         As String   'Pais de Nacimiento       -- tercero.paisnro
Dim Fing            As String   'Fec.Ingreso al Pais      -- terecro.terfecing
Dim EstCivil        As String   'Est.Civil                -- tercero.estcivnro
Dim Sexo            As String   'Sexo                     -- tercero.tersex
Dim FAlta           As String   'Fec. de Alta             -- empleado.empfaltagr y fases.altfec
Dim Estudio         As String   'Estudia                  -- empleado.empestudia
Dim NivEst          As String   'Nivel de Estudio         -- empleado.nivnro
Dim Tdocu           As String   'Tipo Documento           -- ter_dpc.tidnro (DU)
Dim Ndocu           As String   'Nro. Documento           -- ter_doc.nrodoc
Dim Cuil            As String   'CUIL                     -- ter_doc.nrodoc (10)
Dim Calle           As String   'Calle                    -- detdom.calle
Dim Nro             As String   'Número                   -- detdom.nro
Dim Piso            As String   'Piso                     -- detdom.piso
Dim Depto           As String   'Depto                    -- detdom.depto
Dim Torre           As String   'Torre                    -- detdom.torre
Dim Manzana         As String   'Manzana                  -- detdom.manzana
Dim Cpostal         As String   'Cpostal                  -- detdom.codigopostal
Dim Entre           As String   'Entre Calles             -- detdom.entrecalles
Dim Barrio          As String   'Barrio                   -- detdom.barrio
Dim Localidad       As String   'Localidad                -- detdom.locnro
Dim Partido         As String   'Partido                  -- detdom.partnro
Dim Zona            As String   'Zona                     -- detdom.zonanro
Dim Provincia       As String   'Provincia                -- detdom.provnro
Dim Pais            As String   'Pais                     -- detdom.paisnro
Dim Telefono        As String   'Telefono                 -- telefono.telnro
Dim TelLaboral        As String   'Telefono                 -- telefono.telnro
Dim TelCelular        As String   'Telefono                 -- telefono.telnro
Dim Email           As String   'E-mail                   -- empleado.empemail
Dim Sucursal        As String   'Sucursal                 -- his_estructura.estrnro
Dim Sector          As String   'Sector                   -- his_estructura.estrnro
Dim categoria       As String   'Categoria                -- his_estructura.estrnro
Dim Puesto          As String   'Puesto                   -- his_estructura.estrnro
Dim CCosto          As String   'C.Costo                  -- his_estructura.estrnro
Dim Gerencia        As String   'Gerencia                 -- his_estructura.estrnro
Dim Departamento    As String   'Departamento             -- his_estructura.estrnro
Dim Direccion       As String   'Direccion                -- his_estructura.estrnro
Dim CajaJub         As String   'Caja de Jubilacion       -- his_estructura.estrnro
Dim Sindicato       As String   'Sindicato                -- his_estructura.estrnro
Dim OSocialLey         As String   'Obra Social              -- his_estructura.estrnro
Dim PlanOSLey          As String   'Plan OS                  -- his_estructura.estrnro
Dim OSocialElegida         As String   'Obra Social              -- his_estructura.estrnro
Dim PlanOSElegida          As String   'Plan OS                  -- his_estructura.estrnro
Dim Contrato        As String   'Contrato                 -- his_estructura.estrnro
Dim Convenio        As String   'Convenio                 -- his_estructura.estrnro
Dim LPago           As String   'Lugar de Pago            -- his_estructura.estrnro
Dim RegHorario      As String   'Regimen Horario          -- his_estructura.estrnro
Dim FormaLiq        As String   'Forma de Liquidacion     -- his_estructura.estrnro
Dim FormaPago       As String   'Forma de Pago            -- formapago.fpagdescabr
Dim SucBanco        As String   'Sucursal del Banco       -- ctabancaria.ctabsuc
Dim BancoPago       As String   'Banco Pago               -- his_estructura.estrnro, formapago.fpagbanc (siempre y cuando el Banco sea <> 0) y ctabancaria.banco
Dim NroCuenta       As String   'Nro. Cuenta              -- ctabancario.ctabnro
Dim Actividad       As String   'Actividad                -- his_estructura.estrnro
Dim CondSIJP        As String   'Condicion SIJP           -- his_estructura.estrnro
Dim SitRev          As String   'Sit. de Revista SIJP     -- his_estructura.estrnro
Dim ModCont         As String   'Mod. de Contrat. SIJP    -- his_estructura.estrnro
Dim ART             As String   'ART                      -- his_estructura.estrnro
Dim Estado          As String   'Estado                   -- empleado.empest y fases.estado
Dim CausaBaja       As String   'Causa de Baja            -- fases.caunro
Dim FBaja           As String   'Fecha de Baja            -- fases.bajfec
Dim Empresa         As String   'Empresa                  -- his_estructura.estrnro
Dim ModOrg          As String   'Empresa                  -- his_estructura.estrnro
Dim OSL             As String   'Empresa                  -- his_estructura.estrnro
Dim OSE             As String   'Empresa                  -- his_estructura.estrnro
Dim PlanOdon        As String   'Empresa                  -- his_estructura.estrnro
Dim Locacion        As String   'Empresa                  -- his_estructura.estrnro
Dim Area            As String   'Empresa                  -- his_estructura.estrnro
Dim SubDepto        As String   'Empresa                  -- his_estructura.estrnro
Dim NroCBU          As String   'Empresa                  -- his_estructura.estrnro
Dim Empremu         As String   'Remuneración del empleado

Dim ternro As Long

Dim NroTercero          As Long

Dim Nro_Legajo          As Long
Dim nro_tdocumento      As Long
Dim nro_nivest          As Long
Dim nro_estudio         As Long

'Dim nro_nrodom          as long
Dim nro_nrodom          As String

Dim nro_barrio          As Long
Dim nro_localidad       As Long
Dim nro_partido         As Long
Dim nro_zona            As Long
Dim nro_provincia       As Long
Dim nro_pais            As Long
Dim nro_paisnac         As Long

Dim nro_sucursal        As Long
Dim nro_sector          As Long
Dim nro_categoria       As Long
Dim nro_puesto          As Long
Dim nro_ccosto          As Long
Dim nro_gerencia        As Long
Dim nro_cajajub         As Long
Dim nro_sindicato       As Long
Dim nro_osocial_ley     As Long
Dim nro_planos_ley      As Long
Dim nro_osocial_elegida As Long
Dim nro_planos_elegida  As Long
Dim nro_contrato        As Long
Dim nro_convenio        As Long
Dim nro_reghorario      As Long
Dim nro_formaliq        As Long
Dim nro_bancopago       As Long
Dim nro_actividad       As Long
Dim nro_sitrev          As Long
Dim nro_modcont         As Long
Dim nro_art             As Long
Dim nro_departamento    As Long
Dim nro_direccion       As Long
Dim nro_lpago           As Long
Dim nro_condsijp        As Long
Dim nro_formapago       As Long
Dim nro_causabaja       As Long
Dim nro_empresa         As Long
Dim NroDom              As Long
Dim nro_osl             As Long
Dim nro_odon            As Long
Dim nro_ose             As Long
Dim nro_locacion        As Long
Dim nro_area            As Long
Dim nro_SubDepto        As Long
Dim nro_empremu         As Long

Dim nro_estcivil        As Long
Dim nro_nacionalidad    As Long

Dim F_Nacimiento        As String
Dim F_Fallecimiento     As String
Dim F_Alta              As String
Dim F_Baja              As String
Dim F_Ingreso           As String

Dim Inserto_estr        As Boolean

Dim ter_sucursal        As Long
Dim ter_empresa         As Long
Dim ter_cajajub         As Long
Dim ter_sindicato       As Long
Dim ter_osocial_ley         As Long
Dim ter_osocial_elegida         As Long
Dim ter_bancopago       As Long
Dim ter_art             As Long
Dim ter_sexo            As Long
Dim ter_estudio         As Long
Dim ter_estado          As Long

Dim fpgo_bancaria       As Long

Dim rs As New ADODB.Recordset
Dim rs_sql As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Tel As New ADODB.Recordset
Dim rs_repl As New ADODB.Recordset


Dim SucDesc             As Boolean   'Sucursal                 -- his_estructura
Dim SecDesc             As Boolean   'Sector                   -- his_estructura
Dim CatDesc             As Boolean   'Categoria                -- his_estructura
Dim PueDesc             As Boolean   'Puesto                   -- his_estructura
Dim CCoDesc             As Boolean   'C.Costo                  -- his_estructura
Dim GerDesc             As Boolean   'Gerencia                 -- his_estructura
Dim DepDesc             As Boolean   'Departamento             -- his_estructura
Dim DirDesc             As Boolean   'Direccion                -- his_estructura
Dim CaJDesc             As Boolean   'Caja de Jubilacion       --
Dim SinDesc             As Boolean   'Sindicato                -- his_estructura
Dim OSoElegidaDesc             As Boolean   'Obra Social              -- his_estructura
Dim PoSElegidaDesc             As Boolean   'Plan OS                  -- his_estructura
Dim OSoLeyDesc             As Boolean   'Obra Social              -- his_estructura
Dim PoSLeyDesc             As Boolean   'Plan OS                  -- his_estructura
Dim CotDesc             As Boolean   'Contrato                 -- his_estructura
Dim CovDesc             As Boolean   'Convenio                 -- his_estructura
Dim LPaDesc             As Boolean   'Lugar de Pago            -- his_estructura
Dim RegDesc             As Boolean   'Regimen Horario          -- his_estructura
Dim FLiDesc             As Boolean   'Forma de Liquidacion     -- his_estructura
Dim FPaDesc             As Boolean      'Forma de Pago            -- his_estructura
Dim BcoDesc             As Boolean      'Banco Pago               --
Dim ActDesc             As Boolean      'Actividad                --
Dim CSJDesc             As Boolean      'Condicion SIJP           --
Dim SReDesc             As Boolean      'Sit. de Revista SIJP     --
Dim MCoDesc             As Boolean      'Mod. de Contrat. SIJP    --
Dim ARTDesc             As Boolean      'ART                      --
Dim EmpDesc             As Boolean      'Empresa                  --
Dim OSLDesc             As Boolean      'Empresa                  --
Dim POdoDesc             As Boolean     'Empresa                  --
Dim OSEDesc             As Boolean      'Empresa                  --
Dim LocDesc             As Boolean      'Empresa                  --
Dim AreaDesc             As Boolean     'Empresa                  --
Dim SubDepDesc             As Boolean   'Empresa                  --

Dim IngresoDom          As Boolean

Dim rs_tdoc As New ADODB.Recordset
Dim rs_emp  As New ADODB.Recordset
Dim rs_tpl  As New ADODB.Recordset
Dim rs_leg  As New ADODB.Recordset

Dim Nroadtemplate As Long

Dim Sigue As Boolean
Dim ExisteLeg As Boolean
Dim CalculaLegajo As Boolean



    On Error GoTo SaltoLinea


    ' True indica que se hace por Descripcion. False por Codigo Externo

    SucDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    SecDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CatDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    PueDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CCoDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    GerDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    DepDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    DirDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CaJDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    SinDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    OSoElegidaDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    PoSElegidaDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    OSoLeyDesc = True   ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    PoSLeyDesc = True   ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CotDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CovDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    LPaDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    RegDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    FLiDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    FPaDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    BcoDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    ActDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    CSJDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    SReDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    MCoDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    ARTDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    EmpDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    OSLDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    POdoDesc = True     ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    OSEDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    LocDesc = True      ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    AreaDesc = True     ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    SubDepDesc = True   ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    
    Sigue = True 'Indica que si en el archivo viene mas de una vez un empleado, le crea las fases
    ExisteLeg = False
    
    RegLeidos = RegLeidos + 1
    LineaCarga = LineaCarga + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
    
    ' Recupero los Valores del Archivo
    
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador) - 1
    Legajo = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Legajo = "N/A" Or Legajo = "" Then
    
        CalculaLegajo = True
        
    Else
        StrSql = "SELECT * FROM empleado WHERE empleado.empleg = " & Legajo
        OpenRecordset StrSql, rs_emp
        If (Not rs_emp.EOF) Then
            If (Not Sigue) Then
                LineaError.Writeline Mid(strReg, 1, Len(strReg))
                ErrCarga.Writeline "Linea: " & LineaCarga & " - El Empleado ya Existe."
                Ok = False
                Exit Sub
            Else
                NroTercero = rs_emp!ternro
                ExisteLeg = True
            End If
        End If
    End If
    
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Apellido = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    nombre = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If (Apellido = "" Or Apellido = "N/A") And (nombre = "" Or nombre = "N/A") Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - Debe Ingresar un Nombre y Apellido."
        Ok = False
        Exit Sub
    End If
    
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Fnac = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Fnac = "N/A" Then
       F_Nacimiento = "Null"
    Else
       F_Nacimiento = ConvFecha(Fnac)
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    PaisNac = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If PaisNac <> "N/A" Then
        StrSql = " SELECT paisnro FROM pais WHERE UPPER(paisdesc) = '" & UCase(PaisNac) & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_paisnac = rs_sql!paisnro
        Else
            StrSql = " INSERT INTO pais(paisdesc,paisdef) VALUES ('" & UCase(PaisNac) & "',0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nro_paisnac = getLastIdentity(objConn, "pais")
            
        End If
    Else
        nro_paisnac = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Nacionalidad = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Nacionalidad <> "N/A" Then
        StrSql = " SELECT nacionalnro FROM nacionalidad WHERE UPPER(nacionaldes) = '" & UCase(Nacionalidad) & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_nacionalidad = rs_sql!nacionalnro
        Else
            StrSql = " INSERT INTO nacionalidad(nacionaldes) VALUES ('" & UCase(Nacionalidad) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nro_nacionalidad = getLastIdentity(objConn, "nacionalidad")
            
        End If
    Else
        nro_nacionalidad = 0
    End If
    
    If nro_nacionalidad = 0 Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - La Nacionalidad no Existe."
        Ok = False
        Exit Sub
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Fing = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If (Fing = "N/A") Then
        F_Ingreso = "Null"
    Else
        F_Ingreso = ConvFecha(Fing)
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    EstCivil = Mid(strReg, pos1, pos2 - pos1 + 1)
    EstCivil = Mid(EstCivil, 1, 30)
    
    
    If EstCivil <> "N/A" Then
        StrSql = " SELECT estcivnro FROM estcivil WHERE UPPER(estcivdesabr) = '" & UCase(EstCivil) & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_estcivil = rs_sql!estcivnro
        Else
            StrSql = " INSERT INTO estcivil(estcivdesabr) VALUES ('" & UCase(EstCivil) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nro_estcivil = getLastIdentity(objConn, "estcivil")
            
        End If
    Else
        nro_estcivil = 0
    End If
    
    If nro_estcivil = 0 Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - El Estado Civil no Existe."
        Ok = False
        Exit Sub
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Sexo = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If (Sexo = "M") Or (Sexo = "Masculino") Or (Sexo = "-1") Or (Sexo = "MASCULINO") Then
        ter_sexo = -1
    Else
        ter_sexo = 0
    End If
                                                            
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FAlta = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If FAlta = "N/A" Then
        F_Alta = "Null"
    Else
        F_Alta = ConvFecha(FAlta)
    End If
   
    If FAlta = "N/A" Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - La Fecha de Alta es Obligatoria."
        Ok = False
        Exit Sub
    End If
   
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Estudio = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Estudio <> "N/A" Then
        If Estudio = "SI" Then
            ter_estudio = -1
        Else
            ter_estudio = 0
        End If
    Else
        ter_estudio = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    NivEst = Mid(strReg, pos1, pos2 - pos1 + 1)
    NivEst = Mid(NivEst, 1, 40)
    
    If NivEst <> "N/A" Then
        StrSql = " SELECT nivnro FROM nivest WHERE UPPER(nivdesc) = '" & UCase(NivEst) & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_nivest = rs_sql!nivnro
        Else
            StrSql = " INSERT INTO nivest(nivdesc,nivsist,nivobligatorio,nivestfli) VALUES ('" & UCase(NivEst) & "',-1,0,-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_nivest = getLastIdentity(objConn, "nivest")
        
        End If
    Else
        nro_nivest = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Tdocu = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Tdocu = Mid(Tdocu, 1, 8)
    
    If Tdocu <> "N/A" Then
        StrSql = " SELECT tidnro FROM tipodocu WHERE UPPER(tidsigla) = '" & UCase(Tdocu) & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_tdocumento = rs_sql!tidnro
        Else
            StrSql = " INSERT INTO tipodocu(tidnom,tidsigla,tidsist,instnro,tidunico) VALUES ('" & UCase(Tdocu) & "','" & UCase(Tdocu) & "',0,0,0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_tdocumento = getLastIdentity(objConn, "tipodocu")
            
        End If
    Else
        nro_tdocumento = 0
    End If
    
    If nro_tdocumento = 0 Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - El Tipo de Documento no Existe."
        Ok = False
        Exit Sub
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Ndocu = Mid(strReg, pos1, pos2 - pos1 + 1)
    Ndocu = Mid(Ndocu, 1, 30)
    
    If Ndocu = "N/A" Then
        Ndocu = ""
    End If
    
    StrSql = "SELECT * FROM empleado "
    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro "
    StrSql = StrSql & " WHERE ter_doc.tidnro = " & nro_tdocumento & " AND ter_doc.nrodoc = '" & Ndocu & "'"
    OpenRecordset StrSql, rs_tdoc
    
    If (Not rs_tdoc.EOF) Then
        If (Not Sigue) Then
            LineaError.Writeline Mid(strReg, 1, Len(strReg))
            ErrCarga.Writeline "Linea: " & LineaCarga & " - Ese Tipo y Numero de Documento estan Asignados a otro Empleado."
            Ok = False
            Exit Sub
        Else
            NroTercero = rs_tdoc!ternro
            ExisteLeg = True
            ErrCarga.Writeline "Linea: " & LineaCarga & " - Empleado: " & Legajo & " - Ese Tipo y Numero de Documento estan Asignados a otro Empleado."
        End If
    End If
    
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Cuil = Mid(strReg, pos1, pos2 - pos1 + 1)
    Cuil = Mid(Cuil, 1, 30)
    
    If Cuil = "N/A" Then
        Cuil = Ndocu
        CalcularCUIL (Cuil)
    End If
    
    ' Hasta Aqui los Datos Obligatorios del Empleado
    
    IngresoDom = True
        
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Calle = Mid(strReg, pos1, pos2 - pos1 + 1)
    Calle = Mid(Calle, 1, 30)
    
    If Calle = "N/A" Then
        Calle = ""
        IngresoDom = False
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Nro = Mid(strReg, pos1, pos2 - pos1 + 1)
    Nro = Mid(Nro, 1, 8)
    
    If (Nro <> "N/A") Then
        nro_nrodom = Nro
    Else
        nro_nrodom = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Piso = Mid(strReg, pos1, pos2 - pos1 + 1)
    Piso = Mid(Piso, 1, 8)
    
    If Piso = "N/A" Then
        Piso = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Depto = Mid(strReg, pos1, pos2 - pos1 + 1)
    Depto = Mid(Depto, 1, 8)

    If Depto = "N/A" Then
        Depto = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Torre = Mid(strReg, pos1, pos2 - pos1 + 1)
    Torre = Mid(Torre, 1, 8)
    
    If Torre = "N/A" Then
        Torre = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Manzana = Mid(strReg, pos1, pos2 - pos1 + 1)
    Manzana = Mid(Manzana, 1, 8)
    
    If Manzana = "N/A" Then
        Manzana = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Cpostal = Mid(strReg, pos1, pos2 - pos1 + 1)
    Cpostal = Mid(Cpostal, 1, 12)

    If Cpostal = "N/A" Then
        Cpostal = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Entre = Mid(strReg, pos1, pos2 - pos1 + 1)
    Entre = Mid(Entre, 1, 80)

    If Entre = "N/A" Then
        Entre = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Barrio = Mid(strReg, pos1, pos2 - pos1 + 1)
    Barrio = Mid(Barrio, 1, 30)

    If Barrio = "N/A" Then
        Barrio = ""
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Localidad = Mid(strReg, pos1, pos2 - pos1 + 1)
    Localidad = Mid(Localidad, 1, 30)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Partido = Mid(strReg, pos1, pos2 - pos1 + 1)
    Partido = Mid(Partido, 1, 30)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Zona = Mid(strReg, pos1, pos2 - pos1 + 1)
    Zona = Mid(Zona, 1, 20)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Provincia = Mid(strReg, pos1, pos2 - pos1 + 1)
    Provincia = Mid(Provincia, 1, 20)
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Pais = Mid(strReg, pos1, pos2 - pos1 + 1)
    Pais = Mid(Pais, 1, 20)
    
    If Pais <> "N/A" Then
        Call ValidarPais(Pais, nro_pais)
    Else
        nro_pais = 0
    End If
    
    If Provincia <> "N/A" Then
        Call ValidarProvincia(Provincia, nro_provincia, nro_pais)
    Else
        nro_provincia = 0
    End If
    
    If Localidad <> "N/A" Then
        Call ValidarLocalidad(Localidad, nro_localidad, nro_pais, nro_provincia)
    Else
        nro_localidad = 0
    End If
    
    If Partido <> "N/A" Then
        Call ValidarPartido(Partido, nro_partido)
    Else
        nro_partido = 0
    End If
    
    If Zona <> "N/A" Then
        Call ValidarZona(Zona, nro_zona, nro_provincia)
    Else
        nro_zona = 0
    End If
    
    If (IngresoDom = True) And (nro_localidad = 0) Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - Debe Ingresar la Localidad."
        Ok = False
        Exit Sub
    End If
    
    If (IngresoDom = True) And (nro_provincia = 0) Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - Debe Ingresar la Provincia."
        Ok = False
        Exit Sub
    End If
    
    If (IngresoDom = True) And (nro_pais = 0) Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - Debe Ingresar la Pais."
        Ok = False
        Exit Sub
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Telefono = Mid(strReg, pos1, pos2 - pos1 + 1)
    Telefono = Mid(Telefono, 1, 20)
    
    If Telefono = "N/A" Then
        Telefono = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    TelLaboral = Mid(strReg, pos1, pos2 - pos1 + 1)
    TelLaboral = Mid(TelLaboral, 1, 20)
    
    If TelLaboral = "N/A" Then
        TelLaboral = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    TelCelular = Mid(strReg, pos1, pos2 - pos1 + 1)
    TelCelular = Mid(TelCelular, 1, 20)
    
    If TelCelular = "N/A" Then
        TelCelular = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Email = Mid(strReg, pos1, pos2 - pos1 + 1)
    Email = Mid(Email, 1, 100)

    If Email = "N/A" Then
        Email = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Sucursal = Mid(strReg, pos1, pos2 - pos1 + 1)

    ' Validacion y Creacion de la Sucursal (junto con sus Complementos)

    If Sucursal <> "N/A" Then
        If SucDesc Then
            Call ValidaEstructura(1, Sucursal, nro_sucursal, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(1, Sucursal, nro_sucursal, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(10, Sucursal, ter_sucursal)
            Call CreaComplemento(1, ter_sucursal, nro_sucursal, Sucursal)
            Inserto_estr = False
        End If
    Else
        nro_sucursal = 0
    End If
    
    ' Validacion y Creacion del Sector
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Sector = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Sector <> "N/A" Then
        If SecDesc Then
            Call ValidaEstructura(2, Sector, nro_sector, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(2, Sector, nro_sector, Inserto_estr)
        End If
    Else
        nro_sector = 0
    End If

    ' Validacion, Creacion del Convenio
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Convenio = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Convenio <> "N/A" Then
        If CovDesc Then
            Call ValidaEstructura(19, Convenio, nro_convenio, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(19, Convenio, nro_convenio, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(19, 0, nro_convenio, Convenio)
            Inserto_estr = False
        End If
    Else
        nro_convenio = 0
    End If
    
    ' Validacion y Creacion de la Categoria

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    categoria = Mid(strReg, pos1, pos2 - pos1 + 1)

    If (categoria <> "N/A" And nro_convenio <> 0) Then
        If CatDesc Then
            'Call ValidaEstructura(3, categoria, nro_categoria, Inserto_estr)
            Call ValidaCategoria(3, categoria, nro_convenio, nro_categoria, Inserto_estr)
            
        Else
            'Call ValidaEstructuraCodExt(3, categoria, nro_categoria, Inserto_estr)
            Call ValidaCategoriaCodExt(3, categoria, nro_convenio, nro_categoria, Inserto_estr)
        End If
    Else
        nro_categoria = 0
    End If
    
    ' Validacion y Creacion del Puesto (junto con sus Complementos)

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Puesto = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Puesto <> "N/A" Then
        If PueDesc Then
            Call ValidaEstructura(4, Puesto, nro_puesto, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(4, Puesto, nro_puesto, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(4, 0, nro_puesto, Puesto)
            Inserto_estr = False
        End If
    Else
        nro_puesto = 0
    End If

    ' Validacion y Creacion del Centro de Costo
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    CCosto = Mid(strReg, pos1, pos2 - pos1 + 1)

    If CCosto <> "N/A" Then
        If CCoDesc Then
            Call ValidaEstructura(5, CCosto, nro_ccosto, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(5, CCosto, nro_ccosto, Inserto_estr)
        End If
    Else
        nro_ccosto = 0
    End If

    ' Validacion y Creacion de la Gerencia
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Gerencia = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Gerencia <> "N/A" Then
        If GerDesc Then
            Call ValidaEstructura(6, Gerencia, nro_gerencia, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(6, Gerencia, nro_gerencia, Inserto_estr)
        End If
    Else
        nro_gerencia = 0
    End If

    ' Validacion y Creacion del Departamento
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Departamento = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Departamento <> "N/A" Then
        If DepDesc Then
            Call ValidaEstructura(9, Departamento, nro_departamento, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(9, Departamento, nro_departamento, Inserto_estr)
        End If
    Else
        nro_departamento = 0
    End If

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Direccion = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Direccion <> "N/A" Then
        If DirDesc Then
            Call ValidaEstructura(35, Direccion, nro_direccion, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(35, Direccion, nro_direccion, Inserto_estr)
        End If
    Else
        nro_direccion = 0
    End If
    
    ' Validacion y Creacion de la Caja de Jubilacion
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    CajaJub = Mid(strReg, pos1, pos2 - pos1 + 1)

    If CajaJub <> "N/A" Then
        If CaJDesc Then
            Call ValidaEstructura(15, CajaJub, nro_cajajub, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(15, CajaJub, nro_cajajub, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(6, CajaJub, ter_cajajub)
            Call CreaComplemento(15, ter_cajajub, nro_cajajub, CajaJub)
        End If
    Else
        nro_cajajub = 0
    End If

    ' Validacion y Creacion del Sindicato
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Sindicato = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Sindicato <> "N/A" Then
        If SinDesc Then
            Call ValidaEstructura(16, Sindicato, nro_sindicato, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(16, Sindicato, nro_sindicato, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(5, Sindicato, ter_sindicato)
            Call CreaComplemento(16, ter_sindicato, nro_sindicato, Sindicato)
        End If
    Else
        nro_sindicato = 0
    End If
    
    ' Validacion y Creacion de la Obra Social por Ley
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    OSocialLey = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))

    If OSocialLey <> "N/A" Then
        If OSoLeyDesc Then
            Call ValidaEstructura(24, OSocialLey, nro_osocial_ley, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(24, OSocialLey, nro_osocial_ley, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(4, OSocialLey, ter_osocial_ley)
            Call CreaComplemento(24, ter_osocial_ley, nro_osocial_ley, OSocialLey)
        Else
            StrSql = " SELECT origen FROM replica_estr WHERE estrnro = " & nro_osocial_ley
            OpenRecordset StrSql, rs_repl
            
            If Not rs_repl.EOF Then
                ter_osocial_ley = rs_repl!Origen
                rs_repl.Close
            End If
            
        End If
    Else
        nro_osocial_ley = 0
    End If

    ' Validacion y Creacion del Plan de Obra Social por Ley
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    PlanOSLey = Mid(strReg, pos1, pos2 - pos1 + 1)

    If (PlanOSLey <> "N/A" And nro_osocial_ley <> 0) Then
        If PoSLeyDesc Then
            Call ValidaEstructura(25, PlanOSLey, nro_planos_ley, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(25, PlanOSLey, nro_planos_ley, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(25, ter_osocial_ley, nro_planos_ley, PlanOSLey)
            Inserto_estr = False
        End If
    Else
        nro_planos_ley = 0
    End If
    
    ' Validacion y Creacion de la Obra Social Elegida
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    OSocialElegida = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    
    If OSocialElegida <> "N/A" Then
        If OSoElegidaDesc Then
            Call ValidaEstructura(17, OSocialElegida, nro_osocial_elegida, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(17, OSocialElegida, nro_osocial_elegida, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(4, OSocialElegida, ter_osocial_elegida)
            Call CreaComplemento(17, ter_osocial_elegida, nro_osocial_elegida, OSocialElegida)
        Else
            StrSql = " SELECT origen FROM replica_estr WHERE estrnro = " & nro_osocial_elegida
            OpenRecordset StrSql, rs_sql
            ter_osocial_elegida = rs_sql!Origen
            rs_sql.Close
        End If
    Else
        nro_osocial_elegida = 0
    End If

    ' Validacion y Creacion del Plan de Obra Social Elegida
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    PlanOSElegida = Mid(strReg, pos1, pos2 - pos1 + 1)
    

    If (PlanOSElegida <> "N/A" And nro_osocial_elegida <> 0) Then
        If PoSElegidaDesc Then
            Call ValidaEstructura(23, PlanOSElegida, nro_planos_elegida, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(23, PlanOSElegida, nro_planos_elegida, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(23, ter_osocial_elegida, nro_planos_elegida, PlanOSElegida)
            Inserto_estr = False
        End If
    Else
        nro_planos_elegida = 0
    End If
    
    ' Validacion y Creacion del Contrato

    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Contrato = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Contrato <> "N/A" Then
        If CotDesc Then
            Call ValidaEstructura(18, Contrato, nro_contrato, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(18, Contrato, nro_contrato, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaComplemento(18, 0, nro_contrato, Contrato)
            Inserto_estr = False
        End If
    Else
        nro_contrato = 0
    End If
    
    ' Validacion y Creacion del Lugar de Pago
        
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    LPago = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))

    If LPago <> "N/A" Then
        If LPaDesc Then
            Call ValidaEstructura(20, LPago, nro_lpago, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(20, LPago, nro_lpago, Inserto_estr)
        End If
    Else
        nro_lpago = 0
    End If

    ' Validacion y Creacion del Regimen Horario
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    RegHorario = Mid(strReg, pos1, pos2 - pos1 + 1)

    If RegHorario <> "N/A" Then
        If RegDesc Then
            Call ValidaEstructura(21, RegHorario, nro_reghorario, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(21, RegHorario, nro_reghorario, Inserto_estr)
        End If
    Else
        nro_reghorario = 0
    End If

    ' Validacion y Creacion de la Forma de Liquidacion
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FormaLiq = Mid(strReg, pos1, pos2 - pos1 + 1)

    If FormaLiq <> "N/A" Then
        If FLiDesc Then
            Call ValidaEstructura(22, FormaLiq, nro_formaliq, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(22, FormaLiq, nro_formaliq, Inserto_estr)
        End If
    Else
        nro_formaliq = 0
    End If

    ' Validacion y Creacion de la Forma de Pago
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FormaPago = Mid(strReg, pos1, pos2 - pos1 + 1)

    If FormaPago <> "N/A" Then
        StrSql = " SELECT fpagnro FROM formapago WHERE fpagdescabr = '" & FormaPago & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_formapago = rs_sql!fpagnro
        Else
            StrSql = " INSERT INTO formapago(fpagdescabr,fpagbanc,acunro,monnro) VALUES ('" & FormaPago & "',0,6,1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nro_formapago = getLastIdentity(objConn, "formapago")
            
        End If
    Else
        nro_formapago = 0
    End If
    
    ' Validacion y Creacion de los Bancos de Pago
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    BancoPago = Mid(strReg, pos1, pos2 - pos1 + 1)

    If BancoPago <> "N/A" Then
        If BcoDesc Then
            Call ValidaEstructura(41, BancoPago, nro_bancopago, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(41, BancoPago, nro_bancopago, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(13, BancoPago, ter_bancopago)
            Call CreaComplemento(41, ter_bancopago, nro_bancopago, BancoPago)
        End If
        fpgo_bancaria = -1
    Else
        nro_bancopago = 0
        fpgo_bancaria = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    NroCuenta = Mid(strReg, pos1, pos2 - pos1 + 1)
    If NroCuenta = "N/A" Then
        NroCuenta = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    NroCBU = Mid(strReg, pos1, pos2 - pos1 + 1)
    If NroCBU = "N/A" Then
        NroCBU = ""
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    SucBanco = Mid(strReg, pos1, pos2 - pos1 + 1)
    If SucBanco = "N/A" Then
        SucBanco = ""
    End If

    ' Validacion y Creacion de la Actividad
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Actividad = Mid(strReg, pos1, pos2 - pos1 + 1)

    If Actividad <> "N/A" Then
        If ActDesc Then
            Call ValidaEstructura(29, Actividad, nro_actividad, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(29, Actividad, nro_actividad, Inserto_estr)
        End If
    Else
        nro_actividad = 0
    End If

    ' Validacion y Creacion de la Condicion SIJP
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    CondSIJP = Mid(strReg, pos1, pos2 - pos1 + 1)

    If CondSIJP <> "N/A" Then
        If CSJDesc Then
            Call ValidaEstructura(31, CondSIJP, nro_condsijp, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(31, CondSIJP, nro_condsijp, Inserto_estr)
        End If
    Else
        nro_condsijp = 0
    End If

    ' Validacion y Creacion de la Situacion de Revista
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    SitRev = Mid(strReg, pos1, pos2 - pos1 + 1)

    If SitRev <> "N/A" Then
        If SReDesc Then
            Call ValidaEstructura(30, SitRev, nro_sitrev, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(30, SitRev, nro_sitrev, Inserto_estr)
        End If
    Else
        nro_sitrev = 0
    End If
    
    ' Validacion y Creacion de la ART
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    ART = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If ART <> "N/A" Then
        If ARTDesc Then
            Call ValidaEstructura(40, ART, nro_art, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(40, ART, nro_art, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(8, ART, ter_art)
            Call CreaComplemento(40, ter_art, nro_art, ART)
        End If
    Else
        nro_art = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Estado = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If UCase(Estado) = "ACTIVO" Then
        ter_estado = -1
    Else
        ter_estado = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    CausaBaja = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Not EsNulo(CausaBaja) And CausaBaja <> "N/A" Then
        StrSql = " SELECT caunro FROM causa WHERE caudes = '" & CausaBaja & "'"
        OpenRecordset StrSql, rs_sql
        If Not rs_sql.EOF Then
            nro_causabaja = rs_sql!caunro
        Else
            StrSql = " INSERT INTO causa(caudes,causist,caudesvin,empnro) VALUES ('" & CausaBaja & "',0,-1,1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nro_causabaja = getLastIdentity(objConn, "causa")
            
        End If
    Else
        nro_causabaja = 0
    End If
    
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FBaja = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If EsNulo(FBaja) Or FBaja = "N/A" Then
        F_Baja = "Null"
    Else
        F_Baja = ConvFecha(FBaja)
    End If
    
    ' Validacion y Creacion de la Empresa
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    If pos2 > 0 Then
        Empresa = Mid(strReg, pos1, pos2 - pos1 + 1)
    Else
        pos2 = Len(strReg)
        Empresa = Mid(strReg, pos1, pos2 - pos1 + 1)
    End If

    If Empresa <> "N/A" Then
        If EmpDesc Then
            Call ValidaEstructura(10, Empresa, nro_empresa, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(10, Empresa, nro_empresa, Inserto_estr)
        End If
        
        If Inserto_estr Then
            Call CreaTercero(10, Empresa, ter_empresa)
            Call CreaComplemento(10, ter_empresa, nro_empresa, Empresa)
        End If
    Else
        nro_empresa = 0
    End If
    
    'Remuneración del Empleado
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Empremu = Mid(strReg, pos1, pos2 - pos1 + 1)
    
    If Empremu = "N/A" Then
        Empremu = "Null"
    End If
   
    ' Modelo de Organizacion
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    
    If pos2 > 0 Then
        ModOrg = Mid(strReg, pos1, pos2 - pos1 + 1)
        
        StrSql = "SELECT * FROM adptemplate WHERE tplatedesabr = '" & ModOrg & "'"
        OpenRecordset StrSql, rs_tpl
        
        If rs_tpl.EOF Then
        
            StrSql = "INSERT INTO adptemplate (tplatedesabr,tplatedefault) VALUES ('" & ModOrg & "',-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nro_ModOrg = getLastIdentity(objConn, "adptemplate")
            
        Else
        
            nro_ModOrg = rs_tpl!tplatenro
        
        End If
        
    Else
        nro_ModOrg = 0
    End If


  ' Inserto el Tercero
  If F_Nacimiento = "Null" Then
    F_Nacimiento = "''"
  End If
  If F_Ingreso = "Null" Then
    F_Ingreso = "''"
  End If
  
  If CalculaLegajo Then
    Call CalcularLegajo(nro_empresa, Legajo)
  End If

    If Not ExisteLeg Then

        StrSql = " INSERT INTO tercero(ternom,terape,terfecnac,tersex,terestciv,estcivnro,nacionalnro,paisnro,terfecing)"
        StrSql = StrSql & " VALUES('" & nombre & "','" & Apellido & "'," & F_Nacimiento & "," & ter_sexo & "," & nro_estcivil & "," & nro_estcivil & ","
        If nro_nacionalidad <> 0 Then
            StrSql = StrSql & nro_nacionalidad & ","
        Else
            StrSql = StrSql & "null,"
        End If
        If nro_paisnac <> 0 Then
            StrSql = StrSql & nro_paisnac & ","
        Else
            StrSql = StrSql & "null,"
        End If
        StrSql = StrSql & F_Ingreso & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

        NroTercero = getLastIdentity(objConn, "tercero")
    Else
      ErrCarga.Writeline "Linea: " & LineaCarga & " - Empleado: " & Legajo & " - Ese Empleado ya existe en la base."
    End If

    Flog.Writeline "Codigo de Tercero = " & NroTercero

    If Not ExisteLeg Then
        StrSql = " INSERT INTO empleado(empleg,empfecalta,empfecbaja,empest,empfaltagr,"
        StrSql = StrSql & "ternro,nivnro,empestudia,terape,ternom,empinterno,empemail,"
        StrSql = StrSql & "empnro,tplatenro,empremu) VALUES("
        StrSql = StrSql & Legajo & "," & F_Alta & "," & F_Baja & "," & ter_estado & "," & F_Alta & ","
        StrSql = StrSql & NroTercero & "," & nro_nivest & "," & ter_estudio & ",'" & Apellido & "','"
        StrSql = StrSql & nombre & "',Null,'" & Email & "',1," & nro_ModOrg & "," & Empremu & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

    Flog.Writeline "Inserte el Empleado - " & Legajo & " - " & Apellido & " - " & nombre

    ' Inserto el Registro correspondiente en ter_tip
    
    If Not ExisteLeg Then
    
        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & NroTercero & ",1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If

' Inserto los Documentos
    
    If Not ExisteLeg Then
    
        If nro_tdocumento <> 0 Then
            StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
            StrSql = StrSql & " VALUES(" & NroTercero & "," & nro_tdocumento & ",'" & Ndocu & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline "Inserte el DU - "
        End If
            
    End If
  

    If Not ExisteLeg Then
    
        If Cuil <> "" Then
            StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
            StrSql = StrSql & " VALUES(" & NroTercero & ",10,'" & Cuil & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.Writeline "Inserte el CUIL - "
        End If
        
    End If

' Inserto el Domicilio

  If rs.State = adStateOpen Then
    rs.Close
  End If
  
  If Not ExisteLeg Then
  
    If (nro_localidad <> 0 And nro_provincia <> 0 And nro_pais <> 0) Then
        StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
        StrSql = StrSql & " VALUES(1," & NroTercero & ",-1,2)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        NroDom = getLastIdentity(objConn, "cabdom")
      
        StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,entrecalles,"
        StrSql = StrSql & "locnro,provnro,paisnro,barrio,partnro,zonanro) "
        StrSql = StrSql & " VALUES (" & NroDom & ",'" & Calle & "','" & nro_nrodom & "','" & Piso & "','"
        StrSql = StrSql & Depto & "','" & Torre & "','" & Manzana & "','" & Cpostal & "','" & Entre & "'," & nro_localidad & ","
        StrSql = StrSql & nro_provincia & "," & nro_pais & ",'" & Barrio & "'," & nro_partido & "," & nro_zona & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
      
        Flog.Writeline "Inserte el Domicilio - "
        
        If Telefono <> "" Then
          StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
          StrSql = StrSql & " VALUES(" & NroDom & ",'" & Telefono & "',0,-1,0)"
          objConn.Execute StrSql, , adExecuteNoRecords
          Flog.Writeline "Inserte el Telefono - "
        End If
        If TelLaboral <> "" Then
          StrSql = "SELECT * FROM telefono "
          StrSql = StrSql & " WHERE domnro =" & NroDom
          StrSql = StrSql & " AND telnro ='" & TelLaboral & "'"
          If rs_Tel.State = adStateOpen Then rs_Tel.Close
          OpenRecordset StrSql, rs_Tel
          If rs_Tel.EOF Then
              StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
              StrSql = StrSql & " VALUES(" & NroDom & ",'" & TelLaboral & "',0,0,0)"
              objConn.Execute StrSql, , adExecuteNoRecords
              Flog.Writeline "Inserte el Telefono Laboral - "
          End If
        End If
        If TelCelular <> "" Then
              StrSql = "SELECT * FROM telefono "
              StrSql = StrSql & " WHERE domnro =" & NroDom
              StrSql = StrSql & " AND telnro ='" & TelCelular & "'"
              If rs_Tel.State = adStateOpen Then rs_Tel.Close
              OpenRecordset StrSql, rs_Tel
              If rs_Tel.EOF Then
                  StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                  StrSql = StrSql & " VALUES(" & NroDom & ",'" & TelCelular & "',0,0,-1)"
                  objConn.Execute StrSql, , adExecuteNoRecords
                  Flog.Writeline "Inserte el Telefono Celular - "
              End If
        End If
        
    End If
    
  End If
  
  If Not ExisteLeg Then
    ' Inserto las Fases
    StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
    StrSql = StrSql & " VALUES( " & NroTercero & "," & F_Alta & "," & F_Baja & ","
    If nro_causabaja <> 0 Then
      StrSql = StrSql & nro_causabaja
    Else
      StrSql = StrSql & "null"
    End If
    StrSql = StrSql & "," & ter_estado & ",-1,-1,-1,-1,-1)"
    objConn.Execute StrSql, , adExecuteNoRecords
  End If
  
  'Inserto la cuenta bancaria
    If Not ExisteLeg Then
  
        If (nro_formapago <> 0 And nro_bancopago <> 0 And NroCuenta <> "") Then
        
          StrSql = " INSERT INTO ctabancaria (ternro,fpagnro,banco,ctabestado,"
          StrSql = StrSql & "ctabsuc,ctabnro,ctabporc,ctabcbu) VALUES ("
          StrSql = StrSql & NroTercero & "," & nro_formapago & "," & ter_bancopago & ","
          StrSql = StrSql & "-1,'" & SucBanco & "','" & NroCuenta & "',100,'" & NroCBU & "')"
          objConn.Execute StrSql, , adExecuteNoRecords
          Flog.Writeline "Inserte la Cuenta Bancaria - "
          
        End If
        
    End If
  ' Inserto las Estructuras

'
  Call AsignarEstructura(1, nro_sucursal, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(2, nro_sector, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(3, nro_categoria, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(4, nro_puesto, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(5, nro_ccosto, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(6, nro_gerencia, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(9, nro_departamento, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(10, nro_empresa, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(15, nro_cajajub, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(16, nro_sindicato, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(17, nro_osocial_elegida, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(18, nro_contrato, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(19, nro_convenio, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(20, nro_lpago, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(21, nro_reghorario, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(22, nro_formaliq, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(23, nro_planos_elegida, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(24, nro_osocial_ley, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(25, nro_planos_ley, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(29, nro_actividad, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(30, nro_sitrev, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(31, nro_condsijp, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(35, nro_direccion, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(40, nro_art, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura(41, nro_bancopago, NroTercero, F_Alta, F_Baja)
  
  LineaOK.Writeline Mid(strReg, 1, Len(strReg))
  Ok = True
         
  If rs.State = adStateOpen Then
      rs.Close
  End If

  Exit Sub

SaltoLinea:

    LineaError.Writeline Mid(strReg, 1, Len(strReg))
    ErrCarga.Writeline "Linea: " & LineaCarga & " - " & Err.Description
    Resume Next
    MyRollbackTrans
    Ok = False
    

End Sub


Public Sub LineaModelo_630(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Estructuras
' Autor      : FGZ
' Fecha      : 21/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As String   'LEGAJO                        -- empleado.empleg
Dim Estructura      As String   'Estructura                    -- his_estructura.estrnro
Dim TipoEstructura  As String   'Tipo de Estructura            -- his_estructura.tenro
Dim FAlta           As String   'Fecha Desde en la Estructura  -- his_estructura.htetdesde
Dim FBaja           As String   'Fecha Hasta en la Estructura  -- his_estructura.htethasta

Dim ternro As Long

Dim pos1 As Long
Dim pos2 As Long

Dim NroTercero          As Long
Dim NroLegajo           As Long
Dim nro_estructura      As Long
Dim F_Alta              As String
Dim F_Baja              As String

Dim Inserto_estr        As Boolean

Dim rs As New ADODB.Recordset
Dim rs_sql As New ADODB.Recordset
Dim rs_tes As New ADODB.Recordset

Dim nro_tenro As Long

' True indica que se hace por Descripcion. False por Codigo Externo

Dim EstrDesc             As Boolean   'Sucursal                 -- his_estructura


    EstrDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo

    ' Recupero los Valores del Archivo
    
    pos1 = 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    TipoEstructura = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Estructura = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    FAlta = Mid(strReg, pos1, pos2 - pos1)
    
    If FAlta = "N/A" Or FAlta = "" Then
        F_Alta = "Null"
    Else
        F_Alta = ConvFecha(FAlta)
    End If
    
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    FBaja = Mid(strReg, pos1, pos2 - pos1)
    
    If FBaja = "N/A" Or FBaja = "" Then
        F_Baja = "Null"
    Else
        F_Baja = ConvFecha(FBaja)
    End If
    
    ' Valida que los campos obligatorios este cargados
    
    If TipoEstructura = "" Or Legajo = "" Or Estructura = "" Or FAlta = "" Then
        Exit Sub
    End If
    
    ' Busca el Tercero
    StrSql = "SELECT ternro FROM empleado WHERE empleado.empleg = " & Legajo
    OpenRecordset StrSql, rs
    
    If rs.EOF Then Exit Sub
    
    NroTercero = rs!ternro

    StrSql = "SELECT tenro FROM tipoestructura WHERE UPPER(tedabr) = '" & UCase(TipoEstructura) & "'"
    OpenRecordset StrSql, rs_tes
    If rs_tes.EOF Then
        StrSql = "INSERT INTO tipoestructura(tedabr,tesist,tedepbaja,cenro) VALUES('" & UCase(TipoEstructura) & "',0,0,1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        nro_tenro = getLastIdentity(objConn, "tipoestructura")

        
    Else
        nro_tenro = rs_tes!Tenro
    End If


    ' Validacion y Creacion de la Sucursal (junto con sus Complementos)
    If Estructura <> "N/A" Then
        If EstrDesc Then
            Call ValidaEstructura(nro_tenro, Mid(Estructura, 1, 60), nro_estructura, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(nro_tenro, Mid(Estructura, 1, 20), nro_estructura, Inserto_estr)
        End If
    End If
    
  ' Inserto las Estructuras
  Call AsignarEstructura(nro_tenro, nro_estructura, NroTercero, F_Alta, F_Baja)
         
  If rs.State = adStateOpen Then
      rs.Close
  End If
End Sub

Public Sub LineaModelo_610()
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de DesmenFamiliar
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Long
Dim pos2            As Long
Dim rs              As New ADODB.Recordset
Dim rsa             As New ADODB.Recordset
Dim Legajo          As Long ' Legajo del Empleado
Dim Anio            As String
Dim FecDesde        As String
Dim FecHasta        As String
Dim NroItem         As String
Dim Monto           As String
Dim NroTercero      As Long
'Dim StrSql          As String

MyBeginTrans
  StrSql = " SELECT terfecnac,empleado,parenro FROM familiar "
  StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = familiar.ternro "
  StrSql = StrSql & " WHERE famcargadgi = -1 "
  OpenRecordset StrSql, rs
  
  Do While Not rs.EOF:
  
    NroTercero = rs!Empleado
    
    If rs!parenro = 1 Then
        NroItem = 10
    Else
        If rs!parenro = 2 Then
            NroItem = 11
        Else
            NroItem = 12
        End If
    End If
    If rs!terfecnac > CDate("01/01/2004") Then
        FecDesde = rs!terfecnac
    Else
        FecDesde = "01/01/2004"
    End If
    FecHasta = "31/12/2004"
    
    ' Inserto el Desmen
    StrSql = " SELECT desmondec FROM desmen WHERE empleado = " & NroTercero
    StrSql = StrSql & " AND itenro = " & NroItem
    StrSql = StrSql & " AND desano = 2004 "
    OpenRecordset StrSql, rsa
    
    If rsa.EOF Then
        ' Inserto el Desmen
        Monto = 1
        StrSql = " INSERT INTO desmen(empleado,itenro,desano,desfecdes,desfechas,desmenprorra,desmondec)"
        StrSql = StrSql & " VALUES(" & NroTercero & "," & NroItem & ",2004,'" & FecDesde & "','" & FecHasta & "',0," & Monto & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        Monto = Int(rsa!desmondec) + 1
        ' Actualizo el Desmen
        StrSql = " UPDATE desmen SET desmondec = " & Monto
        StrSql = StrSql & " WHERE empleado = " & NroTercero
        StrSql = StrSql & " AND itenro = " & NroItem
        StrSql = StrSql & " AND desano = 2004 "
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rs.MoveNext
    
  Loop
  
  MyCommitTrans
  
  If rs.State = adStateOpen Then rs.Close
  If rsa.State = adStateOpen Then rsa.Close
End Sub

Public Sub LineaModelo_615(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As Long ' Legajo del Empleado
Dim Anio            As String
Dim FecDesde        As String
Dim FecHasta        As String
Dim NroItem         As String
Dim Monto           As String

Dim pos1            As Long
Dim pos2            As Long

Dim NroTercero      As Long

Dim StrSql          As String
Dim rs              As New ADODB.Recordset

'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
        
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Anio = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    FecDesde = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    FecHasta = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    NroItem = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Monto = Mid(strReg, pos1, pos2 - pos1)
    
' Busco el empleado

  StrSql = " SELECT ternro FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, rs
  
  If Not (rs.EOF) Then
      NroTercero = rs!ternro
  Else
    Flog.Writeline "Legajo Inexistente= " & Legajo
    Exit Sub
  End If
  
  Flog.Writeline "Legajo = " & Legajo & "Codigo de Tercero = " & NroTercero

' Inserto el Desmen
  StrSql = " INSERT INTO desmen(empleado,itenro,desano,desfecdes,desfechas,desmenprorra,desmondec)"
  StrSql = StrSql & " VALUES(" & NroTercero & "," & NroItem & "," & Anio & "," & ConvFecha(FecDesde) & "," & ConvFecha(FecHasta) & ",0," & Monto & ")"
  objConn.Execute StrSql, , adExecuteNoRecords
  Flog.Writeline "Inserte el item - " & NroItem & " - " & Anio & " - " & Monto
  
  If rs.State = adStateOpen Then rs.Close
End Sub

Public Sub LineaModelo_620(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As Long ' Legajo del Empleado
Dim Anio            As String  ' Apellido del Familiar
Dim mes             As String  ' Nombre del Familiar
Dim Item1           As String  ' Item 1
Dim Item2           As String  ' Item 2
Dim Item3           As String  ' Item 3
Dim Item4           As String  ' Item 4
Dim Item5           As String  ' Item 5
Dim Item6           As String  ' Item 6
Dim Item7           As String  ' Item 7
Dim Item8           As String  ' Item 8
Dim Item9           As String  ' Item 9
Dim Item10           As String  ' Item 10
Dim Item11           As String  ' Item 11
Dim Item12           As String  ' Item 12
Dim Item13           As String  ' Item 13
Dim Item14           As String  ' Item 14
Dim Item15           As String  ' Item 15
Dim Item16           As String  ' Item 16
Dim Item17           As String  ' Item 17
Dim Item18           As String  ' Item 18
Dim Item19           As String  ' Item 19
Dim Item20           As String  ' Item 20
Dim Item21           As String  ' Item 21
Dim Item22           As String  ' Item 22

Dim pos1            As Long
Dim pos2            As Long

Dim NroTercero      As Long
Dim FecHasta_Peri   As String

Dim StrSql          As String
Dim rs              As New ADODB.Recordset

    Item1 = ""
    Item2 = ""
    Item3 = ""
    Item4 = ""
    Item5 = ""
    Item6 = ""
    Item7 = ""
    Item8 = ""
    Item9 = ""
    Item10 = ""
    Item11 = ""
    Item12 = ""
    Item13 = ""
    Item14 = ""
    Item15 = ""
    Item16 = ""
    Item17 = ""
    Item18 = ""
    Item19 = ""
    Item20 = ""
    Item21 = ""
    Item22 = ""
    
'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
        
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Anio = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    mes = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item1 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item2 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item3 = Mid(strReg, pos1, pos2 - pos1)
    
    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, separador)
    'Item4 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item5 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item6 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item7 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item8 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item9 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item10 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item11 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item12 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item13 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item14 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item15 = Mid(strReg, pos1, pos2 - pos1)
    
    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, separador)
    'Item16 = Mid(strReg, pos1, pos2 - pos1)
    
    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, separador)
    'Item17 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item18 = Mid(strReg, pos1, pos2 - pos1)
    
    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, separador)
    'Item19 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item20 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Item22 = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Item21 = Mid(strReg, pos1, pos2 - pos1)
    
    If (mes = 1) Then
        FecHasta_Peri = "31/01/" & Anio
    Else
        If (mes = 2) Then
            FecHasta_Peri = "28/02/" & Anio
        Else
            If (mes = 3) Then
                FecHasta_Peri = "31/03/" & Anio
            Else
                If (mes = 4) Then
                    FecHasta_Peri = "30/04/" & Anio
                Else
                    If (mes = 5) Then
                        FecHasta_Peri = "31/05/" & Anio
                    Else
                        If (mes = 6) Then
                            FecHasta_Peri = "30/06/" & Anio
                        Else
                            If (mes = 7) Then
                                FecHasta_Peri = "31/07/" & Anio
                            Else
                                If (mes = 8) Then
                                    FecHasta_Peri = "31/08/" & Anio
                                Else
                                    If (mes = 9) Then
                                        FecHasta_Peri = "30/09/" & Anio
                                    Else
                                        If (mes = 10) Then
                                            FecHasta_Peri = "31/10/" & Anio
                                        Else
                                            If (mes = 11) Then
                                                FecHasta_Peri = "30/11/" & Anio
                                            Else
                                                If (mes = 12) Then
                                                    FecHasta_Peri = "31/12/" & Anio
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
' Busco el empleado

  StrSql = " SELECT ternro FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, rs
  NroTercero = rs!ternro

  Flog.Writeline "Legajo = " & Legajo & "Codigo de Tercero = " & NroTercero

' Inserto el Desliq item 1
  If Item1 <> "0" And Item1 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",1,'" & FecHasta_Peri & "',Null," & Item1 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item1 - " & Legajo
  End If
' Inserto el Desliq item 2
  If Item2 <> "0" And Item2 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",2,'" & FecHasta_Peri & "',Null," & Item2 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item2 - " & Legajo
  End If
' Inserto el Desliq item 3
  If Item3 <> "0" And Item3 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",3,'" & FecHasta_Peri & "',Null," & Item3 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item3 - " & Legajo
  End If
' Inserto el Desliq item 4
  If Item4 <> "0" And Item4 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",4,'" & FecHasta_Peri & "',Null," & Item4 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item4 - " & Legajo
  End If
' Inserto el Desliq item 5
  If Item5 <> "0" And Item5 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",5,'" & FecHasta_Peri & "',Null,-" & Item5 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item5 - " & Legajo
  End If
' Inserto el Desliq item 6
  If Item6 <> "0" And Item6 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",6,'" & FecHasta_Peri & "',Null,-" & Item6 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item6 - " & Legajo
  End If
' Inserto el Desliq item 7
  If Item7 <> "0" And Item7 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",7,'" & FecHasta_Peri & "',Null," & Item7 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item7 - " & Legajo
  End If
' Inserto el Desliq item 8
  If Item8 <> "0" And Item8 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",8,'" & FecHasta_Peri & "',Null," & Item8 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item8 - " & Legajo
  End If
' Inserto el Desliq item 9
  If Item9 <> "0" And Item9 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",9,'" & FecHasta_Peri & "',Null," & Item9 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item9 - " & Legajo
  End If
' Inserto el Desliq item 10
  If Item10 <> "0" And Item10 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",10,'" & FecHasta_Peri & "',Null," & Item10 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item10 - " & Legajo
  End If
' Inserto el Desliq item 11
  If Item11 <> "0" And Item11 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",11,'" & FecHasta_Peri & "',Null," & Item11 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item11 - " & Legajo
  End If
' Inserto el Desliq item 12
  If Item12 <> "0" And Item12 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",12,'" & FecHasta_Peri & "',Null," & Item12 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item12 - " & Legajo
  End If
' Inserto el Desliq item 13
  If Item13 <> "0" And Item13 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",13,'" & FecHasta_Peri & "',Null," & Item13 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item13 - " & Legajo
  End If
' Inserto el Desliq item 14
  If Item14 <> "0" And Item14 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",14,'" & FecHasta_Peri & "',Null," & Item14 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item14 - " & Legajo
  End If
' Inserto el Desliq item 15
  If Item15 <> "0" And Item15 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",15,'" & FecHasta_Peri & "',Null," & Item15 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item15 - " & Legajo
  End If
' Inserto el Desliq item 16
  If Item16 <> "0" And Item16 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",16,'" & FecHasta_Peri & "',Null," & Item16 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item16 - " & Legajo
  End If
' Inserto el Desliq item 17
  If Item17 <> "0" And Item17 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",17,'" & FecHasta_Peri & "',Null," & Item17 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item17 - " & Legajo
  End If
' Inserto el Desliq item 18
  If Item18 <> "0" And Item18 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",18,'" & FecHasta_Peri & "',Null," & Item18 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item18 - " & Legajo
  End If
' Inserto el Desliq item 19
  If Item19 <> "0" And Item19 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",19,'" & FecHasta_Peri & "',Null," & Item19 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item19 - " & Legajo
  End If
' Inserto el Desliq item 20
  If Item20 <> "0" And Item20 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",20,'" & FecHasta_Peri & "',Null," & Item20 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item20 - " & Legajo
  End If
' Inserto el Desliq item 21
  If Item21 <> "0" And Item21 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",21,'" & FecHasta_Peri & "',Null," & Item21 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item21 - " & Legajo
  End If
' Inserto el Desliq item 22
  If Item22 <> "0" And Item22 <> "" Then
    StrSql = " INSERT INTO desliq(empleado,itenro,dlfecha,pronro,dlmonto,dlprorratea)"
    StrSql = StrSql & " values(" & NroTercero & ",22,'" & FecHasta_Peri & "',Null," & Item22 & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.Writeline "Inserte el item22 - " & Legajo
  End If
  
  If rs.State = adStateOpen Then rs.Close

End Sub


Public Sub LineaModelo_625(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Long
Dim pos2            As Long
Dim Legajo          As String   'Legajo                   -- empleado.empleg
Dim Anio            As String   'Año de la liquidacion    -- periodo.pliqanio
Dim mes             As String   'Mes de la liquidacion    -- periodo.pliqmes
Dim Proceso         As String   'Proceso de liquidacion   -- proceso.prodesc
Dim CtoCodigo       As String   'Código de concepto       -- concepto.conccod
Dim Monto           As String   'Monto liquidado          -- detliq.dlimonto
Dim Cantidad        As String   'Cantidad liquidada       -- detliq.dlicant

Dim Desc_Periodo    As String   'Descripcion del Periodo de liquidacion
Dim FecDesde_Peri   As String   'Fecha desde del periodo
Dim FecHasta_Peri   As String   'Fecha desde del periodo

Dim NroTercero          As Long

Dim Nro_Legajo          As Long
Dim nro_concepto        As Long
Dim nro_periodo         As Long
Dim nro_proceso         As Long
Dim nro_cabecera        As Long
Dim nro_tipoconc        As Long

Dim RsPeriodo    As New ADODB.Recordset
Dim RsPeri       As New ADODB.Recordset
Dim RsConcepto   As New ADODB.Recordset
Dim RsCabecera   As New ADODB.Recordset
Dim RsCabe       As New ADODB.Recordset
Dim RsCabliq     As New ADODB.Recordset
Dim RsProceso    As New ADODB.Recordset
Dim RsPro        As New ADODB.Recordset
Dim RsEmple      As New ADODB.Recordset

Dim CodPeri As Long
Dim CodPro  As Long
Dim CodCabe As Long

'    RegLeidos = RegLeidos + 1
    
    Flog.Writeline "Numero de Linea = " & RegLeidos
    
    ' Recupero los Valores del Archivo
    
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Anio = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    mes = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Proceso = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    CtoCodigo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Cantidad = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    Monto = Mid(strReg, pos1, pos2 - pos1)
    
' Busco al empleado asociado

  StrSql = " SELECT ternro,empleg FROM empleado WHERE empleg = " & Legajo
  OpenRecordset StrSql, RsEmple
  
  If RsEmple.EOF Then 'No existe el empleado
    NroTercero = 0
  Else
    NroTercero = RsEmple!ternro
  End If
  
  If NroTercero <> 0 Then
      Flog.Writeline "Procesando al empleado = " & RsEmple!empleg
    
    ' Busco el periodo de liquidacion
    
      StrSql = " SELECT pliqnro FROM periodo WHERE pliqanio = " & Anio
      StrSql = StrSql & " AND pliqmes = " & mes
      OpenRecordset StrSql, RsPeriodo
      
      If RsPeriodo.EOF Then  'No existe el periodo => lo creo
        If (mes = 1) Then
            Desc_Periodo = "Enero "
            FecDesde_Peri = "01/01/" & Anio
            FecHasta_Peri = "31/01/" & Anio
        Else
            If (mes = 2) Then
                Desc_Periodo = "Febrero "
                FecDesde_Peri = "01/02/" & Anio
                FecHasta_Peri = "28/02/" & Anio
            Else
                If (mes = 3) Then
                    Desc_Periodo = "Marzo "
                    FecDesde_Peri = "01/03/" & Anio
                    FecHasta_Peri = "31/03/" & Anio
                Else
                    If (mes = 4) Then
                        Desc_Periodo = "Abril "
                        FecDesde_Peri = "01/04/" & Anio
                        FecHasta_Peri = "30/04/" & Anio
                    Else
                        If (mes = 5) Then
                            Desc_Periodo = "Mayo "
                            FecDesde_Peri = "01/05/" & Anio
                            FecHasta_Peri = "31/05/" & Anio
                        Else
                            If (mes = 6) Then
                                Desc_Periodo = "Junio "
                                FecDesde_Peri = "01/06/" & Anio
                                FecHasta_Peri = "30/06/" & Anio
                            Else
                                If (mes = 7) Then
                                    Desc_Periodo = "Julio "
                                    FecDesde_Peri = "01/07/" & Anio
                                    FecHasta_Peri = "31/07/" & Anio
                                Else
                                    If (mes = 8) Then
                                        Desc_Periodo = "Agosto "
                                        FecDesde_Peri = "01/08/" & Anio
                                        FecHasta_Peri = "31/08/" & Anio
                                    Else
                                        If (mes = 9) Then
                                            Desc_Periodo = "Septiembre "
                                            FecDesde_Peri = "01/09/" & Anio
                                            FecHasta_Peri = "30/09/" & Anio
                                        Else
                                            If (mes = 10) Then
                                                Desc_Periodo = "Octubre "
                                                FecDesde_Peri = "01/10/" & Anio
                                                FecHasta_Peri = "31/10/" & Anio
                                            Else
                                                If (mes = 11) Then
                                                    Desc_Periodo = "Noviembre "
                                                    FecDesde_Peri = "01/11/" & Anio
                                                    FecHasta_Peri = "30/11/" & Anio
                                                Else
                                                    If (mes = 12) Then
                                                        Desc_Periodo = "Diciembre "
                                                        FecDesde_Peri = "01/12/" & Anio
                                                        FecHasta_Peri = "31/12/" & Anio
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Desc_Periodo = Desc_Periodo & Anio
        
        StrSql = " INSERT INTO periodo(pliqmes,pliqanio,pliqdesc,pliqdesde,pliqhasta) "
        StrSql = StrSql & " VALUES(" & mes & "," & Anio & ",'" & Desc_Periodo & "','" & FecDesde_Peri & "','" & FecHasta_Peri & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        nro_periodo = getLastIdentity(objConn, "periodo")
        
      Else
        nro_periodo = RsPeriodo!PliqNro
      End If
      
    ' Busco el proceso dentro del periodo de liquidacion
    
      StrSql = " SELECT pronro FROM proceso WHERE prodesc = '" & Proceso & "'"
      StrSql = StrSql & " AND pliqnro = " & nro_periodo
      OpenRecordset StrSql, RsProceso
      
      If RsProceso.EOF Then  'No existe el proceso => lo creo
        StrSql = " INSERT INTO proceso(pliqnro,tprocnro,profecpago,prodesc,profecplan,profecini,profecfin) "
        StrSql = StrSql & " VALUES(" & nro_periodo & ",3,'" & FecHasta_Peri & "','" & Proceso & "','"
        StrSql = StrSql & FecHasta_Peri & "','" & FecDesde_Peri & "','" & FecHasta_Peri & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        nro_proceso = getLastIdentity(objConn, "proceso")
        
      Else
        nro_proceso = RsProceso!pronro
      End If
      
    ' Busco el concepto
    
      StrSql = " SELECT concnro,tconnro FROM concepto WHERE conccod = '" & CtoCodigo & "'"
      OpenRecordset StrSql, RsConcepto
      
      If RsConcepto.EOF Then  'No existe el concepto => Error
        nro_concepto = 0
      Else
        nro_concepto = RsConcepto!concnro
'Ajuste para Accor porque los conceptos son todos positivos -------------------------------------
        nro_tipoconc = RsConcepto!tconnro
        If nro_tipoconc = 6 Or nro_tipoconc = 8 Or nro_tipoconc = 10 Or nro_tipoconc = 13 Then
            If Mid(Monto, 1, 1) = "-" Then
                Monto = Mid(Monto, 2)
            Else
                Monto = "-" & Monto
            End If
        End If
'Fin ajuste para Accor --------------------------------------------------------------------------
      End If
    
    ' Busco el cabliq del empleado para el proceso y periodo
    
      StrSql = " SELECT cliqnro FROM cabliq WHERE empleado = " & NroTercero
      StrSql = StrSql & " AND pronro = " & nro_proceso
      OpenRecordset StrSql, RsCabliq
      
      If RsCabliq.EOF Then  'No existe el cabliq => lo creo
        StrSql = " INSERT INTO cabliq(empleado,pronro) VALUES("
        StrSql = StrSql & NroTercero & "," & nro_proceso & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " SELECT @@IDENTITY AS CodCabe "
        OpenRecordset StrSql, RsCabe
        
        nro_cabecera = RsCabe!CodCabe
      Else
        nro_cabecera = RsCabliq!cliqnro
      End If
      
      If nro_concepto <> 0 Then ' Inserto el detalle de liquidacion
      
        StrSql = " INSERT INTO detliq(cliqnro,concnro,dlimonto,dlicant,dliqdesde,dliqhasta,tconnro,dlitexto,dlifec) VALUES("
        StrSql = StrSql & nro_cabecera & "," & nro_concepto & "," & Monto & "," & Cantidad
        StrSql = StrSql & ",0,0,0,'','00:00:00')"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        Flog.Writeline "Inserte el Detalle de liquidacion  - " & CtoCodigo & " - " & mes & " - " & Anio
      Else
        Flog.Writeline "Concepto Inexistente = " & CtoCodigo
      End If
  Else
      Flog.Writeline "Empleado Inexistente = " & RsEmple!empleg
  End If
  
End Sub

Public Sub LineaModelo_635(ByVal strReg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Títulos del Empleado
' Autor      : EPL
' Fecha      : 30/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Legajo          As String   'LEGAJO                        -- empleado.empleg
Dim Titulo          As String   'Estructura                    -- his_estructura.estrnro
Dim Institucion     As String   'Tipo de Estructura            -- his_estructura.tenro
Dim Nivel           As String

Dim pos1 As Long
Dim pos2 As Long

Dim NroTercero          As Long
Dim NroLegajo           As Long
Dim NroTitulo           As Long
Dim NroInstitucion      As Long
Dim NroNivel            As Long

Dim rs_tit As New ADODB.Recordset
Dim rs_ins As New ADODB.Recordset
Dim rs_niv As New ADODB.Recordset
Dim rs     As New ADODB.Recordset

    ' Recupero los Valores del Archivo
    
    pos1 = 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Titulo = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Institucion = Mid(strReg, pos1, pos2 - pos1)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Nivel = Mid(strReg, pos1, pos2 - pos1)
    
     
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & Legajo
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
        Exit Sub
    Else
        NroTercero = rs!ternro
    End If
    
    ' Valida que los campos obligatorios este cargados
    
    If (Legajo = "N/A" Or Legajo <> "") Or (Titulo <> "" Or Titulo <> "N/A") Then
        Exit Sub
    End If

    StrSql = "SELECT titnro FROM titulo WHERE titdesabr = '" & Titulo & "'"
    OpenRecordset StrSql, rs_tit
    
    StrSql = "SELECT instnro FROM institucion WHERE instdes = '" & Institucion & "'"
    OpenRecordset StrSql, rs_ins
    
    StrSql = "SELECT nivnro FROM nivest WHERE nivdesc = '" & Nivel & "'"
    OpenRecordset StrSql, rs_niv
    
    ' Busco el nivel de estudio, si no existe lo creo
    
    If rs_niv.EOF Then
    
        If Nivel = "N/A" Then
        
            StrSql = "INSERT INTO nivest (nivdesc,nivsist,nivobligatorio,nivestfli) VALUES ('" & Nivel & "',0,0,0)"
            objConn.Execute StrSql, , adExecuteNoRecords
        
            NroNivel = getLastIdentity(objConn, "nivest")
        
        Else
        
            NroNivel = 0
        
        End If
        
    Else
    
        NroNivel = rs_niv!nivnro
    
    End If
    
    
    ' Busco la institucion, si no existe lo creo
    
    If rs_ins.EOF Then
    
        If Institucion = "N/A" Then
        
            StrSql = "INSERT INTO institucion (instdes,instabre) VALUES('" & Institucion & "',' ')"
            objConn.Execute StrSql, , adExecuteNoRecords
        
            NroInstitucion = getLastIdentity(objConn, "institucion")
    
        Else
        
            NroInstitucion = 0
        
        End If
    Else
    
        NroInstitucion = rs_ins!instnro
    
    End If
    
    ' Busco el Título, si no existe lo creo
    
    If rs_tit.EOF Then
    
        StrSql = "INSERT INTO titulo (titdesabr,instnro,nivnro) VALUES('"
        StrSql = StrSql & Titulo & "'," & NroInstitucion & "," & NroNivel & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        NroTitulo = getLastIdentity(objConn, "titulo")
    
    Else
    
        NroTitulo = rs_ins!titnro
    
    End If
    
    
    ' Controlo si el empleado tiene el titulo asociado, si no lo asocio.
    
    StrSql = "SELECT emp_tit WHERE ternro = " & NroTercero & " AND titnro = " & NroTitulo
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
    
        StrSql = "INSERT INTO emp_tit(ternro,titnro) VALUES (" & NroTercero & "," & NroTitulo & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    
    End If
    

End Sub

Public Sub LineaModelo_640(ByVal strReg As String)

Dim Legajo          As String
Dim Causa           As String
Dim FAlta           As String
Dim FBaja           As String
Dim Estado          As String
Dim Sueldo          As String
Dim Vacaciones      As String
Dim Indemnizacion   As String
Dim Real            As String
Dim Fasrecofec      As String
Dim Empantnro       As String

Dim NroTercero        As Long
Dim NroCausa          As Long
Dim F_Alta            As String
Dim F_Baja            As String
Dim V_Estado          As Long
Dim V_Sueldo          As Long
Dim V_Vacaciones      As Long
Dim V_Indemnizacion   As Long
Dim V_Real            As Long
Dim V_Fasrecofec      As Long
Dim V_Empantnro       As Long


Dim rs      As New ADODB.Recordset
Dim rs_cau  As New ADODB.Recordset
Dim rs_emp  As New ADODB.Recordset


    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & Legajo
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
        Exit Sub
    Else
        NroTercero = rs!ternro
    End If

    If Not EsNulo(Causa) And Causa <> "N/A" Then
        StrSql = " SELECT caunro FROM causa WHERE caudes = '" & Causa & "'"
        OpenRecordset StrSql, rs_cau
        If Not rs_cau.EOF Then
            NroCausa = rs_sql!caunro
        Else
            StrSql = " INSERT INTO causa(caudes,causist,caudesvin,empnro) VALUES ('" & Causa & "',0,-1,1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            NroCausa = getLastIdentity(objConn, "causa")
            
        End If
    Else
        NroCausa = 0
    End If

    If FAlta <> "N/A" Then
        F_Alta = ConvFecha(FAlta)
    Else
        Exit Sub
    End If
    
    If FBaja <> "N/A" Then
        F_Baja = ConvFecha(FBaja)
    Else
        F_Baja = "Null"
    End If
    
    If Estado <> "N/A" Then
        If Estado = "Si" Or Estado = "SI" Then
            V_Estado = -1
        Else
            V_Estado = 0
        End If
    Else
        Exit Sub
    End If
    
    V_Sueldo = 0
    
    If (Sueldo = "Si") Or (Sueldo = "SI") Then
            V_Sueldo = -1
    End If
    
    V_Vacaciones = 0
    
    If (Vacaciones = "Si") Or (Vacaciones = "SI") Then
        V_Vacaciones = -1
    End If

    V_Indemnizacion = 0
    
    If (Indemnizacion = "Si") Or (Indemnizacion = "SI") Then
        V_Indemnizacion = -1
    End If
    
    V_Real = 0
    
    If (Real = "Si") Or (Real = "SI") Then
        V_Real = -1
    End If

    StrSql = "INSERT INTO fases(empleado,caunro,altfec,bajfec,estado,empantnro,"
    StrSql = StrSql & " (sueldo,vacaciones,indemnizacion,real,fasrecofec)"
    StrSql = StrSql & " VALUES(" & NroTercero & "," & NroCausa & "," & F_Alta & "," & F_Baja & ","
    StrSql = StrSql & V_Estado & ",0," & V_Sueldo & "," & v_vacacion & "," & V_Indemnizacion & ","
    StrSql = StrSql & V_Real & ",Null)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub

Public Sub LineaModelo_645(ByVal strReg As String)

' Levanta los Acumulados de Liquidacion Mensual


Dim Legajo          As String
Dim Anio            As String
Dim mes             As String
Dim TBruto          As String
Dim TNeto           As String
Dim TDescuentos     As String
Dim TVariables      As String
Dim Remuneracion1   As String
Dim Remuneracion2   As String
Dim Remuneracion3   As String
Dim Remuneracion4   As String

Dim NroTercero        As Long

Dim ACBruto         As Long
Dim ACNeto          As Long
Dim ACDescuentos    As Long
Dim ACVariables     As Long
Dim ACRemuneracion1 As Long
Dim ACRemuneracion2 As Long
Dim ACRemuneracion3 As Long
Dim ACRemuneracion4 As Long


Dim rs      As New ADODB.Recordset
Dim rs_emp  As New ADODB.Recordset
Dim rs_acu  As New ADODB.Recordset

    On Error GoTo SaltoLinea
    LineaCarga = LineaCarga + 1

    ACBruto = 7
    ACNeto = 6
    ACDescuentos = 12
    ACVariables = 2
    ACRemuneracion1 = 34
    ACRemuneracion2 = 35
    ACRemuneracion3 = 36
    ACRemuneracion4 = 37

    pos1 = 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Legajo = Mid(strReg, pos1, pos2 - pos1)
    
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & Legajo
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - El LEGAJO no Existe."
        Ok = False
        Exit Sub
    Else
        NroTercero = rs!ternro
    End If
    
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Anio = Mid(strReg, pos1, pos2 - pos1)
    
    If Anio = "N/A" Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - Debe Ingresar el AÑO!!."
        Ok = False
        Exit Sub
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    mes = Mid(strReg, pos1, pos2 - pos1)
    
    If mes = "N/A" Then
        LineaError.Writeline Mid(strReg, 1, Len(strReg))
        ErrCarga.Writeline "Linea: " & LineaCarga & " - Debe Ingresar el MES!!."
        Ok = False
        Exit Sub
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    TBruto = Mid(strReg, pos1, pos2 - pos1)

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    TNeto = Mid(strReg, pos1, pos2 - pos1)

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    TDescuentos = Mid(strReg, pos1, pos2 - pos1)

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    TVariables = Mid(strReg, pos1, pos2 - pos1)

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Remuneracion1 = Mid(strReg, pos1, pos2 - pos1)

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Remuneracion2 = Mid(strReg, pos1, pos2 - pos1)

    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Remuneracion3 = Mid(strReg, pos1, pos2 - pos1)

    pos1 = pos2 + 1
    pos2 = Len(strReg)
    Remuneracion4 = Mid(strReg, pos1, pos2)


    If (TBruto <> "N/A") And (TBruto <> "0") Then
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 0) & "Linea Insertando: " & LineaCarga
        Flog.Writeline
            
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & NroTercero & " AND ammes = " & mes
        StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACBruto
        OpenRecordset StrSql, rs_acu
            
        If rs_acu.EOF Then
            StrSql = "INSERT INTO acu_mes (ternro,acunro,amanio,ammes,ammonto,amcant,ammontoreal) values( "
            StrSql = StrSql & NroTercero & "," & ACBruto & "," & Anio & "," & mes & ","
            StrSql = StrSql & TBruto & "," & "30," & TBruto & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE acu_mes SET ammonto = " & rs_acu!ammonto + TBruto
            StrSql = StrSql & " WHERE ternro = " & NroTercero & " AND ammes = " & mes
            StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACBruto
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    End If
    
    If TNeto <> "N/A" And TNeto <> "0" Then
    
    
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & NroTercero & " AND ammes = " & mes
        StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACNeto
        OpenRecordset StrSql, rs_acu
            
        If rs_acu.EOF Then
            StrSql = "INSERT INTO acu_mes (ternro,acunro,amanio,ammes,ammonto,amcant,ammontoreal) values( "
            StrSql = StrSql & NroTercero & "," & ACNeto & "," & Anio & "," & mes & ","
            StrSql = StrSql & TNeto & "," & "30," & TNeto & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE acu_mes SET ammonto = " & rs_acu!ammonto + TNeto
            StrSql = StrSql & " WHERE ternro = " & NroTercero & " AND ammes = " & mes
            StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACNeto
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    End If
    
    If TDescuentos <> "N/A" And TDescuentos <> "0" Then
    
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & NroTercero & " AND ammes = " & mes
        StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACDescuentos
        OpenRecordset StrSql, rs_acu
            
        If rs_acu.EOF Then
        
            StrSql = "INSERT INTO acu_mes (ternro,acunro,amanio,ammes,ammonto,amcant,ammontoreal) values( "
            StrSql = StrSql & NroTercero & "," & ACDescuentos & "," & Anio & "," & mes & ","
            StrSql = StrSql & TDescuentos & "," & "30," & TDescuentos & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE acu_mes SET ammonto = " & rs_acu!ammonto + TDescuentos
            StrSql = StrSql & " WHERE ternro = " & NroTercero & " AND ammes = " & mes
            StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACDescuentos
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    End If
    
    If TVariables <> "N/A" And TVariables <> "0" Then
    
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & NroTercero & " AND ammes = " & mes
        StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACVariables
        OpenRecordset StrSql, rs_acu
            
        If rs_acu.EOF Then
            StrSql = "INSERT INTO acu_mes (ternro,acunro,amanio,ammes,ammonto,amcant,ammontoreal) values( "
            StrSql = StrSql & NroTercero & "," & ACVariables & "," & Anio & "," & mes & ","
            StrSql = StrSql & TVariables & "," & "30," & TVariables & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE acu_mes SET ammonto = " & rs_acu!ammonto + TVariables
            StrSql = StrSql & " WHERE ternro = " & NroTercero & " AND ammes = " & mes
            StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACVariables
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    End If
    
    If Remuneracion1 <> "N/A" And Remuneracion1 <> "0" Then
    
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & NroTercero & " AND ammes = " & mes
        StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACRemuneracion1
        OpenRecordset StrSql, rs_acu
            
        If rs_acu.EOF Then
            StrSql = "INSERT INTO acu_mes (ternro,acunro,amanio,ammes,ammonto,amcant,ammontoreal) values( "
            StrSql = StrSql & NroTercero & "," & ACRemuneracion1 & "," & Anio & "," & mes & ","
            StrSql = StrSql & Remuneracion1 & "," & "30," & Remuneracion1 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE acu_mes SET ammonto = " & rs_acu!ammonto + Remuneracion1
            StrSql = StrSql & " WHERE ternro = " & NroTercero & " AND ammes = " & mes
            StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACRemuneracion1
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    End If
    
    If Remuneracion2 <> "N/A" And Remuneracion2 <> "0" Then
    
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & NroTercero & " AND ammes = " & mes
        StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACRemuneracion2
        OpenRecordset StrSql, rs_acu
            
        If rs_acu.EOF Then
            StrSql = "INSERT INTO acu_mes (ternro,acunro,amanio,ammes,ammonto,amcant,ammontoreal) values( "
            StrSql = StrSql & NroTercero & "," & ACRemuneracion2 & "," & Anio & "," & mes & ","
            StrSql = StrSql & Remuneracion2 & "," & "30," & Remuneracion2 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE acu_mes SET ammonto = " & rs_acu!ammonto + Remuneracion2
            StrSql = StrSql & " WHERE ternro = " & NroTercero & " AND ammes = " & mes
            StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACRemuneracion2
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    If Remuneracion3 <> "N/A" And Remuneracion3 <> "0" Then
    
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & NroTercero & " AND ammes = " & mes
        StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACRemuneracion3
        OpenRecordset StrSql, rs_acu
            
        If rs_acu.EOF Then
            StrSql = "INSERT INTO acu_mes (ternro,acunro,amanio,ammes,ammonto,amcant,ammontoreal) values( "
            StrSql = StrSql & NroTercero & "," & ACRemuneracion3 & "," & Anio & "," & mes & ","
            StrSql = StrSql & Remuneracion3 & "," & "30," & Remuneracion3 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE acu_mes SET ammonto = " & rs_acu!ammonto + Remuneracion3
            StrSql = StrSql & " WHERE ternro = " & NroTercero & " AND ammes = " & mes
            StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACRemuneracion3
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    If Remuneracion4 <> "N/A" And Remuneracion4 <> "0" Then
    
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & NroTercero & " AND ammes = " & mes
        StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACRemuneracion4
        OpenRecordset StrSql, rs_acu
            
        If rs_acu.EOF Then
            StrSql = "INSERT INTO acu_mes (ternro,acunro,amanio,ammes,ammonto,amcant,ammontoreal) values( "
            StrSql = StrSql & NroTercero & "," & ACRemuneracion4 & "," & Anio & "," & mes & ","
            StrSql = StrSql & Remuneracion4 & "," & "30," & Remuneracion4 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE acu_mes SET ammonto = " & rs_acu!ammonto + Remuneracion4
            StrSql = StrSql & " WHERE ternro = " & NroTercero & " AND ammes = " & mes
            StrSql = StrSql & " AND amanio = " & Anio & " AND acunro = " & ACRemuneracion4
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    End If
    

  LineaOK.Writeline Mid(strReg, 1, Len(strReg))
  Ok = True
         
  If rs.State = adStateOpen Then
      rs.Close
  End If

  Exit Sub

SaltoLinea:

    LineaError.Writeline Mid(strReg, 1, Len(strReg))
    ErrCarga.Writeline "Linea: " & LineaCarga & " - " & Err.Description
    MyRollbackTrans
    Ok = False


End Sub

Public Sub LineaModelo_650(ByVal strReg As String, ByRef Ok As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Empleados - CODELCO
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.: 15/06/2005 CCR - Controlar Apellido2 y Nombre2<>N/A, sino ponerle vacios. Y la
'              actualizacion del RUT.
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Long
Dim pos2            As Long

Dim Legajo          As String   'LEGAJO                   -- empleado.empleg
Dim Apellido        As String   'APELLIDO                 -- empleado.terape y tercero.terape
Dim Apellido2       As String   'APELLIDO2                -- empleado.terape y tercero.terape
Dim nombre          As String   'NOMBRE                   -- empleado.ternom y tercero.ternom
Dim Nombre2         As String   'NOMBRE                   -- empleado.ternom y tercero.ternom
Dim Reporta         As String   'REPORTA A                -- empleado.empreporta
Dim Fing            As String   'FECHA DE INGRESO         -- terecro.terfecing
Dim FBaja           As String   '
Dim RUT             As String   'DOCUMENTO                -- ter_doc.nrodoc
Dim Email           As String   'EMAIL                    -- empleado.empemail
Dim NTUser          As String   'NT User                  -- Tidnro  = 32 en ter_doc
Dim Emppass         As String   'PASSWORD                 -- empleado.emppass

Dim ternro As Long

Dim NroTercero          As Long

Dim Nro_Legajo          As Long
Dim nro_tdocumento      As Long

Dim nro_tenro           As Long
Dim nro_estrnro         As Long

Dim Inserto_estr        As Boolean

Dim Str                 As String

Dim rs As New ADODB.Recordset
Dim rs_sql As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim rs_Tel As New ADODB.Recordset
Dim rs_tdo As New ADODB.Recordset


Dim TipoEstr As String
Dim EstrDesc As String

Dim IngresoDom          As Boolean

Dim rs_tdoc As New ADODB.Recordset
Dim rs_emp  As New ADODB.Recordset
Dim rs_ten  As New ADODB.Recordset
Dim rs_leg  As New ADODB.Recordset
Dim rs_rep  As New ADODB.Recordset
Dim rs_inst As New ADODB.Recordset
Dim rs_fas  As New ADODB.Recordset
Dim rs_Doc  As New ADODB.Recordset



Dim Sigue As Boolean
Dim ExisteLeg As Boolean
Dim CalculaLegajo As Boolean

Dim NroInstitucion As Long
Dim NroTdocum As Long

Dim F_Nacimiento        As String
Dim F_Fallecimiento     As String
Dim F_Alta              As String
Dim F_Baja              As String
Dim F_Ingreso           As String

Dim Fecha_Desde As String
Dim Fecha_Hasta As String

Dim empestado           As Long

Dim Actualizo_Supervisor As Boolean
Dim Supervisor As Long
Dim Ultimo As Boolean
Dim rs_Eva As New ADODB.Recordset
Dim rs_Eventos As New ADODB.Recordset


    On Error GoTo SaltoLinea

    Pisa = True
    
    StrSql = " SELECT * FROM confrep WHERE repnro = 120"
    StrSql = StrSql & " AND conftipo = 'TE'"
    StrSql = StrSql & " ORDER BY confnrocol"
    If rs_rep.State = adStateOpen Then rs_rep.Close
    OpenRecordset StrSql, rs_rep
    
    ' True indica que se hace por Descripcion. False por Codigo Externo
    
    Sigue = True 'Indica que si en el archivo viene mas de una vez un empleado, le crea las fases
    ExisteLeg = False
    
    NroColumna = 0
    RegLeidos = RegLeidos + 1
    LineaCarga = LineaCarga + 1
    
    Flog.Writeline
    FlogE.Writeline
    FlogP.Writeline

    'Texto = ": " & "Numero de Linea = " & RegLeidos
    'Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)

    
    'Recupero los Valores del Archivo
    NroColumna = NroColumna + 1
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador) - 1
    Legajo = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If Legajo = "N/A" Or Legajo = "" Then
    
        CalculaLegajo = True
        
    Else
        StrSql = "SELECT * FROM empleado WHERE empleado.empleg = " & Legajo
        OpenRecordset StrSql, rs_emp
        If (Not rs_emp.EOF) Then
            If (Not Sigue) Then
                Texto = ": " & " - El Empleado ya Existe."
                Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
                Ok = False
                Exit Sub
            Else
                NroTercero = rs_emp!ternro
                ExisteLeg = True
            End If
        End If
    End If
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Apellido = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Apellido = Left(Replace(Apellido, "'", "`"), 25)
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Apellido2 = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Apellido2 = Left(Replace(Apellido2, "'", "`"), 25)
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    nombre = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    nombre = Left(Replace(nombre, "'", "`"), 25)
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Nombre2 = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Nombre2 = Left(Replace(Nombre2, "'", "`"), 25)
    
    If (Apellido = "" Or Apellido = "N/A") And (nombre = "" Or nombre = "N/A") Then
        Texto = ": " & " - Debe Ingresar un Nombre y Apellido."
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        Ok = False
        Exit Sub
    End If
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Reporta = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Fing = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If (Fing = "N/A") Then
        F_Ingreso = "Null"
        Fecha_Desde = ""
    Else
        F_Ingreso = ConvFecha(Fing)
        Fecha_Desde = Fing
    End If
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FBaja = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If (FBaja = "N/A") Then
        F_Baja = "Null"
        Fecha_Hasta = ""
        empestado = -1
    Else
        F_Baja = ConvFecha(FBaja)
        Fecha_Hasta = FBaja
        empestado = 0
    End If
    
    ter_sexo = -1
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    RUT = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    RUT = Mid(RUT, 1, 30)
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Email = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Email <> "N/A" And Email <> "") Then
        NTUser = Mid(Email, 1, InStr(1, Email, "@") - 1)
    End If
    'Si nos dan el Password de la persona en la Interface General
    Emppass = "''"
    
    ' Inserto el Tercero, para luego poder insertar las estructuras
    If Trim(Apellido2) = "N/A" Then
        Apellido2 = ""
    End If
    
    If Trim(Nombre2) = "N/A" Then
        Nombre2 = ""
    End If
    
    If Not ExisteLeg Then

        StrSql = " INSERT INTO tercero(ternom,terape,ternom2,terape2,tersex)"
        StrSql = StrSql & " VALUES('" & nombre & "','" & UCase(Apellido) & "','" & Nombre2 & "','" & Apellido2 & "'," & ter_sexo & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

        NroTercero = getLastIdentity(objConn, "tercero")
        
        Texto = " Inserte en Tercero en la Base para el Legajo: " & Legajo
        Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
    Else
    
        StrSql = "UPDATE tercero SET terape = '" & Apellido & "', "
        StrSql = StrSql & " ternom = '" & nombre & "', "
        StrSql = StrSql & " terape2 = '" & Apellido2 & "', "
        StrSql = StrSql & " ternom2 = '" & Nombre2 & "'"
        StrSql = StrSql & " WHERE ternro = " & NroTercero
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Texto = " Modifique el Tercero en la Base para el Legajo: " & Legajo
        Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If

    ' Fin de la insercion del Legajo
  
    If Not rs_emp.EOF Then
        
        rs_rep.MoveFirst
    
    End If
  
    Ultimo = False
    Do While Not rs_rep.EOF And Not Ultimo
    
        NroColumna = NroColumna + 1
        pos1 = pos2 + 2
        pos2 = InStr(pos1, strReg, Separador) - 1
        If pos2 < 0 Then
            pos2 = Len(strReg)
            Ultimo = True
        End If
        EstrDesc = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
            
        If UCase(EstrDesc) <> "N/A" Then
            
            StrSql = "SELECT * FROM tipoestructura WHERE UPPER(tedabr) = '" & UCase(rs_rep!confetiq) & "'"
            OpenRecordset StrSql, rs_ten
            If rs_ten.EOF Then
                StrSql = "INSERT INTO tipoestructura(tedabr,tesist,tedepbaja,cenro) VALUES("
                StrSql = StrSql & "'" & UCase(rs_rep!confetiq) & "',0,0,1)"
                objConn.Execute StrSql, , adExecuteNoRecords
                nro_tenro = getLastIdentity(objConn, "tipoestructura")
            Else
                nro_tenro = rs_ten!Tenro
            End If
            
            If Nombre_EstructuraValido(nro_tenro, EstrDesc) Then
                Call ValidaEstructura(nro_tenro, EstrDesc, nro_estrnro, Inserto_estr)
                'Inserto las Estructuras
                'Call AsignarEstructura(nro_tenro, nro_estrnro, NroTercero, F_Ingreso, F_Baja)
                'Call AsignarEstructura_NEW(nro_tenro, nro_estrnro, NroTercero, F_Ingreso, F_Baja)
                Call Insertar_His_Estructura(nro_tenro, nro_estrnro, NroTercero, Fecha_Desde, Fecha_Hasta)
            Else
                Texto = ": " & " - Nombre de estructura " & EstrDesc & " Incorrecta para tipo " & nro_tenro & "."
                Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
            End If
        Else
            nro_estrnro = 0
        End If
        
        rs_rep.MoveNext
    Loop
  
  
    If Not ExisteLeg Then
        StrSql = " INSERT INTO empleado(empleg, ternro, terape, ternom, terape2,ternom2,empreporta,empnro,emppass,empest,empemail)"
        StrSql = StrSql & " VALUES(" & Legajo & "," & NroTercero & ",'" & Apellido & "','" & nombre
        StrSql = StrSql & "','" & Apellido2 & "','" & Nombre2 & "',"
        
        If UCase(Reporta) <> "N/A" Then
           Str = "SELECT ternro FROM empleado WHERE empleado.empleg = " & Reporta
           OpenRecordset Str, rs
           If Not rs.EOF Then
               StrSql = StrSql & rs!ternro & ",1," & Emppass & "," & empestado & ",'" & Email & "')"
           Else
               StrSql = StrSql & "Null,1," & Emppass & "," & empestado & ",'" & Email & "')"
           End If
        Else
           StrSql = StrSql & "Null,1," & Emppass & "," & empestado & ",'" & Email & "')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
                       
        Texto = "Inserte el Empleado - " & Legajo & " - " & Apellido & " - " & nombre
        Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
        
        Actualizo_Supervisor = False
        Supervisor = 0
    Else
                       
        StrSql = "UPDATE empleado SET  terape = '" & Apellido & "', "
        StrSql = StrSql & "ternom = '" & nombre & "', "
        StrSql = StrSql & "terape2 = '" & Apellido2 & "', "
        StrSql = StrSql & "ternom2 = '" & Nombre2 & "', "
        StrSql = StrSql & "empest = " & empestado & ", "
        StrSql = StrSql & "empemail = '" & Email & "'"
        'StrSql = StrSql & " WHERE empleado.ternro = " & NroTercero
        
        Actualizo_Supervisor = False
        Supervisor = 0
        If UCase(Reporta) <> "N/A" Then
           Str = "SELECT ternro FROM empleado WHERE empleado.ternro = " & Reporta
           OpenRecordset Str, rs
           If Not rs.EOF Then
               StrSql = StrSql & ", empreporta = " & rs!ternro
           End If
           Actualizo_Supervisor = True
           Supervisor = rs!ternro
        End If
        
        StrSql = StrSql & " WHERE ternro = " & NroTercero
        objConn.Execute StrSql, , adExecuteNoRecords

        Texto = " Se Modificaron los Datos del Empleado " & Legajo
        Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    
    'FGZ - 19/01/2006
    'Actualizo la relacion supervisor - Supervisado en todos los eventos activos del legajo
    If Actualizo_Supervisor Then
        'Busco todos los eventos no aprobados (que tienen al menos una cabecera no aprobada)
        StrSql = " SELECT distinct(evaevento.evaevenro) FROM evaevento "
        StrSql = StrSql & " INNER JOIN evacab ON evaevento.evaevenro = evacab.evaevenro "
        StrSql = StrSql & " WHERE evacab.cabaprobada = 0"
        If rs_Eventos.State = adStateOpen Then rs_Eventos.Close
        OpenRecordset StrSql, rs_Eventos
        Do While Not rs_Eventos.EOF
            StrSql = " SELECT evacabnro FROM evacab "
            StrSql = StrSql & " WHERE   evacab.empleado = " & NroTercero
            StrSql = StrSql & " AND evacab.evaevenro = " & rs_Eventos!evaevenro
            If rs_Eva.State = adStateOpen Then rs_Eva.Close
            OpenRecordset StrSql, rs_Eva
            Do While Not rs_Eva.EOF
                StrSql = "UPDATE evadetevldor SET "
                StrSql = StrSql & " evaluador = " & Supervisor
                StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & rs_Eva!evacabnro
                StrSql = StrSql & " AND evadetevldor.evatevnro = 2"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                rs_Eva.MoveNext
            Loop

            rs_Eventos.MoveNext
        Loop
    End If
    

    'Inserto el Registro correspondiente en ter_tip
    
    If Not ExisteLeg Then
    
        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & NroTercero & ",1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If

    ' Inserto los Documentos
    
    If Not ExisteLeg Then
    
    
        StrSql = "SELECT * FROM tipodocu WHERE UPPER(tidsigla) = 'RUT'"
        OpenRecordset StrSql, rs_tdo
        If Not rs_tdo.EOF Then
            
            NroTdocum = rs_tdo!tidnro
        
        Else
        
            StrSql = "SELECT * FROM institucion WHERE UPPER(instdes) = 'NO INFORMADO'"
            OpenRecordset StrSql, rs_inst
            If rs_inst.EOF Then
            
                StrSql = "INSERT INTO institucion(instdes,instabre) VALUES("
                StrSql = StrSql & "'NO INFORMADO','N/I')"
                objConn.Execute StrSql, , adExecuteNoRecords
                NroInstitucion = getLastIdentity(objConn, "institucion")
                
            Else
            
                NroInstitucion = rs_inst!instnro
            
            End If
        
            StrSql = "INSERT INTO tipodocu(tidnom,tidsigla,instnro) VALUES("
            StrSql = StrSql & "'RUT','RUT'," & NroInstitucion & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            NroTdocum = getLastIdentity(objConn, "tipodocu")
        
        End If
    
    
        If RUT <> "" And UCase(RUT) <> "N/A" Then
            StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
            StrSql = StrSql & " VALUES(" & NroTercero & "," & NroTdocum & ",'" & RUT & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = "Inserte el RUT - "
            Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
        Else
            Texto = "No se Inserte el RUT - "
            Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
        End If
        
        If (Email <> "N/A" And Email <> "") Then
            StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
            StrSql = StrSql & " VALUES(" & NroTercero & ",32,'" & NTUser & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = " Inserte el NTUser - "
            Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
        End If
        
    Else ' existe el empleado
        
        If RUT <> "" And UCase(RUT) <> "N/A" Then
        
            StrSql = " UPDATE ter_doc SET nrodoc = '" & RUT
            StrSql = StrSql & "' WHERE tidnro = 21 AND ternro = " & NroTercero
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = " Modifique el RUT del Empleado: " & Legajo
            Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
        Else
            Texto = " No se Modifique el RUT del Empleado: " & Legajo
            Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
        End If
        
        If (Email <> "N/A" And Email <> "") Then
            StrSql = "SELECT * FROM ter_doc WHERE ternro = " & NroTercero & " AND tidnro =  32"
            OpenRecordset StrSql, rs_Doc
            If rs_Doc.EOF Then
                StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                StrSql = StrSql & " VALUES(" & NroTercero & ",32,'" & NTUser & "')"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Texto = " Inserte el NTUser del Empleado: " & Legajo
                Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
            Else
                StrSql = " UPDATE ter_doc SET nrodoc = '" & NTUser
                StrSql = StrSql & "' WHERE tidnro = 32 AND ternro = " & NroTercero
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Texto = " Modifique el NTUser del Empleado: " & Legajo
                Call Escribir_Log("flog", LineaCarga, NroColumna, Texto, Tabs, strReg)
            End If
        End If

    End If
    If rs.State = adStateOpen Then rs.Close
  
    If Not ExisteLeg Then
     ' Inserto las Fases
     
         StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
         StrSql = StrSql & " VALUES( " & NroTercero & "," & F_Ingreso & "," & F_Baja & ","
         If nro_causabaja <> 0 Then
                StrSql = StrSql & nro_causabaja & ","
         Else
                StrSql = StrSql & "null" & ","
         End If
         StrSql = StrSql & empestado & ",-1,-1,-1,-1,-1)"
         objConn.Execute StrSql, , adExecuteNoRecords
     
    Else
    
'        StrSql = "SELECT fasnro from fases  WHERE empleado =" & NroTercero
'        StrSql = StrSql & " and ( (altfec <= " & F_Ingreso
'        StrSql = StrSql & " and " & F_Ingreso & " <= bajfec) "
        'Si esta cargada la fecha de baja verifico
'        If Len(F_Baja) > 0 Then
'            StrSql = StrSql & " OR (altfec <= " & F_Baja
'            StrSql = StrSql & " and " & F_Baja & "<= bajfec) "
'            StrSql = StrSql & " OR (" & F_Ingreso
'            StrSql = StrSql & " <= altfec and altfec <= " & F_Baja & ") "
'            OpenRecordset StrSql, rs_fas
        
'        End If
'
'        If rs_fas.EOF Then
'
'            StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
'            StrSql = StrSql & " VALUES( " & NroTercero & "," & F_Ingreso & "," & F_Baja & ","
'            If nro_causabaja <> 0 Then
'                   StrSql = StrSql & nro_causabaja & ","
'            Else
'                   StrSql = StrSql & "null" & ","
'            End If
 '           StrSql = StrSql & empestado & ",-1,-1,-1,-1,-1)"
'            objConn.Execute StrSql, , adExecuteNoRecords
        
'        Else
'
'            LineaError.Writeline Mid(strReg, 1, Len(strReg))
'            ErrCarga.Writeline "Linea: " & LineaCarga & " - Hay Problemas con las Fechas de Ingreso/Egreso del Legajo: " & Legajo
'            Ok = False
'            Exit Sub
'
'        End If
    End If
  
    Texto = ": " & "Linea procesada correctamente "
    Call Escribir_Log("flogp", LineaCarga, NroColumna, Texto, Tabs + 1, strReg)
    Ok = True

    If rs.State = adStateOpen Then rs.Close
  Exit Sub

SaltoLinea:
    Texto = ": " & " - " & Err.Description
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)

    Texto = ": " & " - Ultimo SQl Ejecutado: " & StrSql
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    Ok = False
End Sub

Public Sub LineaModelo_651(ByVal strReg As String, ByRef Ok As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de PSW
' Autor      :
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Legajo          As String   'LEGAJO                        -- empleado.empleg
Dim RUT             As String


Dim ternro As Long

Dim pos1 As Long
Dim pos2 As Long

Dim NroTercero          As Long
Dim NroLegajo           As Long


Dim rs As New ADODB.Recordset
Dim rs_tdo As New ADODB.Recordset


    On Error GoTo SaltoLinea
    
    NroColumna = 0
    RegLeidos = RegLeidos + 1
    LineaCarga = LineaCarga + 1
    
    Flog.Writeline
    FlogE.Writeline
    FlogP.Writeline
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    RUT = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 1
    pos2 = Len(strReg)
    Emppass = Trim(Mid(strReg, pos1, pos2 - pos1))
    
    'Tipo de Documento
    StrSql = "SELECT * FROM tipodocu WHERE tidsigla = 'RUT'"
    OpenRecordset StrSql, rs_tdo
    
    'Busca el Tercero
    StrSql = "SELECT empleado.ternro FROM ter_doc "
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = ter_doc.ternro"
    StrSql = StrSql & " WHERE tidnro = " & rs_tdo!tidnro & " AND nrodoc = '" & RUT & "'"
    OpenRecordset StrSql, rs
    
    If rs.EOF Then Exit Sub
    
    NroTercero = rs!ternro
    StrSql = "UPDATE empleado SET emppass = '" & Emppass & "' WHERE empleado.ternro = " & NroTercero
    objConn.Execute StrSql, , adExecuteNoRecords
         
    Texto = ": " & "Linea procesada correctamente "
    Call Escribir_Log("flogp", LineaCarga, NroColumna, Texto, Tabs + 1, strReg)
    Ok = True
    
    If rs.State = adStateOpen Then rs.Close
    Exit Sub
    
SaltoLinea:
    Texto = ": " & " - " & Err.Description
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)

    Texto = ": " & " - Ultimo SQl Ejecutado: " & StrSql
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    Ok = False
End Sub


Public Sub LineaModelo_653(ByVal strReg As String, ByRef Ok As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Reporta A
' Autor      :
' Fecha      :
' Ultima Mod.: FGZ - 26/01/2006
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim pos1            As Long
Dim pos2            As Long

Dim Legajo          As String   'LEGAJO                   -- empleado.empleg
Dim Apellido        As String   'APELLIDO                 -- empleado.terape y tercero.terape
Dim Apellido2       As String   'APELLIDO2                -- empleado.terape y tercero.terape
Dim nombre          As String   'NOMBRE                   -- empleado.ternom y tercero.ternom
Dim Nombre2         As String   'NOMBRE                   -- empleado.ternom y tercero.ternom
Dim Reporta         As String   'REPORTA A                -- empleado.empreporta
Dim Fing            As String   'FECHA DE INGRESO         -- terecro.terfecing
Dim FBaja           As String
Dim RUT             As String   'DOCUMENTO                -- ter_doc.nrodoc
Dim Email           As String   'EMAIL                    -- empleado.empemail
Dim Emppass         As String   'PASSWORD                 -- empleado.emppass

Dim ternro As Long

Dim NroTercero          As Long
Dim NroLegajo           As Long


Dim rs As New ADODB.Recordset
Dim rs_tdo As New ADODB.Recordset
Dim rs_rep As New ADODB.Recordset
Dim rs_emp As New ADODB.Recordset


Dim Actualizo_Supervisor As Boolean
Dim Supervisor As Long
Dim rs_Eva As New ADODB.Recordset
Dim rs_Eventos As New ADODB.Recordset

    
    
    On Error GoTo SaltoLinea
    
    Sigue = True
    Pisa = True
    
    NroColumna = 0
    RegLeidos = RegLeidos + 1
    LineaCarga = LineaCarga + 1
    
    Flog.Writeline
    FlogE.Writeline
    FlogP.Writeline
    
    NroColumna = NroColumna + 1
    pos1 = 1
    pos2 = InStr(pos1, strReg, Separador) - 1
    Legajo = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If Legajo = "N/A" Or Legajo = "" Then
    
        CalculaLegajo = True
        
    Else
        StrSql = "SELECT * FROM empleado WHERE empleado.empleg = " & Legajo
        OpenRecordset StrSql, rs_emp
        If (Not rs_emp.EOF) Then
            If (Not Sigue) Then
                Texto = ": " & " - El Empleado ya Existe."
                Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
                Ok = False
                Exit Sub
            Else
                NroTercero = rs_emp!ternro
                ExisteLeg = True
            End If
        End If
    End If
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Apellido = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Apellido = Replace(Apellido, "'", "`")
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Apellido2 = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Apellido2 = Replace(Apellido2, "'", "`")
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    nombre = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    nombre = Replace(nombre, "'", "`")
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Nombre2 = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    Nombre2 = Replace(Nombre2, "'", "`")
    
    If (Apellido = "" Or Apellido = "N/A") And (nombre = "" Or nombre = "N/A") Then
        Texto = ": " & " - Debe Ingresar un Nombre y Apellido."
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        Ok = False
        Exit Sub
    End If
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Reporta = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Fing = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If (Fing = "N/A") Then
        F_Ingreso = "Null"
    Else
        F_Ingreso = ConvFecha(Fing)
    End If
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    FBaja = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    If (FBaja = "N/A") Then
        F_Baja = "Null"
        empestado = -1
    Else
        F_Baja = ConvFecha(FBaja)
        empestado = 0
    End If
    
    ter_sexo = -1
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    RUT = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    RUT = Mid(RUT, 1, 30)
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, Separador) - 1
    Email = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    
    'Si nos dan el Password de la persona en la Interface General
    Emppass = "''"
    
    'Tipo de Documento
    
    StrSql = "SELECT * FROM tipodocu WHERE tidsigla = 'RUT'"
    OpenRecordset StrSql, rs_tdo
    
    ' Busca el Tercero
    StrSql = "SELECT empleado.ternro FROM ter_doc "
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = ter_doc.ternro"
    StrSql = StrSql & " WHERE tidnro = " & rs_tdo!tidnro & " AND nrodoc = '" & RUT & "'"
    OpenRecordset StrSql, rs
    
    If rs.EOF Then Exit Sub
    
    NroTercero = rs!ternro

    Actualizo_Supervisor = False
    Supervisor = 0

    If UCase(Reporta) <> "N/A" Then
       StrSql = "SELECT ternro FROM empleado WHERE empleado.empleg = " & Reporta
       OpenRecordset StrSql, rs_rep
       If Not rs_rep.EOF Then
       
            StrSql = "UPDATE empleado SET empreporta = " & rs_rep!ternro & " WHERE empleado.ternro = " & NroTercero
            objConn.Execute StrSql, , adExecuteNoRecords
           
            Actualizo_Supervisor = True
            Supervisor = rs_rep!ternro
           
           
            'FGZ - 19/01/2006
            'Actualizo la relacion supervisor - Supervisado en todos los eventos activos del legajo
            If Actualizo_Supervisor Then
                'Busco todos los eventos no aprobados (que tienen al menos una cabecera no aprobada)
                StrSql = " SELECT distinct(evaevento.evaevenro) FROM evaevento "
                StrSql = StrSql & " INNER JOIN evacab ON evaevento.evaevenro = evacab.evaevenro "
                StrSql = StrSql & " WHERE evacab.cabaprobada = 0"
                If rs_Eventos.State = adStateOpen Then rs_Eventos.Close
                OpenRecordset StrSql, rs_Eventos
                Do While Not rs_Eventos.EOF
                    StrSql = " SELECT evacabnro FROM evacab "
                    StrSql = StrSql & " WHERE   evacab.empleado = " & NroTercero
                    StrSql = StrSql & " AND evacab.evaevenro = " & rs_Eventos!evaevenro
                    If rs_Eva.State = adStateOpen Then rs_Eva.Close
                    OpenRecordset StrSql, rs_Eva
                    Do While Not rs_Eva.EOF
                        StrSql = "UPDATE evadetevldor SET "
                        StrSql = StrSql & " evaluador = " & Supervisor
                        StrSql = StrSql & " WHERE evadetevldor.evacabnro = " & rs_Eva!evacabnro
                        StrSql = StrSql & " AND evadetevldor.evatevnro = 2"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        rs_Eva.MoveNext
                    Loop
        
                    rs_Eventos.MoveNext
                Loop
            End If
           
       End If
    End If
         
    Texto = ": " & "Linea procesada correctamente "
    Call Escribir_Log("flogp", LineaCarga, NroColumna, Texto, Tabs + 1, strReg)
    Ok = True
    
    If rs.State = adStateOpen Then rs.Close
    If rs_rep.State = adStateOpen Then rs_rep.Close
    Exit Sub
    
SaltoLinea:
    Texto = ": " & " - " & Err.Description
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)

    Texto = ": " & " - Ultimo SQl Ejecutado: " & StrSql
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    Ok = False
End Sub

Public Sub LineaModelo_652(ByVal strReg As String, ByRef Ok As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Historico de Estructuras
' Autor      : FGZ
' Fecha      : 21/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim RUT             As String   'LEGAJO                        -- empleado.empleg
Dim Estructura      As String   'Estructura                    -- his_estructura.estrnro
Dim TipoEstructura  As String   'Tipo de Estructura            -- his_estructura.tenro
Dim FAlta           As String   'Fecha Desde en la Estructura  -- his_estructura.htetdesde
Dim FBaja           As String   'Fecha Hasta en la Estructura  -- his_estructura.htethasta

Dim ternro As Long

Dim pos1 As Long
Dim pos2 As Long

Dim NroTercero          As Long
Dim NroLegajo           As Long
Dim nro_estructura      As Long
Dim F_Alta              As String
Dim F_Baja              As String

Dim Fecha_Desde As String
Dim Fecha_Hasta As String

Dim Inserto_estr        As Boolean

Dim rs As New ADODB.Recordset
Dim rs_sql As New ADODB.Recordset
Dim rs_tes As New ADODB.Recordset
Dim rs_tdo As New ADODB.Recordset


Dim nro_tenro As Long

' True indica que se hace por Descripcion. False por Codigo Externo

Dim EstrDesc             As Boolean   'Sucursal                 -- his_estructura

    On Error GoTo SaltoLinea

    EstrDesc = True ' Indica si la Validacion de la Estructura es por Descripcion o Codigo Externo
    Pisa = True

    Ok = True
    
    NroColumna = 0
    RegLeidos = RegLeidos + 1
    LineaCarga = LineaCarga + 1
    
    Flog.Writeline
    FlogE.Writeline
    FlogP.Writeline
    
    NroColumna = NroColumna + 1
    pos1 = 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    TipoEstructura = Mid(strReg, pos1, pos2 - pos1)
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    RUT = Mid(strReg, pos1, pos2 - pos1)
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    Estructura = Mid(strReg, pos1, pos2 - pos1)
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strReg, Separador)
    FAlta = Mid(strReg, pos1, pos2 - pos1)
    
    If FAlta = "N/A" Or FAlta = "" Then
        F_Alta = "Null"
        Fecha_Desde = ""
    Else
        F_Alta = ConvFecha(FAlta)
        Fecha_Desde = FAlta
    End If
    
    NroColumna = NroColumna + 1
    pos1 = pos2 + 1
    pos2 = Len(strReg) + 1
    FBaja = Mid(strReg, pos1, pos2 - pos1)
    
    If FBaja = "N/A" Or FBaja = "" Then
        F_Baja = "Null"
        Fecha_Hasta = ""
    Else
        F_Baja = ConvFecha(FBaja)
        Fecha_Desde = FBaja
    End If
    
    'Valida que los campos obligatorios este cargados
    If TipoEstructura = "" Or TipoEstructura = "N/A" Or RUT = "" Or RUT = "N/A" Or Estructura = "" Or Estructura = "N/A" Or FAlta = "" Or FAlta = "N/A" Then
        Texto = ": " & " - Faltan campos obligatorios, revisar: Tipo de estructura, estrucrura, RUT y Fecha de alta."
        Call Escribir_Log("floge", LineaCarga, 0, Texto, Tabs, strReg)
        Exit Sub
    End If
    
    StrSql = "SELECT * FROM tipodocu WHERE tidsigla = 'RUT'"
    OpenRecordset StrSql, rs_tdo
    
    ' Busca el Tercero
    StrSql = "SELECT empleado.ternro FROM ter_doc "
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = ter_doc.ternro"
    StrSql = StrSql & " WHERE tidnro = " & rs_tdo!tidnro & " AND nrodoc = '" & RUT & "'"
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
        
        Ok = False
        Exit Sub
    
    End If
    
    NroTercero = rs!ternro

    StrSql = "SELECT tenro FROM tipoestructura WHERE UPPER(tedabr) = '" & UCase(TipoEstructura) & "'"
    OpenRecordset StrSql, rs_tes
    If rs_tes.EOF Then
        StrSql = "INSERT INTO tipoestructura(tedabr,tesist,tedepbaja,cenro) VALUES('" & UCase(TipoEstructura) & "',0,0,1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        nro_tenro = getLastIdentity(objConn, "tipoestructura")
    Else
        nro_tenro = rs_tes!Tenro
    End If


    ' Validacion y Creacion de la Sucursal (junto con sus Complementos)
    If Estructura <> "N/A" Then
        If Nombre_EstructuraValido(nro_tenro, Mid(Estructura, 1, 60)) Then
            If EstrDesc Then
                Call ValidaEstructura(nro_tenro, Mid(Estructura, 1, 60), nro_estructura, Inserto_estr)
            Else
                Call ValidaEstructuraCodExt(nro_tenro, Mid(Estructura, 1, 20), nro_estructura, Inserto_estr)
            End If
            Call Insertar_His_Estructura(nro_tenro, nro_estructura, NroTercero, Fecha_Desde, Fecha_Hasta)
        Else
            Texto = ": " & " - Nombre de estructura " & Estructura & " Incorrecta para tipo " & nro_tenro & "."
            Call Escribir_Log("floge", LineaCarga, 3, Texto, Tabs, strReg)
        End If
    End If

'    ' Validacion y Creacion de la Sucursal (junto con sus Complementos)
'    If Estructura <> "N/A" Then
'        If EstrDesc Then
'            Call ValidaEstructura(nro_tenro, Mid(Estructura, 1, 60), nro_estructura, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(nro_tenro, Mid(Estructura, 1, 20), nro_estructura, Inserto_estr)
'        End If
'    End If
'  ' Inserto las Estructuras
'  Call AsignarEstructura(nro_tenro, nro_estructura, NroTercero, F_Alta, F_Baja)
         
    Texto = ": " & "Linea procesada correctamente "
    Call Escribir_Log("flogp", LineaCarga, NroColumna, Texto, Tabs + 1, strReg)
    Ok = True

    If rs.State = adStateOpen Then rs.Close
    Exit Sub

SaltoLinea:
    Texto = ": " & " - " & Err.Description
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)

    Texto = ": " & " - Ultimo SQl Ejecutado: " & StrSql
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    Ok = False
End Sub


Public Sub CalcularCUIL(ByRef Cuil As String)

End Sub


Public Sub Insertar_His_Estructura(ByVal Tenro As Long, ByVal Estrnro As Long, ByVal Tercero As Long, ByVal FechaDesde As String, ByVal FechaHasta As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta en his_estructura la estructura del tipo especificada
'               en el rango de fechas especificado. Si ya existe una estructura del mismo tipo abierta ==> se cierra el dia anterior al la
'               fecha_desde y se abre la nueva estructura.
' Autor      : FGZ
' Fecha      : 14/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_His_Estructura As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Fecha_Desde As Date
Dim Fecha_Hasta As Date


Fecha_Desde = CDate(FechaDesde)
If Not EsNulo(FechaHasta) Then
    Fecha_Hasta = CDate(FechaHasta)
End If

    If Estrnro <> 0 Then
        If nro_ModOrg <> 0 Then
            StrSql = "SELECT * FROM adptte_estr WHERE tplatenro = " & nro_ModOrg & " AND tenro = " & TipoEstr
            OpenRecordset StrSql, rs
            If rs.EOF Then
                tplaorden = tplaorden + 1
                StrSql = "INSERT INTO adptte_estr(tplatenro,tenro,tplaestroblig,tplaestrorden) VALUES (" & nro_ModOrg & "," & TipoEstr & ",0," & tplaorden & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If

        StrSql = " SELECT * FROM his_estructura "
        StrSql = StrSql & " WHERE ternro = " & Tercero
        StrSql = StrSql & " AND tenro = " & Tenro
        StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Fecha_Desde) & ") "
        StrSql = StrSql & " AND ((" & ConvFecha(Fecha_Desde) & " <= htethasta) or (htethasta is null))"
        StrSql = StrSql & " ORDER BY htetdesde "
        If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
        OpenRecordset StrSql, rs_His_Estructura
        If Not rs_His_Estructura.EOF Then
            rs_His_Estructura.MoveLast
            If rs_His_Estructura!Estrnro = Estrnro Then
                'la estructura es la misma, reviso las fechas
                If Fecha_Desde < rs_His_Estructura!htetdesde Then
                    StrSql = "UPDATE his_estructura SET htetdesde = " & ConvFecha(CDate(Fecha_Desde))
                    StrSql = StrSql & " WHERE ternro = " & Tercero
                    StrSql = StrSql & " AND tenro = " & Tenro
                    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Fecha_Desde) & ") "
                    StrSql = StrSql & " AND ((" & ConvFecha(Fecha_Desde) & " <= htethasta) or (htethasta is null))"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                If Not EsNulo(FechaHasta) Then
                    If Not EsNulo(rs_His_Estructura!htethasta) Then
                        If Fecha_Hasta < rs_His_Estructura!htethasta Then
                            StrSql = "UPDATE his_estructura SET htethasta = " & ConvFecha(CDate(Fecha_Hasta))
                            StrSql = StrSql & " WHERE ternro = " & Tercero
                            StrSql = StrSql & " AND tenro = " & Tenro
                            StrSql = StrSql & " AND estrnro = " & Estrnro
                            StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
                            StrSql = StrSql & " AND htethasta = " & ConvFecha(rs_His_Estructura!htethasta)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'no actualizo porque la que estaba abarca mayor rango
                            FlogE.Writeline Espacios(Tabulador * 3) & "Estructura no insertada. No actualizo porque la que estaba abarca mayor rango"
                        End If
                    Else
                        StrSql = "UPDATE his_estructura SET htethasta = " & ConvFecha(CDate(Fecha_Hasta))
                        StrSql = StrSql & " WHERE ternro = " & Tercero
                        StrSql = StrSql & " AND estrnro = " & Estrnro
                        StrSql = StrSql & " AND tenro = " & Tenro
                        StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
                        StrSql = StrSql & " AND htethasta is null"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                Else
                    If Not EsNulo(rs_His_Estructura!htethasta) Then
                        'si noy hay alguna otra estructura con fechas desde posterior ==> actualizo sino queda como esta
                        StrSql = " SELECT * FROM his_estructura "
                        StrSql = StrSql & " WHERE ternro = " & Tercero
                        StrSql = StrSql & " AND tenro = " & Tenro
                        StrSql = StrSql & " AND (htetdesde > " & ConvFecha(Fecha_Desde) & ") "
                        StrSql = StrSql & " ORDER BY htetdesde "
                        If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            'Entonces no actualizo
                            
                        Else
                            StrSql = "UPDATE his_estructura SET htethasta = NULL "
                            StrSql = StrSql & " WHERE ternro = " & Tercero
                            StrSql = StrSql & " AND tenro = " & Tenro
                            StrSql = StrSql & " AND estrnro = " & Estrnro
                            StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
                            StrSql = StrSql & " AND htethasta = " & ConvFecha(rs_His_Estructura!htethasta)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    Else
    '                    StrSql = "UPDATE his_estructura SET htethasta = NULL "
    '                    StrSql = StrSql & " WHERE ternro = " & Tercero
    '                    StrSql = StrSql & " AND tenro = " & Tenro
    '                    StrSql = StrSql & " AND estrnro = " & Estrnro
    '                    StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
    '                    objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            Else    'tiene abierta otra estructura, la cierro y le abro la nueva
                If Fecha_Desde = rs_His_Estructura!htetdesde Then
                    StrSql = "UPDATE his_estructura SET estrnro = " & Estrnro
                    If Not EsNulo(FechaHasta) Then
                        StrSql = StrSql & " ,htethasta = " & ConvFecha(CDate(Fecha_Hasta))
                    Else
                        StrSql = StrSql & " ,htethasta = null"
                    End If
                    StrSql = StrSql & " WHERE ternro = " & Tercero
                    StrSql = StrSql & " AND tenro = " & Tenro
                    StrSql = StrSql & " AND (htetdesde = " & ConvFecha(Fecha_Desde) & ") "
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = "UPDATE his_estructura SET htethasta = " & ConvFecha(CDate(Fecha_Desde - 1))
                    StrSql = StrSql & " WHERE ternro = " & Tercero
                    StrSql = StrSql & " AND tenro = " & Tenro
                    StrSql = StrSql & " AND estrnro = " & rs_His_Estructura!Estrnro
                    StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_His_Estructura!htetdesde)
                    If Not EsNulo(rs_His_Estructura!htethasta) Then
                        StrSql = StrSql & " AND htethasta = " & ConvFecha(rs_His_Estructura!htethasta)
                    Else
                        StrSql = StrSql & " AND htethasta is null"
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                
                    'Inserto la nueva estructura
                    If Not EsNulo(FechaHasta) Then
                        StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                        StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & "," & ConvFecha(Fecha_Hasta) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
                        StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        Else
            StrSql = " SELECT * FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & Tercero
            StrSql = StrSql & " AND tenro = " & Tenro
            StrSql = StrSql & " AND (htetdesde > " & ConvFecha(Fecha_Desde) & ") "
            StrSql = StrSql & " ORDER BY htetdesde "
            If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
            OpenRecordset StrSql, rs_His_Estructura
            If Not rs_His_Estructura.EOF Then
                'Inserto la nueva estructura
                If Not EsNulo(FechaHasta) Then
                    If FechaHasta < rs_His_Estructura!htetdesde Then
                        StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                        StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & "," & ConvFecha(Fecha_Hasta) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'no way
                        FlogE.Writeline Espacios(Tabulador * 3) & "Estructura no insertada. El rango se superpone con una ya existente."
                    End If
                Else
                    StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                    StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & "," & ConvFecha(rs_His_Estructura!htetdesde - 1) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else
                'Inserto la nueva estructura
                If Not EsNulo(FechaHasta) Then
                    StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde,htethasta) VALUES("
                    StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & "," & ConvFecha(Fecha_Hasta) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = " INSERT INTO his_estructura(ternro,estrnro,tenro,htetdesde) VALUES("
                    StrSql = StrSql & Tercero & "," & Estrnro & "," & Tenro & "," & ConvFecha(Fecha_Desde) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
        
    End If
    
    
    
'cierro y libero
If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
Set rs_His_Estructura = Nothing
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub


