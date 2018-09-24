Attribute VB_Name = "MdlInterfacesPortugal"

Public Sub LineaModelo_612(ByVal strReg As String, ByRef OK As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Empleados - Portugal
' Autor      : Gonzalez Nicolás
' Fecha      : 18/01/2012
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
'FORMATO
'Data Venc.Contrato  (Falta encontrar esta variable)
'---------------------
Dim pos1            As Long
Dim pos2            As Long

Dim Legajo          As String   'LEGAJO              |Nº mec                      -- empleado.empleg
Dim Apellido        As String   'APELLIDO            |Sobrenome                   -- empleado.terape y tercero.terape
Dim nombre          As String   'NOMBRE              |Nome                        -- empleado.ternom y tercero.ternom
Dim Fnac            As String   'FNAC                |Data Nasc.                  -- tercero.terfecna
Dim Nacionalidad    As String   'Nacionalidad        |País de Nascimento          -- tercero.nacionalnro
Dim PaisNac         As String   'Pais de Nacimiento  |Nacionalidade               -- tercero.paisnro
Dim EstCivil        As String   'Est.Civil           |Est.Civil                   -- tercero.estcivnro
Dim Sexo            As String   'Sexo                |Sexo                        -- tercero.tersex
Dim FAlta           As String   'Fec. de Alta        |Data de Admissão            -- empleado.empfaltagr y fases.altfec
Dim NivEst          As String   'Nivel de Estudio    |Nivel de Estudo             -- empleado.nivnro
Dim Tdocu           As String   'Tipo Documento      |tipo Documento              -- ter_dpc.tidnro (DU)
Dim Ndocu           As String   'Nro. Documento      |Nro.Documento               -- ter_doc.nrodoc
Dim Calle           As String   'Calle               |Rua                         -- detdom.calle
Dim Nro             As String   'Número              |N°                          -- detdom.nro
Dim Piso            As String   'Piso                |Andar                       -- detdom.piso
Dim Partido         As String   'Partido             |Conselho                    -- detdom.partnro
Dim Cpostal         As String   'Cpostal             |Cód.Postal                  -- detdom.codigopostal
Dim Localidad       As String   'Localidad           |Freguesia                   -- detdom.locnro
Dim Provincia       As String   'Provincia           |Distrito                    -- detdom.provnro
Dim Pais            As String   'Pais                |País                        -- detdom.paisnro
Dim Email           As String   'E-mail              |E-mail                      -- empleado.empemail
Dim Convenio        As String   'Convenio            |Acordo Colectivo            -- his_estructura.estrnro
Dim categoria       As String   'Categoria           |Categoría                   -- his_estructura.estrnro
Dim CajaJub         As String   'Caja de Jubilacion  |Contribuição Previdenciária -- his_estructura.estrnro
Dim Contrato        As String   'Contrato            |tipo Contrato               -- his_estructura.estrnro
'----
'Data Venc.Contrato
'----
Dim RegHorario      As String   'Regimen Horario     |Horário                     -- his_estructura.estrnro
Dim FormaLiq        As String   'Forma de Liquidacion|Forma de Cálculo            -- his_estructura.estrnro
Dim FormaPago       As String   'Forma de Pago       |Forma de Pagto              -- formapago.fpagdescabr
Dim BancoPago       As String   'Banco Pago          |Banco Pagto                 -- his_estructura.estrnro, formapago.fpagbanc (siempre y cuando el Banco sea <> 0) y ctabancaria.banco
Dim NroCuenta       As String   'Nro. Cuenta         |Nº Conta                    -- ctabancario.ctabnro
Dim Estado          As String   'Estado              |Estado do Funcionário       -- empleado.empest y fases.estado
Dim CausaBaja       As String   'Causa de Baja       |Motivo de Demissão          -- fases.caunro
Dim FBaja           As String   'Fecha de Baja       |Data de Demissão            -- fases.bajfec
Dim Empresa         As String   'Empresa             |Empresa Remuneração         -- his_estructura.estrnro



Dim Sucursal        As String   'Sucursal                 -- his_estructura.estrnro
Dim Sector          As String   'Sector                   -- his_estructura.estrnro


Dim LPago           As String   'Lugar de Pago            -- his_estructura.estrnro



Dim SucBanco        As String   'Sucursal del Banco       -- ctabancaria.ctabsuc


Dim NroCuentaAcreditacionE As String
Dim Actividad       As String   'Actividad                -- his_estructura.estrnro
Dim CondSIJP        As String   'Condicion SIJP           -- his_estructura.estrnro
Dim SitRev          As String   'Sit. de Revista SIJP     -- his_estructura.estrnro
Dim ModCont         As String   'Mod. de Contrat. SIJP    -- his_estructura.estrnro
Dim ART             As String   'ART                      -- his_estructura.estrnro




Dim ModOrg          As String   'Empresa                  -- his_estructura.estrnro
Dim OSL             As String   'Empresa                  -- his_estructura.estrnro
Dim OSE             As String   'Empresa                  -- his_estructura.estrnro
Dim PlanOdon        As String   'Empresa                  -- his_estructura.estrnro
Dim Locacion        As String   'Empresa                  -- his_estructura.estrnro
Dim Area            As String   'Empresa                  -- his_estructura.estrnro
Dim SubDepto        As String   'Empresa                  -- his_estructura.estrnro
Dim NroCBU          As String   'Empresa                  -- his_estructura.estrnro
Dim Empremu         As String   'Remuneración del empleado
Dim GrupoSeguridad  As String   'Grupo de Seguridad
Dim Nro_GrupoSeguridad  As Long 'Grupo de Seguridad       -- his_estructura.estrnro





'----------------------ESTAS EN TEORIA HAY QUE ELIMINARLAS LUEGO
Dim Fing            As String   'Fec.Ingreso al Pais      -- terecro.terfecing
Dim Estudio         As String   'Estudia                 -- empleado.empestudia
Dim Cuil            As String   'CUIL                     -- ter_doc.nrodoc (10)
Dim Depto           As String   'Depto                              -- detdom.depto
Dim Torre           As String   'Torre                    -- detdom.torre
Dim Entre           As String   'Entre Calles             -- detdom.entrecalles
Dim Zona            As String   'Zona                     -- detdom.zonanro
Dim Telefono        As String   'Telefono                 -- telefono.telnro
Dim TelLaboral      As String   'Telefono                 -- telefono.telnro
Dim TelCelular      As String   'Telefono                 -- telefono.telnro
Dim Puesto          As String   'Puesto                   -- his_estructura.estrnro
Dim CCosto          As String   'C.Costo                  -- his_estructura.estrnro
Dim Gerencia        As String   'Gerencia                 -- his_estructura.estrnro
Dim Departamento    As String   'Departamento             -- his_estructura.estrnro
Dim Direccion       As String   'Direccion                -- his_estructura.estrnro
Dim Sindicato       As String   'Sindicato                -- his_estructura.estrnro
Dim OSocialLey      As String   'Obra Social              -- his_estructura.estrnro
Dim PlanOSLey       As String   'Plan OS                  -- his_estructura.estrnro
Dim OSocialElegida  As String   'Obra Social              -- his_estructura.estrnro
Dim PlanOSElegida   As String   'Plan OS                  -- his_estructura.estrnro
Dim Manzana         As String   'Manzana             |                            -- detdom.manzana
Dim Barrio          As String   'Barrio              |                            -- detdom.barrio

'----------------------

Dim FFinContrato    As String
Dim Fecha_FinContrato As String
Dim Reporta_a       As String
Dim Nro_Reporta_a   As Long

Dim Ternro As Long

Dim NroTercero          As Long

Dim Nro_Legajo          As Long
Dim Nro_TDocumento      As Long
Dim nro_nivest          As Long
Dim nro_estudio         As Long

Dim Nro_Nrodom          As String

Dim Nro_Barrio          As Long
Dim Nro_Localidad       As Long
Dim Nro_Partido         As Long
Dim Nro_Zona            As Long
Dim Nro_Provincia       As Long
Dim Nro_Pais            As Long
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
Dim Ter_Empresa         As Long
Dim ter_cajajub         As Long
Dim ter_sindicato       As Long
Dim ter_osocial_ley     As Long
Dim ter_osocial_elegida As Long
Dim ter_bancopago       As Long
Dim ter_art             As Long
Dim ter_sexo            As Long
Dim ter_estudio         As Long
Dim ter_estado          As Long
Dim os_vacio As Long
Dim os_bool As Boolean

Dim fpgo_bancaria       As Long

Dim rs As New ADODB.Recordset
Dim rs_Sql As New ADODB.Recordset
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
Dim OSoElegidaDesc      As Boolean   'Obra Social              -- his_estructura
Dim PoSElegidaDesc      As Boolean   'Plan OS                  -- his_estructura
Dim OSoLeyDesc          As Boolean   'Obra Social              -- his_estructura
Dim PoSLeyDesc          As Boolean   'Plan OS                  -- his_estructura
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
Dim SubDepDesc           As Boolean     'Empresa                  --

Dim IngresoDom          As Boolean

Dim rs_tdoc As New ADODB.Recordset
Dim rs_emp  As New ADODB.Recordset
Dim rs_tpl  As New ADODB.Recordset
Dim rs_leg  As New ADODB.Recordset

Dim Nroadtemplate As Long
Dim Nro_Institucion As Long

Dim Sigue As Boolean
Dim ExisteLeg As Boolean
Dim CalculaLegajo As Boolean
Dim Valida_CUIL As Boolean

Dim F_NacAux As Date
Dim F_AltaAux As Date
Dim Edad As Integer
Dim MaxEmpl As Long
Dim CantEmpl As Long

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
    
    
    'FGZ - 21/02/2009 - reviso si debo validar el CUIL o no --------------------
    Valida_CUIL = True
'    StrSql = " SELECT tidvalida FROM tipodocu WHERE tidnro = 10"
'    OpenRecordset StrSql, rs_Sql
'    If Not rs_Sql.EOF Then
'        If Not EsNulo(rs_Sql!tidvalida) Then
'            Valida_CUIL = True
'        Else
'            Valida_CUIL = False
'        End If
'    Else
'        Valida_CUIL = False
'    End If
    'FGZ - 21/02/2009 - reviso si debo validar el CUIL o no --------------------
    
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    'GUARDO LOS MENSAJES QUE SE REPITEN + DE 1 VEZ.
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    Dim Mje_EsNuloObli As String
    Dim Mje_Campo As String
    
    Mje_EsNuloObli = EscribeLogMI("con valor Nulo y es obligatorio")
    Mje_Campo = EscribeLogMI("campo")
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    
    
    Sigue = True 'Indica que si en el archivo viene mas de una vez un empleado, le crea las fases
    ExisteLeg = False
    LineaCarga = LineaCarga + 1
    
    Flog.writeline
    FlogE.writeline
    FlogP.writeline
    
    '__________________________
    'LEGAJO - N° Mec
    '--------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Legajo")
    pos1 = 1
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Legajo = Aux
    If Legajo = "N/A" Or EsNulo(Legajo) Then
        CalculaLegajo = True
    Else
        StrSql = "SELECT * FROM empleado WHERE empleado.empleg = " & Legajo
        OpenRecordset StrSql, rs_emp
        If (Not rs_emp.EOF) Then
            If (Not Sigue) Then
                Texto = ": " & " - " & EscribeLogMI("El Empleado ya Existe.")
                NroColumna = 1
                Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
            Else
                NroTercero = rs_emp!Ternro
                ExisteLeg = True
            End If
        End If
    End If
    '__________________________
    '--------------------------
    
    '__________________________
    'APELLIDO - SOBRENOME
    '--------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Apellido")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    If (Aux <> "N/A") Then
    Aux = EliminarCHInvalidosII(Aux, 0, "")
    End If
    Apellido = Left(Aux, 25)
    '__________________________
    '--------------------------
    
    '__________________________
    'NOMBRE - NOME
    '--------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Nombre")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    If (Aux <> "N/A") Then
    Aux = EliminarCHInvalidosII(Aux, 0, "")
    End If
    nombre = Left(Aux, 25)
    If (Apellido = "" Or Apellido = "N/A") And (nombre = "" Or nombre = "N/A") Then
        Texto = ": " & " - " & EscribeLogMI("Debe Ingresar Nombre y Apellido.")
        NroColumna = 2
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        'Ok = False
        RegError = RegError + 1
        Exit Sub
    End If
    '__________________________
    '--------------------------
    
    '________________________________________
    'FECHA DE NACIMIENTO - Data de Nascimento
    '----------------------------------------
    NroColumna = NroColumna + 1
    'Obligatorio = False
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Fecha de Nacimiento")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    If Aux <> "N/A" Then
    Aux = EliminarCHInvalidosII(Aux, 2, "")
    End If
    Fnac = Aux
    If Fnac = "N/A" Then
       F_Nacimiento = "Null"
    Else
       F_Nacimiento = ConvFecha(Fnac)
       F_NacAux = CDate(Fnac)
    End If
    '__________________________
    '--------------------------
    
    '_______________________________________
    'PAIS DE NACIMIENTO - PAIS DE NASCIMENTO
    '---------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Pais de Nacimiento")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    PaisNac = Aux
    If PaisNac <> "N/A" Then
        StrSql = " SELECT paisnro FROM pais WHERE UPPER(paisdesc) = '" & UCase(PaisNac) & "'"
        OpenRecordset StrSql, rs_Sql
        If Not rs_Sql.EOF Then
            nro_paisnac = rs_Sql!paisnro
        Else
            StrSql = " INSERT INTO pais(paisdesc,paisdef) VALUES ('" & UCase(PaisNac) & "',0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_paisnac = getLastIdentity(objConn, "pais")
        End If
    Else
        nro_paisnac = 0
    End If
    '__________________________
    '--------------------------
    '_____________________________
    'NACIONALIDAD - NACIONALIDADE
    '-----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Nacionalidad")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Nacionalidad = Aux
    If Nacionalidad <> "N/A" Then
        StrSql = " SELECT nacionalnro FROM nacionalidad WHERE UPPER(nacionaldes) = '" & UCase(Nacionalidad) & "'"
        OpenRecordset StrSql, rs_Sql
        If Not rs_Sql.EOF Then
            nro_nacionalidad = rs_Sql!nacionalnro
        Else
            StrSql = " INSERT INTO nacionalidad(nacionaldes) VALUES ('" & UCase(Nacionalidad) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_nacionalidad = getLastIdentity(objConn, "nacionalidad")
        End If
    Else
        nro_nacionalidad = 0
    End If
    If nro_nacionalidad = 0 Then
        Texto = ": " & " - " & EscribeLogMI("La Nacionalidad no Existe.")
        NroColumna = 6
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        'Ok = False
        RegError = RegError + 1
        Exit Sub
    End If
    
    '__________________________
    '--------------------------
    
    'Fecha de Ingreso al Pais
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Fecha de Ingreso al Pais"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    If aux <> "N/A" Then
'    aux = EliminarCHInvalidosII(aux, 2, "")
'    End If
'    Fing = aux
'    If (Fing = "N/A") Or EsNulo(Fing) Then
        F_Ingreso = "Null"
'    Else
'        F_Ingreso = ConvFecha(Fing)
'    End If
    
    
    '____________________________
    'ESTADO CIVIL - ESTADO CIVIL
    '----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Estado Civil")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    EstCivil = Left(Aux, 30)
    If EstCivil <> "N/A" And Not EsNulo(EstCivil) Then
        StrSql = " SELECT estcivnro FROM estcivil WHERE UPPER(estcivdesabr) = '" & UCase(EstCivil) & "'"
        OpenRecordset StrSql, rs_Sql
        If Not rs_Sql.EOF Then
            nro_estcivil = rs_Sql!estcivnro
        Else
            StrSql = " INSERT INTO estcivil(estcivdesabr) VALUES ('" & UCase(EstCivil) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_estcivil = getLastIdentity(objConn, "estcivil")
        End If
    Else
        nro_estcivil = 0
    End If
    If nro_estcivil = 0 Then
        Texto = ": " & " - " & EscribeLogMI("El Estado Civil no Existe.")
        NroColumna = 8
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        'Ok = False
        RegError = RegError + 1
        Exit Sub
    End If
    '__________________________
    '--------------------------
    
    '__________________________
    'SEXO - SEXO
    '--------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Sexo")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Sexo = UCase(Aux)
    If (Sexo = "M") Or (Sexo = "-1") Or (Sexo = "MASCULINO") Then
        ter_sexo = -1
    Else
        ter_sexo = 0
    End If
    
    '__________________________
    '--------------------------
    
    '________________________________
    'FECHA DE ALTA - DATA DE ADMISSÃO
    '--------------------------------
    NroColumna = NroColumna + 1
    'Obligatorio = False
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Fecha de alta")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    If Aux <> "N/A" Then
        Aux = EliminarCHInvalidosII(Aux, 2, "")
    End If
    FAlta = Aux
    If FAlta = "N/A" Or EsNulo(FAlta) Then
        F_Alta = "Null"
        Texto = ": " & " - " & EscribeLogMI("La Fecha de Alta es Obligatoria.")
        NroColumna = 10
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        'Ok = False
        RegError = RegError + 1
        Exit Sub
    Else
        F_Alta = ConvFecha(FAlta)
        F_AltaAux = CDate(FAlta)
        
        '____________________________________________
        'VALIDA QUE EL EMPLEADO SEA MAYOR DE 14 AÑOS
        '--------------------------------------------
        Edad = 0
        If (Month(F_AltaAux) > Month(F_NacAux)) Then
           Edad = DateDiff("yyyy", F_NacAux, F_AltaAux)
        Else
           If (Month(F_AltaAux) = Month(F_NacAux)) And (Day(F_AltaAux) >= Day(F_NacAux)) Then
             Edad = DateDiff("yyyy", F_NacAux, F_AltaAux)
           Else
             Edad = DateDiff("yyyy", F_NacAux, F_AltaAux) - 1
           End If
        End If
        
        If Edad < 14 Then
            Texto = ": " & " - " & EscribeLogMI("El empleado es menor a 14 años.")
            NroColumna = 10
            Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
            'Ok = False
            RegError = RegError + 1
            Exit Sub
        End If
        
    End If
    
    '__________________________
    '--------------------------
   
   '-
    'Estudia?
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Estudia?"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Estudio = UCase(aux)
'    If Estudio <> "N/A" And Estudio <> "NO" Then
'        If Estudio = "SI" Then
'            ter_estudio = -1
'        Else
            ter_estudio = 0
'        End If
'    Else
'        ter_estudio = 0
'    End If
    
    '____________________________________
    'NINVEL DE ESTUDIO - NIVEL DE ESTUDO
    '------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Nivel de estudio")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    NivEst = Left(Aux, 40)
    If NivEst <> "N/A" Then
        StrSql = " SELECT nivnro FROM nivest WHERE UPPER(nivdesc) = '" & UCase(NivEst) & "'"
        OpenRecordset StrSql, rs_Sql
        If Not rs_Sql.EOF Then
            nro_nivest = rs_Sql!nivnro
        Else
            StrSql = " INSERT INTO nivest(nivdesc,nivsist,nivobligatorio,nivestfli) VALUES ('" & UCase(NivEst) & "',-1,0,-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_nivest = getLastIdentity(objConn, "nivest")
        End If
    Else
        nro_nivest = 0
    End If
    
    '__________________________
    '--------------------------
    
    '______________________________________
    'TIPO DE DOCUMENTO - TIPO DE DOCUMENTO
    '--------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Tipo de Doc")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Tdocu = Left(Aux, 8)
    If Tdocu <> "N/A" Then
        StrSql = " SELECT tidnro FROM tipodocu WHERE UPPER(tidsigla) = '" & UCase(Tdocu) & "'"
        OpenRecordset StrSql, rs_Sql
        If Not rs_Sql.EOF Then
            Nro_TDocumento = rs_Sql!tidnro
        Else
            '__________________________________________________
            'BUSCO LA PRIMERA INSTITUCION, SI NO EXISTE LA CREO             'VER QUE HACE ESTO!!
            '--------------------------------------------------
            StrSql = " SELECT * FROM institucion "
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                Nro_Institucion = rs!InstNro
            Else
                'creo una
                StrSql = " INSERT INTO institucion (instdes,instabre) VALUES ('NACIONAL','NAC')"
                objConn.Execute StrSql, , adExecuteNoRecords
                Nro_Institucion = getLastIdentity(objConn, "institucion")
            End If
            'creo el tipo de documento
            StrSql = " INSERT INTO tipodocu(tidnom,tidsigla,tidsist,instnro,tidunico) VALUES ('" & UCase(Tdocu) & "','" & UCase(Tdocu) & "',0," & Nro_Institucion & ",0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Nro_TDocumento = getLastIdentity(objConn, "tipodocu")
        End If
    Else
        Nro_TDocumento = 0
    End If
    If Nro_TDocumento = 0 Then
        Texto = ": " & " - " & EscribeLogMI("El Tipo de Documento no Existe.")
        NroColumna = 13
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        'Ok = False
        RegError = RegError + 1
        Exit Sub
    End If
    
    '__________________________
    '--------------------------
    
    '__________________________________
    'N° DE DOCUMENTO - N° DE DOCUMENTO
    '----------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Nro de Documento")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Ndocu = Left(Aux, 30)
    If Ndocu = "N/A" Then
        Ndocu = ""
    End If
    
    StrSql = "SELECT * FROM empleado "
    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro "
    StrSql = StrSql & " WHERE ter_doc.tidnro = " & Nro_TDocumento & " AND ter_doc.nrodoc = '" & Ndocu & "'"
    OpenRecordset StrSql, rs_tdoc
    If (Not rs_tdoc.EOF) Then
        If (Not Sigue) Then
            Texto = ": " & " - " & EscribeLogMI("Tipo y Numero de Documento Asignados a otro Empleado")
            NroColumna = 14
            Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
            'Ok = False
            RegError = RegError + 1
            Exit Sub
        Else
            'FGZ - 11/11/2010 - Si es el mismo legajo no debe poner ningun cartel
            If NroTercero <> rs_tdoc!Ternro Then
                NroTercero = rs_tdoc!Ternro
                ExisteLeg = True
                Texto = ": " & " - " & EscribeLogMI("Empleado") & ": " & Legajo & " - " & EscribeLogMI("Tipo y Numero de Documento Asignados a otro Empleado")
                NroColumna = 14
                Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
            End If
        End If
    End If
    
    'CUIL
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "CUIL"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    Cuil = Replace(Left(aux, 30), "-", "")
'    If Cuil = "N/A" Or EsNulo(Cuil) Then
'        'FGZ - 23/02/2009 - validacion de cuil configurable ---------------
'        If Valida_CUIL Then
'            Cuil = Generar_Cuil(Ndocu, CBool(ter_sexo))
'            'Cuil = CalcularCUIL(Ndocu)
'            If Cuil = 0 Then
'                texto = ": " & " - Campo " & Campoetiqueta & " no se pudo generar automaticamente " & Ndocu
'                Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'                RegError = RegError + 1
'                Exit Sub
'            End If
'        Else
'            Cuil = ""
'        End If
'        'FGZ - 23/02/2009 - validacion de cuil configurable ---------------
'    Else
'        If Valida_CUIL Then
'            'OK = Cuil_Valido(Cuil, Texto)
'            OK = Cuil_Valido605(Cuil, Ndocu, texto, Tdocu, nro_nacionalidad)
'            If Not OK Then
'                OK = True
'                'Texto = "El CUIL no es valido"
'                'Call Escribir_Log("floge", LineaCarga, nrocolumna, Texto, Tabs, strReg)
'                'Cuil = CalcularCUIL(Ndocu)
'
'                'FGZ - 26/08/2010 ---------------
'                'Cuil = Generar_Cuil(Ndocu, CBool(ter_sexo))
'                Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'                RegError = RegError + 1
'                Exit Sub
'                'FGZ - 26/08/2010 ---------------
'            End If
'        End If
'    End If
    'FGZ - 23/02/2009 - validacion de cuil configurable ---------------
'    If Cuil = "" And Not Valida_CUIL Then
'        Cuil = ""
'    Else
'        Cuil = Left(Cuil, 2) & "-" & Mid(Cuil, 3, 8) & "-" & Right(Cuil, 1)
'    End If
    'FGZ - 23/02/2009 - validacion de cuil configurable ---------------
    
    
    'Hasta Aqui los Datos Obligatorios del Empleado
    IngresoDom = True
        
        
    '_____________________________
    'CALLE - RUA
    '-----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Direccion.Calle")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Calle = Left(Aux, 250)
    If Calle = "N/A" Then
        Calle = ""
        IngresoDom = False
    End If
    '_____________________________
    '-----------------------------
    
    
    '_____________________________
    'NUMERO - NUMERO
    '-----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Direccion.Nro")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Nro = Left(Aux, 8)
    If (Nro <> "N/A") Then
        Nro_Nrodom = Nro
    Else
        Nro_Nrodom = 0
    End If
    '_____________________________
    '-----------------------------
    
    
    '____________________________
    'PISO - ANDAR
    '----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Direccion.Piso")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Piso = Left(Aux, 8)
    If Piso = "N/A" Then
        Piso = ""
    End If

    
    
    'Departamento
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Direccion.Departamento"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Depto = Left(aux, 8)
'    If Depto = "N/A" Then
        Depto = ""
'    End If

    'Torre
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Direccion.Torre"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Torre = Left(aux, 8)
'    If Torre = "N/A" Then
        Torre = ""
'    End If
    
    '________________________
    'MANZANA
    '------------------------
'   Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = EscribeLogMI("Direccion.Manzana")
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, Texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Manzana = Left(aux, 8)
'    If Manzana = "N/A" Then
        Manzana = ""
'    End If
    '____________________________
    'PROVINCIA -
    '----------------------------
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Provincia"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Provincia = Left(aux, 20)
'
'    '_____________________________
'    '-----------------------------
    
    
    '______________________________
    'CODIGO POSTAL - CODIGO POSTAL
    '------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Codigo Postal")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Cpostal = Left(Aux, 12)
    If Cpostal = "N/A" Then
        Cpostal = ""
    End If
    '_____________________________
    '-----------------------------
    
    'Entre calles
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Entre calles"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Entre = Left(aux, 80)
'    If Entre = "N/A" Then
        Entre = ""
'    End If

    '________________________
    'BARRIO -
    '------------------------
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = EscribeLogMI("Barrio")
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Barrio = Left(aux, 30)
'    If Barrio = "N/A" Then
'        Barrio = ""
'    End If
    '_____________________________
    '-----------------------------


    '____________________________
    'PROVINCIA - DISTRITO
    '----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Provincia")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Provincia = Left(Aux, 20)
    '_____________________________
    '-----------------------------
    'PARTIDO - CONSELHO
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Partido")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Partido = Left(Aux, 30)
    
    '_____________________________
    '-----------------------------
    
    '_____________________________
    'LOCALIDAD - FREGUESIA
    '-----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Localidad")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Localidad = Left(Aux, 60)
    '_____________________________
    '-----------------------------
    

'    '_____________________________
'    '-----------------------------
    
    'Partido
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Partido"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Partido = Left(aux, 30)
'    Partido = ""
    '_____________________________
    '-----------------------------
    
    'Zona
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Zona"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    '19/03/2010 - Se cambio la longitud a 60
'    'Zona = Left(aux, 20)
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Zona = Left(aux, 60)
    Zona = "N/A"

    
    '_____________________
    'PAIS - PAIS
    '---------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Pais")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Pais = Left(Aux, 20)
    If Pais <> "N/A" Then
        Call ValidarPais(Pais, Nro_Pais)
    Else
        Nro_Pais = 0
    End If
    If Provincia <> "N/A" Then
        Call ValidarProvincia(Provincia, Nro_Provincia, Nro_Pais)
    Else
        Nro_Provincia = 0
    End If
    If Localidad <> "N/A" Then
        Call ValidarLocalidad(Localidad, Nro_Localidad, Nro_Pais, Nro_Provincia)
    Else
        Nro_Localidad = 0
    End If
    If Partido <> "N/A" Then
        Call ValidarPartido(Partido, Nro_Partido)
    Else
        Nro_Partido = 0
    End If
    If Zona <> "N/A" Then
        Call ValidarZona(Zona, Nro_Zona, Nro_Provincia)
    Else
        Nro_Zona = 0
    End If
    If (IngresoDom = True) And (Nro_Localidad = 0) Then
        Texto = ": " & " - " & EscribeLogMI("Debe Ingresar la Localidad.")
        NroColumna = 25
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        'Ok = False
        RegError = RegError + 1
        Exit Sub
    End If
    If (IngresoDom = True) And (Nro_Provincia = 0) Then
        Texto = ": " & " - " & EscribeLogMI("Debe Ingresar la Provincia.")
        NroColumna = 28
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        'Ok = False
        RegError = RegError + 1
        Exit Sub
    End If
    If (IngresoDom = True) And (Nro_Pais = 0) Then
        Texto = ": " & " - " & EscribeLogMI("Debe Ingresar la Pais.")
        NroColumna = 29
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        'Ok = False
        RegError = RegError + 1
        Exit Sub
    End If
    
    
    'Tel Particular
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Tel Particular"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    If aux = "N/A" Then
        Telefono = ""
'    Else
'        aux = EliminarCHInvalidosII(aux, 3, "")
'        Telefono = Left(aux, 20)
'    End If
    
    
    'Tel Laboral
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Tel Laboral"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    If aux = "N/A" Then
        TelLaboral = ""
'    Else
'        aux = EliminarCHInvalidosII(aux, 3, "")
'        TelLaboral = Left(aux, 20)
'    End If
    
    'Tel Celular
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Tel Celular"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    If aux = "N/A" Then
        TelCelular = ""
'    Else
'        aux = EliminarCHInvalidosII(aux, 3, "")
'        TelCelular = Left(aux, 20)
'    End If

    '________________________
    'EMAIL - EMAIL
    '------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("E-mail")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    If Aux <> "N/A" Then
        Aux = EliminarCHInvalidosII(Aux, 4, "")
        Email = Left(Aux, 100)
    Else
        Email = ""
    End If
    '_________________________
    '-------------------------
    
    
    'Sucursal
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Sucursal"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Sucursal = aux
'    ' Validacion y Creacion de la Sucursal (junto con sus Complementos)
'    If Sucursal <> "N/A" Then
'        If SucDesc Then
'            Call ValidaEstructura(1, Sucursal, nro_sucursal, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(1, Sucursal, nro_sucursal, Inserto_estr)
'        End If
'        Call CreaTercero(10, Sucursal, ter_sucursal)
'
'        If Inserto_estr Then
'            Call CreaComplemento(1, ter_sucursal, nro_sucursal, Sucursal)
'            Inserto_estr = False
'        End If
'    Else
        nro_sucursal = 0
'    End If
    
    'Sector
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Sector"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Sector = aux
'    ' Validacion y Creacion del Sector
'    If Sector <> "N/A" Then
'        If SecDesc Then
'            Call ValidaEstructura(2, Sector, nro_sector, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(2, Sector, nro_sector, Inserto_estr)
'        End If
'    Else
        nro_sector = 0
'    End If
    
    '___________________________
    'CONVENIO - ACORDO COLETIVO
    '---------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Convenio")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Convenio = Aux
    ' Validacion, Creacion del Convenio
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
    '___________________________
    '---------------------------
    
    
    '___________________________
    'CATEGORIA - CATEGORIA
    '---------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Categoria")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    categoria = Aux
    '_____________________________________
    'VALIDACION, CREACION DE LA CATEGORIA
    If (categoria <> "N/A" And nro_convenio <> 0) Then
        If CatDesc Then
            Call ValidaCategoria(3, categoria, nro_convenio, nro_categoria, Inserto_estr)
        Else
            Call ValidaCategoriaCodExt(3, categoria, nro_convenio, nro_categoria, Inserto_estr)
        End If
    Else
        nro_categoria = 0
    End If
    '___________________________
    '---------------------------
    
    'Puesto
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Puesto"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Puesto = aux
'    'Validacion y Creacion del Puesto (junto con sus Complementos)
'    If Puesto <> "N/A" Then
'        If PueDesc Then
'            Call ValidaEstructura(4, Puesto, nro_puesto, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(4, Puesto, nro_puesto, Inserto_estr)
'        End If
'
'        If Inserto_estr Then
'            Call CreaComplemento(4, 0, nro_puesto, Puesto)
'            Inserto_estr = False
'        End If
'    Else
        nro_puesto = 0
'    End If

    'Centro de Costo
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Centro de Costo"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    CCosto = aux
'    ' Validacion y Creacion del Centro de Costo
'    If CCosto <> "N/A" Then
'        If CCoDesc Then
'            Call ValidaEstructura(5, CCosto, nro_ccosto, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(5, CCosto, nro_ccosto, Inserto_estr)
'        End If
'    Else
        nro_ccosto = 0
'    End If

    'Gerencia
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Gerencia"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Gerencia = aux
'    ' Validacion y Creacion de la Gerencia
'    If Gerencia <> "N/A" Then
'        If GerDesc Then
'            Call ValidaEstructura(6, Gerencia, nro_gerencia, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(6, Gerencia, nro_gerencia, Inserto_estr)
'        End If
'    Else
        nro_gerencia = 0
'    End If

    
    'Departamento
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Departamento"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Departamento = aux
'    ' Validacion y Creacion del Departamento
'    If Departamento <> "N/A" Then
'        If DepDesc Then
'            Call ValidaEstructura(9, Departamento, nro_departamento, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(9, Departamento, nro_departamento, Inserto_estr)
'        End If
'    Else
        nro_departamento = 0
'    End If


    'Direccion
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Direccion"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Direccion = aux
'    ' Validacion y Creacion de direccion
'    If Direccion <> "N/A" Then
'        If DirDesc Then
'            Call ValidaEstructura(35, Direccion, nro_direccion, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(35, Direccion, nro_direccion, Inserto_estr)
'        End If
'    Else
        nro_direccion = 0
'    End If
    
    '________________________________________________
    'CAJA DE JUBILACION - CONTRIBUÇÃO PREVIDENCIÁRIA
    '------------------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Caja de Jubilacion")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    CajaJub = Aux
    '____________________________________________________
    ' VALIDACION Y CREACION DE CONTRIBUÇÃO PREVIDENCIÁRIA
    If CajaJub <> "N/A" Then
        If CaJDesc Then
            Call ValidaEstructura(15, CajaJub, nro_cajajub, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(15, CajaJub, nro_cajajub, Inserto_estr)
        End If
        Call CreaTercero(6, CajaJub, ter_cajajub)
        
        If Inserto_estr Then
            Call CreaComplemento(15, ter_cajajub, nro_cajajub, CajaJub)
        End If
    Else
        nro_cajajub = 0
    End If

    'Sindicato
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Sindicato"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Sindicato = aux
'    ' Validacion y Creacion del Sindicato
'    If Sindicato <> "N/A" Then
'        If SinDesc Then
'            Call ValidaEstructura(16, Sindicato, nro_sindicato, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(16, Sindicato, nro_sindicato, Inserto_estr)
'        End If
'        Call CreaTercero(5, Sindicato, ter_sindicato)
'
'        If Inserto_estr Then
'            Call CreaComplemento(16, ter_sindicato, nro_sindicato, Sindicato)
'        End If
'    Else
        nro_sindicato = 0
'    End If
    
    os_vacio = 0
    os_bool = False
    
    'Obra social por Ley
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Obra social por Ley"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    OSocialLey = aux
'    ' Validacion y Creacion de la Obra Social por Ley
'    If OSocialLey <> "N/A" Then
'        If OSoLeyDesc Then
'            Call ValidaEstructura(24, OSocialLey, nro_osocial_ley, Inserto_estr)
'            Call ValidaEstructura2(17, OSocialLey, os_vacio) 'Agregado ver 3.96
'        Else
'            Call ValidaEstructuraCodExt(24, OSocialLey, nro_osocial_ley, Inserto_estr)
'        End If
'        Call CreaTercero(4, OSocialLey, ter_osocial_ley)
'
'        If Inserto_estr Then
'            Call CreaComplemento(24, ter_osocial_ley, nro_osocial_ley, OSocialLey)
'            Call CreaComplemento(24, ter_osocial_ley, os_vacio, OSocialLey) 'Agregado ver 3.96
'        Else
'            StrSql = " SELECT origen FROM replica_estr WHERE estrnro = " & nro_osocial_ley
'            OpenRecordset StrSql, rs_repl
'
'            If Not rs_repl.EOF Then
'                ter_osocial_ley = rs_repl!Origen
'                rs_repl.Close
'            End If
'
'        End If
'    Else
        nro_osocial_ley = 0
'    End If

    os_vacio = 0
    os_bool = False
    'Plan de OS por Ley
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Plan de Obra social por Ley"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    PlanOSLey = aux
'    ' Validacion y Creacion del Plan de Obra Social por Ley
'    If (PlanOSLey <> "N/A" And nro_osocial_ley <> 0) Then
'        If PoSLeyDesc Then
'            Call ValidaEstructura(25, PlanOSLey, nro_planos_ley, Inserto_estr)
'            Call ValidaEstructura2(23, PlanOSLey, os_vacio) 'Agregado ver 3.96
'        Else
'            Call ValidaEstructuraCodExt(25, PlanOSLey, nro_planos_ley, Inserto_estr)
'         End If
'
'        If Inserto_estr Then
'            'Manterola Maria Magdalena (29/06/2011)
'            Call CreaComplemento(23, ter_osocial_ley, nro_planos_ley, PlanOSLey)
'            Call CreaComplemento(23, ter_osocial_ley, os_vacio, PlanOSLey) 'Agregado ver 3.96
'            Inserto_estr = False
'        End If
'    Else
        nro_planos_ley = 0
'    End If
    os_vacio = 0
    os_bool = False
    
    'OS Elegida
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Obra Social elegida"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    OSocialElegida = aux
'    ' Validacion y Creacion de la Obra Social Elegida
'    If OSocialElegida <> "N/A" Then
'        If OSoElegidaDesc Then
'            Call ValidaEstructura(17, OSocialElegida, nro_osocial_elegida, Inserto_estr)
'            Call ValidaEstructura2(25, OSocialElegida, os_vacio) 'Agregado ver 3.96
'        Else
'            Call ValidaEstructuraCodExt(17, OSocialElegida, nro_osocial_elegida, Inserto_estr)
'        End If
'        Call CreaTercero(4, OSocialElegida, ter_osocial_elegida)
'
'        If Inserto_estr Then
'            Call CreaComplemento(17, ter_osocial_elegida, nro_osocial_elegida, OSocialElegida)
'            Call CreaComplemento(17, ter_osocial_elegida, os_vacio, OSocialElegida) 'Agregado ver 3.96
'        Else
'            StrSql = " SELECT origen FROM replica_estr WHERE estrnro = " & nro_osocial_elegida
'            OpenRecordset StrSql, rs_Sql
'            ter_osocial_elegida = rs_Sql!Origen
'            rs_Sql.Close
'        End If
'    Else
        nro_osocial_elegida = 0
'    End If
    os_vacio = 0
    os_bool = False
    
    'Plan de OS Elegida
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Plan Obra Social elegida"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    PlanOSElegida = aux
'    ' Validacion y Creacion del Plan de Obra Social Elegida
'    If (PlanOSElegida <> "N/A" And nro_osocial_elegida <> 0) Then
'        If PoSElegidaDesc Then
'            Call ValidaEstructura(23, PlanOSElegida, nro_planos_elegida, Inserto_estr)
'            Call ValidaEstructura2(25, PlanOSElegida, os_vacio) 'Agregado ver 3.96
'        Else
'            Call ValidaEstructuraCodExt(23, PlanOSElegida, nro_planos_elegida, Inserto_estr)
'        End If
'
'        If Inserto_estr Then
'            Call CreaComplemento(23, ter_osocial_elegida, nro_planos_elegida, PlanOSElegida)
'            Call CreaComplemento(23, ter_osocial_elegida, os_vacio, PlanOSElegida) 'Agregado ver 3.96
'            Inserto_estr = False
'        End If
'    Else
        nro_planos_elegida = 0
'    End If
    
    '__________________________
    'CONTRATO - TIPO CONTRATO
    '--------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Contrato")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Contrato = Aux
    '____________________________________
    ' Validacion y Creacion del Contrato
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
    '__________________________
    '--------------------------
    
    
    '______________________________________________________
    'FECHA DE FIN DE CONTRATO - DATA VENCIMENTO DE CONTRATO
    '------------------------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Fecha de fin de contrato")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    If Aux <> "N/A" Then
        Aux = EliminarCHInvalidosII(Aux, 2, "")
    End If
    FFinContrato = Aux
    If FFinContrato = "N/A" Or EsNulo(FFinContrato) Then
        Fecha_FinContrato = "Null"
    Else
        Fecha_FinContrato = ConvFecha(FFinContrato)
    End If
    '__________________________
    '--------------------------
    
    
    'Lugar de pago
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Lugar de pago"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    LPago = aux
'    ' Validacion y Creacion del Lugar de Pago
'    If LPago <> "N/A" Then
'        If LPaDesc Then
'            Call ValidaEstructura(20, LPago, nro_lpago, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(20, LPago, nro_lpago, Inserto_estr)
'        End If
'    Else
        nro_lpago = 0
'    End If
    '_______________________________
    'REGIMEN HORARIO - HORÁRIO
    '-------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Regimen Horario")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    RegHorario = Aux
    If RegHorario <> "N/A" Then
        If RegDesc Then
            Call ValidaEstructura(21, RegHorario, nro_reghorario, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(21, RegHorario, nro_reghorario, Inserto_estr)
        End If
    Else
        nro_reghorario = 0
    End If
    '------------------------------
    '------------------------------

    '________________________________________
    'FORMA DE LIQUIDACION - FORMA DE CÁLCULO
    '----------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Forma de Liquidacion")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    FormaLiq = Aux
    If FormaLiq <> "N/A" Then
        If FLiDesc Then
            Call ValidaEstructura(22, FormaLiq, nro_formaliq, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(22, FormaLiq, nro_formaliq, Inserto_estr)
        End If
        ' Agregado por MB 10/08/2006
        If Inserto_estr Then
            Call CreaComplemento(22, 0, nro_formaliq, FormaLiq)
            Inserto_estr = False
        End If
    Else
        nro_formaliq = 0
    End If
    '------------------------------
    '------------------------------

    '___________________________________
    'FORMA DE PAGO - FORMA DE PAGAMENTO
    '-----------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Forma de Pago")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    FormaPago = Aux
    If FormaPago <> "N/A" Then
        StrSql = " SELECT fpagnro FROM formapago WHERE fpagdescabr = '" & FormaPago & "'"
        OpenRecordset StrSql, rs_Sql
        If Not rs_Sql.EOF Then
            nro_formapago = rs_Sql!fpagnro
        Else
            StrSql = " INSERT INTO formapago(fpagdescabr,fpagbanc,acunro,monnro) VALUES ('" & FormaPago & "',0,6,1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nro_formapago = getLastIdentity(objConn, "formapago")
        End If
    Else
        nro_formapago = 0
    End If
    '------------------------------
    '------------------------------
    
    '__________________________________
    'BANCO DE PAGO - BANCO DE PAGAMENTO
    '----------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Banco de Pago")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    BancoPago = Aux
    If BancoPago <> "N/A" Then
        If BcoDesc Then
            Call ValidaEstructura(41, BancoPago, nro_bancopago, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(41, BancoPago, nro_bancopago, Inserto_estr)
        End If
        Call CreaTercero(13, BancoPago, ter_bancopago)
        
        If Inserto_estr Then
            Call CreaComplemento(41, ter_bancopago, nro_bancopago, BancoPago)
        End If
        fpgo_bancaria = -1
    Else
        nro_bancopago = 0
        fpgo_bancaria = 0
    End If
    '------------------------------
    '------------------------------
    
    '______________________________
    'N° DE CUENTA - N° DE CONTA
    '------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Numero de cuenta")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    NroCuenta = Aux
    If NroCuenta = "N/A" Then
        NroCuenta = ""
    Else
        NroCuenta = Left(NroCuenta, 30)
    End If
    '--------------------------
    '--------------------------
    
    
    'CBU
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "CBU"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    NroCBU = aux
'    If NroCBU = "N/A" Then
        NroCBU = ""
'    Else
'        NroCBU = Left(NroCBU, 30)
'    End If
    
    'Sucursal del banco
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Sucursal del banco"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    SucBanco = aux
'    If SucBanco = "N/A" Then
        SucBanco = ""
'    Else
'        SucBanco = Left(SucBanco, 10)
'    End If


    'Nro de cuenta de acreditacion empresa
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Nro de cuenta de acreditacion empresa"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    NroCuentaAcreditacionE = aux
'    If NroCuentaAcreditacionE = "N/A" Or EsNulo(NroCuentaAcreditacionE) Then
'        NroCuentaAcreditacionE = ""
'    Else
'        StrSql = "SELECT * FROM ctabancaria WHERE cbnro ='" & NroCuentaAcreditacionE & "'"
'        If rs.State = adStateOpen Then rs.Close
'        OpenRecordset StrSql, rs
'        If rs.EOF Then
'            texto = ": " & " - Nro de cuenta de acreditacion empresa no existe."
'            Nrocolumna = 59
'            Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
            NroCuentaAcreditacionE = ""
'        End If
'    End If
    
    'Actividad SIJP
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Actividad SIJP"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    Actividad = aux
'    If Actividad <> "N/A" Then
'        If ActDesc Then
'            Call ValidaEstructura(29, Actividad, nro_actividad, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(29, Actividad, nro_actividad, Inserto_estr)
'        End If
'    Else
        nro_actividad = 0
'    End If

    'Condicion SIJP
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Condicion SIJP"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    CondSIJP = aux
'    If CondSIJP <> "N/A" And Not EsNulo(CondSIJP) Then
'        If CSJDesc Then
'            Call ValidaEstructura(31, CondSIJP, nro_condsijp, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(31, CondSIJP, nro_condsijp, Inserto_estr)
'        End If
'    Else
        nro_condsijp = 0
'    End If

    'Situacion de Revista SIJP
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Situacion de Revista SIJP"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    SitRev = aux
'    If SitRev <> "N/A" And Not EsNulo(SitRev) Then
'        If SReDesc Then
'            Call ValidaEstructura(30, SitRev, nro_sitrev, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(30, SitRev, nro_sitrev, Inserto_estr)
'        End If
'    Else
        nro_sitrev = 0
'    End If
    
    
'    'ART
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "ART"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'        texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    aux = EliminarCHInvalidosII(aux, 1, "")
'    ART = aux
'    If ART <> "N/A" And Not EsNulo(ART) Then
'        If ARTDesc Then
'            Call ValidaEstructura(40, ART, nro_art, Inserto_estr)
'        Else
'            Call ValidaEstructuraCodExt(40, ART, nro_art, Inserto_estr)
'        End If
'        Call CreaTercero(8, ART, ter_art)
'
'        If Inserto_estr Then
'            Call CreaComplemento(40, ter_art, nro_art, ART)
'        End If
'    Else
        nro_art = 0
'    End If
    
    '____________________________________________
    'ESTADO DEL EMPLEADO - ESTADO DO FUNCIONÁRIO
    '--------------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Estado del empleado")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    Estado = Aux
    If UCase(Estado) = "ATIVO" Then
        ter_estado = -1
    Else
        ter_estado = 0
    End If
    '-------------------------
    '-------------------------
    
    
    '____________________________________
    'CAUSA DE BAJA  - MOTIVO DA DEMISSÂO
    '------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Causa de baja")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Aux = EliminarCHInvalidosII(Aux, 1, "")
    CausaBaja = Aux
    If Not EsNulo(CausaBaja) And CausaBaja <> "N/A" Then
        StrSql = " SELECT caunro FROM causa WHERE caudes = '" & CausaBaja & "'"
        OpenRecordset StrSql, rs_Sql
        If Not rs_Sql.EOF Then
            nro_causabaja = rs_Sql!caunro
        Else
            StrSql = " INSERT INTO causa(caudes,causist,caudesvin,empnro) VALUES ('" & CausaBaja & "',0,-1,1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nro_causabaja = getLastIdentity(objConn, "causa")
        End If
    Else
        nro_causabaja = 0
    End If
    '-----------------------------------
    '-----------------------------------
    
    
    '________________________________
    'FECHA DE BAJA - DATA DE DEMISSÁO
    '--------------------------------
    NroColumna = NroColumna + 1
        ' Si hay una causa de baja, se tiene que cargar ssi la Fecha de Baja (tablero adp)
    If nro_causabaja <> 0 Then
        Obligatorio = True
    Else
        Obligatorio = False
    End If
    Campoetiqueta = EscribeLogMI("Fecha de baja")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    If Aux <> "N/A" Then
    Aux = EliminarCHInvalidosII(Aux, 2, "")
    End If
    FBaja = Aux
    If EsNulo(FBaja) Or FBaja = "N/A" Then
        F_Baja = "Null"
    Else
        F_Baja = ConvFecha(FBaja)
    End If
    '-------------------------
    '-------------------------
        
    '__________________________
    'EMPRESA - EMPRESA
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Empresa")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    If pos2 > 0 Then
        Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
        If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
            Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
            Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
            RegError = RegError + 1
            Exit Sub
        End If
        Aux = EliminarCHInvalidosII(Aux, 1, "")
        Empresa = Aux
    Else
        pos2 = Len(strReg)
        Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
        If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
            Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
            Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
            RegError = RegError + 1
            Exit Sub
        End If
        Aux = EliminarCHInvalidosII(Aux, 1, "")
        Empresa = Aux
    End If
    If Empresa <> "N/A" Or EsNulo(Empresa) Then
        If EmpDesc Then
            Call ValidaEstructura(10, Empresa, nro_empresa, Inserto_estr)
        Else
            Call ValidaEstructuraCodExt(10, Empresa, nro_empresa, Inserto_estr)
        End If
        Call CreaTercero(10, Empresa, Ter_Empresa)
        
        If Inserto_estr Then
            Call CreaComplemento(10, Ter_Empresa, nro_empresa, Empresa)
        End If
    Else
        nro_empresa = 0
    End If
    '-----------------------------
    '-----------------------------
    
    '_______________________________________________________
    'REMUNERACIÓN DEL EMPLEADO - REMUNERAÇÃO DO FUNCIONÁRIO
    '-------------------------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Remuneración del Empleado")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, Len(strReg)))
    'aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Empremu = Aux
    If Empremu = "N/A" Or EsNulo(Empremu) Then
        Empremu = "Null"
    End If
    '----------------------------------------
    '----------------------------------------
   
    'Modelo de Organizacion
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Modelo de Organizacion"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    If pos2 > 0 Then
'        aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'        If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'            texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'            Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'            RegError = RegError + 1
'            Exit Sub
'        End If
'        aux = EliminarCHInvalidosII(aux, 1, "")
'        ModOrg = aux
'        'desde aca
'        If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'            texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'            Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'            RegError = RegError + 1
'            Exit Sub
'        End If
'        ModOrg = aux
'        If ModOrg = "N/A" Or EsNulo(ModOrg) Then
'            'agrego codigo sebastian stremel 15/09/2011
'            StrSql = "SELECT * FROM adptemplate WHERE tplatedefault = -1"
'            OpenRecordset StrSql, rs_tpl
'            If Not rs_tpl.EOF Then
'                nro_ModOrg = rs_tpl!tplatenro
'            Else
'                StrSql = "SELECT top 1 * FROM adptemplate"
'                OpenRecordset StrSql, rs_tpl
'                If rs_tpl.EOF Then
'                    texto = ": no hay modelos de organizacion"
'                    Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'                    RegError = RegError + 1
'                    Exit Sub
'                Else
'                    nro_ModOrg = rs_tpl!tplatenro
'                End If
'
'            End If
'
'
'        Else
'            StrSql = "SELECT * FROM adptemplate WHERE tplatedesabr = '" & ModOrg & "'"
'            OpenRecordset StrSql, rs_tpl
'            If rs_tpl.EOF Then
'                StrSql = "INSERT INTO adptemplate (tplatedesabr,tplatedefault) VALUES ('" & ModOrg & "',-1)"
'                objConn.Execute StrSql, , adExecuteNoRecords
'                nro_ModOrg = getLastIdentity(objConn, "adptemplate")
'            Else
'                nro_ModOrg = rs_tpl!tplatenro
'            End If
'        End If
'    Else
'            StrSql = "SELECT * FROM adptemplate WHERE tplatedefault = -1"
'            OpenRecordset StrSql, rs_tpl
'            If Not rs_tpl.EOF Then
'                nro_ModOrg = rs_tpl!tplatenro
'            Else
'                StrSql = "SELECT top 1 * FROM adptemplate"
'                OpenRecordset StrSql, rs_tpl
'                If rs_tpl.EOF Then
'                    texto = ": no hay modelos de organizacion"
'                    Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'                    RegError = RegError + 1
'                    Exit Sub
'                Else
'                    nro_ModOrg = rs_tpl!tplatenro
'                End If
'
'            End If
'
'    End If
        
        
        'hasta aca
'        If ModOrg = "N/A" Or EsNulo(ModOrg) Then
'            nro_ModOrg = 0
'        Else
'            StrSql = "SELECT * FROM adptemplate WHERE tplatedesabr = '" & ModOrg & "'"
'            OpenRecordset StrSql, rs_tpl
'            If rs_tpl.EOF Then
'               StrSql = "INSERT INTO adptemplate (tplatedesabr,tplatedefault) VALUES ('" & ModOrg & "',-1)"
'                objConn.Execute StrSql, , adExecuteNoRecords
'                nro_ModOrg = getLastIdentity(objConn, "adptemplate")
'            Else
'                nro_ModOrg = rs_tpl!tplatenro
'            End If
'        End If
'    Else
'        nro_ModOrg = 0
'    End If

    'Reporta_a  campo empleado.empreporta va el tercero
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Reporta_a"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    If pos2 > 0 Then
'        aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'        If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'            texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'            Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'            RegError = RegError + 1
'            Exit Sub
'        End If
'        Reporta_a = aux
'        If Not EsNulo(Reporta_a) And Reporta_a <> "N/A" Then
'            If IsNumeric(Reporta_a) Then
'                StrSql = "SELECT * FROM empleado WHERE empleado.empleg = " & Reporta_a
'                If rs_emp.State = adStateOpen Then rs_emp.Close
'                OpenRecordset StrSql, rs_emp
'                If (Not rs_emp.EOF) Then
'                    Nro_Reporta_a = rs_emp!Ternro
'                Else
'                    Nro_Reporta_a = 0
'                    texto = ": " & "El Empleado " & Reporta_a & " no existe."
'                    Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'                End If
'            Else
'                Nro_Reporta_a = 0
'                texto = "El valor no es numérico."
'                Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'            End If
'        End If
'    Else
        Nro_Reporta_a = 0
'        'no es obligatorio
'    End If
    If rs_emp.State = adStateOpen Then rs_emp.Close
    

    'Grupo de Seguridad
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Grupo de Seguridad"
'    pos1 = pos2 + 2
'    pos2 = Len(strReg) 'InStr(pos1, strReg, Separador) - 1
'    If pos2 > 0 Then
'        aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'        If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'            texto = ": " & " - Campo " & Campoetiqueta & " con valor Nulo y es obligatorio"
'            Call Escribir_Log("floge", LineaCarga, Nrocolumna, texto, Tabs, strReg)
'            RegError = RegError + 1
'            Exit Sub
'        End If
'        aux = EliminarCHInvalidosII(aux, 1, "")
'        GrupoSeguridad = aux
'        If GrupoSeguridad <> "N/A" And Not EsNulo(GrupoSeguridad) Then
'            Call ValidaEstructura(7, GrupoSeguridad, Nro_GrupoSeguridad, Inserto_estr)
'        Else
'            Nro_GrupoSeguridad = 0
'        End If
'    Else
        Nro_GrupoSeguridad = 0
'    End If

' =====================================================================================================
'                                         FIN LECTURA DE CAMPOS
' =====================================================================================================
' =====================================================================================================
' =====================================================================================================

' =====================================================================================================
'                                    COMIENZO A INSERTAR LOS REGISTROS
' =====================================================================================================
' =====================================================================================================

  'JPB - Inicializo el pass+perfil (l_ess_Pass_Estandar y l_ess_Perfil_Estandar) por defecto del Autogestion en el caso que este habilitado el confper 12
  '--------------------------------
  Call ESS_Configuracion_Default
  '-------------------------------

  '_____________________________
  ' INSERTO EL TERCERO
  '-----------------------------
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
        '________________________________________________
        'Busco si esta config la maxima cant de empleados
        MaxEmpl = 0
        StrSql = "SELECT lib_generica FROM sistema"
        OpenRecordset StrSql, rs_Sql
        If Not EsNulo(rs_Sql!lib_generica) Then
            If IsNumeric(Decryptar("56238", rs_Sql!lib_generica)) Then MaxEmpl = Decryptar("56238", rs_Sql!lib_generica)
        End If
        
        If MaxEmpl > 0 Then
            '_______________________________________
            'Calculo la cantidad actual de empleados
            StrSql = "SELECT COUNT(empleg) cant FROM empleado"
            OpenRecordset StrSql, rs_Sql
            If Not rs_Sql.EOF Then
                If Not EsNulo(rs_Sql!cant) Then CantEmpl = rs_Sql!cant
            End If
            '_______________________________________
            'Controlo cantidad
            If CantEmpl >= MaxEmpl Then
                Texto = ": " & Replace(EscribeLogMI("El sistema alcanzo el maximo de @@NUM@@ empleados permitidos"), "@@NUM@@", MaxEmpl)
                'Texto = ": " & "El sistema alcanzo el maximo de " & MaxEmpl & " empleados permitidos."
                Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
                RegError = RegError + 1
                Exit Sub
            End If
        End If
    
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
        
        Texto = ": " & EscribeLogMI("Codigo de Tercero") & " = " & NroTercero
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
    Else
        StrSql = " UPDATE tercero SET "
        StrSql = StrSql & " ternom = " & "'" & nombre & "'"
        StrSql = StrSql & ", terape = " & "'" & Apellido & "'"
        StrSql = StrSql & ", terfecnac = " & F_Nacimiento
        StrSql = StrSql & ", tersex = " & ter_sexo
        StrSql = StrSql & ", terestciv =" & nro_estcivil
        StrSql = StrSql & ", estcivnro =" & "'" & nro_estcivil & "'"
        If nro_nacionalidad <> 0 Then
            StrSql = StrSql & ", nacionalnro =" & nro_nacionalidad
        End If
        StrSql = StrSql & ", terfecing =" & F_Ingreso
        If nro_paisnac <> 0 Then
            StrSql = StrSql & ", paisnro =" & nro_paisnac
        End If
        StrSql = StrSql & " WHERE ternro = " & NroTercero
        objConn.Execute StrSql, , adExecuteNoRecords
      
        Texto = ": " & " - " & EscribeLogMI("Empleado") & ": " & Legajo & " - " & EscribeLogMI("Ese Empleado ya existe en la base.")
        Texto = Texto & " " & EscribeLogMI("Datos de tercero actualizados")
        NroColumna = 1
        Call Escribir_Log("flogp", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    
    If Not ExisteLeg Then
        StrSql = " INSERT INTO empleado(empleg,empfecalta,empfecbaja,empest,empfaltagr,empfbajaprev,"
        StrSql = StrSql & "ternro,nivnro,empestudia,terape,ternom,empinterno,empemail,"
        StrSql = StrSql & "empnro,tplatenro,empremu"
        If Nro_Reporta_a <> 0 Then
            StrSql = StrSql & ",empreporta"
        End If
        'JPB - Si esta habilitado el confper 12 asigno el pass+perfil por defecto al empleado
        '------------------------------------------------------------------------------------
        If (l_ess_Pass_Estandar <> "") And (l_ess_Perfil_Estandar <> "") Then
             StrSql = StrSql & ",empessactivo,emppass,perfnro"
        End If
        '------------------------------------------------------------------------------------
        StrSql = StrSql & ") VALUES("
        StrSql = StrSql & Legajo & "," & F_Alta & "," & F_Baja & "," & ter_estado & "," & F_Alta & "," & Fecha_FinContrato & ","
        StrSql = StrSql & NroTercero & "," & nro_nivest & "," & ter_estudio & ",'" & Apellido & "','"
        StrSql = StrSql & nombre & "',Null,'" & Email & "',1," & nro_ModOrg & "," & Empremu
        If Nro_Reporta_a <> 0 Then
            StrSql = StrSql & "," & Nro_Reporta_a
        End If
        'JPB - Si esta habilitado el confper 12 asigno el pass+perfil por defecto al empleado
        '------------------------------------------------------------------------------------
        If (l_ess_Pass_Estandar <> "") And (l_ess_Perfil_Estandar <> "") Then
             StrSql = StrSql & ",-1,'" & l_ess_Pass_Estandar & "'," & l_ess_Perfil_Estandar
        End If
        '------------------------------------------------------------------------------------
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Texto = ": " & EscribeLogMI("Empleado insertado") & " - " & Legajo & " - " & Apellido & " - " & nombre
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
    Else
        StrSql = " UPDATE empleado SET "
        StrSql = StrSql & " empleg = " & Legajo
        StrSql = StrSql & " ,empfecalta = " & F_Alta
        StrSql = StrSql & " ,empfecbaja = " & F_Baja
        StrSql = StrSql & " ,empfbajaprev = " & Fecha_FinContrato
        StrSql = StrSql & " ,empest = " & ter_estado
        StrSql = StrSql & " ,empfaltagr = " & F_Alta
        StrSql = StrSql & " ,nivnro = " & nro_nivest
        StrSql = StrSql & " ,empestudia = " & ter_estudio
        StrSql = StrSql & " ,terape = " & "'" & Apellido & "'"
        StrSql = StrSql & " ,ternom = " & "'" & nombre & "'"
        StrSql = StrSql & " ,empemail = " & "'" & Email & "'"
        StrSql = StrSql & " ,empnro = 1 "
        StrSql = StrSql & " ,tplatenro =" & nro_ModOrg
        StrSql = StrSql & " ,Empremu =" & Empremu
        If Nro_Reporta_a <> 0 Then
            StrSql = StrSql & ", empreporta =" & Nro_Reporta_a
        End If
    
        'JPB - Si esta habilitado el confper 12 asigno el pass+perfil por defecto al empleado
        '------------------------------------------------------------------------------------
        If (l_ess_Pass_Estandar <> "") And (l_ess_Perfil_Estandar <> "") Then
            StrSql = StrSql & ", empessactivo = -1"
            StrSql = StrSql & ", emppass = '" & l_ess_Pass_Estandar & "'"
            StrSql = StrSql & ", perfnro =" & l_ess_Perfil_Estandar
        End If
        '------------------------------------------------------------------------------------
        StrSql = StrSql & " WHERE ternro = " & NroTercero
        objConn.Execute StrSql, , adExecuteNoRecords

        Texto = ": " & EscribeLogMI("Empleado Actualizado") & " - " & Legajo & " - " & Apellido & " - " & nombre
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
    End If
    
    ' Inserto el Registro correspondiente en ter_tip
    If Not ExisteLeg Then
        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & NroTercero & ",1)"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

' Inserto los Documentos
    If Not ExisteLeg Then
        If Nro_TDocumento <> 0 Then
            StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
            StrSql = StrSql & " VALUES(" & NroTercero & "," & Nro_TDocumento & ",'" & Ndocu & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            Texto = ": " & EscribeLogMI("Inserte el Documento") & " - "
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
        End If
    Else
        If Nro_TDocumento <> 0 Then
            StrSql = "SELECT * FROM ter_doc WHERE ternro = " & NroTercero
            StrSql = StrSql & " AND tidnro = " & Nro_TDocumento
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If rs.EOF Then
                StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                StrSql = StrSql & " VALUES(" & NroTercero & "," & Nro_TDocumento & ",'" & Ndocu & "')"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Texto = ": " & EscribeLogMI("Inserte el Documento") & " - "
                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
            Else
                StrSql = " UPDATE ter_doc SET "
                StrSql = StrSql & " nrodoc = '" & Ndocu & "'"
                StrSql = StrSql & " WHERE ternro = " & NroTercero
                StrSql = StrSql & " AND tidnro = " & Nro_TDocumento
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Texto = Texto = ": " & EscribeLogMI("Documento actualizado") & " - "
                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
            End If
        End If
    End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'    If Not ExisteLeg Then                   ' VER ESTOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
'        If Cuil <> "" Then
'            StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
'            StrSql = StrSql & " VALUES(" & NroTercero & ",10,'" & Cuil & "')"
'            objConn.Execute StrSql, , adExecuteNoRecords
'            texto = ": " & "Inserte el CUIL - "
'            Call Escribir_Log("flogp", LineaCarga, 1, texto, Tabs + 1, strReg)
'        End If
'    Else
'        If Cuil <> "" Then
'            StrSql = "SELECT * FROM ter_doc WHERE ternro = " & NroTercero
'            StrSql = StrSql & " AND tidnro = 10 "
'            If rs.State = adStateOpen Then rs.Close
'            OpenRecordset StrSql, rs
'            If rs.EOF Then
'                StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
'                StrSql = StrSql & " VALUES(" & NroTercero & ",10,'" & Cuil & "')"
'                objConn.Execute StrSql, , adExecuteNoRecords
'                texto = ": " & "Inserte el CUIL - "
'                Call Escribir_Log("flogp", LineaCarga, 1, texto, Tabs + 1, strReg)
'            Else
'                StrSql = " UPDATE ter_doc SET "
'                StrSql = StrSql & " nrodoc = '" & Cuil & "'"
'                StrSql = StrSql & " WHERE ternro = " & NroTercero
'                StrSql = StrSql & " AND tidnro = 10"
'                objConn.Execute StrSql, , adExecuteNoRecords
'
'                texto = texto = ": " & "CUIL actualizado - "
'                Call Escribir_Log("flogp", LineaCarga, 1, texto, Tabs + 1, strReg)
'            End If
'        End If
'    End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
' Inserto el Domicilio
  If rs.State = adStateOpen Then
    rs.Close
  End If
  
  If Not ExisteLeg Then
    'If (Nro_Localidad <> 0 And Nro_Provincia <> 0 And Nro_Pais <> 0) Then
    If (Nro_Localidad <> 0 And Nro_Pais <> 0) Then
        StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
        StrSql = StrSql & " VALUES(1," & NroTercero & ",-1,2)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        NroDom = getLastIdentity(objConn, "cabdom")
      
        StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,entrecalles,"
        StrSql = StrSql & "locnro,provnro,paisnro,barrio,partnro,zonanro) "
        StrSql = StrSql & " VALUES (" & NroDom & ",'" & Calle & "','" & Nro_Nrodom & "','" & Piso & "','"
        StrSql = StrSql & Depto & "','" & Torre & "','" & Manzana & "','" & Cpostal & "','" & Entre & "'," & Nro_Localidad & ","
        StrSql = StrSql & Nro_Provincia & "," & Nro_Pais & ",'" & Barrio & "'," & Nro_Partido & "," & Nro_Zona & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
      
        Texto = ": " & EscribeLogMI("Inserte el Domicilio") & " - "
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
        
'        If Telefono <> "" Then
'            'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
'            StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
'            StrSql = StrSql & " VALUES(" & NroDom & ",'" & Telefono & "',0,-1,0,1)"
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            Texto = ": " & "Inserte el Telefono Principal - "
'            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'        End If
'        If TelLaboral <> "" Then
'            StrSql = "SELECT * FROM telefono "
'            StrSql = StrSql & " WHERE domnro =" & NroDom
'            StrSql = StrSql & " AND telnro ='" & TelLaboral & "'"
'            If rs_Tel.State = adStateOpen Then rs_Tel.Close
'            OpenRecordset StrSql, rs_Tel
'            If rs_Tel.EOF Then
'                'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
'                StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
'                StrSql = StrSql & " VALUES(" & NroDom & ",'" & TelLaboral & "',0,0,0,3)"
'                objConn.Execute StrSql, , adExecuteNoRecords
'
'                Texto = ": " & "Inserte el Telefono Laboral - "
'                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'            End If
'        End If
'        If TelCelular <> "" Then
'            StrSql = "SELECT * FROM telefono "
'            StrSql = StrSql & " WHERE domnro =" & NroDom
'            StrSql = StrSql & " AND telnro ='" & TelCelular & "'"
'            If rs_Tel.State = adStateOpen Then rs_Tel.Close
'            OpenRecordset StrSql, rs_Tel
'            If rs_Tel.EOF Then
'                'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
'                StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
'                StrSql = StrSql & " VALUES(" & NroDom & ",'" & TelCelular & "',0,0,-1,2)"
'                objConn.Execute StrSql, , adExecuteNoRecords
'                Texto = ": " & "Inserte el Telefono Celular - "
'                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'            End If
'        End If
    End If
  Else
    If (Nro_Localidad <> 0 And Nro_Provincia <> 0 And Nro_Pais <> 0) Then
    'If (Nro_Localidad <> 0 And Nro_Pais <> 0) Then
        StrSql = "SELECT * FROM cabdom  "
        StrSql = StrSql & " WHERE tipnro = 1"
        StrSql = StrSql & " AND ternro = " & NroTercero
        StrSql = StrSql & " AND domdefault = -1"
        StrSql = StrSql & " AND tidonro = 2"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If rs.EOF Then
          StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
          StrSql = StrSql & " VALUES(1," & NroTercero & ",-1,2)"
          objConn.Execute StrSql, , adExecuteNoRecords
          
          NroDom = getLastIdentity(objConn, "cabdom")
        
          StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,entrecalles,"
          StrSql = StrSql & "locnro,provnro,paisnro,barrio,partnro,zonanro) "
          StrSql = StrSql & " VALUES (" & NroDom & ",'" & Calle & "','" & Nro_Nrodom & "','" & Piso & "','"
          StrSql = StrSql & Depto & "','" & Torre & "','" & Manzana & "','" & Cpostal & "','" & Entre & "'," & Nro_Localidad & ","
          StrSql = StrSql & Nro_Provincia & "," & Nro_Pais & ",'" & Barrio & "'," & Nro_Partido & "," & Nro_Zona & ")"
          objConn.Execute StrSql, , adExecuteNoRecords
        
          Texto = ": " & EscribeLogMI("Inserte el Domicilio") & " - "
          Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
          
'          If Telefono <> "" Then
'            'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
'              StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
'              StrSql = StrSql & " VALUES(" & NroDom & ",'" & Telefono & "',0,-1,0,1)"
'              objConn.Execute StrSql, , adExecuteNoRecords
'
'              Texto = ": " & "Inserte el Telefono Principal - "
'              Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'          End If
'          If TelLaboral <> "" Then
'              StrSql = "SELECT * FROM telefono "
'              StrSql = StrSql & " WHERE domnro =" & NroDom
'              StrSql = StrSql & " AND telnro ='" & TelLaboral & "'"
'              If rs_Tel.State = adStateOpen Then rs_Tel.Close
'              OpenRecordset StrSql, rs_Tel
'              If rs_Tel.EOF Then
'                'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
'                  StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
'                  StrSql = StrSql & " VALUES(" & NroDom & ",'" & TelLaboral & "',0,0,0,3)"
'                  objConn.Execute StrSql, , adExecuteNoRecords
'
'                  Texto = ": " & "Inserte el Telefono Laboral - "
'                  Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'              End If
'          End If
'          If TelCelular <> "" Then
'              StrSql = "SELECT * FROM telefono "
'              StrSql = StrSql & " WHERE domnro =" & NroDom
'              StrSql = StrSql & " AND telnro ='" & TelCelular & "'"
'              If rs_Tel.State = adStateOpen Then rs_Tel.Close
'              OpenRecordset StrSql, rs_Tel
'              If rs_Tel.EOF Then
'                'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
'                  StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
'                  StrSql = StrSql & " VALUES(" & NroDom & ",'" & TelCelular & "',0,0,-1,2)"
'                  objConn.Execute StrSql, , adExecuteNoRecords
'                  Texto = ": " & "Inserte el Telefono Celular - "
'                  Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'              End If
'          End If
        Else
            StrSql = " UPDATE detdom SET "
            StrSql = StrSql & " calle =" & "'" & Calle & "'"
            StrSql = StrSql & ",nro =" & "'" & Nro_Nrodom & "'"
            StrSql = StrSql & ",piso =" & "'" & Piso & "'"
            StrSql = StrSql & ",oficdepto =" & "'" & Depto & "'"
            StrSql = StrSql & ",torre =" & "'" & Torre & "'"
            StrSql = StrSql & ",manzana =" & "'" & Manzana & "'"
            StrSql = StrSql & ",codigopostal =" & "'" & Cpostal & "'"
            StrSql = StrSql & ",entrecalles =" & "'" & Entre & "'"
            StrSql = StrSql & ",locnro =" & Nro_Localidad
            StrSql = StrSql & ",provnro =" & Nro_Provincia
            StrSql = StrSql & ",paisnro =" & Nro_Pais
            StrSql = StrSql & ", partnro = " & Nro_Partido
            StrSql = StrSql & ", zonanro =" & Nro_Zona
            StrSql = StrSql & " WHERE domnro = " & rs!domnro
            objConn.Execute StrSql, , adExecuteNoRecords
        
            Texto = ": " & EscribeLogMI("Domicilio Actualizado") & " - "
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
        
'            If Telefono <> "" Then
'                StrSql = "SELECT * FROM telefono "
'                StrSql = StrSql & " WHERE domnro =" & rs!domnro
'                StrSql = StrSql & " AND telnro ='" & Telefono & "'"
'                If rs_Tel.State = adStateOpen Then rs_Tel.Close
'                OpenRecordset StrSql, rs_Tel
'                If rs_Tel.EOF Then
'                    'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
'                    StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
'                    StrSql = StrSql & " VALUES(" & rs!domnro & ",'" & Telefono & "',0,-1,0,1)"
'                    objConn.Execute StrSql, , adExecuteNoRecords
'
'                    Texto = ": " & "Inserte el Telefono Principal - "
'                    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'                End If
'            End If
'            If TelLaboral <> "" Then
'                StrSql = "SELECT * FROM telefono "
'                StrSql = StrSql & " WHERE domnro =" & rs!domnro
'                StrSql = StrSql & " AND telnro ='" & TelLaboral & "'"
'                If rs_Tel.State = adStateOpen Then rs_Tel.Close
'                OpenRecordset StrSql, rs_Tel
'                If rs_Tel.EOF Then
'                    'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
'                    StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
'                    StrSql = StrSql & " VALUES(" & rs!domnro & ",'" & TelLaboral & "',0,0,0,3)"
'                    objConn.Execute StrSql, , adExecuteNoRecords
'
'                    Texto = ": " & "Inserte el Telefono Laboral - "
'                    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'                End If
'            End If
'            If TelCelular <> "" Then
'                StrSql = "SELECT * FROM telefono "
'                StrSql = StrSql & " WHERE domnro =" & rs!domnro
'                StrSql = StrSql & " AND telnro ='" & TelCelular & "'"
'                If rs_Tel.State = adStateOpen Then rs_Tel.Close
'                OpenRecordset StrSql, rs_Tel
'                If rs_Tel.EOF Then
'                    'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
'                    StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
'                    StrSql = StrSql & " VALUES(" & rs!domnro & ",'" & TelCelular & "',0,0,-1,2)"
'                    objConn.Execute StrSql, , adExecuteNoRecords
'                    Texto = ": " & "Inserte el Telefono Celular - "
'                    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'                End If
'            End If
        End If
    End If
  End If
  


If Not ExisteLeg Then
    ' Inserto las Fases
    StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
    StrSql = StrSql & " VALUES( " & NroTercero & "," & F_Alta & "," & F_Baja & ","
    If nro_causabaja <> 0 Then
        StrSql = StrSql & nro_causabaja
        StrSql = StrSql & ",0,-1,-1,-1,-1,-1)"  ' estado fase=0  - no mira ter_estado
    Else
        StrSql = StrSql & "null"
        StrSql = StrSql & "," & ter_estado & ",-1,-1,-1,-1,-1)"
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    If nro_causabaja <> 0 Then
        Call AsignarSitRevistaAsoc(nro_causabaja, NroTercero, F_Baja)
    End If
    
Else
    'FGZ - 23/07/2010
    'SI existe el legajo reviso si estaba activo o inactivo.
    '   Si estaba inactivo y ahora el estado es activo dedeuzco que se trata de un reingreso
    '   ==> intento insertar la fase.
    '   Problemas potenciales
    '
    '   Existe Fase cerrada *****
    '   Fase anterior  ------[------------------]--------
    'Casos
    '   Fecha ingreso  ---FI----------------------------- ==> no se puede insertar (informar error)
    '   Fecha ingreso  ---------   FI-------------------- ==> tenfo 2 posibilidades
    '                                                           Cierro fase un dia antes de FI y creo nueva fase
    '                                                           no se puede insertar (informar error)
    '   Fecha ingreso  -----------------------------FI--- ==> inserto la nueva fase
    
    '   Existe Fase abierta *****
    '   Fase anterior  ------[---------------------------
    'Casos
    '   Fecha ingreso  ---FI----------------------------- ==> no se puede insertar (informar error)
    '   Fecha ingreso  -------------FI------------------- ==> Cierro fase un dia antes de FI y creo nueva fase
   
    'Si no existe fase ==> simplemente crea la fase
    
    StrSql = "SELECT * FROM fases WHERE empleado = " & NroTercero
    StrSql = StrSql & " ORDER BY altfec DESC"
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
        'No Existe ==> Inserto
        StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
        StrSql = StrSql & " VALUES( " & NroTercero & "," & F_Alta & "," & F_Baja & ","
        If nro_causabaja <> 0 Then
          StrSql = StrSql & nro_causabaja
          StrSql = StrSql & ",0,-1,-1,-1,-1,-1)"  ' estado fase=0  - no mira ter_estado
        Else
          StrSql = StrSql & "null"
          StrSql = StrSql & "," & ter_estado & ",-1,-1,-1,-1,-1)"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        If nro_causabaja <> 0 Then
            Call AsignarSitRevistaAsoc(nro_causabaja, NroTercero, F_Baja)
        End If
        
    Else
        'Existe
        If CBool(rs!Estado) Then
            Texto = ": " & EscribeLogMI("Existe Fase activa.")
            'Activa
            Texto = Texto & " " & rs!altfec & " - " & IIf(EsNulo(rs!bajfec), "#", rs!bajfec)
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
            
            
            'Si ahora está inactivo ==> actualiza la fase
            If Not CBool(ter_estado) Then
                'la cierro y abro otro
                StrSql = " UPDATE fases SET "
                StrSql = StrSql & " bajfec =" & F_Baja
                StrSql = StrSql & ",estado = 0 "
                If nro_causabaja <> 0 Then
                StrSql = StrSql & ", caunro =" & nro_causabaja
                End If
                StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
                objConn.Execute StrSql, , adExecuteNoRecords
            
                Texto = ": " & EscribeLogMI("Fase Actualizada.")
                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                
                If nro_causabaja <> 0 Then
                    Call AsignarSitRevistaAsoc(nro_causabaja, NroTercero, F_Baja)
                End If
                
            Else
                Texto = ": " & EscribeLogMI("Si desea actualizar debe corregir la situacion manualmente.")
                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
            End If
            
        Else
            'Inactiva
            'Texto = Texto & " inactiva "
            If EsNulo(rs!bajfec) Then
                Texto = EscribeLogMI("Existe fase inactiva abierta")
                Texto = Texto & " " & rs!altfec & " - #"
                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                
                If CDate(rs!altfec) < CDate(FAlta) Then
                    Texto = ": " & EscribeLogMI("Cierro la fase anterior (un dia antes) y creo la nueva fase")
                    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    
                    'la cierro y abro otro - (se cierra a FAlta-1)
                    StrSql = " UPDATE fases SET "
                    StrSql = StrSql & "bajfec =" & ConvFecha(DateAdd("d", -1, FAlta))
                    StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                        
                    'abro una nueva
                    StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                    StrSql = StrSql & " VALUES( " & NroTercero & "," & F_Alta & "," & F_Baja & ","
                    If nro_causabaja <> 0 Then
                      StrSql = StrSql & nro_causabaja
                      StrSql = StrSql & ",0,-1,-1,-1,-1,-1)"
                    Else
                      StrSql = StrSql & "null"
                      StrSql = StrSql & "," & ter_estado & ",-1,-1,-1,-1,-1)"
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Texto = ": " & EscribeLogMI("Fase Actualizada.")
                    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    
                    If nro_causabaja <> 0 Then
                        Call AsignarSitRevistaAsoc(nro_causabaja, NroTercero, F_Baja)
                    End If
                    
                Else
                    If CDate(rs!altfec) > CDate(FAlta) Then
                        Texto = ": " & EscribeLogMI("No se puede actualizar las fases.") & " " & EscribeLogMI("Debe corregir la situacion manualmente.")
                        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    Else
                        'es la misma fase, no hago nada.
                        ' actualiza
                        
                        StrSql = " UPDATE fases SET "
                        StrSql = StrSql & " bajfec =" & F_Baja
                        If nro_causabaja <> 0 Then
                        StrSql = StrSql & ", caunro =" & nro_causabaja
                        End If
                        StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    
                        Texto = ": " & EscribeLogMI("Fase Actualizada.") & " - "
                        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                            
                        If nro_causabaja <> 0 Then
                            Call AsignarSitRevistaAsoc(nro_causabaja, NroTercero, F_Baja)
                        End If
                            
                    End If
                End If
            Else
                Texto = EscribeLogMI("Existe fase cerrada") & " " & rs!altfec & " - " & rs!bajfec
                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
            
                If CDate(rs!altfec) < CDate(FAlta) And CDate(rs!bajfec) < CDate(FAlta) Then
                    'Fase existente ------[------]-----------
                    'Nueva fase     ---------------[-----]---
                    
                    'abro una nueva
                    StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
                    StrSql = StrSql & " VALUES( " & NroTercero & "," & F_Alta & "," & F_Baja & ","
                    If nro_causabaja <> 0 Then
                      StrSql = StrSql & nro_causabaja
                      StrSql = StrSql & ",0,-1,-1,-1,-1,-1)"
                    Else
                      StrSql = StrSql & "null"
                      StrSql = StrSql & "," & ter_estado & ",-1,-1,-1,-1,-1)"
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Texto = ": " & EscribeLogMI("Fase creada.")
                    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    
                    If nro_causabaja <> 0 Then
                        Call AsignarSitRevistaAsoc(nro_causabaja, NroTercero, F_Baja)
                    End If
                    
                Else
                    'FGZ - 11/11/2010 --------------------------
                    If CDate(rs!altfec) = CDate(FAlta) And CDate(rs!bajfec) = CDate(FBaja) Then
                        If nro_causabaja <> 0 Then
                            StrSql = " UPDATE fases SET "
                            StrSql = StrSql & " caunro =" & nro_causabaja
                            StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                    
                            Texto = ": " & EscribeLogMI("Fase Actualizada") & " - "
                            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                            
                            Call AsignarSitRevistaAsoc(nro_causabaja, NroTercero, F_Baja)
                            
                        Else
                        
                        End If
                    Else
                        Texto = ": " & EscribeLogMI("No se puede crear la fase nueva.") & " " & EscribeLogMI("Debe corregir la situacion manualmente.")
                        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                    End If
                End If
            End If
        End If
    End If
'        StrSql = " UPDATE fases SET "
'        StrSql = StrSql & " altfec =" & F_Alta
'        StrSql = StrSql & ",bajfec =" & F_Baja
'        StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        Texto = ": " & "Fase Actualizada - "
'        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'    End If
End If
    
    
    
'  18/02/2010 - No hacer nada con las fases cuando existe el empleado
'  Else
'    StrSql = "SELECT * FROM fases WHERE estado = -1 AND sueldo = -1 AND vacaciones = -1 AND indemnizacion = -1 AND real = -1 AND empleado = " & NroTercero
'    OpenRecordset StrSql, rs
'    If rs.EOF Then
'        StrSql = " INSERT INTO fases(empleado,altfec,bajfec,caunro,estado,sueldo,vacaciones,indemnizacion,real,fasrecofec)"
'        StrSql = StrSql & " VALUES( " & NroTercero & "," & F_Alta & "," & F_Baja & ","
'        If nro_causabaja <> 0 Then
'          StrSql = StrSql & nro_causabaja
'        Else
'          StrSql = StrSql & "null"
'        End If
'        StrSql = StrSql & "," & ter_estado & ",-1,-1,-1,-1,-1)"
'        objConn.Execute StrSql, , adExecuteNoRecords
'    Else
'        StrSql = " UPDATE fases SET "
'        StrSql = StrSql & " altfec =" & F_Alta
'        StrSql = StrSql & ",bajfec =" & F_Baja
'        StrSql = StrSql & " WHERE fasnro = " & rs!fasnro
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        Texto = ": " & "Fase Actualizada - "
'        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
'    End If
  
  'Inserto la cuenta bancaria
    If Not ExisteLeg Then
        'FGZ - 09/04/2010 ------------------------------------------------------------------------------
        'FGZ - 09/04/2010 - Es obligatorio el nro de cuenta o el CBU -----------------------------------
        'If (nro_formapago <> 0 And ter_bancopago <> 0 And NroCuenta <> "") Then
        If (nro_formapago <> 0 And ter_bancopago <> 0 And (NroCuenta <> "" Or NroCBU <> "")) Then
                StrSql = " INSERT INTO ctabancaria (ternro,fpagnro,banco,ctabestado,"
                StrSql = StrSql & "ctabsuc,ctabnro,ctabporc,ctabcbu"
                If Not EsNulo(NroCuentaAcreditacionE) Then
                    StrSql = StrSql & ",ctabacred"
                End If
                StrSql = StrSql & ") VALUES ("
                StrSql = StrSql & NroTercero & "," & nro_formapago & "," & ter_bancopago & ","
                StrSql = StrSql & "-1,'" & Left(SucBanco, 10) & "','" & NroCuenta & "',100,'" & NroCBU & "'"
                If Not EsNulo(NroCuentaAcreditacionE) Then
                    StrSql = StrSql & ",'" & NroCuentaAcreditacionE & "'"
                End If
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                Texto = ": " & EscribeLogMI("Inserte la Cuenta Bancaria") & " - "
                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs, strReg)
        End If
    Else
        'If (nro_formapago <> 0 And ter_bancopago <> 0 And NroCuenta <> "") Then
        If (nro_formapago <> 0 And ter_bancopago <> 0 And (NroCuenta <> "" Or NroCBU <> "")) Then
            StrSql = "SELECT * FROM ctabancaria"
            StrSql = StrSql & " WHERE ctabancaria.ternro =" & NroTercero
            StrSql = StrSql & " AND ctabestado = -1 "
            StrSql = StrSql & " AND banco =" & ter_bancopago
            StrSql = StrSql & " AND fpagnro =" & nro_formapago
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If rs.EOF Then
                StrSql = " INSERT INTO ctabancaria (ternro,fpagnro,banco,ctabestado,"
                StrSql = StrSql & "ctabsuc,ctabnro,ctabporc,ctabcbu"
                If Not EsNulo(NroCuentaAcreditacionE) Then
                    StrSql = StrSql & ",ctabacred"
                End If
                StrSql = StrSql & ") VALUES ("
                StrSql = StrSql & NroTercero & "," & nro_formapago & "," & ter_bancopago & ","
                StrSql = StrSql & "-1,'" & Left(SucBanco, 10) & "','" & NroCuenta & "',100,'" & NroCBU & "'"
                If Not EsNulo(NroCuentaAcreditacionE) Then
                    StrSql = StrSql & ",'" & NroCuentaAcreditacionE & "'"
                End If
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                Texto = ": " & EscribeLogMI("Inserte la Cuenta Bancaria") & " - "
                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
            Else
                If rs!ctabnro = NroCuenta Then
                    StrSql = "UPDATE ctabancaria SET "
                    StrSql = StrSql & " ctabsuc = '" & SucBanco & "'"
                    If Not EsNulo(NroCuentaAcreditacionE) Then
                        StrSql = StrSql & ", ctabacred = '" & NroCuentaAcreditacionE & "'" '20-12-2010
                    End If
                    StrSql = StrSql & ", ctabnro = '" & NroCuenta & "'"
                    StrSql = StrSql & ", ctabporc = 100 "
                    StrSql = StrSql & ", ctabcbu = '" & NroCBU & "'"
                    StrSql = StrSql & " WHERE ctabancaria.ternro =" & NroTercero
                    StrSql = StrSql & " AND ctabestado = -1 "
                    StrSql = StrSql & " AND banco =" & ter_bancopago
                    StrSql = StrSql & " AND fpagnro =" & nro_formapago
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Texto = ": " & EscribeLogMI("Cuenta Bancaria actualizada") & " - "
                    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                Else
                    'Desactivo la anterior
                    StrSql = " UPDATE ctabancaria SET "
                    StrSql = StrSql & " ctabestado = 0 "
                    StrSql = StrSql & " WHERE cbnro = " & rs!Cbnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'inserto la nueva
                    StrSql = " INSERT INTO ctabancaria (ternro,fpagnro,banco,ctabestado,"
                    StrSql = StrSql & "ctabsuc,ctabnro,ctabporc,ctabcbu"
                    If Not EsNulo(NroCuentaAcreditacionE) Then
                        StrSql = StrSql & ",ctabacred"
                    End If
                    StrSql = StrSql & ") VALUES ("
                    StrSql = StrSql & NroTercero & "," & nro_formapago & "," & ter_bancopago & ","
                    StrSql = StrSql & "-1,'" & Left(SucBanco, 10) & "','" & NroCuenta & "',100,'" & NroCBU & "'"
                    If Not EsNulo(NroCuentaAcreditacionE) Then
                        StrSql = StrSql & ",'" & NroCuentaAcreditacionE & "'"
                    End If
                    StrSql = StrSql & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Texto = ": " & EscribeLogMI("Inserte la Cuenta Bancaria") & " - "
                    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
                End If
            End If
        End If
    End If
  
  'Inserto las Estructuras
  'FGZ - 02/03/2011 - se sacó la fecha de baja para el manejo de estructuras
  F_Baja = "Null"
  Call AsignarEstructura_NEW(1, nro_sucursal, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(2, nro_sector, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(3, nro_categoria, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(4, nro_puesto, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(5, nro_ccosto, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(6, nro_gerencia, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(7, Nro_GrupoSeguridad, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(9, nro_departamento, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(10, nro_empresa, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(15, nro_cajajub, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(16, nro_sindicato, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(17, nro_osocial_elegida, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(18, nro_contrato, NroTercero, F_Alta, F_Baja)
  'Call AsignarEstructura_NEW(18, nro_contrato, NroTercero, F_Alta, Fecha_FinContrato)
  Call AsignarEstructura_NEW(19, nro_convenio, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(20, nro_lpago, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(21, nro_reghorario, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(22, nro_formaliq, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(23, nro_planos_elegida, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(24, nro_osocial_ley, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(25, nro_planos_ley, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(29, nro_actividad, NroTercero, F_Alta, F_Baja)
  
  If ter_estado = -1 Then
    'SOLO SI EL EMPLEADO ESTA ACTIVO
    Call AsignarEstructura_SitRev2(30, nro_sitrev, NroTercero, F_Alta, F_Baja)
  End If
  
  Call AsignarEstructura_NEW(31, nro_condsijp, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(35, nro_direccion, NroTercero, F_Alta, F_Baja)
  Call AsignarEstructura_NEW(40, nro_art, NroTercero, F_Alta, F_Baja)
 
  
Texto = ": " & EscribeLogMI("Linea procesada correctamente") & " "
Call Escribir_Log("flogp", LineaCarga, NroColumna, Texto, Tabs + 1, strReg)
'LineaOK.Writeline Mid(strReg, 1, Len(strReg))
OK = True
         
FinLinea:
If rs.State = adStateOpen Then
    rs.Close
End If
Exit Sub

SaltoLinea:
    Texto = ": " & " - " & EscribeLogMI("Error") & ":" & Err.Description & " -- " & EscribeLogMI("Última consulta") & ": " & StrSql
    NroColumna = 1
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    MyRollbackTrans
    OK = False
    GoTo FinLinea
End Sub
Public Sub LineaModelo_912(ByVal strReg As String, ByRef OK As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Migracion de Familiares - Portugal
' Autor      : Gonzalez Nicolás
' Fecha      : 26/01/2012
' Ultima Mod.:
' Descripcion:
'            :
'
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
Dim Nro_Nrodom      As String
Dim NroDom          As Long
Dim Nro_TDocumento  As Long
Dim nro_nacionalidad As Long
Dim nro_paisnac      As Long
Dim nro_estciv      As Long
Dim nro_Sexo        As Long
Dim nro_estudia     As Long
Dim nro_nivest      As String
Dim nro_osocial     As Long
Dim nro_planos      As Long
Dim nro_aviso       As Long
Dim nro_salario     As Long
Dim nro_gan         As Long
Dim nro_disc        As Long
Dim nro_paren        As Long
Dim Nro_Barrio          As Long
Dim Nro_Localidad       As Long
Dim Nro_Partido         As Long
Dim Nro_Zona            As Long
Dim Nro_Provincia       As Long
Dim Nro_Pais            As Long
Dim OSocial             As String
Dim ter_osocial         As Long
Dim Inserto_estr        As Boolean
Dim IngresoDom          As Boolean
Dim Nro_Institucion As Long

Dim rs_Sql          As New ADODB.Recordset
Dim rs              As New ADODB.Recordset
Dim rs_Tel          As New ADODB.Recordset


    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    'GUARDO LOS MENSAJES QUE SE REPITEN + DE 1 VEZ.
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    Dim Mje_EsNuloObli As String
    Dim Mje_Campo As String
    
    Mje_EsNuloObli = EscribeLogMI("con valor Nulo y es obligatorio")
    Mje_Campo = EscribeLogMI("campo")
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------


    LineaCarga = LineaCarga + 1
    Tabs = 1
    
    On Error GoTo SaltoLinea
    
    Flog.writeline
    FlogE.writeline
    FlogP.writeline
    '_____________________________________________
    'LEGAJO | N° MEC
    '---------------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Legajo")
    pos1 = 1
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    If IsNumeric(Aux) Then
        Legajo = Aux
    Else
        Texto = ": " & EscribeLogMI("El legajo debe ser numérico")
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        InsertaError 1, 8
        Exit Sub
    End If
    
    '-----------------------
    '-----------------------
    
    '_________________________
    'APELLIDO | SOBRENOME
    '-------------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Apellido")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
        Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Apellido = Left(Aux, 25)
    '------------------------
    '------------------------
    
    
    '___________________________
    'NOMBRE | NOMBRE
    '---------------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Nombre")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    nombre = Left(Aux, 25)
    '_________________________________________
    'FECHA DE NACIMIENTO | DATA DE NASCIMENTO
    '-----------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Fecha de Nacimiento")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Fnac = Aux
    If Fnac = "N/A" Or EsNulo(Fnac) Then
        Fnac = "Null"
    Else
       Fnac = ConvFecha(Fnac)
    End If
    '_________________________________________
    'PAIS DE NACIMIENTO | PAIS DE NASCIMENTO
    '-----------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Pais de Nacimiento")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    PaisNac = Aux
    If PaisNac <> "N/A" Then
        StrSql = " SELECT paisnro FROM pais WHERE UPPER(paisdesc) = '" & UCase(PaisNac) & "'"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            nro_paisnac = rs!paisnro
        Else
            StrSql = " INSERT INTO pais(paisdesc,paisdef) VALUES ('" & UCase(PaisNac) & "',0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_paisnac = getLastIdentity(objConn, "pais")
        End If
    Else
        nro_paisnac = 0
    End If
    '_____________________________
    'NACIONALIDAD | NACIONALIDADE
    '-----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Nacionalidad")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    NAC = UCase(Aux)
    If NAC <> "N/A" And Not EsNulo(NAC) Then
        StrSql = " SELECT nacionalnro FROM nacionalidad WHERE upper(nacionaldes) = '" & NAC & "'"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            nro_nacionalidad = rs!nacionalnro
        Else
            StrSql = " INSERT INTO nacionalidad(nacionaldes) VALUES ('" & UCase(NAC) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_nacionalidad = getLastIdentity(objConn, "nacionalidad")
        End If
    Else
        nro_nacionalidad = 0
    End If
    If nro_nacionalidad = 0 Then
        Texto = ": " & " - " & EscribeLogMI("Debe ingresar una Nacionalidad.")
        NroColumna = 5
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    
    '-------------------
    '-------------------
    
    '_____________________________
    'ESTADO CIVIL | ESTADO CIVIL
    '-----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Estado Civil")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    EstCiv = Left(UCase(Aux), 30)
    If EstCiv <> "N/A" And Not EsNulo(EstCiv) Then
        StrSql = " SELECT estcivnro FROM estcivil WHERE upper(estcivdesabr) = '" & UCase(EstCiv) & "'"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            nro_estciv = rs!estcivnro
        Else
            StrSql = " INSERT INTO estcivil(estcivdesabr) VALUES ('" & UCase(EstCiv) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_estciv = getLastIdentity(objConn, "estcivil")
        End If
    Else
        nro_estciv = 0
    End If
    If nro_estciv = 0 Then
        Texto = ": " & " - " & EscribeLogMI("Debe ingresar Estado Civil.")
        NroColumna = 6
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    
    '_______________________
    'SEXO | SEXO
    '-----------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Sexo")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    End If
    Sexo = UCase(Aux)
    If Sexo = "M" Or Sexo = "MASCULINO" Or Sexo = "-1" Then
        nro_Sexo = -1
    Else
        nro_Sexo = 0
    End If
    '____________________________
    'PARENTESCO | PARENTESCO
    '----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Parentesco")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    GPare = UCase(Aux)
    StrSql = " SELECT parenro FROM parentesco WHERE upper(paredesc) = '" & UCase(GPare) & "'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        nro_paren = rs!parenro
    Else
        nro_paren = 0
    End If
    If nro_paren = 0 Then
        Texto = ": " & " - " & EscribeLogMI("El parentesco ingresado no existe, verifíquelo.") & ":" & GPare
        NroColumna = 8
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    '______________________________
    'DISCAPACITADO - DISCAPACITADO
    '------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Discapacitado")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Disc = UCase(Aux)
    If Disc = "N/A" Or Disc = "NAO" Or Disc = "NÃO" Or Disc = "NO" Then
        nro_disc = 0
    Else
        nro_disc = -1
    End If
    
    '________________________
    'ESTUDIA | ESTUDA
    '------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Estudia")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = UCase(Trim(Mid(strReg, pos1, pos2 - pos1 + 1)))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Estudia = Aux
    If Estudia = "N/A" Or Estudia = "NÃO" Or Estudia = "NAO" Or Estudia = "NO" Then
        nro_estudia = 0
    Else
        nro_estudia = -1
    End If
    
    '___________________________________
    'NIVEL DE ESTUDIO | NIVEL DE ESTUDO
    '-----------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Nivel de estudio")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = UCase(Trim(Mid(strReg, pos1, pos2 - pos1 + 1)))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    NivEst = Aux
    'Por ahora no hago nada con el nivel de estudio porque en Accor no lo pasaron
    If NivEst = "N/A" Or EsNulo(NivEst) Or NivEst = "NAO" Or NivEst = "NÃO" Or NivEst = "NO" Then
        'StrSql = " SELECT nivnro FROM nivest WHERE nivdes = '" & NivEst & "'"
        'OpenRecordset StrSql, rs
        nro_nivest = 0
    Else
        'busco el primer novel de estudio, si no existe la creo
        StrSql = " SELECT * FROM nivest WHERE nivdesc = '" & NivEst & "'"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            nro_nivest = rs!nivnro
        Else
            'creo una
            StrSql = " INSERT INTO nivest (nivdesc) VALUES ('" & NivEst & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            nro_nivest = getLastIdentity(objConn, "nivest")
        End If
     End If

    '______________________________________
    'TIPO DE DOCUMENTO | TIPO DE DOCUMENTO
    '--------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Tipo de documento")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    TipDoc = Aux
    If TipDoc <> "N/A" And Not EsNulo(TipDoc) Then
        StrSql = " SELECT tidnro FROM tipodocu WHERE UPPER(tidsigla) = '" & UCase(TipDoc) & "'"
        If rs_Sql.State = adStateOpen Then rs_Sql.Close
        OpenRecordset StrSql, rs_Sql
        If Not rs_Sql.EOF Then
            Nro_TDocumento = rs_Sql!tidnro
        Else
            'busco la primera institucion, si no existe la creo
            StrSql = " SELECT * FROM institucion "
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                Nro_Institucion = rs!InstNro
            Else
                'creo una
                StrSql = " INSERT INTO institucion (instdes,instabre) VALUES ('NACIONAL','NAC')"
                objConn.Execute StrSql, , adExecuteNoRecords
                Nro_Institucion = getLastIdentity(objConn, "institucion")
            End If
            'creo el tipo de documento
            StrSql = " INSERT INTO tipodocu(tidnom,tidsigla,tidsist,instnro,tidunico) VALUES ('" & UCase(TipDoc) & "','" & UCase(TipDoc) & "',0," & Nro_Institucion & ",0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Nro_TDocumento = getLastIdentity(objConn, "tipodocu")
        End If
    Else
        Nro_TDocumento = 0
    End If
    '_________________________________
    'N° DE DOCUMENTO | N° DE DOCUMENTO
    '---------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = True
    Campoetiqueta = EscribeLogMI("Nro de Documento")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    NroDoc = Aux
    If NroDoc = "N/A" Or EsNulo(NroDoc) Then
        NroDoc = "0"
    End If
    '____________________________
    'CALLE | RUA/AVENIDA
    '----------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Direccion.Calle")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Calle = Left(Aux, 250)
    IngresoDom = True
    If Calle = "N/A" Or EsNulo(Calle) Then
        Calle = ""
        IngresoDom = False
    End If
    '___________________________
    'NUMERO | NUMERO
    '---------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Direccion.Nro")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Nro = Left(Aux, 8)
    If (Nro <> "N/A") And Not EsNulo(Nro) Then
        Nro_Nrodom = Nro
    Else
        Nro_Nrodom = "S/N"
    End If
    '_________________
    'PISO | ANDAR
    '-----------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Piso")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Piso = Left(Aux, 8)
    If Piso = "N/A" Then
        Piso = ""
    End If
    
    'Depto
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Depto"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, Texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    Depto = Left(aux, 8)
'    If Depto = "N/A" Then
        Depto = ""
'    End If

    'Torre
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Torre"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, Texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    Torre = Left(aux, 8)
'    If Torre = "N/A" Then
        Torre = ""
'    End If
    
    'Manzana
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Manzana"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, Texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    Manzana = Left(aux, 8)
'    If Manzana = "N/A" Then
        Manzana = ""
'    End If
    '______________________________
    'CODIGO POSTAL | CODIGO POSTAL
    '------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = IngresoDom
    Campoetiqueta = EscribeLogMI("Codigo postal")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Cpostal = Left(Aux, 12)
    If Cpostal = "N/A" Then
        Cpostal = ""
    End If

    'Entre calles
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Entre calles"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, Texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    Entre = Left(aux, 80)
'    If Entre = "N/A" Then
        Entre = ""
'    End If

    'Barrio
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Barrio"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, Texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    Barrio = Left(aux, 30)
'    If Barrio = "N/A" Then
        Barrio = ""
'    End If
    '___________________________
    'LOCALIDAD | FREGUESIA
    '---------------------------
    NroColumna = NroColumna + 1
    Obligatorio = IngresoDom
    Campoetiqueta = EscribeLogMI("Localidad")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    '19/03/2010 - Se cambio la longitud a 60
    'Localidad = aux
    Localidad = Left(Aux, 60)
    
    
    '__________________________
    'PARTIDO | CONCELHO
    '--------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Partido")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Partido = Aux
    
    'Zona
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Zona"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, Texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    '19/03/2010 - Se cambio la longitud a 60
'    'Zona = aux
'    Zona = Left(aux, 60)
     Zona = ""
     
    '__________________________
    'PROVINCIA | DISTRITO
    '--------------------------
    NroColumna = NroColumna + 1
    Obligatorio = IngresoDom
    Campoetiqueta = EscribeLogMI("Provincia")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Provincia = Aux
    '__________________________
    'PAIS | PAIS
    '--------------------------
    NroColumna = NroColumna + 1
    Obligatorio = IngresoDom
    Campoetiqueta = EscribeLogMI("Pais")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Pais = Aux
    If Pais <> "N/A" Then
        Call ValidarPais(Pais, Nro_Pais)
    Else
        Nro_Pais = 0
    End If
    If (Provincia <> "N/A") And (Pais <> "N/A") Then
        Call ValidarProvincia(Provincia, Nro_Provincia, Nro_Pais)
    Else
        Nro_Provincia = 0
    End If
    If (Localidad <> "N/A") And (Provincia <> "N/A") And (Pais <> "N/A") Then
        Call ValidarLocalidad(Localidad, Nro_Localidad, Nro_Pais, Nro_Provincia)
    Else
        Nro_Localidad = 0
    End If
    If Partido <> "N/A" Then
        Call ValidarPartido(Partido, Nro_Partido)
    Else
        Nro_Partido = 0
    End If
    If Zona <> "N/A" Then
        Call ValidarZona(Zona, Nro_Zona, Nro_Provincia)
    Else
        Nro_Zona = 0
    End If
    
    If (IngresoDom = True) And (Nro_Localidad = 0) Then
        Texto = ": " & " - " & EscribeLogMI("Debe Ingresar la Localidad.")
        NroColumna = 23
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    
    If (IngresoDom = True) And (Nro_Provincia = 0) Then
        Texto = ": " & " - " & EscribeLogMI("Debe Ingresar la Provincia.")
        NroColumna = 26
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    
    If (IngresoDom = True) And (Nro_Pais = 0) Then
        Texto = ": " & " - " & EscribeLogMI("Debe Ingresar la Pais.")
        NroColumna = 27
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    '___________________________
    'TELEFONO | TELEFONE
    '---------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Telefono")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Telefono = Aux
    If Telefono = "N/A" Then
        Telefono = ""
    End If

    'Obra Social
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Obra Social"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, Texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    ObraSocial = aux
'    If ObraSocial = "N/A" Or ObraSocial = "" Then
'        nro_osocial = 0
'    Else
'        StrSql = " SELECT ternro FROM osocial WHERE UPPER(osdesc) = '" & UCase(ObraSocial) & "'"
'        If rs.State = adStateOpen Then rs.Close
'        OpenRecordset StrSql, rs
'        If Not rs.EOF Then
'            nro_osocial = rs!Ternro
'        Else
            nro_osocial = 0
'        End If
'    End If

    'Plan de OS
'    Nrocolumna = Nrocolumna + 1
'    Obligatorio = False
'    Campoetiqueta = "Plan de Obra Social"
'    pos1 = pos2 + 2
'    pos2 = InStr(pos1, strReg, separador) - 1
'    aux = Trim(Mid(strReg, pos1, pos2 - pos1 + 1))
'    If (aux = "N/A" Or EsNulo(aux)) And Obligatorio Then
'         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
'        Call Escribir_Log("floge", LineaCarga, Nrocolumna, Texto, Tabs, strReg)
'        RegError = RegError + 1
'        Exit Sub
'    End If
'    PlanOSocial = aux
'    If PlanOSocial = "N/A" Or PlanOSocial = "" Then
'        nro_planos = 0
'    Else
'        If nro_osocial <> 0 Then
'            StrSql = " SELECT plnro FROM planos WHERE UPPER(plnom) = '" & UCase(PlanOSocial) & "'"
'            StrSql = StrSql & " AND osocial = " & nro_osocial
'            If rs.State = adStateOpen Then rs.Close
'            OpenRecordset StrSql, rs
'            If Not rs.EOF Then
'                nro_planos = rs!plnro
'            Else
'                nro_planos = 0
'            End If
'        Else
            nro_planos = 0
'        End If
'    End If
    '________________________________________________
    'AVISO ANTE EMERGENCIA  |AVISAR ANTE EMERGËNCIA
    '------------------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Avisar ante emergencia")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = UCase(Trim(Mid(strReg, pos1, pos2 - pos1 + 1)))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    AvisoEmer = Aux
    If AvisoEmer = "" Or AvisoEmer = "N/A" Or AvisoEmer = "NAO" Or AvisoEmer = "NÃO" Then
        nro_aviso = 0
    Else
        nro_aviso = -1
    End If
    '______________________________________________
    'PAGA SALARIO FAMILIAR | PAGA SALÁRIO FAMILIA
    '----------------------------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Paga Salario Familiar")
    pos1 = pos2 + 2
    pos2 = InStr(pos1, strReg, separador) - 1
    Aux = UCase(Trim(Mid(strReg, pos1, pos2 - pos1 + 1)))
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    PagaSalario = Aux
    If PagaSalario = "N/A" Or PagaSalario = "" Or PagaSalario = "NÃO" Or PagaSalario = "NAO" Then
        nro_salario = 0
    Else
        nro_salario = -1
    End If
    '_______________________
    'GANANCIAS | IRS
    '------------------------
    NroColumna = NroColumna + 1
    Obligatorio = False
    Campoetiqueta = EscribeLogMI("Ganancias")
    pos1 = pos2 + 2
    pos2 = Len(strReg)
    If pos1 > pos2 Then
        Aux = ""
    Else
        Aux = Trim(Mid(strReg, pos1, pos2 - pos1))
    End If
    
    If (Aux = "N/A" Or EsNulo(Aux)) And Obligatorio Then
         Texto = ": " & " - " & Mje_Campo & " " & Campoetiqueta & " " & Mje_EsNuloObli
        Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
        RegError = RegError + 1
        Exit Sub
    End If
    Ganancias = Aux
    If Ganancias = "N/A" Or Ganancias = "" Or Ganancias = "NO" Or Ganancias = "NÃO" Then
        nro_gan = 0
    Else
        nro_gan = -1
    End If

' ==================================================================================
' ==================================================================================

'Busco el empleado asociado
StrSql = "SELECT ternro FROM empleado WHERE empleg = " & Legajo
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If Not rs.EOF Then
    NroEmpleado = rs!Ternro
Else
    Texto = ": " & " - Campo " & Campoetiqueta & " " & EscribeLogMI("El Legajo no Existe")
    NroColumna = 1
    Call Escribir_Log("floge", LineaCarga, NroColumna, Texto, Tabs, strReg)
    RegError = RegError + 1
    Exit Sub
End If
  
StrSql = "SELECT * FROM tercero "
StrSql = StrSql & " INNER JOIN ter_tip ON tercero.ternro = ter_tip.ternro AND ter_tip.tipnro = 3 "
StrSql = StrSql & " INNER JOIN familiar ON familiar.ternro = tercero.ternro AND familiar.empleado = " & NroEmpleado
StrSql = StrSql & " WHERE ternom = '" & nombre & "'"
StrSql = StrSql & " AND terape = '" & Apellido & "'"
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    'Inserto el tercero asociado al familiar
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
    
    'Inserto el Registro correspondiente en ter_tip
    StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & NroTercero & ",3)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Inserto el Familiar
    StrSql = " INSERT INTO familiar(empleado,ternro,parenro,famest,famestudia,famcernac,faminc,famsalario,famemergencia,famcargadgi,osocial,plnro,famternro)"
    StrSql = StrSql & " values(" & NroEmpleado & "," & NroTercero & "," & nro_paren & ",-1," & nro_estudia & ",0," & nro_disc & "," & nro_salario & "," & nro_aviso & "," & nro_gan & "," & nro_osocial & "," & nro_planos & ",0)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Inserto los estudios de familiar
    If nro_estudia = -1 Then
        StrSql = " INSERT INTO estudio_actual (ternro, nivnro) VALUES (" & NroTercero & "," & nro_nivest & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Texto = EscribeLogMI("Familiar Insertado") & " - " & Legajo & " - " & Apellido & " - " & nombre
    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
Else
    'Actualizo los datos
    StrSql = "UPDATE tercero SET "
    StrSql = StrSql & " terfecnac = " & Fnac
    StrSql = StrSql & " ,tersex = " & nro_Sexo
    If nro_nacionalidad <> 0 Then
        StrSql = StrSql & " ,nacionalnro = " & nro_nacionalidad
    End If
    If nro_paisnac <> 0 Then
        StrSql = StrSql & " ,paisnro = " & nro_paisnac
    End If
    StrSql = StrSql & " WHERE ternro = " & rs!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords

    NroTercero = rs!Ternro

    StrSql = "UPDATE familiar SET "
    StrSql = StrSql & " parenro = " & nro_paren
    StrSql = StrSql & " ,famestudia = " & nro_estudia
    StrSql = StrSql & " ,faminc = " & nro_disc
    StrSql = StrSql & " ,famsalario = " & nro_salario
    StrSql = StrSql & " ,famemergencia = " & nro_aviso
    StrSql = StrSql & " ,famcargadgi = " & nro_gan
    StrSql = StrSql & " ,osocial = " & nro_osocial
    StrSql = StrSql & " ,plnro = " & nro_planos
    StrSql = StrSql & " ,famternro = 0"
    StrSql = StrSql & " WHERE empleado = " & NroEmpleado
    StrSql = StrSql & " AND ternro = " & NroTercero
    objConn.Execute StrSql, , adExecuteNoRecords

    If nro_estudia = -1 Then
        StrSql = " SELECT ternro FROM estudio_actual WHERE ternro = " & NroTercero
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If rs.EOF Then
            StrSql = " INSERT INTO estudio_actual (ternro, nivnro) VALUES (" & NroTercero & "," & nro_nivest & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = " UPDATE estudio_actual SET nivnro = " & nro_nivest
            StrSql = StrSql & "WHERE ternro = " & NroTercero
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    Else
        'StrSql = " DELETE FROM estudio_actual WHERE ternro = " & NroTercero
        'objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Texto = EscribeLogMI("Familiar actualizado") & " - " & Legajo & " - " & Apellido & " - " & nombre
    Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
End If


'Inserto los Documentos
If NroDoc <> "" And NroDoc <> "N/A" And TipDoc <> "N/A" Then
    StrSql = "SELECT * FROM ter_doc WHERE ternro = " & NroTercero
    StrSql = StrSql & " AND tidnro = " & Nro_TDocumento
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
        StrSql = StrSql & " VALUES(" & NroTercero & "," & Nro_TDocumento & ",'" & NroDoc & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Texto = ": " & EscribeLogMI("Inserte el Documento") & " - "
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
    Else
        StrSql = " UPDATE ter_doc SET "
        StrSql = StrSql & " nrodoc = '" & NroDoc & "'"
        StrSql = StrSql & " WHERE ternro = " & NroTercero
        StrSql = StrSql & " AND tidnro = " & Nro_TDocumento
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Texto = ": " & EscribeLogMI("Documento actualizado") & " - "
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
    End If
End If


'Inserto el Domicilio
If Not IngresoDom = False Then
    StrSql = "SELECT * FROM cabdom  "
    StrSql = StrSql & " WHERE tipnro = 1"
    StrSql = StrSql & " AND ternro = " & NroTercero
    StrSql = StrSql & " AND domdefault = -1"
    StrSql = StrSql & " AND tidonro = 2"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
        StrSql = StrSql & " VALUES(1," & NroTercero & ",-1,2)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        NroDom = getLastIdentity(objConn, "cabdom")
        
        StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,"
        StrSql = StrSql & "locnro,provnro,paisnro,barrio,partnro,zonanro) "
        StrSql = StrSql & " VALUES (" & NroDom & ",'" & Calle & "','" & Nro_Nrodom & "','" & Piso & "','"
        StrSql = StrSql & Depto & "','" & Torre & "','" & Manzana & "','" & Cpostal & "'," & Nro_Localidad & ","
        StrSql = StrSql & Nro_Provincia & "," & Nro_Pais & ",'" & Barrio & "'," & Nro_Partido & "," & Nro_Zona & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Texto = ": " & EscribeLogMI("Inserte el Domicilio") & " - "
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)

        If Telefono <> "" Then
            'sebastian stremel le agrege tipo de telefono
            'StrSql = " INSERT INTO telefono(domnro,telnro,teldefault) "
            StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
            'StrSql = StrSql & " VALUES(" & NroDom & ",'" & Telefono & "',-1)"
            StrSql = StrSql & " VALUES(" & NroDom & ",'" & Telefono & "',0,-1,0,1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Texto = ": " & EscribeLogMI("Telefono insertado") & " - "
            Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
        End If
    Else
        StrSql = " UPDATE detdom SET "
        StrSql = StrSql & " calle =" & "'" & Calle & "'"
        StrSql = StrSql & ",nro =" & "'" & Nro_Nrodom & "'"
        StrSql = StrSql & ",piso =" & "'" & Piso & "'"
        StrSql = StrSql & ",oficdepto =" & "'" & Depto & "'"
        StrSql = StrSql & ",torre =" & "'" & Torre & "'"
        StrSql = StrSql & ",manzana =" & "'" & Manzana & "'"
        StrSql = StrSql & ",codigopostal =" & "'" & Cpostal & "'"
        StrSql = StrSql & ",entrecalles =" & "'" & Entre & "'"
        StrSql = StrSql & ",locnro =" & Nro_Localidad
        StrSql = StrSql & ",provnro =" & Nro_Provincia
        StrSql = StrSql & ",paisnro =" & Nro_Pais
        StrSql = StrSql & ", partnro = " & Nro_Partido
        StrSql = StrSql & ", zonanro =" & Nro_Zona
        StrSql = StrSql & " WHERE domnro = " & rs!domnro
        objConn.Execute StrSql, , adExecuteNoRecords
    
        Texto = ": " & EscribeLogMI("Domicilio Actualizado") & " - "
        Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
    
        If Telefono <> "" Then
            StrSql = "SELECT * FROM telefono "
            StrSql = StrSql & " WHERE domnro =" & rs!domnro
            StrSql = StrSql & " AND telnro ='" & Telefono & "'"
            If rs_Tel.State = adStateOpen Then rs_Tel.Close
            OpenRecordset StrSql, rs_Tel
            If rs_Tel.EOF Then
                'FGZ - 04/05/2011 --- le agregué el tipo de telefono ---------------------
                StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular,tipotel) "
                StrSql = StrSql & " VALUES(" & rs!domnro & ",'" & Telefono & "',0,-1,0,1)"
                objConn.Execute StrSql, , adExecuteNoRecords
                  
                Texto = ": " & EscribeLogMI("Telefono insertado") & " - "
                Call Escribir_Log("flogp", LineaCarga, 1, Texto, Tabs + 1, strReg)
            End If
        End If
    End If
End If

OK = True

'Cierro y libero
If rs.State = adStateOpen Then rs.Close
If rs_Tel.State = adStateOpen Then rs_Tel.Close
Set rs = Nothing
Set rs_Tel = Nothing


Exit Sub
SaltoLinea:
    Texto = ": " & Err.Description
    Call Escribir_Log("floge", LineaCarga, 1, Texto, Tabs + 1, strReg)
    MyRollbackTrans
    OK = False
End Sub



