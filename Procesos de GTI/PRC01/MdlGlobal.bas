Attribute VB_Name = "MdlGlobal"
Option Explicit

Global Const Presicion_HC = 2

Public Type TipoSubTDia
    Trabaja As Boolean
    Orden_Dia As Long
    Nro_Dia As Long
    Nro_Subturno As Long
    NombreSubTurno As String
    Dia_Libre As Boolean
End Type

Public Type Templeado
    Legajo As Long
    Ternro As Long
    Grupo As Long
    NombreGrupo As String
End Type

Public Type TRegistracionesProcesadas
    Reg_Ent As Long
    Reg_Sal As Long
End Type

'FGZ - 18/04/2006
Public Type Tjust
    Ent As String
    Sal As String
    Cantidad As Double
End Type


Public Type TRegOblig
    Ent As String
    Sal As String
    Justificada As Boolean
End Type

Public Type TAD
    thnro As Long
    Cant As Double
End Type

'FGZ - 07/10/2010 -----
Public Type TEstr
    Tenro As Long
    estrnro As Long
    rel As Long 'estrnro correspondiente al relevo
    relCantHs As Double  'Cantidad de horas relevadas
End Type
'FGZ - 07/10/2010 -----

'FGZ - 16/06/2011 -----
Public Type TTarjeta
    tptrnro As Long
    hstjnrotar As String
End Type
'FGZ - 16/06/2011 -----

Global fs
Global Flog
Global Const FechaNula = "null"
Global Const ValorNulo = "null"
Global Empleado As Templeado
Global p_fecha As Date
Global depurar As Boolean
Global blnTerminarProceso As Boolean

'Variables globales de Politicas
Global UsaConversionHoras As Boolean
Global usaTurnoTrasnoche As Boolean
Global usaDesgloseAP As Boolean
Global usaSabadoDomingo As Boolean
Global usaHorasExtras As Boolean
Global usaTopesHorasExtras As Boolean
Global usaExcedentesHorasNormales As Boolean
Global usaControlDias As Boolean
Global usaRelojEntradaSalida As Boolean
Global usaCompensacionAP As Boolean
'FGZ- 10/01/2006
Global usaTopesGralHorasExtras As Boolean
'FGZ- 07/05/2009
'FGZ- 07/10/2010 ---------
Global UsaRedondeoHoras As Boolean
Global UsaDesgloseMovilidad As Boolean
'FGZ- 07/10/2010 ---------
Global HayMinimoExtrasSinAutorizar As Boolean
Global MinimoExtrasSinAutorizar As Single
Global ListaNoAutorizable As String
Global RegAuto_Permanentes As Boolean

' FGZ - 22/10/2003
' las pasé acá porque las usa el modulo de politicas
Global E1 As String
Global E2 As String
Global E3 As String
Global S1 As String
Global S2 As String
Global S3 As String
Global FE1 As Date
Global FE2 As Date
Global FE3 As Date
Global FS1 As Date
Global FS2 As Date
Global FS3 As Date

Global Bien As Boolean
Global valor As Single

Global NroVac As Long
Global Reproceso As Boolean
Global parametros As String

Global Pliq_Nro As Integer
Global Pliq_Anio As Integer
Global Pliq_Mes As Integer

Global GeneraPorLicencia As Boolean
Global Todas As Boolean
Global TipoLicencia As Long
Global nrolicencia As Long

Global TipDiaPago As Integer
Global TipDiaDescuento As Integer
Global Tipo_Dia_Maternidad As Integer
Global Factor As Double

Global NroGrilla As Long

'FGZ - 01/12/2003
Global CantidadDeRegistraciones As Integer

'FGZ - 09/01/2004
Global primer_mes As Integer
Global primer_ano As Integer
Global PrimeraVez As Boolean

'FGZ - 20/09/2004
Global UsaFeriadoConControl As Boolean
Global GeneraLaborable_y_Feriado As Boolean
Global NroProceso As Long
Global fecha_proceso As Date

'FGZ - 02/11/2004
Global hab_pol220 As Boolean

Global objFechasHoras As New FechasHoras

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'FGZ - 20/01/2005
Global Etiqueta

Global Tipo_Hora As Long

Global Horario_Movil As Boolean
'FGZ - 28/09/2011 -------------
Global Horario_Flexible_sinParte As Boolean
'FGZ - 28/09/2011 -------------

Global Pasa_de_Dia As Boolean
Global Horas_del_Turno As Single
Global TipoRedondeo As Integer

'FGZ - 23/12/2005
Global Tipo_de_Justificacion As Integer

'FGZ - 10/01/2006
Global Tipo_de_Desgloce As Integer

'FGZ - 18/04/2006
Global Total_Hs_Justificadas As Double
Global Arr_Justificaciones(1 To 10) As Tjust
Global Indice_Justif As Integer
Global Arr_Oblig(1 To 3) As TRegOblig

'FGZ - 30/10/2006
Global Horario_Flexible_Rotativo As Boolean

'FGZ - 14/11/2006
Global Feriado_Por_Estructura As Boolean
'FGZ - 10/08/2010
Global Feriado_Laborable As Boolean

'FGZ - 26/03/2007
Global Cantidad_de_OpenRecordset As Long
Global Cantidad_Call_Politicas As Long

'Diego Rosso - 12/11/2007
Global ModificaHT As Boolean

''FGZ - 01/06/2007 - Mejoras ------
'Global Cantidad_Feriados As Long
'Global Cantidad_Turnos As Long
'Global Cantidad_Dias As Long
'Global Cantidad_Empl_Dias_Proc As Long
''FGZ - 01/06/2007 - Mejoras ------

'FGZ - 15/01/2008 -----
Global Reg_Afected
Global Nivel_Tab_Log As Long
'FGZ - 15/01/2008 -----

'FGZ - 13/02/2008 -----
Global Arr_AD() As TAD
'FGZ - 13/02/2008 -----

'FGZ - 05/11/2008 -----
Global Continua_Procesando As Boolean
'FGZ - 05/11/2008 -----
'FGZ - 05/08/2009 -----
Global Version_Valida As Boolean
'FGZ - 05/08/2009 -----

'FGZ - 18/0/2009 -----
Global ReprocesarFT As Boolean
Global ListaTHAP As String

'FGZ - 12/10/2010 --------------------------
Global TolDtoLLT As String
Global TolDtoST As String

'FGZ - 15/06/2011
Global Firma_Novedades As Boolean
Global Firma_Partes_Turno As Boolean
Global Firma_Partes_AsigHor As Boolean
Global Firma_Partes_Extras As Boolean
Global Firma_Licencias As Boolean
Global Firma_RegManuales As Boolean
Global Firma_InputFT As Boolean
Global Firma_NovLiq As Boolean

Global FIN_Firma_Novedades As Boolean
Global FIN_Firma_Partes_Turno As Boolean
Global FIN_Firma_Partes_AsigHor As Boolean
Global FIN_Firma_Partes_Extras As Boolean
Global FIN_Firma_Licencias As Boolean
Global FIN_Firma_RegManuales As Boolean
Global FIN_Firma_InputFT As Boolean
Global FIN_Firma_NovLiq As Boolean

'FGZ - 16/06/2011 -----
Global Tarjetas() As TTarjeta
Global TiempoInicialPolitica As Long
Global TiempoFinalPolitica As Long

Global User_Proceso
Global Firma_User_Destino

Global cysfirautoriza
Global cysfirusuario
Global cysfirdestino
Global cysfirfin
Global cysfiryaaut
Global cysfirrecha
Global cystipnro

'FGZ - 08/09/2011 --------------------
Global tol_LLT_Grado2 As String
Global tol_LLT_Grado3 As String
Global Tiene_tol_LLT_Grado2 As Boolean
Global Tiene_tol_LLT_Grado3 As Boolean

Global tol_ST_Grado2 As String
Global tol_ST_Grado3 As String
Global Tiene_tol_ST_Grado2 As Boolean
Global Tiene_tol_ST_Grado3 As Boolean
'FGZ - 08/09/2011 --------------------
'FGZ - 18/05/2012 -------
Global TurnoNocturno_HaciaAtras As Boolean  'EAM- Esto se setea de la version 13 de la politica 14 y se usa para insertar las registraciones en tabla temporales segun el HT

'=============================================================================================
'=============================================================================================
'=============================================================================================

Public Sub GeneraTraza(Ternro As Long, Fecha As Date, Desc As String, Optional valor As String = "?")

'    'FGZ - 31/05/2007 -------------------
'    If Not depurar Then Exit Sub
'    StrSql = "INSERT INTO gti_traza(ternro,fecproc,descripcion,valor) VALUES (" & _
'             Ternro & "," & ConvFecha(Fecha) & ",'" & Desc & "','" & valor & "')"
'    CnTraza.Execute StrSql, , adExecuteNoRecords
'    'FGZ - 31/05/2007 -------------------
   
End Sub

' no se està utilizando
Public Sub Borrar_WF()
    objConn.Execute "DELETE FROM #wf_turno", , adExecuteNoRecords
    objConn.Execute "DELETE FROM #wf_dia", , adExecuteNoRecords
End Sub


Public Function FIRST_OF(SQLString As String, campo As String, Codigo As Variant) As Boolean
Dim objRs As New ADODB.Recordset

    FIRST_OF = False
    OpenRecordset SQLString, objRs
    If Not objRs.EOF Then
        objRs.MoveFirst
        FIRST_OF = objRs.Fields(campo).Value = Codigo
    End If
End Function

Public Function LAST_OF(SQLString As String, campo As String, Codigo As Variant) As Boolean
Dim objRs As New ADODB.Recordset

    LAST_OF = False
    OpenRecordset SQLString, objRs
    If Not objRs.EOF Then
        objRs.MoveLast
        LAST_OF = objRs.Fields(campo).Value = Codigo
    End If
End Function


Public Function EsNulo(ByVal Objeto) As Boolean
    If IsNull(Objeto) Then
        EsNulo = True
    Else
        If UCase(Objeto) = "NULL" Or UCase(Objeto) = "" Then
            EsNulo = True
        Else
            EsNulo = False
        End If
    End If
End Function

Public Function Espacios(ByVal Cantidad As Integer) As String
    Espacios = Space(Cantidad)
End Function



Public Function ValidarV(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que determina si el proceso esta en condiciones de ejecutarse.
' Autor      : FGZ
' Fecha      : 05/08/2009
' Modificado : 04/05/2012 - Gonzalez Nicolás -Se comentó el conteido de CASE 10 (Dias corr).
'                           SE RESUELVE LOCALMENTE DESDE EL PROCEDIMIENTO DESDE LA FUNCION ValidarVBD
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True

Select Case TipoProceso
Case 1: 'Horario cumplido
    If Version >= "3.49" Then
        'Revisar los campos
        'gti_subturno.subtgen
        'gti_registracion.fechagen
        'gti_horcumplido.horfecgen
        Texto = "Revisar los campos: gti_subturno.subtgen, gti_registracion.fechagen, gti_horcumplido.horfecgen"
        
        StrSql = "Select subtgen from gti_subturno WHERE turnro = 1"
        OpenRecordset StrSql, rs
        
        StrSql = "Select fechagen from gti_registracion WHERE ternro = 1 AND regfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs
        
        StrSql = "Select horfecgen from gti_horcumplido WHERE ternro = 1 AND thnro = 1 AND horfecrep = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "4.00" Then
        'gti_horcumplido.Horas
        'gti_hishc.Horas
        'gti_acumdiario.Horas
        'gti_hisad.Horas
        'gti_achdiario.Horas
        'gti_his_achdiario.Horas
        'gti_det.Horas
        
        Texto = ""
        Texto = Texto & " gti_horcumplido.Horas, gti_hishc.Horas, gti_acumdiario.Horas"
        Texto = Texto & " gti_hisad.Horas , gti_achdiario.Horas,"
        Texto = Texto & " gti_his_achdiario.Horas, gti_det.Horas"
        
        StrSql = "Select horas from gti_horcumplido WHERE ternro = 1 AND thnro = 1 AND horfecrep = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        StrSql = "Select horas from gti_hishc WHERE ternro = 1 AND thnro = 1 AND horfecrep = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        StrSql = "Select horas from gti_acumdiario WHERE ternro = 1 AND thnro = 1 AND adfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        StrSql = "Select horas from gti_hisad WHERE ternro = 1 AND thnro = 1 AND adfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        StrSql = "Select horas from gti_achdiario WHERE ternro = 1 AND thnro = 1 AND achdfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs
        

        StrSql = "Select horas from gti_his_achdiario WHERE ternro = 1 AND thnro = 1 AND achdfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs
        
        StrSql = "Select horas from gti_det WHERE cgtinro = 1"
        OpenRecordset StrSql, rs
       
        V = True
    End If
    
    If Version >= "4.09" Then
        'gti_vales.valida_hcorr
        'gti_vales.hc1_ent
        'gti_vales.hc1_sal
        'gti_vales.hc2_ent
        'gti_vales.hc2_sal
        'gti_vales.hc3_ent
        'gti_vales.hc3_sal
        
        Texto = ""
        Texto = Texto & " gti_vales.valida_hcorr,"
        Texto = Texto & " gti_vales.hc1_ent , gti_vales.hc1_sal,"
        Texto = Texto & " gti_vales.hc2_ent , gti_vales.hc2_sal,"
        Texto = Texto & " gti_vales.hc3_ent , gti_vales.hc3_sal,"
        
        StrSql = "Select valida_hcorr, hc1_ent, hc1_sal, hc2_ent, hc2_sal, hc3_ent, hc3_sal from gti_vales WHERE valenro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "5.00" Then
        'gti_registracion.ft
        'gti_registracion.ftap
        'gti_novedad.ft
        'gti_novedad.ftap
        'emp_lic.ft
        'emp_lic.ftap
        'gti_cabparte.ft
        'gti_cabparte.ftap
        
        'registracion
        Texto = ""
        Texto = Texto & " gti_registracion.ft,gti_registracion.ftap "
        StrSql = "Select gti_registracion.ft,gti_registracion.ftap FROM gti_registracion WHERE regnro = 1"
        OpenRecordset StrSql, rs

        'novedades
        Texto = ""
        Texto = Texto & " gti_novedad.ft,gti_novedad.ftap "
        StrSql = "Select gti_novedad.ft,gti_novedad.ftap FROM gti_novedad WHERE gnovnro = 1"
        OpenRecordset StrSql, rs


        'licencias
        Texto = ""
        Texto = Texto & " emp_lic.ft,emp_lic.ftap "
        StrSql = "Select emp_lic.ft,emp_lic.ftap FROM emp_lic WHERE emp_licnro = 1"
        OpenRecordset StrSql, rs


        'Partes
        Texto = ""
        Texto = Texto & " gti_cabparte.ft,gti_cabparte.ftap "
        StrSql = "Select gti_cabparte.ft,gti_cabparte.ftap FROM gti_cabparte WHERE gcpnro = 1"
        OpenRecordset StrSql, rs


        V = True
    End If
    
    If Version >= "5.02" Then
        'gti_registracion.tiporeg
        
        'registracion
        Texto = ""
        Texto = Texto & " gti_registracion.tiporeg "
        StrSql = "Select gti_registracion.tiporeg FROM gti_registracion WHERE regnro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "5.03" Then
        'feriado.ferilaborable
        
        'Feriado
        Texto = ""
        Texto = Texto & " feriado.ferilaborable "
        StrSql = "Select feriado.ferilaborable FROM feriado WHERE ferinro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "5.20" Then
        'tablas nuevas
        '   WC_MOV_HORARIOS
        
        'WC_MOV_HORARIOS
        Texto = " Revisar tabla WC_MOV_HORARIOS "
        StrSql = "Select WC_MOV_HORARIOS.bpronro FROM WC_MOV_HORARIOS WHERE bpronro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "5.21" Then
        'tablas nuevas
        '   WC_LUNCH
        
        'WC_MOV_HORARIOS
        Texto = " Revisar tabla WC_LUNCH "
        StrSql = "Select WC_LUNCH.lunchnro FROM WC_LUNCH WHERE lunchnro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "5.23" Then
        'tablas nuevas
        '   WC_LUNCH
        'ALTER TABLE [dbo].[wc_lunch] ADD [tptnro] [int] NULL GO
        
        'WC_LUNCH
        Texto = " Revisar campo tptnro de la tabla WC_LUNCH "
        StrSql = "Select WC_LUNCH.tptnro FROM WC_LUNCH WHERE lunchnro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
Case 2: 'Acumulado Diarios
    If Version >= "3.32" Then
        'Revisar los campos
        'gti_subturno.subtgen
        'gti_registracion.fechagen
        'gti_horcumplido.horfecgen
        Texto = "Revisar los campos: gti_subturno.subtgen, gti_registracion.fechagen, gti_horcumplido.horfecgen"
        
        StrSql = "Select subtgen from gti_subturno WHERE turnro = 1"
        OpenRecordset StrSql, rs
        
        StrSql = "Select fechagen from gti_registracion WHERE ternro = 1 AND regfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs
        
        StrSql = "Select horfecgen from gti_horcumplido WHERE ternro = 1 AND thnro = 1 AND horfecrep = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        V = True
    End If

    If Version >= "4.00" Then
        'gti_horcumplido.Horas
        'gti_hishc.Horas
        'gti_acumdiario.Horas
        'gti_hisad.Horas
        'gti_achdiario.Horas
        'gti_his_achdiario.Horas
        'gti_det.Horas
        Texto = ""
        Texto = Texto & " gti_horcumplido.Horas, gti_hishc.Horas, gti_acumdiario.Horas"
        Texto = Texto & " gti_hisad.Horas , gti_achdiario.Horas,"
        Texto = Texto & " gti_his_achdiario.Horas, gti_det.Horas"
        
        StrSql = "Select horas from gti_horcumplido WHERE ternro = 1 AND thnro = 1 AND horfecrep = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        StrSql = "Select horas from gti_hishc WHERE ternro = 1 AND thnro = 1 AND horfecrep = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        StrSql = "Select horas from gti_acumdiario WHERE ternro = 1 AND thnro = 1 AND adfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        StrSql = "Select horas from gti_hisad WHERE ternro = 1 AND thnro = 1 AND adfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs

        StrSql = "Select horas from gti_achdiario WHERE ternro = 1 AND thnro = 1 AND achdfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs
        
        StrSql = "Select horas from gti_his_achdiario WHERE ternro = 1 AND thnro = 1 AND achdfecha = " & ConvFecha(Date)
        OpenRecordset StrSql, rs
        
        StrSql = "Select horas from gti_det WHERE cgtinro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If

    If Version >= "5.00" Then
        'gti_registracion.ft
        'gti_registracion.ftap
        'gti_novedad.ft
        'gti_novedad.ftap
        'emp_lic.ft
        'emp_lic.ftap
        'gti_cabparte.ft
        'gti_cabparte.ftap
        
        'registracion
        Texto = ""
        Texto = Texto & " gti_registracion.ft,gti_registracion.ftap "
        StrSql = "Select gti_registracion.ft,gti_registracion.ftap FROM gti_registracion WHERE regnro = 1"
        OpenRecordset StrSql, rs

        'novedades
        Texto = ""
        Texto = Texto & " gti_novedad.ft,gti_novedad.ftap "
        StrSql = "Select gti_novedad.ft,gti_novedad.ftap FROM gti_novedad WHERE gnovnro = 1"
        OpenRecordset StrSql, rs


        'licencias
        Texto = ""
        Texto = Texto & " emp_lic.ft,emp_lic.ftap "
        StrSql = "Select emp_lic.ft,emp_lic.ftap FROM emp_lic WHERE emp_licnro = 1"
        OpenRecordset StrSql, rs


        'Partes
        Texto = ""
        Texto = Texto & " gti_cabparte.ft,gti_cabparte.ftap "
        StrSql = "Select gti_cabparte.ft,gti_cabparte.ftap FROM gti_cabparte WHERE gcpnro = 1"
        OpenRecordset StrSql, rs


        V = True
    End If
    
    If Version >= "5.03" Then
        'feriado.ferilaborable
        
        'Feriado
        Texto = ""
        Texto = Texto & " feriado.ferilaborable "
        StrSql = "Select feriado.ferilaborable FROM feriado WHERE ferinro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If

    If Version >= "5.05" Then
        'Desgloses de Movilidad (Politica 589)
        'Nuevas tablas:
        
        'gti_relevos
            'CREATE TABLE gti_relevos(
            '    relnro int IDENTITY(1,1) NOT NULL,
            '    gcpnro int NOT NULL,
            '    relfecdesde datetime NULL,
            '    relfechasta datetime NULL,
            '    ternro int NOT NULL,
            '    tenro int NOT NULL,
            '    estrnro int NOT NULL
            ')
            'GO

        Texto = ""
        Texto = Texto & " relnro,gcpnro,relfecdesde,relfechasta,ternro,tenro,estrnro "
        StrSql = "Select relnro,gcpnro,relfecdesde,relfechasta,ternro,tenro,estrnro FROM gti_relevos WHERE ternro = 1"
        OpenRecordset StrSql, rs


        'gti_relevos_det
            'CREATE TABLE gti_relevos_det(
            '    reldetnro int IDENTITY(1,1) NOT NULL,
            '    relnro int NOT NULL,
            '    fecha datetime NULL
            ')
            'GO
            
        Texto = ""
        Texto = Texto & " reldetnro,relnro,fecha "
        StrSql = "Select reldetnro,relnro,fecha FROM gti_relevos_det WHERE relnro = 1"
        OpenRecordset StrSql, rs
            
            
        'gti_desgldiario
            'CREATE TABLE gti_desgldiario(
            '    desgnro int IDENTITY(1,1) NOT NULL,
            '    canthoras decimal (5, 2) NULL,
            '    fecha datetime NOT NULL,
            '    manual smallint NOT NULL,
            '    valido smallint NOT NULL,
            '    ternro int NOT NULL,
            '    thnro int NOT NULL,
            '    horas varchar(10) NULL,
            '    te1 int NOT NULL DEFAULT 0,
            '    estrnro1 int NOT NULL DEFAULT 0,
            '    te2 int NOT NULL DEFAULT 0,
            '    estrnro2 int NOT NULL DEFAULT 0,
            '    te3 int NOT NULL DEFAULT 0,
            '    estrnro3 int NOT NULL DEFAULT 0,
            '    te4 int NOT NULL DEFAULT 0,
            '    estrnro4 int NOT NULL DEFAULT 0,
            '    te5 int NOT NULL DEFAULT 0,
            '    estrnro5 int NOT NULL DEFAULT 0
            ')
            'GO
        
        Texto = ""
        Texto = Texto & " desgnro,canthoras,fecha,manual,valido,ternro,thnro,horas,te1,estrnro1,te2,estrnro2,te3,estrnro3,te4,estrnro4,te5,estrnro5 "
        StrSql = "Select desgnro,canthoras,fecha,manual,valido,ternro,thnro,horas,te1,estrnro1,te2,estrnro2,te3,estrnro3,te4,estrnro4,te5,estrnro5 FROM gti_desgldiario WHERE desgnro = 1"
        OpenRecordset StrSql, rs


        V = True
    End If
    
    'EAM- Sanciones - Monresa
    If Version >= "5.09" Then
        'Sanciones (Politica 300)
        'script -> ALTER TABLE [dbo].[gti_acumdiario] ADD [infraccion] [smallint] NOT NULL default 0 GO
        
        Texto = ""
        Texto = Texto & " gti_acumdiario.infraccion "
        StrSql = "Select infraccion FROM gti_acumdiario WHERE ternro = 1"
        OpenRecordset StrSql, rs

        
        'script -> ALTER TABLE [dbo].[gti_hisad] ADD [infraccion] [smallint] NOT NULL default 0 GO
        Texto = ""
        Texto = Texto & " gti_hisad.infraccion "
        StrSql = "Select infraccion FROM gti_hisad WHERE ternro = 1"
        OpenRecordset StrSql, rs


        V = True
    End If

    'FGZ - 20/04/2011 - Francos compensatorios
    If Version >= "5.10" Then
        'Francos compensatorios (Politica 465)
        'CREATE TABLE [dbo].[emp_fr_comp](
        '    [frannro] [int] IDENTITY(1,1) NOT NULL,
        '    [ternro] [int] NOT NULL,
        '    [fecha] [date] NOT BULL,
        '    [unidad] [int] NOT NULL,
        '    [Cantidad] Not [decimal](19, 4)
        '    [comentario] [varchar](200) NULL
        ') ON [PRIMARY]
        'GO

        Texto = ""
        Texto = Texto & " emp_fr_comp "
        StrSql = "Select frannro FROM emp_fr_comp WHERE frannro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If

'Esto se agregó para una custom de AGD
'    If Version >= "5.06" Then
'        'Distribucion de hs (Politica 710)
'        'Nuevas tablas:
'
'        'his_OT
'            'CREATE TABLE [dbo].[his_OT](
'            '    [hisotnro] [int] NOT NULL IDENTITY (1,1),
'            '    [ternro] [int] NOT NULL,
'            '    [ot] [int] NOT NULL,
'            '    [hdesde] [datetime] NOT NULL,
'            '    [hhasta] [datetime] NULL,
'            '    [hismotivo] [varchar](100) NULL,
'            '    [tipmotnro] [int] NULL
'            ') ON [PRIMARY]
'            'GO
'
'        Texto = ""
'        Texto = Texto & " hisotnro,ternro,ot,hdesde,hhasta "
'        StrSql = "Select hisotnro,ternro,ot,hdesde,hhasta FROM his_OT WHERE ternro = 1"
'        OpenRecordset StrSql, rs
'
'        'gti_disthor_det
'            'CREATE TABLE [dbo].[gti_disthor_det](
'            '    [disthordetnro] [int] NOT NULL IDENTITY (1,1),
'            '    [ternro] [int] NOT NULL,
'            '    [ot] [int] NOT NULL,
'            '    [fecha] [datetime] NOT NULL,
'            '    [autom] [smallint] NOT NULL default 0,
'            '    [canthoras] Not [decimal](5, 2)
'            ') ON [PRIMARY]
'            'GO
'
'        Texto = ""
'        Texto = Texto & " disthordetnro,ternro,fecha,ot,autom,canthoras "
'        StrSql = "Select disthordetnro,ternro,fecha,ot,autom,canthoras FROM gti_disthor_det WHERE ternro = 1"
'        OpenRecordset StrSql, rs
'
'
'        'gti_disthor
'            'CREATE TABLE [dbo].[gti_disthor](
'            '    [disthornro] [int] NOT NULL IDENTITY (1,1),
'            '    [ternro] [int] NOT NULL,
'            '    [fecha] [datetime] NOT NULL,
'            '    [normnro] Not [Int]
'            ') ON [PRIMARY]
'            'GO
'
'        Texto = ""
'        Texto = Texto & " disthornro,ternro,fecha,normnro"
'        StrSql = "Select disthornro,ternro,fecha,normnro FROM gti_disthor WHERE ternro = 1"
'        OpenRecordset StrSql, rs
'
'
'        V = True
'    End If

Case 4: 'Acumulado Parcial
    If Version >= "5.00" Then
        'gti_registracion.ft
        'gti_registracion.ftap
        'gti_novedad.ft
        'gti_novedad.ftap
        'emp_lic.ft
        'emp_lic.ftap
        'gti_cabparte.ft
        'gti_cabparte.ftap
        
        'registracion
        Texto = ""
        Texto = Texto & " gti_registracion.ft,gti_registracion.ftap "
        StrSql = "Select gti_registracion.ft,gti_registracion.ftap FROM gti_registracion WHERE regnro = 1"
        OpenRecordset StrSql, rs

        'novedades
        Texto = ""
        Texto = Texto & " gti_novedad.ft,gti_novedad.ftap "
        StrSql = "Select gti_novedad.ft,gti_novedad.ftap FROM gti_novedad WHERE gnovnro = 1"
        OpenRecordset StrSql, rs


        'licencias
        Texto = ""
        Texto = Texto & " emp_lic.ft,emp_lic.ftap "
        StrSql = "Select emp_lic.ft,emp_lic.ftap FROM emp_lic WHERE emp_licnro = 1"
        OpenRecordset StrSql, rs


        'Partes
        Texto = ""
        Texto = Texto & " gti_cabparte.ft,gti_cabparte.ftap "
        StrSql = "Select gti_cabparte.ft,gti_cabparte.ftap FROM gti_cabparte WHERE gcpnro = 1"
        OpenRecordset StrSql, rs


        V = True
    End If
    If Version >= "5.01" Then
        'feriado.ferilaborable
        
        'Feriado
        Texto = ""
        Texto = Texto & " feriado.ferilaborable "
        StrSql = "Select feriado.ferilaborable FROM feriado WHERE ferinro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
        
Case 5: 'Novedades de GTI
    
    If Version >= "5.00" Then
        'gti_registracion.ft
        'gti_registracion.ftap
        'gti_novedad.ft
        'gti_novedad.ftap
        'emp_lic.ft
        'emp_lic.ftap
        'gti_cabparte.ft
        'gti_cabparte.ftap
        
        'registracion
        Texto = ""
        Texto = Texto & " gti_registracion.ft,gti_registracion.ftap "
        StrSql = "Select gti_registracion.ft,gti_registracion.ftap FROM gti_registracion WHERE regnro = 1"
        OpenRecordset StrSql, rs

        'novedades
        Texto = ""
        Texto = Texto & " gti_novedad.ft,gti_novedad.ftap "
        StrSql = "Select gti_novedad.ft,gti_novedad.ftap FROM gti_novedad WHERE gnovnro = 1"
        OpenRecordset StrSql, rs


        'licencias
        Texto = ""
        Texto = Texto & " emp_lic.ft,emp_lic.ftap "
        StrSql = "Select emp_lic.ft,emp_lic.ftap FROM emp_lic WHERE emp_licnro = 1"
        OpenRecordset StrSql, rs


        'Partes
        Texto = ""
        Texto = Texto & " gti_cabparte.ft,gti_cabparte.ftap "
        StrSql = "Select gti_cabparte.ft,gti_cabparte.ftap FROM gti_cabparte WHERE gcpnro = 1"
        OpenRecordset StrSql, rs


        V = True
    End If


Case 10: 'Vacaciones - Dias Correspondientes
    'SE RESUELVE LOCALMENTE DESDE EL PROCEDIMIENTO DESDE LA FUNCION ValidarVBD
'    If Version >= "2.16" Then
'        'Revisar los campos
'        'vacdiascor.venc
'        Texto = "Revisar los campos: vacdiascor.venc"
'
'        StrSql = "Select venc from vacdiascor where ternro = 1"
'        OpenRecordset StrSql, rs
'
'        V = True
'    End If
'
'    If Version >= "3.18" Then
'        'dias correspondientes
'        Texto = ""
'        Texto = Texto & " vacdiascor.vdiascorcantcorr,  vacdiascor.tipvacnrocorr"
'        StrSql = "Select vacdiascor.vdiascorcantcorr,  vacdiascor.tipvacnrocorr FROM vacdiascor WHERE ternro = 1"
'        OpenRecordset StrSql, rs
'
'        V = True
'    End If
'
'    If Version >= "3.19" Then
'        'dias correspondientes
'        Texto = ""
'        Texto = Texto & " vacdiascor.vdiasfechasta"
'        StrSql = "Select vacdiascor.vdiasfechasta FROM vacdiascor WHERE ternro = 1"
'        OpenRecordset StrSql, rs
'
'
'        Texto = ""
'        Texto = Texto & " vacacion.ternro"
'        StrSql = "Select vacacion.ternro FROM vacacion WHERE ternro = 1"
'        OpenRecordset StrSql, rs
'
'        V = True
'    End If
    
Case 11: 'Vacaciones - Dias Pedidos
    If Version >= "2.06" Then
        'Revisar los campos
        'vacdiascor.venc
        Texto = "Revisar los campos: vacdiascor.venc"
        
        StrSql = "Select venc from vacdiascor where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
Case 12: 'Vacaciones - Dias Licencias
    If Version >= "2.12" Then
        'Revisar los campos
        'vacdiascor.venc
        Texto = "Revisar los campos: vacdiascor.venc"
        
        StrSql = "Select venc from vacdiascor where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "3.01" Then
        'Revisar los campos
        Texto = "Revisar los campos: tipovacac.excferilab"

        StrSql = "Select excferilab from tipovacac"
        OpenRecordset StrSql, rs

        V = True
    End If
    
Case 13: 'Vacaciones - Notificaciones
    If Version >= "2.04" Then
        'Revisar los campos

        Texto = "Revisar los campos: vacdiascor.venc"
        
        StrSql = "Select venc from vacdiascor where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
Case 14: 'Vacaciones - Pago /Descuento
    If Version >= "2.35" Then
        'Revisar los campos
        'vacdiascor.venc
        Texto = "Revisar los campos: vacdiascor.venc"
        
        StrSql = "Select venc from vacdiascor where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "3.11" Then
        'Revisar los campos vacdiascor.venc
        Texto = "Revisar los campos: vacdiascor.vdiascorcantcorr"
        
        StrSql = "Select vdiascorcantcorr from vacdiascor where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If

Case 17: 'Reporte de Ausentismo
    If Version >= "4.00" Then
        'Revisar los campos
        'rep_asp_01.horastr
        Texto = "Revisar los campos: rep_asp_01.horastr"
        
        StrSql = "Select horastr from rep_asp_01 where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
Case 256: 'Vacaciones - Vencimiento
    If Version >= "1.01" Then
        'Revisar los campos
        'vacdiascor.venc
        Texto = "Revisar los campos: vacdiascor.venc"
        
        StrSql = "Select venc from vacdiascor where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If

Case Else:
    Texto = "version correcta"
    V = True
End Select



    ValidarV = V
Exit Function

ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function



Public Sub InsertarFT(ByVal ID As Long, ByVal Tipo As Long, ByVal Origen As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: sub que inserta en la lista de entradas fuera de termino tomadas en el proceso.
' Autor      : FGZ
' Fecha      : 14/06/2010
' ---------------------------------------------------------------------------------------------
'Uso
'Call InsertarFT(ID, Tipo, Origen)

'-- Tabla de inputs ft
'input_ft
'        [idnro] Not [Int], --identity
'        [idtipoinput] [int] NOT NULL,   -- tipo de input
'        [origen] [int] NOT NULL,    -- (emplicnro, regnro, etc)
'        [ternro] [int] NULL,
'        [feccarga] [datetime] NULL ,    -- fecha de carga del input
'        [fecorigen] [datetime] NULL ,   -- fecha del origen
'        [inputestrnro] Not [Int]

'-- Tabla de tipos de input
'tipoinput_ft
'    [idtipoinput] [int] NOT NULL,
'    [origendesabr] [varchar] (200) NULL
'Datos
'   1   Registraciones
'   2   Licencias
'   3   Novedades Horarias
'   4   ABM de HC
'   5   ABM de AD
'   6   Parte de Asignacion Horaria
'   7   Parte de Cambio de Turno
'   8   Parte de Autorizacion de Hs extras
'   6   Parte de Movilidad

'-- tabla de estados de inputs
'estado_input_ft
'    [inputestrnro] [int] NOT NULL,
'    [estdesabr] Not [VarChar](200)
'Datos
'   1 Pendiente
'   2 Visto sin Autorizar (rechazado)
'   3 Autorizado
'   4 Reprocesado

On Error GoTo ME_FT

    StrSql = "INSERT INTO " & TTempWFInputFT & "(idnro, idtipoinput, origen) VALUES (" & _
             ID & "," & Tipo & "," & Origen & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
       
Exit Sub

ME_FT:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "No se puede actualizar la lista de entrada FT."
    Flog.writeline
End Sub


Public Sub ActualizarFT(ByVal Proceso As Long, ByVal Fecha As Date, ByVal Ternro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: sub que actualiza el estado de las entradas fuera de termino contempladas durante el procesamiento.
' Autor      : FGZ
' Fecha      : 14/06/2010
'Parametros  :
'   Proceso:
'           1 Horario Cumplido
'           2 Acumulado Diarioç
'           4 Acumulado Parcial
'           5 Novedades de GTI
' ---------------------------------------------------------------------------------------------
'-- Tabla de inputs ft
'input_ft
'        [idnro] Not [Int], --identity
'        [idtipoinput] [int] NOT NULL,   -- tipo de input
'        [origen] [int] NOT NULL,    -- (emplicnro, regnro, etc)
'        [ternro] [int] NULL,
'        [feccarga] [datetime] NULL ,    -- fecha de carga del input
'        [fecorigen] [datetime] NULL ,   -- fecha del origen
'        [inputestrnro] Not [Int]

'-- Tabla de tipos de input
'tipoinput_ft
'    [idtipoinput] [int] NOT NULL,
'    [origendesabr] [varchar] (200) NULL
'Datos
'   1   Registraciones
'   2   Licencias
'   3   Novedades Horarias
'   4   ABM de HC
'   5   ABM de AD
'   6   Parte de Asignacion Horaria
'   7   Parte de Cambio de Turno
'   8   Parte de Autorizacion de Hs extras
'   9   Parte de Movilidad

'-- tabla de estados de inputs
'estado_input_ft
'    [inputestrnro] [int] NOT NULL,
'    [estdesabr] Not [VarChar](200)
'Datos
'   1 Pendiente
'   2 Visto sin Autorizar (rechazado)
'   3 Autorizado
'   4 Reprocesado

Dim Tipo As Long
Dim ID As Long

Dim Rs_WF_ft As New ADODB.Recordset
Dim rs As New ADODB.Recordset

On Error GoTo ME_UFT


If Proceso = 1 Then
    'Registraciones
    StrSql = "SELECT regnro, regfecha,reghora FROM gti_registracion "
    StrSql = StrSql & " WHERE  ternro = " & Ternro
    StrSql = StrSql & " AND ( regfecha >= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (regfecha <=" & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (gti_registracion.ft = -1 AND gti_registracion.ftap = -1)"
    OpenRecordset StrSql, rs
    Do While Not rs.EOF
        ID = 0
        Tipo = 1
        
        StrSql = "INSERT INTO " & TTempWFInputFT & "(idnro, idtipoinput, origen) VALUES ("
        StrSql = StrSql & ID & "," & Tipo & "," & rs!Regnro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        rs.MoveNext
    Loop
End If

If Proceso = 1 Then
    'Licencias
    StrSql = "SELECT emp_lic.emp_licnro FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN emp_lic ON gti_justificacion.juscodext = emp_lic.emp_licnro "
    StrSql = StrSql & " WHERE (ternro = " & Ternro & ") "
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND emp_lic.licestnro = 2"
    StrSql = StrSql & " AND jussigla = 'LIC'"
    StrSql = StrSql & " AND (emp_lic.ft = -1 AND emp_lic.ftap = -1)"
    OpenRecordset StrSql, rs
    Do While Not rs.EOF
        ID = 0
        Tipo = 2
        
        StrSql = "INSERT INTO " & TTempWFInputFT & "(idnro, idtipoinput, origen) VALUES ("
        StrSql = StrSql & ID & "," & Tipo & "," & rs!emp_licnro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        rs.MoveNext
    Loop
End If

If Proceso = 1 Then
    'Novedades
    StrSql = " SELECT gti_novedad.gnovnro FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN gti_novedad ON gti_justificacion.juscodext = gti_novedad.gnovnro "
    StrSql = StrSql & " WHERE (Ternro = " & Ternro & ")"
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND jussigla = 'NOV'"
    StrSql = StrSql & " AND jussigla <> 'ALM'"
    StrSql = StrSql & " AND (gti_novedad.ft = -1 AND gti_novedad.ftap = -1)"
    OpenRecordset StrSql, rs
    Do While Not rs.EOF
        ID = 0
        Tipo = 3
        
        StrSql = "INSERT INTO " & TTempWFInputFT & "(idnro, idtipoinput, origen) VALUES ("
        StrSql = StrSql & ID & "," & Tipo & "," & rs!gnovnro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        rs.MoveNext
    Loop
End If

'Partes diarios
    'Los partes ya se insertaron cuando se buscaron
    
        StrSql = "SELECT DISTINCT idnro,idtipoinput,origen FROM " & TTempWFInputFT
        OpenRecordset StrSql, Rs_WF_ft
        Do While Not Rs_WF_ft.EOF
        
            'Inserto en batch_proceso un HC
            StrSql = "UPDATE input_ft SET inputestrnro = 4"
            StrSql = StrSql & " WHERE idtipoinput = " & Rs_WF_ft!idtipoinput
            'StrSql = StrSql & " AND idnro = " & Rs_WF_ft!idnro
            StrSql = StrSql & " AND origen = " & Rs_WF_ft!Origen
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Rs_WF_ft.MoveNext
        Loop
              
'Cierro y libero
    If rs.State = adStateOpen Then rs.Close
    If Rs_WF_ft.State = adStateOpen Then Rs_WF_ft.Close
    Set rs = Nothing
    Set Rs_WF_ft = Nothing
Exit Sub

ME_UFT:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "No se pueden actualizar las entradas fuenta de termino a estado reprocesado."
    Flog.writeline
End Sub

'EAM- Obtiene el país configurado para el cálculo del Medelo de vacaciones.
Public Function Pais_Modelo(nroModelo As Integer) As Integer
 Dim rsPais As New ADODB.Recordset
 '0,"" - Argentina
 '1 - Uruguay
 '2 - Chile
 '3 - Colombia
 '4 - Costa Rica
 
    StrSql = "SELECT confint FROM confper WHERE confnro= " & nroModelo
    OpenRecordset StrSql, rsPais
    
    If Not rsPais.EOF Then
        Pais_Modelo = rsPais!confint
    Else
        Pais_Modelo = 0
    End If
End Function


Public Sub ParametrosGlobales()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que Carga da variables globales.
' Autor      : FGZ
' Fecha      : 15/06/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Firma As New ADODB.Recordset
Dim rs_Fin As New ADODB.Recordset

    Firma_Novedades = False
    Firma_Partes_Extras = False
    Firma_Partes_Turno = False
    Firma_Partes_AsigHor = False
    Firma_Licencias = False
    Firma_RegManuales = False
    Firma_InputFT = False
    Firma_NovLiq = False
    
    FIN_Firma_Novedades = False
    FIN_Firma_Partes_Extras = False
    FIN_Firma_Partes_Turno = False
    FIN_Firma_Partes_AsigHor = False
    FIN_Firma_Licencias = False
    FIN_Firma_RegManuales = False
    FIN_Firma_InputFT = False
    FIN_Firma_NovLiq = False
    
    StrSql = "SELECT * FROM cystipo WHERE cystipact = -1 ORDER BY cystipnro"
    OpenRecordset StrSql, rs_Firma
    Do While Not rs_Firma.EOF
        Select Case rs_Firma!cystipnro
        Case 1: 'Parte de hs extras
            Firma_Partes_Extras = True
            
            StrSql = "SELECT cystipnro FROM cysfincirc"
            StrSql = StrSql & " WHERE cystipnro = " & rs_Firma!cystipnro & " AND upper(userid) = '" & UCase(User_Proceso) & "'"
            OpenRecordset StrSql, rs_Fin
            If Not rs_Fin.EOF Then
                FIN_Firma_Partes_Extras = True
            End If
        Case 4: 'Parte de cambios de turno
            Firma_Partes_Turno = True
        
            StrSql = "SELECT cystipnro FROM cysfincirc"
            StrSql = StrSql & " WHERE cystipnro = " & rs_Firma!cystipnro & " AND upper(userid) = '" & UCase(User_Proceso) & "'"
            OpenRecordset StrSql, rs_Fin
            If Not rs_Fin.EOF Then
                FIN_Firma_Partes_Turno = True
            End If
        Case 5: 'Novedades de Liquidacion
            Firma_NovLiq = True
        
            StrSql = "SELECT cystipnro FROM cysfincirc"
            StrSql = StrSql & " WHERE cystipnro = " & rs_Firma!cystipnro & " AND upper(userid) = '" & UCase(User_Proceso) & "'"
            OpenRecordset StrSql, rs_Fin
            If Not rs_Fin.EOF Then
                FIN_Firma_NovLiq = True
            End If
        
        Case 6: 'Licencias
            Firma_Licencias = True
        
            StrSql = "SELECT cystipnro FROM cysfincirc"
            StrSql = StrSql & " WHERE cystipnro = " & rs_Firma!cystipnro & " AND upper(userid) = '" & UCase(User_Proceso) & "'"
            OpenRecordset StrSql, rs_Fin
            If Not rs_Fin.EOF Then
                FIN_Firma_Licencias = True
            End If
        Case 7: 'Novedades Horarias
            Firma_Novedades = True
        
            StrSql = "SELECT cystipnro FROM cysfincirc"
            StrSql = StrSql & " WHERE cystipnro = " & rs_Firma!cystipnro & " AND upper(userid) = '" & UCase(User_Proceso) & "'"
            OpenRecordset StrSql, rs_Fin
            If Not rs_Fin.EOF Then
                FIN_Firma_Novedades = True
            End If
        Case 17: 'Parte de asignacion horaria
            Firma_Partes_AsigHor = True
        
            StrSql = "SELECT cystipnro FROM cysfincirc"
            StrSql = StrSql & " WHERE cystipnro = " & rs_Firma!cystipnro & " AND upper(userid) = '" & UCase(User_Proceso) & "'"
            OpenRecordset StrSql, rs_Fin
            If Not rs_Fin.EOF Then
                FIN_Firma_Partes_AsigHor = True
            End If
        Case 30: 'Registraciones manuales
            Firma_RegManuales = True
        
            StrSql = "SELECT cystipnro FROM cysfincirc"
            StrSql = StrSql & " WHERE cystipnro = " & rs_Firma!cystipnro & " AND upper(userid) = '" & UCase(User_Proceso) & "'"
            OpenRecordset StrSql, rs_Fin
            If Not rs_Fin.EOF Then
                FIN_Firma_RegManuales = True
            End If
        Case 32: 'entradas fuera de termino (input_FT)
            Firma_InputFT = True
        
            StrSql = "SELECT cystipnro FROM cysfincirc"
            StrSql = StrSql & " WHERE cystipnro = " & rs_Firma!cystipnro & " AND upper(userid) = '" & UCase(User_Proceso) & "'"
            OpenRecordset StrSql, rs_Fin
            If Not rs_Fin.EOF Then
                FIN_Firma_InputFT = True
            End If
        End Select
        rs_Firma.MoveNext
    Loop

If rs_Firma.State = adStateOpen Then rs_Firma.Close
Set rs_Firma = Nothing

If rs_Fin.State = adStateOpen Then rs_Fin.Close
Set rs_Fin = Nothing

End Sub

