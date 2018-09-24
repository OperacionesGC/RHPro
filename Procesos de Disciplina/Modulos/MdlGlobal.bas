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
    ternro As Long
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
Global Parametros As String

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


Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'FGZ - 20/01/2005
Global Etiqueta

Global Tipo_Hora As Long

Global Horario_Movil As Boolean
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


'=============================================================================================
'=============================================================================================
'=============================================================================================

Public Sub GeneraTraza(ternro As Long, Fecha As Date, Desc As String, Optional valor As String = "?")

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
        '   gti_notifhor
        '   gti_notifhor_det
        
        'gti_notifhor
        Texto = " Revisar tablas gti_notifhor y gti_notifhor_det "
        StrSql = "Select gti_notifhor.notifnro FROM gti_notifhor WHERE notifnro = 1"
        OpenRecordset StrSql, rs
        
        'gti_notifhor_det
        Texto = " Revisar tablas gti_notifhor y gti_notifhor_det "
        StrSql = "Select gti_notifhor_det.notifnro FROM gti_notifhor_det WHERE notifnro = 1"
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
    If Version >= "2.16" Then
        'Revisar los campos
        'vacdiascor.venc
        Texto = "Revisar los campos: vacdiascor.venc"
        
        StrSql = "Select venc from vacdiascor where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
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


Public Sub ActualizarFT(ByVal Proceso As Long, ByVal Fecha As Date, ByVal ternro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: sub que actualiza el estado de las entradas fuera de termino contempladas durante el procesamiento.
' Autor      : FGZ
' Fecha      : 14/06/2010
'Parametros  :
'   Proceso:
'           1 Horario Cumplido
'           2 Acumulado Diario
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
    StrSql = StrSql & " WHERE  ternro = " & ternro
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
    StrSql = StrSql & " WHERE (ternro = " & ternro & ") "
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
    StrSql = StrSql & " WHERE (Ternro = " & ternro & ")"
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
        StrSql = StrSql & ID & "," & Tipo & "," & rs!Gnovnro & ")"
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
 
    StrSql = "SELECT confint FROM confper WHERE confnro= " & nroModelo
    OpenRecordset StrSql, rsPais
    
    If Not rsPais.EOF Then
        Pais_Modelo = rsPais!confint
    Else
        Pais_Modelo = 0
    End If
End Function
