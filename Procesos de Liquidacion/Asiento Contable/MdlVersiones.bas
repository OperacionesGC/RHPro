Attribute VB_Name = "MdlVersiones"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "15/08/2005"   'Martin Ferraro - Version Inicial

'Const Version = 1.02
'Const FechaVersion = "22/02/2007"   'Martin Ferraro - Permitir * en los nro de cuenta

'Const Version = 1.03
'Const FechaVersion = "12/05/2007"   'Martin Ferraro - Guardar la descripcion de la linea en el detalle y debe/haber
                                    'para la exportacion alcoa
'Const Version = 1.04
'Const FechaVersion = "11/05/2007"   'Martin Ferraro - En GuardarDetalleAsi se resetean los indicies al final

'Const Version = 1.05
'Const FechaVersion = "13/06/2007 - Custom SMT - Modelo TARJA"   'Fernando Favre - modelo TARJA custom para SMT

'Const Version = 1.06
'Const FechaVersion = "10/07/2007"  'Fernando Favre - Se agrego la proporcionalidad en las cantidades

'Const Version = 1.07
'Const FechaVersion = "29/10/2007"   'Martin Ferraro - En GuardarDetalleAsi se agruparon las lineas del mismo empleado, modelo
                                    '                 proceso, cuenta, tipo origen, origen y proceso (guardado en dlcosto4)
                                    '                 Cuando hay dist contable el detalle asi simpre tenia el 100%
'Const Version = 1.08
'Const FechaVersion = "05/12/2008"   'Martin Ferraro - Movimiento de Golondrinas

'Global Const Version = 1.09
'Global Const FechaVersion = "19/08/2009"   'Encriptacion de string connection
'Global Const UltimaModificacion = "Manuel Lopez"
'Global Const UltimaModificacion1 = "Encriptacion de string connection"

'Global Const Version = "1.10"
'Global Const FechaVersion = "15/07/2010"
'Global Const UltimaModificacion = "Martin"
'Global Const UltimaModificacion1 = "Calculo de cuenta niveladora con 2 decimales"

'Global Const Version = "1.11"
'Global Const FechaVersion = "24/08/2011"
'Global Const UltimaModificacion = "" 'Juan A. Zamarbide
'Global Const UltimaModificacion1 = "" 'Se agregó el Proceso "PRESUPUESTO" dentro de la función Procesar Modelo

'Global Const Version = "1.12"
'Global Const FechaVersion = "03/07/2011"
'Global Const UltimaModificacion = "" 'Deluchi Ezequiel
'Global Const UltimaModificacion1 = " Se modifico la longitud de cuenta "

'Global Const Version = "1.13"
'Global Const FechaVersion = "04/03/2013"
'Global Const UltimaModificacion = "" 'Deluchi Ezequiel - CAS-13764 - H&A - Imputacion Contable
'Global Const UltimaModificacion1 = " Se agrego vigencia a las imputaciones (std) "


'Global Const Version = "1.14"
'Global Const FechaVersion = "19/04/2013"
'Global Const UltimaModificacion = "" 'Margiotta Emanuel - CAS-13764 - H&A - Imputacion Contable
'Global Const UltimaModificacion1 = " Se agrego vigencia a las imputaciones (std) "

'---------------------------------------------------------------------------------------------------------------
' Funcionalidad aun no liberada
'Global Const Version = "2.00"
'Global Const FechaVersion = "28/01/2014"
'Global Const UltimaModificacion = "" 'FGZ - CAS-22808 - SGS - Distribución Contable
'Global Const UltimaModificacion1 = " Nueva forma de Imputacion Por conceptos. sub AcumularXConceptos "
''                                    "EAM 28/01/2014 - Se agrego la funcionalidad para que haga distribucion por acumuladores"
'
'Global Const Version = "2.01"
'Global Const FechaVersion = "26/02/2014"
'Global Const UltimaModificacion = "" 'EAM 26/02/2014 - Se modifico el asiento para que tenga alcance por 3 niveles (Global,Estructura,Individual)
'Global Const UltimaModificacion1 = " Nueva forma de Imputacion Por conceptos. sub AcumularXConceptos "

'Global Const Version = "2.02"
'Global Const FechaVersion = "29/04/2014"
'Global Const UltimaModificacion = "" 'EAM 26/02/2014 - Se modifico el asiento para que tenga alcance por 3 niveles (Global,Estructura,Individual)
'Global Const UltimaModificacion1 = " Se corrigio el asiento para el modelo AcumularXConceptos ya que el estaba guardando mal los datos en detalle_asiaux"
'
'Global Const Version = "2.03"
'Global Const FechaVersion = "19/05/2014"
'Global Const UltimaModificacion = " 'Miriam Ruiz- CAS-13713 - MONRESA - Gestion Presupuestaria-"
'Global Const UltimaModificacion1 = " Al buscar el proceso modificó para que tuviera en cuenta si era simulación"
'

'Global Const Version = "2.04"
'Global Const FechaVersion = "30/05/2014"
'Global Const UltimaModificacion = " 'Margiotta, Emanuel - CAS-22808 - SGS-"
'Global Const UltimaModificacion1 = "Se modifico la distribucin estandar por concepto cuando tenía un concepto de pago y otro de descuento no distribuia éste último"

'Global Const Version = "2.05" versión No liberada 2.05
'Global Const FechaVersion = "28/08/2014"
'Global Const UltimaModificacion =
'Global Const UltimaModificacion1 = versión No liberada 2.05

'Global Const Version = "2.06"
'Global Const FechaVersion = "17/12/2014"
'Global Const UltimaModificacion = " 'Margiotta, Emanuel - CAS-22808 - SGS- Prueba"
'Global Const UltimaModificacion1 = "Se modifico el asieto cuando distribuye por acumulador para que tome la apertura estandar y tambien la distribuya aunque sea negativa"
'                           Se agrego una nueva funcion de balancelo para la cuenta niveladora que tenga en cuenta los 4 decimales ya que con el modelo estandar no balancea."

'Global Const Version = "2.07"
'Global Const FechaVersion = "30/12/2014"
'Global Const UltimaModificacion = " Margiotta, Emanuel - Moro, Lisandro - CAS-28361 - SGS - Bug en generar asiento contable"
'Global Const UltimaModificacion1 = "Se modifico el asiento cuando distribuye por acumulador y el monto sea 0 "

'Global Const Version = "2.08"
'Global Const FechaVersion = "05/01/2015"
'Global Const UltimaModificacion = " Moro, Lisandro - CAS-19483 - BANCO INDUSTRIAL - BUG EN ITEM COMPROBANTE [Entrega 2] "
'Global Const UltimaModificacion1 = " Llamo al BalanceModelo en el else al ProcesoGeneral "

'Global Const Version = "2.09"
'Global Const FechaVersion = "27/04/2015"
'Global Const UltimaModificacion = " Carmen Quinteo - CAS-30312 - SGS - BUG EN ASIENTO CONTABLE "
'Global Const UltimaModificacion1 = " Se aumento el tamaño del arreglo vec_imputacion2 "

'Global Const Version = "2.10"
'Global Const FechaVersion = "17/06/2015"
'Global Const UltimaModificacion = " LM - CAS-31414 - SGS - Bug en generación de asiento"
'Global Const UltimaModificacion1 = " Se aumento el tamaño del arreglo vec_imputacion "

'Global Const Version = "2.11"
'Global Const FechaVersion = "20/07/2015"
'Global Const UltimaModificacion = " Carmen Quintero - CAS-30759 - SGS - CUSTOM REPORTE HEADCOUNT"
'Global Const UltimaModificacion1 = " Se actualiza el campo linaD_H ubicado en detalle_asi, con el mismo valor que fue registrado en la tabla linea_asi "

''Global Const Version = "2.12"
''Global Const FechaVersion = "31/07/2015"
''Global Const UltimaModificacion = " Margiotta, Emanuel - CAS-31970 - SYKES El Salvador - LIQ - Error en distribucion contable"
''Global Const UltimaModificacion1 = " Se modifico en el asiento estandar el alcance (periodo desde y hasta) el inner del query de la fecha hasta por un left"

'Global Const Version = "2.13"
'Global Const FechaVersion = "13/10/2015"
'Global Const UltimaModificacion = " CAS-31519 & CAS-33398 & CAS-30312 " ' FGZ - CAS-31519 - UTDT - Bug en distribución contable - Se solucionó un error en la distribución de conceptos por novedad y su correcta distribución en los asientos cuando hereda de un acumulador.
'Global Const UltimaModificacion1 = " " ' LM  - CAS-33398 - BAPRO - Bug en generación de asiento - Integers a longs

Global Const Version = "2.14"
Global Const FechaVersion = "14/11/2015"
Global Const UltimaModificacion = " Sebastian Stremel CAS-33601 - RH Pro ( Producto ) - Peru – Item Contables"
Global Const UltimaModificacion1 = " funcion armarCuenta se busca el documento con tipodocu_pais, si no encuentra hace la logica anterior."

'
'---------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------

Public Function ValidarV(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que determina si el proceso esta en condiciones de ejecutarse.
' Autor      : FGZ
' Fecha      : 08/01/2014
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True

Select Case TipoProceso
Case 6: 'Asiento Contable
    
    If Version >= "2.00" Then
        'Estas tablas tienen que ver con el liquidador en primera medida aunque sin esas tablas el asiento no puede distribuir por concepto
        
        'Tabla nueva
        'CREATE TABLE [dbo].[nov_dist](
        '    [nedistnro] [int] IDENTITY(1,1) NOT NULL,
        '    [novnro] [int] NOT NULL,                -- FK (novemp o novaju)
        '    [auto] [smallint] NOT NULL default 0,
        '    [tiponov] [int] NOT NULL default 1,     -- {1 novemp, 2 novaju}
        '    [concnro] [int] NOT NULL,               -- FK {concepto}
        '    [tpanro] [int] NULL,                    -- FK {parametro del concepto} -- no se si es FK porque si es novaju ==> va 0
        '    [Masinro] Not [Int], --FK(mod_asiento)
        '    [tenro] [int] NOT NULL,                 -- FK (estructura) ---- no se si es FK porque necesitamos ponerle 0 cuando no tiene distr
        '    [Estrnro] Not [Int], --FK(estructura)
        '    [tenro2] [int] NULL,                    -- FK (estructura)
        '    [estrnro2] [int] NULL,                  -- FK (estructura)
        '    [tenro3] [int] NULL,                    -- FK (estructura)
        '    [estrnro3] [int] NULL                   -- FK (estructura)
        ') ON [PRIMARY]
        'GO


        'CREATE TABLE [dbo].[concepto_dist](
        '    [Ternro] Not [Int], --FK(Tercero)
        '    [ConcNro] Not [Int], --FK(Concepto)
        '    [pronro] Not [Int], --FK(Proceso)
        '    [Masinro] Not [Int], --FK(mod_asiento)
        '    [tenro] [int] NULL,         -- FK (estructura)
        '    [estrnro] [int] NULL,       -- FK (estructura)
        '    [tenro2] [int] NULL,        -- FK (estructura)
        '    [estrnro2] [int] NULL,      -- FK (estructura)
        '    [tenro3] [int] NULL,        -- FK (estructura)
        '    [estrnro3] [int] NULL,      -- FK (estructura)
        '    [porcentaje] [decimal](5, 2) NOT NULL Default(100),
        '    [Monto] Not [decimal](19, 4)
        ') ON [PRIMARY]
        'GO
        
        Texto = "Revisar que exista y tenga permisos la tabla nov_dist "
        StrSql = "Select * FROM nov_dist WHERE nedistnro = 1"
        OpenRecordset StrSql, rs
        
        Texto = "Revisar que exista y tenga permisos la tabla concepto_dist "
        StrSql = "Select * FROM concepto_dist WHERE Ternro = 1"
        OpenRecordset StrSql, rs
        
        
        'CREATE TABLE [dbo].[imputacion_conc](
        '    [concnro] [int] NOT NULL,
        '    [tipo_dist] [int] NOT NULL,
        '    [concnro_depende] [int] NULL
        ') ON [PRIMARY]
        'GO
        Texto = "Revisar que exista y tenga permisos la tabla imputacion_conc "
        StrSql = "Select * FROM imputacion_conc WHERE concnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If
    
   
    
    




'Cambio en busqueda de embargos bus_embargos
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

