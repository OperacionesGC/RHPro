Attribute VB_Name = "mdlValidarBD"
Option Explicit
'*********************************************************************************************************************
'***********************************************                       ***********************************************
'*********************************************** -- CONSIDERACIONES -- ***********************************************
'***********************************************                       ***********************************************
'*********************************************************************************************************************
'1 - SI NO EXISTE EL MODULO DEL PAIS, SE DEBE CREAR --> NOMENCLATURA: mdlPed_Nombrepais
'2 - LA POLITICA 1515 ES OBLIGATORIA | PERMITE UTILIZAR MÁS DE UN MODELO (SOLO COMO INDIVIDUAL O ESTRUCTURA)
'3 - SI SE AGREGA UN MODELO NUEVO - AGREGAR EN CASE politica1515()
'4 - VALIDAR VERSIONES EN FUNCION -->  ValidarVBD()
'5 - AGREGAR ETIQUETAS CON MULTILENGUAJE (ANTES DE CREAR UNA ETIQUETA VERIFICAR QUE NO EXISTA UNA SIMILAR)

'En todos los procedimientos de calculo de dias correspondientes se deben validar las políticas:
'-> 1515 -
'-> 1510 -

' MODELOS DE PAISES/CUSTOM (Al agregar Nuevo modelo verificar fuentes ASP de vacaciones en shared/inc/vacaciones_XXX.inc)
' y en politica1515()
'0 - Argentina | 1 - Uruguay | 2 - Chile | 3 - Colombia | 4 - Costa Rica 5 - Portugal | 6 - Paraguay | 7 - Perú

'****************************************                              ******************************************
'**************************************** -- ULT. VERSION AL INICIO -- ******************************************
'****************************************                              ******************************************

'Global Const Version = "13.00"
'Global Const FechaVersion = "15/11/2012" 'Gonzalez Nicolás - Version inicial R4

'Global Const Version = "13.01"
'Global Const FechaVersion = "15/11/2012" 'Gonzalez Nicolás - Version inicial R4
'30/03/2015 - EAM- CAS-16645 - Se ajusto el proceso para argentina version r4
'Global Const Version = "13.02"
'Global Const FechaVersion = "18/12/2015" 'Gonzalez Nicolás - CAS-34211 - VISION PY - Pedido de vacaciones - Se calcula a 12 meses de la fecha de cause

Global Const Version = "13.03"
Global Const FechaVersion = "16/02/2016  - MDZ CAS-33780 - Se ajusto el proceso para operara con periodos de GIV R4"

                                         

'---------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------FIN VERSIONES ------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ValidarVBD(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer, ByVal codPais As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : Gonzalez Nicolás
' Fecha      : 08/11/2012
' Modificado :
'            :
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True
' CODIGOS DE PAISES
 '0 - Argentina
 '1 - Uruguay
 '2 - Chile
 '3 - Colombia
 '4 - Costa Rica
 '5 - Portugal
 '6 - Paraguay
 '7 - Perú
 
  If Version >= "2.06" Then
    'Revisar los campos
    'vacdiascor.venc
    Texto = "Revisar los campos: vacdiascor.venc"
        
    StrSql = "Select venc from vacdiascor where ternro = 1"
    OpenRecordset StrSql, rs

    V = True
    End If
 

    If Version >= "13.00" Then
        'dias correspondientes
        Texto = "Revisar los campos: "
        Texto = Texto & " vacacion.alcannivel"
        StrSql = "Select vacacion.alcannivel FROM vacacion WHERE alcannivel = 1"
        OpenRecordset StrSql, rs
        V = True
        
        Texto = "Revisar los campos: "
        Texto = Texto & " vacdiascortipo.tdnro,vacdiascortipo.progval"
        StrSql = "Select vacdiascortipo.tdnro,vacdiascortipo.progval FROM vacdiascortipo WHERE tdnro = 1"
        OpenRecordset StrSql, rs
        V = True
        
        Texto = "Revisar Existencia de Tabla vac_alcan. Campos: "
        Texto = Texto & " vac_alcan.vacnro,vac_alcan.vacfecdesde,vac_alcan.vacfechasta,vac_alcan.alcannivel,vac_alcan.origen,vac_alcan.vacestado"
        StrSql = "Select vac_alcan.vacnro,vac_alcan.vacfecdesde,vac_alcan.vacfechasta,vac_alcan.alcannivel,vac_alcan.origen,vac_alcan.vacestado FROM vac_alcan WHERE vacnro = 0"
        OpenRecordset StrSql, rs
        V = True
    End If


ValidarVBD = V
    
If V = True Then
    Texto = "Versión correcta"
    Exit Function
End If




ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & "Revisar estructura de Base de Datos para el modelo N°: " & codPais
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function
