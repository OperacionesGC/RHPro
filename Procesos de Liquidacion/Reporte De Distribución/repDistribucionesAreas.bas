Attribute VB_Name = "repDistribucionesAreas"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "05/05/2011" 'Manterola Maria Magdalena
'Global Const UltimaModificacion = "Inicial"

'Global Const Version = "1.01"
'Global Const FechaModificacion = "26/05/2011" 'Manterola Maria Magdalena
'Global Const UltimaModificacion = "Se agregó la inicialización de estrConfrep en 0, para cada uno de sus elementos"

'Global Const Version = "1.02"
'Global Const FechaModificacion = "13/06/2011" 'Manterola Maria Magdalena
'Global Const UltimaModificacion = "Se guardan los empleados en una lista solo si no se seleccionaron todos los empleados"

'Global Const Version = "1.03"
'Global Const FechaModificacion = "26/07/2011" 'Manterola Maria Magdalena
'Global Const UltimaModificacion = "Se agregó distinción entre Mano de Obra Productiva y No Productiva"

'Global Const Version = "1.04"
'Global Const FechaModificacion = "12/09/2011" 'Manterola Maria Magdalena
'Global Const UltimaModificacion = "Se corrigió la forma de contabilizar los empleados"

'Global Const Version = "1.05"
'Global Const FechaModificacion = "27/12/2011" 'Manterola Maria Magdalena
'Global Const UltimaModificacion = "Se estandarizo el reporte"

'Global Const Version = "1.06"
'Global Const FechaModificacion = "08/08/2012" 'Manterola Maria Magdalena
'Global Const UltimaModificacion = "Se corrigió el error en la generación de las listas, para que calcule correctamente los empleados."

Global Const Version = "1.07"
Global Const FechaModificacion = "10/09/2013" 'Deluchi Ezequiel
Global Const UltimaModificacion = "Correcion en armado de listaagencia y listaPropios, cuando no hay datos."
'-------------------------------------------------------------------------------

Dim fs, f
'Global Flog

Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

'LAS LISTAS DE LOS EMPLEADOS
'LE PONGO LONG = 13 PARA ESTANDARIZARLO CON 10000 EMPLEADOS
'10000 / 800 = 12.5
Dim Listas(13) As String
Dim Longitud As Integer
Dim HastaTodos As Boolean

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global Formato As Integer

Dim repboldesnro As Long
Dim bpronro As Long

Dim Lista_Pro As String
Dim tipoprocesos As Integer

Dim COAC_FONDO As String
Dim COAC_TIPO As String
Dim EST_BCO_FONDO As String
Dim TD_IERIC As Long
Dim TD_RNIC As Long


Private Sub Main()

    Dim NombreArchivo As String
    Dim directorio As String
    Dim CArchivos
    Dim archivo
    Dim Folder
    Dim strCmdLine As String
    Dim Nombre_Arch As String
    
    
    Dim objRs As New ADODB.Recordset
    Dim objRs2 As New ADODB.Recordset
    Dim rsPeriodos As New ADODB.Recordset

    Dim tipoDepuracion
    Dim historico As Boolean
    Dim param
    Dim listapronro
    Dim proNro
    Dim Ternro
    Dim arrpronro
    Dim Periodos
    Dim rsEmpl As New ADODB.Recordset
    Dim I
    Dim totalEmpleados
    Dim cantRegistros
    Dim PID As String
    Dim tituloReporte As String
    Dim Parametros As String
    Dim ArrParametros
    Dim strTempo As String
    Dim Orden
    Dim bprcparam As String
    Dim iduser
    
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProceso = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProceso = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    Call ObtenerConfiguracionRegional
    
    TiempoInicialProceso = GetTickCount
    
    Nombre_Arch = PathFLog & "Rep_Distribucion_Areas" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
     ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    On Error GoTo ME_Main
    
    HuboErrores = False
    
    Flog.writeline "Inicio Proceso de Reporte Distribución De Areas : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
        'Obtengo los parametros del proceso
        Parametros = objRs!bprcparam
        iduser = objRs!iduser
        objRs.Close
        Set objRs = Nothing
        Call Calcular(NroProcesoBatch, Parametros, iduser)
        
        
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
ME_Main:

        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans

End Sub
Public Sub Calcular(ByVal bpronro As Long, ByVal Parametros As String, ByVal iduser As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : Manterola Maria Magdalena
' Fecha      : 06/05/2011
' Modificado : 23/06/2011 - Gonzalez Nicolás
'                           Se modificó el sql cuando tipoEmpleados = 1 (Se agregó UNION)
                 
' --------------------------------------------------------------------------------------------
Dim cantNiv As Integer
Dim VecActual(5) As Integer

Dim rs_Periodo As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset
Dim rsAux As New ADODB.Recordset
Dim rsAgencia As New ADODB.Recordset
Dim rsPropios As New ADODB.Recordset

Dim CantPeriodos As Integer

Dim cantPropios As Integer
Dim cantAgencia As Integer
Dim total As Integer

Dim cantEmpleadosPropios As Integer
Dim cantEmpleadosAgencia As Integer

Dim pos1 As Integer
Dim pos2 As Integer
Dim legdesde As Long

Dim leghasta As Long
Dim estado As Integer
Dim tipoEmpleados As Integer
Dim perdesde As Integer
Dim perhasta As Integer

Dim estadosproc As String
Dim tenro1 As String
Dim estrnro1 As String
Dim tenro2 As String
Dim estrnro2 As String
Dim tenro3 As String
Dim estrnro3 As String
Dim tenro4 As String
Dim estrnro4 As String
Dim fecestr As String
Dim Orden As String
Dim estadoprocdes As String
Dim ordenado As String

Dim pliqnroMenor As Integer
Dim pliqAnioMenor As Integer
Dim pliqMesMenor As Integer

Dim pliqnroMayor As Integer
Dim pliqAnioMayor As Integer
Dim pliqMesMayor As Integer

Dim Empnro As Long

Dim filas As Integer

Dim k As Integer
Dim I As Integer

Dim primero As Boolean
Dim estrConfrep(5) As Integer
Dim estr1Desc As String
Dim estrConfrep2 As Integer
Dim estr2Desc As String
Dim estrConfrep3 As Integer
Dim estr3Desc As String
Dim estrConfrep4 As Integer
Dim estr4Desc As String
Dim estrConfrep5 As Integer
Dim estr5Desc As String

Dim longi As Integer
Dim listaEmp2 As String
Dim listaaux
Dim StrSql2 As String
Dim StrSql3 As String


Dim tipoCalcCol(5) As String
Dim NroCA(5)
Dim DescCol(5) As String
Dim colUltValor(5) As Double
Dim Matriz(1000, 22) As Variant
Dim DescEst(5) As String
Dim DescTE(5) As String
Dim DesEstActual(5) As String
Dim Vec(5) As Integer
    
Dim listaEmp As String
Dim listaAgencia As String
Dim listaPropios As String
Dim listPer As String

Dim cantColumnasFinales As Integer

Dim cantProductivo As Integer
Dim cantNoProductivo As Integer
Dim MDOProdConfigurada As Boolean
Dim MDONProdConfigurada As Boolean

Dim cant As Integer

Dim Fecha As String
Dim hora As String

MDOProdConfigurada = False
MDONProdConfigurada = False


'FECHA Y HORA ACTUAL
Fecha = ConvFecha(Date)
hora = Format(Now, "hh:mm:ss")

' PARAMETROS------------------------------------------------------------------------------
' Levanto cada parametro por separado, el separador de parametros es "."
' El formato es:
' LEGDESDE.LEGHASTA.ESTADOEMPLEADO.TIPOEMPLEADO.EMPRESA.PERIODODESDE.PERIODOHASTA.PROCESOS.
' ESTADOPROC.TENROEST1.ESTRNROEST1.TENROEST2.ESTRNROEST2.TENROEST3.ESTRNROEST3.FECHACORTE.
'----------------------------------------------------------------------------------------
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
    
    If Len(Parametros) >= 1 Then
    
        'LEG DESDE
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        legdesde = CLng(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro legdesde = " & legdesde
        
        'LEG HASTA
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        If Mid(Parametros, pos1, pos2) = "9999999999" Then
            HastaTodos = True
            Flog.writeline "Parametro leghasta = TODOS "
        Else
            HastaTodos = False
            leghasta = CLng(Mid(Parametros, pos1, pos2))
            Flog.writeline "Parametro leghasta = " & leghasta
        End If
        
        'ESTADO EMPLEADO
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        estado = CInt(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro estado = " & estado
        
        'TIPO EMPLEADO (TODOS/AGENCIA/PROPIOS)
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        tipoEmpleados = CInt(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro tipoEmpleados = " & tipoEmpleados
        
        'EMPRESA
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        Empnro = CLng(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro Empresa(Nro) = " & Empnro
        
        'PERIODO DESDE
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        perdesde = CInt(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro perdesde = " & perdesde
        
        'PERIODO HASTA
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        perhasta = CInt(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro perhasta = " & perhasta
        
        'LISTA DE PROCESOS
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        Lista_Pro = Mid(Parametros, pos1, pos2)
        ' esta lista tiene los nro de procesos separados por comas
        Flog.writeline "Parametro Lista_Pro = " & Lista_Pro
       
        'ESTADO DE PROCESOS
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        estadosproc = Mid(Parametros, pos1, pos2)
        Select Case estadosproc
            Case "":
                estadosproc = 0
                estadoprocdes = "Todos"
                Flog.writeline "Parametro Estado de Procesos = Todos"
            Case 1:
                estadoprocdes = "Liquidado"
                Flog.writeline "Parametro Estado de Procesos = Liquidado"
            Case 2:
                estadoprocdes = "Aprobado Provisorio"
                Flog.writeline "Parametro Estado de Procesos = Aprobado Provisorio"
            Case 3:
                estadoprocdes = "Aprobado Definitivo"
                Flog.writeline "Parametro Estado de Procesos = Aprobado Definitivo"
        End Select
        
        'TENRO ESTRUCTURA 1
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        tenro1 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro tenro1 = " & tenro1
       
        'ESTRNRO ESTRUCTURA 1
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        estrnro1 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro estrnro1 = " & estrnro1
        
        'TENRO ESTRUCTURA 2
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        tenro2 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro tenro2 = " & tenro2
       
        'ESTRNRO ESTRUCTURA 2
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        estrnro2 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro estrnro2 = " & estrnro2
        
        'TENRO ESTRUCTURA 3
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        tenro3 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro tenro3 = " & tenro3
       
        'ESTRNRO ESTRUCTURA 3
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        estrnro3 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro estrnro3 = " & estrnro3
        
        'FECHA CORTE
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = Len(Parametros)
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        fecestr = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro fecestr = " & fecestr
       
                        
    End If
Else
    Flog.writeline "parametros nulos"
End If
Flog.writeline "terminó de levantar los parametros"

'TERMINO DE LEVANTAR LOS PARAMETROS-----------------------------------------------
       
'EMPIEZA EL PROCESO---------------------------------------------------------------
    
'Me quedo con el menor periodo de liquidación elegido
StrSql = "SELECT pliqnro,pliqmes,pliqanio "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " WHERE pliqnro = '" & perdesde & "'"
OpenRecordset StrSql, rs_Periodo
    
If Not rs_Periodo.EOF Then
    pliqnroMenor = rs_Periodo!pliqnro
    pliqAnioMenor = rs_Periodo!pliqanio
    pliqMesMenor = rs_Periodo!pliqmes
End If
rs_Periodo.Close

'Ahora busco el mayor periodo de liquidación elegido
StrSql = "SELECT pliqnro,pliqmes,pliqanio "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " WHERE pliqnro = '" & perhasta & "'"
OpenRecordset StrSql, rs_Periodo
    
If Not rs_Periodo.EOF Then
    pliqnroMayor = rs_Periodo!pliqnro
    pliqAnioMayor = rs_Periodo!pliqanio
    pliqMesMayor = rs_Periodo!pliqmes
End If
rs_Periodo.Close


'Traigo todos los periodos de liquidación válidos para la selección del usuario
StrSql = "SELECT pliqnro,pliqmes,pliqanio,pliqdesde,pliqhasta,pliqdesc "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " WHERE pliqanio >= '" & pliqAnioMenor & "'"
StrSql = StrSql & " AND pliqmes >= (CASE WHEN pliqanio = " & pliqAnioMenor & " THEN " & pliqMesMenor & " ELSE 1 END)"
StrSql = StrSql & " AND pliqanio >= '" & pliqAnioMenor & "'"
StrSql = StrSql & " AND pliqmes < = (CASE WHEN pliqmes = " & pliqAnioMayor & " THEN " & pliqMesMayor & " ELSE 12 END)"
StrSql = StrSql & " ORDER BY pliqanio DESC "

'Flog.writeline StrSql

OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If

listPer = ""

rs_Periodo.MoveFirst
Do
    If listPer = "" Then
        listPer = rs_Periodo!pliqnro
    Else
        If Len(listPer) = 100 Then
            listPer = listPer & ") OR proceso.pliqnro IN (" & rs_Periodo!pliqnro
        Else
            listPer = listPer & "," & rs_Periodo!pliqnro
        End If
    End If
    
    rs_Periodo.MoveNext
Loop Until rs_Periodo.EOF
    
rs_Periodo.Close


'26/05/2011 - Manterola Maria Magdalena
'Se agregó la inicialización del arreglo estrConfrep en 0
k = 0
Do Until k = 5
    estrConfrep(k) = 0
    k = k + 1
Loop


'CONFREP---------------------------------------------------
Flog.writeline "Configuración del Reporte"
StrSql = "SELECT * FROM confrep WHERE repnro = 347 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Confrep

k = 1
If Not rs_Confrep.EOF Then
    rs_Confrep.MoveFirst
    cantColumnasFinales = 0
    Do Until rs_Confrep.EOF
        Select Case UCase(rs_Confrep!conftipo)
            Case "NIV":
                'Tengo que ver cuantos niveles de apertura estan configurados
                cantNiv = rs_Confrep!confval
                Flog.writeline "Cantidad de niveles de apertura:" & cantNiv
                If (cantNiv > 5) Then
                    Exit Sub
                End If
            Case "TE1":
                'Primera Estructura
                estrConfrep(0) = rs_Confrep!confval
                Flog.writeline "Primer Tipo de Estructura Configurado:" & estrConfrep(0)
            Case "TE2":
                'Segunda Estructura
                estrConfrep(1) = rs_Confrep!confval
                Flog.writeline "Segundo Tipo de Estructura Configurado:" & estrConfrep(1)
            Case "TE3":
                'Tercer Estructura
                estrConfrep(2) = rs_Confrep!confval
                Flog.writeline "Tercer Tipo de Estructura Configurado:" & estrConfrep(2)
            Case "TE4":
                'Cuarta Estructura
                estrConfrep(3) = rs_Confrep!confval
                Flog.writeline "Cuarto Tipo de Estructura Configurado:" & estrConfrep(3)
            Case "AC", "CO", "CCO", "CAC":
                'Columnas Finales
                cantColumnasFinales = cantColumnasFinales + 1
                Select Case rs_Confrep!confval
                    Case "1":
                        'Primera Columna
                        tipoCalcCol(0) = UCase(rs_Confrep!conftipo)
                        NroCA(0) = rs_Confrep!confval2
                        DescCol(0) = rs_Confrep!confetiq
                    Case "2":
                        'Segunda Columna
                        tipoCalcCol(1) = UCase(rs_Confrep!conftipo)
                        NroCA(1) = rs_Confrep!confval2
                        DescCol(1) = rs_Confrep!confetiq
                    Case "3":
                        'Tercera Columna
                        tipoCalcCol(2) = UCase(rs_Confrep!conftipo)
                        NroCA(2) = rs_Confrep!confval2
                        DescCol(2) = rs_Confrep!confetiq
                    Case "4":
                        'Cuarta Columna
                        tipoCalcCol(3) = UCase(rs_Confrep!conftipo)
                        NroCA(3) = rs_Confrep!confval2
                        DescCol(3) = rs_Confrep!confetiq
                    Case "5":
                        'Quinta Columna
                        tipoCalcCol(4) = UCase(rs_Confrep!conftipo)
                        NroCA(4) = rs_Confrep!confval2
                        DescCol(4) = rs_Confrep!confetiq
                End Select
            Case "MOP":
                'Mano de Obra Productiva
                MDOProdConfigurada = True
            Case "MON":
                'Mano de Obra No Productiva
                MDONProdConfigurada = True
        End Select
        rs_Confrep.MoveNext
    Loop
   
    rs_Confrep.MoveFirst
Else
    Flog.writeline "No se encontró la configuración del reporte. Abortando"
    Exit Sub
End If
'TERMINO CONFREP-------------------------------------------------------------
Dim elEstado
If estado = -1 Then
    'Activo
    If Not HastaTodos Then
        elEstado = " empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec <=" & ConvFecha(fecestr) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fecestr) & ")) ) AND (empleg >= " & legdesde & ") AND (empleg <= " & leghasta & ") "
    Else
        elEstado = " empleado.ternro in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec <=" & ConvFecha(fecestr) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fecestr) & ")) ) AND (empleg >= " & legdesde & ") "
    End If
ElseIf estado = 0 Then
    'Inactivo
    If Not HastaTodos Then
        elEstado = " empleado.ternro not in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec <=" & ConvFecha(fecestr) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fecestr) & ")) ) AND (empleg >= " & legdesde & ") AND (empleg <= " & leghasta & ") "
    Else
        elEstado = " empleado.ternro not in (SELECT empleado from fases WHERE fases.empleado = empleado.ternro AND (fases.altfec <=" & ConvFecha(fecestr) & " AND (fases.bajfec is null or fases.bajfec>=" & ConvFecha(fecestr) & ")) ) AND (empleg >= " & legdesde & ") "
    End If
Else
    'Ambos
    If Not HastaTodos Then
        elEstado = "(0 = 0) And (empleg >= " & legdesde & ") And (empleg <= " & leghasta & ")"
    Else
        elEstado = "(0 = 0) And (empleg >= " & legdesde & ") "
    End If
End If

'------------------------------------------
'SELECCION DE EMPLEADOS SEGUN FILTRO
'------------------------------------------

StrSql2 = ""
StrSql3 = ""


'Primero filtro por Empresa si se selecciono una empresa
StrSql = "SELECT DISTINCT empleado.ternro"
StrSql = StrSql & " FROM empleado  "

'***************ACLARACION!!!!! VER QUE SE DEFINE DPS DEL MAIL ***********
'Ahora lo dejo con left pero usar lo que esta comentado si se considera que los de agencia tienen empresa
'If Empnro <> 0 Then
'    StrSql = StrSql & " INNER JOIN his_estructura estEmpresa ON empleado.ternro = estEmpresa.ternro AND estEmpresa.tenro = 10"
'    StrSql = StrSql & " AND estEmpresa.estrnro = " & Empnro
'Else
'    StrSql = StrSql & " LEFT JOIN his_estructura estEmpresa ON empleado.ternro = estEmpresa.ternro AND estEmpresa.tenro = 10"
'End If
'StrSql = StrSql & " AND (estEmpresa.htetdesde<=" & ConvFecha(fecestr) & " AND (estEmpresa.htethasta is null or estEmpresa.htethasta>=" & ConvFecha(fecestr) & "))"
'************************************************************************

If estrConfrep(0) = 10 Or estrConfrep(1) = 10 Or estrConfrep(2) = 10 Or estrConfrep(3) = 10 Then
    StrSql = StrSql & " INNER JOIN his_estructura estEmpresa ON empleado.ternro = estEmpresa.ternro AND estEmpresa.tenro = 10"
Else
    StrSql = StrSql & " LEFT JOIN his_estructura estEmpresa ON empleado.ternro = estEmpresa.ternro AND estEmpresa.tenro = 10"
End If
If Empnro <> 0 Then
    StrSql = StrSql & " AND estEmpresa.estrnro = " & Empnro
End If

StrSql = StrSql & " AND (estEmpresa.htetdesde<=" & ConvFecha(fecestr) & " AND (estEmpresa.htethasta is null or estEmpresa.htethasta>=" & ConvFecha(fecestr) & "))"

'--> ARMO LA LISTA AGENCIA
        listaAgencia = "0"
        StrSql3 = " SELECT DISTINCT empleado.ternro from empleado "
        StrSql3 = StrSql3 & " INNER JOIN his_estructura agencia ON agencia.ternro = empleado.ternro "
        StrSql3 = StrSql3 & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(fecestr)
        StrSql3 = StrSql3 & " AND (agencia.htethasta is null or agencia.htethasta>=" & ConvFecha(fecestr) & ")) "
        
        Flog.writeline
        Flog.writeline "CONSULTA LISTA AGENCIA!!!" & StrSql3
        Flog.writeline
        
        OpenRecordset StrSql3, rsAgencia
        If rsAgencia.EOF Then
            Flog.writeline "No hay empleados con estructura agencia asignada."
        End If
        Do Until rsAgencia.EOF
            listaAgencia = listaAgencia & "," & rsAgencia!Ternro
            rsAgencia.MoveNext
        Loop
        
        '--> ARMO LA LISTA PROPIOS
        listaPropios = "0"
        StrSql3 = " SELECT DISTINCT empleado.ternro from empleado "
        StrSql3 = StrSql3 & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
        StrSql3 = StrSql3 & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
        If listPer <> "" Then
            StrSql3 = StrSql3 & " AND proceso.pliqnro IN (" & listPer & ")"
        End If
        If listaAgencia <> "" Then
            StrSql3 = StrSql3 & " WHERE empleado.ternro NOT IN (" & listaAgencia & ")"
        End If
       
        Flog.writeline
        Flog.writeline "CONSULTA LISTA PROPIOS!!!" & StrSql3
        Flog.writeline
        
        OpenRecordset StrSql3, rsPropios
        
        Do Until rsPropios.EOF
            listaPropios = listaPropios & "," & rsPropios!Ternro
            rsPropios.MoveNext
        Loop
        
Select Case tipoEmpleados
    Case 1:
        'AMBOS
        StrSql2 = "AND (empleado.ternro IN (" & listaPropios & ")"
        StrSql2 = StrSql2 & " OR empleado.ternro IN (" & listaAgencia & "))"
        
    Case 2:
        'PROPIOS
        'Busco que esten liquidados --> ACLARACION DEL MAIL
        'StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
        'StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
        'If listPer <> "" Then
        '    StrSql = StrSql & " AND proceso.pliqnro IN (" & listPer & ")"
        'End If
        
        'StrSql2 = " AND empleado.ternro not in (SELECT ternro from his_estructura agencia "
        'StrSql2 = StrSql2 & " WHERE agencia.tenro= 28 AND (agencia.htetdesde<=" & ConvFecha(fecestr)
        'StrSql2 = StrSql2 & " AND (agencia.htethasta is null or agencia.htethasta>=" & ConvFecha(fecestr) & " )) )"
        StrSql2 = "AND empleado.ternro IN (" & listaPropios & ")"

    Case 3:
        'AGENCIA
        StrSql2 = "AND empleado.ternro IN (" & listaAgencia & ")"
        'StrSql2 = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
        'StrSql2 = StrSql2 & " WHERE agencia.tenro=28 "
        'StrSql2 = StrSql2 & " AND agencia.htetdesde<=" & ConvFecha(fecestr)
        'StrSql2 = StrSql2 & " AND (agencia.htethasta is null or agencia.htethasta>=" & ConvFecha(fecestr) & ") )"
        
End Select

'____________________________________
'SOLO CUANDO SE SELECCIONAN 3 NIVELES
If tenro3 <> "" And tenro3 <> "0" Then
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
    '_______________________________
    'SE LE ASIGNA 1 VALOR AL NIVEL 1
    If estrnro1 <> "" And estrnro1 <> "0" Then
        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
    End If
    StrSql = StrSql & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecestr) & "))"
    
    '_______________________________
    'SE LE ASIGNA 1 VALOR AL NIVEL 2
    If estrnro2 <> "" And estrnro2 <> "0" Then
        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
    End If
    
    StrSql = StrSql & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3
    StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(fecestr) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(fecestr) & "))"
    
    '_______________________________
    'SE LE ASIGNA 1 VALOR AL NIVEL 3
    If estrnro3 <> "" And estrnro3 <> "0" Then
        StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
    End If
    
    StrSql = StrSql & " INNER JOIN estructura estructura3 ON estructura3.estrnro=estact3.estrnro "
    StrSql = StrSql & " WHERE " & elEstado & StrSql2
    'StrSql = StrSql & " ORDER BY estructura1.tenro,estructura1.estrdabr,estructura2.tenro,estructura2.estrdabr,estructura3.tenro,estructura3.estrdabr"
'_______________________________________
'CUANDO SE SELECCIONA HASTA EL 2DO NIVEL
ElseIf tenro2 <> "" And tenro2 <> "0" Then
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
        
        If estrnro1 <> "" And estrnro1 <> "0" Then
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        
        StrSql = StrSql & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
        StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecestr) & "))"
        
        If estrnro2 <> "" And estrnro2 <> "0" Then
            StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
        End If
        
        StrSql = StrSql & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
        StrSql = StrSql & " WHERE " & elEstado & StrSql2
        'StrSql = StrSql & " ORDER BY estructura1.tenro,estructura1.estrdabr,estructura2.tenro,estructura2.estrdabr"
'______________________________________
'CUANDO SOLO SE SELECCIONA EL 1ER NIVEL
ElseIf tenro1 <> "" And tenro1 <> "0" Then
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
        If estrnro1 <> "" And estrnro1 <> "0" Then
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        
        StrSql = StrSql & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
        StrSql = StrSql & " WHERE " & elEstado & StrSql2
        'StrSql = StrSql & " ORDER BY estructura1.tenro,estructura1.estrdabr"
'______________________________________________
'CUANDO NO HAY NIVEL DE ESTRUCTURA SELECCIONADO
Else
    StrSql = StrSql & " WHERE " & elEstado & StrSql2
End If

'Lo ordeno de forma ascendente para tener un mejor orden de los empleados
StrSql = StrSql & " ORDER BY empleado.ternro ASC "

Flog.writeline StrSql
OpenRecordset StrSql, rsConsult
listaEmp = ""

'If Not HastaTodos Then
    Do Until rsConsult.EOF

        If listaEmp = "" Then
            listaEmp = rsConsult!Ternro
        Else
            listaEmp = listaEmp & "," & rsConsult!Ternro
        End If
        rsConsult.MoveNext
    Loop
    Call cargarArreglo(listaEmp)
    
cant = 0

Flog.writeline
Flog.writeline "LISTA DE EMPLEADOS SELECCIONADOS DEL FILTRO (listaEmp) = " & listaEmp
Flog.writeline

'LUEGO VOY A GENERAR LA CONSULTA CON LAS 3/4 ESTRUCTURAS CONFIGURADAS EN EL
'CONFREP. ------------------------------------------
   
StrSql = " SELECT DISTINCT empleado.ternro"

cant = 1
Do Until cant > cantNiv
    StrSql = StrSql & " , tipoest" & cant & " .tedabr TE" & cant & "DESC, estact" & cant & " .estrnro ESTRNRO_" & cant & " ,est" & cant & " .estrdabr DESCESTRNRO" & cant & "  "
    cant = cant + 1
Loop

StrSql = StrSql & " FROM empleado "
cant = 1
Do Until cant > cantNiv
    StrSql = StrSql & " INNER JOIN his_estructura estact" & cant & " ON empleado.ternro = estact" & cant & ".ternro  AND estact" & cant & ".tenro  = " & estrConfrep(cant - 1)
    StrSql = StrSql & " AND (estact" & cant & ".htetdesde<=" & ConvFecha(fecestr) & " AND (estact" & cant & ".htethasta is null or estact" & cant & ".htethasta>=" & ConvFecha(fecestr) & "))"
    StrSql = StrSql & " INNER JOIN estructura est" & cant & " ON est" & cant & ".estrnro = estact" & cant & ".estrnro AND est" & cant & ".tenro = estact" & cant & ".tenro"
    StrSql = StrSql & " INNER JOIN tipoestructura tipoest" & cant & " ON tipoest" & cant & ".tenro = estact" & cant & ".tenro "
    cant = cant + 1
Loop

cant = 0
StrSql = StrSql & " WHERE ("

Do Until cant > Longitud
    Flog.writeline "Listas(cant) = " & Listas(cant)
    Flog.writeline "Con cant = " & cant
    Flog.writeline
    If Listas(cant) = "" Then
        StrSql = StrSql & " (0 = 0) "
    Else
        StrSql = StrSql & " empleado.ternro IN (" & Listas(cant) & ")"
    End If
    cant = cant + 1
    'If cant <= Longitud Then
    If cant <= Longitud And Listas(cant) <> "" Then
        'If Listas(cant) = "" Then
        '   StrSql = StrSql & " AND "
        'Else
            StrSql = StrSql & " OR "
        'End If
    End If
    If Listas(cant) = "" Then
        cant = Longitud + 1
    End If
Loop
StrSql = StrSql & ")"


cant = 1
StrSql = StrSql & " ORDER BY "

Do Until cant > cantNiv
    StrSql = StrSql & " ESTRNRO_" & cant
    cant = cant + 1
    If cant <= cantNiv Then
        StrSql = StrSql & ","
    End If
Loop

Flog.writeline
Flog.writeline "CONSULTA GENERAL!" & StrSql
Flog.writeline

OpenRecordset StrSql, rsConsult

If listaEmp <> "" Then


If Not rsConsult.EOF Then
    'Ahora tengo que calcular los conceptos/acumuladores...dependiendo de la cantidad
    'de ultimas columnas configuradas en el confrep

    'Recupero la cantidad de empleados
    cantEmpleadosPropios = 0
    cantEmpleadosAgencia = 0
    
    'Ahora comienzo a armar la matriz donde voy a guardar:
    'Ej:
    '|ger  dep  suc mobra agencia  propio
    '|1270 1289 1241 1539   0       1 |
    '|...|
    
    filas = 0
    
    'Inicializo todos los arreglos en nulo
    I = 0
    Do Until I > 4
        VecActual(I) = 0
        Vec(I) = 0
        DescTE(I) = ""
        DescEst(I) = ""
        DesEstActual(I) = ""
        I = I + 1
    Loop
    
    primero = True
    
    cantProductivo = 0
    cantNoProductivo = 0
    
    cantPropios = 0
    total = 0
    cantAgencia = 0
    Dim cargar As Boolean
    listaEmp = ""
    Dim buscar As Boolean
    buscar = True
    
    cargar = False
    
    cant = 1
    Do Until cant > cantNiv
        Select Case cant
            Case 1: DescTE(0) = rsConsult!TE1DESC
            Case 2: DescTE(1) = rsConsult!TE2DESC
            Case 3: DescTE(2) = rsConsult!TE3DESC
            Case 4: DescTE(3) = rsConsult!TE4DESC
        End Select
        cant = cant + 1
    Loop
    
    'Recupero la cantidad de filas
    
    Do Until rsConsult.EOF
                                                              
        Flog.writeline "-----------------------------------------------------------------------------------"
        Flog.writeline "EMPLEADO ACTUAL " & rsConsult!Ternro
        Flog.writeline
                    
        cant = 1
        Do Until cant > cantNiv
            
            Select Case cant
                Case 1: Vec(0) = rsConsult!ESTRNRO_1
                        DescEst(0) = rsConsult!DESCESTRNRO1
                Case 2: Vec(1) = rsConsult!ESTRNRO_2
                        DescEst(1) = rsConsult!DESCESTRNRO2
                Case 3: Vec(2) = rsConsult!ESTRNRO_3
                        DescEst(2) = rsConsult!DESCESTRNRO3
                Case 4: Vec(3) = rsConsult!ESTRNRO_4
                        DescEst(3) = rsConsult!DESCESTRNRO4
            End Select
            Flog.writeline "VEC ACTUAL (" & cant & ") = " & VecActual(cant)
            Flog.writeline "VEC (" & cant & ") = " & Vec(cant)
            Flog.writeline
            cant = cant + 1
        Loop
        
        Flog.writeline "If Iguales(VecActual, Vec) Or primero Then"
        
        If Iguales(VecActual, Vec) Or primero Then
        
            'TENGO QUE SUMAR UN EMPLEADO MAS
            If listaEmp = "" Then
                listaEmp = rsConsult!Ternro
            Else
                listaEmp = listaEmp & "," & rsConsult!Ternro
            End If
            
            Flog.writeline "(THEN) LISTA EMP ACTUAL " & listaEmp
            Flog.writeline
            
            If tipoEmpleados = 1 Then
                'AMBOS
                If perteneceLista(rsConsult!Ternro, listaAgencia) Then
                    cantAgencia = cantAgencia + 1
                    Flog.writeline
                    Flog.writeline Espacios(Tabulador * 0) & "EL EMPLEADO ACTUAL ES DE AGENCIA! " & rsConsult!Ternro
                    Flog.writeline Espacios(Tabulador * 0) & "CANT AGENCIA " & cantAgencia
                    Flog.writeline
                Else
                    If perteneceLista(rsConsult!Ternro, listaPropios) Then
                        cantPropios = cantPropios + 1
                        Flog.writeline
                        Flog.writeline Espacios(Tabulador * 0) & "EL EMPLEADO ACTUAL ES PROPIO! " & rsConsult!Ternro
                        Flog.writeline Espacios(Tabulador * 0) & "CANT PROPIO " & cantPropios
                        Flog.writeline
                    Else
                        Flog.writeline
                        Flog.writeline Espacios(Tabulador * 0) & "EL EMPLEADO ACTUAL NO ES DE AGENCIA PERO TAMPOCO ES PROPIO! ERROR! " & rsConsult!Ternro
                        Flog.writeline
                    End If
                End If

            ElseIf tipoEmpleados = 2 Then
                'PROPIOS
                cantPropios = cantPropios + 1
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 0) & "CANT PROPIO " & cantPropios
                Flog.writeline
            Else
                'AGENCIA
                cantAgencia = cantAgencia + 1
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 0) & "CANT AGENCIA " & cantAgencia
                Flog.writeline
            End If
            
                        
            If primero Then
                I = 0
                Do Until I > cantNiv
                    VecActual(I) = Vec(I)
                    DesEstActual(I) = DescEst(I)
                    I = I + 1
                Loop
                primero = False
            End If
            
            rsConsult.MoveNext
            If rsConsult.EOF Then
               cargar = True
            End If
            rsConsult.MovePrevious
        Else
        
            If listaEmp = "" Then
                listaEmp = rsConsult!Ternro
                buscar = False
            Else
                buscar = True
            End If

            Flog.writeline " (ELSE) LISTA EMP ACTUAL " & listaEmp
            Flog.writeline
            
            If MDOProdConfigurada Then
                cantProductivo = BuscarManoDeObraProd(listaEmp, fecestr)
            Else
                cantProductivo = -1
            End If
            
            If MDONProdConfigurada Then
                cantNoProductivo = BuscarManoDeObraNoProd(listaEmp, fecestr)
            Else
                cantNoProductivo = -1
            End If
            
            cantEmpleadosPropios = cantEmpleadosPropios + cantPropios
            cantEmpleadosAgencia = cantEmpleadosAgencia + cantAgencia
            total = cantAgencia + cantPropios
            
            'GUARDO LOS DATOS EN LA MATRIZ, PARA LUEGO
            'UTILIZARLA AL INSERTAR EN LA BASE DE DATOS
            
            Matriz(filas, 0) = VecActual(0)
            Matriz(filas, 1) = DesEstActual(0)
            
            Matriz(filas, 2) = VecActual(1)
            Matriz(filas, 3) = DesEstActual(1)
            
            Matriz(filas, 4) = VecActual(2)
            Matriz(filas, 5) = DesEstActual(2)
            
            Matriz(filas, 6) = VecActual(3)
            Matriz(filas, 7) = DesEstActual(3)
            
            Matriz(filas, 8) = cantProductivo
            Matriz(filas, 9) = cantNoProductivo
            
            Matriz(filas, 10) = cantPropios
            Matriz(filas, 11) = cantAgencia
            
            Matriz(filas, 12) = 0
            Matriz(filas, 13) = 0
            Matriz(filas, 14) = 0
            Matriz(filas, 15) = 0
            Matriz(filas, 16) = 0
                                    
            cant = 1
            
            Do Until cant > cantColumnasFinales
    
                Select Case tipoCalcCol(cant - 1)
                    Case "CO":
                        Matriz(filas, 11 + cant) = BuscarMontoConcepto(listPer, listaEmp, CInt(NroCA(cant - 1)))
                    Case "AC":
                        Matriz(filas, 11 + cant) = BuscarMontoAcu(listPer, listaEmp, CInt(NroCA(cant - 1)))
                    Case "CCO"
                        Matriz(filas, 11 + cant) = BuscarCantConcepto(listPer, listaEmp, CInt(NroCA(cant - 1)))
                    Case "CAC"
                        Matriz(filas, 11 + cant) = BuscarCantAcu(listPer, listaEmp, CInt(NroCA(cant - 1)))
                End Select
            
                cant = cant + 1
            Loop

            total = 0
            If tipoEmpleados = 1 Then
                'AMBOS
                If perteneceLista(rsConsult!Ternro, listaAgencia) Then
                    Flog.writeline
                    Flog.writeline Espacios(Tabulador * 0) & "EL EMPLEADO ACTUAL ES DE AGENCIA! " & rsConsult!Ternro
                    Flog.writeline
                    cantAgencia = 1
                    cantPropios = 0
                Else
                    If perteneceLista(rsConsult!Ternro, listaPropios) Then
                        Flog.writeline
                        Flog.writeline Espacios(Tabulador * 0) & "EL EMPLEADO ACTUAL ES PROPIO! " & rsConsult!Ternro
                        Flog.writeline
                        cantPropios = 1
                        cantAgencia = 0
                    Else
                        Flog.writeline
                        Flog.writeline Espacios(Tabulador * 0) & "EL EMPLEADO ACTUAL NO ES DE AGENCIA PERO TAMPOCO ES PROPIO! ERROR! " & rsConsult!Ternro
                        Flog.writeline
                        cantPropios = 0
                        cantAgencia = 0
                    End If
                End If
            ElseIf tipoEmpleados = 2 Then
                'PROPIOS
                cantPropios = 1
                cantAgencia = 0
            Else
                'AGENCIA
                cantAgencia = 1
                cantPropios = 0
            End If
            
            filas = filas + 1
                        
            listaEmp = rsConsult!Ternro
            
            I = 0
            Do Until I > cantNiv
                VecActual(I) = Vec(I)
                DesEstActual(I) = DescEst(I)
                I = I + 1
            Loop
            
        End If
        
        rsConsult.MoveNext
        Flog.writeline "-----------------------------------------------------------------------------------"
            
    Loop
    
    If cargar Then
    
        I = 0
        Do Until I > cantNiv
            VecActual(I) = Vec(I)
            DesEstActual(I) = DescEst(I)
        I = I + 1
        Loop
    
    End If
        
    If buscar Then
        If MDOProdConfigurada Then
            cantProductivo = BuscarManoDeObraProd(listaEmp, fecestr)
        Else
            cantProductivo = -1
        End If
            
        If MDONProdConfigurada Then
            cantNoProductivo = BuscarManoDeObraNoProd(listaEmp, fecestr)
        Else
            cantNoProductivo = -1
        End If
    End If
            
    'Tengo que cargar los ultimos ingresados
    cantEmpleadosPropios = cantEmpleadosPropios + cantPropios
    cantEmpleadosAgencia = cantEmpleadosAgencia + cantAgencia
    total = cantAgencia + cantPropios
                 
    'GUARDO LOS DATOS EN LA MATRIZ, PARA LUEGO
    'UTILIZARLA AL INSERTAR EN LA BASE DE DATOS
          
    Matriz(filas, 0) = VecActual(0)
    Matriz(filas, 1) = DesEstActual(0)
            
    Matriz(filas, 2) = VecActual(1)
    Matriz(filas, 3) = DesEstActual(1)
            
    Matriz(filas, 4) = VecActual(2)
    Matriz(filas, 5) = DesEstActual(2)
            
    Matriz(filas, 6) = VecActual(3)
    Matriz(filas, 7) = DesEstActual(3)
            
    Matriz(filas, 8) = cantProductivo
    Matriz(filas, 9) = cantNoProductivo
            
    Matriz(filas, 10) = cantPropios
    Matriz(filas, 11) = cantAgencia
            
    Matriz(filas, 12) = 0
    Matriz(filas, 13) = 0
    Matriz(filas, 14) = 0
    Matriz(filas, 15) = 0
    Matriz(filas, 16) = 0
            
    cant = 1
        
    Flog.writeline
    Flog.writeline "CANTIDAD COLUMNAS FINALES = " & cantColumnasFinales
    Flog.writeline
    Do Until cant > cantColumnasFinales
    
    
        Select Case tipoCalcCol(cant - 1)
            Case "CO":
                Matriz(filas, 11 + cant) = BuscarMontoConcepto(listPer, listaEmp, CInt(NroCA(cant - 1)))
            Case "AC":
                Matriz(filas, 11 + cant) = BuscarMontoAcu(listPer, listaEmp, CInt(NroCA(cant - 1)))
            Case "CCO"
                Matriz(filas, 11 + cant) = BuscarCantConcepto(listPer, listaEmp, CInt(NroCA(cant - 1)))
            Case "CAC"
                Matriz(filas, 11 + cant) = BuscarCantAcu(listPer, listaEmp, CInt(NroCA(cant - 1)))
        End Select
            
        cant = cant + 1
    Loop
    

    
    'Calculo los porcentajes propios y de agencia para cada fila ingresada en
    'la matriz y voy insertando en la base de datos
    I = 0
    Dim totalIngresar As Integer
    Do Until I > filas
        
        If cantEmpleadosPropios = 0 Then
            Matriz(I, 17) = "0"
        Else
            Matriz(I, 17) = (Matriz(I, 10) * 100) / cantEmpleadosPropios
        End If
        If cantEmpleadosAgencia = 0 Then
            Matriz(I, 18) = "0"
        Else
            Matriz(I, 18) = (Matriz(I, 11) * 100) / cantEmpleadosAgencia
        End If
        
        totalIngresar = cantEmpleadosAgencia + cantEmpleadosPropios
        
        If totalIngresar = 0 Then
            Matriz(I, 19) = "0"
        Else
            Matriz(I, 19) = (Matriz(I, 10) + Matriz(I, 11)) * 100 / totalIngresar
        End If
               
        
        '------------------------------------------------------
        'INSERTO LOS DATOS EN LA BASE DE DATOS
        '------------------------------------------------------
        
        StrSql = " INSERT INTO rep_distribucion ( "
        StrSql = StrSql & " bpronro,Fecha,Hora,iduser,bprcparam,te1,tedesc1,estrnro1,estrdesc1  "
        StrSql = StrSql & " ,te2,tedesc2,estrnro2,estrdesc2,te3,tedesc3,estrnro3,estrdesc3 "
        StrSql = StrSql & " ,te4,tedesc4,estrnro4,estrdesc4,cantProd,cantNoProd "
        StrSql = StrSql & " ,cant,cant_propios,cant_agencia,porc,porc_propios,porc_agencia  "
        StrSql = StrSql & " ,col1,col1_desc,col2,col2_desc,col3,col3_desc,col4,col4_desc,col5,col5_desc  "
        StrSql = StrSql & " ) values ( "
        StrSql = StrSql & NroProceso & ","
        StrSql = StrSql & Fecha & ","
        StrSql = StrSql & "'" & Mid(hora, 1, 8) & "',"
        StrSql = StrSql & "'" & Mid(iduser, 1, 20) & "',"
        StrSql = StrSql & "'" & Mid(Parametros, 1, 3000) & "',"
        
        'ESTRUCTURA 1
        StrSql = StrSql & estrConfrep(0) & ","
        StrSql = StrSql & "'" & Mid(DescTE(0), 1, 60) & "',"
        StrSql = StrSql & Matriz(I, 0) & ","
        StrSql = StrSql & "'" & Matriz(I, 1) & "',"
                
        'ESTRUCTURA 2
        StrSql = StrSql & estrConfrep(1) & ","
        StrSql = StrSql & "'" & Mid(DescTE(1), 1, 60) & "',"
        StrSql = StrSql & Matriz(I, 2) & ","
        StrSql = StrSql & "'" & Mid(Matriz(I, 3), 1, 60) & "',"
    
        'ESTRUCTURA 3
        StrSql = StrSql & estrConfrep(2) & ","
        StrSql = StrSql & "'" & Mid(DescTE(2), 1, 60) & "',"
        StrSql = StrSql & Matriz(I, 4) & ","
        StrSql = StrSql & "'" & Mid(Matriz(I, 5), 1, 60) & "',"
        
        'ESTRUCTURA 4
        StrSql = StrSql & estrConfrep(3) & ","
        StrSql = StrSql & "'" & Mid(DescTE(3), 1, 60) & "',"
        StrSql = StrSql & Matriz(I, 6) & ","
        StrSql = StrSql & "'" & Mid(Matriz(I, 7), 1, 60) & "',"
    
        'PRODUCTIVOS Y NO PRODUCTIVOS
        StrSql = StrSql & Matriz(I, 8) & ","
        StrSql = StrSql & Matriz(I, 9) & ","
    
        'CANTIDADES
        StrSql = StrSql & Matriz(I, 10) + Matriz(I, 11) & ","
        StrSql = StrSql & Matriz(I, 10) & ","
        StrSql = StrSql & Matriz(I, 11) & ","
            
        'PORCENTAJES
        StrSql = StrSql & Matriz(I, 19) & ","
        StrSql = StrSql & Matriz(I, 17) & ","
        StrSql = StrSql & Matriz(I, 18) & ","
            
        'ULTIMAS COLUMNAS
        StrSql = StrSql & Matriz(I, 12) & ","
        StrSql = StrSql & "'" & Mid(DescCol(0), 1, 60) & "',"
            
        StrSql = StrSql & Matriz(I, 13) & ","
        StrSql = StrSql & "'" & Mid(DescCol(1), 1, 60) & "',"
            
        StrSql = StrSql & Matriz(I, 14) & ","
        StrSql = StrSql & "'" & Mid(DescCol(2), 1, 60) & "',"
            
        StrSql = StrSql & Matriz(I, 15) & ","
        StrSql = StrSql & "'" & Mid(DescCol(3), 1, 60) & "',"
            
        StrSql = StrSql & Matriz(I, 16) & ","
        StrSql = StrSql & "'" & Mid(DescCol(4), 1, 60) & "'"
            
        StrSql = StrSql & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords

        I = I + 1
    Loop
    
End If

End If

End Sub
Function BuscarManoDeObraProd(listEmp As String, fecestr As String)

    Dim rsAux As New ADODB.Recordset
    Dim Valor As Integer
    Valor = 0
        
    If listEmp <> "" Then
        StrSql = " SELECT * "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN confrep ON confrep.repnro = 347 AND confrep.conftipo = 'MOP'"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND estructura.tenro = 11 "
        StrSql = StrSql & " WHERE his_estructura.estrnro = confrep.confval2"
        StrSql = StrSql & " AND (his_estructura.htetdesde<=" & ConvFecha(fecestr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta>=" & ConvFecha(fecestr) & "))"
        StrSql = StrSql & " AND his_estructura.ternro IN (" & listEmp & ")"
        
        OpenRecordset StrSql, rsAux
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "CONSULTA BUSQUEDA MANODEOBRA PROD " & StrSql
        

        Do Until rsAux.EOF
            Valor = Valor + 1
            rsAux.MoveNext
        Loop
        
        Flog.writeline Espacios(Tabulador * 0) & "VALOR FINAL DE MANO DE OBRA PROD " & Valor
        Flog.writeline
    
    End If
    
    BuscarManoDeObraProd = Valor
End Function
Function BuscarManoDeObraNoProd(listEmp As String, fecestr As String)

    Dim rsAux As New ADODB.Recordset
    Dim Valor As Integer
    Valor = 0
    
    If listEmp <> "" Then
        StrSql = " SELECT * "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN confrep ON confrep.repnro = 347 AND confrep.conftipo = 'MON'"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
        StrSql = StrSql & " AND estructura.tenro = 11 "
        StrSql = StrSql & " WHERE his_estructura.estrnro = confrep.confval2"
        StrSql = StrSql & " AND (his_estructura.htetdesde<=" & ConvFecha(fecestr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta>=" & ConvFecha(fecestr) & "))"
        StrSql = StrSql & " AND his_estructura.ternro IN (" & listEmp & ")"
        
        OpenRecordset StrSql, rsAux
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "CONSULTA BUSQUEDA MANODEOBRA NO PROD " & StrSql

        Do Until rsAux.EOF
            Valor = Valor + 1
            rsAux.MoveNext
        Loop
        
        Flog.writeline Espacios(Tabulador * 0) & "VALOR FINAL DE MANO DE OBRA PROD " & Valor
        Flog.writeline
    End If
    
    BuscarManoDeObraNoProd = Valor
    
End Function

Function BuscarMontoConcepto(listPliq As String, listEmp As String, nroconc As Integer) As Double

    Dim rsConsult As New ADODB.Recordset
    Dim Valor As Double
    Dim cant As Integer
    
    Valor = 0
    
    StrSql = " SELECT detliq.dlimonto "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " INNER JOIN concepto ON  concepto.concnro = detliq.concnro "
    StrSql = StrSql & " WHERE (1=1) "
     
    If listEmp <> "" Then
        StrSql = StrSql & " AND cabliq.empleado IN (" & listEmp & ")"
    End If
        
    If listPliq <> "" Then
        StrSql = StrSql & " AND proceso.pliqnro IN (" & listPliq & ")"
    End If
    If tipoprocesos <> -1 Then
        StrSql = StrSql & " AND proceso.pronro in (" & CStr(Replace(Lista_Pro, "-", ",")) & ")"
    End If
    StrSql = StrSql & " AND concepto.concnro = " & nroconc
    
    
    OpenRecordset StrSql, rsConsult

    Do Until rsConsult.EOF
        Valor = Valor + rsConsult!dlimonto
        rsConsult.MoveNext
    Loop
        
    
    BuscarMontoConcepto = Valor
    
End Function

Function perteneceLista(emp As Long, Lista As String) As Boolean
        
    Dim Valor As Boolean
    Dim listaaux
    Dim primer
    
    Valor = False
    listaaux = Split(Lista, ",", -1)
    primer = 0
    Do Until primer > UBound(listaaux) Or Valor = True
        If listaaux(primer) = CStr(emp) Then
            Valor = True
        Else
            primer = primer + 1
        End If
    Loop

    
    perteneceLista = Valor
        
    
End Function
Function BuscarMontoAcu(listPliq As String, listEmp As String, nroacu As Integer) As Double

    Dim rsConsult As New ADODB.Recordset
    Dim Valor As Double
    Dim cant As Integer
    
    Valor = 0
    
    StrSql = " SELECT acu_liq.almonto "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro"
    StrSql = StrSql & " WHERE (1=1) "
    
    If listEmp <> "" Then
        StrSql = StrSql & " AND cabliq.empleado IN (" & listEmp & ")"
    End If
    
    If listPliq <> "" Then
        StrSql = StrSql & " AND proceso.pliqnro IN (" & listPliq & ")"
    End If
    If tipoprocesos <> -1 Then
        StrSql = StrSql & " AND proceso.pronro in (" & CStr(Replace(Lista_Pro, "-", ",")) & ")"
    End If
    StrSql = StrSql & " AND acu_liq.acunro = " & nroacu
    
   
    OpenRecordset StrSql, rsConsult
   
    Do Until rsConsult.EOF
        Valor = Valor + rsConsult!almonto
        rsConsult.MoveNext
    Loop
        
    
    BuscarMontoAcu = Valor
   
    
End Function
Function BuscarCantConcepto(listPliq As String, listEmp As String, nroconc As Integer) As Double

    Dim rsConsult As New ADODB.Recordset
    Dim Valor As Double
    Dim cant As Integer
    
    Valor = 0
    
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " INNER JOIN concepto ON  concepto.concnro = detliq.concnro "
    StrSql = StrSql & " WHERE (1=1) "
    
    If listEmp <> "" Then
        StrSql = StrSql & " AND cabliq.empleado IN (" & listEmp & ")"
    End If
    
    If listPliq <> "" Then
        StrSql = StrSql & " AND proceso.pliqnro IN (" & listPliq & ")"
    End If
    If tipoprocesos <> -1 Then
        StrSql = StrSql & " AND proceso.pronro in (" & CStr(Replace(Lista_Pro, "-", ",")) & ")"
    End If
    StrSql = StrSql & " AND concepto.concnro = " & nroconc
    
    
    OpenRecordset StrSql, rsConsult
        
    Do Until rsConsult.EOF
        Valor = Valor + 1
        rsConsult.MoveNext
    Loop
    
   
    BuscarCantConcepto = Valor
    
End Function
Function BuscarCantAcu(listPliq As String, listEmp As String, nroacu As Integer) As Double

    Dim rsConsult As New ADODB.Recordset
    Dim Valor As Double
    Dim cant As Integer
    
    Valor = 0
    
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro"
    StrSql = StrSql & " WHERE (1=1) "
    
    If listEmp <> "" Then
        StrSql = StrSql & " AND cabliq.empleado IN (" & listEmp & ")"
    End If
    
    If listPliq <> "" Then
        StrSql = StrSql & " AND proceso.pliqnro IN (" & listPliq & ")"
    End If
    If tipoprocesos <> -1 Then
        StrSql = StrSql & " AND proceso.pronro in (" & CStr(Replace(Lista_Pro, "-", ",")) & ")"
    End If
    StrSql = StrSql & " AND acu_liq.acunro = " & nroacu
    
    
    OpenRecordset StrSql, rsConsult

    Valor = 0
    
    Do Until rsConsult.EOF
        Valor = Valor + 1
        rsConsult.MoveNext
    Loop
    
    
    BuscarCantAcu = Valor
    
End Function


Function Iguales(Arr1, Arr2) As Boolean

    Dim I As Integer
    Dim es As Boolean
    es = True
    
    I = 0
    
    Do Until (Not es) Or (I = 5)
        If Arr1(I) <> Arr2(I) Then
            es = False
        Else
            I = I + 1
        End If
    Loop
    
    Iguales = es
    
End Function

Sub cargarArreglo(listEmp As String)

    Dim Div As Integer
    Dim listaaux
    Dim j As Integer
    Dim I As Integer
    Dim termino As Boolean
    Dim total As Integer
    Dim losQQuedan As Integer
    
    'Div = Len(listEmp)
    Flog.writeline
        
    listaaux = Split(listEmp, ",", -1)
    
    Div = UBound(listaaux)
    
    Flog.writeline listEmp
    Flog.writeline "LA CANTIDAD DE EMPLEADOS SELECCIONADA ES UBound(listaaux) " & Div
    
    If Div >= 800 Then
    
        
        Longitud = CInt(Div / 800)
        If ((Div / 800) - Longitud) > 0 Then
            Longitud = Longitud + 1
        End If
        
        Flog.writeline "LEN ES MAYOR O IGUAL A 800 "
        Flog.writeline "Longitud = " & Longitud
        
        'listaaux = Split(listEmp, ",", -1)
        
        j = 0
        I = 0
        'total = CInt(UBound(listaaux) / Longitud)
        total = 800
        
        Flog.writeline "TOTAL = " & total
        
        
        Do Until j = Longitud Or termino
        
            Listas(j) = ""
            termino = False
            
            Flog.writeline
            Flog.writeline "Total Actual " & total
            Flog.writeline "UBound(listaaux) " & UBound(listaaux)
            Flog.writeline
            
            'If total > UBound(listaaux) Then
            'If UBound(listaaux) > total Then
            '    Flog.writeline "TERMINOOO total > UBound(listaaux) "
            '    Flog.writeline
            '    termino = True
            'Else
                Do Until I > total Or termino
                    If listaaux(I) = "" Then
                        Flog.writeline "TERMINOOO listaaux(i) vacio!"
                        Flog.writeline
                        termino = True
                    Else
                        If Listas(j) = "" Then
                            Listas(j) = listaaux(I)
                        Else
                            Listas(j) = Listas(j) & "," & listaaux(I)
                        End If
                        'Flog.writeline "Listaaux(i) = " & listaaux(I)
                        'Flog.writeline " CON i = " & I
                        'Flog.writeline
                        I = I + 1
                    End If
                Loop
            'End If
            Flog.writeline "Listas(j) = " & Listas(j)
            Flog.writeline " CON j = " & j
            Flog.writeline
            
            losQQuedan = UBound(listaaux) - total
            
            If losQQuedan < 800 Then
                total = total + losQQuedan
            Else
                total = total + 800
            End If
            j = j + 1
            
        Loop

    Else
        Listas(0) = listEmp
        Flog.writeline "LEN ES MENOR A 800 "
        Flog.writeline "Listas(0) = " & Listas(0)
        Flog.writeline
    End If
            
         
End Sub

