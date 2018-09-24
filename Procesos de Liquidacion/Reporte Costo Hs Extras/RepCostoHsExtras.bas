Attribute VB_Name = "RepCostoHsExtras"
'Global Const Version = "1.00"
'Global Const FechaModificacion = "08/08/2012"
'Global Const UltimaModificacion = " " 'Sebastian Stremel - Version inicial

'Global Const Version = "1.01"
'Global Const FechaModificacion = "09/08/2012"
'Global Const UltimaModificacion = " " 'Sebastian Stremel - cambio en la busqueda de la cantidad de los conceptos

'Global Const Version = "1.02"
'Global Const FechaModificacion = "13/08/2012"
'Global Const UltimaModificacion = " " 'Sebastian Stremel - cambio en la busqueda de la cantidad de los conceptos para los graficos

'Global Const Version = "1.03"
'Global Const FechaModificacion = "28/09/2012"
'Global Const UltimaModificacion = " " 'Sebastian Stremel - cambios para sumar las columnas 2,3,4,5

'Global Const Version = "1.04"
'Global Const FechaModificacion = "23/12/2013"
'Global Const UltimaModificacion = " " 'Carmen Quintero (CAS-22838 - MIMO - ERROR EN REPORTE COSTO DE HORAS EXTRAS) -
                                      'Se saco la columna sedenestr1 de los procedimientos CargarEmpleados, CargarEmpleadosGraficos

'Global Const Version = "1.05"
'Global Const FechaModificacion = "07/02/2014"
'Global Const UltimaModificacion = " " 'Carmen Quintero (CAS-22838 - MIMO - ERROR EN REPORTE COSTO DE HORAS EXTRAS) -
                                      'Se realizaron modificaciones para mejorar su procesamiento
                                      
'Global Const Version = "1.06"
'Global Const FechaModificacion = "19/02/2014"
'Global Const UltimaModificacion = " " 'Carmen Quintero (CAS-22838 - MIMO - ERROR EN REPORTE COSTO DE HORAS EXTRAS) -
                                      'Se agregaron varias validaciones

'Global Const Version = "1.07"
'Global Const FechaModificacion = "30/05/2014"
'Global Const UltimaModificacion = " " 'Carmen Quintero CAS-22838 - MIMO - ERROR EN REPORTE COSTO DE HORAS EXTRAS [Entrega 4] (CAS-15298) -
                                      'Se agregaron varias validaciones

'Global Const Version = "1.08"
'Global Const FechaModificacion = "29/09/2014"
'Global Const UltimaModificacion = " " 'Carmen Quintero CAS-22838 - MIMO - OVERFLOW EN REPORTE COSTO DE HORAS EXTRAS -
                                      'Se agregaron varias validaciones
'Global Const Version = "1.09"
'Global Const FechaModificacion = "09/10/2014"
'Global Const UltimaModificacion = " " 'Fernandez, Matias - CAS-22838 - MIMO - ERROR EN COMPARACION DE MESES EN REPORTE -
                                      ' nada nuevo, solo imprimo el log para detectar posibles errores
                                      
'Global Const Version = "1.10"
'Global Const FechaModificacion = "02/12/2014"
'Global Const UltimaModificacion = " " 'Fernandez, Matias - CAS-22838 - MIMO - ERROR EN COMPARACION DE MESES EN REPORTE -
                                      'se recalcula el porcentaje de progreso y se agrega mas info al log
                                      
'Global Const Version = "1.11"
'Global Const FechaModificacion = "28/01/2015"
'Global Const UltimaModificacion = " " 'Fernandez, Matias - CAS-22838 - MIMO - ERROR EN COMPARACION DE MESES EN REPORTE -
                                      'se acomodan calculos de porcentaje y progreso
                                      
'Global Const Version = "1.12"
'Global Const FechaModificacion = "06/03/2015"
'Global Const UltimaModificacion = " " 'Fernandez, Matias - CAS-22838 - MIMO - ERROR EN COMPARACION DE MESES EN REPORTE -
                                      'Se optimizo consultas sql, se  comento codigo que no era necesario.
                                      
'Global Const Version = "1.13"
'Global Const FechaModificacion = "16/03/2015"
'Global Const UltimaModificacion = " " 'Fernandez, Matias - CAS-22838 - MIMO - ERROR EN COMPARACION DE MESES EN REPORTE -
                                      'mejora en consultas sql y tiempo de procesamiento
                                      


                                      
'Global Const Version = "1.14"
'Global Const FechaModificacion = "08/06/2015"
'Global Const UltimaModificacion = " " 'Fernandez, Matias - CAS-22838 - MIMO - ERROR EN COMPARACION DE MESES EN REPORTE -
                                      'optimizacion de reporte, reduccion de consultas a la base de datos y de inserts.
                                      

Global Const Version = "1.15"
Global Const FechaModificacion = "06/07/2015"
Global Const UltimaModificacion = " " 'Fernandez, Matias - CAS-22838 - MIMO - ERROR EN COMPARACION DE MESES EN REPORTE -
                                      'Se agrego un incremento de indice de arreglos de sedes q faltaba

'---------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------

Option Explicit

Dim fs, f
'Global Flog

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global arrTipoConc(1000) As Integer

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String

Global empresa
Global emprNro

Global emprActiv
Global param_empresa
Global listapronro
Global l_orden
Global filtro
Global totalEmpleados
Global cantRegistros
'Global fecEstr
Global tituloFiltro
Global detalle
Global agrupado
'Global empresa
Global Orden
Global proaprob
Global fechadesde
Global descDesde
Global fechahasta
Global descHasta


'variables confrep
Global esconcA As Boolean
Global concA
Global esconcMontoA As Boolean
Global montoA
Global esconcB As Boolean
Global cantidadB
Global esconcMontoB As Boolean
Global montoB
Global esconcAcuA As Boolean
Global acuA
Global porc As Double
Global porcentaje As Double
Global estrdet
Global estragrup

' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial
' Autor      : Sebastian Stremel
' Fecha      : 17/07/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Public Sub Main()

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset

Dim PID As String
Dim bprcparam As String
Dim ArrParametros
Dim parametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProcesoBatch = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProcesoBatch = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    Nombre_Arch = PathFLog & "Reporte Costo Hs Extras" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline

    On Error Resume Next
    
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo 0
    
    'FGZ - 11/11/2011 --------- Control de versiones ------
    'Version_Valida = ValidarV(Version, 371, TipoBD)
    'If Not Version_Valida Then
        'Actualizo el progreso
    '    MyBeginTrans
    '        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    '        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    '    MyCommitTrans
    '    Flog.writeline
    '    GoTo Fin
    'End If
    'FGZ - 11/11/2011 --------- Control de versiones ------
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 372 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
       Flog.writeline "Hora de inicio: " & Now()

        Call DatosReporte(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    Flog.writeline "---------------------------------------------------"
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General(1): " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
    
End Sub
    
Sub DatosReporte(NroProcesoBatch, bprcparam)
Dim rsEmpl As New ADODB.Recordset
Dim objRs As New ADODB.Recordset
Dim parametros
Dim ArrParametros
Dim sectorAnt
Dim sedeAnt
Dim i As Integer
Dim sql As String
Dim StrSqlAux As String

'Dim arrEmpleados(2000) As String
'Dim arrSector(2000) As Integer
'Dim arrSede(2000) As Integer
'Dim arrMeses(2000) As Integer
'Dim arrAnio(2000) As Integer

Dim arrEmpleados()
Dim arrSector()
Dim arrSede()
Dim arrMeses()
Dim arrAnio()

Dim valorCantidadA
Dim valorMontoA
Dim valorCantidadB
Dim valorMontoB
Dim valorAcuA
Dim sqlAux As String
On Error GoTo CE

    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProcesoBatch
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
        
        'Obtengo los parametros del proceso
        parametros = objRs!bprcparam
    
        Flog.writeline "Lista de Parametros = " & parametros
        
        ArrParametros = Split(parametros, "@")
               
        'Obtengo la lista de procesos
        Flog.writeline "Obtengo la Lista de Procesos"
        
        listapronro = ArrParametros(0)
            
        Flog.writeline "Lista de Procesos = " & listapronro
        
        'obtengo el tipo de estructura 1
        tenro1 = ArrParametros(1)
        Flog.writeline "Obtengo el tipo de estructura 1 = " & tenro1
            
        'obtengo la estructura 1
        estrnro1 = ArrParametros(2)
        Flog.writeline "Obtengo la estructura 1 = " & estrnro1
         
        'obtengo el tipo de estructura 2
        tenro2 = ArrParametros(3)
        Flog.writeline "Obtengo el tipo de estructura 2 = " & tenro2
            
        'obtengo la estructura 2
        estrnro2 = ArrParametros(4)
        Flog.writeline "Obtengo la estructura 2 = " & estrnro2
        
        'obtengo el tipo de estructura 3
        tenro3 = ArrParametros(5)
        Flog.writeline "Obtengo el tipo de estructura 3 = " & tenro3
           
        'obtengo la estructura 2
        estrnro3 = ArrParametros(6)
        Flog.writeline "Obtengo la estructura 3 = " & estrnro3
        
        'obtengo la fecha de la estructura
        fecEstr = ArrParametros(7)
        Flog.writeline "Obtengo la fecha de la estructura = " & fecEstr
         
        'obtengo el titulo del filtro
        tituloFiltro = ArrParametros(8)
        Flog.writeline "Obtengo el titulo del filtro = " & tituloFiltro
        
        'obtengo la estructura de detalle
        detalle = ArrParametros(9)
        Flog.writeline "Obtengo la estructura de detalle = " & detalle
        
        'obtengo la estructura de agrupado
        agrupado = ArrParametros(10)
        Flog.writeline "Obtengo la estructura de agrupado = " & agrupado
        
        'obtengo el orden
        l_orden = " pliqanio, pliqmes, "
        l_orden = l_orden & ArrParametros(11)
        
        Flog.writeline "Obtengo el orden = " & Orden
        
        'empresa
        empresa = ArrParametros(12)
        Flog.writeline "Obtengo la empresa = " & empresa
         
        'proceso aprobado
        proaprob = ArrParametros(13)
        Flog.writeline "Obtengo que estado de proceso busco = " & proaprob
        
        'Obtengo el periodo desde
        Flog.writeline "Obtengo el Período Desde"
        pliqdesde = CLng(ArrParametros(14))
        Flog.writeline "Período Desde = " & pliqdesde
            
        'Obtengo el periodo hasta
        Flog.writeline "Obtengo el Período Hasta"
        pliqhasta = CLng(ArrParametros(15))
        Flog.writeline "Período Hasta = " & pliqhasta
            
            
        'obtengo el filtro
        Flog.writeline "Obtengo el filtro"
        filtro = ArrParametros(16)
        Flog.writeline "Filtro = " & filtro
        
        

        
        'EMPIEZA EL PROCESO
        'Busco el periodo desde
        StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqdesde
        OpenRecordset StrSql, objRs
         
        If Not objRs.EOF Then
           fechadesde = objRs!pliqdesde
           descDesde = objRs!pliqDesc
        Else
           Flog.writeline "No se encontro el periodo desde."
           Exit Sub
        End If
         
        objRs.Close
            
        'Busco el periodo hasta
        StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqhasta
        OpenRecordset StrSql, objRs
         
        If Not objRs.EOF Then
           fechahasta = objRs!pliqhasta
           descHasta = objRs!pliqDesc
        Else
           Flog.writeline "No se encontro el periodo hasta."
           Exit Sub
        End If
             
        objRs.Close
    
    End If
    

    
    
    'levanto el confrep
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro=374"
    OpenRecordset StrSql, objRs
    concA = "0"
    montoA = "0"
    cantidadB = "0"
    montoB = "0"
    Do While Not objRs.EOF
        Select Case objRs!confnrocol
            Case 2:
                'cantidad concepto a
                If objRs!conftipo = "CO" Then
                    esconcA = True
                    concA = concA & ", " & objRs!confval
                Else
                    esconcA = False
                    concA = concA & ", " & objRs!confval2
                End If
            
            Case 3:
                'monto concepto a
                If objRs!conftipo = "CO" Then
                    esconcMontoA = True
                    montoA = montoA & ", " & objRs!confval
                Else
                    esconcMontoA = False
                    montoA = montoA & ", " & objRs!confval2
                End If
            
            Case 4:
                'cantidad concepto B
                If objRs!conftipo = "CO" Then
                    esconcB = True
                    cantidadB = cantidadB & ", " & objRs!confval
                Else
                    esconcB = False
                    cantidadB = cantidadB & ", " & objRs!confval2
                End If
        
            Case 5:
                'monto concepto B
                If objRs!conftipo = "CO" Then
                    esconcMontoB = True
                    montoB = montoB & ", " & objRs!confval
                Else
                    esconcMontoB = False
                    montoB = montoB & ", " & objRs!confval2
                End If
        
            Case 6:
                'monto acumulador A
                If objRs!conftipo = "CO" Then
                    esconcAcuA = True
                    acuA = objRs!confval
                Else
                    esconcAcuA = False
                    acuA = objRs!confval2
                End If
            Case 7:
                'tipo estructura detalle
                If objRs!conftipo = "TE" Then
                    estrdet = objRs!confval
                Else
                    Flog.writeline "Falta configurar tipo de estructura detalle"
                End If
            Case 8:
                'tipo estructura agrupado
                If objRs!conftipo = "TE" Then
                    estragrup = objRs!confval
                Else
                    Flog.writeline "Falta configurar tipo de estructura agrupado"
                End If
        
        End Select

    objRs.MoveNext
    Loop
    'hasta aca
    
    
    
    'Obtengo los empleados sobre los que tengo que generar el reporte
    Flog.writeline "Cargo los Empleados "
    
    Call CargarEmpleados(listapronro, rsEmpl)

    '_______________________ARREGLO DE EMPLEADOS_____________________
    'buscar los distintos
    i = 0
    Dim emplAnt
    emplAnt = ""
    sedeAnt = ""
    sectorAnt = ""
    
    If rsEmpl.EOF Then
        Flog.writeline "No hay empleados para procesar"
        Exit Sub
    Else
        Do While Not rsEmpl.EOF
           ' Flog.writeline "Empleado: " & rsEmpl!empleg
           'If (emplAnt <> rsEmpl!empleg) Then 05/03/2015
           '     I = I + 1
           '     ReDim Preserve arrEmpleados(I)
           '     arrEmpleados(I) = rsEmpl!empleg & "@" & rsEmpl!Ternro & "@"
           '     emplAnt = rsEmpl!empleg
           ' End If 05/03/2015
            'sedes
            If (sedeAnt <> rsEmpl!sedeestrnro) Then
                ReDim Preserve arrSede(i)
                Flog.writeline "Sede: " & rsEmpl!sedeestrnro
                arrSede(i) = rsEmpl!sedeestrnro
                sedeAnt = rsEmpl!sedeestrnro
                i = i + 1
            End If
            'sector 05/03/2015
           '  If (sectorAnt <> rsEmpl!sectorestrnro) Then
           '     'I = I + 1
           '     ReDim Preserve arrSector(I)
           '     Flog.writeline "Sector: " & rsEmpl!sectorestrnro
           '     arrSector(I) = rsEmpl!sectorestrnro
           '    sectorAnt = rsEmpl!sectorestrnro
           ' End If 05/03/2015
            
            rsEmpl.MoveNext
        Loop
    End If
    '________________________HASTA ACA_______________________________
    
    
    'GUARDO EN LA CABECERA DEL REPORTE_______________________________
    StrSql = " INSERT INTO rep_costo_HS_Extras "
    StrSql = StrSql & "(bpronro, repdescabr, repdescext, cantEmp)"
    StrSql = StrSql & " VALUES ( "
    StrSql = StrSql & NroProcesoBatch & ", "
    StrSql = StrSql & "'" & tituloFiltro & "'  , " & ""
    StrSql = StrSql & "'" & tituloFiltro & "'  , " & ""
    StrSql = StrSql & i & " )"
    Flog.writeline "SQL INSERT: " & StrSql
    objConn.Execute StrSql, , adExecuteNoRecords
    '________________________________________________________________
    
    
    '______________________ARREGLO DE SEDES__________________________
    'If rsEmpl.EOF Then
    '    GoTo CE
    'Else
    
'    rsEmpl.MoveFirst
'    sedeAnt = ""
'    I = 0
'    Do While Not rsEmpl.EOF
'        If (sedeAnt <> rsEmpl!sedeestrnro) Then
'            I = I + 1
'            'ReDim Preserve arrSede(I)
'            arrSede(I) = rsEmpl!sedeestrnro
'            sedeAnt = rsEmpl!sedeestrnro
'        End If
'
'    rsEmpl.MoveNext
'    Loop
    'End If
    '________________________HASTA ACA_______________________________
    
    
    
    '______________________ARREGLO DE SECTORES______________________
'    rsEmpl.MoveFirst
'    sectorAnt = ""
'    I = 0
'    Do While Not rsEmpl.EOF
'        If (sectorAnt <> rsEmpl!sectorestrnro) Then
'            I = I + 1
'            'ReDim Preserve arrSector(I)
'            arrSector(I) = rsEmpl!sectorestrnro
'            sectorAnt = rsEmpl!sectorestrnro
'        End If
'
'    rsEmpl.MoveNext
'    Loop
    '________________________HASTA ACA_______________________________
    
    
    '______________________ARREGLO DE AÑOS___________________________
    
    i = 0
    Dim anioAnt
    anioAnt = ""
   ' Do While Not rsEmpl.EOF 05/03/2015
   '     Flog.writeline "AÑO: " & rsEmpl!pliqanio
   '     If (anioAnt <> rsEmpl!pliqanio) Then
   '         I = I + 1
   '         ReDim Preserve arrAnio(I)
   '         arrAnio(I) = rsEmpl!pliqanio
   '         anioAnt = rsEmpl!pliqanio
   '     End If
   '     rsEmpl.MoveNext
   ' Loop  05/03/2015
    
    '________________________HASTA ACA_______________________________
    
    
    
    
    '______________________ARREGLO DE MESES__________________________
    
    i = 0
    Dim mesAnt
    mesAnt = ""
   ' Do While Not rsEmpl.EOF 05/03/2015
   '     Flog.writeline "Mes: " & rsEmpl!pliqmes
   '     If (mesAnt <> rsEmpl!pliqmes) Then
   '         I = I + 1
   '         ReDim Preserve arrMeses(I)
   '         arrMeses(I) = rsEmpl!pliqmes
   '         mesAnt = rsEmpl!pliqmes
   '     End If
   '     rsEmpl.MoveNext
   ' Loop 05/03/2015
    
    '________________________HASTA ACA______________________________
    


    '_____________CICLO POR CADA SEDE Y BUSCO LOS DATOS_____________
    Dim j
    'Dim sede As Integer
    Dim sede As Long
    Dim sector As Long
    Dim sectoresactuales As String
    'Dim sectorAnt
    sectorAnt = ""
    mesAnt = ""
    anioAnt = ""
    sectoresactuales = "0"
    j = 0
    
    porcentaje = 0
    'porc = 5 / (totalEmpleados * (UBound(arrSede) + 1))
    
    Flog.writeline "Datos del reporte " & (CDbl(totalEmpleados) * (CDbl(UBound(arrSede) + 1)))
    Flog.writeline CDbl(totalEmpleados)
    Flog.writeline CDbl(UBound(arrSede) + 1)
    If (CDbl(totalEmpleados) * (CDbl(UBound(arrSede) + 1))) > 0 Then
        'porc = 5 / (CDbl(totalEmpleados) * CDbl((UBound(arrSede) + 1))) 29/05/2015
         porc = 2 / (CDbl(totalEmpleados) * CDbl((UBound(arrSede) + 1))) '29/05/2015
    Else
        porc = 0
        'Flog.writeline "EL PORCENTAJE ES 0"
    End If
    Flog.writeline "----------------------------------------------------------------"
    For j = 0 To UBound(arrSede)
       Flog.writeline arrSede(j)
    Next
    Flog.writeline "----------------------------------------------------------------"
    For j = 0 To UBound(arrSede)
       rsEmpl.MoveFirst
       sede = CLng(arrSede(j))
       If sede <> CLng(0) Then
        Do While Not rsEmpl.EOF
            If (rsEmpl!sedeestrnro = sede) Then
                If ((sectorAnt <> rsEmpl!sectorestrnro) Or (mesAnt <> rsEmpl!pliqmes) Or (anioAnt <> rsEmpl!pliqanio)) Then
                    sector = rsEmpl!sectorestrnro
                    sectoresactuales = sectoresactuales & "," & sector
                    StrSql = "INSERT INTO rep_costo_HS_Extras_det "
                    StrSql = StrSql & "(bpronro,sede,sector,anio,mes)  "
                    StrSql = StrSql & " VALUES ( "
                    StrSql = StrSql & NroProcesoBatch & ", "
                    StrSql = StrSql & sede & ", "
                    StrSql = StrSql & sector & ", "
                    StrSql = StrSql & rsEmpl!pliqanio & ", "
                    StrSql = StrSql & rsEmpl!pliqmes & " )"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    sectorAnt = sector
                    mesAnt = rsEmpl!pliqmes
                    anioAnt = rsEmpl!pliqanio
                End If
            End If
            rsEmpl.MoveNext
                        
            'progreso seba 01/08/2012
            'If porcentaje < 10 Then MDF
             porcentaje = porcentaje + porc
            'Actualizo el estado del proceso
            'End If
            TiempoAcumulado = GetTickCount
            sqlAux = "UPDATE batch_proceso SET bprcprogreso = " & Replace(porcentaje, ",", ".")
            sqlAux = sqlAux & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
            sqlAux = sqlAux & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
            objConn.Execute sqlAux, , adExecuteNoRecords
        Loop
       End If
    Next
    
    Flog.writeline "Sectores actuales: " & sectoresactuales
    '________________________________________________________________
    Flog.writeline "----------------------------------------------------------------"
    Flog.writeline "----------------------------------------------------------------"
    
    '________calculo la cantidad para cada registro_______
    StrSql = "SELECT * FROM rep_costo_HS_Extras_det "
    StrSql = StrSql & " WHERE bpronro=" & NroProcesoBatch
    OpenRecordset StrSql, rs
    Dim Anio
    Dim Mes
    Dim cantEmpl
    Dim ternroAnt
    Dim total
    ternroAnt = ""
    total = rs.RecordCount
    'porc = 5 / (totalEmpleados * total)
    Flog.writeline "--------------------------"
    Flog.writeline "datos de reporte segunda parte: " & (CDbl(totalEmpleados) * CDbl(total))
    Flog.writeline porcentaje
    Flog.writeline "---------------------------"
    If (CDbl(totalEmpleados) * CDbl(total)) > 0 Then
        'porc = 5 / (CDbl(totalEmpleados) * CDbl(total)) '29/05/2015
        porc = 1 / (CDbl(totalEmpleados) * CDbl(total)) '29/05/2015
    Else
        porc = 0

    End If
    
    Do While Not rs.EOF
        rsEmpl.MoveFirst
        cantEmpl = 0
        sede = rs!sede
        sector = rs!sector
        Anio = rs!Anio
        Mes = rs!Mes
        ternroAnt = ""
        Do While Not rsEmpl.EOF
            
            If (sede = rsEmpl!sedeestrnro) And (sector = rsEmpl!sectorestrnro) And (Mes = rsEmpl!pliqmes) And (Anio = rsEmpl!pliqanio) And (ternroAnt <> rsEmpl!Ternro) Then
                cantEmpl = cantEmpl + 1
                ternroAnt = rsEmpl!Ternro
            End If
            rsEmpl.MoveNext
            
            porcentaje = porcentaje + porc
        
            If Int(porcentaje) > 100 Then
               'Inserto progreso
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = 100"
                sqlAux = sqlAux & " WHERE bpronro = " & NroProcesoBatch
            Else
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = " & Replace(porcentaje, ",", ".")
                sqlAux = sqlAux & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                sqlAux = sqlAux & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
            End If
            objConn.Execute sqlAux, , adExecuteNoRecords
             
             
             
             
        Loop
        StrSql = "UPDATE rep_costo_HS_Extras_det"
        StrSql = StrSql & " SET cantEmp=" & cantEmpl
        StrSql = StrSql & " WHERE sede=" & sede & " AND sector= " & sector & " AND anio=" & Anio & " AND mes=" & Mes & "AND bpronro=" & NroProcesoBatch
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'porc = 20 / ((cantEmpl - 1) * total)
        'Flog.writeline "Porcentaje por Sede 3 " & (CDbl(cantEmpl - 1) * CDbl(total))
        
        If (CDbl(cantEmpl - 1) * CDbl(total)) > 0 Then
            'porc = 20 / (CDbl(cantEmpl - 1) * CDbl(total)) mdf 29/05/2015
            ' porc = 1 / (CDbl(cantEmpl - 1) * CDbl(total)) 'mdf 29/05/2015
        Else
            porc = 0
        End If
        
        porcentaje = porcentaje + porc
        

        If Int(porcentaje) > 100 Then
           'Inserto progreso
            sqlAux = "UPDATE batch_proceso SET bprcprogreso = 100"
            sqlAux = sqlAux & " WHERE bpronro = " & NroProcesoBatch
        Else
            'Actualizo el estado del proceso
            TiempoAcumulado = GetTickCount
            sqlAux = "UPDATE batch_proceso SET bprcprogreso = " & Replace(porcentaje, ",", ".")
            sqlAux = sqlAux & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
            sqlAux = sqlAux & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
        End If
        objConn.Execute sqlAux, , adExecuteNoRecords
        
        rs.MoveNext
    Loop
   

    rs.Close

    '________________________________________________________________
    
    'CALCULO LOS ACU Y CONCEPTOS PARA CADA REGISTRO__________________
    Dim bpronro As Long
    Dim Proceso As Integer
    StrSql = "SELECT * FROM rep_costo_HS_Extras_det "
    StrSql = StrSql & " WHERE bpronro=" & NroProcesoBatch
    StrSql = StrSql & " ORDER BY mes "
    OpenRecordset StrSql, rs
    Flog.writeline "datos reporte tercera parte"
      Do While Not rs.EOF

        sede = rs!sede
        sector = rs!sector
        Anio = rs!Anio
        Mes = rs!Mes
        bpronro = rs!bpronro
        rsEmpl.MoveFirst
        valorCantidadA = 0
        valorMontoA = 0
        valorCantidadB = 0
        valorMontoB = 0
        valorAcuA = 0
        total = rs.RecordCount
        'porc = 10 / (totalEmpleados * total)
   
        
        If (CDbl(totalEmpleados) * CDbl(total)) > 0 Then
            'porc = 10 / (CDbl(totalEmpleados) * CDbl(total)) mdf 29/05/2015
             porc = 3 / (CDbl(totalEmpleados) * CDbl(total)) 'MDF 29/05/2015
        Else
            porc = 0
        End If
        

        Do While Not rsEmpl.EOF
            If (sede = rsEmpl!sedeestrnro) And (sector = rsEmpl!sectorestrnro) And (Anio = rsEmpl!pliqanio) And (Mes = rsEmpl!pliqmes) Then
                Proceso = rsEmpl!pronro
                'busco la cantidad de concepto a
                If esconcA Then
                    sql = "SELECT sum(dlicant) cantidad FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND concepto.conccod IN(" & concA & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not EsNulo(objRs!Cantidad) Then
                            valorCantidadA = valorCantidadA + objRs!Cantidad
                        Else
                            valorCantidadA = valorCantidadA
                        End If
                    End If
                    
                Else
                    sql = "SELECT sum(alcant) cantidad FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND acu_liq.acunro IN(" & concA & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Cantidad) Then
                            valorCantidadA = valorCantidadA + objRs!Cantidad
                        Else
                            valorCantidadA = valorCantidadA
                        End If
                    End If
                End If
                
                'busco el monto del concepto a
                If esconcMontoA Then
                    sql = "SELECT sum(dlimonto) monto FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND concepto.conccod IN(" & montoA & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Monto) Then
                            valorMontoA = valorMontoA + objRs!Monto
                        Else
                            valorMontoA = valorMontoA
                        End If
                    End If
                    
                Else
                    sql = "SELECT sum(almonto) monto FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND acu_liq.acunro IN(" & montoA & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Monto) Then
                            valorMontoA = valorMontoA + objRs!Monto
                        Else
                            valorMontoA = valorMontoA
                        End If
                    End If
                End If
                
                'busco la cantidad del concepto B
                If esconcB Then
                    sql = "SELECT sum(dlicant) cantidad FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND concepto.conccod IN(" & cantidadB & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Cantidad) Then
                            valorCantidadB = valorCantidadB + objRs!Cantidad
                        Else
                            valorCantidadB = valorCantidadB
                        End If
                    End If
                    
                Else
                    sql = "SELECT sum(alcant) cantidad FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND acu_liq.acunro IN(" & cantidadB & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Cantidad) Then
                            valorCantidadB = valorCantidadB + objRs!Cantidad
                        Else
                            valorCantidadB = valorCantidadB
                        End If
                    End If
                End If
                
                'busco el monto del concepto B
                If esconcMontoB Then
                    sql = "SELECT sum(dlimonto) monto FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND concepto.conccod IN(" & montoB & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Monto) Then
                            valorMontoB = valorMontoB + objRs!Monto
                        Else
                            valorMontoB = valorMontoB
                        End If
                    End If
                    
                Else
                    sql = "SELECT sum(almonto) monto FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND acu_liq.acunro IN(" & montoB & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Monto) Then
                            valorMontoB = valorMontoB + objRs!Monto
                        Else
                            valorMontoB = valorMontoB
                        End If
                    End If
                End If
                
                'busco el valor del acumulador A
                If esconcAcuA Then
                    sql = "SELECT dlimonto FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND concepto.conccod =" & acuA
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        valorAcuA = valorAcuA + objRs!dlimonto
                    End If
                    
                Else
                    sql = "SELECT almonto FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro =" & Proceso & " And Empleado = " & rsEmpl!Ternro
                    sql = sql & " AND acu_liq.acunro =" & acuA
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        valorAcuA = valorAcuA + objRs!almonto
                    End If
                End If
                'hasta aca
            End If

            rsEmpl.MoveNext
            
            total = total - 1
            
            'porc = 10 / (totalEmpleados * total)
            
            If (CDbl(totalEmpleados) * CDbl(total)) > 0 Then
                'porc = 10 / (CDbl(totalEmpleados) * CDbl(total)) MDF 29/05/2015
                porc = 2 / (CDbl(totalEmpleados) * CDbl(total))
            Else
                porc = 0
            End If
            
            porcentaje = porcentaje + porc


            If Int(porcentaje) > 100 Then
                'Inserto progreso
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = 100"
                sqlAux = sqlAux & " WHERE bpronro = " & NroProcesoBatch
            Else
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = " & Replace(porcentaje, ",", ".")
                sqlAux = sqlAux & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                sqlAux = sqlAux & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
            End If
            objConn.Execute sqlAux, , adExecuteNoRecords
        Loop
        rs.MoveNext
        'actualizo los valores de la tabla det
        StrSqlAux = "UPDATE rep_costo_HS_Extras_det"
        StrSqlAux = StrSqlAux & " SET cantidad1=" & Replace(valorCantidadA, ",", ".")
        StrSqlAux = StrSqlAux & " , monto1=" & Replace(valorMontoA, ",", ".")
        StrSqlAux = StrSqlAux & " , cantidad2=" & Replace(valorCantidadB, ",", ".")
        StrSqlAux = StrSqlAux & " , monto2=" & Replace(valorMontoB, ",", ".")
        StrSqlAux = StrSqlAux & " , monto3=" & Replace(valorAcuA, ",", ".")
        StrSqlAux = StrSqlAux & " WHERE sede=" & sede & " AND sector= " & sector & " AND anio=" & Anio & " AND mes=" & Mes & " AND bpronro=" & NroProcesoBatch
        objConn.Execute StrSqlAux, , adExecuteNoRecords
        'hasta aca
        
        total = total - 1
            
        'porc = 10 / (totalEmpleados * total)
        
        If (CDbl(totalEmpleados) * CDbl(total)) > 0 Then
            'porc = 10 / (CDbl(totalEmpleados) * CDbl(total))
            porc = 1 / (CDbl(totalEmpleados) * CDbl(total))
        Else
            porc = 0

        End If
        
        porcentaje = porcentaje + porc
        
        If Int(porcentaje) > 100 Then
           'Inserto progreso
           sqlAux = "UPDATE batch_proceso SET bprcprogreso = 100"
           sqlAux = sqlAux & " WHERE bpronro = " & NroProcesoBatch
        Else
           'Actualizo el estado del proceso
           TiempoAcumulado = GetTickCount
           sqlAux = "UPDATE batch_proceso SET bprcprogreso = " & Replace(porcentaje, ",", ".")
           sqlAux = sqlAux & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
           sqlAux = sqlAux & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
        End If
        objConn.Execute sqlAux, , adExecuteNoRecords
    Loop

    Flog.writeline "Fin tercera parte"
    '________________________________________________________________
    
    Call graficos
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline
    Flog.writeline " Error(2): " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
    Flog.writeline
    
    
End Sub

Sub graficos()
'variables para graficos
Dim pliqdesdeinicial
Dim pliqdesdehasta
Dim mesinicial
Dim mesfinal
Dim anioInicial
Dim aniofinal
Dim MesDesde
Dim fechadesde1
Dim fechahasta1
Dim listaProc
Dim rsEmpl As New ADODB.Recordset
Dim i
'Dim arrEmpleados(2000) As String
Dim sedeAnt
'Dim arrSede(2000) As Integer
Dim sectorAnt
'Dim arrSector(2000) As Integer
'Dim arrAnio(2000) As Integer
'Dim arrMeses(2000) As Integer

Dim arrEmpleados()
Dim arrSede()
Dim arrSector()
Dim arrAnio()
Dim arrMeses()

Dim valorCantidadA
Dim valorMontoA
Dim valorCantidadB
Dim valorMontoB
Dim valorAcuA
Dim sql As String
Dim StrSqlAux As String
Dim sqlAux As String
'Dim Total As Integer
Dim total As Double

If porcentaje > 100 Then
 Flog.writeline "No da el tiempo........ :("
 Exit Sub
End If
On Error GoTo CE

'valores para graficos
pliqdesdeinicial = fechadesde
'pliqdesdehasta = fechadesde - 1
pliqdesdehasta = DateAdd("m", -1, fechadesde)
mesinicial = Month(pliqdesdeinicial)
mesfinal = Month(pliqdesdehasta)
MesDesde = Month(fechadesde)
If mesfinal < MesDesde Then
    anioInicial = Year(DateAdd("yyyy", -1, pliqdesdeinicial))
Else
    anioInicial = Year(pliqdesdeinicial)
End If
'anioInicial = Year(pliqdesdeinicial)
If mesfinal >= mesinicial Then
    anioInicial = anioInicial - 1
End If
aniofinal = Year(pliqdesdeinicial)

fechadesde1 = 1 & "/" & mesinicial & "/" & anioInicial
fechahasta1 = 1 & "/" & mesinicial & "/" & aniofinal
'busco los periodos y procesos entre esas fechas
StrSql = "SELECT periodo.pliqnro,pliqmes,pliqanio,proceso.pronro "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " INNER JOIN proceso on proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " WHERE ((pliqdesde >=" & ConvFecha(fechadesde1) & ") and (pliqhasta<=" & ConvFecha(fechahasta1) & "))"
OpenRecordset StrSql, objRs
listaProc = 0
Do While Not objRs.EOF
    listaProc = listaProc & "," & objRs!pronro
objRs.MoveNext
Loop
'hasta aca

Call CargarEmpleadosGraficos(listaProc, rsEmpl)

'codigo copiado
    '_______________________ARREGLO DE EMPLEADOS_____________________
    'buscar los distintos
    i = 0
    Dim emplAnt
    emplAnt = ""
    sedeAnt = ""
    sectorAnt = ""
'    Do While Not rsEmpl.EOF
'        Flog.writeline "Empleado: " & rsEmpl!empleg
'        If (emplAnt <> rsEmpl!empleg) Then
'            I = I + 1
'            'ReDim Preserve arrEmpleados(I)
'            arrEmpleados(I) = rsEmpl!empleg & "@" & rsEmpl!Ternro & "@"
'            emplAnt = rsEmpl!empleg
'        'Else
'
'        End If
'        rsEmpl.MoveNext
'    Loop
    
    Do While Not rsEmpl.EOF
        '-----------------------MDF
        'If (emplAnt <> rsEmpl!empleg) Then
        '    i = i + 1
        '    Flog.writeline "Empleado: " & rsEmpl!empleg
        '    ReDim Preserve arrEmpleados(i)
        '    arrEmpleados(i) = rsEmpl!empleg & "@" & rsEmpl!Ternro & "@"
        '    emplAnt = rsEmpl!empleg
        'Else
            
        'End If
        '------------------------------
        'sedes
        
        If (sedeAnt <> rsEmpl!sedeestrnro) And (CInt(rsEmpl!sedeestrnro) <> 0) Then
            'I = I + 1
            ReDim Preserve arrSede(i)
            Flog.writeline "Sede: " & rsEmpl!sedeestrnro
            arrSede(i) = rsEmpl!sedeestrnro
            sedeAnt = rsEmpl!sedeestrnro
            i = i + 1
        End If
        'sector   05/03/2014
        'If (sectorAnt <> rsEmpl!sectorestrnro) Then
            'I = I + 1 ya estaba
         '   ReDim Preserve arrSector(I)
         '  Flog.writeline "Sector: " & rsEmpl!sectorestrnro
         '   arrSector(I) = rsEmpl!sectorestrnro
         '   sectorAnt = rsEmpl!sectorestrnro
        ' End If
        
        rsEmpl.MoveNext
    Loop
    '________________________HASTA ACA_______________________________
    
    
    'GUARDO EN LA CABECERA DEL REPORTE_______________________________
'    StrSql = " INSERT INTO rep_costo_HS_Extras "
'    StrSql = StrSql & "(bpronro, repdescabr, repdescext, cantEmp)"
'    StrSql = StrSql & " VALUES ( "
'    StrSql = StrSql & NroProcesoBatch & ", "
'    StrSql = StrSql & "'" & tituloFiltro & "'  , " & ""
'    StrSql = StrSql & "'" & tituloFiltro & "'  , " & ""
'    StrSql = StrSql & i & " )"
'    Flog.writeline "SQL INSERT: " & StrSql
'    objConn.Execute StrSql, , adExecuteNoRecords
    '________________________________________________________________
    
    
    '______________________ARREGLO DE SEDES__________________________
'    rsEmpl.MoveFirst
'    sedeAnt = ""
'    I = 0
'    Do While Not rsEmpl.EOF
'        If (sedeAnt <> rsEmpl!sedeestrnro) Then
'            I = I + 1
'           ' ReDim Preserve arrSede(I)
'            arrSede(I) = rsEmpl!sedeestrnro
'            Flog.writeline "sede: " & rsEmpl!sedeestrnro
'            sedeAnt = rsEmpl!sedeestrnro
'        End If
'
'    rsEmpl.MoveNext
'    Loop
    '________________________HASTA ACA_______________________________
    
    
    
    '______________________ARREGLO DE SECTORES______________________
'    rsEmpl.MoveFirst
'    sectorAnt = ""
'    I = 0
'    Do While Not rsEmpl.EOF
'        If (sectorAnt <> rsEmpl!sectorestrnro) Then
'            I = I + 1
'            'ReDim Preserve arrSector(I)
'            arrSector(I) = rsEmpl!sectorestrnro
'            Flog.writeline "sector: " & rsEmpl!sectorestrnro
'            sectorAnt = rsEmpl!sectorestrnro
'        End If
'
'    rsEmpl.MoveNext
'    Loop
    '________________________HASTA ACA_______________________________
    
    
    '______________________ARREGLO DE AÑOS___________________________
    
    'I = 0  05/03/2015
    Dim anioAnt
    'anioAnt = ""
    'Do While Not rsEmpl.EOF
    '    Flog.writeline "AÑO: " & rsEmpl!pliqanio
    '    If (anioAnt <> rsEmpl!pliqanio) Then
    '        I = I + 1
    '        ReDim Preserve arrAnio(I)
    '        arrAnio(I) = rsEmpl!pliqanio
    '        anioAnt = rsEmpl!pliqanio
    '    End If
    '    rsEmpl.MoveNext
   ' Loop 05/03/2015
    
    '________________________HASTA ACA_______________________________
    
    
    
    
    '______________________ARREGLO DE MESES__________________________
    
    'I = 0 05/03/2015
     Dim mesAnt
    'mesAnt = ""
    'Do While Not rsEmpl.EOF
    '    Flog.writeline "Mes: " & rsEmpl!pliqmes
    '    If (mesAnt <> rsEmpl!pliqmes) Then
    '        I = I + 1
    '        ReDim Preserve arrMeses(I)
    '        arrMeses(I) = rsEmpl!pliqmes
    '        mesAnt = rsEmpl!pliqmes
    '    End If
    '    rsEmpl.MoveNext
    'Loop  05/03/2015
    
    '________________________HASTA ACA______________________________
    


    '_____________CICLO POR CADA SEDE Y BUSCO LOS DATOS_____________
    Dim j
    Dim sede As Integer
    Dim sector As Integer
    'Dim sectorAnt
    sectorAnt = ""
    mesAnt = ""
    anioAnt = ""
    j = 0
    Dim SectoresHistorico As String
    SectoresHistorico = "0"
    'porcentaje = 0
    'porc = 5 / (totalEmpleados * (UBound(arrSede) + 1))
    
    Flog.writeline "DATOS ANTIGUOS  primera parte: " & (CDbl(totalEmpleados) * CDbl(UBound(arrSede) + 1))
    If (CDbl(totalEmpleados) * CDbl(UBound(arrSede) + 1)) > 0 Then
        'porc = 5 / (CDbl(totalEmpleados) * CDbl((UBound(arrSede) + 1))) MDF 29/05/2015
        porc = 2 / (CDbl(totalEmpleados) * CDbl((UBound(arrSede) + 1)))
    Else
        porc = 0
    End If
    Flog.writeline "----------------------Sedes historico"
    For j = 0 To UBound(arrSede)
       Flog.writeline arrSede(j)
    Next
    Flog.writeline "----------------------Sedes historico"
    For j = 0 To UBound(arrSede)
        sede = arrSede(j)
        rsEmpl.MoveFirst
        'busco los sectores de la sede
       If CInt(sede) <> 0 Then
        Do While Not rsEmpl.EOF

            If (rsEmpl!sedeestrnro = sede) Then
                If ((CStr(sectorAnt) <> CStr(rsEmpl!sectorestrnro)) Or (CStr(mesAnt) <> CStr(rsEmpl!pliqmes)) Or (CStr(anioAnt) <> CStr(rsEmpl!pliqanio))) Then
                    SectoresHistorico = SectoresHistorico & "," & rsEmpl!sectorestrnro
                    sector = rsEmpl!sectorestrnro
                    StrSql = "INSERT INTO rep_costo_HS_Extras_det "
                    StrSql = StrSql & "(bpronro,tipo,sede,sector,anio,mes)  "
                    StrSql = StrSql & " VALUES ( "
                    StrSql = StrSql & NroProcesoBatch & ", "
                    StrSql = StrSql & 1 & ", "
                    StrSql = StrSql & sede & ", "
                    StrSql = StrSql & sector & ", "
                    StrSql = StrSql & rsEmpl!pliqanio & ", "
                    StrSql = StrSql & rsEmpl!pliqmes & " )"
                    'Flog.writeline StrSql
                    objConn.Execute StrSql, , adExecuteNoRecords
                    sectorAnt = sector
                    mesAnt = rsEmpl!pliqmes
                    anioAnt = rsEmpl!pliqanio
                End If
            End If
            rsEmpl.MoveNext
            
            porcentaje = porcentaje + porc
            
            If Int(porcentaje) > 100 Then
                'Inserto progreso
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = 100"
                sqlAux = sqlAux & " WHERE bpronro = " & NroProcesoBatch
            Else
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = " & Replace(porcentaje, ",", ".")
                sqlAux = sqlAux & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                sqlAux = sqlAux & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
            End If
            'Flog.writeline StrSql
            objConn.Execute sqlAux, , adExecuteNoRecords
        Loop

       End If
    Next
    Flog.writeline "Sectores en el historico a comparar: " & SectoresHistorico
    Flog.writeline "Fin datos antiguos primera parte"
     
    '________________________________________________________________
    
    
    '________calculo la cantidad para cada registro_______
    StrSql = "SELECT sede,sector,anio,mes FROM rep_costo_HS_Extras_det " 'antes *
    StrSql = StrSql & " WHERE bpronro=" & NroProcesoBatch
    StrSql = StrSql & "group by sede, sector, anio, mes, cantEmp "   '<----------------MDF nueva linea
    
    OpenRecordset StrSql, rs
    total = rs.RecordCount
    Dim Anio
    Dim Mes
    Dim cantEmpl
    Dim ternroAnt
    ternroAnt = ""
    'porc = 5 / (totalEmpleados * total)
    
    Flog.writeline " datos antiguos segunda parte " & (CDbl(totalEmpleados) * CDbl(total))
    
    If (CDbl(totalEmpleados) * CDbl(total)) > 0 Then
        porc = 20 / (CDbl(totalEmpleados) * CDbl(total))
    Else
        porc = 0
        'Flog.writeline "EL PORCENTAJE ES 0"
    End If
    Flog.writeline "Primera iteracion: " & rs.RecordCount
    Dim k
    k = 1
    Do While Not rs.EOF
        rsEmpl.MoveFirst
        cantEmpl = 0
        sede = rs!sede
        sector = rs!sector
        Anio = rs!Anio
        Mes = rs!Mes
        ternroAnt = ""
        If k = 1 Then
          Flog.writeline "Segunda iteracion: " & rsEmpl.RecordCount
          k = 2
        End If
        Do While Not rsEmpl.EOF
            
            If (sede = rsEmpl!sedeestrnro) And (sector = rsEmpl!sectorestrnro) And (Mes = rsEmpl!pliqmes) And (Anio = rsEmpl!pliqanio) And (ternroAnt <> rsEmpl!Ternro) Then
                cantEmpl = cantEmpl + 1
                ternroAnt = rsEmpl!Ternro
            End If
            rsEmpl.MoveNext
           ' If porcentaje < 80 Then
            porcentaje = porcentaje + porc
           ' End If
        Loop
            If Int(porcentaje) > 100 Then
                'Inserto progreso
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = 100"
                sqlAux = sqlAux & " WHERE bpronro = " & NroProcesoBatch
               ' Flog.writeline sqlAux
            Else
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = " & Replace(porcentaje, ",", ".")
                sqlAux = sqlAux & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                sqlAux = sqlAux & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
               ' Flog.writeline sqlAux
            End If
            objConn.Execute sqlAux, , adExecuteNoRecords
        
        ' Loop 04/06/2015
      
        StrSql = "UPDATE rep_costo_HS_Extras_det"
        StrSql = StrSql & " SET cantEmp=" & cantEmpl
        StrSql = StrSql & " WHERE sede=" & sede & " AND sector= " & sector & " AND anio=" & Anio & " AND mes=" & Mes & "AND bpronro=" & NroProcesoBatch & " AND tipo=1"
        objConn.Execute StrSql, , adExecuteNoRecords
        rs.MoveNext
    Loop
    rs.Close
    Flog.writeline "fin datos antiguos segunda parte "
    '________________________________________________________________
    
    'CALCULO LOS ACU Y CONCEPTOS PARA CADA REGISTRO__________________
    
    Flog.writeline "calculo de Acumuladores y Conceptos para cada registro"
    Dim bpronro As Long
    Dim Proceso As Integer
    
    Dim Procesos As String
    Dim Empleados As String
    'Dim total As Integer
    
    StrSql = "SELECT sede,sector,anio,mes, bpronro FROM rep_costo_HS_Extras_det "
    StrSql = StrSql & " WHERE bpronro=" & NroProcesoBatch
    StrSql = StrSql & "group by sede, sector, anio, mes,  bpronro"   '<----------------MDF nueva linea
    StrSql = StrSql & " ORDER BY mes "
    OpenRecordset StrSql, rs
    total = rs.RecordCount
    
    'porc = 20 / (totalEmpleados * total)
    Flog.writeline "datos antiguos tercera parte" & CDbl(totalEmpleados) * CDbl(total)
    If (CDbl(totalEmpleados) * CDbl(total)) > 0 Then
        'Flog.writeline "--0000-- "
        porc = 3 / (CDbl(totalEmpleados) * CDbl(total))
    Else
        porc = 0
    End If
    'cantRegistros = CInt(cantRegistros / 2)
    Do While Not rs.EOF
        sede = rs!sede
        sector = rs!sector
        Anio = rs!Anio
        Mes = rs!Mes
        bpronro = rs!bpronro
        rsEmpl.MoveFirst
        valorCantidadA = 0
        valorMontoA = 0
        valorCantidadB = 0
        valorMontoB = 0
        valorAcuA = 0
        
        Procesos = "0" 'MDF
        Empleados = "0" 'MDF
        
        Do While Not rsEmpl.EOF

            If (sede = rsEmpl!sedeestrnro) And (sector = rsEmpl!sectorestrnro) And (Anio = rsEmpl!pliqanio) And (Mes = rsEmpl!pliqmes) Then
                
                Procesos = Procesos & "," & rsEmpl!pronro
                Empleados = Empleados & "," & rsEmpl!Ternro
             
             End If 'MDF
             rsEmpl.MoveNext 'MDF
            If (CDbl(totalEmpleados) * CDbl(total)) > 0 Then
               porc = 20 / (CDbl(totalEmpleados) * CDbl(total))
            Else
                porc = 0
            End If
        Loop
        '----------------------------------------------------
                If esconcA Then
                    sql = "SELECT sum(dlicant) cantidad FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND concepto.conccod IN(" & concA & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not EsNulo(objRs!Cantidad) Then
                            
                            valorCantidadA = valorCantidadA + objRs!Cantidad
                        Else
                            valorCantidadA = valorCantidadA
                        End If
                    End If
                    
                Else
                    sql = "SELECT sum(alcant) cantidad FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND acu_liq.acunro IN(" & concA & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Cantidad) Then
                            valorCantidadA = valorCantidadA + objRs!Cantidad
                        Else
                            valorCantidadA = valorCantidadA
                        End If
                    End If
                End If
                
                'busco el monto del concepto a
                If esconcMontoA Then
                    sql = "SELECT sum(dlimonto) monto FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND concepto.conccod IN(" & montoA & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Monto) Then
                            valorMontoA = valorMontoA + objRs!Monto
                        Else
                             valorMontoA = valorMontoA
                        End If
                    End If
                    
                Else
                    sql = "SELECT sum(almonto) monto FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND acu_liq.acunro IN(" & montoA & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Monto) Then
                            valorMontoA = valorMontoA + objRs!Monto
                        Else
                            valorMontoA = valorMontoA
                        End If
                    End If
                End If
                
                'busco la cantidad del concepto B
                If esconcB Then
                    sql = "SELECT sum(dlicant) cantidad FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND concepto.conccod IN(" & cantidadB & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Cantidad) Then
                            valorCantidadB = valorCantidadB + objRs!Cantidad
                        Else
                            valorCantidadB = valorCantidadB
                        End If
                    End If
                    
                Else
                    sql = "SELECT sum(alcant) cantidad FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND acu_liq.acunro IN(" & cantidadB & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Cantidad) Then
                            valorCantidadB = valorCantidadB + objRs!Cantidad
                        Else
                            valorCantidadB = valorCantidadB
                        End If
                    End If
                End If
                
                'busco el monto del concepto B
                If esconcMontoB Then
                    sql = "SELECT sum(dlimonto) monto FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND concepto.conccod IN(" & montoB & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Monto) Then
                            valorMontoB = valorMontoB + objRs!Monto
                        Else
                            valorMontoB = valorMontoB
                        End If
                    End If
                    
                Else
                    sql = "SELECT sum(almonto) monto FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND acu_liq.acunro IN(" & montoB & ")"
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        If Not IsNull(objRs!Monto) Then
                            valorMontoB = valorMontoB + objRs!Monto
                        Else
                            valorMontoB = valorMontoB
                        End If
                    End If
                End If
                
                'busco el valor del acumulador A
                If esconcAcuA Then
                    sql = "SELECT dlimonto FROM cabliq "
                    sql = sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                    sql = sql & " INNER JOIN concepto ON concepto.concnro=detliq.concnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND concepto.conccod =" & acuA
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        valorAcuA = valorAcuA + objRs!dlimonto
                    End If
                    
                Else
                    sql = "SELECT almonto FROM cabliq "
                    sql = sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    sql = sql & " WHERE pronro in (" & Procesos & ") And Empleado in (" & Empleados & ")"
                    sql = sql & " AND acu_liq.acunro =" & acuA
                    OpenRecordset sql, objRs
                    If Not objRs.EOF Then
                        valorAcuA = valorAcuA + objRs!almonto
                    End If
                End If
                
                'hasta aca
       '-----------------------------------------MDF inicio comentada
       '     End If
       '     rsEmpl.MoveNext

       '     'progreso seba 01/08/2012
            
       '     'porcentaje = porcentaje + porc
       '     'Flog.writeline "--19-- "
       '
       '
       '     'porc = 20 / (totalEmpleados * total)
       '
       '     'Flog.writeline "Porcentaje por Sede 10" & (CDbl(totalEmpleados) * CDbl(total))
       '     If (CDbl(totalEmpleados) * CDbl(total)) > 0 Then
       '         'Flog.writeline "--20-- "
       '         porc = 20 / (CDbl(totalEmpleados) * CDbl(total))
       '     Else
       '         porc = 0
       '         'Flog.writeline "EL PORCENTAJE ES 0"
       '     End If
       '   Loop
       '-----------------------------------------MDF Fin comentada
            'Redondeo a 100%
            If Int(porcentaje) > 100 Then
                'Inserto progreso
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = 100"
                sqlAux = sqlAux & " WHERE bpronro = " & NroProcesoBatch
            Else
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                sqlAux = "UPDATE batch_proceso SET bprcprogreso = " & Replace(porcentaje, ",", ".")
                sqlAux = sqlAux & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                sqlAux = sqlAux & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
            End If
            'Flog.writeline sqlAux
            objConn.Execute sqlAux, , adExecuteNoRecords
        'Loop MDF 04/06/2015
        rs.MoveNext
        'actualizo los valores de la tabla det
        StrSqlAux = "UPDATE rep_costo_HS_Extras_det"
        StrSqlAux = StrSqlAux & " SET cantidad1=" & Replace(valorCantidadA, ",", ".")
        StrSqlAux = StrSqlAux & " , monto1=" & Replace(valorMontoA, ",", ".")
        StrSqlAux = StrSqlAux & " , cantidad2=" & Replace(valorCantidadB, ",", ".")
        StrSqlAux = StrSqlAux & " , monto2=" & Replace(valorMontoB, ",", ".")
        StrSqlAux = StrSqlAux & " , monto3=" & Replace(valorAcuA, ",", ".")
        StrSqlAux = StrSqlAux & " WHERE sede=" & sede & " AND sector= " & sector & " AND anio=" & Anio & " AND mes=" & Mes & " AND bpronro=" & NroProcesoBatch & "AND tipo=1"
       
        objConn.Execute StrSqlAux, , adExecuteNoRecords
         porcentaje = porcentaje + porc
        'End If
        If Int(porcentaje) > 100 Then
           'Inserto progreso
           sqlAux = "UPDATE batch_proceso SET bprcprogreso = 100"
           sqlAux = sqlAux & " WHERE bpronro = " & NroProcesoBatch
        Else
           'Actualizo el estado del proceso
           TiempoAcumulado = GetTickCount
           sqlAux = "UPDATE batch_proceso SET bprcprogreso = " & Replace(porcentaje, ",", ".")
           sqlAux = sqlAux & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
           sqlAux = sqlAux & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch
        End If
       ' objConn.Execute sqlAux, , adExecuteNoRecords   MDF 04/06/2016
        total = total - 1
    Loop
'hasta aca
    

Flog.writeline "fin datos antiguos tercera parte"
Flog.writeline "FIN PROCESO CON EXITO"
Flog.writeline "Hora de finalizacion: " & Now()
Exit Sub
CE:
    HuboErrores = True
    Flog.writeline
    Flog.writeline " Error(3): " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
    Flog.writeline
    

End Sub
'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String



If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
        StrEmpl = " SELECT DISTINCT empleado.ternro, empleg, terape, estact1.tenro   tenro1, estact1.estrnro   estrnro1, " & _
        " estact2.tenro  tenro2, estact2.estrnro  estrnro2, estact3.tenro  tenro3, estact3.estrnro  estrnro3, nestact1.estrdabr   nestr1, nestact2.estrdabr   nestr2, nestact3.estrdabr   nestr3  "
        
        'If ((detalle <> "0") And (detalle <> "")) Then
        StrEmpl = StrEmpl & " , sector.tenro sectortenro, sector.estrnro sectorestrnro "
        'End If
        
        'If ((agrupado <> "0") And (agrupado <> "")) Then
        StrEmpl = StrEmpl & " , sede.tenro sedetenro, sede.estrnro sedeestrnro "
        'End If
        
        StrEmpl = StrEmpl & " , periodo.pliqmes,periodo.pliqanio,proceso.pronro "
        
        StrEmpl = StrEmpl & " FROM cabliq "
        
        'seba
        StrEmpl = StrEmpl & " INNER JOIN proceso on proceso.pronro=cabliq.pronro "
        StrEmpl = StrEmpl & " INNER JOIN periodo on periodo.pliqnro=proceso.pliqnro "
        'hasta aca
        
        
        If listapronro = "" Or listapronro = "0" Then
            StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
        Else
            StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
        End If
         
        StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
        StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If empresa > 0 Then
            StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
        End If
        
        StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
        StrEmpl = StrEmpl & " AND (estact1.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If estrnro1 <> 0 Then   'cuando se le asigna un valor al nivel 1
            StrEmpl = StrEmpl & " AND estact1.estrnro =" & estrnro1
        End If
         
        StrEmpl = StrEmpl & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
        StrEmpl = StrEmpl & " AND (estact2.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If estrnro2 <> 0 Then   'cuando se le asigna un valor al nivel 2
            StrEmpl = StrEmpl & " AND estact2.estrnro =" & estrnro2
        End If
         
        StrEmpl = StrEmpl & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3 & _
        " AND (estact3.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If estrnro3 <> 0 Then   'cuando se le asigna un valor al nivel 3
            StrEmpl = StrEmpl & " AND estact3.estrnro =" & estrnro3
        End If
        
        If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
            StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro AND ter_doc.tidnro=10"
        End If
         
        StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro "
        StrEmpl = StrEmpl & " INNER JOIN estructura nestact2 ON nestact2.estrnro = estact2.estrnro "
        StrEmpl = StrEmpl & " INNER JOIN estructura nestact3 ON nestact3.estrnro = estact3.estrnro "
         
        'filtro agregado por sebastian stremel verifica
        StrEmpl = StrEmpl & " INNER JOIN his_estructura sector ON his_estructura.ternro = empleado.ternro "
        StrEmpl = StrEmpl & " INNER JOIN his_estructura sede ON his_estructura.ternro = empleado.ternro"
        
        StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro =" & estrdet
        StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If ((detalle <> "0") And (detalle <> "")) Then
            StrEmpl = StrEmpl & " AND sector.estrnro = " & detalle
        End If
        
        
        StrEmpl = StrEmpl & " AND sede.ternro = empleado.ternro AND sede.tenro =" & estragrup
        StrEmpl = StrEmpl & " AND (sede.htetdesde<=" & ConvFecha(fecEstr) & " AND (sede.htethasta is null OR sede.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If ((agrupado <> "0") And (agrupado <> "")) Then
            StrEmpl = StrEmpl & " AND sede.estrnro = " & agrupado
        End If
        
        'hasta aca
    
         
        StrEmpl = StrEmpl & " WHERE " & filtro
        'If ((agrupado <> "0") And (agrupado <> "")) Then
        StrEmpl = StrEmpl & " ORDER BY sedetenro,sedeestrnro,sectortenro,sectorestrnro," & l_orden
        'Else
        

Else
        If tenro2 <> 0 Then   ' ocurre cuando se selecciono hasta el segundo nivel
            
            StrEmpl = "SELECT DISTINCT empleado.ternro, empleg, terape, estact1.tenro   tenro1, estact1.estrnro   estrnro1, " & _
                " estact2.tenro   tenro2, estact2.estrnro   estrnro2, nestact1.estrdabr   nestr1, nestact2.estrdabr   nestr2 "
            
            'If ((detalle <> "0") And (detalle <> "")) Then
            StrEmpl = StrEmpl & " , sector.tenro sectortenro, sector.estrnro sectorestrnro "
            'End If
            
            'If ((agrupado <> "0") And (agrupado <> "")) Then
            StrEmpl = StrEmpl & " , sede.tenro sedetenro, sede.estrnro sedeestrnro "
            'End If
            
            StrEmpl = StrEmpl & " , periodo.pliqmes,periodo.pliqanio,proceso.pronro "
            
            StrEmpl = StrEmpl & " FROM cabliq "
            
            'seba
            StrEmpl = StrEmpl & " INNER JOIN proceso on proceso.pronro=cabliq.pronro "
            StrEmpl = StrEmpl & " INNER JOIN periodo on periodo.pliqnro=proceso.pliqnro "
            'hasta aca
            
            If listapronro = "" Or listapronro = "0" Then
                StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
            Else
                StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
            End If
            
            StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
            StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
        
            If empresa > 0 Then
             StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
            End If
        
            StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
            StrEmpl = StrEmpl & " AND (estact1.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
            
            If estrnro1 <> 0 Then
                 StrEmpl = StrEmpl & " AND estact1.estrnro =" & estrnro1
            End If
            
            StrEmpl = StrEmpl & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & _
            " AND (estact2.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecEstr) & "))"
            
            If estrnro2 <> 0 Then
                StrEmpl = StrEmpl & " AND estact2.estrnro =" & estrnro2
            End If
            
            If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro AND ter_doc.tidnro=10"
            End If
            
            StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro "
            StrEmpl = StrEmpl & " INNER JOIN estructura nestact2 ON nestact2.estrnro = estact2.estrnro "
            
            'filtro agregado por sebastian stremel verifica
            StrEmpl = StrEmpl & " INNER JOIN his_estructura sector ON his_estructura.ternro = empleado.ternro "
            StrEmpl = StrEmpl & " INNER JOIN his_estructura sede ON his_estructura.ternro = empleado.ternro"
            
            StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro =" & estrdet
            StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
            
            
            If ((detalle <> "0") And (detalle <> "")) Then
                StrEmpl = StrEmpl & " AND sector.estrnro = " & detalle
            End If
            
            
            StrEmpl = StrEmpl & " AND sede.ternro = empleado.ternro AND sede.tenro = " & estragrup
            StrEmpl = StrEmpl & " AND (sede.htetdesde<=" & ConvFecha(fecEstr) & " AND (sede.htethasta is null OR sede.htethasta>=" & ConvFecha(fecEstr) & "))"
            
            If ((agrupado <> "0") And (agrupado <> "")) Then
                StrEmpl = StrEmpl & " AND sede.estrnro = " & agrupado
            End If
            
            'hasta aca
            
            StrEmpl = StrEmpl & " WHERE " & filtro
            StrEmpl = StrEmpl & " ORDER BY sedetenro,sedeestrnro,sectortenro,sectorestrnro," & l_orden

        Else
            If tenro1 <> 0 Then   ' Cuando solo selecionamos el primer nivel
                StrEmpl = "SELECT DISTINCT empleado.ternro, empleg, terape, estact1.tenro   tenro1, estact1.estrnro   estrnro1, nestact1.estrdabr   nestr1"
                
                'If ((detalle <> "0") And (detalle <> "")) Then
                StrEmpl = StrEmpl & " , sector.tenro sectortenro, sector.estrnro sectorestrnro "
                'End If
                
                'If ((agrupado <> "0") And (agrupado <> "")) Then
                StrEmpl = StrEmpl & " , sede.tenro sedetenro, sede.estrnro sedeestrnro "
                'End If
                
                StrEmpl = StrEmpl & " , periodo.pliqmes,periodo.pliqanio,proceso.pronro "
                
                StrEmpl = StrEmpl & " FROM cabliq "
                
                'seba
                StrEmpl = StrEmpl & " INNER JOIN proceso on proceso.pronro=cabliq.pronro "
                StrEmpl = StrEmpl & " INNER JOIN periodo on periodo.pliqnro=proceso.pliqnro "
                'hasta aca
                
                If listapronro = "" Or listapronro = "0" Then
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
                Else
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
                End If
                
                StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
                StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
        
                If empresa > 0 Then
                 StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
                End If
        
                StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                StrEmpl = StrEmpl & " AND (estact1.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If estrnro1 <> 0 Then
                    StrEmpl = StrEmpl & " AND estact1.estrnro =" & estrnro1
                End If
                
                If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                    StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro AND ter_doc.tidnro=10"
                End If
                
                StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro "
                
                'filtro agregado por sebastian stremel verifica
                StrEmpl = StrEmpl & " INNER JOIN his_estructura sector ON his_estructura.ternro = empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN his_estructura sede ON his_estructura.ternro = empleado.ternro"
                
                StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro =" & estrdet
                StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If ((detalle <> "0") And (detalle <> "")) Then
                    StrEmpl = StrEmpl & " AND sector.estrnro = " & detalle
                End If
                
                StrEmpl = StrEmpl & " AND sede.ternro = empleado.ternro AND sede.tenro = " & estragrup
                StrEmpl = StrEmpl & " AND (sede.htetdesde<=" & ConvFecha(fecEstr) & " AND (sede.htethasta is null OR sede.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If ((agrupado <> "0") And (agrupado <> "")) Then
                    StrEmpl = StrEmpl & " AND sede.estrnro = " & agrupado
                End If
                
                'hasta aca
                
                
                
                StrEmpl = StrEmpl & " WHERE " & filtro
                StrEmpl = StrEmpl & " ORDER BY sedetenro,sedeestrnro,sectortenro,sectorestrnro," & l_orden

            Else ' cuando no hay nivel de estructura seleccionado
                StrEmpl = " SELECT DISTINCT empleado.ternro, empleg, terape,his_estructura.tenro emp "
                'If ((detalle <> "0") And (detalle <> "")) Then
                StrEmpl = StrEmpl & " , sector.tenro sectortenro, sector.estrnro sectorestrnro "
                'End If
                
                'If ((agrupado <> "0") And (agrupado <> "")) Then
                StrEmpl = StrEmpl & " , sede.tenro sedetenro, sede.estrnro sedeestrnro "
                'End If
                
                StrEmpl = StrEmpl & " , periodo.pliqmes,periodo.pliqanio,proceso.pronro "
                
                StrEmpl = StrEmpl & " FROM cabliq "
                
                'seba
                StrEmpl = StrEmpl & " INNER JOIN proceso on proceso.pronro=cabliq.pronro "
                StrEmpl = StrEmpl & " INNER JOIN periodo on periodo.pliqnro=proceso.pliqnro "
                'hasta aca
                
                If listapronro = "" Or listapronro = "0" Then
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
                Else
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
                End If
                
                StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
                StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
        
                If empresa > 0 Then
                 StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
                End If
                                        
                If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                    StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro AND ter_doc.tidnro=10"
                End If
                
                'filtro agregado por sebastian stremel verifica
                StrEmpl = StrEmpl & " INNER JOIN his_estructura sector ON his_estructura.ternro = empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN his_estructura sede ON his_estructura.ternro = empleado.ternro"
                StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro =  " & estrdet
                StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If ((detalle <> "0") And (detalle <> "")) Then
'                    StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro = 2 "
'                    StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
                    StrEmpl = StrEmpl & " AND sector.estrnro = " & detalle
                End If
                
                StrEmpl = StrEmpl & " AND sede.ternro = empleado.ternro AND sede.tenro =  " & estragrup
                StrEmpl = StrEmpl & " AND (sede.htetdesde<=" & ConvFecha(fecEstr) & " AND (sede.htethasta is null OR sede.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If ((agrupado <> "0") And (agrupado <> "")) Then
                    StrEmpl = StrEmpl & " AND sede.estrnro = " & agrupado
                End If
                
                'hasta aca

                
                StrEmpl = StrEmpl & " WHERE " & filtro
                
                StrEmpl = StrEmpl & " ORDER BY sedetenro,sedeestrnro,sectortenro,sectorestrnro," & l_orden
            End If
        
        End If

End If
    
Flog.writeline "SQL :" & StrEmpl
   
OpenRecordset StrEmpl, rsEmpl

'----------------------------MDF---------------
Dim anterior
Dim cont
cont = 0
anterior = "0"
Do While Not rsEmpl.EOF
 If anterior <> rsEmpl("ternro") Then
  cont = cont + 1
  anterior = rsEmpl("ternro")
 End If
 rsEmpl.MoveNext
Loop

'----------------------------------------------
'cantRegistros = rsEmpl.RecordCount
 rsEmpl.MoveFirst
cantRegistros = cont
totalEmpleados = cantRegistros
    

    
    
End Sub


'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleadosGraficos(ByVal listapronro, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String



If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
        StrEmpl = " SELECT DISTINCT empleado.ternro, empleg, terape, estact1.tenro AS tenro1, estact1.estrnro AS estrnro1, " & _
        " estact2.tenro AS tenro2, estact2.estrnro AS estrnro2, estact3.tenro AS tenro3, estact3.estrnro AS estrnro3, nestact1.estrdabr AS nestr1, nestact2.estrdabr AS nestr2, nestact3.estrdabr AS nestr3  "
        
        'If ((detalle <> "0") And (detalle <> "")) Then
        StrEmpl = StrEmpl & " , sector.tenro sectortenro, sector.estrnro sectorestrnro "
        'End If
        
        'If ((agrupado <> "0") And (agrupado <> "")) Then
        StrEmpl = StrEmpl & " , sede.tenro sedetenro, sede.estrnro sedeestrnro "
        'End If
        
        StrEmpl = StrEmpl & " , periodo.pliqmes,periodo.pliqanio,proceso.pronro "
        
        StrEmpl = StrEmpl & " FROM cabliq "
        
        'seba
        StrEmpl = StrEmpl & " INNER JOIN proceso on proceso.pronro=cabliq.pronro "
        StrEmpl = StrEmpl & " INNER JOIN periodo on periodo.pliqnro=proceso.pliqnro "
        'hasta aca
        
        
        If listapronro = "" Or listapronro = "0" Then
            StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
        Else
            StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
        End If
         
        StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
        StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If empresa > 0 Then
            StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
        End If
        
        StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
        StrEmpl = StrEmpl & " AND (estact1.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If estrnro1 <> 0 Then   'cuando se le asigna un valor al nivel 1
            StrEmpl = StrEmpl & " AND estact1.estrnro =" & estrnro1
        End If
         
        StrEmpl = StrEmpl & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
        StrEmpl = StrEmpl & " AND (estact2.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If estrnro2 <> 0 Then   'cuando se le asigna un valor al nivel 2
            StrEmpl = StrEmpl & " AND estact2.estrnro =" & estrnro2
        End If
         
        StrEmpl = StrEmpl & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3 & _
        " AND (estact3.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If estrnro3 <> 0 Then   'cuando se le asigna un valor al nivel 3
            StrEmpl = StrEmpl & " AND estact3.estrnro =" & estrnro3
        End If
        
        If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
            StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro AND ter_doc.tidnro=10"
        End If
         
        StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro "
        StrEmpl = StrEmpl & " INNER JOIN estructura nestact2 ON nestact2.estrnro = estact2.estrnro "
        StrEmpl = StrEmpl & " INNER JOIN estructura nestact3 ON nestact3.estrnro = estact3.estrnro "
         
        'filtro agregado por sebastian stremel verifica
        StrEmpl = StrEmpl & " INNER JOIN his_estructura sector ON his_estructura.ternro = empleado.ternro "
        StrEmpl = StrEmpl & " INNER JOIN his_estructura sede ON his_estructura.ternro = empleado.ternro"
        
        StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro = " & estrdet
        StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If ((detalle <> "0") And (detalle <> "")) Then
            StrEmpl = StrEmpl & " AND sector.estrnro = " & detalle
        End If
        
        
        StrEmpl = StrEmpl & " AND sede.ternro = empleado.ternro AND sede.tenro = " & estragrup
        StrEmpl = StrEmpl & " AND (sede.htetdesde<=" & ConvFecha(fecEstr) & " AND (sede.htethasta is null OR sede.htethasta>=" & ConvFecha(fecEstr) & "))"
        
        If ((agrupado <> "0") And (agrupado <> "")) Then
            StrEmpl = StrEmpl & " AND sede.estrnro = " & agrupado
        End If
        
        'hasta aca
    
         
        StrEmpl = StrEmpl & " WHERE " & filtro
        'If ((agrupado <> "0") And (agrupado <> "")) Then
        StrEmpl = StrEmpl & " ORDER BY sedetenro,sedeestrnro,sectortenro,sectorestrnro," & l_orden
        'Else
        

Else
        If tenro2 <> 0 Then   ' ocurre cuando se selecciono hasta el segundo nivel
            
            StrEmpl = "SELECT DISTINCT empleado.ternro, empleg, terape, estact1.tenro AS tenro1, estact1.estrnro AS estrnro1, " & _
                " estact2.tenro AS tenro2, estact2.estrnro AS estrnro2, nestact1.estrdabr AS nestr1, nestact2.estrdabr AS nestr2 "
            
            'If ((detalle <> "0") And (detalle <> "")) Then
            StrEmpl = StrEmpl & " , sector.tenro sectortenro, sector.estrnro sectorestrnro "
            'End If
            
            'If ((agrupado <> "0") And (agrupado <> "")) Then
            StrEmpl = StrEmpl & " , sede.tenro sedetenro, sede.estrnro sedeestrnro "
            'End If
            
            StrEmpl = StrEmpl & " , periodo.pliqmes,periodo.pliqanio,proceso.pronro "
            
            StrEmpl = StrEmpl & " FROM cabliq "
            
            'seba
            StrEmpl = StrEmpl & " INNER JOIN proceso on proceso.pronro=cabliq.pronro "
            StrEmpl = StrEmpl & " INNER JOIN periodo on periodo.pliqnro=proceso.pliqnro "
            'hasta aca
            
            If listapronro = "" Or listapronro = "0" Then
                StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
            Else
                StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
            End If
            
            StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
            StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
        
            If empresa > 0 Then
             StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
            End If
        
            StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
            StrEmpl = StrEmpl & " AND (estact1.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
            
            If estrnro1 <> 0 Then
                 StrEmpl = StrEmpl & " AND estact1.estrnro =" & estrnro1
            End If
            
            StrEmpl = StrEmpl & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & _
            " AND (estact2.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecEstr) & "))"
            
            If estrnro2 <> 0 Then
                StrEmpl = StrEmpl & " AND estact2.estrnro =" & estrnro2
            End If
            
            If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro AND ter_doc.tidnro=10"
            End If
            
            StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro "
            StrEmpl = StrEmpl & " INNER JOIN estructura nestact2 ON nestact2.estrnro = estact2.estrnro "
            
            'filtro agregado por sebastian stremel verifica
            StrEmpl = StrEmpl & " INNER JOIN his_estructura sector ON his_estructura.ternro = empleado.ternro "
            StrEmpl = StrEmpl & " INNER JOIN his_estructura sede ON his_estructura.ternro = empleado.ternro"
            
            StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro = " & estrdet
            StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
            
            
            If ((detalle <> "0") And (detalle <> "")) Then
                StrEmpl = StrEmpl & " AND sector.estrnro = " & detalle
            End If
            
            
            StrEmpl = StrEmpl & " AND sede.ternro = empleado.ternro AND sede.tenro = " & estragrup
            StrEmpl = StrEmpl & " AND (sede.htetdesde<=" & ConvFecha(fecEstr) & " AND (sede.htethasta is null OR sede.htethasta>=" & ConvFecha(fecEstr) & "))"
            
            If ((agrupado <> "0") And (agrupado <> "")) Then
                StrEmpl = StrEmpl & " AND sede.estrnro = " & agrupado
            End If
            
            'hasta aca
            
            StrEmpl = StrEmpl & " WHERE " & filtro
            StrEmpl = StrEmpl & " ORDER BY sedetenro,sedeestrnro,sectortenro,sectorestrnro," & l_orden

        Else
            If tenro1 <> 0 Then   ' Cuando solo selecionamos el primer nivel
                StrEmpl = "SELECT DISTINCT empleado.ternro, empleg, terape, estact1.tenro AS tenro1, estact1.estrnro AS estrnro1, nestact1.estrdabr AS nestr1"
                
                'If ((detalle <> "0") And (detalle <> "")) Then
                StrEmpl = StrEmpl & " , sector.tenro sectortenro, sector.estrnro sectorestrnro "
                'End If
                
                'If ((agrupado <> "0") And (agrupado <> "")) Then
                StrEmpl = StrEmpl & " , sede.tenro sedetenro, sede.estrnro sedeestrnro "
                'End If
                
                StrEmpl = StrEmpl & " , periodo.pliqmes,periodo.pliqanio,proceso.pronro "
                
                StrEmpl = StrEmpl & " FROM cabliq "
                
                'seba
                StrEmpl = StrEmpl & " INNER JOIN proceso on proceso.pronro=cabliq.pronro "
                StrEmpl = StrEmpl & " INNER JOIN periodo on periodo.pliqnro=proceso.pliqnro "
                'hasta aca
                
                If listapronro = "" Or listapronro = "0" Then
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
                Else
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
                End If
                
                StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
                StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
        
                If empresa > 0 Then
                 StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
                End If
        
                StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                StrEmpl = StrEmpl & " AND (estact1.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If estrnro1 <> 0 Then
                    StrEmpl = StrEmpl & " AND estact1.estrnro =" & estrnro1
                End If
                
                If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                    StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro AND ter_doc.tidnro=10"
                End If
                
                StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro "
                
                'filtro agregado por sebastian stremel verifica
                StrEmpl = StrEmpl & " INNER JOIN his_estructura sector ON his_estructura.ternro = empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN his_estructura sede ON his_estructura.ternro = empleado.ternro"
                
                StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro = " & estrdet
                StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If ((detalle <> "0") And (detalle <> "")) Then
                    StrEmpl = StrEmpl & " AND sector.estrnro = " & detalle
                End If
                
                StrEmpl = StrEmpl & " AND sede.ternro = empleado.ternro AND sede.tenro = " & estragrup
                StrEmpl = StrEmpl & " AND (sede.htetdesde<=" & ConvFecha(fecEstr) & " AND (sede.htethasta is null OR sede.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If ((agrupado <> "0") And (agrupado <> "")) Then
                    StrEmpl = StrEmpl & " AND sede.estrnro = " & agrupado
                End If
                
                'hasta aca
                
                
                
                StrEmpl = StrEmpl & " WHERE " & filtro
                StrEmpl = StrEmpl & " ORDER BY sedetenro,sedeestrnro,sectortenro,sectorestrnro," & l_orden

            Else ' cuando no hay nivel de estructura seleccionado
                StrEmpl = " SELECT DISTINCT empleado.ternro, empleg, terape,his_estructura.tenro emp "
                'If ((detalle <> "0") And (detalle <> "")) Then
                StrEmpl = StrEmpl & " , sector.tenro sectortenro, sector.estrnro sectorestrnro "
                'End If
                
                'If ((agrupado <> "0") And (agrupado <> "")) Then
                StrEmpl = StrEmpl & " , sede.tenro sedetenro, sede.estrnro sedeestrnro "
                'End If
                
                StrEmpl = StrEmpl & " , periodo.pliqmes,periodo.pliqanio,proceso.pronro "
                
                StrEmpl = StrEmpl & " FROM cabliq "
                
                'seba
                StrEmpl = StrEmpl & " INNER JOIN proceso on proceso.pronro=cabliq.pronro "
                StrEmpl = StrEmpl & " INNER JOIN periodo on periodo.pliqnro=proceso.pliqnro "
                'hasta aca
                
                If listapronro = "" Or listapronro = "0" Then
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado "
                Else
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
                End If
                
                StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
                StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
        
                If empresa > 0 Then
                 StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
                End If
                                        
                If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                    StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro AND ter_doc.tidnro=10"
                End If
                
                'filtro agregado por sebastian stremel verifica
                StrEmpl = StrEmpl & " INNER JOIN his_estructura sector ON his_estructura.ternro = empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN his_estructura sede ON his_estructura.ternro = empleado.ternro"
                StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro = " & estrdet
                StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If ((detalle <> "0") And (detalle <> "")) Then
'                    StrEmpl = StrEmpl & " AND sector.ternro = empleado.ternro AND sector.tenro = 2 "
'                    StrEmpl = StrEmpl & " AND (sector.htetdesde<=" & ConvFecha(fecEstr) & " AND (sector.htethasta is null OR sector.htethasta>=" & ConvFecha(fecEstr) & "))"
                    StrEmpl = StrEmpl & " AND sector.estrnro = " & detalle
                End If
                
                StrEmpl = StrEmpl & " AND sede.ternro = empleado.ternro AND sede.tenro = " & estragrup
                StrEmpl = StrEmpl & " AND (sede.htetdesde<=" & ConvFecha(fecEstr) & " AND (sede.htethasta is null OR sede.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If ((agrupado <> "0") And (agrupado <> "")) Then
                    StrEmpl = StrEmpl & " AND sede.estrnro = " & agrupado
                End If
                
                'hasta aca

                
                StrEmpl = StrEmpl & " WHERE " & filtro
                
                StrEmpl = StrEmpl & " ORDER BY sedetenro,sedeestrnro,sectortenro,sectorestrnro," & l_orden
            End If
        
        End If

End If
    
Flog.writeline "SQL : " & StrEmpl
OpenRecordset StrEmpl, rsEmpl

'--------------------------------------- MDFF
'Dim anterior  05/03/2015
'Dim cont
'cont = 0
'anterior = "0"
'Do While Not rsEmpl.EOF
' If anterior <> rsEmpl("ternro") Then
'  cont = cont + 1
'  anterior = rsEmpl("ternro")
' End If
' rsEmpl.MoveNext
'Loop
'---------------------------------------
' rsEmpl.MoveFirst
'cantRegistros = rsEmpl.RecordCount
'totalEmpleados = cantRegistros
totalEmpleados = rsEmpl.RecordCount
    

    
    
End Sub


