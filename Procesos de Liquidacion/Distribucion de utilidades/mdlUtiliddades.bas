Attribute VB_Name = "mdlUtiliddades"
'Global Const Version = "1.00"
'Global Const FechaModificacion = "30/01/2013"
'Global Const UltimaModificacion = " Version inicial - CAS-16441- H&A - PERU - DISTRIBUCION DE UTILIDADES "

'Global Const Version = "1.01"
'Global Const FechaModificacion = "25/02/2013"
'Global Const UltimaModificacion = " cambio en variable fuera de indice y en parametros que se recuperan. - CAS-16441- H&A - PERU - DISTRIBUCION DE UTILIDADES"

'Global Const Version = "1.02"
'Global Const FechaModificacion = "19/03/2013"
'Global Const UltimaModificacion = " se corrigio variable duplicada y se limpio variable de arregloHoras. - CAS-16441- H&A - PERU - DISTRIBUCION DE UTILIDADES"

'Global Const Version = "1.03"
'Global Const FechaModificacion = "09/04/2013"
'Global Const UltimaModificacion = " Se agrego fecha de generacion  y hora a la tabla liq_emputil - CAS-16441- H&A - PERU - DISTRIBUCION DE UTILIDADES"

Global Const Version = "1.04"
Global Const FechaModificacion = "03/07/2013"
'Global Const UltimaModificacion = " Cambio en la busqueda de base de calculo - Sebastian Stremel - CAS-16441- H&A - PERU - DISTRIBUCION DE UTILIDADES"
'Global Const UltimaModificacion = " Se agrego la busqueda por ACM de de dias trab y rem. - Sebastian Stremel - CAS-16441- H&A - PERU - DISTRIBUCION DE UTILIDADES"


Global empresa
Global pliqdesde
Global pliqhasta
Global fechaproinicio
Global fechaprofin
Global l_fechadesde
Global l_fechahasta
Global arregloEmpleados(13)
Global NroProceso

Private Sub main()

    strcmdline = Command()
    arrparametros = Split(strcmdline, " ", -1)
    If UBound(arrparametros) > 1 Then
        If IsNumeric(arrparametros(0)) Then
            NroProceso = arrparametros(0)
            Etiqueta = arrparametros(1)
            EncriptStrconexion = CBool(arrparametros(2))
            c_seed = arrparametros(2)
        Else
             Exit Sub
        End If
    Else
        If UBound(arrparametros) > 0 Then
            If IsNumeric(arrparametros(0)) Then
                NroProceso = arrparametros(0)
                Etiqueta = arrparametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strcmdline) Then
                NroProceso = strcmdline
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    HuboErrores = False
    
    Nombre_Arch = PathFLog & "RHProUtilidades" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de distribucion de utilidades : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    TiempoInicialProceso = GetTickCount
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo CE
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
        'Obtengo los parametros del proceso
         parametros = objRs!bprcparam
         Flog.writeline " parametros --> " & parametros
         arrparametros = Split(parametros, "@")
         
         Flog.writeline " limite del array --> " & UBound(arrparametros)
         
         'obtengo el numero de estructura de la empresa
         empresa = arrparametros(0)
         
         'obtengo el año que se quier calcular
         Anio = arrparametros(1)
         
         'con el año busco el pliqdesde y el pliqhasta
         StrSql = " SELECT MIN(pliqmes) mesmin, MAX(pliqmes) mesmax "
         StrSql = StrSql & " FROM periodo WHERE pliqanio=" & Anio
         OpenRecordset StrSql, objRs
         If Not objRs.EOF Then
            mesdesde = objRs!mesmin
            meshasta = objRs!mesmax
         End If
         objRs.Close
         
         'busco el pliqdesde
         StrSql = "select pliqnro FROM periodo "
         StrSql = StrSql & "where pliqmes=" & mesdesde & "AND pliqanio=" & Anio
         OpenRecordset StrSql, objRs
         If Not objRs.EOF Then
            pliqdesde = objRs!pliqnro
         End If
         objRs.Close
         
         'busco el pliqhasta
         StrSql = "select pliqnro FROM periodo "
         StrSql = StrSql & "where pliqmes=" & meshasta & "AND pliqanio=" & Anio
         OpenRecordset StrSql, objRs
         If Not objRs.EOF Then
            pliqhasta = objRs!pliqnro
         End If
         objRs.Close
         
         
         'periodo desde
         'pliqdesde = arrparametros(1)
         
         'periodo hasta
         'pliqhasta = arrparametros(2)
         
         Call generarDatos
            
            
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        Flog.writeline "Proceso Finalizado correctamente"
        objConn.CommitTrans
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    MyRollbackTrans
    Flog.writeline
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & l_sql
    Flog.writeline



End Sub
Sub generarDatos()

Dim l_sql As String
Dim rsPeriodos As New ADODB.Recordset
Dim rsConfrep As New ADODB.Recordset
Dim rsProcesos As New ADODB.Recordset
Dim rsBajas As New ADODB.Recordset

Dim l_mesdesde As Integer
Dim l_aniodesde As Integer
Dim l_meshasta As Integer
Dim l_aniohasta As Integer
Dim l_meshastaaux As Integer
Dim k As Integer
Dim I As Integer
Dim k_aux As Integer
Dim l_anio As Integer
'Dim l_fechadesde As String
'Dim l_fechahasta As String
Dim arreglo(13) As Integer
Dim arregloProcesos() As String
Dim arregloRem(13) As Double
Dim arregloHoras(13) As Double
'Dim arregloEmpleados(13) As String
Dim arregloRenta(13) As Double
Dim arregloPerdidas(13) As Double
Dim arregloBaseCalculo(13) As Double
Dim diastrabajados1() As Integer
Dim j As Integer
Dim totalEmp As Double
Dim parte_entera As Integer
Dim parte_decimal As Integer
Dim l_pliqnro As Integer
Dim listaProc As String
Dim l_nroEmpresa As Integer
Dim listaEmpleados As String
Dim empleados
Dim Ternro As Long

Dim baseCalculo As Double

'variables tabla detalle
Dim remuneracion As Double
Dim diastrabajados As Integer


'variables del confrep
Dim RentaParametro As Integer
Dim RentaConcepto As String
Dim PerdidaParametro As Integer
Dim PerdidaConcepto As String
Dim porcentajeParametro As Integer
Dim porcentajeConcepto As String
Dim BaseCalculoParametro As Integer
Dim BaseCalculoConcepto As String
Dim esConceptoRemTotal As Boolean
Dim RemTotal As String
Dim tipoHora As Integer
Dim concepto As String
Dim acumulaor As Integer
Dim tipoBusqueda As String
Dim tipo As String
Dim tipoFam As String
'Dim tipo As String
Dim porcentaje As Integer
Dim esConceptoCargaFam As Boolean
Dim cargaFam As String

Dim iduser As String
Dim fechaGen As String
Dim hora As String

Dim r As Integer

Dim remTotalMensual As Boolean
Dim valorRem As Double

hora = Format(Now, "hh:mm:ss ")
fechaGen = ConvFecha(Now())
'iduser =


'LEVANTO DATOS DEL CONFREP
'levanto el concepto o acumulador que va a calcular las remuneraciones totales
l_sql = " SELECT * FROM confrep "
l_sql = l_sql & " WHERE repnro=391 "
l_sql = l_sql & " ORDER BY confnrocol ASC "
OpenRecordset l_sql, rsConfrep
If Not rsConfrep.EOF Then
    Do While Not rsConfrep.EOF
        Select Case rsConfrep!confnrocol
            Case 1:
                RentaParametro = rsConfrep!confval
                RentaConcepto = rsConfrep!confval2
            
            Case 2:
                PerdidaParametro = rsConfrep!confval
                PerdidaConcepto = rsConfrep!confval2
            
            Case 3:
                BaseCalculoParametro = rsConfrep!confval
                BaseCalculoConcepto = rsConfrep!confval2
            '========================================================
            'col4 dias trab
            'col5 remuneracion
            'col6 cargaflia
            'col7 cargaflia2
            'COLUMNAS 4 A 7 FIJAS PARA LAS COLUMNAS DEL REPORTE
            Case 4:
                Select Case rsConfrep!conftipo
                    Case "TH":
                        tipoHora = rsConfrep!confval
                        tipo = "TH"
                    Case "BUS":
                        tipoBusqueda = rsConfrep!confval
                        tipo = "BUS"
                    Case "CO":
                        concepto = rsConfrep!confval2
                        tipo = "CO"
                    Case "ACM":
                        acumulador = rsConfrep!confval
                        tipo = "ACM"
                    Case "AC":
                        acumulador = rsConfrep!confval
                        tipo = "AC"

                End Select
            If tipo <> "" Then
                Flog.writeline "Los dias trabajados se buscaran  mediante el tipo:" & tipo
            Else
                Flog.writeline "No se configuro los dias trabajados"
            End If
            
            Case 5:
                If rsConfrep!conftipo = "CO" Then
                    esConceptoRemTotal = True
                    RemTotal = rsConfrep!confval2
                Else
                    If rsConfrep!conftipo = "AC" Then
                        esConceptoRemTotal = False
                        RemTotal = rsConfrep!confval
                    Else
                        If rsConfrep!conftipo = "ACM" Then
                            remTotalMensual = True
                            RemTotal = rsConfrep!confval
                        End If
                    End If
                End If
                If RemTotal <> "" Then
                    Flog.writeline "La remuneracion total se configuro con el codigo:" & RemTotal
                Else
                    Flog.writeline "No se configuro la remuneracion total"
                End If
                
            Case 6:
                Select Case rsConfrep!conftipo
                    Case "TH":
                        tipoFam = rsConfrep!confval
                        tipoF = "TH"
                    Case "BUS":
                        tipoBusquedaFam = rsConfrep!confval
                        tipoF = "BUS"
                    Case "CO":
                        conceptoFam = rsConfrep!confval2
                        tipoF = "CO"
                    Case "ACM":
                        acumuladorFam = rsConfrep!confval
                        tipoF = "ACM"
                    Case "AC":
                        acumuladorFam = rsConfrep!confval
                        tipoF = "AC"
                End Select
            If tipoF <> "" Then
                Flog.writeline "Los cargas familiares por dias trabajados se buscaran  mediante el tipo:" & tipo
            Else
                Flog.writeline "No se configuro la carga fliar por dias trabajados"
            End If
            
            Case 7:
                If rsConfrep!conftipo = "CO" Then
                    esConceptoCargaFam = True
                    cargaFam = rsConfrep!confval2
                Else
                    esConceptoCargaFam = False
                    cargaFam = rsConfrep!confval
                End If
                If cargaFam <> "" Then
                    Flog.writeline "La remuneracion total por carga familiar se configuro con el codigo:" & cargaFam
                Else
                    Flog.writeline "No se configuro la remuneracion total por carga familiar"
                End If
            '====================================================
            Case 8:
                porcentajeParametro = rsConfrep!confval
                porcentajeConcepto = rsConfrep!confval2
        End Select
    rsConfrep.MoveNext
    Loop
Else
    Flog.writeline "no esta configurado el confrep"
End If

'busco el nro de la empresa
l_sql = " SELECT empnro FROM empresa "
l_sql = l_sql & " WHERE estrnro= " & empresa
OpenRecordset l_sql, rsPeriodos
If Not rsPeriodos.EOF Then
    l_nroEmpresa = rsPeriodos!Empnro
    Flog.writeline "El nro de la empresa es: " & rsPeriodos!Empnro
Else
    Flog.writeline "no se encontro el nro de la empresa "
End If
'hasta aca
rsPeriodos.Close

'busco el mes y anio del periodo desde
l_sql = " SELECT pliqmes, pliqanio "
l_sql = l_sql & " FROM periodo "
l_sql = l_sql & " WHERE pliqnro =" & pliqdesde
OpenRecordset l_sql, rsPeriodos
If Not rsPeriodos.EOF Then
    l_mesdesde = rsPeriodos!pliqmes
    l_aniodesde = rsPeriodos!pliqanio
    Flog.writeline "Mes desde: " & l_mesdesde
    Flog.writeline "Año Desde: " & l_aniodesde
End If
rsPeriodos.Close
'hasta aca

'busco el mes y anio del periodo hasta
l_sql = " SELECT pliqmes, pliqanio "
l_sql = l_sql & " FROM periodo "
l_sql = l_sql & " WHERE pliqnro =" & pliqhasta
OpenRecordset l_sql, rsPeriodos
If Not rsPeriodos.EOF Then
    l_meshasta = rsPeriodos!pliqmes
    l_aniohasta = rsPeriodos!pliqanio
    Flog.writeline "Mes hasta: " & l_meshasta
    Flog.writeline "Año hasta: " & l_aniohasta
End If
rsPeriodos.Close
'hasta aca

'======================================================================
'SEBASTIAN STREMEL 25/02/2013
'ME FIJO SI YA EXISTE UN HISTORICO PARA LA EMPRESA Y EL AÑO Y LO BORRO
objConn.BeginTrans
StrSql = " SELECT * FROM liq_emputil "
StrSql = StrSql & " WHERE anio=" & l_aniodesde
StrSql = StrSql & " AND estrnro=" & empresa
OpenRecordset StrSql, rsPeriodos
If Not rsPeriodos.EOF Then
    Flog.writeline " Existe un historico para el mismo periodo y empresa, sera borrado"
    Do While Not rsPeriodos.EOF
        StrSql = "DELETE FROM liq_emputil"
        StrSql = StrSql & " WHERE utilnro=" & rsPeriodos!utilnro
        StrSql = StrSql & " AND bpronro=" & rsPeriodos!bpronro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'borro el detalle
        StrSql = "SELECT utilnro FROM liq_emputil_Det "
        StrSql = StrSql & " WHERE utilnro=" & rsPeriodos!utilnro
        OpenRecordset StrSql, rsBajas
        Do While Not rsBajas.EOF
            StrSql = "DELETE FROM liq_emputil_Det"
            StrSql = StrSql & " WHERE utilnro=" & rsBajas!utilnro
            objConn.Execute StrSql, , adExecuteNoRecords
        rsBajas.MoveNext
        Loop
        rsBajas.Close
    rsPeriodos.MoveNext
    Loop
Else
    Flog.writeline " No hay un historico generado para el periodo y empresa seleccionado."
End If
rsPeriodos.Close
'======================================================================


If l_aniohasta > l_aniodesde Then
    l_meshastaaux = l_meshasta
    l_meshasta = l_meshasta + 12
End If

'busco la cantidad de empleados para cada uno de los meses
I = 1
listaEmpleados = "0"
For k = l_mesdesde To l_meshasta
    
    If k <= 12 Then
        l_anio = l_aniodesde
        k_aux = k
    Else
        k_aux = k - 12
        l_anio = l_aniohasta
    End If
    l_sql = " SELECT pliqdesde, pliqhasta "
    l_sql = l_sql & " FROM periodo "
    l_sql = l_sql & " WHERE pliqmes=" & k_aux & " AND pliqanio=" & l_anio
    OpenRecordset l_sql, rsPeriodos
    
    If Not rsPeriodos.EOF Then
        l_fechadesde = rsPeriodos!pliqdesde
        l_fechahasta = rsPeriodos!pliqhasta
        Flog.writeline "Periodos correctos"
    Else
        Flog.writeline "Error en los periodos"
    End If
    rsPeriodos.Close
    
    'cuento los empleados para ese mes
    l_sql = " SELECT count(distinct ternro) cant FROM his_estructura "
    l_sql = l_sql & " WHERE his_estructura.estrnro=" & empresa
    l_sql = l_sql & " AND (his_estructura.htetdesde <=" & ConvFecha(l_fechadesde) & " AND ( "
    l_sql = l_sql & " his_estructura.htethasta >= " & ConvFecha(l_fechahasta) & " OR his_estructura.htethasta IS NULL)) "
    l_sql = l_sql & " AND his_estructura.tenro  = 10 "
    OpenRecordset l_sql, rsPeriodos
    If Not rsPeriodos.EOF Then
        arreglo(I) = rsPeriodos!cant
    End If
    rsPeriodos.Close
    
    'guardo los empleados de cada mes
    l_sql = " SELECT distinct his_estructura.ternro "
    l_sql = l_sql & " From Empleado"
    l_sql = l_sql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = 10"
    l_sql = l_sql & " AND ("
    l_sql = l_sql & " (his_estructura.htetdesde <=" & ConvFecha(l_fechadesde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(l_fechahasta) & ""
    l_sql = l_sql & " or his_estructura.htethasta >= " & ConvFecha(l_fechadesde) & ")) OR"
    l_sql = l_sql & " (his_estructura.htetdesde >= " & ConvFecha(l_fechadesde) & " AND (his_estructura.htetdesde <= " & ConvFecha(l_fechahasta) & "))"
    l_sql = l_sql & " )"
    l_sql = l_sql & " WHERE his_estructura.estrnro=" & empresa
    
'    l_sql = " SELECT distinct ternro FROM his_estructura "
'    l_sql = l_sql & " WHERE his_estructura.estrnro=" & empresa
'    l_sql = l_sql & " AND (his_estructura.htetdesde <=" & ConvFecha(l_fechadesde) & " AND ( "
'    l_sql = l_sql & " his_estructura.htethasta >= " & ConvFecha(l_fechahasta) & " OR his_estructura.htethasta IS NULL)) "
'    l_sql = l_sql & " AND his_estructura.tenro  = 10 "
    OpenRecordset l_sql, rsPeriodos
    Do While Not rsPeriodos.EOF
         listaEmpleados = listaEmpleados & ", " & rsPeriodos!Ternro
    rsPeriodos.MoveNext
    Loop
    arregloEmpleados(I) = listaEmpleados
    listaEmpleados = "0"
    'hasta aca
    I = I + 1
Next
'hasta aca

'calculo la cantidad de empleados por mes
For j = 1 To UBound(arreglo)
    totalEmp = FormatNumber(totalEmp + arreglo(j), 2)
Next

totalEmp = FormatNumber((totalEmp / 12), 2)
parte_entera = CInt(totalEmp)
parte_decimal = (totalEmp) - parte_entera
'hasta aca

'busco la remuneracion total por mes y anio
listaProc = "0"
I = 0
For k = l_mesdesde To l_meshasta
    I = I + 1
    If k <= 12 Then
        l_anio = l_aniodesde
        k_aux = k
    Else
        l_anio = l_aniohasta
        k_aux = k - 12
    End If
    l_sql = " SELECT pliqdesde, pliqhasta, pliqnro "
    l_sql = l_sql & " FROM periodo "
    l_sql = l_sql & " WHERE pliqmes=" & k_aux & " AND pliqanio=" & l_anio
    OpenRecordset l_sql, rsPeriodos

    If Not rsPeriodos.EOF Then
        l_fechadesde = rsPeriodos!pliqdesde
        l_fechahasta = rsPeriodos!pliqhasta
        l_pliqnro = rsPeriodos!pliqnro
        Flog.writeline "Periodos correctos 2"
    Else
        Flog.writeline "Error en los periodos"
    End If
    rsPeriodos.Close
    
    'busco todos los procesos entre esas fechas que corresponda a la empresa
    l_sql = "SELECT pronro FROM proceso "
    l_sql = l_sql & " WHERE empnro= " & l_nroEmpresa
    l_sql = l_sql & " AND pliqnro= " & l_pliqnro
    OpenRecordset l_sql, rsProcesos
    If Not rsProcesos.EOF Then
        Do While Not rsProcesos.EOF
            listaProc = listaProc & ", " & rsProcesos!pronro
        rsProcesos.MoveNext
        Loop
        Flog.writeline " Se encontraron procesos para ese periodo y esa empresa "
    Else
        Flog.writeline " No hay procesos para ese periodo y esa empresa "
        listaProc = "0"
        'Exit Sub
    End If
    ReDim Preserve arregloProcesos(I)
    arregloProcesos(I) = listaProc
    listaProc = "0"
    rsProcesos.Close
    Flog.writeline "paso 1"
    'hasta aca

    'busco la renta total
    StrSql = " SELECT SUM(ntevalor) renta FROM novestr "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novestr.concnro"
    StrSql = StrSql & " WHERE estrnro=" & empresa
    StrSql = StrSql & " AND (ntevigencia <> 0 "
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " (ntedesde <= " & ConvFecha(l_fechadesde) & " AND (ntehasta is null or ntehasta >=" & ConvFecha(l_fechahasta)
    StrSql = StrSql & " or ntehasta >=" & ConvFecha(l_fechadesde) & ")) OR"
    StrSql = StrSql & " (ntedesde >= " & ConvFecha(l_fechadesde) & " AND (ntedesde <= " & ConvFecha(l_fechahasta) & "))) OR ntevigencia=0)"
    'StrSql = StrSql & " AND concepto.conccod=" & RentaParametro & " AND tpanro = " & RentaConcepto
    StrSql = StrSql & " AND concepto.conccod=" & RentaConcepto & " AND tpanro = " & RentaParametro
    OpenRecordset StrSql, rsProcesos
    If Not rsProcesos.EOF Then
        arregloRenta(I) = IIf(EsNulo(rsProcesos!renta), 0, rsProcesos!renta)
    End If
    rsProcesos.Close
    Flog.writeline "paso 2"
    'hasta aca

    'busco las perdidas
    StrSql = " SELECT SUM(ntevalor) perdida FROM novestr "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novestr.concnro"
    StrSql = StrSql & " WHERE estrnro=" & empresa
    StrSql = StrSql & " AND (ntevigencia <> 0 "
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " (ntedesde <= " & ConvFecha(l_fechadesde) & " AND (ntehasta is null or ntehasta >=" & ConvFecha(l_fechahasta)
    StrSql = StrSql & " or ntehasta >=" & ConvFecha(l_fechadesde) & ")) OR"
    StrSql = StrSql & " (ntedesde >= " & ConvFecha(l_fechadesde) & " AND (ntedesde <= " & ConvFecha(l_fechahasta) & "))) OR ntevigencia=0)"
    'StrSql = StrSql & " AND concepto.conccod=" & PerdidaParametro & " AND tpanro = " & PerdidaConcepto
    StrSql = StrSql & " AND concepto.conccod=" & PerdidaConcepto & " AND tpanro = " & PerdidaParametro
    OpenRecordset StrSql, rsProcesos
    If Not rsProcesos.EOF Then
        arregloPerdidas(I) = IIf(EsNulo(rsProcesos!perdida), 0, rsProcesos!perdida)
    End If
    rsProcesos.Close
    Flog.writeline "paso 3"
    'hasta aca

    'busco las base de calculo
    StrSql = " SELECT SUM(ntevalor) calculo FROM novestr "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novestr.concnro"
    StrSql = StrSql & " WHERE estrnro=" & empresa
    StrSql = StrSql & " AND (ntevigencia <> 0 "
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " (ntedesde <= " & ConvFecha(l_fechadesde) & " AND (ntehasta is null or ntehasta >=" & ConvFecha(l_fechahasta)
    StrSql = StrSql & " or ntehasta >=" & ConvFecha(l_fechadesde) & ")) OR"
    StrSql = StrSql & " (ntedesde >= " & ConvFecha(l_fechadesde) & " AND (ntedesde <= " & ConvFecha(l_fechahasta) & "))) OR ntevigencia=0)"
    'StrSql = StrSql & " AND concepto.conccod=" & BaseCalculoParametro & " AND tpanro = " & BaseCalculoConcepto
    StrSql = StrSql & " AND concepto.conccod=" & BaseCalculoConcepto & " AND tpanro = " & BaseCalculoParametro
    OpenRecordset StrSql, rsProcesos
    If Not rsProcesos.EOF Then
        arregloBaseCalculo(I) = IIf(EsNulo(rsProcesos!calculo), 0, rsProcesos!calculo)
    End If
    rsProcesos.Close
    Flog.writeline "paso 4"
    'hasta aca
    
    
'busco las hs trabajadas
If tipo = "TH" Then
    For j = 1 To (UBound(arregloEmpleados) - 1)
        StrSql = "SELECT adcanthoras Horas FROM gti_acumdiario"
        StrSql = StrSql & " WHERE ternro IN (" & arregloEmpleados(j) & ")"
        StrSql = StrSql & " AND thnro =" & tipoHora
        StrSql = StrSql & " AND " & ConvFecha(FechaDesde) & " <= adfecha"
        StrSql = StrSql & " AND adfecha <= " & ConvFecha(FechaHasta)
        OpenRecordset StrSql, rsProcesos
        If Not rsProcesos.EOF Then
            horas = IIf(EsNulo(rsProcesos!horas), 0, rsProcesos!horas)
        End If
        arregloHoras(j) = horas
        rsProcesos.Close
    Next
End If
Flog.writeline "paso 5"


If tipo = "BUS" Then
    'Validar que la búsqueda configurada sea del tipo correcto (Case 83: 'Dias Habiles Trabajados)
    StrSql = "SELECT tprognro FROM programa "
    StrSql = StrSql & " WHERE prognro=" & tipoBusqueda
    OpenRecordset StrSql, rsProcesos
    If Not rsProcesos.EOF Then
        If rsProcesos!Tprognro = 83 Then
        '    For j = 1 To UBound(arregloEmpleados)
            Flog.writeline "numero de busqueda correcta"
            NroProg = tipoBusqueda
            Empleado = Split(arregloEmpleados(k), ",")
                canthoras = 0
                For r = 1 To UBound(Empleado)
                    'Call bus_DiasHabiles_Trabajados(Empleado(I))
                    Call bus_DiasHabiles_Trabajados(Empleado(r))
                    canthoras = canthoras + Valor
                Next
            arregloHoras(k) = canthoras
            'Next
        Else
            Flog.writeline "numero de busqueda incorrecta"
        End If
    Else
        Flog.writeline "No se encontro la busqueda"
    End If
End If
Flog.writeline "paso 6"

If tipo = "ACM" Then
    'For j = 1 To (UBound(arregloEmpleados) - 1)
        StrSql = " SELECT sum(ammonto) valor FROM acu_mes "
        StrSql = StrSql & " WHERE acunro=" & acumulador & " "
        StrSql = StrSql & " AND amanio=" & l_anio & " AND ammes=" & k
        StrSql = StrSql & " AND ternro IN(" & arregloEmpleados(k) & ")"
        Flog.writeline "strsql: " & StrSql
        OpenRecordset StrSql, rsProcesos
        
        If Not rsProcesos.EOF Then
            horas = IIf(EsNulo(rsProcesos!Valor), 0, rsProcesos!Valor)
            Flog.writeline "Dias trabajados, acumulador mensual encontrado"
            
        Else
            Flog.writeline "Dias trabajados, acumulador mensual no encontrado"
            horas = "0"
        End If
    arregloHoras(k) = horas
    'Next
    
End If
Flog.writeline "paso 7"


'NUEVO BUSCO LOS DIAS TRAB Y REM POR ACU MENSUAL-----------------------
If remTotalMensual Then
    'For j = 1 To (UBound(arregloEmpleados) - 1)
        StrSql = " SELECT sum(ammonto) valor FROM acu_mes "
        StrSql = StrSql & " WHERE acunro=" & RemTotal & " "
        StrSql = StrSql & " AND amanio=" & l_anio & " AND ammes=" & k
        StrSql = StrSql & " AND ternro IN(" & arregloEmpleados(k) & ")"
        Flog.writeline "strsql: " & StrSql
        OpenRecordset StrSql, rsProcesos
        
        If Not rsProcesos.EOF Then
            valorRem = IIf(EsNulo(rsProcesos!Valor), 0, rsProcesos!Valor)
            Flog.writeline "Remuneracion, acumulador mensual encontrado"
        Else
            valorRem = "0"
            Flog.writeline "Remuneracion, acumulador mensual No encontrado"
        End If
    arregloRem(k) = valorRem
    'Next
    
End If
Flog.writeline "paso 8"
'HASTA ACA-------------------------------------------------------------


'hasta aca
    

Next
'hasta aca

'calculo la remuneracion para cada uno de los procesos

totalEmpleados = UBound(arregloProcesos)
cantRegistros = totalEmpleados

For j = 1 To UBound(arregloProcesos)
    If esConceptoRemTotal Then
        l_sql = " SELECT SUM(dlimonto) valor FROM cabliq "
        l_sql = l_sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
        l_sql = l_sql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
        If arregloProcesos(j) <> "" Then
            l_sql = l_sql & " WHERE cabliq.pronro in (" & arregloProcesos(j) & ")"
        Else
            l_sql = l_sql & " WHERE cabliq.pronro = 0 "
        End If
        l_sql = l_sql & " AND concepto.conccod='" & RemTotal & "'"
        OpenRecordset l_sql, rsProcesos
        If Not rsProcesos.EOF Then
            arregloRem(j) = IIf(EsNulo(rsProcesos!Valor), 0, rsProcesos!Valor)
        Else
            arregloRem(j) = "0"
            Flog.writeline "No hay datos"
        End If
        rsProcesos.Close
    Else
        If Not remTotalMensual Then
            l_sql = " select SUM(almonto) valor from cabliq "
            l_sql = l_sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
            'l_sql = l_sql & " WHERE cabliq.pronro in (" & arregloProcesos(j) & ")"
            If arregloProcesos(j) <> "" Then
                l_sql = l_sql & " WHERE cabliq.pronro in (" & arregloProcesos(j) & ")"
            Else
                l_sql = l_sql & " WHERE cabliq.pronro = 0 "
            End If
            l_sql = l_sql & " AND acu_liq.acunro=" & RemTotal
            OpenRecordset l_sql, rsProcesos
            If Not rsProcesos.EOF Then
                arregloRem(j) = IIf(EsNulo(rsProcesos!Valor), 0, rsProcesos!Valor)
            Else
                arregloRem(j) = "0"
                Flog.writeline "No hay datos"
            End If
            rsProcesos.Close
        End If
    End If
   
    
    
    ' si la busqueda de dias trabajados es por concepto la busca aca
    If tipo = "CO" Then
        l_sql = " SELECT SUM(dlimonto) horas FROM cabliq "
        l_sql = l_sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
        l_sql = l_sql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
        l_sql = l_sql & " WHERE cabliq.pronro in (" & arregloProcesos(j) & ")"
        l_sql = l_sql & " AND concepto.conccod='" & concepto & "'"
        OpenRecordset l_sql, rsProcesos
        If Not rsProcesos.EOF Then
            arregloHoras(j) = IIf(EsNulo(rsProcesos!horas), 0, rsProcesos!horas)
        Else
            Flog.writeline "No hay datos"
        End If
        rsProcesos.Close
    Else
        If tipo = "AC" Then
            l_sql = " SELECT SUM(almonto) horas FROM cabliq "
            l_sql = l_sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
            l_sql = l_sql & " WHERE cabliq.pronro in (" & arregloProcesos(j) & ")"
            l_sql = l_sql & " AND acu_liq.acunro='" & acumulador & "'"
            OpenRecordset l_sql, rsProcesos
            If Not rsProcesos.EOF Then
                arregloHoras(j) = IIf(EsNulo(rsProcesos!horas), 0, rsProcesos!horas)
            Else
                Flog.writeline "No hay datos"
            End If
            rsProcesos.Close
        End If
    End If
    
    'hasta aca
    
    
'Actualizo el estado del proceso
TiempoAcumulado = GetTickCount
cantRegistros = cantRegistros - 1

StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
         ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
   
objConn.Execute StrSql, , adExecuteNoRecords
    
Next
'hasta aca


'==============================================================
'BUSCO LA BASE DE CALCULO
'busco las base de calculo
StrSql = " SELECT SUM(ntevalor) calculo FROM novestr "
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novestr.concnro"
StrSql = StrSql & " WHERE estrnro=" & empresa
StrSql = StrSql & " AND (ntevigencia <> 0 "
StrSql = StrSql & " AND ("
StrSql = StrSql & " (ntedesde <= " & ConvFecha(l_fechadesde) & " AND (ntehasta is null or ntehasta >=" & ConvFecha(l_fechahasta)
StrSql = StrSql & " or ntehasta >=" & ConvFecha(l_fechadesde) & ")) OR"
StrSql = StrSql & " (ntedesde >= " & ConvFecha(l_fechadesde) & " AND (ntedesde <= " & ConvFecha(l_fechahasta) & "))) OR ntevigencia=0)"
'StrSql = StrSql & " AND concepto.conccod=" & BaseCalculoParametro & " AND tpanro = " & BaseCalculoConcepto
StrSql = StrSql & " AND concepto.conccod=" & BaseCalculoConcepto & " AND tpanro = " & BaseCalculoParametro
OpenRecordset StrSql, rsProcesos
If Not rsProcesos.EOF Then
    baseCalculo = IIf(EsNulo(rsProcesos!calculo), 0, rsProcesos!calculo)
End If
rsProcesos.Close
Flog.writeline "se busco la base de calculo"
'hasta aca
'==============================================================
    
    
'==============================================================
'BUSCO EL PORCENTAJE DE PARTICIPACION DE LA EMPRESA
'busco las base de calculo
StrSql = " SELECT SUM(ntevalor) calculo FROM novestr "
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novestr.concnro"
StrSql = StrSql & " WHERE estrnro=" & empresa
StrSql = StrSql & " AND (ntevigencia <> 0 "
StrSql = StrSql & " AND ("
StrSql = StrSql & " (ntedesde <= " & ConvFecha(l_fechadesde) & " AND (ntehasta is null or ntehasta >=" & ConvFecha(l_fechahasta)
StrSql = StrSql & " or ntehasta >=" & ConvFecha(l_fechadesde) & ")) OR"
StrSql = StrSql & " (ntedesde >= " & ConvFecha(l_fechadesde) & " AND (ntedesde <= " & ConvFecha(l_fechahasta) & "))) OR ntevigencia=0)"
StrSql = StrSql & " AND concepto.conccod=" & porcentajeConcepto & " AND tpanro = " & porcentajeParametro
OpenRecordset StrSql, rsProcesos
If Not rsProcesos.EOF Then
    porcentaje = IIf(EsNulo(rsProcesos!calculo), 0, rsProcesos!calculo)
End If
rsProcesos.Close
'hasta aca
'==============================================================
    

'inserto los datos en la tabla [liq_emputil]
I = 0
For k = l_mesdesde To l_meshasta
    I = I + 1
    If k <= 12 Then
        l_anio = l_aniodesde
        k_aux = k
    Else
        l_anio = l_aniohasta
        k_aux = k - 12
    End If
    
    l_sql = " SELECT pliqdesde, pliqhasta, pliqnro "
    l_sql = l_sql & " FROM periodo "
    l_sql = l_sql & " WHERE pliqmes=" & k_aux & " AND pliqanio=" & l_anio
    OpenRecordset l_sql, rsPeriodos

    If Not rsPeriodos.EOF Then
        l_fechadesde = rsPeriodos!pliqdesde
        l_fechahasta = rsPeriodos!pliqhasta
        l_pliqnro = rsPeriodos!pliqnro
    Else
        Flog.writeline "Error en los periodos"
    End If
    rsPeriodos.Close
    
    
    
''==============================================================
''BUSCO LA BASE DE CALCULO
''busco las base de calculo
'StrSql = " SELECT SUM(ntevalor) calculo FROM novestr "
'StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novestr.concnro"
'StrSql = StrSql & " WHERE estrnro=" & empresa
'StrSql = StrSql & " AND (ntevigencia <> 0 "
'StrSql = StrSql & " AND ("
'StrSql = StrSql & " (ntedesde <= " & ConvFecha(l_fechadesde) & " AND (ntehasta is null or ntehasta >=" & ConvFecha(l_fechahasta)
'StrSql = StrSql & " or ntehasta >=" & ConvFecha(l_fechadesde) & ")) OR"
'StrSql = StrSql & " (ntedesde >= " & ConvFecha(l_fechadesde) & " AND (ntedesde <= " & ConvFecha(l_fechahasta) & "))) OR ntevigencia=0)"
'StrSql = StrSql & " AND concepto.conccod=" & BaseCalculoParametro & " AND tpanro = " & BaseCalculoConcepto
''StrSql = StrSql & " AND concepto.conccod=" & BaseCalculoConcepto & " AND tpanro = " & BaseCalculoParametro
'OpenRecordset StrSql, rsProcesos
'If Not rsProcesos.EOF Then
'    baseCalculo = IIf(EsNulo(rsProcesos!calculo), 0, rsProcesos!calculo)
'End If
'rsProcesos.Close
'Flog.writeline "se busco la base de calculo"
''hasta aca
''==============================================================
'
'
''==============================================================
''BUSCO EL PORCENTAJE DE PARTICIPACION DE LA EMPRESA
''busco las base de calculo
'StrSql = " SELECT SUM(ntevalor) calculo FROM novestr "
'StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novestr.concnro"
'StrSql = StrSql & " WHERE estrnro=" & empresa
'StrSql = StrSql & " AND (ntevigencia <> 0 "
'StrSql = StrSql & " AND ("
'StrSql = StrSql & " (ntedesde <= " & ConvFecha(l_fechadesde) & " AND (ntehasta is null or ntehasta >=" & ConvFecha(l_fechahasta)
'StrSql = StrSql & " or ntehasta >=" & ConvFecha(l_fechadesde) & ")) OR"
'StrSql = StrSql & " (ntedesde >= " & ConvFecha(l_fechadesde) & " AND (ntedesde <= " & ConvFecha(l_fechahasta) & "))) OR ntevigencia=0)"
'StrSql = StrSql & " AND concepto.conccod=" & porcentajeParametro & " AND tpanro = " & porcentajeConcepto
'OpenRecordset StrSql, rsProcesos
'If Not rsProcesos.EOF Then
'    porcentaje = IIf(EsNulo(rsProcesos!calculo), 0, rsProcesos!calculo)
'End If
'rsProcesos.Close
''hasta aca
''==============================================================
'
    StrSql = " INSERT INTO liq_emputil "
    StrSql = StrSql & " (estrnro, anio, mes, empcantempleados, empremempleados, empdiastrabempleados, "
    StrSql = StrSql & " emprenta, empperdidas, basecalculoutil,empporcpart, bpronro, fecha, hora )"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "("
    StrSql = StrSql & empresa & ", "
    StrSql = StrSql & l_anio & ", "
    StrSql = StrSql & k_aux & ", "
    StrSql = StrSql & arreglo(I) & ", "
    StrSql = StrSql & arregloRem(I) & ", "
    StrSql = StrSql & arregloHoras(I) & ", "
    StrSql = StrSql & arregloRenta(I) & ", "
    StrSql = StrSql & arregloPerdidas(I) & ", "
    'StrSql = StrSql & arregloBaseCalculo(I) & ", "
    StrSql = StrSql & baseCalculo & ", "
    StrSql = StrSql & porcentaje & ", "
    StrSql = StrSql & NroProceso & ", "
    StrSql = StrSql & fechaGen & ", "
    StrSql = StrSql & "'" & hora & "'"
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


    'BUSCO LOS DATOS PARA LA TABLA DETALLE
        
        'busco el numero de utilnro de la tabla cabecera
        StrSql = " SELECT max(utilnro) utilnro FROM liq_emputil "
        StrSql = StrSql & " WHERE anio=" & l_anio & " AND mes=" & k_aux
        StrSql = StrSql & " AND bpronro=" & NroProceso
        OpenRecordset StrSql, rsProcesos
        If Not rsProcesos.EOF Then
            utilnro = rsProcesos!utilnro
        Else
            Flog.writeline "no se encontro el utilnro"
        End If
        rsProcesos.Close
        'hasta aca
        
        
        empleados = Split(arregloEmpleados(I), ",")
        For j = 1 To UBound(empleados)
            Ternro = empleados(j)
            diastrabajados = 0
            'para cada empleado del mes y anio busco los datos correspondientes
            
            'para cada empleado en año y mes busco la remuneracion
            If esConceptoRemTotal Then
                l_sql = " SELECT SUM(dlimonto) valor FROM cabliq "
                l_sql = l_sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                l_sql = l_sql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
                l_sql = l_sql & " WHERE cabliq.pronro in (" & arregloProcesos(I) & ")"
                l_sql = l_sql & " AND concepto.conccod='" & RemTotal & "'"
                l_sql = l_sql & " AND cabliq.empleado =" & Ternro
                OpenRecordset l_sql, rsProcesos
                If Not rsProcesos.EOF Then
                    remuneracion = IIf(EsNulo(rsProcesos!Valor), 0, rsProcesos!Valor)
                Else
                    remuneracion = 0
                    Flog.writeline "No hay datos"
                End If
                 
            Else
                If remTotalMensual Then
                    l_sql = " SELECT sum(ammonto) valor FROM acu_mes "
                    l_sql = l_sql & " WHERE acunro=" & RemTotal & " "
                    l_sql = l_sql & " AND amanio=" & l_anio & " AND ammes=" & k_aux
                    l_sql = l_sql & " AND ternro =" & Ternro
                    OpenRecordset l_sql, rsProcesos
                    If Not rsProcesos.EOF Then
                        remuneracion = IIf(EsNulo(rsProcesos!Valor), 0, rsProcesos!Valor)
                        Flog.writeline "Acumulador mensual remuneracion del empleado"
                    Else
                        remuneracion = 0
                        Flog.writeline "No se encontro Acumulador mensual remuneracion del empleado"
                    End If
                Else
                    l_sql = " select SUM(almonto) valor from cabliq "
                    l_sql = l_sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    l_sql = l_sql & " WHERE cabliq.pronro in (" & arregloProcesos(I) & ")"
                    l_sql = l_sql & " AND acu_liq.acunro=" & RemTotal
                    l_sql = l_sql & " AND cabliq.empleado =" & Ternro
                    OpenRecordset l_sql, rsProcesos
                    If Not rsProcesos.EOF Then
                        remuneracion = IIf(EsNulo(rsProcesos!Valor), 0, rsProcesos!Valor)
                    Else
                        remuneracion = 0
                        Flog.writeline "No hay datos"
                    End If
                End If
            End If
            rsProcesos.Close
            'hasta aca
            
            
            'para cada empledo en año y mes busco los dias trabajados
            ' si la busqueda de dias trabajados es por concepto la busca aca
            If tipo = "CO" Then
                l_sql = " SELECT SUM(dlimonto) horas FROM cabliq "
                l_sql = l_sql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                l_sql = l_sql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
                l_sql = l_sql & " WHERE cabliq.pronro in (" & arregloProcesos(I) & ")"
                l_sql = l_sql & " AND concepto.conccod='" & concepto & "'"
                l_sql = l_sql & " AND cabliq.empleado =" & Ternro
                OpenRecordset l_sql, rsProcesos
                If Not rsProcesos.EOF Then
                    diastrabajados = IIf(EsNulo(rsProcesos!horas), 0, rsProcesos!horas)
                Else
                    Flog.writeline "No hay datos"
                End If
                rsProcesos.Close
            Else
                If tipo = "AC" Then
                    l_sql = " SELECT SUM(almonto) horas FROM cabliq "
                    l_sql = l_sql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                    l_sql = l_sql & " WHERE cabliq.pronro in (" & arregloProcesos(I) & ")"
                    l_sql = l_sql & " AND acu_liq.acunro='" & acumulador & "'"
                    l_sql = l_sql & " AND cabliq.empleado =" & Ternro
                    OpenRecordset l_sql, rsProcesos
                    If Not rsProcesos.EOF Then
                        diastrabajados = IIf(EsNulo(rsProcesos!horas), 0, rsProcesos!horas)
                        Flog.writeline " Acumulador mensual de remuneracion encontrado"
                    Else
                        Flog.writeline "No hay datos"
                        Flog.writeline " Acumulador mensual de remuneracion no encontrado"
                    End If
                    rsProcesos.Close
                Else
                    If tipo = "ACM" Then
                        l_sql = " SELECT sum(ammonto) valor FROM acu_mes "
                        l_sql = l_sql & " WHERE acunro=" & acumulador & " "
                        l_sql = l_sql & " AND amanio=" & l_anio & " AND ammes=" & k_aux
                        l_sql = l_sql & " AND ternro =" & Ternro
                        Flog.writeline "l_sql: " & l_sql
                        OpenRecordset l_sql, rsProcesos
                        If Not rsProcesos.EOF Then
                            diastrabajados = IIf(EsNulo(rsProcesos!Valor), 0, rsProcesos!Valor)
                            Flog.writeline " Acumulador mensual de dias encontrado"
                        Else
                            diastrabajados = "0"
                            Flog.writeline " Acumulador mensual de dias no encontrado"
                        End If
                    End If
                End If
            End If
            
            If tipo = "TH" Then
                StrSql = "SELECT adcanthoras Horas FROM gti_acumdiario"
                StrSql = StrSql & " WHERE ternro = (" & Ternro & ")"
                StrSql = StrSql & " AND thnro =" & tipoHora
                StrSql = StrSql & " AND " & ConvFecha(FechaDesde) & " <= adfecha"
                StrSql = StrSql & " AND adfecha <= " & ConvFecha(FechaHasta)
                OpenRecordset StrSql, rsProcesos
                If Not rsProcesos.EOF Then
                    diastrabajados = IIf(EsNulo(rsProcesos!horas), 0, rsProcesos!horas)
                End If
                rsProcesos.Close
            End If
            
            
            If tipo = "BUS" Then
                'Validar que la búsqueda configurada sea del tipo correcto (Case 83: 'Dias Habiles Trabajados)
                StrSql = "SELECT tprognro FROM programa "
                StrSql = StrSql & " WHERE prognro=" & tipoBusqueda
                OpenRecordset StrSql, rsProcesos
                If Not rsProcesos.EOF Then
                    If rsProcesos!Tprognro = 83 Then
                        Call bus_DiasHabiles_Trabajados(Ternro)
                        diastrabajados = diastrabajados + Valor
                    Else
                        Flog.writeline "No se encontro la busqueda"
                    End If
                Else
                    Flog.writeline "No se encontro la busqueda"
                End If
            End If
            
            
            'hasta aca
                    
            'INSERTO EN LA TABLA DETALLE
            StrSql = " INSERT INTO liq_emputil_det "
            StrSql = StrSql & "("
            StrSql = StrSql & " utilnro, ternro, terrem, terdiastrab"
            StrSql = StrSql & ")"
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "("
            StrSql = StrSql & utilnro & ", "
            StrSql = StrSql & Ternro & ", "
            StrSql = StrSql & remuneracion & ", "
            StrSql = StrSql & diastrabajados
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            'HASTA ACA
        
        Next
    'Next
    'HASTA ACA

Next
'hasta aca



End Sub

Public Sub bus_DiasHabiles_Trabajados(ByVal Ternro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Dias Habiles trabajados entre dos fechas teniendo en cuenta las fases
'              Total = DiasDelMes - Licencias - Fin de semana - Feriados
' Autor      : FGZ
' Fecha      : 31/10/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Tipofecha As Integer    '1- Periodo
                            '2- Proceso
Dim DiasHabiles As Double
Dim Dia As Date
Dim EsFeriado As Boolean

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim objFeriado As New Feriado
Dim ConFases As Boolean
Dim Aux_FechaDesde As Date
Dim Aux_FechaHasta As Date
Dim rs_Fases As New ADODB.Recordset
Dim Param_cur As New ADODB.Recordset
Dim IncluyeSabados As Boolean
Dim PorcentajeSabados As Double
Dim IncluyeDomingos As Boolean
Dim PorcentajeDomingos As Double
Dim IncluyeFeriados As Boolean
Dim PorcentajeFeriados As Double


    ConFases = True
    DiasHabiles = 0
    Bien = False
    
    Set objFeriado.Conexion = objConn
    
    ' Obtener los parametros de la Busqueda
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    OpenRecordset StrSql, Param_cur

    If Param_cur!Prognro <> 0 Then
        Tipofecha = Param_cur!Auxint1
        ConFases = CBool(Param_cur!Auxlog1)
        IncluyeSabados = CBool(Param_cur!Auxlog2)
        PorcentajeSabados = Param_cur!Auxint2
        IncluyeFeriados = CBool(Param_cur!Auxlog3)
        PorcentajeFeriados = Param_cur!Auxint3
        IncluyeDomingos = CBool(Param_cur!Auxlog4)
        PorcentajeDomingos = Param_cur!Auxint4
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada "
        End If
        Exit Sub
    End If

    If Tipofecha = 1 Then
        FechaDesde = l_fechadesde 'reemplazar por la fecha de mi periodo
        FechaHasta = l_fechahasta
    Else
        FechaDesde = fechaproinicio 'reemplazar por la fecha de mi proceso
        FechaHasta = fechaprofin
    End If
    
If ConFases Then
    'Busco las fases del periodo
    'StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!Ternro
        StrSql = "SELECT altfec,bajfec FROM fases WHERE empleado = " & Ternro & _
                 " AND real = -1 AND Fases.altfec <= " & ConvFecha(FechaHasta) & _
                 " AND (Fases.bajfec >= " & ConvFecha(FechaDesde) & " OR Fases.bajfec is null )" & _
                 " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    
    Do While Not rs_Fases.EOF
        Aux_FechaDesde = IIf(rs_Fases!altfec < FechaDesde, FechaDesde, rs_Fases!altfec)
        If Not EsNulo(rs_Fases!bajfec) Then
            Aux_FechaHasta = IIf(rs_Fases!bajfec < FechaHasta, rs_Fases!bajfec, FechaHasta)
        Else
            Aux_FechaHasta = FechaHasta
        End If
        
        Dia = Aux_FechaDesde
        Do While Dia <= Aux_FechaHasta
            If Not Esta_de_Licencia(Dia, Ternro) Then
                EsFeriado = objFeriado.Feriado(Dia, Ternro, False)
                
                If Not EsFeriado Then   'No es feriado
                    If Not Weekday(Dia) = 1 Then 'Domingo
                        If Weekday(Dia) = 7 Then 'Sabado
                            If IncluyeSabados Then
                                DiasHabiles = DiasHabiles + (1 * PorcentajeSabados / 100)
                            End If
                        Else
                            DiasHabiles = DiasHabiles + 1
                        End If
                    Else
                        If IncluyeDomingos Then
                            DiasHabiles = DiasHabiles + (1 * PorcentajeDomingos / 100)
                        End If
                    End If
                Else    'Incluye feriados
                    DiasHabiles = DiasHabiles + (1 * PorcentajeFeriados / 100)
                End If
            Else
                'Ese dia esta de licencia
            End If
            Dia = Dia + 1
        Loop
        
       rs_Fases.MoveNext
    Loop
Else
    Dia = FechaDesde
    Do While Dia <= FechaHasta
        If Not Esta_de_Licencia(Dia, Ternro) Then
            EsFeriado = objFeriado.Feriado(Dia, Ternro, False)
            If Not EsFeriado Then   'No es feriado
                If Not Weekday(Dia) = 1 Then 'Domingo
                    If Weekday(Dia) = 7 Then 'Sabado
                        If IncluyeSabados Then
                            DiasHabiles = DiasHabiles + (1 * PorcentajeSabados / 100)
                        End If
                    Else
                        DiasHabiles = DiasHabiles + 1
                    End If
                Else
                    If IncluyeDomingos Then
                        DiasHabiles = DiasHabiles + (1 * PorcentajeDomingos / 100)
                    End If
                End If
            Else    'Incluye feriados
                DiasHabiles = DiasHabiles + (1 * PorcentajeFeriados / 100)
            End If
        Else
            'Esta de licencia ese dia
        End If
        Dia = Dia + 1
    Loop
End If

If DiasHabiles > 30 Then
    DiasHabiles = 30
End If
Bien = True
Valor = DiasHabiles

'Cierro y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing
End Sub


