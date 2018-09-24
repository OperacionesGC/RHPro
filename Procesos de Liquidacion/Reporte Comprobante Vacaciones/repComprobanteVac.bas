Attribute VB_Name = "repComprobanteVac"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "10/05/2011"
'Global Const UltimaModificacion = "Stankunas Cesar - Se crea el modelo inicial para Multivoice Colombia a partir del proceso de Comprobante de Liquidación Final."
'Global Const Version = "1.01"
'Global Const FechaModificacion = "05/01/2012"
'Global Const UltimaModificacion = "Gonzalez  Nicolás - Se crea modelo N° 2 - ASTRAZENECA COLOMBIA"

'Global Const Version = "1.02"
'Global Const FechaModificacion = "05/12/2014"
'Global Const UltimaModificacion = "Fernandez, Matias - CAS-26972 - H&A - Bug en  Comprobante vacaciones Colombia-Se cambiaron mensajes de error "
                                  
'Global Const Version = "1.03"
'Global Const FechaModificacion = "11/12/2014"
'Global Const UltimaModificacion = "Fernandez, Matias - CAS-26972 - H&A - Bug en  Comprobante vacaciones Colombia-Se corrigieron mensajes de error/se omite el logo."
                           
Global Const Version = "1.04"
Global Const FechaModificacion = "19/12/2014"
Global Const UltimaModificacion = "Borrelli Facunado - CAS-26972 - H&A - Bug en Comprobante vacaciones Colombia [Entrega 3] - Se agregan consultas para obtener SUCURSAL, REGIMEN y TIPO EMPLEADO"

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------

Dim fs, f

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
Global arrTipoConc(1 To 1000) As Long

Global tenro1 As Long
Global estrnro1 As Long
Global tenro2 As Long
Global estrnro2 As Long
Global tenro3 As Long
Global estrnro3 As Long
Global fecEstr As String
Global tipoRecibo As Long
Global acumSubTotalRemun As Long
Global zonaDomicilio0
Global zonaDomicilio1
Global zonaDomicilio2
Global ConcImpIRPF
Global ConcImpDescuentoIRPF
Global Tickets As Long
Global Gratificacion As Long
Global ValNetoTalon As Long
Global SueldoRecibo As Long
Global esSueldoRecibo As Long
Global esTicketsConc As Long
Global esSueldoConc As Long
Global esGratificacionConc As Long
Global TicketsNeto
Global TicketsBruto

Global tidDNRP

Global acumGrupo1(1 To 100) As Long
Global acumGrupo2(1 To 100) As Long
Global acumGrupo3(1 To 100) As Long
Global cantAcumGrupo1 As Long
Global cantAcumGrupo2 As Long
Global cantAcumGrupo3 As Long

Global acumHaberesConAp
Global acumNetoImp
Global acumRetImpuesto

Global ConceptoHsExtra
Global ConcBaseTributable
Global SueldoDelMes
Global ConcDiasTrabajados
Global ConcAFC_Empresa
Global ConcPactado_ISAPRE

Global acumAuxDeci1
Global acumAuxDeci2
Global acumAuxDeci3

Global BasicoEscala
Global VacacionesDisfr
Global VacacionesRet
Global DiasVacMesSig
Global VacMesSig
Global ApVacaMesSig
Global ApSaludMesSig
Global SalarioPrimaServ
Global DiasVacCausadas
Global DiasVacPendientes
Global SueldoBaseVacTC
Global SalarioBaseVac
Global SalarioIndemn

Global concnroVacPend As Long
Global paramVacPend As Long

Global arrFormaPago(1 To 1000) As Long
Global generoRecibo As Boolean



Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim objRs3 As New ADODB.Recordset
Dim fechadesde
Dim fechahasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim listapronro
Dim Pronro
Dim Ternro
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim acunroSueldo
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim tituloReporte As String
Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden


'Dim generoRecibo As Boolean

    On Error GoTo CE
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

    
    NroProceso = NroProcesoBatch
    

    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    tituloReporte = ""

    TiempoInicialProceso = GetTickCount

    Nombre_Arch = PathFLog & "ComprobantePagoVac" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Comprobante de Pago de Vacaciones: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
   
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline

    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    HuboErrores = False
    
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CLng(objRs!bprcempleados)
    totalEmpleados = cantRegistros
    
    objRs.Close
   
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       
       Flog.writeline "Parametros del proceso: " & parametros
       
       ArrParametros = Split(parametros, "@")
       
       'Obtengo la lista de procesos
       listapronro = ArrParametros(0)
       
       tipoRecibo = CLng(ArrParametros(1))
       
       tenro1 = CLng(ArrParametros(2))
       estrnro1 = CLng(ArrParametros(3))
       tenro2 = CLng(ArrParametros(4))
       estrnro2 = CLng(ArrParametros(5))
       tenro3 = CLng(ArrParametros(6))
       estrnro3 = CLng(ArrParametros(7))
       fecEstr = ArrParametros(8)
       
       'Armo el titulo del reporte
       strTempo = ArrParametros(9)
       
       ArrParametros = Split(strTempo, "<br>")
       If UBound(ArrParametros) >= 1 Then
          tituloReporte = ArrParametros(1)
       Else
          tituloReporte = ""
       End If
       
       ArrParametros = Split(ArrParametros(0), "- PerigenerarDatosComprobante01odos")
       tituloReporte = ArrParametros(0) & tituloReporte
       
       'EMPIEZA EL PROCESO
       
       'Busco la configuracion del confrep
       Flog.writeline "Obtengo los datos del confrep"
       
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 311"
       OpenRecordset StrSql, objRs2
       If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep"
          Exit Sub
       End If

       acunroSueldo = 0
       zonaDomicilio0 = 1
       zonaDomicilio1 = 1
       zonaDomicilio2 = 1
       tidDNRP = 0
       cantAcumGrupo1 = 0
       cantAcumGrupo2 = 0
       cantAcumGrupo3 = 0
       
       concnroVacPend = 0
       paramVacPend = 0
       
       esTicketsConc = True
       TicketsBruto = 0
       TicketsNeto = 0
       Tickets = 0
       esGratificacionConc = True
       Gratificacion = 0
       
       ValNetoTalon = 0
       
       esSueldoRecibo = True
       SueldoRecibo = 0
       
       acumHaberesConAp = 0
       acumNetoImp = 0
       acumRetImpuesto = 0
       
       acumAuxDeci1 = 0
       acumAuxDeci2 = 0
       acumAuxDeci3 = 0
       
       Do Until objRs2.EOF
          Flog.writeline "Columna " & objRs2!confnrocol
          Select Case objRs2!confnrocol
             Case 1
                'Busco en el confrep el numero de cuenta que se va a usar para
                'buscar el valor del sueldo
                If objRs2!conftipo = "CO" Then
                    esSueldoConc = True
                    
                    StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                    If Not EsNulo(objRs2!confval2) Then
                        StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                    End If
                    OpenRecordset StrSql, objRs3
                    
                    If objRs3.EOF Then
                        acunroSueldo = 0
                    Else
                        acunroSueldo = objRs3!ConcNro
                    End If
                    
                    objRs3.Close
                Else
                    esSueldoConc = False
                    acunroSueldo = objRs2!confval
                End If
             
             Case 30
                'zonaDomicilio 0
                zonaDomicilio0 = objRs2!confval
             Case 31
                'zonaDomicilio 1
                zonaDomicilio1 = objRs2!confval
             Case 32
                'zonaDomicilio 2
                zonaDomicilio2 = objRs2!confval
             Case 37
                'Concepto de Monto Imponible para IRPF
                ConcImpIRPF = objRs2!confval2
             Case 38
                'Concepto de Monto Imponible para Deducciones IRPF
                ConcImpDescuentoIRPF = objRs2!confval2
             Case 39
                'Concepto de Tickets Bruto
                TicketsBruto = objRs2!confval2
             Case 40
                'Concepto de Tickets Neto
                TicketsNeto = objRs2!confval2
             Case 41
                'Nro. de tipo de documento del DNRP
                tidDNRP = objRs2!confval
             Case 42
                'Concepto o Acum de Tickets
                If objRs2!conftipo = "CO" Then
                   esTicketsConc = True
                  
                   StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                    If Not EsNulo(objRs2!confval2) Then
                        StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                    End If
                   OpenRecordset StrSql, objRs3
                  
                   If objRs3.EOF Then
                     Tickets = 0
                   Else
                     Tickets = objRs3!ConcNro
                   End If
                  
                   objRs3.Close
                Else
                   esTicketsConc = False
                   Tickets = objRs2!confval
                End If
             
             Case 43
                'Concepto o Acum de Gratificaciones
                Gratificacion = objRs2!confval2
                If objRs2!conftipo = "CO" Then
                   esGratificacionConc = True
                  
                   StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                    If Not EsNulo(objRs2!confval2) Then
                        StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                    End If
                   OpenRecordset StrSql, objRs3
                  
                   If objRs3.EOF Then
                     Gratificacion = 0
                   Else
                     Gratificacion = objRs3!ConcNro
                   End If
                  
                   objRs3.Close
                Else
                   esGratificacionConc = False
                   Gratificacion = objRs2!confval
                End If
             Case 44
                'Concepto o Acum de Sueldo
                SueldoRecibo = objRs2!confval2
                If objRs2!conftipo = "CO" Then
                  
                   StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                    If Not EsNulo(objRs2!confval2) Then
                        StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                    End If
                   OpenRecordset StrSql, objRs3
                  
                   If objRs3.EOF Then
                     SueldoRecibo = 0
                   Else
                     SueldoRecibo = objRs3!ConcNro
                   End If
                  
                   objRs3.Close
                Else
                   esSueldoRecibo = False
                   SueldoRecibo = objRs2!confval
                End If
             Case 45
                'Concepto o Acum de Neto de Talon
                ValNetoTalon = objRs2!confval
             Case 46
                'Nro de Concepto de HS Extras
                ConceptoHsExtra = objRs2!confval
             Case 47
                'Nro de Concepto de la Base Tributable
                ConcBaseTributable = objRs2!confval2
             Case 48
                'Nro de Concepto del Sueldo del Mes
                SueldoDelMes = objRs2!confval2
             Case 49
                'Busco el Concepto de Dias Trabajados
                ConcDiasTrabajados = objRs2!confval2
             Case 50
                'acumGrupo 1
                cantAcumGrupo1 = cantAcumGrupo1 + 1
                acumGrupo1(cantAcumGrupo1) = objRs2!confval
             Case 51
                'acumGrupo 2
                cantAcumGrupo2 = cantAcumGrupo2 + 1
                acumGrupo2(cantAcumGrupo2) = objRs2!confval
             Case 52
                'acumGrupo 3
                cantAcumGrupo3 = cantAcumGrupo3 + 1
                acumGrupo3(cantAcumGrupo3) = objRs2!confval
             Case 53
                'Busco el concnro del concepto
                StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                If Not EsNulo(objRs2!confval2) Then
                    StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                End If
                OpenRecordset StrSql, objRs3
                  
                If objRs3.EOF Then
                    concnroVacPend = 0
                Else
                     concnroVacPend = objRs3!ConcNro
                End If
                  
                objRs3.Close
             Case 54
                paramVacPend = objRs2!confval
             Case 55
                ConcPactado_ISAPRE = objRs2!confval2
             Case 56
                ConcAFC_Empresa = objRs2!confval2
             Case 60
                'Busco el acumulador de haberes con aportes
                acumHaberesConAp = objRs2!confval
                
             Case 61
                'Busco el acumulador de netos imponibles
                acumNetoImp = objRs2!confval
                
             Case 62
                'Busco el acumulador de ret. impuesto
                acumRetImpuesto = objRs2!confval
                
             Case 80
                'Busco el acumulador de acumAuxDeci1
                acumAuxDeci1 = objRs2!confval
                
             Case 81
                'Busco el acumulador de acumAuxDeci2
                acumAuxDeci2 = objRs2!confval
             Case 82
                'Busco el acumulador de acumAuxDeci3
                acumAuxDeci3 = objRs2!confval
                
             Case 100
                'Sueldo Basico Por Escala
                BasicoEscala = objRs2!confval2
             Case 101
                'Salario de Vacaciones Disfrutadas
                VacacionesDisfr = objRs2!confval2
             Case 102
                'Salario de Vacaciones en Dinero
                VacacionesRet = objRs2!confval2
             Case 103
                'Dias de Vacaciones Mes Sig
                DiasVacMesSig = objRs2!confval2
             Case 104
                'Vacaciones Mes Siguiente
                VacMesSig = objRs2!confval2
             Case 105
                'Aportes de Vacaciones del Mes Siguiente
                ApVacaMesSig = objRs2!confval2
             Case 106
                'Aportes de Salud del Mes Siguiente
                ApSaludMesSig = objRs2!confval2
                
          End Select
          
          objRs2.MoveNext
       Loop
       
       objRs2.Close
       
       'Inicializo los tipos de conceptos
       For I = 1 To 1000
           arrTipoConc(I) = 0
       Next
        
       Flog.writeline "Busco el tipo de cada concepto "
       'Busco el tipo de cada concepto
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 311 "
       OpenRecordset StrSql, objRs2
       acumSubTotalRemun = 0
        
       Do Until objRs2.EOF
           Flog.writeline "Tipo encontrado " & objRs2!conftipo & " en Columna: " & objRs2!confnrocol
           Select Case objRs2!conftipo
              'Remunerativo
              Case "RE"
                 arrTipoConc(objRs2!confval) = 1
              'No Remunerativo
              Case "NR"
                 arrTipoConc(objRs2!confval) = 2
              'Descuento
              Case "DS"
                 arrTipoConc(objRs2!confval) = 3
              'Sub. Total Remunerativo
              Case "ST"
                 acumSubTotalRemun = objRs2!confval
              Case "TL"
                 arrTipoConc(objRs2!confval) = 4
           End Select
        
           objRs2.MoveNext
        Loop
        
        objRs2.Close
        
       'Inicializo las forma de pago
       For I = 1 To 1000
           arrFormaPago(I) = 0
       Next

       'Busco el tipo de cada concepto
       
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 311 AND confnrocol >= 60"

       OpenRecordset StrSql, objRs2
       
       acumSubTotalRemun = 0

       Do Until objRs2.EOF
           
           If objRs2!conftipo = "FP" Then
              Flog.writeline "FP encontrada en Columna " & objRs2!confnrocol & " con valor " & objRs2!confval2
              arrFormaPago(objRs2!confval) = objRs2!confval2
           End If
        
           objRs2.MoveNext
       Loop
       objRs2.Close
        
        
       'Obtengo los empleados sobre los que tengo que generar los recibos
       Flog.writeline "Obtengo los empleados sobre los que tengo que generar los Comprobantes"
       Call CargarEmpleados(NroProceso, rsEmpl)
       
       Flog.writeline "Inicializo progreso"
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
       
       orden = 0
       
       Flog.writeline "Genero por cada empleado un Comprobante de Pago de Vacaciones"
       'Genero por cada empleado un recibo de sueldo
       Do Until rsEmpl.EOF
          Flog.writeline "Lista de procesos " & listapronro
          Flog.writeline "tercero " & rsEmpl!Ternro
          
          arrpronro = Split(listapronro, ",")
          EmpErrores = False
          Ternro = rsEmpl!Ternro
          generoRecibo = False
                    
          'Genero un recibo de sueldo para el empleado por cada proceso
          Flog.writeline "Genero un Comprobante de Pago de Vacaciones para el empleado por cada proceso"
          For I = 0 To UBound(arrpronro)
             Pronro = arrpronro(I)
             
             If tieneConceptosImprimibles(Ternro, Pronro) Then
                Flog.writeline "Generando Comprobante de Pago de Vacacion para el tercero " & Ternro & " para el proceso " & Pronro
             
                'De acuerdo al tipo del recibo son los datos que se guardan en la tabla
                Flog.writeline "Tipo de Comprobante: " & tipoRecibo
                Select Case tipoRecibo
                    Case 1
                        'Generacion del comprobante estandar
                        Call generarDatosComprobante01(Pronro, Ternro, acunroSueldo, tituloReporte, orden)
                    Case 2
                        Call generarDatosComprobante02(Pronro, Ternro, acunroSueldo, tituloReporte, orden)
                    Case Else
                        Flog.writeline "No se encontro el modelo [" & tipoRecibo & "]"
                End Select
            Else
                Flog.writeline "El tercero " & Ternro & " no tiene conceptos imprimibles para el proceso " & Pronro
            End If
             
        Next
          
        If generoRecibo Then
            orden = orden + 1
        End If
             
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
           
        cantRegistros = cantRegistros - 1
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
           
        objConn.Execute StrSql, , adExecuteNoRecords
         
        'Si se generaron todos los recibos de sueldo del empleado correctamente lo borro
        If Not EmpErrores Then
              StrSql = " DELETE FROM batch_empleado "
              StrSql = StrSql & " WHERE bpronro = " & NroProceso
              StrSql = StrSql & " AND ternro = " & Ternro
    
              objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rsEmpl.MoveNext
    Loop
Else
    Exit Sub
End If
   
    If Not HuboErrores Then
        'Actualizo el estado del proceso
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline "************************************************************"
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo sql ejecutado: " & StrSql
    Flog.writeline "************************************************************"
End Sub


Public Sub bus_Anti0(Ternro, pliqanio, pliqmes, ByRef Dia As Integer, ByRef Mes As Integer, ByRef Anio As Integer)

Dim q As Integer

Dim FechaAux As Date

    Bien = False
    Valor = 0
        
    If pliqmes = 12 Then
        FechaAux = CDate("1/1/" & pliqanio + 1) - 1
    Else
        FechaAux = CDate("01/" & pliqmes + 1 & "/" & pliqanio) - 1
    End If
        
    Call bus_Antiguedad(Ternro, "REAL", FechaAux, Dia, Mes, Anio, q)
    
End Sub


Public Sub bus_Antiguedad(ByVal Ternro As Integer, ByVal TipoAnt As String, ByVal fechafin As String, ByRef Dia As Integer, ByRef Mes As Integer, ByRef Anio As Integer, ByRef DiasHabiles As Integer)

Dim aux1 As Long
Dim aux2 As Long
Dim aux3 As Long
Dim fecalta As Date
Dim fecbaja As Date
Dim Seguir As Date
Dim q As Integer

Dim NombreCampo As String

Dim rs_Fases As New ADODB.Recordset

NombreCampo = ""
DiasHabiles = 0

Select Case UCase(TipoAnt)
Case "SUELDO":
    NombreCampo = "sueldo"
Case "INDEMNIZACION":
    NombreCampo = "indemnizacion"
Case "VACACIONES":
    NombreCampo = "vacaciones"
Case "REAL":
    NombreCampo = "real"
Case Else
End Select


' FGZ -27/01/2004
StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro & _
         " AND " & NombreCampo & " = -1 " & _
         " AND not altfec is null " & _
         " AND not (bajfec is null AND estado = 0)" & _
         " AND altfec <= " & ConvFecha(fechafin)

OpenRecordset StrSql, rs_Fases

Do While Not rs_Fases.EOF
    fecalta = rs_Fases!altfec
    
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If rs_Fases!estado Then
        fecbaja = fechafin ' solo es un alta, tomar el fecha-fin
    ElseIf rs_Fases!bajfec <= CDate(fechafin) Then
        fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
    Else
        fecbaja = fechafin ' hasta la fecha ingresada
    End If
    
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    
    If rs_Fases.RecordCount = 1 Then
        Dia = aux1
        Mes = aux2
        Anio = aux3
    Else
        Dia = Dia + aux1
        Mes = Mes + aux2 + Int(Dia / 30)
        Anio = Anio + aux3 + Int(Mes / 12)
        Dia = Dia Mod 30
        Mes = Mes Mod 12
    End If
        
    If Anio = 0 Then
        Call DiasTrab(Ternro, fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
siguiente:
    rs_Fases.MoveNext
Loop

If Anio <> 0 Then
    DiasHabiles = 0
End If

' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub

Public Sub DiasTrab(ByVal Ternro As Integer, ByVal Desde As Date, ByVal Hasta As Date, ByRef DiasH As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias trabajados de acuerdo al turno en que se trabaja y
'              de acuerdo a los dias que figuran como feriados en la tabla de feriados.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim d1 As Integer
Dim d2 As Integer
Dim aux As Integer
Dim aux2 As Integer
Dim dxsem As Integer

Dim rs_pais As New ADODB.Recordset
Dim rs_feriados As New ADODB.Recordset

    dxsem = 5
    
    d1 = Weekday(Desde)
    d2 = Weekday(Hasta)
    
    aux = DateDiff("d", Desde, Hasta) + 1
    If aux < 7 Then
        DiasH = Minimo(aux, dxsem)
    Else
        If aux = 7 Then
            DiasH = dxsem
        Else
            aux2 = 8 - d1 + d2
            If aux2 < 7 Then
                aux2 = Minimo(aux2, dxsem)
            Else
                If aux2 = 7 Then
                    aux2 = dxsem
                End If
            End If
            
            If aux2 >= 7 Then
                aux2 = Abs(aux2 - 7) + Int(aux2 / 7) * dxsem
            Else
                aux2 = aux2 + Int((aux2 - aux2) / 7) * dxsem
            End If
        End If
    End If
    
    aux = 0
    
    StrSql = "SELECT * FROM pais INNER JOIN tercero ON tercero.paisnro = pais.paisnro WHERE tercero.ternro = " & Ternro
    OpenRecordset StrSql, rs_pais
    
    If Not rs_pais.EOF Then
        ' Resto los Feriados Nacionales
        StrSql = "SELECT * FROM feriado WHERE tipferinro = 2 " & _
                 " AND fericodext = " & rs_pais!paisnro & _
                 " AND ferifecha >= " & ConvFecha(Desde) & _
                 " AND ferifecha < " & ConvFecha(Hasta)
        OpenRecordset StrSql, rs_feriados
        
        Do While Not rs_feriados.EOF
            If Weekday(rs_feriados!ferifecha) > 1 Then
                DiasH = DiasH - 1
            End If
            
            ' Siguiente Feriado
            rs_feriados.MoveNext
        Loop
    End If


    ' Resto los feriados por Convenio
    StrSql = "SELECT * FROM empleado INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro " & _
             " INNER JOIN fer_estr ON fer_estr.tenro = his_estructura.tenro " & _
             " INNER JOIN feriado ON fer_estr.ferinro = feriado.ferinro " & _
             " WHERE empleado.ternro = " & Ternro & _
             " AND feriado.tipferinro = 2" & _
             " AND feriado.ferifecha >= " & ConvFecha(Desde) & _
             " AND feriado.ferifecha < " & ConvFecha(Hasta)
    OpenRecordset StrSql, rs_feriados
    
    Do While Not rs_feriados.EOF
        If Weekday(rs_feriados!ferifecha) > 1 Then
            DiasH = DiasH - 1
        End If
        
        ' Siguiente Feriado
        rs_feriados.MoveNext
    Loop
    
    
    ' cierro todo y libero
    If rs_pais.State = adStateOpen Then rs_pais.Close
    If rs_feriados.State = adStateOpen Then rs_feriados.Close
        
    Set rs_feriados = Nothing
    Set rs_pais = Nothing

End Sub


'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    If tipoRecibo = 20 Then
       StrEmpl = " SELECT DISTINCT empleado.ternro, empleado.empleg, empleado.terape, nestact1.estrdext AS nestr1, nestact2.estrcodext AS nestr2"
       StrEmpl = StrEmpl & " From batch_empleado"
       StrEmpl = StrEmpl & " INNER JOIN empleado ON batch_empleado.ternro = empleado.ternro AND batch_empleado.bpronro = " & NroProc
       StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.tenro = 6 "
       StrEmpl = StrEmpl & "      AND (estact1.htetdesde <= " & ConvFecha(fecEstr)
       StrEmpl = StrEmpl & "      AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
       StrEmpl = StrEmpl & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro AND estact2.tenro = 5 "
       StrEmpl = StrEmpl & "      AND (estact2.htetdesde<=" & ConvFecha(fecEstr)
       StrEmpl = StrEmpl & "      AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecEstr) & "))"
       StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro"
       StrEmpl = StrEmpl & " INNER JOIN estructura nestact2 ON nestact2.estrnro = estact2.estrnro"
       StrEmpl = StrEmpl & " Where batch_empleado.bpronro = " & NroProc
       StrEmpl = StrEmpl & " ORDER BY nestr1,nestr2,terape "
    
    Else
       StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc & " ORDER BY progreso,estado "
       
    End If
    
    OpenRecordset StrEmpl, rsEmpl
End Sub

Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function


Public Function Minimo(ByVal X, ByVal Y)
    If X <= Y Then
        Minimo = X
    Else
        Minimo = Y
    End If
End Function

Sub antigfec(Ternro, Anio, Mes, ByRef antdia As Long, ByRef antmes As Long, ByRef antanio As Long)

Dim sql As String
Dim rs1 As New ADODB.Recordset
Dim acum
Dim auxiliar
Dim tipo

sql = "SELECT * FROM confrep WHERE repnro = 311 "

OpenRecordset sql, rs1

tipo = 1
acum = 0

Do Until rs1.EOF
   Select Case rs1!confnrocol
      Case 1
        acum = CInt(rs1!confval)
      Case 2
        tipo = CInt(rs1!confval)
   End Select

   rs1.MoveNext
Loop

rs1.Close

If tipo = 1 Then
   sql = " SELECT sum(ammonto) as suma "
Else
   sql = " SELECT sum(amcant) as suma "
End If

sql = sql & " FROM acu_mes "
sql = sql & " WHERE ternro = " & Ternro
sql = sql & "   AND acunro = " & acum
sql = sql & "   AND ( amanio < " & Anio
sql = sql & "    OR ( ammes <= " & Mes
sql = sql & "   AND amanio = " & Anio
sql = sql & "   ) ) "
  
OpenRecordset sql, rs1
  
If rs1.EOF Then
   antdia = 0
Else
   If IsNull(rs1!Suma) Then
      antdia = 0
   Else
      antdia = CLng(rs1!Suma)
   End If
End If

rs1.Close

auxiliar = antdia
If antdia > 0 Then
    ' calcular año
     If antdia > 360 Then
        antanio = Int(antdia / 360)
        antdia = antdia - 360 * antanio
     End If
     If antdia > 30 Then
        antmes = Int(antdia / 30)
        antdia = antdia - Int(antmes * 30)
     Else
        antdia = antdia
     End If
End If

End Sub


Function sinDatos(Str)

  If IsNull(Str) Then
     sinDatos = True
  Else
     If Trim(Str) = "" Then
        sinDatos = True
     Else
        sinDatos = False
     End If
  End If

End Function

' Recibo de Sueldo para Dabra. En vez de mostrar el puesto, se muestra la sucursal.
Function tieneConceptosImprimibles(Ternro, Pronro)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'--------------------------------------------------------------------
'Controlo si el empleado tiene conceptos imprimibles para el proceso
'--------------------------------------------------------------------

StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
    
OpenRecordset StrSql, rsConsult

tieneConceptosImprimibles = Not rsConsult.EOF

rsConsult.Close

End Function

'--------------------------------------------------------------------
' Se encarga de generar los datos para Multivoice Colombia
'--------------------------------------------------------------------
Sub generarDatosComprobante01(Pronro, Ternro, acunroSueldo, tituloReporte, orden)
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsFec As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim nombre
Dim Direccion
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim Cuil
Dim empFecAlta
Dim empFecBaja
Dim CausaDes
Dim Sueldo
Dim Categoria
Dim CentroCosto
Dim Localidad
Dim proFecPago
Dim pliqhasta
Dim FormaPago
Dim Puesto
Dim Sucursal
Dim TipoSalario
Dim Desde As Date
Dim Hasta As Date
Dim FecHasta As Date
Dim FecDesde As Date
Dim VacFecDde As Date
Dim VacFecHta As Date
Dim CantidadDias
Dim CantidadDiasHab

Dim EmpEstrnro
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim TipoCuenta As String
Dim Cuenta As String
Dim Banco As String
Dim BancoDesc As String
Dim nroCuenta As String

Dim MontoBasicoEscala As Double
Dim MontoVacacionesDisfr As Double
Dim MontoVacacionesRet As Double
Dim MontoDiasVacMesSig As Double
Dim MontoVacMesSig As Double
Dim MontoApVacaMesSig As Double
Dim MontoApSaludMesSig As Double

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset

On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

'------------------------------------------------------------------
'Obtengo el nro de cabezera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   cliqnro = rsConsult!cliqnro
   proFecPago = rsConsult!proFecPago
   profecini = rsConsult!profecini
   proDesc = rsConsult!proDesc
Else
   Flog.writeline "No se encontro el empleado en la cabecera de liquidacion "
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    nombre = rsConsult!ternom & " " & rsConsult!ternom2
    Apellido = rsConsult!terape & " " & rsConsult!terape2
    Legajo = rsConsult!empleg
    
    StrSql = " SELECT top(1) bajfec, caudes FROM fases"
    StrSql = StrSql & " INNER JOIN causa ON causa.caunro = fases.caunro"
    StrSql = StrSql & " WHERE Empleado = " & Ternro
    StrSql = StrSql & " ORDER BY altfec DESC"
    OpenRecordset StrSql, rsFec
    If Not rsFec.EOF Then
        empFecBaja = rsFec!bajfec
        CausaDes = rsFec!caudes
    Else
        empFecBaja = ""
        CausaDes = ""
    End If
    rsFec.Close
    
    StrSql = " SELECT top(1) altfec FROM fases"
    StrSql = StrSql & " WHERE Empleado = " & Ternro
    StrSql = StrSql & " ORDER BY altfec ASC"
    OpenRecordset StrSql, rsFec
    If Not rsFec.EOF Then
        empFecAlta = rsFec!altfec
    Else
        empFecAlta = ""
    End If
    rsFec.Close
    
    If IsNull(rsConsult!empremu) Then
        Sueldo = 0
    Else
        Sueldo = rsConsult!empremu
    End If
Else
    Flog.writeline "No se encontro el empleado."
    GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & Pronro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If

Desde = "01/" & pliqmes & "/" & pliqanio
Hasta = DateAdd("m", 1, pliqhasta)

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

If tenro1 <> 0 Then
    If estrnro1 <> 0 Then
        EmpEstrnro1 = estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro1
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

If tenro2 <> 0 Then
    If estrnro2 <> 0 Then
        EmpEstrnro2 = estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro2
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------


If tenro3 <> 0 Then
    If estrnro3 <> 0 Then
        EmpEstrnro3 = estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro3
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco el valor del cuil
'------------------------------------------------------------------
StrSql = " SELECT cuil.nrodoc "
StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   Cuil = rsConsult!NroDoc
Else
'   Flog.writeline "Error al obtener los datos del cuil"
'   GoTo MError
End If


'------------------------------------------------------------------
'Dias pedidos de Vacaciones
'------------------------------------------------------------------

StrSql = "SELECT DISTINCT  elfechadesde,elfechahasta, elcantdias, vacnotifestado,vacdesc,licestdesabr,emp_lic.emp_licnro"
StrSql = StrSql & " ,vacacion.vacfecdesde, vacacion.vacfechasta FROM emp_lic"
StrSql = StrSql & " INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro"
StrSql = StrSql & " LEFT JOIN vacnotif ON vacnotif.emp_licnro = emp_lic.emp_licnro"
StrSql = StrSql & " LEFT JOIN lic_estado ON lic_estado.licestnro = emp_lic.licestnro"
StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = lic_vacacion.vacnro"
StrSql = StrSql & " Where Empleado = " & Ternro
StrSql = StrSql & " AND tdnro = 2"
StrSql = StrSql & " AND ("
StrSql = StrSql & " (elfechadesde >= '" & Desde & "'"
StrSql = StrSql & " and elfechahasta <= '" & Hasta & "')"
StrSql = StrSql & " or (elfechadesde <  '" & Desde & "'"
StrSql = StrSql & " and elfechahasta <= '" & Hasta & "'"
StrSql = StrSql & " and elfechahasta >= '" & Desde & "')"
StrSql = StrSql & " or (elfechadesde >= '" & Desde & "'"
StrSql = StrSql & " and elfechahasta >  '" & Hasta & "'"
StrSql = StrSql & " and elfechadesde <= '" & Hasta & "')"
StrSql = StrSql & " or (elfechadesde <  '" & Desde & "'"
StrSql = StrSql & " and elfechahasta >  '" & Hasta & "')"
StrSql = StrSql & " )"
StrSql = StrSql & " and vacestado= -1"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   ' VacFecDde = rsConsult!vacfecdesde
   ' VacFecHta = rsConsult!vacfechasta
    VacFecDde = IIf(IsNull(rsConsult!vacfecdesde), " ", rsConsult!vacfecdesde)
    VacFecHta = IIf(IsNull(rsConsult!vacfechasta), " ", rsConsult!vacfechasta)
    FecHasta = rsConsult!elfechahasta
    FecDesde = rsConsult!elfechadesde
    CantidadDiasHab = rsConsult!elcantdias
    CantidadDias = DateDiff("d", CDate(FecDesde), CDate(FecHasta)) + 1
Else
    Flog.writeline "Error al obtener la fecha desde y hasta de las vacaciones"
    Exit Sub
End If


'------------------------------------------------------------------
'Busco el valor de la direccion y localidad
'------------------------------------------------------------------
Direccion = ""
Sucursal = ""
Localidad = ""

StrSql = " SELECT detdom.calle,detdom.nro,detdom.piso,detdom.oficdepto,localidad.locdesc, estrdabr,estructura.estrcodext"
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(proFecPago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(proFecPago) & ") AND his_estructura.tenro=1 AND his_estructura.ternro=" & Ternro
StrSql = StrSql & " INNER JOIN estructura ON sucursal.estrnro = estructura.estrnro"
StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    Sucursal = rsConsult!estrcodext
    Direccion = rsConsult!calle & " " & rsConsult!nro
    
    '02/10/2006 - Martin Ferraro - Se agrego piso y dpto a la dir del la empresa
    If Not EsNulo(rsConsult!piso) Then
        Direccion = Direccion & " P. " & rsConsult!piso
    End If
    If Not EsNulo(rsConsult!oficdepto) Then
        Direccion = Direccion & " Dpto. " & rsConsult!oficdepto
    End If
    
    Direccion = Direccion & ", " & rsConsult!locdesc
    
    Localidad = Direccion
Else
'   Flog.writeline "Error al obtener los datos de la localidad"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

If Sueldo = 0 Then
    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
       Sueldo = rsConsult!almonto
    Else
       Flog.writeline "Error al obtener los datos del sueldo"
       Sueldo = 0
       'GoTo MError
    End If
End If

'------------------------------------------------------------------
'Busco el valor del Monto BasicoEscala
'------------------------------------------------------------------
If BasicoEscala <> "" And Not IsNull(BasicoEscala) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & BasicoEscala & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoBasicoEscala = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoBasicoEscala = 0
    End If
Else
    MontoBasicoEscala = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto VacacionesDisfr
'------------------------------------------------------------------
If VacacionesDisfr <> "" And Not IsNull(VacacionesDisfr) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & VacacionesDisfr & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoVacacionesDisfr = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoVacacionesDisfr = 0
    End If
Else
    MontoVacacionesDisfr = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto VacacionesRet
'------------------------------------------------------------------
If VacacionesRet <> "" And Not IsNull(VacacionesRet) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & VacacionesRet & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoVacacionesRet = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoVacacionesRet = 0
    End If
Else
    MontoVacacionesRet = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto DiasVacMesSig
'------------------------------------------------------------------
If DiasVacMesSig <> "" And Not IsNull(DiasVacMesSig) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & DiasVacMesSig & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoDiasVacMesSig = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoDiasVacMesSig = 0
    End If
Else
    MontoDiasVacMesSig = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto VacMesSig
'------------------------------------------------------------------
If VacMesSig <> "" And Not IsNull(VacMesSig) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & VacMesSig & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoVacMesSig = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoVacMesSig = 0
    End If
Else
    MontoVacMesSig = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto ApVacaMesSig
'------------------------------------------------------------------
If ApVacaMesSig <> "" And Not IsNull(ApVacaMesSig) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & ApVacaMesSig & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoApVacaMesSig = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoApVacaMesSig = 0
    End If
Else
    MontoApVacaMesSig = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto ApSaludMesSig
'------------------------------------------------------------------
If ApSaludMesSig <> "" And Not IsNull(ApSaludMesSig) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & ApSaludMesSig & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoApSaludMesSig = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoApSaludMesSig = 0
    End If
Else
    MontoApSaludMesSig = 0
End If

'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 3 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   Categoria = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la categoria"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del puesto
'------------------------------------------------------------------
Puesto = ""

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   Puesto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If


'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------
StrSql = " SELECT estructura.estrcodext, estrdabr From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrcodext
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
'   GoTo MError
End If


'------------------------------------------------------------------
'Busco el banco y la cuenta del empleado
'------------------------------------------------------------------
Banco = ""
TipoCuenta = ""
Cuenta = ""

StrSql = " SELECT banco.bandesc, ctabancaria.ctabnro, ctabancaria.ctabcbu "
StrSql = StrSql & " From ctabancaria "
StrSql = StrSql & " INNER JOIN banco ON ctabancaria.banco = banco.ternro "
StrSql = StrSql & " WHERE ctabancaria.ternro= " & Ternro
StrSql = StrSql & " AND ctabancaria.ctabestado = -1"
OpenRecordset StrSql, rsConsult
'Forma de Pago: (Estas formas de pago fueron tomadas de la forma de pago de la Base de Test)
'3256 - Cheque
'3255 - Cuenta Corriente
'3254 - Cuenta de Ahorros
'3257 - Efectivo
If Not rsConsult.EOF Then
   If Not EsNulo(rsConsult!ctabcbu) Then
      ' TipoCuenta = "CBU"
       Cuenta = rsConsult!ctabcbu
       BancoDesc = rsConsult!Bandesc
   Else
       If Not EsNulo(rsConsult!ctabnro) Then
           ' TipoCuenta = "Cta. Nro."
            Cuenta = rsConsult!ctabnro
            BancoDesc = rsConsult!Bandesc
       Else
           ' TipoCuenta = ""
            Cuenta = ""
            BancoDesc = ""
       End If
   End If
   Flog.writeline "    los datos de Cuenta + Banco (OK)"
Else
  '  TipoCuenta = ""
    Cuenta = ""
    BancoDesc = ""
    Flog.writeline "Error al obtener los datos de Cuenta + Banco"
End If


'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
StrSql = " SELECT ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc, fpagbanc "
StrSql = StrSql & " From pago"
StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro  AND pago.pagorigen=" & cliqnro
StrSql = StrSql & " LEFT JOIN banco ON banco.ternro = pago.banternro"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!fpagbanc = "-1" Then
        TipoCuenta = rsConsult!fpagdescabr ' & " " & rsConsult!terrazsoc & " " & rsConsult!ctabnro
        nroCuenta = rsConsult!ctabnro
        Banco = rsConsult!terrazsoc
    Else
        TipoCuenta = ""
        nroCuenta = ""
        Banco = ""
    End If
Else
    Banco = ""
    nroCuenta = ""
    TipoCuenta = ""
    Flog.writeline "Error al obtener los datos de la forma de pago "
End If
rsConsult.Close


FormaPago = ""

StrSql = " SELECT estrdabr From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 76 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   FormaPago = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos TipoSalario "
'   GoTo MError
End If
rsConsult.Close


'------------------------------------------------------------------
'Busco el valor de la obra social elegida (Entidad Promotora de Salud - Estr. 17)
'------------------------------------------------------------------
TipoSalario = ""

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 63 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   TipoSalario = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos TipoSalario "
'   GoTo MError
End If
rsConsult.Close


' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = 10"
OpenRecordset StrSql, rs_estructura
EmpEstrnro = 0
If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    Exit Sub
Else
    EmpNombre = rs_estructura!empnom
    EmpEstrnro = rs_estructura!Estrnro
End If

'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc, detdom.piso, detdom.oficdepto From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    EmpDire = "   "
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro
    '02/10/2006 - Martin Ferraro - Se agrego piso y dpto a la dir del la empresa
    If Not EsNulo(rs_Domicilio!piso) Then
        EmpDire = EmpDire & " P. " & rs_Domicilio!piso
    End If
    If Not EsNulo(rs_Domicilio!oficdepto) Then
        EmpDire = EmpDire & " Dpto. " & rs_Domicilio!oficdepto
    End If
    EmpDire = EmpDire & " - " & rs_Domicilio!locdesc
End If

'Consulta para obtener el cuit de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa"
    'Exit Sub
    EmpCuit = "  "
Else
    EmpCuit = rs_cuit!NroDoc
End If

'Consulta para buscar el logo de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_logo
If rs_logo.EOF Then
   ' Flog.writeline "No se encontró el Logo de la Empresa"
    'Exit Sub
    EmpLogo = ""
    EmpLogoAlto = 0
    EmpLogoAncho = 0
Else
    EmpLogo = rs_logo!tipimdire & rs_logo!terimnombre
    EmpLogoAlto = rs_logo!tipimaltodef
    EmpLogoAncho = rs_logo!tipimanchodef
End If

'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    'Exit Sub
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If

'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & EmpEstrnro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   pliqdepant = rsConsult!periodoant
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!Banco
Else
   pliqdepant = ""
   pliqfecdep = ""
   pliqbco = ""
   Flog.writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If
rsConsult.Close


'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------
StrSql = " INSERT INTO rep_comprobante_vac "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,direccion,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,pliqdepant,"
StrSql = StrSql & " pliqfecdep,pliqbco,cuil,empfecalta,empfecbaja,causadespido,"
StrSql = StrSql & " sueldo,categoria,centrocosto,localidad,"
StrSql = StrSql & " profecpago,empnombre,empdire,empcuit,emplogo,emplogoalto,emplogoancho,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,formapago,prodesc,descripcion,puesto, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden, auxchar1, auxchar2, auxchar3, auxchar4, auxchar5)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & ",'" & empFecBaja & "'"
StrSql = StrSql & ",'" & CausaDes & "'"
StrSql = StrSql & "," & Sueldo
StrSql = StrSql & ",'" & Mid(Categoria, 1, 20) & "'"
StrSql = StrSql & ",'" & Mid(CentroCosto, 1, 25) & "'"
StrSql = StrSql & ",'" & Mid(Localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & FormaPago & "'"
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden

StrSql = StrSql & ",'" & Sucursal & "'"
StrSql = StrSql & ",'" & VacFecDde & "'"
StrSql = StrSql & ",'" & VacFecHta & "'"
StrSql = StrSql & ",'" & FecDesde & "'"
StrSql = StrSql & ",'" & FecHasta & "'"
StrSql = StrSql & ")"
    
'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

Flog.writeline "SQL INSERT: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado
'------------------------------------------------------------------

StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    generoRecibo = True
End If
Do Until rsConsult.EOF
    StrSql = " INSERT INTO rep_comprobante_vac_det "
    StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
    StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
    StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & "," & Pronro
    StrSql = StrSql & "," & rsConsult!cliqnro
    StrSql = StrSql & ",'" & rsConsult!concabr & "'"
    StrSql = StrSql & ",'" & rsConsult!ConcCod & "'"
    StrSql = StrSql & "," & rsConsult!ConcNro
    StrSql = StrSql & "," & rsConsult!tconnro
    StrSql = StrSql & "," & rsConsult!concimp
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
    StrSql = StrSql & ",'" & arrTipoConc(rsConsult!tconnro) & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rsConsult.MoveNext
Loop
rsConsult.Close


'------------------------------------------------------------------
'Inserto Detalle 2
'------------------------------------------------------------------
StrSql = " INSERT INTO rep_comprobante_vac_det2"
StrSql = StrSql & " (bpronro, ternro,"
StrSql = StrSql & " SueldoBasico, SalVacDisfrutadas, SalarioVacDinero, DiasVacMesSig, MontoVacMesSig, MontoApVacaMesSig"
StrSql = StrSql & " , MontoApSaludMesSig, CantidadDiasHab, CantidadDias)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & " (" & NroProceso
StrSql = StrSql & " ," & Ternro
StrSql = StrSql & " ," & MontoBasicoEscala
StrSql = StrSql & " ," & MontoVacacionesDisfr
StrSql = StrSql & " ," & MontoVacacionesRet
StrSql = StrSql & " ," & MontoDiasVacMesSig
StrSql = StrSql & " ," & MontoVacMesSig
StrSql = StrSql & " ," & MontoApVacaMesSig
StrSql = StrSql & " ," & MontoApSaludMesSig
StrSql = StrSql & " ," & CantidadDiasHab
StrSql = StrSql & " ," & CantidadDias & ")"

Flog.writeline "SQL INSERT DETALLE 2: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub
'--------------------------------------------------------------------
' Se encarga de generar los datos para ASTRAZENECA COLOMBIA
'--------------------------------------------------------------------
Sub generarDatosComprobante02(Pronro, Ternro, acunroSueldo, tituloReporte, orden)
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsFec As New ADODB.Recordset
Dim cliqnro
Dim profecini As String

'Variables donde se guardan los datos del INSERT final
Dim Apellido
Dim nombre
Dim Direccion
Dim Legajo
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim pliqdepant
Dim pliqfecdep
Dim pliqbco
Dim Cuil
Dim empFecAlta
Dim empFecBaja
Dim CausaDes
Dim Sueldo
Dim Categoria
Dim CentroCosto
Dim Localidad
Dim proFecPago
Dim pliqhasta
Dim FormaPago
Dim Puesto
Dim Sucursal
Dim TipoSalario
Dim Desde As Date
Dim Hasta As Date
Dim FecHasta As Date
Dim FecDesde As Date
'Dim VacFecDde As Date
'Dim VacFecHta As Date
Dim VacFecDde
Dim VacFecHta

Dim CantidadDias
Dim CantidadDiasHab

Dim EmpEstrnro
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer
Dim proDesc As String

Dim EmpEstrnro1
Dim EmpEstrnro2
Dim EmpEstrnro3

Dim TipoCuenta As String
Dim Cuenta As String
Dim Banco As String
Dim BancoDesc As String
Dim nroCuenta As String

Dim MontoBasicoEscala As Double
Dim MontoVacacionesDisfr As Double
Dim MontoVacacionesRet As Double
Dim MontoDiasVacMesSig As Double
Dim MontoVacMesSig As Double
Dim MontoApVacaMesSig As Double
Dim MontoApSaludMesSig As Double

Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_firma As New ADODB.Recordset
Dim Regimen
Dim tipocontrato
Dim Tipo_empleado
On Error GoTo MError

EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0

'------------------------------------------------------------------
'Obtengo el nro de cabecera de liquidacion y la fecha de pago del proceso
'------------------------------------------------------------------
StrSql = " SELECT cabliq.cliqnro, proceso.profecpago, proceso.prodesc, proceso.profecini FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON cabliq.pronro = proceso.pronro AND proceso.pronro = " & Pronro
StrSql = StrSql & " WHERE cabliq.empleado=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   cliqnro = rsConsult!cliqnro
   proFecPago = rsConsult!proFecPago
   profecini = rsConsult!profecini
   proDesc = rsConsult!proDesc
Else
   Flog.writeline "El empleado no se encuentra en la cabecera de liquidacion"
   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    nombre = rsConsult!ternom & " " & rsConsult!ternom2
    Apellido = rsConsult!terape & " " & rsConsult!terape2
    Legajo = rsConsult!empleg
    
    StrSql = " SELECT top(1) bajfec, caudes FROM fases"
    StrSql = StrSql & " INNER JOIN causa ON causa.caunro = fases.caunro"
    StrSql = StrSql & " WHERE Empleado = " & Ternro
    StrSql = StrSql & " ORDER BY altfec DESC"
    OpenRecordset StrSql, rsFec
    If Not rsFec.EOF Then
        empFecBaja = rsFec!bajfec
        CausaDes = rsFec!caudes
    Else
        empFecBaja = ""
        CausaDes = ""
    End If
    rsFec.Close
    
    StrSql = " SELECT top(1) altfec FROM fases"
    StrSql = StrSql & " WHERE Empleado = " & Ternro
    StrSql = StrSql & " ORDER BY altfec ASC"
    OpenRecordset StrSql, rsFec
    If Not rsFec.EOF Then
        empFecAlta = rsFec!altfec
    Else
        empFecAlta = ""
    End If
    rsFec.Close
    
    If IsNull(rsConsult!empremu) Then
        Sueldo = 0
    Else
        Sueldo = rsConsult!empremu
    End If
Else
    Flog.writeline "No se encontro el Empleado."
    GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.* FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & Pronro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo MError
End If

Desde = "01/" & pliqmes & "/" & pliqanio
Hasta = DateAdd("m", 1, pliqhasta)

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

If tenro1 <> 0 Then
    If estrnro1 <> 0 Then
        EmpEstrnro1 = estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro1
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

If tenro2 <> 0 Then
    If estrnro2 <> 0 Then
        EmpEstrnro2 = estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro2
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------


If tenro3 <> 0 Then
    If estrnro3 <> 0 Then
        EmpEstrnro3 = estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro3
        StrSql = StrSql & " AND (htetdesde <=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!Estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco el valor del cuil
'------------------------------------------------------------------
StrSql = " SELECT cuil.nrodoc "
StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   Cuil = rsConsult!NroDoc
Else
'   Flog.writeline "Error al obtener los datos del cuil"
'   GoTo MError
End If


'------------------------------------------------------------------
'Dias pedidos de Vacaciones
'------------------------------------------------------------------

StrSql = "SELECT DISTINCT  elfechadesde,elfechahasta, elcantdias, vacnotifestado,vacdesc,licestdesabr,emp_lic.emp_licnro"
StrSql = StrSql & " ,vacacion.vacfecdesde, vacacion.vacfechasta FROM emp_lic"
StrSql = StrSql & " INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro"
StrSql = StrSql & " LEFT JOIN vacnotif ON vacnotif.emp_licnro = emp_lic.emp_licnro"
StrSql = StrSql & " LEFT JOIN lic_estado ON lic_estado.licestnro = emp_lic.licestnro"
StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = lic_vacacion.vacnro"
StrSql = StrSql & " Where Empleado = " & Ternro
StrSql = StrSql & " AND tdnro = 2"
StrSql = StrSql & " AND ("
StrSql = StrSql & " (elfechadesde >= '" & Desde & "'"
StrSql = StrSql & " and elfechahasta <= '" & Hasta & "')"
StrSql = StrSql & " or (elfechadesde <  '" & Desde & "'"
StrSql = StrSql & " and elfechahasta <= '" & Hasta & "'"
StrSql = StrSql & " and elfechahasta >= '" & Desde & "')"
StrSql = StrSql & " or (elfechadesde >= '" & Desde & "'"
StrSql = StrSql & " and elfechahasta >  '" & Hasta & "'"
StrSql = StrSql & " and elfechadesde <= '" & Hasta & "')"
StrSql = StrSql & " or (elfechadesde <  '" & Desde & "'"
StrSql = StrSql & " and elfechahasta >  '" & Hasta & "')"
StrSql = StrSql & " )"
StrSql = StrSql & " and vacestado= -1"
Flog.writeline "--------------"
Flog.writeline StrSql
Flog.writeline "--------------"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    'VacFecDde = rsConsult!vacfecdesde
    VacFecDde = IIf(IsNull(rsConsult!vacfecdesde), " ", rsConsult!vacfecdesde)
    'VacFecHta = rsConsult!vacfechasta
    VacFecHta = IIf(IsNull(rsConsult!vacfechasta), " ", rsConsult!vacfechasta)
    FecHasta = rsConsult!elfechahasta
    FecDesde = rsConsult!elfechadesde
    CantidadDiasHab = rsConsult!elcantdias
    CantidadDias = DateDiff("d", CDate(FecDesde), CDate(FecHasta)) + 1
Else
    Flog.writeline "Error al obtener la fecha desde y hasta de las vacaciones"
    Exit Sub
End If


'------------------------------------------------------------------
'Busco el valor de la direccion y localidad
'------------------------------------------------------------------
Direccion = ""
Sucursal = "0"
Localidad = ""

StrSql = " SELECT detdom.calle,detdom.nro,detdom.piso,detdom.oficdepto,localidad.locdesc, estrdabr,estructura.estrnro"
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(proFecPago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(proFecPago) & ") AND his_estructura.tenro=1 AND his_estructura.ternro=" & Ternro
StrSql = StrSql & " INNER JOIN estructura ON sucursal.estrnro = estructura.estrnro"
StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    'Sucursal = rsConsult!estrcodext
    If rsConsult!Estrnro = "" Then
        Sucursal = "0"
    Else
        Sucursal = rsConsult!Estrnro
    End If
    
    Direccion = rsConsult!calle & " " & rsConsult!nro
    
    '02/10/2006 - Martin Ferraro - Se agrego piso y dpto a la dir del la empresa
    If Not EsNulo(rsConsult!piso) Then
        Direccion = Direccion & " P. " & rsConsult!piso
    End If
    If Not EsNulo(rsConsult!oficdepto) Then
        Direccion = Direccion & " Dpto. " & rsConsult!oficdepto
    End If
    
    Direccion = Direccion & ", " & rsConsult!locdesc
    
    Localidad = Direccion
Else
'   Flog.writeline "Error al obtener los datos de la localidad"
'   GoTo MError
End If

'FB ---------------------------------------------------------------
'Se comenta desde aca
'------------------------------------------------------------------
'BUSCO EL TIPO DE ESTRUCTURA Regimen del confrep
'------------------------------------------------------------------
'StrSql = " SELECT * FROM confrep "
'StrSql = StrSql & " WHERE repnro = 311 and conftipo = 'TE'"
'OpenRecordset StrSql, rsConsult
'If rsConsult.EOF Then
'    Regimen = ""
'    Flog.writeline "El tipo de estructura Régimen No esta configurado en el ConfRep"
'Else
'    Regimen = rsConsult!confval
'End If
'rsConsult.Close

'------------------------------------------------------------------
'BUSCO EL TIPO DE ESTRUCTURA Regimen
'------------------------------------------------------------------
'If Regimen <> "" Then
'    StrSql = "SELECT estrnro From his_estructura"
'    StrSql = StrSql & " WHERE htetdesde <= " & ConvFecha(proFecPago) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(proFecPago) & ")"
'    StrSql = StrSql & " AND his_estructura.Tenro = " & Regimen & " AND his_estructura.Ternro = " & Ternro
'    OpenRecordset StrSql, rsConsult
'    If rsConsult.EOF Then
'        Regimen = 0
'        Flog.writeline "El empleado no tiene un regimen asociado o esta mal configurado en el confrep."
'    Else
'        Regimen = rsConsult!Estrnro
'    End If
'    rsConsult.Close
'End If
'Hasta aca
'FB ---------------------------------------------------------------

'FB - Se agrega esto ---------------------------------------------------------------
Regimen = ""
Sucursal = ""
Tipo_empleado = ""
'------------------------------------------------------------------
'BUSCO EL TIPO DE ESTRUCTURA Regimen del confrep
'------------------------------------------------------------------
StrSql = " SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 311 and conftipo = 'TE'"
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    Flog.writeline "El tipo de estructura Régimen No esta configurado en el ConfRep"
Else
    Do While Not rsConsult.EOF
        Select Case UCase(rsConsult!confetiq)
            Case "SUCURSAL"
                Sucursal = rsConsult!confval
            Case "REGIMEN"
                Regimen = rsConsult!confval
            Case "TIPO EMPLEADO"
                 Tipo_empleado = rsConsult!confval
        End Select
        rsConsult.MoveNext
    Loop
    
End If
rsConsult.Close
EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0
'------------------------------------------------------------------
'BUSCO EL TIPO DE ESTRUCTURA sucursal
'------------------------------------------------------------------
If Sucursal <> "" Then
    StrSql = "SELECT estrnro From his_estructura"
    StrSql = StrSql & " WHERE htetdesde <= " & ConvFecha(proFecPago) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(proFecPago) & ")"
    StrSql = StrSql & " AND his_estructura.Tenro = " & Sucursal & " AND his_estructura.Ternro = " & Ternro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        'Sucursal = ""
        EmpEstrnro1 = 0
        Flog.writeline "No se encontró la estructura sucursal. Debe configurar el codigo externo de la estructura"
    Else
        EmpEstrnro1 = rsConsult!Estrnro
    End If
    rsConsult.Close
End If

'------------------------------------------------------------------
'BUSCO EL TIPO DE ESTRUCTURA TIPO EMPLEADO
'------------------------------------------------------------------
If Tipo_empleado <> "" Then
    StrSql = "SELECT estrnro From his_estructura"
    StrSql = StrSql & " WHERE htetdesde <= " & ConvFecha(proFecPago) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(proFecPago) & ")"
    StrSql = StrSql & " AND his_estructura.Tenro = " & Tipo_empleado & " AND his_estructura.Ternro = " & Ternro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        EmpEstrnro2 = 0
        Flog.writeline "No se encontró la  estructura Tipo empleado. Debe configurar el codigo externo de la estructura"
    Else
        EmpEstrnro2 = rsConsult!Estrnro
    End If
    rsConsult.Close
End If

'------------------------------------------------------------------
'BUSCO EL TIPO DE ESTRUCTURA Regimen
'------------------------------------------------------------------
If Regimen <> "" Then
    StrSql = "SELECT estrnro From his_estructura"
    StrSql = StrSql & " WHERE htetdesde <= " & ConvFecha(proFecPago) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(proFecPago) & ")"
    StrSql = StrSql & " AND his_estructura.Tenro = " & Regimen & " AND his_estructura.Ternro = " & Ternro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        'Regimen = "0"
        EmpEstrnro3 = 0
        Flog.writeline "No se encontró la estructura Régimen. Debe configurar el codigo externo de la estructura"
    Else
        EmpEstrnro3 = rsConsult!Estrnro
    End If
    rsConsult.Close
End If
'FB - Hasta aca ---------------------------------------------------------------

'------------------------------------------------------------------
'BUSCO EL TIPO DE ESTRUCTURA TIPO DE CONTRATO
'------------------------------------------------------------------
StrSql = "SELECT estrnro From his_estructura"
StrSql = StrSql & " WHERE htetdesde <= " & ConvFecha(proFecPago) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(proFecPago) & ")"
StrSql = StrSql & " AND his_estructura.Tenro = 18 AND his_estructura.Ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    tipocontrato = 0
    Flog.writeline "No se encuentra el tipo de contrato."
Else
    tipocontrato = rsConsult!Estrnro
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del sueldo basico
'------------------------------------------------------------------
'si el valor sueldo es cero en los datos del empleado entonces tengo que
'buscar el valor del sueldo

If Sueldo = 0 Then
    StrSql = " SELECT almonto"
    StrSql = StrSql & " From acu_liq"
    StrSql = StrSql & " Where acunro = " & acunroSueldo
    StrSql = StrSql & " AND cliqnro = " & cliqnro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
       Sueldo = rsConsult!almonto
    Else
       Flog.writeline "Error al obtener los datos del sueldo"
       Sueldo = 0
       'GoTo MError
    End If
End If

'------------------------------------------------------------------
'Busco el valor del Monto BasicoEscala
'------------------------------------------------------------------
If BasicoEscala <> "" And Not IsNull(BasicoEscala) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & BasicoEscala & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoBasicoEscala = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoBasicoEscala = 0
    End If
Else
    MontoBasicoEscala = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto VacacionesDisfr
'------------------------------------------------------------------
If VacacionesDisfr <> "" And Not IsNull(VacacionesDisfr) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & VacacionesDisfr & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoVacacionesDisfr = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoVacacionesDisfr = 0
    End If
Else
    MontoVacacionesDisfr = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto VacacionesRet
'------------------------------------------------------------------
If VacacionesRet <> "" And Not IsNull(VacacionesRet) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & VacacionesRet & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoVacacionesRet = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoVacacionesRet = 0
    End If
Else
    MontoVacacionesRet = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto DiasVacMesSig
'------------------------------------------------------------------
If DiasVacMesSig <> "" And Not IsNull(DiasVacMesSig) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & DiasVacMesSig & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoDiasVacMesSig = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoDiasVacMesSig = 0
    End If
Else
    MontoDiasVacMesSig = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto VacMesSig
'------------------------------------------------------------------
If VacMesSig <> "" And Not IsNull(VacMesSig) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & VacMesSig & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoVacMesSig = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoVacMesSig = 0
    End If
Else
    MontoVacMesSig = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto ApVacaMesSig
'------------------------------------------------------------------
If ApVacaMesSig <> "" And Not IsNull(ApVacaMesSig) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & ApVacaMesSig & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoApVacaMesSig = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoApVacaMesSig = 0
    End If
Else
    MontoApVacaMesSig = 0
End If

'------------------------------------------------------------------
'Busco el valor del Monto ApSaludMesSig
'------------------------------------------------------------------
If ApSaludMesSig <> "" And Not IsNull(ApSaludMesSig) Then
    StrSql = " SELECT detliq.dlimonto valor FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & cliqnro & " And concepto.conccod = '" & ApSaludMesSig & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        MontoApSaludMesSig = rsConsult!Valor
    Else
        Flog.writeline "Error al obtener el Monto Ticket Bruto."
        MontoApSaludMesSig = 0
    End If
Else
    MontoApSaludMesSig = 0
End If

'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------
StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 3 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   Categoria = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la categoria"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del puesto
'------------------------------------------------------------------
Puesto = ""

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 4 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   Puesto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del puesto"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------
StrSql = " SELECT estructura.estrcodext, estrdabr From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   CentroCosto = rsConsult!estrcodext
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco el banco y la cuenta del empleado
'------------------------------------------------------------------
Banco = ""
TipoCuenta = ""
Cuenta = ""

StrSql = " SELECT banco.bandesc, ctabancaria.ctabnro, ctabancaria.ctabcbu "
StrSql = StrSql & " From ctabancaria "
StrSql = StrSql & " INNER JOIN banco ON ctabancaria.banco = banco.ternro "
StrSql = StrSql & " WHERE ctabancaria.ternro= " & Ternro
StrSql = StrSql & " AND ctabancaria.ctabestado = -1"
OpenRecordset StrSql, rsConsult
'Forma de Pago: (Estas formas de pago fueron tomadas de la forma de pago de la Base de Test)
'3256 - Cheque
'3255 - Cuenta Corriente
'3254 - Cuenta de Ahorros
'3257 - Efectivo
If Not rsConsult.EOF Then
   If Not EsNulo(rsConsult!ctabcbu) Then
      ' TipoCuenta = "CBU"
       Cuenta = rsConsult!ctabcbu
       BancoDesc = rsConsult!Bandesc
   Else
       If Not EsNulo(rsConsult!ctabnro) Then
           ' TipoCuenta = "Cta. Nro."
            Cuenta = rsConsult!ctabnro
            BancoDesc = rsConsult!Bandesc
       Else
           ' TipoCuenta = ""
            Cuenta = ""
            BancoDesc = ""
       End If
   End If
   Flog.writeline "    los datos de Cuenta + Banco (OK)"
Else
  '  TipoCuenta = ""
    Cuenta = ""
    BancoDesc = ""
    Flog.writeline "Error al obtener los datos de Cuenta + Banco"
End If


'------------------------------------------------------------------
'Busco los datos de la forma de pago
'------------------------------------------------------------------
StrSql = " SELECT ctabnro,fpagdescabr,tercero.terrazsoc,fpagbanc, fpagbanc "
StrSql = StrSql & " From pago"
StrSql = StrSql & " INNER JOIN formapago ON formapago.fpagnro = pago.fpagnro  AND pago.pagorigen=" & cliqnro
StrSql = StrSql & " LEFT JOIN banco ON banco.ternro = pago.banternro"
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = banco.ternro"
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    If rsConsult!fpagbanc = "-1" Then
        TipoCuenta = rsConsult!fpagdescabr ' & " " & rsConsult!terrazsoc & " " & rsConsult!ctabnro
        nroCuenta = rsConsult!ctabnro
        Banco = rsConsult!terrazsoc
    Else
        TipoCuenta = ""
        nroCuenta = ""
        Banco = ""
    End If
Else
    Banco = ""
    nroCuenta = ""
    TipoCuenta = ""
    Flog.writeline "Error al obtener los datos de la forma de pago "
End If
rsConsult.Close


FormaPago = ""

StrSql = " SELECT estrdabr From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 76 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   FormaPago = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos TipoSalario "
'   GoTo MError
End If
rsConsult.Close


'------------------------------------------------------------------
'Busco el valor de la obra social elegida (Entidad Promotora de Salud - Estr. 17)
'------------------------------------------------------------------
TipoSalario = ""

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(pliqhasta) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(pliqhasta) & ") And his_estructura.tenro = 63 And his_estructura.ternro = " & Ternro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   TipoSalario = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos TipoSalario "
'   GoTo MError
End If
rsConsult.Close


' -------------------------------------------------------------------------
' Busco los datos de la empresa
'--------------------------------------------------------------------------

StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
    " From his_estructura" & _
    " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
    " WHERE his_estructura.htetdesde <=" & ConvFecha(pliqhasta) & " AND " & _
    " (his_estructura.htethasta >= " & ConvFecha(pliqhasta) & " OR his_estructura.htethasta IS NULL)" & _
    " AND his_estructura.ternro = " & Ternro & _
    " AND his_estructura.tenro  = 10"
OpenRecordset StrSql, rs_estructura
EmpEstrnro = 0
If rs_estructura.EOF Then
    Flog.writeline "No se encontró la empresa"
    Exit Sub
Else
    EmpNombre = rs_estructura!empnom
    EmpEstrnro = rs_estructura!Estrnro
End If

'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc, detdom.piso, detdom.oficdepto From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & rs_estructura!Ternro & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
OpenRecordset StrSql, rs_Domicilio
If rs_Domicilio.EOF Then
    Flog.writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    EmpDire = "   "
Else
    EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!nro
    '02/10/2006 - Martin Ferraro - Se agrego piso y dpto a la dir del la empresa
    If Not EsNulo(rs_Domicilio!piso) Then
        EmpDire = EmpDire & " P. " & rs_Domicilio!piso
    End If
    If Not EsNulo(rs_Domicilio!oficdepto) Then
        EmpDire = EmpDire & " Dpto. " & rs_Domicilio!oficdepto
    End If
    EmpDire = EmpDire & " - " & rs_Domicilio!locdesc
End If

'Consulta para obtener el cuit de la empresa
StrSql = "SELECT cuit.nrodoc FROM tercero " & _
         " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
         " Where tercero.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_cuit
If rs_cuit.EOF Then
    Flog.writeline "No se encontró el CUIT de la Empresa"
    'Exit Sub
    EmpCuit = "  "
Else
    EmpCuit = rs_cuit!NroDoc
End If

'Consulta para buscar el logo de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_logo
If rs_logo.EOF Then
    Flog.writeline "No se encontró el Logo de la Empresa"
    'Exit Sub
    EmpLogo = ""
    EmpLogoAlto = 0
    EmpLogoAncho = 0
Else
    EmpLogo = rs_logo!tipimdire & rs_logo!terimnombre
    EmpLogoAlto = rs_logo!tipimaltodef
    EmpLogoAncho = rs_logo!tipimanchodef
End If

'Consulta para buscar la firma de la empresa
StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
    " From ter_imag " & _
    " INNER JOIN tipoimag ON tipoimag.tipimnro = 2 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
    " AND ter_imag.ternro =" & rs_estructura!Ternro
OpenRecordset StrSql, rs_firma
If rs_firma.EOF Then
    Flog.writeline "No se encontró el Firma de la Empresa"
    'Exit Sub
    EmpFirma = ""
    EmpFirmaAlto = 0
    EmpFirmaAncho = 0
Else
    EmpFirma = rs_firma!tipimdire & rs_firma!terimnombre
    EmpFirmaAlto = rs_firma!tipimaltodef
    EmpFirmaAncho = rs_firma!tipimanchodef
End If

'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & EmpEstrnro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   pliqdepant = rsConsult!periodoant
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!Banco
Else
   pliqdepant = ""
   pliqfecdep = ""
   pliqbco = ""
   Flog.writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If
rsConsult.Close


'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------
StrSql = " INSERT INTO rep_comprobante_vac "
StrSql = StrSql & " (bpronro,ternro,pronro,"
StrSql = StrSql & " apellido,Nombre,direccion,Legajo,"
StrSql = StrSql & " pliqnro,pliqmes,pliqanio,pliqdepant,"
StrSql = StrSql & " pliqfecdep,pliqbco,cuil,empfecalta,empfecbaja,causadespido,"
StrSql = StrSql & " sueldo,categoria,centrocosto,localidad,"
StrSql = StrSql & " profecpago,empnombre,empdire,empcuit,emplogo,emplogoalto,emplogoancho,empfirma,"
StrSql = StrSql & " empfirmaalto,empfirmaancho,formapago,prodesc,descripcion,puesto, "
StrSql = StrSql & " tenro1 , estrnro1, tenro2, estrnro2, tenro3, estrnro3, orden, auxchar1, auxchar2, auxchar3, auxchar4, auxchar5,auxdeci1,auxdeci2,auxdeci3)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & Pronro
StrSql = StrSql & ",'" & Apellido & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & Mid(Direccion, 1, 100) & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & ",'" & pliqdepant & "'"
StrSql = StrSql & ",'" & pliqfecdep & "'"
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & ",'" & empFecAlta & "'"
StrSql = StrSql & ",'" & empFecBaja & "'"
StrSql = StrSql & ",'" & CausaDes & "'"
StrSql = StrSql & "," & Sueldo
StrSql = StrSql & ",'" & Mid(Categoria, 1, 20) & "'"
StrSql = StrSql & ",'" & Mid(CentroCosto, 1, 25) & "'"
StrSql = StrSql & ",'" & Mid(Localidad, 1, 100) & "'"
StrSql = StrSql & ",'" & proFecPago & "'"
StrSql = StrSql & ",'" & EmpNombre & "'"
StrSql = StrSql & ",'" & EmpDire & "'"
StrSql = StrSql & ",'" & EmpCuit & "'"
StrSql = StrSql & ",'" & EmpLogo & "'"
StrSql = StrSql & "," & EmpLogoAlto
StrSql = StrSql & "," & EmpLogoAncho
StrSql = StrSql & ",'" & EmpFirma & "'"
StrSql = StrSql & "," & EmpFirmaAlto
StrSql = StrSql & "," & EmpFirmaAncho
StrSql = StrSql & ",'" & FormaPago & "'"
StrSql = StrSql & ",'" & proDesc & "'"
StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
StrSql = StrSql & ",'" & Puesto & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & EmpEstrnro1
StrSql = StrSql & "," & tenro2
StrSql = StrSql & "," & EmpEstrnro2
StrSql = StrSql & "," & tenro3
StrSql = StrSql & "," & EmpEstrnro3
StrSql = StrSql & "," & orden

StrSql = StrSql & ",'" & Sucursal & "'"
StrSql = StrSql & ",'" & VacFecDde & "'"
StrSql = StrSql & ",'" & VacFecHta & "'"
StrSql = StrSql & ",'" & FecDesde & "'"
StrSql = StrSql & ",'" & FecHasta & "'"
StrSql = StrSql & "," & Regimen
StrSql = StrSql & "," & tipocontrato
StrSql = StrSql & "," & tipocontrato

StrSql = StrSql & ")"
    
'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

Flog.writeline "SQL INSERT: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado
'------------------------------------------------------------------

StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & Pronro
StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 " 'mdf ojoooo es -1
StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
    generoRecibo = True
End If
Do Until rsConsult.EOF
    StrSql = " INSERT INTO rep_comprobante_vac_det "
    StrSql = StrSql & " (bpronro, ternro, pronro, cliqnro,"
    StrSql = StrSql & " concabr, conccod, concnro, tconnro,"
    StrSql = StrSql & " concimp , dlicant, dlimonto,conctipo) "
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & "," & Pronro
    StrSql = StrSql & "," & rsConsult!cliqnro
    StrSql = StrSql & ",'" & rsConsult!concabr & "'"
    StrSql = StrSql & ",'" & rsConsult!ConcCod & "'"
    StrSql = StrSql & "," & rsConsult!ConcNro
    StrSql = StrSql & "," & rsConsult!tconnro
    StrSql = StrSql & "," & rsConsult!concimp
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
    StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
    StrSql = StrSql & ",'" & arrTipoConc(rsConsult!tconnro) & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rsConsult.MoveNext
Loop
rsConsult.Close


'------------------------------------------------------------------
'Inserto Detalle 2
'------------------------------------------------------------------
StrSql = " INSERT INTO rep_comprobante_vac_det2"
StrSql = StrSql & " (bpronro, ternro,"
StrSql = StrSql & " SueldoBasico, SalVacDisfrutadas, SalarioVacDinero, DiasVacMesSig, MontoVacMesSig, MontoApVacaMesSig"
StrSql = StrSql & " , MontoApSaludMesSig, CantidadDiasHab, CantidadDias)"
StrSql = StrSql & " VALUES"
StrSql = StrSql & " (" & NroProceso
StrSql = StrSql & " ," & Ternro
StrSql = StrSql & " ," & MontoBasicoEscala
StrSql = StrSql & " ," & MontoVacacionesDisfr
StrSql = StrSql & " ," & MontoVacacionesRet
StrSql = StrSql & " ," & MontoDiasVacMesSig
StrSql = StrSql & " ," & MontoVacMesSig
StrSql = StrSql & " ," & MontoApVacaMesSig
StrSql = StrSql & " ," & MontoApSaludMesSig
StrSql = StrSql & " ," & CantidadDiasHab
StrSql = StrSql & " ," & CantidadDias
StrSql = StrSql & ")"

Flog.writeline "SQL INSERT DETALLE 2: " & StrSql

objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboErrores = True
    EmpErrores = True
End Sub

