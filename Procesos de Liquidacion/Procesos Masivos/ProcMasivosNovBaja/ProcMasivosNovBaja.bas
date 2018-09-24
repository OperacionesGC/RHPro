Attribute VB_Name = "ProcMasivosNovBaja"
Option Explicit

Global Const Version = 1
Global Const FechaVersion = "21/10/2009"   'Encriptacion de string connection
Global Const UltimaModificacion = "Encriptacion de string connection"
Global Const UltimaModificacion1 = "Manuel Lopez"

' ---------------------------------------------------------------------------------------------
' Descripcion: Modulo que se encarga de generar los procesos masivos (insertar en NovBaja)
' Autor      : GdeCos
' Fecha      : 28/04/2005
' Ultima Mod :
' ---------------------------------------------------------------------------------------------

Dim fs, f

Dim NroProceso As Long

Dim nro_concepto As Integer
Dim nro_parametro As Integer

Global Path As String
Global HuboErrores As Boolean
Global depurar As Boolean
Global p_fecha As Date
Global NroGrilla As Long

Global simpronro  As Integer
Global empternro  As Integer
Global Fec_Fin  As Date
Global fecha_desde  As Date
Global pnovnro  As Integer



Private Sub Main()

Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim StrSql2 As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim rsEmpleados As New ADODB.Recordset
Dim i
Dim cantConcPar
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim proccalculo As Integer
Dim valor_proccalc As Double


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

    On Error GoTo CE
    
    TiempoInicialProceso = GetTickCount
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ProcMasivos" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Procesos Masivos (Novedades Baja): " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
     
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    depurar = False
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs2
    
    If Not objRs2.EOF Then
       
       'Obtengo los parametros del proceso
       Parametros = objRs2!bprcparam
       ArrParametros = Split(Parametros, "@")

       'Obtengo el Nro de simpronro
        pnovnro = CInt(ArrParametros(0))
        
       'Obtengo el Nro de Pronov
        simpronro = CInt(ArrParametros(1))
              
        Flog.writeline "Datos Obtenidos."
        Flog.writeline "Comienza el procesamiento de los datos"
              
        'EMPIEZA EL PROCESO

        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA SQL QUE BUSCA EL CONJUNTO DE CONCEPTOS Y SUS PARAMETROS
        '------------------------------------------------------------------------------------------------------------------------
        
        StrSql2 = " SELECT * FROM pnov_conf "
        StrSql2 = StrSql2 & " WHERE pnovnro = " & pnovnro

        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA Sql QUE BUSCA LOS EMPLEADOS
        '------------------------------------------------------------------------------------------------------------------------
        
        StrSql = "SELECT empleg, empleado, simcabliq.cliqdesde, simproceso.* "
        StrSql = StrSql & " FROM simcabliq"
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = simcabliq.empleado "
        StrSql = StrSql & " INNER JOIN simproceso ON simproceso.pronro = simcabliq.pronro "
        StrSql = StrSql & " WHERE simcabliq.pronro = " & simpronro
        StrSql = StrSql & " ORDER BY empleg "

        'Agrego una entrada en novBaja para cada empleado
        OpenRecordset StrSql, rsEmpleados
        
        Flog.writeline "Proceso Masivo Nro.: " & pnovnro
                                                     
       If rsEmpleados.EOF Then
              Flog.writeline "No se encontraron empleados asignados a este proceso: " & simpronro

       Else
             
             cantConcPar = CInt(rsEmpleados.RecordCount)
             i = 1
            
            ' Para cada empleado relacionado al proceso
            Do Until rsEmpleados.EOF
            
                  empternro = CInt(rsEmpleados!Empleado)
                  Fec_Fin = Format(rsEmpleados!profecbaja, "DD/MM/YYYY")
                  fecha_desde = CDate(rsEmpleados!profecini)
                  p_fecha = CDate(rsEmpleados!profecini)
                                          
                  Flog.writeline "Proceso el empleado Nro: " & empternro
                  
                 'Busco el Conjunto de conceptos y parametros
                 OpenRecordset StrSql2, objRs
                                   
                 If objRs.EOF Then
                    Flog.writeline "No se encontraron Conceptos para el Proceso Masivo: " & pnovnro
                    GoTo Fin
                    
                 Else
                 
                      'Recorro todos los conceptos y parametros del proceso masivo
                      Do While Not objRs.EOF
                     
                          nro_concepto = CInt(objRs!concnro)
                          nro_parametro = CInt(objRs!tpanro)
                          proccalculo = CInt(objRs!tippnovnro)
                          
                          'Obtengo el valor dado por el prorama de calculo
                          valor_proccalc = 0
                          Select Case proccalculo
                              Case 0: valor_proccalc = 0
                              Case 1: valor_proccalc = DiasSAC_Proporcional()
                              Case 2: valor_proccalc = Vacaciones()
                              Case 3: valor_proccalc = Preaviso()
                              Case 4: valor_proccalc = IndemAntig()
                              Case 5: valor_proccalc = MesIntegra()
                          End Select
                        
                          ' Agrego entrada en el log
                          Flog.writeline "Insertando Novedad de Baja para el Empleado:" & rsEmpleados!empleg
                          Flog.writeline "Nro. Concepto:" & nro_concepto & " Nro. Parametro:" & nro_parametro
                          Flog.writeline "Valor Calculado: " & valor_proccalc & "  por el programa de calculo " & proccalculo
                                          
                          ' Inserto en NovBaja
                          InsertarNovBaja nro_concepto, nro_parametro, valor_proccalc
                          
                          objRs.MoveNext
                          
                      Loop
                     
                      TiempoAcumulado = GetTickCount
                        
                      StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix((i / cantConcPar) * 100) & _
                               ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                               " WHERE bpronro = " & NroProceso
                      objConn.Execute StrSql, , adExecuteNoRecords
                  
                      i = i + 1
                                     
                      rsEmpleados.MoveNext
                      
                      Flog.writeline "Se termino de procesar el empleado " & empternro
                
                End If
               
            Loop
               
       End If
    
    Else

       Exit Sub

    End If
    
Fin:
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline "Fin :" & Now
    Flog.Close
    If objRs.State = adStateOpen Then objRs.Close
    If rsEmpleados.State = adStateOpen Then rsEmpleados.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

Sub InsertarNovBaja(ByVal nro_concepto As Integer, ByVal nro_parametro As Integer, ByVal valor_proccalc As Double)
' ---------------------------------------------------------------------------------------------
' Descripcion: Se encarga de Insertar los datos en la tabla NovBaja
' Autor      : GdeCos
' Fecha      : 28/04/2005
' Ultima Mod :
' ---------------------------------------------------------------------------------------------

    Dim StrSql As String
    
    On Error GoTo MError
    
    '-------------------------------------------------------------------------------
    'Inserto los datos en la BD
    '-------------------------------------------------------------------------------
    StrSql = "INSERT INTO novbaja "
    StrSql = StrSql & "(empleado, concnro, tpanro, nevalor, nevigencia)"
    StrSql = StrSql & "VALUES (" & _
             empternro & "," & nro_concepto & "," & nro_parametro & ",'" & valor_proccalc & "',0)"
    objConn.Execute StrSql, , adExecuteNoRecords
        
    Exit Sub
                
MError:
    Flog.writeline " Error: " & Err.Description
    HuboErrores = True
    Exit Sub

End Sub

Function DiasSAC_Proporcional() As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Se encarga de calcular el valor del parametro de un concepto
'               copnfigurado como SAC
' Autor      : GdeCos
' Fecha      : 28/04/2005
' Ultima Mod :
' ---------------------------------------------------------------------------------------------

Dim Fec_Ini_Sem     As Date
Dim Fec_Fin_Sem     As Date
Dim Fec_Ini_1_Sem   As Date
Dim Fec_Ini_2_Sem   As Date
Dim Fec_Ini_Calc    As Date
Dim Fec_Fin_Calc    As Date
Dim Fec_Fin_1_Sem   As Date
Dim Fec_Fin_2_Sem   As Date
Dim Dias_Sac        As Single
Dim Dias_Aus        As Single

Dim Maximo As Double
Dim A_fecha        As Date

Dim rs_Lic As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rsPnov_conf As New ADODB.Recordset
' Dim rsEmp As New ADODB.Recordset
Dim rsPeriodo As New ADODB.Recordset

    
    Bien = False
    Valor = 0
    
    

    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM pnov_conf INNER JOIN pronov ON pronov.pnovnro = pnov_conf.pnovnro"
    StrSql = StrSql & " WHERE pnov_conf.pnovnro = " & pnovnro
    OpenRecordset StrSql, rsPnov_conf

    Maximo = rsPnov_conf!Maximo
    
    'Obtengo los datos del periodo
    StrSql = "SELECT pliqdesde, pliqhasta FROM periodo INNER JOIN simproceso "
    StrSql = StrSql & " ON simproceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " WHERE pronro = " & simpronro
    OpenRecordset StrSql, rsPeriodo
    
'    A_fecha = CDate(CStr(Day(rsPeriodo!pliqhasta) & "/" & Month(rsPeriodo!pliqhasta) & "/" & Year(rsPeriodo!pliqhasta)))
    A_fecha = rsPeriodo!pliqhasta
    
    rsPeriodo.Close
    
'    StrSql = "SELECT * FROM empleado"
'    StrSql = StrSql & " WHERE empleado.ternro = " & empternro
'    OpenRecordset StrSql, rsEmp


    'calculo de inicio del semetre
    Fec_Ini_1_Sem = CDate("01/01/" & Year(A_fecha))
    Fec_Ini_2_Sem = CDate("01/07/" & Year(A_fecha))
    Fec_Ini_Sem = IIf(A_fecha >= Fec_Ini_2_Sem, Fec_Ini_2_Sem, Fec_Ini_1_Sem)
    Fec_Fin_1_Sem = CDate("30/06/" & Year(A_fecha))
    Fec_Fin_2_Sem = CDate("31/12/" & Year(A_fecha))
    Fec_Fin_Sem = IIf(A_fecha >= Fec_Ini_2_Sem, Fec_Fin_2_Sem, Fec_Fin_1_Sem)
    Fec_Fin_Sem = IIf(A_fecha < Fec_Fin_Sem, A_fecha, Fec_Fin_Sem)
    ' SE AGREGARON ESTAS 2 INICIALIZACIONES
    Fec_Ini_Calc = Fec_Ini_Sem
    Fec_Fin_Calc = Fec_Fin_Sem


    'Busco la ultima fase inactiva
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & empternro
    StrSql = StrSql & " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        'Busco la fecha de Inicio
        If rs_Fases!altfec > Fec_Ini_Sem Then
            Fec_Ini_Calc = rs_Fases!altfec
        Else
            Fec_Ini_Calc = Fec_Ini_Sem
        End If
        'Busco la fecha de fin
        If Not EsNulo(rs_Fases!bajfec) Then
            Fec_Fin_Calc = rs_Fases!bajfec
        Else
            If Not EsNulo(Fec_Fin) Then
                Fec_Fin_Calc = Fec_Fin
            Else
                Fec_Fin_Calc = Fec_Fin_Sem
            End If
        End If
    Else
        Fec_Fin_Calc = Fec_Fin_Sem
    End If


    Dias_Sac = DateDiff("d", Fec_Ini_Calc, Fec_Fin_Calc) + 1
    Dias_Sac = IIf(Dias_Sac >= Maximo, Maximo, Dias_Sac)

    Dias_Aus = 0
    StrSql = "SELECT * FROM emp_lic WHERE empleado = " & empternro & _
             " AND elfechadesde <=" & ConvFecha(Fec_Fin_Sem) & _
             " AND elfechahasta >= " & ConvFecha(Fec_Ini_Calc)
    OpenRecordset StrSql, rs_Lic
    
    Do While Not rs_Lic.EOF
        Dias_Aus = Dias_Aus + CantidadDeDias(Fec_Ini_Calc, Fec_Fin_Sem, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
        
        rs_Lic.MoveNext
    Loop

    Dias_Sac = Dias_Sac - Dias_Aus
    
    DiasSAC_Proporcional = IIf(Dias_Sac > 0, Dias_Sac, 0)
    Bien = True

End Function

Public Function Vacaciones() As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Se encarga de calcular el valor del parametro de un concepto
'               copnfigurado como Vacaciones
' Autor      : GdeCos
' Fecha      : 28/04/2005
' Ultima Mod :
' ---------------------------------------------------------------------------------------------

Dim Maximo As Double
Dim Diasvac As Double
Dim Diasvactomados As Double
Dim Genera As Boolean
Dim Propor As Double

Dim rsPnov_conf As New ADODB.Recordset
Dim rs_Emp_Lic As New ADODB.Recordset
Dim StrSql As String


    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM pnov_conf INNER JOIN pronov ON pronov.pnovnro = pnov_conf.pnovnro"
    StrSql = StrSql & " WHERE pnov_conf.pnovnro = " & pnovnro
    OpenRecordset StrSql, rsPnov_conf

    Vacaciones = 0

    Maximo = rsPnov_conf!Maximo

    'A pedido de Analia
    ' run pronov/des01.p(empleado.ternro, FEC-fin, output diasvac,output genera).
    
    Diasvac = bus_DiasVac()
    
    Diasvactomados = Diasvac
    Propor = True


    'Se le descuenta los dias de vacaciones que ya estan marcados como liquidados en el pago /dto de la Gestion integral
    StrSql = "SELECT * FROM emp_lic "
    StrSql = StrSql & " INNER JOIN vacpagdesc ON vacpagdesc.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = lic_vacacion.vacnro "
    StrSql = StrSql & " WHERE (empleado = " & empternro & " )"
    StrSql = StrSql & " AND (tdnro = 2) "
    StrSql = StrSql & " AND elfechahasta < " & ConvFecha(Fec_Fin)
    StrSql = StrSql & " AND vacpagdesc.pago_dto = 3 and not vacpagdesc.pronro is null "
    OpenRecordset StrSql, rs_Emp_Lic
    
    Do While Not rs_Emp_Lic.EOF
        If rs_Emp_Lic!vacanio = Year(Fec_Fin) Then
            Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
            Propor = True
        Else
            If rs_Emp_Lic!vacanio + 1 = Year(Fec_Fin) Then
                Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
                Propor = False
            End If
        End If
        
        rs_Emp_Lic.MoveNext
    Loop
   

    'PROPORCIONAR  LA CANTIDAD TOTAL DE DIAS CORRESPONDIENTES O LA CANT. PENDIENTE EN FUNCION  A LA FECHA DE BAJA
    If Propor Then
        Diasvac = Diasvactomados / 365 * ((Month(Fec_Fin) - 1) * 30 + Day(Fec_Fin))
    Else
        Diasvac = Diasvac / 365 * ((Month(Fec_Fin) - 1) * 30 + Day(Fec_Fin)) + Diasvactomados
    End If

    Diasvac = IIf(Fix(Diasvac) = Diasvac, Diasvac, Fix(Diasvac + 1))

    Vacaciones = IIf(Diasvac < 0, 0, Diasvac)

End Function

Public Function IndemAntig() As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Se encarga de calcular el valor del parametro de un concepto
'               copnfigurado como Indemnizacion por Antiguedad para seg. Fallec
' Autor      : GdeCos
' Fecha      : 28/04/2005
' Ultima Mod :
' ---------------------------------------------------------------------------------------------

'/************************* Integraci¢n Mensual ************************/
'/*******  ADAPTACION NUEVA LEY REFORMA LABORAL #25.013  ***************
'          CON VIGENCIA PARA LAS ALTA POSTERIORES AL 3/10/1998
'          no se paga mas la integracion mensual
'***********************************************************************/

Dim Maximo As Double
Dim Tope As Double
Dim Sueldo  As Double
Dim Monto As Double
Dim anio  As Integer
Dim mes  As Integer
Dim dia  As Integer
Dim diashab  As Integer
Dim i  As Integer

Dim rsPnov_conf As New ADODB.Recordset
Dim rsEmpConv As New ADODB.Recordset
Dim rsAcuMes As New ADODB.Recordset
Dim StrSql As String


    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM pnov_conf INNER JOIN pronov ON pronov.pnovnro = pnov_conf.pnovnro"
    StrSql = StrSql & " WHERE pnov_conf.pnovnro = " & pnovnro
    OpenRecordset StrSql, rsPnov_conf

    Maximo = rsPnov_conf!Maximo
    
    StrSql = "SELECT * FROM empleado " & _
            " INNER JOIN his_estructura ON ((empleado.ternro = his_estructura.ternro) " & _
            " AND (19 = his_estructura.tenro)) " & _
            " INNER JOIN convenios ON convenios.estrnro = his_estructura.estrnro " & _
    StrSql = StrSql & " WHERE empleado.ternro = " & empternro
    OpenRecordset StrSql, rsEmpConv
    

    If rsEmpConv.EOF Then
        Flog.writeline "Imposible realizar el calculo. El empleado no pertenece a un convenio."
        IndemAntig = 0
        Exit Function
    Else
        Tope = CDbl(rsEmpConv!convsueldop)
    End If
     
     Sueldo = 0
     mes = Month(Fec_Fin)
     anio = Year(Fec_Fin)
       
    StrSql = "SELECT * FROM acu_mes  WHERE acu_mes.ternro = " & empternro
    StrSql = StrSql & " AND acu_mes.amanio = " & anio
    ' Ver si no se cablea este acumulador
    StrSql = StrSql & " AND acu_mes.acunro = 9"
    OpenRecordset StrSql, rsAcuMes

    If Not rsAcuMes.EOF Then
        If (rsAcuMes!ammes = mes And rsAcuMes!ammonto > 0) Then
                Sueldo = rsAcuMes!ammonto
                i = 11
                mes = mes - 1
        Else
                i = 12

                Do While i > 0
                                    
                    If mes = 0 Then
                        anio = anio - 1
                        mes = 12
                        rsAcuMes.Close
                        StrSql = "SELECT * FROM acu_mes  WHERE acu_mes.ternro = " & empternro
                        StrSql = StrSql & " AND acu_mes.amanio = " & anio
                        ' Ver si no se cablea este acumulador
                        StrSql = StrSql & " AND acu_mes.acunro = 9"
                        OpenRecordset StrSql, rsAcuMes
                    End If
                    If Not rsAcuMes.EOF Then
                        If (rsAcuMes!ammes = mes And rsAcuMes!ammonto > 0) Then
                            Sueldo = rsAcuMes!ammonto
                        End If
                        mes = mes - 1
                        i = i - 1
                    Else
                        i = 0
                    End If
                Loop
        End If

        Call bus_Antiguedad("INDEMNIZACION", Fec_Fin, dia, mes, anio, diashab)

        If (rsEmpConv!empfaltagr < CDate("10/3/1998")) Then
            If mes > 3 Then
              anio = anio + 1
            End If
            Monto = varios.Maximo(CDbl(Minimo(Sueldo * anio, Tope * anio)), CDbl(Sueldo * 2))
        Else
                If dia > 10 Then
                    mes = mes + 1 + anio * 12
                Else
                    mes = mes + anio * 12
                End If
                Monto = varios.Maximo(Minimo(Sueldo * mes / 12, Tope * mes / 12), Sueldo * 2 / 12)
        End If

    End If
        
    IndemAntig = Monto

End Function

Function Preaviso() As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Se encarga de calcular el valor del parametro de un concepto
'               copnfigurado como Preaviso
' Autor      : GdeCos
' Fecha      : 28/04/2005
' Ultima Mod :
' ---------------------------------------------------------------------------------------------

'/************************* Integraci¢n Mensual ************************/
'/*******  ADAPTACION NUEVA LEY REFORMA LABORAL #25.013  ***************
'          CON VIGENCIA PARA LAS ALTA POSTERIORES AL 3/10/1998
'          no se paga mas la integracion mensual
'***********************************************************************/

Dim anio  As Integer
Dim mes  As Integer
Dim dia  As Integer
Dim diashab  As Integer
Dim rsEmp As New ADODB.Recordset

    anio = 0
    mes = 0
    dia = 0
    diashab = 0

    StrSql = "SELECT * FROM empleado"
    StrSql = StrSql & " WHERE empleado.ternro = " & empternro
    OpenRecordset StrSql, rsEmp

    Call bus_Antiguedad("REAL", Fec_Fin, dia, mes, anio, diashab)

    mes = mes + anio * 12
  
    If mes >= 60 Then
      Monto = 2
    Else
      If mes >= 3 Then
        Monto = 1
      Else
        If (rsEmp!empfaltagr < CDate("10/3/1998")) Then
          Monto = 0.5
        Else
          Monto = 0
        End If
      End If
    End If
        
    Preaviso = Monto

    rsEmp.Close
    
End Function

Function MesIntegra() As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Se encarga de calcular el valor del parametro de un concepto
'               copnfigurado como Indemnizacion por Antiguedad para seg. Fallec
' Autor      : GdeCos
' Fecha      : 28/04/2005
' Ultima Mod :
' ---------------------------------------------------------------------------------------------

'/************************* Integraci¢n Mensual ************************/
'/*******  ADAPTACION NUEVA LEY REFORMA LABORAL #25.013  ***************
'          CON VIGENCIA PARA LAS ALTA POSTERIORES AL 3/10/1998
'          no se paga mas la integracion mensual
'***********************************************************************/
Dim dia As Integer

    dia = Day(Fec_Fin)
    
    MesIntegra = IIf(dia > 30, 0, 30 - dia)
    
End Function

'Public Function Licencias() As Double
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Se encarga de calcular el valor del parametro de un concepto
''               copnfigurado como Licencias
'' Autor      : GdeCos
'' Fecha      : 28/04/2005
'' Ultima Mod :
'' ---------------------------------------------------------------------------------------------
'
'Licencias = bus_Licencias()
'
'End Function



Public Sub DiasTrab(ByVal Desde As Date, ByVal hasta As Date, ByRef DiasH As Integer)
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
Dim Aux As Integer
Dim aux2 As Integer
Dim dxsem As Integer

Dim rs_pais As New ADODB.Recordset
Dim rs_feriados As New ADODB.Recordset

    dxsem = 5
    
    d1 = Weekday(Desde)
    d2 = Weekday(hasta)
    
    Aux = DateDiff("d", Desde, hasta) + 1
    If Aux < 7 Then
        DiasH = Minimo(Aux, dxsem)
    Else
        If Aux = 7 Then
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
    
    Aux = 0
    
    StrSql = "SELECT * FROM pais INNER JOIN tercero ON tercero.paisnro = pais.paisnro WHERE tercero.ternro = " & empternro
    OpenRecordset StrSql, rs_pais
    
    If Not rs_pais.EOF Then
        ' Resto los Feriados Nacionales
        StrSql = "SELECT * FROM feriado WHERE tipferinro = 2 " & _
                 " AND fericodext = " & rs_pais!paisnro & _
                 " AND ferifecha >= " & ConvFecha(Desde) & _
                 " AND ferifecha < " & ConvFecha(hasta)
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
             " WHERE empleado.ternro = " & empternro & _
             " AND feriado.tipferinro = 2" & _
             " AND feriado.ferifecha >= " & ConvFecha(Desde) & _
             " AND feriado.ferifecha < " & ConvFecha(hasta)
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




Function bus_DiasVac() As Integer

' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim Continuar As Boolean
Dim Parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Integer
Dim antmes As Integer
Dim antanio As Integer
Dim q As Integer

Dim ValorCoord As Single
Dim Encontro As Boolean

Dim ternro As Long
Dim NroVac As Long
 
Dim cantdias As Integer
Dim Columna As Integer
Dim Genera As Boolean

Dim dias_trabajados As Integer
Dim DiasProporcion As Integer

    Genera = False
    
    Call Politica(1502)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Function
    End If
    
    StrSql = "SELECT * FROM tipdia WHERE tdnro = 2 " '2 es vacaciones
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        NroGrilla = objRs!tdgrilla
        tdinteger3 = objRs!tdinteger3

        If tdinteger3 <> 20 And tdinteger3 <> 365 And tdinteger3 <> 360 Then
            'El campo auxiliar3 del Tipo de Día para Vacaciones no está configurado para Proporcionar la cant. de días de Vacaciones.
            Exit Function
        End If
    End If

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Exit Function
    End If
    
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop


    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
            Parametros(j) = (antanio * 12) + antmes
        Case Else:
            Select Case j
            Case 1:
                Call bus_Estructura(rs_cabgrilla!grparnro_1)
            Case 2:
                Call bus_Estructura(rs_cabgrilla!grparnro_2)
            Case 3:
                Call bus_Estructura(rs_cabgrilla!grparnro_3)
            Case 4:
                Call bus_Estructura(rs_cabgrilla!grparnro_4)
            Case 5:
                Call bus_Estructura(rs_cabgrilla!grparnro_5)
            End Select
            Parametros(j) = bus_DiasVac
        End Select
    Next j

    'Busco la primera antiguedad de la escala menor a la del empleado
    ' de abajo hacia arriba
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
    For j = 1 To rs_cabgrilla!cgrdimension
        If j <> ant Then
            StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
        End If
    Next j
        StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
    OpenRecordset StrSql, rs_valgrilla


    Encontro = False
    Do While Not rs_valgrilla.EOF And Not Encontro
        Select Case ant
        Case 1:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 2:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 3:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 4:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 5:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        End Select
                    
        rs_valgrilla.MoveNext
    Loop
    
    If Not Encontro Then
    
        'Busco si existe algun valor para la estructura y ...
        'si hay que carga la columna correspondiente
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        StrSql = StrSql & " AND vgrvalor is not null"
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            End If
        Next j
        'StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
        OpenRecordset StrSql, rs_valgrilla
        If Not rs_valgrilla.EOF Then
            Columna = rs_valgrilla!vgrorden
        Else
            Columna = 1
        End If
        
        
        dias_trabajados = ((antanio * 365) + (antmes * 30) + antdia)
    
        If Parametros(ant) <= 6 Then
            'setea la proporcion de dias
            Call Politica(1501)

            If DiasProporcion = 20 Then
                If (dias_trabajados / DiasProporcion) / 7 * 5 > Fix((dias_trabajados / DiasProporcion) / 7 * 5) Then
                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5) + 1
                Else
                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5)
                End If
            Else
                cantdias = Round((dias_trabajados / DiasProporcion) / FactorDivision, 0)
            End If
                       
        Else
            Flog.writeline "No se encontro la escala para el convenio"
            Genera = False
        End If
    End If
   
Genera = True
    
bus_DiasVac = cantdias

' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing

End Function


'Function bus_Licencias() As Integer
'
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Dias de Licencias entre dos fechas (de un tipo o de todos los tipos)
'' Autor      : FGZ
'' Fecha      : 14/01/2004
'' Ultima Mod.: GdeCos - 4/5/2005
'' Descripcion: Ajustes necesarios de acuerdo a los datos, para que funcione en este modulo
'' ---------------------------------------------------------------------------------------------
'
'Dim TipoLicencia As Long    'Tipo de Estructura
'Dim Todas As Boolean        'Todos los tipos
'
'Dim dias As Integer
'Dim SumaDias As Integer
'Dim SumaDiasYaGenerados As Integer
'Dim fecha_desde As Date
'Dim Fec_Fin As Date
'Dim TipoDia_Ok As Boolean
'Dim Dias_Mes_Anterior As Integer
'
'Dim rs_Estructura As New ADODB.Recordset
'Dim rs_tipd_con As New ADODB.Recordset
'Dim rs_Lic As New ADODB.Recordset
'Dim rs_Simproceso As New ADODB.Recordset
'
'
'    Bien = False
'    Valor = 0
'
'    StrSql = "SELECT * FROM simproceso "
'    StrSql = StrSql & " WHERE simpronro = " & simpronro
'
'    OpenRecordset StrSql, rs_Simproceso
'
'    If rs_Simproceso.EOF Then
'    ' Entrada Log
'    End If
'
''FGZ - 29/01/2004
'fecha_desde = rs_Simproceso!profecini
'Fec_Fin = rs_Simproceso!profecfin
'
'' Primero Busco  los tipos de dias asociados a los conceptos
'' Todos las Licencias del tipo especificado
'    StrSql = " SELECT * FROM tipd_con " & _
'             " WHERE concnro =" & nro_concepto & _
'             " AND tdnro > 2"
'
'OpenRecordset StrSql, rs_tipd_con
'If rs_tipd_con.EOF Then
'        Flog.writeline Espacios(Tabulador * 4) & "no hay tipos de dias asociados a los conceptos "
'End If
'
'Do While Not rs_tipd_con.EOF
'        Flog.writeline Espacios(Tabulador * 4) & "Tipo de dia: " & rs_tipd_con!tdnro
'    TipoDia_Ok = True
'    If Not EsNulo(rs_tipd_con!tenro) Then
'        If rs_tipd_con!tenro <> 0 Then
'            StrSql = " SELECT * FROM his_estructura " & _
'                     " WHERE ternro = " & empternro & " AND " & _
'                     " tenro =" & rs_tipd_con!tenro & " AND " & _
'                     " estrnro = " & rs_tipd_con!estrnro & " AND " & _
'                     " (htetdesde <= " & ConvFecha(Fec_Fin) & ") AND " & _
'                     " ((" & ConvFecha(Fec_Fin) & " <= htethasta) or (htethasta is null))"
'            OpenRecordset StrSql, rs_Estructura
'            If rs_Estructura.EOF Then
'                Flog.writeline Espacios(Tabulador * 4) & "Tipo de dia " & rs_tipd_con!tdnro & " no valido. No tiene estructura del tipo " & rs_tipd_con!tenro
'                TipoDia_Ok = False
'            End If
'        End If
'    End If
'
'    If CBool(TipoDia_Ok) Then
'        StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & empternro & " )" & _
'                 " AND tdnro =" & rs_tipd_con!tdnro & _
'                 " AND elfechadesde <=" & ConvFecha(Fec_Fin) & _
'                 " AND elfechahasta >= " & ConvFecha(fecha_desde)
'        OpenRecordset StrSql, rs_Lic
'
'        dias = 0
'        Do While Not rs_Lic.EOF
'            dias = CantidadDeDias(fecha_desde, Fec_Fin, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
'            'reviso si la licencia es completa
'            If Todas Then 'Todos los tipos de Licencias
'                Dias_Mes_Anterior = Dias_Licencias_Mes_Anterior(empternro, DateAdd("m", -1, fecha_desde), fecha_desde - 1)
'                If Dias_Mes_Anterior = 30 Then
'                    'calculo los dias reales del mes
'                    Dias_Mes_Anterior = DateDiff("d", DateAdd("m", -1, fecha_desde), fecha_desde - 1) + 1
'                    dias = dias + (Dias_Mes_Anterior - 30)
'                End If
'            Else
'                ' solo este tipo
'                If rs_Lic!elfechadesde <= DateAdd("m", -1, fecha_desde) Then
'                    If rs_Lic!elfechahasta >= DateAdd("m", -1, Fec_Fin) Then
'                        'Para ajustar la cantidad de dias cuando la lic sobrepasa al mes y fue topeada
'                        Dias_Mes_Anterior = DateDiff("d", DateAdd("m", -1, fecha_desde), fecha_desde - 1) + 1
'                        dias = dias + (Dias_Mes_Anterior - 30)
'                    End If
'                End If
'            End If
'            SumaDias = SumaDias + dias
'
'            'Marco la licencia para que no se pueda Borrar
''            StrSql = "UPDATE emp_lic SET pronro = " & simpronro & _
' '                    " WHERE emp_licnro = " & rs_Lic!emp_licnro
'  '          objConn.Execute StrSql, , adExecuteNoRecords
'
'            rs_Lic.MoveNext
'        Loop
'    End If
'    rs_tipd_con.MoveNext
'Loop
'
'' --------------------------------------------
'' FGZ - 29/01/2004
'' Buscar todas las licencias
'' Busco los detliq (campo dlicant "cantidad") del periodo cuyas licencias emp_lic esten marcadas (en pronro)
'' este valor +  SumaDias no debe seperar 30 dias
'' FGZ - 29/01/2004
'' QUEDA PENDIENTE
'' --------------------------------------------
'
'If Month(fecha_desde) = 2 Then 'Febrero
'    If Biciesto(Year(fecha_desde)) Then
'        If SumaDias >= 29 Then
'            Valor = 30
'        Else
'            Valor = SumaDias
'        End If
'    Else
'        If SumaDias >= 28 Then
'            Valor = 30
'        Else
'            Valor = SumaDias
'        End If
'    End If
'Else
'    If SumaDias > 30 Then
'        Valor = 30
'    Else
'        Valor = SumaDias
'    End If
'End If
'Bien = True
'
'bus_Licencias = Valor
'
'' Cierro todo y libero
'If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
'Set rs_Estructura = Nothing
'
'If rs_Lic.State = adStateOpen Then rs_Lic.Close
'Set rs_Lic = Nothing
'
'If rs_tipd_con.State = adStateOpen Then rs_tipd_con.Close
'Set rs_tipd_con = Nothing
'
'End Function
'
'
'Public Function Dias_Licencias_Mes_Anterior(ByVal Tercero As Long, ByVal fecha_desde As Date, ByVal Fec_Fin As Date) As Integer
'Dim rs_Lic As New ADODB.Recordset
'Dim dias As Integer
'
'    dias = 0
'
'    StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & Tercero & " )" & _
'             " AND elfechadesde <=" & ConvFecha(Fec_Fin) & _
'             " AND elfechahasta >= " & ConvFecha(fecha_desde)
'    OpenRecordset StrSql, rs_Lic
'
'    Do While Not rs_Lic.EOF
'        dias = CantidadDeDias(fecha_desde, Fec_Fin, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
'        rs_Lic.MoveNext
'    Loop
'
'    If Month(fecha_desde) = 2 Then 'Febrero
'        If Biciesto(Year(fecha_desde)) Then
'            If dias >= 29 Then
'                dias = 30
'            End If
'        Else
'            If dias >= 28 Then
'                dias = 30
'            End If
'        End If
'    Else
'        If dias > 30 Then
'            dias = 30
'        End If
'    End If
'
'    Dias_Licencias_Mes_Anterior = dias
'
'    'cierro
'    If rs_Lic.State = adStateOpen Then rs_Lic.Close
'    Set rs_Lic = Nothing
'End Function
'
'

