Attribute VB_Name = "ExpWord"
' ---------------------------------------------------------------------------------------------
' Descripcion: Proyecto encargado de generar la exportacion del Word Meeting
' Autor      : Martin Ferraro
' Fecha      : 21/12/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "21/12/2005"
'Global Const UltimaModificacion = " " 'Version Inicial

Global Const Version = "1.02"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = " " 'MB - Encriptacion de string connection

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f
'Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

'NUEVAS
Global EmpErrores As Boolean

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global tprocNro As Integer
Global tprocDesc As String
Global proDesc As String
Global concnro As Integer
Global Conccod As String
Global concabr As String
Global tconnro As Integer
Global tconDesc As String
Global concimp As Integer
Global concpuente As Integer
Global fecEstr As String
Global Formato As Integer
Global Modelo As Long
Global TituloRep As String
Global descDesde
Global descHasta
Global FechaHasta
Global FechaDesde
Global ArchExp
Global UsaEncabezado As Integer
Global Encabezado As Boolean

Global detalleSexo As Boolean
Global detallePobAct As Boolean
Global Sep As String
Global listaEmp As String
Global ArrIng(24) As Long
Global ArrIngSt(24) As Long
Global ArrEgre(24) As Long
Global ArrPact(24) As Long
Global ArrPactAnt(24) As Long
Global ArrMasc(24) As Long
Global ArrFem(24) As Long
Global IngTotal As Long
Global IngStTotal As Long
Global EgrTotal As Long
Global MascTotal As Long
Global FemTotal As Long
Global PactivaTotal As Long

Global errConfrep As Boolean



Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial exportacion de Word Meeting.
' Autor      : Martin Ferraro
' Fecha      : 21/12/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim pliqNro As Long
Dim Lista_Pronro As String
Dim PID As String
Dim Parametros As String
Dim ArrParametros
Dim param
Dim listapronro
Dim proNro
Dim ternro As Long
Dim arrpronro
Dim Periodos
Dim rsEmpl As New ADODB.Recordset
Dim totalEmpleados
Dim cantRegistros
Dim objRs As New ADODB.Recordset
Dim rsPeriodos As New ADODB.Recordset
Dim orden
Dim NroModelo As Long
Dim Directorio As String
Dim Carpeta
'Dim Nombre_Arch As String
'Dim rs As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim fs1

Dim mes As Integer
Dim Anio As Integer
Dim cant As Integer
Dim mes_actual As Integer
Dim anio_actual As Integer

Dim total_estr_mes_egr As Long
Dim total_estr_mes_ing As Long
Dim total_estr_mes_ing_sin_trans As Long
Dim total_estr_mes_pactiva As Long
Dim total_estr_mes_masc As Long
Dim total_estr_mes_fem As Long
Dim total_estr_mes_pactiva_ant As Long

Dim total_estr_egr As Long
Dim total_estr_ing As Long
Dim total_estr_ing_sin_trans As Long
Dim total_estr_pactiva As Long
Dim total_estr_masc As Long
Dim total_estr_fem As Long
Dim total_estr_pactiva_ant As Long

Dim rsEstr As New ADODB.Recordset
Dim I As Long
Dim estrnro As Long
Dim estrdabr
Dim Desde As Date
Dim Hasta As Date
Dim linea As String
Dim cantMasc As Long
Dim cantFem As Long

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProceso = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProceso = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If
    
    
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

    On Error Resume Next
    'Abro la conexion
'    OpenConnection strconexion, objConn
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
'    OpenConnection strconexion, objconnProgreso
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
    On Error GoTo 0

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExportacionWordMeeting" & "-" & NroProceso & ".log"
    
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
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo 0

    On Error GoTo ME_Main
    
    Flog.writeline "Inicio Proceso de Exportación Word Meeting : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 119"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       'Obtengo el mes
       mes = CInt(ArrParametros(0))
       'Obtengo el anio
       Anio = CInt(ArrParametros(1))
       'Obtengo la cantidad de meses
       cant = CInt(ArrParametros(2))
       
       Flog.writeline Espacios(Tabulador * 0) & "PARAMETROS"
       Flog.writeline Espacios(Tabulador * 1) & "Mes = " & mes
       Flog.writeline Espacios(Tabulador * 1) & "Anio = " & Anio
       Flog.writeline Espacios(Tabulador * 1) & "Cant = " & cant
       
       NroModelo = 271
    
       'Directorio de exportacion
       StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
       If rs.State = adStateOpen Then rs.Close
       OpenRecordset StrSql, rs
       If Not rs.EOF Then
          Directorio = Trim(rs!sis_dirsalidas)
       End If
     
       StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
       OpenRecordset StrSql, rs_Modelo
       If Not rs_Modelo.EOF Then
          If Not IsNull(rs_Modelo!modarchdefault) Then
             Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
          Else
             Flog.writeline Espacios(Tabulador * 0) & "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
          End If
       Else
          Flog.writeline Espacios(Tabulador * 0) & "No se encontró el modelo " & NroModelo & ". El archivo será generado en el directorio default"
       End If
                
       'Obtengo los datos del separador
       Sep = rs_Modelo!modseparador
       UsaEncabezado = rs_Modelo!modencab
       Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep
       
       If UsaEncabezado = -1 Then
          Encabezado = True
          Flog.writeline Espacios(Tabulador * 0) & "Con Encabezado."
       Else
          Encabezado = False
          Flog.writeline Espacios(Tabulador * 0) & "Sin Encabezado."
       End If
       
       Nombre_Arch = Directorio & "\word_meeting" & NroProceso & ".csv"
       Flog.writeline Espacios(Tabulador * 0) & "Se crea el archivo: " & Nombre_Arch
       Set fs = CreateObject("Scripting.FileSystemObject")
       On Error Resume Next
       If Err.Number <> 0 Then
          Flog.writeline Espacios(Tabulador * 0) & "La carpeta Destino no existe. Se creará."
          Set Carpeta = fs1.CreateFolder(Directorio)
       End If
       'desactivo el manejador de errores
       Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
       
       errConfrep = False
       'Cargo la configuracion del reporte
       Call CargarConfiguracionReporte
       If errConfrep Then
            Exit Sub
       End If
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       
       'Obtengo los empleados sobre del filtro
       Flog.writeline Espacios(Tabulador * 0) & "Se van a cargar los empleados."
       Call CargarEmpleados(NroProceso, rsEmpl)
       'Obtengo las estructuras a analizar
       Flog.writeline Espacios(Tabulador * 0) & "Se van a cargar las estruturas."
       Call CargarEstructuras(tenro1, rsEstr)
       
       'seteo de las variables de progreso
       Progreso = 0
       cantRegistros = cant * rsEmpl.RecordCount * rsEstr.RecordCount
       totalEmpleados = cant * rsEmpl.RecordCount * rsEstr.RecordCount
       If cantRegistros = 0 Then
          cantRegistros = 1
          Flog.writeline Espacios(Tabulador * 0) & "No se encontraron datos a Exportar."
       End If
       IncPorc = (100 / cantRegistros)
          
       If Encabezado Then
            Call imprime_encabezado(mes, Anio, cant)
       End If
       
       'Armo la lista de empleados del filtro
       listaEmp = "0"
       Do Until rsEmpl.EOF
            listaEmp = listaEmp & "," & rsEmpl!ternro
            rsEmpl.MoveNext
       Loop
       If listaEmp = "0" Then
            Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los empleados."
       End If
       rsEmpl.MoveFirst
       
       'Inicializo el total por columna
       For I = 0 To (cant - 1)
            ArrIng(I) = 0
            ArrEgre(I) = 0
            ArrIngSt(I) = 0
            ArrPact(I) = 0
            ArrPactAnt(I) = 0
            ArrMasc(I) = 0
            ArrFem(I) = 0
       Next
       
       EgrTotal = 0
       IngTotal = 0
       IngStTotal = 0
       MascTotal = 0
       FemTotal = 0
       PactivaTotal = 0
    
       'Por cada estructura
       Do Until rsEstr.EOF
           
           'Inicializo variables de la estructura
           estrnro = rsEstr!estrnro
           estrdabr = rsEstr!estrdabr
           total_estr_egr = 0
           total_estr_ing = 0
           total_estr_ing_sin_trans = 0
           total_estr_pactiva = 0
           total_estr_masc = 0
           total_estr_fem = 0
           
           linea = estrdabr
           
           Flog.writeline Espacios(Tabulador * 0) & "Se comenzo a procesar las estructura = " & estrnro & " - " & estrdabr
           
           'Por cada mes
           For I = 0 To (cant - 1)
                           
               'Inicializo variables del mes
               Call calcula_anio_mes(mes, Anio, I, mes_actual, anio_actual)
               Desde = primer_dia_mes(mes_actual, anio_actual)
               Hasta = ultimo_dia_mes(mes_actual, anio_actual)
               total_estr_mes_egr = 0
               total_estr_mes_ing = 0
               total_estr_mes_ing_sin_trans = 0
               total_estr_mes_pactiva = 0
               total_estr_mes_masc = 0
               total_estr_mes_fem = 0
               total_estr_mes_pactiva_ant = 0
               
               Flog.writeline Espacios(Tabulador * 1) & "Procesando el mes= " & mes_actual & " Año= " & anio_actual
               
               'Por cada empleado
               Do Until rsEmpl.EOF
                   'Inicializo variables del empleado
                   ternro = rsEmpl!ternro
                   Flog.writeline Espacios(Tabulador * 2) & "Procesando el empleado= " & ternro
                   
                   Call ContarIngEgr(Desde, Hasta, estrnro, ternro, total_estr_mes_egr, total_estr_mes_ing, total_estr_mes_ing_sin_trans)

                   rsEmpl.MoveNext
                   
                   'Actualizo el estado del proceso
                   TiempoAcumulado = GetTickCount
                   cantRegistros = cantRegistros - 1
                   StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                            ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                            ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                      
                   objConn.Execute StrSql, , adExecuteNoRecords
                   
               Loop
               rsEmpl.MoveFirst
               
               If detallePobAct Then
                    Call CalculaPoblacionActiva(Desde, estrnro, total_estr_mes_masc, total_estr_mes_fem, total_estr_mes_pactiva_ant)
                    Call CalculaPoblacionActiva(Hasta, estrnro, total_estr_mes_masc, total_estr_mes_fem, total_estr_mes_pactiva)
                   
                    'inserto p activa anterior en la linea
                    linea = linea & Sep & total_estr_mes_pactiva_ant
               End If
               
               'Armo la linea del mes
               linea = linea & Sep & total_estr_mes_ing & Sep & total_estr_mes_egr & Sep & total_estr_mes_ing_sin_trans
               
               If detallePobAct Then
                    linea = linea & Sep & total_estr_mes_pactiva
                    If detalleSexo Then
                        linea = linea & Sep & total_estr_mes_masc & Sep & total_estr_mes_fem
                    End If
               End If
               
               'Totalizo por estructura
               total_estr_egr = total_estr_egr + total_estr_mes_egr
               total_estr_ing = total_estr_ing + total_estr_mes_ing
               total_estr_ing_sin_trans = total_estr_ing_sin_trans + total_estr_mes_ing_sin_trans
               total_estr_pactiva = total_estr_pactiva + total_estr_mes_pactiva
               total_estr_pactiva_ant = total_estr_pactiva_ant + total_estr_mes_pactiva_ant
               total_estr_masc = total_estr_masc + total_estr_mes_masc
               total_estr_fem = total_estr_fem + total_estr_mes_fem
               
               'Totalizo por columna
               ArrIng(I) = ArrIng(I) + total_estr_mes_ing
               ArrEgre(I) = ArrEgre(I) + total_estr_mes_egr
               ArrIngSt(I) = ArrIngSt(I) + total_estr_mes_ing_sin_trans
               ArrPact(I) = ArrPact(I) + total_estr_mes_pactiva
               ArrPactAnt(I) = ArrPactAnt(I) + total_estr_mes_pactiva_ant
               ArrMasc(I) = ArrMasc(I) + total_estr_mes_masc
               ArrFem(I) = ArrFem(I) + total_estr_mes_fem
               IngTotal = IngTotal + total_estr_mes_ing
               IngStTotal = IngStTotal + total_estr_mes_ing_sin_trans
               EgrTotal = EgrTotal + total_estr_mes_egr
               MascTotal = MascTotal + total_estr_mes_masc
               FemTotal = FemTotal + total_estr_mes_fem
               PactivaTotal = PactivaTotal + total_estr_mes_pactiva
               
           Next
           
           'Agrego el total de la estructura
           linea = linea & Sep & total_estr_ing & Sep & total_estr_egr & Sep & total_estr_ing_sin_trans
           
           If detallePobAct Then
                linea = linea & Sep & total_estr_pactiva
                If detalleSexo Then
                    linea = linea & Sep & total_estr_masc & Sep & total_estr_fem
                End If
           End If
           
            
           'Imprimo la linea
           ArchExp.writeline linea
            

           rsEstr.MoveNext
           
       Loop
       rsEmpl.Close
       rsEstr.Close
          
       Call ImprimeTotalesColum(cant)
          
       ArchExp.Close
    
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
          
    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    Set rs_Modelo = Nothing

    'Actualizo el estado del proceso
    If Not HuboErrores Then
        
        'Borro los empleado
        StrSql = " DELETE FROM batch_empleado "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND btprcnro = 119"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "cant open " & Cantidad_de_OpenRecordset

    TiempoFinalProceso = GetTickCount
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    objconnProgreso.Close
    objConn.Close
    
Exit Sub
    
ME_Main:
    HuboErrores = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
    
End Sub


Sub CargarConfiguracionReporte()

    Dim objRs As New ADODB.Recordset
    Dim objRs2 As New ADODB.Recordset
    Dim StrSql As String
    Dim I
    Dim columnaActual
    Dim Nro_col
    Dim Valor As Long
    
    StrSql = " SELECT * FROM confrep WHERE confrep.repnro= 153 "
    StrSql = StrSql & "ORDER BY confnrocol "
    
    OpenRecordset StrSql, objRs
    
    tenro1 = 0
    detalleSexo = False
    detallePobAct = False
    
    If objRs.EOF Then
        errConfrep = True
        Flog.writeline Espacios(Tabulador * 0) & "No se encontro la configuracion del reporte 153"
        Exit Sub
    End If
    
    Do Until objRs.EOF
       
       columnaActual = CLng(objRs!confnrocol)
       
       If columnaActual = 1 Then     'Tipo Estructura 1
          tenro1 = objRs!confval
       ElseIf columnaActual = 2 Then 'Si lleva detalle de poblacion activa
          If objRs!confval = -1 Then
              detallePobAct = True
          End If
       ElseIf columnaActual = 3 Then 'Si lleva detalle de sexo
          If objRs!confval = -1 Then
              detalleSexo = True
          End If
       End If
       
       objRs.MoveNext
       
    Loop
    objRs.Close
    
    If tenro1 = 0 Then
        errConfrep = True
        Flog.writeline Espacios(Tabulador * 0) & "Falta configurar la estructura en el reporte 153"
        Exit Sub
    End If
    

End Sub


Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    StrEmpl = StrEmpl & " ORDER BY ternro "
    
    OpenRecordset StrEmpl, rsEmpl
End Sub


Sub calcula_anio_mes(mes, Anio, cuenta, ByRef calc_mes As Integer, ByRef calc_anio As Integer)

    If mes + cuenta <= 12 Then
        calc_mes = mes + cuenta
        calc_anio = Anio
    Else
        calc_mes = (mes + cuenta) Mod 12
        calc_anio = Anio + Fix((mes + cuenta) / 12)
    End If
    
End Sub


Sub CargarEstructuras(tenro, ByRef rsEstr As ADODB.Recordset)

Dim StrEstr As String

    StrEstr = "SELECT * FROM estructura WHERE tenro = " & tenro
    StrEstr = StrEstr & " ORDER BY estrdabr "
    
    OpenRecordset StrEstr, rsEstr
End Sub


Function primer_dia_mes(mes As Integer, Anio As Integer) As Date
Dim aux As String
    primer_dia_mes = C_Date("01/" & mes & "/" & Anio)
    
End Function



Function ultimo_dia_mes(mes As Integer, Anio As Integer) As Date

Dim mes_sgt As Integer
Dim anio_sgt As Integer

    If mes = 12 Then
        mes_sgt = 1
        anio_sgt = Anio + 1
    Else
        mes_sgt = mes + 1
        anio_sgt = Anio
    End If
    
    ultimo_dia_mes = DateAdd("d", -1, primer_dia_mes(mes_sgt, anio_sgt))
    
End Function


Sub ContarIngEgr(fecdesde As Date, fechasta As Date, estrnro As Long, ternro As Long, ByRef total_egr As Long, ByRef total_ing As Long, ByRef total_ing_sin_trans As Long)
          
Dim SqlCont As String
Dim rsCont As New ADODB.Recordset
Dim rsFases As New ADODB.Recordset

    'Busco la estructura del empleado con la fecha desde o hasta dentro del rango
    SqlCont = " SELECT ternro, estrnro, htetdesde, htethasta "
    SqlCont = SqlCont & " FROM his_estructura "
    SqlCont = SqlCont & " WHERE his_estructura.ternro = " & ternro
    SqlCont = SqlCont & " AND his_estructura.estrnro = " & estrnro
    SqlCont = SqlCont & " AND ( (his_estructura.htetdesde <= " & ConvFecha(fechasta) & " AND " & ConvFecha(fecdesde) & " <= his_estructura.htetdesde) "
    SqlCont = SqlCont & "   OR (his_estructura.htethasta <= " & ConvFecha(fechasta) & " AND " & ConvFecha(fecdesde) & " <= his_estructura.htethasta) )"
    OpenRecordset SqlCont, rsCont
    
    Do Until rsCont.EOF
        'Ingreso
        If (rsCont!htetdesde <= fechasta And fecdesde <= rsCont!htetdesde) Then
        
            total_ing = total_ing + 1
            'Busco la fase del empleado con la fecha desde igual a la de la fase
            SqlCont = " SELECT fasnro, empleado, altfec "
            SqlCont = SqlCont & " FROM fases "
            SqlCont = SqlCont & " WHERE fases.empleado = " & ternro
            SqlCont = SqlCont & " AND fases.altfec = " & ConvFecha(rsCont!htetdesde)
            OpenRecordset SqlCont, rsFases
            If Not rsFases.EOF Then
                total_ing_sin_trans = total_ing_sin_trans + 1
            End If
            rsFases.Close
        
        End If
        
        'Egreso
        If (rsCont!htethasta <= fechasta And fecdesde <= rsCont!htethasta) Then
        
            'Busco la fase del empleado con la fecha hasta igual a la de la fase
            SqlCont = " SELECT fasnro, empleado, bajfec "
            SqlCont = SqlCont & " FROM fases "
            SqlCont = SqlCont & " WHERE fases.empleado = " & ternro
            SqlCont = SqlCont & " AND fases.bajfec = " & ConvFecha(rsCont!htethasta)
            OpenRecordset SqlCont, rsFases
            If Not rsFases.EOF Then
                total_egr = total_egr + 1
            End If
            rsFases.Close
        
        End If
        
        rsCont.MoveNext
        
    Loop
    rsCont.Close
    
          

End Sub



Sub imprime_encabezado(mes As Integer, Anio As Integer, cant As Integer)

Dim linea As String
Dim linea2 As String
Dim I As Integer
Dim mes_actual As Integer
Dim anio_actual As Integer

    linea = " "
    linea2 = " "
    For I = 0 To (cant - 1)
    
        Call calcula_anio_mes(mes, Anio, I, mes_actual, anio_actual)
        
        linea = linea & Sep & NombreMes(mes_actual, anio_actual)
        
        If detallePobAct Then
            linea2 = linea2 & Sep & "Pob. Activa Anterior"
            linea = linea & Sep & " " & Sep & " "
        End If
        
        linea2 = linea2 & Sep & "Ingreso" & Sep & "Egreso" & Sep & "Ingreso s/trans"
        
        If detallePobAct Then
            linea2 = linea2 & Sep & "Pob. Activa"
            If detalleSexo Then
                linea2 = linea2 & Sep & "Masc" & Sep & "Fem"
                linea = linea & Sep & " " & Sep & " "
            End If
        End If
        
        linea = linea & Sep & " " & Sep & " "
        
    Next
        
    'Total
    linea = linea & Sep & "Total" & Sep & " " & Sep & " "
    linea2 = linea2 & Sep & "Ingreso" & Sep & "Egreso" & Sep & "Ingreso s/trans"
    
    If detallePobAct Then
        linea2 = linea2 & Sep & "Pob. Activa"
        linea = linea & Sep & " "
        If detalleSexo Then
            linea2 = linea2 & Sep & "Masc" & Sep & "Fem"
            linea = linea & Sep & " " & Sep & " "
        End If
    End If

    ArchExp.writeline linea
    ArchExp.writeline linea2
    
End Sub


Function NombreMes(mes As Integer, Anio As Integer) As String

Dim Texto As String

    Select Case mes
        Case 1
            Texto = "Enero "
        Case 2
            Texto = "Febrero "
        Case 3
            Texto = "Marzo "
        Case 4
            Texto = "Abril "
        Case 5
            Texto = "Mayo "
        Case 6
            Texto = "Junio "
        Case 7
            Texto = "Julio "
        Case 8
            Texto = "Agosto "
        Case 9
            Texto = "Septiembre "
        Case 10
            Texto = "Octubre "
        Case 11
            Texto = "Noviembre "
        Case 12
            Texto = "Diciembre "
    End Select
    
    NombreMes = Texto & CStr(Anio)
End Function


Sub CalculaPoblacionActiva(Fecha As Date, estrnro As Long, ByRef masc As Long, ByRef fem As Long, ByRef total As Long)
Dim SqlAct As String
Dim rsAct As New ADODB.Recordset

    SqlAct = " SELECT distinct his_estructura.ternro "
    SqlAct = SqlAct & " FROM his_estructura "
    SqlAct = SqlAct & " INNER JOIN tercero ON tercero.ternro = his_estructura.ternro AND tersex = -1 "
    SqlAct = SqlAct & " AND tercero.ternro in (" & listaEmp & ")"
    SqlAct = SqlAct & " WHERE his_estructura.estrnro = " & estrnro
    SqlAct = SqlAct & " AND his_estructura.htetdesde < " & ConvFecha(Fecha)
    SqlAct = SqlAct & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR htethasta is null)"
    OpenRecordset SqlAct, rsAct
    
    masc = rsAct.RecordCount
    rsAct.Close
    
    SqlAct = " SELECT distinct his_estructura.ternro "
    SqlAct = SqlAct & " FROM his_estructura "
    SqlAct = SqlAct & " INNER JOIN tercero ON tercero.ternro = his_estructura.ternro AND tersex <> -1 "
    SqlAct = SqlAct & " AND tercero.ternro in (" & listaEmp & ")"
    SqlAct = SqlAct & " WHERE his_estructura.estrnro = " & estrnro
    SqlAct = SqlAct & " AND his_estructura.htetdesde < " & ConvFecha(Fecha)
    SqlAct = SqlAct & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR htethasta is null)"
    OpenRecordset SqlAct, rsAct
    
    fem = rsAct.RecordCount
    rsAct.Close
    
    total = fem + masc
    
End Sub


Sub ImprimeTotalesColum(Cantidad As Integer)
Dim I As Integer
Dim linea As String

    linea = "Total"
    
    For I = 0 To Cantidad - 1
        If detallePobAct Then
            linea = linea & Sep & ArrPactAnt(I)
        End If
        linea = linea & Sep & ArrIng(I) & Sep & ArrEgre(I) & Sep & ArrIngSt(I)
        If detallePobAct Then
            linea = linea & Sep & ArrPact(I)
            If detalleSexo Then
                linea = linea & Sep & ArrMasc(I) & Sep & ArrFem(I)
            End If
        End If
    Next
    
    linea = linea & Sep & IngTotal & Sep & EgrTotal & Sep & IngStTotal
    
    If detallePobAct Then
        linea = linea & Sep & PactivaTotal
        If detalleSexo Then
            linea = linea & Sep & MascTotal & Sep & FemTotal
        End If
    End If

    ArchExp.writeline linea
    
End Sub
