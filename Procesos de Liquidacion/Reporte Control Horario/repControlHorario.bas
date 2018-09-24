Attribute VB_Name = "repControlHorario"
Option Explicit

'Version: 1.01
'

'Const Version = 1.01
'Const FechaVersion = "14/09/2005"

'Const Version = "1.02"
'Const FechaVersion = "04/07/2007"   'FGZ - problemas de longitud de campos cuando se insertan datos en las tablas rep_ctrl_hor

Global Const Version = "1.03" ' Cesar Stankunas
Global Const FechaVersion = "05/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

'================================================================================
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

Global tipoHora As String
Global TipoDia As Long
Global FechaDesde As Date
Global FechaHasta As Date

Global Tenro1 As Long
Global Tenro2 As Long
Global Tenro3 As Long

Global Estrnro1 As Long
Global Estrnro2 As Long
Global Estrnro3 As Long

Global EmpEstrnro1 As Long
Global EmpEstrnro2 As Long
Global EmpEstrnro3 As Long


Global rsConsult As New ADODB.Recordset
Global rsConsult2 As New ADODB.Recordset
Global rsConsult3 As New ADODB.Recordset
Global rsConsult4 As New ADODB.Recordset
Global rsConsult5 As New ADODB.Recordset
Global rsConsult6 As New ADODB.Recordset

Global Columnas(10) As String
Global ColumnasTitulos(10) As String
Global ValorColumna(10) As Single

'DATOS DE LA TABLA batch_proceso
Global bpfecha As Date
Global bphora As String
Global bpusuario As String

Global repNro As Long
Global conceptos As String
Global acumuladores As String
Global procesos As String
Global idUser As String

Private Sub Main()
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim ternro
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim rsEmpl As New ADODB.Recordset
Dim objRs As New ADODB.Recordset

Dim PID As String
Dim ArrParametros
Dim parametros As String

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

    Nombre_Arch = PathFLog & "ReporteControlHorario" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
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
    
    On Error GoTo CE
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CLng(objRs!bprcempleados)
    totalEmpleados = cantRegistros
    
    objRs.Close
    
    Flog.writeline "Inicio Proceso de Control Pagos : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
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
       ArrParametros = Split(parametros, "@")
       
       'Obtengo el Tipo de Hora
       TipoDia = ArrParametros(0)
       
       'Obtengo el Tipo de Dia
       tipoHora = ArrParametros(1)
       
       'Obtengo las estructuras
       Tenro1 = CLng(ArrParametros(2))
       Estrnro1 = CLng(ArrParametros(3))
       Tenro2 = CLng(ArrParametros(4))
       Estrnro2 = CLng(ArrParametros(5))
       Tenro3 = CLng(ArrParametros(6))
       Estrnro3 = CLng(ArrParametros(7))
       
       'Obtengo las fechas
       FechaDesde = objRs!bprcfecdesde
       FechaHasta = objRs!bprcfechasta
       
       'Obtengo la fecha del proceso
       bpfecha = objRs!bprcfecha
       
       'Obtengo la hora del proceso
       bphora = objRs!bprchora
       
       'Obtengo el usuario del proceso
       bpusuario = objRs!idUser
                     
       'Obtengo el titulo del reporte
       'tituloReporte = arrParametros(5)
    
       'EMPIEZA EL PROCESO
       
       'Cargo los datos del confrep
       Call cargarConfRep(62)
       
       'Obtengo los empleados sobre los que tengo que generar los recibos
       Call CargarEmpleados(NroProceso, rsEmpl)
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
       
       If Not rsEmpl.EOF Then
          Call crearCabezera
       End If
       
       'Genero por cada empleado un recibo de sueldo
       Do Until rsEmpl.EOF

         ternro = rsEmpl!ternro
         Flog.writeline "Generando datos para el empleado: " & rsEmpl!empleg
         
         Call generarControlHorario(CLng(rsEmpl!empleg), ternro, FechaDesde, FechaHasta, TipoDia, tipoHora)
         
         'Actualizo el estado del proceso
          TiempoAcumulado = GetTickCount
           
          cantRegistros = cantRegistros - 1
        
          Flog.writeline "Progreso " & (((totalEmpleados - cantRegistros) * 100) / totalEmpleados)
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & (((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
          objConn.Execute StrSql, , adExecuteNoRecords
          Flog.writeline "Progreso actualizado"
         
          'Si se generaron todos los recibos de sueldo del empleado correctamente lo borro
          If Not EmpErrores Then
              StrSql = " DELETE FROM batch_empleado "
              StrSql = StrSql & " WHERE bpronro = " & NroProceso
              StrSql = StrSql & " AND ternro = " & ternro
    
              objConn.Execute StrSql, , adExecuteNoRecords
          End If
          
          rsEmpl.MoveNext
       Loop
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
    
    'cierro y libero
    If rsEmpl.State = adStateOpen Then rsEmpl.Close
    If objRs.State = adStateOpen Then objRs.Close
    If objConn.State = adStateOpen Then objConn.Close
    
    Set rsEmpl = Nothing
    Set objRs = Nothing
    
    Exit Sub
    
CE:
    HuboErrores = True

    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function

'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT batch_empleado.ternro, empleado.empleg FROM batch_empleado "
    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
    StrEmpl = StrEmpl & " WHERE batch_empleado.bpronro = " & NroProc & " ORDER BY batch_empleado.progreso"
    
    OpenRecordset StrEmpl, rsEmpl
End Sub


'FUNCION: Convierte un string que contiene una hora al formato string hora
Function convHora(Str)
  convHora = Mid(Str, 1, 2) & ":" & Mid(Str, 3, 2)
End Function 'convHora(str)
'FUNCION: Se encarga de controlar si la persona tiene los tipos de horas indicados para el rango de fechas
Function tieneHoras(ternro, Fecha)
    Dim Salida

    Salida = False
    
    If Trim(tipoHora) = "" Then
        Salida = True
    Else

        StrSql = " SELECT DISTINCT ternro, thnro, adcanthoras "
        StrSql = StrSql & " FROM gti_acumdiario "
        StrSql = StrSql & " WHERE ternro = " & ternro
        StrSql = StrSql & " AND thnro IN (" & tipoHora & ")"
        StrSql = StrSql & " AND adfecha = " & ConvFecha(Fecha)

        OpenRecordset StrSql, rsConsult2
        Salida = Not rsConsult2.EOF
        rsConsult2.Close

    End If

    tieneHoras = Salida

End Function
'Rutina que se encarga de buscar el horario teorico
Function obtenerTeorico(ByVal dia, ByVal ternro, ByRef TipoDiaActual)
Dim Salida
Dim hay_datos

    Flog.writeline "Buscando Horario Teorico"
    
    Salida = ""
    hay_datos = False
    StrSql = " SELECT * FROM gti_proc_emp "
    StrSql = StrSql & " LEFT JOIN gti_dias ON gti_dias.dianro = gti_proc_emp.dianro "
    StrSql = StrSql & " WHERE gti_proc_emp.ternro = " & ternro
    StrSql = StrSql & " AND gti_proc_emp.fecha  = " & ConvFecha(dia)
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       If rsConsult!Feriado = -1 Then
          Salida = "Feriado"
          TipoDiaActual = 2
       Else
          If rsConsult!pasig = -1 Then
               'Busco el parte
               StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & ternro & ") AND " & _
                        " (gttempdesde <= " & ConvFecha(dia) & ") AND " & _
                        " (" & ConvFecha(dia) & " <= gttemphasta)"
               
               rsConsult.Close
               
               OpenRecordset StrSql, rsConsult
               
               TipoDiaActual = 4

               If Not rsConsult.EOF Then
               
                   If CLng(rsConsult!ttemplibre) = -1 Then
                      TipoDiaActual = 1
                   Else
                      TipoDiaActual = 0
                   End If

                   If Not IsNull(rsConsult!ttemphdesde1) Then
                      If rsConsult!ttemphdesde1 <> "" Then
                         Salida = Salida & convHora(rsConsult!ttemphdesde1) & "-" & convHora(rsConsult!ttemphhasta1)
                         hay_datos = True
                      End If
                   End If
                   If Not IsNull(rsConsult!ttemphdesde2) Then
                      If rsConsult!ttemphdesde2 <> "" Then
                         If hay_datos Then
                            Salida = Salida & "<br>"
                         End If
                         hay_datos = True
                         Salida = Salida & convHora(rsConsult!ttemphdesde2) & "-" & convHora(rsConsult!ttemphhasta2)
                      End If
                   End If
                   If Not IsNull(rsConsult!ttemphdesde3) Then
                      If rsConsult!ttemphdesde3 <> "" Then
                         If hay_datos Then
                            Salida = Salida & "<br>"
                         End If
                         hay_datos = True
                         Salida = Salida & convHora(rsConsult!ttemphdesde3) & "-" & convHora(rsConsult!ttemphhasta3)
                      End If
                   End If

               Else
                   Salida = Salida & "Parte Asig. Hor. Eliminado"
               End If
           Else
             If rsConsult!dialibre = -1 Then
                 Salida = "Franco"
                 TipoDiaActual = 1
             Else
                 TipoDiaActual = 0
                 
                   If Not IsNull(rsConsult!diahoradesde1) Then
                      If Replace(rsConsult!diahoradesde1, "0", "") <> "" And Replace(rsConsult!diahorahasta1, "0", "") <> "" Then
                         Salida = Salida & convHora(rsConsult!diahoradesde1) & "-" & convHora(rsConsult!diahorahasta1)
                         hay_datos = True
                      End If
                   End If
                   If Not IsNull(rsConsult!diahoradesde2) Then
                      If Replace(rsConsult!diahoradesde2, "0", "") <> "" And Replace(rsConsult!diahorahasta2, "0", "") <> "" Then
                         If hay_datos Then
                            Salida = Salida & "<br>"
                         End If
                         hay_datos = True
                         Salida = Salida & convHora(rsConsult!diahoradesde2) & "-" & convHora(rsConsult!diahorahasta2)
                      End If
                   End If
                   If Not IsNull(rsConsult!diahoradesde3) Then
                      If Replace(rsConsult!diahoradesde3, "0", "") <> "" And Replace(rsConsult!diahorahasta3, "0", "") <> "" Then
                         If hay_datos Then
                            Salida = Salida & "<br>"
                         End If
                         hay_datos = True
                         Salida = Salida & convHora(rsConsult!diahoradesde3) & "-" & convHora(rsConsult!diahorahasta3)
                      End If
                   End If
             End If
          End If
       End If
    Else
       Salida = "Sin Procesar"
       TipoDiaActual = 4
    End If
    
    rsConsult.Close
    
    obtenerTeorico = Salida
End Function 'obtenerTeorico(dia,ternro)

'Rutina que muestra las registraciones de un empleado en un dia
Function mostrarRegistracion(dia, ternro)
 Dim Str
 Dim cantreg
 Dim Salida
     Salida = ""
     
    Flog.writeline "Buscando Registraciones"

    StrSql = "SELECT ternro, regfecha, reghora, regentsal, reldabr, regestado "
    StrSql = StrSql & " FROM gti_registracion INNER JOIN gti_reloj ON gti_registracion.relnro = gti_reloj.relnro "
    StrSql = StrSql & " WHERE gti_registracion.ternro = " & ternro
    StrSql = StrSql & " AND regfecha = " & ConvFecha(dia)
    StrSql = StrSql & " ORDER BY reghora ASC "
    
    OpenRecordset StrSql, rsConsult3

    If Not rsConsult3.EOF Then
       cantreg = 1
       
       'Cargo los valores en las columnas
        Do While Not rsConsult3.EOF And cantreg < 9
            If rsConsult3!reghora <> "" Then
               Str = Mid(rsConsult3!reghora, 1, 2) & ":" & Mid(rsConsult3!reghora, 3, 2) & "-" & rsConsult3!regentsal
               If CStr(rsConsult3!regestado) <> "P" Then
                  Str = Str
               Else
                  Str = Str
               End If
               If Salida = "" Then
                  Salida = Str
               Else
                  If (cantreg Mod 2) = 0 Then
                     Salida = Salida & "&nbsp;" & Str
                  Else
                     Salida = Salida & "<br>" & Str
                  End If
               End If
            End If
            cantreg = cantreg + 1
            rsConsult3.MoveNext
        Loop
    End If

    rsConsult3.Close
    
    mostrarRegistracion = Salida

End Function 'mostrarRegistracion(dia,ternro)

'Rutina que se encarga de buscar las anormalidades
Function buscarAnormalida(dia, ternro)
Dim Salida

    Salida = ""

    StrSql = " SELECT 'Licencia' as tipo,tipdia.tddesc as descripcion "
    StrSql = StrSql & " from gti_justificacion "
    StrSql = StrSql & " INNER JOIN emp_lic ON gti_justificacion.juscodext=emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro=tipdia.tdnro "
    StrSql = StrSql & " WHERE jussigla ='LIC' and gti_justificacion.ternro=" & ternro
    StrSql = StrSql & " and gti_justificacion.jusdesde <= " & ConvFecha(dia) & " and gti_justificacion.jushasta >= " & ConvFecha(dia)
    StrSql = StrSql & " UNION "
    StrSql = StrSql & " SELECT 'Novedad' as tipo, gti_novedad.gnovdesabr  as descripcion "
    StrSql = StrSql & " from gti_justificacion "
    StrSql = StrSql & " INNER JOIN  gti_novedad ON gti_justificacion.juscodext= gti_novedad.gnovnro "
    StrSql = StrSql & " where jussigla ='NOV' and gti_justificacion.ternro=" & ternro
    StrSql = StrSql & " and gti_justificacion.jusdesde <= " & ConvFecha(dia) & " and gti_justificacion.jushasta >= " & ConvFecha(dia)
    StrSql = StrSql & " UNION "
    StrSql = StrSql & " SELECT 'Curso' as tipo,'                         '  as descripcion "
    StrSql = StrSql & " from gti_justificacion where jussigla ='CUR' and gti_justificacion.ternro=" & ternro
    StrSql = StrSql & " and gti_justificacion.jusdesde <= " & ConvFecha(dia) & " and gti_justificacion.jushasta >= " & ConvFecha(dia)
    StrSql = StrSql & " UNION "
    StrSql = StrSql & " SELECT 'Suspencion' as tipo,'                        '  as descripcion "
    StrSql = StrSql & " from gti_justificacion where jussigla ='SUS' and gti_justificacion.ternro=" & ternro
    StrSql = StrSql & " and gti_justificacion.jusdesde <= " & ConvFecha(dia) & " and gti_justificacion.jushasta >= " & ConvFecha(dia)
    StrSql = StrSql & " UNION "
    StrSql = StrSql & " SELECT 'Anorm' as tipo,  gti_anormalidad.normdesabr as descripcion "
    StrSql = StrSql & " from gti_horcumplido "
    StrSql = StrSql & " INNER JOIN gti_anormalidad ON gti_horcumplido.normnro= gti_anormalidad.normnro "
    StrSql = StrSql & " where gti_horcumplido.ternro=" & ternro
    StrSql = StrSql & " and gti_horcumplido.horfecrep=" & ConvFecha(dia)
    StrSql = StrSql & " UNION "
    StrSql = StrSql & " SELECT 'Anorm' as tipo,  gti_anormalidad.normdesabr as descripcion "
    StrSql = StrSql & " from gti_horcumplido "
    StrSql = StrSql & " INNER JOIN gti_anormalidad ON gti_horcumplido.normnro2= gti_anormalidad.normnro "
    StrSql = StrSql & " WHERE gti_horcumplido.normnro2<> gti_horcumplido.normnro and gti_horcumplido.ternro=" & ternro
    StrSql = StrSql & " and gti_horcumplido.horfecrep=" & ConvFecha(dia)

    OpenRecordset StrSql, rsConsult4

    Do Until rsConsult4.EOF
        If Salida = "" Then
           Salida = "<b>" & rsConsult4!tipo & "</b>: " & rsConsult4!descripcion
        Else
           Salida = Salida & "<br><b>" & rsConsult4!tipo & "</b>: " & rsConsult4!descripcion
        End If
        rsConsult4.MoveNext
    Loop
    rsConsult4.Close

    buscarAnormalida = Salida
End Function 'buscarAnormalida(dia,ternro)

'Rutina que muestra las Licencias de un empleado en un dia
Function mostrarLicencias(dia, ternro)
 Dim Str
 Dim cantreg
 Dim Salida
     Salida = ""

    Flog.writeline "Buscando Licencias"
    
    StrSql = "SELECT gtnovdesabr "
    StrSql = StrSql & " FROM gti_novedad INNER JOIN gti_tiponovedad ON gti_tiponovedad.gtnovnro = gti_novedad.gtnovnro "
    StrSql = StrSql & " WHERE gti_novedad.gnovotoa = " & ternro
    StrSql = StrSql & " AND gnovdesde >= " & ConvFecha(dia)
    StrSql = StrSql & " AND gnovhasta <= " & ConvFecha(dia)
    StrSql = StrSql & " ORDER BY gtnovdesabr ASC "
    
    OpenRecordset StrSql, rsConsult5
    
    Salida = ""
    Do While Not rsConsult.EOF And cantreg < 9
        If Salida = "" Then
           Salida = rsConsult5!gtnovdesabr
        Else
            Salida = Salida & "<BR>" & rsConsult5!gtnovdesabr
        End If
        rsConsult5.MoveNext
    Loop
    rsConsult5.Close
    mostrarLicencias = Salida
End Function 'mostrarLicencias(dia,ternro)

'Rutina que muestra las Licencias de un empleado en un dia
Function mostrarAutorizadas(dia, ternro)
 Dim Str
 Dim cantreg
 Dim Salida
     Salida = ""

    Flog.writeline "Buscando Autorizadas"
    
    StrSql = "SELECT thdesc, gadhoras "
    StrSql = StrSql & " FROM gti_autdet INNER JOIN tiphora ON gti_autdet.thnro = tiphora.thnro "
    StrSql = StrSql & " WHERE gti_autdet.ternro = " & ternro
    StrSql = StrSql & " AND gadfecdesde >= " & ConvFecha(dia)
    StrSql = StrSql & " AND gadfechasta <= " & ConvFecha(dia)
    StrSql = StrSql & " ORDER BY thdesc ASC "
    
    OpenRecordset StrSql, rsConsult6
    
    Salida = ""
    Do While Not rsConsult6.EOF And cantreg < 9
        If Salida = "" Then
            Salida = rsConsult6!thdesc & " " & FormatNumber(rsConsult6!gadhoras, 2)
        Else
            Salida = Salida & "<BR>" & rsConsult6!thdesc & " " & FormatNumber(rsConsult6!gadhoras, 2)
        End If
        rsConsult6.MoveNext
    Loop
    rsConsult6.Close
    mostrarAutorizadas = Salida

End Function 'mostrarAutorizadas(dia,ternro)

Sub generarControlHorario(ByVal Legajo As Long, ByVal ternro As Long, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal TipoDia As Long, TipoHoras As String)
'--------------------------------------------------------------------
' Se encarga de buscar las auditorias
'--------------------------------------------------------------------
Dim StrSql As String
Dim rsConsult1 As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim Cantidad As Long
Dim cantidadProcesada As Long
Dim FechaActual As Date
Dim Contador
Dim Teorico As String
Dim Registraciones As String
Dim HorarioTeorico As String
Dim Fichada As String
Dim MotivoInasistencias As String
Dim HorasAutorizadas As String
Dim TipoDiaActual As Long
Dim Mostrar As Boolean
Dim cantCols
Dim I

Dim Insertar_Legajo As Boolean
 
On Error GoTo MError

Insertar_Legajo = False
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------
If Tenro1 <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "Buscando el tipo de estructura 1"
    
    If Estrnro1 <> 0 Then
        EmpEstrnro1 = Estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & ternro & " AND tenro = " & Tenro1
        StrSql = StrSql & " AND (htethasta is null )"
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------
If Tenro2 <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "Buscando el tipo de estructura 2"
    
    If Estrnro2 <> 0 Then
        EmpEstrnro2 = Estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & ternro & " AND tenro = " & Tenro2
        StrSql = StrSql & " AND (htethasta is null )"
               
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!estrnro
        End If
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------
If Tenro3 <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "Buscando el tipo de estructura 3"
    
    If Estrnro3 <> 0 Then
        EmpEstrnro3 = Estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & ternro & " AND tenro = " & Tenro3
        StrSql = StrSql & " AND (htethasta is null )"
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!estrnro
        End If
    End If
End If

FechaActual = CDate(FechaDesde)
Contador = DateDiff("d", FechaDesde, FechaHasta)
        
Do
   If tieneHoras(ternro, FechaActual) Then
       Flog.writeline Espacios(Tabulador * 1) & "El emplado tiene horas para la fecha " & FechaActual
       
       'Busco el teorico y las registraciones para la fecha
       Teorico = obtenerTeorico(FechaActual, ternro, TipoDiaActual)
       Registraciones = mostrarRegistracion(FechaActual, ternro)
       
       ' SI (l_tipodiaactual = 4) significa que no esta procesado el dia
       If CLng(TipoDia) = 3 Then 'Esta selectada la opcion todos
          Mostrar = True
       Else
          If TipoDiaActual <> 4 Then 'Si la fecha esta procesada
             If CLng(TipoDia) = 4 Then 'Si hay que mostrar los dias francos con registraciones
                Mostrar = ((CLng(TipoDiaActual) = 1) And (Trim(Registraciones) <> ""))
             Else
                Mostrar = (CLng(TipoDia) = CLng(TipoDiaActual))
             End If
           Else
                Mostrar = False
           End If
       End If
       
       If Mostrar Then
       
          'calculardia (factual)
          HorarioTeorico = Teorico
          Fichada = Registraciones
          MotivoInasistencias = buscarAnormalida(FechaActual, ternro)
          HorasAutorizadas = mostrarAutorizadas(FechaActual, ternro)
          
          '------------------------------------------------------------------
          'Busco los valores de los tipos de hora para el dia
          '------------------------------------------------------------------
          
          Flog.writeline Espacios(Tabulador * 1) & "Busco el acumulado diario para la fecha"
                        
          StrSql = " SELECT DISTINCT ternro, adfecha, thnro, adcanthoras "
          StrSql = StrSql & " FROM gti_acumdiario "
          StrSql = StrSql & " WHERE ternro = " & ternro & " "
          StrSql = StrSql & " AND EXISTS (SELECT * FROM confrep WHERE repnro = 62 AND "
          '09/12/2005 - Busco solo las horas indicadas en el filtro-------------------------
          If tipoHora <> "" Then
            StrSql = StrSql & " confval = thnro AND conftipo = 'TH'"
            StrSql = StrSql & " AND confval IN (" & tipoHora & ") )"
          Else
            StrSql = StrSql & " confval = thnro AND conftipo = 'TH')"
          End If
          '---------------------------------------------------------------------------------
          StrSql = StrSql & " AND adfecha = " & ConvFecha(FechaActual)
          OpenRecordset StrSql, rsConsult
          
          Call vaciar
          Do Until rsConsult.EOF
                Call setValorCol(rsConsult!thnro, CDbl(rsConsult!adcanthoras))
          
                rsConsult.MoveNext
          Loop
          
          rsConsult.Close
          
          StrSql = " INSERT INTO rep_ctrl_hor "
          StrSql = StrSql & " (bpronro , ternro, fecha_desde, fecha_hasta, horario_teorico, "
          StrSql = StrSql & " fichada, motivo_inasist, horas_autorizadas, fecha_actual, "
          StrSql = StrSql & " tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3 "
          
          cantCols = 0
          For I = 1 To 10
             If Trim(Columnas(I)) <> "" Then
                cantCols = cantCols + 1
                StrSql = StrSql & ",valor" & cantCols & ",columna" & cantCols
             End If
          Next
          
          StrSql = StrSql & ",cant_columna) VALUES "
          StrSql = StrSql & "(" & NroProceso
          StrSql = StrSql & "," & ternro
          StrSql = StrSql & "," & ConvFecha(FechaDesde)
          StrSql = StrSql & "," & ConvFecha(FechaHasta)
          'FGZ - 04/07/2007 - Faltaban limitar la longitud de los strings
          StrSql = StrSql & ",'" & Left(HorarioTeorico, 40) & "'"
          StrSql = StrSql & ",'" & Left(Fichada, 100) & "'"
          StrSql = StrSql & ",'" & Left(MotivoInasistencias, 100) & "'"
          StrSql = StrSql & ",'" & Left(HorasAutorizadas, 50) & "'"
          StrSql = StrSql & "," & ConvFecha(FechaActual)
          StrSql = StrSql & ",'" & Tenro1 & "'"
          StrSql = StrSql & ",'" & EmpEstrnro1 & "'"
          StrSql = StrSql & ",'" & Tenro2 & "'"
          StrSql = StrSql & ",'" & EmpEstrnro2 & "'"
          StrSql = StrSql & ",'" & Tenro3 & "'"
          StrSql = StrSql & ",'" & EmpEstrnro3 & "'"
          
          For I = 1 To 10
             If Trim(Columnas(I)) <> "" Then
                StrSql = StrSql & "," & numberForSQL(ValorColumna(I)) & ",'" & Left(ColumnasTitulos(I), 50) & "'"
             End If
          Next
          'FGZ - 04/07/2007 - Faltaban limitar la longitud de los strings
          
          StrSql = StrSql & "," & cantCols
          StrSql = StrSql & ")"
          
          Flog.writeline Espacios(Tabulador * 1) & "Guardo los datos en la tabla"
    
          objConn.Execute StrSql, , adExecuteNoRecords
          
       End If
       FechaActual = DateAdd("d", 1, FechaActual)
       Contador = Contador - 1
   Else
       Flog.writeline Espacios(Tabulador * 1) & "El emplado No tiene horas para la fecha " & FechaActual
       
       FechaActual = DateAdd("d", 1, FechaActual)
       Contador = Contador - 1
   End If
Loop While Contador >= 0
                    
Exit Sub

MError:
    Flog.writeline Espacios(Tabulador * 1) & "Error en el tercero " & ternro & " Error: " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "----------------------------------------------------------"
    
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub


Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
  End If

End Function

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

Function agregar(Lista, Dato)
Dim Salida

   Salida = Lista

   If Not existe(Lista, Dato) Then
     If (Lista = "") Then
       Salida = Dato
     Else
       Salida = Lista & "," & Dato
     End If
   End If
   
   agregar = Salida

End Function

Function existe(Lista, Dato)
Dim Salida
Dim arr
Dim I
Dim tmp

   Salida = False
   arr = Split(Lista, ",")

   For I = 0 To UBound(arr)

       tmp = Split(arr(I), "@")

       If (CStr(tmp(0)) = CStr(Dato)) Then
          Salida = True
       End If
   Next

   existe = Salida

End Function 'existe(lista,dato)

Function existe2(Lista, Dato, ByRef Accion)

Dim Salida
Dim arr
Dim I
Dim tmp

   Salida = False
   arr = Split(Lista, ",")
   Accion = ""

   For I = 0 To UBound(arr)

       tmp = Split(arr(I), "@")

       If (CStr(tmp(0)) = CStr(Dato)) Then
          Salida = True
          Accion = tmp(1)
       End If
   Next

   existe2 = Salida

End Function 'existe(lista,dato)

'-----------------------------------------------------------------------------------------------
'inicializa las columnas
Sub inicializar()
Dim I

   For I = 1 To 10
      Columnas(I) = ""
      ColumnasTitulos(I) = ""
      ValorColumna(I) = 0
   Next

End Sub 'inicializar

'-----------------------------------------------------------------------------------------------
'vaciar las columnas
Sub vaciar()
Dim I

   For I = 1 To 10
      ValorColumna(I) = 0
   Next

End Sub 'vaciar
'-------------------------------------------------------------------------------------------------------------------
Sub cargarConfRep(nroRep)

Dim objRs As New ADODB.Recordset

    StrSql = " SELECT confnrocol, conftipo, confaccion, confetiq, confval, confval2 "
    StrSql = StrSql & " FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & nroRep
    StrSql = StrSql & " AND conftipo = 'TH'"
    '09/12/2005 - Busco solo las horas indicadas en el filtro
    If tipoHora <> "" Then
        StrSql = StrSql & " AND confval IN (" & tipoHora & ")"
    End If


    OpenRecordset StrSql, objRs
    
    inicializar

    Do Until objRs.EOF

        Columnas(CLng(objRs!confnrocol)) = agregar(Columnas(CLng(objRs!confnrocol)), objRs!confval & "@" & objRs!confaccion)
        ColumnasTitulos(CLng(objRs!confnrocol)) = objRs!confetiq
        
        objRs.MoveNext
    Loop
    
    objRs.Close
    
End Sub 'cargarConfRep

Sub setValorCol(tipoHora, Valor)
Dim Accion As String
Dim I

  For I = 1 To 10
    If existe2(Columnas(I), tipoHora, Accion) Then
       
       If UCase(Accion) = UCase("sumar") Then
          ValorColumna(I) = ValorColumna(I) + Valor
       Else
          If UCase(Accion) = UCase("restar") Then
             ValorColumna(I) = ValorColumna(I) - Valor
          Else
             ValorColumna(I) = ValorColumna(I) + Valor
          End If
       End If
    End If
  Next
  
End Sub


Sub crearCabezera()

    StrSql = " INSERT INTO rep_ctrl_hor_cab "
    StrSql = StrSql & " (bpronro , fecha_desde, fecha_hasta,"
    StrSql = StrSql & " tenro1, tenro2, tenro3) VALUES "
    StrSql = StrSql & "(" & NroProceso
    StrSql = StrSql & "," & ConvFecha(FechaDesde)
    StrSql = StrSql & "," & ConvFecha(FechaHasta)
    StrSql = StrSql & ",'" & Tenro1 & "'"
    StrSql = StrSql & ",'" & Tenro2 & "'"
    StrSql = StrSql & ",'" & Tenro3 & "'"
    StrSql = StrSql & ")"
    
    Flog.writeline "Guardo los datos en la tabla cabezera"

    objConn.Execute StrSql, , adExecuteNoRecords

End Sub
