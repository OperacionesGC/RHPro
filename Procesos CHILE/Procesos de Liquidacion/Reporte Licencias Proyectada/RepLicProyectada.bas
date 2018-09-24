Attribute VB_Name = "RepLicProyectada"

Option Explicit


Global Const Version = "1.00" ' Leticia Amadio
Global Const FechaModificacion = "04-10-2011"
'Global Const FechaVersión = "04-10-2011"
Global Const UltimaModificacion = ""   ' reporte de Licencias Proyectada


' ________________________________________________________
Global EmpErrores As Boolean
Global IncPorc As Double
Global Progreso


' Parametros
Global repnro
Global licfecha As Date


' confrep
Global licencias As String
Global licproyectada As String
Global liccantdias As Integer
Global licestado As String


Dim rs_batchpr As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset





' _______________________________________________________
' procedimiento ppal que genera el reporte
' _______________________________________________________

Public Sub Main()
Dim objconnMain As New ADODB.Connection

Dim rsEmpl As New ADODB.Recordset

Dim strCmdLine
Dim Nombre_Arch As String
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

Dim btprcnro

Dim cantRegistros


    
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

    
    
    Nombre_Arch = PathFLog & "RepLicProyectada" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    
    'Abro la conexion
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
    
        
    On Error GoTo ME_Main
 
 
 
 
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    
    Flog.writeline "-------------------------------------------------"
    Flog.writeline " Version     : " & Version
    Flog.writeline " Modificacion: " & FechaModificacion
    Flog.writeline " PID         : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline
        
    cantRegistros = 0
     'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProcesoBatch
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        cantRegistros = CInt(objRs!total)
       ' totalEmpleados = cantRegistros
    End If
    objRs.Close
    
    'seteo de las variables de progreso
    IncPorc = (99 / (cantRegistros))   ' o IncPorc = (99 / (cantRegistros))
    Progreso = 0
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej='" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej=" & ConvFecha(Date) & ", bprcprogreso=0, bprcestado='Procesando', bprcpid=" & PID & " WHERE bpronro=" & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    
    'Obtengo los datos del proceso
    btprcnro = 313
    
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro=" & btprcnro & " AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batchpr
    
    
    TiempoInicialProceso = GetTickCount
    
    
    If Not rs_batchpr.EOF Then
        
        bprcparam = rs_batchpr!bprcparam
        
        cargarParametros bprcparam
        
        datosReporte repnro
            
        borrarRepAnt
                
        insertarCabecera NroProcesoBatch, licfecha
        
        
        StrSql = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProcesoBatch
        OpenRecordset StrSql, rsEmpl
     
        Do While Not rsEmpl.EOF
        
            EmpErrores = False
           
            Flog.writeline "       "
            Flog.writeline " Chequeo de Licencias - Empleado(ternro) " & rsEmpl!ternro
            
            
            reporteLicProy rsEmpl!ternro
           

            'Actualizo el estado del proceso
            TiempoAcumulado = GetTickCount
            cantRegistros = cantRegistros - 1
            
            Progreso = CDbl(Progreso) + IncPorc
            Progreso = Fix(Progreso)
           
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProcesoBatch ' NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
          
            'Si se generaron todos los datos del empleado correctamente lo borro
            If Not EmpErrores Then
                StrSql = " DELETE FROM batch_empleado "
                StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
                StrSql = StrSql & " AND ternro = " & rsEmpl!ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If

            
       rsEmpl.MoveNext
       Loop
       rsEmpl.Close
    
    Else
        Flog.writeline "no encontró el proceso"
    End If
            
    rs_batchpr.Close
    Set rs_batchpr = Nothing
    
    
    
     
    TiempoFinalProceso = GetTickCount
          
          
    'Flog.writeline
    'Flog.writeline "**********************************************************"
    'Flog.writeline
    'Flog.writeline "Cantidad de Empleados Insertados: " & CantEmplSinError
    'Flog.writeline "Cantidad de Empleados Con ERRORES: " & CantEmplError
    'Flog.writeline
    Flog.writeline
    Flog.writeline
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline
    
    
    If Not HuboError Then 'If Not Errores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100  WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub



ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & " SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords

End Sub




' __________________________________________________________________
' ________________________________________________________________
Sub cargarParametros(ByVal Parametros As String)
Dim arrParam

Flog.writeline "Lista de Parametros:  "
Flog.writeline Espacios(Tabulador * 1) & Parametros
Flog.writeline


If Not IsNull(Parametros) Then
    
    If Len(Parametros) >= 1 Then
        
        arrParam = Split(Parametros, "@")
        
        repnro = arrParam(0)
        Flog.writeline Espacios(Tabulador * 1) & "ConfRep=" & repnro
        
        licfecha = arrParam(1)
        Flog.writeline Espacios(Tabulador * 1) & "Fecha para chequeo de Licencia=" & licfecha
        
    End If
    
Else
    Flog.writeline Espacios(Tabulador * 1) & "Parametros Nulos "
End If

Flog.writeline
Flog.writeline

End Sub



' ________________________________________________________
' borrar reporte anterior - x ahora no se guarda historico
' ________________________________________________________
Sub borrarRepAnt()

On Error GoTo borrarep

    
    StrSql = " DELETE FROM rep_licproyectada_det WHERE bpronro <>" & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords

    
    StrSql = " DELETE FROM rep_licproyectada WHERE bpronro <>" & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords


Exit Sub

borrarep:
    MensajeError StrSql
    HuboError = True


End Sub


' ________________________________________________________
' borrar reporte anterior - x ahora no se guarda historico
' ________________________________________________________
Sub borrarLicencia(ByVal ternro)

On Error GoTo borrarLic

    
    StrSql = " DELETE FROM rep_licproyectada_det "
    StrSql = StrSql & " WHERE bpronro=" & NroProcesoBatch
    StrSql = StrSql & "     AND ternro=" & ternro
    objConn.Execute StrSql, , adExecuteNoRecords

    
    'StrSql = " DELETE FROM rep_licproyectada "
    'StrSql = StrSql & " WHERE bpronro=" & NroProcesoBatch
    'StrSql = StrSql & "     AND ternro=" & ternro
    'objConn.Execute StrSql, , adExecuteNoRecords


Exit Sub

borrarLic:
    MensajeError StrSql
    HuboError = True


End Sub

' Usamos la funcion DAteSerial para obtener el primero y el ultimo dia
' Primer = DateSerial(Year(Fecha), Month(Fecha) + 0, 1)
' Ultimo = DateSerial(Year(Fecha), Month(Fecha) + 1, 0)

' __________________________________________________________________
' __________________________________________________________________
Public Sub reporteLicProy(ByVal ternro As Long) ' bpronro

Dim Desde As Date
Dim Hasta As Date

Dim fecdesde As Date
Dim fechasta As Date

Dim totalDias, dias As Integer



On Error GoTo CE

      
    Desde = DateSerial(Year(licfecha), Month(licfecha) - 1, 1) ' primer dia del mes anterior
    Hasta = DateSerial(Year(licfecha), Month(licfecha) + 1, 0) ' último dia del mes actual



    StrSql = "SELECT empleg, empleado.terape, empleado.ternom, emp_lic.emp_licnro, emp_lic.tdnro, "
    StrSql = StrSql & " emp_lic.elfechadesde, emp_lic.elfechahasta, emp_lic.eltipo, licestnro "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN emp_lic ON empleado.ternro = emp_lic.empleado "
    StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro "
    StrSql = StrSql & " WHERE ternro=" & ternro
    StrSql = StrSql & "     AND tipdia.tdnro IN (" & licencias & ") "  ' l_tipodia
    StrSql = StrSql & "     AND emp_lic.licestnro = 2 "   ' licencias aprobadas
    StrSql = StrSql & " AND ( "
    StrSql = StrSql & " ( ( elfechadesde >= " & ConvFecha(Desde) & " ) AND ( elfechadesde <= " & ConvFecha(Hasta) & " )  ) "
    StrSql = StrSql & " OR "
    StrSql = StrSql & " ( ( elfechadesde <= " & ConvFecha(Desde) & " ) AND ( elfechahasta >= " & ConvFecha(Desde) & " ) ) "
    StrSql = StrSql & " ) "
    OpenRecordset StrSql, rs
   
    If rs.EOF Then
        ' Flog.writeline "no hay empleados"
        ' CEmpleadosAProc = 1
    Else
    
    
        dias = 0
        totalDias = 0
        fecdesde = Desde
        fechasta = Hasta
        
        
        Flog.writeline "    " & rs!empleg & " - " & rs!terape & " " & rs!ternom
        

        Do While Not rs.EOF

            If rs!elfechadesde > Desde Then
                calcularDias rs!elfechadesde, rs!elfechahasta, dias
            Else
                calcularDias Desde, rs!elfechahasta, dias
            End If
           
            If rs!elfechahasta > fecdesde Then
                fecdesde = rs!elfechahasta
            End If
               
            totalDias = totalDias + dias
           
            insertarDetLicencia ternro, rs!emp_licnro, rs!tdnro, rs!elfechadesde, rs!elfechahasta, rs!eltipo, rs!licestnro
            
            
        rs.MoveNext
        Loop
        
   
    End If 're_empleados.eof
    rs.Close
    

    
    fecdesde = DateSerial(Year(fecdesde), Month(fecdesde), Day(fecdesde) + 1)
    
        
    If totalDias >= liccantdias Then
        
        If fecdesde <= fechasta Then
            Flog.writeline "    - Se realiza proyección de Licencia Médica. "
            insertarDetLicencia ternro, 0, licproyectada, fecdesde, fechasta, 1, licestado
        Else
            Flog.writeline "    - No se realiza proyección de Licencia Médica."
            borrarLicencia ternro
        End If
    
    Else
    
        Flog.writeline "    - No se realiza proyección de Licencia Médica (cantidad de días de licencia menor a lo especificado)."
        borrarLicencia ternro
    
    End If
    
    
    

If rs.State = adStateOpen Then rs.Close
If rs1.State = adStateOpen Then rs1.Close

Set rs = Nothing
Set rs1 = Nothing



Exit Sub

CE:
    MensajeError StrSql

    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = " UPDATE batch_proceso SET "
    StrSql = StrSql & " bprcprogreso=" & Progreso & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
    StrSql = StrSql & " WHERE bpronro=" & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    HuboError = True
    EmpErrores = True
    Flog.writeline " Error: " & Err.Description


End Sub



' ______________________________________________
' ______________________________________________
Sub calcularDias(fecdesde, fechasta, dias)

dias = 0

dias = DateDiff("d", fecdesde, fechasta + 1)  ' fecha hasta +1


End Sub




' _______________________________________________________________________________
' _______________________________________________________________________________
Sub insertarCabecera(ByVal NroProcesoBatch, ByVal licfecha As Date)

On Error GoTo errorCab

    StrSql = " INSERT INTO rep_licproyectada(bpronro, licfecha)"
    StrSql = StrSql & " VALUES (" & NroProcesoBatch & "," & ConvFecha(licfecha) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


Exit Sub

errorCab:
    MensajeError StrSql
    HuboError = True

End Sub



' _______________________________________________________________________________
' _______________________________________________________________________________
Sub insertarDetLicencia(ternro, licnro, tdnro, fecdesde, fechasta, eltipo, licestnro)

On Error GoTo errorDet


    StrSql = " INSERT INTO rep_licproyectada_det(bpronro, ternro, licnro, tdnro, elfecdesde,elfechasta,eltipo, licestnro) "
    StrSql = StrSql & " VALUES (" & NroProcesoBatch & "," & ternro & "," & licnro & "," & tdnro & "," & ConvFecha(fecdesde) & "," & ConvFecha(fechasta) & "," & eltipo & ", " & licestnro & ")"
    objConn.Execute StrSql, , adExecuteNoRecords


Exit Sub

errorDet:
    MensajeError StrSql
    HuboError = True
    
End Sub





' ______________________________________________________
'
' estado: aprobada=2 (x default), pendiente=1, rechazada=3, autorizado x supervisor=4
' ______________________________________________________
Sub datosReporte(ByRef repnro)
Dim I

On Error GoTo CR


licencias = 0
licproyectada = 0
licestado = 2   ' aprobada
liccantdias = 0

            
'Configuracion del Reporte
StrSql = "SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=" & repnro
StrSql = StrSql & " ORDER BY confrep.confnrocol "
OpenRecordset StrSql, rs1

If rs1.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte " & repnro
    Exit Sub
Else
    
    Do While Not rs1.EOF
            
        Select Case rs1("conftipo")
            
            Case "LIC"
                        
               If Not IsNull(rs1("confval2")) Then
                    'confrep(I).valfa = rs1("confval2")
                    
                    If rs1("confval2") = "-1" Then
                        licproyectada = rs1("confval")
                    Else
                        licencias = licencias & "," & rs1("confval")
                    End If
               Else
                    licencias = licencias & "," & rs1("confval")
               End If
                
        
            Case "VAL"
                liccantdias = rs1("confval")
            
            Case "EST"
                licestado = rs1("confval")
            
        End Select
         
            
        
    rs1.MoveNext
    Loop
    
End If
rs1.Close


Flog.writeline "Configuración del Reporte:"
Flog.writeline "    Licencias:" & licencias
Flog.writeline "    Licencia Proyectada:" & licproyectada
Flog.writeline "    Cantidad de días de la licencia:" & liccantdias
Flog.writeline "    Estado de la Lic. Proyectada:" & licestado



Exit Sub

CR:
    MensajeError StrSql
    
    HuboError = True

End Sub




' _________________________________________________________
' _________________________________________________________
Sub MensajeError(StrSql)

    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="

End Sub


