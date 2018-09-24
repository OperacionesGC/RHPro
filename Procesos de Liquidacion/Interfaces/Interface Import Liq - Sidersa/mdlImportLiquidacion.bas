Attribute VB_Name = "mdlImportLiquidacion"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "21/09/2006"
'Global Const UltimaModificacion = "Versión inicial"

'Global Const Version = "1.02"
'Global Const FechaModificacion = "12/10/2006"
'Global Const UltimaModificacion = "" ' FAF - Error en el insert de la cabecera de liquidacion

'Global Const Version = "1.03"
'Global Const FechaModificacion = "07/01/2008"
'Global Const UltimaModificacion = "" ' Lisandro Moro - Se agregaron las planillas 1, 2 y 3.

'Global Const Version = "1.04"
'Global Const FechaModificacion = "11/08/2008"
'Global Const UltimaModificacion = "" ' Lisandro Moro - Se modifico para que permita insertar valores en 0 (Pedido por el cliente).

Global Const Version = "1.05"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = "" ' MB - Encriptacion de string connection

'------------------------------------------------------------------------
Dim fs, f
Dim a As Integer
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global Tabulador As Long
Global TiempoInicialProceso
Global TiempoAcumulado

Global NroLinea
Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegFecha As Date
Global SeparadorDecimal As String
Global DescripcionModelo As String

Global NroModelo As Integer
Global pliqnro As Integer
Global pronro As Integer
Global tprocnro As Integer
Global tipoLiq As Integer
Global NombArchivo As String

'Conceptos
Const conliq1 = "1000"
Const conliq2 = "1010"
Const conliq3 = "1020"
Const conliq4 = "1030"
Const conliq5 = "1040"
Const conliq6 = "1050"
Const conliq7 = "1060"
Const conliq8 = "1070"
Const conliq9 = "1080"
Const conliq10 = "1090"
Const conliq11 = "1100"
Const conliq12 = "1110"
Const conliq13 = "1120"
Const conliq14 = "3000"
Const conliq15 = "7000"
Const conliq16 = "8000"
Const conliq17 = "9000"

'Acumuladores
Const aculiq1 = 1
Const aculiq2 = 2
Const aculiq3 = 3
Const aculiq4 = 4
Const aculiq5 = 5
Const aculiq6 = 6

Dim conacu(35, 3) As String
'0 = confnrocol
'1 = conftipo
'2 = confval
'3 = confaccion

'P1
'1   Sueldo básico
'2   Horas Extras
'3   Horas Nocturnas
'4   Horas Guardia
'5   Inasistencia
'6   Trabajo por Equipo
'7   Enfermedad / Accidente
'8   Licencias Especiales
'9   Antigüedad
'10  A Cuenta de futuros aumentos
'11  Comisión
'12  Ajuste Horas Nocturnas
'13  Ajuste Horas Extras
'
'P2
'14  Sueldo Resto Mes
'15  Adicional Variable
'16  Premio Asistencia
'17  Adicional Fijo
'18  Adicional x Título
'19  Adicional x Turno
'20  Guardia Objetivo
'21  Diferencia Reemplazos
'22  Valorizado KM Recorridos
'23  Descarga
'24  SubTotal (P1 + P2)
'
'P3
'25  Transporte (total planilla 2)
'26  Retención Judicial
'27  Retención Ganancias
'28  Resta Retenciones
'29  SubTotal (Transporte - Retenciones)
'30  Permanencia
'31  Acuerdo
'32  Viático
'33  No Remun.
'34  Neto


Private Sub Main()

Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim cantRegistros
Dim PID As String
Dim Parametros As String
Dim ArrParametros

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

    TiempoInicialProceso = GetTickCount
'    OpenConnection strconexion, objConn
'    OpenConnection strconexion, objconnProgreso
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ImportacionLiquidaciones" & "-" & NroProceso & ".log"
    
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
    
    Flog.writeline "Inicio Proceso Importación Liquidación: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
        'Busca la configuracion por confrep
        StrSql = " SELECT * FROM confrep "
        StrSql = StrSql & " WHERE repnro = 225 "
        OpenRecordset StrSql, objRs2
        If objRs2.EOF Then
            Flog.writeline " No esta configurado el ConfRep "
            Exit Sub
        End If
        
        'Inicializo el arreglo con los acumuladores
        
        For a = 0 To UBound(conacu) - 1
            conacu(a, 0) = ""
            conacu(a, 1) = ""
            conacu(a, 2) = ""
            conacu(a, 3) = ""
        Next a
        
        Flog.writeline "Obtengo los datos del confrep (225) "
        
        'Lleno el arreglo con los valores
        Do Until objRs2.EOF
            If IsNull(objRs2!confval2) Or objRs2!confval2 = "" Then
                If IsNull(objRs2!confval) Or objRs2!confval = "" Then
                    conacu(CInt(objRs2!confnrocol), 2) = ""
                Else
                    conacu(CInt(objRs2!confnrocol), 2) = CStr(objRs2!confval)
                End If
            Else
                If IsNull(objRs2!confval2) Or objRs2!confval2 = "" Then
                    conacu(CInt(objRs2!confnrocol), 2) = ""
                Else
                    conacu(CInt(objRs2!confnrocol), 2) = CStr(objRs2!confval2)
                End If
            End If
            conacu(CInt(objRs2!confnrocol), 0) = CStr(objRs2!confnrocol)
            conacu(CInt(objRs2!confnrocol), 1) = CStr(objRs2!conftipo)
            conacu(CInt(objRs2!confnrocol), 3) = CStr(objRs2!confaccion)
            
            objRs2.MoveNext
        Loop
        
        objRs2.Close
        Set objRs2 = Nothing
        
       'Obtengo los parametros del proceso
       Parametros = objRs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       'Nro Modelo
       NroModelo = ArrParametros(0)
       Flog.writeline "     Modelo --> " & NroModelo
       
       'Archivo
       NombArchivo = ArrParametros(1)
       Flog.writeline "     Archivo --> " & NombArchivo
       
       'Período
       pliqnro = ArrParametros(2)
       Flog.writeline "     Período --> " & pliqnro
       
       'Proceso
       pronro = ArrParametros(3)
       Flog.writeline "     Proceso --> " & pronro
       
       'Tipo Proceso
       tprocnro = ArrParametros(4)
       Flog.writeline "     Tipo Proceso --> " & tprocnro
       
       'Tipo de Liquidacion
       tipoLiq = ArrParametros(5)
       Flog.writeline "     Tipo Liquidación --> " & tipoLiq
       
       ' Proceso que genera los datos
       Call ComenzarTransferencia
       
    Else
       Exit Sub
    End If
    
    If objRs.State = adStateOpen Then objRs.Close
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
End Sub

Public Sub ComenzarTransferencia()
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim EncontroAlguno As Boolean

    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el registro de la tabla sistema nro 1"
        Exit Sub
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
        SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
        DescripcionModelo = objRs!moddesc
        
        Flog.writeline Espacios(Tabulador * 1) & "Modelo " & NroModelo & " - " & objRs!moddesc
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Directorio de importación :  " & Directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo
        Exit Sub
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Path = Directorio
        
    Dim fc, F1, s2
    Set Folder = fs.GetFolder(Directorio)
    Set CArchivos = Folder.Files
        
    HuboError = False
    EncontroAlguno = False
    For Each archivo In CArchivos
        EncontroAlguno = True
        If UCase(archivo.Name) = UCase(NombArchivo) Then
            NArchivo = archivo.Name
            Flog.writeline Espacios(Tabulador * 1) & "Procesando archivo " & archivo.Name
            Call LeeArchivo(Directorio & "\" & archivo.Name)
        End If
    Next
    
    If Not EncontroAlguno Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el archivo " & NombArchivo
    End If
    
End Sub

Private Sub LeeArchivo(ByVal NombreArchivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset


    If App.PrevInstance Then
        Flog.writeline Espacios(Tabulador * 0) & "Hay una instancia previa del proceso corriendo "
        Exit Sub
    End If
    'Espero hasta que se crea el archivo
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then
            Flog.writeline Espacios(Tabulador * 0) & "No anda el getfile "
            Err.Number = 1
        End If
    Loop
    On Error GoTo 0
    Flog.writeline Espacios(Tabulador * 0) & "Archivo creado " & NombreArchivo
   
   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProceso & "," & NroModelo & ",'" & Left(NombreArchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "inter_pin")
        Flog.writeline Espacios(Tabulador * 0) & "Ultimo inter_pin " & crpNro
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se pudo abrir el archivo " & NombreArchivo
    End If
                
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If rs_Modelo.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "No esta el modelo " & NroModelo
        Exit Sub
    End If
                
    StrSql = "SELECT * FROM modelo_filas WHERE bpronro =" & NroProceso
    StrSql = StrSql & " ORDER BY fila "
    OpenRecordset StrSql, rs_Lineas
    If Not rs_Lineas.EOF Then
        rs_Lineas.MoveFirst
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No hay filas seleccionadas"
    End If
    
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = rs_Lineas.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (99 / CEmpleadosAProc)
    
    If Not Validar_tipo() Then
        HuboError = True
        GoTo Fin
    End If

    Do While Not f.AtEndOfStream And Not rs_Lineas.EOF
        strLinea = f.ReadLine
        NroLinea = NroLinea + 1
        If Trim(strLinea) <> "" And NroLinea = rs_Lineas!fila Then
            Select Case tipoLiq
                Case 1:
                    Call Insertar_Linea_tipo1(strLinea)
                Case 2:
                    Call Insertar_Linea_tipo2(strLinea)
                Case 3:
                    Call Insertar_Linea_tipo3(strLinea)
                Case 4:
                    Call Insertar_Linea_tipo4(strLinea)
                Case 5:
                    Call Insertar_Linea_tipo5(strLinea)
                Case 6:
                    Call Insertar_Linea_tipo6(strLinea)
                Case Else
                    Flog.writeline Espacios(Tabulador * 0) & "Tipo de Liquidación no definida"
                    HuboError = True
                    GoTo Fin
            End Select
            
            RegLeidos = RegLeidos + 1
            
            rs_Lineas.MoveNext
            
            'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
            Progreso = Progreso + IncPorc
            Flog.writeline Espacios(Tabulador * 0) & "Progreso = " & CLng(Progreso) & " (Incremento = " & IncPorc & ")"
            
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
    Loop
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    
    f.Close
    'Borrar el archivo
    fs.Deletefile NombreArchivo, True
    
    Exit Sub
    
CE:
    HuboError = True
    
    MyRollbackTrans
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Private Function checkconc(concod As String) As Boolean
Dim objRs As New ADODB.Recordset

    StrSql = "SELECT * FROM concepto WHERE conccod = '" & concod & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No existe el concepto --> " & concod & ". El proceso abortara."
        checkconc = False
    Else
        checkconc = True
    End If

End Function
    

Private Function checkacu(ByVal acuNro As Integer) As Boolean
Dim objRs As New ADODB.Recordset

    StrSql = "SELECT * FROM acumulador WHERE acunro = '" & acuNro & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No existe el acumulador --> " & acuNro & ". El proceso abortara."
        checkacu = False
    Else
        checkacu = True
    End If

End Function
    
Private Sub insertConc(cabliq, conliq As String, val As Double)
Dim objRs As New ADODB.Recordset
Dim concnro As Integer

    'If val = 0 Then
    '    Exit Sub
    'End If
    
    StrSql = "SELECT * FROM concepto "
    StrSql = StrSql & " WHERE concepto.conccod = '" & conliq & "'"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        concnro = objRs!concnro
    Else
        Flog.writeline "       ****** No existe el concepto --> " & conliq & ". No se ingresara el valor asociado."
        Exit Sub
    End If
    objRs.Close
    
    StrSql = "SELECT * FROM detliq WHERE detliq.cliqnro = " & cabliq
    StrSql = StrSql & " AND detliq.concnro = " & concnro
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "INSERT INTO detliq (cliqnro,concnro,dlimonto,dlicant) VALUES (" & cabliq & _
                 "," & concnro & "," & val & ",0)"
        
    Else
        StrSql = "UPDATE detliq SET dlimonto = " & val & ",dlicant=0 " & _
                 " WHERE cliqnro = " & cabliq & " AND concnro = " & concnro
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub

Private Sub insertAcu(cabliq, ByVal aculiq As Integer, val As Double)
Dim objRs As New ADODB.Recordset

    'If val = 0 Then
    '    Exit Sub
    'End If
    
    StrSql = "SELECT * FROM acu_liq WHERE acu_liq.cliqnro = " & cabliq
    StrSql = StrSql & " AND acu_liq.acunro = " & aculiq
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "INSERT INTO acu_liq (cliqnro,acunro,almonto,alcant) VALUES (" & cabliq & _
                 "," & aculiq & "," & val & ",0)"
        
    Else
        StrSql = "UPDATE acu_liq SET almonto = " & val & ",alcant=0 " & _
                 " WHERE cliqnro = " & cabliq & " AND acunro = " & aculiq
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub

Private Sub insertConAcu(indice As Integer, cabliq, val As Double)

    'If val = 0 Then
    '    Exit Sub
    'End If
    
    If conacu(indice, 3) = "sumar" Then
    ElseIf conacu(indice, 3) = "restar" Then
        val = val * -1
    End If
    
    If conacu(indice, 1) = "CO" Then
        insertConc cabliq, conacu(indice, 2), val
    ElseIf conacu(indice, 1) = "AC" Then
        insertAcu cabliq, conacu(indice, 2), val
    Else
        'Si no es concepto ni acumulador no hace nada
    End If
    
End Sub


Private Function Validar_tipo() As Boolean
Dim val1 As Integer
Dim val2 As String
Dim val3 As Double
Dim val4 As Double
Dim val5 As Double
Dim val6 As Double
Dim val7 As Double
Dim val8 As Double
Dim val9 As Double
Dim val10 As Double
Dim val11 As Double
Dim val12 As Double
Dim val13 As Double
Dim val14 As Double
Dim val15 As Double
Dim val16 As Double
Dim val17 As Double
Dim val18 As Double
    
    ' CHEQUEA LA EXISTENCIA DE CONCEPTOS
    If Not checkconc(conliq1) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq2) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq3) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq4) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq5) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq6) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq7) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq8) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq9) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq10) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq11) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq12) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq13) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq14) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq15) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq16) Then
        GoTo Fin
    End If
    
    If Not checkconc(conliq17) Then
        GoTo Fin
    End If
    
    '/* CHEQUEA LA EXISTENCIA DE LOS ACUMULADORES */
    
    If Not checkacu(aculiq1) Then
        GoTo Fin
    End If
    
    If Not checkacu(aculiq2) Then
        GoTo Fin
    End If
    
    If Not checkacu(aculiq3) Then
        GoTo Fin
    End If
    
    If Not checkacu(aculiq4) Then
        GoTo Fin
    End If
    
    If Not checkacu(aculiq5) Then
        GoTo Fin
    End If
    
    If Not checkacu(aculiq6) Then
        GoTo Fin
    End If
    
    'Agregados por confrep
    For a = 1 To UBound(conacu) - 1
        If conacu(a, 1) = "CO" Then
            If Not checkconc(conacu(a, 2)) Then
                GoTo Fin
            End If
        ElseIf conacu(a, 1) = "AC" Then
            If Not checkacu(conacu(a, 2)) Then
                GoTo Fin
            End If
        Else
            GoTo Fin
        End If
    Next a
    
    
    Validar_tipo = True
    Exit Function
    
Fin:
    Validar_tipo = False
    Exit Function
    
End Function


Private Sub Insertar_Linea_tipo1(strReg As String)
'/* Proceso de importacion para Planilla - Q1 */
Dim val1 As Long
Dim val2 As String
Dim val3 As Double
Dim val4 As Double
Dim val5 As Double
Dim val6 As Double
Dim val7 As Double
Dim val8 As Double
Dim val9 As Double
Dim val10 As Double
Dim val11 As Double
Dim val12 As Double
Dim val13 As Double
Dim val14 As Double
Dim val15 As Double
Dim val16 As Double
Dim val17 As Double
Dim val18 As Double
Dim cliqnro As Long
Dim ternro As Long

Dim objRs As New ADODB.Recordset

    On Error GoTo MError
    
'    RegLeidos = RegLeidos + 1
    
    val1 = CLng(Mid(strReg, 1, 7))
    val2 = Mid(strReg, 8, 29)
    val3 = CDbl(Mid(strReg, 37, 10))
    val4 = CDbl(Mid(strReg, 47, 10))
    val5 = CDbl(Mid(strReg, 57, 9)) * (-1)
    val6 = CDbl(Mid(strReg, 66, 9))
    val7 = CDbl(Mid(strReg, 75, 9))
    val8 = CDbl(Mid(strReg, 84, 9))
    val9 = CDbl(Mid(strReg, 93, 11))
    val10 = CDbl(Mid(strReg, 104, 10))
    val11 = CDbl(Mid(strReg, 114, 9)) * (-1)
    val12 = CDbl(Mid(strReg, 123, 13))
        
    '    IMPORT DELIMITER " " val1             /* Legajo */
    '                         val2             /* Apellido y nombre */
    '                         val3             /* Bruto */
    '                         val4             /* Hs. Extras */
    '                         val5             /* Inasistencia */
    '                         val6             /* Antiguedad */
    '                         val7             /* Varios */
    '                         val8             /* Dif. horas */
    '                         val9             /* Adicional */
    '                         val10            /* Subtotal */
    '                         val11            /* Retenciones */
    '                         val12.           /* Neto */
    
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & val1
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & val1
        GoTo Fin
    Else
        ternro = objRs!ternro
    End If
    
    'verifico si existe o no la liquidacion
    StrSql = "SELECT * FROM cabliq where pronro = " & pronro
    StrSql = StrSql & " AND empleado = " & ternro
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "INSERT INTO cabliq (pronro, empleado) VALUES (" & pronro & "," & ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        cliqnro = getLastIdentity(objConn, "cabliq")
    Else
    
        cliqnro = objRs!cliqnro
        
    End If
    
    Call insertConc(cliqnro, conliq1, val3)
    Call insertConc(cliqnro, conliq2, val4)
    Call insertConc(cliqnro, conliq3, val5)
    Call insertConc(cliqnro, conliq4, val6)
    Call insertConc(cliqnro, conliq11, val7)
    Call insertConc(cliqnro, conliq12, val8)
    Call insertConc(cliqnro, conliq13, val9)
    Call insertConc(cliqnro, conliq16, val11)
    
    Call insertAcu(cliqnro, aculiq6, val12)
    
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub


Private Sub Insertar_Linea_tipo2(strReg As String)
'/* Proceso de importacion para Planilla 1 - Q2 */
Dim val1 As Long
Dim val2 As String
Dim val3 As Double
Dim val4 As Double
Dim val5 As Double
Dim val6 As Double
Dim val7 As Double
Dim val8 As Double
Dim val9 As Double
Dim val10 As Double
Dim val11 As Double
Dim val12 As Double
Dim val13 As Double
Dim val14 As Double
Dim val15 As Double
Dim val16 As Double
Dim val17 As Double
Dim val18 As Double
Dim cliqnro As Long
Dim ternro As Long

Dim objRs As New ADODB.Recordset

    On Error GoTo MError
    
'    RegLeidos = RegLeidos + 1
    
    val1 = CLng(Mid(strReg, 1, 4))
    val2 = Mid(strReg, 5, 29)
    val3 = CDbl(Mid(strReg, 34, 8))
    val4 = CDbl(Mid(strReg, 42, 8))
    val5 = CDbl(Mid(strReg, 50, 8)) * (-1)
    val6 = CDbl(Mid(strReg, 58, 8))
    val7 = CDbl(Mid(strReg, 66, 9)) * (-1)
    val8 = CDbl(Mid(strReg, 75, 9))
    val9 = CDbl(Mid(strReg, 84, 9))
    val10 = CDbl(Mid(strReg, 93, 11))
    val11 = CDbl(Mid(strReg, 104, 11))
    val12 = CDbl(Mid(strReg, 115, 11))
    val13 = CDbl(Mid(strReg, 126, 11))
    
'    IMPORT DELIMITER " " val1             /* Legajo */
'                         val2             /* Apellido y nombre */
'                         val3             /* Bruto */
'                         val4             /* Hs. Extras */
'                         val5             /* Inasistencia */
'                         val6             /* Antiguedad */
'                         val7             /* Inasistencia vacaciones */
'                         val8             /* Adic. Viaje - Choferes */
'                         val9             /* Guardia - Supervisores */
'                         val10            /* Asistencia - Presentismo */
'                         val11            /* Adic. Fijo - Personal de planta */
'                         val12            /* Dec. 392/03 */
'                         val13.           /* Subtotal */
    
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & val1
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & val1
        GoTo Fin
    Else
        ternro = objRs!ternro
    End If
    
    'verifico si existe o no la liquidacion
    StrSql = "SELECT * FROM cabliq where pronro = " & pronro
    StrSql = StrSql & " AND empleado = " & ternro
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "INSERT INTO cabliq (pronro, empleado) VALUES (" & pronro & "," & ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        cliqnro = getLastIdentity(objConn, "cabliq")
        
    Else
    
        cliqnro = objRs!cliqnro
        
    
    End If
    
    Call insertConc(cliqnro, conliq1, val3)
    Call insertConc(cliqnro, conliq2, val4)
    Call insertConc(cliqnro, conliq3, val5)
    Call insertConc(cliqnro, conliq4, val6)
    Call insertConc(cliqnro, conliq5, val7)
    Call insertConc(cliqnro, conliq6, val8)
    Call insertConc(cliqnro, conliq7, val9)
    Call insertConc(cliqnro, conliq8, val10)
    Call insertConc(cliqnro, conliq9, val11)
    Call insertConc(cliqnro, conliq10, val12)
    
'    /* Asignaci¢n de los valores de acumulados
'
'    RUN asign-acu(PER.cabliq.cliqnro, aculiq1, val13).
'    RUN asign-acu(PER.cabliq.cliqnro, aculiq2, val14).
'    RUN asign-acu(PER.cabliq.cliqnro, aculiq3, val15).
'    RUN asign-acu(PER.cabliq.cliqnro, aculiq4, val16).
'    RUN asign-acu(PER.cabliq.cliqnro, aculiq5, val17).
'    RUN asign-acu(PER.cabliq.cliqnro, aculiq6, val18).
'    */

'    Call insertAcu(cliqnro, aculiq6, val12)

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub


Private Sub Insertar_Linea_tipo3(strReg As String)
'/* Proceso de importacion para Planilla 2 - Q2 */
Dim val1 As Long
Dim val2 As String
Dim val3 As Double
Dim val4 As Double
Dim val5 As Double
Dim val6 As Double
Dim val7 As Double
Dim val8 As Double
Dim val9 As Double
Dim val10 As Double
Dim val11 As Double
Dim val12 As Double
Dim val13 As Double
Dim val14 As Double
Dim val15 As Double
Dim val16 As Double
Dim val17 As Double
Dim val18 As Double
Dim cliqnro As Long
Dim ternro As Long

Dim objRs As New ADODB.Recordset

    On Error GoTo MError
    
'    RegLeidos = RegLeidos + 1
    
    val1 = CLng(Mid(strReg, 1, 4))
    val2 = Mid(strReg, 5, 32)
    val3 = CDbl(Mid(strReg, 37, 13))
    val4 = CDbl(Mid(strReg, 50, 13))
    val5 = CDbl(Mid(strReg, 63, 13)) * (-1)
    val6 = CDbl(Mid(strReg, 76, 13))
    val7 = CDbl(Mid(strReg, 89, 13))
    val8 = CDbl(Mid(strReg, 102, 13))
    val9 = CDbl(Mid(strReg, 115, 13))
        
'    IMPORT DELIMITER " " val1             /* Legajo */
'                         val2             /* Apellido y nombre */
'                         val3             /* Transporte */
'                         val4             /* SAC */
'                         val5             /* Retenciones */
'                         val6             /* Subtotal */
'                         val7             /* No remunerativo */
'                         val8             /* Asignaciones */
'                         val9.            /* Neto */
    
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & val1
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & val1
        GoTo Fin
    Else
        ternro = objRs!ternro
    End If
    
    'verifico si existe o no la liquidacion
    StrSql = "SELECT * FROM cabliq where pronro = " & pronro
    StrSql = StrSql & " AND empleado = " & ternro
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "INSERT INTO cabliq (pronro, empleado) VALUES (" & pronro & "," & ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        cliqnro = getLastIdentity(objConn, "cabliq")
        
    Else
    
        cliqnro = objRs!cliqnro
        
    End If
    
    Call insertConc(cliqnro, conliq14, val4)
    Call insertConc(cliqnro, conliq16, val5)
    Call insertConc(cliqnro, conliq17, val7)
    Call insertConc(cliqnro, conliq15, val8)
    
    Call insertAcu(cliqnro, aculiq6, val9)
    
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub

Private Sub Insertar_Linea_tipo4(strReg As String)
'/* Proceso de importacion para Planilla - P1 */

    Dim val1 As Long
    Dim val2 As String
    Dim val3 As Double
    Dim val4 As Double
    Dim val5 As Double
    Dim val6 As Double
    Dim val7 As Double
    Dim val8 As Double
    Dim val9 As Double
    Dim val10 As Double
    Dim val11 As Double
    Dim val12 As Double
    Dim val13 As Double
    Dim val14 As Double
    Dim val15 As Double
    
    Dim cliqnro As Long
    Dim ternro As Long
    
    Dim objRs As New ADODB.Recordset

    On Error GoTo MError
    
    'RegLeidos = RegLeidos + 1
    
    val1 = CLng(Mid(strReg, 1, 4))
    val2 = CStr(Mid(strReg, 5, 20))
    val3 = CDbl(Mid(strReg, 25, 10))
    val4 = CDbl(Mid(strReg, 35, 7))
    val5 = CDbl(Mid(strReg, 42, 7)) '* (-1)
    val6 = CDbl(Mid(strReg, 49, 7))
    val7 = CDbl(Mid(strReg, 56, 7))
    val8 = CDbl(Mid(strReg, 63, 7))
    val9 = CDbl(Mid(strReg, 70, 9))
    val10 = CDbl(Mid(strReg, 79, 9))
    val11 = CDbl(Mid(strReg, 88, 7)) '* (-1)
    val12 = CDbl(Mid(strReg, 95, 7))
    val13 = CDbl(Mid(strReg, 102, 11))
    val14 = CDbl(Mid(strReg, 113, 7))
    val15 = CDbl(Mid(strReg, 120, 7))
        
    '4   1   4   Cod.Legajo                     Numerico
    '20  5   24  Nombre                         Texto
    '10  25  34  Sueldo básico                  Decimal (con 2 decimales)   Suma
    '7   35  41  Horas Extras                   Decimal (con 2 decimales)   Suma
    '7   42  48  Horas Nocturnas                Decimal (con 2 decimales)   Suma
    '7   49  55  Horas Guardia                  Decimal (con 2 decimales)   Suma
    '7   56  62  Inasistencia                   Decimal (con 2 decimales)   Resta
    '7   63  69  Trabajo por Equipo             Decimal (con 2 decimales)   Suma
    '9   70  78  Enfermedad / Accidente         Decimal (con 2 decimales)   Suma
    '9   79  87  Licencias Especiales           Decimal (con 2 decimales)   Suma
    '7   88  94  Antigüedad                     Decimal (con 2 decimales)   Suma
    '7   95  101 A Cuenta de futuros aumentos   Decimal (con 2 decimales)   Suma
    '11  102 112 Comisión                       Decimal (con 2 decimales)   Suma
    '7   113 119 Ajuste Horas Nocturnas         Decimal (con 2 decimales)   Suma
    '7   120 126 Ajuste Horas Extras            Decimal (con 2 decimales)   Suma
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & val1
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & val1
        GoTo Fin
    Else
        ternro = objRs!ternro
    End If
    
    'verifico si existe o no la liquidacion
    StrSql = "SELECT * FROM cabliq where pronro = " & pronro
    StrSql = StrSql & " AND empleado = " & ternro
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "INSERT INTO cabliq (pronro, empleado) VALUES (" & pronro & "," & ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        cliqnro = getLastIdentity(objConn, "cabliq")
    Else
    
        cliqnro = objRs!cliqnro
        
    End If
    
    Call insertConAcu(1, cliqnro, val3)
    Call insertConAcu(2, cliqnro, val4)
    Call insertConAcu(3, cliqnro, val5)
    Call insertConAcu(4, cliqnro, val6)
    Call insertConAcu(5, cliqnro, val7)
    Call insertConAcu(6, cliqnro, val8)
    Call insertConAcu(7, cliqnro, val9)
    Call insertConAcu(8, cliqnro, val10)
    Call insertConAcu(9, cliqnro, val11)
    Call insertConAcu(10, cliqnro, val12)
    Call insertConAcu(11, cliqnro, val13)
    Call insertConAcu(12, cliqnro, val14)
    Call insertConAcu(13, cliqnro, val15)
    
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub

Private Sub Insertar_Linea_tipo5(strReg As String)
'/* Proceso de importacion para Planilla - P2 */

    Dim val1 As Long
    Dim val2 As String
    Dim val3 As Double
    Dim val4 As Double
    Dim val5 As Double
    Dim val6 As Double
    Dim val7 As Double
    Dim val8 As Double
    Dim val9 As Double
    Dim val10 As Double
    Dim val11 As Double
    Dim val12 As Double
    Dim val13 As Double
    
    Dim cliqnro As Long
    Dim ternro As Long
    
    Dim objRs As New ADODB.Recordset

    On Error GoTo MError
    
    'RegLeidos = RegLeidos + 1
    
    val1 = CLng(Mid(strReg, 1, 4))
    val2 = CStr(Mid(strReg, 6, 25))
    val3 = CDbl(Mid(strReg, 31, 12))
    val4 = CDbl(Mid(strReg, 43, 9))
    val5 = CDbl(Mid(strReg, 52, 9))
    val6 = CDbl(Mid(strReg, 61, 9))
    val7 = CDbl(Mid(strReg, 70, 9))
    val8 = CDbl(Mid(strReg, 79, 9))
    val9 = CDbl(Mid(strReg, 88, 9))
    val10 = CDbl(Mid(strReg, 97, 9))
    val11 = CDbl(Mid(strReg, 106, 9))
    val12 = CDbl(Mid(strReg, 115, 9))
    val13 = CDbl(Mid(strReg, 124, 13))
        
    '4   1   4   Cod.Legajo                 Numerico
    '25  6   30  Nombre                     Texto
    '12  31  42  Sueldo Resto Mes           Decimal (con 2 decimales)   Suma
    '9   43  51  Adicional Variable         Decimal (con 2 decimales)   Suma
    '9   52  60  Premio Asistencia          Decimal (con 2 decimales)   Suma
    '9   61  69  Adicional Fijo             Decimal (con 2 decimales)   Suma
    '9   70  78  Adicional x Título         Decimal (con 2 decimales)   Suma
    '9   79  87  Adicional x Turno          Decimal (con 2 decimales)   Suma
    '9   88  96  Guardia Objetivo           Decimal (con 2 decimales)   Suma
    '9   97  105 Diferencia Reemplazos      Decimal (con 2 decimales)   Suma
    '9   106 114 Valorizado KM Recorridos   Decimal (con 2 decimales)   Suma
    '9   115 123 Descarga                   Decimal (con 2 decimales)   Suma
    '13  124 136 SUBTOTAL (P1+P2)           Decimal (con 2 decimales)
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & val1
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & val1
        GoTo Fin
    Else
        ternro = objRs!ternro
    End If
    
    'verifico si existe o no la liquidacion
    StrSql = "SELECT * FROM cabliq where pronro = " & pronro
    StrSql = StrSql & " AND empleado = " & ternro
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "INSERT INTO cabliq (pronro, empleado) VALUES (" & pronro & "," & ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        cliqnro = getLastIdentity(objConn, "cabliq")
    Else
    
        cliqnro = objRs!cliqnro
        
    End If
    
    'Call insertConc(cliqnro, conliq1, val3)
    'Call insertConc(cliqnro, conliq2, val4)
    'Call insertConc(cliqnro, conliq3, val5)
    'Call insertConc(cliqnro, conliq4, val6)
    'Call insertConc(cliqnro, conliq11, val7)
    'Call insertConc(cliqnro, conliq12, val8)
    'Call insertConc(cliqnro, conliq13, val9)
    'Call insertConc(cliqnro, conliq16, val11)
    
    'Call insertAcu(cliqnro, aculiq6, val12)
    
    Call insertConAcu(14, cliqnro, val3)
    Call insertConAcu(15, cliqnro, val4)
    Call insertConAcu(16, cliqnro, val5)
    Call insertConAcu(17, cliqnro, val6)
    Call insertConAcu(18, cliqnro, val7)
    Call insertConAcu(19, cliqnro, val8)
    Call insertConAcu(20, cliqnro, val9)
    Call insertConAcu(21, cliqnro, val10)
    Call insertConAcu(22, cliqnro, val11)
    Call insertConAcu(23, cliqnro, val12)
    Call insertConAcu(24, cliqnro, val13)
    
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub

Private Sub Insertar_Linea_tipo6(strReg As String)
'/* Proceso de importacion para Planilla - P3 */

    Dim val1 As Long
    Dim val2 As String
    Dim val3 As Double
    Dim val4 As Double
    Dim val5 As Double
    Dim val6 As Double
    Dim val7 As Double
    Dim val8 As Double
    Dim val9 As Double
    Dim val10 As Double
    Dim val11 As Double
    Dim val12 As Double

    Dim cliqnro As Long
    Dim ternro As Long
    
    Dim objRs As New ADODB.Recordset

    On Error GoTo MError
    
    'RegLeidos = RegLeidos + 1
    
    val1 = CLng(Mid(strReg, 1, 4))
    val2 = CStr(Mid(strReg, 6, 20))
    val3 = CDbl(Mid(strReg, 26, 10))
    val4 = CDbl(Mid(strReg, 36, 11))
    val5 = CDbl(Mid(strReg, 47, 9))
    val6 = CDbl(Mid(strReg, 56, 11))
    val7 = CDbl(Mid(strReg, 67, 13))
    val8 = CDbl(Mid(strReg, 80, 11))
    val9 = CDbl(Mid(strReg, 91, 11))
    val10 = CDbl(Mid(strReg, 102, 11))
    val11 = CDbl(Mid(strReg, 113, 11))
    val12 = CDbl(Mid(strReg, 124, 13))
    
    ' 4   1   4 Cod.Legajo                             Numerico
    '20   6  25 Nombre                                 Texto
    '10  26  35 Transporte (total planilla 2)          Decimal (con 2 decimales)
    '11  36  46 Retención Judicial                     Decimal (con 2 decimales)   Resta
    ' 9  47  55 Retención Ganancias                    Decimal (con 2 decimales)   Resta
    '11  56  66 Resta Retenciones                      Decimal (con 2 decimales)   Resta
    '13  67  79 SubTotal (Transporte - Retenciones)    Decimal (con 2 decimales)
    '11  80  90 Permanencia Decimal (con 2 decimales)  Suma
    '11  91 101 Acuerdo     Decimal (con 2 decimales)  Suma
    '11 102 112 Viático Decimal (con 2 decimales)      Suma
    '11 113 123 No Remun.   Decimal (con 2 decimales)  Suma
    '13 124 136 Neto                                   Decimal (con 2 decimales)
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & val1
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & val1
        GoTo Fin
    Else
        ternro = objRs!ternro
    End If
    
    'verifico si existe o no la liquidacion
    StrSql = "SELECT * FROM cabliq where pronro = " & pronro
    StrSql = StrSql & " AND empleado = " & ternro
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "INSERT INTO cabliq (pronro, empleado) VALUES (" & pronro & "," & ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        cliqnro = getLastIdentity(objConn, "cabliq")
    Else
        cliqnro = objRs!cliqnro
    End If
    
    Call insertConAcu(25, cliqnro, val3)
    Call insertConAcu(26, cliqnro, val4)
    Call insertConAcu(27, cliqnro, val5)
    Call insertConAcu(28, cliqnro, val6)
    Call insertConAcu(29, cliqnro, val7)
    Call insertConAcu(30, cliqnro, val8)
    Call insertConAcu(31, cliqnro, val9)
    Call insertConAcu(32, cliqnro, val10)
    Call insertConAcu(33, cliqnro, val11)
    Call insertConAcu(34, cliqnro, val12)
    
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub

