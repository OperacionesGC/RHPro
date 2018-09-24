Attribute VB_Name = "MdlGenCuotasPrestamos"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaVersion = "21/10/2009"
'Global Const UltimaModificacion = "Encriptacion de string connection"
'Global Const UltimaModificacion1 = "Manuel Lopez"

'Global Const Version = "1.02"
'Global Const FechaVersion = "16/02/2010"
'Global Const UltimaModificacion = "Se agregó el método *Interes_Pre_Sac*, y las correspondientes llamadas al mismo dentro del método *Generar_Cuotas*"
'Global Const UltimaModificacion1 = "Domingo Pacheco"

'Global Const Version = "1.03"
'Global Const FechaVersion = "26/03/2010"
'Global Const UltimaModificacion = "Se corrigió el formato de la fecha de vencimiento de cuota, en el método *Interes_pre_Sac*"
'Global Const UltimaModificacion1 = "Domingo Pacheco"

'Global Const Version = "1.04"
'Global Const FechaVersion = "23/11/2011"
'Global Const UltimaModificacion = "Se modificó para que cuando Sistema = 3 genere las cuotas por SAC."
'Global Const UltimaModificacion1 = "Manterola Maria Magdalena"

'Global Const Version = "1.05"
'Global Const FechaVersion = "05/11/2013"
'Global Const UltimaModificacion = "se corrigio el case para el sistema en el que entra para calcular las cuotas, se modifico el calculo de las cuotas en sistema simple, se quito redondeo a sistema frances"
'Global Const UltimaModificacion1 = "Fernandez, Matias"

'Global Const Version = "1.06"
'Global Const FechaVersion = "30/06/2015"
'Global Const UltimaModificacion = " Se modifica para el sistema Interes Simple y Frances, el valor que se guarda de interes, (se quita *100), y se modifica el valor del capital"
'Global Const UltimaModificacion1 = " CAS-29961 - VISION - Error en busqueda de prestamos"
'Global Const UltimaModificacion2 = " Borrelli Facundo"

'Global Const Version = "1.07"
'Global Const FechaVersion = "20/01/2016"
'Global Const UltimaModificacion = " Se agrega un mensaje de log para los 3 sistemas para cuando la cantidad de cuotas del prestamo es cero"
'Global Const UltimaModificacion1 = " CAS-34507 - MONASTERIO BASE CORVEN - Bug en generar cuotas masivamente"
'Global Const UltimaModificacion2 = " Borrelli Facundo"

'Global Const Version = "1.08"
'Global Const FechaVersion = "14/03/2016"
'Global Const UltimaModificacion = " Se cambian los tipos de datos para variables de los procedimientos generar_cuotas, interes_simple, e interes_Frances"
'Global Const UltimaModificacion1 = " CAS-34507 - MONASTERIO BASE CORVEN - Bug en generar cuotas masivamente - [Entrega 2]"
'Global Const UltimaModificacion2 = " Borrelli Facundo"

'Global Const Version = "1.09"
'Global Const FechaVersion = "21/03/2016"
'Global Const UltimaModificacion = " Se cambian los tipos de datos de las variables de Single a Double"
'Global Const UltimaModificacion1 = " CAS-34507 - MONASTERIO BASE CORVEN - Bug en generar cuotas masivamente - [Entrega 3]"
'Global Const UltimaModificacion2 = " Borrelli Facundo"

Global Const Version = "1.10"
Global Const FechaVersion = "28/03/2016"
Global Const UltimaModificacion = " Se cambian los tipos de datos de las variables a Double Para el Sistema Frances"
Global Const UltimaModificacion1 = " CAS-34507 - MONASTERIO BASE CORVEN - Bug en generar cuotas masivamente - [Entrega 4]"
Global Const UltimaModificacion2 = " Borrelli Facundo"

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial.
' Autor      : FGZ
' Fecha      : 09/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

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
    
    Nombre_Arch = PathFLog & "Prestamos-Cuotas" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Ultima Modificacion      : " & UltimaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 57 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Generar_Cuotas(NroProcesoBatch, bprcparam)
    End If
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    objConn.Close
    objconnProgreso.Close
    Flog.Close

End Sub


Public Sub Generar_Cuotas(ByVal NroProceso As Long, ByVal Parametros As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Proceso de generacion de cuotas de Prestamos.
' Autor      : FGZ
' Fecha      : 09/12/2004
' Ultima Mod.:
' Descripcion: 14/03/2016 - Borrelli Facundo - Se cambiaron los tipos de las variables Monto y deuda de Single a Long
' Descripcion: 21/03/2016 - Borrelli Facundo - Se cambiaron los tipos de las variables a Double
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String
Dim Cantidad As Long
Dim FechaDesde As Date
Dim FechaHasta As Date

Dim PreNro As Long
Dim DiaVenc As Long
Dim Monto As Double
Dim deuda As Double
Dim cuotas As Integer
Dim PreTNA As Double 'Single 'Interes
Dim PreIVA As Double 'Single
Dim PreOtrosGas As Double 'Single
Dim Quincenal As Boolean
Dim PreMes As Integer
Dim PreAnio As Integer
Dim Sistema As Long
Dim Descripcion As String
Dim PreQuin As Integer
'Dim Dueda As Double No se utiliza

Dim CantCuoSal As Integer
Dim SaldoCanc As Double 'Single
Dim Mes As Integer
Dim Anio As Integer

Dim Re_Generar As Boolean   'si se regeneran las cuotas que aun no esten generadas y no caceladas
Dim Todos As Boolean        'Todos los empleados

Dim rs_Prestamos As New ADODB.Recordset
Dim rs_Cuotas As New ADODB.Recordset

' Parametros


' Levanto cada parametro por separado, el separador de parametros es "."

Separador = "@"
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        'si se re generan las cuotas
        pos1 = 1
        pos2 = InStr(pos1, Parametros, Separador) - 1
        Re_Generar = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        'todos los empleados
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        Todos = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
    End If
End If

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'Comienzo la transaccion
MyBeginTrans
Flog.writeline Espacios(Tabulador * 0) & "Comienza transaccion"

On Error GoTo Manejador:

'-------------------------------------------------
TiempoInicialProceso = GetTickCount

If Todos Then
    StrSql = "SELECT prestamo.prenro, predesc, quincenal, prediavto, estnro "
    StrSql = StrSql & " ,precantcuo, premes, preanio, preimp, prediavto, preiva "
    StrSql = StrSql & " , pretna, preotrosgas, pre_formula.fcintnro, pre_formula.fcintprgcuot "
    StrSql = StrSql & " FROM  prestamo "
    StrSql = StrSql & " INNER JOIN pre_linea ON pre_linea.lnprenro = prestamo.lnprenro "
    StrSql = StrSql & " INNER JOIN pre_formula ON pre_linea.fcintnro = pre_formula.fcintnro "
    StrSql = StrSql & " WHERE prestamo.estnro = 3"
Else
    StrSql = "SELECT prestamo.prenro, predesc, quincenal, prediavto, estnro "
    StrSql = StrSql & " ,precantcuo, premes, preanio, preimp, prediavto, preiva "
    StrSql = StrSql & " , pretna, preotrosgas, pre_formula.fcintnro, pre_formula.fcintprgcuot"
    StrSql = StrSql & " FROM  prestamo "
    StrSql = StrSql & " INNER JOIN batch_empleado ON prestamo.ternro = batch_empleado.ternro "
    StrSql = StrSql & " INNER JOIN pre_linea ON pre_linea.lnprenro = prestamo.lnprenro "
    StrSql = StrSql & " INNER JOIN pre_formula ON pre_linea.fcintnro = pre_formula.fcintnro "
    StrSql = StrSql & " WHERE prestamo.estnro = 3"
    StrSql = StrSql & " AND batch_empleado.bpronro = " & NroProcesoBatch
End If
If rs_Prestamos.State = adStateOpen Then rs_Prestamos.Close
OpenRecordset StrSql, rs_Prestamos

If rs_Prestamos.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No hay ningun Prestamo"
End If
Cantidad = rs_Prestamos.RecordCount
If Cantidad = 0 Then
    Cantidad = 1
End If
IncPorc = (100 / Cantidad)
Progreso = 0
    
    
Do Until rs_Prestamos.EOF
'Inicializo
CantCuoSal = 0
SaldoCanc = 0
deuda = 0
Mes = 0
Anio = 0

    PreNro = rs_Prestamos("prenro")
    DiaVenc = rs_Prestamos("prediavto")
    Monto = rs_Prestamos("preimp")
    cuotas = IIf(Not EsNulo(rs_Prestamos("precantcuo")), rs_Prestamos("precantcuo"), 0)
    PreTNA = rs_Prestamos("pretna")
    PreIVA = rs_Prestamos("preiva")
    PreOtrosGas = IIf(Not EsNulo(rs_Prestamos("preotrosgas")), rs_Prestamos("preotrosgas"), 0)
    Quincenal = IIf(Not EsNulo(rs_Prestamos("quincenal")), CBool(rs_Prestamos("quincenal")), False)
    PreMes = rs_Prestamos("premes")
    PreAnio = rs_Prestamos("preanio")
    Sistema = rs_Prestamos("fcintnro")
    Descripcion = rs_Prestamos("predesc")
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Prestamo (" & PreNro & ") " & Descripcion
    
    If Not Re_Generar Then
        StrSql = " SELECT cuonro "
        StrSql = StrSql & " FROM  pre_cuota "
        StrSql = StrSql & " WHERE  pre_cuota.prenro = " & PreNro
        OpenRecordset StrSql, rs_Cuotas
    
        If rs_Cuotas.EOF Then
            
            Flog.writeline Espacios(Tabulador * 2) & "Generacion de cuotas... "
            
            Select Case Sistema
                Case 2: 'Sistema Frances
                    Call Sistema_Frances(PreNro, PreTNA, PreIVA, PreOtrosGas, Monto, cuotas, Quincenal, PreAnio, PreMes, DiaVenc)
                Case 1: 'Interes Simple
                    Call Sistema_InteresSimple(PreNro, PreTNA, PreIVA, PreOtrosGas, Monto, cuotas, Quincenal, PreAnio, PreMes, DiaVenc)
                Case 3:
                    Call Interes_Pre_Sac(PreNro, Monto, cuotas, PreTNA, PreIVA, PreOtrosGas, Quincenal, DiaVenc, PreMes, PreAnio)
                    Flog.writeline Espacios(Tabulador * 2) & "***SAC*** "
            Case Else
                Call Sistema_InteresSimple(PreNro, PreTNA, PreIVA, PreOtrosGas, Monto, cuotas, Quincenal, PreAnio, PreMes, DiaVenc)
            End Select
        
        Else
            Flog.writeline Espacios(Tabulador * 2) & "El Prestamo tiene cuotas generadas y no se regeneraran"
        End If
        If rs_Cuotas.State = adStateOpen Then rs_Cuotas.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "ReGeneracion de cuotas... "
                
        StrSql = " SELECT cuoimp, cuocapital, cuocancela, cuoano, cuomes, cuoquin, cuonro "
        StrSql = StrSql & " FROM  pre_cuota "
        StrSql = StrSql & " WHERE  pre_cuota.prenro = " & PreNro
        OpenRecordset StrSql, rs_Cuotas
        
        If Not Quincenal Then
            Do Until rs_Cuotas.EOF
                If CBool(rs_Cuotas("cuocancela")) Then
                    CantCuoSal = CantCuoSal + 1
                    SaldoCanc = SaldoCanc + CSng(rs_Cuotas("cuoimp"))
                    If Mes = 12 Then
                        Anio = Anio + 1
                        Mes = 1
                    Else
                        Mes = Mes + 1
                        Anio = Anio
                    End If
                Else
                    StrSql = " DELETE pre_cuota "
                    StrSql = StrSql & " WHERE  pre_cuota.cuonro = " & rs_Cuotas("cuonro")
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
        
               rs_Cuotas.MoveNext
            Loop
        Else
            Do Until rs_Cuotas.EOF
                If CBool(rs_Cuotas("cuocancela")) Then
                    CantCuoSal = CantCuoSal + 1
                    SaldoCanc = SaldoCanc + CSng(rs_Cuotas("cuoimp"))
                    If PreQuin = 2 Then
                        PreQuin = 1
                        If Mes = 12 Then
                            Anio = Anio + 1
                            Mes = 1
                        Else
                            Mes = Mes + 1
                            Anio = Anio
                        End If
                    Else
                        PreQuin = 2
                    End If
                Else
                    StrSql = " DELETE pre_cuota "
                    StrSql = StrSql & " WHERE  pre_cuota.cuonro = " & rs_Cuotas("cuonro")
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
    
               rs_Cuotas.MoveNext
            Loop
        End If
        If rs_Cuotas.State = adStateOpen Then rs_Cuotas.Close
        
        cuotas = cuotas - CantCuoSal
        'Deuda = Deuda - SaldoCanc
        deuda = Monto - SaldoCanc
        
        Select Case Sistema
            Case 2: 'Sistema Frances
                Call Sistema_Frances(PreNro, PreTNA, PreIVA, PreOtrosGas, deuda, cuotas, Quincenal, PreAnio, PreMes, DiaVenc)
            Case 1: 'Interes Simple
                Call Sistema_InteresSimple(PreNro, PreTNA, PreIVA, PreOtrosGas, deuda, cuotas, Quincenal, PreAnio, PreMes, DiaVenc)
            Case 3:
                Call Interes_Pre_Sac(PreNro, Monto, cuotas, PreTNA, PreIVA, PreOtrosGas, Quincenal, DiaVenc, PreMes, PreAnio)
                Flog.writeline Espacios(Tabulador * 2) & "***SAC2*** "
        Case Else
            Call Sistema_InteresSimple(PreNro, PreTNA, PreIVA, PreOtrosGas, deuda, cuotas, Quincenal, PreAnio, PreMes, DiaVenc)
        End Select
                
    End If
    
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Prestamos.MoveNext
Loop

'-------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Commit transaccion"
MyCommitTrans

Fin:
If rs_Cuotas.State = adStateOpen Then rs_Cuotas.Close
If rs_Prestamos.State = adStateOpen Then rs_Prestamos.Close

Set rs_Cuotas = Nothing
Set rs_Prestamos = Nothing
Exit Sub

Manejador:
    MyRollbackTrans
    Flog.writeline Espacios(Tabulador * 0) & "Rollback transaccion"
    
    HuboError = True
   
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error: " & Err.Description
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    GoTo Fin
End Sub


Private Sub Sistema_Frances(ByVal PreNro As Long, ByVal interes As Double, ByVal iva As Double, ByVal otroGasto As Double, ByVal deuda As Double, ByVal cuotas As Integer, ByVal Quincenal As Boolean, ByVal Anio As Integer, ByVal Mes As Integer, ByVal diaVto As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta cuota con calculo en sistema frances.
' Autor      : FGZ
' Fecha      : 09/12/2004
' Ultima Mod.:
' Descripcion: 14/03/2016 - Borrelli Facundo - Se cambiaron los tipos de las variables Vcuota, Vcuotatot y capital de Single a Long
' Descripcion: 21/03/2016 - Borrelli Facundo - Se cambiaron los tipos de las variables a Double
' Descripcion: 28/03/2016 - Borrelli Facundo - Se cambiaron los tipos de las variables que faltaban a Double
' ---------------------------------------------------------------------------------------------
Flog.writeline "Ingresa por  sistema frances"
Dim i

Dim IntCuo As Double 'Single
Dim Vcuota As Double
Dim Vcuotatot As Double
Dim Fechavto As Date
Dim Cquin As Integer
Dim capital As Double

Dim rs_Prestamo As New ADODB.Recordset

StrSql = "SELECT estnro, preimp, precantcuo,quincenal,prequin"
StrSql = StrSql & " FROM prestamo "
StrSql = StrSql & " WHERE prestamo.prenro =" & PreNro
OpenRecordset StrSql, rs_Prestamo
Flog.writeline "Busca prestamo frances:" & StrSql
If Not rs_Prestamo.EOF Then
    Cquin = rs_Prestamo("prequin")
    If cuotas <> 0 Then
        If Not Quincenal Then
            interes = CDbl(Round((interes / 12) / 100, 9))
        Else
            interes = CDbl(Round((interes / 24) / 100, 9))
        End If
        
        If interes = 0 Then
            Vcuota = deuda / cuotas
        Else
            Vcuota = (CLng(deuda) * CDbl(interes)) / (1 - ((1 + (CDbl(interes))) ^ ((-1) * CSng(cuotas))))
        End If
        Vcuotatot = Vcuota + iva + otroGasto
        
        i = 1
        Do While CInt(i) <= CInt(cuotas)
            IntCuo = deuda * interes
            'capital = Vcuota - IntCuo
            'FB - 30/06/2015 - Se modifica el valor del Capital para Interes Frances: Valor de la cuota.
            capital = Vcuota
            'Hasta aca
            deuda = deuda - capital
            Fechavto = CDate(diaVto & "/" & Mes & "/" & Anio)
        
            StrSql = "INSERT INTO pre_cuota "
            StrSql = StrSql & "(prenro,cuoimp,cuocancela, cuonrocuo,cuogastos,cuoiva,cuototal"
            StrSql = StrSql & ",cuocapital,cuointeres,cuosaldo,cuomes,cuoano,cuofecvto, cuoquin) "
            StrSql = StrSql & "VALUES ("
            StrSql = StrSql & PreNro
            StrSql = StrSql & "," & Round(Vcuota, 2)
            StrSql = StrSql & "," & "0"
            StrSql = StrSql & "," & CInt(i)
            StrSql = StrSql & "," & Round(otroGasto, 2)
            StrSql = StrSql & "," & Round(iva, 2)
            'StrSql = StrSql & "," & Round(Vcuotatot, 2)
            StrSql = StrSql & "," & Vcuotatot
            StrSql = StrSql & "," & Round(capital, 2)
            StrSql = StrSql & "," & Round(IntCuo, 2)
            StrSql = StrSql & "," & Round(deuda, 2)
            StrSql = StrSql & "," & Mes
            StrSql = StrSql & "," & Anio
            StrSql = StrSql & "," & ConvFecha(Fechavto)
            StrSql = StrSql & "," & Cquin
            StrSql = StrSql & ")"
            
            Flog.writeline "Inserto datos:" & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        
            If Not Quincenal Then
                If Mes = 12 Then
                    Mes = 1
                    Anio = Anio + 1
                Else
                    Mes = Mes + 1
                End If
            Else
                If Cquin = 2 Then
                    Cquin = 1
                    diaVto = diaVto - 15
                    If Mes = 12 Then
                        Mes = 1
                        Anio = Anio + 1
                    Else
                        Mes = Mes + 1
                    End If
                Else
                    Cquin = 2
                    diaVto = diaVto + 15
                End If
            End If
            i = i + 1
        Loop
    Else 'cuota <> 0
        Flog.writeline Espacios(Tabulador * 1) & "El prestamo: (" & PreNro & ") no posee configurada la cantidad de cuotas, o la cantidad de cuotas es 0(cero), revisar el prestamo "
    End If
End If
Flog.writeline "sale por  sistema frances"
'FB - 20/01/2016 - Se comenta para que siga procesando si encuentra un error
'If rs_Prestamo.State = adStateOpen Then rs_Prestamo.Close

End Sub

Private Sub Sistema_InteresSimple(ByVal PreNro As Long, ByVal interes As Double, ByVal iva As Double, ByVal otroGasto As Double, ByVal deuda As Double, ByVal cuotas As Integer, ByVal Quincenal As Boolean, ByVal Anio As Integer, ByVal Mes As Integer, ByVal diaVto As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inserta cuota con calculo en sistema de Interes Simple.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion: 14/03/2016 - Borrelli Facundo - Se cambiaron los tipos de las variables capital, Valor_cuota, de Single a Long
' Descripcion: 21/03/2016 - Borrelli Facundo - Se cambiaron los tipos de las variables a Double
' ---------------------------------------------------------------------------------------------
Flog.writeline "Ingresa por el sistema simple"
Dim i

Dim IntCuo As Single
'Dim Vcuota As Single
'Dim Vcuotatot As Single
Dim Fechavto As Date
Dim Cquin As Integer
Dim capital As Double

Dim Valor_Cuota As Double
Dim Monto_a_Cancelar As Double
Dim Saldo_a_Pagar As Double
Dim cambio As Double
Dim Cuovtotal As Double

Dim rs_Prestamo As New ADODB.Recordset

If cuotas <> 0 Then
    Flog.writeline "Hay cuotas"
    StrSql = "SELECT estnro, preimp, precantcuo,quincenal,prequin"
    StrSql = StrSql & " FROM prestamo "
    StrSql = StrSql & " WHERE prestamo.prenro =" & PreNro
    Flog.writeline "prestamo:" & StrSql
    OpenRecordset StrSql, rs_Prestamo
    If Not rs_Prestamo.EOF Then
        Flog.writeline "Hay prestamo"
        If Not Quincenal Then
            Valor_Cuota = Round(deuda / cuotas, 2)
            Monto_a_Cancelar = deuda
            Saldo_a_Pagar = deuda
            
            i = 1
            Do While CInt(i) <= CInt(cuotas) - 1
                
                cambio = cambio + Round(Valor_Cuota, 2)
                IntCuo = (((1 + (interes / 12 * rs_Prestamo("precantcuo") / 100)) * Monto_a_Cancelar) - (Monto_a_Cancelar)) / rs_Prestamo("precantcuo")
                
                'Cuovtotal = (Valor_Cuota + (Valor_Cuota * IntCuo)) + iva + otroGasto
                Cuovtotal = (Valor_Cuota + IntCuo) + iva + otroGasto 'mdf
                'capital = Saldo_a_Pagar - Valor_Cuota
                'FB - 30/06/2015 - Se modifica el valor del Capital para Interes Simple: capital = Valor de la cuota.
                capital = Valor_Cuota
                'Hasta aca
                
                Saldo_a_Pagar = Saldo_a_Pagar - Valor_Cuota
                Fechavto = CDate(diaVto & "/" & Mes & "/" & Anio)
                
                StrSql = "INSERT INTO pre_cuota "
                StrSql = StrSql & "(prenro,cuoimp,cuocancela, cuonrocuo,cuogastos,cuoiva,cuototal,cuocapital,cuointeres,cuosaldo,cuomes,cuoano,cuofecvto) "
                StrSql = StrSql & "VALUES ("
                StrSql = StrSql & PreNro
                StrSql = StrSql & "," & Round(Valor_Cuota, 2)
                StrSql = StrSql & "," & "0"
                StrSql = StrSql & "," & CInt(i)
                StrSql = StrSql & "," & Round(otroGasto, 2)
                StrSql = StrSql & "," & Round(iva, 2)
                'StrSql = StrSql & "," & Round(Cuovtotal, 2)
                StrSql = StrSql & "," & Cuovtotal
                StrSql = StrSql & "," & Round(capital, 2)
                'StrSql = StrSql & "," & Round(IntCuo, 2) * 100 'FB - 30/06/2015 - Se quita el * 100 para mostrar correctamente el interes
                StrSql = StrSql & "," & Round(IntCuo, 2)
                StrSql = StrSql & "," & Round(Saldo_a_Pagar, 2)
                StrSql = StrSql & "," & Mes
                StrSql = StrSql & "," & Anio
                StrSql = StrSql & "," & ConvFecha(Fechavto)
                StrSql = StrSql & ")"
                Flog.writeline "Inserto datos si es Mensual:" & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
                
                If Mes = 12 Then
                    Mes = 1
                    Anio = Anio + 1
                Else
                    Mes = Mes + 1
                End If
    
               i = i + 1
            Loop
    
            Valor_Cuota = Valor_Cuota + (Monto_a_Cancelar - (cambio + Round(Valor_Cuota, 2)))
            IntCuo = (((1 + (interes / 12 * rs_Prestamo("precantcuo") / 100)) * Monto_a_Cancelar) - (Monto_a_Cancelar)) / rs_Prestamo("precantcuo")
            'Cuovtotal = (Valor_Cuota + (Valor_Cuota * IntCuo)) + iva + otroGasto
            Cuovtotal = (Valor_Cuota + IntCuo) + iva + otroGasto 'mdf
            'capital = Saldo_a_Pagar - Valor_Cuota
            'FB - 30/06/2015 - Se modifica el valor del Capital para Interes Simple: capital = Valor de la cuota.
            capital = Valor_Cuota
            'Hasta aca
            Saldo_a_Pagar = Saldo_a_Pagar - Valor_Cuota
            Fechavto = CDate(diaVto & "/" & Mes & "/" & Anio)
    
            StrSql = "INSERT INTO pre_cuota "
            StrSql = StrSql & "(prenro,cuoimp,cuocancela, cuonrocuo,cuogastos,cuoiva,cuototal,cuocapital,cuointeres,cuosaldo,cuomes,cuoano,cuofecvto) "
            StrSql = StrSql & "VALUES ("
            StrSql = StrSql & PreNro
            StrSql = StrSql & "," & Round(Valor_Cuota, 2)
            StrSql = StrSql & "," & "0"
            StrSql = StrSql & "," & CInt(i)
            StrSql = StrSql & "," & Round(otroGasto, 2)
            StrSql = StrSql & "," & Round(iva, 2)
            StrSql = StrSql & "," & Round(Cuovtotal, 2)
            StrSql = StrSql & "," & Round(capital, 2)
            'StrSql = StrSql & "," & Round(IntCuo, 2) * 100 'FB - 30/06/2015 - Se quita el * 100 para mostrar correctamente el interes
            StrSql = StrSql & "," & Round(IntCuo, 2)
            StrSql = StrSql & "," & Round(Saldo_a_Pagar, 2)
            StrSql = StrSql & "," & Mes
            StrSql = StrSql & "," & Anio
            StrSql = StrSql & "," & ConvFecha(Fechavto)
            StrSql = StrSql & ")"
            Flog.writeline "Inserto datos Mensual:" & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'QUINCENAL
            Cquin = rs_Prestamo("prequin")
            Valor_Cuota = Round(deuda / cuotas, 2)
            Monto_a_Cancelar = deuda
            Saldo_a_Pagar = deuda
    
            i = 1
            Do While CInt(i) <= CInt(cuotas) - 1
                cambio = cambio + Round(Valor_Cuota, 2)
                IntCuo = (((1 + (interes / 12 * rs_Prestamo("precantcuo") / 100)) * Monto_a_Cancelar) - (Monto_a_Cancelar)) / rs_Prestamo("precantcuo") / 2
                'Cuovtotal = (Valor_Cuota + (Valor_Cuota * IntCuo)) + iva + otroGasto
                Cuovtotal = (Valor_Cuota + IntCuo) + iva + otroGasto 'mdf
                'capital = Saldo_a_Pagar - Valor_Cuota
                'FB - 30/06/2015 - Se modifica el valor del Capital para Interes Simple: capital = Valor de la cuota.
                capital = Valor_Cuota
                'Hasta aca
                Saldo_a_Pagar = Saldo_a_Pagar - Valor_Cuota
                Fechavto = CDate(diaVto & "/" & Mes & "/" & Anio)
                
                StrSql = "INSERT INTO pre_cuota "
                StrSql = StrSql & "(prenro,cuoimp,cuocancela, cuonrocuo,cuogastos,cuoiva,cuototal,cuocapital,cuointeres,cuosaldo,cuomes,cuoano,cuofecvto, cuoquin) "
                StrSql = StrSql & "VALUES ("
                StrSql = StrSql & PreNro
                StrSql = StrSql & "," & Round(Valor_Cuota, 2)
                StrSql = StrSql & "," & "0"
                StrSql = StrSql & "," & CInt(i)
                StrSql = StrSql & "," & Round(otroGasto, 2)
                StrSql = StrSql & "," & Round(iva, 2)
                StrSql = StrSql & "," & Round(Cuovtotal, 2)
                StrSql = StrSql & "," & Round(capital, 2)
                'StrSql = StrSql & "," & Round(IntCuo, 2) * 100 'FB - 30/06/2015 - Se quita el * 100 para mostrar correctamente el interes
                StrSql = StrSql & "," & Round(IntCuo, 2)
                StrSql = StrSql & "," & Round(Saldo_a_Pagar, 2)
                StrSql = StrSql & "," & Mes
                StrSql = StrSql & "," & Anio
                StrSql = StrSql & "," & ConvFecha(Fechavto)
                StrSql = StrSql & "," & Cquin
                StrSql = StrSql & ")"
                Flog.writeline "Inserto datos para cuando es Quincenal:" & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
                
                If Cquin = 2 Then
                   Cquin = 1
                   diaVto = diaVto - 15
                   If Mes = 12 Then
                      Mes = 1
                      Anio = Anio + 1
                      Else: Mes = Mes + 1
                   End If
                   Else: Cquin = 2
                        diaVto = diaVto + 15
                End If
                i = i + 1
            Loop
    
            Valor_Cuota = Valor_Cuota + (Monto_a_Cancelar - (cambio + Round(Valor_Cuota, 2)))
            IntCuo = (((1 + (interes / 12 * rs_Prestamo("precantcuo") / 100)) * Monto_a_Cancelar) - (Monto_a_Cancelar)) / rs_Prestamo("precantcuo") / 2
            'Cuovtotal = (Valor_Cuota + ((Valor_Cuota * IntCuo * 100) / 100)) + iva + otroGasto
            Cuovtotal = (Valor_Cuota + IntCuo) + iva + otroGasto 'mdf
            'capital = Saldo_a_Pagar - Valor_Cuota
            'FB - 30/06/2015 - Se modifica el valor del Capital para Interes Simple: capital = Valor de la cuota.
            capital = Valor_Cuota
            'Hasta aca
            Saldo_a_Pagar = Saldo_a_Pagar - Valor_Cuota
            Fechavto = CDate(diaVto & "/" & Mes & "/" & Anio)
    
            StrSql = "INSERT INTO pre_cuota "
            StrSql = StrSql & "(prenro,cuoimp,cuocancela, cuonrocuo,cuogastos,cuoiva,cuototal,cuocapital,cuointeres,cuosaldo,cuomes,cuoano,cuofecvto, cuoquin) "
            StrSql = StrSql & "VALUES ("
            StrSql = StrSql & PreNro
            StrSql = StrSql & "," & Round(Valor_Cuota, 2)
            StrSql = StrSql & "," & "0"
            StrSql = StrSql & "," & CInt(i)
            StrSql = StrSql & "," & Round(otroGasto, 2)
            StrSql = StrSql & "," & Round(iva, 2)
            StrSql = StrSql & "," & Round(Cuovtotal, 2)
            StrSql = StrSql & "," & Round(capital, 2)
            'StrSql = StrSql & "," & Round(IntCuo, 2) * 100 'FB - 30/06/2015 - Se quita el * 100 para mostrar correctamente el interes
            StrSql = StrSql & "," & Round(IntCuo, 2)
            StrSql = StrSql & "," & Round(Saldo_a_Pagar, 2)
            StrSql = StrSql & "," & Mes
            StrSql = StrSql & "," & Anio
            StrSql = StrSql & "," & ConvFecha(Fechavto)
            StrSql = StrSql & "," & Cquin
            StrSql = StrSql & ")"
            Flog.writeline "Inserto datos Quincenal:" & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
Else 'cuota <> 0
    Flog.writeline Espacios(Tabulador * 1) & "El prestamo: (" & PreNro & ") no posee configurada la cantidad de cuotas, o la cantidad de cuotas es 0(cero), revisar el prestamo "
End If
Flog.writeline "Sale por el sistema simple"
'FB - 20/01/2016 - Se comenta para que siga procesando si encuentra un error
'If rs_Prestamo.State = adStateOpen Then rs_Prestamo.Close
End Sub

Private Sub Interes_Pre_Sac(ByVal l_prenro As Integer, _
                                        ByVal deuda As Double, _
                                        ByVal cuotas As Double, _
                                        ByVal interes As Double, _
                                        ByVal iva As Double, _
                                        ByVal otroGasto As Double, _
                                        ByVal Quincenal As Integer, _
                                        ByVal diaVto As Integer, _
                                        ByVal Mes As Integer, _
                                        ByVal Anio As Integer)
Flog.writeline "Ingresa por Pre_sac - SAC2"
Dim Valor_Cuota As Double
Dim Monto_a_Cancelar As Double
Dim Saldo_a_Pagar As Double
Dim IntCuo As Double
Dim Cuovtotal As Double
Dim l_sql As String
Dim l_rs As New ADODB.Recordset
Dim i As Integer
Dim cambio As Double
Dim capital As Double
Dim l_diavto As Long
Dim Fechavto As String
Dim Cquin As Long

                                        
If Quincenal <> "-1" Then
   Quincenal = 0
End If
                                        
If cuotas <> 0 Then
    
    l_sql = "SELECT * "
    l_sql = l_sql & " FROM prestamo "
    l_sql = l_sql & " WHERE prestamo.prenro =" & l_prenro
    
    OpenRecordset l_sql, l_rs
    
    If Not l_rs.EOF Then
    
        deuda = CDbl(l_rs("preimp"))
        cuotas = CInt(l_rs("precantcuo"))
        interes = CDbl(l_rs("pretna"))
        iva = CDbl(l_rs("preiva"))
        otroGasto = CDbl(l_rs("preotrosgas"))
        Quincenal = CInt(l_rs("quincenal"))
        diaVto = l_rs("prediavto")
        Mes = CInt(l_rs("premes"))
        Anio = CInt(l_rs("preanio"))
        

        If l_rs("quincenal") = 0 Then
          
            Valor_Cuota = Round(deuda / cuotas, 2)
            Monto_a_Cancelar = deuda
            Saldo_a_Pagar = deuda
            
            i = 1

            If Mes > 6 Then
            Mes = 12
            Else: Mes = 6
            End If
            
            Do While CInt(i) <= CInt(cuotas) - 1
               
               cambio = cambio + Round(Valor_Cuota, 2)
               IntCuo = (((1 + (interes / 12 * CDbl(l_rs("precantcuo")) / 100)) * Monto_a_Cancelar) - (Monto_a_Cancelar)) / CDbl(l_rs("precantcuo"))
               Cuovtotal = (Valor_Cuota + (Valor_Cuota * IntCuo)) + iva + otroGasto
               
               'capital = Saldo_a_Pagar - Valor_Cuota
               'FB - 30/06/2015 - Se modifica el valor del Capital para Interes Pre_Sac: capital = Valor de la cuota.
               capital = Valor_Cuota
               'Hasta aca
               Saldo_a_Pagar = Saldo_a_Pagar - Valor_Cuota
               l_diavto = diaVto
               
               If Mes = 1 Or Mes = 3 Or Mes = 5 Or Mes = 7 Or Mes = 8 Or Mes = 10 Or Mes = 12 Then
                  If diaVto > 31 Then
                     l_diavto = 31
                  End If
               ElseIf Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
                      If diaVto > 30 Then
                         l_diavto = 30
                      End If
               ElseIf Mes = 2 Then
                      'If anioBisiesto(Anio) Then
                       '  If diaVto > 29 Then
                        '    l_diavto = 29
                         'End If
                      'ElseIf diaVto > 28 Then
                             l_diavto = 28
                      'End If
               End If
               
               Fechavto = "'" & CStr(l_diavto) & "/" & CStr(Mes) & "/" & CStr(Anio) & "'"
                
                
                l_sql = "INSERT INTO pre_cuota "
                l_sql = l_sql & "(prenro,cuoimp,cuocancela, cuonrocuo,cuogastos,cuoiva,cuototal,cuocapital,cuointeres,cuosaldo,cuomes,cuoano,cuofecvto) "
                l_sql = l_sql & "VALUES (" & l_prenro & "," & Valor_Cuota & "," & 0 & "," & CInt(i) & ","
                'l_sql = l_sql & otroGasto & "," & iva & "," & Cuovtotal & "," & capital & "," & IntCuo * 100 & ","
                'FB - 30/06/2015 - Se quita el * 100 para mostrar correctamente el interes
                l_sql = l_sql & otroGasto & "," & iva & "," & Cuovtotal & "," & capital & "," & IntCuo & ","
                l_sql = l_sql & Saldo_a_Pagar & "," & Mes & "," & Anio & "," & Fechavto & ")"
                
                Flog.writeline "Inserto datos Mensual:" & StrSql
                objConn.Execute l_sql, , adExecuteNoRecords
                
                If Mes = 12 Then
                   Mes = 6
                   Anio = Anio + 1
                   Else: Mes = Mes + 6
                End If
                
               i = i + 1
                
            Loop
            
            Valor_Cuota = Valor_Cuota + (Monto_a_Cancelar - (cambio + Round(Valor_Cuota, 2)))
            IntCuo = (((1 + (interes / 12 * CDbl(l_rs("precantcuo")) / 100)) * Monto_a_Cancelar) - (Monto_a_Cancelar)) / CDbl(l_rs("precantcuo"))
            Cuovtotal = (Valor_Cuota + (Valor_Cuota * IntCuo)) + iva + otroGasto
            
            'capital = Saldo_a_Pagar - Valor_Cuota
            'FB - 30/06/2015 - Se modifica el valor del Capital para Interes Pre_Sac: capital = Valor de la cuota.
            capital = Valor_Cuota
            'Hasta aca
            Saldo_a_Pagar = Saldo_a_Pagar - Valor_Cuota
            
            l_diavto = diaVto
            
            If Mes = 1 Or Mes = 3 Or Mes = 5 Or Mes = 7 Or Mes = 8 Or Mes = 10 Or Mes = 12 Then
               If diaVto > 31 Then
                  l_diavto = 31
               End If
            ElseIf Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
                   If diaVto > 30 Then
                      l_diavto = 30
                   End If
            ElseIf Mes = 2 Then
                   'If anioBisiesto(Anio) Then
                    '  If diaVto > 29 Then
                     '    l_diavto = 29
                      'End If
                   'ElseIf diaVto > 28 Then
                          l_diavto = 28
                   'End If
            End If
               
            Fechavto = "'" & CStr(l_diavto) & "/" & CStr(Mes) & "/" & CStr(Anio) & "'"
            'fechavto =  Cdate(diavto & "/" & mes & "/" & anio)
                        
            l_sql = "INSERT INTO pre_cuota "
            l_sql = l_sql & "(prenro,cuoimp,cuocancela, cuonrocuo,cuogastos,cuoiva,cuototal,cuocapital,cuointeres,cuosaldo,cuomes,cuoano,cuofecvto) "
            l_sql = l_sql & "VALUES (" & l_prenro & "," & Valor_Cuota & "," & 0 & "," & CInt(i) & ","
            'l_sql = l_sql & otroGasto & "," & iva & "," & Cuovtotal & "," & capital & "," & IntCuo * 100 & ","
            'FB - 30/06/2015 - Se quita el * 100 para mostrar correctamente el interes
            l_sql = l_sql & otroGasto & "," & iva & "," & Cuovtotal & "," & capital & "," & IntCuo & ","
            l_sql = l_sql & Saldo_a_Pagar & "," & Mes & "," & Anio & "," & Fechavto & ")"
            
            Flog.writeline "Inserto datos:" & StrSql
            objConn.Execute l_sql, , adExecuteNoRecords
            
        Else
            'QUINCENAL
            
            Cquin = l_rs("prequin")
            
            Valor_Cuota = Round(deuda / cuotas, 2)
            Monto_a_Cancelar = deuda
            Saldo_a_Pagar = deuda
            
            i = 1
            
            Do While CInt(i) <= CInt(cuotas) - 1
                
               cambio = cambio + Round(Valor_Cuota, 2)
               IntCuo = (((1 + (interes / 12 * CDbl(l_rs("precantcuo")) / 100)) * Monto_a_Cancelar) - (Monto_a_Cancelar)) / CDbl(l_rs("precantcuo")) / 2
               Cuovtotal = (Valor_Cuota + (Valor_Cuota * IntCuo)) + iva + otroGasto
                
               'capital = Saldo_a_Pagar - Valor_Cuota
               'FB - 30/06/2015 - Se modifica el valor del Capital para Interes Pre_Sac: capital = Valor de la cuota.
               capital = Valor_Cuota
               'Hasta aca
               Saldo_a_Pagar = Saldo_a_Pagar - Valor_Cuota
               
               l_diavto = diaVto
               
               If Mes = 1 Or Mes = 3 Or Mes = 5 Or Mes = 7 Or Mes = 8 Or Mes = 10 Or Mes = 12 Then
                  If diaVto > 31 Then
                     l_diavto = 31
                  End If
               ElseIf Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
                      If diaVto > 30 Then
                         l_diavto = 30
                      End If
               ElseIf Mes = 2 Then
                    '  If anioBisiesto(Anio) Then
                     '    If diaVto > 29 Then
                      '      l_diavto = 29
                       '  End If
                      'ElseIf diaVto > 28 Then
                             l_diavto = 28
                      'End If
               End If
               
               Fechavto = "'" & CStr(l_diavto) & "/" & CStr(Mes) & "/" & CStr(Anio) & "'"
                                
                l_sql = "INSERT INTO pre_cuota "
                l_sql = l_sql & "(prenro,cuoimp,cuocancela, cuonrocuo,cuogastos,cuoiva,cuototal,cuocapital,cuointeres,cuosaldo,cuomes,cuoano,cuofecvto, cuoquin) "
                l_sql = l_sql & "VALUES (" & l_prenro & "," & Valor_Cuota & "," & 0 & "," & CInt(i) & ","
                'l_sql = l_sql & otroGasto & "," & iva & "," & Cuovtotal & "," & capital & "," & IntCuo * 100 & ","
                'FB - 30/06/2015 - Se quita el * 100 para mostrar correctamente el interes
                l_sql = l_sql & otroGasto & "," & iva & "," & Cuovtotal & "," & capital & "," & IntCuo & ","
                l_sql = l_sql & Saldo_a_Pagar & "," & Mes & "," & Anio & "," & Fechavto & "," & Cquin & ")"
                
                Flog.writeline "Inserto datos Quincenal:" & StrSql
                objConn.Execute l_sql, , adExecuteNoRecords
                
                If Cquin = 2 Then
                   Cquin = 1
                   diaVto = diaVto - 15
                   If Mes = 12 Then
                      Mes = 1
                      Anio = Anio + 1
                      Else: Mes = Mes + 1
                   End If
                   Else: Cquin = 2
                        diaVto = diaVto + 15
                End If
                
               i = i + 1
                
            Loop
            
            Valor_Cuota = Valor_Cuota + (Monto_a_Cancelar - (cambio + Round(Valor_Cuota, 2)))
            IntCuo = (((1 + (interes / 12 * CDbl(l_rs("precantcuo")) / 100)) * Monto_a_Cancelar) - (Monto_a_Cancelar)) / CDbl(l_rs("precantcuo")) / 2
            Cuovtotal = (Valor_Cuota + ((Valor_Cuota * IntCuo * 100) / 100)) + iva + otroGasto
            
            'capital = Saldo_a_Pagar - Valor_Cuota
            'FB - 30/06/2015 - Se modifica el valor del Capital para Interes Pre_Sac: capital = Valor de la cuota.
            capital = Valor_Cuota
            'Hasta aca
            Saldo_a_Pagar = Saldo_a_Pagar - Valor_Cuota
            
            l_diavto = diaVto
            
            If Mes = 1 Or Mes = 3 Or Mes = 5 Or Mes = 7 Or Mes = 8 Or Mes = 10 Or Mes = 12 Then
               If diaVto > 31 Then
                  l_diavto = 31
               End If
            ElseIf Mes = 4 Or Mes = 6 Or Mes = 9 Or Mes = 11 Then
                   If diaVto > 30 Then
                      l_diavto = 30
                   End If
            ElseIf Mes = 2 Then
                   'If anioBisiesto(Anio) Then
                    '  If diaVto > 29 Then
                     '    l_diavto = 29
                      'End If
                   'ElseIf diaVto > 28 Then
                          l_diavto = 28
                   'End If
            End If
               
            Fechavto = "'" & CStr(l_diavto) & "/" & CStr(Mes) & "/" & CStr(Anio) & "'"
                        
            l_sql = "INSERT INTO pre_cuota "
            l_sql = l_sql & "(prenro,cuoimp,cuocancela, cuonrocuo,cuogastos,cuoiva,cuototal,cuocapital,cuointeres,cuosaldo,cuomes,cuoano,cuofecvto, cuoquin) "
            l_sql = l_sql & "VALUES (" & l_prenro & "," & Valor_Cuota & "," & 0 & "," & CInt(i) & ","
            'l_sql = l_sql & otroGasto & "," & iva & "," & Cuovtotal & "," & capital & "," & IntCuo * 100 & ","
            'FB - 30/06/2015 - Se quita el * 100 para mostrar correctamente el interes
            l_sql = l_sql & otroGasto & "," & iva & "," & Cuovtotal & "," & capital & "," & IntCuo & ","
            l_sql = l_sql & Saldo_a_Pagar & "," & Mes & "," & Anio & "," & Fechavto & "," & Cquin & ")"
            
            Flog.writeline "Inserto datos:" & StrSql
            objConn.Execute l_sql, , adExecuteNoRecords
            
        End If
    
    End If ' not l_rs.eof
Else 'cuota <> 0
    Flog.writeline Espacios(Tabulador * 1) & "El prestamo: (" & l_prenro & ") no posee configurada la cantidad de cuotas, o la cantidad de cuotas es 0(cero), revisar el prestamo "
End If ' cuotas <> 0
                                        
End Sub
