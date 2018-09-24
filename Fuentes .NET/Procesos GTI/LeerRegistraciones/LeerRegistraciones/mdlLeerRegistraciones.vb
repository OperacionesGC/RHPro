Imports System.Data.OleDb
Imports System
Imports System.Diagnostics
Imports System.ComponentModel
Imports System.Data
Imports System.Collections.Generic
Imports LeerRegistraciones.ServiceReference1
Imports System.ServiceModel
Imports System.Configuration
Imports System.Reflection
Imports RHPro.Shared.Sys.Crypt


Module mdlLeerRegistraciones
    ''Version = "1.0.0.1"
    ''Const UltimaModificacion = "Se agregaron mejoras varias ademas de un modulo para imprimir vesion"
    'Const FechaModificacion = "29/01/2012"

    ''Version = "1.0.0.2"
    ''Const UltimaModificacion = "Se modificó el proceso para que levante la fecha en la que se esta ejecuntado para utilizarlo del planificador"
    'Const FechaModificacion = "19/06/2012"

    'Version = "1.0.0.2"
    'Const UltimaModificacion = "Se modificó el proceso para que levante la fecha en la que se esta ejecuntado para utilizarlo del planificador"
    'Const FechaModificacion = "05/06/2013"

    'Version = "1.0.0.3"
    'Const UltimaModificacion = "Se creo el modelo insertarFormatoSpec"
    'Const FechaModificacion = "11/12/2013"

    'Version = "1.0.0.4"
    'Const UltimaModificacion = "Se corrige modelo insertarFormatoSpec"
    'Const FechaModificacion = "17/01/2014"

    'Version = "1.0.0.5"
    'Const UltimaModificacion = "Se corrige la barra de progreso"
    'Const FechaModificacion = "25/04/2014"

    'Version = "2.0"
    'Const UltimaModificacion = "Mejora de codigo"
    'Const FechaModificacion = "08/05/2014"

    'Version = "2.0.0.1"
    'Const UltimaModificacion = "Importacion de Novedades - CAS-25711 - DABRA - INTERFACE PARA IMPORTACION DE NOVEDADES - LED "
    'Const FechaModificacion = "27/06/2014"

    'Version = "2.0.0.2"
    'Const UltimaModificacion = "Importacion de Novedades - CAS-25711 - DABRA - LED - Importacion de licencias. "
    'Const FechaModificacion = "23/12/2014"


    'Version = "2.0.0.3"
    'Const UltimaModificacion = "CAS-28477 - Raffo - Problema de procesamiento ON-LINE con Interface SPEC- Matias, Fernandez."
    'Const FechaModificacion = "06/01/2014"

    'Version = "2.0.0.4"
    'Const UltimaModificacion = "CAS-29210 - RAFFO - Error Interface SPEC - LecturaRegistraciones - LED."
    'Const FechaModificacion = "09/02/2015"

    'version = "2.0.0.5"
    'Const UltimaModificacion = CAS-29837 - RAFFO - Error en procesamiento de lectura de registraciones - Matias Fernandez
    'Const FechaModificacion = "26/03/2015"

    'version = "2.0.0.6"
    'Const UltimaModificacion = CAS-29837 - RAFFO - Error en procesamiento de lectura de registraciones - Matias Fernandez
    ' Const FechaModificacion = "23/04/2015"

    'version = "2.0.0.7"
    'Const UltimaModificacion = CAS-36895 - GAMA - Error en procesamiento planificado - Matias Fernandez
    'Const FechaModificacion = "26/04/2016"

    'version = "2.0.0.8"
    'Const UltimaModificacion = CAS-36895 - GAMA - Error en procesamiento planificado - Matias Fernandez
    'Const UltimaModificacion = Se corrige la fecha de insersion en la tabla batch proceso para el prc30 y prc01 spec
    Const FechaModificacion = "29/04/2016"

    Public Path As String
    Public NArchivo As String
    'Dim Rta
    'Dim ObjetoVentana As Object
    'Dim HuboError As Boolean
    'Dim Nro_Modelo As Integer
    Public Etiqueta
    'Dim Separador As String
    Public PathSAP As String
    'Dim PathProcesos As String
    Public NroProceso As Long
    Public conexion As OleDbConnection
    Dim StrSql As String
    Dim da As OleDbDataAdapter
    Dim cmd As New OleDbCommand()
    Dim transaction As OleDbTransaction
    Public Terceros(0) As ProcesamientoOnline
    Public FLog As New ManejoArchivo()
    Dim Errores As Integer = 0
    Public fechaAltaSist As String
    Public rhpro_nombreEmpresa As String
    'variables del progreso
    Public IncPorc As Single = 0
    Public Progreso As Single = 0

    Dim g_fdesde As Date    'MDF
    Dim g_fhasta As Date    'MDF
    Dim g_spec As Boolean




    Sub Main()
        Dim strCmdLine As String
        Dim Proc_ONLINE As Boolean = False 'Si el procesamiento On Line está o no activo
        Dim HC_ONLINE As Boolean = False 'Si genera o no procesos de Horario Cumplido por procesamiento On Line
        Dim AD_ONLINE As Boolean = False 'Si genera o no procesos de Acumulado Diario por procesamiento On Line
        Dim fechaaux As String
        Dim ArrParametros
        Dim dtTabla As New DataTable

        Dim usuario As String
        Dim NroProcesoHC As Long
        Dim NroProcesoAD As Long
        Dim EncriptStrconexion As Boolean
        Dim c_seed As String
        Dim fecProcDesde As DateTime
        Dim fecProcHasta As DateTime
        Dim Fechas As New Fechas

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


        'Carga las configuraciones basicas, formato de fecha, string de conexion,
        'tipo de BD y ubicacion del archivo de log
        'Call CargarConfiguracionesBasicas()
        Dim confiBasica As New DataAccess(True)
        'Dim asm As Assembly = Assembly.LoadFrom("RHPro.Shared.Sys.dll")
        'EAM- Crea el archivo de log
        FLog.CrearArchivo(confiBasica.PathFLog, "LecturaReg " & "-" & NroProceso & "-" & Format(Convert.ToDateTime(Now.ToString), "dd-MM-yyyy") & ".log")

        Try            
            Try
                If EncriptStrconexion Then
                    Dim crypt As New RHPro.Shared.Sys.Crypt
                    conexion = New OleDbConnection(crypt.RHDecrypt(c_seed, confiBasica.Conexion))                    
                Else
                    conexion = New OleDbConnection(confiBasica.Conexion)
                End If
                conexion.Open()
                cmd.Connection = conexion
                conexion.Close()
            Catch ex As Exception
                Throw New Exception("Error de Conexión" & conexion.ConnectionString.ToString)

            End Try


            'EAM- Imprime los datos de la versión y seguido el PID
            mdlVersion.Main(FechaModificacion)

            Dim PID As Process = Process.GetCurrentProcess()
            FLog.EscribirLinea("")
            FLog.EscribirLinea("PID = " & PID.Id)

            'Verifica si ya hay algun proceso corriendo
            If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
                FLog.EscribirLinea("Ya hay una instancia del proceso corriendo. Queda Pendiente." & Format(Now, "dd/mm/yyyy hh:mm:ss"))

                transaction = conexion.BeginTransaction()
                StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & confiBasica.ConvFecha(Now) & _
                        ", bprcprogreso = 0, bprcestado = 'Pendiente', bprcpid = " & PID.ToString & " WHERE bpronro = " & NroProceso
                conexion.Open()
                cmd = New OleDbCommand(StrSql, conexion)
                cmd.ExecuteNonQuery()

                transaction.Commit()
                conexion.Close()
                FLog.EscribirLinea("")
            End If

            'Cambio el estado del proceso a Procesando
            conexion.Open()
            da = New OleDbDataAdapter(StrSql, conexion)
            StrSql = "UPDATE batch_proceso SET bprcpid = " & PID.Id & ", bprchorainicioej = '" & DateTime.Today.ToString("dd/MM/yyyy") & "', bprcfecinicioej = " & confiBasica.ConvFecha(Now) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
            cmd.CommandText = StrSql
            cmd.ExecuteNonQuery()
            conexion.Close()


            StrSql = "SELECT iduser,bprcfecdesde,bprcfechasta FROM batch_proceso WHERE bpronro = " & NroProceso
            da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
            da.Fill(dtTabla)
            If (dtTabla.Rows.Count > 0) Then
                usuario = dtTabla.Rows(0).Item(0).ToString
                fecProcDesde = dtTabla.Rows(0).Item("bprcfecdesde").ToString()
                fecProcHasta = dtTabla.Rows(0).Item("bprcfechasta").ToString()

            End If
            FLog.EscribirLinea("Inicio Transferencia " & Format(Now, "dd-MM-yyyy hh:mm:ss"))

            'EAM- Se obtiene el moledo e inserta con el formato correspondiente
            Call ComenzarTransferenciaReg()

            FLog.EscribirLinea("")
            FLog.EscribirLinea("Procesamiento ON-LINE")
            StrSql = "SELECT * FROM GTI_puntos_proc " & _
                     " INNER JOIN GTI_proc_online ON GTI_puntos_proc.ptoprcnro = GTI_proc_online.ptoprcnro " & _
                     " WHERE GTI_puntos_proc.ptoprcid = 19 AND GTI_puntos_proc.ptoprcact = -1 AND GTI_proc_online.prconlineact=-1"
            da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
            dtTabla = New DataTable
            da.Fill(dtTabla)

            'EAM- Verifica que punto de procesamiento esta activo
            If (dtTabla.Rows.Count = 0) Then
                Proc_ONLINE = False
                FLog.EscribirLinea("Procesamiento ON-LINE, Lectura de Registraciones, punto de procesamiento Inactivo.", 1)
            Else
                FLog.EscribirLinea("Procesamiento ON-LINE, hay puntos de procesamiento Activos ==>", 1)
                Proc_ONLINE = True
                For Each MiDataRow As DataRow In dtTabla.Rows
                    Select Case MiDataRow("btprcnro")
                        Case 1
                            HC_ONLINE = True
                            FLog.EscribirLinea("Puntos de procesamiento: Horario Cumplido", 1)
                        Case 2
                            AD_ONLINE = True
                            FLog.EscribirLinea("Puntos de procesamiento: Acumulado Diario", 1)
                        Case Else
                            FLog.EscribirLinea("Puntos de procesamiento desconocido. " & MiDataRow("btprcnro"), 1)
                    End Select
                Next
            End If
            FLog.EscribirLinea("Cantidad Empleado para procesar On-Line: " & UBound(Terceros), 1)
            If Proc_ONLINE Then


                'EAM- Recorre todos los terceros que se insertaron registraciones e inserta los regitros si tiene configurado los puntos de procesamiento
                StrSql = "SELECT DISTINCT regfecha FROM gti_registracion WHERE regfecha >= " & confiBasica.ConvFecha(DateAdd("d", -1, fecProcDesde)) & " AND regfecha <= " & confiBasica.ConvFecha(fecProcHasta)
                da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                dtTabla = New DataTable
                da.Fill(dtTabla)
                FLog.EscribirLinea("LaSS " & StrSql)
                FLog.EscribirLinea("Cant lass " & dtTabla.Rows.Count)
                For Each Fila As DataRow In dtTabla.Rows
                    If HC_ONLINE Then
                        FLog.EscribirLinea("Genero HC para el " & Fila.Item("regfecha"))
                        'Inserto en batch_proceso un HC

                        If Not g_spec Then


                            StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprcparam,bprchora, bprcfecdesde, bprcfechasta, " &
                                        "bprcestado, empnro) " &
                                        "VALUES (1," & confiBasica.ConvFecha(Now) & ",'" & usuario & "','-1.0','" & Format(Now, "hh:mm:ss ") & "' " &
                                 ", " & confiBasica.ConvFecha(DateAdd("d", -1, Fila.Item("regfecha"))) & ", " & confiBasica.ConvFecha(Fila.Item("regfecha")) &
                                 ", 'Temp', 0)"
                        Else

                            StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprcparam,bprchora, bprcfecdesde, bprcfechasta, " &
                                        "bprcestado, empnro) " &
                                        "VALUES (1," & confiBasica.ConvFecha(Now) & ",'" & usuario & "','-1.0','" & Format(Now, "hh:mm:ss ") & "' " &
                                 ", '" & g_fdesde & "', '" & g_fhasta & "', 'Temp', 0)"

                        End If

                        conexion.Open()
                        cmd.CommandText = StrSql
                        cmd.ExecuteNonQuery()

                        NroProcesoHC = confiBasica.getLastIdentity(conexion, "Batch_Proceso")
                            conexion.Close()
                            FLog.EscribirLinea("Disparo HC. Nro de proceso: " & NroProcesoHC)
                        End If

                        'Procese el dia anterior y el actual
                        If AD_ONLINE Then
                        FLog.EscribirLinea("genero AD para el " & Fila.Item("regfecha"))
                        'Inserto en batch_proceso un AD
                        If Not g_spec Then

                            StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprcparam, bprchora, bprcfecdesde, bprcfechasta, " &
                                         "bprcestado, empnro) " &
                                         "VALUES (2," & confiBasica.ConvFecha(Now) & ",'" & usuario & "','-1.0','" & Format(Now, "hh:mm:ss ") & "' " &
                                 ", " & confiBasica.ConvFecha(DateAdd("d", -1, Fila.Item("regfecha"))) & ", " & confiBasica.ConvFecha(Fila.Item("regfecha")) &
                                 ", 'Temp', 0)"
                        Else
                            StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprcparam, bprchora, bprcfecdesde, bprcfechasta, " &
                                   "bprcestado, empnro) " &
                                   "VALUES (2," & confiBasica.ConvFecha(Now) & ",'" & usuario & "','-1.0','" & Format(Now, "hh:mm:ss ") & "' " &
                           ", '" & g_fdesde & "', '" & g_fhasta & "', 'Temp', 0)"
                        End If
                        conexion.Open()
                            cmd.CommandText = StrSql
                            cmd.ExecuteNonQuery()

                            'Recupera el numero de proceso generado
                            NroProcesoAD = confiBasica.getLastIdentity(conexion, "Batch_Proceso")
                            conexion.Close()
                            FLog.EscribirLinea("Disparo AD. Nro de proceso: " & NroProcesoAD)
                        End If


                        FLog.EscribirLinea("")
                    FLog.EscribirLinea("Inserto en batch_empleados los empleados de los procesos generados.")

                    'Inserta en Batch_empelado para cada fecha del proceso de AD y HC
                    FLog.EscribirLinea("--------------->hay terceros: " & UBound(Terceros))

                    For j = 1 To UBound(Terceros)
                        Try
                            FLog.EscribirLinea("---------------------------------------------------------------------------------------------------------")
                            FLog.EscribirLinea("regfecha: " & Fila.Item("regfecha") & " y fecha tercero: " & Terceros(j).Fecha)
                            FLog.EscribirLinea("---------------------------------------------------------------------------------------------------------")

                            If (CDate(Fila.Item("regfecha").ToString().Replace("'", "")) = CDate(Terceros(j).Fecha.Replace("'", ""))) Then
                                'Para HC
                                If HC_ONLINE Then
                                    StrSql = "INSERT INTO batch_empleado (bpronro, ternro, estado) VALUES (" & _
                                             NroProcesoHC & "," & Terceros(j).Ternro & ", NULL )"
                                    conexion.Open()
                                    cmd.CommandText = StrSql
                                    cmd.ExecuteNonQuery()
                                    conexion.Close()
                                    FLog.EscribirLinea("Inserto para HC ternro: " & Terceros(j).Ternro & " y bpronro: " & NroProcesoHC)
                                End If

                                'Para AD
                                If AD_ONLINE Then
                                    StrSql = "INSERT INTO batch_empleado (bpronro, ternro, estado) VALUES (" & _
                                             NroProcesoAD & "," & Terceros(j).Ternro & ", NULL )"
                                    conexion.Open()
                                    cmd.CommandText = StrSql
                                    cmd.ExecuteNonQuery()
                                    conexion.Close()
                                    FLog.EscribirLinea("Inserto para AD ternro: " & Terceros(j).Ternro & " y bpronro: " & NroProcesoAD)
                                End If
                            End If
                        Catch ex As Exception
                            FLog.EscribirLinea(ex.Message)
                            FLog.EscribirLinea("Sql: " & StrSql)
                            conexion.Close()
                        End Try
                    Next

                    'Actualizo el estado de los procesos a Pendiente
                    StrSql = "UPDATE Batch_Proceso SET bprcestado ='Pendiente'"
                    StrSql = StrSql & " WHERE bpronro = " & NroProcesoHC
                    conexion.Open()
                    cmd.CommandText = StrSql
                    cmd.ExecuteNonQuery()
                    conexion.Close()

                    'Actualizo el estado de los procesos a Pendiente
                    StrSql = "UPDATE Batch_Proceso "
                    StrSql = StrSql & " SET bprcestado ='Pendiente'"
                    StrSql = StrSql & " WHERE bpronro = " & NroProcesoAD
                    conexion.Open()
                    cmd.CommandText = StrSql
                    cmd.ExecuteNonQuery()
                    conexion.Close()
                Next
            End If


        Catch ex As Exception
            Errores = -1
            FLog.EscribirLinea(ex.Message)
            FLog.EscribirLinea("Sql: " & StrSql)
        Finally
            If Not Errores Then
                StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & confiBasica.ConvFecha(Now) & ",bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
                FLog.EscribirLinea("Proceso terminado Correctamente")
            Else
                StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & confiBasica.ConvFecha(Now) & ", bprcprogreso =100, bprcestado = 'Error' WHERE bpronro = " & NroProceso
            End If

            conexion.Open()
            cmd.CommandText = StrSql
            cmd.ExecuteNonQuery()
            conexion.Close()

            TerminarTransferencia()
        End Try

    End Sub

    Public Sub ComenzarTransferenciaReg()
        Dim dtDatos As New DataTable
        Dim dtConfrep As New DataTable
        Dim Directorio As String
        Dim parametros As String
        Dim arrayModelo(3) As String
        Dim idEmpresa As String
        Dim psw As String
        g_spec = False

        'por defecto cargamos lectura o spec depende cual esta activo
        parametros = "-1@0@0"
        Call levantarParametros(parametros)


        'levanto los parametros del confrep
        StrSql = "SELECT confval, confval2 FROM confrep WHERE repnro = 441 and confnrocol = 1 and upper(conftipo) = 'PSW' "
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtConfrep)
        If (dtConfrep.Rows.Count > 0) Then
            idEmpresa = dtConfrep.Rows(0).Item("confval").ToString
            psw = Trim(dtConfrep.Rows(0).Item("confval2").ToString)
        Else
            idEmpresa = ""
            psw = ""
            FLog.EscribirLinea("No se encuentran configurados los datos para conexion con quickpass, reporte 441")
        End If

        StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)
        If (dtDatos.Rows.Count > 0) Then
            Directorio = Trim(dtDatos.Rows(0).Item("sis_dirsalidas").ToString)
        Else            
            Exit Sub
            'cambiar por el manejador de error
        End If

        arrayModelo = parametros.Split("@")

        For indice = 0 To UBound(arrayModelo)
            Select Case indice
                Case 0
                    If arrayModelo(indice) = "-1" Then
                        'Controlo si esta activo el modelo
                        dtDatos = New DataTable
                        StrSql = "SELECT * FROM modelo WHERE modtipo = 3 and modestado = -1 AND Modnro = 190"
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        da.Fill(dtDatos)
                        If dtDatos.Rows.Count > 0 Then
                            InsertaFormatoQPass(idEmpresa, psw)
                        End If
                    End If
                Case 1
                    If arrayModelo(indice) = "-1" Then
                        'Controlo si esta activo el modelo
                        dtDatos = New DataTable
                        StrSql = "SELECT * FROM modelo WHERE modtipo = 3 and modestado = -1 AND Modnro = 168"
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        da.Fill(dtDatos)
                        If dtDatos.Rows.Count > 0 Then
                            FLog.EscribirLinea("llamo a spec()")
                            g_spec = True
                            spec()
                            FLog.EscribirLinea("fin de spec()")
                        End If
                    End If
                Case 2 'Importacion de novedades
                    If arrayModelo(indice) = "-1" Then
                        InsertarNovedades(idEmpresa, psw)
                    End If

                Case 3 'Importacion de licencias
                    If arrayModelo(indice) = "-1" Then
                        InsertarLicencias(idEmpresa, psw)
                    End If

            End Select

        Next
    End Sub

    Public Sub InsertarNovedades(ByVal idempresa As String, ByVal psw As String)
        Dim dsDatos As New DataSet
        Dim dtDatos As New DataTable
        Dim dtDatosAux As DataTable
        Dim dtTercero As DataTable
        Dim dtLicencia As DataTable
        Dim dtMotivo, dtNovedad As DataTable
        Dim objLecturaQP As New ServiceMovimientos
        Dim FDesde, FHasta As String
        Dim Fechas As New Fechas
        Dim IncPorc As Single = 0
        Dim Progreso As Single = 0
        Dim diaFijo As Integer
        Dim gnovnro As Integer
        Dim confiBasica As New DataAccess(True)
        'Dim agregarProcOnline As Boolean

        FLog.EscribirLinea("Ingresa al modelo de Novedades")
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)
        If (dtDatos.Rows.Count > 0) Then
            'Le resta uno a la fecha desde porque si es online puede que las ultimas registraciones sean del dia anterior y no las toma
            FDesde = Fechas.cambiaFecha(dtDatos.Rows(0).Item("bprcfecdesde").ToString)
            FHasta = Fechas.cambiaFecha(dtDatos.Rows(0).Item("bprcfechasta").ToString)
            dsDatos = objLecturaQP.wsListarNovedades(idempresa, psw, "-1", FDesde, FHasta)

            '------------------ carga de datos manual para pruebas
            'Dim dr As DataRow = dsDatos.Tables(0).NewRow()
            'dr("FechaDesde") = "201406270000"
            'dr("FechaHasta") = "201406280000"
            'dr("Motivo") = "Enfermedad Familiar"
            'dr("Legajo") = "2"
            'dr("Observaciones") = ""
            'dsDatos.Tables(0).Rows.Add(dr)
            '------------------

            'Calcula en incremento del porcentaje
            If (dsDatos.Tables(0).Rows.Count > 0) Then
                FLog.EscribirLinea(dsDatos.Tables(0).Rows.Count & " Archivo de Novedades encontrado " & Format(Now, "dd/mm/yyyy hh:mm:ss"), 1)
                IncPorc = 100 / dsDatos.Tables(0).Rows.Count
            ElseIf (dsDatos.Tables(0).Rows.Count = 0) Then
                IncPorc = 100
            End If

            For Each MiDataRow As DataRow In dsDatos.Tables(0).Rows
                Try
                    FLog.EscribirLinea("Fecha desde de la novedad: " & MiDataRow.Item("FechaDesde"), 1)
                    FLog.EscribirLinea("Fecha hasta de la novedad: " & MiDataRow.Item("FechaHasta"), 1)
                    'busco que no exista ninguna novedad en el rango de fechas
                    StrSql = " SELECT gnovnro FROM gti_novedad " &
                            " INNER JOIN empleado on empleado.ternro = gti_novedad.gnovotoa AND empleado.empleg = " & MiDataRow.Item("Legajo") & _
                            " WHERE (gnovdesde <= " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & _
                            " AND (gnovhasta >= " & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & " or gnovhasta >= " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & ")) " & _
                            " OR (gnovdesde >= " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & " AND (gnovdesde <= " & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & "))"
                    dtDatosAux = New DataTable
                    da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                    da.Fill(dtDatosAux)
                    FLog.EscribirLinea("Legajo: " & MiDataRow.Item("Legajo") & " fecha desde: " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & " fecha hasta: " & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy"), 1)

                    If dtDatosAux.Rows.Count > 0 Then
                        FLog.EscribirLinea("Error. Ya existe novedad horaria para el rango de fecha.", 1)
                        FLog.EscribirLinea("SQL: " & StrSql, 1)
                    Else
                        'Controlo que exista el empleado
                        StrSql = "SELECT ternro FROM empleado WHERE empleg = " & MiDataRow.Item("Legajo")
                        dtTercero = New DataTable
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        da.Fill(dtTercero)

                        If dtTercero.Rows.Count > 0 Then
                            FLog.EscribirLinea("Legajo: " & MiDataRow.Item("Legajo") & " Encontrado", 1)
                            'controlo que no tenga licencias cargadas en el rango de fechas
                            StrSql = " SELECT emp_licnro FROM emp_lic WHERE empleado = " & dtTercero.Rows(0).Item("ternro") & " AND " & _
                                    " ((elfechadesde <= " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & _
                                    " AND (elfechahasta >= " & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & " or elfechahasta >= " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & ")) " & _
                                    " OR (elfechadesde >= " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & " AND (elfechadesde <= " & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & ")))"
                            dtLicencia = New DataTable
                            da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                            da.Fill(dtLicencia)
                            If dtLicencia.Rows.Count <= 0 Then

                                'busco el mapeo motivo - cod motivo configurado
                                StrSql = " SELECT codinterno FROM mapeo_sap WHERE upper(codexterno) = '" & UCase(MiDataRow.Item("Motivo")) & "' AND upper(tablaref) = 'tipo_motivo'"
                                dtMotivo = New DataTable
                                da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                da.Fill(dtMotivo)

                                If dtMotivo.Rows.Count > 0 Then
                                    FLog.EscribirLinea("Mapeo Motivo: " & MiDataRow.Item("Motivo") & " Encontrado", 1)
                                    'busco el mapeo motivo - cod tipo novedad configurado
                                    StrSql = " SELECT codinterno FROM mapeo_sap WHERE upper(codexterno) = '" & UCase(MiDataRow.Item("Motivo")) & "' AND upper(tablaref) = 'tipo_novedad'"
                                    dtNovedad = New DataTable
                                    da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                    da.Fill(dtNovedad)
                                    If dtNovedad.Rows.Count > 0 Then
                                        FLog.EscribirLinea("Mapeo Motivo - Tipo novedad: " & MiDataRow.Item("Motivo") & " Encontrado", 1)
                                        'controlo q el motivo exista 
                                        StrSql = " SELECT motdesabr, motnro FROM gti_motivo WHERE motnro = " & dtMotivo.Rows(0).Item("codinterno")
                                        dtMotivo = New DataTable
                                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                        da.Fill(dtMotivo)
                                        If dtMotivo.Rows.Count > 0 Then
                                            FLog.EscribirLinea("Motivo en Rhpro: " & dtMotivo.Rows(0).Item("motnro") & " Encontrado en Rhpro", 1)
                                            'controlo q el tipo de novedad exista 
                                            StrSql = " SELECT gtnovnro FROM gti_tiponovedad WHERE gtnovnro = " & dtNovedad.Rows(0).Item("codinterno")
                                            dtNovedad = New DataTable
                                            da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                            da.Fill(dtNovedad)

                                            If dtNovedad.Rows.Count > 0 Then
                                                FLog.EscribirLinea("Tipo de novedad en Rhpro: " & dtNovedad.Rows(0).Item("gtnovnro") & " Encontrado en Rhpro", 1)
                                                'Se controla si es dia completo o parcial fija 
                                                If MiDataRow.Item("FechaDesde").ToString.Substring(8, 4) <> "0000" Or MiDataRow.Item("FechaHasta").ToString.Substring(8, 4) <> "0000" Then
                                                    diaFijo = 2
                                                    StrSql = " INSERT INTO gti_novedad (gnovdesabr,motnro,gnovotoa,gnovdesde,gnovhasta,gnovestado,gtnovnro,gnovdiacompleto,gnovtipo,gnovhoradesde, gnovhorahasta,gnovdesext) VALUES " & _
                                                         " ('" & dtMotivo.Rows(0).Item("motdesabr") & "'," & dtMotivo.Rows(0).Item("motnro") & _
                                                         " ," & dtTercero.Rows(0).Item("ternro") & "," & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & _
                                                         " ," & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & ",' '," & dtNovedad.Rows(0).Item("gtnovnro") & ",-1," & diaFijo & _
                                                         " ,'" & MiDataRow.Item("FechaDesde").ToString.Substring(8, 4) & "','" & MiDataRow.Item("FechaHasta").ToString.Substring(8, 4) & "'" & _
                                                         " , '" & MiDataRow.Item("Observaciones") & "') "
                                                Else
                                                    diaFijo = 1
                                                    StrSql = " INSERT INTO gti_novedad (gnovdesabr,motnro,gnovotoa,gnovdesde,gnovhasta,gnovestado,gtnovnro,gnovdiacompleto,gnovtipo,gnovdesext) VALUES " & _
                                                         " ('" & dtMotivo.Rows(0).Item("motdesabr") & "'," & dtMotivo.Rows(0).Item("motnro") & _
                                                         " ," & dtTercero.Rows(0).Item("ternro") & "," & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & _
                                                         " ," & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & ",' '," & dtNovedad.Rows(0).Item("gtnovnro") & ",-1," & diaFijo & _
                                                         " , '" & MiDataRow.Item("Observaciones") & "') "
                                                End If
                                                conexion.Open()
                                                cmd.CommandText = StrSql
                                                cmd.ExecuteNonQuery()


                                                'inserto en la tabla gti_justificacion
                                                gnovnro = confiBasica.getLastIdentity(conexion, "gti_novedad")
                                                conexion.Close()
                                                If diaFijo = 2 Then
                                                    StrSql = " INSERT INTO gti_justificacion ( jusanterior,juscodext,jusdesde,jusdiacompleto,jushasta,jussigla,jussistema,ternro,tjusnro,turnro,jushoradesde,jushorahasta,juseltipo,juselorden,juselmaxhoras ) " & _
                                                             " VALUES( -1," & gnovnro & "," & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & ",-1," & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & _
                                                             " ,'NOV',-1," & dtTercero.Rows(0).Item("ternro") & ",1,0,'" & MiDataRow.Item("FechaDesde").ToString.Substring(8, 4) & "','" & MiDataRow.Item("FechaHasta").ToString.Substring(8, 4) & "'" & _
                                                             " ," & diaFijo & ",null,null)"
                                                Else
                                                    StrSql = " INSERT INTO gti_justificacion ( jusanterior,juscodext,jusdesde,jusdiacompleto,jushasta,jussigla,jussistema,ternro,tjusnro,turnro,jushoradesde,jushorahasta,juseltipo,juselorden,juselmaxhoras ) " & _
                                                             " VALUES( -1," & gnovnro & "," & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & ",-1," & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & _
                                                             " ,'NOV',-1," & dtTercero.Rows(0).Item("ternro") & ",1,0,null,null" & _
                                                             " ," & diaFijo & ",null,null)"
                                                End If
                                                conexion.Open()
                                                cmd.CommandText = StrSql
                                                cmd.ExecuteNonQuery()
                                                conexion.Close()

                                                FLog.EscribirLinea("Novedad insertada para el legajo: " & MiDataRow.Item("Legajo") & " fecha: " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & " - " & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy"))
                                            Else
                                                FLog.EscribirLinea("Error. El tipo de novedad no existe en rhpro.")
                                            End If
                                            Else
                                                FLog.EscribirLinea("Error. El empleado con legajo: " & MiDataRow.Item("Legajo") & " posee licencias en el rango de fechas de la novedad.")
                                            End If
                                    Else
                                        FLog.EscribirLinea("Error. El tipo de motivo no existe en rhpro.")
                                    End If
                                Else
                                    FLog.EscribirLinea("Error. El tipo de novedad para el motivo: " & MiDataRow.Item("Motivo") & " no esta configurado.")
                                End If
                            Else
                                FLog.EscribirLinea("Error. El motivo: " & MiDataRow.Item("Motivo") & " no esta configurado.")
                            End If

                        Else
                            FLog.EscribirLinea("Error. El empleado no existe en el sistema.")
                        End If
                    End If

                Catch ex As Exception
                    FLog.EscribirLinea("ERROR SQL: " & StrSql)
                    FLog.EscribirLinea("Error al insertar la novedad del legajo " & MiDataRow.Item("Legajo"))
                Finally
                    'EAM- Actualiza el avance del proceso
                    Progreso = Progreso + IncPorc
                    conexion.Open()
                    da = New OleDbDataAdapter(StrSql, conexion)
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
                    cmd.CommandText = StrSql
                    cmd.ExecuteNonQuery()
                    conexion.Close()
                    FLog.EscribirLinea("----------------------------------------------")
                    FLog.EscribirLinea("")
                End Try
            Next
        End If

    End Sub

    Public Sub InsertarLicencias(ByVal idempresa As String, ByVal psw As String)
        Dim dsDatos As New DataSet
        Dim dtDatos As New DataTable
        'Dim dtDatosAux As DataTable
        Dim dtTercero As DataTable
        Dim dtLicencia As DataTable
        'Dim dtMotivo As DataTable
        Dim dtVacacion As DataTable
        Dim dtSitRevista As DataTable
        Dim dtHis_Estr As DataTable
        Dim objLecturaQP As New ServiceMovimientos
        Dim FDesde, FHasta As String
        Dim FecDesde, FecHasta As date
        Dim Fechas As New Fechas
        Dim IncPorc As Single = 0
        Dim Progreso As Single = 0
        'Dim diaFijo As Integer
        ' Dim gnovnro As Integer
        Dim confiBasica As New DataAccess(True)
        'Dim agregarProcOnline As Boolean
        Dim diaCompleto As Boolean
        Dim horaDesde As String
        Dim horaHasta As String
        Dim LicEstNro As String
        Dim Licencia As New Licencias
        Dim diasCorresp As Double
        Dim diasGozados As Double
        Dim diasBeneficio As Double
        Dim vacvendidos As Double
        Dim saldo As Double
        Dim insertar As Boolean
        Dim Lictipo As Integer
        Dim emp_licnro As String
        Dim PeriodoVac As Long
        Dim Estrnro_SitRev As Long

        FLog.EscribirLinea("Ingresa al modelo de Licencias")
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)
        If (dtDatos.Rows.Count > 0) Then
            'Le resta uno a la fecha desde porque si es online puede que las ultimas registraciones sean del dia anterior y no las toma
            FDesde = Fechas.cambiaFecha(DateAdd(DateInterval.Day, -1, CDate(dtDatos.Rows(0).Item("bprcfecdesde").ToString)))
            FHasta = Fechas.cambiaFecha(dtDatos.Rows(0).Item("bprcfechasta").ToString)
            dsDatos = objLecturaQP.wsListarNovedades(idempresa, psw, "-1", FDesde, FHasta)
            'dsDatos = objLecturaQP.wsListarNovedades(idempresa, psw, "1115", "201410070000", "201410070000")

            '------------------ carga de datos manual para pruebas
            'Dim dr As DataRow = dsDatos.Tables(0).NewRow()
            'dr("FechaDesde") = "201412040000"
            'dr("FechaHasta") = "201412040000"
            'dr("Motivo") = "Enfermedad Familiar"
            'dr("Legajo") = "2"
            'dr("Observaciones") = ""
            'dsDatos.Tables(0).Rows.Add(dr)
            '------------------

            'Calcula en incremento del porcentaje
            If (dsDatos.Tables(0).Rows.Count > 0) Then
                FLog.EscribirLinea(dsDatos.Tables(0).Rows.Count & " Licencias encontradas en Quickpass " & Format(Now, "dd/mm/yyyy hh:mm:ss"), 1)
                IncPorc = 100 / dsDatos.Tables(0).Rows.Count
            ElseIf (dsDatos.Tables(0).Rows.Count = 0) Then
                IncPorc = 100
            End If

            For Each MiDataRow As DataRow In dsDatos.Tables(0).Rows
                Try
                    FLog.EscribirLinea("Fecha desde de la novedad: " & MiDataRow.Item("FechaDesde"), 1)
                    FLog.EscribirLinea("Fecha hasta de la novedad: " & MiDataRow.Item("FechaHasta"), 1)

                    FecDesde = MiDataRow.Item("FechaDesde").substring(6, 2) & "/" & MiDataRow.Item("FechaDesde").substring(4, 2) & "/" & MiDataRow.Item("FechaDesde").substring(0, 4)

                    FecHasta = MiDataRow.Item("FechaHasta").substring(6, 2) & "/" & MiDataRow.Item("FechaHasta").substring(4, 2) & "/" & MiDataRow.Item("FechaHasta").substring(0, 4)
                    'controlo si es dia completo o parcial fija con los ultimos 4 caracteres de las fechas
                    If (Right(Trim(MiDataRow.Item("FechaDesde")), 4) = "0000") And ((Right(Trim(MiDataRow.Item("FechaHasta")), 4) = "0000") Or (Right(Trim(MiDataRow.Item("FechaHasta")), 4) = "2359")) Then
                        diaCompleto = True
                        Lictipo = 1
                        horaDesde = "0000"
                        horaHasta = "0000"
                    Else
                        diaCompleto = False
                        Lictipo = 2
                        horaDesde = Right(Trim(MiDataRow.Item("FechaDesde")), 4)
                        horaHasta = Right(Trim(MiDataRow.Item("FechaHasta")), 4)
                    End If

                    'Controlo que exista el empleado
                    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & MiDataRow.Item("Legajo")
                    dtTercero = New DataTable
                    da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                    da.Fill(dtTercero)

                    If dtTercero.Rows.Count > 0 Then
                        FLog.EscribirLinea("Legajo: " & MiDataRow.Item("Legajo") & " Encontrado", 1)
                        'controlo que no tenga licencias cargadas en el rango de fechas
                        StrSql = " SELECT emp_licnro FROM emp_lic WHERE empleado = " & dtTercero.Rows(0).Item("ternro") & " AND " & _
                                 " (elfechadesde <= " & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & _
                                 " AND elfechahasta >= " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & ")"
                        dtLicencia = New DataTable
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        da.Fill(dtLicencia)
                        If dtLicencia.Rows.Count <= 0 Then

                            'controlo q el tipo de licencia exista 
                            StrSql = " SELECT tdnro FROM tipdia WHERE upper(tddesc) = '" & UCase(MiDataRow.Item("Motivo")) & "'"
                            dtLicencia = New DataTable
                            da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                            da.Fill(dtLicencia)

                            If dtLicencia.Rows.Count > 0 Then
                                FLog.EscribirLinea("Tipo de Licencia: " & MiDataRow.Item("Motivo") & " Encontrado en Rhpro", 1)
                                'seteo el estado de la licencia en aprobado
                                LicEstNro = 2

                                'es licencia de vacaciones                
                                If (dtLicencia.Rows(0).Item("tdnro") = 2) Then
                                    'si al licencia es de vacaciones busco los dias gozados

                                    diasCorresp = Licencia.TotalDiasCorrespondientes(dtTercero.Rows(0).Item("ternro"))
                                    diasGozados = Licencia.TotalDiasGozados(dtTercero.Rows(0).Item("ternro"))
                                    diasBeneficio = Licencia.TotalDiasBeneficio(dtTercero.Rows(0).Item("ternro"))
                                    vacvendidos = Licencia.TotalDiasVendidos(dtTercero.Rows(0).Item("ternro"))
                                    saldo = Math.Round((CDbl(diasCorresp) + CDbl(diasBeneficio)) - CDbl(diasGozados) - CDbl(vacvendidos), 2)

                                    If saldo < 1 Then
                                        FLog.EscribirLinea(" No tiene saldo disponible para la licencia para el tercero: " & dtTercero.Rows(0).Item("ternro"))
                                        insertar = False
                                    Else
                                        'controlo si el empleado tiene periodod de vacaciones para el año
                                        StrSql = " SELECT vacnro FROM vacacion WHERE vacanio = " & Year(CDate(Fechas.convFechaQP(FecHasta, "dd-mm-yyyy").Replace("'", ""))) & " AND vacestado = -1 "
                                        dtVacacion = New DataTable
                                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                        da.Fill(dtVacacion)

                                        If dtVacacion.Rows.Count > 0 Then
                                            PeriodoVac = dtVacacion.Rows(0).Item("vacnro")
                                            insertar = True
                                        Else
                                            FLog.EscribirLinea(" El tercero: " & dtTercero.Rows(0).Item("ternro") & " no posee periodo de vacaciones para el año: " & Year(CDate(Fechas.convFechaQP(FecDesde, "dd-mm-yyyy").Replace("'", ""))))
                                            insertar = False
                                        End If
                                    End If
                                Else
                                    insertar = True
                                End If

                                If insertar Then
                                    'Inserto la Licencia
                                    StrSql = "INSERT INTO emp_lic (empleado,elfechadesde,elfechahasta,tdnro,eldiacompleto,eltipo,elorden "
                                    If Lictipo = 2 Then
                                        StrSql = StrSql & ",elhoradesde,elhorahasta"
                                    End If
                                    StrSql = StrSql & ",elcantdias,licestnro) VALUES ("
                                    StrSql = StrSql & dtTercero.Rows(0).Item("ternro")
                                    StrSql = StrSql & "," & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy")
                                    StrSql = StrSql & "," & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy")
                                    StrSql = StrSql & "," & dtLicencia.Rows(0).Item("tdnro")
                                    StrSql = StrSql & "," & CLng(diaCompleto)
                                    StrSql = StrSql & "," & Lictipo ' 1 = Dia Completo
                                    StrSql = StrSql & ",1" ' 

                                    If Lictipo = 2 Then ' Parcial Fija
                                        StrSql = StrSql & ",'" & horaDesde & "'"
                                        StrSql = StrSql & ",'" & horaHasta & "'"
                                    End If

                                    StrSql = StrSql & "," & (DateDiff("d", CDate(Fechas.convFechaQP(FecDesde, "dd-mm-yyyy").Replace("'", "")), CDate(Fechas.convFechaQP(FecHasta, "dd-mm-yyyy").Replace("'", ""))) + 1)
                                    StrSql = StrSql & "," & LicEstNro
                                    StrSql = StrSql & " )"

                                    conexion.Open()
                                    cmd.CommandText = StrSql
                                    cmd.ExecuteNonQuery()
                                    emp_licnro = confiBasica.getLastIdentity(conexion, "emp_lic")
                                    conexion.Close()

                                    FLog.EscribirLinea("Licencia insertada para el tercero: " & dtTercero.Rows(0).Item("ternro") & "fechas desde: " & MiDataRow.Item("FechaDesde") & " fecha hasta: " & MiDataRow.Item("FechaHasta"))

                                    '________________________________________________
                                    'INSERTO COMPLEMENTOS
                                    '------------------------------------------------
                                    Select Case dtLicencia.Rows(0).Item("tdnro")
                                        Case 2
                                            'Inserto Complemento de vacaciones
                                            StrSql = "INSERT INTO lic_vacacion  (emp_licnro,vacnro,vacnotifnro,licvacmanual) "
                                            StrSql = StrSql & " VALUES ("
                                            StrSql = StrSql & emp_licnro & "," & PeriodoVac & ",NULL,-1)"
                                            conexion.Open()
                                            cmd.CommandText = StrSql
                                            cmd.ExecuteNonQuery()
                                            conexion.Close()

                                            FLog.EscribirLinea("Complemento de Vacaciones Insertado ")
                                            '------------------------------------
                                    End Select

                                    '________________________________________________
                                    'Genero la Justificacion
                                    '------------------------------------------------
                                    StrSql = " INSERT INTO gti_justificacion ( jusanterior,juscodext,jusdesde,jusdiacompleto,jushasta,jussigla,jussistema,ternro,tjusnro,turnro " & _
                                             " ,jushoradesde,jushorahasta,juseltipo,juselmaxhoras,juselorden ) VALUES (" & _
                                             " -1," & emp_licnro & "," & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & _
                                             " ,-1," & Fechas.convFec(MiDataRow.Item("FechaHasta"), "dd-mm-yyyy") & _
                                             " ,'LIC',-1," & dtTercero.Rows(0).Item("ternro") & ",1,0" & _
                                             " ,'" & horaDesde & "','" & horaHasta & "'" & _
                                             " ," & Lictipo & ",0,1)"
                                    conexion.Open()
                                    cmd.CommandText = StrSql
                                    cmd.ExecuteNonQuery()
                                    conexion.Close()
                                    FLog.EscribirLinea("Justificacion insertada ")

                                    '________________________________________________
                                    ' Codigo de Sit. Revista
                                    '------------------------------------------------
                                    FLog.EscribirLinea("Situacion de revista")


                                    StrSql = "SELECT estrnro, tdnro FROM csijp_srtd "
                                    StrSql = StrSql & " WHERE tdnro =" & dtLicencia.Rows(0).Item("tdnro")
                                    dtSitRevista = New DataTable
                                    da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                    da.Fill(dtSitRevista)

                                    If dtSitRevista.Rows.Count > 0 Then
                                        Estrnro_SitRev = dtSitRevista.Rows(0).Item("estrnro")
                                    Else
                                        Estrnro_SitRev = 0
                                    End If

                                    If Estrnro_SitRev <> 0 Then
                                        'Busco el tipo de la situacion de revista anterior
                                        StrSql = " SELECT * FROM his_estructura WHERE tenro = 30 " & _
                                                 " AND ternro = " & dtTercero.Rows(0).Item("ternro") & _
                                                 " AND htetdesde <= " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & _
                                                 " AND ( htethasta >= " & Fechas.convFec(MiDataRow.Item("FechaDesde"), "dd-mm-yyyy") & " OR htethasta IS Null )"
                                        dtHis_Estr = New DataTable
                                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                        da.Fill(dtHis_Estr)

                                        If dtHis_Estr.Rows.Count > 0 Then
                                            'la cierro un dia antes
                                            If IsDBNull(dtHis_Estr.Rows(0).Item("htethasta")) Then
                                                ' If Not (CDate(dtHis_Estr.Rows(0).Item("htetdesde")) = CDate(MiDataRow.Item("FechaDesde"))) Then
                                                If Not (CDate(Fechas.convFechaQP(CDate(dtHis_Estr.Rows(0).Item("htetdesde")), "dd-mm-yyyy").Replace("'", "")) = CDate(FecDesde)) Then

                                                    StrSql = " UPDATE his_estructura SET "
                                                    StrSql = StrSql & " htethasta = " & Fechas.convFechaQP(DateAdd("d", -1, FecDesde))
                                                    StrSql = StrSql & " WHERE tenro   = 30 "
                                                    StrSql = StrSql & " AND   estrnro  = " & dtHis_Estr.Rows(0).Item("estrnro")
                                                    StrSql = StrSql & " AND   ternro  = " & dtTercero.Rows(0).Item("ternro")
                                                    StrSql = StrSql & " AND   htetdesde = " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htetdesde"))
                                                    StrSql = StrSql & " AND   htethasta  is null "
                                                Else
                                                    'la borro porque se va superponer con la licencia
                                                    StrSql = " DELETE his_estructura "
                                                    StrSql = StrSql & " WHERE tenro   = 30 "
                                                    StrSql = StrSql & " AND   estrnro  = " & dtHis_Estr.Rows(0).Item("estrnro")
                                                    StrSql = StrSql & " AND   ternro =" & dtTercero.Rows(0).Item("ternro")
                                                    StrSql = StrSql & " AND   htetdesde = " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htetdesde"))
                                                    StrSql = StrSql & " AND   htethasta  is null "
                                                End If

                                                conexion.Open()
                                                cmd.CommandText = StrSql
                                                cmd.ExecuteNonQuery()
                                                conexion.Close()

                                                'Inserto la misma situacion despues de la nueva situacion (la de la licencia)
                                                StrSql = "INSERT INTO his_estructura(tenro, ternro, estrnro, htetdesde)"
                                                StrSql = StrSql & " VALUES (30, " & dtTercero.Rows(0).Item("ternro") & ", " & dtHis_Estr.Rows(0).Item("estrnro") & ", "
                                                StrSql = StrSql & Fechas.convFechaQP(CDate(FecHasta).AddDays(1)) & ")"

                                                conexion.Open()
                                                cmd.CommandText = StrSql
                                                cmd.ExecuteNonQuery()
                                                conexion.Close()

                                            Else
                                                If CDate(dtHis_Estr.Rows(0).Item("htethasta")) > CDate(MiDataRow.Item("FechaHasta")) Then
                                                    If CDate(dtHis_Estr.Rows(0).Item("htetdesde")) > CDate(MiDataRow.Item("FechaDesde")) Then
                                                        StrSql = " UPDATE his_estructura SET "
                                                        StrSql = StrSql & " htethasta = " & Fechas.convFechaQP(CDate(FecDesde).AddDays(-1))
                                                        StrSql = StrSql & " WHERE tenro   = 30 "
                                                        StrSql = StrSql & " AND   ternro  = " & dtTercero.Rows(0).Item("ternro")
                                                        StrSql = StrSql & " AND   htetdesde = " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htetdesde"))
                                                        StrSql = StrSql & " AND   htethasta  = " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htethasta"))

                                                    Else
                                                        'la borro porque se va superponer con la licencia
                                                        StrSql = " DELETE his_estructura "
                                                        StrSql = StrSql & " WHERE tenro = 30 "
                                                        StrSql = StrSql & " AND   estrnro = " & dtHis_Estr.Rows(0).Item("estrnro")
                                                        StrSql = StrSql & " AND   ternro  = " & dtTercero.Rows(0).Item("ternro")
                                                        StrSql = StrSql & " AND   htetdesde = " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htetdesde"))
                                                        StrSql = StrSql & " AND   htethasta  = " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htethasta"))

                                                    End If

                                                    conexion.Open()
                                                    cmd.CommandText = StrSql
                                                    cmd.ExecuteNonQuery()
                                                    conexion.Close()

                                                    'Inserto la misma situacion despues de la nueva situacion (la de la licencia)
                                                    StrSql = "INSERT INTO his_estructura "
                                                    StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
                                                    StrSql = StrSql & " VALUES (30, " & dtTercero.Rows(0).Item("ternro") & ", " & dtHis_Estr.Rows(0).Item("estrnro") & ", "
                                                    StrSql = StrSql & Fechas.convFechaQP(CDate(FecHasta).AddDays(1)) & ", " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htethasta")) & ")"

                                                    conexion.Open()
                                                    cmd.CommandText = StrSql
                                                    cmd.ExecuteNonQuery()
                                                    conexion.Close()

                                                Else
                                                    If CDate(dtHis_Estr.Rows(0).Item("htetdesde")) > CDate(MiDataRow.Item("FechaDesde")) Then
                                                        StrSql = " UPDATE his_estructura SET "
                                                        StrSql = StrSql & " htethasta = " & Fechas.convFechaQP(CDate(FecDesde).AddDays(-1))
                                                        StrSql = StrSql & " WHERE tenro   = 30 "
                                                        StrSql = StrSql & " AND   ternro  = " & dtTercero.Rows(0).Item("ternro")
                                                        StrSql = StrSql & " AND   htetdesde = " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htetdesde"))
                                                        StrSql = StrSql & " AND   htethasta  is null "

                                                    Else
                                                        'la borro porque se va superponer con la licencia
                                                        StrSql = " DELETE his_estructura "
                                                        StrSql = StrSql & " WHERE tenro = 30 "
                                                        StrSql = StrSql & "     AND estrnro  = " & dtHis_Estr.Rows(0).Item("estrnro")
                                                        StrSql = StrSql & "     AND ternro  = " & dtTercero.Rows(0).Item("ternro")
                                                        StrSql = StrSql & "     AND htetdesde = " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htetdesde"))
                                                        StrSql = StrSql & "     AND htethasta  = " & Fechas.convFechaQP(dtHis_Estr.Rows(0).Item("htethasta"))

                                                    End If

                                                    conexion.Open()
                                                    cmd.CommandText = StrSql
                                                    cmd.ExecuteNonQuery()
                                                    conexion.Close()
                                                End If
                                            End If
                                        End If
                                        StrSql = "INSERT INTO his_estructura(tenro, ternro, estrnro, htetdesde,htethasta) "
                                        StrSql = StrSql & " VALUES (30, " & dtTercero.Rows(0).Item("ternro") & ", " & Estrnro_SitRev & ", "
                                        StrSql = StrSql & Fechas.convFec(MiDataRow.Item("FechaDesde")) & ", " & Fechas.convFec(MiDataRow.Item("FechaHasta")) & ")"

                                        conexion.Open()
                                        cmd.CommandText = StrSql
                                        cmd.ExecuteNonQuery()
                                        conexion.Close()
                                    Else
                                        FLog.EscribirLinea("La Licencia no tienen Situacion de Revista asociado")
                                    End If
                                End If
                                Else
                                    FLog.EscribirLinea("No tiene saldo de vacaciones para insertar la licencia o  no tiene periodo de vacaciones para el año")
                                End If
                        Else
                            FLog.EscribirLinea("Error. El empleado con legajo: " & MiDataRow.Item("Legajo") & " posee licencias en el rango de fechas.")
                        End If

                    Else
                        FLog.EscribirLinea("Error. El empleado con legajo: " & MiDataRow.Item("Legajo") & " no existe en el sistema.")
                    End If
        'End If

                Catch ex As Exception
                    FLog.EscribirLinea("ERROR SQL: " & StrSql)
                    FLog.EscribirLinea("Error al insertar la licencia del legajo " & MiDataRow.Item("Legajo"))
        Finally
            'EAM- Actualiza el avance del proceso
            Progreso = Progreso + IncPorc
            conexion.Open()
            da = New OleDbDataAdapter(StrSql, conexion)
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
            cmd.CommandText = StrSql
            cmd.ExecuteNonQuery()
            conexion.Close()
            FLog.EscribirLinea("----------------------------------------------")
            FLog.EscribirLinea("")
        End Try
            Next
        End If

    End Sub
    Public Sub spec()

        'variables web service
        Dim wsContract As New WebServiceContractClient
        Dim listEmp As New WSElement
        Dim fields() As String
        Dim ws As New WSElement
        Dim empleado As New WSElement
        Dim tarjeta As New WSElement
        Dim agregarProcOnline As Boolean 'mdf
        'variables query
        Dim StrSql3 As String
        Dim StrSql2 As String
        Dim dtDatos As New DataTable
        Dim dtDatosAux2 As New DataTable
        Dim dtDatosAux3 = New DataTable
        Dim dtDatosAux4 = New DataTable

        'variables parametros
        Dim parametros As String
        Dim param
        Dim rhpro_ternro As Integer
        Dim rhpro_nrotarj As String
        Dim operacion As String
        Dim validoDesde As String
        Dim validoHasta As String

        'variables empleado

        Dim rhpro_dni_emp As String
        Dim rhpro_ternom As String
        Dim rhpro_ternom2 As String
        Dim rhpro_terape As String
        Dim rhpro_terape2 As String
        Dim rhpro_fechaAltaEmpresa As String
        Dim nronivel(0) As String
        Dim nivel As String
        Dim empId As Integer

        'variables del empleado de spec
        Dim dni As String
        Dim encontro As Boolean
        Dim SpecEmpId As Integer
        Dim SpecEmpDoc As String
        Dim SpecEmpTarj As String
        Dim elimino As Boolean
        Dim modifico As Boolean
        Dim lista As String
        Dim listaRelojes As String
        Dim permitido As Boolean
        Dim rhpro_legajo As String


        'variables nuevas
        Dim listaTarjetas As WSElement
        Dim empleados As WSElement
        Dim wsCards As WSElement
        Dim idTarj
        Dim empTarj
        Dim tarjetaAnt
        Dim tarjetaNue
        Dim tarjetas
        Dim bpronro

        'variables del progreso
        Dim IncPorc As Single = 0
        Dim Progreso As Single = 0

        'Dim webService
        Dim endpointAddress
        Dim newEndPointAddress
        Dim endpoint
        Dim nroCon As Integer

        'variables Organigrama
        Dim departs As WSElement
        Dim empDep As New WSElement()
        Dim validity As New WSElement()
        Dim dep As New WSElement()
        Dim niveles()
        Dim ndate As Integer
        Dim enddate As Integer
        Dim i As Integer
        Dim Dia As String   '<---------- MDF 25/03/2015

        listaRelojes = ""
        lista = ""


        FLog.EscribirLinea("Ingresa al modelo de spec", 5)
        endpoint = "http://cstest.grupospec.com:8097/WebService"
        'busco en el confrep la configuracion del endpoint
        StrSql = " SELECT * FROM confrep"
        StrSql += " WHERE repnro=421"
        StrSql += " and (conftipo = 'BDO' or conftipo = 'DIA') order by conftipo asc"
        'StrSql += " AND  conftipo='BDO'"
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        dtDatosAux2 = New DataTable
        da.Fill(dtDatosAux2)
        If (dtDatosAux2.Rows.Count > 0) Then
            nroCon = dtDatosAux2.Rows(0).Item("confval")
            If (dtDatosAux2.Rows.Count > 1) Then '------MDF
                Dia = dtDatosAux2.Rows(1).Item("confval")
                FLog.EscribirLinea("Dia configurado en confrep,  0 = lee a partir del dia anterior, -1= solo el dia", 5)
            Else
                FLog.EscribirLinea("Dia no configurado en confrep, se leera a partir del dia anterior", 5)
                Dia = "0"
            End If '------MDF
            StrSql = " SELECT * FROM conexion "
            StrSql += " WHERE cnnro=" & nroCon
            da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
            dtDatosAux2 = New DataTable
            da.Fill(dtDatosAux2)
            If (dtDatosAux2.Rows.Count > 0) Then
                endpoint = dtDatosAux2.Rows(0).Item("cnstring")
            End If
        End If

        endpointAddress = wsContract.Endpoint.Address
        newEndPointAddress = New EndpointAddressBuilder(endpointAddress)
        newEndPointAddress.uri = New Uri(endpoint)
        wsContract = New WebServiceContractClient("WSHttpBinding_IWebServiceContract", newEndPointAddress.ToEndPointAddress().ToString)


        encontro = False
        fields = {"id", "name", "Number", "cards"}
        FLog.EscribirLinea("Se empieza a procesar la interfaz spec", 5)
        fechaAltaSist = Format(Now(), "yyyy/MM/dd")
        FLog.EscribirLinea("Fecha de alta en el sistema: " & fechaAltaSist, 5)
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)
        If (dtDatos.Rows.Count > 0) Then
            'Levanto los parametros, de aca dedusco si es un alta/baja/mod o alta masiva de empleados
            'nro de tercero, nro tarjeta, operacion, fechadesde, fechahasta
            parametros = dtDatos.Rows(0).Item("bprcparam").ToString
            param = Split(parametros, "@")
            If UBound(param) > 3 Then  'aca decia > 1 ---- mdf
                IncPorc = 99 / 1
                FLog.EscribirLinea("La cantidad de parametros es mayor a 1, Vienen los datos del trigger", 5)
                rhpro_ternro = param(4) '0
                FLog.EscribirLinea("Numero de tercero en RHPro: " & rhpro_ternro, 5)
                rhpro_nrotarj = param(5) '1
                FLog.EscribirLinea("Numero de tarjeta de RHPro: " & rhpro_nrotarj, 5)
                operacion = param(6) '2
                FLog.EscribirLinea("Tipo de oparacion: " & operacion, 5)
                validoDesde = param(7) '3
                FLog.EscribirLinea("Tarjeta valida desde:" & validoDesde, 5)
                validoHasta = param(8) '4
                FLog.EscribirLinea("Tarjeta valida Hasta:" & validoHasta, 5)

                'BUSCO EL NUMERO DE TERCERO DEL EMPLEADO
                'busco con el numero de doc del empleado si existe en spec y si existe y la tarjeta coincide lo elimino
                StrSql2 = " SELECT * FROM ter_doc "
                StrSql2 += " INNER JOIN  tercero ON tercero.ternro = ter_doc.ternro "
                StrSql2 += " INNER JOIN empleado ON empleado.ternro = tercero.ternro "
                StrSql2 += " WHERE tercero.ternro=" & rhpro_ternro
                StrSql2 += " AND ter_doc.tidnro <=5 ORDER BY tidnro ASC "
                FLog.EscribirLinea("Se busca los datos del tercero, el mismo debe tener documento tipo menor a 5", 5)
                FLog.EscribirLinea(StrSql2, 5)
                da = New OleDbDataAdapter(StrSql2, conexion.ConnectionString)
                dtDatosAux2 = New DataTable
                da.Fill(dtDatosAux2)
                If (dtDatosAux2.Rows.Count > 0) Then
                    FLog.EscribirLinea("Se encontro un tercero en RHPro", 5)
                    rhpro_dni_emp = dtDatosAux2.Rows(0).Item("nrodoc").ToString()
                    FLog.EscribirLinea("Dni del tercero: " & rhpro_dni_emp, 5)
                    rhpro_ternom = dtDatosAux2.Rows(0).Item("ternom").ToString()
                    rhpro_ternom2 = dtDatosAux2.Rows(0).Item("ternom2").ToString()
                    rhpro_terape = dtDatosAux2.Rows(0).Item("terape").ToString()
                    rhpro_terape2 = dtDatosAux2.Rows(0).Item("terape2").ToString()
                    rhpro_legajo = dtDatosAux2.Rows(0).Item("empleg").ToString
                    If rhpro_ternom2 <> "" Then
                        rhpro_ternom = rhpro_ternom & " " & rhpro_ternom2
                    End If
                    FLog.EscribirLinea("Nombre: " & rhpro_ternom)
                    If rhpro_terape2 <> "" Then
                        rhpro_terape = rhpro_terape & " " & rhpro_terape2
                    End If
                    FLog.EscribirLinea("Apellido: " & rhpro_terape)
                    FLog.EscribirLinea("El tercero encontrado es: " & rhpro_ternom & ", " & rhpro_terape & "Doc: " & rhpro_dni_emp)
                Else
                    FLog.EscribirLinea("No se encontro ningun tercero en RHPro, se aborta el proceso", 5)
                    Exit Sub
                End If

                'Busco la empresa del tercero
                StrSql3 = " SELECT estrdabr FROM his_estructura "
                StrSql3 += " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
                StrSql3 += " WHERE his_estructura.ternro =" & rhpro_ternro
                StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                StrSql3 += " AND his_estructura.tenro=10"
                da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                dtDatosAux2 = New DataTable
                da.Fill(dtDatosAux2)
                FLog.EscribirLinea("Busco el nombre de la empresa (TE 10 en RHPro):" & StrSql3, 5)
                If (dtDatosAux2.Rows.Count > 0) Then
                    rhpro_nombreEmpresa = dtDatosAux2.Rows(0).Item("estrdabr").ToString
                Else
                    rhpro_nombreEmpresa = ""
                End If
                'HASTA ACA

                'LA OPERACION VA A SER SOBRE UN EMPLEADO EN PARTICULAR 29/12/2013
                Select Case operacion
                    Case "B" 'La operacion es una baja
                        elimino = False
                        'BUSCO CON EL DNI SI EL EMPLEADO ESTA EN SPEC
                        listEmp = wsContract.ListFields(WSContainer.Employee, fields, "this.name=""" + rhpro_dni_emp + """")
                        If listEmp.Data.Count = 0 Then
                            FLog.EscribirLinea("No hay ningun empleado con el doc " & rhpro_dni_emp & "en Spec")
                        Else
                            empId = -1
                            For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                empleado = pair.Value
                                SpecEmpId = empleado.Data("id")
                                SpecEmpDoc = empleado.Data("name")
                                'obtengo las tarjetas del empleado
                                listaTarjetas = empleado.Data("Cards")
                                For Each pair2 As KeyValuePair(Of String, Object) In listaTarjetas.Data
                                    tarjeta = wsContract.Get(WSContainer.Card, pair2.Value)
                                    SpecEmpTarj = tarjeta.Data("Number")
                                    If rhpro_nrotarj = SpecEmpTarj Then
                                        idTarj = tarjeta.Data("id")
                                        'elimino la tarjeta
                                        wsContract.Delete(WSContainer.Card, idTarj)
                                        FLog.EscribirLinea("Se elimino la tarjeta numero :" & idTarj)
                                        elimino = True
                                        Exit For
                                    End If
                                Next
                                'hasta aca
                            Next
                            If Not elimino Then
                                'la tarjeta no se pudo eliminar
                                FLog.EscribirLinea("La tarjeta " & rhpro_nrotarj & " no se pudo eliminar porque no se encontro en Spec.", 5)
                            End If
                        End If
                        'HASTA ACA

                    Case "A", "M"
                        'Valido si en spec ya hay empleados insertados
                        listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                        If (listEmp.Data.Count = 0) Then 'no hay empleados en spec aun
                            FLog.EscribirLinea("No hay empleados cargado en Spec, se inserta el empleado", 5)
                            'Busco los datos basicos a ingresar del empleado
                            'con el numero de tercero busco la fecha de alta en la empresa
                            StrSql3 = " SELECT * FROM his_estructura "
                            StrSql3 += " WHERE ternro =" & rhpro_ternro
                            StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                            StrSql3 += " AND tenro=10"
                            da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                            dtDatosAux2 = New DataTable
                            da.Fill(dtDatosAux2)
                            FLog.EscribirLinea("Busco la fecha de ingreso a la empresa (TE 10 en RHPro):" & StrSql3, 5)
                            If (dtDatosAux2.Rows.Count > 0) Then
                                rhpro_fechaAltaEmpresa = dtDatosAux2.Rows(0).Item("htetdesde").ToString
                                rhpro_fechaAltaEmpresa = Format(dtDatosAux2.Rows(0).Item("htetdesde"), "yyyy-MM-dd")
                                FLog.EscribirLinea("Se encontro la fecha de ingreso en la empresa: " & rhpro_fechaAltaEmpresa, 5)
                            Else
                                rhpro_fechaAltaEmpresa = ""
                                FLog.EscribirLinea("No se encontro la fecha de ingreso a la empresa", 5)
                            End If
                            'hasta aca

                            'busco los niveles en el organigrama
                            StrSql3 = "SELECT * FROM confrep "
                            StrSql3 += "WHERE repnro=421 "
                            StrSql3 += "AND conftipo='TE'"
                            StrSql3 += "ORDER BY confnrocol ASC "
                            da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                            dtDatosAux2 = New DataTable
                            da.Fill(dtDatosAux2)
                            For i = 0 To dtDatosAux2.Rows.Count - 1
                                ReDim Preserve nronivel(dtDatosAux2.Rows(i).Item("confnrocol"))
                                nronivel(dtDatosAux2.Rows(i).Item("confnrocol")) = dtDatosAux2.Rows(i).Item("confval").ToString()
                            Next
                            'hasta aca

                            'para cada uno de los niveles busco la descripcion
                            For j = 1 To UBound(nronivel)
                                StrSql3 = " SELECT * FROM his_estructura "
                                StrSql3 += " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                                StrSql3 += " WHERE ternro =" & rhpro_ternro
                                StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                                StrSql3 += " AND estructura.tenro=" & nronivel(j)
                                FLog.EscribirLinea("para cada nivel busca la descripcion:" & StrSql3, 5)
                                da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                                dtDatosAux2 = New DataTable
                                da.Fill(dtDatosAux2)
                                FLog.EscribirLinea("Busco la descripcion del nivel  (TE: " & nronivel(j) & " en RHPro):" & StrSql3, 5)
                                If (dtDatosAux2.Rows.Count > 0) Then
                                    nivel += dtDatosAux2.Rows(0).Item("estrdabr").ToString & "/"
                                    FLog.EscribirLinea("Nivel: " & j & "Descripcion Nivel: " & dtDatosAux2.Rows(0).Item("estrdabr").ToString, 5)
                                Else
                                    nivel = ""
                                    FLog.EscribirLinea("No se la descripcion del nivel ", 5)
                                End If
                            Next
                            'inserto el empleado
                            ws = wsContract.Get(WSContainer.Employee, -1)
                            ws.Data("name") = rhpro_dni_emp
                            ws.Data("nameEmployee") = rhpro_ternom
                            ws.Data("LastName") = rhpro_terape
                            ws.Data("REGISTERSYSTEMDATE") = fechaAltaSist
                            ws.Data("ACTIVEDAYS") = rhpro_fechaAltaEmpresa
                            ws.Data("DEPARTAMENTS") = nivel
                            ws.Data("employeeCode") = rhpro_legajo
                            ws.Data("companyCode") = rhpro_nombreEmpresa
                            '------------------------------------------
                            'armo el arbol organizacional
                            'Get all deparments
                            niveles = Split(nivel, "/")
                            i = 0
                            empDep.Data = New Dictionary(Of String, Object)
                            FLog.EscribirLinea("cantidad de niveles: " & UBound(niveles) - 1, 5)
                            For k = 0 To UBound(niveles) - 1
                                departs = wsContract.ListFields(WSContainer.StructureTree, fields, "this.name=""" + niveles(k) + """")
                                'Add department
                                'empDep.Data = New Dictionary(Of String, Object)
                                'i = 0

                                For Each pair3 As KeyValuePair(Of String, Object) In departs.Data
                                    'Create validity
                                    FLog.EscribirLinea("entro al for pair3", 5)
                                    'validity.Data = New Dictionary(Of String, Object)
                                    'ndate = 2013
                                    'enddate = 2016
                                    'Do While (ndate < enddate)
                                    '    FLog.EscribirLinea("ndate: " & ndate.ToString(), 5)
                                    '    validity.Data.Add(ndate.ToString(), ndate)
                                    '    ndate = ndate + 1                                        
                                    'Loop
                                    dep = pair3.Value

                                    'dep.Data.Add("validity", validity)
                                    'Add dep to employee
                                    i = i + 1

                                    empDep.Data.Add((i).ToString(), dep)
                                    FLog.EscribirLinea("despues de agregar depto ", 5)

                                    'Exit For
                                Next
                            Next
                            ws.Data("Departments") = empDep
                            wsContract.Set(WSContainer.Employee, ws)

                            'busco el id del empleado
                            listEmp = wsContract.ListFields(WSContainer.Employee, fields, "this.name=""" + rhpro_dni_emp + """")
                            For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                empleado = pair.Value
                                empId = empleado.Data("id")
                            Next

                            tarjeta = wsContract.Get(WSContainer.Card, -1)
                            tarjeta.Data("Number") = rhpro_nrotarj
                            tarjeta.Data("employee") = empId
                            wsContract.Set(WSContainer.Card, tarjeta)

                            ws = wsContract.Get(WSContainer.Employee, rhpro_ternro)

                            ws.Data("cards") = rhpro_nrotarj
                            ws.Data("cards_dateini") = validoDesde
                            ws.Data("cards_dateend") = validoHasta
                            wsContract.Set(WSContainer.Employee, ws)
                            FLog.EscribirLinea("Se inserta el empleado " & rhpro_terape & " con id " & rhpro_ternro, 5)
                            'hasta aca
                        Else 'ya hay empleados en spec
                            If operacion = "A" Then
                                FLog.EscribirLinea("La operacion es un ALTA, se busco si existe un empleado con mismo doc.", 5)
                                listEmp = wsContract.ListFields(WSContainer.Employee, fields, "this.name=""" + rhpro_dni_emp + """")
                                If listEmp.Data.Count = 0 Then
                                    FLog.EscribirLinea("No hay ningun empleado con el doc " & rhpro_dni_emp & "en Spec, se inserta un nuevo empleado")

                                    'con el numero de tercero busco la fecha de alta en la empresa
                                    StrSql3 = " SELECT * FROM his_estructura "
                                    StrSql3 += " WHERE ternro =" & rhpro_ternro
                                    StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                                    StrSql3 += " AND tenro=10"
                                    da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                                    dtDatosAux2 = New DataTable
                                    da.Fill(dtDatosAux2)
                                    FLog.EscribirLinea("Busco la fecha de ingreso a la empresa (TE 10 en RHPro):" & StrSql3, 5)
                                    If (dtDatosAux2.Rows.Count > 0) Then
                                        rhpro_fechaAltaEmpresa = dtDatosAux2.Rows(0).Item("htetdesde").ToString
                                        rhpro_fechaAltaEmpresa = Format(dtDatosAux2.Rows(0).Item("htetdesde"), "yyyy-MM-dd")
                                        FLog.EscribirLinea("Se encontro la fecha de ingreso en la empresa: " & rhpro_fechaAltaEmpresa, 5)
                                    Else
                                        rhpro_fechaAltaEmpresa = ""
                                        FLog.EscribirLinea("No se encontro la fecha de ingreso a la empresa", 5)
                                    End If
                                    'hasta aca

                                    'busco los niveles en el organigrama
                                    StrSql3 = "SELECT * FROM confrep "
                                    StrSql3 += "WHERE repnro=421 "
                                    StrSql3 += "AND conftipo='TE'"
                                    StrSql3 += "ORDER BY confnrocol ASC "
                                    da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                                    dtDatosAux2 = New DataTable
                                    da.Fill(dtDatosAux2)
                                    For i = 0 To dtDatosAux2.Rows.Count - 1
                                        ReDim Preserve nronivel(dtDatosAux2.Rows(i).Item("confnrocol"))
                                        nronivel(dtDatosAux2.Rows(i).Item("confnrocol")) = dtDatosAux2.Rows(i).Item("confval").ToString()
                                    Next
                                    'hasta aca

                                    'para cada uno de los niveles busco la descripcion
                                    For j = 1 To UBound(nronivel)
                                        StrSql3 = " SELECT * FROM his_estructura "
                                        StrSql3 += " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                                        StrSql3 += " WHERE ternro =" & rhpro_ternro
                                        StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                                        StrSql3 += " AND estructura.tenro=" & nronivel(j)
                                        FLog.EscribirLinea("para cada nivel busca la descripcion:" & StrSql3, 5)
                                        da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                                        dtDatosAux2 = New DataTable
                                        da.Fill(dtDatosAux2)
                                        FLog.EscribirLinea("Busco la descripcion del nivel  (TE: " & nronivel(j) & " en RHPro):" & StrSql3, 5)
                                        If (dtDatosAux2.Rows.Count > 0) Then
                                            nivel += dtDatosAux2.Rows(0).Item("estrdabr").ToString & "/"
                                            FLog.EscribirLinea("Nivel: " & j & "Descripcion Nivel: " & dtDatosAux2.Rows(0).Item("estrdabr").ToString, 5)
                                        Else
                                            nivel = ""
                                            FLog.EscribirLinea("No se la descripcion del nivel ", 5)
                                        End If
                                    Next
                                    'inserto el empleado
                                    ws = wsContract.Get(WSContainer.Employee, -1)
                                    ws.Data("name") = rhpro_dni_emp
                                    ws.Data("nameEmployee") = rhpro_ternom
                                    ws.Data("LastName") = rhpro_terape
                                    ws.Data("REGISTERSYSTEMDATE") = fechaAltaSist
                                    ws.Data("ACTIVEDAYS") = rhpro_fechaAltaEmpresa
                                    'ws.Data("DEPARTAMENTS") = nivel
                                    ws.Data("employeeCode") = rhpro_legajo
                                    ws.Data("companyCode") = rhpro_nombreEmpresa
                                    '------------------------------------------
                                    'armo el arbol organizacional
                                    'Get all deparments
                                    niveles = Split(nivel, "/")
                                    i = 0
                                    empDep.Data = New Dictionary(Of String, Object)
                                    FLog.EscribirLinea("cantidad de niveles: " & UBound(niveles) - 1, 5)
                                    For k = 0 To UBound(niveles) - 1
                                        departs = wsContract.ListFields(WSContainer.StructureTree, fields, "this.name=""" + niveles(k) + """")
                                        'Add department
                                        'empDep.Data = New Dictionary(Of String, Object)
                                        'i = 0

                                        For Each pair3 As KeyValuePair(Of String, Object) In departs.Data
                                            'Create validity
                                            FLog.EscribirLinea("entro al for pair3", 5)
                                            'validity.Data = New Dictionary(Of String, Object)
                                            'ndate = 2013
                                            'enddate = 2016
                                            'Do While (ndate < enddate)
                                            '    FLog.EscribirLinea("ndate: " & ndate.ToString(), 5)
                                            '    validity.Data.Add(ndate.ToString(), ndate)
                                            '    ndate = ndate + 1                                        
                                            'Loop
                                            dep = pair3.Value

                                            'dep.Data.Add("validity", validity)
                                            'Add dep to employee
                                            i = i + 1

                                            empDep.Data.Add((i).ToString(), dep)
                                            FLog.EscribirLinea("despues de agregar depto ", 5)

                                            'Exit For
                                        Next
                                    Next
                                    ws.Data("Departments") = empDep
                                    wsContract.Set(WSContainer.Employee, ws)
                                    '------------------------------------------
                                    FLog.EscribirLinea("Se inserto el empleado " & rhpro_ternom & ", " & rhpro_terape, 5)
                                    'busco si la tarjeta existe en Spec
                                    tarjeta = wsContract.ListFields(WSContainer.Card, fields, "")
                                    If tarjeta.Data.Count > 0 Then
                                        'armo una lista de tarjetas
                                        For Each pair As KeyValuePair(Of String, Object) In tarjeta.Data
                                            wsCards = pair.Value
                                            lista = lista + "'" + wsCards.Data("Number") + "'"
                                        Next
                                        Dim x As Integer
                                        x = lista.IndexOf("'" & rhpro_nrotarj & "'")
                                        If x <> -1 Then
                                            FLog.EscribirLinea("La tarjeta :" & rhpro_nrotarj & "ya existe en Spec, no se le inserta al empleado")
                                        Else
                                            'busco el id del empleado
                                            listEmp = wsContract.ListFields(WSContainer.Employee, fields, "this.name=""" + rhpro_dni_emp + """")
                                            For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                                empleado = pair.Value
                                                empId = empleado.Data("id")
                                            Next
                                            tarjeta = wsContract.Get(WSContainer.Card, -1)
                                            tarjeta.Data("Number") = rhpro_nrotarj
                                            tarjeta.Data("employee") = empId
                                            wsContract.Set(WSContainer.Card, tarjeta)
                                            ws = wsContract.Get(WSContainer.Employee, empId)
                                            ws.Data("cards") = rhpro_nrotarj
                                            ws.Data("cards_dateini") = validoDesde
                                            ws.Data("cards_dateend") = validoHasta
                                            wsContract.Set(WSContainer.Employee, ws)
                                            FLog.EscribirLinea("Se inserta la tarjeta " & rhpro_nrotarj & " al empleado " & rhpro_terape, 5)
                                            'hasta aca
                                        End If
                                    End If
                                    'hasta aca
                                Else
                                    FLog.EscribirLinea("Ya existe un empleado en Spec con Doc:" & rhpro_dni_emp & " no se puede insertar")
                                    For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                        empleado = pair.Value
                                        empId = empleado.Data("id")
                                    Next
                                    'se busca si existe una tarjeta con el mismo numero
                                    tarjeta = wsContract.ListFields(WSContainer.Card, fields, "")
                                    If tarjeta.Data.Count > 0 Then
                                        'armo una lista de tarjetas
                                        For Each pair As KeyValuePair(Of String, Object) In tarjeta.Data
                                            wsCards = pair.Value
                                            lista = lista + "'" + wsCards.Data("Number") + "'"
                                        Next
                                        Dim x As Integer
                                        x = lista.IndexOf("'" & rhpro_nrotarj & "'")
                                        If x <> -1 Then
                                            FLog.EscribirLinea("La tarjeta :" & rhpro_nrotarj & "ya existe en Spec, no se le inserta al empleado")
                                        Else
                                            tarjeta = wsContract.Get(WSContainer.Card, -1)
                                            tarjeta.Data("Number") = rhpro_nrotarj
                                            tarjeta.Data("employee") = empId
                                            wsContract.Set(WSContainer.Card, tarjeta)

                                            ws = wsContract.Get(WSContainer.Employee, empId)
                                            ws.Data("cards") = rhpro_nrotarj
                                            ws.Data("cards_dateini") = validoDesde
                                            ws.Data("cards_dateend") = validoHasta
                                            wsContract.Set(WSContainer.Employee, ws)
                                            FLog.EscribirLinea("Se inserta la tarjeta " & rhpro_nrotarj & " al empleado " & rhpro_terape, 5)
                                            'hasta aca
                                        End If
                                    End If
                                    'hasta aca
                                End If
                            Else
                                If operacion = "M" Then
                                    FLog.EscribirLinea("La operacion es una modificacion", 5)
                                    listEmp = wsContract.ListFields(WSContainer.Employee, fields, "this.name=""" + rhpro_dni_emp + """")
                                    'busco los niveles en el organigrama
                                    StrSql3 = "SELECT * FROM confrep "
                                    StrSql3 += "WHERE repnro=421 "
                                    StrSql3 += "AND conftipo='TE'"
                                    StrSql3 += "ORDER BY confnrocol ASC "
                                    da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                                    dtDatosAux2 = New DataTable
                                    da.Fill(dtDatosAux2)
                                    For i = 0 To dtDatosAux2.Rows.Count - 1
                                        ReDim Preserve nronivel(dtDatosAux2.Rows(i).Item("confnrocol"))
                                        nronivel(dtDatosAux2.Rows(i).Item("confnrocol")) = dtDatosAux2.Rows(i).Item("confval").ToString()
                                    Next
                                    'hasta aca

                                    'para cada uno de los niveles busco la descripcion
                                    For j = 1 To UBound(nronivel)
                                        StrSql3 = " SELECT * FROM his_estructura "
                                        StrSql3 += " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                                        StrSql3 += " WHERE ternro =" & rhpro_ternro
                                        StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                                        StrSql3 += " AND estructura.tenro=" & nronivel(j)
                                        FLog.EscribirLinea("para cada nivel busca la descripcion:" & StrSql3, 5)
                                        da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                                        dtDatosAux2 = New DataTable
                                        da.Fill(dtDatosAux2)
                                        FLog.EscribirLinea("Busco la descripcion del nivel  (TE: " & nronivel(j) & " en RHPro):" & StrSql3, 5)
                                        If (dtDatosAux2.Rows.Count > 0) Then
                                            nivel += dtDatosAux2.Rows(0).Item("estrdabr").ToString & "/"
                                            FLog.EscribirLinea("Nivel: " & j & "Descripcion Nivel: " & dtDatosAux2.Rows(0).Item("estrdabr").ToString, 5)
                                        Else
                                            nivel = ""
                                            FLog.EscribirLinea("No se la descripcion del nivel ", 5)
                                        End If
                                    Next
                                    If listEmp.Data.Count = 0 Then
                                        FLog.EscribirLinea("No hay ningun empleado con el doc " & rhpro_dni_emp & "en Spec")
                                        'aca podria insertar el empleado
                                    Else
                                        empId = -1
                                        tarjetas = Split(rhpro_nrotarj, ",")
                                        tarjetaNue = tarjetas(0)
                                        tarjetaAnt = tarjetas(1)
                                        For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                            empleado = pair.Value
                                            SpecEmpId = empleado.Data("id")
                                            SpecEmpDoc = empleado.Data("name")
                                            'obtengo las tarjetas del empleado
                                            listaTarjetas = empleado.Data("Cards")
                                            modifico = False
                                            For Each pair2 As KeyValuePair(Of String, Object) In listaTarjetas.Data
                                                tarjeta = wsContract.Get(WSContainer.Card, pair2.Value)
                                                SpecEmpTarj = tarjeta.Data("Number")
                                                If tarjetaAnt = SpecEmpTarj Then
                                                    idTarj = tarjeta.Data("id")
                                                    'modifico la tarjeta
                                                    tarjeta = wsContract.Get(WSContainer.Card, idTarj)
                                                    tarjeta.Data("Number") = tarjetaNue
                                                    tarjeta.Data("employee") = SpecEmpId
                                                    wsContract.Set(WSContainer.Card, tarjeta)

                                                    ws = wsContract.Get(WSContainer.Employee, SpecEmpId)
                                                    ws.Data("cards") = tarjetaNue
                                                    ws.Data("cards_dateini") = validoDesde
                                                    ws.Data("cards_dateend") = validoHasta
                                                    'ws.Data("DEPARTAMENTS") = nivel
                                                    ws.Data("employeeCode") = rhpro_legajo
                                                    ws.Data("companyCode") = rhpro_nombreEmpresa

                                                    '------------------------------------------
                                                    'armo el arbol organizacional

                                                    niveles = Split(nivel, "/")
                                                    i = 0
                                                    empDep.Data = New Dictionary(Of String, Object)
                                                    FLog.EscribirLinea("cantidad de niveles: " & UBound(niveles) - 1, 5)
                                                    For k = 0 To UBound(niveles) - 1
                                                        departs = wsContract.ListFields(WSContainer.StructureTree, fields, "this.name=""" + niveles(k) + """")
                                                        'Add department
                                                        'empDep.Data = New Dictionary(Of String, Object)
                                                        'i = 0

                                                        For Each pair3 As KeyValuePair(Of String, Object) In departs.Data
                                                            'Create validity
                                                            FLog.EscribirLinea("entro al for pair3", 5)
                                                            'validity.Data = New Dictionary(Of String, Object)
                                                            'ndate = 2013
                                                            'enddate = 2016
                                                            'Do While (ndate < enddate)
                                                            '    FLog.EscribirLinea("ndate: " & ndate.ToString(), 5)
                                                            '    validity.Data.Add(ndate.ToString(), ndate)
                                                            '    ndate = ndate + 1                                        
                                                            'Loop
                                                            dep = pair3.Value

                                                            'dep.Data.Add("validity", validity)
                                                            'Add dep to employee
                                                            i = i + 1

                                                            empDep.Data.Add((i).ToString(), dep)
                                                            FLog.EscribirLinea("despues de agregar depto ", 5)

                                                            'Exit For
                                                        Next
                                                    Next
                                                    ws.Data("Departments") = empDep
                                                    wsContract.Set(WSContainer.Employee, ws)
                                                    '------------------------------------------

                                                    'wsContract.Set(WSContainer.Employee, ws)
                                                    modifico = True
                                                    FLog.EscribirLinea("Se modifico la tarjeta " & tarjetaAnt, 5)
                                                    Exit For
                                                End If
                                            Next
                                            'hasta aca
                                        Next
                                        If Not modifico Then
                                            'la tarjeta no se pudo eliminar
                                            FLog.EscribirLinea("La tarjeta " & tarjetaAnt & " no se pudo modificar porque la tarjeta no pertenece al empleado.", 5)
                                            'se busca si existe una tarjeta con el mismo numero
                                            tarjeta = wsContract.ListFields(WSContainer.Card, fields, "")
                                            If tarjeta.Data.Count > 0 Then
                                                'armo una lista de tarjetas
                                                For Each pair As KeyValuePair(Of String, Object) In tarjeta.Data
                                                    wsCards = pair.Value
                                                    lista = lista + "'" + wsCards.Data("Number") + "'"
                                                Next
                                                Dim x As Integer
                                                x = lista.IndexOf("'" & tarjetaNue & "'")
                                                If x <> -1 Then
                                                    FLog.EscribirLinea("La tarjeta :" & tarjetaNue & "ya existe en Spec, no se le inserta al empleado")
                                                Else
                                                    tarjeta = wsContract.Get(WSContainer.Card, -1)
                                                    tarjeta.Data("Number") = tarjetaNue
                                                    tarjeta.Data("employee") = SpecEmpId
                                                    wsContract.Set(WSContainer.Card, tarjeta)

                                                    ws = wsContract.Get(WSContainer.Employee, SpecEmpId)
                                                    ws.Data("cards") = tarjetaNue
                                                    ws.Data("cards_dateini") = validoDesde
                                                    ws.Data("cards_dateend") = validoHasta
                                                    ws.Data("employeeCode") = rhpro_legajo
                                                    wsContract.Set(WSContainer.Employee, ws)
                                                    FLog.EscribirLinea("Se inserta la tarjeta " & tarjetaNue & " al empleado " & rhpro_terape, 5)
                                                    'hasta aca
                                                End If
                                            End If

                                        End If
                                    End If
                                End If
                            End If
                        End If
                End Select
                'HASTA ACA
                'ACA ACTUALIZO EL PORCENTAJE DE LAS TARJETAS 25/04/2014
                conexion.Open()
                da = New OleDbDataAdapter(StrSql, conexion)
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
                cmd.CommandText = StrSql
                cmd.ExecuteNonQuery()
                conexion.Close()
                'HASTA ACA
            Else    '----mdf 
                If UBound(param) = 0 Then
                    If parametros.Length > 0 Then
                        'es carga masiva 
                        FLog.EscribirLinea("CARGA MASIVA DE TARJETAS")
                        bpronro = param(0)
                        Call cargaMasiva(bpronro)
                    End If
                    'hasta aca
                    'End If
                Else
                    '--------------------------------------------------------------------
                    'aca se hace la lectura de registraciones para todos los empleados
                    Dim FDesde, FHasta As String
                    Dim fechadesde As Date
                    Dim fechaHasta As Date
                    Dim Fechas As New Fechas
                    Dim clock As New ClockingType
                    Dim Clockings() As Clocking
                    Dim WsLectura As New WebServiceContractClient
                    Dim employee As New WSElement
                    Dim reloj As String
                    Dim dtDatosAux As DataTable
                    Dim empDni As String
                    Dim StrSqlEmp As String
                    Dim cantEmpleados As Long
                    'endpoint = "http://200.61.13.22:8091/WebService"
                    'webService = New WebServiceContractClient()

                    FLog.EscribirLinea("levanta los datos del endpoint", 5)
                    StrSql = " SELECT * FROM confrep"
                    StrSql += " WHERE repnro=421"
                    StrSql += " AND  conftipo='BDO'"
                    da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                    dtDatosAux2 = New DataTable
                    da.Fill(dtDatosAux2)
                    If (dtDatosAux2.Rows.Count > 0) Then
                        nroCon = dtDatosAux2.Rows(0).Item("confval")
                        StrSql = " SELECT * FROM conexion "
                        StrSql += " WHERE cnnro=" & nroCon
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        dtDatosAux2 = New DataTable
                        da.Fill(dtDatosAux2)
                        If (dtDatosAux2.Rows.Count > 0) Then
                            endpoint = dtDatosAux2.Rows(0).Item("cnstring")
                        End If
                    End If
                    FLog.EscribirLinea("endpoint:" & endpoint, 5)
                    endpointAddress = WsLectura.Endpoint.Address
                    newEndPointAddress = New EndpointAddressBuilder(endpointAddress)
                    newEndPointAddress.uri = New Uri(endpoint)

                    'armo una lista de los relojes permitidos '30/12/2013
                    StrSql3 = "SELECT relcodext FROM gti_reloj "
                    da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                    dtDatosAux2 = New DataTable
                    da.Fill(dtDatosAux2)
                    For i = 0 To dtDatosAux2.Rows.Count - 1
                        listaRelojes = listaRelojes & ",'" & dtDatosAux2.Rows(i).Item("relcodext").ToString() & "'"
                    Next
                    'hasta aca

                    WsLectura = New WebServiceContractClient("WSHttpBinding_IWebServiceContract", newEndPointAddress.ToEndPointAddress().ToString)
                    fields = {"name", "id"}
                    reloj = ""
                    'FDesde = Fechas.cambiaFecha(DateAdd(DateInterval.Day, -1, CDate(dtDatos.Rows(0).Item("bprcfecdesde").ToString)))
                    If Dia = "0" Then   'MDF raffo solo quiere procesar un dia
                        fechadesde = DateAdd(DateInterval.Day, -1, CDate(dtDatos.Rows(0).Item("bprcfecdesde").ToString))
                    Else
                        'fechadesde = CDate(dtDatos.Rows(0).Item("bprcfecdesde").ToString)
                        fechadesde = dtDatos.Rows(0).Item("bprcfecdesde").ToString
                    End If

                    'FHasta = Fechas.cambiaFecha(dtDatos.Rows(0).Item("bprcfechasta").ToString)
                    fechaHasta = dtDatos.Rows(0).Item("bprcfechasta").ToString

                    FLog.EscribirLinea("la fecha desde es:" & fechadesde) 'mdf
                    FLog.EscribirLinea("la fecha Hasta es:" & fechaHasta) 'mdf
                    '-------------------------------
                    g_fdesde = fechadesde
                    g_fhasta = fechaHasta
                    '-------------------------------

                    If fechaHasta = fechadesde Then
                        FLog.EscribirLinea("La fecha desde y la fecha hasta son iguales.")
                        fechaHasta = DateAdd(DateInterval.Day, 1, CDate(dtDatos.Rows(0).Item("bprcfechasta").ToString))
                    End If

                    FDesde = Format(fechadesde, "yyyy-MM-dd")
                    FHasta = Format(fechaHasta, "yyyy-MM-dd")
                    clock = ClockingType.Access
                    'para c/u de los empleados busco las registraciones
                    employee = WsLectura.ListFields(WSContainer.Employee, fields, "")
                    cantEmpleados = employee.Data.Count()
                    Progreso = 0
                    If cantEmpleados <> 0 Then
                        IncPorc = 100 / cantEmpleados
                    Else
                        cantEmpleados = 1
                        IncPorc = 100 / cantEmpleados
                    End If
                    For Each pair As KeyValuePair(Of String, Object) In employee.Data
                        'busco las registraciones para cada uno de los empleados
                        Progreso = Progreso + IncPorc
                        employee = pair.Value
                        empId = employee.Data("id")
                        empDni = employee.Data("name")
                        Clockings = WsLectura.Clockings(empId, FDesde, FHasta, clock)
                        FLog.EscribirLinea("Emp DNI:" & empDni)
                        For j = 0 To UBound(Clockings)
                            Try

                                FLog.EscribirLinea("idReader:" & Clockings(j).IdReader, 5)
                                FLog.EscribirLinea("idClocking:" & Clockings(j).IdClocking, 5)
                                FLog.EscribirLinea("idTerminal:" & Clockings(j).IdTerminal, 5)
                                FLog.EscribirLinea("idZone:" & Clockings(j).IdZone, 5)
                                FLog.EscribirLinea("idIP:" & Clockings(j).IP, 5)
                                'If (Clockings(j).IdReader = 29) Or (Clockings(j).IdReader = 31) Or (Clockings(j).IdReader = 32) Then
                                'Dim x = 0
                                'x = listaRelojes.IndexOf("'" & Clockings(j).IdReader & "'")
                                'If (x <> -1) Then
                                'reloj = Clockings(j).IdReader
                                If Clockings(j).IdTerminal = -1 Then
                                    'busco el reloj por default
                                    FLog.EscribirLinea("Busco el reloj por defecto, ya que la registracion no tiene reloj", 5)
                                    StrSql = " SELECT * FROM gti_reloj "
                                    StrSql += " WHERE reldefault=-1"
                                    dtDatosAux = New DataTable
                                    da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                    da.Fill(dtDatosAux)
                                    If dtDatosAux.Rows.Count <= 0 Then
                                        FLog.EscribirLinea("No hay un reloj asiganado por defecto", 5)
                                        permitido = False
                                    Else
                                        FLog.EscribirLinea("El reloj por defecto es: " & dtDatosAux.Rows(0).Item("relcodext").ToString(), 5)
                                        reloj = dtDatosAux.Rows(0).Item("relnro").ToString()
                                        permitido = True
                                    End If
                                Else
                                    Dim x = 0
                                    x = listaRelojes.IndexOf("'" & Clockings(j).IdTerminal & "'")
                                    If x <> -1 Then
                                        FLog.EscribirLinea("Busco el reloj: " & Clockings(j).IdTerminal)
                                        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & Clockings(j).IdTerminal.ToString & "'"
                                        dtDatosAux = New DataTable
                                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                        da.Fill(dtDatosAux)
                                        FLog.EscribirLinea("Legajo Interno: " & empId & " registración: " & Clockings(j).Datetime, 1)
                                        If dtDatosAux.Rows.Count <= 0 Then
                                            FLog.EscribirLinea("Error. Reloj no encontrado: " & Clockings(j).IdTerminal, 3)
                                            FLog.EscribirLinea("SQL: " & StrSql, 3)
                                            permitido = False
                                        Else
                                            reloj = dtDatosAux.Rows(0).Item("relnro").ToString()
                                            permitido = True
                                        End If
                                    Else
                                        FLog.EscribirLinea("El reloj no esta en la lista de relojes permitidos", 5)
                                        permitido = False
                                    End If
                                End If
                                If permitido = True Then
                                    'busco el numero del empleado de la registracion
                                    StrSqlEmp = " SELECT * FROM tercero "
                                    StrSqlEmp += "INNER JOIN ter_doc ON ter_doc.ternro = tercero.ternro "
                                    StrSqlEmp += " WHERE ter_doc.nrodoc='" & empDni & "'"
                                    dtDatosAux = New DataTable
                                    da = New OleDbDataAdapter(StrSqlEmp, conexion.ConnectionString)
                                    da.Fill(dtDatosAux)
                                    If dtDatosAux.Rows.Count <= 0 Then
                                        FLog.EscribirLinea("No se encontro el tercero")
                                    Else
                                        rhpro_ternro = dtDatosAux.Rows(0).Item("ternro")
                                        ' Verifica si ya existe la registracion
                                        StrSql = "SELECT * FROM gti_registracion WHERE ternro= " & rhpro_ternro & " AND regfecha= " & Fechas.convFechaQP(Clockings(j).Datetime, "dd-mm-yyyy") & _
                                                " AND reghora= " & Fechas.ObtenerHoraDeFecha(Clockings(j).Datetime) & " AND relnro= " & reloj
                                        dtDatosAux = New DataTable
                                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                                        da.Fill(dtDatosAux)
                                        If dtDatosAux.Rows.Count > 0 Then
                                            FLog.EscribirLinea("La registración  ya Existe. Tercero: " & rhpro_ternro & " registración: " & Fechas.convFechaQP(Clockings(j).Datetime, "dd-mm-yyyy") & " - " & Fechas.ObtenerHoraDeFecha(Clockings(j).Datetime))
                                        Else
                                            StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,relnro)" & _
                                                    " VALUES ( " & rhpro_ternro & "," & Fechas.convFechaQP(Clockings(j).Datetime, "dd-mm-yyyy") & ",'" & Fechas.ObtenerHoraDeFecha(Clockings(j).Datetime) & _
                                                    "'," & reloj & ")"
                                            conexion.Open()
                                            cmd.CommandText = StrSql
                                            cmd.ExecuteNonQuery()
                                            conexion.Close()
                                            FLog.EscribirLinea("Se inserta la registracion: " & StrSql, 5)



                                            '------------mdff
                                            FLog.EscribirLinea("Insert empleados en lista")
                                            agregarProcOnline = True
                                            If UBound(Terceros) > 0 Then
                                                For i = 1 To UBound(Terceros)
                                                    If (Terceros(i).Ternro = rhpro_ternro) And (Terceros(i).Fecha = Fechas.convFechaQP(Clockings(j).Datetime, "dd-mm-yyyy").Replace("'", "")) Then
                                                        agregarProcOnline = False
                                                        Exit For
                                                    End If
                                                Next
                                            End If

                                            If agregarProcOnline Then
                                                ReDim Preserve Terceros(UBound(Terceros) + 1)
                                                Terceros(UBound(Terceros)) = New ProcesamientoOnline
                                                Terceros(UBound(Terceros)).Ternro = rhpro_ternro
                                                Terceros(UBound(Terceros)).Fecha = Fechas.convFechaQP(Clockings(j).Datetime, "dd-mm-yyyy")
                                                FLog.EscribirLinea("Ternro" & Terceros(UBound(Terceros)).Ternro)
                                                FLog.EscribirLinea("fecha" & Terceros(UBound(Terceros)).Fecha)
                                            End If
                                            '---------mdff
                                            FLog.EscribirLinea("fin de Insert empleados en lista")
                                        End If
                                    End If
                                    'End If
                                End If
                                ' FLog.EscribirLinea("---Sigo en el segundo for---")

                            Catch e As Exception
                                FLog.EscribirLinea(e.Message)
                            End Try
                        Next 'avanza las registraciones del empleado
                        'FLog.EscribirLinea("---Sali del segundo for---")
                        conexion.Open()
                        da = New OleDbDataAdapter(StrSql, conexion)
                        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & "  WHERE bpronro = " & NroProceso
                        cmd.CommandText = StrSql
                        cmd.ExecuteNonQuery()
                        conexion.Close()
                        'FLog.EscribirLinea("---Sigo en el primer for---")
                    Next 'avanza los empleados
                    'FLog.EscribirLinea("---Sali del primer  for---")
                    'actualizo el progreso
                    conexion.Open()
                    da = New OleDbDataAdapter(StrSql, conexion)
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = 100 WHERE bpronro = " & NroProceso
                    cmd.CommandText = StrSql
                    cmd.ExecuteNonQuery()
                    conexion.Close()
                    'hasta aca

                    'hasta aca
                    '-------------------------------------------------------------------
                    'End If
                End If
            End If
            'End If
        End If
        FLog.EscribirLinea("---Termino Spec!!!!---")
    End Sub
    Public Sub InsertaFormatoSpec()
        'variables web service
        Dim wsContract As New WebServiceContractClient
        Dim listEmp As New WSElement
        Dim fields() As String
        Dim ws As New WSElement
        Dim empleado As New WSElement
        Dim tarjeta As New WSElement
        'variables query
        Dim StrSql3 As String
        Dim StrSql2 As String
        Dim dtDatos As New DataTable
        Dim dtDatosAux2 As New DataTable
        Dim dtDatosAux3 = New DataTable
        Dim dtDatosAux4 = New DataTable
        'variables parametros
        Dim parametros As String
        Dim param
        Dim rhpro_ternro As Integer
        Dim rhpro_nrotarj As String
        Dim operacion As String
        Dim validoDesde As String
        Dim validoHasta As String

        'variables empleado
        Dim fechaAltaSist As String
        Dim rhpro_dni_emp As String
        Dim rhpro_ternom As String
        Dim rhpro_ternom2 As String
        Dim rhpro_terape As String
        Dim rhpro_terape2 As String
        Dim rhpro_fechaAltaEmpresa As String
        Dim nronivel(0) As String
        Dim nivel As String
        Dim empId As Integer

        'variables del empleado de spec
        Dim dni As String
        Dim encontro As Boolean

        'variables nuevas
        Dim listaTarjetas As WSElement
        Dim empleados As WSElement
        Dim wsCards As WSElement
        Dim idTarj
        Dim empTarj
        Dim tarjetaAnt
        Dim tarjetaNue
        Dim tarjetas
        Dim bpronro



        'Dim webService
        Dim endpointAddress
        Dim newEndPointAddress
        Dim endpoint
        Dim nroCon As Integer
        FLog.EscribirLinea("Ingresa al modelo de spec", 5)
        endpoint = "http://cstest.grupospec.com:8097/WebService"
        'busco en el confrep la configuracion del endpoint
        StrSql = " SELECT * FROM confrep"
        StrSql += " WHERE repnro=421"
        StrSql += " AND  conftipo='BDO'"
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        dtDatosAux2 = New DataTable
        da.Fill(dtDatosAux2)
        If (dtDatosAux2.Rows.Count > 0) Then
            nroCon = dtDatosAux2.Rows(0).Item("confval")
            StrSql = " SELECT * FROM conexion "
            StrSql += " WHERE cnnro=" & nroCon
            da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
            dtDatosAux2 = New DataTable
            da.Fill(dtDatosAux2)
            If (dtDatosAux2.Rows.Count > 0) Then
                endpoint = dtDatosAux2.Rows(0).Item("cnstring")
            End If
        End If

        endpointAddress = wsContract.Endpoint.Address
        newEndPointAddress = New EndpointAddressBuilder(endpointAddress)
        newEndPointAddress.uri = New Uri(endpoint)
        wsContract = New WebServiceContractClient("WSHttpBinding_IWebServiceContract", newEndPointAddress.ToEndPointAddress().ToString)


        encontro = False
        fields = {"id", "name"}
        'listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
        FLog.EscribirLinea("Se empieza a procesar la interfaz spec", 5)
        fechaAltaSist = Format(Now(), "yyyy/MM/dd")
        FLog.EscribirLinea("Fecha de alta en el sistema: " & fechaAltaSist, 5)
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)
        If (dtDatos.Rows.Count > 0) Then
            'Levanto los parametros, de aca dedusco si es un alta/baja/mod o alta masiva de empleados
            'nro de tercero, nro tarjeta, operacion, fechadesde, fechahasta
            parametros = dtDatos.Rows(0).Item("bprcparam").ToString
            param = Split(parametros, "@")
            If UBound(param) > 1 Then
                IncPorc = 90 / 1
                FLog.EscribirLinea("La cantidad de parametros es mayor a 1, Vienen los datos del trigger", 5)
                rhpro_ternro = param(0)
                FLog.EscribirLinea("Numero de tercero en RHPro: " & rhpro_ternro, 5)
                rhpro_nrotarj = param(1)
                FLog.EscribirLinea("Numero de tarjeta de RHPro: " & rhpro_nrotarj, 5)
                operacion = param(2)
                FLog.EscribirLinea("Tipo de oparacion: " & operacion, 5)
                validoDesde = param(3)
                FLog.EscribirLinea("Tarjeta valida desde:" & validoDesde, 5)
                validoHasta = param(4)
                FLog.EscribirLinea("Tarjeta valida Hasta:" & validoHasta, 5)

                'desde aca
                'consulto el documento del empleado
                StrSql2 = " SELECT * FROM ter_doc "
                StrSql2 += " INNER JOIN  tercero ON tercero.ternro = ter_doc.ternro "
                StrSql2 += " WHERE tercero.ternro=" & rhpro_ternro
                StrSql2 += " AND ter_doc.tidnro = 1 "
                FLog.EscribirLinea("Se busca los datos del tercero, el mismo debe tener documento tipo 1(DNI)", 5)
                FLog.EscribirLinea(StrSql2, 5)
                da = New OleDbDataAdapter(StrSql2, conexion.ConnectionString)
                dtDatosAux2 = New DataTable
                da.Fill(dtDatosAux2)
                If (dtDatosAux2.Rows.Count > 0) Then
                    FLog.EscribirLinea("Se encontro un tercero en RHPro", 5)
                    rhpro_dni_emp = dtDatosAux2.Rows(0).Item("nrodoc").ToString()
                    FLog.EscribirLinea("Dni del tercero: " & rhpro_dni_emp, 5)
                    rhpro_ternom = dtDatosAux2.Rows(0).Item("ternom").ToString()
                    rhpro_ternom2 = dtDatosAux2.Rows(0).Item("ternom2").ToString()
                    rhpro_terape = dtDatosAux2.Rows(0).Item("terape").ToString()
                    rhpro_terape2 = dtDatosAux2.Rows(0).Item("terape2").ToString()
                    If rhpro_ternom2 <> "" Then
                        rhpro_ternom = rhpro_ternom & " " & rhpro_ternom2
                    End If
                    FLog.EscribirLinea("Nombre: " & rhpro_ternom)
                    If rhpro_terape2 <> "" Then
                        rhpro_terape2 = rhpro_terape & " " & rhpro_terape2
                    End If
                    FLog.EscribirLinea("Apellido: " & rhpro_terape)

                    'con el numero de tercero busco la fecha de alta en la empresa
                    StrSql3 = " SELECT * FROM his_estructura "
                    StrSql3 += " WHERE ternro =" & rhpro_ternro
                    StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                    StrSql3 += " AND tenro=10"
                    da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                    dtDatosAux2 = New DataTable
                    da.Fill(dtDatosAux2)
                    FLog.EscribirLinea("Busco la fecha de ingreso a la empresa (TE 10 en RHPro):" & StrSql3, 5)
                    If (dtDatosAux2.Rows.Count > 0) Then
                        rhpro_fechaAltaEmpresa = dtDatosAux2.Rows(0).Item("htetdesde").ToString
                        rhpro_fechaAltaEmpresa = Format(dtDatosAux2.Rows(0).Item("htetdesde"), "yyyy-MM-dd")
                        FLog.EscribirLinea("Se encontro la fecha de ingreso en la empresa: " & rhpro_fechaAltaEmpresa, 5)
                    Else
                        rhpro_fechaAltaEmpresa = ""
                        FLog.EscribirLinea("No se encontro la fecha de ingreso a la empresa", 5)
                    End If
                    'hasta aca

                    'busco los niveles en el organigrama
                    StrSql3 = "SELECT * FROM confrep "
                    StrSql3 += "WHERE repnro=421 "
                    StrSql3 += "AND conftipo='TE'"
                    StrSql3 += "ORDER BY confnrocol ASC "
                    da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                    dtDatosAux2 = New DataTable
                    da.Fill(dtDatosAux2)
                    For i = 0 To dtDatosAux2.Rows.Count - 1
                        ReDim Preserve nronivel(dtDatosAux2.Rows(i).Item("confnrocol"))
                        nronivel(dtDatosAux2.Rows(i).Item("confnrocol")) = dtDatosAux2.Rows(i).Item("confval").ToString()
                    Next
                    'hasta aca

                    'para cada uno de los niveles busco la descripcion
                    For j = 1 To UBound(nronivel)
                        StrSql3 = " SELECT * FROM his_estructura "
                        StrSql3 += " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                        StrSql3 += " WHERE ternro =" & rhpro_ternro
                        StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                        StrSql3 += " AND estructura.tenro=" & nronivel(j)
                        FLog.EscribirLinea("para cada nivel busca la descripcion:" & StrSql3, 5)
                        da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                        dtDatosAux2 = New DataTable
                        da.Fill(dtDatosAux2)
                        FLog.EscribirLinea("Busco la descripcion del nivel  (TE: " & nronivel(j) & " en RHPro):" & StrSql3, 5)
                        If (dtDatosAux2.Rows.Count > 0) Then
                            nivel += dtDatosAux2.Rows(0).Item("estrdabr").ToString & "/"
                            FLog.EscribirLinea("Nivel: " & j & "Descripcion Nivel: " & dtDatosAux2.Rows(0).Item("estrdabr").ToString, 5)
                        Else
                            nivel = ""
                            FLog.EscribirLinea("No se la descripcion del nivel ", 5)
                        End If
                    Next
                    'hasta aca
                    'End If
                    'hasta aca

                    Select Case operacion
                        Case "A", "M"
                            'Valido si en spec ya hay empleados insertados
                            listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                            If (listEmp.Data.Count = 0) Then
                                'no hay empleados en spec aun
                                FLog.EscribirLinea("No hay empleados en spec cargados, se inserta el empleado", 5)

                                'inserto el empleado
                                ws = wsContract.Get(WSContainer.Employee, -1)
                                ws.Data("name") = rhpro_dni_emp
                                ws.Data("nameEmployee") = rhpro_ternom
                                ws.Data("LastName") = rhpro_terape
                                ws.Data("REGISTERSYSTEMDATE") = fechaAltaSist
                                ws.Data("ACTIVEDAYS") = rhpro_fechaAltaEmpresa
                                ws.Data("DEPARTAMENTS") = nivel
                                wsContract.Set(WSContainer.Employee, ws)

                                tarjeta = wsContract.Get(WSContainer.Card, -1)
                                tarjeta.Data("Number") = rhpro_nrotarj
                                tarjeta.Data("employee") = empId
                                wsContract.Set(WSContainer.Card, tarjeta)

                                ws = wsContract.Get(WSContainer.Employee, rhpro_ternro)
                                ws.Data("cards") = rhpro_nrotarj
                                ws.Data("cards_dateini") = validoDesde
                                ws.Data("cards_dateend") = validoHasta
                                wsContract.Set(WSContainer.Employee, ws)
                                FLog.EscribirLinea("Se inserta el empleado " & rhpro_terape & " con id " & rhpro_ternro, 5)
                                'hasta aca

                                'End If
                            Else
                                'ya hay empleados en la base de spec
                                If operacion = "A" Then
                                    FLog.EscribirLinea("Se va a realizar un alta", 5)
                                    FLog.EscribirLinea("Se busca si el empleado de RHPro ya existe en el sistema", 5)
                                    For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                        'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                                        empleado = pair.Value
                                        dni = empleado.Data("name")
                                        If dni = rhpro_dni_emp Then
                                            FLog.EscribirLinea("El empleado ya existe, se asigna una nueva tarjeta", 5)
                                            encontro = True
                                            empId = empleado.Data("id")
                                            Exit For
                                        Else
                                            encontro = False
                                        End If
                                    Next
                                    If Not encontro Then
                                        'el empleado no existe lo inserto
                                        ws = New WSElement()
                                        ws = wsContract.Get(WSContainer.Employee, -1)
                                        ws.Data("name") = rhpro_dni_emp
                                        ws.Data("nameEmployee") = rhpro_ternom
                                        ws.Data("LastName") = rhpro_terape
                                        ws.Data("REGISTERSYSTEMDATE") = fechaAltaSist
                                        ws.Data("ACTIVEDAYS") = rhpro_fechaAltaEmpresa
                                        ws.Data("DEPARTAMENTS") = nivel
                                        wsContract.Set(WSContainer.Employee, ws)

                                        'tengo que buscar el id del empleado para insertarle la tarjeta
                                        listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                                        FLog.EscribirLinea("Se busca si el empleado de RHPro ya existe en el sistema", 5)
                                        For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                            'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                                            empleado = pair.Value
                                            dni = empleado.Data("name")
                                            If dni = rhpro_dni_emp Then
                                                empId = empleado.Data("id")
                                                tarjeta = wsContract.Get(WSContainer.Card, -1)
                                                tarjeta.Data("Number") = rhpro_nrotarj
                                                tarjeta.Data("employee") = empId
                                                wsContract.Set(WSContainer.Card, tarjeta)

                                                ws = wsContract.Get(WSContainer.Employee, empId)
                                                ws.Data("cards") = rhpro_nrotarj
                                                ws.Data("cards_dateini") = validoDesde
                                                ws.Data("cards_dateend") = validoHasta
                                                wsContract.Set(WSContainer.Employee, ws)
                                                FLog.EscribirLinea("Se inserta el empleado " & rhpro_terape & " con id " & rhpro_ternro, 5)
                                                Exit For
                                            End If
                                        Next
                                        'hasta aca
                                    Else
                                        'encontro el empleado le agrego la tarjeta
                                        tarjeta = wsContract.Get(WSContainer.Card, -1)
                                        tarjeta.Data("Number") = rhpro_nrotarj
                                        tarjeta.Data("employee") = empId
                                        wsContract.Set(WSContainer.Card, tarjeta)

                                        ws = wsContract.Get(WSContainer.Employee, empId)
                                        ws.Data("cards") = rhpro_nrotarj
                                        ws.Data("cards_dateini") = validoDesde
                                        ws.Data("cards_dateend") = validoHasta
                                        wsContract.Set(WSContainer.Employee, ws)
                                        FLog.EscribirLinea("se le asigna una nueva tarjeta al empleado", 5)
                                        'hasta aca
                                    End If
                                Else
                                    If operacion = "M" Then
                                        FLog.EscribirLinea("La operacion es una Modificacion", 5)
                                        tarjetas = Split(rhpro_nrotarj, ",")
                                        tarjetaNue = tarjetas(0)
                                        tarjetaAnt = tarjetas(1)
                                        fields = {"Cards", "id", "name"}
                                        'buscar el id del empleado para reemplazar abajo
                                        '2048@0000000000000069,0000000000000069@M@Ene  1 2014 12:00AM@
                                        listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                                        FLog.EscribirLinea("Se busca el id del empleado con el nro de tarjeta", 5)
                                        For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                            empleados = pair.Value
                                            empId = empleados.Data("id")

                                            empleados = wsContract.Get(WSContainer.Employee, empId)
                                            wsCards = empleados.Data("Cards")
                                            dni = empleados.Data("name")
                                            listaTarjetas = wsContract.List(WSContainer.Card)
                                            For Each pair2 As KeyValuePair(Of String, Object) In wsCards.Data
                                                'Console.WriteLine(listaTarjetas.Data("cards"))
                                                wsCards = wsContract.Get(WSContainer.Card, pair2.Value)
                                                idTarj = wsCards.Data("id")
                                                empTarj = wsCards.Data("Number")
                                                If empTarj = tarjetaAnt Or (dni = rhpro_dni_emp) Then
                                                    If (dni = rhpro_dni_emp) And (empTarj <> tarjetaAnt) Then
                                                        FLog.EscribirLinea("el empleado no tenia la tarjeta a modificar, se le inserta una nueva", 5)
                                                        tarjeta = wsContract.Get(WSContainer.Card, -1)
                                                        tarjeta.Data("Number") = rhpro_nrotarj
                                                        tarjeta.Data("employee") = empId
                                                        wsContract.Set(WSContainer.Card, tarjeta)
                                                        FLog.EscribirLinea("Se le inserto una tarjeta al empleado", 5)
                                                        encontro = True
                                                        Exit For

                                                    Else
                                                        wsCards.Data("Number") = tarjetaNue
                                                        wsContract.Set(WSContainer.Card, wsCards)
                                                        FLog.EscribirLinea("Se actualizo la tarjeta del empleado: " & rhpro_terape & ", " & rhpro_ternom, 5)
                                                        encontro = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next
                                        Next
                                        If Not encontro Then
                                            FLog.EscribirLinea("No se encontro ningun empleado con el numero de tarjeta modificado en RHPro", 5)
                                            'tengo que buscar el id del empleado para insertarle la tarjeta
                                            listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                                            FLog.EscribirLinea("Se busca si el empleado de RHPro ya existe en el sistema", 5)
                                            For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                                'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                                                empleado = pair.Value
                                                dni = empleado.Data("name")
                                                If dni = rhpro_dni_emp Then
                                                    empId = empleado.Data("id")
                                                    tarjeta = wsContract.Get(WSContainer.Card, -1)
                                                    tarjeta.Data("Number") = tarjetaNue
                                                    tarjeta.Data("employee") = empId
                                                    wsContract.Set(WSContainer.Card, tarjeta)

                                                    ws = wsContract.Get(WSContainer.Employee, empId)
                                                    ws.Data("cards") = rhpro_nrotarj
                                                    ws.Data("cards_dateini") = validoDesde
                                                    ws.Data("cards_dateend") = validoHasta
                                                    wsContract.Set(WSContainer.Employee, ws)
                                                    FLog.EscribirLinea("Se inserta la tarjeta al empleado " & rhpro_terape & " con id " & empId, 5)
                                                    encontro = True
                                                    Exit For
                                                End If
                                            Next

                                            If Not encontro Then
                                                FLog.EscribirLinea("El empleado de RHPro no esta en spec, se inserta", 5)
                                                '19/12/2013
                                                'el empleado no existe lo inserto
                                                ws = New WSElement()
                                                ws = wsContract.Get(WSContainer.Employee, -1)
                                                ws.Data("name") = rhpro_dni_emp
                                                ws.Data("nameEmployee") = rhpro_ternom
                                                ws.Data("LastName") = rhpro_terape
                                                ws.Data("REGISTERSYSTEMDATE") = fechaAltaSist
                                                ws.Data("ACTIVEDAYS") = rhpro_fechaAltaEmpresa
                                                ws.Data("DEPARTAMENTS") = nivel
                                                wsContract.Set(WSContainer.Employee, ws)

                                                'tengo que buscar el id del empleado para insertarle la tarjeta
                                                listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                                                FLog.EscribirLinea("Se busca si el empleado de RHPro ya existe en el sistema", 5)
                                                For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                                    'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                                                    empleado = pair.Value
                                                    dni = empleado.Data("name")
                                                    If dni = rhpro_dni_emp Then
                                                        empId = empleado.Data("id")
                                                        tarjeta = wsContract.Get(WSContainer.Card, -1)
                                                        tarjeta.Data("Number") = tarjetaNue
                                                        tarjeta.Data("employee") = empId
                                                        wsContract.Set(WSContainer.Card, tarjeta)

                                                        ws = wsContract.Get(WSContainer.Employee, empId)
                                                        ws.Data("cards") = tarjetaNue
                                                        ws.Data("cards_dateini") = validoDesde
                                                        ws.Data("cards_dateend") = validoHasta
                                                        wsContract.Set(WSContainer.Employee, ws)
                                                        FLog.EscribirLinea("Se inserta el empleado  con id " & empId, 5)
                                                        Exit For
                                                    End If
                                                Next
                                                'fin 19/12/2013
                                            End If
                                        End If
                                        'hasta aca
                                    End If
                                End If
                            End If
                        Case "B"
                            'se desea dar de baja la tarjeta del empleado
                            FLog.EscribirLinea("La operacion es una baja de tarjeta", 5)
                            listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                            FLog.EscribirLinea("Se busca el id del empleado con el nro de tarjeta", 5)
                            For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                empleados = pair.Value
                                empId = empleados.Data("id")

                                empleados = wsContract.Get(WSContainer.Employee, empId)
                                wsCards = empleados.Data("Cards")
                                listaTarjetas = wsContract.List(WSContainer.Card)
                                For Each pair2 As KeyValuePair(Of String, Object) In wsCards.Data
                                    'Console.WriteLine(listaTarjetas.Data("cards"))
                                    wsCards = wsContract.Get(WSContainer.Card, pair2.Value)
                                    idTarj = wsCards.Data("id")
                                    empTarj = wsCards.Data("Number")
                                    If empTarj = rhpro_nrotarj Then
                                        'wsCards.Data("Number") = tarjetaNue
                                        wsContract.Delete(WSContainer.Card, idTarj)
                                        FLog.EscribirLinea("Se elimino la tarjeta del empleado: " & rhpro_terape & ", " & rhpro_ternom, 5)
                                        Exit For
                                    End If
                                    'empleado = pair.Value
                                Next
                            Next
                            'hasta aca la baja
                    End Select
                End If
                'actualizo el progreso
                Progreso = Progreso + IncPorc
                conexion.Open()
                da = New OleDbDataAdapter(StrSql, conexion)
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
                cmd.CommandText = StrSql
                cmd.ExecuteNonQuery()
                conexion.Close()
                'hasta aca
            Else
                If UBound(param) = 0 Then
                    If parametros.Length > 0 Then
                        'es carga masiva 
                        FLog.EscribirLinea("CARGA MASIVA DE TARJETAS")
                        bpronro = param(0)
                        '------------------------------------------------------------------------------------------
                        '------------------------------CARGA MASIVA------------------------------------------------
                        '------------------------------------------------------------------------------------------
                        'busco en gti_histarjeta los empleados con el campo sinc = 0 y bpronro igual al que busque
                        StrSql = "SELECT * FROM gti_histarjeta WHERE bpronro = " & bpronro
                        StrSql += " AND sinc=0"
                        dtDatosAux2 = New DataTable
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        da.Fill(dtDatosAux2)
                        If (dtDatosAux2.Rows.Count > 0) Then
                            IncPorc = 90 / dtDatosAux2.Rows.Count
                            For Each row As DataRow In dtDatosAux2.Rows
                                Progreso = Progreso + IncPorc
                                rhpro_nrotarj = row("hstjnrotar").ToString
                                'rhpro_nrotarj = dtDatosAux2.Rows(0).Item("hstjnrotar")
                                'actualizo en la tabla gti_histarjeta el sincronizado
                                StrSql3 = "UPDATE gti_histarjeta "
                                StrSql3 += "SET sinc=-1"
                                StrSql3 += "WHERE bpronro=" & bpronro
                                StrSql3 += "AND ternro=" & row("ternro")
                                StrSql3 += "AND hstjnrotar=" & rhpro_nrotarj
                                conexion.Open()
                                cmd = New OleDbCommand(StrSql3, conexion)
                                cmd.ExecuteNonQuery()
                                conexion.Close()
                                'hasta aca
                                validoDesde = row("hstjfecdes")
                                If IsDBNull(row("hstjfechas")) Then
                                    validoHasta = ""
                                Else
                                    validoHasta = row("hstjfechas")
                                End If
                                'consulto el documento del empleado
                                StrSql2 = " SELECT * FROM ter_doc "
                                StrSql2 += " INNER JOIN  tercero ON tercero.ternro = ter_doc.ternro "
                                StrSql2 += " INNER JOIN  empleado ON empleado.ternro = tercero.ternro "
                                StrSql2 += " WHERE tercero.ternro=" + row("ternro").ToString
                                StrSql2 += " AND ter_doc.tidnro = 1 "
                                dtDatosAux3 = New DataTable
                                da = New OleDbDataAdapter(StrSql2, conexion.ConnectionString)
                                da.Fill(dtDatosAux3)
                                If (dtDatosAux3.Rows.Count > 0) Then
                                    rhpro_dni_emp = dtDatosAux3.Rows(0).Item("nrodoc").ToString
                                    rhpro_ternom = dtDatosAux3.Rows(0).Item("ternom").ToString
                                    rhpro_ternom2 = dtDatosAux3.Rows(0).Item("ternom2").ToString
                                    rhpro_terape = dtDatosAux3.Rows(0).Item("terape").ToString
                                    rhpro_terape2 = dtDatosAux3.Rows(0).Item("terape2").ToString

                                    If rhpro_ternom2 <> "" Then
                                        rhpro_ternom = rhpro_ternom & " " & rhpro_ternom2
                                    End If

                                    If rhpro_terape2 <> "" Then
                                        rhpro_terape2 = rhpro_terape & " " & rhpro_terape2
                                    End If

                                    'con el numero de tercero busco la fecha de alta en la empresa
                                    StrSql3 = " SELECT * FROM his_estructura "
                                    StrSql3 += " WHERE ternro =" & dtDatosAux3.Rows(0).Item("ternro").ToString
                                    StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                                    StrSql3 += " AND tenro=10"
                                    dtDatosAux4 = New DataTable
                                    da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                                    da.Fill(dtDatosAux4)
                                    If (dtDatosAux4.Rows.Count > 0) Then
                                        rhpro_fechaAltaEmpresa = dtDatosAux4.Rows(0).Item("htetdesde").ToString
                                        'fechaAltaEmpresa = fechaAltaEmpresa.ToString("yyyy/MM/dd")
                                        rhpro_fechaAltaEmpresa = Format(dtDatosAux4.Rows(0).Item("htetdesde"), "yyyy-MM-dd")
                                    Else
                                        rhpro_fechaAltaEmpresa = ""
                                    End If
                                    'hasta aca

                                Else
                                    rhpro_dni_emp = 0
                                End If
                                'hasta aca

                                'traigo todos los empleados del web service
                                fields = New String() {"name", "id"}
                                listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                                If (listEmp.Data.Count = 0) Then
                                    'no hay empleados en spec aun
                                    FLog.EscribirLinea("No hay empleados en spec cargados, se inserta el empleado", 5)
                                    'inserto el empleado
                                    ws = wsContract.Get(WSContainer.Employee, -1)
                                    'ws.Data("id") = rhpro_ternro
                                    ws.Data("name") = rhpro_dni_emp
                                    ws.Data("nameEmployee") = rhpro_ternom
                                    ws.Data("LastName") = rhpro_terape
                                    ws.Data("REGISTERSYSTEMDATE") = fechaAltaSist
                                    ws.Data("ACTIVEDAYS") = rhpro_fechaAltaEmpresa
                                    ws.Data("DEPARTAMENTS") = nivel
                                    wsContract.Set(WSContainer.Employee, ws)

                                    'busco el id del empleado para insertarle la tarjeta 19/12/2013
                                    listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                                    For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                        'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                                        empleado = pair.Value
                                        dni = empleado.Data("name")
                                        FLog.EscribirLinea("DNI empleado: " & dni, 5)
                                        If dni = rhpro_dni_emp Then
                                            empId = empleado.Data("id")
                                            tarjeta = wsContract.Get(WSContainer.Card, -1)
                                            tarjeta.Data("Number") = rhpro_nrotarj
                                            tarjeta.Data("employee") = empId
                                            wsContract.Set(WSContainer.Card, tarjeta)

                                            ws = wsContract.Get(WSContainer.Employee, empId)
                                            ws.Data("cards") = rhpro_nrotarj
                                            ws.Data("cards_dateini") = validoDesde
                                            ws.Data("cards_dateend") = validoHasta
                                            ws.Data("employeeCode") = ""
                                            wsContract.Set(WSContainer.Employee, ws)
                                            FLog.EscribirLinea("Se inserta el empleado " & rhpro_terape & " con id " & empId, 5)
                                            Exit For
                                        End If
                                    Next
                                    'hasta aca

                                Else
                                    'ya hay empleados en spec
                                    FLog.EscribirLinea("Se busca si el empleado de RHPro ya existe en el sistema", 5)
                                    For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                        'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                                        empleado = pair.Value
                                        'empId = wsEmp.Data("id")
                                        dni = empleado.Data("name")
                                        If dni = rhpro_dni_emp Then
                                            FLog.EscribirLinea("El empleado ya existe, se asigna una nueva tarjeta", 5)
                                            encontro = True
                                            empId = empleado.Data("id")
                                            Exit For
                                        Else
                                            encontro = False
                                        End If
                                    Next
                                    If Not encontro Then
                                        'el empleado no existe lo inserto
                                        ws = New WSElement()
                                        ws = wsContract.Get(WSContainer.Employee, -1)
                                        ws.Data("name") = rhpro_dni_emp
                                        ws.Data("nameEmployee") = rhpro_ternom
                                        ws.Data("LastName") = rhpro_terape
                                        ws.Data("REGISTERSYSTEMDATE") = fechaAltaSist
                                        ws.Data("ACTIVEDAYS") = rhpro_fechaAltaEmpresa
                                        ws.Data("DEPARTAMENTS") = nivel
                                        wsContract.Set(WSContainer.Employee, ws)

                                        'tengo que buscar el id del empleado para insertarle la tarjeta
                                        listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                                        FLog.EscribirLinea("Se busca si el empleado de RHPro ya existe en el sistema", 5)
                                        For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                            'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                                            empleado = pair.Value
                                            'empId = wsEmp.Data("id")
                                            dni = empleado.Data("name")
                                            If dni = rhpro_dni_emp Then
                                                empId = empleado.Data("id")
                                                tarjeta = wsContract.Get(WSContainer.Card, -1)
                                                tarjeta.Data("Number") = rhpro_nrotarj
                                                tarjeta.Data("employee") = empId
                                                wsContract.Set(WSContainer.Card, tarjeta)

                                                ws = wsContract.Get(WSContainer.Employee, empId)
                                                ws.Data("cards") = rhpro_nrotarj
                                                ws.Data("cards_dateini") = validoDesde
                                                ws.Data("cards_dateend") = validoHasta
                                                wsContract.Set(WSContainer.Employee, ws)
                                                FLog.EscribirLinea("Se inserta el empleado " & rhpro_terape & " con id " & empId, 5)
                                                Exit For
                                            End If
                                        Next
                                        'hasta aca
                                    Else
                                        'encontro el empleado le agrego la tarjeta
                                        tarjeta = wsContract.Get(WSContainer.Card, -1)
                                        tarjeta.Data("Number") = rhpro_nrotarj
                                        tarjeta.Data("employee") = empId
                                        wsContract.Set(WSContainer.Card, tarjeta)

                                        ws = wsContract.Get(WSContainer.Employee, empId)
                                        ws.Data("cards") = rhpro_nrotarj
                                        ws.Data("cards_dateini") = validoDesde
                                        ws.Data("cards_dateend") = validoHasta
                                        wsContract.Set(WSContainer.Employee, ws)
                                        FLog.EscribirLinea("se le asigna una nueva tarjeta al empleado", 5)

                                        'hasta aca
                                    End If
                                End If
                                'actualizo el progreso
                                conexion.Open()
                                da = New OleDbDataAdapter(StrSql, conexion)
                                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
                                cmd.CommandText = StrSql
                                cmd.ExecuteNonQuery()
                                conexion.Close()
                                'hasta aca
                            Next
                        End If
                        'hasta aca
                    End If
                End If
                'HASTA ACA----------------------------------------------------
            End If
        End If
        '--------------------------------------------------------------------
        'ace se hace la lectura de registraciones para todos los empleados
        Dim FDesde, FHasta As String
        Dim fechadesde As Date
        Dim fechaHasta As Date
        Dim Fechas As New Fechas
        Dim clock As New ClockingType
        Dim Clockings() As Clocking
        Dim WsLectura As New WebServiceContractClient
        Dim employee As New WSElement
        Dim reloj As String
        Dim dtDatosAux As DataTable
        Dim empDni As String
        Dim StrSqlEmp As String
        Dim cantEmpleados As Long
        'endpoint = "http://200.61.13.22:8091/WebService"
        'webService = New WebServiceContractClient()

        FLog.EscribirLinea("levanta los datos del endpoint", 5)
        StrSql = " SELECT * FROM confrep"
        StrSql += " WHERE repnro=421"
        StrSql += " AND  conftipo='BDO'"
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        dtDatosAux2 = New DataTable
        da.Fill(dtDatosAux2)
        If (dtDatosAux2.Rows.Count > 0) Then
            nroCon = dtDatosAux2.Rows(0).Item("confval")
            StrSql = " SELECT * FROM conexion "
            StrSql += " WHERE cnnro=" & nroCon
            da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
            dtDatosAux2 = New DataTable
            da.Fill(dtDatosAux2)
            If (dtDatosAux2.Rows.Count > 0) Then
                endpoint = dtDatosAux2.Rows(0).Item("cnstring")
            End If
        End If
        FLog.EscribirLinea("endpoint:" & endpoint, 5)
        endpointAddress = WsLectura.Endpoint.Address
        newEndPointAddress = New EndpointAddressBuilder(endpointAddress)
        newEndPointAddress.uri = New Uri(endpoint)
        WsLectura = New WebServiceContractClient("WSHttpBinding_IWebServiceContract", newEndPointAddress.ToEndPointAddress().ToString)



        fields = {"name", "id"}
        reloj = ""
        'FDesde = Fechas.cambiaFecha(DateAdd(DateInterval.Day, -1, CDate(dtDatos.Rows(0).Item("bprcfecdesde").ToString)))
        fechadesde = DateAdd(DateInterval.Day, -1, CDate(dtDatos.Rows(0).Item("bprcfecdesde").ToString))
        'FHasta = Fechas.cambiaFecha(dtDatos.Rows(0).Item("bprcfechasta").ToString)
        fechaHasta = dtDatos.Rows(0).Item("bprcfechasta").ToString
        FDesde = Format(fechadesde, "yyyy-MM-dd")
        FHasta = Format(fechaHasta, "yyyy-MM-dd")
        clock = ClockingType.Access
        'para c/u de los empleados busco las registraciones
        employee = WsLectura.ListFields(WSContainer.Employee, fields, "")
        cantEmpleados = employee.Data.Count()
        Progreso = 0
        If cantEmpleados <> 0 Then
            IncPorc = 100 / cantEmpleados
        Else
            cantEmpleados = 1
            IncPorc = 100 / cantEmpleados
        End If
        For Each pair As KeyValuePair(Of String, Object) In employee.Data
            Progreso = Progreso + IncPorc
            'busco las registraciones para cada uno de los empleados
            employee = pair.Value
            empId = employee.Data("id")
            empDni = employee.Data("name")
            Clockings = WsLectura.Clockings(empId, FDesde, FHasta, clock)
            FLog.EscribirLinea("Emp DNI:" & empDni)
            For j = 0 To UBound(Clockings)
                FLog.EscribirLinea("idReader:" & Clockings(j).IdReader, 5)
                FLog.EscribirLinea("idClocking:" & Clockings(j).IdClocking, 5)
                FLog.EscribirLinea("idTerminal:" & Clockings(j).IdTerminal, 5)
                FLog.EscribirLinea("idZone:" & Clockings(j).IdZone, 5)
                FLog.EscribirLinea("idIP:" & Clockings(j).IP, 5)
                If (Clockings(j).IdReader = 29) Or (Clockings(j).IdReader = 31) Or (Clockings(j).IdReader = 32) Then
                    'reloj = Clockings(j).IdReader
                    If Clockings(j).IdReader = -1 Then
                        'busco el reloj por default
                        FLog.EscribirLinea("Busco el reloj por defecto, ya que la registracion no tiene reloj", 5)
                        StrSql = " SELECT * FROM gti_reloj "
                        StrSql += " WHERE reldefault=-1"
                        dtDatosAux = New DataTable
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        da.Fill(dtDatosAux)
                        If dtDatosAux.Rows.Count <= 0 Then
                            FLog.EscribirLinea("No hay un reloj asiganado por defecto", 5)
                        Else
                            FLog.EscribirLinea("El reloj por defecto es: " & dtDatosAux.Rows(0).Item("relcodext").ToString(), 5)
                            reloj = dtDatosAux.Rows(0).Item("relnro").ToString()
                        End If
                    Else
                        FLog.EscribirLinea("Busco el reloj: " & Clockings(j).IdReader)
                        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & Clockings(j).IdReader.ToString & "'"
                        dtDatosAux = New DataTable
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        da.Fill(dtDatosAux)
                        FLog.EscribirLinea("Legajo Interno: " & empId & " registración: " & Clockings(j).Datetime, 1)
                        If dtDatosAux.Rows.Count <= 0 Then
                            FLog.EscribirLinea("Error. Reloj no encontrado: " & Clockings(j).IdReader, 3)
                            FLog.EscribirLinea("SQL: " & StrSql, 3)
                        Else
                            reloj = dtDatosAux.Rows(0).Item("relnro").ToString()
                        End If
                    End If
                    'busco el numero del empleado de la registracion
                    StrSqlEmp = " SELECT * FROM tercero "
                    StrSqlEmp += "INNER JOIN ter_doc ON ter_doc.ternro = tercero.ternro "
                    StrSqlEmp += " WHERE ter_doc.nrodoc='" & empDni & "'"
                    dtDatosAux = New DataTable
                    da = New OleDbDataAdapter(StrSqlEmp, conexion.ConnectionString)
                    da.Fill(dtDatosAux)
                    If dtDatosAux.Rows.Count <= 0 Then
                        FLog.EscribirLinea("No se encontro el tercero")
                    Else
                        rhpro_ternro = dtDatosAux.Rows(0).Item("ternro")
                        ' Verifica si ya existe la registracion
                        StrSql = "SELECT * FROM gti_registracion WHERE ternro= " & rhpro_ternro & " AND regfecha= " & Fechas.convFechaQP(Clockings(j).Datetime, "dd-mm-yyyy") & _
                                " AND reghora= " & Fechas.ObtenerHoraDeFecha(Clockings(j).Datetime) & " AND relnro= " & reloj
                        dtDatosAux = New DataTable
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        da.Fill(dtDatosAux)
                        If dtDatosAux.Rows.Count > 0 Then
                            FLog.EscribirLinea("La registración  ya Existe. Tercero: " & rhpro_ternro & " registración: " & Fechas.convFechaQP(Clockings(j).Datetime, "dd-mm-yyyy") & " - " & Fechas.ObtenerHoraDeFecha(Clockings(j).Datetime))
                        Else
                            StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,relnro)" & _
                                    " VALUES ( " & rhpro_ternro & "," & Fechas.convFechaQP(Clockings(j).Datetime, "dd-mm-yyyy") & ",'" & Fechas.ObtenerHoraDeFecha(Clockings(j).Datetime) & _
                                    "'," & reloj & ")"
                            conexion.Open()
                            cmd.CommandText = StrSql
                            cmd.ExecuteNonQuery()
                            conexion.Close()
                            FLog.EscribirLinea("Se inserta la registracion: " & StrSql, 5)
                        End If
                    End If
                End If
            Next 'avanza las registraciones del empleado
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & "  WHERE bpronro = " & NroProceso
            cmd.CommandText = StrSql
            cmd.ExecuteNonQuery()
        Next 'avanza los empleados
        'actualizo el progreso
        conexion.Open()
        da = New OleDbDataAdapter(StrSql, conexion)
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100 WHERE bpronro = " & NroProceso
        cmd.CommandText = StrSql
        cmd.ExecuteNonQuery()
        conexion.Close()
        'hasta aca

        'hasta aca
        '-------------------------------------------------------------------
        'End If
    End Sub
    'Inserta lasgistracines con el formato QuickPass
    Public Sub InsertaFormatoQPass(ByVal idempresa As String, ByVal psw As String)
        Dim dsDatos As New DataSet
        Dim dtDatos As New DataTable
        Dim dtDatosAux As DataTable
        Dim dtTercero As DataTable
        Dim objLecturaQP As New ServiceMovimientos
        Dim FDesde, FHasta As String
        Dim Fechas As New Fechas
        Dim IncPorc As Single = 0
        Dim Progreso As Single = 0
        Dim relnro As Long = 0
        Dim i As Long
        Dim agregarProcOnline As Boolean

        StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatos)

        If (dtDatos.Rows.Count > 0) Then
            'EAM- Le resta uno a la fecha desde porque si es online puede que las ultimas registraciones sean del dia anterior y no las toma
            FDesde = Fechas.cambiaFecha(DateAdd(DateInterval.Day, -1, CDate(dtDatos.Rows(0).Item("bprcfecdesde").ToString)))
            FHasta = Fechas.cambiaFecha(dtDatos.Rows(0).Item("bprcfechasta").ToString)
            dsDatos = objLecturaQP.wsListarMovimientos(idempresa, psw, "-1", FDesde, FHasta, -1, -1, 1)

            'EAM- Calcula en incremento del porcentaje
            If (dsDatos.Tables(0).Rows.Count >= 0) Then
                FLog.EscribirLinea(dsDatos.Tables(0).Rows.Count & " Archivos de registraciones encontrados " & Format(Now, "dd/mm/yyyy hh:mm:ss"), 3)
                IncPorc = 100 / dsDatos.Tables(0).Rows.Count
            End If

            For Each MiDataRow As DataRow In dsDatos.Tables(0).Rows
                Try
                    FLog.EscribirLinea("Busco el reloj: " & MiDataRow.Item("LectorSerialNumber"))
                    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & MiDataRow.Item("LectorSerialNumber") & "'"
                    dtDatosAux = New DataTable
                    da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                    da.Fill(dtDatosAux)
                    FLog.EscribirLinea("Legajo: " & MiDataRow.Item("Legajo") & " registración: " & Fechas.convFechaQP(MiDataRow.Item("FechaMovimiento"), "dd-mm-yyyy"), 1)

                    If dtDatosAux.Rows.Count <= 0 Then
                        FLog.EscribirLinea("Error. Reloj no encontrado: " & MiDataRow.Item("LectorSerialNumber"), 3)
                        FLog.EscribirLinea("SQL: " & StrSql, 3)
                    Else
                        'StrSql = "SELECT ternro FROM tercero WHERE ternro = " & MiDataRow.Item("Legajo")
                        StrSql = "SELECT ternro FROM empleado WHERE empleg = " & MiDataRow.Item("Legajo")
                        dtTercero = New DataTable
                        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                        da.Fill(dtTercero)

                        If dtDatosAux.Rows.Count >= 0 Then
                            relnro = dtDatosAux.Rows(0).Item("relnro")
                            'EAM- Verifica si ya existe la registracion
                            StrSql = "SELECT * FROM gti_registracion WHERE ternro= " & dtTercero.Rows(0).Item("ternro") & " AND regfecha= " & Fechas.convFechaQP(MiDataRow.Item("FechaMovimiento"), "dd-mm-yyyy") & _
                                    " AND reghora= " & Fechas.ObtenerHoraDeFecha(MiDataRow.Item("FechaMovimiento")) & " AND relnro= " & relnro
                            dtDatosAux = New DataTable
                            da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
                            da.Fill(dtDatosAux)

                            If dtDatosAux.Rows.Count > 0 Then
                                FLog.EscribirLinea("La registración  ya Existe. Tercero: " & dtTercero.Rows(0).Item("ternro") & " registración: " & Fechas.convFechaQP(MiDataRow.Item("FechaMovimiento"), "dd-mm-yyyy") & " - " & Fechas.ObtenerHoraDeFecha(MiDataRow.Item("FechaMovimiento")))
                            Else
                                StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,relnro)" & _
                                        " VALUES ( " & dtTercero.Rows(0).Item("ternro") & "," & Fechas.convFechaQP(MiDataRow.Item("FechaMovimiento"), "dd-mm-yyyy") & ",'" & Fechas.ObtenerHoraDeFecha(MiDataRow.Item("FechaMovimiento")) & _
                                        "'," & relnro & ")"
                                conexion.Open()
                                cmd.CommandText = StrSql
                                cmd.ExecuteNonQuery()
                                conexion.Close()

                                agregarProcOnline = True
                                If UBound(Terceros) > 0 Then
                                    For i = 1 To UBound(Terceros)
                                        If (Terceros(i).Ternro = dtTercero.Rows(0).Item("ternro")) And (Terceros(i).Fecha = CDate(Fechas.convFechaQP(MiDataRow.Item("FechaMovimiento"), "dd-mm-yyyy").Replace("'", ""))) Then
                                            agregarProcOnline = False
                                            Exit For
                                        End If
                                    Next
                                End If

                                If agregarProcOnline Then
                                    ReDim Preserve Terceros(UBound(Terceros) + 1)
                                    Terceros(UBound(Terceros)) = New ProcesamientoOnline
                                    Terceros(UBound(Terceros)).Ternro = dtTercero.Rows(0).Item("ternro")
                                    Terceros(UBound(Terceros)).Fecha = CDate(Fechas.convFechaQP(MiDataRow.Item("FechaMovimiento"), "dd-mm-yyyy").Replace("'", ""))
                                End If
                            End If
                        End If
                    End If

                Catch ex As Exception
                    FLog.EscribirLinea("ERROR SQL: " & StrSql)
                    FLog.EscribirLinea("Error al insertar la registración del legajo " & MiDataRow.Item("Legajo") & " lector: " & MiDataRow.Item("LectorSerialNumber"))
                Finally
                    'EAM- Actualiza el avance del proceso
                    Progreso = Progreso + IncPorc
                    conexion.Open()
                    da = New OleDbDataAdapter(StrSql, conexion)
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
                    cmd.CommandText = StrSql
                    cmd.ExecuteNonQuery()
                    conexion.Close()
                    FLog.EscribirLinea("----------------------------------------------")
                    FLog.EscribirLinea("")
                End Try
            Next
        End If
    End Sub

    Public Sub TerminarTransferencia()
        If Not Errores Then
            StrSql = "DELETE FROM batch_proceso WHERE bpronro = " & NroProceso
            conexion.Open()
            cmd.CommandText = StrSql
            cmd.ExecuteNonQuery()
            conexion.Close()
            FLog.EscribirLinea("Se ha eliminado de bach_Proceso el Proceso Nro: " & NroProceso)
        End If
    End Sub

    Public Sub cargaMasiva(ByVal bpronro As Long)
        Dim dtDatosAux2 As DataTable
        Dim dtDatosAux3 As DataTable
        Dim dtDatosAux4 As DataTable
        Dim StrSql3 As String
        Dim StrSql2 As String
        Dim rhpro_nrotarj As String
        Dim departs As WSElement
        Dim wsContract As New WebServiceContractClient
        Dim tarjeta As New WSElement
        Dim fields() As String
        Dim wsCards As WSElement
        Dim lista As String
        Dim nronivel(0) As String
        Dim nivel As String
        Dim niveles()
        Dim validoDesde As String
        Dim validoHasta As String
        Dim rhpro_dni_emp As String
        Dim rhpro_ternom As String
        Dim rhpro_ternom2 As String
        Dim rhpro_terape As String
        Dim rhpro_terape2 As String
        Dim rhpro_legajo As Long
        Dim rhpro_fechaAltaEmpresa As String
        Dim rhpro_ternro As Long
        Dim listEmp As New WSElement
        Dim validity As New WSElement()
        Dim empDep As New WSElement()
        Dim SpecEmpId As Long
        Dim ndate As Integer
        Dim enddate As Integer
        Dim dep As New WSElement()
        Dim ws As New WSElement
        Dim empleado As New WSElement
        Dim i As Integer
        Dim dni As String
        Dim encontro As Boolean
        encontro = False
        'busco en gti_histarjeta los empleados con el campo sinc = 0 y bpronro igual al que busque
        StrSql = "SELECT * FROM gti_histarjeta WHERE bpronro = " & bpronro
        StrSql += " AND sinc=0"
        FLog.EscribirLinea("consulta de tarjetas por Interfaz: " & StrSql)
        dtDatosAux2 = New DataTable
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        da.Fill(dtDatosAux2)
        If (dtDatosAux2.Rows.Count > 0) Then
            IncPorc = 99 / dtDatosAux2.Rows.Count
            For Each row As DataRow In dtDatosAux2.Rows
                Progreso = Progreso + IncPorc
                FLog.EscribirLinea("actualizo la tabla de tarjetas")
                rhpro_nrotarj = row("hstjnrotar").ToString
                'actualizo en la tabla gti_histarjeta el sincronizado
                Dim nulo As String
                nulo = "NULL"
                StrSql3 = "UPDATE gti_histarjeta "
                StrSql3 += "SET sinc=-1, "
                StrSql3 += " bpronro =" & nulo
                StrSql3 += " WHERE bpronro=" & bpronro
                StrSql3 += " AND ternro=" & row("ternro")
                StrSql3 += " AND hstjnrotar=" & rhpro_nrotarj
                FLog.EscribirLinea("query update" & StrSql3)
                conexion.Open()
                cmd = New OleDbCommand(StrSql3, conexion)
                cmd.ExecuteNonQuery()
                conexion.Close()
                'hasta aca

                'armo una lista de tarjetas
                tarjeta = wsContract.ListFields(WSContainer.Card, fields, "")
                If tarjeta.Data.Count > 0 Then
                    'armo una lista de tarjetas
                    For Each pair2 As KeyValuePair(Of String, Object) In tarjeta.Data
                        wsCards = pair2.Value
                        lista = lista + "'" + wsCards.Data("Number") + "'"
                    Next
                End If
                'hasta aca

                validoDesde = row("hstjfecdes")
                If IsDBNull(row("hstjfechas")) Then
                    validoHasta = ""
                Else
                    validoHasta = row("hstjfechas")
                End If
                'consulto el documento del empleado
                StrSql2 = " SELECT * FROM ter_doc "
                StrSql2 += " INNER JOIN tercero ON tercero.ternro = ter_doc.ternro "
                StrSql2 += " INNER JOIN empleado ON empleado.ternro = tercero.ternro "
                StrSql2 += " WHERE tercero.ternro=" + row("ternro").ToString
                StrSql2 += " AND ter_doc.tidnro = 1 "
                dtDatosAux3 = New DataTable
                da = New OleDbDataAdapter(StrSql2, conexion.ConnectionString)
                da.Fill(dtDatosAux3)
                If (dtDatosAux3.Rows.Count > 0) Then
                    rhpro_dni_emp = dtDatosAux3.Rows(0).Item("nrodoc").ToString
                    rhpro_ternom = dtDatosAux3.Rows(0).Item("ternom").ToString
                    rhpro_ternom2 = dtDatosAux3.Rows(0).Item("ternom2").ToString
                    rhpro_terape = dtDatosAux3.Rows(0).Item("terape").ToString
                    rhpro_terape2 = dtDatosAux3.Rows(0).Item("terape2").ToString
                    rhpro_legajo = dtDatosAux3.Rows(0).Item("empleg").ToString
                    rhpro_ternro = dtDatosAux3.Rows(0).Item("ternro").ToString
                    If rhpro_ternom2 <> "" Then
                        rhpro_ternom = rhpro_ternom & " " & rhpro_ternom2
                    End If

                    If rhpro_terape2 <> "" Then
                        rhpro_terape2 = rhpro_terape & " " & rhpro_terape2
                    End If

                    'con el numero de tercero busco la fecha de alta en la empresa
                    StrSql3 = " SELECT * FROM his_estructura "
                    StrSql3 += " WHERE ternro =" & dtDatosAux3.Rows(0).Item("ternro").ToString
                    StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                    StrSql3 += " AND tenro=10"
                    dtDatosAux4 = New DataTable
                    da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                    da.Fill(dtDatosAux4)
                    If (dtDatosAux4.Rows.Count > 0) Then
                        rhpro_fechaAltaEmpresa = dtDatosAux4.Rows(0).Item("htetdesde").ToString
                        'fechaAltaEmpresa = fechaAltaEmpresa.ToString("yyyy/MM/dd")
                        rhpro_fechaAltaEmpresa = Format(dtDatosAux4.Rows(0).Item("htetdesde"), "yyyy-MM-dd")
                    Else
                        rhpro_fechaAltaEmpresa = ""
                    End If
                    'hasta aca
                    'busco los niveles en el organigrama
                    StrSql3 = "SELECT * FROM confrep "
                    StrSql3 += "WHERE repnro=421 "
                    StrSql3 += "AND conftipo='TE'"
                    StrSql3 += "ORDER BY confnrocol ASC "
                    da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                    dtDatosAux2 = New DataTable
                    da.Fill(dtDatosAux2)
                    For i = 0 To dtDatosAux2.Rows.Count - 1
                        ReDim Preserve nronivel(dtDatosAux2.Rows(i).Item("confnrocol"))
                        nronivel(dtDatosAux2.Rows(i).Item("confnrocol")) = dtDatosAux2.Rows(i).Item("confval").ToString()
                    Next
                    'hasta aca

                    'para cada uno de los niveles busco la descripcion
                    For j = 1 To UBound(nronivel)
                        StrSql3 = " SELECT * FROM his_estructura "
                        StrSql3 += " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                        StrSql3 += " WHERE ternro =" & rhpro_ternro
                        StrSql3 += " AND ((htetdesde <= '" & Format(Now(), "dd/MM/yyyy") & "') AND ((htethasta >= '" & Format(Now(), "dd/MM/yyyy") & "') OR (htethasta is null)))"
                        StrSql3 += " AND estructura.tenro=" & nronivel(j)
                        FLog.EscribirLinea("para cada nivel busca la descripcion:" & StrSql3, 5)
                        da = New OleDbDataAdapter(StrSql3, conexion.ConnectionString)
                        dtDatosAux2 = New DataTable
                        da.Fill(dtDatosAux2)
                        FLog.EscribirLinea("Busco la descripcion del nivel  (TE: " & nronivel(j) & " en RHPro):" & StrSql3, 5)
                        If (dtDatosAux2.Rows.Count > 0) Then
                            nivel += dtDatosAux2.Rows(0).Item("estrdabr").ToString & "/"
                            FLog.EscribirLinea("Nivel: " & j & "Descripcion Nivel: " & dtDatosAux2.Rows(0).Item("estrdabr").ToString, 5)
                        Else
                            nivel = ""
                            FLog.EscribirLinea("No se la descripcion del nivel ", 5)
                        End If
                    Next

                Else
                    rhpro_dni_emp = 0
                End If
                'hasta aca


                'traigo todos los empleados del web service
                fields = New String() {"name", "id", "Number"}
                listEmp = wsContract.ListFields(WSContainer.Employee, fields, "")
                If (listEmp.Data.Count = 0) Then
                    'no hay empleados en spec aun
                    FLog.EscribirLinea("No hay empleados en spec cargados, se inserta el empleado", 5)
                    'inserto el empleado
                    ws = wsContract.Get(WSContainer.Employee, -1)
                    ws.Data("name") = rhpro_dni_emp
                    ws.Data("nameEmployee") = rhpro_ternom
                    ws.Data("LastName") = rhpro_terape
                    ws.Data("REGISTERSYSTEMDATE") = fechaAltaSist
                    ws.Data("ACTIVEDAYS") = rhpro_fechaAltaEmpresa
                    ws.Data("DEPARTAMENTS") = nivel
                    ws.Data("employeeCode") = rhpro_legajo
                    ws.Data("companyCode") = rhpro_nombreEmpresa
                    '------------------------------------------
                    'armo el arbol organizacional

                    niveles = Split(nivel, "/")
                    i = 0
                    empDep.Data = New Dictionary(Of String, Object)
                    For k = 0 To UBound(niveles) - 1
                        departs = wsContract.ListFields(WSContainer.StructureTree, fields, "this.name=""" + niveles(k) + """")
                        'Add department
                        'empDep.Data = New Dictionary(Of String, Object)
                        'i = 0
                        For Each pair3 As KeyValuePair(Of String, Object) In departs.Data
                            'Create validity
                            validity.Data = New Dictionary(Of String, Object)
                            ndate = Year(DateTime.Now())
                            enddate = Year(DateTime.Now()) + 1
                            Do While (ndate < enddate)
                                validity.Data.Add(ndate.ToString(), ndate)
                                ndate = ndate + 1
                            Loop
                            dep = pair3.Value
                            dep.Data.Add("validity", validity)
                            'Add dep to employee
                            empDep.Data.Add((i).ToString(), dep)
                            i = i + 1
                            'Exit For
                        Next
                    Next
                    ws.Data("Departments") = empDep
                    wsContract.Set(WSContainer.Employee, ws)
                    '------------------------------------------

                    'busco el id del empleado para insertarle la tarjeta 19/12/2013
                    listEmp = wsContract.ListFields(WSContainer.Employee, fields, "this.name=""" + rhpro_dni_emp + """")
                    For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                        'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                        empleado = pair.Value
                        SpecEmpId = empleado.Data("id")

                        '30/12/2013 codigo del empleado
                        'tarjeta = wsContract.ListFields(WSContainer.Card, fields, "")
                        'If tarjeta.Data.Count > 0 Then
                        'armo una lista de tarjetas
                        'For Each pair2 As KeyValuePair(Of String, Object) In tarjeta.Data
                        'wsCards = pair2.Value
                        'lista = lista + "'" + wsCards.Data("Number") + "'"
                        'Next
                        Dim x As Integer
                        x = lista.IndexOf("'" & rhpro_nrotarj & "'")
                        If x <> -1 Then
                            FLog.EscribirLinea("La tarjeta :" & rhpro_nrotarj & "ya existe en Spec, no se le inserta al empleado")
                            ws = wsContract.Get(WSContainer.Employee, SpecEmpId)
                            ws.Data("employeeCode") = rhpro_legajo
                            wsContract.Set(WSContainer.Employee, ws)
                            FLog.EscribirLinea("Se actualizo el codigo del empleado en spec")
                        Else
                            tarjeta = wsContract.Get(WSContainer.Card, -1)
                            tarjeta.Data("Number") = rhpro_nrotarj
                            tarjeta.Data("employee") = SpecEmpId
                            wsContract.Set(WSContainer.Card, tarjeta)

                            ws = wsContract.Get(WSContainer.Employee, SpecEmpId)
                            ws.Data("cards") = rhpro_nrotarj
                            ws.Data("cards_dateini") = validoDesde
                            ws.Data("cards_dateend") = validoHasta
                            wsContract.Set(WSContainer.Employee, ws)
                            FLog.EscribirLinea("Se inserta la tarjeta " & rhpro_nrotarj & " al empleado " & rhpro_terape, 5)
                            'hasta aca
                        End If
                        'End If

                    Next
                    'hasta aca

                Else
                    'ya hay empleados en spec
                    listEmp = wsContract.ListFields(WSContainer.Employee, fields, "this.name=""" + rhpro_dni_emp + """")
                    FLog.EscribirLinea("Se busca si el empleado de RHPro ya existe en el sistema", 5)
                    If listEmp.Data.Count > 0 Then
                        For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                            'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                            empleado = pair.Value
                            FLog.EscribirLinea("El empleado ya existe, se asigna una nueva tarjeta", 5)
                            encontro = True
                            SpecEmpId = empleado.Data("id")
                            Exit For
                        Next
                    Else
                        encontro = False
                    End If
                    If Not encontro Then
                        'el empleado no existe lo inserto
                        ws = New WSElement()
                        ws = wsContract.Get(WSContainer.Employee, -1)
                        ws.Data("name") = rhpro_dni_emp
                        ws.Data("nameEmployee") = rhpro_ternom
                        ws.Data("LastName") = rhpro_terape
                        ws.Data("REGISTERSYSTEMDATE") = fechaAltaSist
                        ws.Data("ACTIVEDAYS") = rhpro_fechaAltaEmpresa
                        ws.Data("DEPARTAMENTS") = nivel
                        ws.Data("employeeCode") = rhpro_legajo
                        ws.Data("companyCode") = rhpro_nombreEmpresa
                        '------------------------------------------
                        'armo el arbol organizacional

                        niveles = Split(nivel, "/")
                        i = 0
                        empDep.Data = New Dictionary(Of String, Object)
                        For k = 0 To UBound(niveles) - 1
                            departs = wsContract.ListFields(WSContainer.StructureTree, fields, "this.name=""" + niveles(k) + """")
                            'Add department
                            'empDep.Data = New Dictionary(Of String, Object)
                            'i = 0
                            For Each pair3 As KeyValuePair(Of String, Object) In departs.Data
                                'Create validity
                                validity.Data = New Dictionary(Of String, Object)
                                ndate = 2013
                                enddate = 2015
                                Do While (ndate < enddate)
                                    validity.Data.Add(ndate.ToString(), ndate)
                                    ndate = ndate + 1
                                Loop
                                dep = pair3.Value
                                dep.Data.Add("validity", validity)
                                'Add dep to employee
                                empDep.Data.Add((i).ToString(), dep)
                                i = i + 1
                                'Exit For
                            Next
                        Next
                        ws.Data("Departments") = empDep
                        wsContract.Set(WSContainer.Employee, ws)
                        '------------------------------------------


                        'tengo que buscar el id del empleado para insertarle la tarjeta
                        listEmp = wsContract.ListFields(WSContainer.Employee, fields, "this.name=""" + rhpro_dni_emp + """")
                        FLog.EscribirLinea("Se busca si el empleado de RHPro ya existe en el sistema", 5)
                        If listEmp.Data.Count > 0 Then
                            For Each pair As KeyValuePair(Of String, Object) In listEmp.Data
                                'si hay empleados ciclo por cada uno y me fijo si tiene el doc de nuestro empleado
                                empleado = pair.Value
                                dni = empleado.Data("name")
                                SpecEmpId = empleado.Data("id")

                                '30/12/2013
                                Dim x As Integer
                                x = lista.IndexOf("'" & rhpro_nrotarj & "'")
                                If x <> -1 Then
                                    FLog.EscribirLinea("La tarjeta :" & rhpro_nrotarj & "ya existe en Spec, no se le inserta al empleado")
                                    ws = wsContract.Get(WSContainer.Employee, SpecEmpId)
                                    ws.Data("employeeCode") = rhpro_legajo
                                    wsContract.Set(WSContainer.Employee, ws)
                                    FLog.EscribirLinea("Se actualizo el codigo del empleado en spec")
                                Else
                                    tarjeta = wsContract.Get(WSContainer.Card, -1)
                                    tarjeta.Data("Number") = rhpro_nrotarj
                                    tarjeta.Data("employee") = SpecEmpId
                                    wsContract.Set(WSContainer.Card, tarjeta)

                                    ws = wsContract.Get(WSContainer.Employee, SpecEmpId)
                                    ws.Data("cards") = rhpro_nrotarj
                                    ws.Data("cards_dateini") = validoDesde
                                    ws.Data("cards_dateend") = validoHasta
                                    wsContract.Set(WSContainer.Employee, ws)
                                    FLog.EscribirLinea("Se inserta el empleado " & rhpro_terape & " con id " & SpecEmpId, 5)
                                    'hasta aca
                                End If
                                'hasta aca
                                Exit For
                            Next
                        End If
                    Else
                        'encontro el empleado le agrego la tarjeta
                        '30/12/2013
                        Dim x As Integer
                        x = lista.IndexOf("'" & rhpro_nrotarj & "'")
                        If x <> -1 Then
                            FLog.EscribirLinea("La tarjeta :" & rhpro_nrotarj & "ya existe en Spec, no se le inserta al empleado")
                            ws = wsContract.Get(WSContainer.Employee, SpecEmpId)
                            ws.Data("employeeCode") = rhpro_legajo
                            ws.Data("companyCode") = rhpro_nombreEmpresa
                            wsContract.Set(WSContainer.Employee, ws)
                            FLog.EscribirLinea("Se actualizo el codigo del empleado en spec")
                        Else
                            tarjeta = wsContract.Get(WSContainer.Card, -1)
                            tarjeta.Data("Number") = rhpro_nrotarj
                            tarjeta.Data("employee") = SpecEmpId
                            ws.Data("companyCode") = rhpro_nombreEmpresa
                            wsContract.Set(WSContainer.Card, tarjeta)

                            ws = wsContract.Get(WSContainer.Employee, SpecEmpId)
                            ws.Data("cards") = rhpro_nrotarj
                            ws.Data("cards_dateini") = validoDesde
                            ws.Data("cards_dateend") = validoHasta
                            wsContract.Set(WSContainer.Employee, ws)
                            FLog.EscribirLinea("se le asigna una nueva tarjeta al empleado", 5)
                            'hasta aca
                        End If
                        'hasta aca
                    End If
                End If
                'actualizo el progreso
                conexion.Open()
                da = New OleDbDataAdapter(StrSql, conexion)
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
                cmd.CommandText = StrSql
                cmd.ExecuteNonQuery()
                conexion.Close()
                'hasta aca
            Next
        End If
    End Sub

    Public Sub levantarParametros(ByRef modtipo As String)
        Dim dtTabla As New DataTable

        StrSql = "SELECT bprcparam FROM batch_proceso WHERE bpronro = " & NroProceso
        da = New OleDbDataAdapter(StrSql, conexion.ConnectionString)
        dtTabla = New DataTable
        da.Fill(dtTabla)
        If (dtTabla.Rows.Count > 0) Then
            modtipo = dtTabla.Rows(0).Item(0).ToString
        Else
            modtipo = "-1@0@0@0"
        End If
        FLog.EscribirLinea("Se obtuvieron los parametros")

    End Sub
End Module
