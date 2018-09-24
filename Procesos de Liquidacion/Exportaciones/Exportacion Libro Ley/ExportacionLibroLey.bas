Attribute VB_Name = "ExpLibroLey"
' ---------------------------------------------------------------------------------------------
' Descripcion: Proyecto encargado de generar la exportacion de Libro de Ley
' Autor      : GdeCos
' Fecha      : 1/06/2005
' Ultima Mod : 7/6/2005 - GdeCos - Se cambio el nombre de los archivos generados
' ---------------------------------------------------------------------------------------------

Option Explicit

'*************************************************************************************

Global Const Version = "1.01"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = " " 'MB - Encriptacion de string connection

'*************************************************************************************

Dim fs, f
'Global Flog

Dim NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

Global l_bpronro As Long
Global l_tipo As String
Dim l_rs As New ADODB.Recordset
Dim l_rs2 As New ADODB.Recordset
Dim l_rs3 As New ADODB.Recordset
Dim l_rs4 As New ADODB.Recordset

Dim l_Max_Lineas_X_Pag
Dim l_Max_Cols_X_Linea

Dim ArchExp

Dim l_nrolinea
Dim l_nropagina
Dim l_poner_nro_pag

Dim l_total_mbasico
Dim l_total_mneto
Dim l_total_mmsr
Dim l_total_masi_flia
Dim l_total_mDtos
Dim l_total_mbruto

Dim l_linea
Dim l_sql As String
 
Dim l_cambiaest1
Dim l_cambiaest2
Dim l_cambiaest3

 Dim l_hay_te1
 Dim l_hay_te2
 Dim l_hay_te3

 Dim l_estr1ant
 Dim l_estr2ant
 Dim l_estr3ant

 Dim l_tedabr1
 Dim l_tedabr2
 Dim l_tedabr3

 Dim l_estrdabr1
 Dim l_estrdabr2
 Dim l_estrdabr3


Const l_nro_col = 4


Private Sub Main()

Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim StrSql2 As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim fs2
Dim I As Integer
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim procesos_sep
Dim CantRegistros As Long
Dim Empresa As String
Dim Directorio As String
Dim Carpeta
Dim cambiar_Empresa As Boolean


Dim l_i

Dim l_encabezado
Dim l_corte
Dim l_cambioEmp
Dim l_conc_detdom
Dim l_listaproc
Dim l_procant
Dim l_terant
Dim l_pliqdesc
Dim l_pliqant
Dim l_pliqmesant
Dim l_pliqanioant

'Parametros
 Dim l_filtro ' Viene el filtro comun: empest, legajo,
 Dim l_tenro1
 Dim l_estrnro1
 Dim l_tenro2
 Dim l_estrnro2
 Dim l_tenro3
 Dim l_estrnro3
 Dim l_orden

 Dim l_pliqdesde  ' periodo desde (nro de periodo)
 Dim l_pliqhasta  ' periodo hasta (nro de periodo)
 Dim l_desde      ' periodo desde (fecha)
 Dim l_hasta      ' periodo hasta (fecha)
 Dim l_proaprob   ' indica estado de aprobacion del proceso
 Dim l_pronro     ' proceso o lista de procesos
 Dim l_empresa    '
 Dim l_fecestr    ' fecha para his_estructura
 Dim l_titulofiltro ' Viene el titulo armado segun filtro
 Dim l_conceptonombre ' Viene el nombre del concepto para el titulo
 Dim l_concnro

    
    l_Max_Lineas_X_Pag = 50
    l_Max_Cols_X_Linea = 80


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

    On Error GoTo CE
    
    TiempoInicialProceso = GetTickCount
    'OpenConnection strconexion, objConn
    
       
    On Error GoTo CE
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExportacionLibroLey" & "-" & NroProceso & ".log"
    
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
    
    Flog.writeline "Inicio Proceso de Exportación LibroLey : " & Now
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
        
    On Error GoTo CE
    
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs2
    
    If Not objRs2.EOF Then
       
       'Obtengo los parametros del proceso
       parametros = objRs2!bprcparam
       ArrParametros = Split(parametros, "@")
              
       'Obtengo el nro. de batch proceso
        l_bpronro = ArrParametros(0)
        Flog.writeline "batch_pronro : " & l_bpronro
       
       'Obtengo el tipo
        l_tipo = ArrParametros(1)
        Flog.writeline "Tipo: " & l_tipo
        
        'Directorio de exportacion
         StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
         OpenRecordset StrSql, rs
         If Not rs.EOF Then
             Directorio = Trim(rs!sis_dirsalidas)
         End If
         
         StrSql = "SELECT * FROM modelo WHERE modnro = 260"
         OpenRecordset StrSql, rs_Modelo
         If Not rs_Modelo.EOF Then
             If Not IsNull(rs_Modelo!modarchdefault) Then
                 Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
             Else
                 Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
             End If
         Else
             Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
         End If
         rs_Modelo.Close
         
         Flog.writeline "Se comienza a procesar los datos"

        '------------------------------------------------------------------------------------------------------------------
        'COMIENZA
        '------------------------------------------------------------------------------------------------------------------
        
        If l_tipo = "" Then
            l_tipo = 1
        Else
            l_tipo = CInt(l_tipo)
        End If
        
        'Cargo la estructura del reporte
        Dim l_msr
        Dim l_neto
        Dim l_basico
        Dim l_asi_flia
        Dim l_Dtos
        Dim l_bruto
        
        Dim l_hay_datos
        
        Dim p1
        Dim p2
        
        l_sql = " SELECT * FROM confrep WHERE repnro = 61 AND confnrocol = 8 "
        
        OpenRecordset l_sql, l_rs
    
        If Not l_rs.EOF Then
           l_Max_Lineas_X_Pag = CLng(l_rs("confval"))
        End If
        
        l_rs.Close
        
        l_nrolinea = 1
        l_nropagina = 0
        l_encabezado = True
        l_corte = False
        
        l_hay_datos = False
        
        l_total_mbasico = 0
        l_total_mneto = 0
        l_total_mmsr = 0
        l_total_masi_flia = 0
        l_total_mDtos = 0
        l_total_mbruto = 0
                
        l_sql = " SELECT * FROM rep_libroley WHERE bpronro= " & l_bpronro
        l_sql = l_sql & " ORDER BY orden "
        
        OpenRecordset l_sql, l_rs
                
            If Not l_rs.EOF Then
               If Not IsNull(l_rs("tedabr1")) Then
                  l_hay_te1 = True
                  l_estr1ant = ""
                  l_tedabr1 = ""
                  l_estrdabr1 = ""
               End If
               If Not IsNull(l_rs("tedabr2")) Then
                  l_hay_te2 = True
                  l_estr2ant = ""
                  l_tedabr2 = ""
                  l_estrdabr2 = ""
               End If
               If Not IsNull(l_rs("tedabr3")) Then
                  l_hay_te3 = True
                  l_estr3ant = ""
                  l_tedabr3 = ""
                  l_estrdabr3 = ""
               End If
            End If
            
            l_poner_nro_pag = True
            
            If Not l_rs.EOF Then
               If IsNull(l_rs("ultima_pag_impr")) Then
                  l_nropagina = 0
                  l_poner_nro_pag = False
               Else
                  l_nropagina = CLng(l_rs("ultima_pag_impr"))
                  l_poner_nro_pag = (l_nropagina <> -1)
               End If
            End If
                
            If Not l_rs.EOF Then
                CantRegistros = CLng(l_rs.RecordCount)
                                
                Empresa = CLng(l_rs!emprnro)
                                
                l_sql = "SELECT estrnro, empresa.ternro, empnom, nrodoc FROM empresa "
                l_sql = l_sql & " INNER JOIN ter_doc ON ter_doc.ternro = empresa.ternro AND ter_doc.tidnro = 6 "
                l_sql = l_sql & " WHERE estrnro = " & Empresa
                
                OpenRecordset l_sql, objRs
                                
                Nombre_Arch = Directorio & "\" & CStr(objRs!nrodoc) & "_Libro Ley.txt"
                                
                objRs.Close
                                                                
                Flog.writeline "Se crea el archivo: " & Nombre_Arch
                ' Se Crea el archivo con nombre de la empresa y estrnro
                Set fs2 = CreateObject("Scripting.FileSystemObject")
                'Activo el manejador de errores
                On Error Resume Next
                If Err.Number <> 0 Then
                    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
                    Set Carpeta = fs2.CreateFolder(Directorio)
                End If
                'desactivo el manejador de errores
                Set ArchExp = fs2.CreateTextFile(Nombre_Arch, True)
                On Error GoTo CE
            
            
            End If
    
            I = 1
            cambiar_Empresa = False
                                 
            Do Until l_rs.EOF
                
                    If Empresa <> CLng(l_rs!emprnro) Then
                        cambiar_Empresa = True
                        Empresa = CLng(l_rs!emprnro)
                        
                        ArchExp.Close
                        
                        l_sql = "SELECT estrnro, empresa.ternro, empnom, nrodoc FROM empresa "
                        l_sql = l_sql & " INNER JOIN ter_doc ON ter_doc.ternro = empresa.ternro AND ter_doc.tidnro = 6 "
                        l_sql = l_sql & " WHERE estrnro = " & Empresa
                        
                        OpenRecordset l_sql, objRs
                                        
                        Nombre_Arch = Directorio & "\" & CStr(objRs!nrodoc) & "_Libro Ley.txt"
                                        
                        objRs.Close
                                        
                        Flog.writeline "Se crea el archivo: " & Nombre_Arch
                        ' Se Crea el archivo con nombre de la empresa y estrnro
                        Set fs2 = CreateObject("Scripting.FileSystemObject")
                        'Activo el manejador de errores
                        On Error Resume Next
                        If Err.Number <> 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
                            Set Carpeta = fs2.CreateFolder(Directorio)
                        End If
                        'desactivo el manejador de errores
                        Set ArchExp = fs2.CreateTextFile(Nombre_Arch, True)
                        On Error GoTo CE
                        
                    End If
    
                    l_sql = " SELECT * FROM rep_libroley_det "
                    l_sql = l_sql & " WHERE bpronro = " & l_bpronro
                    l_sql = l_sql & "   AND pronro  = " & l_rs("pronro")
                    l_sql = l_sql & "   AND ternro  = " & l_rs("ternro")
                
                    OpenRecordset l_sql, l_rs2
                    
                    l_cambioEmp = True
                
                    If Not l_rs2.EOF Then
                       If l_pliqant <> CLng(l_rs("pliqnro")) Then
                          l_corte = True
                          l_encabezado = True
                          l_cambioEmp = True
                          l_nropagina = l_nropagina + 1
                       End If
                
                       l_procant = CLng(l_rs("pronro"))
                       l_pliqant = CLng(l_rs("pliqnro"))
                       l_pliqmesant = CInt(l_rs("pliqmes"))
                       l_pliqanioant = CInt(l_rs("pliqanio"))
                    End If
                
                    Do Until l_rs2.EOF
                        If l_encabezado Then
                
                            l_procant = CLng(l_rs("pronro"))
                            l_pliqant = CLng(l_rs("pliqnro"))
                            l_pliqmesant = CInt(l_rs("pliqmes"))
                            l_pliqanioant = CInt(l_rs("pliqanio"))
                
                            If l_corte And l_hay_datos Then
                                For l_i = l_nrolinea To l_Max_Lineas_X_Pag
                                    imprNVeces " ", l_Max_Cols_X_Linea
                                Next
                                l_cambioEmp = True
                                ArchExp.Write Chr(12)
                                l_nrolinea = 1
                            End If
                
                            encabezado
                
                            l_nrolinea = l_nrolinea + 1
                
                        End If
                
                        l_hay_datos = True
                        l_nrolinea = l_nrolinea + 1
                        If l_cambioEmp Then
                            If cambioEstructura(True) Then
                
                                If l_tipo = 1 Or l_tipo = 3 Then
                                    If l_hay_te1 And l_cambiaest1 Then
                                        borrarLinea l_linea
                                        addStrLinea l_linea, l_rs("tedabr1") & " :" & l_rs("estrdabr1"), 1, 0
                                        imprLinea l_linea
                                        l_nrolinea = l_nrolinea + 1
                                        l_tedabr1 = l_rs("tedabr1")
                                        l_estrdabr1 = l_rs("estrdabr1")
                                    End If
                                    If l_hay_te2 And l_cambiaest2 Then
                                        borrarLinea l_linea
                                        addStrLinea l_linea, "** " & l_rs("tedabr2") & " :" & l_rs("estrdabr2"), 1, 0
                                        imprLinea l_linea
                                        l_nrolinea = l_nrolinea + 1
                                        l_tedabr2 = l_rs("tedabr2")
                                        l_estrdabr2 = l_rs("estrdabr2")
                                    End If
                                    If l_hay_te3 And l_cambiaest3 Then
                                        borrarLinea l_linea
                                        addStrLinea l_linea, "**** " & l_rs("tedabr3") & " :" & l_rs("estrdabr3"), 1, 0
                                        imprLinea l_linea
                                        l_nrolinea = l_nrolinea + 1
                                        l_tedabr3 = l_rs("tedabr3")
                                        l_estrdabr3 = l_rs("estrdabr3")
                                    End If
                                Else
                                    If l_hay_te1 And l_cambiaest1 Then
                                        imprNVeces " ", l_Max_Cols_X_Linea
                                        l_nrolinea = l_nrolinea + 1
                                    End If
                                    If l_hay_te2 And l_cambiaest2 Then
                                        imprNVeces " ", l_Max_Cols_X_Linea
                                        l_nrolinea = l_nrolinea + 1
                                    End If
                                    If l_hay_te3 And l_cambiaest3 Then
                                        imprNVeces " ", l_Max_Cols_X_Linea
                                        l_nrolinea = l_nrolinea + 1
                                    End If
                                End If
                
                            End If
                
                            titulo_empleado l_rs("ternro"), l_rs("legajo"), l_rs("apellido"), l_rs("apellido2"), l_rs("nombre"), l_rs("nombre2"), l_rs("prodesc"), l_rs("pliqdesc"), l_rs("profecpago")
                            l_nrolinea = l_nrolinea + 1
                            l_cambioEmp = False
                        End If
                
                        If l_tipo = 1 Or l_tipo = 3 Then
                            borrarLinea l_linea
                
                            addStrLinea l_linea, l_rs2("conccod"), 1, 0
                            addStrLinea l_linea, l_rs2("concabr"), 10, 0
                            addStrLinea l_linea, l_rs2("dlicant"), 60, 1
                            addStrLinea l_linea, FormatNumber(CDbl(l_rs2("dlimonto")), 2), 70, 1
                
                            imprLinea l_linea
                        Else
                            borrarLinea l_linea
                            imprLinea l_linea
                        End If
                
                        If l_nrolinea > l_Max_Lineas_X_Pag Then
                            l_corte = True
                            l_encabezado = True
                            l_cambioEmp = True
                            l_nropagina = l_nropagina + 1
                        Else
                            l_encabezado = False
                        End If
                
                        l_rs2.MoveNext
                
                        If Not l_rs2.EOF Then
                           l_cambioEmp = CLng(l_procant) <> CLng(l_rs("pronro"))
                           l_procant = CLng(l_rs("pronro"))
                           l_pliqant = CLng(l_rs("pliqnro"))
                           l_pliqmesant = CInt(l_rs("pliqmes"))
                           l_pliqanioant = CInt(l_rs("pliqanio"))
                        Else
                           l_cambioEmp = True
                           If (l_Max_Lineas_X_Pag - l_nrolinea) < 12 Then
                               l_corte = True
                               l_encabezado = True
                               l_nropagina = l_nropagina + 1
                           End If
                        End If
                
                        If l_cambioEmp Then
                           'Cierro la tabla que habri en la funcion titulo_empleado
                           If l_tipo = 1 Or l_tipo = 3 Then
                               imprNVeces "-", l_Max_Cols_X_Linea
                           Else
                               imprNVeces " ", l_Max_Cols_X_Linea
                           End If
                
                           'Muestro los acumuladores del proceso indicados en el confrep.
                           mostrarAcumuladores l_rs("ternro"), l_procant, l_pliqmesant, l_pliqanioant
                
                           If l_tipo = 1 Or l_tipo = 3 Then
                               imprNVeces "=", l_Max_Cols_X_Linea
                           Else
                               imprNVeces " ", l_Max_Cols_X_Linea
                           End If
                
                
                        End If
                    Loop
                    l_rs2.Close
                
                    l_rs.MoveNext
                
                    If cambioEstructura(False) Then
                        If l_tipo = 1 Or l_tipo = 3 Then
                            If l_hay_te3 And l_cambiaest3 Then
                                borrarLinea l_linea
                                addStrLinea l_linea, "**** " & l_tedabr3 & " :" & l_estrdabr3, 1, 0
                                imprLinea l_linea
                                l_nrolinea = l_nrolinea + 1
                            End If
                            If l_hay_te2 And l_cambiaest2 Then
                                borrarLinea l_linea
                                addStrLinea l_linea, "** " & l_tedabr2 & " :" & l_estrdabr2, 1, 0
                                imprLinea l_linea
                                l_nrolinea = l_nrolinea + 1
                            End If
                            If l_hay_te1 And l_cambiaest1 Then
                                borrarLinea l_linea
                                addStrLinea l_linea, l_tedabr1 & " :" & l_estrdabr1, 1, 0
                                imprLinea l_linea
                                l_nrolinea = l_nrolinea + 1
                            End If
                
                        Else
                            If l_hay_te3 And l_cambiaest3 Then
                                imprNVeces " ", l_Max_Cols_X_Linea
                                l_nrolinea = l_nrolinea + 1
                            End If
                            If l_hay_te2 And l_cambiaest2 Then
                                imprNVeces " ", l_Max_Cols_X_Linea
                                l_nrolinea = l_nrolinea + 1
                            End If
                            If l_hay_te1 And l_cambiaest1 Then
                                imprNVeces " ", l_Max_Cols_X_Linea
                                l_nrolinea = l_nrolinea + 1
                            End If
                        End If
                
                        l_corte = True
                        l_encabezado = True
                        l_cambioEmp = True
                        l_nropagina = l_nropagina + 1
                    End If
                
                    If cambiar_Empresa Then
                    
                        If Not l_hay_datos Then
                            ArchExp.writeline "No se encontraron datos"
                            Flog.writeline "No se encontraron datos a Exportar para esta empresa."
                        Else
                        
                            p1 = 25
                            p2 = 55
                            
                            If l_tipo = 1 Or l_tipo = 3 Then
                            
                                borrarLinea l_linea
                                imprLinea l_linea
                            
                                borrarLinea l_linea
                                addStrLinea l_linea, "Totales Generales del Período:", 20, 0
                                imprLinea l_linea
                            
                                borrarLinea l_linea
                                imprLinea l_linea
                            
                                borrarLinea l_linea
                                addStrLinea l_linea, "Bruto:", p1, 0
                                addStrLinea l_linea, FormatNumber(CDbl(l_total_mbruto), 2), p2, 1
                                imprLinea l_linea
                            
                                borrarLinea l_linea
                                addStrLinea l_linea, "Rem. Básica:", p1, 0
                                addStrLinea l_linea, FormatNumber(CDbl(l_total_mbasico), 2), p2, 1
                                imprLinea l_linea
                            
                                borrarLinea l_linea
                                addStrLinea l_linea, "Sal. Familiar:", p1, 0
                                addStrLinea l_linea, FormatNumber(CDbl(l_total_masi_flia), 2), p2, 1
                                imprLinea l_linea
                            
                                borrarLinea l_linea
                                addStrLinea l_linea, "Monto Suj. Ret.:", p1, 0
                                addStrLinea l_linea, FormatNumber(CDbl(l_total_mmsr), 2), p2, 1
                                imprLinea l_linea
                            
                                borrarLinea l_linea
                                addStrLinea l_linea, "Descuentos:", p1, 0
                                addStrLinea l_linea, FormatNumber(CDbl(l_total_mDtos), 2), p2, 1
                                imprLinea l_linea
                            
                                borrarLinea l_linea
                                addStrLinea l_linea, "Neto Pagado:", p1, 0
                                addStrLinea l_linea, FormatNumber(CDbl(l_total_mneto), 2), p2, 1
                                imprLinea l_linea
                            
                            End If
                        
                        End If
    
                        cambiar_Empresa = False
                        
                    End If
                
                    TiempoAcumulado = GetTickCount

                    I = I + 1
                
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix((I / CantRegistros) * 100) & _
                             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                             " WHERE bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                
                Loop
                
                If Not l_hay_datos Then
                    ArchExp.writeline "No se encontraron datos"
                    Flog.writeline "No se encontraron datos a Exportar para esta empresa."
                Else
                
                    p1 = 25
                    p2 = 55
                    
                    If l_tipo = 1 Or l_tipo = 3 Then
                    
                        borrarLinea l_linea
                        imprLinea l_linea
                    
                        borrarLinea l_linea
                        addStrLinea l_linea, "Totales Generales del Período:", 20, 0
                        imprLinea l_linea
                    
                        borrarLinea l_linea
                        imprLinea l_linea
                    
                        borrarLinea l_linea
                        addStrLinea l_linea, "Bruto:", p1, 0
                        addStrLinea l_linea, FormatNumber(CDbl(l_total_mbruto), 2), p2, 1
                        imprLinea l_linea
                    
                        borrarLinea l_linea
                        addStrLinea l_linea, "Rem. Básica:", p1, 0
                        addStrLinea l_linea, FormatNumber(CDbl(l_total_mbasico), 2), p2, 1
                        imprLinea l_linea
                    
                        borrarLinea l_linea
                        addStrLinea l_linea, "Sal. Familiar:", p1, 0
                        addStrLinea l_linea, FormatNumber(CDbl(l_total_masi_flia), 2), p2, 1
                        imprLinea l_linea
                    
                        borrarLinea l_linea
                        addStrLinea l_linea, "Monto Suj. Ret.:", p1, 0
                        addStrLinea l_linea, FormatNumber(CDbl(l_total_mmsr), 2), p2, 1
                        imprLinea l_linea
                    
                        borrarLinea l_linea
                        addStrLinea l_linea, "Descuentos:", p1, 0
                        addStrLinea l_linea, FormatNumber(CDbl(l_total_mDtos), 2), p2, 1
                        imprLinea l_linea
                    
                        borrarLinea l_linea
                        addStrLinea l_linea, "Neto Pagado:", p1, 0
                        addStrLinea l_linea, FormatNumber(CDbl(l_total_mneto), 2), p2, 1
                        imprLinea l_linea
                    
                    End If
                
                End If
                   
                Flog.writeline "Se Terminaron de Procesar los datos"
                
                l_rs.Close
                ArchExp.Close
        
                   
                   
'               Flog.writeline "Se escribe el registro Nro.: " & i & " de la Empresa: " & Empresa

       
    
    Else

       Exit Sub

    End If
    
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
    If objRs2.State = adStateOpen Then objRs2.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

Function cambiarFormato(cadena, separador)

 Dim l_salida
 
 l_salida = cadena

 If (InStr(cadena, ".")) Then
    If separador = "," Then
       l_salida = Replace(cadena, ".", ",")
    End If
 Else
     If (InStr(cadena, ",")) Then
        If separador = "." Then
           l_salida = Replace(cadena, ",", ".")
        End If
     End If
 End If

 cambiarFormato = l_salida

End Function 'cambiarFormato(cadena,separador)

Sub imprimirTexto(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 0
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    u = CInt(Longitud) - longTexto
    If u < 0 Then
        archivo.Write cadena
    Else
        If derecha Then
            archivo.Write cadena & String(u, " ")
        Else
            archivo.Write String(u, " ") & cadena
        End If
    End If

End Sub

        '----------------------------------------------------------------------------------------------------------------
        ' SUB: imprime un caracter N veces
        '----------------------------------------------------------------------------------------------------------------
        Sub imprNVeces(caracter, veces)
          Dim linea
          linea = String(veces, caracter)
          imprLinea (linea)
        End Sub 'borrarLinea(ByRef linea)
        
        '----------------------------------------------------------------------------------------------------------------
        ' SUB: borra un linea
        '----------------------------------------------------------------------------------------------------------------
        Sub borrarLinea(ByRef linea)
          linea = String(l_Max_Cols_X_Linea, " ")
        End Sub 'borrarLinea(ByRef linea)
        
        '----------------------------------------------------------------------------------------------------------------
        ' SUB: imprime una linea
        '----------------------------------------------------------------------------------------------------------------
        Sub imprLinea(linea)
          ArchExp.writeline linea
        End Sub 'imprLinea(linea)
        
        '----------------------------------------------------------------------------------------------------------------
        ' SUB: agrega un string a una linea en una posicion y con una alineacion
        '----------------------------------------------------------------------------------------------------------------
        Sub addStrLinea(ByRef linea, ByVal Str, ByVal posicion, ByVal alineacion)
          Dim nuevo
          Dim pos
        
          If Not IsNull(Str) Then
        
              Select Case alineacion
                Case 0 'Inserto a la derecha
                   nuevo = Mid(linea, 1, posicion)
                   nuevo = nuevo & CStr(Str)
                   pos = CInt(posicion) + Len(Str)
                   If CInt(l_Max_Cols_X_Linea - pos) > 0 Then
                      nuevo = nuevo & Right(CStr(linea), CInt(l_Max_Cols_X_Linea - pos))
                   End If
                   linea = Mid(CStr(nuevo), 1, l_Max_Cols_X_Linea)
                Case 1 'Inserto a la izquierda
                   posicion = posicion - Len(Str)
                   If posicion < 0 Then
                      posicion = 0
                   End If
        
                   nuevo = Mid(linea, 1, posicion)
                   nuevo = nuevo & Str
                   nuevo = nuevo & Mid(linea, posicion + Len(Str), 200)
                   linea = Mid(nuevo, 1, l_Max_Cols_X_Linea)
              End Select
          End If
        End Sub 'addStrLinea(ByRef linea,str,posicion,alineacion)
        
        'Muestra al final de la pagina el basico,neto,msr
        Sub mostrarAcumuladores(ternro, procnro, pliqmes, pliqanio)
        Dim l_mbasico
        Dim l_mneto
        Dim l_mmsr
        Dim l_masi_flia
        Dim l_mDtos
        Dim l_mbruto
        
        l_nrolinea = l_nrolinea + 3
        
        l_mbasico = l_rs("basico")
        l_mneto = l_rs("neto")
        l_mmsr = l_rs("msr")
        l_masi_flia = l_rs("asi_flia")
        l_mDtos = l_rs("dtos")
        l_mbruto = l_rs("bruto")
        
        l_total_mbasico = l_total_mbasico + CDbl(l_mbasico)
        l_total_mneto = l_total_mneto + CDbl(l_mneto)
        l_total_mmsr = l_total_mmsr + CDbl(l_mmsr)
        l_total_masi_flia = l_total_masi_flia + CDbl(l_masi_flia)
        l_total_mDtos = l_total_mDtos + CDbl(l_mDtos)
        l_total_mbruto = l_total_mbruto + CDbl(l_mbruto)
        
        borrarLinea l_linea
        If l_tipo = 1 Or l_tipo = 3 Then
            addStrLinea l_linea, "Bruto:", 1, 0
            addStrLinea l_linea, FormatNumber(CDbl(l_mbruto), 2), 22, 1
            addStrLinea l_linea, "T.Sal.Familiar:", 25, 0
            addStrLinea l_linea, FormatNumber(CDbl(l_masi_flia), 2), 54, 1
            addStrLinea l_linea, "MSR:", 57, 0
            addStrLinea l_linea, FormatNumber(CDbl(l_mmsr), 2), 78, 1
        End If
        imprLinea l_linea
        
        borrarLinea l_linea
        If l_tipo = 1 Or l_tipo = 3 Then
            addStrLinea l_linea, "Básico:", 1, 0
            addStrLinea l_linea, FormatNumber(CDbl(l_mbasico), 2), 22, 1
            addStrLinea l_linea, "T.Descuentos:", 25, 0
            addStrLinea l_linea, FormatNumber(CDbl(l_mDtos), 2), 54, 1
            addStrLinea l_linea, "Neto:", 57, 0
            addStrLinea l_linea, FormatNumber(CDbl(l_mneto), 2), 78, 1
        End If
        imprLinea l_linea
        
        End Sub 'mostrarAcumuladores(ternro,procnro,pliqmes,pliqanio)
        
        ' Imprime el encabezado de cada pagina
        Sub encabezado()
            Dim f_str
            Dim l_empcuit
            Dim l_empdire
            Dim l_empactiv
        
            l_nrolinea = l_nrolinea + 7
        
            If IsNull(l_rs("auxchar1")) Then
              l_empcuit = ""
            Else
              l_empcuit = l_rs("auxchar1")
            End If
        
            If IsNull(l_rs("auxchar2")) Then
              l_empdire = ""
            Else
              l_empdire = l_rs("auxchar2")
            End If
        
            If IsNull(l_rs("auxchar3")) Then
              l_empactiv = ""
            Else
              l_empactiv = l_rs("auxchar3")
            End If
        
            If l_tipo = 1 Or l_tipo = 2 Then
                imprNVeces "*", l_Max_Cols_X_Linea
        
                borrarLinea l_linea
                If l_poner_nro_pag Then
                   addStrLinea l_linea, l_rs("empresa"), 0, 0
                   addStrLinea l_linea, "Página:" & l_nropagina, l_Max_Cols_X_Linea - 15, 0
                End If
                imprLinea l_linea
        
                borrarLinea l_linea
                If l_poner_nro_pag Then
                   addStrLinea l_linea, l_empcuit, 0, 0
                End If
                imprLinea l_linea
        
                borrarLinea l_linea
                If l_poner_nro_pag Then
                   addStrLinea l_linea, l_empdire, 0, 0
                End If
                imprLinea l_linea
        
                borrarLinea l_linea
                If l_poner_nro_pag Then
                   addStrLinea l_linea, l_empactiv, 0, 0
                End If
                imprLinea l_linea
        
                f_str = "Libro Ley - Articulo 52 - Ley 20744"
                borrarLinea l_linea
                If l_poner_nro_pag Then
                   addStrLinea l_linea, f_str, (l_Max_Cols_X_Linea / 2) - (Len(f_str) / 2), 0
                End If
                imprLinea l_linea
        
                f_str = "Periodo de Liquidación: " & l_rs("pliqdesc")
                borrarLinea l_linea
                addStrLinea l_linea, f_str, (l_Max_Cols_X_Linea / 2) - (Len(f_str) / 2), 0
                imprLinea l_linea
        
                f_str = "Fecha de Pago: " & l_rs("pliqfecdep") & "      Banco: " & l_rs("pliqbco")
                borrarLinea l_linea
                addStrLinea l_linea, f_str, (l_Max_Cols_X_Linea / 2) - (Len(f_str) / 2), 0
                imprLinea l_linea
        
                imprNVeces "*", l_Max_Cols_X_Linea
        
            Else
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
            End If
        End Sub 'encabezado

        'Imprime el titulo de la seccion de cada empleado
        Sub titulo_empleado(ternro, Legajo, apellido, apellido2, nombre, nombre2, Proceso, periodo, fecpago)
        
         Dim l_fecalta
         Dim l_fecbaja
         Dim l_contrato
         Dim l_categoria
        
         Dim l_direccion
         Dim l_puesto
         Dim l_documento
         Dim l_fecha_nac
         Dim l_est_civil
         Dim l_cuil
         Dim l_estado
         Dim l_reg_prev
         Dim l_lug_trab
         Dim l_est_texto
        
         l_nrolinea = l_nrolinea + 11
        
            l_contrato = l_rs("contrato")
            l_categoria = controlNull(l_rs("categoria"))
            l_puesto = l_rs("puesto")
            If IsNull(l_puesto) Then
               l_puesto = ""
            Else
               If Len(l_puesto) > 15 Then
                  l_puesto = Mid(l_puesto, 1, 15)
               End If
            End If
        
            l_reg_prev = l_rs("reg_prev")
            l_lug_trab = l_rs("lug_trab")
            l_estado = l_rs("estado")
            l_direccion = l_rs("direccion")
            l_cuil = l_rs("cuil")
            l_documento = l_rs("documento")
            l_fecha_nac = l_rs("fecha_nac")
            l_est_civil = l_rs("est_civil")
            l_fecalta = l_rs("fecalta")
            l_fecbaja = l_rs("fecbaja")
        
            If l_tipo = 1 Or l_tipo = 3 Then
        
                imprNVeces "=", l_Max_Cols_X_Linea
                borrarLinea l_linea
                addStrLinea l_linea, Legajo & " - " & apellido & " " & apellido2 & ", " & nombre & " " & nombre2, 1, 0
                imprLinea l_linea
                imprNVeces "-", l_Max_Cols_X_Linea
        
                borrarLinea l_linea
                addStrLinea l_linea, "Categoria:", 1, 0
                addStrLinea l_linea, l_categoria, 12, 0
                addStrLinea l_linea, "Puesto:", 26, 0
                addStrLinea l_linea, l_puesto, 34, 0
                addStrLinea l_linea, "Est.Civil:", 50, 0
                addStrLinea l_linea, l_est_civil, 65, 0
                imprLinea l_linea
        
                borrarLinea l_linea
                addStrLinea l_linea, "Documento:", 1, 0
                addStrLinea l_linea, l_documento, 12, 0
                addStrLinea l_linea, "Nacimiento:", 26, 0
                addStrLinea l_linea, l_fecha_nac, 38, 0
                addStrLinea l_linea, "Lug. de Trab.:", 50, 0
                addStrLinea l_linea, l_lug_trab, 65, 0
                imprLinea l_linea
        
                borrarLinea l_linea
                addStrLinea l_linea, "CUIL:", 1, 0
                addStrLinea l_linea, l_cuil, 12, 0
                addStrLinea l_linea, "Ingreso:", 26, 0
                addStrLinea l_linea, l_fecalta, 38, 0
                addStrLinea l_linea, "Egreso:", 50, 0
                addStrLinea l_linea, l_fecbaja, 65, 0
                imprLinea l_linea
        
                borrarLinea l_linea
                addStrLinea l_linea, "Contrato:", 1, 0
                addStrLinea l_linea, l_contrato, 12, 0
                addStrLinea l_linea, "Régimen Previsional:", 26, 0
                addStrLinea l_linea, l_reg_prev, 47, 0
                imprLinea l_linea
        
                If CInt(l_estado) = -1 Then
                   l_est_texto = "Activo"
                Else
                   l_est_texto = "Inactivo"
                End If
        
                borrarLinea l_linea
                addStrLinea l_linea, "Estado:", 1, 0
                addStrLinea l_linea, l_est_texto, 12, 0
                addStrLinea l_linea, "Dirección:", 26, 0
                addStrLinea l_linea, l_direccion, 38, 0
                imprLinea l_linea
        
                borrarLinea l_linea
                addStrLinea l_linea, "Proceso:", 1, 0
                addStrLinea l_linea, Proceso, 10, 0
                imprLinea l_linea
        
                impr_familiares ternro
        
                imprNVeces "-", l_Max_Cols_X_Linea
        
                borrarLinea l_linea
                addStrLinea l_linea, "Concepto", 1, 0
                addStrLinea l_linea, "Descripción", 10, 0
                addStrLinea l_linea, "Parámetro", 60, 1
                addStrLinea l_linea, "Importe", 70, 1
                imprLinea l_linea
            Else
        
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
        
                Call impr_familiares(ternro)
        
                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea
            End If
        
        End Sub 'titulo_empleado
        
        Function controlNull(Str)
          If IsNull(Str) Then
             controlNull = " "
          Else
             controlNull = Str
          End If
        End Function 'controlNull(str)



'Imprime los familiares del empleado
Sub impr_familiares(ternro)

Dim l_hd
    
    l_hd = False

    l_sql = " SELECT * FROM rep_libroley_fam "
    l_sql = l_sql & " WHERE bpronro = " & l_bpronro
    l_sql = l_sql & "   AND pronro  = " & l_rs("pronro")
    l_sql = l_sql & "   AND ternro  = " & l_rs("ternro")


    OpenRecordset l_sql, l_rs3

Do Until l_rs3.EOF

        If Not l_hd Then
            l_hd = True
            l_nrolinea = l_nrolinea + 2

            If l_tipo = 1 Or l_tipo = 3 Then

                imprNVeces "-", l_Max_Cols_X_Linea
                borrarLinea l_linea
                addStrLinea l_linea, "Apellido y Nombre", 1, 0
                addStrLinea l_linea, "Parentesco", 20, 0
                addStrLinea l_linea, "Documento", 31, 0
                addStrLinea l_linea, "Fec. Nac.", 44, 0
                addStrLinea l_linea, "Fec. Cargo", 55, 0
                addStrLinea l_linea, "Sexo", 66, 0
                addStrLinea l_linea, "Inc.", 71, 0
                addStrLinea l_linea, "Trab", 76, 0
                imprLinea l_linea

            Else

                imprNVeces " ", l_Max_Cols_X_Linea
                imprNVeces " ", l_Max_Cols_X_Linea

            End If

        End If

        If l_tipo = 1 Or l_tipo = 3 Then
            borrarLinea l_linea
            addStrLinea l_linea, l_rs3("terape") & ", " & l_rs3("ternom"), 1, 0
            addStrLinea l_linea, controlNull(l_rs3("paredesc")), 20, 0
            addStrLinea l_linea, controlNull(l_rs3("sigladoc")) & "-" & controlNull(l_rs3("nrodoc")), 31, 0
            addStrLinea l_linea, controlNull(l_rs3("terfecnac")), 44, 0
            addStrLinea l_linea, controlNull(l_rs3("famDGIdesde")), 55, 0

            If CInt(l_rs3("tersex")) = -1 Then
               addStrLinea l_linea, "Mas.", 66, 0
            Else
               addStrLinea l_linea, "Fem.", 66, 0
            End If

            If CInt(l_rs3("faminc")) = -1 Then
               addStrLinea l_linea, "S", 71, 0
            Else
               addStrLinea l_linea, "N", 71, 0
            End If

            If CInt(l_rs3("famtrab")) = -1 Then
               addStrLinea l_linea, "S", 76, 0
            Else
               addStrLinea l_linea, "N", 76, 0
            End If

            imprLinea l_linea
        Else
            imprNVeces " ", l_Max_Cols_X_Linea
        End If

   l_nrolinea = l_nrolinea + 1
   l_rs3.MoveNext
Loop

l_rs3.Close

End Sub 'impr_familiares(ternro)


Function cambioEstructura(asigna)

l_cambiaest1 = False
l_cambiaest2 = False
l_cambiaest3 = False

If Not l_rs.EOF Then
    If l_hay_te1 Then
        If l_estr1ant <> l_rs("estrdabr1") Then
            l_cambiaest1 = True
            l_cambiaest2 = True
            l_cambiaest3 = True
            If asigna Then
               l_estr1ant = l_rs("estrdabr1")
            End If
        End If
    End If

    If l_hay_te2 Then
        If l_estr2ant <> l_rs("estrdabr2") Then
            l_cambiaest2 = True
            l_cambiaest3 = True
            If asigna Then
               l_estr2ant = l_rs("estrdabr2")
            End If
        End If
    End If

    If l_hay_te3 Then
        If l_estr3ant <> l_rs("estrdabr3") Then
            l_cambiaest3 = True
            If asigna Then
               l_estr3ant = l_rs("estrdabr3")
            End If
        End If
    End If
Else
    l_cambiaest1 = True
    l_cambiaest2 = True
    l_cambiaest3 = True
End If
cambioEstructura = (l_cambiaest1 Or l_cambiaest2 Or l_cambiaest3)

End Function 'cambioEstructura()

