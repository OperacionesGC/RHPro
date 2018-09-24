Attribute VB_Name = "ExpSIJP"
' ---------------------------------------------------------------------------------------------
' Descripcion: Proyecto encargado de generar la exportacion de SIJP
' Autor      : FGZ
' Fecha      : 16/11/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "16/11/2005"
Global Const UltimaModificacion = " " 'Version Inicial

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f
Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean


Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial exportacion de SIJP.
' Autor      : FGZ
' Fecha      : 16/11/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim Pliqnro As Long
Dim Lista_Pronro As String
Dim Sep As String
Dim PID As String
Dim Parametros As String
Dim ArrParametros

    
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.Writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.Writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExportacionSIJP" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.Writeline "Inicio Proceso de Exportación SIJP : " & Now
    Flog.Writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline "-----------------------------------------------------------------"
    Flog.Writeline "Version = " & Version
    Flog.Writeline "Modificacion = " & UltimaModificacion
    Flog.Writeline "Fecha = " & FechaModificacion
    Flog.Writeline "-----------------------------------------------------------------"
    Flog.Writeline
    Flog.Writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 53"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
              
       'Obtengo el Periodo
        Pliqnro = ArrParametros(0)
        Flog.Writeline Espacios(Tabulador * 0) & "Periodo: " & Pliqnro
       
       'Obtengo la lista de procesos
        Lista_Pronro = ArrParametros(1)
        Flog.Writeline Espacios(Tabulador * 0) & "Procesos: " & Lista_Pronro
        
       'Obtengo los datos del separador
        Sep = ArrParametros(2)
        Flog.Writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Sep
       
       Call Generar_Archivo_SIJP(Pliqnro, Lista_Pronro, Sep)
    Else
        Flog.Writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.Writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.Writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    TiempoFinalProceso = GetTickCount
    Flog.Writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.Writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    objconnProgreso.Close
    objConn.Close
Exit Sub
    
ME_Main:
    HuboErrores = True
    Flog.Writeline "Error: " & Err.Description
    Flog.Writeline "Ultimo SQL: " & StrSql
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
        longTexto = 1
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    u = CInt(Longitud) - longTexto
    If u <= 0 Then
        archivo.Write cadena
    Else
        If derecha Then
            archivo.Write cadena & String(u, " ")
        Else
            archivo.Write String(u, " ") & cadena
        End If
    End If

End Sub
Sub imprimirTextoConCeros(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 1
        cadena = "0"
    Else
        longTexto = Len(Texto)
        cadena = CStr(Texto)
    End If
    
    u = CInt(Longitud) - longTexto
    If u <= 0 Then
        archivo.Write cadena
    Else
        If derecha Then
            archivo.Write cadena & String(u, "0")
        Else
            archivo.Write String(u, "0") & cadena
        End If
    End If

End Sub



Private Sub Generar_Archivo_SIJP(ByVal Pliqnro As Long, ByVal Lista_Pronro As String, ByVal Sep As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion de SIJP
' Autor      : FGZ
' Fecha      : 16/11/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Formato
'------------------------------------------------------------
'Nro Campo                                  Desde   Longitud
'------------------------------------------------------------
'2   Apellido y nombre                      0       0
'9   Porcentaje de aporte adicional SS      0       0
'17  importe adicional OS                   0       0
'18  Importe excedente aporte SS            0       0
'19  Importe excedente aporte OS            0       0
'20  provincia localidad                    0       0
'25  marca de corresponde reduccion         0       0
'26  capital de recomposicion de LRT        0       0
'28  aporte adicional de Obra social        0       0
'44  remuneracion imponible 6               0       0
'45  Tipo de operación                      0       0
'1   CUIL                                   1       11
'    Espacios                                       2
'11  codigo de obra social                  14      6
'4   cantidad de hijos                      20      2
'3   conyuge                                22      1
'12  cantidad de adeherentes                23      2
'10  codigo de modalidad de contratacion    25      3
'8   codigo de zona                         28      2
'7   codigo de actividad                    30      2
'6   codigo de condicion                    32      2
'5   Codigo de situacion                    34      2
'13  remuneracion total                     36      9
'14  remuneracion imponible 1               45      9
'    Espacios                                       2
'16  importe aporte voluntario              56      9
'15  asignaciones fliares pagadas           65      9
'21  remuneracion imponible 2               74      9
'    Espacios                                       14
'24  codigo de siniestrado                  97      2
'    Espacios                                       9
'22  remuneracion imponible 3               108     9
'23  remuneracion imponible 4               117     9
'27  tipo de empresa                        126     1
'29  regimen                                127     1
'42  remuneracion imponible 5               128     9
'43  Trabajador convencionado 0-NO 1-SI     137     1
'30  situacon de revista 1                  138     2
'31  dia de inicio sit. De revista 1        140     2
'32  situacon de revista 2                  142     2
'33  dia de inicio sit. De revista 2        144     2
'34  situacon de revista 3                  146     2
'35  dia de inicio sit. De revista 3        148     2
'36  sueldo + adicionales                   150     9
'37  sac                                    159     9
'38  horas extra                            168     9
'39  zona desfavorable                      177     9
'40  vacaciones                             186     9
'41  cantidad de dias trabajados            195     9
'------------------------------------------------------------
'TOTAL                                              203 caracteres
'------------------------------------------------------------
Dim NroModelo As Long
Dim Directorio As String
Dim Carpeta

Dim Nombre_Arch As String
Dim fs1
Dim ArchExp
Dim i As Integer
Dim CantRegistros As Long
Dim Empresa As String

Dim rs As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

    NroModelo = 258
    
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
            Flog.Writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
        End If
    Else
        Flog.Writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo & ". El archivo será generado en el directorio default"
    End If
                
    Nombre_Arch = Directorio & "\edenor_sijp.txt"
    Flog.Writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If Err.Number <> 0 Then
        Flog.Writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Set Carpeta = fs1.CreateFolder(Directorio)
    End If
    'desactivo el manejador de errores
    Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
    
    On Error GoTo ME_Local


    '------------------------------------------------------------------------------------------------------------------------
    'BUSCO LOS DATOS
    '------------------------------------------------------------------------------------------------------------------------
    StrSql = " SELECT repsijp.*, empresa.empnom, periodo.pliqdesc, empresa.ternro, ter_doc.nrodoc "
    StrSql = StrSql & " FROM repsijp "
    StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = repsijp.empresa "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = repsijp.pliqnro"
    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empresa.ternro AND ter_doc.tidnro = 6 "
    StrSql = StrSql & " WHERE repsijp.pliqnro = " & Pliqnro & " AND repsijp.lista_pronro = '" & Lista_Pronro & "' "
    StrSql = StrSql & " ORDER BY ter_doc.nrodoc "
    OpenRecordset StrSql, objRs
                    
    'seteo de las variables de progreso
    Progreso = 0
    CantRegistros = objRs.RecordCount
    If CantRegistros = 0 Then
        CantRegistros = 1
        Flog.Writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
    End If
    IncPorc = (100 / CantRegistros)
            
    Flog.Writeline Espacios(Tabulador * 1) & "Se comienza a procesar los datos"
           

    Do While Not objRs.EOF
        Empresa = CStr(objRs!empnom)
        i = 1
    
        Flog.Writeline "Se escribe el registro Nro.: " & i & " de la Empresa: " & Empresa
        
        Call imprimirTexto(objRs!Cuil, ArchExp, 11, True)               'CUIL
        
        Call imprimirTexto(Espacios(2), ArchExp, 2, True)               'Espacios en blanco
        
        Call imprimirTextoConCeros(objRs!cod_obra_social, ArchExp, 6, False)    'Codigo de Obra Social
        Call imprimirTextoConCeros(objRs!cant_hijos, ArchExp, 2, False)         'Cantidad de Hjos
        Call imprimirTextoConCeros(objRs!Conyuges, ArchExp, 1, True)            'Conyugue
        Call imprimirTextoConCeros(objRs!adherentes, ArchExp, 2, False)         'Cantidad de Adherentes
        Call imprimirTextoConCeros(objRs!cod_cont, ArchExp, 3, False)           'Codigo de modalidad de Contratacion
        Call imprimirTextoConCeros(objRs!zona, ArchExp, 2, False)               'Codigo de Zona
        Call imprimirTextoConCeros(objRs!actividad, ArchExp, 2, False)          'Codigo de Actividad
        Call imprimirTextoConCeros(objRs!cod_cond, ArchExp, 2, False)           'Codigo de Condicion
        Call imprimirTextoConCeros(objRs!cod_sitr, ArchExp, 2, False)           'Codigo de Situacion
        Call imprimirTextoConCeros(cambiarFormato(objRs!rem_total, Sep), ArchExp, 9, False) 'Remuneracion Total
        Call imprimirTextoConCeros(cambiarFormato(objRs!imp_ss, Sep), ArchExp, 9, False)    'Remuneracion Imponible 1
        
        Call imprimirTexto("00", ArchExp, 2, True)                              '00
        'Call imprimirTextoConCeros(Espacios(2), ArchExp, 2, True)              'Espacios en blanco
        
        Call imprimirTextoConCeros(cambiarFormato(objRs!aporte_voluntario, Sep), ArchExp, 9, False) 'Aporte Voluntario
        Call imprimirTextoConCeros(cambiarFormato(objRs!asig_fliar, Sep), ArchExp, 9, False) 'Asignaciones Familiares Pagadas
        Call imprimirTextoConCeros(cambiarFormato(objRs!imp_ss_con, Sep), ArchExp, 9, False) 'Remuneracion Imponible SS Contribuciones
        
        Call imprimirTexto("000000", ArchExp, 6, True)                          '000000
        Call imprimirTexto(".", ArchExp, 1, True)                               '.
        Call imprimirTexto("0000000", ArchExp, 7, True)                         '0000000
        'Call imprimirTextoConCeros(Espacios(14), ArchExp, 14, True)              'Espacios en blanco
        
        Call imprimirTextoConCeros(objRs!codsiniestro, ArchExp, 2, False) 'Codigo Siniestrado
        
        Call imprimirTexto("000000", ArchExp, 6, True)                          '000000
        Call imprimirTexto(".", ArchExp, 1, True)                               '.
        Call imprimirTexto("00", ArchExp, 2, True)                              '00
        'Call imprimirTextoConCeros(Espacios(9), ArchExp, 9, True)              'Espacios en blanco
        
        Call imprimirTextoConCeros(cambiarFormato(objRs!imprem3, Sep), ArchExp, 9, False) 'Remineracion Imponible3 Tope 75 Mopres
        Call imprimirTextoConCeros(cambiarFormato(objRs!imprem4, Sep), ArchExp, 9, False) 'Remineracion Imponible4 Tope 60 Mopres
        Call imprimirTextoConCeros(objRs!tipempnro, ArchExp, 1, False) 'Codigo de DGI del tipo de empleador
        
        'FAF - 08/02/2006 - Estaba mal la referencia. Se cambio objRs!secrep por objRs!con_fne
        Call imprimirTextoConCeros(objRs!con_fne, ArchExp, 1, False)  'Regimen Reparto o AFJP
        Call imprimirTextoConCeros(cambiarFormato(objRs!imprem5, Sep), ArchExp, 9, False) 'Remuneracion Imponible 5
        Call imprimirTextoConCeros(objRs!xconv, ArchExp, 1, False)    'Trabajador Convenciodado
        Call imprimirTextoConCeros(objRs!cod_sitr1, ArchExp, 2, False) 'Codigo de Situacion de Revista 1
        Call imprimirTextoConCeros(objRs!diainisr1, ArchExp, 2, False) 'Dia inicio SR 1
        Call imprimirTextoConCeros(objRs!cod_sitr2, ArchExp, 2, False) 'Codigo de Situacion de Revista 2
        Call imprimirTextoConCeros(objRs!diainisr2, ArchExp, 2, False) 'Dia inicio SR 2
        Call imprimirTextoConCeros(objRs!cod_sitr3, ArchExp, 2, False) 'Codigo de Situacion de Revista 3
        Call imprimirTextoConCeros(objRs!diainisr3, ArchExp, 2, False) 'Dia inicio SR 3
        Call imprimirTextoConCeros(cambiarFormato(objRs!sue_adic, Sep), ArchExp, 9, False) 'Sueldo + Adicionales
        Call imprimirTextoConCeros(cambiarFormato(objRs!sac, Sep), ArchExp, 9, False)     'SAC
        Call imprimirTextoConCeros(cambiarFormato(objRs!hrsextras, Sep), ArchExp, 9, False) 'Horas Extras
        Call imprimirTextoConCeros(cambiarFormato(objRs!zonadesf, Sep), ArchExp, 9, False) 'Zona Desfavorable
        Call imprimirTextoConCeros(cambiarFormato(objRs!lar, Sep), ArchExp, 9, False)   'Vacaciones
        Call imprimirTextoConCeros(cambiarFormato(Trim(objRs!diastrab), Sep), ArchExp, 9, False) 'Dias Trabajados
        
'        Call imprimirTexto(objRs!apenom, ArchExp, 30, True)   'Apellido y nombre
'        Call imprimirTexto(objRs!PORC_ADI, ArchExp, 5, False) 'Porcentaje de Aporte adicional SS
'        Call imprimirTexto(cambiarFormato(objRs!adi_os, Sep), ArchExp, 9, False)  'Adicional OS
'        Call imprimirTexto(cambiarFormato(objRs!exc_ss, Sep), ArchExp, 9, False)  'Excedentes Aportes SS
'        Call imprimirTexto(cambiarFormato(objRs!exc_os, Sep), ArchExp, 9, False)  'Excedentes Aportes OS
'        Call imprimirTexto(objRs!Localidad, ArchExp, 50, False) 'Provincia Localidad
'        Call imprimirTexto(objRs!correspred, ArchExp, 1, False) 'Corresp. Reduccion
'        Call imprimirTexto(cambiarFormato(objRs!caprecomlrt, Sep), ArchExp, 9, False) 'Capital Recomp. de LRT
'        Call imprimirTexto(cambiarFormato(objRs!apo_adi_os, Sep), ArchExp, 9, False) 'Aportes adicionales OS
'        Call imprimirTexto(cambiarFormato("0.00", Sep), ArchExp, 9, False) 'Remuneracion Imponible 6
'        Call imprimirTexto(cambiarFormato("0", Sep), ArchExp, 1, False) 'Tipo de operacion
        
        'Salto de linea
        ArchExp.Writeline ""
                      
        'Actualizo el progreso
        TiempoAcumulado = GetTickCount
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
         
        i = i + 1
        objRs.MoveNext
    Loop
    ArchExp.Close
    Flog.Writeline Espacios(Tabulador * 1) & "Se Terminaron de Procesar los datos"


    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    If objRs.State = adStateOpen Then objRs.Close
    
    Set rs = Nothing
    Set rs_Modelo = Nothing
    Set objRs = Nothing
Exit Sub

ME_Local:
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.Writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.Writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.Writeline
End Sub
