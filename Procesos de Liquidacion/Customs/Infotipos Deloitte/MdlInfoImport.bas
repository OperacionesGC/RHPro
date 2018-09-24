Attribute VB_Name = "MdlInfoImport"
Option Explicit

Private Type TEmpleadoInfotipo
    Legajo As Long
    Tercero As Long
End Type

Public Type TReg_Tercero
    Ternom As String
    Ternom2 As String
    Terape As String
    Terape2 As String
    Terfecnac As Date
    Tersex As Integer
    Terestciv As String
    EstCivNro As Long
    NacionalNro As Long
    PaisNro As Long
    TerFecIng As Date
End Type

Public Type TEstructuras
    Tenro As Long
    Estrnro As Long
    Desde As String
    Hasta As String
End Type

Public Type TNomina
    Nomina As String
    Monto As Single
    Cantidad As Single
    Unidad As String
    Operacion As String
End Type

Public Type TCCosto
    Sociedad As String
    Division As String
    CCosto As String
    Porcentaje As String
End Type

Global Empleado As TEmpleadoInfotipo
Global Infotipo As String

Global Accion As String
Global SubAccion As String
Global Status As String
Global Contador_Familiar As Long

Global Medida_Clase As String
Global Medida_Motivo As String
Global Reg_Tercero As TReg_Tercero
Global Estructuras(100) As TEstructuras
Global rs_Empleado As New ADODB.Recordset
Global ExisteLegajo As Boolean
Global UltimoLegajo As Long
Global Continuar As Boolean
Global Fecha_Alta
Global Fecha_BajaPrevista

Global Infotipo_0000 As Boolean
Global Infotipo_0001 As Boolean
Global Infotipo_0002 As Boolean
Global Infotipo_0006 As Boolean
Global Infotipo_0007 As Boolean 'No se hace
Global Infotipo_0008 As Boolean
Global Infotipo_0009 As Boolean
Global Infotipo_0014 As Boolean
Global Infotipo_0015 As Boolean
Global Infotipo_0016 As Boolean 'No se hace
Global Infotipo_0021 As Boolean
Global Infotipo_0023 As Boolean 'No se hace
Global Infotipo_0027 As Boolean
Global Infotipo_0041 As Boolean
Global Infotipo_0050 As Boolean 'No se hace
Global Infotipo_0057 As Boolean
Global Infotipo_0105 As Boolean 'No se hace
Global Infotipo_0185 As Boolean
Global Infotipo_0267 As Boolean 'No se hace
Global Infotipo_0389 As Boolean
Global Infotipo_0390 As Boolean
Global Infotipo_0391 As Boolean
Global Infotipo_0392 As Boolean
Global Infotipo_0393 As Boolean
Global Infotipo_0394 As Boolean
Global Infotipo_0416 As Boolean 'No se hace
Global Infotipo_2001 As Boolean
Global Infotipo_2006 As Boolean 'No se hace
Global Infotipo_2010 As Boolean
Global Infotipo_2013 As Boolean 'No se hace
Global Infotipo_9004 As Boolean 'Benefits Brasil ???????????
Global Infotipo_9302 As Boolean 'Prestamos

Global PrimeraVez_Infotipo_0000 As Boolean
Global PrimeraVez_Infotipo_0001 As Boolean
Global PrimeraVez_Infotipo_0002 As Boolean
Global PrimeraVez_Infotipo_0006 As Boolean
Global PrimeraVez_Infotipo_0007 As Boolean 'No se hace
Global PrimeraVez_Infotipo_0008 As Boolean
Global PrimeraVez_Infotipo_0009 As Boolean
Global PrimeraVez_Infotipo_0014 As Boolean
Global PrimeraVez_Infotipo_0015 As Boolean
Global PrimeraVez_Infotipo_0016 As Boolean 'No se hace
Global PrimeraVez_Infotipo_0021 As Boolean
Global PrimeraVez_Infotipo_0023 As Boolean 'No se hace
Global PrimeraVez_Infotipo_0027 As Boolean
Global PrimeraVez_Infotipo_0041 As Boolean
Global PrimeraVez_Infotipo_0050 As Boolean 'No se hace
Global PrimeraVez_Infotipo_0057 As Boolean
Global PrimeraVez_Infotipo_0105 As Boolean 'No se hace
Global PrimeraVez_Infotipo_0185 As Boolean
Global PrimeraVez_Infotipo_0267 As Boolean 'No se hace
Global PrimeraVez_Infotipo_0389 As Boolean
Global PrimeraVez_Infotipo_0390 As Boolean
Global PrimeraVez_Infotipo_0391 As Boolean
Global PrimeraVez_Infotipo_0392 As Boolean
Global PrimeraVez_Infotipo_0393 As Boolean
Global PrimeraVez_Infotipo_0394 As Boolean
Global PrimeraVez_Infotipo_0416 As Boolean 'No se hace
Global PrimeraVez_Infotipo_2001 As Boolean
Global PrimeraVez_Infotipo_2006 As Boolean 'No se hace
Global PrimeraVez_Infotipo_2010 As Boolean
Global PrimeraVez_Infotipo_2013 As Boolean 'No se hace
Global PrimeraVez_Infotipo_9004 As Boolean 'Benefits Brasil ???????????
Global PrimeraVez_Infotipo_9302 As Boolean 'Prestamos

Global Cant_Acumulada As Integer
Global Monto_Acumulado As Single




Public Sub LeeArchivo(ByVal NombreArchivo As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento
' Autor      : FGZ
' Fecha      : 23/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0
Dim strlinea As String
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

    If App.PrevInstance Then Exit Sub

    'Espero hasta que se crea el archivo
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then Err.Number = 1
    Loop
    On Error GoTo 0
   
   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProcesoBatch & "," & NroModelo & ",'" & Left(NombreArchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "inter_pin")
    End If
                
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If rs_Modelo.EOF Then
        Exit Sub
    End If
                
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = 0
    Do While Not f.AtEndOfStream
        strlinea = f.ReadLine
        If f.AtEndOfStream Then
            CEmpleadosAProc = f.Line
        End If
    Loop
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (99 / CEmpleadosAProc)
    f.Close
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    
    'FGZ - 05/07/2005
    'Crea la planilla de Excel que contendrá la informacion leida
    Call CrearArchivoExcel(ArchivoAGenerar)
    
    'Creo el csv con la informacion levantada
    Set fNovedades = fs.CreateTextFile(ArchivoNovedades, True)
    Set fCambios = fs.CreateTextFile(ArchivoCambios, True)
    Call InsertarLogEncabezadoNovedad
    Call InsertarLogEncabezadoCambios
    UltimoLegajo = -1
    Do While Not f.AtEndOfStream
        strlinea = f.ReadLine
        NroLinea = NroLinea + 1
        If NroLinea = 1 And UsaEncabezado Then
            strlinea = f.ReadLine
        End If
        If Trim(strlinea) <> "" Then
            RegLeidos = RegLeidos + 1
            
            Select Case rs_Modelo!modinterface
                Case 4:
                    Call Insertar_Linea_Segun_Infotipo(strlinea)
                Case Else
                    'para levantar las tablas globales (MAPEOS)
                    Call Insertar_Linea_Segun_Modelo(strlinea)
            End Select
        End If
        
        'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
        'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
        'como colgado
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Round(CSng(Progreso), 3) & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Loop
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    f.Close
    Flog.writeline Espacios(Tabulador * 1) & "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'Borrar el archivo
    fs.Deletefile NombreArchivo, True
    Call CerrarArchivoExcel(ArchivoAGenerar)
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
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


Public Sub Insertar_Linea_Segun_Infotipo(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador segun infotipo
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1            As Integer
Dim pos2            As Integer

Dim rs_Empleado As New ADODB.Recordset
' ----------------------------------------------------------------
' tengo que leer el string y determinar el infotipo que corresponde
' Todos los infotipos vienen con el legajo al inicio seguido del infotipo
' El formato es:
'   PERNR   NUMC    8   Personnel Number
'   INFTY   CHAR    6   Constant name infotype
'   El resto de los campos depende del infotipo

On Error GoTo Manejador_De_Error


If Primera_vez Then
    Fila_Infotipo_0000 = 1
    Fila_Infotipo_0001 = 1
    Fila_Infotipo_0002 = 1
    Fila_Infotipo_0006 = 1
    Fila_Infotipo_0008 = 1
    Fila_Infotipo_0009 = 1
    Fila_Infotipo_0014 = 1
    Fila_Infotipo_0015 = 1
    Fila_Infotipo_0021 = 1
    'Fila_Infotipo_0027 = 1
    Fila_Infotipo_0041 = 1
    Fila_Infotipo_0057 = 1
    Fila_Infotipo_0185 = 1
    Fila_Infotipo_0389 = 1
    Fila_Infotipo_0390 = 1
    Fila_Infotipo_0391 = 1
    Fila_Infotipo_0392 = 1
    Fila_Infotipo_0393 = 1
    Fila_Infotipo_0394 = 1
    Fila_Infotipo_2001 = 1
    Fila_Infotipo_2010 = 1
    Fila_Infotipo_9004 = 1
    Fila_Infotipo_9302 = 1
    
    'Call MapearNominasAutomaticamente
    Primera_vez = False
End If
    'Nro de Legajo
    pos1 = 1
    pos2 = 8
    If IsNumeric(Mid$(strlinea, pos1, pos2)) Then
        Empleado.Legajo = Mid$(strlinea, pos1, pos2)
        If UltimoLegajo <> Empleado.Legajo Then
            Flog.writeline Espacios(Tabulador * 1) & "Legajo " & Empleado.Legajo
            
            'Inicializo todas las variables globales que interesan al proceso
            Call Inicializar_Globales
            
            UltimoLegajo = Empleado.Legajo
            StrSql = "SELECT empleg, ternro FROM empleado"
            StrSql = StrSql & " WHERE empleg = " & Empleado.Legajo
            OpenRecordset StrSql, rs_Empleado
            
            If rs_Empleado.EOF Then
                ExisteLegajo = False
                Empleado.Tercero = 0
            Else
                ExisteLegajo = True
                Empleado.Tercero = rs_Empleado!ternro
            End If
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El legajo no es numerico"
        FlogE.writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El legajo no es numerico"
        InsertaError 1, 8
        HuboError = True
        Exit Sub
    End If
    
    'Infotipos
    pos1 = 9
    pos2 = 6
    Infotipo = Mid(strlinea, pos1, pos2)

    If Continuar Then
        Select Case Infotipo
        Case "IT0000":  '- ACTIONS
            Call Leer_Infotipo_IT0000(strlinea)
        Case "IT0001":  '– ORGANIZATIONAL ASSIGNMENT
            Call Leer_Infotipo_IT0001(strlinea)
        Case "IT0002":  '- PERSONAL DATA
            Call Leer_Infotipo_IT0002(strlinea)
        Case "IT0006":  '– ADDRESSES
            Call Leer_Infotipo_IT0006(strlinea)
        Case "IT0007":  '– PLANNED WORKING TIME
            'Call Leer_Infotipo_IT0007(strLinea)
            ' No se levanta, no trae nada que interese
        Case "IT0008":  '– BASIC PAY
            Call Leer_Infotipo_IT0008(strlinea)
        Case "IT0009":  '– BANK DETAILS
            Call Leer_Infotipo_IT0009(strlinea)
        Case "IT0014":  '– DEVENGOS Y DEDUCCIONES PERIODICAS
            Call Leer_Infotipo_IT0014(strlinea)
        Case "IT0015":  '– DEVENGOS COMPLEMENTARIOS
            Call Leer_Infotipo_IT0015(strlinea)
        Case "IT0016":  '– CONTRACT ELEMENTS
            'Call Leer_Infotipo_IT0016(strLinea)
            ' No se levanta, no trae nada que interese
        Case "IT0021":  '- FAMILY/RELEATED PERSON
            Call Leer_Infotipo_IT0021(strlinea)
        Case "IT0023":  '– OTHER/PREVIOUS EMPLOYERS
            'Call Leer_Infotipo_IT0023(strLinea)
            ' No se levanta, no trae nada que interese
        Case "IT0027":  '- Distribucion de Costos
            Call Leer_Infotipo_IT0027(strlinea)
        Case "IT0041":  '– DATE SPECIFICATIONS
            Call Leer_Infotipo_IT0041(strlinea)
        Case "IT0050":  '– TIME RECORDING INFO
            'Call Leer_Infotipo_IT0050(strLinea)
            ' No se levanta, no trae nada que interese
        Case "IT0057":  '– MEMBERSHIP FEES
            Call Leer_Infotipo_IT0057(strlinea)
        Case "IT0105":  '– COMUNICACION
            'Call Leer_Infotipo_IT0105(strLinea)
            'No se levanta, no trae nada que interese
        Case "IT0185":  '- PERSONAL IDs
            Call Leer_Infotipo_IT0185(strlinea)
        Case "IT0267":  '- PAGOS COMPLEMENTARIOS NOMINA ESPECIAL
            'Call Leer_Infotipo_IT0267(strLinea)
            ' No se levanta, no trae nada que interese
        Case "IT0389":  '- IMPUESTO A LAS GANANCIAS (ARGENTINA)
            Call Leer_Infotipo_IT0389(strlinea)
        Case "IT0390":  '– IMPUESTO A LAS GANANCIAS - DEDUCCIONES (ARGENTINA)
            Call Leer_Infotipo_IT0390(strlinea)
            'OJO, por ahora esta desactivado. no lo va a levantar
        Case "IT0391":  '– IMPUESTO A LAS GANANCIAS - OTRO EMPLEADOR
            Call Leer_Infotipo_IT0391(strlinea)
        Case "IT0392":  '– SEGURIDAD SOCIAL - ARGENTINA
            Call Leer_Infotipo_IT0392(strlinea)
        Case "IT0393":  '– DATOS DE FAMILIA AYUDA ESCOLAR (ARGENTINA)
            Call Leer_Infotipo_IT0393(strlinea)
        Case "IT0394":  '– DATOS FAMILIA: INFORMACION ADICIONAL -ARGENTINA
            Call Leer_Infotipo_IT0394(strlinea)
        Case "IT0416":  '– COMPENSACION CONTINGENTES DE TIEMPOS
            'Call Leer_Infotipo_IT0416(strLinea)
            ' No se levanta, no trae nada que interese
        Case "IT2001":  '– AUSENTISMOS
            Call Leer_Infotipo_IT2001(strlinea)
        Case "IT2006":  '– CONTINGENTES DE AUSENTISMOS
            'Call Leer_Infotipo_IT2006(strLinea)
            ' No se levanta, no trae nada que interese
        Case "IT2010":  '– COMPROBANTE DE REMUNERACION
            Call Leer_Infotipo_IT2010(strlinea)
        Case "IT2013":  '– CORRECCION CONTINGENTES
            'Call Leer_Infotipo_IT2013(strLinea)
            ' No se levanta, no trae nada que interese
        Case "IT9004":  '– BENEFICIOS - BRASIL
            Call Leer_Infotipo_IT9004(strlinea)
            ' No se levanta, no trae nada que interese
        Case "IT9302":  '– Prestamos
            Call Leer_Infotipo_IT9302(strlinea)
        Case Else
            Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
            Flog.writeline Espacios(Tabulador * 0) & "Infotipo " & Infotipo & " desconocido "
            Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        End Select
    End If
    
Fin:
    'Flog.Writeline Espacios(Tabulador * 3) & "Infotipo actualizado correctamente "
Exit Sub

Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Public Sub Leer_Infotipo_IT0000(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0000. Actions.
' Autor      : FGZ
' Fecha      : 23/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
'PERNR   NUMC               8           Personnel Number
'INFTY   CHAR               6           Constant name infotype
'SUBTY   CHAR               4           Subtipo
'BEGDA   DATS               8           Inicio de Validez
'ENDDA   DATS               8           Fin de Validez
'PREAS   CHAR               2           Motivo para el cambio de datos maestros
'MASSN   CHAR               2           Clase de medida                                                 1   T529A
'MASSG   CHAR               2           Motivo de la medida                                             2   T530
'STAT1   CHAR               1           Status individual de cliente
'STAT2   CHAR               1           Status del empleado
'STAT3   CHAR               1           Status pagas extraordinarias
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Motivo As String
Dim Status_Individual As String
Dim Status_Empleado As String
Dim Status_Pagas As String

Dim Causa As Long
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0000"
    Columna = 2
    Infotipo_0000 = False
    Fila_Infotipo_0000 = Fila_Infotipo_0000 + 1
    Hoja = 1
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)
    
    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 2) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Motivo para el cambio de datos maestros
    Columna = Columna + 1
    Texto = "Motivo para el cambio de datos maestros"
    pos1 = 35
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, Columna, Aux)
    Motivo = Mid(strlinea, pos1, pos2)

    'Clase de Medida
    Columna = Columna + 1
    Texto = "Clase de Medida"
        '29  HA  Hiring Argentina
        '29  LA  Leaving Argentina
        '29  CA  Change of pay Argentina
        '29  OA  Organizational reassignment Ar
    pos1 = 37
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, Columna, Aux)
    Medida_Clase = Mid(strlinea, pos1, pos2)

    'Motivo de la medida
    Columna = Columna + 1
    Texto = "Motivo de la medida"
        'HA      01      Expansion
        'HA      02      Substitution of workplaces
        'HA      03      Re-entry temporary suspension
        'HA      04      Department reorganisation
        'LA      01      Resignation
        'LA      02      Dismissal
        'LA      03      Dismissal
        'LA      04      Dissolution contract
        'LA      05      Firing
        'LA      06      Disability retirement
        'LA      07      Old age (normal) retirement
        'LA      08      Antecipated retirement
        'LA      09      Death of employee
        'LA      20      Immediate termination
        'LA      K1      Temp. or Permanent Layofff
        'LA      K2      Temporary closure
        'LA      K3      Seasonal closure
        'LA      M1      Unjustified dismissal
        'CA      01      Pay Scale Reclassification
        'CA      02      Pay Scale Increase
        'CA      03      Annual increase
        'CA      04      Promotion
        'OA      01      Change of employer (leg.pers.)
        'OA      02      Lateral move
        'OA      03      Change employee subgroup
        'OA      04      Temporary to regular
        'OA      05      Regular to temporary
        'OA      06      Demotion
        'OA      07      Company code change
        'OA      08      Personnel area change
        'OA      09      Personnel subarea change
        'OA      10      Cost center change
        'OA      11      Organization unit change
        'OA      12      Position change
        'OA      13      Job code change
    pos1 = 39
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, Columna, Aux)
    Medida_Motivo = Mid(strlinea, pos1, pos2)

    'Status Individual del cliente
    Columna = Columna + 1
    Texto = "Status Individual del cliente"
    pos1 = 41
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, Columna, Aux)
    Status_Individual = Mid(strlinea, pos1, pos2)

    'Status del empleado
    Columna = Columna + 1
    Texto = "Status del empleado"
    pos1 = 42
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, Columna, Aux)
    Status_Empleado = Mid(strlinea, pos1, pos2)

    'Status pagas extraordinarias
    Columna = Columna + 1
    Texto = "Status pagas extraordinarias"
    pos1 = 43
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0000, Columna, Aux)
    Status_Pagas = Mid(strlinea, pos1, pos2)

Select Case Medida_Clase
    Case "12"   'Reentry into company (esto supuestamente es para Mexico pero .....)
        Accion = "ALTA"
        Fecha_Alta = Inicio_Validez
        Select Case Medida_Motivo
        Case "01":  '12      01      Expansion
            SubAccion = "Expansion 01"
        Case "02":  '12      02      Substitution of workplaces
            SubAccion = "Substitución de lugares de trabajo 02"
        Case "03":  '12      03      Re-entry temporary suspension
            Accion = "REINCORPORACION"
            SubAccion = "Reincorporacion de suspension 03"
        Case "04":  '12      04      Department reorganisation
            SubAccion = "Reorganizacion departamental 04"
        Case Else:
            SubAccion = ""
            FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": " & "Motivo de Medida desconocido " & Medida_Motivo
        End Select
    Case "HA"   'HA  Hiring Argentina
        Accion = "ALTA"
        Fecha_Alta = Inicio_Validez
        'Fecha_BajaPrevista = Fin_Validez
        
        Select Case Medida_Motivo
        Case "01":  'HA      01      Expansion
            SubAccion = "Expansion 01"
            'Call Action_HA_01(Subtipo, Inicio_Validez, Fin_Validez, Motivo, Status_Individual, Status_Empleado, Status_Pagas)
        Case "02":  'HA      02      Substitution of workplaces
            SubAccion = "Substitución de lugares de trabajo 02"
        Case "03":  'HA      03      Re-entry temporary suspension
            Accion = "REINCORPORACION"
            SubAccion = "Reincorporacion de suspension 03"
        Case "04":  'HA      04      Department reorganisation
            SubAccion = "Reorganizacion departamental 04"
        Case Else:
            SubAccion = ""
            FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": " & "Motivo de Medida desconocido " & Medida_Motivo
        End Select
    Case "LA":  'LA  Leaving Argentina
        Accion = "BAJA"
        Select Case Medida_Motivo
        Case "01":  'LA      01      Resignation
            SubAccion = "Renuncia 01"
        Case "02":  'LA      02      Dismissal
            SubAccion = "Despido 02"
        Case "03":  'LA      03      Dismissal
            SubAccion = "Despido 03"
        Case "04":  'LA      04      Dissolution contract
            SubAccion = "Disolucion de contrato 04"
        Case "05":  'LA      05      Firing
            SubAccion = "Despido 05"
        Case "06":  'LA      06      Disability retirement
            SubAccion = "Retiro por incapacidad 06"
        Case "07":  'LA      07      Old age (normal) retirement
            SubAccion = "Retiro por edad 07"
        Case "08":  'LA      08      Antecipated retirement
            SubAccion = "Retiro anticipado 08"
        Case "09":  'LA      09      Death of employee
            SubAccion = "Muerte del empleado 09"
        Case "20":  'LA      20      Immediate termination
            SubAccion = "Terminacion Inmediata 20"
        Case "K1":  'LA      K1      Temp. or Permanent Layofff
            SubAccion = "Despido temporal o permanente K1"
        Case "K2":  'LA      K2      Temporary closure
            SubAccion = "Cierre Temporal K2"
        Case "K3":  'LA      K3      Seasonal closure
            SubAccion = "Cierre de Temporada K3"
        Case "M1":  'LA      M1      Unjustified dismissal
            SubAccion = "Despido Injustificado M1"
        Case Else:
            FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": " & "Motivo de Medida desconocido " & Medida_Motivo
        End Select
        Causa = CLng(CalcularMapeoInv(Medida_Motivo, "CAUSAB", "0"))
        OK = False
        If Causa <> 0 Then
            OK = Baja_Empleado(Empleado.Tercero, Inicio_Validez, Causa)
        End If
        If Not OK Then
            Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
            FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": Error dando de baja el legajo. Infotipo abortado"
            'InsertaError 0, 8
            'HuboError = True
            'Continuar = False
            'Exit Sub
'        Else
'            Flog.Writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
'            FlogE.Writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": Error dando de baja el legajo. Infotipo abortado"
        End If
        
    Case "CA":  'CA  Change of pay Argentina
        Accion = "MODIFICACION"
        Select Case Medida_Motivo
        Case "01":  'CA      01      Pay Scale Reclassification
            SubAccion = "Reclasificacion de escala salarial 01"
        Case "02":  'CA      02      Pay Scale Increase
            SubAccion = "Aumento De la Escala salarial 02"
        Case "03":  'CA      03      Annual increase
            SubAccion = "Aumento anual"
        Case "04":  'CA      04      Promotion
            SubAccion = "Promocion"
        Case Else:
            SubAccion = ""
            Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
            FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": " & "Motivo de Medida desconocido " & Medida_Motivo
        End Select
        'Call Cambiar_Sueldo
        
    Case "OA":  'OA  Organizational reassignment Ar
        Accion = "MODIFICACION"
        Select Case Medida_Motivo
        Case "01":  'OA      01      Change of employer (leg.pers.)
            SubAccion = "Cambio de empleador (leg. pers.) 01"
        Case "02":  'OA      02      Lateral move
            SubAccion = "Movimiento lateral 02"
        Case "03":  'OA      03      Change employee subgroup
            SubAccion = "Cambio de subgrupo del empleado 03"
        Case "04":  'OA      04      Temporary to regular
            SubAccion = "Temporal a Regular 04"
        Case "05":  'OA      05      Regular to temporary
            SubAccion = "Regular a temporal 05"
        Case "06":  'OA      06      Demotion
            SubAccion = "Demotion 06"
        Case "07":  'OA      07      Company code change
            SubAccion = "Cambio del código de la compañía 07"
        Case "08":  'OA      08      Personnel area change
            SubAccion = "Cambio del área del personal 08"
        Case "09":  'OA      09      Personnel subarea change
            SubAccion = "Cambio del subárea del personal 09"
        Case "10":  'OA      10      Cost center change
            SubAccion = "Cambio de Centro de costo 10"
        Case "11":  'OA      11      Organization unit change
            SubAccion = "Cambio de unidad de la organización 11"
        Case "12":  'OA      12      Position change
            SubAccion = "Cambio de Puesto 12"
        Case "13":  'OA      13      Job code change
            SubAccion = "Cambio de Codigo de trabajo 13"
        Case Else:
            SubAccion = ""
            Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
            FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": " & "Motivo de Medida desconocido " & Medida_Motivo
        End Select
Case Else:
    Accion = ""
    SubAccion = ""
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": " & "Clase de Medida desconocido " & Medida_Clase
End Select

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
End Sub


Public Sub Leer_Infotipo_IT0001(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0001. Organizational Assignament.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION                Código Tabla    Nombre Técnico
'PERNR   NUMC               8       Personnel Number
'INFTY   CHAR               6       Constant name infotype
'SUBTY   CHAR               4       Subtipo
'BEGDA   DATS               8       Inicio de Validez
'ENDDA   DATS               8       Fin de la Validez
'BUKRS   CHAR               4       Sociedad                                                       3   T001
'WERKS   CHAR               4       División de personal                                           4   T500P
'PERSG   CHAR               1       Grupo de personal                                              5   T501
'PERSK   CHAR               2       Área de personal                                               6   T503K
'VDSK1   CHAR               14      Clave de Organización
'GSBER   CHAR               4       División
'BTRTL   CHAR               4       Subdivisión de personal             7           T001P
'JUPER   CHAR               4       Legal Person
'ABKRS   CHAR               2       Área de nómina                      8           T549A
'ANSVH   CHAR               2       Relacion Laboral                    35          T542A
'KOSTL   CHAR               10      Centro de coste                     9           CSKS
'ORGEH   NUMC               8       Unidad Oganizativa                  12          T527X
'PLANS   NUMC               8       Posición                            10          T528B
'STELL   NUMC               8       Función                             11          T513
'MSTBR   CHAR               8       Area maestro
'SACHA   CHAR               3       Administrador de nómina
'SACHP   CHAR               3       Administrador de personal
'SACHZ   CHAR               3       Administrador de tiempos
'SNAME   CHAR               30      Employee's Name (Sortable by LAST NAME FIRST NAME)
'ENAME   CHAR               40      Formatted Name of Employee or Applicant
'OTYPE   CHAR               2       Tipo de objeto
'SBMOD   CHAR               4       Grupo de Administración
'KOKRS   CHAR               4       Area de contabilidad
'FISTL   CHAR               16      Funds Center
'GEBER   CHAR               10      Fund
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte
Dim i As Integer

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Sociedad As String
Dim Division_de_Personal As String
Dim Grupo_de_Personal As String
Dim Area_de_Personal As String
Dim Nro_Grupo_de_Personal As Long
Dim Nro_Area_de_Personal As Long
Dim Clave_de_Organizacion As String
Dim Division As String
Dim Subdivision_de_Personal As String
Dim Legal_Person As String
Dim Area_de_Nomina As String
Dim Relacion_Laboral As String
Dim Centro_de_Costo As String
Dim Unidad_Organizativa As Long
Dim Posicion As Long
Dim Funcion As Long
Dim Area_Maestro As String
Dim Administrador_de_Nomina As String
Dim Administrador_de_Personal As String
Dim Administrador_de_Tiempos As String
Dim Nombre_de_Empleado As String
Dim Nombre_Empleado_Formateado As String
Dim Tipo_de_Objeto As String
Dim Grupo_de_Administracion As String
Dim Area_de_Contabilidad As String
Dim Funds_center As String
Dim Fund As String

'Auxiliares
Dim Tenro As Long
Dim Empresa As Long
Dim Gerencia As Long
Dim Tipo_Contrato As Long
Dim Convenio As String
Dim nro_convenio As Long
Dim Sucursal As Long
Dim CCosto As Long
Dim Sector As Long
Dim Nro_Area_de_Nomina As Long
Dim Hoja As Integer

Dim rs_Estructura As New ADODB.Recordset


'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0001"
    Columna = 2
    Infotipo_0001 = False
    Fila_Infotipo_0001 = Fila_Infotipo_0001 + 1
    Hoja = 2
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Sociedad
    Columna = 6
    pos1 = 35
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Sociedad = Mid(strlinea, pos1, pos2)

    'Division de Personal
    Columna = 7
    pos1 = 39
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Division_de_Personal = Mid(strlinea, pos1, pos2)

    'Grupo de Personal
    Columna = 8
    pos1 = 43
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Grupo_de_Personal = Mid(strlinea, pos1, pos2)

    'Area de Personal
    Columna = 9
    pos1 = 44
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Area_de_Personal = Mid(strlinea, pos1, pos2)

    'Clave de Organizacion
    Columna = 10
    pos1 = 46
    pos2 = 14
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Clave_de_Organizacion = Mid(strlinea, pos1, pos2)

    'Division
    Columna = 11
    pos1 = 60
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Division = Mid(strlinea, pos1, pos2)

    'Subdivision de Personal
    Columna = 12
    pos1 = 64
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Subdivision_de_Personal = Mid(strlinea, pos1, pos2)

    'Legal Person
    Columna = 13
    pos1 = 68
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Legal_Person = Mid(strlinea, pos1, pos2)

    'Area de Nomina
    Columna = 14
    pos1 = 72
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Area_de_Nomina = Mid(strlinea, pos1, pos2)

    'Relacion Laboral
    Columna = 15
    pos1 = 74
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Relacion_Laboral = Mid(strlinea, pos1, pos2)

    'Centro de Costo
    Columna = 16
    pos1 = 76
    pos2 = 10
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Centro_de_Costo = Mid(strlinea, pos1, pos2)

    'Unidad Organizativa
    Columna = 17
    pos1 = 86
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, Aux)
    Unidad_Organizativa = Mid(strlinea, pos1, pos2)

    'Posicion
    pos1 = 94
    pos2 = 8
    Posicion = Mid(strlinea, pos1, pos2)

    'Funcion
    pos1 = 102
    pos2 = 8
    Funcion = Mid(strlinea, pos1, pos2)

    'Area Maestro
    pos1 = 110
    pos2 = 8
    Area_Maestro = Mid(strlinea, pos1, pos2)

    'Administrador de Nomina
    pos1 = 118
    pos2 = 3
    Administrador_de_Nomina = Mid(strlinea, pos1, pos2)

    'Administrador de Personal
    pos1 = 121
    pos2 = 3
    Administrador_de_Personal = Mid(strlinea, pos1, pos2)

    'Administrador de Tiempos
    pos1 = 124
    pos2 = 3
    Administrador_de_Tiempos = Mid(strlinea, pos1, pos2)

    'Nombre de Empleado
    pos1 = 127
    pos2 = 30
    Nombre_de_Empleado = Mid(strlinea, pos1, pos2)

    'Nombre de Empleado con Formato
    pos1 = 157
    pos2 = 40
    Nombre_Empleado_Formateado = Mid(strlinea, pos1, pos2)

    'Tipo de Objeto
    pos1 = 197
    pos2 = 2
    Tipo_de_Objeto = Mid(strlinea, pos1, pos2)

    'Grupo de Administracion
    pos1 = 199
    pos2 = 4
    Grupo_de_Administracion = Mid(strlinea, pos1, pos2)

    'Area de Contabilidad
    pos1 = 203
    pos2 = 4
    Area_de_Contabilidad = Mid(strlinea, pos1, pos2)

    'Funds Center
    pos1 = 207
    pos2 = 16
    Funds_center = Mid(strlinea, pos1, pos2)

    'Fund
    pos1 = 223
    pos2 = 10
    Fund = Mid(strlinea, pos1, pos2)

    'Completo las columnas vacias o que no tienen importancia
    For Columna = 18 To 30
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0001, Columna, "")
    Next Columna

'----------------------------------------------------------------------------------------
' Validaciones
'----------------------------------------------------------------------------------------

'Sociedad - [ RHPro(Empresa)]
Texto = "Sociedad - [ RHPro(Empresa)] " & Sociedad
Empresa = CLng(CalcularMapeoInv(Sociedad, "T001", "-1"))
Tenro = 10
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = Empresa
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
If Empresa <> 0 Then
'    Call Insertar_His_Estructura(Tenro, Empresa, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If

'Division de Personal - [ RHPro(Gerencia)]
Texto = "Division de Personal - [ RHPro(Gerencia)] " & Division_de_Personal
Gerencia = CLng(CalcularMapeoInv(Division_de_Personal, "T500P", "0"))
Tenro = 6
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = Gerencia
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
If Gerencia <> 0 Then
'    Call Insertar_His_Estructura(Tenro, Gerencia, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If

'Grupo de Personal - [ RHPro(Tipo de personal)]
Texto = "Grupo de Personal - [ RHPro(Grupo de Personal)] " & Grupo_de_Personal
Nro_Grupo_de_Personal = CLng(CalcularMapeoInv(Grupo_de_Personal, "T501", "-1"))
Tenro = 48
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = Nro_Grupo_de_Personal
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
If Nro_Grupo_de_Personal <> 0 Then
'    Call Insertar_His_Estructura(Tenro, Tipo_Contrato, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If

'Area de Personal - [ RHPro (Area de Personal)]
Texto = "Area de Personal - [ RHPro (Area de Personal)] " & Area_de_Personal
Nro_Area_de_Personal = CLng(CalcularMapeoInv(Area_de_Personal, "T503K", "-1"))
Tenro = 49
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = Nro_Area_de_Personal
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
If Nro_Area_de_Personal <> 0 Then
'    Call Insertar_His_Estructura(Tenro, Convenio, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If

'Convenio
'Tambien del Area de Personal saco el Convenio de RHPro
Convenio = Area_de_Personal
'Ver como convertir


Texto = "Convenio " & Convenio
nro_convenio = CLng(CalcularMapeoInv(Convenio, "TCONV", "0"))
Tenro = 19
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = nro_convenio
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
If nro_convenio <> 0 Then
'    Call Insertar_His_Estructura(Tenro, Convenio, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If

'Clave de Organizacion (por ahora no se hace nada)


'Division (por ahora no se hace nada)


'Subdivision de Personal - [RHPro (Sucursal)]
Texto = "Subdivision de Personal - [ RHPro (Sucursal)] " & Subdivision_de_Personal
Sucursal = CLng(CalcularMapeoInv(Subdivision_de_Personal, "T001P", "-1"))
Tenro = 1
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = Sucursal
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
If Sucursal <> 0 Then
'    Call Insertar_His_Estructura(Tenro, Sucursal, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If

'Legal Person (por ahora no se hace nada)

'Area de nomina - [RHPro (Centro de Costo)]
Texto = "Centro de Coste - [RHPro (Centro de Costo)] " & Area_de_Nomina
Nro_Area_de_Nomina = CLng(CalcularMapeoInv(Area_de_Nomina, "T549A", "-1"))
Tenro = 22
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = Nro_Area_de_Nomina
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
If Nro_Area_de_Nomina <> 0 Then
'    Call Insertar_His_Estructura(Tenro, CCosto, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If

'Relacion Laboral (No se hace nada. Aplicable solo para Colombia)
'99  NA Not APPLY

'Centro de Coste - [Heidt (Centro de Costo)]
Texto = "Centro de Coste - [RHPro (Centro de Costo)] " & Centro_de_Costo
CCosto = CLng(CalcularMapeoInv(Centro_de_Costo, "CSKS", "-1"))
Tenro = 5
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = CCosto
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
If CCosto <> 0 Then
'    Call Insertar_His_Estructura(Tenro, CCosto, Empleado.Tercero, Inicio_Validez, Fin_Validez)

    'FGZ - 06/02/2007 - Le agrego automaticamente el Profit Center de acuerdo a la siguiente regla
    '                   Si El codigo externo del CC comienza con 11* ==> Pe pongo un Profit 0000110000
    '                   Si El codigo externo del CC comienza con 45* ==> Pe pongo un Profit 0000453000
    StrSql = " SELECT estrcodext FROM estructura "
    StrSql = StrSql & " WHERE estrnro = " & CCosto
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        If Not EsNulo(rs_Estructura!estrcodext) Then
            If Left(rs_Estructura!estrcodext, 2) = "11" Then
                Aux = "0000110000"
            Else
                If Left(rs_Estructura!estrcodext, 2) = "45" Then
                    Aux = "0000453000"
                Else
                    Aux = ""
                End If
            End If
            
            'Busco el Profit con este codigo externo
            StrSql = " SELECT * FROM estructura "
            StrSql = StrSql & " WHERE tenro = 50"
            StrSql = StrSql & " AND estrcodext = '" & Aux & "'"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                Tenro = 50
                Estructuras(Tenro).Tenro = Tenro
                Estructuras(Tenro).Estrnro = rs_Estructura!Estrnro
                Estructuras(Tenro).Desde = Inicio_Validez
                Estructuras(Tenro).Hasta = Fin_Validez
            Else
                Flog.writeline Espacios(Tabulador * 3) & "No se actualizará automaticamente el PROFIT pues no existe ninguno con codigo externo " & Aux
            End If
        Else
            Flog.writeline Espacios(Tabulador * 3) & "No se actualizará automaticamente el PROFIT. CC con codigo externo nulo "
        End If
    End If
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If

'Unidad Organizativa - [RHPro (Sector)]
Texto = "Unidad Organizativa - [RHPro (Sector)] " & Unidad_Organizativa
Sector = CLng(CalcularMapeoInv(Unidad_Organizativa, "T527X", "-1"))
Tenro = 2
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = Sector
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
If Sector <> 0 Then
'    Call Insertar_His_Estructura(Tenro, Sector, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If

'---------------------------------------------------------------
'Estructuras por default
'Contrato
Tenro = 18
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = 5312
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez

'Situacion de Revista
Tenro = 30
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = 530
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez

'Condicion SIJP
Tenro = 31
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = 529
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez

'Siniestro SIJP
Tenro = 42
Estructuras(Tenro).Tenro = Tenro
Estructuras(Tenro).Estrnro = 375
Estructuras(Tenro).Desde = Inicio_Validez
Estructuras(Tenro).Hasta = Fin_Validez
'---------------------------------------------------------------

'Busco el legajo existe ==> Actualizo las estructuras
If ExisteLegajo Then
    'Inserto las estructuras (quedaron pendientes del IT0001)
    For i = 1 To UBound(Estructuras)
        If Estructuras(i).Estrnro <> 0 And Estructuras(i).Estrnro <> -1 Then
            Call Insertar_His_Estructura(Estructuras(i).Tenro, Estructuras(i).Estrnro, Empleado.Tercero, Estructuras(i).Desde, Estructuras(i).Hasta)
        End If
        Infotipo_0001 = True
    Next i
End If

    'cierro y libero
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    Set rs_Estructura = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub


Public Sub Leer_Infotipo_IT0002(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0002. Personal Data.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION            Código Tabla    Nombre Técnico
'PERNR      NUMC                8   Personnel Number
'INFTY      CHAR                6   Constant name infotype
'SUBTY      CHAR                4   Subtipo
'BEGDA      DATS                8   Inicio de Validez
'ENDDA      DATS                8   Fin de Validez
'INITS      CHAR                10  Initials
'NACHN      CHAR                40  Apellidos
'NAME2      CHAR                40  Nombre c/o
'NACH2      CHAR                40  Segundo apellido
'VORNA      CHAR                40  Nombre
'CNAME      CHAR                80  Nombre completo.
'TITEL      CHAR                15  Título
'TITL2      CHAR                15  Second Title
'NAMZU      CHAR                15  Other Title
'VORSW      CHAR                15  Name Prefix
'VORS2      CHAR                15  Second Name Prefix
'RUFNM      CHAR                40  Known as
'MIDNM      CHAR                40  Segundo nombre
'KNZNM      NUMC                2   Name Format Indicator for Employee in a List
'ANRED      CHAR                1   Clave de dirección
'GESCH      CHAR                1   Clave de sexo                                                       GESCH
'GBDAT      DATS                8   Fecha de nacimiento
'GBLND      CHAR                3   País de nacimiento                                              13  T005
'GBDEP      CHAR                3   Estado o Departamento           14          T005S
'GBORT      CHAR                40  Lugar de nacimiento
'NATIO      CHAR                3   Nacionalidad                                                    15  T005
'NATI2      CHAR                3   Segunda nacionalidad            15          T005
'NATI3      CHAR                3   Tercera nacionalidad            15          T005
'SPRSL      LANG                1   Lenguaje de comunicación
'KONFE      CHAR                2   Religious Denomination Key
'FAMST      CHAR                1   Clave para el estado civil                                      17  T502T
'FAMDT      DATS                8   Inicio de validez del estado civil actual
'ANZKD      DEC                 3   Número de hijos
'NACON      CHAR                1   Name Connection
'PERMO      CHAR                2   Modifier for Personnel Identifier
'PERID      CHAR                20  Personnel ID Number
'GBPAS      DATS                8   Date of Birth According to Passport
'FNAMK      CHAR                40  First name (Katakana)
'LNAMK      CHAR                40  Last name (Katakana)
'FNAMR      CHAR                40  First Name (Romaji)
'LNAMR      CHAR                40  Last Name (Romaji)
'NABIK      CHAR                40  Name of Birth (Katakana)
'NABIR      CHAR                40  Name of Birth (Romaji)
'NICKK      CHAR                40  Koseki (Katakana)
'NICKR      CHAR                40  Koseki (Romaji)
'GBJHR      NUMC                4   Year of Birth
'GBMON      NUMC                2   Month of Birth
'GBTAG      NUMC                2   Birth Date (to Month/Year)
'NCHMC      CHAR                25  Last Name (Field for Search Help)
'VNAMC      CHAR                25  First Name (Field for Search Help)
'NAMZ2      CHAR                15  Name Affix for Name at Birth
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte
Dim i As Integer

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Initials As String
Dim Apellidos As String
Dim Nombre_c_o As String
Dim Segundo_apellido As String
Dim nombre As String
Dim nombre_completo As String
Dim Titulo As String
Dim Second_Title As String
Dim Other_Title As String
Dim Name_Prefix As String
Dim Second_Name_Prefix As String
Dim Known_as As String
Dim Segundo_nombre As String
Dim Name_Format_Indicator As String
Dim Clave_de_Direccion As String
Dim Clave_de_sexo As String
Dim Fecha_de_nacimiento As String
Dim Pais_de_nacimiento As String
Dim Estado_o_Departamento As String
Dim Lugar_de_nacimiento As String
Dim Nacionalidad As String
Dim Segunda_Nacionalidad As String
Dim Tercera_Nacionalidad As String
Dim Lenguaje_de_comunicacion As String
Dim Religious_Denomination_Key As String
Dim Clave_Estado_Civil As String
Dim Inicio_Validez_Estado_Civil_Actual As String
Dim Numero_de_Hijos As String
Dim Name_Connection As String
Dim Modifier_for_Personnel_Identifier As String
Dim Personnel_ID_Number As String
Dim Date_of_Birth_Passport As String
Dim First_Name_Katakana As String
Dim Last_Name_Katakana As String
Dim First_Name_Romaji As String
Dim Last_Name_Romaji As String
Dim Name_of_Birth_Katakana As String
Dim Name_of_Birth_Romaji As String
Dim Koseki_Katakana As String
Dim Koseki_Romaji As String
Dim Year_of_Birth As String
Dim Month_of_Birth As String
Dim Birth_Date_Month_Year As String
Dim Last_Name_Search_Help As String
Dim First_Name_Search_Help As String
Dim Name_Affix_Birth As String
Dim FueraDeConvenio As Long
Dim Inserto_estr As Boolean
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0002"
    Columna = 2
    Infotipo_0002 = False
    Fila_Infotipo_0002 = Fila_Infotipo_0002 + 1
    Hoja = 3
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Subtipo = Trim(Mid(strlinea, pos1, pos2))

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Trim(Mid(strlinea, pos1, pos2))
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Trim(Mid(strlinea, pos1, pos2))
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Initials
    Columna = Columna + 1
    Texto = "Initials"
    pos1 = 35
    pos2 = 10
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Initials = Trim(Mid(strlinea, pos1, pos2))

    'Apellidos
    Columna = Columna + 1
    Texto = "Apellidos"
    pos1 = 45
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Apellidos = Trim(Mid(strlinea, pos1, pos2))
    Reg_Tercero.Terape = EliminarCHInvalidos(Trim(Format_Str(Apellidos, 25, True, " ")))

    'Nombre c/o
    Columna = Columna + 1
    Texto = "Nombre c/o"
    pos1 = 85
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Nombre_c_o = Trim(Mid(strlinea, pos1, pos2))
    
    'Segundo Apellido
    Columna = Columna + 1
    Texto = "Segundo Apellido"
    pos1 = 125
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Segundo_apellido = Trim(Mid(strlinea, pos1, pos2))
    Reg_Tercero.Terape2 = EliminarCHInvalidos(Trim(Format_Str(Segundo_apellido, 25, True, " ")))
        
    'Nombre
    Columna = Columna + 1
    Texto = "Nombre"
    pos1 = 165
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    nombre = Trim(Mid(strlinea, pos1, pos2))
    Reg_Tercero.Ternom = EliminarCHInvalidos(Trim(Format_Str(nombre, 25, True, " ")))
    
    'Nombre Completo
    Columna = Columna + 1
    Texto = "Nombre Completo"
    pos1 = 205
    pos2 = 80
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    nombre_completo = Trim(Mid(strlinea, pos1, pos2))

    'Titulo
    Columna = Columna + 1
    Texto = "Titulo"
    pos1 = 285
    pos2 = 15
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Titulo = Trim(Mid(strlinea, pos1, pos2))

    'Second title
    Columna = Columna + 1
    Texto = "Segundo titulo"
    pos1 = 300
    pos2 = 15
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Second_Title = Trim(Mid(strlinea, pos1, pos2))

    'Other title
    Columna = Columna + 1
    Texto = "Other title"
    pos1 = 315
    pos2 = 15
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Other_Title = Trim(Mid(strlinea, pos1, pos2))

    'Name Prefix
    Columna = Columna + 1
    Texto = "Name Prefix"
    pos1 = 330
    pos2 = 15
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Name_Prefix = Trim(Mid(strlinea, pos1, pos2))
    
    'Second Name Prefix
    Columna = Columna + 1
    Texto = "Second Name Prefix"
    pos1 = 345
    pos2 = 15
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Second_Name_Prefix = Trim(Mid(strlinea, pos1, pos2))
    
    'Known as
    Columna = Columna + 1
    Texto = "Known as"
    pos1 = 360
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Known_as = Trim(Mid(strlinea, pos1, pos2))

    'Segundo Nombre
    Columna = Columna + 1
    Texto = "Segundo Nombre"
    pos1 = 400
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Segundo_nombre = Trim(Mid(strlinea, pos1, pos2))
    Reg_Tercero.Ternom2 = EliminarCHInvalidos(Trim(Format_Str(Segundo_nombre, 25, True, " ")))
        
    'Name Format Indicator for Employee in a List
    Columna = Columna + 1
    Texto = "Name Format Indicator for Employee in a List"
    pos1 = 440
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Name_Format_Indicator = Trim(Mid(strlinea, pos1, pos2))

    'Clave de dirección
    Columna = Columna + 1
    Texto = "Clave de dirección"
    pos1 = 442
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Clave_de_Direccion = Trim(Mid(strlinea, pos1, pos2))
    
    'Clave de sexo
    Columna = Columna + 1
    Texto = "Clave de Sexo"
    pos1 = 443
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Clave_de_sexo = Trim(Mid(strlinea, pos1, pos2))
    If Clave_de_sexo = 1 Then   'Masculino
        Reg_Tercero.Tersex = -1
    Else
        Reg_Tercero.Tersex = 0
    End If
    
    'Fecha de nacimiento
    Columna = Columna + 1
    Texto = "Fecha de nacimiento"
    pos1 = 444
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Fecha_de_nacimiento = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Trim(Mid(strlinea, pos1, pos2))
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    Reg_Tercero.Terfecnac = CDate(Fecha_de_nacimiento)
    
    'Pais de Nacimiento
    Columna = Columna + 1
    Texto = "Pais de Nacimiento"
    pos1 = 452
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Pais_de_nacimiento = Trim(Trim(Mid(strlinea, pos1, pos2)))
    If Not EsNulo(Pais_de_nacimiento) Then
        Reg_Tercero.PaisNro = CLng(CalcularMapeoInv(Pais_de_nacimiento, "T005", "0"))
    Else
        Reg_Tercero.PaisNro = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Pais de Nacimiento"
    End If

    'Estado o Departamento
    Columna = Columna + 1
    Texto = "Estado o Departamento"
    pos1 = 455
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Estado_o_Departamento = Trim(Mid(strlinea, pos1, pos2))

    'Lugar de Nacimiento
    Columna = Columna + 1
    Texto = "Lugar de Nacimiento"
    pos1 = 458
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Lugar_de_nacimiento = Trim(Mid(strlinea, pos1, pos2))
    
    'Nacionalidad
    Columna = Columna + 1
    Texto = "Nacionalidad"
    pos1 = 498
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Nacionalidad = Trim(Mid(strlinea, pos1, pos2))
    Reg_Tercero.NacionalNro = CLng(CalcularMapeoInv(Nacionalidad, "T005", "-1"))
    
    'Segunda Nacionalidad
    Columna = Columna + 1
    Texto = "Segunda Nacionalidad"
    pos1 = 501
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Segunda_Nacionalidad = Trim(Mid(strlinea, pos1, pos2))

    'Tercera Nacionalidad
    Columna = Columna + 1
    Texto = "Tercera Nacionalidad"
    pos1 = 504
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Tercera_Nacionalidad = Trim(Mid(strlinea, pos1, pos2))

    'Lenguaje de comunicacion
    Columna = Columna + 1
    Texto = "Lenguaje de comunicacion"
    pos1 = 507
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Lenguaje_de_comunicacion = Trim(Mid(strlinea, pos1, pos2))

    'Religious Denomination Key
    Columna = Columna + 1
    Texto = "Religious Denomination Key"
    pos1 = 508
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Religious_Denomination_Key = Trim(Mid(strlinea, pos1, pos2))
    
    'Clave para el estado civil
    Columna = Columna + 1
    Texto = "Clave para el estado civil"
    pos1 = 510
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Clave_Estado_Civil = Trim(Mid(strlinea, pos1, pos2))
    Reg_Tercero.EstCivNro = CLng(CalcularMapeoInv(Clave_Estado_Civil, "T502T", "-1"))
    
    'Inicio Validez Estado Civil Actual
    Columna = Columna + 1
    Texto = "Inicio Validez Estado Civil Actual"
    pos1 = 511
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Inicio_Validez_Estado_Civil_Actual = Trim(Mid(strlinea, pos1, pos2))

    'Numero de Hijos
    Columna = Columna + 1
    Texto = "Numero de Hijos"
    pos1 = 519
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Numero_de_Hijos = Trim(Mid(strlinea, pos1, pos2))

    'Name Connection
    Columna = Columna + 1
    Texto = "Name Connection"
    pos1 = 522
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Name_Connection = Trim(Mid(strlinea, pos1, pos2))

    'Modifier for Personnel Identifier
    Columna = Columna + 1
    Texto = "Modifier for Personnel Identifier"
    pos1 = 523
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Modifier_for_Personnel_Identifier = Trim(Mid(strlinea, pos1, pos2))
    
    'Personnel ID Number
    Columna = Columna + 1
    Texto = "Personnel ID Number"
    pos1 = 525
    pos2 = 20
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Personnel_ID_Number = Trim(Mid(strlinea, pos1, pos2))
    
    'Date of Birth Passport
    Columna = Columna + 1
    Texto = "Date of Birth Passport"
    pos1 = 545
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Date_of_Birth_Passport = CDate(Mid(strlinea, pos1, pos2))

    'First Name Katakana
    Columna = Columna + 1
    Texto = "First Name Katakana"
    pos1 = 553
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    First_Name_Katakana = Trim(Mid(strlinea, pos1, pos2))

    'Last Name Katakana
    Columna = Columna + 1
    Texto = "Last Name Katakana"
    pos1 = 593
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Last_Name_Katakana = Trim(Mid(strlinea, pos1, pos2))

    'First Name Romaji
    Columna = Columna + 1
    Texto = "First Name Romaji"
    pos1 = 633
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    First_Name_Romaji = Trim(Mid(strlinea, pos1, pos2))
    
    'Last Name Romaji
    Columna = Columna + 1
    Texto = "Last Name Romaji"
    pos1 = 673
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Last_Name_Romaji = Trim(Mid(strlinea, pos1, pos2))
    
    'Name of Birth Katakana
    Columna = Columna + 1
    Texto = "Name of Birth Katakana"
    pos1 = 713
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Name_of_Birth_Katakana = Trim(Mid(strlinea, pos1, pos2))

    'Name of Birth Romaji
    Columna = Columna + 1
    Texto = "Name of Birth Romaji"
    pos1 = 753
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Name_of_Birth_Romaji = Trim(Mid(strlinea, pos1, pos2))

    'Koseki Katakana
    Columna = Columna + 1
    Texto = "Koseki Katakana"
    pos1 = 793
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Koseki_Katakana = Trim(Mid(strlinea, pos1, pos2))

    'Koseki Romaji
    Columna = Columna + 1
    Texto = "Koseki Romaji"
    pos1 = 833
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Koseki_Romaji = Trim(Mid(strlinea, pos1, pos2))
    
    'Year of Birth
    Columna = Columna + 1
    Texto = "Year of Birth"
    pos1 = 873
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Year_of_Birth = Trim(Mid(strlinea, pos1, pos2))
    
    'Month of Birth
    Columna = Columna + 1
    Texto = "Month of Birth"
    pos1 = 877
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Month_of_Birth = Trim(Mid(strlinea, pos1, pos2))

    'Birth Date Month Year
    Columna = Columna + 1
    Texto = "Birth Date Month Year"
    pos1 = 879
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Birth_Date_Month_Year = Trim(Mid(strlinea, pos1, pos2))

    'Last_Name_Search_Help
    Columna = Columna + 1
    Texto = "Last_Name_Search_Help"
    pos1 = 881
    pos2 = 25
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Last_Name_Search_Help = Trim(Mid(strlinea, pos1, pos2))

    'First_Name_Search_Help
    Columna = Columna + 1
    Texto = "First_Name_Search_Help"
    pos1 = 906
    pos2 = 25
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    First_Name_Search_Help = Trim(Mid(strlinea, pos1, pos2))
    
    'Name_Affix_Birth
    Columna = Columna + 1
    Texto = "Name_Affix_Birth"
    pos1 = 931
    pos2 = 15
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0002, Columna, Aux)
    Name_Affix_Birth = Trim(Mid(strlinea, pos1, pos2))
'------------------------------------------------------------------------


    If Not ExisteLegajo Then
        'Inserto el tercero
        StrSql = " INSERT INTO tercero ("
        StrSql = StrSql & " ternom,"
        If Not EsNulo(Reg_Tercero.Ternom2) Then
            StrSql = StrSql & " ternom2,"
        End If
        StrSql = StrSql & " terape,"
        If Not EsNulo(Reg_Tercero.Terape2) Then
            StrSql = StrSql & " terape2,"
        End If
        StrSql = StrSql & " terfecnac,"
        StrSql = StrSql & " tersex,"
        StrSql = StrSql & " estcivnro,"
        StrSql = StrSql & " nacionalnro,"
        StrSql = StrSql & " paisnro"
        StrSql = StrSql & " )"
        
        StrSql = StrSql & " VALUES("
        
        StrSql = StrSql & "'" & Reg_Tercero.Ternom & "'"
        If Not EsNulo(Trim(Reg_Tercero.Ternom2)) Then
            StrSql = StrSql & ",'" & Trim(Reg_Tercero.Ternom2) & "'"
        End If
        StrSql = StrSql & ",'" & Reg_Tercero.Terape & "'"
        If Not EsNulo(Trim(Reg_Tercero.Terape2)) Then
            StrSql = StrSql & ",'" & Trim(Reg_Tercero.Terape2) & "'"
        End If
        StrSql = StrSql & "," & ConvFecha(Reg_Tercero.Terfecnac)
        StrSql = StrSql & "," & Reg_Tercero.Tersex
        StrSql = StrSql & "," & Reg_Tercero.EstCivNro
        StrSql = StrSql & "," & Reg_Tercero.NacionalNro
        StrSql = StrSql & "," & Reg_Tercero.PaisNro
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Empleado.Tercero = getLastIdentity(objConn, "tercero")

        StrSql = " INSERT INTO empleado("
        StrSql = StrSql & " empleg"
        StrSql = StrSql & " ,empfecalta"
'        If Not EsNulo(Fecha_BajaPrevista) Then
'            StrSql = StrSql & " ,empfecbaja"
'        End If
        StrSql = StrSql & " ,empest"
        StrSql = StrSql & " ,empfaltagr"
        StrSql = StrSql & " ,ternro"
        StrSql = StrSql & " ,terape"
        If Not EsNulo(Reg_Tercero.Terape2) Then
            StrSql = StrSql & " ,terape2"
        End If
        StrSql = StrSql & " ,ternom"
        If Not EsNulo(Reg_Tercero.Ternom2) Then
            StrSql = StrSql & " ,ternom2"
        End If
        StrSql = StrSql & " ,empnro"
        StrSql = StrSql & " ) VALUES( "
        StrSql = StrSql & Empleado.Legajo
        StrSql = StrSql & "," & ConvFecha(Fecha_Alta)
'        If Not EsNulo(Fecha_BajaPrevista) Then
'            StrSql = StrSql & "," & ConvFecha(Fecha_BajaPrevista)
'        End If
        StrSql = StrSql & ",-1 "
        StrSql = StrSql & "," & ConvFecha(Fecha_Alta)
        StrSql = StrSql & "," & Empleado.Tercero
        StrSql = StrSql & ",'" & Reg_Tercero.Terape & "'"
        If Not EsNulo(Reg_Tercero.Terape2) Then
            StrSql = StrSql & ",'" & Reg_Tercero.Terape2 & "'"
        End If
        StrSql = StrSql & ",'" & Reg_Tercero.Ternom & "'"
        If Not EsNulo(Reg_Tercero.Ternom2) Then
            StrSql = StrSql & ",'" & Reg_Tercero.Ternom2 & "'"
        End If
        StrSql = StrSql & ",1"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & Empleado.Tercero & ",1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
    Else    'Actualizo los datos
    
        'Tercero
        StrSql = " UPDATE tercero SET "
        StrSql = StrSql & " ternom = '" & Reg_Tercero.Ternom & "'"
        If Not EsNulo(Reg_Tercero.Ternom2) Then
            StrSql = StrSql & " ,ternom2 = '" & Reg_Tercero.Ternom2 & "'"
        End If
        StrSql = StrSql & " ,terape = '" & Reg_Tercero.Terape & "'"
        If Not EsNulo(Reg_Tercero.Terape2) Then
            StrSql = StrSql & " ,terape2 ='" & Reg_Tercero.Terape2 & "'"
        End If
        StrSql = StrSql & " ,terfecnac = " & ConvFecha(Reg_Tercero.Terfecnac)
        StrSql = StrSql & " ,tersex = " & Reg_Tercero.Tersex
        StrSql = StrSql & " ,estcivnro = " & Reg_Tercero.EstCivNro
        StrSql = StrSql & " ,nacionalnro = " & Reg_Tercero.NacionalNro
        StrSql = StrSql & " ,paisnro = " & Reg_Tercero.PaisNro
        StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
    
        'Empleado
        StrSql = " UPDATE empleado SET "
        StrSql = StrSql & " empleg = " & Empleado.Legajo
'        StrSql = StrSql & ",empfecalta = " & ConvFecha(Inicio_Validez)
        If Not EsNulo(Fin_Validez) Then
            StrSql = StrSql & ",empfecbaja = " & ConvFecha(Fin_Validez)
        End If
        StrSql = StrSql & ",empest = -1 "
'        StrSql = StrSql & ",empfaltagr = " & ConvFecha(Inicio_Validez)
        StrSql = StrSql & ",terape = '" & Reg_Tercero.Terape & "'"
        If Not EsNulo(Reg_Tercero.Terape2) Then
            StrSql = StrSql & ",terape2 = '" & Reg_Tercero.Terape2 & "'"
        End If
        StrSql = StrSql & ",ternom = '" & Reg_Tercero.Ternom & "'"
        If Not EsNulo(Reg_Tercero.Ternom2) Then
            StrSql = StrSql & ",ternom2 = '" & Reg_Tercero.Ternom2 & "'"
        End If
        StrSql = StrSql & ",empnro = 1"
        StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
  
    If Not ExisteLegajo Then
        'Inserto las estructuras (quedaron pendientes del IT0001)
        For i = 1 To UBound(Estructuras)
            If Estructuras(i).Estrnro <> 0 And Estructuras(i).Estrnro <> -1 Then
                Call Insertar_His_Estructura(Estructuras(i).Tenro, Estructuras(i).Estrnro, Empleado.Tercero, Estructuras(i).Desde, Estructuras(i).Hasta)
            End If
        Next i
    End If


Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
End Sub


Public Sub Leer_Infotipo_IT0006(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0006. Addresses.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION            Código Tabla    Nombre Técnico
'PERNR      NUMC            8       Personnel Number
'INFTY      CHAR            6       Constant name infotype
'SUBTY      CHAR            4       Subtipo     T591A
'BEGDA      DATS            8       Inicio de Validez
'ENDDA      DATS            8       Fin de Validez
'ANSSA      CHAR            4       Clase de registro de direcciones
'NAME2      CHAR            40      C/O name
'STRAS      CHAR            60      Calle y número
'ORT01      CHAR            40      Población
'ORT02      CHAR            40      Distrito
'PSTLZ      CHAR            10      Código postal
'LAND1      CHAR            3       Clave de país                                                   13  T005
'TELNR      CHAR            14      Número de teléfono
'ENTKM      DEC             3       Distance in Kilometers
'WKWNG      CHAR            1       Company Housing
'BUSRT      CHAR            3       Ruta de Bus
'LOCAT      CHAR            40      Campo adicional p.dirección
'ADR03      CHAR            40      Calle 2
'ADR04      CHAR            40      Calle 3
'STATE      CHAR            3       Región (Estado federal, "land", provincia, condado)             14  T005S
'HSNMR      CHAR            10      Nº (edificio)
'POSTA      CHAR            10      Identificación de una vivienda en una casa
'BLDNG      CHAR            10      Edificio (número o sigla)
'FLOOR      CHAR            10      Floor in building
'STRDS      CHAR            2       Street Abbreviation
'ENTK2      DEC             3       Distance in Kilometers
'COM01      CHAR            4       Communications type
'NUM01      CHAR            20      Communications Number
'COM02      CHAR            4       Communications type
'NUM02      CHAR            20      Communications number
'COM03      CHAR            4       Communications type
'NUM03      CHAR            20      Communications Number
'COM04      CHAR            4       Communication Type
'NUM04      CHAR            20      Communications Number
'COM05      CHAR            4       Communication Type
'NUM05      CHAR            20      Communications Number
'COM06      CHAR            4       Communication Type
'NUM06      CHAR            20      Communications Number
'INDRL      CHAR            2       Indicator for relationship (specification code)
'COUNC      CHAR            3       Country Code
'RCTVC      CHAR            6       Municipal city code
'OR2KK      CHAR            40      Second address line (Katakana)
'CONKK      CHAR            40      Contact Person (Katakana) (Japan)
'OR1KK      CHAR            40      First address line (Katakana)
'RAILW      NUMC            1       Social Subscription Railway
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Clase_de_registro_de_direcciones As String
Dim C_O_name As String
Dim Calle_y_numero As String
Dim Poblacion As String
Dim Distrito As String
Dim Codigo_postal As String
Dim Clave_de_pais As String
Dim Numero_de_telefono As String
Dim Distance_in_Kilometers As String
Dim Company_Housing As String
Dim Ruta_de_Bus As String
Dim Campo_Adicional_Direccion As String
Dim Calle_2 As String
Dim Calle_3 As String
Dim Region As String
Dim Nro_Edificio As String
Dim Identificacion_de_vivienda As String
Dim Edificio As String
Dim Floor_in_building As String
Dim Street_Abbreviation As String
Dim Distance_in_Kilometers2 As String
Dim Communications_Type1 As String
Dim Communications_Number1 As String
Dim Communications_type2 As String
Dim Communications_Number2 As String
Dim Communications_type3 As String
Dim Communications_Number3 As String
Dim Communications_Type4 As String
Dim Communications_Number4 As String
Dim Communications_Type5 As String
Dim Communications_Number5 As String
Dim Communications_Type6 As String
Dim Communications_Number6 As String
Dim Indicator_for_relationship_code As String
Dim Country_code As String
Dim Municipal_city_code As String
Dim Second_address_line As String
Dim Contact_Person As String
Dim First_address_line As String
Dim Social_Subscription_Railway As String

Dim TipoDomi As Long
Dim Nro_Pais As Long
Dim Nro_Localidad As Long
Dim Nro_Provincia As Long
Dim Nro_Partido As Long
Dim NroDom As Long


Dim rs_Tel As New ADODB.Recordset
Dim rs_CabDom As New ADODB.Recordset
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0006"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0006 = False
    Fila_Infotipo_0006 = Fila_Infotipo_0006 + 1
    Hoja = 4
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Subtipo = Trim(Mid(strlinea, pos1, pos2))
    TipoDomi = CLng(CalcularMapeoSubtipo("IT0006", Subtipo, "T591A", "0"))

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

'    'Clase de registro de direcciones
'    pos1 = pos1 + pos2
'    pos2 = 4
'    Clase_de_registro_de_direcciones = Mid(strLinea, pos1, pos2)
'
'    'C/O Name
'    pos1 = pos1 + pos2
'    pos2 = 40
'    C_O_name = Mid(strLinea, pos1, pos2)
    For Columna = 6 To 7
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, "")
    Next Columna

    'Calle y Numero
    Columna = 8
    pos1 = 79
    pos2 = 60
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Calle_y_numero = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

    'Poblacion
    Columna = 9
    pos1 = pos1 + pos2
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Poblacion = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

    'Distrito
    Columna = 10
    pos1 = pos1 + pos2
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Distrito = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

    'Codigo Postal
    Columna = 11
    pos1 = pos1 + pos2
    pos2 = 10
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Codigo_postal = Trim(Mid(strlinea, pos1, pos2))

    'Clave de Pais
    Columna = 12
    pos1 = pos1 + pos2
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Clave_de_pais = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Clave_de_pais) Then
        Nro_Pais = CLng(CalcularMapeoInv(Clave_de_pais, "T005", "-1"))
    Else
        Nro_Pais = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Clave de Pais"
    End If
        
    'Numero de Telefono
    Columna = 13
    pos1 = pos1 + pos2
    pos2 = 14
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Numero_de_telefono = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

'    'Distance in Kilometers
'    pos1 = pos1 + pos2
'    pos2 = 3
'    Distance_in_Kilometers = Mid(strLinea, pos1, pos2)
'
'    'Company Housing
'    pos1 = pos1 + pos2
'    pos2 = 1
'    Company_Housing = Mid(strLinea, pos1, pos2)
'
'    'Ruta de Bus
'    pos1 = pos1 + pos2
'    pos2 = 3
'    Ruta_de_Bus = Mid(strLinea, pos1, pos2)
    For Columna = 14 To 16
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, "")
    Next Columna
    
    'Campo adicional dirección
    Columna = 17
    pos1 = 253
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Campo_Adicional_Direccion = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

    'Calle 2
    Columna = 18
    pos1 = pos1 + pos2
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Calle_2 = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

    'Calle 3
    Columna = 19
    pos1 = pos1 + pos2
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Calle_3 = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

    'Region
    Columna = 20
    pos1 = pos1 + pos2
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Region = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Region) Then
        Nro_Provincia = CLng(CalcularMapeoInv(Region, "T005S", "-1"))
    Else
        Nro_Provincia = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Region"
    End If
    
    'Nro Edificio
    Columna = 21
    pos1 = pos1 + pos2
    pos2 = 10
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Nro_Edificio = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))
    
    'Identificacion de vivienda
    Columna = 22
    pos1 = pos1 + pos2
    pos2 = 10
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Identificacion_de_vivienda = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

    'Edificio
    Columna = 23
    pos1 = pos1 + pos2
    pos2 = 10
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Edificio = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

    'Floor in building
    Columna = 24
    pos1 = pos1 + pos2
    pos2 = 10
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, Aux)
    Floor_in_building = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))

    'Completo las columnas vacias o que no tienen importancia
    For Columna = 25 To 45
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0006, Columna, "")
    Next Columna
'-------------------------------------------------------------------------

    If Not EsNulo(Trim(Poblacion)) Then
        Call ValidarLocalidad(Poblacion, Nro_Localidad, Nro_Pais, Nro_Provincia)
    Else
        Nro_Localidad = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Region"
    End If
    
    If Distrito <> "N/A" And Not EsNulo(Trim(Distrito)) Then
        Call ValidarPartido(Distrito, Nro_Partido)
    Else
        Nro_Partido = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Region"
    End If

    If Not ExisteLegajo Then
        'Inserto el Domicilio
        If (TipoDomi <> 0 And Nro_Localidad <> 0 And Nro_Provincia <> 0 And Nro_Pais <> 0) Then
            'StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
            'StrSql = StrSql & " VALUES(" & TipoDomi & "," & Empleado.Tercero & ",-1,2)"
            StrSql = " INSERT INTO cabdom(tipnro,tidonro,ternro,domdefault) "
            StrSql = StrSql & " VALUES(1," & TipoDomi & "," & Empleado.Tercero & ",-1)"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            NroDom = getLastIdentity(objConn, "cabdom")
          
            StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,entrecalles,"
            StrSql = StrSql & "locnro,provnro,paisnro,partnro) "
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & NroDom
            StrSql = StrSql & ",'" & Format_Str(Calle_y_numero, 30, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Nro_Edificio, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Floor_in_building, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Identificacion_de_vivienda, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Edificio, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Nro_Edificio, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Codigo_postal, 12, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Campo_Adicional_Direccion, 80, False, "") & "'"
            StrSql = StrSql & "," & Nro_Localidad
            StrSql = StrSql & "," & Nro_Provincia
            StrSql = StrSql & "," & Nro_Pais
            StrSql = StrSql & "," & Nro_Partido
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
          
          
            If Numero_de_telefono <> "" And TipoDomi = 2 Then
              StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
              StrSql = StrSql & " VALUES(" & NroDom & ",'" & Numero_de_telefono & "',0,-1,0)"
              objConn.Execute StrSql, , adExecuteNoRecords
            Else
                If Numero_de_telefono <> "" And TipoDomi = 4 Then
                      StrSql = "SELECT * FROM telefono "
                      StrSql = StrSql & " WHERE domnro =" & NroDom
                      StrSql = StrSql & " AND telnro ='" & Numero_de_telefono & "'"
                      If rs_Tel.State = adStateOpen Then rs_Tel.Close
                      OpenRecordset StrSql, rs_Tel
                      If rs_Tel.EOF Then
                          StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                          StrSql = StrSql & " VALUES(" & NroDom & ",'" & Numero_de_telefono & "',0,0,-1)"
                          objConn.Execute StrSql, , adExecuteNoRecords
                      End If
                Else
                    If Numero_de_telefono <> "" Then
                      StrSql = "SELECT * FROM telefono "
                      StrSql = StrSql & " WHERE domnro =" & NroDom
                      StrSql = StrSql & " AND telnro ='" & Numero_de_telefono & "'"
                      If rs_Tel.State = adStateOpen Then rs_Tel.Close
                      OpenRecordset StrSql, rs_Tel
                      If rs_Tel.EOF Then
                          StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                          StrSql = StrSql & " VALUES(" & NroDom & ",'" & Numero_de_telefono & "',0,0,0)"
                          objConn.Execute StrSql, , adExecuteNoRecords
                      End If
                    End If
                End If
            End If
        End If
    Else 'Actualizo
        'Domicilio
        If (Nro_Localidad <> 0 And Nro_Provincia <> 0 And Nro_Pais <> 0) Then
            StrSql = " SELECT * FROM cabdom "
            'StrSql = StrSql & " WHERE tipnro = " & TipoDomi & " AND domdefault = -1 AND tidonro = 2 "
            StrSql = StrSql & " WHERE tipnro = 1 AND domdefault = -1 AND tidonro = " & TipoDomi
            StrSql = StrSql & " AND ternro = " & Empleado.Tercero
            If rs_CabDom.State = adStateOpen Then rs_CabDom.Close
            OpenRecordset StrSql, rs_CabDom
            If rs_CabDom.EOF Then
                'StrSql = " INSERT INTO cabdom(tipnro,ternro,domdefault,tidonro) "
                'StrSql = StrSql & " VALUES(" & TipoDomi & "," & Empleado.Tercero & ",-1,2)"
                StrSql = " INSERT INTO cabdom(tipnro,tidonro,ternro,domdefault) "
                StrSql = StrSql & " VALUES(1," & TipoDomi & "," & Empleado.Tercero & ",-1)"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                NroDom = getLastIdentity(objConn, "cabdom")
                
                StrSql = " INSERT INTO detdom(domnro,calle,nro,piso,oficdepto,torre,manzana,codigopostal,entrecalles,"
                StrSql = StrSql & "locnro,provnro,paisnro,partnro) "
                StrSql = StrSql & " VALUES ("
                StrSql = StrSql & NroDom
                StrSql = StrSql & ",'" & Format_Str(Calle_y_numero, 30, False, "") & "'"
                StrSql = StrSql & ",'" & Format_Str(Nro_Edificio, 8, False, "") & "'"
                StrSql = StrSql & ",'" & Format_Str(Floor_in_building, 8, False, "") & "'"
                StrSql = StrSql & ",'" & Format_Str(Identificacion_de_vivienda, 8, False, "") & "'"
                StrSql = StrSql & ",'" & Format_Str(Edificio, 8, False, "") & "'"
                StrSql = StrSql & ",'" & Format_Str(Nro_Edificio, 8, False, "") & "'"
                StrSql = StrSql & ",'" & Format_Str(Codigo_postal, 12, False, "") & "'"
                StrSql = StrSql & ",'" & Format_Str(Campo_Adicional_Direccion, 80, False, "") & "'"
                StrSql = StrSql & "," & Nro_Localidad
                StrSql = StrSql & "," & Nro_Provincia
                StrSql = StrSql & "," & Nro_Pais
                StrSql = StrSql & "," & Nro_Partido
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                
                If Numero_de_telefono <> "" And TipoDomi = 2 Then
                  StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                  StrSql = StrSql & " VALUES(" & NroDom & ",'" & Numero_de_telefono & "',0,-1,0)"
                  objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    If Numero_de_telefono <> "" And TipoDomi = 4 Then
                          StrSql = "SELECT * FROM telefono "
                          StrSql = StrSql & " WHERE domnro =" & NroDom
                          StrSql = StrSql & " AND telnro ='" & Numero_de_telefono & "'"
                          If rs_Tel.State = adStateOpen Then rs_Tel.Close
                          OpenRecordset StrSql, rs_Tel
                          If rs_Tel.EOF Then
                              StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                              StrSql = StrSql & " VALUES(" & NroDom & ",'" & Numero_de_telefono & "',0,0,-1)"
                              objConn.Execute StrSql, , adExecuteNoRecords
                          End If
                    Else
                        If Numero_de_telefono <> "" Then
                          StrSql = "SELECT * FROM telefono "
                          StrSql = StrSql & " WHERE domnro =" & NroDom
                          StrSql = StrSql & " AND telnro ='" & Numero_de_telefono & "'"
                          If rs_Tel.State = adStateOpen Then rs_Tel.Close
                          OpenRecordset StrSql, rs_Tel
                          If rs_Tel.EOF Then
                              StrSql = " INSERT INTO telefono(domnro,telnro,telfax,teldefault,telcelular) "
                              StrSql = StrSql & " VALUES(" & NroDom & ",'" & Numero_de_telefono & "',0,0,0)"
                              objConn.Execute StrSql, , adExecuteNoRecords
                          End If
                        End If
                    End If
                End If
            Else
                NroDom = rs_CabDom!Domnro
              
                StrSql = " UPDATE detdom SET "
                StrSql = StrSql & " calle = '" & Format_Str(Calle_y_numero, 30, False, "") & "'"
                StrSql = StrSql & " ,nro = '" & Format_Str(Nro_Edificio, 8, False, "") & "'"
                StrSql = StrSql & " ,piso = '" & Format_Str(Floor_in_building, 8, False, "") & "'"
                StrSql = StrSql & " ,oficdepto = '" & Format_Str(Identificacion_de_vivienda, 8, False, "") & "'"
                StrSql = StrSql & " ,torre = '" & Format_Str(Edificio, 8, False, "") & "'"
                StrSql = StrSql & " ,manzana = '" & Format_Str(Nro_Edificio, 8, False, "") & "'"
                StrSql = StrSql & " ,codigopostal = '" & Format_Str(Codigo_postal, 12, False, "") & "'"
                StrSql = StrSql & " ,entrecalles = '" & Format_Str(Campo_Adicional_Direccion, 80, False, "") & "'"
                StrSql = StrSql & " ,locnro = " & Nro_Localidad
                StrSql = StrSql & " ,provnro = " & Nro_Provincia
                StrSql = StrSql & " ,paisnro = " & Nro_Pais
                StrSql = StrSql & " ,partnro = " & Nro_Partido
                StrSql = StrSql & " WHERE domnro = " & NroDom
                objConn.Execute StrSql, , adExecuteNoRecords
                
                If Numero_de_telefono <> "" And TipoDomi = 2 Then
                    StrSql = " UPDATE telefono SET "
                    StrSql = StrSql & " telnro = '" & Numero_de_telefono & "'"
                    StrSql = StrSql & " WHERE domnro = " & NroDom
                    StrSql = StrSql & " AND teldefault = -1 "
                    StrSql = StrSql & " AND telcelular = 0 "
                    StrSql = StrSql & " AND telfax = 0 "
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    If Numero_de_telefono <> "" And TipoDomi = 4 Then
                        StrSql = " UPDATE telefono SET "
                        StrSql = StrSql & " telnro = '" & Numero_de_telefono & "'"
                        StrSql = StrSql & " WHERE domnro = " & NroDom
                        StrSql = StrSql & " AND teldefault = 0 "
                        StrSql = StrSql & " AND telcelular = -1 "
                        StrSql = StrSql & " AND telfax = 0 "
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        If Numero_de_telefono <> "" Then
                            StrSql = " UPDATE telefono SET "
                            StrSql = StrSql & " telnro = '" & Numero_de_telefono & "'"
                            StrSql = StrSql & " WHERE domnro = " & NroDom
                            StrSql = StrSql & " AND teldefault = 0 "
                            StrSql = StrSql & " AND telcelular = 0 "
                            StrSql = StrSql & " AND telfax = 0 "
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
            End If
        End If
    End If

Exit Sub
Manejador_De_Error:
    HuboError = True
    'Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub


Public Sub Leer_Infotipo_IT0007(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0007. Planned Working Time.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion: por ahora no sa hace porque no trae nada relevante
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Regla_Plan_Horario_de_Trabajo As String
Dim Status_Empleado As String
Dim Porcentaje_Horario_de_Trabajo As String
Dim Horas_Mensuales As String
Dim Horas_Semanales As String
Dim Horas_de_Trabajo_diarias As String
Dim Dias_Laborales_Semanales As String
Dim Horas_de_Trabajo_Anuales As String
Dim Empleado_a_Tiempo_Parcial As String
Dim Horas_de_Trabajo_Minimas_dia As String
Dim Maximo_Horas_x_Dia As String
Dim Minimo_Horas_x_Dia As String
Dim Maximo_Horas_x_Semana As String
Dim Minimo_Horas_x_Mes As String
Dim Maximo_Horas_x_Mes As String
Dim Minimo_Horas_x_Ano As String
Dim Maximo_Horas_x_Ano As String
Dim Plan_Horario_de_Trabajo_diario As String
Dim Indicador_Adicional As String
Dim Semana_de_Trabajo As String


'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

'    Flog.Writeline Espacios(Tabulador * 2) & "Infotipo 0002"
'    Columna = 2
'    Infotipo_0007 = False
'
'    'Subtipo
'    Columna = Columna + 1
'    Texto = "Subtipo"
'    pos1 = 15
'    pos2 = 4
'    Subtipo = Mid(strLinea, pos1, pos2)
'
'    'Inicio de Validez
'    Columna = Columna + 1
'    Texto = "Inicio de Validez"
'    pos1 = 19
'    pos2 = 8
'    Inicio_Validez = StrToFecha(Mid(strLinea, pos1, pos2), Ok)
'    If Not Ok Then
'        Flog.Writeline Espacios(Tabulador * 2) & ""
'        FlogE.Writeline Espacios(Tabulador * 2) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strLinea, pos1, pos2)
'        InsertaError Columna, 8
'        HuboError = True
'        Exit Sub
'    End If
'
'    'Fin de Validez
'    Columna = Columna + 1
'    Texto = "Fin de Validez"
'    pos1 = 27
'    pos2 = 8
'    Fin_Validez = StrToFecha(Mid(strLinea, pos1, pos2), Ok)
'    If Not Ok Then
'        Flog.Writeline Espacios(Tabulador * 2) & ""
'        FlogE.Writeline Espacios(Tabulador * 2) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strLinea, pos1, pos2)
'        InsertaError Columna, 8
'        HuboError = True
'        Exit Sub
'    End If
'
'    'Regla para plan horario de trabajo
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Regla_Plan_Horario_de_Trabajo = Mid(strLinea, pos1, pos2)
'
'    'Status del empleado para gestión de tiempos
'    pos1 = pos1 + pos2
'    pos2 = 1
'    Status_Empleado = Mid(strLinea, pos1, pos2)
'
'    'Porcentaje de horario de trabajo
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Porcentaje_Horario_de_Trabajo = Mid(strLinea, pos1, pos2)
'
'    'horas mensuales
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Horas_Mensuales = Mid(strLinea, pos1, pos2)
'
'    'horas semanales
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Horas_Semanales = Mid(strLinea, pos1, pos2)
'
'    'Horas de trabajo diarias
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Horas_de_Trabajo_diarias = Mid(strLinea, pos1, pos2)
'
'    'Días laborales semanales
'    pos1 = pos1 + pos2
'    pos2 = 7
'    Dias_Laborales_Semanales = Mid(strLinea, pos1, pos2)
'
'    'Horas de trabajo anuales
'    pos1 = pos1 + pos2
'    pos2 = 10
'    Horas_de_Trabajo_Anuales = Mid(strLinea, pos1, pos2)
'
'    'Casilla selección empleado a tiempo parcial
'    pos1 = pos1 + pos2
'    pos2 = 1
'    Empleado_a_Tiempo_Parcial = Mid(strLinea, pos1, pos2)
'
'    'Horas de trabajo mínimas por dia
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Horas_de_Trabajo_Minimas_dia = Mid(strLinea, pos1, pos2)
'
'    'Máximo de horas de trabajo por dia
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Maximo_Horas_x_Dia = Mid(strLinea, pos1, pos2)
'
'    'Minimo de horas de trabajo por dia
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Minimo_Horas_x_Dia = Mid(strLinea, pos1, pos2)
'
'    'Maximo de horas de trabajo por semana
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Maximo_Horas_x_Semana = Mid(strLinea, pos1, pos2)
'
'    'Minimo de horas de trabajo por mes
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Minimo_Horas_x_Mes = Mid(strLinea, pos1, pos2)
'
'    'Máximo de horas de trabajo por mes
'    pos1 = pos1 + pos2
'    pos2 = 8
'    Maximo_Horas_x_Mes = Mid(strLinea, pos1, pos2)
'
'    'Mínimo de horas de trabajo por ano
'    pos1 = pos1 + pos2
'    pos2 = 10
'    Minimo_Horas_x_Ano = Mid(strLinea, pos1, pos2)
'
'    'Maximo de horas de trabajo por ano
'    pos1 = pos1 + pos2
'    pos2 = 10
'    Maximo_Horas_x_Ano = Mid(strLinea, pos1, pos2)
'
'    'Creación dinamica del plan de horario de trabajo diario
'    pos1 = pos1 + pos2
'    pos2 = 1
'    Plan_Horario_de_Trabajo_diario = Mid(strLinea, pos1, pos2)
'
'    'Indicador adicional para gestión de tiempos
'    pos1 = pos1 + pos2
'    pos2 = 2
'    Indicador_Adicional = Mid(strLinea, pos1, pos2)
'
'    'Semana de trabajo
'    pos1 = pos1 + pos2
'    pos2 = 2
'    Semana_de_Trabajo = Mid(strLinea, pos1, pos2)
'-------------------------------------------------------------------------

    On Error GoTo Manejador_De_Error
    
'por ahora no hago nada ya que no veo ningun campo que resulte de utilidad

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub


Public Sub Leer_Infotipo_IT0008(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte
Dim i As Integer

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Sueldo_Anual As Single
Dim Remuneracion As Single
Dim CC_Nominas As TNomina
Dim Termino As Boolean
Dim Clase_Convenio_Colectivo As String
Dim Area_Convenio_Colectivo As String
Dim Grupo_Profesional As String
Dim Subgrupo_Profesional As String
Dim Categoria As String
Dim Nro_Categoria As Long

Dim Inserto_estr As Boolean
Dim rs_Empleado As New ADODB.Recordset
Dim Hoja As Integer
Dim Convenio As Long
Dim GrupoLiq As Long

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0008"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0008 = False
    Fila_Infotipo_0008 = Fila_Infotipo_0008 + 1
    Hoja = 5
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    For Columna = 6 To 8
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, "")
    Next Columna

    'Grupo profesional
    Columna = 9
    pos1 = 41
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
    Grupo_Profesional = Trim(Mid(strlinea, pos1, pos2))

    'subgrupo profesional
    Columna = 10
    pos1 = 49
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
    Subgrupo_Profesional = Trim(Mid(strlinea, pos1, pos2))
    
    ' *******************************************************
    'De estos 2 campos anteriores sale la categoria
    If Grupo_Profesional = "1 AUX 1°" And Subgrupo_Profesional = "1" Then Categoria = "03"
    If Grupo_Profesional = "1 AUX 2°" And Subgrupo_Profesional = "2" Then Categoria = "04"
    If Grupo_Profesional = "1 AUX 3°" And Subgrupo_Profesional = "3" Then Categoria = "05"
    If Grupo_Profesional = "1 AUX 4°" And Subgrupo_Profesional = "4" Then Categoria = "06"
    If Grupo_Profesional = "1A OPN/C" And Subgrupo_Profesional = "5" Then Categoria = "12"
    If Grupo_Profesional = "1B OPS/C" And Subgrupo_Profesional = "6" Then Categoria = "11"
    If Grupo_Profesional = "1C OPC/F" And Subgrupo_Profesional = "7" Then Categoria = "10"
    If Grupo_Profesional = "1D OPC/E" And Subgrupo_Profesional = "8" Then Categoria = "09"
    If Grupo_Profesional = "1E OPC/O" And Subgrupo_Profesional = "9" Then Categoria = "08"
    If Grupo_Profesional = "1F OPP/Q" And Subgrupo_Profesional = "10" Then Categoria = "13"
    If Grupo_Profesional = "1G OPC/T" And Subgrupo_Profesional = "11" Then Categoria = "07"
    If Grupo_Profesional = "2 ENTRAN" And Subgrupo_Profesional = "12" Then Categoria = "50"
    If Grupo_Profesional = "2 ENTRAN" And Subgrupo_Profesional = "13" Then Categoria = "50"
    If Grupo_Profesional = "2 ENTRAN" And Subgrupo_Profesional = "14" Then Categoria = "50"
    If Grupo_Profesional = "2 ENTRAN" And Subgrupo_Profesional = "15" Then Categoria = "50"
    If Grupo_Profesional = "2 ENTRAN" And Subgrupo_Profesional = "16" Then Categoria = "50"
    If Grupo_Profesional = "2 ENTRAN" And Subgrupo_Profesional = "17" Then Categoria = "50"
    If Grupo_Profesional = "2 ESPEC." And Subgrupo_Profesional = "18" Then Categoria = "50"
    If Grupo_Profesional = "2 ESPEC." And Subgrupo_Profesional = "19" Then Categoria = "50"
    If Grupo_Profesional = "2 ESPEC." And Subgrupo_Profesional = "20" Then Categoria = "50"
    If Grupo_Profesional = "2 ESPEC." And Subgrupo_Profesional = "21" Then Categoria = "50"
    If Grupo_Profesional = "2 ESPEC." And Subgrupo_Profesional = "22" Then Categoria = "50"
    If Grupo_Profesional = "2 ESPEC." And Subgrupo_Profesional = "23" Then Categoria = "50"
    If Grupo_Profesional = "2 JUNIOR" And Subgrupo_Profesional = "24" Then Categoria = "50"
    If Grupo_Profesional = "2 JUNIOR" And Subgrupo_Profesional = "25" Then Categoria = "50"
    If Grupo_Profesional = "2 JUNIOR" And Subgrupo_Profesional = "26" Then Categoria = "50"
    If Grupo_Profesional = "2 JUNIOR" And Subgrupo_Profesional = "27" Then Categoria = "50"
    If Grupo_Profesional = "2 JUNIOR" And Subgrupo_Profesional = "28" Then Categoria = "50"
    If Grupo_Profesional = "2 JUNIOR" And Subgrupo_Profesional = "29" Then Categoria = "50"
    If Grupo_Profesional = "2 MEDIOD" And Subgrupo_Profesional = "30" Then Categoria = "50"
    If Grupo_Profesional = "2 MEDIOD" And Subgrupo_Profesional = "31" Then Categoria = "50"
    If Grupo_Profesional = "2 MEDIOD" And Subgrupo_Profesional = "32" Then Categoria = "50"
    If Grupo_Profesional = "2 MEDIOD" And Subgrupo_Profesional = "33" Then Categoria = "50"
    If Grupo_Profesional = "2 MEDIOD" And Subgrupo_Profesional = "34" Then Categoria = "50"
    If Grupo_Profesional = "2 MEDIOD" And Subgrupo_Profesional = "35" Then Categoria = "50"
    If Grupo_Profesional = "2 PLENO" And Subgrupo_Profesional = "36" Then Categoria = "50"
    If Grupo_Profesional = "2 PLENO" And Subgrupo_Profesional = "37" Then Categoria = "50"
    If Grupo_Profesional = "2 PLENO" And Subgrupo_Profesional = "38" Then Categoria = "50"
    If Grupo_Profesional = "2 PLENO" And Subgrupo_Profesional = "39" Then Categoria = "50"
    If Grupo_Profesional = "2 PLENO" And Subgrupo_Profesional = "40" Then Categoria = "50"
    If Grupo_Profesional = "2 PLENO" And Subgrupo_Profesional = "41" Then Categoria = "50"
    If Grupo_Profesional = "2 SENIOR" And Subgrupo_Profesional = "42" Then Categoria = "50"
    If Grupo_Profesional = "2 SENIOR" And Subgrupo_Profesional = "43" Then Categoria = "50"
    If Grupo_Profesional = "2 SENIOR" And Subgrupo_Profesional = "44" Then Categoria = "50"
    If Grupo_Profesional = "2 SENIOR" And Subgrupo_Profesional = "45" Then Categoria = "50"
    If Grupo_Profesional = "2 SENIOR" And Subgrupo_Profesional = "46" Then Categoria = "50"
    If Grupo_Profesional = "2 SENIOR" And Subgrupo_Profesional = "47" Then Categoria = "50"
    If Grupo_Profesional = "3AMEDIOD" And Subgrupo_Profesional = "48" Then Categoria = "00"
    If Grupo_Profesional = "3BSUPERV" And Subgrupo_Profesional = "49" Then Categoria = "02"
    If Grupo_Profesional = "3C JEFE" And Subgrupo_Profesional = "50" Then Categoria = "17"
    If Grupo_Profesional = "3DENTRAN" And Subgrupo_Profesional = "51" Then Categoria = "00"
    If Grupo_Profesional = "3EJUNIOR" And Subgrupo_Profesional = "52" Then Categoria = "00"
    If Grupo_Profesional = "3FPLENO" And Subgrupo_Profesional = "53" Then Categoria = "00"
    If Grupo_Profesional = "3GGERENT" And Subgrupo_Profesional = "54" Then Categoria = "01"
    If Grupo_Profesional = "3HSENIOR" And Subgrupo_Profesional = "55" Then Categoria = "00"
    If Grupo_Profesional = "3I FC" And Subgrupo_Profesional = "56" Then Categoria = "00"
    If Grupo_Profesional = "3JRCH GR" And Subgrupo_Profesional = "57" Then Categoria = "00"
    If Grupo_Profesional = "3KRCH GR" And Subgrupo_Profesional = "58" Then Categoria = "00"
    If Grupo_Profesional = "3LRX GD" And Subgrupo_Profesional = "59" Then Categoria = "00"
    If Grupo_Profesional = "3MRX GDN" And Subgrupo_Profesional = "60" Then Categoria = "00"
    If Grupo_Profesional = "3NRX GR" And Subgrupo_Profesional = "61" Then Categoria = "00"
    If Grupo_Profesional = "3ODIRECT" And Subgrupo_Profesional = "62" Then Categoria = "18"
    
    
    'FGZ - 07/07/2005
    Select Case Categoria
    Case "00":  'Fuera de Convenio
        Convenio = 5238
        GrupoLiq = 294
    Case "01":  'Fuera de Convenio
        Convenio = 5238
        GrupoLiq = 294
    Case "02":  'Fuera de Convenio
        Convenio = 5238
        GrupoLiq = 294
    Case "09":  'APM
        Convenio = 5239
        GrupoLiq = 439
    Case "17":  'Fuera de Convenio
        Convenio = 5238
        GrupoLiq = 294
    Case "18":  'Fuera de Convenio
        Convenio = 5238
        GrupoLiq = 294
    Case "50":  'ATSA
        Convenio = 5240
        GrupoLiq = 440
    End Select
    
    'Categoria
    Call ValidaEstructura(3, Categoria, Nro_Categoria, Inserto_estr)
    Texto = "Categoria - [ RHPRO(Categoria)] " & Grupo_Profesional & "-" & Subgrupo_Profesional
    'Nro_Categoria = CLng(CalcularMapeoInv(Categoria, "T547V", "0"))
    If Nro_Categoria <> 0 Then
        Call Insertar_His_Estructura(3, Nro_Categoria, Empleado.Tercero, Inicio_Validez, Fin_Validez)
        'FGZ - 07/07/2005
        Call Insertar_His_Estructura(19, Convenio, Empleado.Tercero, Inicio_Validez, Fin_Validez)
        Call Insertar_His_Estructura(32, GrupoLiq, Empleado.Tercero, Inicio_Validez, Fin_Validez)
    Else
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
    End If
    ' *******************************************************
    
    For Columna = 11 To 21
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, "")
    Next Columna
    
    'Sueldo Anual
    Columna = 22
    pos1 = 105
    pos2 = 18
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
    Sueldo_Anual = CSng(Mid(strlinea, pos1, pos2))
    Remuneracion = Sueldo_Anual / 12
    
    'a continuacion vienen 20 cc-nominas
    pos1 = 134
    i = 1
    Columna = 24
    Termino = False
    Do While i <= 20 And Not Termino
        'CC-Nomina
        Columna = Columna + 1
        Aux = Mid(strlinea, pos1, 4)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
        CC_Nominas.Nomina = Trim(Aux)
        pos1 = pos1 + 4
        
        'Importe CC-Nomina
        Columna = Columna + 1
        Aux = Trim(Mid(strlinea, pos1, 14))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
        If IsNumeric(Aux) Then
            CC_Nominas.Monto = CSng(Mid(Aux, 2, 11) & "." & Mid(Aux, 13, 2))
        Else
            CC_Nominas.Monto = 0
        End If
        If Mid(Aux, 1, 1) = "-" Then
            CC_Nominas.Monto = CC_Nominas.Monto * -1
        End If
        pos1 = pos1 + 14
        
        'Numero
        Columna = Columna + 1
        Aux = Mid(strlinea, pos1, 8)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
        Aux = Trim(Mid(strlinea, pos1, 8))
        CC_Nominas.Cantidad = CSng(Mid(Aux, 2, 5) & "." & Mid(Aux, 7, 2))
        If Mid(Aux, 1, 1) = "-" Then
            CC_Nominas.Cantidad = CC_Nominas.Cantidad * -1
        End If
        pos1 = pos1 + 9
        
        'Unidad de Medida/Tiempo
        Columna = Columna + 1
        Aux = Mid(strlinea, pos1, 3)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
        Aux = Trim(Mid(strlinea, pos1, 3))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
        If Not EsNulo(Aux) Then
            CC_Nominas.Unidad = CLng(CalcularMapeoInv(Aux, "T538A", "0"))
        Else
            CC_Nominas.Unidad = 2
        End If
        'Actualizo la cantidad de acuerdo a la unidad de medida
        CC_Nominas.Cantidad = Calcular_Cantidad(CC_Nominas.Cantidad, CC_Nominas.Unidad)
        pos1 = pos1 + 3
        
        'Indicador de Operacion para CC-Nomina
        Columna = Columna + 1
        Aux = Mid(strlinea, pos1, 1)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, Aux)
        Aux = Trim(Mid(strlinea, pos1, 1))
        CC_Nominas.Operacion = Aux
        'pos1 = pos1 + 1
        
        If Not EsNulo(CC_Nominas.Nomina) And (CC_Nominas.Monto <> 0 Or CC_Nominas.Cantidad <> 0) Then
            Call Insertar_Novedad(CC_Nominas.Nomina, CC_Nominas.Monto, CC_Nominas.Cantidad, Inicio_Validez, Fin_Validez, "IT0008")
        Else
            If EsNulo(CC_Nominas.Nomina) Then
                Termino = True
            End If
        End If
        
        i = i + 1
    Loop
    'este infotipo solo viene en un hiring ==> .....
    
    'Completo las columnas vacias o que no tienen importancia
    For Columna = 125 To 146
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0008, Columna, "")
    Next Columna
    

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    
End Sub


Public Sub Borrar_Infotipo0014(ByVal Tercero As Long, ByVal FechaDesde, ByVal FechaHasta)
' ---------------------------------------------------------------------------------------------
' Descripcion: Depura todas las novedades cargadas del empleado para el periodo especificado.
' Autor      : FGZ
' Fecha      : 23/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

        StrSql = "DELETE FROM novemp WHERE "
        StrSql = StrSql & " empleado = " & Tercero
'        StrSql = StrSql & " AND nevigencia = -1 "
'        StrSql = StrSql & " AND (nedesde >= " & ConvFecha(FechaDesde)
'        StrSql = StrSql & " AND nehasta <= " & ConvFecha(FechaHasta) & ")"
        StrSql = StrSql & " AND netexto = 'IT0014'"
        objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub Leer_Infotipo_IT0009(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0009. Banked Details.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD             DESCRIPCION           Código Tabla    Nombre Técnico
'PERNR      NUMC        8                   Personnel Number
'INFTY      CHAR        6                   Constant name infotype
'SUBTY      CHAR        4                   Subtipo                                         T591A
'BEGDA      DATS        8                   Inicio de Validez
'ENDDA      DATS        8                   Fin de Validez
'OPKEN      CHAR        1                   Operation Indicator for Wage Types
'BETRG      CURR        13.2                Valor prefijado
'WAERS      CUKY        5                   Clave de moneda                   19            TCURC
'ANZHL      DEC         5.2                 Porcentaje prefijado
'ZEINH      CHAR        3                   Unidad de medida/Tiempo                         T538A
'BNKSA      CHAR        4                   Clase de registro de relación bancaria
'ZLSCH      CHAR        1                   Vía de pago                                     T042Z
'EMFTX      CHAR        40                  Texto para el receptor
'BKPLZ      CHAR        10                  Código postal
'BKORT      CHAR        25                  Población
'BANKS      CHAR        3                   Clave de país del banco                                         13  T005
'BANKL      CHAR        15                  Clave de banco                                                  18  T012
'BANKN      CHAR        18                  Nº cuenta bancaria
'BANKP      CHAR        2                   Dígito de control Código bancario/cuenta
'BKONT      CHAR        2                   Clave de control de bancos
'SWIFT      CHAR        11                  Código SWIFT para pagos internacionales
'DTAWS      CHAR        2                   Clave de instrucción para ISD
'DTAMS      CHAR        1                   Clave de notificación para ISD
'STCD1      CHAR        16                  Número de identificación fiscal 1
'STCD2      CHAR        11                  Número de identificación fiscal suplementario
'PSKTO      CHAR        16                  Nº cta.giro caja postal
'ESRNR      CHAR        11                  Nº de usuario ESR
'ESRRE      CHAR        27                  Nº referencia ESR
'ESRPZ      CHAR        2                   Dígito de control ESR
'EMFSL      CHAR        8                   Clave de receptor para transferencias
'ZWECK      CHAR        40                  Destino para transferencias
'BTTYP      NUMC        2                   Transfer Type
'PAYTY      CHAR        1                   Tipo de Pago
'PAYID      CHAR        1                   Identificador Del pago
'OCRSN      CHAR        4                   Reason for Off-Cycle Payroll
'BONDT      DATS        8                   Off-cycle payroll payment date
'BKREF      CHAR        20                  Reference specifications for bank details
'STRAS      CHAR        30                  House number and street
'STATE      CHAR        3                   Region (State, Province, County)
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Operation_Indicator As String
Dim Valor_prefijado As String
Dim Clave_de_Moneda As String
Dim Porcentaje_prefijado As Integer
Dim Unidad_de_MedidaTiempo As String
Dim Clase_registro_relacion_bancaria As String
Dim Via_de_pago As String
Dim Texto_receptor As String
Dim Codigo_postal As String
Dim Poblacion As String
Dim Clave_de_pais As String
Dim Clave_de_banco As String
Dim Nro_cuenta_bancaria As String
Dim Digito_control_Codigo_bancario As String
Dim Clave_control_de_bancos As String
Dim Codigo_SWIFT As String
Dim Clave_de_instruccion_para_ISD As String
Dim Clave_de_notificacion_para_ISD As String
Dim Numero_de_id_fiscal_1 As String
Dim Numero_de_id_fiscal_suplementario As String
Dim Nro_cta_caja_postal As String
Dim Nro_de_usuario_ESR As String
Dim Nro_referencia_ESR As String
Dim Digito_de_control_ESR As String
Dim Clave_de_receptor_para_transferencias As String
Dim Destino_para_transferencias As String
Dim CBU As String
Dim Transfer_Type As String
Dim Tipo_de_Pago As String
Dim Identificador_Del_pago As String
Dim Reason_for_Off_Cycle_Payroll As String
Dim Off_cycle_payroll_payment_date As String
Dim Reference_specifications As String
Dim House_Number_And_street As String
Dim Region As String

Dim Nro_FormaPago As Long
Dim Nro_Bancopago As Long
Dim NroCuenta As String
Dim Nro_Moneda As Long
Dim Nro_Banco As Long

Dim rs_Cta As New ADODB.Recordset
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0009"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0009 = False
    Fila_Infotipo_0009 = Fila_Infotipo_0009 + 1
    Hoja = 6
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

'    'Operation Indicator for Wage Types
'    pos1 = pos1 + pos2
'    pos2 = 1
'    Operation_Indicator = Mid(strLinea, pos1, pos2)
'
    Columna = 6
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, "")

    'Valor prefijado
    Columna = 7
    pos1 = 36
    pos2 = 14   '13.2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Valor_prefijado = CSng(Mid(Aux, 2, 12) & "." & Mid(Aux, 13, 2))
    
    'Clave de moneda
    Columna = 8
    pos1 = 50
    pos2 = 5
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Clave_de_Moneda = Trim(Mid(strlinea, pos1, pos2))
    Nro_Moneda = CLng(CalcularMapeoInv(Clave_de_Moneda, "TCURC", "-1"))
        
    'Porcentaje prefijado
    Columna = 9
    pos1 = 55
    pos2 = 6    '5.2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    Porcentaje_prefijado = CInt(CSng(Mid(Aux, 2, 4) & "." & Mid(Aux, 5, 2)))
    If Porcentaje_prefijado = 0 Then
        Porcentaje_prefijado = 100
    End If
    
    'Unidad de medida/Tiempo
    Columna = 10
    pos1 = 61
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Unidad_de_MedidaTiempo = Trim(Mid(strlinea, pos1, pos2))

'    'Clase de registro de relación bancaria
'    pos1 = pos1 + pos2
'    pos2 = 4
'    Clase_registro_relacion_bancaria = Mid(strLinea, pos1, pos2)
    Columna = 11
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, "")
    
    'Vía de pago
    Columna = 12
    pos1 = 68
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Via_de_pago = Mid(strlinea, pos1, pos2)
    Nro_FormaPago = CLng(CalcularMapeoInv(Via_de_pago, "T042Z", "0"))

    For Columna = 13 To 16
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, "")
    Next Columna

    'Clave de banco
    Columna = 17
    pos1 = 147
    pos2 = 15
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Clave_de_banco = Trim(Mid(strlinea, pos1, pos2))
    Nro_Banco = CLng(CalcularMapeoInv(Clave_de_banco, "T012", "-1"))
    
    'Nº cuenta bancaria - BANKN
    Columna = 18
    pos1 = 162
    pos2 = 18
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Nro_cuenta_bancaria = Trim(Mid(strlinea, pos1, pos2))
       
    For Columna = 19 To 30
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, "")
    Next Columna
       
    'Destino para transferencias (primera parte del CBU) - ZWECK
    Columna = 31
    pos1 = 289
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, Aux)
    Destino_para_transferencias = Trim(Mid(strlinea, pos1, pos2))
    'CBU = Left(Replace(Destino_para_transferencias, "-", "") & "-" & Replace(Nro_cuenta_bancaria, "-", ""), 30)
    'CBU = Left(Replace(Nro_cuenta_bancaria, "-", "") & Replace(Destino_para_transferencias, "-", ""), 22)
    
    'FGZ -
    If Not EsNulo(Destino_para_transferencias) Then
        CBU = Left(Left(Replace(Destino_para_transferencias, "-", ""), 8) & "-" & Left(Replace(Nro_cuenta_bancaria, "-", ""), 14), 22)
    Else
        CBU = ""
    End If
    
    'Completo las columnas vacias o que no tienen importancia
    For Columna = 32 To 39
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0009, Columna, "")
    Next Columna
    
'-------------------------------------------------------------------------
'Revisar si existe el banco
'teoricamente deberia existir dado que antes de correr los infotipos se qctualizan las tablas planas
        
        'Inserto la cuenta bancaria
        If (Nro_FormaPago <> 0 And Nro_Banco <> 0 And Nro_cuenta_bancaria <> "") Then
        
            'primero deberia buscar si existe
            StrSql = "SELECT * FROM ctabancaria "
            StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
            'StrSql = StrSql & " AND fpagnro = " & Nro_FormaPago
            'StrSql = StrSql & " AND banco = " & Nro_Banco
            StrSql = StrSql & " AND ctabestado = -1 "
            OpenRecordset StrSql, rs_Cta
            If rs_Cta.EOF Then
                StrSql = " INSERT INTO ctabancaria (ternro,fpagnro,banco,ctabestado,ctabnro,ctabcbu,ctabporc"
                StrSql = StrSql & " ) VALUES ("
                StrSql = StrSql & Empleado.Tercero
                StrSql = StrSql & "," & Nro_FormaPago
                StrSql = StrSql & "," & Nro_Banco
                StrSql = StrSql & "," & "-1"
                StrSql = StrSql & ",'" & Replace(Nro_cuenta_bancaria, "-", "") & "'"
                StrSql = StrSql & ",'" & Replace(CBU, "-", "") & "'"
                StrSql = StrSql & "," & Porcentaje_prefijado
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                If rs_Cta!ctabnro = Nro_cuenta_bancaria Then
                    StrSql = " UPDATE ctabancaria SET "
                    StrSql = StrSql & " ctabcbu = '" & Replace(CBU, "-", "") & "'"
                    StrSql = StrSql & " ,ctabporc = " & Porcentaje_prefijado
                    StrSql = StrSql & " WHERE cbnro = " & rs_Cta!cbnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    'Desactivo la anterior
                    StrSql = " UPDATE ctabancaria SET "
                    StrSql = StrSql & " ctabestado = 0 "
                    StrSql = StrSql & " WHERE cbnro = " & rs_Cta!cbnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'inserto la nueva
                    StrSql = " INSERT INTO ctabancaria (ternro,fpagnro,banco,ctabestado,ctabnro,ctabcbu,ctabporc"
                    StrSql = StrSql & " ) VALUES ("
                    StrSql = StrSql & Empleado.Tercero
                    StrSql = StrSql & "," & Nro_FormaPago
                    StrSql = StrSql & "," & Nro_Banco
                    StrSql = StrSql & "," & "-1"
                    StrSql = StrSql & ",'" & Replace(Nro_cuenta_bancaria, "-", "") & "'"
                    StrSql = StrSql & ",'" & Replace(CBU, "-", "") & "'"
                    StrSql = StrSql & "," & Porcentaje_prefijado
                    StrSql = StrSql & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If

'cierro y libero
If rs_Cta.State = adStateOpen Then rs_Cta.Close
Set rs_Cta = Nothing

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub

Public Sub Leer_Infotipo_IT0014(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0014. Devengos y Deducciones periodicas.
' Autor      : FGZ
' Fecha      : 08/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Periodo
Dim Fin_Periodo
Dim Clave_de_Moneda As String
Dim CC_Nominas As TNomina
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0014"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0014 = False
    Fila_Infotipo_0014 = Fila_Infotipo_0014 + 1
    Hoja = 7
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio Periodo
    Columna = Columna + 1
    Texto = "Inicio Periodo"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, Aux)
    Inicio_Periodo = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin Peiodo
    Columna = Columna + 1
    Texto = "Fin Periodo"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, Aux)
    Fin_Periodo = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'CC-Nomina
    Columna = 6
    pos1 = 35
    Aux = Mid(strlinea, pos1, 4)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, Aux)
    Aux = Mid(strlinea, pos1, 4)
    CC_Nominas.Nomina = Trim(Aux)
    pos1 = pos1 + 4
    
    'Indicador de Operacion para CC-Nomina
    Columna = 7
    Aux = Mid(strlinea, pos1, 1)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, 1))
    CC_Nominas.Operacion = Aux
    pos1 = pos1 + 1
    
    'Importe CC-Nomina
    Columna = 8
    Aux = Mid(strlinea, pos1, 14)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, 14))
    If IsNumeric(Aux) Then
        CC_Nominas.Monto = CSng(Mid(Aux, 2, 11) & "." & Mid(Aux, 13, 2))
    Else
        CC_Nominas.Monto = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Importe CC-Nomina"
    End If
    If Mid(Aux, 1, 1) = "-" Then
        CC_Nominas.Monto = CC_Nominas.Monto * -1
    End If
    pos1 = pos1 + 14
    
    'Clave de Moneda
    Columna = 9
    Aux = Mid(strlinea, pos1, 5)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, Aux)
    Clave_de_Moneda = Trim(Mid(strlinea, pos1, 5))
    pos1 = pos1 + 5
    
    'Numero
    Columna = 10
    Aux = Mid(strlinea, pos1, 8)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, 8))
    CC_Nominas.Cantidad = CSng(Mid(Aux, 2, 5) & "." & Mid(Aux, 7, 2))
    If Mid(Aux, 1, 1) = "-" Then
        CC_Nominas.Cantidad = CC_Nominas.Cantidad * -1
    End If
    pos1 = pos1 + 8
    
    'Unidad de Medida/Tiempo
    Columna = 11
    Aux = Mid(strlinea, pos1, 3)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, 3))
    If Not EsNulo(Aux) Then
        CC_Nominas.Unidad = CLng(CalcularMapeoInv(Aux, "T538A", "0"))
    Else
        CC_Nominas.Unidad = 2
    End If
    'Actualizo la cantidad de acuerdo a la unidad de medida
    CC_Nominas.Cantidad = Calcular_Cantidad(CC_Nominas.Cantidad, CC_Nominas.Unidad)
    pos1 = pos1 + 3
    
    'Completo las columnas vacias o que no tienen importancia
    For Columna = 12 To 19
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0014, Columna, "")
    Next Columna
    
'    'Caso particular para la nomina 1176
'    'Si viene un 1 en cantidad el sistema tiene que calcular el 10 o 20 segun corresponda
'    ' y sino tiene el 1 se liquida el monto que informen.
'    If CStr(CC_Nominas.Nomina) = "1176" Then
'       If CSng(CC_Nominas.Cantidad) = 1 Then
'          If CLng(CC_Nominas.Unidad) = 10 Then
'             CC_Nominas.Monto = -10
'          Else
'             CC_Nominas.Monto = -20
'          End If
'       End If
'    End If
    
    'Toda la informacion Grabada anteriormente(en actualizaciones de PU12 anteriores) debe ser borrada
    ' Cada PU12 trae la inf completa para el periodo
'    If PrimeraVez_Infotipo_0014 Then
'        PrimeraVez_Infotipo_0014 = False
'        Call Borrar_Infotipo0014(Empleado.Tercero, Inicio_Periodo, Fin_Periodo)
'    End If
    
    'Inserto la novedad
    If Not EsNulo(CC_Nominas.Nomina) And (CC_Nominas.Monto <> 0 Or CC_Nominas.Cantidad <> 0) Then
        Call Insertar_Novedad(CC_Nominas.Nomina, CC_Nominas.Monto, CC_Nominas.Cantidad, Inicio_Periodo, Fin_Periodo, "IT0014")
    End If
        
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
        
End Sub

Public Sub Leer_Infotipo_IT0015(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0015. Devengos complementarios.
' Autor      : FGZ
' Fecha      : 08/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Periodo
Dim Fin_Periodo
Dim Clave_de_Moneda As String
Dim CC_Nominas As TNomina
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0015"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0015 = False
    Fila_Infotipo_0015 = Fila_Infotipo_0015 + 1
    Hoja = 8
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio Periodo
    Columna = Columna + 1
    Texto = "Inicio Periodo"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, Aux)
    Inicio_Periodo = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin Peiodo
    Columna = Columna + 1
    Texto = "Fin Periodo"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, Aux)
    Fin_Periodo = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'CC-Nomina
    Columna = 6
    pos1 = 35
    Aux = Mid(strlinea, pos1, 4)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, Aux)
    Aux = Mid(strlinea, pos1, 4)
    CC_Nominas.Nomina = Trim(Aux)
    pos1 = pos1 + 4
    
    'Indicador de Operacion para CC-Nomina
    Columna = 7
    Aux = Mid(strlinea, pos1, 1)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, 1))
    CC_Nominas.Operacion = Aux
    pos1 = pos1 + 1
    
    'Importe CC-Nomina
    Columna = 8
    Aux = Mid(strlinea, pos1, 14)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, 14))
    If IsNumeric(Aux) Then
        CC_Nominas.Monto = CSng(Mid(Aux, 2, 11) & "." & Mid(Aux, 13, 2))
    Else
        CC_Nominas.Monto = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Importe CC-Nomina"
    End If
    If Mid(Aux, 1, 1) = "-" Then
        CC_Nominas.Monto = CC_Nominas.Monto * -1
    End If
    
    pos1 = pos1 + 14
    
    'Clave de Moneda
    Columna = 9
    Aux = Mid(strlinea, pos1, 5)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, Aux)
    Clave_de_Moneda = Trim(Mid(strlinea, pos1, 5))
    pos1 = pos1 + 5
    
    'Numero
    Columna = 10
    Aux = Mid(strlinea, pos1, 8)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, 8))
    CC_Nominas.Cantidad = CSng(Mid(Aux, 2, 5) & "." & Mid(Aux, 7, 2))
    If Mid(Aux, 1, 1) = "-" Then
        CC_Nominas.Cantidad = CC_Nominas.Cantidad * -1
    End If
    
    pos1 = pos1 + 8
    
    'Unidad de Medida/Tiempo
    Columna = 11
    Aux = Mid(strlinea, pos1, 3)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, 3))
    If Not EsNulo(Aux) Then
        CC_Nominas.Unidad = CLng(CalcularMapeoInv(Aux, "T538A", "0"))
    Else
        CC_Nominas.Unidad = 2
    End If
    'Actualizo la cantidad de acuerdo a la unidad de medida
    CC_Nominas.Cantidad = Calcular_Cantidad(CC_Nominas.Cantidad, CC_Nominas.Unidad)
    pos1 = pos1 + 3
    
    'Completo las columnas vacias o que no tienen importancia
    For Columna = 12 To 18
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0015, Columna, "")
    Next Columna
    
    'Inserto la novedad
    If Not EsNulo(CC_Nominas.Nomina) And (CC_Nominas.Monto <> 0 Or CC_Nominas.Cantidad <> 0) Then
        Call Insertar_Novedad(CC_Nominas.Nomina, CC_Nominas.Monto, CC_Nominas.Cantidad, Inicio_Periodo, Fin_Periodo, "IT0015")
    End If
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub


Public Sub Leer_Infotipo_IT0016(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0016. Contract Elements.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD           DESCRIPCION                             Código Tabla        Nombre Técnico
'PERNR      NUMC        8               Personnel Number
'INFTY      CHAR        6               Constant name infotype
'SUBTY      CHAR        4               Subtipo
'BEGDA      DATS        8               Inicio de Validez
'ENDDA      DATS        8               Fin de Validez
'NBTGK      CHAR        1               Actividad secundaria
'WTTKL      CHAR        1               Cláusula de no competencia
'LFZFR      DEC         3               Plazo p.continuación pago del salario (cantidad)
'LFZZH      CHAR        3               Plazo p.continuación pago del salario (unidad)                      T538A
'LFZSO      NUMC        2               Salario de enfermedad Regla especial
'KGZFR      DEC         3               Plazo p.plus del subsidio de enfermedad (cantidad)
'KGZZH      CHAR        3               Plazo p.plus del subsidio de enfermedad (unidad)                    T538A
'PRBZT      DEC         3               Período de prueba (cantidad)
'PRBEH      CHAR        3               Período de prueba (unidad)                                          T538A
'KDGFR      CHAR        2               Plazo de preaviso Empresario
'KDGF2      CHAR        2               Plazo de preaviso Empleado
'ARBER      DATS        8               Final permiso de trabajo
'EINDT      DATS        8               Primera alta
'KONDT      DATS        8               Fecha de alta en grupo de empresas
'KONSL      CHAR        2               Clv.gr.empresas
'CTTYP      CHAR        2               Clase de contrato                                   36              T547V
'CTEDT      DATS        8               Fecha expir.contrato
'PERSG      CHAR        1               Employee Group                                      5               T501
'PERSK      CHAR        2               Employee Subgroup                                   6               T503K
'WRKPL      CHAR        40              Work location (Contract Elements infotype)
'CTBEG      DATS        8               Start of contract
'CTNUM      CHAR        20              Contract number
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
'Dim Actividad_Secundaria As String
'Dim Clausula As String
'Dim Plazo_Salario_Cantidad As String
'Dim Plazo_Salario_Unidad As String
'Dim Salario_de_Enfermedad As String
'Dim Plazo_plus_Subsidio_Enfermedad_Cantidad As String
'Dim Plazo_plus_Subsidio_Enfermedad_Unidad As String
'Dim Periodo_de_Prueba_Cantidad As String
'Dim Periodo_de_Prueba_Unidad As String
'Dim Plazo_Preaviso_Empresario As String
'Dim Plazo_Preaviso_Empleado As String
'Dim Final_Permiso_de_Trabajo As String
'Dim Primera_Alta As String
'Dim Fecha_Alta_empresas As String
'Dim Clv_Empresas As String
Dim Clase_de_Contrato As String
'Dim Fecha_Expir_Contrato As String
'Dim Employee_Group As String
'Dim Employee_Subgroup As String
'Dim Work_Location As String
'Dim Start_of_Contract As String
Dim Contract_Number As String

Dim Nro_Contrato As Long

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0016"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0016 = False
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Clase de contrato
    pos1 = 86
    pos2 = 2
    Clase_de_Contrato = Mid(strlinea, pos1, pos2)
        
    'Contract Number
    pos1 = pos1 + pos2
    pos2 = 20
    Contract_Number = Mid(strlinea, pos1, pos2)

'----------------------------------------------------------------------------
''Clase de Contrato - [ RHPro(Contrato Actual)]
'Texto = "Clase de Contrato - [ RHPro(Contrato Actual)] " & Clase_de_Contrato
'Nro_Contrato = CLng(CalcularMapeoInv(Clase_de_Contrato, "T547V", "0"))
'If Nro_Contrato <> 0 Then
'    Call Insertar_His_Estructura(18, Nro_Contrato, Empleado.Tercero, Inicio_Validez, Fin_Validez)
'Else
'    Flog.Writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
'    FlogE.Writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
'End If


' no se hace

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub


Public Sub Leer_Infotipo_IT0021(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0021. Family/Related Person.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION                            Código Tabla    Nombre Técnico
'PERNR      NUMC                8   Personnel Number
'INFTY      CHAR                6   Constant name infotype
'SUBTY      CHAR                4   Subtipo     T591A
'BEGDA      DATS                8   Inicio de Validez
'ENDDA      DATS                8   Fin de Validez
'OBJPS      CHAR                2   Identificación de objeto
'FAMSA      CHAR                4   Clase de registro de familia
'FGBDT      DATS                8   Fecha de nacimiento
'FGBLD      CHAR                3   País de nacimiento                               13         T005
'FANAT      CHAR                3   Nacionalidad                                     13         T005
'FASEX      CHAR                1   Clave de sexo
'FAVOR      CHAR                40  Nombre de pila
'FANAM      CHAR                40  Apellidos
'FGBOT      CHAR                40  Lugar de nacimiento
'FGDEP      CHAR                3   Estado federado                                  14         T005S
'ERBNR      CHAR                12  No. Empleado Sustituto de Pensión
'FGBNA      CHAR                40  Name at Birth
'FNAC2      CHAR                40  Segunda fecha de nacimiento
'FCNAM      CHAR                80  Nombre completo
'FKNZN      NUMC                2   Name Format Indicator for Employee in a List
'FINIT      CHAR                10  Initials
'FVRSW      CHAR                15  Name Prefix
'FVRS2      CHAR                15  Name Prefix
'FNMZU      CHAR                15  Other Title
'AHVNR      CHAR                11  PDS number
'KDSVH      CHAR                2   Relationship to child
'KDBSL      CHAR                2   Extra pay entitlement
'KDUTB      CHAR                2   Address of child
'KDGBR      CHAR                2   Child allowance entitlement
'KDART      CHAR                2   Child type
'KDZUG      CHAR                2   Child bonuses
'KDZUL      CHAR                2   Child allowances
'KDVBE      CHAR                2   Sickness certificate entitlement
'ERMNR      CHAR                8   Authority number
'AUSVL      NUMC                4   1st part of SI number (sequential number)
'AUSVG      NUMC                8   End of family member's education/training
'FASDT      DATS                8   End of family member's education/training
'FASAR      CHAR                2   Nivel de estudios del miembro de la familia
'FASIN      CHAR                20  Educational institute
'EGAGA      CHAR                8   Employer of family member
'FANA2      CHAR                3   Segunda Nacionalidad                               13  T005
'FANA3      CHAR                3   Tercera nacionalidad                               13  T005
'BETRG      CURR                9.2 Amount
'TITEL      CHAR                15  Título
'EMRGN      CHAR                1   Emergency contact
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Identificacion_de_objeto As String
Dim Clase_registro_familia As String
Dim Fecha_de_nacimiento As String
Dim Pais_de_nacimiento As String
Dim Nacionalidad As String
Dim Clave_de_sexo As String
Dim Nombre_de_pila As String
Dim Apellidos As String
Dim Lugar_de_nacimiento As String
Dim Estado_federado As String
Dim Nro_Empleado_Sustituto_de_Pension As String
Dim Name_at_Birth As String
Dim Segunda_fecha_de_nacimiento As String
Dim nombre_completo As String
Dim Name_Format_Indicator_Employee As String
Dim Initials As String
Dim Name_Prefix As String
Dim Name_Prefix2 As String
Dim Other_Title As String
Dim PDS_Number As String
Dim Relationship_to_child As String
Dim Extra_pay_entitlement As String
Dim Address_of_child As String
Dim Child_allowance_entitlement As String
Dim Child_type As String
Dim Child_bonuses As String
Dim Child_allowances As String
Dim Sickness_certificate_entitlement As String
Dim Authority_Number As String
Dim sequential_number As String
Dim family_members_education_training As String
Dim family_members_education_training2 As String
Dim Nivel_de_estudios As String
Dim Educational_institute As String
Dim Employer_of_family_member As String
Dim Segunda_Nacionalidad As String
Dim Tercera_Nacionalidad As String
Dim Amount As String
Dim Titulo As String
Dim Emergency_contact As String

Dim Parentesco As Long
Dim Estado_civil As Long
Dim Tersex As Integer
Dim PaisNro As Long
Dim NacionalNro As Long
Dim Aux_Tercero As Long

Dim rs_Tercero As New ADODB.Recordset
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0021"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0021 = False
    Fila_Infotipo_0021 = Fila_Infotipo_0021 + 1
    Hoja = 9
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Subtipo = Trim(Mid(strlinea, pos1, pos2))
    Parentesco = CLng(CalcularMapeoSubtipo("IT0021", Subtipo, "T591A", "0"))
    
    'cargo el estado civil por default
        Select Case Parentesco
        Case "2", "02":  'Hijo
            Estado_civil = 1
        Case "3", "03":  'Cónyuge
            Estado_civil = 2
        Case "4", "04":  'Conviviente
            Estado_civil = 7
        Case "5", "05":  'Menor Bajo Guarda/Tutela
            Estado_civil = 1
        Case "6", "06":  'Adherente
            Estado_civil = 7
        Case "7", "07":  'Hermano
            Estado_civil = 7
        Case "10": 'PreNatal
            Estado_civil = 1
        Case "11": 'Otro
            Estado_civil = 7
        Case "12": 'A Cargo
            Estado_civil = 7
        Case Else
            Estado_civil = 7
        End Select

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Identificación de objeto (este es el nro del hijo)
    Columna = 6
    pos1 = 35
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Identificacion_de_objeto = Trim(Mid(strlinea, pos1, pos2))
    If EsNulo(Identificacion_de_objeto) Then
        Identificacion_de_objeto = 0
    End If
        
    'Clase de registro de familia
    Columna = 7
    pos1 = 37
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Clase_registro_familia = Trim(Mid(strlinea, pos1, pos2))
        
    'Fecha de nacimiento
    Columna = 8
    pos1 = 41
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Texto = "Fecha de nacimiento"
    Fecha_de_nacimiento = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
        
    'País de nacimiento
    Columna = 9
    pos1 = 49
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Pais_de_nacimiento = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Pais_de_nacimiento) Then
        PaisNro = CLng(CalcularMapeoInv(Pais_de_nacimiento, "T005", "0"))
    Else
        PaisNro = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. País de nacimiento"
    End If
    
    'Nacionalidad
    Columna = 10
    pos1 = pos1 + pos2
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Nacionalidad = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Nacionalidad) Then
        NacionalNro = CLng(CalcularMapeoInv(Nacionalidad, "T005", "0"))
    Else
        NacionalNro = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Nacionalidad"
    End If
    
    'Si alguno de los dos es 0 ==> toma el valor del otro
    If (PaisNro <> 0 Or NacionalNro <> 0) And PaisNro <> NacionalNro And (PaisNro = 0 Or NacionalNro = 0) Then
        If PaisNro = 0 Then
            PaisNro = NacionalNro
        Else
            NacionalNro = PaisNro
        End If
    End If
    
    'Clave de sexo
    Columna = 11
    pos1 = pos1 + pos2
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Clave_de_sexo = Mid(strlinea, pos1, pos2)
    If Clave_de_sexo = 1 Then   'Masculino
        Tersex = -1
    Else
        Tersex = 0
    End If
    
    'Nombre de pila
    Columna = 12
    pos1 = pos1 + pos2
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Nombre_de_pila = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))
        
    'Apellidos
    Columna = 13
    pos1 = pos1 + pos2
    pos2 = 40
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Apellidos = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))
        
    For Columna = 14 To 29
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Next Columna
        
    'Child type
    Columna = 30
    pos1 = 427
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Child_type = Trim(Mid(strlinea, pos1, pos2))
        
    'child bonuses
    Columna = 31
    pos1 = pos1 + pos2
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Child_bonuses = Trim(Mid(strlinea, pos1, pos2))
        
    For Columna = 32 To 37
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Next Columna
        
    'Nivel de estudios del miembro de la familia
    Columna = 38
    pos1 = 463
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Nivel_de_estudios = Trim(Mid(strlinea, pos1, pos2))
            
    For Columna = 39 To 44
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Next Columna
            
    'Emergency contact
    Columna = 45
    pos1 = 526
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0021, Columna, Aux)
    Emergency_contact = Trim(Mid(strlinea, pos1, pos2))

'-----------------------------------------------------------------
'Como se si el pariente existe o no???
' por ahora solo se inserta cuando es un hiring... mas adelante vemos como tomar las modificaciones

'If UCase(Accion) = "ALTA" Then
'        'Inserto el tercero
'        StrSql = " INSERT INTO tercero ("
'        StrSql = StrSql & " ternom,"
'        StrSql = StrSql & " terape,"
'        StrSql = StrSql & " terfecnac,"
'        StrSql = StrSql & " tersex,"
'        StrSql = StrSql & " nacionalnro,"
'        StrSql = StrSql & " paisnro"
'        StrSql = StrSql & " )"
'
'        StrSql = StrSql & " VALUES("
'        StrSql = StrSql & "'" & Format_Str(Nombre_de_pila, 25, False, "") & "'"
'        StrSql = StrSql & ",'" & Format_Str(Apellidos, 25, False, "") & "'"
'        StrSql = StrSql & "," & ConvFecha(Fecha_de_nacimiento)
'        StrSql = StrSql & "," & Tersex
'        StrSql = StrSql & "," & NacionalNro
'        StrSql = StrSql & "," & PaisNro
'        StrSql = StrSql & ")"
'        objConn.Execute StrSql, , adExecuteNoRecords
'        Aux_Tercero = getLastIdentity(objConn, "tercero")
'
'        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & Aux_Tercero & ",3)"
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        'crear el complemento
'        StrSql = "INSERT INTO familiar (empleado,ternro,parenro,famnrocorr,famest,famtrab,famestudia,famcernac) "
'        StrSql = StrSql & " VALUES ( "
'        StrSql = StrSql & Empleado.Tercero
'        StrSql = StrSql & "," & Aux_Tercero
'        StrSql = StrSql & "," & Parentesco
'        StrSql = StrSql & "," & Identificacion_de_objeto
'        StrSql = StrSql & ",-1" 'estado
'        StrSql = StrSql & ",0"  'trabaja
'        StrSql = StrSql & "," & IIf(EsNulo(Nivel_de_estudios), 0, -1) 'estudia
'        StrSql = StrSql & ",0"
'        StrSql = StrSql & " )"
'        objConn.Execute StrSql, , adExecuteNoRecords
'Else
    'Es una modificacion o cambio de algo ==> Actualizo
    StrSql = "SELECT * FROM tercero "
    StrSql = StrSql & " INNER JOIN ter_tip ON tercero.ternro = ter_tip.ternro AND ter_tip.tipnro = 3 "
    StrSql = StrSql & " INNER JOIN familiar ON familiar.ternro = tercero.ternro AND familiar.empleado = " & Empleado.Tercero
    StrSql = StrSql & " WHERE ternom = '" & Format_Str(Nombre_de_pila, 25, False, "") & "'"
    StrSql = StrSql & " AND terape = '" & Format_Str(Apellidos, 25, False, "") & "'"
    StrSql = StrSql & " AND terape = '" & Format_Str(Apellidos, 25, False, "") & "'"
    OpenRecordset StrSql, rs_Tercero
    
    If rs_Tercero.EOF Then
        'Inserto el tercero
        StrSql = " INSERT INTO tercero ("
        StrSql = StrSql & " ternom,"
        StrSql = StrSql & " terape,"
        StrSql = StrSql & " terfecnac,"
        StrSql = StrSql & " tersex,"
        StrSql = StrSql & " nacionalnro,"
        StrSql = StrSql & " paisnro,estcivnro"
        StrSql = StrSql & " )"
        
        StrSql = StrSql & " VALUES("
        StrSql = StrSql & "'" & Format_Str(Nombre_de_pila, 25, False, "") & "'"
        StrSql = StrSql & ",'" & Format_Str(Apellidos, 25, False, "") & "'"
        StrSql = StrSql & "," & ConvFecha(Fecha_de_nacimiento)
        StrSql = StrSql & "," & Tersex
        StrSql = StrSql & "," & NacionalNro
        StrSql = StrSql & "," & PaisNro
        StrSql = StrSql & "," & Estado_civil
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Aux_Tercero = getLastIdentity(objConn, "tercero")

        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & Aux_Tercero & ",3)"
        objConn.Execute StrSql, , adExecuteNoRecords

        'crear el complemento
        StrSql = "INSERT INTO familiar (empleado,ternro,parenro,famnrocorr,famest,famtrab,famestudia,famcernac) "
        StrSql = StrSql & " VALUES ( "
        StrSql = StrSql & Empleado.Tercero
        StrSql = StrSql & "," & Aux_Tercero
        StrSql = StrSql & "," & Parentesco
        StrSql = StrSql & "," & Identificacion_de_objeto
        StrSql = StrSql & ",-1" 'estado
        StrSql = StrSql & ",0"  'trabaja
        StrSql = StrSql & "," & IIf(EsNulo(Nivel_de_estudios), 0, -1) 'estudia
        StrSql = StrSql & ",0"
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
        'Flog.Writeline Espacios(Tabulador * 3) & "Fliar no encontrado " & Format_Str(Nombre_de_pila, 25, False, "") & ", " & Format_Str(Apellidos, 25, False, "")
        'Flog.Writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado. "
    Else
        StrSql = "UPDATE tercero SET "
        StrSql = StrSql & " terfecnac = " & ConvFecha(Fecha_de_nacimiento)
        StrSql = StrSql & " ,tersex = " & Tersex
        StrSql = StrSql & " ,nacionalnro = " & NacionalNro
        StrSql = StrSql & " ,paisnro = " & PaisNro
        StrSql = StrSql & " WHERE ternro = " & rs_Tercero!ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "UPDATE familiar SET "
        StrSql = StrSql & " parenro = " & Parentesco
        StrSql = StrSql & " ,famestudia = " & IIf(EsNulo(Nivel_de_estudios), 0, -1) 'estudia
        StrSql = StrSql & " WHERE empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND ternro = " & rs_Tercero!ternro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
'End If

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline


End Sub


Public Sub Leer_Infotipo_IT0023(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0023. Other/Previous Employers.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION                Código Tabla    Nombre Técnico
'PERNR      NUMC            8       Personnel Number
'INFTY      CHAR            6       Constant name infotype
'SUBTY      CHAR            4       Subtipo
'BEGDA      DATS            8       Inicio de Validez
'ENDDA      DATS            8       Fin de Validez
'ARBGB      CHAR            20      Nombre del empleador
'ORT01      CHAR            25      City
'LAND1      CHAR            3       Country key                         13      T005
'BRANC      CHAR            4       Industry key
'TAETE      NUMC            8       Job at former employer(s)
'ANSVX      CHAR            2       Work Contract - Other Emplo
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Nombre_del_empleador As String
Dim City As String
Dim Country_Key As String
Dim Industry_Key As String
Dim Job_at_former_employer As String
Dim Work_Contract As String

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0023"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0023 = False
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Nombre del empleador
    pos1 = pos1 + pos2
    pos2 = 20
    Nombre_del_empleador = Mid(strlinea, pos1, pos2)
        
    'City
    pos1 = pos1 + pos2
    pos2 = 25
    City = Mid(strlinea, pos1, pos2)
        
    'Country Key
    pos1 = pos1 + pos2
    pos2 = 3
    Country_Key = Mid(strlinea, pos1, pos2)
        
    'Industry Key
    pos1 = pos1 + pos2
    pos2 = 4
    Industry_Key = Mid(strlinea, pos1, pos2)
        
    'Job at former employer(s)
    pos1 = pos1 + pos2
    pos2 = 8
    Job_at_former_employer = Mid(strlinea, pos1, pos2)
        
    'Work Contract - Other Emplo
    pos1 = pos1 + pos2
    pos2 = 2
    Work_Contract = Mid(strlinea, pos1, pos2)

'----------------------------------------------------------------------------
'Por ahora no se hace


Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline


End Sub


Public Sub Leer_Infotipo_IT0027(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0027. Distribucion de Costos.
' Autor      : FGZ
' Fecha      : 22/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION                Código Tabla    Nombre Técnico
'PERNR      NUMC            8       Personnel Number
'INFTY      CHAR            6       Constant name infotype
'SUBTY      CHAR            4       Subtipo
'BEGDA      DATS            8       Inicio de Validez
'ENDDA      DATS            8       Fin de Validez
' .....
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte
Dim i As Integer
Dim Seguir As Boolean

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Costos_a_Distribuir As String
Dim CC(25) As TCCosto
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0027"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0027 = False
    Fila_Infotipo_0027 = Fila_Infotipo_0027 + 1
    Hoja = 10
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0027, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0027, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Costos a Distribuir
    pos1 = 35
    pos2 = 4
    Costos_a_Distribuir = Mid(strlinea, pos1, pos2)
    
    pos1 = 39
    i = 1
    Seguir = True
    Do While i <= 25 And Seguir
        pos2 = 4
        CC(i).Sociedad = Trim(Mid(strlinea, pos1, pos2))
        pos1 = pos1 + pos2
        
        pos2 = 4
        CC(i).Division = Trim(Mid(strlinea, pos1, pos2))
        pos1 = pos1 + pos2
    
        pos2 = 10
        CC(i).CCosto = Trim(Mid(strlinea, pos1, pos2))
        pos1 = pos1 + pos2
    
        pos2 = 7
        CC(i).Porcentaje = Trim(Mid(strlinea, pos1, pos2))
        pos1 = pos1 + pos2
    
    
        If Not EsNulo(CC(i).CCosto) Then
            
        Else
            Seguir = False
        End If
        i = i + 1
    Loop
'----------------------------------------------------------------------------
'Por ahora no se hace

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline


End Sub


Public Sub Leer_Infotipo_IT0041(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0041. Date Specifications.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
'PERNR   NUMC    8   Personnel Number
'INFTY   CHAR    6   Constant name infotype
'SUBTY   CHAR    4   Subtipo
'BEGDA   DATS    8   Inicio de Validez
'ENDDA   DATS    8   Fin de Validez
'DAR01   CHAR    2   Clase de fecha                                                      T548Y
'DAT01   DATS    8   Fecha por clase de fecha
'DAR02   CHAR    2   Clase de fecha                                                      T548Y
'DAT02   DATS    8   Fecha por clase de fecha
'DAR03   CHAR    2   Clase de fecha                                                      T548Y
'DAT03   DATS    8   Fecha por clase de fecha
'DAR04   CHAR    2   Clase de fecha                                                      T548Y
'DAT04   DATS    8   Fecha por clase de fecha
'DAR05   CHAR    2   Clase de fecha                                                      T548Y
'DAT05   DATS    8   Fecha por clase de fecha
'DAR06   CHAR    2   Clase de fecha                                                      T548Y
'DAT06   DATS    8   Fecha por clase de fecha
'DAR07   CHAR    2   Clase de fecha                                                      T548Y
'DAT07   DATS    8   Fecha por clase de fecha
'DAR08   CHAR    2   Clase de fecha                                                      T548Y
'DAT08   DATS    8   Fecha por clase de fecha
'DAR09   CHAR    2   Clase de fecha                                                      T548Y
'DAT09   DATS    8   Fecha por clase de fecha
'DAR10   CHAR    2   Clase de fecha                                                      T548Y
'DAT10   DATS    8   Fecha por clase de fecha
'DAR11   CHAR    2   Clase de fecha                                                      T548Y
'DAT11   DATS    8   Fecha por clase de fecha
'DAR12   CHAR    2   Clase de fecha                                                      T548Y
'DAT12   DATS    8   Fecha por clase de fecha
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Clase_fecha As String
Dim Fecha As String

Dim i As Integer
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0041"
    Columna = 2
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Infotipo_0041 = False
    Fila_Infotipo_0041 = Fila_Infotipo_0041 + 1
    Hoja = 10
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0041, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0041, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0041, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0041, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0041, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    pos1 = 35
    For i = 1 To 12
        'Clase de fecha
        pos2 = 2
        Columna = Columna + 1
        Aux = Mid(strlinea, pos1, pos2)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0041, Columna, Aux)
        Clase_fecha = Trim(Mid(strlinea, pos1, pos2))
        pos1 = pos1 + pos2
        
        'Fecha por clase de fecha
        pos2 = 8
        Columna = Columna + 1
        Aux = Mid(strlinea, pos1, pos2)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0041, Columna, Aux)
        Fecha = StrToFecha(Mid(strlinea, pos1, pos2), OK)
        pos1 = pos1 + pos2
        
        If Not EsNulo(Clase_fecha) Then
            Call Insertar_Fecha(Clase_fecha, CDate(Fecha), Inicio_Validez, Fin_Validez)
        End If
    Next i
    
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub


Public Sub Leer_Infotipo_IT0050(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0050. Time Recording Info.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
'PERNR   NUMC    8   Personnel Number
'INFTY   CHAR    6   Constant name infotype
'SUBTY   CHAR    4   Subtipo
'BEGDA   DATS    8   Inicio de Validez
'ENDDA   DATS    8   Fin de Validez
'SEQNR   NUMC    3   Número de registro de Infotipo con la misma clave
'ZAUSW   NUMC    8   No. de la tajeta del registro de tiempos
'ZAUVE   CHAR    1   Versión del ID
'ZABAR   CHAR    1   Employee grouping for the time evaluation rule
'BDEGR   CHAR    3   Grouping for Connection to Subsystem
'ZANBE   CHAR    2   Access control group
'ZDGBE   CHAR    1   Off-site work athorization
'ZMAIL   CHAR    1   Mail indicator
'ZPINC   CHAR    4   Personal code
'GLMAX   DEC 5.2 Flextime maximum
'GLMIN   DEC 5.2 Flextime minimum
'ZTZUA   DEC 5.2 Time Bonus/Deduction
'ZMGEN   CHAR    1   Standard overtime
'ZUSKZ   CHAR    1   Additional indicator
'PMBDE   NUMC    2   Work time event type group
'GRAWG   CHAR    3   Grouping of Attendance/Absence Reasons
'GRELG   CHAR    3   Grouping for Employee Expenses
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Numero_de_registro_Infotipo As String
Dim Nro_tarjeta_registro_de_tiempos As String
Dim Version_ID As String
Dim Employee_grouping As String
Dim Grouping_for_Connection As String
Dim Access_control_group As String
Dim Off_site As String
Dim Mail_indicator As String
Dim Personal_code As String
Dim Flextime_Maximum As String
Dim Flextime_Minimum As String
Dim Time_Bonus_Deduction As String
Dim Standard_overtime As String
Dim Additional_indicator As String
Dim Work_time As String
Dim Grouping_of_Attendance_Absence_Reasons As String
Dim Grouping_for_Employee_Expenses As String


'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    
    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0050"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0050 = False
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
'-------------------------------------------------------------------------------

'por ahora no se hace porque no me sirve nada de lo que informa para los efectos de la liquidacion

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub


Public Sub Leer_Infotipo_IT0057(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0057. Membership Fees
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
'PERNR   NUMC    8   Personnel Number
'INFTY   CHAR    6   Constant name infotype
'SUBTY   CHAR    4   Subtipo     T591A
'BEGDA   DATS    8   Fecha Inicio
'ENDDA   DATS    8   Fecha Fin
'EMFSL   CHAR    8   Clave de receptor para trasferencia
'MTGLN   CHAR    20  Numero asociado
'GRPRG   CHAR    10  Agrupacion
'BETRG   CURR    9.2 Contribucion de Beneficiario
'WAERS   CUKY    5   Clave de Moneda                     19  TCURC
'LGART   CHAR    4   CC-Nomina       T512Z
'ANZHL   DEC 7.2 Cantidad
'ZEINH   CHAR    3   Unidad de medida/tiempo     T538A
'ZFPER   NUMC    2   Primer periodo de pago
'ZDATE   DATS    8   Primera fecha de pago
'ZANZL   DEC 3   Cantidad para determinar los otros momentos de pago
'ZEINZ   CHAR    3   Unidad de medida/tiempo                                 T538A
'PRITY   CHAR    1   Prioridad
'UFUNC   CHAR    2   Funcion beneficiario
'UNLOC   CHAR    4   Subdivision de la asociacion
'USTAT   CHAR    2   Status de la asociacion
'ESRNR   CHAR    11  No. de usuario ESR
'ESRRE   CHAR    27  No. de referencia ESR
'ESRPZ   CHAR    2   Digito de control ESR
'ZWECK   CHAR    40  Destino de utilizacion para trasferencias
'OPKEN   CHAR    1   Indicador de operacion para CC-Nomina
'INDBW   CHAR    1   Indicador para valoracion indirecta
'ZSCHL   CHAR    1   Via de pago     T042Z
'UWDAT   DATS    8   Fecha de trasferencia
'MODEL   CHAR    4   Modelo de pago
'MGART   CHAR    4   Member Type
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Agrupacion As String
Dim Contribucion_de_Beneficiario As String
Dim Clave_de_Moneda As String
Dim Cantidad As String
Dim Unidad_medida_tiempo As String
Dim CC_Nominas As TNomina

Dim nro_convenio As Long
Dim Inserto_estr As Boolean
Dim Nro_Tercero As Long
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0057"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0057 = False
    Fila_Infotipo_0057 = Fila_Infotipo_0057 + 1
    Hoja = 11
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0057, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0057, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0057, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0057, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0057, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Agrupacion (Convenio)
    pos1 = 63
    pos2 = 10
    Agrupacion = Trim(Mid(strlinea, pos1, pos2))
    
    If Trim(Agrupacion) = "" Then
       Agrupacion = "Fuera de Convenio"
    End If
    
'    Call ValidaEstructura(19, Agrupacion, nro_convenio, Inserto_estr)
'    If Inserto_estr Then
'        'Descripcion = Format_Str(Agrupacion, 40, False, "")
'        Call CreaTercero(19, Agrupacion, Nro_Tercero)
'        Call CreaComplemento(19, Nro_Tercero, nro_convenio, Agrupacion)
'    End If
'    'Descripcion = Format_Str(Agrupacion, 40, False, "")
'    'Call Mapear("T547V", Codigo, CStr(nro_convenio))
'
'    'Convenio
'    Texto = "Agrupacion - [ RHPRO(Convenio)] " & Agrupacion
'    'Nro_Convenio = CLng(CalcularMapeoInv(Agrupacion, "T547V", "0"))
'    If nro_convenio <> 0 Then
'        Call Insertar_His_Estructura(19, nro_convenio, Empleado.Tercero, Inicio_Validez, Fin_Validez)
'    Else
'        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
'        Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
'    End If


'    'Contribucion de Beneficiario (Importe CC-Nomina)
'    pos1 = 73
'    pos2 = 10   '9.2
'    Contribucion_de_Beneficiario = Trim(Mid(strLinea, pos1, 10))
'    If IsNumeric(Contribucion_de_Beneficiario) Then
'        CC_Nominas.Monto = CSng(Mid(Contribucion_de_Beneficiario, 1, 8) & "." & Mid(Contribucion_de_Beneficiario, 9, 2))
'    Else
'        CC_Nominas.Monto = 0
'    End If
'    pos1 = pos1 + 10
'
'
'    'Clave de Moneda
'    pos1 = pos1 + pos2
'    pos2 = 5
'    Clave_de_Moneda = Mid(strLinea, pos1, pos2)
'
'    'CC -Nomina
'    pos1 = 86
'    Aux = Mid(strLinea, pos1, 4)
'    CC_Nominas.Nomina = Trim(Aux)
'    pos1 = pos1 + 4
'
'    'Cantidad
'    pos1 = pos1 + pos2
'    pos2 = 10   '7.2
'    Cantidad = Trim(Mid(strLinea, pos1, pos2))
'    CC_Nominas.Cantidad = CSng(Mid(Cantidad, 2, 8) & "." & Mid(Cantidad, 9, 2))
'
'    'Unidad de medida/tiempo
'    pos1 = pos1 + pos2
'    pos2 = 3
'    Unidad_medida_tiempo = Mid(strLinea, pos1, pos2)
'    CC_Nominas.Unidad = CLng(CalcularMapeoInv(Unidad_medida_tiempo, "T538A", "0"))
'    'Actualizo la cantidad de acuerdo a la unidad de medida
'    CC_Nominas.Cantidad = Calcular_Cantidad(CC_Nominas.Cantidad, CC_Nominas.Unidad)
'

    'Completo las columnas vacias o que no tienen importancia
    For Columna = 6 To 31
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0057, Columna, "")
    Next Columna

''---------------------------------------------------------------------------------
'    'Inserto la novedad
'    If Not EsNulo(CC_Nominas.Nomina) And (CC_Nominas.Monto <> 0 Or CC_Nominas.Cantidad <> 0) Then
'        Call Insertar_Novedad(CC_Nominas.Nomina, CC_Nominas.Monto, CC_Nominas.Cantidad, Inicio_Validez, Fin_Validez, "IT0057")
'    End If


Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub



Public Sub Leer_Infotipo_IT0105(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez

Dim Comunication_Type As String
Dim Comunication_Id As String
Dim Comunication_SMTP As String
Dim Nro_Comunication_Type As Long

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0105"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0105 = False
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Tipo de Comunicacion
    pos1 = 35
    pos2 = 4
    Comunication_Type = Trim(Mid(strlinea, pos1, pos2))
    Nro_Comunication_Type = CLng(CalcularMapeoSubtipo("IT0105", Aux, "T591A", "0"))
    
    'ID de Comunicacion
    pos1 = pos1 + pos2
    pos2 = 30
    Comunication_Id = Trim(Mid(strlinea, pos1, pos2))
    
    'SMTP
    pos1 = pos1 + pos2
    pos2 = 241
    Comunication_SMTP = Mid(strlinea, pos1, pos2)
    
    '---------------------------------------------------------------
    If Not EsNulo(Comunication_Id) Then
        Select Case UCase(Comunication_Type)
        Case "0001":    'user name
        Case "0002":    'SAP2
        Case "0003":    'Net pass
        Case "0004":    'TSO1
        Case "0005":    'Fax
            Call Insertar_Telefono(Comunication_Id, False, False, True)
        Case "0006":    'Voice mail
        Case "0010":    'E-mail
            Call Insertar_Mail(Comunication_Id)
        Case "0011":    'Credit card number
        Case "0020":    'Telefono de trabajo
            Call Insertar_Telefono(Comunication_Id, False, False, False)
        Case "9NCK":    'Nick name
        Case "9RUD":    'Persistent ID
        Case "CELL":    'Celular
            Call Insertar_Telefono(Comunication_Id, True, False, False)
        Case "MAIL":    'E-mail
            Call Insertar_Mail(Comunication_Id)
        Case "MPHN":    'Telefono Mobil
            Call Insertar_Telefono(Comunication_Id, False, False, False)
        Case "PAGR":    'Pager
        Case Else
            'subtipo desconocido
            Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
            FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": subtipo desconocido " & Comunication_Type
        End Select
    End If

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub


Public Sub Leer_Infotipo_IT0185(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0185. Personal ID's.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
'PERNR   NUMC    8   Personnel Number
'INFTY   CHAR    6   Constant name infotype
'SUBTY   CHAR    4   Subtipo                                                             T5R05
'BEGDA   DATS    8   Inicio de Validez
'ENDDA   DATS    8   Fecha Fin
'ICTYP   CHAR    2   Tipo de Identificación (Tipo ID)                                    T5R05
'ICNUM   CHAR    30  Nº Identific.
'ICOLD   CHAR    20  Antigüo Nº ID
'AUTH1   CHAR    30  Autoridad competente
'DOCN1   CHAR    20  Número emisión documento
'FPDAT   DATS    8   Fecha de emisión para ID personal
'EXPID   DATS    8   Final validez de ID de personal
'ISSPL   CHAR    30  Lugar emisión de la identificación
'ISCOT   CHAR    3   País de emisión                                                 13  T005
'IDCOT   CHAR    3   País de ID                                                      15  T005
'OVCHK   CHAR    1   Indicador para sustituir verificación consistencia
'ASTAT   CHAR    1   Application status
'AKIND   CHAR    1   Single/multiple
'REJEC   CHAR    20  Reject reason
'USEFR   DATS    8   Used from -date
'USETO   DATS    8   Used to -date
'DATEN   DEC 3   Valid length of multiple visa
'DATEU   CHAR    3   Time unit for determining next payment      T538A
'TIMES   DATS    8   Application date
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Tipo_ID As String
Dim Nro_Identific As String
Dim Antiguo_Nro_ID As String
Dim Autoridad_competente As String
Dim Numero_emision_documento As String
Dim Fecha_de_emision As String
Dim Fecha_Final_Validez As String
Dim Lugar_emision As String
Dim Pais_de_emision As String
Dim Pais_de_ID As String

Dim Nro_Tipo_ID As Long
Dim Nro_Pais_Emision As Long
Dim Nro_Pais_ID As Long
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0185"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0185 = False
    Fila_Infotipo_0185 = Fila_Infotipo_0185 + 1
    Hoja = 12
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, Columna, Aux)
    Subtipo = Trim(Mid(strlinea, pos1, pos2))

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Tipo de Identificación (Tipo ID)
    Columna = 6
    pos1 = 35
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, Columna, Aux)
    Tipo_ID = Trim(Mid(strlinea, pos1, pos2))
    Nro_Tipo_ID = CLng(CalcularMapeoInv(Tipo_ID, "T5R05", "0"))
    
    'Nº Identific.
    Columna = 7
    pos1 = pos1 + pos2
    pos2 = 30
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, Columna, Aux)
    Nro_Identific = Trim(Mid(strlinea, pos1, pos2))
        
    'País de emisión
    Columna = 14
    pos1 = 183
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, Columna, Aux)
    Pais_de_emision = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Pais_de_emision) Then
        Nro_Pais_Emision = CLng(CalcularMapeoInv(Pais_de_emision, "T005", "0"))
    Else
        Nro_Pais_Emision = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. País de emisión"
    End If
    
    'País de ID
    Columna = 15
    pos1 = pos1 + pos2
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, Columna, Aux)
    Pais_de_ID = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Pais_de_ID) Then
        Nro_Pais_ID = CLng(CalcularMapeoInv(Pais_de_ID, "T005", "0"))
    Else
        Nro_Pais_ID = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. País de ID"
    End If
        
    'Completo las columnas vacias o que no tienen importancia
    For Columna = 16 To 24
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0185, Columna, "")
    Next Columna
        
'--------------------------------------------------------------------------
    If Not EsNulo(Nro_Identific) Then
        Call Insertar_Documento(Nro_Identific, Nro_Tipo_ID)
    End If
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub


Public Sub Leer_Infotipo_IT0267(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0267"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0267 = False
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    '
    pos1 = pos1 + pos2
    pos2 = 6
    Aux = Mid(strlinea, pos1, pos2)
    '---------------------------------------------------------------
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub

Public Sub Leer_Infotipo_IT0389(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Indicador_Agente_Retencion As String
Dim Agente_Retencion As Boolean
Dim IRS As String
Dim Nro_IRS As Long
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0389"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0389 = False
    Fila_Infotipo_0389 = Fila_Infotipo_0389 + 1
    Hoja = 13
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0389, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0389, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0389, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0389, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0389, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If


    'Indicador empleado es Agente Retencion o no
    Columna = 6
    pos1 = 35
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0389, Columna, Aux)
    Indicador_Agente_Retencion = Trim(Mid(strlinea, pos1, pos2))
    Agente_Retencion = IIf(Indicador_Agente_Retencion = "X", True, False)
    
    For Columna = 7 To 9
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0389, Columna, "")
    Next Columna
    
    'IRS (DGI) agency
    Columna = 10
    pos1 = 39
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0389, Columna, Aux)
    IRS = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(IRS) Then
        Nro_IRS = CLng(CalcularMapeoInv(IRS, "T7AR66", "0"))
    Else
        Nro_IRS = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. IRS (DGI) agency"
    End If
        
    'Completo las columnas vacias o que no tienen importancia
    For Columna = 11 To 13
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0389, Columna, "")
    Next Columna
'---------------------------------------------------------------
'IRS (DGI) Agency  - [ RHPro(Agencia)]
Texto = "IRS (DGI) Agency - [ RHPro(Agencia)] " & IRS
If Nro_IRS <> 0 Then
    Call Insertar_His_Estructura(28, Nro_IRS, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If
    
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    
    
End Sub


Public Sub Leer_Infotipo_IT0390(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Aux_Inicio_Validez
Dim Mes_Inicial As String
Dim CC_Nominas As TNomina

Dim NroDoc As String
Dim Persona_Juridica As String
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0390"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0390 = False
    Fila_Infotipo_0390 = Fila_Infotipo_0390 + 1
    Hoja = 14
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'CC -Nomina
    Columna = 6
    pos1 = 35
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
    Aux = Mid(strlinea, pos1, pos2)
    CC_Nominas.Nomina = Trim(Aux)

    'Monto
    Columna = 7
    pos1 = 39
    pos2 = 14
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
    Aux = Mid(strlinea, pos1, pos2)
    If IsNumeric(Aux) Then
        CC_Nominas.Monto = CSng(Mid(Aux, 2, 11) & "." & Mid(Aux, 13, 2))
    Else
        CC_Nominas.Monto = 0
    End If
    If Mid(Aux, 1, 1) = "-" Then
        CC_Nominas.Monto = CC_Nominas.Monto * -1
    End If
    pos1 = pos1 + 14
            
    Columna = 8
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
        
    'Mes inicial de la aplicacion
    Columna = 9
    pos1 = 58
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
    Mes_Inicial = Trim(Mid(strlinea, pos1, pos2))
    If Mes_Inicial = "00" Then
        Mes_Inicial = "01"
    End If
    If Not EsNulo(Mes_Inicial) Then
        Aux_Inicio_Validez = CDate("01/" & Mes_Inicial & "/" & Year(CDate(Inicio_Validez)))
    Else
        Aux_Inicio_Validez = Inicio_Validez
    End If
        
    For Columna = 10 To 13
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
    Next Columna
    
    'Nro de Documento Persona Juridica
    Columna = 14
    pos1 = 76
    pos2 = 20
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
    NroDoc = Trim(Mid(strlinea, pos1, pos2))
    
    'Nombre de la Persona Juridica
    Columna = 15
    pos1 = 96
    pos2 = 30
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, Aux)
    Persona_Juridica = EliminarCHInvalidos(Trim(Mid(strlinea, pos1, pos2)))
    
    'Completo las columnas vacias o que no tienen importancia
    For Columna = 16 To 17
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0390, Columna, "")
    Next Columna
'---------------------------------------------------------------------------------
    'Inserto la DDJJ
    If Not EsNulo(CC_Nominas.Nomina) Then
        If CC_Nominas.Monto <> 0 Then
            'FGZ 12/09/2005 - Por ahorta esto no se levanta
            'Call Insertar_DDJJ(CC_Nominas.Nomina, CC_Nominas.Monto, Aux_Inicio_Validez, Fin_Validez, NroDoc, Persona_Juridica)
        End If
    End If

Exit Sub
Manejador_De_Error:
    HuboError = True
    'Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub

Public Sub Leer_Infotipo_IT0391(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte
Dim i As Integer

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez

Dim CC_Nominas As TNomina
Dim NroDoc As String
Dim Persona_Juridica As String
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0391"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0391 = False
    Fila_Infotipo_0391 = Fila_Infotipo_0391 + 1
    Hoja = 15
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    Columna = 6
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, Columna, Aux)

    'Nro de Documento Persona Juridica
    Columna = 7
    pos1 = 37
    pos2 = 20
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, Columna, Aux)
    NroDoc = Trim(Mid(strlinea, pos1, pos2))
    
    'Nombre de la Persona Juridica
    Columna = 8
    pos1 = 57
    pos2 = 30
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, Columna, Aux)
    Persona_Juridica = Trim(Mid(strlinea, pos1, pos2))
        
    For Columna = 9 To 18
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, Columna, Aux)
    Next Columna
    
    
    pos1 = 278
    Columna = 18
    For i = 1 To 20
        'CC -Nomina
        pos2 = 4
        Columna = Columna + 1
        Aux = Mid(strlinea, pos1, pos2)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, Columna, Aux)
        Aux = Mid(strlinea, pos1, pos2)
        CC_Nominas.Nomina = Trim(Aux)
        pos1 = pos1 + pos2
        
        'Monto
        pos2 = 14
        Columna = Columna + 1
        Aux = Mid(strlinea, pos1, pos2)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0391, Columna, Aux)
        Aux = Mid(strlinea, pos1, pos2)
        If IsNumeric(Aux) Then
            CC_Nominas.Monto = CSng(Mid(Aux, 2, 11) & "." & Mid(Aux, 13, 2))
        Else
            CC_Nominas.Monto = 0
        End If
        If Mid(Aux, 1, 1) = "-" Then
            CC_Nominas.Monto = CC_Nominas.Monto * -1
        End If
        pos1 = pos1 + pos2
            
        'Inserto la DDJJ
        If Not EsNulo(CC_Nominas.Nomina) Then
            'FGZ 12/09/2005 - Por ahorta esto no se levanta
            'Call Insertar_DDJJ(CC_Nominas.Nomina, CC_Nominas.Monto, Inicio_Validez, Fin_Validez, NroDoc, Persona_Juridica)
        End If
    Next i
        
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
        
End Sub


Public Sub Leer_Infotipo_IT0392(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo IT0392. Seguridad Social - Argentina.
' Autor      : FGZ
' Fecha      : 10/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
'PERNR   NUMC    8   Personnel Number
'INFTY   CHAR    6   Constant name infotype
'SUBTY   CHAR    4   Subtipo
'BEGDA   DATS    8   Inicio de Validez
'ENDDA   DATS    8   Fin de Validez
'OBRAS   CHAR    6   Obra social 41  T7AR34
'OSNOA   CHAR    20  Número de afiliación a la obra social
'OBRAO   CHAR    6   Obra social original del empleado   41  T7AR34
'SYJUB   CHAR    1   Indicador : el empleado pertenece al esquema de la distribución
'CAFJP   CHAR    4   Administradoras de fondos de jubilaciones y pensiones   42  T7AR36
'TYACT   CHAR    2   Código de actividad del empleado    43  T7AR38
'PLANS   CHAR    10  Plan de obra social
'CSERV   NUMC    2   Caracter de servicio
'AFJUB   CHAR    15  Número de afiliación retiro por jubilación
'ASPCE   CHAR    4   Agrupación de empleados para la seguridad social
'TPUOS   CHAR    1   Cálculo de los aportes y contribuciones
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Obra_Social As String
Dim Nro_de_Afiliacion As String
Dim Obra_Social_Original As String
Dim Indicador As String
Dim Administradoras_FJP As String
Dim Codigo_de_actividad As String
Dim Plan_Obra_Social As String
Dim Caracter_de_Servicio As String
Dim Nro_Afiliacion_Retiro As String
Dim Agrupacion_de_Empleados As String
Dim Calculo_Aportes_y_Contribuciones As String

Dim Nro_Obra_Social As Long
Dim Nro_Obra_Social_Original As Long
Dim Nro_Administradoras_FJP As Long
Dim Nro_Codigo_de_actividad As Long
Dim Nro_Plan_Obra_Social As Long
Dim Inserto_estr As Boolean
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0392"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0392 = False
    Fila_Infotipo_0392 = Fila_Infotipo_0392 + 1
    Hoja = 16
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'obra social
    Columna = Columna + 1
    pos1 = 35
    pos2 = 6
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    Obra_Social = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Obra_Social) Then
        Nro_Obra_Social = CLng(CalcularMapeoInv(Obra_Social, "T7AR34", "0"))
    Else
        Nro_Obra_Social = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. obra social"
    End If
    
    Columna = Columna + 1
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    
    'Obra social original del empleado
    Columna = Columna + 1
    pos1 = 61
    pos2 = 6
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    Obra_Social_Original = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Obra_Social_Original) Then
        Nro_Obra_Social_Original = CLng(CalcularMapeoInv(Obra_Social_Original, "T7AR34", "0"))
    Else
        Nro_Obra_Social_Original = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. obra social Original"
    End If
        
    'Si alguno de las dos OS es 0 ==> toma el valor de la otra
    If (Nro_Obra_Social <> 0 Or Nro_Obra_Social_Original <> 0) And Nro_Obra_Social <> Nro_Obra_Social_Original And (Nro_Obra_Social = 0 Or Nro_Obra_Social_Original = 0) Then
        If Nro_Obra_Social = 0 Then
            Nro_Obra_Social = Nro_Obra_Social_Original
        Else
            Nro_Obra_Social_Original = Nro_Obra_Social
        End If
    End If
        
    Columna = Columna + 1
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
        
    'Administradoras de fondos de jubilaciones y pensiones
    Columna = Columna + 1
    pos1 = 68
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    Administradoras_FJP = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Administradoras_FJP) Then
        Nro_Administradoras_FJP = CLng(CalcularMapeoInv(Administradoras_FJP, "T7AR36", "0"))
    Else
        Nro_Administradoras_FJP = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Administradoras de fondos de jubilaciones y pensiones"
    End If
    
    'Código de actividad del empleado
    Columna = Columna + 1
    pos1 = 72
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    Codigo_de_actividad = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Codigo_de_actividad) Then
        Nro_Codigo_de_actividad = CLng(CalcularMapeoInv(Codigo_de_actividad, "T7AR38", "0"))
    Else
        Nro_Codigo_de_actividad = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Código de actividad del empleado"
    End If
    
    'Plan de obra social
    Columna = Columna + 1
    pos1 = 74
    pos2 = 10
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    Plan_Obra_Social = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Plan_Obra_Social) Then
        Call ValidaEstructura(23, Plan_Obra_Social, Nro_Plan_Obra_Social, Inserto_estr)
    Else
        Nro_Plan_Obra_Social = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Código de actividad del empleado"
    End If

    For Columna = 13 To 16
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0392, Columna, Aux)
    Next Columna

'---------------------------------------------------------------
    
'Obra social - [ RHPro(Obra social)]
Texto = "Obra Social - [ RHPro(Obra social)] " & Obra_Social
If Nro_Obra_Social <> 0 Then
    Call Insertar_His_Estructura(17, Nro_Obra_Social, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If
    
'Plan de obra social - [ RHPRO(Plan de obra social)]
Texto = "Plan de obra social - [ RHPRO(Plan de obra social)] " & Plan_Obra_Social
If Nro_Plan_Obra_Social <> 0 Then
    Call Insertar_His_Estructura(23, Nro_Plan_Obra_Social, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo del " & Texto
End If
    
'Obra social Original- [ RHPro(Obra social por ley)]
Texto = "Obra Social - [ RHPro(Obra social Original)] " & Obra_Social_Original
If Nro_Obra_Social_Original <> 0 Then
    Call Insertar_His_Estructura(24, Nro_Obra_Social_Original, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If
    
'Plan de obra social - [ RHPRO(Plan de obra social)]
Texto = "Plan de obra social - [ RHPRO(Plan de obra social)] " & Plan_Obra_Social
If Nro_Plan_Obra_Social <> 0 Then
    Call Insertar_His_Estructura(25, Nro_Plan_Obra_Social, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo del " & Texto
End If
    
    
    
'AFJP- [ RHPro(Caja de Jubilacion)]
Texto = "AFJP - [ RHPro(Caja de Jubilacion)] " & Administradoras_FJP
If Nro_Administradoras_FJP <> 0 Then
    Call Insertar_His_Estructura(15, Nro_Administradoras_FJP, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If
    
'Codigo de actividad- [ RHPro(Actividad)]
Texto = "Codigo_de_actividad - [ RHPro(Actividad)] " & Codigo_de_actividad
If Nro_Codigo_de_actividad <> 0 Then
    Call Insertar_His_Estructura(29, Nro_Codigo_de_actividad, Empleado.Tercero, Inicio_Validez, Fin_Validez)
Else
    Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
    Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
End If
    
Exit Sub
Manejador_De_Error:
    HuboError = True
    'Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    
End Sub

Public Sub Leer_Infotipo_IT0393(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim FAMSA As String
Dim CERAP As String
Dim CERAA As String
Dim MESLI As String
Dim ANOES As String
Dim OBJPS As String

Dim Parentesco As Long
Dim rs_Familiar As New ADODB.Recordset
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0393"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0393 = False
    Fila_Infotipo_0393 = Fila_Infotipo_0393 + 1
    Hoja = 17
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Tipo de Hijo - (FAMSA)
    Columna = Columna + 1
    pos1 = 35
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, Columna, Aux)
    FAMSA = Trim(Mid(strlinea, pos1, pos2))
    Parentesco = CLng(CalcularMapeoSubtipo("IT0393", FAMSA, "T591A", "0"))
    
    'Año anterior - (CERAP)
    Columna = Columna + 1
    pos1 = 39
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, Columna, Aux)
    CERAP = Trim(Mid(strlinea, pos1, pos2))
        
    'Año actual - (CERAA)
    Columna = Columna + 1
    pos1 = 40
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, Columna, Aux)
    CERAA = Trim(Mid(strlinea, pos1, pos2))
        
    'Mes de Pago - (MESLI)
    Columna = Columna + 1
    pos1 = 41
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, Columna, Aux)
    MESLI = Trim(Mid(strlinea, pos1, pos2))
        
    'Identificador del hijo - (OBJPS)
    Columna = Columna + 1
    pos1 = 43
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0393, Columna, Aux)
    OBJPS = Trim(Mid(strlinea, pos1, pos2))
    'Esto es provisiorio hasta que definan el campo y la posicion en donde va a venir
    If EsNulo(OBJPS) Then
        OBJPS = 1
    End If
    
    '---------------------------------------------------------------
    StrSql = "SELECT * FROM familiar "
    StrSql = StrSql & " WHERE empleado = " & Empleado.Tercero
    StrSql = StrSql & " AND famnrocorr = " & OBJPS
    StrSql = StrSql & " AND parenro = " & Parentesco
    If rs_Familiar.State = adStateOpen Then rs_Familiar.Close
    OpenRecordset StrSql, rs_Familiar
    
    If Not rs_Familiar.EOF Then
        StrSql = "UPDATE familiar SET "
        StrSql = StrSql & "  famestudia = -1 " 'estudia
        StrSql = StrSql & " ,famsalario = -1 " 'paga salario familiar
        StrSql = StrSql & " WHERE empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND ternro = " & rs_Familiar!ternro
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        'No se puede dar, tiene que haber sido informado con anterioridad por el infotipo 21
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "No se encontró el familiar " & OBJPS
    End If
    
Exit Sub
Manejador_De_Error:
    HuboError = True
    'Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    
End Sub


Public Sub Leer_Infotipo_IT0394(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez

Dim ASFAX As String
Dim DISCP As String
Dim NADOC As String
Dim OBJPS As String
Dim ICTYP As String
Dim Nro_ICTYP As Long
Dim ICNUM As String
Dim Parentesco As Long

Dim rs_Familiar As New ADODB.Recordset
Dim rs_TDoc As New ADODB.Recordset
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0394"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0394 = False
    Fila_Infotipo_0394 = Fila_Infotipo_0394 + 1
    Hoja = 18
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    Subtipo = Trim(Mid(strlinea, pos1, pos2))
    Parentesco = CLng(CalcularMapeoSubtipo("IT0021", Subtipo, "T591A", "0"))

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Paga asignacion por hijo - (ASFAX)
    Columna = 6
    pos1 = 35
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    ASFAX = Trim(Mid(strlinea, pos1, pos2))
        
    'hijo discapacitado - (DISCP)
    Columna = 7
    pos1 = 36
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    DISCP = Trim(Mid(strlinea, pos1, pos2))
        
    For Columna = 7 To 10
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    Next Columna
        
    'Certificado - (NADOC)
    Columna = 11
    pos1 = 47
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    NADOC = Trim(Mid(strlinea, pos1, pos2))
        
    For Columna = 12 To 15
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    Next Columna
        
    'Tipo de Documento
    Columna = 16
    pos1 = 65
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    ICTYP = Trim(Mid(strlinea, pos1, pos2))
    Nro_ICTYP = CLng(CalcularMapeoInv(ICTYP, "T5R05", "0"))
    
    'Tipo de Documento
    Columna = 17
    pos1 = 67
    pos2 = 20
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    ICNUM = Trim(Mid(strlinea, pos1, pos2))
    
    'Identificador del hijo - (OBJPS)
    Columna = 18
    pos1 = 87
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_0394, Columna, Aux)
    OBJPS = Trim(Mid(strlinea, pos1, pos2))
    
    If EsNulo(OBJPS) Then
        OBJPS = 0
    End If
    '---------------------------------------------------------------
    
    If OBJPS <> 0 Then
        'busco por nro de familiar
        StrSql = "SELECT * FROM familiar "
        StrSql = StrSql & " WHERE familiar.empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND famnrocorr = " & OBJPS
    Else
        'Busco por tipo de fliar
        StrSql = "SELECT * FROM familiar "
        StrSql = StrSql & " INNER JOIN tercero ON familiar.ternro = tercero.ternro "
        StrSql = StrSql & " WHERE familiar.empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND familiar.parenro = " & Parentesco
    End If
    If rs_Familiar.State = adStateOpen Then rs_Familiar.Close
    OpenRecordset StrSql, rs_Familiar
    
    If Not rs_Familiar.EOF Then
        StrSql = "UPDATE familiar SET "
        StrSql = StrSql & "   faminc = " & CInt(IIf(Not EsNulo(DISCP), True, False)) 'Hijo incapacitado
        StrSql = StrSql & " , famsalario = " & CInt(IIf(Not EsNulo(ASFAX), True, False)) 'paga salario familiar
        StrSql = StrSql & " , famcernac = " & CInt(IIf(Not EsNulo(NADOC), True, False)) 'paga salario familiar
        If Not EsNulo(ASFAX) Then
            StrSql = StrSql & " ,famcargaDGI = -1 " 'DDJJ
            StrSql = StrSql & " ,famDGIdesde = " & ConvFecha(Inicio_Validez) 'DDJJ
            If Not EsNulo(Fin_Validez) Then
                StrSql = StrSql & " ,famDGIhasta = " & ConvFecha(Fin_Validez) 'DDJJ
            Else
                StrSql = StrSql & " ,famDGIhasta = " & ConvFecha(CDate("31/12/" & Year(CDate(Inicio_Validez)))) 'DDJJ
            End If
        Else
            StrSql = StrSql & " ,famcargaDGI = 0 " 'DDJJ
        End If
        StrSql = StrSql & " WHERE empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND ternro = " & rs_Familiar!ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Actualizo los documentos
        If Nro_ICTYP <> 0 Then
            ICNUM = Format_Str(ICNUM, 30, False, "")
            
            StrSql = "SELECT * FROM ter_doc  "
            StrSql = StrSql & " WHERE ter_doc.tidnro = " & Nro_ICTYP '& " AND ter_doc.nrodoc = '" & Doc & "'"
            StrSql = StrSql & " AND ternro = " & rs_Familiar!ternro
            OpenRecordset StrSql, rs_TDoc
                
            If rs_TDoc.EOF Then
                StrSql = " INSERT INTO ter_doc(ternro,tidnro,nrodoc) "
                StrSql = StrSql & " VALUES(" & rs_Familiar!ternro & "," & Nro_ICTYP & ",'" & ICNUM & "')"
            Else 'Actualizo
                StrSql = " UPDATE ter_doc SET "
                StrSql = StrSql & " nrodoc = '" & ICNUM & "'"
                StrSql = StrSql & " WHERE ternro = " & rs_Familiar!ternro
                StrSql = StrSql & " AND tidnro = " & Nro_ICTYP
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Else
        'No se puede dar, tiene que haber sido informado con anterioridad por el infotipo 21
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "No se encontró el familiar " & OBJPS
    End If
   
   
'cierro y libero
If rs_Familiar.State = adStateOpen Then rs_Familiar.Close
Set rs_Familiar = Nothing


Exit Sub
Manejador_De_Error:
    HuboError = True
    'Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub

Public Sub Leer_Infotipo_IT0416(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 0416"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_0416 = False
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    '
    pos1 = pos1 + pos2
    pos2 = 6
    Aux = Mid(strlinea, pos1, pos2)
        
    
    '---------------------------------------------------------------
    
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    
End Sub


Public Sub Leer_Infotipo_IT2001(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez

Dim TipoLic As Long
Dim Hora_Inicio As String
Dim Hora_Fin As String
Dim Clase_Ausentismo As String
Dim Dias_Ausentismo As String
Dim Horas_Ausentismo  As String
Dim Dias_Nomina As Single
Dim Horas_Nomina As Single
Dim Dias_Calculables As String
Dim Inicio_Certificado_Enfermedad As String
Dim Fecha_Notificacion_Enfermedad As String
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 2001"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_2001 = False
    Fila_Infotipo_2001 = Fila_Infotipo_2001 + 1
    Hoja = 19
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Subtipo = Trim(Mid(strlinea, pos1, pos2))
    TipoLic = CLng(CalcularMapeoSubtipo("IT2001", Subtipo, "TLIC", "0"))
    
    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    'Hora Inicio
    Columna = 6
    pos1 = 35
    pos2 = 6
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Hora_Inicio = Trim(Mid(strlinea, pos1, pos2))
        
    'Hora Final
    Columna = 7
    pos1 = 41
    pos2 = 6
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Hora_Fin = Trim(Mid(strlinea, pos1, pos2))
        
    Columna = 8
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
        
    'Clase Ausentismo
    Columna = 9
    pos1 = 48
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Clase_Ausentismo = Trim(Mid(strlinea, pos1, pos2))
        
    'Dias Ausentismo
    Columna = 10
    pos1 = 52
    pos2 = 7
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    Dias_Ausentismo = CSng(Mid(Aux, 2, 4) & "." & Mid(Aux, 6, 2))
    
    'Horas Ausentismo
    Columna = 11
    pos1 = 59
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    Horas_Ausentismo = CSng(Mid(Aux, 2, 5) & "." & Mid(Aux, 7, 2))
        
    'Dias Nomina
    Columna = 12
    pos1 = 67
    pos2 = 7
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    Dias_Nomina = CSng(Mid(Aux, 2, 4) & "." & Mid(Aux, 6, 2))
    
    'Horas Nomina
    Columna = 13
    pos1 = 74
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    Horas_Nomina = CSng(Mid(Aux, 2, 5) & "." & Mid(Aux, 7, 2))
    
    'Dias Calculables para salario
    Columna = 14
    pos1 = 82
    pos2 = 7
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    Dias_Calculables = CSng(Mid(Aux, 2, 4) & "." & Mid(Aux, 6, 2))
     
    Columna = 15
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Columna = 16
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    
    'Inicio certificado enfermedad
    pos1 = 105
    pos2 = 8
    Columna = 17
    Texto = "Inicio certificado enfermedad"
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Inicio_Certificado_Enfermedad = StrToFecha(Trim(Mid(strlinea, pos1, pos2)), OK)
    If Not OK Then
        'Flog.Writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        'Flog.Writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strLinea, pos1, pos2)
        'InsertaError Columna, 8
        'HuboError = True
        'Exit Sub
    End If
        
    'Fecha Notificacion enfermedad
    pos1 = 113
    pos2 = 8
    Columna = 18
    Texto = "Fecha Notificacion enfermedad"
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Fecha_Notificacion_Enfermedad = StrToFecha(Trim(Mid(strlinea, pos1, pos2)), OK)
    If Not OK Then
        'Flog.Writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        'Flog.Writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strLinea, pos1, pos2)
        'InsertaError Columna, 8
        'HuboError = True
        'Exit Sub
    End If
        
    For Columna = 19 To 66
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2001, Columna, Aux)
    Next Columna
        
    If TipoLic <> 0 Then
        Call Insertar_Licencia(TipoLic, Inicio_Validez, Fin_Validez, 0, Dias_Nomina, Empleado.Tercero, True, OK)
        If Not OK Then
            'Flog.Writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
            'Flog.Writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strLinea, pos1, pos2)
            'HuboError = True
            'Exit Sub
        End If
    End If
    
Exit Sub
Manejador_De_Error:
    HuboError = True
    'Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    
End Sub

Public Sub Leer_Infotipo_IT2006(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 2006"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_2006 = False
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    '
    pos1 = pos1 + pos2
    pos2 = 6
    Aux = Mid(strlinea, pos1, pos2)
        
    
    '---------------------------------------------------------------
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub

Public Sub Leer_Infotipo_IT2010(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez

Dim CC_Nominas As TNomina
Dim horas_comprobantes As String
Dim Hoja As Integer

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 2010"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_2010 = False
    Fila_Infotipo_2010 = Fila_Infotipo_2010 + 1
    Hoja = 20
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    For Columna = 6 To 8
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Next Columna

    'Horas para comprobantes
    Columna = 9
    pos1 = 48
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    If IsNumeric(Aux) Then
        horas_comprobantes = CSng(Mid(Aux, 2, 5) & "." & Mid(Aux, 7, 2))
    Else
        horas_comprobantes = 0
    End If
    pos1 = pos1 + 8
    
    'CC-Nomina
    Columna = 10
    pos1 = 56
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Aux = Mid(strlinea, pos1, pos2)
    CC_Nominas.Nomina = Trim(Aux)

    'Cantidad por unidades de tiempo
    Columna = 11
    pos1 = 60
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    If IsNumeric(Aux) Then
        CC_Nominas.Cantidad = CSng(Mid(Aux, 2, 5) & "." & Mid(Aux, 7, 2))
    Else
        CC_Nominas.Cantidad = 0
    End If
    pos1 = pos1 + 8

    'Unidad de Medida/Tiempo
    Columna = 12
    pos1 = 68
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    If Not EsNulo(Aux) Then
        CC_Nominas.Unidad = CLng(CalcularMapeoInv(Aux, "T538A", "0"))
    Else
        CC_Nominas.Unidad = 2
    End If
    'Actualizo la cantidad de acuerdo a la unidad de medida
    CC_Nominas.Cantidad = Calcular_Cantidad(CC_Nominas.Cantidad, CC_Nominas.Unidad)
    
    For Columna = 13 To 14
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Next Columna
    
    'Monto
    Columna = 15
    pos1 = 86
    pos2 = 14
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, pos2))
    If IsNumeric(Aux) Then
        CC_Nominas.Monto = CSng(Mid(Aux, 2, 11) & "." & Mid(Aux, 13, 2))
    Else
        CC_Nominas.Monto = 0
    End If
    If Mid(Aux, 1, 1) = "-" Then
        CC_Nominas.Monto = CC_Nominas.Monto * -1
    End If
    pos1 = pos1 + 14
        
    For Columna = 16 To 37
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_2010, Columna, Aux)
    Next Columna
        
        
    'Inserto la novedad
    If Not EsNulo(CC_Nominas.Nomina) And (CC_Nominas.Monto <> 0 Or CC_Nominas.Cantidad <> 0) Then
        Call Insertar_Novedad(CC_Nominas.Nomina, CC_Nominas.Monto, CC_Nominas.Cantidad, Inicio_Validez, Fin_Validez, "IT2010")
    End If
Exit Sub
Manejador_De_Error:
    HuboError = True
    'Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
        
End Sub

Public Sub Leer_Infotipo_IT2013(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 2013"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_2013 = False
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If

    '
    pos1 = pos1 + pos2
    pos2 = 6
    Aux = Mid(strlinea, pos1, pos2)
        
    
    '---------------------------------------------------------------
Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline


End Sub

Public Sub Leer_Infotipo_IT9004(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
'PERNR   NUMC    8   Personnel Number
'INFTY   CHAR    6   Constant name infotype
'SUBTY   CHAR    4   Subtipo
'BEGDA   DATS    8   Inicio de Validez
'ENDDA   DATS    8   Fin de Validez
'TYPE    CHAR    4   This field indicates the type of plan of health offered for
'BEGDA_T DATS    8   Start Date
'ENDDA_T DATS    8   End Date
'APPLICY CURR    13.2    This field indicates the value of the Insurance policy
'BEGDA_AP    DATS    8   Start Date
'ENDDA_AP    DATS    8   End Date
'AMOUNT  CURR    13.2    This field indicates the amount Roche Connect
'BEGDA_AM    DATS    8   Start Date
'ENDDA_AM    DATS    8   End Date
'PERCENT_PLAN    DEC 7.2 Percentage
'STATUS  CHAR    2   Status of Plan
'BEGDA_PLAN  DATS    8   Start Date
'ENDDA_PLAN  DATS    8   End Date
'MONTH_SAVIN NUMC    3   Month of the Savings.
'PERCENT_SAVIN   DEC 7.2 Percentage
'BEGDA_SAVIN DATS    8   Start Date
'ENDDA_SAVIN DATS    8   End Date
'BENE01  CHAR    40  Benefit dependent name
'DATE01  DATS    8   Vesting Date
'PERC01  DEC 7.2 Percentage
'BENE02  CHAR    40  Benefit dependent name
'DATE02  DATS    8   Vesting Date
'PERC02  DEC 7.2 Percentage
'BENE03  CHAR    40  Benefit dependent name
'DATE03  DATS    8   Vesting Date
'PERC03  DEC 7.2 Percentage
'BENE04  CHAR    40  Benefit dependent name
'DATE04  DATS    8   Vesting Date
'PERC04  DEC 7.2 Percentage
'BENE05  CHAR    40  Benefit dependent name
'DATE05  DATS    8   Vesting Date
'PERC05  DEC 7.2 Percentage
'BENE06  CHAR    40  Benefit dependent name
'DATE06  DATS    8   Vesting Date
'PERC06  DEC 7.2 Percentage
'BENE07  CHAR    40  Benefit dependent name
'DATE07  DATS    8   Vesting Date
'PERC07  DEC 7.2 Percentage
'BENE08  CHAR    40  Benefit dependent name
'DATE08  DATS    8   Vesting Date
'PERC08  DEC 7.2 Percentage
'BENE09  CHAR    40  Benefit dependent name
'DATE09  DATS    8   Vesting Date
'PERC09  DEC 7.2 Percentage
'BENE10  CHAR    40  Benefit dependent name
'DATE10  DATS    8   Vesting Date
'PERC10  DEC 7.2 Percentage
'RECRE   CHAR    1   if employee has option for pay and participated in employee
'TICKET  CHAR    1   Flag for indicate employee that participated in this benefit
'RESTAUR CHAR    6   Restaurante
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Hoja As Integer
Dim Monto As Single
Dim concepto As String
Dim concnro As Long
Dim tpanro As Long

Dim rs As New ADODB.Recordset

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 9004"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_9004 = False
    Fila_Infotipo_9004 = Fila_Infotipo_9004 + 1
    Hoja = 21
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9004, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9004, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9004, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9004, Columna, Aux)
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9004, Columna, Aux)
    
    For Columna = 6 To 11
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9004, Columna, Aux)
    Next Columna
    
    'Monto Roche Connect
    Columna = 12
    pos1 = 85
    pos2 = 14
    Aux = Mid(strlinea, pos1, 14)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9004, Columna, Aux)
    Aux = Trim(Mid(strlinea, pos1, 14))
    If IsNumeric(Aux) Then
        Monto = CSng(Mid(Aux, 2, 11) & "." & Mid(Aux, 13, 2))
    Else
        Monto = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Importe CC-Nomina"
    End If
    If Mid(Aux, 1, 1) = "-" Then
        Monto = Monto * -1
    End If
    
    'Inicio de Validez
    Columna = 13
    Texto = "Inicio de Validez"
    pos1 = 99
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9004, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = 14
    Texto = "Fin de Validez"
    pos1 = 107
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9004, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    '---------------------------------------------------------------
    'Esto esta cableado porque tiene un tratamiento especial
    concepto = "0612"
    Parametro = 51
    StrSql = "SELECT * FROM concepto "
    StrSql = StrSql & " WHERE conccod = '" & concepto & "'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        concnro = rs!concnro
    Else
        concnro = 0
    End If
    
    'Inserto la novedad
    If Not EsNulo(concepto) And Monto <> 0 Then
        Call Insertar_Novedad_2(concnro, Parametro, Monto, Inicio_Validez, Fin_Validez, "IT9004")
    End If

Exit Sub
Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline

End Sub

Public Sub Leer_Infotipo_IT9302(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Infotipo
' Autor      : FGZ
' Fecha      : 13/07/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
'PERNR   NUMC    8   Personnel Number
'INFTY   CHAR    6   Constant name infotype
'SUBTY   CHAR    4   Tipo de ausentismo
'BEGDA   DATS    8   Inicio de validez
'ENDDA   DATS    8   Fin de la validez
'NROPR   NUMC    3   Numero de prestamo
'TIPPR   CHAR    2   Tipo de prestamo                    ZCO_TTPRW
'MONTO   CURR    16.2    Monto
'MONEDA  CUKY    5   Moneda  19  TCURC
'NOMTES  CHAR    1   Indicador si paga por nomina o no
'MODLD   CHAR    2   Codigos de modalidad de prestamo
'NCUOTAS NUMC    3   Numero de cuotas
'CODIN   CHAR    10  Codigo de intereses
'PTAJE   DEC 5.2 Porcentaje de interes o puntos
'NPOLIZA CHAR    10  Numero de poliza de seguros
'LAND    CHAR    3   Codigo de pais de Cia de Seguros
'BANKL   CHAR    15  Clave o Nit de Cia de Seguros
'BANKN   CHAR    18  No. Cta Cia de Seguros
'VLRSEG  CURR    16.2    Valor del seguro
'NPOLIZA01   CHAR    10  Numero de poliza de seguros
'LAND01  CHAR    3   Codigo de pais de Cia de Seguros
'BANKL01 CHAR    15  Clave o Nit de Cia de Seguros
'BANKN01 CHAR    18  No. Cta Cia de Seguros
'VLRSEG01    CURR    16.2    Valor del seguro
'GARANTIA    NUMC    2   Codigo de garantia
'DETGAR01    CHAR    100 Detalle de garantia
'VLRGAR  CURR    16.2    Valor de garantia
'VLRBIEN CURR    16.2    Valor del bien dado en garantia
'GR_PTVL CHAR    1   Porcentaje o valor de gradiente
'GR_VALOR    CURR    16.2    Valor de gradiente
'GR_PERDC    NUMC    3   Periodicidad de gradiente
'CPT01   CHAR    4   Concepto de nomina
'PTV01   CHAR    1   Porcentaje o valor
'VAL01   CURR    16.2    Monto
'BAE01   CHAR    1
'FRE01   NUMC    3   Frecuencia de periodicidad
'GRA01   NUMC    3   Gracia para aplicacion del concepto
'NCT01   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT01   CURR    16.2    Monto
'CPT02   CHAR    4   Concepto de nomina
'PTV02   CHAR    1   Porcentaje o valor
'VAL02   CURR    16.2    Monto
'BAE02   CHAR    1
'FRE02   NUMC    3   Frecuencia de periodicidad
'GRA02   NUMC    3   Gracia para aplicacion del concepto
'NCT02   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT02   CURR    16.2    Monto
'CPT03   CHAR    4   Concepto de nomina
'PTV03   CHAR    1   Porcentaje o valor
'VAL03   CURR    16.2    Monto
'BAE03   CHAR    1
'FRE03   NUMC    3   Frecuencia de periodicidad
'GRA03   NUMC    3   Gracia para aplicacion del concepto
'NCT03   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT03   CURR    16.2    Monto
'CPT04   CHAR    4   Concepto de nomina
'PTV04   CHAR    1   Porcentaje o valor
'VAL04   CURR    16.2    Monto
'BAE04   CHAR    1
'FRE04   NUMC    3   Frecuencia de periodicidad
'GRA04   NUMC    3   Gracia para aplicacion del concepto
'NCT04   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT04   CURR    16.2    Monto
'CPT05   CHAR    4   Concepto de nomina
'PTV05   CHAR    1   Porcentaje o valor
'VAL05   CURR    16.2    Monto
'BAE05   CHAR    1
'FRE05   NUMC    3   Frecuencia de periodicidad
'GRA05   NUMC    3   Gracia para aplicacion del concepto
'NCT05   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT05   CURR    16.2    Monto
'CPT06   CHAR    4   Concepto de nomina
'PTV06   CHAR    1   Porcentaje o valor
'VAL06   CURR    16.2    Monto
'BAE06   CHAR    1
'FRE06   NUMC    3   Frecuencia de periodicidad
'GRA06   NUMC    3   Gracia para aplicacion del concepto
'NCT06   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT06   CURR    16.2    Monto
'CPT07   CHAR    4   Concepto de nomina
'PTV07   CHAR    1   Porcentaje o valor
'VAL07   CURR    16.2    Monto
'BAE07   CHAR    1
'FRE07   NUMC    3   Frecuencia de periodicidad
'GRA07   NUMC    3   Gracia para aplicacion del concepto
'NCT07   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT07   CURR    16.2    Monto
'CPT08   CHAR    4   Concepto de nomina
'PTV08   CHAR    1   Porcentaje o valor
'VAL08   CURR    16.2    Monto
'BAE08   CHAR    1
'FRE08   NUMC    3   Frecuencia de periodicidad
'GRA08   NUMC    3   Gracia para aplicacion del concepto
'NCT08   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT08   CURR    16.2    Monto
'CPT09   CHAR    4   Concepto de nomina
'PTV09   CHAR    1   Porcentaje o valor
'VAL09   CURR    16.2    Monto
'BAE09   CHAR    1
'FRE09   NUMC    3   Frecuencia de periodicidad
'GRA09   NUMC    3   Gracia para aplicacion del concepto
'NCT09   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT09   CURR    16.2    Monto
'CPT10   CHAR    4   Concepto de nomina
'PTV10   CHAR    1   Porcentaje o valor
'VAL10   CURR    16.2    Monto
'BAE10   CHAR    1
'FRE10   NUMC    3   Frecuencia de periodicidad
'GRA10   NUMC    3   Gracia para aplicacion del concepto
'NCT10   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT10   CURR    16.2    Monto
'CPT11   CHAR    4   Concepto de nomina
'PTV11   CHAR    1   Porcentaje o valor
'VAL11   CURR    16.2    Monto
'BAE11   CHAR    1
'FRE11   NUMC    3   Frecuencia de periodicidad
'GRA11   NUMC    3   Gracia para aplicacion del concepto
'NCT11   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT11   CURR    16.2    Monto
'CPT12   CHAR    4   Concepto de nomina
'PTV12   CHAR    1   Porcentaje o valor
'VAL12   CURR    16.2    Monto
'BAE12   CHAR    1
'FRE12   NUMC    3   Frecuencia de periodicidad
'GRA12   NUMC    3   Gracia para aplicacion del concepto
'NCT12   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT12   CURR    16.2    Monto
'CPT13   CHAR    4   Concepto de nomina
'PTV13   CHAR    1   Porcentaje o valor
'VAL13   CURR    16.2    Monto
'BAE13   CHAR    1
'FRE13   NUMC    3   Frecuencia de periodicidad
'GRA13   NUMC    3   Gracia para aplicacion del concepto
'NCT13   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT13   CURR    16.2    Monto
'CPT14   CHAR    4   Concepto de nomina
'PTV14   CHAR    1   Porcentaje o valor
'VAL14   CURR    16.2    Monto
'BAE14   CHAR    1
'FRE14   NUMC    3   Frecuencia de periodicidad
'GRA14   NUMC    3   Gracia para aplicacion del concepto
'NCT14   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT14   CURR    16.2    Monto
'CPT15   CHAR    4   Concepto de nomina
'PTV15   CHAR    1   Porcentaje o valor
'VAL15   CURR    16.2    Monto
'BAE15   CHAR    1
'FRE15   NUMC    3   Frecuencia de periodicidad
'GRA15   NUMC    3   Gracia para aplicacion del concepto
'NCT15   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT15   CURR    16.2    Monto
'CPT16   CHAR    4   Concepto de nomina
'PTV16   CHAR    1   Porcentaje o valor
'VAL16   CURR    16.2    Monto
'BAE16   CHAR    1
'FRE16   NUMC    3   Frecuencia de periodicidad
'GRA16   NUMC    3   Gracia para aplicacion del concepto
'NCT16   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT16   CURR    16.2    Monto
'CPT17   CHAR    4   Concepto de nomina
'PTV17   CHAR    1   Porcentaje o valor
'VAL17   CURR    16.2    Monto
'BAE17   CHAR    1
'FRE17   NUMC    3   Frecuencia de periodicidad
'GRA17   NUMC    3   Gracia para aplicacion del concepto
'NCT17   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT17   CURR    16.2    Monto
'CPT18   CHAR    4   Concepto de nomina
'PTV18   CHAR    1   Porcentaje o valor
'VAL18   CURR    16.2    Monto
'BAE18   CHAR    1
'FRE18   NUMC    3   Frecuencia de periodicidad
'GRA18   NUMC    3   Gracia para aplicacion del concepto
'NCT18   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT18   CURR    16.2    Monto
'CPT19   CHAR    4   Concepto de nomina
'PTV19   CHAR    1   Porcentaje o valor
'VAL19   CURR    16.2    Monto
'BAE19   CHAR    1
'FRE19   NUMC    3   Frecuencia de periodicidad
'GRA19   NUMC    3   Gracia para aplicacion del concepto
'NCT19   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT19   CURR    16.2    Monto
'CPT20   CHAR    4   Concepto de nomina
'PTV20   CHAR    1   Porcentaje o valor
'VAL20   CURR    16.2    Monto
'BAE20   CHAR    1
'FRE20   NUMC    3   Frecuencia de periodicidad
'GRA20   NUMC    3   Gracia para aplicacion del concepto
'NCT20   NUMC    3   Numero de cuotas-concepto de aplicacion
'TOT20   CURR    16.2    Monto
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim Aux
Dim OK As Boolean
Dim Columna As Byte

Dim Subtipo As String
Dim Inicio_Validez
Dim Fin_Validez
Dim Hoja As Integer

Dim Concepto1 As String
Dim Concepto2 As String
Dim Cuotas1 As Integer
Dim Cuotas2 As Integer
Dim Monto1 As Double
Dim Monto2 As Double

Dim Nro_Prestamo As String
Dim Tipo_Prestamo As String
Dim Monto_Prestamo As Double
Dim Modalidad_Prestamo As String
Dim Cuotas_Prestamo As Integer
Dim Poliza As String
Dim Nro_Moneda As Long

Dim rs As New ADODB.Recordset

'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador
'    'Empleado
'    pos1 = 1
'    pos2 = 8
'    Legajo = Mid$(strLinea, pos1, pos2)
'
'    'Infotipo
'    pos1 = 9
'    pos2 = 6
'    Infotipo = Mid(strLinea, pos1, pos2)
'Las dos primeras no las evaluo porque ya se evaluaron en el procedimiento llamador

    On Error GoTo Manejador_De_Error
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo 9302"
    If Not EsNulo(Empleado.Tercero) And Empleado.Tercero = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Imposible insertar o Modificar datos. Legajo inexistente."
        Exit Sub
    End If
    Columna = 2
    Infotipo_9302 = False
    Fila_Infotipo_9302 = Fila_Infotipo_9302 + 1
    Hoja = 22
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, 1, Empleado.Legajo)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, 2, Infotipo)
    
    'Subtipo
    Columna = Columna + 1
    Texto = "Subtipo"
    pos1 = 15
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Subtipo = Mid(strlinea, pos1, pos2)

    'Inicio de Validez
    Columna = Columna + 1
    Texto = "Inicio de Validez"
    pos1 = 19
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Inicio_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    'Fin de Validez
    Columna = Columna + 1
    Texto = "Fin de Validez"
    pos1 = 27
    pos2 = 8
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Fin_Validez = StrToFecha(Mid(strlinea, pos1, pos2), OK)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        InsertaError Columna, 8
        HuboError = True
        Exit Sub
    End If
    
    
   'NROPR   NUMC    3   Numero de prestamo
    Columna = Columna + 1
    pos1 = 35
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Nro_Prestamo = Aux
    
    'TIPPR   CHAR    2   Tipo de prestamo                    ZCO_TTPRW
    Columna = Columna + 1
    pos1 = 38
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Tipo_Prestamo = Aux

    'MONTO   CURR    16.2    Monto
    Columna = Columna + 1
    pos1 = 41
    pos2 = 16
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    If IsNumeric(Aux) Then
        Monto_Prestamo = CSng(Mid(Aux, 2, 13) & "." & Mid(Aux, 15, 2))
        'Monto_Prestamo = CDbl(Mid(aux, 2, Len(aux)))
    Else
        Monto_Prestamo = 0
    End If
    If Mid(Aux, 1, 1) = "-" Then
        Monto_Prestamo = Monto_Prestamo * -1
    End If

    'MONEDA  CUKY    5   Moneda  19  TCURC
    Columna = Columna + 1
    pos1 = 57
    pos2 = 5
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Nro_Moneda = CLng(CalcularMapeoInv(Aux, "TCURC", "-1"))

    'Columnas vacias
    For Columna = 10 To 10
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Next Columna
    
    'MODLD   CHAR    2   Codigos de modalidad de prestamo
    Columna = 11
    pos1 = 63
    pos2 = 2
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Modalidad_Prestamo = Aux
    
    'NCUOTAS NUMC    3   Numero de cuotas
    Columna = Columna + 1
    pos1 = 65
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Cuotas_Prestamo = Aux
    
    'Columnas vacias
    For Columna = 13 To 14
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Next Columna
    
    'NPOLIZA CHAR    10  Numero de poliza de seguros
    Columna = 15
    pos1 = 84
    pos2 = 10
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Poliza = Aux
    
    'Columnas vacias
    For Columna = 16 To 31
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Next Columna
    
    
    'CPT01   CHAR    4   Concepto de nomina
    Columna = 32
    pos1 = 367
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Concepto1 = Aux
    
    'PTV01   CHAR    1   Porcentaje o valor
    Columna = Columna + 1
    pos1 = 371
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'VAL01   CURR    16.2    Monto
    Columna = Columna + 1
    pos1 = 373
    pos2 = 16
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    If IsNumeric(Aux) Then
        Monto1 = CSng(Mid(Aux, 2, 13) & "." & Mid(Aux, 15, 2))
        'Monto1 = CDbl(Mid(aux, 2, Len(aux)))
    Else
        Monto1 = 0
    End If
    If Mid(Aux, 1, 1) = "-" Then
        Monto1 = Monto1 * -1
    End If

    'BAE01   CHAR    1
    Columna = Columna + 1
    pos1 = 389
    pos2 = 1
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'FRE01   NUMC    3   Frecuencia de periodicidad
    Columna = Columna + 1
    pos1 = 390
    pos2 = 3
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'GRA01   NUMC    3   Gracia para aplicacion del concepto
    Columna = Columna + 1
    pos1 = 393
    pos2 = 3
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'NCT01   NUMC    3   Numero de cuotas-concepto de aplicacion
    Columna = Columna + 1
    pos1 = 396
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Cuotas1 = Aux
    
    'TOT01   CURR    16.2    Monto
    Columna = Columna + 1
    pos1 = 400
    pos2 = 16
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'CPT02   CHAR    4   Concepto de nomina
    Columna = 40
    pos1 = 416
    pos2 = 4
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Concepto2 = Aux
    
    'PTV02   CHAR    1   Porcentaje o valor
    Columna = Columna + 1
    pos1 = 420
    pos2 = 1
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'VAL02   CURR    16.2    Monto
    Columna = Columna + 1
    pos1 = 422
    pos2 = 16
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    If IsNumeric(Aux) Then
        Monto2 = CSng(Mid(Aux, 2, 13) & "." & Mid(Aux, 15, 2))
    Else
        Monto2 = 0
    End If
    If Mid(Aux, 1, 1) = "-" Then
        Monto2 = Monto2 * -1
    End If

    'BAE02   CHAR    1
    Columna = Columna + 1
    pos1 = 438
    pos2 = 1
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'FRE02   NUMC    3   Frecuencia de periodicidad
    Columna = Columna + 1
    pos1 = 439
    pos2 = 3
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'GRA02   NUMC    3   Gracia para aplicacion del concepto
    Columna = Columna + 1
    pos1 = 442
    pos2 = 3
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'NCT02   NUMC    3   Numero de cuotas-concepto de aplicacion
    Columna = Columna + 1
    pos1 = 445
    pos2 = 3
    Aux = Mid(strlinea, pos1, pos2)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Cuotas2 = Aux
    
    'TOT02   CURR    16.2    Monto
    Columna = Columna + 1
    pos1 = 438
    pos2 = 16
    Aux = ""
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    
    'Columnas vacias
    For Columna = 48 To 191
        Aux = ""
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo_9302, Columna, Aux)
    Next Columna
    
       
    
    '---------------------------------------------------------------
    'Inserto Cuota Prestamo
    Call Insertar_Prestamo(Empleado.Tercero, Nro_Prestamo, Tipo_Prestamo, Poliza, Monto_Prestamo, Modalidad_Prestamo, Cuotas_Prestamo, Concepto1, Monto1, Cuotas1, Concepto2, Monto2, Cuotas2, Nro_Moneda, Inicio_Validez, Fin_Validez)

Exit Sub

Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error en infotipo " & Infotipo
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
End Sub


Public Sub Insertar_Linea_Segun_Modelo(ByVal Linea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador de acurdo al modelo
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
MyBeginTrans
'    Flog.Writeline Espacios(Tabulador * 1) & "Comienza Transaccion"

    HuboError = False
    
    Select Case NroModelo
    Case 248: 'infotipos
        'Reservado en otro proceso
    Case 249: 'Interfase de Mapeos para Infotipos
        Call LineaModelo_249(Linea)
    End Select

MyCommitTrans
If Not HuboError Then
    'MyCommitTrans
'    Flog.Writeline Espacios(Tabulador * 1) & "Transaccion Cometida"
Else
    'MyRollbackTrans
'    Flog.Writeline Espacios(Tabulador * 1) & "Transaccion Abortada"
End If
End Sub




Public Sub LineaModelo_249(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : FGZ
' Fecha      : 09/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Aux As String
Dim Aux2 As String

Dim Tabla As String
Dim Codigo As String
Dim Descripcion As String
Dim Tablaref As String
Dim Tenro As Long
Dim Nro_Pais As Long
Dim Nro_Provincia As Long
Dim Nro_Estr As Long
Dim Nro_Tercero As Long
Dim Inserto_estr As Boolean
Dim Nro_EstadoCivil As Long
Dim Nro_Moneda As Long

Dim rs As New ADODB.Recordset

' El formato es:    tabla [auxiliar] codigo descripcion
    
    On Error GoTo Manejador_De_Error
    
    'TABLA
    pos1 = 1
    pos2 = 2
    Tabla = Mid$(strlinea, pos1, pos2)
    
    'Auxiliar
    pos1 = 3
    pos2 = 4
    Aux = Mid$(strlinea, pos1, pos2)
    
    'Auxiliar 2
    pos1 = 38
    pos2 = 4
    Aux2 = Mid$(strlinea, pos1, pos2)
    
    'Codigo
    pos1 = 38
    pos2 = 10
    Codigo = Trim(Mid(strlinea, pos1, pos2))
    Codigo = Replace(Codigo, "'", "´")
    
    'Descripcion
    pos1 = 53
    pos2 = Len(strlinea)
    If pos2 < pos1 Then
        Descripcion = ""
    Else
        Descripcion = Mid(strlinea, pos1, (pos2 - pos1) + 1)
        Descripcion = Replace(Descripcion, "'", "´")
    End If

' ====================================================================
'   Validar los parametros Levantados

'Solo se levantan algunas de las tablas
Select Case CLng(Tabla)
Case 1:  'Clase de Medida T529A           NO
    Texto = "Clase de Medida T529A           NO"
    FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto
Case 2:  'Motivo de la Medida T530            NO
    Texto = "Motivo de la Medida T530            NO"
    FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto
Case 3: 'Sociedad    T001    10  Empresa si
    Tablaref = "T001"
    Tenro = 10
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
            Call CreaTercero(12, Descripcion, Nro_Tercero)
            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 4: 'Division de Personal    T500P   6   Gerencia    si
    Tablaref = "T500P"
    Tenro = 6
    If Not EsNulo(Trim(Codigo)) And Mid(UCase(Trim(Aux2)), 1, 2) = "AR" Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
'            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 5: 'Grupo de Personal   T501    48  Grupo de Personal   si
    Tablaref = "T501"
    Tenro = 48
    If Not EsNulo(Trim(Codigo)) And (UCase(Mid(Trim(Descripcion), 1, 2)) = "AR" Or Trim(Codigo) = "1" Or Trim(Codigo) = "2" Or Trim(Codigo) = "4" Or Trim(Codigo) = "9") Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
'            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If


Case 6: 'Area de Personal    T503K   49  Area de Personal    si
    Tablaref = "T503K"
    Tenro = 49
    If Not EsNulo(Trim(Codigo)) And UCase(Mid(Trim(Descripcion), 1, 2)) = "AR" Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
'            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 7: 'Subdivision de Personal T001P   1   Sucursal    si
    Tablaref = "T001P"
    Tenro = 1
    If Not EsNulo(Trim(Codigo)) And UCase(Mid(Trim(Aux), 1, 2)) = "AR" Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
            Call CreaTercero(10, Descripcion, Nro_Tercero)
            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 8: 'Area de Nomina  T549A   22  FORMA DE LIQUIDACION    NO
    Tablaref = "T549A"
    Tenro = 22
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
'            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If
Case 9: 'Centro de Costo CSKS    5   Centro de Costo si
    Tablaref = "CSKS"
    Tenro = 5
    If Not EsNulo(Trim(Codigo)) Then
        Descripcion = Trim(Codigo) & " $ " & Trim(Descripcion)
        Call ValidaEstructura(Tenro, Trim(Descripcion), Nro_Estr, Inserto_estr)
        If Inserto_estr Then
'            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

'10 Posicion    T528B           NO
'11 Funcion T513            NO
Case 12:    'Unidad Organizativa T527X   2   SECTOR  SI
    Tablaref = "T527X"
    Tenro = 2
    If Not EsNulo(Codigo) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
'            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 13:  'Pais    T005        PAIS    SI
    Tablaref = "T005"
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidarPais(Descripcion, Nro_Pais)
        Call Mapear(Tablaref, Codigo, CStr(Nro_Pais))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 14:  'Estados - Provincia T005S       PROVINCIA   SI
    Tablaref = "T005S"
    If Not EsNulo(Trim(Codigo)) And Mid(Descripcion, 1, 2) = "AR" Then
        Call ValidarPais("ARGENTINA", Nro_Pais)
        Call ValidarProvincia(Mid(Descripcion, 4, Len(Descripcion) - 3), Nro_Provincia, Nro_Pais)
        Call Mapear(Tablaref, Codigo, CStr(Nro_Provincia))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

'15  Nacionalidad (NO SE HACE PORQUE ES IGUAL AL 13)
'16
Case 17:    'Estado Civil    T502T       ESTCIVIL    SI
    Tablaref = "T502T"
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidarEstadoCivil(Descripcion, Nro_EstadoCivil)
        Call Mapear(Tablaref, Codigo, CStr(Nro_EstadoCivil))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 18:    'Bancos  T012    41  BANCO   SI
    Tablaref = "T012"
    Tenro = 41
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
            Descripcion = Format_Str(Descripcion, 40, False, "")
            Call CreaTercero(13, Descripcion, Nro_Tercero)
            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 19:    'Clave de Moneda TCURC       MONEDA  SI
    Tablaref = "TCURC"
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidarMoneda(Descripcion, Codigo, 1, Nro_Moneda)
        Call Mapear(Tablaref, Codigo, CStr(Nro_Moneda))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If
Case 20:
Case 21:
Case 22:
Case 23:
Case 24:
Case 25:
Case 26:
Case 27:  'Conceptos

Case 28:
Case 29:
Case 30:
Case 31:
Case 32:
Case 33:
Case 34:
Case 35:  'Relacion Laboral    T542A           NO
Case 36:  'Contrato Actual T547V   18  CONTRATO ACTUAL SI
    Tablaref = "T547V"
    Tenro = 18
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
'            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If
Case 37:
Case 38:
Case 39:
Case 40:
Case 41:  'Obra Social
    Tablaref = "T7AR34"
    Tenro = 17
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
            Call CreaTercero(4, Descripcion, Nro_Tercero)
            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If
Case 42:  'AFJP
    Tablaref = "T7AR36"
    Tenro = 15
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
            Call CreaTercero(6, Descripcion, Nro_Tercero)
            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 43:  'Codigo de Actividad del Empleado
    Tablaref = "T7AR38"
    Tenro = 29
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
'            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If
Case 44:  'Agencia IRS
    Tablaref = "T7AR66"
    Tenro = 28
    If Not EsNulo(Trim(Codigo)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
            Call CreaTercero(7, Descripcion, Nro_Tercero)
            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If
Case Else
End Select

Fin:
    'Cierro todo y libero
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strlinea
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

