Attribute VB_Name = "MdlSAP"
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
Global UltimoInfotipo As String

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

Global Cant_Acumulada As Integer
Global Monto_Acumulado As Single
Global Infotipo_Campo As Integer
Global IndiceArr As Long

Global FormatoFechaSap1 As String '= "dd.mm.yyyy"
Global NuloFechaSap1 As String '= "31.12.9999"
Global FormatoFechaSap2 As String '= "yyyymmdd"
Global NuloFechaSap2 As String '= "99991231"
Global TipoDocLSAP As Long


Private Function Es_ABM(ByVal Linea As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que determina si la linea corresponde a
'               Altas, Bajas & Modificaciones ó
'               Salario Basico, Adicionales, Ausencias y presencias
' Autor      : FGZ
' Fecha      : 09/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim ABM As Boolean
Dim Lista
Dim i As Integer
Dim pos As Integer

Lista = Split("IT0008;IT0014;IT0015;IT0021;IT0185;IT0392;IT0394;IT2001;IT2002", ";")

ABM = True
For i = LBound(Lista) To UBound(Lista)
    pos = InStr(1, Linea, Lista(i))
    If pos > 0 Then
        ABM = False
    End If
Next i
Es_ABM = ABM
End Function

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
Dim UltimaLinea As Boolean
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset

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
    
    
    'FGZ - 30/03/2006
    'Configuracion (confrep)
    'Valores por default
    FormatoFechaSap1 = "dd.mm.yyyy"
    NuloFechaSap1 = "31.12.9999"
    
    FormatoFechaSap2 = "dd.mm.yyyy"
    NuloFechaSap2 = "31.12.9999"
    
    StrSql = "SELECT conftipo,confval2  FROM confrep "
    StrSql = StrSql & " WHERE repnro = 159 "
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    Do While Not rs_Confrep.EOF
        Select Case UCase(rs_Confrep!conftipo)
        Case "FD1":
            FormatoFechaSap1 = rs_Confrep!confval2
        Case "FD2":
            FormatoFechaSap2 = rs_Confrep!confval2
        Case "FN1":
            NuloFechaSap1 = rs_Confrep!confval2
        Case "FN2":
            NuloFechaSap2 = rs_Confrep!confval2
        Case "LEG":
            TipoDocLSAP = CLng(rs_Confrep!confval2)
        Case Else
        End Select
        
        rs_Confrep.MoveNext
    Loop
    
    'Crea la planilla de Excel que contendrá la informacion leida
    Call CrearArchivoExcel(ArchivoAGenerar)
    
    'Creo el csv con la informacion levantada
    Set fNovedades = fs.CreateTextFile(ArchivoNovedades, True)
    Set fCambios = fs.CreateTextFile(ArchivoCambios, True)
    Call InsertarLogEncabezadoNovedad
    Call InsertarLogEncabezadoCambios
    UltimoLegajo = -1
    Fila_Infotipo = 1
    If Not f.AtEndOfStream Then
        strlinea = f.ReadLine
    End If
    Do While Not f.AtEndOfStream
        strlinea = f.ReadLine
        UltimaLinea = f.AtEndOfStream
        NroLinea = NroLinea + 1
        If NroLinea = 1 And UsaEncabezado Then
            strlinea = f.ReadLine
            UltimaLinea = f.AtEndOfStream
        End If
        If Trim(strlinea) <> "" Then
            RegLeidos = RegLeidos + 1
            
            Select Case rs_Modelo!modinterface
                Case 4:
                    'Debo decidir por el tipo de linea leida si se trata de un formato u otro
                    If Es_ABM(strlinea) Then
                        Call Insertar_Linea_ABM(strlinea)
                    Else
                        Call Insertar_Linea_Infotipo(strlinea, UltimaLinea)
                    End If
                Case Else
                    'para levantar las tablas globales (MAPEOS)
                    Call Insertar_Linea_Segun_Modelo(strlinea)
                    'Flog.writeline Espacios(Tabulador * 1) & "para levantar las tablas globales (MAPEOS). No activo"
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
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    Set rs_Lineas = Nothing
    Set rs_Confrep = Nothing
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
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
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
    Case 277: 'Interfase de Mapeos de tablas planas
        Call LineaModelo_277(Linea)
    Case Else
    
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



Public Sub Insertar_Linea_ABM(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: ABM
' Autor      : FGZ
' Fecha      : 09/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim aux
Dim OK As Boolean
Dim Columna As Byte
Dim Obligatorio As Boolean
Dim Texto
Dim Inicio_Validez
Dim Fin_Validez
Dim Fecha_Medida
Dim Tenro As Long
Dim TipoFechaFase As String

Dim Obra_Social As String
Dim Descripcion_OSocial As String
Dim Nro_Obra_Social As Long

Dim Nro_Plan_Obra_Social As Long
Dim Plan_Obra_Social As String

Dim Grupo_SIJP As String
Dim Nro_Grupo_SIJP As Long
Dim Descripcion_Grupo_SIJP As String
Dim Capitalizacion As Boolean
Dim Reparto As Boolean
Dim Nro_caja As Long
Dim Actividad_SIJP As String
Dim Nro_Actividad_SIJP As Long

Dim Categoria As String
Dim Nro_Categoria As Long

Dim Reg_Tercero As TReg_Tercero
Dim Estructuras(100) As TEstructuras

Dim Hoja As Integer
Dim Lista
Dim i As Integer

'--------------------
'IT0185
'--------------------
Dim Nro_Tipo_ID As Long
Dim Descripcion As String
Dim Nro_Identific As String
'--------------------
'IT0002
'--------------------
Dim Apellido As String
Dim Apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim Cuil As String
Dim Sexo As String  '{Femenino - Masculino}
Dim Fecha_de_nacimiento As Date
Dim Estado_civil As String
Dim Nacionalidad As String
Dim Pais_de_nacimiento As String

'--------------------
'IT0001
'--------------------
Dim Division_de_Personal As String
Dim Gerencia As Long
Dim Descripcion_Division_de_Personal As String
Dim Grupo_de_Personal As String
Dim Nro_Grupo_de_Personal As Long
Dim Area_de_Personal As String
Dim Nro_Area_de_Personal As Long
Dim Grupo_Liq As Long
Dim Sector As Long
Dim Area_Funcional As String
'--------------------
'IT0000
'--------------------
Dim Medida As String
Dim Motivo As String
Dim Status As String
Dim Causa As Long

'--------------------
'IT0001
'--------------------
Dim Codigo_Compania As String
Dim Descripcion_Compania As String
Dim Empresa As Long
Dim SubDivision_de_Personal As String
Dim Descripcion_Subdivision_Personal As String
Dim Sucursal As Long
Dim Centro_de_Costo As String
Dim Descripcion_CC As String
Dim CCosto As Long
Dim Area_de_Nomina As String
Dim Nro_Area_de_Nomina As Long
Dim Relacion_Laboral As String
Dim Descripcion_Relacion_Laboral As String
Dim Convenio As Long
Dim Sindicato As Long
Dim Posicion As String
Dim Descripcion_Posicion As String
Dim Funcion As String
Dim Nro_Puesto As Long
Dim Descripcion_Funcion As String
Dim Unidad_Organizativa As String
Dim Descripcion_Unidad_Organizativa As String
'Dim Sector As Long
'--------------------
'IT0001
'--------------------
Dim Contrato As String
Dim Descripcion_Contrato As String
Dim Fecha_Fin_Contrato As String
Dim Nro_Contrato As Long
Dim Periodo_Prueba As String
Dim Periodo_Prueba_UN As String
'--------------------
'IT0007
'--------------------
Dim Plan_Horario_Trabajo As String
Dim Descripcion_Horario_Trabajo As String
Dim Regimen_Horario As Long
Dim Porcentaje_Horario_Trabajo As String
'--------------------
'IT0009
'--------------------
Dim Clase_Datos_Bancarios As String
Dim Descripcion_Clase_banco As String
Dim Clave_de_Banco As String
Dim Nro_Banco As Long
Dim Descricpion_Clase_Banco As String
Dim Banco As String
Dim Cuenta_Bancaria As String
Dim Clave_Control_Banco As String
Dim Via_de_Pago As String
Dim Nro_FormaPago As Long
Dim Destino_para_Transferencias As String
Dim CBU As String
Dim Porcentaje_Prefijado As Integer
Dim Banco_Sucursal As String


Dim Tipo_Servicio As String
Dim TipoDomi As Long
Dim NroDom As Long
Dim Calle As String 'Street
Dim Nro As String   'number
Dim Piso As String  'Plant
Dim Dpto As String  'House
Dim Torre As String
Dim Manzana As String
Dim Entre As String
Dim Codigo_postal As String    'Zip Code
Dim Estado As String 'State
Dim Region As String 'Region
Dim Nro_Pais As Long
Dim Nro_Provincia As Long
Dim Nro_Localidad As Long
Dim Nro_Partido As Long
Dim Tablaref As String
Dim Inserto_estr As Boolean
Dim Nro_Estr As Long
Dim Nro_Tercero As Long
Dim Aux_Legajo As String

Dim rs_Cta As New ADODB.Recordset
Dim rs_caja  As New ADODB.Recordset


'--------------------
    
    On Error GoTo Manejador_De_Error
    
    Hoja = 1
    Fila_Infotipo = Fila_Infotipo + 1
    
    'Leo todos los campos
    Lista = Split(strlinea, Separador)
    
    
    'Start Date
    Obligatorio = True
    Columna = 52
    Texto = "Start Date"
    If Columna <= UBound(Lista) Then
        Inicio_Validez = StrToDate(Trim(Lista(Columna)), OK, FormatoFechaSap1, NuloFechaSap1)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        If Not OK Then
            Flog.writeline Espacios(Tabulador * 3) & "Valor en Dato incorrecto"
            Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
            If Obligatorio Then Exit Sub
        End If

    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    'End Date
    Obligatorio = True
    Columna = 53
    Texto = "End Date"
    If Columna <= UBound(Lista) Then
        Fin_Validez = StrToDate(Trim(Lista(Columna)), OK, FormatoFechaSap1, NuloFechaSap1)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        If Not OK Then
            Flog.writeline Espacios(Tabulador * 3) & "Valor en Dato incorrecto"
            Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
            If Obligatorio Then Exit Sub
        End If

    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    'CUIL -
    Obligatorio = True
    Columna = 7
    Cuil = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    
    'Nro de Legajo - PRNR
    Obligatorio = True
    Columna = 0
    If IsNumeric(Lista(Columna)) Then
        Aux_Legajo = Lista(Columna)
        Empleado.Legajo = BuscarLegajo(Cuil, Aux_Legajo)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
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
        'Infotipo_Invalido = True
        If Obligatorio Then
            Exit Sub
        End If
    End If
    
    '---------------------------------------
    'Campos de Infotipo 0000 - Medidas
    '---------------------------------------
    'columnas 16..18
        
        
    'Tipo de Medida - MASSG
    Obligatorio = True
    Columna = 17
    Texto = "Tipo de Medida"
    Medida = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    '1L     Contratación
    '2L     Reasignacion Organizacional
    '3L     Transferencia
    '4L     Baja
    '5L     Reingreso a la Cia
    '6L     Cambio de Salario
    '7L     Licencia / Ausencias
    '8L     Reincorporación Ausencia
        
    'Motivo de la medida - MGTXT
    Obligatorio = True
    Columna = 18
    Texto = "Motivo para el cambio de datos maestros"
    Motivo = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    'Medida                         Motivo
    'Contratacion                   00  N/A
    '                               01  Fusión
    '                               02  Nueva posición
    '                               03  Reemplazo / Sustitución
    '                               04  Expatriación
    '                               05  Transferencia
    '                               06  Empleado temporario
    '                               07  Empleado Externo
    'Reasignacion Organizacional    00  N/A
    '                               01  Cambio de Posición
    '                               02  Promoción Normal
    '                               03  Plan de desarrollo global (GDP)
    '                               05  Transferencia
    '                               06  Cambio de contratación externo
    '                               07  Control de headcount
    '                               08  posición compartida
    '                               09  Cambio posición compartida externo
    '                               10  Cambio posición compartida local
    '                               11  Fin Extranjero Secondment
    '                               12  Inicio Extranjero Secondment
    'Transferencia                  01  Legal
    '                               02  Transferencia
    '                               03  Centro de costo
    '                               04  GDP
    'Baja                           01  Reasig. Inicio Extranjero Secondment
    '                               02  Reasig. Fin Extranjero Secondment
    '                               03  Despido
    '                               04  Vencimiento de contrato
    '                               05  Renuncia
    '                               06  Jubilación
    '                               07  Muerte
    '                               08  enfermedad
    '                               09  Sindical
    '                               10  Fusión
    '                               12  Reasig.Transferencia Extranjero
    '                               15  IMP
    'Reingreso a la compañia        01  Nueva Posición
    '                               02  Reemplazo / sustitución
    '                               03  Expatriado
    '                               04  Empleado temporario
    '                               05  Fusion
    '                               06  Transferencia
    '                               07  Programa de Movilidad Interna (IMP)
    'Cambio de Salario              01  Promoción / Ascenso
    '                               02  por Ley
    '                               03  Ajuste por inflación
    '                               04  Ajuste por Mérito
    '                               05  Ajuste
    'Licencias / Ausencias          02  Educación
    '                               03  Asunto personales
    '                               04  Licencia por enfermedad
    '                               05  Maternidad
    '                               06  Licencia paga
    '                               07  Licencia no paga
    '                               08  Adopción
    '                               09  Apercibimiento / suspensión
    '                               10  Accidente
    '                               11  Servicio militar
    '                               12  Paternidad
    'Reincorporación de la licencia / Ausencia
    '                               01  Reincorporación
        
    'Status - STAT2
    Obligatorio = False
    Columna = 19
    Status = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Select Case UCase(Medida)
    Case "BAJA":
        TipoFechaFase = ""
        Accion = "BAJA"
        Select Case UCase(Motivo)
        Case "REASIG. INICIO EXTRANJERO SECONDMENT":
            SubAccion = "Reasig. Inicio Extranjero Secondment"
        Case "REASIG. FIN EXTRANJERO SECONDMENT":
            SubAccion = "Reasig. Fin Extranjero Secondment"
        Case "DESPIDO":
            SubAccion = "Despido"
        Case "VENCIMIENTO DE CONTRATO":
            SubAccion = "Vencimiento de contrato"
        Case "RENUNCIA":
            SubAccion = "Renuncia"
        Case "JUBILACION":
            SubAccion = "Jubilacion"
        Case "MUERTE":
            SubAccion = "Muerte"
        Case "ENFERMEDAD":
            SubAccion = "Enfermedad"
        Case "SINDICAL":
            SubAccion = "Sindical"
        Case "FUSION":
            SubAccion = "Fusion"
        Case "REASIG. TRANSFERENCIA EXTRANJERO":
            SubAccion = "Reasig.Transferencia Extranjero"
        Case "IMP":
            SubAccion = "IMP"
        Case Else:
            FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": " & "Motivo de Medida desconocido " & Motivo
        End Select
        Causa = CLng(CalcularMapeoInv(Medida_Motivo, "CAUSAB", "0"))
        OK = False
        If Causa <> 0 Then
            OK = Baja_Empleado(Empleado.Tercero, Inicio_Validez, Causa)
        End If
        If Not OK Then
            Flog.writeline Espacios(Tabulador * 3) & "Error dando de baja el empleado."
            FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ": Error dando de baja el legajo."
        End If
    Case "CONTRATACIÓN", "CONTRATACION", "HIRING":
        TipoFechaFase = "01"
    Case Else
        TipoFechaFase = ""
        'las demas causas no hacen diferencia
    End Select
    
    
    '---------------------------------------
    'Campos de Infotipo 0002 - Datos Personales
    '---------------------------------------
    'columnas 4..10
    
    'Apellido - NACHN
    Obligatorio = True
    Columna = 4
    Apellido = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Reg_Tercero.Terape = EliminarCHInvalidos(Trim(Format_Str(Apellido, 25, True, " ")))

    'Segundo Apellido - NACH2
    Obligatorio = False
    Columna = 5
    Apellido2 = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Reg_Tercero.Terape2 = EliminarCHInvalidos(Trim(Format_Str(Apellido2, 25, True, " ")))
        
    'Nombre - VORNA
    Obligatorio = True
    Columna = 6
    nombre = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Reg_Tercero.Ternom = EliminarCHInvalidos(Trim(Format_Str(nombre, 25, True, " ")))
    
'    'Nombre2 - VORNA
'    Obligatorio = True
'    Columna = 7
'    nombre2 = Trim(Lista(Columna))
'    Reg_Tercero.Ternom2 = EliminarCHInvalidos(Trim(Format_Str(nombre2, 25, True, " ")))

    'Genero - GESC1/2
    Obligatorio = False
    Columna = 8
    Sexo = UCase(Trim(Lista(Columna)))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    'Default = Masculino
    If Sexo = "MASCULINO" Then
        Reg_Tercero.Tersex = -1
    Else
        Reg_Tercero.Tersex = 0
    End If
    
    'Fecha de Nacimiento - GBDAT
    Obligatorio = True
    Columna = 9
    Texto = "Fecha de Nacimiento"
    'Fecha_de_Nacimiento = StrToFecha(Replace(Lista(Columna), ".", ""), OK)
    Fecha_de_nacimiento = StrToDate(Lista(Columna), OK, FormatoFechaSap1, NuloFechaSap1)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Valor en Dato incorrecto"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        If Obligatorio Then Exit Sub
    End If
    Reg_Tercero.Terfecnac = Fecha_de_nacimiento
    
    'Estado civil - FATXT
    Obligatorio = True
    Columna = 10
    Estado_civil = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Reg_Tercero.EstCivNro = CLng(CalcularMapeoInv(Estado_civil, "ESTCIV", "-1"))
    
    'Nacionalidad - NATIO
    Obligatorio = True
    Columna = 11
    Nacionalidad = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Reg_Tercero.NacionalNro = CLng(CalcularMapeoInv(Nacionalidad, "NACION", "-1"))

    'El pais no viene informado y es obligatorio, le pongo lo mismo que el pais de nacimiento
    'Pais_de_Nacimiento = "Argentina"
    Reg_Tercero.PaisNro = Reg_Tercero.NacionalNro
    Pais_de_nacimiento = Nacionalidad
    
'    If Not EsNulo(Pais_de_nacimiento) Then
'        Reg_Tercero.PaisNro = CLng(CalcularMapeoInv(Pais_de_nacimiento, "T005", "0"))
'    Else
'        Reg_Tercero.PaisNro = 0
'        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Pais de Nacimiento"
'    End If

    'Domicilio
    'Calle - Street
    Obligatorio = False
    Columna = 63
    If Columna <= UBound(Lista) Then
        Calle = EliminarCHInvalidos(Trim(Lista(Columna)))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    'Nro - N
    Obligatorio = False
    Columna = 64
    If Columna <= UBound(Lista) Then
        Nro = EliminarCHInvalidos(Trim(Lista(Columna)))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    'Piso - plant
    Obligatorio = False
    Columna = 65
    If Columna <= UBound(Lista) Then
        Piso = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    'Departamento - House
    Obligatorio = False
    Columna = 66
    If Columna <= UBound(Lista) Then
        Dpto = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    'CP - Zip Code
    Obligatorio = False
    Columna = 67
    If Columna <= UBound(Lista) Then
        Codigo_postal = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    'Provincia - State
    Obligatorio = False
    Columna = 68
    If Columna <= UBound(Lista) Then
        Estado = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    'Region - Region
    Obligatorio = False
    Columna = 69
    If Columna <= UBound(Lista) Then
        Region = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    TipoDomi = 2    'Particular
    'Las localidades no tienen cargado el Cp con lo cual con este dato no voy a encontrar la localidad
    'Pais y provincia
    
    Nro_Pais = Reg_Tercero.PaisNro
    If Not EsNulo(Region) Then
        Nro_Provincia = CLng(CalcularMapeoInv(Region, "PROVIN", "-1"))
    Else
        Nro_Provincia = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Region"
    End If
    If Not EsNulo(Trim(Estado)) Then
        Call ValidarLocalidad(Estado, Nro_Localidad, Nro_Pais, Nro_Provincia)
    Else
        Nro_Localidad = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Region"
    End If
    Nro_Partido = 0
    
    If EsNulo(Codigo_postal) Then
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Cod. Postal"
    End If

    If Not ExisteLegajo Then
        'Inserto el tercero
        StrSql = " INSERT INTO tercero ("
        StrSql = StrSql & " ternom,"
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
        StrSql = StrSql & " ,empest"
        StrSql = StrSql & " ,empfaltagr"
        StrSql = StrSql & " ,ternro"
        StrSql = StrSql & " ,terape"
        If Not EsNulo(Reg_Tercero.Terape2) Then
            StrSql = StrSql & " ,terape2"
        End If
        StrSql = StrSql & " ,ternom"
        StrSql = StrSql & " ,empnro"
        StrSql = StrSql & " ) VALUES( "
        StrSql = StrSql & Empleado.Legajo
        StrSql = StrSql & "," & ConvFecha(Inicio_Validez)
        StrSql = StrSql & ",-1 "
        StrSql = StrSql & "," & ConvFecha(Inicio_Validez)
        StrSql = StrSql & "," & Empleado.Tercero
        StrSql = StrSql & ",'" & Reg_Tercero.Terape & "'"
        If Not EsNulo(Reg_Tercero.Terape2) Then
            StrSql = StrSql & ",'" & Reg_Tercero.Terape2 & "'"
        End If
        StrSql = StrSql & ",'" & Reg_Tercero.Ternom & "'"
        StrSql = StrSql & ",1"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " INSERT INTO ter_tip(ternro,tipnro) VALUES(" & Empleado.Tercero & ",1)"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Insertar Domicilios
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
            StrSql = StrSql & ",'" & Format_Str(Calle, 30, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Nro, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Piso, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Dpto, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Torre, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Manzana, 8, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Codigo_postal, 12, False, "") & "'"
            StrSql = StrSql & ",'" & Format_Str(Entre, 80, False, "") & "'"
            StrSql = StrSql & "," & Nro_Localidad
            StrSql = StrSql & "," & Nro_Provincia
            StrSql = StrSql & "," & Nro_Pais
            StrSql = StrSql & "," & Nro_Partido
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'MMMMMM .... no se
        End If
        
    Else    'Actualizo los datos
    
        'Tercero
        StrSql = " UPDATE tercero SET "
        StrSql = StrSql & " ternom = '" & Reg_Tercero.Ternom & "'"
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
        If Not EsNulo(Fin_Validez) Then
            StrSql = StrSql & ",empfecbaja = " & ConvFecha(Fin_Validez)
        End If
        StrSql = StrSql & ",empest = -1 "
        StrSql = StrSql & ",terape = '" & Reg_Tercero.Terape & "'"
        If Not EsNulo(Reg_Tercero.Terape2) Then
            StrSql = StrSql & ",terape2 = '" & Reg_Tercero.Terape2 & "'"
        End If
        StrSql = StrSql & ",ternom = '" & Reg_Tercero.Ternom & "'"
        StrSql = StrSql & ",empnro = 1"
        StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
  
    'Fases
    If Not EsNulo(TipoFechaFase) Then
        Call Insertar_Fecha(TipoFechaFase, CDate(Inicio_Validez), Inicio_Validez, Fin_Validez)
    End If

    '---------------------------------------
    'Campos de Infotipo 0185 - Documentacion
    '---------------------------------------
    'columnas 1..3
        
    'Tipo de ID - ICTYP
    Obligatorio = True
    Columna = 1
    Nro_Tipo_ID = CLng(CalcularMapeoInv(Trim(Lista(1)), "TIPDOC", "0"))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
   
    'Descripcion - ICTXT
    Obligatorio = False
    Columna = 2
    Descripcion = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        
    'Nro de identidad - ICNUM
    Obligatorio = True
    Columna = 3
    Nro_Identific = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    
    If Not EsNulo(Nro_Identific) Then
        Call Insertar_Documento(Nro_Identific, Nro_Tipo_ID)
    End If
    
    'Esto esta puesto a dedo
    If Not EsNulo(Aux_Legajo) Then
        Call Insertar_Documento(Aux_Legajo, TipoDocLSAP)
    End If
    
    
    '---------------------------------------
    'Campos de Infotipo 0001 - Asignacion Organizacional
    '---------------------------------------
    'columnas 12..15, 19..33
        
    'Division de Personal - WERKS
    Obligatorio = True
    Columna = 12
    Division_de_Personal = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Division de Personal - [ RHPro(Gerencia)] " & Division_de_Personal
    Tablaref = "GERENC"
    Tenro = 6
    Gerencia = CLng(CalcularMapeoInv(Division_de_Personal, Tablaref, "-1"))
    If Gerencia = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline
        
        'Descripcion
        Obligatorio = True
        Columna = 13
        Descripcion_Division_de_Personal = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        
        If Not EsNulo(Trim(Descripcion)) Then
            Call ValidaEstructura(Tenro, Descripcion_Division_de_Personal, Nro_Estr, Inserto_estr)
            If Inserto_estr Then
                'Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
                'Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
            End If
            Call Mapear(Tablaref, Division_de_Personal, CStr(Nro_Estr))
            Gerencia = Nro_Estr
        Else
            Flog.writeline Espacios(Tabulador * 4) & "**** Descripción nula, no se puede insertar"
        End If
    Else
        'Descripcion
        Obligatorio = True
        Columna = 13
        Descripcion_Division_de_Personal = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    End If
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Gerencia
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez

        
    'Grupo de Personal - [ RHPro(Tipo de personal)] - PERSG
    Obligatorio = True
    Columna = 14
    Grupo_de_Personal = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Grupo de Personal - [ RHPro(Grupo de Personal)] " & Grupo_de_Personal
    Tablaref = "GRUPER"
    Tenro = 48
    Nro_Grupo_de_Personal = CLng(CalcularMapeoInv(Grupo_de_Personal, Tablaref, "-1"))
    If Nro_Grupo_de_Personal = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline Espacios(Tabulador * 4) & "**** No hay Descripción, no se puede insertar"
        Flog.writeline
    End If
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Nro_Grupo_de_Personal
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez


    'Area de Personal - RHPro(Sector y Grupo de Liquidacion)
    Obligatorio = True
    Columna = 15
    Area_de_Personal = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    'Mapeo Manual tanto para Sector como para grupo de liquidacion
    Select Case UCase(Area_de_Personal)
    Case "1", "01", "SOCIO INTERNACIONAL":
        Sector = 3
        Grupo_Liq = 625
    Case "2", "02", "SOCIO":
        Sector = 3
        Grupo_Liq = 625
    Case "3", "03", "SOCIO ASALARIADO":
        Sector = 3
        Grupo_Liq = 625
    Case "4", "04", "DIRECTOR":
        Sector = 4
        Grupo_Liq = 624
    Case "5", "05", "GERENTE SENIOR (FC)":
        Sector = 4
        Grupo_Liq = 624
    Case "6", "06", "GERENTE SENIOR (N/F)":
        Sector = 4
        Grupo_Liq = 624
    Case "7", "07", "GERENTE (FC)":
        Sector = 4
        Grupo_Liq = 624
    Case "8", "08", "GERENTE (N/F)":
        Sector = 4
        Grupo_Liq = 624
    Case "9", "09", "STAFF (FC)":
        Sector = 4
        Grupo_Liq = 624
    Case "10", "STAFF (N/F)":
        Sector = 4
        Grupo_Liq = 624
    Case "11", "PASANTE (FC)":
        Sector = 4
        Grupo_Liq = 626
    Case "12", "PASANTE (N/F)":
        Sector = 4
        Grupo_Liq = 626
    Case Else
        Sector = 4
        Grupo_Liq = 624
        Flog.writeline Espacios(Tabulador * 3) & "No se encontró el Sector/Grupo de Liquidación " & Area_de_Personal
        Flog.writeline Espacios(Tabulador * 3) & "Se cargarán defaults 09"
    End Select
    'Sector
    Tenro = 2
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Sector
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
    
    'Grupo de liquidacion
    Tenro = 32
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Grupo_Liq
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
        
        
    'Area Funcional - BUS_AR_DES
    Obligatorio = False
    Columna = 16
    Area_Funcional = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Area Funcional - [ RHPro (Sucursal)] " & Area_Funcional
    'Sucursal = CLng(CalcularMapeoInv(Area_Funcional, "TSUC", "-1"))
    'Mapeo Manual
    Select Case UCase(Area_Funcional)
    Case "ARBA":
        Sucursal = 1
    Case "ARRO":
        Sucursal = 2
    Case Else
        Flog.writeline Espacios(Tabulador * 3) & "No se encontró la Sucursal " & Area_Funcional
        Flog.writeline Espacios(Tabulador * 3) & "Se cargará la Sucursal por default 1"
    End Select
    'Sucursal
    Tenro = 1
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Sucursal
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
       
        
    '---------------------------------------
    'Campos de Infotipo 0001 - Asignacion Organizacional
    '---------------------------------------
    'columnas 20..33
        
    'Codigo de Compañia - BUKRS
    Obligatorio = True
    Columna = 20
    Codigo_Compania = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Compañia - [ RHPro(Empresa)] " & Codigo_Compania
    Tablaref = "EMPRES"
    Tenro = 10
    Empresa = CLng(CalcularMapeoInv(Codigo_Compania, Tablaref, "-1"))
    If Empresa = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline
        
        'Descripcion de Compañia - COMP_NAME
        Obligatorio = True
        Columna = 21
        Descripcion_Compania = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        Texto = "Nombre Compañia - [ RHPro(Empresa)] " & Descripcion_Compania
        If Not EsNulo(Trim(Descripcion_Compania)) Then
            Call ValidaEstructura(Tenro, Descripcion_Compania, Nro_Estr, Inserto_estr)
            If Inserto_estr Then
                Call CreaTercero(12, Descripcion_Compania, Nro_Tercero)
                Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion_Compania)
            End If
            Call Mapear(Tablaref, Codigo_Compania, CStr(Nro_Estr))
            Empresa = Nro_Estr
        Else
            Flog.writeline Espacios(Tabulador * 4) & "**** Descripción nula, no se puede insertar"
        End If
    Else
        'Descripcion de Compañia - COMP_NAME
        Obligatorio = True
        Columna = 21
        Descripcion_Compania = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    End If
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Empresa
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
    
        
    'Subdivision de Personal - [RHPro (Sucursal)] - BTRTL
    Obligatorio = True
    Columna = 22
    SubDivision_de_Personal = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
'    Texto = "Subdivision de Personal - [ RHPro (Sucursal)] " & SubDivision_de_Personal
'    Sucursal = CLng(CalcularMapeoInv(SubDivision_de_Personal, "T001P", "-1"))
'    Tenro = 1
'    Estructuras(Tenro).Tenro = Tenro
'    Estructuras(Tenro).Estrnro = Sucursal
'    Estructuras(Tenro).Desde = Inicio_Validez
'    Estructuras(Tenro).Hasta = Fin_Validez
        
    'Descripcion de Subdivision de Personal - BTEXT
    Obligatorio = True
    Columna = 23
    Descripcion_Subdivision_Personal = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
'    Texto = "Descripcion de Subdivision de Personal - [ RHPro(Empresa)] " & Descripcion_Subdivision_Personal
        
'hasta la revision del 15/05/2006 estos campos traian la sucursal pero se combino que lo trae el campo 16 (Business Area)
        
    
    'Centro de Coste - [Heidt (Centro de Costo)] - KOSTL
    Obligatorio = True
    Columna = 24
    Centro_de_Costo = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Centro de Coste - [RHPro (Centro de Costo)] " & Centro_de_Costo
    Tablaref = "CCOSTO"
    Tenro = 5
    CCosto = CLng(CalcularMapeoInv(Centro_de_Costo, Tablaref, "-1"))
    If CCosto = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline
        
        'Descripcion de Centro de Costo
        Obligatorio = True
        Columna = 25
        Descripcion_CC = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        Texto = "Descripcion de Centro de Coste - [RHPro (Centro de Costo)] " & Descripcion_CC
        If Not EsNulo(Trim(Descripcion_CC)) Then
            Call ValidaEstructura(Tenro, Descripcion_CC, Nro_Estr, Inserto_estr)
            If Inserto_estr Then
    '            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
    '            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
            End If
            Call Mapear(Tablaref, Centro_de_Costo, CStr(Nro_Estr))
            CCosto = Nro_Estr
        Else
            Flog.writeline Espacios(Tabulador * 4) & "**** Descripción nula, no se puede insertar"
        End If
    Else
        'Descripcion de Centro de Costo
        Obligatorio = True
        Columna = 25
        Descripcion_CC = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    End If
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = CCosto
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
        
        
    'Area de nomina - [RHPro (Forma de Liquidacion)] - ABKRS
    Obligatorio = True
    Columna = 26
    Area_de_Nomina = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Area de nomina - [RHPro (Forma de Liquidacion)] " & Area_de_Nomina
    Tablaref = "FORLIQ"
    Nro_Area_de_Nomina = CLng(CalcularMapeoInv(Area_de_Nomina, Tablaref, "-1"))
    If Nro_Area_de_Nomina = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline Espacios(Tabulador * 4) & "**** No hay Descripción, no se puede insertar"
        Flog.writeline
    End If
    Tenro = 22
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Nro_Area_de_Nomina
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
        
        
    'Relacion Laboral - ANSVH
    Obligatorio = False
    Columna = 27
    Relacion_Laboral = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    'Mapeo Manual tanto para convenio como para Sindicato
    Select Case UCase(Relacion_Laboral)
    Case "01", "1":
        Convenio = 391
        Sindicato = 335
    Case "02", "2":
        Convenio = 391
        Sindicato = 338
    Case "03", "3":
        Convenio = 391
        Sindicato = 339
    Case "04", "4":
        Convenio = 390
        Sindicato = 335
    Case "05", "5":
        Convenio = 390
        Sindicato = 338
    Case "06", "6":
        Convenio = 390
        Sindicato = 339
    Case "07", "7":
        Convenio = 389
        Sindicato = 338
    Case Else
        Convenio = 391
        Sindicato = 335
        Flog.writeline Espacios(Tabulador * 3) & "No se encontró el Convenio/Sindicato " & Relacion_Laboral
        Flog.writeline Espacios(Tabulador * 3) & "Se cargarán defaults 01"
    End Select
    'Convenio
    Tenro = 19
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Convenio
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
    'Sindicato
    Tenro = 16
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Sindicato
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
        
    'Descripcion - ATX
    Obligatorio = True
    Columna = 28
    Descripcion_Relacion_Laboral = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        
    'Posicion - PLANS
    Obligatorio = True
    Columna = 29
    Posicion = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        
    'Descripcion - PLSTX
    Obligatorio = True
    Columna = 30
    Descripcion_Posicion = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        
    'Funcion - STELL
    Obligatorio = True
    Columna = 31
    Funcion = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Funcion - STELL " & Funcion
    Tablaref = "PUESTO"
    Tenro = 4
    Nro_Puesto = CLng(CalcularMapeoInv(Funcion, Tablaref, "-1"))
    If Nro_Puesto = -1 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        
        'Descripcion - STLTX
        Obligatorio = True
        Columna = 32
        Descripcion_Funcion = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        Texto = "Descripcion de Funcion - [RHPro (Puesto)] " & Descripcion_Funcion
        If Not EsNulo(Trim(Descripcion_Funcion)) Then
            Call ValidaEstructura(Tenro, Descripcion_Funcion, Nro_Estr, Inserto_estr)
            If Inserto_estr Then
    '            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
    '            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
            End If
            Call Mapear(Tablaref, Funcion, CStr(Nro_Estr))
            Nro_Puesto = Nro_Estr
        Else
            Flog.writeline Espacios(Tabulador * 4) & "**** Descripción nula, no se puede insertar"
        End If
    Else
        'Descripcion - STLTX
        Obligatorio = True
        Columna = 32
        Descripcion_Funcion = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    End If
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Nro_Puesto
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
        
        
    'Unidad Organizativa - [RHPro (Sector)] - ORGEH
    Obligatorio = True
    Columna = 33
    Unidad_Organizativa = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
'    Texto = "Unidad Organizativa - [RHPro (Sector)] " & Unidad_Organizativa
'    Sector = CLng(CalcularMapeoInv(Unidad_Organizativa, "T527X", "-1"))
'    Tenro = 2
'    Estructuras(Tenro).Tenro = Tenro
'    Estructuras(Tenro).Estrnro = Sector
'    Estructuras(Tenro).Desde = Inicio_Validez
'    Estructuras(Tenro).Hasta = Fin_Validez
'
    'Descripcion - ORGTX
    Obligatorio = True
    Columna = 34
    Descripcion_Unidad_Organizativa = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    
    'Por ahora no hago nada
    
        
    '---------------------------------------
    'Campos de Infotipo 0016 - Contratacion
    '---------------------------------------
    'columnas 34..37
        
    'Tipo de Contrato - [RHPro (Contrato Actual)] - CTTYP
    Obligatorio = True
    Columna = 35
    Contrato = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Tipo de Contrato - [RHPro (Contrato Actual)] " & Contrato
    Tablaref = "CONTRA"
    Tenro = 18
    Nro_Contrato = CLng(CalcularMapeoInv(Contrato, Tablaref, "-1"))
    If Nro_Contrato = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline
        
        'Descripcion - CTTXT
        Obligatorio = True
        Columna = 36
        Descripcion_Contrato = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        Texto = "Descripcion de Tipo de Contrato - [RHPro (Contrato Actual)] " & Descripcion_Contrato
        If Not EsNulo(Trim(Descripcion_Contrato)) Then
            Call ValidaEstructura(Tenro, Descripcion_Contrato, Nro_Estr, Inserto_estr)
            If Inserto_estr Then
    '            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
    '            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
            End If
            Call Mapear(Tablaref, Contrato, CStr(Nro_Estr))
            Nro_Contrato = Nro_Estr
        Else
            Flog.writeline Espacios(Tabulador * 4) & "**** Descripción nula, no se puede insertar"
        End If
    Else
        'Descripcion - CTTXT
        Obligatorio = True
        Columna = 36
        Descripcion_Contrato = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    End If
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Nro_Contrato
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
        
    'Fecha fin de contrato
    Obligatorio = False
    Columna = 37
    Fecha_Fin_Contrato = StrToDate(Trim(Lista(Columna)), OK, FormatoFechaSap1, NuloFechaSap1)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Fecha de fin de Contrato " & Trim(Lista(Columna))
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 3) & "Valor en Dato incorrecto"
        Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
        If Obligatorio Then Exit Sub
    End If
    Estructuras(Tenro).Hasta = Fecha_Fin_Contrato
    
    'Periodo de Prueba # - PRBZT
    Obligatorio = True
    Columna = 38
    Periodo_Prueba = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        
    'Periodo de Prueba # - PRBEH
    Obligatorio = True
    Columna = 39
    Periodo_Prueba_UN = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        
        
        
    '---------------------------------------
    'Campos de Infotipo 0007 - Horarios de trabajo
    '---------------------------------------
    'columnas 38..40
    'Regla plan horario de trabajo - SCHKN
    Obligatorio = True
    Columna = 40
    Plan_Horario_Trabajo = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Regla plan horario de trabajo - [RHPro (Regimen Horario)] " & Plan_Horario_Trabajo
    Tablaref = "HORAT"
    Tenro = 21
    Regimen_Horario = CLng(CalcularMapeoInv(Plan_Horario_Trabajo, Tablaref, "-1"))
    If Regimen_Horario = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline
        
        'Descripcion - PRBEH
        Obligatorio = True
        Columna = 41
        Descripcion_Horario_Trabajo = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        Texto = "Descripcion de Horario de trabajo - [RHPro (Regimen Horario)] " & Descripcion_Horario_Trabajo
        If Not EsNulo(Trim(Descripcion_Horario_Trabajo)) Then
            Call ValidaEstructura(Tenro, Descripcion_Horario_Trabajo, Nro_Estr, Inserto_estr)
            If Inserto_estr Then
    '            Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
    '            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
            End If
            Call Mapear(Tablaref, Plan_Horario_Trabajo, CStr(Nro_Estr))
            Regimen_Horario = Nro_Estr
        Else
            Flog.writeline Espacios(Tabulador * 4) & "**** Descripción nula, no se puede insertar"
        End If
    Else
        'Descripcion - PRBEH
        Obligatorio = True
        Columna = 41
        Descripcion_Horario_Trabajo = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    End If
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Regimen_Horario
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
    
    'Porcentaje de horario de trabajo - EMPCT
    Obligatorio = True
    Columna = 42
    Porcentaje_Horario_Trabajo = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    
    
    'Obra Social
    Obligatorio = True
    Columna = 54
    Obra_Social = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Obra Social - [RHPro (Obra Social)] " & Obra_Social
    Tablaref = "OSOCIA"
    Tenro = 17
    Nro_Obra_Social = CLng(CalcularMapeoInv(Obra_Social, Tablaref, "-1"))
    If Nro_Obra_Social = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline
        
        'Descripcion O. Social - PLANS
        Obligatorio = False
        Columna = 55
        Descripcion_OSocial = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        Texto = "Descripcion de OS - [RHPro (Obra Social)] " & Descripcion_OSocial
        If Not EsNulo(Trim(Descripcion_OSocial)) Then
            Call ValidaEstructura(Tenro, Descripcion_OSocial, Nro_Estr, Inserto_estr)
            If Inserto_estr Then
                Call CreaTercero(4, Descripcion, Nro_Tercero)
                Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
            End If
            Call Mapear(Tablaref, Obra_Social, CStr(Nro_Estr))
            Nro_Obra_Social = Nro_Estr
        Else
            Flog.writeline Espacios(Tabulador * 4) & "**** Descripción nula, no se puede insertar"
        End If
    Else
        'Descripcion O. Social - PLANS
        Obligatorio = False
        Columna = 55
        Descripcion_OSocial = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    End If
    'Elegida - OJO para DTT es por ley
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Nro_Obra_Social
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
    
'    'Por ley - para DTT es elegida y no se utiliza
'    Tenro = 24
'    Estructuras(Tenro).Tenro = Tenro
'    Estructuras(Tenro).Estrnro = Nro_Obra_Social
'    Estructuras(Tenro).Desde = Inicio_Validez
'    Estructuras(Tenro).Hasta = Fin_Validez
    
    'Plan de Obra Social
    Obligatorio = True
    Columna = 56
    Plan_Obra_Social = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Plan de Obra Social - [RHPro (Obra Social)] " & Obra_Social
    Tablaref = "PLANOS"
    Tenro = 23
    Nro_Plan_Obra_Social = CLng(CalcularMapeoInv(Plan_Obra_Social, Tablaref, "-1"))
    If Nro_Plan_Obra_Social = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline Espacios(Tabulador * 4) & "**** No hay Descripción, no se puede insertar"
        Flog.writeline
    End If
   'Elejida - OJO para DTT por ley
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Nro_Plan_Obra_Social
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
    
    
'    'Funcion - STELL
'    Obligatorio = True
'    Columna = 30
'    Funcion = Trim(Lista(Columna))
'    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
'    Texto = "Funcion - STELL " & Funcion
'    Nro_Puesto = CLng(CalcularMapeoInv(Funcion, "PUESTO", "-1"))
'    Tenro = 4
'    Estructuras(Tenro).Tenro = Tenro
'    Estructuras(Tenro).Estrnro = Nro_Puesto
'    Estructuras(Tenro).Desde = Inicio_Validez
'    Estructuras(Tenro).Hasta = Fin_Validez
    
    'AFJP - Capitalizacion
    Obligatorio = False
    Columna = 57
    Capitalizacion = IIf(Not EsNulo(Trim(Lista(Columna))), True, False)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "AFJP - Capitalizacion " & Capitalizacion
    If Capitalizacion Then
        'debo buscar alguna estructura de caja de jubilacion tal que la descripcion sea Capitalizacion y ademas el tipo de caja se 2
        StrSql = "SELECT estructura.estrnro FROM estructura "
        StrSql = StrSql & " INNER JOIN cajjub ON estructura.estrnro = cajjub.estrnro AND cajjub.ticnro = 2"
        StrSql = StrSql & " WHERE upper(estructura.estrdabr) = 'CAPITALIZACION'"
        StrSql = StrSql & " AND estructura.tenro = 15 "
        OpenRecordset StrSql, rs_caja
        If Not rs_caja.EOF Then
            Nro_caja = rs_caja!Estrnro
            Tenro = 15
            Estructuras(Tenro).Tenro = Tenro
            Estructuras(Tenro).Estrnro = Nro_caja
            Estructuras(Tenro).Desde = Inicio_Validez
            Estructuras(Tenro).Hasta = Fin_Validez
        Else
            Flog.writeline Espacios(Tabulador * 3) & "No Hay ninguna Caja de Jubilacion de capitalización cargada  en RHPRO."
        End If
    End If
    
    'AFJP - Reparto
    Obligatorio = False
    Columna = 58
    Reparto = IIf(Not EsNulo(Trim(Lista(Columna))), True, False)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "AFJP - Reparto " & Reparto
    If Reparto Then
        'debo buscar alguna estructura de caja de jubilacion tal que la descripcion sea Capitalizacion y ademas el tipo de caja se 2
        StrSql = "SELECT estructura.estrnro FROM estructura "
        StrSql = StrSql & " INNER JOIN cajjub ON estructura.estrnro = cajjub.estrnro AND cajjub.ticnro = 1"
        StrSql = StrSql & " WHERE upper(estructura.estrdabr) = 'REPARTO'"
        StrSql = StrSql & " AND estructura.tenro = 15 "
        OpenRecordset StrSql, rs_caja
        If Not rs_caja.EOF Then
            Nro_caja = rs_caja!Estrnro
            Tenro = 15
            Estructuras(Tenro).Tenro = Tenro
            Estructuras(Tenro).Estrnro = Nro_caja
            Estructuras(Tenro).Desde = Inicio_Validez
            Estructuras(Tenro).Hasta = Fin_Validez
        Else
            Flog.writeline Espacios(Tabulador * 3) & "No Hay ninguna Caja de Jubilacion de Reparto cargada  en RHPRO."
        End If
    End If

    'Actividad SIJP
    Obligatorio = False
    Columna = 59
    Actividad_SIJP = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Actividad SIJP - [RHPro (Actividad)] " & Actividad_SIJP
    Tablaref = "ASIJP"
    Tenro = 29
    Nro_Actividad_SIJP = CLng(CalcularMapeoInv(Actividad_SIJP, Tablaref, "-1"))
    If Nro_Actividad_SIJP = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline Espacios(Tabulador * 4) & "**** No hay Descripción, no se puede insertar"
        Flog.writeline
    End If
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Nro_Actividad_SIJP
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
    
    
    'Grupo Empleado SIJP
    Obligatorio = False
    Columna = 60
    Grupo_SIJP = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Texto = "Grupo empleado SIJP - [RHPro (Condicion SIJP)] " & Grupo_SIJP
    Tablaref = "CSIJP"
    Tenro = 31
    Nro_Grupo_SIJP = CLng(CalcularMapeoInv(Grupo_SIJP, "CSIJP", "-1"))
    If Nro_Grupo_SIJP = -1 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline
        
        'Descripcion
        Obligatorio = False
        Columna = 61
        Descripcion_Grupo_SIJP = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        Texto = "Descripcion Grupo SIJP - [RHPro (Condicion SIJP)] " & Descripcion_Grupo_SIJP
        If Not EsNulo(Trim(Descripcion_OSocial)) Then
            Call ValidaEstructura(Tenro, Descripcion_Grupo_SIJP, Nro_Estr, Inserto_estr)
            If Inserto_estr Then
'                Call CreaTercero(4, Descripcion_Grupo_SIJP, Nro_Tercero)
'                Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
            End If
            Call Mapear(Tablaref, Grupo_SIJP, CStr(Nro_Estr))
            Nro_Obra_Social = Nro_Estr
        Else
            Flog.writeline Espacios(Tabulador * 4) & "**** Descripción nula, no se puede insertar"
        End If
    Else
        'Descripcion
        Obligatorio = False
        Columna = 61
        Descripcion_Grupo_SIJP = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    End If
    Estructuras(Tenro).Tenro = Tenro
    Estructuras(Tenro).Estrnro = Nro_Grupo_SIJP
    Estructuras(Tenro).Desde = Inicio_Validez
    Estructuras(Tenro).Hasta = Fin_Validez
 
 
    
    
    
    '---------------------------------------------------------------
    'Estructuras por default
    'Situacion de Revista
    'Se carga Manualmente desde RHPRO
    
    'Siniestro SIJP
    'Se carga Manualmente desde RHPRO
    
    
'    '*************************************************************
'    'Por ahora lo inactivo hasta que nos pasen el mapeo manual
'    'Categoria - Funcion + Work Contract
'    Categoria = Trim(Lista(31))
'    'Mapeo Manual
'    Select Case UCase(Categoria)
'    Case "01":
'        Nro_Categoria = 1
'    Case Else
'        Nro_Categoria = 1
'        Flog.writeline Espacios(Tabulador * 3) & "No se encontró el Categoria " & Relacion_Laboral
'        Flog.writeline Espacios(Tabulador * 3) & "Se cargará default 01"
'    End Select
'    Tenro = 3
'    Estructuras(Tenro).Tenro = Tenro
'    Estructuras(Tenro).Estrnro = Nro_Categoria
'    Estructuras(Tenro).Desde = Inicio_Validez
'    Estructuras(Tenro).Hasta = Fin_Validez
'    '*************************************************************
    
    
    '---------------------------------------------------------------
    
    
    '--------------------------------------------------------------------------
    'Inserto/modifico las estructuras
    For i = 1 To UBound(Estructuras)
        If Estructuras(i).Estrnro <> 0 And Estructuras(i).Estrnro <> -1 Then
            Call Insertar_His_Estructura(Estructuras(i).Tenro, Estructuras(i).Estrnro, Empleado.Tercero, Estructuras(i).Desde, Estructuras(i).Hasta)
        End If
    Next i
    'Inserto/modifico las estructuras
    '--------------------------------------------------------------------------

    '---------------------------------------
    'Campos de Infotipo 0009 - Banco
    '---------------------------------------
    'columnas 41..48
    
    'Clase de Datos Bancarios - SUBTY
    Obligatorio = True
    Columna = 43
    If Columna <= UBound(Lista) Then
        Clase_Datos_Bancarios = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
        
    'Descripcion - STEXT
    Obligatorio = True
    Columna = 44
    If Columna <= UBound(Lista) Then
        Descripcion_Clase_banco = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
    
    
    'Clave de banco - BANKL
    Obligatorio = True
    Columna = 45
    Clave_de_Banco = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Tablaref = "BANCOS"
    Tenro = 41
    Nro_Banco = CLng(CalcularMapeoInv(Clave_de_Banco, Tablaref, "0"))
    If Nro_Banco = 0 Then
        Flog.writeline Espacios(Tabulador * 3) & "Inexistente - No se encuentra " & Texto
        Flog.writeline
        
        'Banco - BANKA
        Obligatorio = True
        Columna = 46
        Banco = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        Texto = "Descripcion Banco - [RHPro (Banco)] " & Banco
        If Not EsNulo(Trim(Banco)) Then
            Call ValidaEstructura(Tenro, Banco, Nro_Estr, Inserto_estr)
            If Inserto_estr Then
                Call CreaTercero(13, Format_Str(Banco, 40, False, ""), Nro_Tercero)
                Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Banco)
            End If
            Call Mapear(Tablaref, Clave_de_Banco, CStr(Nro_Estr))
            Nro_Banco = Nro_Estr
        Else
            Flog.writeline Espacios(Tabulador * 4) & "**** Descripción nula, no se puede insertar"
        End If
    Else
        'Banco - BANKA
        Obligatorio = True
        Columna = 46
        Banco = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    End If
    
    'Cuenta Bancaria - BANKN
    Obligatorio = True
    Columna = 47
    Cuenta_Bancaria = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    
    'Clave de control de bancos - BKONT
    Obligatorio = True
    Columna = 48
    Clave_Control_Banco = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    
    'Via de pago - ZLSCH
    Obligatorio = True
    Columna = 49
    Via_de_Pago = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Nro_FormaPago = CLng(CalcularMapeoInv(Via_de_Pago, "FORMAP", "0"))

    'Sucursal - ZWECK - (en la de Roche esto era primera parte del CBU)
    Obligatorio = True
    Columna = 50
    Destino_para_Transferencias = Trim(Lista(Columna))
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Banco_Sucursal = Trim(Left(Destino_para_Transferencias, 40))
    
    CBU = "0"
    Porcentaje_Prefijado = 100
    
    'Revisar si existe el banco
    'Inserto la cuenta bancaria
    If (Nro_FormaPago <> 0 And Nro_Banco <> 0 And Cuenta_Bancaria <> "") Then
    
        'primero deberia buscar si existe
        StrSql = "SELECT * FROM ctabancaria "
        StrSql = StrSql & " WHERE ternro = " & Empleado.Tercero
        'StrSql = StrSql & " AND fpagnro = " & Nro_FormaPago
        'StrSql = StrSql & " AND banco = " & Nro_Banco
        StrSql = StrSql & " AND ctabestado = -1 "
        OpenRecordset StrSql, rs_Cta
        If rs_Cta.EOF Then
            StrSql = " INSERT INTO ctabancaria (ternro,fpagnro,banco,ctabestado,ctabnro,ctabsucdesc"
            If Not EsNulo(CBU) Then
                StrSql = StrSql & ",ctabcbu"
            End If
            StrSql = StrSql & ", ctabporc"
            StrSql = StrSql & " ) VALUES ("
            StrSql = StrSql & Empleado.Tercero
            StrSql = StrSql & "," & Nro_FormaPago
            StrSql = StrSql & "," & Nro_Banco
            StrSql = StrSql & "," & "-1"
            StrSql = StrSql & ",'" & Replace(Cuenta_Bancaria, "-", "") & "'"
            StrSql = StrSql & ",'" & Banco_Sucursal & "'"
            If Not EsNulo(CBU) Then
                StrSql = StrSql & ",'" & Replace(CBU, "-", "") & "'"
            End If
            StrSql = StrSql & "," & Porcentaje_Prefijado
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            If rs_Cta!ctabnro = Cuenta_Bancaria Then
                StrSql = " UPDATE ctabancaria SET "
                StrSql = StrSql & " ctabporc = " & Porcentaje_Prefijado
                StrSql = StrSql & " ,ctabsucdesc = '" & Banco_Sucursal & "'"
                If Not EsNulo(CBU) Then
                    If EsNulo(rs_Cta!ctabcbu) Then
                        StrSql = StrSql & " ,ctabcbu = '" & Replace(CBU, "-", "") & "'"
                    End If
                End If
                StrSql = StrSql & " WHERE cbnro = " & rs_Cta!cbnro
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                'Desactivo la anterior
                StrSql = " UPDATE ctabancaria SET "
                StrSql = StrSql & " ctabestado = 0 "
                StrSql = StrSql & " WHERE cbnro = " & rs_Cta!cbnro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'inserto la nueva
                StrSql = " INSERT INTO ctabancaria (ternro,fpagnro,banco,ctabestado,ctabnro,ctabsucdesc"
                If Not EsNulo(CBU) Then
                    StrSql = StrSql & ",ctabcbu"
                End If
                StrSql = StrSql & ", ctabporc"
                StrSql = StrSql & " ) VALUES ("
                StrSql = StrSql & Empleado.Tercero
                StrSql = StrSql & "," & Nro_FormaPago
                StrSql = StrSql & "," & Nro_Banco
                StrSql = StrSql & "," & "-1"
                StrSql = StrSql & ",'" & Replace(Cuenta_Bancaria, "-", "") & "'"
                StrSql = StrSql & ",'" & Banco_Sucursal & "'"
                If Not EsNulo(CBU) Then
                    StrSql = StrSql & ",'" & Replace(CBU, "-", "") & "'"
                End If
                StrSql = StrSql & "," & Porcentaje_Prefijado
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    End If
    
    'Fecha Medida
    Obligatorio = False
    Columna = 51
    Texto = "Fecha de Medida"
    If Columna <= UBound(Lista) Then
        Fecha_Medida = StrToDate(Trim(Lista(Columna)), OK, FormatoFechaSap1, NuloFechaSap1)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        If Not OK Then
            Flog.writeline Espacios(Tabulador * 3) & "Valor en Dato incorrecto"
            Flog.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Texto & " inválido " & Mid(strlinea, pos1, pos2)
            If Obligatorio Then Exit Sub
        End If
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
    End If
   
    'Tipo de Servicio - type of service
    Obligatorio = False
    Columna = 62
    If Columna <= UBound(Lista) Then
        Tipo_Servicio = Trim(Lista(Columna))
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        'Nro_FormaPago = CLng(CalcularMapeoInv(Via_de_Pago, "T042Z", "0"))
    Else
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, "")
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
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
End Sub



Public Sub Insertar_Linea_Infotipo(ByVal strlinea As String, ByVal UltimaLinea As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Salrio Basico, Adicionales, presentismo y ausentismos
' Autor      : FGZ
' Fecha      : 09/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'CAMPO   TIPO DE DATO    LONGITUD    DESCRIPCION Código Tabla    Nombre Técnico
' ---------------------------------------------------------------------------------------------
Dim pos1
Dim pos2
Dim aux
Dim OK As Boolean
Dim Columna As Byte

Dim Inicio_Validez
Dim Fin_Validez


Dim Descripcion As String
Dim valor As String
Dim Nombre_Tecnico As String
Dim Tipo_Campo As String
Dim Long_Campo As String
Dim Aux_Legajo As String

Dim Infotipo_Invalido As Boolean
Dim Lista
Dim i As Integer
Dim Hoja As Integer

    On Error GoTo Manejador_De_Error
    
    'Leo todos los campos
    Lista = Split(strlinea, Separador)
    
    
    Inicio_Validez = C_Date(Day(Date) & "/" & Month(Date) & "/" & Year(Date))
'    If Month(Date) = 12 Then
'        Fin_Validez = C_Date("31/12/" & Year(Date))
'    Else
'        Fin_Validez = C_Date("01/" & Month(Date) + 1 & "/" & Year(Date)) - 1
'    End If
        
    
    
    On Error GoTo Manejador_De_Error
    Columna = 0
    Hoja = 2
    Fila_Infotipo = Fila_Infotipo + 1
    
    
    'Nro de Legajo
    Columna = 0
    Texto = "Legajo"
    If IsNumeric(Lista(Columna)) Then
        Aux_Legajo = Lista(Columna)
        Empleado.Legajo = BuscarLegajo2(Aux_Legajo)
        Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
        If Empleado.Legajo = "-1" Then
            ExisteLegajo = False
            Empleado.Tercero = 0
        Else
            If UltimoLegajo <> Empleado.Legajo Then
                IndiceArr = 1
                UltimoInfotipo = ""
                Infotipo_Campo = 1
    
                If UltimoLegajo <> -1 Then
                    Flog.writeline Espacios(Tabulador * 1) & "Actualizo infotipos del legajo " & Empleado.Legajo
                    Call Actualizar_Infotipos(Inicio_Validez, Fin_Validez)  'UltimoLegajo
                End If
                
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 1) & "Legajo " & Empleado.Legajo
                
                'Inicializo todas las variables globales que interesan al proceso
                Call Inicializar_Globales   'limpio array de datos
                
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
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El legajo no es numerico"
        FlogE.writeline Espacios(Tabulador * 1) & "Linea " & NroLinea & ": El legajo no es numerico"
        
        Infotipo_Invalido = True
        Exit Sub
    End If
    
    'Infotipos
    Columna = 1
    Texto = "Infotipo"
    Infotipo = Lista(Columna)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Flog.writeline Espacios(Tabulador * 2) & "Infotipo " & Infotipo
    If Infotipo = "IT0394" Then
        Flog.writeline Espacios(Tabulador * 3) & "*** Lo tomo como IT0021 ***"
        Infotipo = "IT0021"
    End If
    'Descripcion
    Columna = 2
    Texto = "Descripcion"
    aux = Lista(Columna)
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, Lista(Columna))
    Descripcion = Trim(aux)
    
    'Importe
    Columna = 3
    Texto = "Importe"
    aux = Lista(Columna)
    valor = Trim(aux)
    
    'Auxiliar
    Columna = 4
    Texto = "Decimales Importe"
    If UBound(Lista) >= 4 Then
        aux = Lista(Columna)
        If SeparadorDecimal = "," Then
            valor = Replace(valor, ".", "") & "." & Trim(aux)
        Else
            valor = Replace(valor, ",", "") & "." & Trim(aux)
        End If
    End If
    Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, valor)
    
    
'    'Nombre Tecnico
'    Columna = 4
'    Texto = "Nombre Tecnico"
'    Aux = Lista(Columna)
'    'Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, aux)
'    Nombre_Tecnico = Trim(Aux)
    
'    'Longitud
'    Columna = 5
'    Texto = "Long del campo"
'    Aux = Lista(Columna)
'    'Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, aux)
'    Long_Campo = Lista(Columna)
'
'    'Tipo de campo
'    Columna = 6
'    Texto = "Tipo de campo"
'    Aux = Lista(Columna)
'    'Call Insertar_Valor_Excel(Hoja, Fila_Infotipo, Columna, aux)
'    Tipo_Campo = Lista(Columna)


    'llevo la cuenta del nro de registro o campo del infotipo para el legajo
    'If UltimoInfotipo <> Infotipo And Not (Infotipo = 394 And UltimoInfotipo = 21) Then
    If UltimoInfotipo <> Infotipo Then
        If Not EsNulo(UltimoInfotipo) Then
            Flog.writeline Espacios(Tabulador * 1) & "Actualizo infotipos del legajo " & Empleado.Legajo
            Call Actualizar_Infotipos(Inicio_Validez, Fin_Validez)  'UltimoLegajo
            IndiceArr = 1
            Infotipo_Campo = 1
            Infotipo_Invalido = False
        End If
        UltimoInfotipo = Infotipo
    Else
        If Primer_Campo_Infotipo(Infotipo, Infotipo_Campo) Then
            Flog.writeline Espacios(Tabulador * 1) & "Actualizo infotipos del legajo " & Empleado.Legajo
            Call Actualizar_Infotipos(Inicio_Validez, Fin_Validez)  'UltimoLegajo
            
            Infotipo_Campo = 1
            Infotipo_Invalido = False
            IndiceArr = 1
        End If
    End If

    'Si es infotipo es IT2001 O IT2002
    If Infotipo_Campo = 3 And (UCase(Infotipo) = "IT2001" Or UCase(Infotipo) = "IT2002") Then
        'Valor = CalcularMapeoSubtipo("IT2001", Valor, "TLIC", "0")
        valor = CalcularMapeoInv(valor, "TLIC", "0")
    End If
    
    If Infotipo_Campo = 3 And UCase(Infotipo) = "IT0021" Then
        'Valor = CLng(CalcularMapeoSubtipo("IT0021", Valor, "T591A", "0"))
        'Valor = CalcularMapeoInv(Valor, "T591A", "0")
        valor = CalcularMapeoInv(valor, "PARENT", "0")
    End If
    
    
    If Not Infotipo_Invalido Then
        Arr_Datos(Infotipo_Campo).ID_campo = Infotipo_Campo
        'Arr_Datos(Infotipo_Campo).Campo = Nombre_Tecnico
        Arr_Datos(Infotipo_Campo).Descripcion = Descripcion
        Arr_Datos(Infotipo_Campo).valor = valor
        Arr_Datos(Infotipo_Campo).TipoDato = Tipo_Campo
        Arr_Datos(Infotipo_Campo).IT = Infotipo
    Else
        Arr_Datos(Infotipo_Campo).ID_campo = Infotipo_Campo
        'Arr_Datos(Infotipo_Campo).Campo = Nombre_Tecnico
        Arr_Datos(Infotipo_Campo).Descripcion = Descripcion
        Arr_Datos(Infotipo_Campo).valor = "ERROR"
        Arr_Datos(Infotipo_Campo).TipoDato = Tipo_Campo
        Arr_Datos(Infotipo_Campo).IT = Infotipo
    End If
    
    Infotipo_Campo = Infotipo_Campo + 1
    
    If UltimaLinea Then
        Call Actualizar_Infotipos(Inicio_Validez, Fin_Validez)  'UltimoLegajo
    End If
    

Exit Sub
Manejador_De_Error:
    Infotipo_Invalido = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error " 'en infotipo " & Infotipo
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


Public Function Indice(ByVal valor) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que devuelve el indice en el arr_Datos al que corresponde
' Autor      : FGZ
' Fecha      : 27/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Posicion As Integer

Select Case UCase(valor)
'Infotipos
    Case "IT0000":
        Posicion = 0
    Case "IT0001":
        Posicion = 1
    Case "IT0002":
        Posicion = 2
    Case "IT0007":
        Posicion = 3
    Case "IT0008":
        Posicion = 4
    Case "IT0009":
        Posicion = 5
    Case "IT0014":
        Posicion = 6
    Case "IT0015":
        Posicion = 7
    Case "IT0016":
        Posicion = 8
    Case "IT0021":
        Posicion = 9
    Case "IT0185":
        Posicion = 10
    Case "IT0392":
        Posicion = 11
    Case "IT2001":
        Posicion = 12
    Case "IT2002":
        Posicion = 13
'Campos
'IT0008
    Case "IT0008PERNR":
        Posicion = 1
    Case "IT0008INFTY":
        Posicion = 2
    Case "IT0008SUBTY":
        Posicion = 3
    Case "IT0008ENDDA":
        Posicion = 5
    Case "IT0008BEGDA":
        Posicion = 6
    Case "IT0008TRFAR":
        Posicion = 7
    Case "IT0008TRFGB":
        Posicion = 8
    Case "IT0008TRFGR":
        Posicion = 9
    Case "IT0008TRFST":
        Posicion = 10
    Case "IT0008LGART":
        Posicion = 11
    Case "IT0008BETRG":
        Posicion = 12
    Case "IT0008LGTXT":
        Posicion = 13
'IT0014
    Case "IT0014PERNR":
        Posicion = 1
    Case "IT0014INFTY":
        Posicion = 2
    Case "IT0014SUBTY":
        Posicion = 3
    Case "IT0014ENDDA":
        Posicion = 4
    Case "IT0014BEGDA":
        Posicion = 5
    Case "IT0014LGART":
        Posicion = 6
    Case "IT0014LGTXT":
        Posicion = 7
    Case "IT0014BETRG":
        Posicion = 8
    Case "IT0014ANZHL":
        Posicion = 9
    Case "IT0014EITXT":
        Posicion = 10
'IT0015
    Case "IT0015PERNR":
        Posicion = 1
    Case "IT0015INFTY":
        Posicion = 2
    Case "IT0015SUBTY":
        Posicion = 3
    Case "IT0015BEGDA":
        Posicion = 4
    Case "IT0015LGART":
        Posicion = 5
    Case "IT0015BETRG":
        Posicion = 6
    Case "IT0015ANZHL":
        Posicion = 7
    Case "IT0015EITXT":
        Posicion = 8
    Case "IT0015LGTXT":
        Posicion = 9
'IT0021
    Case "IT0021PERNR":
        Posicion = 1
    Case "IT0021INFTY":
        Posicion = 2
    Case "IT0021SUBTY":
        Posicion = 3
    Case "IT0021ENDDA":
        Posicion = 4
    Case "IT0021BEGDA":
        Posicion = 5
    Case "IT0021FAMSA":
        Posicion = 6
    Case "IT0021FANAM":
        Posicion = 7
    Case "IT0021FNAC2":
        Posicion = 8
    Case "IT0021FAVOR":
        Posicion = 9
    Case "IT0021GESCC1 / 2":
        Posicion = 10
    Case "IT0021ASFAX":
        Posicion = 11
    Case "IT0021FAMST":
        Posicion = 12
    Case "IT0021ICTYP":
        Posicion = 13
    Case "IT0021ICTXT":
        Posicion = 14
    Case "IT0021ICNUM":
        Posicion = 15
    Case "IT0021ADHOS":
        Posicion = 16
    Case "IT0021STEXT":
        Posicion = 17
    Case "IT0021OBJPS":
        Posicion = 18
    Case "IT0021FGBDT":
        Posicion = 19
    Case "IT0021DISCP":
        Posicion = 20
'IT0185
    Case "IT0185PERNR":
        Posicion = 1
    Case "IT0185INFTY":
        Posicion = 2
    Case "IT0185SUBTY":
        Posicion = 3
    Case "IT0185ENDDA":
        Posicion = 4
    Case "IT0185BEGDA":
        Posicion = 5
    Case "IT0185ICTYP":
        Posicion = 6
    Case "IT0185ICTXT":
        Posicion = 8
    Case "IT0185ICNUM":
        Posicion = 7
    Case "IT0185ISCOT":
        Posicion = 9
'IT0392
    Case "IT0392PERNR":
        Posicion = 1
    Case "IT0392INFTY":
        Posicion = 2
    Case "IT0392SUBTY":
        Posicion = 3
    Case "IT0392ENDDA":
        Posicion = 4
    Case "IT0392BEGDA":
        Posicion = 5
    Case "IT0392OBRAS":
        Posicion = 6
    Case "IT0392SYJU1":
        Posicion = 7
    Case "IT0392SYJU2":
        Posicion = 8
    Case "IT0392TYACT":
        Posicion = 9
    Case "IT0392ASPCE":
        Posicion = 10
    Case "IT0392TASPC":
        Posicion = 11
    Case "IT0392CSERV":
        Posicion = 12
'IT2001
    Case "IT2001PERNR":
        Posicion = 1
    Case "IT2001INFTY":
        Posicion = 2
    Case "IT2001SUBTY":
        Posicion = 3
    Case "IT2001ENDDA":
        Posicion = 4
    Case "IT2001BEGDA":
        Posicion = 5
    Case "IT2001ATEXT":
        Posicion = 6
    Case "IT2001AWART":
        Posicion = 7
    Case "IT2001ABWTG":
        Posicion = 8
'IT2002
    Case "IT2002PERNR":
        Posicion = 1
    Case "IT2002INFTY":
        Posicion = 2
    Case "IT2002SUBTY":
        Posicion = 3
    Case "IT2002ENDDA":
        Posicion = 4
    Case "IT2002BEGDA":
        Posicion = 5
    Case "IT2002ATEXT":
        Posicion = 6
    Case "IT2002AWART":
        Posicion = 7
    Case "IT2002ABWTG":
        Posicion = 8
    Case Else
        Posicion = -1
End Select
Indice = Posicion
End Function



Public Function Indice2(ByVal valor) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que devuelve el indice en el arr_Datos al que corresponde
' Autor      : FGZ
' Fecha      : 13/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Posicion As Integer

Select Case UCase(valor)
'Infotipos
    Case "IT0000":
        Posicion = 0
    Case "IT0001":
        Posicion = 1
    Case "IT0002":
        Posicion = 2
    Case "IT0007":
        Posicion = 3
    Case "IT0008":
        Posicion = 4
    Case "IT0009":
        Posicion = 5
    Case "IT0014":
        Posicion = 6
    Case "IT0015":
        Posicion = 7
    Case "IT0016":
        Posicion = 8
    Case "IT0021":
        Posicion = 9
    Case "IT0185":
        Posicion = 10
    Case "IT0392":
        Posicion = 11
    Case "IT2001":
        Posicion = 12
    Case "IT2002":
        Posicion = 13
'Campos
'IT0008
    Case "IT0008PERNR":
        Posicion = 1
    Case "IT0008INFTY":
        Posicion = 2
    Case "IT0008SUBTY":
        Posicion = 3
    Case "IT0008ENDDA":
        Posicion = 5
    Case "IT0008BEGDA":
        Posicion = 6
    Case "IT0008TRFAR":
        Posicion = 7
    Case "IT0008TRFGB":
        Posicion = 8
    Case "IT0008TRFGR":
        Posicion = 9
    Case "IT0008TRFST":
        Posicion = 10
    Case "IT0008LGART":
        Posicion = 11
    Case "IT0008BETRG":
        Posicion = 12
    Case "IT0008LGTXT":
        Posicion = 13
'IT0014
    Case "IT0014PERNR":
        Posicion = 1
    Case "IT0014INFTY":
        Posicion = 2
    Case "IT0014SUBTY":
        Posicion = 3
    Case "IT0014ENDDA":
        Posicion = 4
    Case "IT0014BEGDA":
        Posicion = 5
    Case "IT0014LGART":
        Posicion = 6
    Case "IT0014LGTXT":
        Posicion = 7
    Case "IT0014BETRG":
        Posicion = 8
    Case "IT0014ANZHL":
        Posicion = 9
    Case "IT0014EITXT":
        Posicion = 10
'IT0015
    Case "IT0015PERNR":
        Posicion = 1
    Case "IT0015INFTY":
        Posicion = 2
    Case "IT0015SUBTY":
        Posicion = 3
    Case "IT0015BEGDA":
        Posicion = 4
    Case "IT0015LGART":
        Posicion = 5
    Case "IT0015BETRG":
        Posicion = 6
    Case "IT0015ANZHL":
        Posicion = 7
    Case "IT0015EITXT":
        Posicion = 8
    Case "IT0015LGTXT":
        Posicion = 9
'IT0021
    Case "IT0021PERNR":
        Posicion = 1
    Case "IT0021INFTY":
        Posicion = 2
    Case "IT0021SUBTY":
        Posicion = 3
    Case "IT0021ENDDA":
        Posicion = 4
    Case "IT0021BEGDA":
        Posicion = 5
    Case "IT0021FAMSA":
        Posicion = 6
    Case "IT0021FANAM":
        Posicion = 7
    Case "IT0021FNAC2":
        Posicion = 8
    Case "IT0021FAVOR":
        Posicion = 9
    Case "IT0021GESCC1 / 2":
        Posicion = 10
    Case "IT0021ASFAX":
        Posicion = 11
    Case "IT0021FAMST":
        Posicion = 12
    Case "IT0021ICTYP":
        Posicion = 13
    Case "IT0021ICTXT":
        Posicion = 14
    Case "IT0021ICNUM":
        Posicion = 15
    Case "IT0021ADHOS":
        Posicion = 16
    Case "IT0021STEXT":
        Posicion = 17
    Case "IT0021OBJPS":
        Posicion = 18
    Case "IT0021FGBDT":
        Posicion = 19
    Case "IT0021DISCP":
        Posicion = 20
'IT0185
    Case "IT0185PERNR":
        Posicion = 1
    Case "IT0185INFTY":
        Posicion = 2
    Case "IT0185SUBTY":
        Posicion = 3
    Case "IT0185ENDDA":
        Posicion = 4
    Case "IT0185BEGDA":
        Posicion = 5
    Case "IT0185ICTYP":
        Posicion = 6
    Case "IT0185ICTXT":
        Posicion = 8
    Case "IT0185ICNUM":
        Posicion = 7
    Case "IT0185ISCOT":
        Posicion = 9
'IT0392
    Case "IT0392PERNR":
        Posicion = 1
    Case "IT0392INFTY":
        Posicion = 2
    Case "IT0392SUBTY":
        Posicion = 3
    Case "IT0392ENDDA":
        Posicion = 4
    Case "IT0392BEGDA":
        Posicion = 5
    Case "IT0392OBRAS":
        Posicion = 6
    Case "IT0392SYJU1":
        Posicion = 7
    Case "IT0392SYJU2":
        Posicion = 8
    Case "IT0392TYACT":
        Posicion = 9
    Case "IT0392ASPCE":
        Posicion = 10
    Case "IT0392TASPC":
        Posicion = 11
    Case "IT0392CSERV":
        Posicion = 12
'IT2001
    Case "IT2001PERNR":
        Posicion = 1
    Case "IT2001INFTY":
        Posicion = 2
    Case "IT2001SUBTY":
        Posicion = 3
    Case "IT2001ENDDA":
        Posicion = 4
    Case "IT2001BEGDA":
        Posicion = 5
    Case "IT2001ATEXT":
        Posicion = 6
    Case "IT2001AWART":
        Posicion = 7
    Case "IT2001ABWTG":
        Posicion = 8
'IT2002
    Case "IT2002PERNR":
        Posicion = 1
    Case "IT2002INFTY":
        Posicion = 2
    Case "IT2002SUBTY":
        Posicion = 3
    Case "IT2002ENDDA":
        Posicion = 4
    Case "IT2002BEGDA":
        Posicion = 5
    Case "IT2002ATEXT":
        Posicion = 6
    Case "IT2002AWART":
        Posicion = 7
    Case "IT2002ABWTG":
        Posicion = 8
    Case Else
        Posicion = -1
End Select
Indice2 = Posicion
End Function

Public Sub Actualizar_Infotipos(ByVal Inicio_Validez, ByVal Fin_Validez)
' ---------------------------------------------------------------------------------------------
' Descripcion: Actualiza los datos guardados en el array
' Autor      : FGZ
' Fecha      : 27/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim Monto As Double
Dim Cantidad As Double
Dim Nomina As String
Dim Nro_Identific As String
Dim Nro_Tipo_ID As String
Dim Dias_Nomina As Integer
Dim TipoLic As Long
Dim OK As Boolean

    Select Case UltimoInfotipo
    Case "IT0008":
        If Not EsNulo(Arr_Datos(5).valor) Then
            Nomina = Trim(Arr_Datos(5).valor)
            Monto = CDbl(Arr_Datos(7).valor)
            Cantidad = 0
            Inicio_Validez = StrToDate(Arr_Datos(1).valor, OK, FormatoFechaSap1, NuloFechaSap1)
            
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "BEGDA. Formato de fecha invalido"
                Exit Sub
            End If
            Fin_Validez = StrToDate(Arr_Datos(2).valor, OK, FormatoFechaSap1, NuloFechaSap1)
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "ENDDA. Formato de fecha invalido"
                Exit Sub
            End If
            If Monto <> 0 Then
                Call Insertar_Novedad(Nomina, Monto, Cantidad, Inicio_Validez, Fin_Validez, "IT0008")
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Infotipo " & UltimoInfotipo & " no insertado. Nomina con valor nulo."
        End If
    Case "IT0014":
        If Not EsNulo(Arr_Datos(1).valor) Then
            Nomina = Trim(Arr_Datos(1).valor)
            Monto = CDbl(Arr_Datos(5).valor)
            Cantidad = CDbl(Arr_Datos(6).valor)
            Inicio_Validez = StrToDate(Arr_Datos(3).valor, OK, FormatoFechaSap1, NuloFechaSap1)
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "BEGDA. Formato de fecha invalido"
                Exit Sub
            End If
            Fin_Validez = StrToDate(Arr_Datos(4).valor, OK, FormatoFechaSap1, NuloFechaSap1)
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "ENDDA. Formato de fecha invalido"
                Exit Sub
            End If
            If Monto <> 0 Then
                Call Insertar_Novedad(Nomina, Monto, Cantidad, Inicio_Validez, Fin_Validez, "IT0008")
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Infotipo " & UltimoInfotipo & " no insertado. Nomina con valor nulo."
        End If
    Case "IT0015":
        If Not EsNulo(Arr_Datos(1).valor) Then
            Nomina = Trim(Arr_Datos(1).valor)
            Monto = CDbl(Arr_Datos(4).valor)
            Cantidad = CDbl(Arr_Datos(5).valor)
            If Monto <> 0 Then
                Call Insertar_Novedad(Nomina, Monto, Cantidad, Inicio_Validez, Fin_Validez, "IT0008")
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Infotipo " & UltimoInfotipo & " no insertado. Nomina con valor nulo."
        End If
    Case "IT0021":
        Call Insertar_Familiar
    Case "IT0185":
        If Not EsNulo(Arr_Datos(1).valor) Then
            Nro_Tipo_ID = Trim(Arr_Datos(1).valor)
            If Not EsNulo(Arr_Datos(3).valor) Then
                Nro_Identific = Trim(Arr_Datos(3).valor)
                Call Insertar_Documento(Nro_Identific, Nro_Tipo_ID)
            Else
                Flog.writeline Espacios(Tabulador * 2) & "Infotipo " & UltimoInfotipo & " no insertado. Nro de Doc con valor nulo."
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Infotipo " & UltimoInfotipo & " no insertado. Tipo de Doc con valor nulo."
        End If
    Case "IT0392":
        If Not EsNulo(Arr_Datos(3).valor) Then
            Inicio_Validez = StrToDate(Arr_Datos(1).valor, OK, FormatoFechaSap1, NuloFechaSap1)
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "BEGDA. Formato de fecha invalido"
                Exit Sub
            End If
            Fin_Validez = StrToDate(Arr_Datos(2).valor, OK, FormatoFechaSap1, NuloFechaSap1)
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "ENDDA. Formato de fecha invalido"
                Exit Sub
            End If
            Call Insertar_Seguridad_Social(Inicio_Validez, Fin_Validez)
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Infotipo " & UltimoInfotipo & " no insertado. Nro de Familiar Tipo de Doc con valor nulo."
        End If
    Case "IT0394":
        Call Insertar_Familiar
    Case "IT2001":
        If CLng(Arr_Datos(3).valor) <> 0 Then
            Dias_Nomina = CInt(Arr_Datos(5).valor)
            TipoLic = CLng(Arr_Datos(3).valor)
            Inicio_Validez = StrToDate(Arr_Datos(1).valor, OK, FormatoFechaSap2, NuloFechaSap2)
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "BEGDA. Formato de fecha invalido"
                Exit Sub
            End If
            Fin_Validez = StrToDate(Arr_Datos(2).valor, OK, FormatoFechaSap2, NuloFechaSap2)
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "ENDDA. Formato de fecha invalido"
                Exit Sub
            End If
            Call Insertar_Licencia(TipoLic, Inicio_Validez, Fin_Validez, 0, Dias_Nomina, Empleado.Tercero, True, OK)
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
            Flog.writeline Espacios(Tabulador * 3) & "Tipo de Licencia Desconocido"
        End If
    Case "IT2002":
        If CLng(Arr_Datos(3).valor) <> 0 Then
            Dias_Nomina = CInt(Arr_Datos(5).valor)
            TipoLic = CLng(Arr_Datos(3).valor)
            Inicio_Validez = StrToDate(Arr_Datos(1).valor, OK, FormatoFechaSap2, NuloFechaSap2)
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "BEGDA. Formato de fecha invalido"
                Exit Sub
            End If
            Fin_Validez = StrToDate(Arr_Datos(2).valor, OK, FormatoFechaSap2, NuloFechaSap2)
            If Not OK Then
                Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
                Flog.writeline Espacios(Tabulador * 3) & "ENDDA. Formato de fecha invalido"
                Exit Sub
            End If
            Call Insertar_Licencia(TipoLic, Inicio_Validez, Fin_Validez, 0, Dias_Nomina, Empleado.Tercero, True, OK)
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
            Flog.writeline Espacios(Tabulador * 3) & "Tipo de Licencia Desconocido"
        End If
    Case Else
    End Select
    
    'Limpio el array de datos
    For i = 1 To UBound(Arr_Datos)
        Arr_Datos(i).ID_campo = 0
        'Arr_Datos(i).Campo = ""
        Arr_Datos(i).Descripcion = ""
        Arr_Datos(i).valor = ""
        Arr_Datos(i).TipoDato = ""
        Arr_Datos(i).IT = ""
    Next i
    
End Sub



Public Sub Insertar_Seguridad_Social(ByVal Inicio_Validez, ByVal Fin_Validez)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta datos de seguridad social del empleado
' Autor      : FGZ
' Fecha      : 22/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Texto As String
Dim Obra_Social As String
Dim Nro_Obra_Social As Long
Dim aux As String
Dim Reparto As Boolean
Dim Capitalizacion As Boolean
Dim Codigo_de_actividad As String
Dim Nro_Codigo_de_actividad As Long
    
    aux = Trim(Arr_Datos(4).valor)
    Capitalizacion = IIf(aux = 0, False, True)
    
    aux = Trim(Arr_Datos(5).valor)
    Reparto = IIf(aux = 0, False, True)
    
    Obra_Social = Trim(Arr_Datos(3).valor)
    If Not EsNulo(Obra_Social) Then
        Nro_Obra_Social = CLng(CalcularMapeoInv(Obra_Social, "T7AR34", "0"))
    Else
        Nro_Obra_Social = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. obra social"
    End If


    If Capitalizacion Then
        'Obra social - [ RHPro(Obra social)]
        Texto = "Obra Social - [ RHPro(Obra social)] " & Obra_Social
        If Nro_Obra_Social <> 0 Then
            Call Insertar_His_Estructura(17, Nro_Obra_Social, Empleado.Tercero, Inicio_Validez, Fin_Validez)
        Else
            Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
            Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
        End If
    End If
    If Reparto Then
        'Obra social - [ RHPro(Obra social)]
        Texto = "Obra Social - [ RHPro(Obra social)] " & Obra_Social
        If Nro_Obra_Social <> 0 Then
            Call Insertar_His_Estructura(24, Nro_Obra_Social, Empleado.Tercero, Inicio_Validez, Fin_Validez)
        Else
            Flog.writeline Espacios(Tabulador * 3) & "Error. Infotipo no actualizado"
            Flog.writeline Espacios(Tabulador * 3) & "No se encontró el mapeo de la " & Texto
        End If
    End If



    Codigo_de_actividad = Trim(Arr_Datos(6).valor)
    If Not EsNulo(Codigo_de_actividad) Then
        Nro_Codigo_de_actividad = CLng(CalcularMapeoInv(Codigo_de_actividad, "T7AR38", "0"))
    Else
        Nro_Codigo_de_actividad = 0
        Flog.writeline Espacios(Tabulador * 3) & "Valor Nulo. Código de actividad del empleado"
    End If


End Sub

Public Sub Insertar_Familiar()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta el familiar del empleado
' Autor      : FGZ
' Fecha      : 22/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Aux_Tercero As Long

Dim Parentesco As Long
Dim Estado_civil As Long
Dim Tersex As Integer
Dim PaisNro As Long
Dim NacionalNro As Long
Dim Fecha_de_nacimiento As String
Dim Nombre_de_pila As String
Dim Apellidos As String
Dim Identificacion_de_objeto As String
Dim Aux_Incapacitado As String
Dim Incapacitado As Boolean
Dim OK As Boolean
Dim Nivel_de_estudios As String
Dim Nro_Identific As String
Dim Nro_Tipo_ID As String


Dim rs_Tercero As New ADODB.Recordset

    
    Parentesco = CLng(Trim(Arr_Datos(3).valor))
    Identificacion_de_objeto = Trim(Arr_Datos(5).valor)
    Nombre_de_pila = Trim(Arr_Datos(8).valor)
    Apellidos = Trim(Arr_Datos(6).valor) & " " & Trim(Arr_Datos(7).valor)
    Fecha_de_nacimiento = StrToDate(Arr_Datos(9).valor, OK, FormatoFechaSap1, NuloFechaSap1)
    If Not OK Then
        Flog.writeline Espacios(Tabulador * 2) & "Error. Infotipo no actualizado"
        Flog.writeline Espacios(Tabulador * 3) & "FGBDT. Formato de fecha invalido"
        'Exit Sub
    End If
    Tersex = -1    'no viene informado
    Nivel_de_estudios = 0 'no viene informado
    Aux_Incapacitado = IIf(Not EsNulo(Trim(Arr_Datos(13).valor)), Trim(Arr_Datos(13).valor), 0)
    If Aux_Incapacitado = 0 Then
        Incapacitado = False
    Else
        Incapacitado = True
    End If
    
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
    
    
    StrSql = "SELECT * FROM tercero "
    StrSql = StrSql & " INNER JOIN ter_tip ON tercero.ternro = ter_tip.ternro AND ter_tip.tipnro = 3 "
    StrSql = StrSql & " INNER JOIN familiar ON familiar.ternro = tercero.ternro AND familiar.empleado = " & Empleado.Tercero
    StrSql = StrSql & " WHERE ternom = '" & Format_Str(Nombre_de_pila, 25, False, "") & "'"
    StrSql = StrSql & " AND terape = '" & Format_Str(Apellidos, 25, False, "") & "'"
    OpenRecordset StrSql, rs_Tercero
    
    If rs_Tercero.EOF Then
    
        'busco la nacionalidad y pais del tercero
        StrSql = "SELECT NacionalNro, PaisNro FROM tercero "
        StrSql = StrSql & " WHERE tercero.ternro = " & Empleado.Tercero
        OpenRecordset StrSql, rs_Tercero
        If Not rs_Tercero.EOF Then
            NacionalNro = rs_Tercero!NacionalNro
            PaisNro = rs_Tercero!PaisNro
        Else
            'Esto no puede darse porque son campos obligatorios en el alta de empleados
            NacionalNro = 0
            PaisNro = 0
        End If
    
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
        StrSql = "INSERT INTO familiar (empleado,ternro,parenro,famnrocorr,famest,famtrab,famestudia,famcernac,faminc) "
        StrSql = StrSql & " VALUES ( "
        StrSql = StrSql & Empleado.Tercero
        StrSql = StrSql & "," & Aux_Tercero
        StrSql = StrSql & "," & Parentesco
        StrSql = StrSql & "," & Identificacion_de_objeto
        StrSql = StrSql & ",-1" 'estado
        StrSql = StrSql & ",0"  'trabaja
        StrSql = StrSql & "," & IIf(EsNulo(Nivel_de_estudios), 0, -1) 'estudia
        StrSql = StrSql & ",0"
        StrSql = StrSql & "," & CInt(Incapacitado)
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Insertar el documento
        If Not EsNulo(Arr_Datos(9).valor) Then
            Nro_Tipo_ID = Trim(Arr_Datos(9).valor)
            If Not EsNulo(Arr_Datos(10).valor) Then
                Nro_Identific = Trim(Arr_Datos(10).valor)
                Call Insertar_Documento_Familiar(Aux_Tercero, Nro_Identific, Nro_Tipo_ID)
            Else
                Flog.writeline Espacios(Tabulador * 2) & "Infotipo " & UltimoInfotipo & " no insertado. Nro de Doc con valor nulo."
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Infotipo " & UltimoInfotipo & " no insertado. Tipo de Doc con valor nulo."
        End If
        
        
        
        
    Else
        StrSql = "UPDATE tercero SET "
        StrSql = StrSql & " terfecnac = " & ConvFecha(Fecha_de_nacimiento)
        StrSql = StrSql & " ,tersex = " & Tersex
        StrSql = StrSql & " ,nacionalnro = " & NacionalNro
        StrSql = StrSql & " ,paisnro = " & PaisNro
        StrSql = StrSql & " WHERE ternro = " & rs_Tercero!ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Actualizar el documento
        
        
        StrSql = "UPDATE familiar SET "
        StrSql = StrSql & " parenro = " & Parentesco
        StrSql = StrSql & " ,famestudia = " & IIf(EsNulo(Nivel_de_estudios), 0, -1) 'estudia
        StrSql = StrSql & " ,faminc = " & CInt(Incapacitado)
        StrSql = StrSql & " WHERE empleado = " & Empleado.Tercero
        StrSql = StrSql & " AND ternro = " & rs_Tercero!ternro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If



End Sub



Public Sub LineaModelo_277(ByVal strlinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : FGZ
' Fecha      : 09/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim aux As String
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
    pos2 = InStr(pos1, strlinea, Separador)
    Tabla = Trim(Mid$(strlinea, pos1, pos2 - pos1))
    
    'Auxiliar
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strlinea, Separador)
    aux = Trim(Mid$(strlinea, pos1, pos2 - pos1))
    
    'Auxiliar 2
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strlinea, Separador)
    Aux2 = Trim(Mid$(strlinea, pos1, pos2 - pos1))
    
    'Codigo
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strlinea, Separador)
    Codigo = Trim(Mid$(strlinea, pos1, pos2 - pos1))
    Codigo = Replace(Codigo, "'", "´")
    
    'Descripcion
    pos1 = pos2 + 1
    pos2 = Len(strlinea)
    If pos2 < pos1 Then
        Descripcion = ""
    Else
        Descripcion = Mid(strlinea, pos1, (pos2 - pos1) + 1)
        Descripcion = Replace(Descripcion, "'", "´")
    End If

    If EsNulo(Trim(Codigo)) Then
        Codigo = Descripcion
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
    If Not EsNulo(Trim(Descripcion)) Then
        Descripcion = Trim(Codigo) & " $ " & Trim(Descripcion)
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
    'If Not EsNulo(Trim(Codigo)) And Mid(UCase(Trim(Aux2)), 1, 2) = "AR" Then
    If Not EsNulo(Trim(Descripcion)) Then
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
    Tenro = 53
    'If Not EsNulo(Trim(Codigo)) And (UCase(Mid(Trim(Descripcion), 1, 2)) = "AR" Or Trim(Codigo) = "1" Or Trim(Codigo) = "2" Or Trim(Codigo) = "4" Or Trim(Codigo) = "9") Then
    If Not EsNulo(Trim(Descripcion)) Then
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
    Tenro = 53
    'If Not EsNulo(Trim(Codigo)) And UCase(Mid(Trim(Descripcion), 1, 2)) = "AR" Then
    If Not EsNulo(Trim(Descripcion)) Then
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
    Tablaref = "TSUC"
    Tenro = 1
    'If Not EsNulo(Trim(Codigo)) And UCase(Mid(Trim(aux), 1, 2)) = "AR" Then
    If Not EsNulo(Trim(Descripcion)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
'        If Inserto_estr Then
'            Call CreaTercero(10, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
'        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If
Case 8: 'Area de Nomina  T549A   22  FORMA DE LIQUIDACION    NO
    Tablaref = "T549A"
    Tenro = 22
    If Not EsNulo(Trim(Descripcion)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
            'Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If
Case 9: 'Centro de Costo CSKS    5   Centro de Costo si
    Tablaref = "CSKS"
    Tenro = 5
    If Not EsNulo(Trim(Descripcion)) Then
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
    If Not EsNulo(Trim(Descripcion)) Then
        Descripcion = Trim(Codigo) & " $ " & Trim(Descripcion)
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
    If Not EsNulo(Trim(Descripcion)) Then
        Call ValidarPais(Descripcion, Nro_Pais)
        Call Mapear(Tablaref, Codigo, CStr(Nro_Pais))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 14:  'Estados - Provincia T005S       PROVINCIA   SI
    Tablaref = "T005S"
    'If Not EsNulo(Trim(Codigo)) And Mid(Descripcion, 1, 2) = "AR" Then
    If Not EsNulo(Trim(Descripcion)) Then
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
    If Not EsNulo(Trim(Descripcion)) Then
        Call ValidarEstadoCivil(Descripcion, Nro_EstadoCivil)
        Call Mapear(Tablaref, Codigo, CStr(Nro_EstadoCivil))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If

Case 18:    'Bancos  T012    41  BANCO   SI
    Tablaref = "T012"
    Tenro = 41
    If Not EsNulo(Trim(Descripcion)) Then
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
    If Not EsNulo(Trim(Descripcion)) Then
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
    If Not EsNulo(Trim(Descripcion)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
            'Call CreaTercero(Tenro, Descripcion, Nro_Tercero)
            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
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
    If Not EsNulo(Trim(Descripcion)) Then
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
    If Not EsNulo(Trim(Descripcion)) Then
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
    If Not EsNulo(Trim(Descripcion)) Then
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
    If Not EsNulo(Trim(Descripcion)) Then
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
Case 98:  'Relacion Laboral
    Tablaref = "TCONV"
    Tenro = 19
    If Not EsNulo(Trim(Descripcion)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
            Call CreaTercero(19, Descripcion, Nro_Tercero)
            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
        End If
        Call Mapear(Tablaref, Codigo, CStr(Nro_Estr))
    Else
        'No se inserta nada
        FlogE.writeline Espacios(Tabulador * 3) & "Linea " & NroLinea & ":" & Tablaref & ". Codigo nulo"
    End If
Case 99:  'Horario de trabajo
    Tablaref = "HORAT"
    Tenro = 21
    If Not EsNulo(Trim(Descripcion)) Then
        Call ValidaEstructura(Tenro, Descripcion, Nro_Estr, Inserto_estr)
        If Inserto_estr Then
'            Call CreaTercero(7, Descripcion, Nro_Tercero)
'            Call CreaComplemento(Tenro, Nro_Tercero, Nro_Estr, Descripcion)
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
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Function Primer_Campo_Infotipo(ByVal Infotipo As String, ByVal NroLinea As Long) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna si el campo es el primero de la secuencia de cada it
' Autor      : FGZ
' Fecha      : 27/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim CantidadLineas As Long
Dim Resto

Select Case Infotipo
    Case "IT0008":
        CantidadLineas = 7
    Case "IT0014":
        CantidadLineas = 7
    Case "IT0015":
        CantidadLineas = 6
    Case "IT0021":
        CantidadLineas = 14
    Case "IT0185":
        CantidadLineas = 4
    Case "IT0392":
        CantidadLineas = 9
    Case "IT2001":
        CantidadLineas = 5
    Case "IT2002":
        CantidadLineas = 5
    Case Else
End Select

Resto = NroLinea Mod CantidadLineas
If Resto = 1 Or NroLinea = 1 Then
    Primer_Campo_Infotipo = True
Else
    Primer_Campo_Infotipo = False
End If
End Function


Public Function Primer_Campo2_Infotipo(ByVal Infotipo As String, ByVal Campo As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna si el campo es el primero de la secuencia de cada it
' Autor      : FGZ
' Fecha      : 22/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Primer_Campo As String

Select Case Infotipo
    Case "IT0008":
        Primer_Campo = "BEGDA"
    Case "IT0014":
        Primer_Campo = "LGART"
    Case "IT0015":
        Primer_Campo = "LGART"
    Case "IT0021":
        Primer_Campo = "BEGDA"
    Case "IT0185":
        Primer_Campo = "ICTYP"
    Case "IT0392":
        Primer_Campo = "BEGDA"
    Case "IT2001":
        Primer_Campo = "BEGDA"
    Case "IT2002":
        Primer_Campo = "BEGDA"
    Case Else
End Select
If UCase(Campo) = Primer_Campo Then
    Primer_Campo2_Infotipo = True
Else
    Primer_Campo2_Infotipo = False
End If
End Function


Public Function StrToDate(ByVal Str As String, ByRef OK As Boolean, Optional ByVal FormatoEntrada As String, Optional ByVal Nulo As String) 'As Date
' ---------------------------------------------------------------------------------------------
' Descripcion: Convierte el string a fecha
' Autor      : FGZ
' Fecha      : 20/03/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Fecha
Dim dia As String
Dim mes As String
Dim Anio As String
Dim PosD
Dim PosM
Dim PosY

If EsNulo(FormatoEntrada) Then
    FormatoEntrada = "dd/mm/yyyy"
End If
If Str = Nulo Then
    Fecha = ""
End If

'If EsNulo(Nulo) Then
'    Nulo = "31/12/9999"
'End If

FormatoEntrada = LCase(FormatoEntrada)
'Busco las posiciones
PosD = InStr(1, FormatoEntrada, "d")
PosM = InStr(1, FormatoEntrada, "m")
PosY = InStr(1, FormatoEntrada, "y")

If Not EsNulo(Trim(Str)) Then
    dia = Mid(Str, PosD, 2)
    mes = Mid(Str, PosM, 2)
    Anio = Mid(Str, PosY, 4)
    
    If Str = Nulo Then
        Fecha = ""
        OK = True
    Else
        If IsDate(dia & "/" & mes & "/" & Anio) Then
            Fecha = C_Date(dia & "/" & mes & "/" & Anio)
            OK = True
        Else
            Fecha = ""
            OK = False
        End If
    End If
    StrToDate = Fecha
Else
    Fecha = ""
    OK = True
End If
End Function

