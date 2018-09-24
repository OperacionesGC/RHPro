Attribute VB_Name = "MdlInfoExport"
Option Explicit

Global Legajo As Long
Global Tercero As Long
Global InfoNro As Long
Global InfotipoVal As String

Public Sub Generar_Archivo(ByVal NroModelo As Long, ByVal Archivo As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento
' Autor      : FGZ
' Fecha      : 23/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Empleados As New ADODB.Recordset

StrSql = "SELECT * FROM empleado WHERE empest = -1 ORDER BY empleg "
OpenRecordset StrSql, rs_Empleados

'Determino la proporcion de progreso
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
    CEmpleadosAProc = 1
End If
IncPorc = (100 / CEmpleadosAProc)

Do While Not rs_Empleados.EOF
    Legajo = rs_Empleados!empleg
    Tercero = rs_Empleados!Ternro
    
    'Exporta todos los infotipos configurados activos para el empleado
    Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 0) & "Empleado " & Legajo
    Flog.Writeline
    Call Exportar_Infotipos(NroModelo)

    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    rs_Empleados.MoveNext
Loop
    
If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
Set rs_Empleados = Nothing
End Sub



Public Sub Exportar_Infotipos(ByVal NroModelo As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador segun infotipo
' Autor      : FGZ
' Fecha      : 23/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'0000    Informacion de alta
'0001    Asignacion Organizacional
'0002    Datos Personales
'0006    Direcciones
'0008    Emolumentos Basicos
'0009    Relacion Bancaria
'0016    Elementos de Contratos
'0021    Datos Familiares
'0032    Datos Internos de la Empresa
'0041    Datos de Fecha
'0057    Asociaciones
'0185    Identificacion Personal
'0390    Impuesto a las Ganancias - Deducciones
'0392    Seguridad Social
'2006    Derechos pendientes de Tiempos
'9999    Conversión de acumulados históricos de la liquidación

'Dim rs_Infotipos As New ADODB.Recordset
'StrSql = " SELECT infotipos.* FROM modelo_infotipos "
'StrSql = StrSql & " INNER JOIN infotipos ON infotipos.inftipnro = modelo_infotipos.inftipnro "
'StrSql = StrSql & " WHERE modelo_infotipos.modnro = " & NroModelo
'StrSql = StrSql & " ORDER BY infotipos.inftiporden "
'OpenRecordset StrSql, rs_Infotipos
'Do While Not rs_Infotipos.EOF
'   InfoNro = rs_Infotipos!Inftipcod

Select Case InfotipoVal
    
    Case "0000":
        'Datos del empleado
        Call Export_Infotipo_0000
    'Case "0001":
    '    Call Export_Infotipo_0001
    'Case "0002":
    '    Call Export_Infotipo_0002
    Case "0006":
        'Direccion del empleado
        Call Export_Infotipo_0006
    Case "0008":
        'Emolumentos Basicos
        Call Export_Infotipo_0008
    Case "0009":
        'Relacion Bancaria
        Call Export_Infotipo_0009
    Case "0016":
        'Elementos de Contratos
        Call Export_Infotipo_0016
    Case "0021":
        'Familiares del empleado
        Call Export_Infotipo_0021
    Case "0032":
        'Datos Internos de la empresa
        Call Export_Infotipo_0032
    Case "0041":
        'Datos de Fechas
        Call Export_Infotipo_0041
    Case "0057":
        'Asociaciones
        Call Export_Infotipo_0057
    Case "0185":
        'Identificacion Personal
        Call Export_Infotipo_0185
    Case "0390":
        'Impuestos a las ganacias - deducciones
        Call Export_Infotipo_0390
    Case "0392":
        'Seguridad Social
        Call Export_Infotipo_0392
    Case "2006":
        'Derechos pendientes de tiempos
        Call Export_Infotipo_2006
    Case "9999":
        'Conversion de acumulados historicos de la liquidacion
        Call Export_Infotipo_9999
    Case Else
    
End Select
    
'    rs_Infotipos.MoveNext
'Loop

'    If rs_Infotipos.State = adStateOpen Then rs_Infotipos.Close
'    Set rs_Infotipos = Nothing
End Sub



Public Sub Export_Infotipo_0000()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0000. Informacion de alta.
'               Esta conversión debe realizarse en primer lugar.
'               Solo para empleados activos al momento de la fecha de corte.
' Autor      : Scarpa D.
' Fecha      : 01/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'PERNR   Número de Personal                 NUMC    8       Sí
'BEGDA   Fecha Ingreso                      DATS    8       Sí                  Fecha de ingreso del empleado. En caso de reingresantes va la ultima fecha de ingreso y en el caso de cambios de sociedad va la fecha de ingreso a la sociedad vigente.
'ENDDA   Fecha fin de Validez               DATS    8       Sí                  Valor fijo: '99991231'
'WERKS   División de Personal               CHAR    4       Sí      T500P       Zona geografica AFIP
'PERSG   Grupo de Personal                  CHAR    1       Sí      T501
'PERSK   Area de Personal                   CHAR    2       Sí      T503K       Convenio del empleado
'PERNR   Número de Personal                 NUMC    8       Sí
'BEGDA   Fecha de inicio de validez         DATS    8       Sí                  Fecha de ingreso del empleado
'ENDDA   Fecha fin de Validez               DATS    8       Sí                  Valor fijo: '99991231'
'BTRTL   Subdivisión de Personal            CHAR    4       Sí      T001P       Tener en cuenta los dias de vacaciones, para los fuera de convenio, si le correspode como la LCT, colocar el 0001 , si corresponde como sanidad colocar SN01.Para los pasantes colocar 0001.
'ABKRS   Area de Nómina                     CHAR    2       Sí      T549A
'ANSVH   Relación Laboral                   CHAR    2       Sí      T542A       Blanco
'PLANS   Posición                           NUMC    8       Sí                  (Por el momento dejar en blanco)
'VDSK1   Clave de Organización              CHAR    14      SI      T527O       Blanco
'SACHZ   Encargado para entrada de tiempos  CHAR    3       Sí      T526
'PERNR   Número de Personal                 NUMC    8       Sí
'BEGDA   Fecha inicio Validez               DATS    8       Sí                  Fecha de nacimiento
'ENDDA   Fecha fin de Validez               DATS    8       Sí                  Valor fijo: '99991231'
'ANRED   Tratamiento                        CHAR    5       Sí      T522G
'NACHN   Apellido                           CHAR    40      Sí                  En mayúsculas
'VORNA   Nombre                             CHAR    40      Sí                  En mayúsculas
'CCUIL   CUIL                               CHAR    13      Sí                  Formato: "##-########-#"
'GBDAT   Fecha Nacimiento                   DATS    8       Sí
'SPRSL   Id. Comunicación                   CHAR    1       Sí                  Idioma para la comunicación y la correspondencia. S=Español
'NATIO   Nacionalidad                       CHAR    3       Sí      T005
'FATXT   Estado Civil                       CHAR    6       Sí      T502T
'ANZKD   Número de Hijos                    DEC     3       Sí
'GESC1   Sexo Masculino                     CHAR    1       Sí                  Si es verdadero poner una X, sino un espacio
'GESC2   Sexo Femenino                      CHAR    1       Sí                  Si es verdadero poner una X, sino un espacio
' ------------------------------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0000"
        PrimerCampo = True
        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos del empleado "
                
        'Estr.   : PSPAR
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        Salida = Format_StrNro(Legajo, 8, True, "0")
            
        'Estr.   : P0000
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        '          Fecha de ingreso del empleado. Fases.
        Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
        
        'Estr.   : P0000
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor Fijo
        Salida = Salida & Separador & "99991231"
        
        'Estr.   : PSPAR
        'Campo   : WERKS
        'Descrip.: Division de Personal
        '          Zona Geografica AFIP
        '          Mapeo T500P
        '          Busco la zona de la sucursal y luego la mapeo
        Salida = Salida & Separador & busq_ZonaGeograficaAFIP(Tercero)
        
        'Estr.   : PSPAR
        'Campo   : PERSG
        'Descrip.: Grupo de Personal - [ RHPro (Tipo de Contrato)]
        '          Mapeo tabla T501P
        Salida = Salida & Separador & busq_GrupoDePersonal(Tercero)
        
        'Estr.   : PSPAR
        'Campo   : PERSK
        'Descrip.: Area de Personal - [ RHPro (Convenio)]
        '          Mapeo tabla T503K
        Salida = Salida & Separador & busq_AreaDePersonal(Tercero)
                     
        'Estr.   : P0001
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        Salida = Salida & Separador & Format_StrNro(Legajo, 8, True, "0")
            
        'Estr.   : P0001
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        '          Fecha de ingreso del empleado. Fases.
        Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
        
        'Estr.   : P0001
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor Fijo
        Salida = Salida & Separador & "99991231"
                     
        'Estr.   : P0001
        'Campo   : BTRTL
        'Descrip.: Subdivision del personal
        '          Tener en cuenta los dias de vacaciones, para los fuera de convenio,
        '          si le correspode como la LCT, colocar el 0001 ,
        '          si corresponde como sanidad colocar SN01.
        '          Para los pasantes colocar 0001.
        
        'Mapeo tabla T001P
        If busq_AreaDePersonal(Tercero) = "FN" Then
            If DateDiff("d", CDate(Fecha_Alta_Fase), CDate("2001-01-01")) > 0 Then
               Salida = Salida & Separador & "0001"
            Else
               Salida = Salida & Separador & "SN01"
            End If
        Else
            Salida = Salida & Separador & "SN01"
        End If
        
        'Estr.   : P0001
        'Campo   : ABKRS
        'Descrip.: Area de Nomina
        '          Mapeo tabla T549A
        Salida = Salida & Separador & "NP"
                     
        'Estr.   : P0001
        'Campo   : ANSVH
        'Descrip.: Relacion Laboral
        '          Mapeo tabla T549A
        '          Valor Fijo en Blanco
        Salida = Salida & Separador & Space(2)
                     
        'Estr.   : P0001
        'Campo   : PLANS
        'Descrip.: Posicion
        '          por ahora en Blanco
        Salida = Salida & Separador & Space(8)

        'Estr.   : P0001
        'Campo   : VDSK1
        'Descrip.: Clave de organizacion
        '          blanco
        Salida = Salida & Separador & Space(14)

        'Estr.   : P0001
        'Campo   : SACHZ
        'Descrip.: Encargado para entrada de tiempos
        '          Mapeo tabla T526
        Salida = Salida & Separador & "T01"
                            
        'Estr.   : P0002
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        Salida = Salida & Separador & Format_StrNro(Legajo, 8, True, "0")
            
        'Estr.   : P0002
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        '          Fecha de nacimiento
        If IsNull(rs_Tercero!Terfecnac) Then
           Salida = Salida & Separador & Space(8)
        Else
           Salida = Salida & Separador & Format_Fecha(rs_Tercero!Terfecnac, 1)
        End If
        
        'Estr.   : P0002
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor Fijo
        Salida = Salida & Separador & "99991231"
                            
        'Estr.   : P0002
        'Campo   : ANRED
        'Descrip.: Tratamiento
        '          Mapeo tabla T522G
        Salida = Salida & Separador & CalcularMapeo(rs_Tercero!Tersex, "T522G", "1")
                            
        'Estr.   : P0002
        'Campo   : NACHN
        'Descrip.: Apellido
        Salida = Salida & Separador & UCase(rs_Tercero!Terape)
                            
        'Estr.   : P0002
        'Campo   : VORNA
        'Descrip.: Nombre
        Salida = Salida & Separador & UCase(rs_Tercero!Ternom)
        
        'Estr.   : Q0002
        'Campo   : CCUIL
        'Descrip.: CUIL
        Salida = Salida & Separador & busq_CUIL(Tercero)

        'Estr.   : P0002
        'Campo   : GBDAT
        'Descrip.: Fecha de nacimiento
        If IsNull(rs_Tercero!Terfecnac) Then
           Salida = Salida & Separador & Space(8)
        Else
           Salida = Salida & Separador & Format_Fecha(rs_Tercero!Terfecnac, 1)
        End If

        'Estr.   : P0002
        'Campo   : SPRSL
        'Descrip.: Id. Comunicacion
        Salida = Salida & Separador & "S"

        'Estr.   : P0002
        'Campo   : NATIO
        'Descrip.: Nacionalidad
        '          Mapeo Tabla T005
        Salida = Salida & Separador & CalcularMapeo(rs_Tercero!NacionalNro, "T005", "AR")

        'Estr.   : Q0002
        'Campo   : FATXT
        'Descrip.: Estado Civil
        '          Mapeo Tabla T502T
        Salida = Salida & Separador & CalcularMapeo(rs_Tercero!Terestciv, "T502T", "0")

        'Estr.   : P0002
        'Campo   : ANZKD
        'Descrip.: Numero de HIjos
        Salida = Salida & Separador & Format_StrNro(busq_NumeroDeHijos(Tercero), 3, True, "0")

        'Estr.   : Q0002
        'Campo   : GESC1
        'Descrip.: Sexo Masculino
        If CInt(rs_Tercero!Tersex) = -1 Then
           Salida = Salida & Separador & "X"
        Else
           Salida = Salida & Separador & " "
        End If
        
        'Estr.   : Q0002
        'Campo   : GESC2
        'Descrip.: Sexo Femenino
        If CInt(rs_Tercero!Tersex) = 0 Then
           Salida = Salida & Separador & "X"
        Else
           Salida = Salida & Separador & " "
        End If
                
        '-------------------------------------------------------------
        ' Escribo en el erchivo
        fExport.Writeline Salida
    
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


'----------------------------------------------------------------
'Busca la zona de la sucursal del empleado y luego el mapeo
Function busq_ZonaGeograficaAFIP(ByVal Tercero)

    Dim Parametro
    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim Sucursal
    Dim Salida
    Dim Zona
    
    Salida = "0000"
    Parametro = "1"
    
    'Busco la estructura sucursal
    StrSql = " SELECT estrnro FROM his_estructura " & _
             " WHERE ternro = " & Tercero & " AND " & _
             " tenro = 1 AND " & _
             " (htetdesde <= " & ConvFecha(Fecha_Hasta) & ") AND " & _
             " ((" & ConvFecha(Fecha_Hasta) & " <= htethasta) or (htethasta is null))"
    
    OpenRecordset StrSql, rs_Consult
    
    Sucursal = "0000"
    If Not rs_Consult.EOF Then
       Sucursal = rs_Consult!Estrnro
    End If
    
    rs_Consult.Close
    
    'Busco la sucursal
    If Sucursal <> "" Then
     
        StrSql = " SELECT * FROM sucursal " & _
                 " WHERE estrnro =" & Sucursal
                 
        OpenRecordset StrSql, rs_Consult
        
        Sucursal = ""
        If Not rs_Consult.EOF Then
           Sucursal = rs_Consult!Ternro
        End If
        
        rs_Consult.Close
        
    End If
    
    'Busco la zona de la sucursal y el mapeo a SAP
    If Sucursal <> "" Then
            
        StrSql = " SELECT zona.zonanro,zona.zonacod,locdesc FROM detdom " & _
                 " INNER JOIN zona ON zona.zonanro = detdom.zonanro " & _
                 " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                 " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
                 " WHERE cabdom.ternro = " & Sucursal
        
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Salida = CalcularMapeo(rs_Consult!zonacod, "T500P", "0000")
        Else
            Flog.Writeline Espacios(Tabulador * 3) & "No se encontró la Zona de la sucursal " & Sucursal
        End If
        
        rs_Consult.Close
        
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "No se encontró sucursal para el empleado " & Legajo
    End If
    
    busq_ZonaGeograficaAFIP = Salida

End Function


'----------------------------------------------------------------
'Busca el telefono del empleado para el domicilio
Function busq_Telefono(ByVal Domnro, ByVal Default)

    Dim Parametro
    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim Salida
    
    Salida = "0000"
    
    'Busco el telefono del domicilio
    StrSql = " SELECT telnro, "
    StrSql = StrSql & " telfax,  "
    StrSql = StrSql & " teldefault, telcelular   "
    StrSql = StrSql & " FROM  telefono"
    StrSql = StrSql & " WHERE  domnro = " & Domnro
    StrSql = StrSql & "   AND teldefault = -1 "
    
    OpenRecordset StrSql, rs_Consult
    
    If Not rs_Consult.EOF Then
       Salida = rs_Consult!telnro
    Else
       Salida = Default
    End If
    
    rs_Consult.Close
    
    busq_Telefono = Salida

End Function


'----------------------------------------------------------------
'Busca el grupo del personal - contrato
Function busq_GrupoDePersonal(ByVal Tercero)
'----------------------------------------------------------------
'Tabla de Grupos de Personal - T501
'PERSG  PTEXT
'1      Activo/Efectivos
'2      Activo/Contratados
'3      Jubilados Efectivos
'4      Jubilados Contratado
'5      Pasantes
'9      Externo

    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim Salida
    Dim tipoEstructura
    
    tipoEstructura = 18
    
    'Busco el tipo de contrato y Mapeo
    StrSql = "SELECT * FROM his_estructura " & _
    " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " & _
    " WHERE his_estructura.tenro = " & tipoEstructura & _
    " AND his_estructura.ternro = " & Tercero & _
    " AND his_estructura.htetdesde <=" & ConvFecha(Fecha_Hasta) & _
    " AND (his_estructura.htethasta >= " & ConvFecha(Fecha_Desde) & _
    " OR his_estructura.htethasta IS NULL)"
    
    OpenRecordset StrSql, rs_Consult
    
    Salida = "1"
    
    If Not rs_Consult.EOF Then
        Salida = CalcularMapeo(rs_Consult!Estrnro, "T501", "1")
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "No se encontró Contrato para el empleado " & Legajo
    End If

    busq_GrupoDePersonal = Salida

End Function


'----------------------------------------------------------------
'Busca una estructura y su correspondiente mapeo
Function busq_Estructura(ByVal Tercero, ByVal tipoEstructura, ByVal Tablaref, ByVal Default)

    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim Salida
    
    'Busco el tipo de contrato y Mapeo
    StrSql = "SELECT * FROM his_estructura " & _
    " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " & _
    " WHERE his_estructura.tenro = " & tipoEstructura & _
    " AND his_estructura.ternro = " & Tercero & _
    " AND his_estructura.htetdesde <=" & ConvFecha(Fecha_Hasta) & _
    " AND (his_estructura.htethasta >= " & ConvFecha(Fecha_Desde) & _
    " OR his_estructura.htethasta IS NULL)"
    
    OpenRecordset StrSql, rs_Consult
    
    Salida = Default
    
    If Not rs_Consult.EOF Then
        Salida = CalcularMapeo(rs_Consult!Estrnro, Tablaref, Default)
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "No se encontró la estructura de tipo " & tipoEstructura & " para el empleado " & Legajo
    End If

    busq_Estructura = Salida

End Function

'----------------------------------------------------------------
'Busca el area del personal - convenio
Function busq_AreaDePersonal(ByVal Tercero)
'----------------------------------------------------------------
'Tabla de Grupos de Personal - T503K
'PERSK  PTEXT
'FN     Fuera de Convenio - Newprod
'NS     Sanidad Newprod

    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim Salida
    Dim tipoEstructura
    
    tipoEstructura = 19
    
    'Busco el tipo de convenio y Mapeo
    StrSql = "SELECT * FROM his_estructura " & _
    " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " & _
    " WHERE his_estructura.tenro = " & tipoEstructura & _
    " AND his_estructura.ternro = " & Tercero & _
    " AND his_estructura.htetdesde <=" & ConvFecha(Fecha_Hasta) & _
    " AND (his_estructura.htethasta >= " & ConvFecha(Fecha_Desde) & _
    " OR his_estructura.htethasta IS NULL)"
    
    OpenRecordset StrSql, rs_Consult
    
    Salida = "FN"
    
    If Not rs_Consult.EOF Then
        Salida = CalcularMapeo(rs_Consult!Estrnro, "T503K", "FN")
    Else
        Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el convenio para el empleado " & Legajo
    End If

    busq_AreaDePersonal = Salida

End Function

'----------------------------------------------------------------
'Busca el cuil del empleado
Function busq_CUIL(Tercero)

    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim Salida
    
    StrSql = " SELECT cuil.nrodoc "
    StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
    StrSql = StrSql & " WHERE tercero.ternro= " & Tercero
           
    OpenRecordset StrSql, rs_Consult
    
    Salida = Space(13)
    
    If Not rs_Consult.EOF Then
       Salida = rs_Consult!NroDoc
    Else
       Flog.Writeline Espacios(Tabulador * 3) & "No se encontró el cuil para el empleado " & Legajo
    End If
    
    busq_CUIL = Salida

End Function

'----------------------------------------------------------------
'Busca la cantidad de hijos que tiene el empleado
Function busq_NumeroDeHijos(ByVal Tercero)

    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim Salida
    
    StrSql = " SELECT count(familiar.ternro) AS total "
    StrSql = StrSql & " FROM familiar"
    StrSql = StrSql & " WHERE parenro = 2 AND familiar.empleado= " & Tercero
           
    OpenRecordset StrSql, rs_Consult
    
    Salida = "0"
    
    If Not rs_Consult.EOF Then
       If IsNull(rs_Consult!total) Then
          Salida = "0"
       Else
          Salida = rs_Consult!total
       End If
    Else
       Salida = "0"
    End If
    
    busq_NumeroDeHijos = Salida

End Function


Public Sub Export_Infotipo_0021()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0021. Datos Familiares.
'
' Autor      : Scarpa D.
' Fecha      : 02/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P0021   PERNR   Número de Personal                 NUMC    8       Sí
'P0021   BEGDA   Fecha inicio validez               DATS    8       Sí                  Especificar la fecha de ingreso a la empresa del empleado, a menos que la fecha de nacimiento de la persona relacionada sea posterior, en cuyo caso se debe especificar ésta.
'P0021   ENDDA   Fecha fin validez                  DATS    8       Sí                  Valor Fijo: '99991231' ó para la clase de registro de familia AR01(Prenatal), una fecha no superior a diez meses comenzando desde la fecha de inicio
'P0021   FAMSA   Clase de registro de Familia       CHAR    4       Sí      T591A
'P0021   FANAM   Apellido                           CHAR    40      Sí
'P0021   FAVOR   Nombre                             CHAR    40      Sí
'P0021   FGBDT   Fecha de nacimiento                DATS    8       Sí
'Q0021   GESC1   Sexo Masculino                     CHAR    1       Sí                  Especificar "X" para "Sí" ó " " para "No". Mutuamente excluyente con GESC2
'Q0021   GESC2   Sexo Femenino                      CHAR    1       Sí                  Especificar "X" para "Sí" ó " " para "No". Mutuamente excluyente con GESC1
'P0394   ASFAX   Asignación por hijo                CHAR    1       Sí      XFELD       Sólo en caso de hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04)
'P0394   DISCP   Discapacitado                      CHAR    1       Sí      XFELD       Sólo en caso de hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04)
'P0394   FEINF   Fecha de informe                   DATS    8       Sí                  "Sólo en caso de hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04)
'                Esta es la fecha en la cual el empleado informa al empleador acerca del nacimiento o adopción de su hijo. Luego de informar al empleador, el empleado tiene un plazo de 90 días para presentar el certificado de nacimiento o adopción. "
'P0394   NADOC   Certificado de nacimiento present. CHAR    1       Sí      XFELD       Sólo en caso de hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04)
'P0394   ADHOS   Adherente a Obra Social            CHAR    1       Sí      XFELD       Campo que aplica sólamente a los familiares que no pertenecen al grupo primario (El grupo primario está conformado por Cónyuge (1),  hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04))

 '- Si se va a informar más de un hijo, se lo debe hacer en orden descendente de edades.
 '- Se debe revisar el registro de gestión, previo a la carga, para verificar que la fecha es anterior o igual al ingreso del empleado
 '- El sistema paga asignación por hijo, sí y sólo si: Está marcado como sí (X) el campo asignación por hijo; tiene registrada fecha de informe y está marcado como sí (X), el campo certificado de nacimiento presentado
 '- Para el caso de hijo discapacitado, el sistema paga, sí y sólo si: Además de cumplir con los criterios para pago por asignación por hijo, está marcado como sí (X) el campo discapacitado
' ------------------------------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Familiares As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor
Dim tipoFam
Dim fechaInicioVal

Dim Hijo2
Dim Hijo6
Dim HijoAR02
Dim HijoAR03
Dim HijoAR04
Dim Conyuge
Dim PreNatal


Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0021"
        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos de familiares para el empleado " & Legajo
        
        'Busco datos que necesito para algunos campos
        Hijo2 = CInt(CalcularMapeoInv("2", "T591A", "-1"))
        Hijo6 = CInt(CalcularMapeoInv("6", "T591A", "-1"))
        HijoAR02 = CInt(CalcularMapeoInv("AR02", "T591A", "-1"))
        HijoAR03 = CInt(CalcularMapeoInv("AR03", "T591A", "-1"))
        HijoAR04 = CInt(CalcularMapeoInv("AR04", "T591A", "-1"))
        Conyuge = CInt(CalcularMapeoInv("1", "T591A", "-1"))
        PreNatal = CInt(CalcularMapeoInv("AR01", "T591A", "-1"))
        
        'Busco los familiares del empleado
        StrSql = " SELECT tercero.terape, tercero.ternom, tercero.tersex, tercero.terfecnac,"
        StrSql = StrSql & " familiar.parenro,familiar.famsalario,familiar.faminc,familiar.famcernac,"
        StrSql = StrSql & " familiar.osocial"
        StrSql = StrSql & " From familiar"
        StrSql = StrSql & " INNER JOIN tercero ON familiar.ternro = tercero.ternro"
        StrSql = StrSql & " Where familiar.Empleado = " & Tercero
        StrSql = StrSql & " ORDER BY parenro, tercero.terfecnac ASC"

        OpenRecordset StrSql, rs_Familiares
        
        Do Until rs_Familiares.EOF
        
            'Estr.   : P0021
            'Campo   : PERNR
            'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
            Salida = Format_StrNro(Legajo, 8, True, "0")
                
            'Estr.   : P0021
            'Campo   : BEGDA
            'Descrip.: Fecha de Inicio de Validez
            '          Si la fecha de nacimiento es mayor a la Fecha de ingreso del empleado
            '          entonces va la fecha de nacimiento
            '          sino la fecha de ingreso del empleado
            If DateDiff("d", CDate(Fecha_Alta_Fase), CDate(rs_Familiares!Terfecnac)) > 0 Then
               fechaInicioVal = rs_Familiares!Terfecnac
               Salida = Salida & Separador & Format_Fecha(rs_Familiares!Terfecnac, 1)
            Else
               fechaInicioVal = Fecha_Alta_Fase
               Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
            End If
            
            'Estr.   : P0021
            'Campo   : ENDDA
            'Descrip.: Fecha Fin de Validez
            '          Si es prenatal
            '          entonces una fecha no superior a diez meses de la fecha de inicio
            '          sino un valor fijo
            If rs_Familiares!parenro = PreNatal Then
               Salida = Salida & Separador & Format_Fecha(DateAdd("d", 1, CDate(fechaInicioVal)), 1)
            Else
               Salida = Salida & Separador & "99991231"
            End If
            
            'Estr.   : P0021
            'Campo   : FAMSA
            'Descrip.: Clase de registro de familiar
            '          Mapeo tabla T591A
            Salida = Salida & Separador & CalcularMapeo(rs_Familiares!parenro, "T591A", "1")
            
            'Estr.   : P0021
            'Campo   : FANAM
            'Descrip.: Apellido
            Salida = Salida & Separador & UCase(rs_Familiares!Terape)
                                
            'Estr.   : P0021
            'Campo   : FAVOR
            'Descrip.: Nombre
            Salida = Salida & Separador & UCase(rs_Familiares!Ternom)
            
            'Estr.   : Q0021
            'Campo   : FGBDT
            'Descrip.: Fecha de nacimiento
            Salida = Salida & Separador & Format_Fecha(rs_Familiares!Terfecnac, 1)
            
            'Estr.   : Q0021
            'Campo   : GESC1
            'Descrip.: Sexo Masculino
            If CInt(rs_Familiares!Tersex) = -1 Then
               Salida = Salida & Separador & "X"
            Else
               Salida = Salida & Separador & " "
            End If
            
            'Estr.   : Q0021
            'Campo   : GESC2
            'Descrip.: Sexo Femenino
            If CInt(rs_Familiares!Tersex) = 0 Then
               Salida = Salida & Separador & "X"
            Else
               Salida = Salida & Separador & " "
            End If
            
            tipoFam = rs_Familiares!parenro
            
            'Estr.   : P0394
            'Campo   : ASFAX
            'Descrip.: Asignacion por hijo
            '          Sólo en caso de hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04)
            If (tipoFam = Hijo2) Or (tipoFam = Hijo6) Or (tipoFam = HijoAR02) Or (tipoFam = HijoAR03) Or (tipoFam = HijoAR04) Then
               If CBool(rs_Familiares!famsalario) Then
                  Salida = Salida & Separador & "X"
               Else
                  Salida = Salida & Separador & " "
               End If
            Else
               Salida = Salida & Separador & " "
            End If

            'Estr.   : P0394
            'Campo   : DISCP
            'Descrip.: Discapacitado
            '          Sólo en caso de hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04)
            If (tipoFam = Hijo2) Or (tipoFam = Hijo6) Or (tipoFam = HijoAR02) Or (tipoFam = HijoAR03) Or (tipoFam = HijoAR04) Then
               If CBool(rs_Familiares!faminc) Then
                  Salida = Salida & Separador & "X"
               Else
                  Salida = Salida & Separador & " "
               End If
            Else
               Salida = Salida & Separador & " "
            End If

            'Estr.   : P0394
            'Campo   : FEINF
            'Descrip.: Fecha de Informe
            '          Sólo en caso de hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04)
            If (tipoFam = Hijo2) Or (tipoFam = Hijo6) Or (tipoFam = HijoAR02) Or (tipoFam = HijoAR03) Or (tipoFam = HijoAR04) Then
                If DateDiff("d", CDate(Fecha_Alta_Fase), CDate(rs_Familiares!Terfecnac)) > 0 Then
                   Salida = Salida & Separador & Format_Fecha(rs_Familiares!Terfecnac, 1)
                Else
                   Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
                End If
            Else
               Salida = Salida & Separador & " "
            End If
            
            'Estr.   : P0394
            'Campo   : NADOC
            'Descrip.: Certificado de nacimiento presentado
            '          Sólo en caso de hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04)
            If (tipoFam = Hijo2) Or (tipoFam = Hijo6) Or (tipoFam = HijoAR02) Or (tipoFam = HijoAR03) Or (tipoFam = HijoAR04) Then
               If CBool(rs_Familiares!famcernac) Then
                  Salida = Salida & Separador & "X"
               Else
                  Salida = Salida & Separador & " "
               End If
            Else
               Salida = Salida & Separador & " "
            End If
            
            'Estr.   : P0394
            'Campo   : ADHOS
            'Descrip.: Adherente a laobra social
            '          Sólo en caso de hijo (2), hijo adoptivo (6), Menor bajo tutela (AR02), Menor tutela temporaria (AR03), hijo de cónyuge (AR04)
            If Not ((Conyuge = tipoFam) Or (tipoFam = Hijo2) Or (tipoFam = Hijo6) Or (tipoFam = HijoAR02) Or (tipoFam = HijoAR03) Or (tipoFam = HijoAR04)) Then
               If Not IsNull(rs_Familiares!osocial) Then
                  If CLng(rs_Familiares!osocial) <> 0 Then
                     Salida = Salida & Separador & "X"
                  Else
                     Salida = Salida & Separador & " "
                  End If
               Else
                  Salida = Salida & Separador & " "
               End If
            Else
               Salida = Salida & Separador & " "
            End If
            
            '-------------------------------------------------------------
            ' Escribo en el erchivo
            fExport.Writeline Salida
        
            rs_Familiares.MoveNext
        Loop
        
        rs_Familiares.Close
        
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_0006()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0006. Direcciones
'
' Autor      : Scarpa D.
' Fecha      : 02/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P0006   PERNR   Número de Personal                 NUMC    8        Sí
'P0006   BEGDA   Fecha Inicio de Validez            DATS    8        Sí                 Valor:"20010101" ó fecha de alta del empleado si es posterior
'P0006   ENDDA   Fecha Fin de Validez               DATS    8        Sí                 Valor Fijo: "99991231"
'P0006   PSTLZ   Código postal                      CHAR    10       Sí
'P0006   ANSSA   Clase de dirección                 CHAR    4        Sí                 Valor Fijo: "0001" (Residencia Habitual)
'P0006   STRAS   Calle                              CHAR    60       Sí
'P0006   HSNMR   Número                             CHAR    6        Sí
'P0006   FLOOR   Planta                             CHAR    6        No      Piso
'P0006   POSTA   Vivienda                           CHAR    6        No      Depto
'P0006   ORT01   Población                          CHAR    40       Sí      Localidad
'P0006   STATE   Región                             CHAR    3        Sí      T005S       Provincias + Capital Federal
'P0006   LAND1   Clave del Pais                     CHAR    3        Sí      T005
'P0006   TELNR   Número de Teléfono                 CHAR    14       No      Este campo es requerido para Colaboradores
'
'- Debe existir al menos un registro para cada empleado
' ------------------------------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Direcciones As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor
Dim tipoFam

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0006"
        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos de las direcciones para el empleado "
        
        'Busco las direcciones del empleado
        StrSql = " SELECT calle,nro,piso,oficdepto,torre,manzana,barrio,email,detdom.provnro,detdom.paisnro "
        StrSql = StrSql & " ,codigopostal,partnro,zonanro,cabdom.domdefault, detdom.domnro, locdesc, cabdom.tidonro "
        StrSql = StrSql & " FROM detdom INNER JOIN cabdom ON detdom.domnro=cabdom.domnro "
        StrSql = StrSql & " LEFT JOIN localidad on localidad.locnro = detdom.locnro"
        StrSql = StrSql & " WHERE cabdom.domdefault = -1 AND cabdom.ternro=" & Tercero

        OpenRecordset StrSql, rs_Direcciones
        
        If Not rs_Direcciones.EOF Then
        
            'Estr.   : P0006
            'Campo   : PERNR
            'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
            Salida = Format_StrNro(Legajo, 8, True, "0")
                
            'Estr.   : P0006
            'Campo   : BEGDA
            'Descrip.: Fecha de Inicio de Validez
            '          Valor:"20010101" ó fecha de alta del empleado si es posterior
            If DateDiff("d", CDate(Fecha_Alta_Fase), CDate("2001-01-01")) > 0 Then
               Salida = Salida & Separador & "20010101"
            Else
               Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
            End If
            
            'Estr.   : P0006
            'Campo   : ENDDA
            'Descrip.: Fecha Fin de Validez
            '          Valor fijo
            Salida = Salida & Separador & "99991231"
            
            'Estr.   : P0006
            'Campo   : PSTLZ
            'Descrip.: Codigo Postal
            Salida = Salida & Separador & Format_StrNro(IIf(IsNull(rs_Direcciones!codigopostal), "", rs_Direcciones!codigopostal), 10, True, " ")
            
            'Estr.   : P0006
            'Campo   : ANSSA
            'Descrip.: Clase de direccion
            '          Valor Fijo: "0001" Residencial Habitual
            Salida = Salida & Separador & "0001"
                                
            'Estr.   : P0006
            'Campo   : STRAS
            'Descrip.: Calle
            Salida = Salida & Separador & rs_Direcciones!calle
                                
            'Estr.   : P0006
            'Campo   : HSNMR
            'Descrip.: Numero
            Salida = Salida & Separador & rs_Direcciones!Nro
                                
            'Estr.   : P0006
            'Campo   : FLOOR
            'Descrip.: Planta
            Salida = Salida & Separador & rs_Direcciones!piso
                                
            'Estr.   : P0006
            'Campo   : POSTA
            'Descrip.: Vivienda
            Salida = Salida & Separador & IIf(IsNull(rs_Direcciones!oficdepto), "", rs_Direcciones!oficdepto)
            
            'Estr.   : P0006
            'Campo   : ORT01
            'Descrip.: Poblacion
            Salida = Salida & Separador & IIf(IsNull(rs_Direcciones!locdesc), "", rs_Direcciones!locdesc)
            
            'Estr.   : P0006
            'Campo   : STATE
            'Descrip.: Region
            '          Mapeo tabla T005S
            Salida = Salida & Separador & CalcularMapeo(rs_Direcciones!provnro, "T005S", "00")
                                
            'Estr.   : P0006
            'Campo   : LAND1
            'Descrip.: Pais
            '          Mapeo tabla T005
            Salida = Salida & Separador & CalcularMapeo(rs_Direcciones!PaisNro, "T005", "AR")
                                
            'Estr.   : P0006
            'Campo   : TELNR
            'Descrip.: Numero de telefono
            '          Este campo es requerido para colaboraciones
            Salida = Salida & Separador & busq_Telefono(rs_Direcciones!Domnro, "")
                                
            ' Escribo en el erchivo
            fExport.Writeline Salida
            
        End If
        
        rs_Direcciones.Close
        
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_0016()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0016. Elementos de Contrato
'
' Autor      : Scarpa D.
' Fecha      : 02/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P0016   PERNR   Número de Personal                 NUMC    8       Sí
'P0016   BEGDA   Fecha Inicio de Validez            DATS    8       Sí                  Fecha de ingreso del empleado
'P0016   ENDDA   Fecha Fin de Validez               DATS    8       Sí                  Valor Fijo: "99991231"
'P0016   CTTYP   Clase de contrato                  CHAR    2       Sí       T547V
'P0016   CTEDT   Limitado al                        DATS    8       Sí                  Indicar este valor si clase de contrato es diferente a 01 Tiempo Indeterminado
'P0016   PRBZT   Período de prueba (cantidad)       DEC     3       No
'P0016   PRBEH   Período de prueba (unidades)       CHAR    3       No       T538A
'P0016   KDGFR   Plazo preaviso empresario          CHAR    2       No       T547T
'P0016   KDGF2   Plazo preaviso empleado            CHAR    2       No       T547T
'
' - Obligatoriamente todos los empleados deben tener un registro de esta estructura
'
' ------------------------------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Contrato As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor
Dim tipoFam
Dim TiempoIndeterminado
Dim TipoCont

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0016"
        tipoEstructura = 18
        TiempoIndeterminado = CInt(CalcularMapeoInv("01", "T547V", "-1"))
        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos de Elementos de Contrato para el empleado " & Legajo
        
        'Busco el contrato del empleado
        'Busco el tipo de convenio y Mapeo
        StrSql = "SELECT tipocont.* FROM his_estructura " & _
        " INNER JOIN tipocont ON his_estructura.estrnro = tipocont.estrnro " & _
        " WHERE his_estructura.tenro = " & tipoEstructura & _
        " AND his_estructura.ternro = " & Tercero & _
        " AND his_estructura.htetdesde <=" & ConvFecha(Fecha_Hasta) & _
        " AND (his_estructura.htethasta >= " & ConvFecha(Fecha_Desde) & _
        " OR his_estructura.htethasta IS NULL)"
        
        OpenRecordset StrSql, rs_Contrato
        
        If Not rs_Contrato.EOF Then
        
            'Estr.   : P0016
            'Campo   : PERNR
            'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
            Salida = Format_StrNro(Legajo, 8, True, "0")
                
            'Estr.   : P0016
            'Campo   : BEGDA
            'Descrip.: Fecha de Inicio de Validez
            '          Valor:"20010101" ó fecha de alta del empleado si es posterior
            Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
            
            'Estr.   : P0016
            'Campo   : ENDDA
            'Descrip.: Fecha Fin de Validez
            '          Valor fijo
            Salida = Salida & Separador & "99991231"
            
            'Estr.   : P0016
            'Campo   : CTTYP
            'Descrip.: Clase de contrato
            '          Mapeo tabla T547V
            Salida = Salida & Separador & CalcularMapeo(rs_Contrato!Estrnro, "T547V", "01")
            
            'Estr.   : P0016
            'Campo   : CTEDT
            'Descrip.: Limitado al
            '          Indicar este valor si la clase de contrato es diferente al 01
            If CInt(rs_Contrato!Estrnro) <> TiempoIndeterminado Then
               StrSql = " SELECT * FROM empleado WHERE ternro=" & Tercero
               
               OpenRecordset StrSql, rs_Consult
               
               If Not rs_Consult.EOF Then
                  If Not IsNull(rs_Consult!empfbajaprev) Then
                     Salida = Salida & Separador & Format_Fecha(rs_Consult!empfbajaprev, 1)
                  Else
                     Salida = Salida & Separador & "        "
                  End If
               Else
                  Salida = Salida & Separador & "        "
               End If
               
               rs_Consult.Close

            Else
               Salida = Salida & Separador & "        "
            End If
            
            'Estr.   : P0016
            'Campo   : PRBZT
            'Descrip.: Periodo de prueba (Cantidad)
            If CInt(rs_Contrato!Estrnro) = TiempoIndeterminado Then
               If CBool(rs_Contrato!tcind) Then
                  Salida = Salida & Separador & ((CInt(rs_Contrato!tcanios) * 12) + CInt(rs_Contrato!tcmeses))
               Else
                  Salida = Salida & Separador & "   "
               End If
            Else
               Salida = Salida & Separador & "   "
            End If
            
            'Estr.   : P0016
            'Campo   : PRBEH
            'Descrip.: Periodo de prueba (Unidades)
            If CInt(rs_Contrato!Estrnro) = TiempoIndeterminado Then
               If CBool(rs_Contrato!tcind) Then
                  Salida = Salida & Separador & "012"
               Else
                  Salida = Salida & Separador & "   "
               End If
            Else
               Salida = Salida & Separador & "   "
            End If
            
            'Estr.   : P0016
            'Campo   : KDGFR
            'Descrip.: Plazo preaviso empresario
            If CInt(rs_Contrato!Estrnro) = TiempoIndeterminado Then
               If CBool(rs_Contrato!tcind) Then
                  Salida = Salida & Separador & "  "
               Else
                  Salida = Salida & Separador & "04"
               End If
            Else
               Salida = Salida & Separador & "   "
            End If
            
            'Estr.   : P0016
            'Campo   : KDGF2
            'Descrip.: Plazo preaviso empleado
            If CInt(rs_Contrato!Estrnro) = TiempoIndeterminado Then
               If CBool(rs_Contrato!tcind) Then
                  Salida = Salida & Separador & "  "
               Else
                  Salida = Salida & Separador & "04"
               End If
            Else
               Salida = Salida & Separador & "   "
            End If
            
            ' Escribo en el erchivo
            fExport.Writeline Salida
            
        End If
        
        rs_Contrato.Close
        
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_0008()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0008. Emolumentos Basicos
'
' Autor      : Scarpa D.
' Fecha      : 03/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'P0008   PERNR   Número de Personal                 NUMC    8       Sí
'P0008   BEGDA   Fecha inicio                       DATS    8       Sí                  "20010101" ó la fecha de alta del empleado si ésta es posterior a 01/01/2001
'P0008   ENDDA   Fecha fin de Validez               DATS    8       Sí                  Valor Fijo: "99991231"
'P0008   TRFAR   Clase de convenio colectivo        CHAR    2       Sí      T510A - T510G   Se debe tener en cuenta la Clase de Convenio (campo anterior) para designar el Área.
'P0008   TRFGB   Area de convenio                   CHAR    2       Sí
'P0008   TRFGR   Grupo profesional                  CHAR    8       Sí      T510A - T510G
'P0008   TRFST   Subgrupo profesional               CHAR    2       Sí
'P0008   ANSAL   Sueldo mensual conformado (Bruto)  CURR    15.2    Sí                  Colocar la suma del Sueldo Basico + A cuenta.

' ---------------------------------------------------------------------------------------------
' Detalle (por cada concepto de nómina)
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'Q0008   LGART   CC-Nómina                          CHAR    4       Sí      T539A
'Q0008   BETRG   Importe                            CURR    13.2    Sí
'Q0008   ANZHL   Número                             DEC     7.2     No

' - Puede tener un máximo de 15 detalles. Los empleados que van a ser confidenciales desde
'   la puesta en marcha del sistema, no deben tener este infotipo cargado. Todo el resto de
'   los empleados tienen que tener al menos un registro en esta conversión
'----------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Contrato As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_Consult2 As New ADODB.Recordset
Dim rs_Conceptos As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor
Dim tipoFam
Dim Convenio
Dim Periodo
Dim Procesos

Dim PERNR
Dim BEGDA
Dim ENDDA
Dim TRFAR
Dim TRFGB
Dim TRFGR
Dim TRFST
Dim ANSAL

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0008"
        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos de Emolumentos Basicos para el empleado "
        
        'Estr.   : P0008
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        PERNR = Format_StrNro(Legajo, 8, True, "0")
        
        'Estr.   : P0008
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        '          Valor:"20010101" ó fecha de alta del empleado si es posterior
        If DateDiff("d", CDate(Fecha_Alta_Fase), CDate("2001-01-01")) > 0 Then
           BEGDA = "20010101"
        Else
           BEGDA = Format_Fecha(Fecha_Alta_Fase, 1)
        End If
        
        'Estr.   : P0008
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor fijo
        ENDDA = "99991231"
        
        'Estr.   : P0008
        'Campo   : TRFAR
        'Descrip.: Clase de Convenio Colectivo
        '          Se debe tener en cuenta la clase de convenio para designar el area
        'Busco el tipo de convenio
        tipoEstructura = 19
        StrSql = "SELECT convenios.* FROM his_estructura " & _
        " INNER JOIN convenios ON his_estructura.estrnro = convenios.estrnro " & _
        " WHERE his_estructura.tenro = " & tipoEstructura & _
        " AND his_estructura.ternro = " & Tercero & _
        " AND his_estructura.htetdesde <=" & ConvFecha(Fecha_Hasta) & _
        " AND (his_estructura.htethasta >= " & ConvFecha(Fecha_Desde) & _
        " OR his_estructura.htethasta IS NULL)"
        
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
           Convenio = rs_Consult!Estrnro
           TRFAR = CalcularMapeo(rs_Consult!Estrnro, "T510A", "FN")
        Else
           Convenio = 0
           TRFAR = "FN"
        End If
        
        rs_Consult.Close
        
        'Estr.   : P0008
        'Campo   : TRFGB
        'Descrip.: Area de Convenio
        TRFGB = "01"
        
        'Estr.   : P0008
        'Campo   : TRFGR
        'Descrip.: Grupo Profesional
        'Busco la Categoria
        StrSql = " SELECT categoria.estrnro" & _
        " From categoria" & _
        " INNER JOIN convenios ON categoria.convnro = convenios.convnro" & _
        " Where convenios.estrnro = " & Convenio
        
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
           TRFGR = CalcularMapeo(rs_Consult!Estrnro, "CATEGORIA", "")
        Else
           TRFGR = ""
        End If
        
        rs_Consult.Close
        
        'Estr.   : P0008
        'Campo   : TRFST
        'Descrip.: SubGrupo Profesional
        TRFST = ""
        
        'Estr.   : P0008
        'Campo   : ANSAL
        'Descrip.: Sueldo mensaul conformado (Bruto)
        '          Colocar la suma del sueldo basico + a cuenta
        StrSql = " SELECT * FROM empleado WHERE ternro=" & Tercero
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
           ANSAL = rs_Consult!empremu
        Else
           ANSAL = "0"
        End If
        
        rs_Consult.Close
        
        '--------------------------------------------
        'Busco el periodo a conciderar
        StrSql = "SELECT * FROM periodo " & _
        " WHERE pliqhasta <= " & ConvFecha(Fecha_Hasta) & _
        " AND pliqdesde >= " & ConvFecha(Fecha_Desde)
        
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Periodo = rs_Consult!pliqnro
        Else
            Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no se encontro el periodo "
            Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
            MyRollbackTrans
            Exit Sub
        End If
        
        rs_Consult.Close
        
        '--------------------------------------------
        'Busco los procesos del periodo
        StrSql = "SELECT * FROM proceso " & _
        " WHERE pliqnro = " & Periodo
        
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            
            Procesos = ""
            Do Until rs_Consult.EOF
               If Procesos = "" Then
                   Procesos = rs_Consult!pronro
               Else
                   Procesos = Procesos & "," & rs_Consult!pronro
               End If
               rs_Consult.MoveNext
            Loop
            
        Else
            Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no se encontraron procesos en el periodo "
            Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
            MyRollbackTrans
            Exit Sub
        End If
        
        rs_Consult.Close
        
        '--------------------------------------------
        'Busco todos los conceptos del empleado definidos en el mapero
            
        StrSql = " SELECT * FROM infotipos_mapeo WHERE tablaref='T539A' "
        
        OpenRecordset StrSql, rs_Conceptos
        
        Do Until rs_Conceptos.EOF
            Salida = PERNR
            Salida = Salida & Separador & BEGDA
            Salida = Salida & Separador & ENDDA
            Salida = Salida & Separador & TRFAR
            Salida = Salida & Separador & TRFGB
            Salida = Salida & Separador & TRFGR
            Salida = Salida & Separador & TRFST
            Salida = Salida & Separador & ANSAL
           
            'Busco el nro. de concepto
            StrSql = " SELECT * FROM concepto WHERE conccod='" & rs_Conceptos!codinterno & "'"
            
            OpenRecordset StrSql, rs_Consult
            
            If Not rs_Consult.EOF Then
                'Busco el monto de los procesos para el periodo
                
                StrSql = " SELECT sum(detliq.dlimonto) AS monto, sum(detliq.dlicant) AS Cantidad "
                StrSql = StrSql & " From cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro "
                StrSql = StrSql & " AND cabliq.empleado = " & Tercero & " AND cabliq.pronro IN (" & Procesos & ")"
                StrSql = StrSql & " Where detliq.concnro = " & rs_Consult!concnro
                
                OpenRecordset StrSql, rs_Consult2
                
                If Not rs_Consult2.EOF Then
                   If Not IsNull(rs_Consult2!Monto) And Not IsNull(rs_Consult2!Cantidad) Then
                        'Estr.   : Q0008
                        'Campo   : LGART
                        'Descrip.: CC-Nomina
                        Salida = Salida & Separador & rs_Conceptos!codexterno
                        
                        'Estr.   : Q0008
                        'Campo   : BETRG
                        'Descrip.: Importe
                        Salida = Salida & Separador & rs_Consult2!Monto
                    
                        'Estr.   : Q0008
                        'Campo   : ANZHL
                        'Descrip.: Numero
                        Salida = Salida & Separador & rs_Consult2!Cantidad
                
                        ' Escribo en el erchivo
                        fExport.Writeline Salida
                    End If
                End If
                
                rs_Consult2.Close
                
            End If
            
            rs_Consult.Close
           
            rs_Conceptos.MoveNext
        Loop
        
        rs_Conceptos.Close
        
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_0009()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0009. Relacion Bancaria
'
' Autor      : Scarpa D.
' Fecha      : 03/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P0009   PERNR   Número de Personal                 NUMC    8       Sí
'P0009   BEGDA   Fecha Inicio d Validez             DATS    8       Sí                  Valor:"20010101" ó fecha de alta del empleado si es posterior
'P0009   ENDDA   Fecha fin de Validez               DATS    8       Sí                  Valor Fijo: "99991231"
'P0009   BNKSA   Clase de Reg de Relación Bancaria  CHAR    4       Sí                  Valor Fijo: "   0" (Relación Bancaria Principal)
'P0009   ZLSCH   Vía de pago                        CHAR    1       Sí       T042Z      Del valor de este campo depende la obligatoriedad de los campos Clave de Banco y Cuenta Bancaria
'P0009   BANKS   País del Banco                     CHAR    3       Si       T005       Es requerido cuando se va a registrar clave de banco
'P0009   BANKL   Clave de Banco                     CHAR    15      Sí       Bancos
'P0009   BANKN   Cuenta Bancaria                    CHAR    18      Si                  Indicar número cta.bancaria del empleado o la palabra CBU si la cuenta del empleado no es del BCO RIO.
'P0009   BKONT   Clave de Control                   CHAR    2       Si                  "CA-Caja de ahorro"
'                                                                                       "CC-Cta Cte"
'                                                                                       "CE-Caja de especial"
'                                                                                       "CS-Cuenta sueldos (CBU)"
'P0009   ZWECK   Destino para transferencias        CHAR    40      Sí                  Colocar N° CBU del empleado
'
' - Para todas las personas que se vayan a incluir en el maestro de personal de Recursos
'   Humanos, se debe indicar un registro de Relación Bancaria. En el caso de no poseer cuenta
'   bancaria, colocar en Vía de pago "C" y dejar los campos de datos del Banco en blanco
'
' ------------------------------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Cta_Bancaria As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor
Dim tipoFam
Dim TiempoIndeterminado
Dim tieneCtaBancaria As Boolean

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0009"
                        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos de Relacion Bancaria para el empleado " & Legajo
        
        'Busco la cuenta bancaria del empleado
        StrSql = " SELECT fpagbanc,ctabnro,ctabancaria.fpagnro, ctabcbu, banco.estrnro " & _
        " From ctabancaria " & _
        " INNER JOIN formapago ON formapago.fpagnro = ctabancaria.fpagnro " & _
        " INNER JOIN banco ON banco.ternro = ctabancaria.banco " & _
        " Where ctabancaria.ternro = " & Tercero & _
        " AND ctabestado = -1 "
        
        OpenRecordset StrSql, rs_Cta_Bancaria
        
        If rs_Cta_Bancaria.EOF Then
           tieneCtaBancaria = False
        Else
           tieneCtaBancaria = CBool(rs_Cta_Bancaria!fpagbanc)
        End If
        
        'Estr.   : P0009
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        Salida = Format_StrNro(Legajo, 8, True, "0")
            
        'Estr.   : P0009
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        '          Valor:"20010101" ó fecha de alta del empleado si es posterior
        If DateDiff("d", CDate(Fecha_Alta_Fase), CDate("2001-01-01")) > 0 Then
           Salida = Salida & Separador & "20010101"
        Else
           Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
        End If
        
        'Estr.   : P0009
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor fijo
        Salida = Salida & Separador & "99991231"
            
        'Estr.   : P0009
        'Campo   : BNKSA
        'Descrip.: Clase de Reg de relacion bancaria
        '          Valor fijo 0
        Salida = Salida & Separador & "0"
        
        If tieneCtaBancaria Then
            'Estr.   : P0009
            'Campo   : ZLSCH
            'Descrip.: Via de Pago
            '          Mapeo tabla T042Z
            Salida = Salida & Separador & "0"
        
            'Estr.   : P0009
            'Campo   : BANKS
            'Descrip.: Pais del banco
            '          Mapeo tabla T005
            Salida = Salida & Separador & "AR"
        
            'Estr.   : P0009
            'Campo   : BANKL
            'Descrip.: Clave de Banco
            '          Mapeo tabla Bancos
            Salida = Salida & Separador & CalcularMapeo(rs_Cta_Bancaria!Estrnro, "Bancos", "")
        
            'Estr.   : P0009
            'Campo   : BANKN
            'Descrip.: Cuenta bancaria
            '          Mapeo tabla Bancos
            If CalcularMapeo(rs_Cta_Bancaria!Estrnro, "Bancos", "") = "072000" Then
               Salida = Salida & Separador & rs_Cta_Bancaria!ctabnro
            Else
               Salida = Salida & Separador & "CBU"
            End If
                                
            'Estr.   : P0009
            'Campo   : BKONT
            'Descrip.: Clave de control
            '          Mapeo tabla FormaPago
            Salida = Salida & Separador & CalcularMapeo(rs_Cta_Bancaria!fpagnro, "FormaPago", "CC")
        
            'Estr.   : P0009
            'Campo   : ZWECK
            'Descrip.: Destino para transferencias (CBU)
            If IsNull(rs_Cta_Bancaria!ctabcbu) Then
               Salida = Salida & Separador & Space(40)
            Else
               Salida = Salida & Separador & rs_Cta_Bancaria!ctabcbu
            End If
        
        Else
            'Estr.   : P0009
            'Campo   : ZLSCH
            'Descrip.: Via de Pago
            '          Mapeo tabla T042Z
            Salida = Salida & Separador & "C"
        
            'Estr.   : P0009
            'Campo   : BANKS
            'Descrip.: Pais del banco
            '          Mapeo tabla T005
            Salida = Salida & Separador & "AR"
        
            'Estr.   : P0009
            'Campo   : BANKL
            'Descrip.: Clave de Banco
            '          Mapeo tabla Bancos
            Salida = Salida & Separador & "      "
        
            'Estr.   : P0009
            'Campo   : BANKN
            'Descrip.: Cuenta bancaria
            '          Mapeo tabla Bancos
            Salida = Salida & Separador & Space(18)
        
            'Estr.   : P0009
            'Campo   : BKONT
            'Descrip.: Clave de control
            '          Mapeo tabla FormaPago
            Salida = Salida & Separador & "  "
        
            'Estr.   : P0009
            'Campo   : ZWECK
            'Descrip.: Destino para transferencias (CBU)
            Salida = Salida & Separador & Space(40)
        
        End If
        
        rs_Cta_Bancaria.Close
        
        ' Escribo en el erchivo
        fExport.Writeline Salida
        
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_0185()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0185. Identificacion Personal
'
' Autor      : Scarpa D.
' Fecha      : 06/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P0185   PERNR   Nro. Personal                      NUMC    8       Sí
'P0185   BEGDA   Fecha inicio validez               DATS    8       Sí                  Es "20010101" o la fecha de ingreso si es posterior
'P0185   ENDDA   Fecha fin validez                  DATS    8       Sí                  Valor Fijo: "99991231"
'P0185   ICTYP   Tipo de documento de identidad     CHAR    2       Sí      T5R05
'P0185   ICNUM   Número del documento               CHAR    30      Sí
'
' ------------------------------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Documentos As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor
Dim tipoDoc

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0185"
                        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos de Identificacion Personal para el empleado " & Legajo
        
        'Busco los documentos del empleado
        StrSql = " SELECT * " & _
        " From ter_doc " & _
        " Where ternro = " & Tercero & _
        " AND tidnro < 5 "
        
        OpenRecordset StrSql, rs_Documentos
        
        If Not rs_Documentos.EOF Then
            'Estr.   : P0185
            'Campo   : PERNR
            'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
            Salida = Format_StrNro(Legajo, 8, True, "0")
                
            'Estr.   : P0185
            'Campo   : BEGDA
            'Descrip.: Fecha de Inicio de Validez
            '          Valor:"20010101" ó fecha de alta del empleado si es posterior
            If DateDiff("d", CDate(Fecha_Alta_Fase), CDate("2001-01-01")) > 0 Then
               Salida = Salida & Separador & "20010101"
            Else
               Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
            End If
            
            'Estr.   : P0185
            'Campo   : ENDDA
            'Descrip.: Fecha Fin de Validez
            '          Valor fijo
            Salida = Salida & Separador & "99991231"
        
            'Estr.   : P0185
            'Campo   : ICTYP
            'Descrip.: Tipo de Documento de Identidad
            '          Mapeo tabla T5R05
            tipoDoc = CalcularMapeo(rs_Documentos!tidnro, "T5R05", "")
            
            If tipoDoc <> "" Then
                Salida = Salida & Separador & tipoDoc
            
                'Estr.   : P0185
                'Campo   : ICNUM
                'Descrip.: Numero del Documento
                Salida = Salida & Separador & rs_Documentos!NroDoc
                
                'Escribo en el erchivo
                fExport.Writeline Salida
            End If
        
        End If
        
        rs_Documentos.Close
        
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_2006()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 2006. Derechos pendientes de tiempos
'
' Autor      : Scarpa D.
' Fecha      : 06/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P2006   PERNR   Nro. Personal                      NUMC    8       Sí
'P2006   BEGDA   Fecha de inicio de validez         DATS    8       Sí                  Fecha de inicio de validez del contingente
'P2006   ENDDA   Fecha de fin de validez            DATS    8       Sí                  Fecha de fin de validez del contingente
'P2006   KTART   Tipo de Contingente de Absentismo  NUMC    2       Sí      T556A       De acuerdo al grupo al que pertenezca la UDN
'P2006   ANZHL   Cantidad de contingentes de tiemp  DEC     10.5    Sí                  Expresado en días
'P2006   DESTA   Fecha de inicio de liquidación de  DATS    8       Sí
'P2006   DEEND   Fecha final de liquidación de con  DATS    8       Sí
' ------------------------------------------------------------------------------------------------------------------

Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor
Dim FDesde
Dim FHasta
Dim Lista_Vac
Dim diasCorCant
Dim diasTom

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "2006"
                        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos de Derechos pendientes de tiempos para el empleado " & Legajo
        
        
        '------------------------------------------------------------------------------------------------------
        ' Busco los periodos de vacaciones en el rango de fecha
        '------------------------------------------------------------------------------------------------------
        FDesde = "01/01/" & Year(Fecha_Desde)
        FHasta = "31/12/" & Year(Fecha_Desde)
        
        StrSql = " SELECT * FROM vacacion "
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " (  (vacfecdesde >= " & ConvFecha(FDesde)
        StrSql = StrSql & " and vacfechasta <= " & ConvFecha(FHasta) & ") "
        StrSql = StrSql & " or (vacfecdesde <  " & ConvFecha(FDesde)
        StrSql = StrSql & " and vacfechasta <= " & ConvFecha(FHasta)
        StrSql = StrSql & " and vacfechasta >= " & ConvFecha(FDesde) & ") "
        StrSql = StrSql & " or (vacfecdesde >= " & ConvFecha(FDesde)
        StrSql = StrSql & " and vacfechasta >  " & ConvFecha(FHasta)
        StrSql = StrSql & " and vacfecdesde <= " & ConvFecha(FHasta) & ") "
        StrSql = StrSql & " or (vacfecdesde <  " & ConvFecha(FDesde)
        StrSql = StrSql & " and vacfechasta >  " & ConvFecha(FHasta) & ") ) "
        
        Lista_Vac = ""
        
        OpenRecordset StrSql, rs_Consult
        
        If rs_Consult.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró un periodo de vacaciones"
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Do Until rs_Consult.EOF
            If Lista_Vac = "" Then
               Lista_Vac = rs_Consult!vacnro
            Else
               Lista_Vac = Lista_Vac & "," & rs_Consult!vacnro
            End If
        
            rs_Consult.MoveNext
        Loop
        
        rs_Consult.Close
        
        '------------------------------------------------------------------------------------------------------
        ' Busco los dias correspondientes
        '------------------------------------------------------------------------------------------------------
        
        StrSql = " SELECT SUM(vdiascorcant) diascorcant FROM vacdiascor "
        StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiascor.vacnro "
        StrSql = StrSql & " WHERE vacdiascor.ternro = " & Tercero
        StrSql = StrSql & "   AND vacdiascor.vacnro IN (" & Lista_Vac & ") "
        
        OpenRecordset StrSql, rs_Consult
        
        If rs_Consult.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontraron dias correspondientes para el empleado " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        Else
            diasCorCant = rs_Consult!diasCorCant
        End If
        
        rs_Consult.Close
        
        '------------------------------------------------------------------------------------------------------
        ' Busco los dias tomados
        '------------------------------------------------------------------------------------------------------
        
        StrSql = "SELECT elcantdias, elfechadesde, elfechahasta "
        StrSql = StrSql & "FROM emp_lic INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro "
        StrSql = StrSql & "LEFT JOIN vacnotif ON vacnotif.emp_licnro = emp_lic.emp_licnro "
        StrSql = StrSql & "WHERE empleado = " & Tercero
        StrSql = StrSql & "  AND tdnro = 2 "
        StrSql = StrSql & "  AND licestnro = 2 "
        StrSql = StrSql & " AND ( (elfechadesde >= " & ConvFecha(FDesde)
        StrSql = StrSql & " and elfechahasta <= " & ConvFecha(FHasta) & ") "
        StrSql = StrSql & " or (elfechadesde <  " & ConvFecha(FDesde)
        StrSql = StrSql & " and elfechahasta <= " & ConvFecha(FHasta)
        StrSql = StrSql & " and elfechahasta >= " & ConvFecha(FDesde) & ") "
        StrSql = StrSql & " or (elfechadesde >= " & ConvFecha(FDesde)
        StrSql = StrSql & " and elfechahasta >  " & ConvFecha(FHasta)
        StrSql = StrSql & " and elfechadesde <= " & ConvFecha(FHasta) & ") "
        StrSql = StrSql & " or (elfechadesde <  " & ConvFecha(FDesde)
        StrSql = StrSql & " and elfechahasta >  " & ConvFecha(FHasta) & ") ) "
        
        OpenRecordset StrSql, rs_Consult
        
        diasTom = 0
        Do Until rs_Consult.EOF
           diasTom = diasTom + CInt(rs_Consult!elcantdias)
        
           rs_Consult.MoveNext
        Loop
        
        rs_Consult.Close
        
        'Estr.   : P2006
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        Salida = Format_StrNro(Legajo, 8, True, "0")
            
        'Estr.   : P2006
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
        
        'Estr.   : P2006
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor fijo
        Salida = Salida & Separador & "99991231"
    
        'Estr.   : P2006
        'Campo   : KTART
        'Descrip.: Tipo de Contingente de Absentismo
        '          Mapeo tabla T556A
        Salida = Salida & Separador & "32"
        
        'Estr.   : P2006
        'Campo   : ANZHL
        'Descrip.: Cant. de contingentes de tiempos de personal
        '          Expresados en dias
        Salida = Salida & Separador & (diasCorCant - diasTom)
        
        'Estr.   : P2006
        'Campo   : DESTA
        'Descrip.: Fecha de inicio de la liquidacion
        Salida = Salida & Separador & Space(8)
        
        'Estr.   : P2006
        'Campo   : DEEND
        'Descrip.: Fecha final de la liquidacion
        Salida = Salida & Separador & Space(8)
        
        ' Escribo en el erchivo
        fExport.Writeline Salida
        
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_0390()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0390. Impuesto a las ganacias - deducciones
'
' Autor      : Scarpa D.
' Fecha      : 06/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P0390   PERNR   Número de Personal                 NUMC    8       Sí
'P0390   BEGDA   Fecha inicio validez               DATS    8       Sí                  Depende de la deducción. No anterior a 20010101
'P0390   ENDDA   Fecha fin validez                  DATS    8       Sí                  Valor Fijo: 99991231
'P0390   LGART   Concepto de Nómina                 CHAR    4       Sí       T512W      Concepto deducible de Impuesto a las Ganancias
'P0390   IMDED   Importe                            CURR    13.2    Sí                  Campo no obligatorio para cargas de familia
'P0390   EITXT   Unidad de Frecuencia               CHAR    8       Sí       UNFREC     No cargar este campo para cargas de familia.
'P0390   NOMBR   Persona jurídica                   CHAR    30      No                  Nombre de entidades que reciben del empleado pagos deducibles del impuesto a las ganancias
'P0390   ICTYP   Tipo de Documento                  CHAR    2       No       T5R05      De la persona jurídica
'P0390   ICNUM   Número de Documento                CHAR    20      No                  De la persona jurídica
'
' ------------------------------------------------------------------------------------------------------------------

Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_Consult2 As New ADODB.Recordset
Dim rs_Items As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean
Dim Periodo
Dim Procesos

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor

Dim PERNR
Dim BEGDA
Dim ENDDA

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0390"
                        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos de Impuesto a las ganacias - deducciones para el empleado " & Legajo
        
        'Estr.   : P0390
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        PERNR = Format_StrNro(Legajo, 8, True, "0")
        
        'Estr.   : P0390
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        '          Depende de la deducción. No anterior a 20010101
        BEGDA = "20050101"
        'If DateDiff("d", CDate(Fecha_Alta_Fase), CDate("2001-01-01")) > 0 Then
        '   BEGDA = "20010101"
        'Else
        '   BEGDA = Format_Fecha(Fecha_Alta_Fase, 1)
        'End If
        
        'Estr.   : P0390
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor fijo
        ENDDA = "99991231"
        
        '---------------------------------------------------------
        'Busco todos los items del empleado definidos en el mapeo
            
        StrSql = " SELECT * FROM infotipos_mapeo WHERE tablaref='T512W' "
        
        OpenRecordset StrSql, rs_Items
        
        Do Until rs_Items.EOF
            Salida = PERNR
            Salida = Salida & Separador & BEGDA
            Salida = Salida & Separador & ENDDA
           
            'Busco el item
            StrSql = " SELECT * FROM desmen WHERE empleado = " & Tercero
            StrSql = StrSql & " AND itenro=" & rs_Items!codinterno
            StrSql = StrSql & " AND desano=" & Year(Fecha_Desde)
            StrSql = StrSql & " ORDER BY desfecdes DESC "
            
            OpenRecordset StrSql, rs_Consult
            
            If Not rs_Consult.EOF Then
                    
                'Estr.   : P0390
                'Campo   : LGART
                'Descrip.: CC-Nomina
                Salida = Salida & Separador & rs_Items!codexterno
                
                'Estr.   : P0390
                'Campo   : IMDED
                'Descrip.: Importe
                Salida = Salida & Separador & rs_Consult!desmondec
            
                'Estr.   : P0390
                'Campo   : EITXT
                'Descrip.: Unidad de frecuencia
                Salida = Salida & Separador & "Años"
                
                'Estr.   : P0390
                'Campo   : NOMBR
                'Descrip.: Persona juridica
                Salida = Salida & Separador & rs_Consult!desrazsoc
                
                'Estr.   : P0390
                'Campo   : ICTYP
                'Descrip.: Tipo de documento
                Salida = Salida & Separador & "07"
                
                'Estr.   : P0390
                'Campo   : ICNUM
                'Descrip.: Numero de documento
                Salida = Salida & Separador & rs_Consult!descuit
            
            Else
            
                'Estr.   : P0390
                'Campo   : LGART
                'Descrip.: CC-Nomina
                Salida = Salida & Separador & rs_Items!codexterno
                
                'Estr.   : P0390
                'Campo   : IMDED
                'Descrip.: Importe
                Salida = Salida & Separador & ""
            
                'Estr.   : P0390
                'Campo   : EITXT
                'Descrip.: Unidad de frecuencia
                Salida = Salida & Separador & ""
                
                'Estr.   : P0390
                'Campo   : NOMBR
                'Descrip.: Persona juridica
                Salida = Salida & Separador & ""
                
                'Estr.   : P0390
                'Campo   : ICTYP
                'Descrip.: Tipo de documento
                Salida = Salida & Separador & ""
                
                'Estr.   : P0390
                'Campo   : ICNUM
                'Descrip.: Numero de documento
                Salida = Salida & Separador & ""
            
            End If
            
            'Escribo en el erchivo
            fExport.Writeline Salida
            
            rs_Consult.Close
           
            rs_Items.MoveNext
        Loop
        
        rs_Items.Close
            
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_0392()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0392. Seguridad Social
'
' Autor      : Scarpa D.
' Fecha      : 06/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P0392   PERNR   Nro. Personal                      NUMC    8       Sí
'P0392   BEGDA   Fecha Inicio de Validez            DATS    8       Sí                  Es 20010101 o la fecha de ingreso si es posterior
'P0392   ENDDA   Fecha Fin de Validez               DATS    8       Sí                  Valor Fijo: 99991231
'P0392   OBRAS   Obra Social                        CHAR    6       Sí      T7AR34
'P0392   PLANS   Plan de Obra Social                CHAR    10      No                  Colocar Plan de las obras sociales LUIS PASTEUR y OSPOCE-CONSOLIDAR SALUD / OSDO -CONSOLIDAR SALUD
'P0392   TPUOS   Opción paga diferencia             CHAR    1       Sí      PAR_TPUOS   Valor Fijo: " " (Vacío)
'P0392   SYJUB   Sistema Jubilacion                 CHAR    1       Sí      PAR_SYJUB
'P0392   CAFJP   AFJP                               CHAR    4       Sí      T7AR36      Requerido sólo si es Sist. de Capitaliz.
'P0392   TYACT   Actividad del empleado             CHAR    2       Sí      T7AR38
'P0392   ASPCE   Agrupación de empleados para cont. CHAR    4       NO      T7AR26
'
'  - Empleados con grupo de personal 9 (Externos) no deben tener infotipo de Seguridad
'    Social cargada
'
' ------------------------------------------------------------------------------------------------------------------

Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor
Dim SistJubTipo

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0392"
                        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando datos de Seguridad Social para el empleado " & Legajo
        
        'Si el grupo del personal es 9 (Externos) no debe tener infortipo de seguridad social
        If CInt(busq_GrupoDePersonal(Tercero)) <> 9 Then
            'Estr.   : P0392
            'Campo   : PERNR
            'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
            Salida = Format_StrNro(Legajo, 8, True, "0")
            
            'Estr.   : P0392
            'Campo   : BEGDA
            'Descrip.: Fecha de Inicio de Validez
            '          Valor:"20010101" ó fecha de alta del empleado si es posterior
            If DateDiff("d", CDate(Fecha_Alta_Fase), CDate("2001-01-01")) > 0 Then
               Salida = Salida & Separador & "20010101"
            Else
               Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
            End If
            
            'Estr.   : P0392
            'Campo   : ENDDA
            'Descrip.: Fecha Fin de Validez
            '          Valor fijo
            Salida = Salida & Separador & "99991231"
            
            'Estr.   : P0392
            'Campo   : OBRAS
            'Descrip.: Obra Social
            '          Mapeo con tabla T7AR34
            Salida = Salida & Separador & busq_Estructura(Tercero, 17, "T7AR34", Space(6))
            
            'Estr.   : P0392
            'Campo   : PLANS
            'Descrip.: Plan de Obra Social
            '          No requerido
            '          Colocar Plan de las obras sociales LUIS PASTEUR y OSPOCE-CONSOLIDAR SALUD / OSDO -CONSOLIDAR SALUD
            Salida = Salida & Separador & Space(10)
            
            'Estr.   : P0392
            'Campo   : TPUOS
            'Descrip.: Opcion paga diferenciada
            '          Valor fijo " "
            Salida = Salida & Separador & Space(1)
            
            'Estr.   : P0392
            'Campo   : SYJUB
            'Descrip.: Sistema jubilacion
            '          Mapeo tabla PAR_SYJUB
            tipoEstructura = 15
            StrSql = "SELECT cajjub.* FROM his_estructura " & _
            " INNER JOIN cajjub ON his_estructura.estrnro = cajjub.estrnro " & _
            " WHERE his_estructura.tenro = " & tipoEstructura & _
            " AND his_estructura.ternro = " & Tercero & _
            " AND his_estructura.htetdesde <=" & ConvFecha(Fecha_Hasta) & _
            " AND (his_estructura.htethasta >= " & ConvFecha(Fecha_Desde) & _
            " OR his_estructura.htethasta IS NULL)"
            
            OpenRecordset StrSql, rs_Consult
            
            If Not rs_Consult.EOF Then
                SistJubTipo = rs_Consult!ticnro
            Else
                Flog.Writeline Espacios(Tabulador * 2) & "No se encontró la caja de jubilacion para el empleado " & Legajo
                Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
                MyRollbackTrans
                Exit Sub
            End If
            
            rs_Consult.Close
            
            If CInt(SistJubTipo) = 1 Then
               Salida = Salida & Separador & "0"
            Else
               Salida = Salida & Separador & "1"
            End If
            
            'Estr.   : P0392
            'Campo   : CAFJP
            'Descrip.: AFJP
            '          Mapeo tabla T7AR36
            If CInt(SistJubTipo) = 2 Then
               Salida = Salida & Separador & busq_Estructura(Tercero, 15, "T7AR36", "")
            Else
               Salida = Salida & Separador & Space(4)
            End If
            
            'Estr.   : P0392
            'Campo   : TYACT
            'Descrip.: Actividad del empleado
            '          Mapeo tabla T7AR38
            Salida = Salida & Separador & busq_Estructura(Tercero, 29, "T7AR38", "")
            
            'Estr.   : P0392
            'Campo   : ASPCE
            'Descrip.: Agrup. de empleados para contrib. de seg. social
            '          Mapeo tabla T7AR26
            Salida = Salida & Separador & Space(4)
    
            ' Escribo en el erchivo
            fExport.Writeline Salida
        
        End If
            
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub





Public Sub Export_Infotipo_0041()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0041. Datos de Fechas
'
' Autor      : Scarpa D.
' Fecha      : 06/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P0041   PERNR   Nro. Personal                      NUMC    8       Sí
'P0041   BEGDA   Fecha Inicio Validez               DATS    8       Sí                  Valor Fijo: "20010101"
'P0041   ENDDA   Fecha Fin Validez                  DATS    8       Sí                  Valor Fijo: "99991231"
'P0041   DARnn   Clase de Fecha                     CHAR    2       Sí      T548Y       Indicar las dos clases de fecha
'P0041   DATnn   Fecha                              DATS    8       Sí
'
' ------------------------------------------------------------------------------------------------------------------

Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0041"
                        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando Datos de Fechas para el empleado " & Legajo
            
        'Estr.   : P0041
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        Salida = Format_StrNro(Legajo, 8, True, "0")
        
        'Estr.   : P0041
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        '          Valor:"20010101" ó fecha de alta del empleado si es posterior
        Salida = Salida & Separador & "20010101"
        
        'Estr.   : P0041
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor fijo
        Salida = Salida & Separador & "99991231"
            
        'Estr.   : P0041
        'Campo   : DAR01
        'Descrip.: Clase de Fecha
        '          Fecha de alta téc.  (para VACACIONES)
        Salida = Salida & Separador & "01"
            
        'Estr.   : P0041
        'Campo   : DAT01
        'Descrip.: Fecha
        StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
        StrSql = StrSql & " AND vacaciones = -1 "
        StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
        StrSql = StrSql & " ORDER BY altfec ASC "
        
        OpenRecordset StrSql, rs_Consult
        
        If rs_Consult.EOF Then
            Salida = Salida & Separador & "20010101"
        Else
           Salida = Salida & Separador & Format_Fecha(rs_Consult!altfec, 1)
        End If
        
        rs_Consult.Close

        'Estr.   : P0041
        'Campo   : DAR02
        'Descrip.: Clase de Fecha
        '          Antiguedad  (Antigüedad RECONOCIDA)
        Salida = Salida & Separador & "09"
            
        'Estr.   : P0041
        'Campo   : DAT02
        'Descrip.: Fecha
        StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
        StrSql = StrSql & " AND real = -1 "
        StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
        StrSql = StrSql & " ORDER BY altfec ASC "
        
        OpenRecordset StrSql, rs_Consult
        
        If rs_Consult.EOF Then
            Salida = Salida & Separador & "20010101"
        Else
           Salida = Salida & Separador & Format_Fecha(rs_Consult!altfec, 1)
        End If
        
        rs_Consult.Close
            
            
        ' Escribo en el erchivo
        fExport.Writeline Salida
            
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_0032()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0032. Datos Internos de la empresa
'
' Autor      : Scarpa D.
' Fecha      : 07/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo       Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------
'P0032   PERNR       Número de Personal                 NUMC    8       Sí
'P0032   BEGDA       Fecha de inicio de validez         DATS    8       Sí                  Valor Fijo: "20010101"
'P0032   ENDDA       Fecha fin de Validez               DATS    8       Sí                  Valor fijo: '99991231'
'P0032   PNALT       N° de personal anterior            CHAR    12      Sí
'P0032   ZZEDIFICIO  Dirección Edificio                 CHAR    4       Sí                  Dejar en Blanco
'P0032   ZZAGENCIA   Código Agencia PE                  CHAR    3       Sí      ZHR_NAM_AGENCIA Campo requerido si el personal es eventual
'
' ------------------------------------------------------------------------------------------------------------------

Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0032"
                        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando Datos Internos de la empresa para el empleado " & Legajo
            
        'Estr.   : P0032
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        Salida = Format_StrNro(Legajo, 8, True, "0")
        
        'Estr.   : P0032
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        '          Valor:"20010101"
        Salida = Salida & Separador & "20010101"
        
        'Estr.   : P0032
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor fijo
        Salida = Salida & Separador & "99991231"
            
        'Estr.   : P0032
        'Campo   : PNALT
        'Descrip.: Nro. de personal anterior
        Salida = Salida & Separador & Format_StrNro(Legajo, 8, True, "0")
            
        'Estr.   : P0032
        'Campo   : ZZEDIFICIO
        'Descrip.: Direccion edificio
        Salida = Salida & Separador & Space(4)


        'Estr.   : P0032
        'Campo   : ZZAGENCIA
        'Descrip.: Codigo Agencia PE
        '          Mapeo tabla ZHR_AGENCI
        Salida = Salida & Separador & busq_Estructura(Tercero, 28, "ZHR_AGENCI", Space(3))
            
        ' Escribo en el erchivo
        fExport.Writeline Salida
            
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub








Public Sub Export_Infotipo_0057()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 0057. Asociaciones
'
' Autor      : Scarpa D.
' Fecha      : 07/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc. Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------P0057   PERNR   Número de Personal  NUMC    8   Sí
'P0057   BEGDA   Fecha de inicio de validez         DATS    8       Sí                  Valor fijo: '20011001" o la fecha de ingreso si es posterior.
'P0057   ENDDA   Fecha fin de Validez               DATS    8       Sí                  Valor fijo: '99991231' o fecha fin de validez
'P0057   MGART   Clase de beneficiario              CHAR    4       Sí  T591A
'P0057   LGART   CC-Nómina                          CHAR    4       Sí  T591A
'
' - En esta conversión se piden las asociaciones que se quieran tener en cuenta a partir de la liquidación de noviembre
'
' ------------------------------------------------------------------------------------------------------------------

Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "0057"
                        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando Datos de Asociasiones para el empleado " & Legajo
            
        'Estr.   : P0057
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        Salida = Format_StrNro(Legajo, 8, True, "0")
        
        'Estr.   : P0057
        'Campo   : BEGDA
        'Descrip.: Fecha de Inicio de Validez
        '          Valor:"20010101" ó fecha de alta del empleado si es posterior
        If DateDiff("d", CDate(Fecha_Alta_Fase), CDate("2001-01-01")) > 0 Then
           Salida = Salida & Separador & "20010101"
        Else
           Salida = Salida & Separador & Format_Fecha(Fecha_Alta_Fase, 1)
        End If
        
        'Estr.   : P0057
        'Campo   : ENDDA
        'Descrip.: Fecha Fin de Validez
        '          Valor fijo
        Salida = Salida & Separador & "99991231"
            
        'Estr.   : P0057
        'Campo   : MGART
        'Descrip.: Clase de beneficiario
        '          Mapeo tabla T591A
        
            
        'Estr.   : P0057
        'Campo   : LGART
        'Descrip.: CC-Nomina
        '          Mapeo tabla T591A

            
        ' Escribo en el erchivo
        fExport.Writeline Salida
            
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub


Public Sub Export_Infotipo_9999()
' ---------------------------------------------------------------------------------------------
' Descripcion:  Infotipo 9999. Conversion de acumulados historicos de la liquidacion
'
' Autor      : Scarpa D.
' Fecha      : 07/12/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Estruc.     Campo   Descripción                        Tipo    Long.   Requer. Tab.Ref.    Observaciones
'----------------------------------------------------------------------------------------------P0057   PERNR   Número de Personal  NUMC    8   Sí
'HRLOADPAY   PERNR   Nro. Personal                      NUMC    8       Sí
'HRLOADPAY   BEGDA   Fecha Liquidación                  DATS    8       Sí                  Ultimo día del período de liquidación
'HRLOADPAY   ENDDA   Fecha Liquidación                  DATS    8       Sí                  Ultimo día del período de liquidación (debe ser igual que la fecha BEGDA)
'HRLOADPAY   LGART   CC-nómina                          CHAR    4       Sí       Conversiones
'HRLOADPAY   BETPE   Importe por unidad                 CHAR    21      No                  Es requerido si la especificación del campo así lo requiere
'HRLOADPAY   ANZHL   Ctd.                               CHAR    21      No                  Es requerido si la especificación del campo así lo requiere
'HRLOADPAY   BETRG   Importe                            CHAR    21      Sí
'
' ------------------------------------------------------------------------------------------------------------------

Dim rs_Fases As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_Acum  As New ADODB.Recordset

Dim tipoEstructura As Long
Dim NroInfotipo As String
Dim Valido As Boolean
Dim Fecha_Alta_Fase As Date
Dim PrimerCampo As Boolean

Dim Aux
Dim Aux_Linea As String

Dim Parametro
Dim Salida
Dim valor
Dim PERNR
Dim BEGDA
Dim ENDDA
Dim PeriodoHasta
Dim Periodo

Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo
Flog.Writeline

'Validacion
'El empleado debe estar activo a la fecha de corte
'Busco la ultima fase activa
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND real = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha_Hasta)
StrSql = StrSql & " AND ( bajfec is null OR bajfec >= " & ConvFecha(Fecha_Hasta) & ")"

OpenRecordset StrSql, rs_Fases

If Not rs_Fases.EOF Then
    Valido = True
    Fecha_Alta_Fase = rs_Fases!altfec
Else
    Valido = False
End If

If Valido Then
    MyBeginTrans
    
        NroInfotipo = "9999"
                        
        'levanto los datos del tercero correspondientes al legajo
        StrSql = "SELECT * FROM tercero WHERE ternro = " & Tercero
        
        If rs_Tercero.State = adStateOpen Then
           rs_Tercero.Close
        End If
        
        OpenRecordset StrSql, rs_Tercero
        If rs_Tercero.EOF Then
            Flog.Writeline Espacios(Tabulador * 2) & "No se encontró el tercero del legajo " & Legajo
            Flog.Writeline Espacios(Tabulador * 2) & "Infotipo Abortado"
            MyRollbackTrans
            Exit Sub
        End If
        
        Flog.Writeline Espacios(Tabulador * 2) & "Infotipo " & NroInfotipo & " - Generando Datos de Conversion de acumulados historicos de la liquidacion para el empleado " & Legajo
        
        '--------------------------------------------
        'Busco el periodo a conciderar
        StrSql = "SELECT * FROM periodo " & _
        " WHERE pliqhasta <= " & ConvFecha(Fecha_Hasta) & _
        " AND pliqdesde >= " & ConvFecha(Fecha_Desde)
        
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Periodo = rs_Consult!pliqnro
            PeriodoHasta = rs_Consult!pliqhasta
        Else
            Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no se encontro el periodo "
            Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
            MyRollbackTrans
            Exit Sub
        End If
        
        rs_Consult.Close
        
        
        'Estr.   : P9999
        'Campo   : PERNR
        'Descrip.: Numero de Personal - [ RHPro (Legajo del empleado)]
        PERNR = Format_StrNro(Legajo, 8, True, "0")
        
        'Estr.   : P9999
        'Campo   : BEGDA
        'Descrip.: Fecha Liquidacion
        '          Ultimo dia del periodo de liquidacion
        BEGDA = Format_Fecha(PeriodoHasta, 1)
        
        'Estr.   : P9999
        'Campo   : ENDDA
        'Descrip.: Fecha Liquidacion
        '          Ultimo dia del periodo de liquidacion
        ENDDA = Format_Fecha(PeriodoHasta, 1)
        
        '---------------------------------------------------------
        'Busco todos los items del empleado definidos en el mapeo
            
        StrSql = " SELECT * FROM infotipos_mapeo WHERE tablaref='Conversion' "
        
        OpenRecordset StrSql, rs_Acum
        
        Do Until rs_Acum.EOF
        
            StrSql = " SELECT * FROM acu_mes "
            StrSql = StrSql & " WHERE ternro = " & Tercero
            StrSql = StrSql & "   AND acunro = " & rs_Acum!codinterno
            StrSql = StrSql & "   AND amanio = " & Year(Fecha_Desde)
            StrSql = StrSql & "   AND ammes  = " & Month(Fecha_Desde)
            
            OpenRecordset StrSql, rs_Consult
            
            If Not rs_Consult.EOF Then
                Salida = PERNR
                Salida = Salida & Separador & BEGDA
                Salida = Salida & Separador & ENDDA
        
                'Estr.   : P9999
                'Campo   : LGART
                'Descrip.: CC-Nomina
                '          Mapeo tabla Conversiones
                Salida = Salida & Separador & rs_Acum!codexterno
        
                'Estr.   : P9999
                'Campo   : BETPE
                'Descrip.: Importe por Unidad
                If CDbl(rs_Consult!amcant) > 0 Then
                   Salida = Salida & Separador & FormatNumber((CDbl(rs_Consult!ammonto) / CDbl(rs_Consult!amcant)), 2)
                Else
                   Salida = Salida & Separador & ""
                End If
        
                'Estr.   : P9999
                'Campo   : ANZHL
                'Descrip.: Ctd.
                If CDbl(rs_Consult!amcant) > 0 Then
                   Salida = Salida & Separador & rs_Consult!amcant
                Else
                   Salida = Salida & Separador & ""
                End If
                    
                'Estr.   : P9999
                'Campo   : BETRG
                'Descrip.: Importe
                Salida = Salida & Separador & rs_Consult!ammonto
                    
                ' Escribo en el erchivo
                fExport.Writeline Salida
                
            End If
            
            rs_Consult.Close
        
            rs_Acum.MoveNext
        Loop
        
        rs_Acum.Close
            
    MyCommitTrans
    
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Exportado satisfactoriamente"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " no exportado para empleado " & Legajo
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
End If

Fin:
    Exit Sub
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Infotipo " & NroInfotipo & " Abortado por Error"
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Decripcion: " & Err.Description
    Flog.Writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub





